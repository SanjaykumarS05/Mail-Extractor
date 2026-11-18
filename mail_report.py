"""
Mail Report Tool
----------------
Connects to your mailbox over IMAP (read-only, mails stay unread), finds every
mail from (or to) a given address inside a custom date range, and writes two
CSV files:

  1. <name>_mails.csv          -> one row per mail, sorted by date/time
  2. <name>_daily_summary.csv  -> one row per day in the range (mail count,
                                  first/last time, subjects)

Run with no arguments for the window (GUI).  Run with --cli for a text prompt.
Only the Python standard library is used - nothing to install.
"""

import csv
import email
import html
import imaplib
import json
import os
import re
import ssl
import subprocess
import sys
import threading
import queue
from datetime import datetime, date, timedelta
from email import policy
from email.utils import parsedate_to_datetime, getaddresses, parseaddr

APP_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(APP_DIR, "settings.json")
DEFAULT_OUTPUT = os.path.join(APP_DIR, "reports")
MAX_BODY_CHARS = 30000  # Excel cell limit is 32767
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

KNOWN_HOSTS = {
    "gmail.com": "imap.gmail.com",
    "googlemail.com": "imap.gmail.com",
    "outlook.com": "outlook.office365.com",
    "hotmail.com": "outlook.office365.com",
    "live.com": "outlook.office365.com",
    "yahoo.com": "imap.mail.yahoo.com",
    "yahoo.in": "imap.mail.yahoo.com",
    "zoho.com": "imap.zoho.com",
    "zohomail.in": "imap.zoho.in",
}


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def parse_user_date(text):
    """Accept DD-MM-YYYY, DD/MM/YYYY or DD.MM.YYYY (also YYYY-MM-DD)."""
    text = text.strip()
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Invalid date '{text}'. Use DD-MM-YYYY, e.g. 01-11-2026")


def imap_date(d):
    return f"{d.day:02d}-{MONTHS[d.month - 1]}-{d.year}"


def guess_imap_host(address):
    """Guess the IMAP server from the e-mail domain (known list, then MX lookup)."""
    domain = address.split("@")[-1].strip().lower()
    if not domain:
        return ""
    if domain in KNOWN_HOSTS:
        return KNOWN_HOSTS[domain]
    mx = ""
    try:
        flags = 0x08000000 if os.name == "nt" else 0  # CREATE_NO_WINDOW
        out = subprocess.run(["nslookup", "-type=mx", domain], capture_output=True,
                             text=True, timeout=10, creationflags=flags).stdout
        found = re.findall(r"mail exchanger\s*=\s*(\S+)", out)
        if found:
            mx = found[0].rstrip(".").lower()
    except Exception:
        pass
    if "google" in mx:
        return "imap.gmail.com"
    if "outlook" in mx or "office365" in mx:
        return "outlook.office365.com"
    if "zoho" in mx:
        return "imap.zoho.in" if mx.endswith(".in") else "imap.zoho.com"
    return mx or f"mail.{domain}"


def html_to_text(raw):
    raw = re.sub(r"(?is)<(script|style).*?</\1>", " ", raw)
    raw = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>", "\n", raw)
    raw = re.sub(r"<[^>]+>", " ", raw)
    return html.unescape(raw)


def clean_text(text):
    text = text.replace("\r", "")
    text = re.sub(r"[ \t ]+", " ", text)
    text = re.sub(r"\n\s*\n+", "\n", text)
    text = text.strip()
    if len(text) > MAX_BODY_CHARS:
        text = text[:MAX_BODY_CHARS] + " ...[truncated]"
    return text


def safe_header(msg, name):
    try:
        value = msg.get(name)
        return str(value).strip() if value is not None else ""
    except Exception:
        try:
            return str(msg.get_all(name, [""])[0])
        except Exception:
            return ""


def get_body_and_attachments(msg):
    plain, htm, attachments = [], [], []
    for part in msg.walk():
        if part.is_multipart():
            continue
        disposition = (part.get_content_disposition() or "").lower()
        filename = part.get_filename()
        if disposition == "attachment" or (filename and disposition != "inline"):
            attachments.append(filename or "unnamed")
            continue
        ctype = part.get_content_type()
        if ctype not in ("text/plain", "text/html"):
            if filename:
                attachments.append(filename)
            continue
        try:
            content = part.get_content()
        except Exception:
            payload = part.get_payload(decode=True) or b""
            content = payload.decode(part.get_content_charset() or "utf-8", "replace")
        (plain if ctype == "text/plain" else htm).append(str(content))
    body = "\n".join(plain) if plain else html_to_text("\n".join(htm))
    return clean_text(body), attachments


def format_addresses(value):
    pairs = getaddresses([value]) if value else []
    return ", ".join(f"{n} <{a}>" if n else a for n, a in pairs if a or n)


def parse_list_line(line):
    """Parse one IMAP LIST response line -> (flags, folder name)."""
    if isinstance(line, bytes):
        line = line.decode("utf-8", "replace")
    m = re.match(r'\((?P<flags>[^)]*)\)\s+(?P<delim>"[^"]*"|NIL)\s+(?P<name>.+)$', line)
    if not m:
        return "", None
    name = m.group("name").strip()
    if name.startswith('"') and name.endswith('"'):
        name = name[1:-1].replace('\\"', '"')
    return m.group("flags").lower(), name


def quote_folder(name):
    return '"' + name.replace("\\", "\\\\").replace('"', '\\"') + '"'


def safe_filename(text):
    return re.sub(r"[^A-Za-z0-9._@-]+", "_", text)[:80]


# --------------------------------------------------------------------------
# Core work
# --------------------------------------------------------------------------
def run_report(cfg, log=print):
    """cfg keys: email, password, senders, start, end, host, port, use_ssl,
    ignore_cert, all_folders, folder, direction, output_dir."""
    start = parse_user_date(cfg["start"])
    end = parse_user_date(cfg["end"])
    if end < start:
        raise ValueError("End date is before start date.")
    senders = [s.strip() for s in re.split(r"[,;\s]+", cfg["senders"]) if s.strip()]
    if not senders:
        raise ValueError("Enter at least one sender e-mail address.")
    host = cfg["host"].strip() or guess_imap_host(cfg["email"])
    port = int(cfg.get("port") or (993 if cfg.get("use_ssl", True) else 143))

    ctx = ssl.create_default_context()
    if cfg.get("ignore_cert"):
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE

    log(f"Connecting to {host}:{port} ...")
    if cfg.get("use_ssl", True):
        conn = imaplib.IMAP4_SSL(host, port, ssl_context=ctx, timeout=60)
    else:
        conn = imaplib.IMAP4(host, port, timeout=60)
        if "STARTTLS" in conn.capabilities:
            conn.starttls(ssl_context=ctx)
    try:
        conn.login(cfg["email"].strip(), cfg["password"])
    except imaplib.IMAP4.error as e:
        raise RuntimeError(f"Login failed: {e}. Check e-mail/password "
                           "(Gmail/Outlook need an App Password).") from None
    log("Logged in.")

    # Which folders to search
    if cfg.get("all_folders"):
        typ, data = conn.list()
        folders = []
        for line in data or []:
            flags, name = parse_list_line(line)
            if name and "\\noselect" not in flags and "\\nonexistent" not in flags:
                folders.append(name)
        # On Gmail "All Mail" already contains everything - avoid duplicates
        all_mail = [f for f in folders if f.lower().endswith("/all mail")]
        if all_mail:
            folders = all_mail + [f for f in folders if "spam" in f.lower()]
    else:
        folders = [cfg.get("folder") or "INBOX"]
    log(f"Folders to search: {len(folders)}")

    # widen by 1 day each side for time-zone differences; exact filter later
    since = imap_date(start - timedelta(days=1))
    before = imap_date(end + timedelta(days=2))
    keys = {"from": ["FROM"], "to": ["TO"], "both": ["FROM", "TO"]}[cfg.get("direction", "from")]

    rows, seen = [], set()
    for folder in folders:
        try:
            typ, _ = conn.select(quote_folder(folder), readonly=True)
            if typ != "OK":
                log(f"  skip {folder} (cannot open)")
                continue
        except Exception as e:
            log(f"  skip {folder} ({e})")
            continue
        uids = set()
        for sender in senders:
            for key in keys:
                typ, data = conn.uid("SEARCH", None, key, f'"{sender}"',
                                     "SINCE", since, "BEFORE", before)
                if typ == "OK" and data and data[0]:
                    uids.update(data[0].split())
        if not uids:
            continue
        log(f"  {folder}: {len(uids)} candidate mail(s)")
        for i, uid in enumerate(sorted(uids, key=int), 1):
            typ, data = conn.uid("FETCH", uid, "(INTERNALDATE BODY.PEEK[])")
            raw, internal = None, None
            for item in data or []:
                if isinstance(item, tuple):
                    raw = item[1]
                    internal = imaplib.Internaldate2tuple(item[0])
            if raw is None:
                continue
            msg = email.message_from_bytes(raw, policy=policy.default)
            msg_id = safe_header(msg, "Message-ID") or f"{folder}:{uid.decode()}"
            if msg_id in seen:
                continue

            dt = None
            try:
                dt = parsedate_to_datetime(safe_header(msg, "Date"))
            except Exception:
                pass
            if dt is None and internal:
                dt = datetime(*internal[:6])
            if dt is None:
                continue
            dt = dt.astimezone() if dt.tzinfo else dt
            if not (start <= dt.date() <= end):
                continue
            seen.add(msg_id)

            from_name, from_addr = parseaddr(safe_header(msg, "From"))
            body, attachments = get_body_and_attachments(msg)
            rows.append({
                "Date": dt.strftime("%d-%m-%Y"),
                "Day": dt.strftime("%A"),
                "Time": dt.strftime("%I:%M:%S %p"),
                "From Name": from_name,
                "From Email": from_addr,
                "To": format_addresses(safe_header(msg, "To")),
                "CC": format_addresses(safe_header(msg, "Cc")),
                "Subject": safe_header(msg, "Subject"),
                "Attachments": ", ".join(attachments),
                "Attachment Count": len(attachments),
                "Body": body,
                "Folder": folder,
                "Message-ID": msg_id,
                "_sort": dt.replace(tzinfo=None),
            })
            if i % 25 == 0:
                log(f"    read {i}/{len(uids)} ...")
    try:
        conn.logout()
    except Exception:
        pass

    rows.sort(key=lambda r: r["_sort"])
    log(f"Total mails in range: {len(rows)}")

    # ---- write CSVs ----
    out_dir = cfg.get("output_dir") or DEFAULT_OUTPUT
    os.makedirs(out_dir, exist_ok=True)
    tag = (f"{safe_filename(senders[0])}{'_and_more' if len(senders) > 1 else ''}"
           f"_{start:%d-%m-%Y}_to_{end:%d-%m-%Y}")
    mails_csv = os.path.join(out_dir, f"{tag}_mails.csv")
    daily_csv = os.path.join(out_dir, f"{tag}_daily_summary.csv")

    fields = ["S.No", "Date", "Day", "Time", "From Name", "From Email", "To", "CC",
              "Subject", "Attachments", "Attachment Count", "Body", "Folder", "Message-ID"]
    with open(mails_csv, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        for n, r in enumerate(rows, 1):
            w.writerow({**r, "S.No": n})

    with open(daily_csv, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["Date", "Day", "Mail Count", "First Mail Time",
                    "Last Mail Time", "Attachments", "Subjects"])
        d = start
        while d <= end:
            key = d.strftime("%d-%m-%Y")
            day_rows = [r for r in rows if r["Date"] == key]
            w.writerow([
                key, d.strftime("%A"), len(day_rows),
                day_rows[0]["Time"] if day_rows else "",
                day_rows[-1]["Time"] if day_rows else "",
                sum(r["Attachment Count"] for r in day_rows),
                " | ".join(r["Subject"] for r in day_rows),
            ])
            d += timedelta(days=1)
        w.writerow([])
        w.writerow(["TOTAL", "", len(rows), "", "",
                    sum(r["Attachment Count"] for r in rows), ""])

    log(f"Saved: {mails_csv}")
    log(f"Saved: {daily_csv}")
    return mails_csv, daily_csv, len(rows)


# --------------------------------------------------------------------------
# Settings (password is never saved)
# --------------------------------------------------------------------------
def load_settings():
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(cfg):
    keep = {k: v for k, v in cfg.items() if k != "password"}
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(keep, f, indent=2)
    except Exception:
        pass


# --------------------------------------------------------------------------
# GUI
# --------------------------------------------------------------------------
def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    s = load_settings()
    today = date.today()
    root = tk.Tk()
    root.title("Mail Report Tool - mails to CSV (per day)")
    root.geometry("760x640")

    v = {
        "email": tk.StringVar(value=s.get("email", "")),
        "password": tk.StringVar(),
        "senders": tk.StringVar(value=s.get("senders", "")),
        "start": tk.StringVar(value=s.get("start", today.replace(day=1).strftime("%d-%m-%Y"))),
        "end": tk.StringVar(value=s.get("end", today.strftime("%d-%m-%Y"))),
        "host": tk.StringVar(value=s.get("host", "")),
        "port": tk.StringVar(value=str(s.get("port", 993))),
        "folder": tk.StringVar(value=s.get("folder", "INBOX")),
        "output_dir": tk.StringVar(value=s.get("output_dir", DEFAULT_OUTPUT)),
        "direction": tk.StringVar(value=s.get("direction", "from")),
    }
    use_ssl = tk.BooleanVar(value=s.get("use_ssl", True))
    ignore_cert = tk.BooleanVar(value=s.get("ignore_cert", False))
    all_folders = tk.BooleanVar(value=s.get("all_folders", True))

    frm = ttk.Frame(root, padding=12)
    frm.pack(fill="both", expand=True)
    frm.columnconfigure(1, weight=1)

    def row(r, label, widget, extra=None):
        ttk.Label(frm, text=label).grid(row=r, column=0, sticky="w", pady=3)
        widget.grid(row=r, column=1, sticky="ew", pady=3)
        if extra:
            extra.grid(row=r, column=2, sticky="w", padx=4)

    row(0, "Your e-mail (mailbox):", ttk.Entry(frm, textvariable=v["email"]))
    row(1, "Your password:", ttk.Entry(frm, textvariable=v["password"], show="*"))
    row(2, "Sender e-mail(s):", ttk.Entry(frm, textvariable=v["senders"]),
        ttk.Label(frm, text="comma separated"))
    row(3, "From date (DD-MM-YYYY):", ttk.Entry(frm, textvariable=v["start"]))
    row(4, "To date (DD-MM-YYYY):", ttk.Entry(frm, textvariable=v["end"]))

    dir_box = ttk.Combobox(frm, textvariable=v["direction"], state="readonly",
                           values=["from", "to", "both"])
    row(5, "Mails:", dir_box,
        ttk.Label(frm, text="from = received from sender, to = sent to it"))

    def detect():
        v["host"].set(guess_imap_host(v["email"].get()) or "")
    row(6, "IMAP server:", ttk.Entry(frm, textvariable=v["host"]),
        ttk.Button(frm, text="Auto detect", command=detect))
    row(7, "Port:", ttk.Entry(frm, textvariable=v["port"], width=8))

    opts = ttk.Frame(frm)
    ttk.Checkbutton(opts, text="SSL", variable=use_ssl).pack(side="left")
    ttk.Checkbutton(opts, text="Ignore SSL certificate errors",
                    variable=ignore_cert).pack(side="left", padx=10)
    ttk.Checkbutton(opts, text="Search all folders", variable=all_folders).pack(side="left")
    row(8, "Options:", opts)
    row(9, "Folder (if not all):", ttk.Entry(frm, textvariable=v["folder"]))

    def browse():
        d = filedialog.askdirectory(initialdir=v["output_dir"].get() or APP_DIR)
        if d:
            v["output_dir"].set(d)
    row(10, "Save CSV to:", ttk.Entry(frm, textvariable=v["output_dir"]),
        ttk.Button(frm, text="Browse", command=browse))

    btn = ttk.Button(frm, text="Get Mails & Create CSV")
    btn.grid(row=11, column=0, columnspan=3, pady=10, sticky="ew")

    logbox = tk.Text(frm, height=14, wrap="word")
    logbox.grid(row=12, column=0, columnspan=3, sticky="nsew")
    frm.rowconfigure(12, weight=1)

    q = queue.Queue()

    def pump():
        try:
            while True:
                kind, payload = q.get_nowait()
                if kind == "log":
                    logbox.insert("end", payload + "\n")
                    logbox.see("end")
                elif kind == "done":
                    btn.config(state="normal")
                    mails_csv, daily_csv, count = payload
                    if messagebox.askyesno("Done", f"{count} mail(s) exported.\n\n"
                                           f"{mails_csv}\n{daily_csv}\n\nOpen folder?"):
                        os.startfile(os.path.dirname(mails_csv))
                elif kind == "error":
                    btn.config(state="normal")
                    messagebox.showerror("Error", payload)
        except queue.Empty:
            pass
        root.after(150, pump)

    def start_job():
        cfg = {k: var.get() for k, var in v.items()}
        cfg.update(use_ssl=use_ssl.get(), ignore_cert=ignore_cert.get(),
                   all_folders=all_folders.get())
        if not cfg["email"] or not cfg["password"] or not cfg["senders"]:
            messagebox.showwarning("Missing", "Fill your e-mail, password and sender e-mail.")
            return
        save_settings(cfg)
        btn.config(state="disabled")
        logbox.delete("1.0", "end")

        def work():
            try:
                q.put(("done", run_report(cfg, log=lambda m: q.put(("log", m)))))
            except Exception as e:
                q.put(("log", f"ERROR: {e}"))
                q.put(("error", str(e)))
        threading.Thread(target=work, daemon=True).start()

    btn.config(command=start_job)
    if not v["host"].get() and v["email"].get():
        detect()
    pump()
    root.mainloop()


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def run_cli():
    import getpass
    s = load_settings()

    def ask(label, key, default=""):
        d = s.get(key, default)
        val = input(f"{label}{f' [{d}]' if d != '' else ''}: ").strip()
        return val or str(d)

    cfg = {"email": ask("Your e-mail", "email")}
    cfg["password"] = getpass.getpass("Your password (hidden): ")
    cfg["senders"] = ask("Sender e-mail(s), comma separated", "senders")
    cfg["start"] = ask("From date DD-MM-YYYY", "start")
    cfg["end"] = ask("To date DD-MM-YYYY", "end")
    cfg["direction"] = ask("Mails from / to / both", "direction", "from").lower()
    cfg["host"] = ask("IMAP server", "host", guess_imap_host(cfg["email"]))
    cfg["port"] = ask("Port", "port", 993)
    cfg["use_ssl"] = True
    cfg["ignore_cert"] = ask("Ignore SSL certificate errors? y/n", "ignore_cert_cli", "n").lower() == "y"
    cfg["all_folders"] = ask("Search all folders? y/n", "all_folders_cli", "y").lower() == "y"
    cfg["folder"] = "INBOX"
    cfg["output_dir"] = ask("Save CSV to", "output_dir", DEFAULT_OUTPUT)
    save_settings({**cfg, "ignore_cert_cli": "y" if cfg["ignore_cert"] else "n",
                   "all_folders_cli": "y" if cfg["all_folders"] else "n"})
    try:
        run_report(cfg)
    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)


if __name__ == "__main__":
    if "--cli" in sys.argv:
        run_cli()
    else:
        try:
            run_gui()
        except ImportError:
            print("Tkinter not available - using text mode.\n")
            run_cli()
