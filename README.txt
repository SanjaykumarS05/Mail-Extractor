MAIL REPORT TOOL  -  mails from a sender -> CSV report per day
===============================================================

HOW TO RUN
  Double-click "Run Mail Report Tool.bat"   (window)
  or "Run (text mode).bat"                  (command prompt version)

FILL IN
  Your e-mail      : your mailbox, e.g. sanjaykumar@rsoft.in
  Your password    : your mailbox password (NOT saved anywhere)
  Sender e-mail(s) : e.g. rsoftindia1@gmail.com  (several: comma separated)
  From / To date   : DD-MM-YYYY, e.g. 01-11-2026 and 18-11-2026 (both days included)
  Mails            : from = mails received from the sender
                     to   = mails you sent to that address
                     both = both directions
  IMAP server      : click "Auto detect" (rsoft.in -> mail.rsoft.in), port 993, SSL on
  Search all folders: ticked = Inbox + all sub folders (duplicates removed)

Click "Get Mails & Create CSV".  Mails are opened read-only, so they stay unread.

OUTPUT (in the "reports" folder by default)
  <sender>_<from>_to_<to>_mails.csv
      S.No, Date, Day, Time, From Name, From Email, To, CC, Subject,
      Attachments, Attachment Count, Body, Folder, Message-ID
      (sorted by date and time)
  <sender>_<from>_to_<to>_daily_summary.csv
      one row for EVERY day in the range (0 if no mail):
      Date, Day, Mail Count, First Mail Time, Last Mail Time, Attachments, Subjects
      + a TOTAL row at the end
  Both open directly in Excel (UTF-8).

TROUBLESHOOTING
  Login failed        -> check password. Gmail / Outlook accounts need an
                         "App Password" (and IMAP enabled in mail settings).
  SSL / certificate   -> tick "Ignore SSL certificate errors".
  Connection refused  -> ask your mail admin for the IMAP server/port
                         (993 with SSL, or 143 with SSL unticked).

Needs Python 3 (already installed). No extra packages.
