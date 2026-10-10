# Email Archive Tool

A free, fully local Windows tool that connects to Outlook.com (or any IMAP mailbox) and archives selected emails into a government record format:

- **One PDF per conversation thread** - each message on its own page, with a title, form fields and the full message body
- Saved **attachments** (real files only; inline logo/signature images embedded in the body are dropped)
- Optional original `.eml` copies
- A **CSV log with one row per message** - every message of a thread shares the same `Ref No.`, showing who sent to whom at each step of the thread
- An **Excel export** (`log.xlsx`)

No data ever leaves the computer except the IMAP connection to the user's own mailbox.

> **Email Archive Tool 1.0.0.0** - Windows 10/11. Open source (MIT).

## Download (Windows - easiest, nothing to install)

| Download | What it is |
|---|---|
| **[EmailArchiveTool.exe](https://github.com/nassertaher/Email-Archive-Tool-1.0.0.0/releases/latest/download/EmailArchiveTool.exe)** | The program (single file, no Python needed) |
| **[EmailArchiveTool_Manual.pdf](https://github.com/nassertaher/Email-Archive-Tool-1.0.0.0/releases/latest/download/EmailArchiveTool_Manual.pdf)** | The step-by-step manual |
| **[EmailArchiveTool-1.0.0.0.zip](https://github.com/nassertaher/Email-Archive-Tool-1.0.0.0/releases/latest/download/EmailArchiveTool-1.0.0.0.zip)** | Both files together |

**How to use:** download `EmailArchiveTool.exe` and the manual, put them in the same folder, then double-click the exe. That's it - on the first run it creates an **Email Archive Tool** shortcut on your Desktop and keeps its data next to itself. Uninstall any time from the app's **Help** sidebar (it deletes its data, the shortcut and the program).

All versions: **[Releases](https://github.com/nassertaher/Email-Archive-Tool-1.0.0.0/releases)**.

## Features

- Outlook.com OAuth2 sign-in (device-login code) or app-password sign-in for Gmail / Yahoo / any IMAP server
- Conversation **threads**: archiving any message of a thread archives the whole thread as one record; a thread extended later reuses the same `Ref No.` and only appends the new messages
- **Incremental re-archive** (a new reply extends the existing PDF, never duplicates it)
- Hidden-senders list (`exclude_keywords`) - hides by **sender name and email address**, editable with a checkbox browser
- Bilingual records with automatic **Arabic rendering** in PDFs (Windows Arial/Tahoma auto-detected)
- Full **dark/light GUI theme** (remembered)
- **Archive browser** in the GUI - read, sort, filter and **edit** the register inside the app; `log.csv` and `log.xlsx` are rewritten together on every save
- **CSV import / export** - merge another CSV into the register (duplicate `Ref No.` rows are skipped) and export exactly the filtered view to CSV or Excel
- **Dashboard** - totals, a per-agency and per-month breakdown, and the list of outstanding actions at a glance
- Ready-made `.exe` (no Python needed) and a step-by-step **PDF manual** with screenshots for every screen

## Quick start

### A. Ready-made `.exe` (no Python needed)

Download `EmailArchiveTool.exe` from the **Releases** page of this repository, plus the `EmailArchiveTool_Manual.pdf` manual. Put both in any folder and double-click the exe.

- On the **first run** the program adds an **Email Archive Tool** shortcut on the Desktop (once) and creates its data next to the program: `config.json`, `token_cache.json`, the log files (`log.csv`, `log.xlsx`), `seen_ids.json` and the `PDF`, `EML` and `Attachments` folders.
- **Uninstall**: open the **Help** sidebar, press **Uninstall (remove all data)**. It asks whether to also delete the archived emails, then removes the settings, sign-in cache, logs, ledger, the Desktop shortcut and the program file itself on exit.

### B. From source

```bat
cd <this-folder>
python -m pip install -r requirements.txt
python gui.py          :: graphical interface
python archive.py      :: command-line menu
```

Requires Python 3.10+.

## 1. One-time Microsoft sign-in setup (OAuth2)

Outlook.com no longer accepts passwords or app passwords for IMAP - Microsoft requires OAuth2. Register a free app in the Azure portal once (about 5 minutes); the tool then signs in with a browser code and remembers you.

1. Open https://portal.azure.com and sign in with your @outlook.com account.
2. Search for **App registrations** -> **New registration**:
   - Name: anything, e.g. `Mail Archive Tool`
   - Supported account types: **Personal Microsoft accounts only**
   - Redirect URI: leave empty -> click **Register**
3. On the **Overview** page copy the **Application (client) ID**.
4. In `config.json`, put that value in `"client_id"` (keep the quotes).
5. In the app registration open **Authentication** -> under *Advanced settings* set **Allow public client flows** to **Enable** -> **Save**. (Required - otherwise the sign-in code is refused.)
6. Optional: **API permissions** -> **Add a permission** -> **APIs my organization uses** -> search *Office 365 Exchange Online* -> **Delegated permissions** -> tick **IMAP.AccessAsUser.All** -> **Add permissions**. If you cannot find that API, skip this step - the sign-in screen will ask for the permission automatically.

The first run prints:

```
To sign in, use a web browser to open the page https://microsoft.com/devicelogin and enter the code XXXXXXX to continue.
```

Open that page on any device, enter the code, sign in and approve. The sign-in is cached in `token_cache.json`; later runs connect silently (you re-approve roughly once a month).

## 2. Configure

Edit `config.json` (created automatically on first run if missing):

| Field | Meaning |
|---|---|
| `email` | Your mailbox address. |
| `client_id` | Application (client) ID from your Azure app - required for Outlook.com (section 1). |
| `tenant` | Sign-in endpoint: `common` (default, works for work and personal accounts) or `consumers` (personal accounts only; the tool also falls back to it automatically when Microsoft asks for `/consumers`). |
| `app_password` | Only for non-Microsoft IMAP servers. Outlook.com ignores it (basic auth is disabled). |
| `imap_server` | `outlook.office365.com` (change only for other providers). |
| `imap_port` | `993` (SSL). |
| `fetch_limit` | How many of the newest emails to fetch and **show** (default 50). Messages hidden by `exclude_keywords` do not count - the tool fetches older emails until the limit of non-excluded messages is reached. |
| `thread_view` | `true` (default) groups replies into conversation threads in the list; `false` shows a plain flat list. |
| `dark_mode` | `true` starts the GUI in dark mode; `false` (default) uses the light theme. The GUI writes this automatically when you press the theme switch in the Help sidebar. |
| `exclude_keywords` | List of words - any email whose **sender name, sender address or subject** contains one of these words is skipped during fetch (never shown, never archived). Edit freely, `[]` disables. ASCII words match whole words only (so `info` does not hide `information`), Arabic words match anywhere in the text. Hiding a sender saves **both its name and its email address** here, so the exact address is always caught even if the display name is written differently. Use menu **k** to add/remove senders with checkboxes - it writes this list for you. |
| `archive_folder` | Root folder for all output (defaults to the folder containing the program). |
| `save_eml` | `true`/`false` - also keep the original `.eml` file. |
| `agency_map` | Sender domain -> agency name for the **Agency** field (for example `"example.gov.sa": "Example Agency"`). Unknown domains fall back to `Other Government Agency`. |
| `default_priority` | Note appended to the record (default `Medium`). |

## 3. Command line

```bat
python archive.py
```

Menu:
1. **fetch emails** - connect, list the newest messages (`[n] date | sender | subject | attachments`); replies of the same conversation are grouped under a `Thread n (k messages):` header; messages matching `exclude_keywords` in `config.json` are hidden (`Skipped N excluded message(s)...`)
2. **archive selected** - type the numbers to archive, e.g. `1,3,5` (or `all`); `t2` archives every message of thread 2. Selecting any message of a thread archives the **whole thread** as one record; a thread is skipped only when every one of its messages is already archived. Each message becomes its own log row under the shared thread ref; threads extended later reuse the same ref and append only the new rows.
3. **export Excel** - write `log.xlsx` from `log.csv`
4. **quit**

k. **keywords** - **checkbox list of the fetched emails**: type row numbers to toggle (`1,3`), `t2` toggles whole thread 2, `k1` toggles an "other" saved keyword, `a` toggles all, `x` clears all, Enter or `q` saves. Every checked email **adds its sender name and email address to `exclude_keywords`** and writes them into `config.json` (row shows `-> hides "Sender Name", "sender@domain"`), so that sender stays hidden on every future fetch (and doesn't count toward `fetch_limit`). Uncheck to remove the sender again - no separate per-email list, one mechanism only.

Other options: `python archive.py --config D:\path\to\config.json` to use a different config file.

## 4. Graphical interface

The same tool has a **windowed GUI** - the command-line tool above stays unchanged and fully functional.

```bat
python gui.py          :: or:  python gui.py --config D:\path\to\config.json
```

- **Settings bar** - the email address up top plus a **Settings...** button with a dialog for everything else: a **mail-provider dropdown** (Outlook.com / Gmail / Yahoo / Other) that pre-fills the IMAP server, port and SSL, a **sign-in method** line that changes per provider, the **app password** field (Gmail and Yahoo - with links to create one), the **Azure application (client) ID** and tenant (Outlook.com only), fetch limit and the conversation-threads toggle. Every field has plain-text guidance and clickable links, all saved to `config.json`. Changing the email or the application ID warns that the Microsoft sign-in will need to be re-done on the next fetch.
- **Fetch emails** - runs in a background thread; every status line appears in the log pane while the list fills with the conversation threads.
- **Keep / Hide checkboxes** - click the **Keep** column of a thread to mark it for archiving (the whole thread); click the **Hide** column of a sender to add its **name and email address** to `exclude_keywords` immediately. Archived messages show a tick in the Keep column.
- **Filter box** above the list - type any text to keep only messages whose **sender or subject** contains it. Unread messages (IMAP `\Seen` flag) are shown in **bold red**.
- **Files tab** - click a message (or thread) and the Files tab lists its attachments plus the shared thread PDF. Double-click to open; an *Open attachments folder* button is there too.
- **Archive selected** (whole threads, incremental) and **Export Excel**, plus a hidden-senders panel to remove or clear keywords.
- **Help sidebar** - **Open PDF Manual**, **Use dark mode / Use light mode**, **About**, and **Uninstall (remove all data)**.

## 5. Archive browser, dashboard and export

The GUI adds read-and-edit views that work directly on `log.csv` / `log.xlsx`:

- **Archive tab** - shows the whole register (one row per message). Click any heading to sort; filter by free text, transaction-date range, agency, action, or whether the row has attachments.
- **Selected record** - pick a row to see all its fields plus the full message body (read from the saved `.eml`). Edit **Agency**, **Action**, **Response Date** and **Notes** and press **Save changes** - `log.csv` and `log.xlsx` are written together. **Open PDF**, **Open EML** and **Open attachments folder** are next to it.
- **Import CSV** - merge any CSV into the register. Columns are matched by name (`Ref`, `Subject`, `Agency`, `Date`, ...); rows whose `Ref No.` already exists are skipped, so imports are safe to repeat.
- **Export view** - save exactly the rows currently matching the filters to a new `.csv` or `.xlsx`.
- **Dashboard tab** - totals (records, threads, with attachments, outstanding), a breakdown **by agency** and **by month**, and the list of **outstanding actions** (rows still missing an Action or Response Date).
- **Emails tab filters** - besides the text box, filter the live list by transaction **date range**, **unread only**, **has attachment** and **archived / not archived**; **Clear all** resets every filter.

## 6. Output

```
<folder>\
|-- log.csv              (the bilingual record register)
|-- log.xlsx             (Excel view, auto-updated after each run)
|-- seen_ids.json        (de-duplication ledger)
|-- PDF\GOV-2026-001.pdf (one styled PDF per thread - a page per message,
|                         long messages continue onto the next page)
|-- Attachments\name_YYYY-MM-DD.ext   (real attachments only - inline
|                         logo/signature images are never saved or listed)
`-- EML\GOV-2026-001.eml, GOV-2026-001-2.eml, ... (originals, if save_eml is true)
```

The printed body is **cleaned of Outlook web "insert link" clutter** - the block of OneDrive file-share entries (`1drv.ms` links with Microsoft icon/thumbnail URLs) that Outlook.com appends after the message text is removed, so only the real body appears in the PDF. Already-made PDFs keep their old text; delete `seen_ids.json` and re-archive to regenerate them cleaned.

Reference numbers are sequential per year: `GOV-YYYY-NNN`. The log gets **one row per message** - everyone who sent something in that thread appears as its own row with its own date, `Sender`, `Recipients` and `CC`, and every row of the same thread shares the same `Ref No.` (the thread ID). All rows of a thread point to the same per-thread PDF. When a thread gets a *new* reply later, re-archiving it **reuses the same Ref No.**, grows the PDF with the new message's page and appends a row only for the new message. Emails archived individually *before* threads were enabled keep their old rows.

### Field mapping (paper form <-> CSV columns)

| CSV column | Source | Notes |
|---|---|---|
| `Ref No.` | auto | `GOV-YYYY-NNN` |
| `Agency` | `agency_map[sender domain]` | fallback `Other Government Agency` |
| `Subject` | `Subject:` header | `Re:` prefix kept |
| `Transaction Date` | `Date:` header | `YYYY-MM-DD` |
| `Action` / `Response Date` | - | written as `(fill manually)` - complete by hand |
| `Sender Name` | `From:` display name | falls back to address |
| `Sender Email` | `From:` address | |
| `Recipients` | `To:` | display name preferred; `-` if none |
| `CC` | `Cc:` | `-` if none |
| `Attachments` | message parts | **this message's** saved names, `;`-separated; `-` if none. Inline images (`Content-Disposition: inline` - logo/signature pictures embedded in the body) are never saved, so they don't appear here or in the PDF. |
| `Email PDF` | auto | `PDF/GOV-2026-001.pdf` - one PDF per thread, shared by all its rows |
| `Notes` | auto | `Received from {firstname}, subject: {subject}` / `...reply on same thread` for replies |

Empty fields are written as an em dash, matching the paper form.

## Troubleshooting

| Symptom | Fix |
|---|---|
| `Basic authentication is disabled` | Outlook.com requires OAuth2 - complete section 1 (put the Azure app `client_id` in `config.json`). |
| `AADSTS50020 ... tenant 'Microsoft Services'` | You cannot sign into the Azure portal from the browser. Open https://portal.azure.com in a **new InPrivate/incognito window** and sign in there. |
| Sign-in fails with `AADSTS7000218` or similar | Enable **Allow public client flows** for the Azure app (section 1, step 5) and run again. |
| Sign-in fails with `AADSTS9002346 ... use the /consumers endpoint` | The app is "Personal Microsoft accounts only" - set `"tenant": "consumers"` in `config.json` (the tool also retries `/consumers` automatically). |
| `User is authenticated but not connected` | A **known Microsoft server-side regression** for Outlook.com IMAP + OAuth2 (reported repeatedly from Dec 2024 through 2026; also hits Thunderbird, FairEmail, etc.). The OAuth sign-in succeeds but Outlook rejects the IMAP session bind. **The tested workaround: open Outlook settings -> Mail -> Forwarding and IMAP and turn "Let devices and apps use IMAP" OFF, then ON again**, wait about a minute and retry. If it still fails, delete `token_cache.json` next to the program and sign in again. |
| The sign-in code expired | Run again - codes are valid for about 15 minutes. |
| `Cannot connect... firewall/VPN` | Port 993 must be open. Try another network. |
| Manual button says "Manual not found" | Put `EmailArchiveTool_Manual.pdf` in the same folder as the program, then press the button again. |
| Arabic looks wrong in PDFs | The tool uses Windows Arial automatically; on minimal systems install Tahoma or DejaVu Sans. |
| Emails archived twice | Delete `seen_ids.json` only if you *want* to re-archive; otherwise leave it. |
| `msal` / `reportlab is not installed` | Re-run `python -m pip install -r requirements.txt`. |
| Using Gmail | In Settings pick **Gmail**, then put a Google **app password** (enable 2-step verification, then create one at https://support.google.com/mail/answer/185833) in the **App password** field - your normal Gmail password will not work. |
| Using Yahoo | In Settings pick **Yahoo Mail** and use a Yahoo **app password** (Yahoo Account Security). |
| Other IMAP servers | In Settings pick **Other IMAP server** and fill in the server, port, SSL and password your provider gives you - no Azure app ID needed. |
| Console shows `?` for Arabic | Windows console font limitation only; files on disk are correct UTF-8. |

## Security notes

- `token_cache.json` contains your Microsoft sign-in refresh token, and `config.json` may contain a saved app password (non-Microsoft servers only) - keep the folder private.
- Everything (PDF, CSV, EML, attachments) is written locally to `archive_folder`; nothing is uploaded.
- The only network connections are to your own mail provider (IMAP over SSL on port 993) and Microsoft's OAuth2 sign-in endpoints.

## Building the `.exe` (PyInstaller)

```bat
python -m pip install -r requirements.txt pyinstaller
python build.py
:: output: dist\EmailArchiveTool.exe  (single self-contained Windows file)
```

## License

MIT - see [LICENSE](LICENSE). Copyright (c) 2026 Eng. Nasser Taher.