# GrepMail 2.0

Monitors a mailbox and saves tender attachments into one local folder per reference.
Everything is recorded in a SQLite database (`GrepMail.db`); an Excel view (`GrepMail.xlsx`) is generated next to it and refreshed automatically (References, Keywords, Exclusions, Email Log,
Custom Log, Processing Log, Saved Files, Unmatched, Summary).

## How an email is judged (in this order)
1. Subject contains a **Reference**.
2. Subject contains no **exclusion word** (an exclusion always wins).
3. The sender's new text contains a **Body Keyword**.
4. Attachments are scanned, then saved using the **file name format** ({Sender}, {Reference}, {Date}, {Attachment Name}).

Subject layout: `Sender - Reference rest of the subject`. The dash `-` is the default separator; others (`–`, `||`, `|`, `/`, `_` and one custom entry) can be switched on in Settings > Subject Rules, and spaces around a separator are ignored. Per-reference search scope:
Subfolders only (default), Subfolders and Inbox, or Inbox only.

## Platforms
| | Windows | Debian / Ubuntu |
|---|---|---|
| Mail | Outlook (COM) or IMAP | IMAP (Gmail, Proton Mail Bridge, Microsoft 365, other) |
| Virus scan | Microsoft Defender | ClamAV |
| Auto start | Registry Run key | ~/.config/autostart |

## Run from source
    pip install -r requirements.txt
    python grepmail.py
On first run the app asks where to create the database. The Excel view can stay open while
GrepMail runs; edits are made in the app, not in the Excel file. If a file named `GrepMail.xlsx`
from an older version exists in that folder, rename or remove it first.

## Build installers
- Windows (open a terminal on a Windows laptop): `packaging\build_windows.bat` (needs Inno Setup 6).
- Debian (open a terminal on a Linux laptop): `packaging/build_deb.sh`.

## Security notes
Attachments are scanned in a temporary location before saving; threats are never saved.
Macro-enabled files: Allow, Allow and Flag, or Block; suspicious macros are quarantined
(oletools). No scan guarantees safety. Passwords are kept in the system keyring when available.

## Tests
    python tests/test_core.py
    python tests/test_imap.py
    python tests/test_update.py
    xvfb-run -a python tests/gui_smoke.py   # Linux, window test

## Updates
GrepMail checks the latest GitHub release once a day (Settings > Startup turns it off; Help > Check for Updates checks now). It shows the release notes and offers Update Now, Remind Me Later or Skip This Version. The installer is downloaded only from GitHub over HTTPS and its SHA-256 checksum is verified.

### Publishing a new version
1. Raise `APP_VERSION` in `grepmail.py` and `AppVersion` in `packaging/grepmail.iss`.
2. Build the installer (and the .deb), then run `gh release create vX.Y.Z <files> --title ... --notes ...`. The release notes are what users see in the update dialog.

## Importing references and keywords

**File > Import References and Keywords...** copies the Reference and Keyword lists (and exclusion words, when present) from another file into the open database. A preview shows what will be added or updated before anything is written; the source file is never changed.
Supported files: Excel (`.xlsx`, `.xlsm`) with or without a header row, CSV, and another `GrepMail.db`. Without a header row, columns are read as Reference, Local Path, Mail Folder, Deadline. A reference with no mail folder is imported as *Inbox Only*. `tools/import_data.py` offers the same import from the command line.

## License
GNU Affero General Public License v3.0 or later (AGPL-3.0-or-later). See the `LICENSE` file.
