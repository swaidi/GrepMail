# SPDX-License-Identifier: AGPL-3.0-or-later
"""
GrepMail  (v2.0)
================
Watches your mailbox and saves the attachments of tender (or any project) emails
into the right folder on your computer.

Core behavior
- Reads Outlook (Windows) or any IMAP mailbox (Gmail, Proton Mail Bridge, ...)
- For every reference you define, it watches a mail folder (default: that folder
  and its subfolders; optionally the Inbox as well, or the Inbox only)
- An incoming email is processed in this order:
    1. its subject contains a Reference
    2. its subject does NOT contain an exclusion word
    3. the sender's new text contains a body keyword
    4. allowed attachments are scanned for viruses and macros, then saved to the
       Local Path, named with your file name format
- Everything is recorded in a database (emails, processing results,
  saved files, a custom log of your own choosing)

Requirements:  see requirements.txt
Run:           python grepmail.py
"""

import os
import sys
import re
import json
import time
import queue
import sqlite3
import hashlib
import shutil
import threading
import datetime as dt
import tkinter as tk
from tkinter import filedialog, messagebox

try:
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import Font
except ImportError:
    print("Missing dependency: openpyxl. Install with: pip install openpyxl")
    sys.exit(1)

try:
    import ttkbootstrap as tb
    ttk = tb  # styled widgets (bootstyle=...) come from ttkbootstrap
    from ttkbootstrap.style import ThemeDefinition
    try:
        from ttkbootstrap.widgets import DateEntry
    except Exception:  # very old/new releases: fall back to a plain date box
        DateEntry = None
    try:
        from ttkbootstrap.utility import scale_size
    except Exception:
        def scale_size(widget, size):
            return size
except ImportError:
    print("Missing dependency: ttkbootstrap. Install with: pip install ttkbootstrap")
    sys.exit(1)

try:
    import pythoncom
    import win32com.client
    HAS_OUTLOOK = True
except ImportError:
    HAS_OUTLOOK = False

import gm_platform
import gm_security
import gm_mail
import gm_update

try:  # pywin32 needs this module when it converts Outlook dates
    import win32timezone  # noqa: F401
except Exception:
    pass

# =========================================================
# CONSTANTS
# =========================================================
APP_TITLE = "GrepMail"
APP_VERSION = "2.0.1"
INBOX_TOKEN = "(Inbox)"   # stored as the Mail Folder when only the Inbox is read

SHEET_TENDERS = "References"
SHEET_KEYWORDS = "Keywords"
SHEET_EXCLUSIONS = "Exclusions"
SHEET_EMAIL_LOG = "EmailLog"
SHEET_TENDER_LOG = "ProcessingLog"
SHEET_OFFERS = "SavedFiles"
SHEET_UNMATCHED = "Unmatched"
SHEET_SUMMARY = "Summary"
SHEET_CUSTOM = "CustomLog"

SCOPE_SUB = "Subfolders"
SCOPE_SUB_INBOX = "Subfolders And Inbox"
SCOPE_INBOX = "Inbox Only"
SCOPES = [SCOPE_SUB, SCOPE_SUB_INBOX, SCOPE_INBOX]

# every field the Email Log keeps (the Custom Log picks from these)
FIELD_CATALOG = [  # (key, label)
    ("LoggedAt", "Logged At"), ("Direction", "Direction"), ("EmailDate", "Email Date"),
    ("SenderName", "Sender Name"), ("SenderEmail", "Sender Email"), ("To", "To"), ("Cc", "Cc"),
    ("Subject", "Subject"), ("MailFolder", "Mail Folder"), ("Reference", "Reference"),
    ("Name", "Name (from subject)"), ("AttachmentCount", "Attachment Count"),
    ("AttachmentNames", "Attachment Names"), ("NewText", "New Message Text"),
    ("Result", "Result"), ("LocalPath", "Local Path"), ("EmailId", "Email Id"),
]
DEFAULT_CUSTOM_FIELDS = ["EmailDate", "Direction", "SenderEmail", "Subject", "Reference", "Result"]
EMAIL_LOG_HEADERS = [k for k, _ in FIELD_CATALOG]
TENDER_LOG_HEADERS = [
    "Timestamp", "Action", "Reference", "Subject", "LocalPath", "Details", "EmailId",
]
OFFER_HEADERS = [
    "ReceivedDate", "Reference", "Name", "FileNo", "Revision",
    "SavedFileName", "OriginalName", "FilePath", "SenderEmail",
    "FileHash", "Security", "Notes", "EmailId", "MatchedKeyword",
]
UNMATCHED_HEADERS = [
    "Timestamp", "SenderName", "SenderEmail", "Subject", "DetectedRef", "EmailId",
]
SUMMARY_HEADERS = [
    "Reference", "LocalPath", "Deadline", "DaysLeft", "Files",
    "Names", "LastSaved", "Status",
]

SCHEMA = {
    SHEET_TENDERS: ["Reference", "LocalPath", "MailFolder", "Scope", "Deadline"],
    SHEET_KEYWORDS: ["Keyword"],
    SHEET_EXCLUSIONS: ["Word"],
    SHEET_EMAIL_LOG: EMAIL_LOG_HEADERS,
    SHEET_CUSTOM: [],            # columns follow the Fields page
    SHEET_TENDER_LOG: TENDER_LOG_HEADERS,
    SHEET_OFFERS: OFFER_HEADERS,
    SHEET_UNMATCHED: UNMATCHED_HEADERS,
    SHEET_SUMMARY: SUMMARY_HEADERS,
}

ALLOWED_EXTENSIONS = {"pdf", "xls", "xlsx", "xlsm", "xlsb", "doc", "docx"}
INVALID_FILENAME_CHARS = '\\/:*?"<>|'

DEFAULT_POLL_SECONDS = 30
DEFAULT_WARN_DAYS = 3
DEFAULT_HINT_REGEX = ""   # optional, e.g. [A-Z]{2,5}-\d{3,}  (spots reference-like codes)
DEFAULT_HINT_PHRASES = []   # optional, e.g. invitation to tender

if getattr(sys, "frozen", False):
    CONFIG_FILE = os.path.join(gm_platform.config_dir(), "grepmail_config.json")
else:  # running from source: keep the settings next to the script
    CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "grepmail_config.json")
DEFAULT_NAME_SEPARATOR = " - "
DEFAULT_FILE_FORMAT = "{Sender} {Attachment Name}"


# =========================================================
# SMALL HELPERS
# =========================================================
def s(v):
    """Safe text: None -> '', everything else stripped str."""
    return "" if v is None else str(v).strip()


def safe(fn, default=""):
    try:
        v = fn()
        return default if v is None else v
    except Exception:
        return default


def date_str(v):
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, dt.date):
        return v.isoformat()
    return s(v)


def parse_date(v):
    text = date_str(v)
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(text[:10], fmt).date()
        except ValueError:
            continue
    return None


def parse_ts(v):
    if isinstance(v, dt.datetime):
        return v
    text = s(v)
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def days_left(deadline):
    d = parse_date(deadline)
    return None if d is None else (d - dt.date.today()).days


def now_str():
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def to_naive(when):
    return dt.datetime(when.year, when.month, when.day,
                       when.hour, when.minute, when.second)


def clean_filename(name):
    for ch in INVALID_FILENAME_CHARS:
        name = name.replace(ch, "_")
    return name


def get_extension(filename):
    return filename.rsplit(".", 1)[-1].lower() if "." in filename else ""


def is_allowed_file_type(ext):
    return ext.lower() in ALLOWED_EXTENSIONS


def unique_file_path(folder, filename):
    filename = clean_filename(filename)
    base, ext = os.path.splitext(filename)
    candidate = os.path.join(folder, filename)
    counter = 1
    while os.path.exists(candidate):
        candidate = os.path.join(folder, "%s (%d)%s" % (base, counter, ext))
        counter += 1
    return candidate


def file_sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def strip_reply_prefixes(subject):
    remaining = s(subject)
    while True:  # stacked "RE: FW: ..." prefixes
        low = remaining.lower()
        for prefix in ("re:", "fw:", "fwd:"):
            if low.startswith(prefix):
                remaining = remaining[len(prefix):].strip()
                break
        else:
            return remaining


def extract_name(subject, separator=DEFAULT_NAME_SEPARATOR):
    """The 'Name' part of a 'Name - Reference' subject (text before the separator).
    The chosen separator is tried first, then ' - ' and ' || '."""
    remaining = strip_reply_prefixes(subject)
    if not remaining:
        return ""
    for sep in [separator, " - ", " || "]:
        if sep and sep in remaining:
            return clean_filename(remaining[:remaining.find(sep)].strip())
    return clean_filename(remaining)


FORMAT_TOKENS = ("{Sender}", "{Reference}", "{Date}", "{Attachment Name}")
_OLD_TOKENS = {"{Name}": "{Sender}", "{Original}": "{Attachment Name}"}   # names used by earlier builds


def upgrade_format(template):
    for old, new in _OLD_TOKENS.items():
        template = s(template).replace(old, new)
    return template


def build_file_name(template, name, reference, date_text, original):
    """Applies the file name format. The original extension is always kept."""
    template = template if s(template) else DEFAULT_FILE_FORMAT
    template = upgrade_format(template)
    out = (template.replace("{Sender}", name).replace("{Reference}", reference)
           .replace("{Date}", date_text).replace("{Attachment Name}", original))
    out = re.sub(r"\s+", " ", out).strip(" -_")
    ext = os.path.splitext(original)[1]
    if ext and not out.lower().endswith(ext.lower()):
        out += ext
    return clean_filename(out or original)


def matched_exclusion(subject, words_lower):
    """First exclusion word found in the subject ('' if none)."""
    low = s(subject).lower()
    for w in words_lower:
        if w and w in low:
            return w
    return ""


def find_tender_reference(subject, tender_map):
    subj = s(subject).lower()
    for ref_upper in tender_map.keys():
        if ref_upper.lower() in subj:
            return ref_upper
    return ""


def matched_keyword(text, keywords_lower):
    """First keyword found in the text ('' if none)."""
    low = s(text).lower()
    for kw in keywords_lower:
        if kw in low:
            return kw
    return ""


def body_contains_keyword(body, keywords_lower):
    return bool(matched_keyword(body, keywords_lower))


# Where the quoted earlier conversation starts (everything from there is ignored)
_QUOTE_MARKERS = [
    re.compile(r"^[ \t]*-{2,}[ \t]*(?:original message|forwarded message)[ \t]*-{2,}", re.I | re.M),
    re.compile(r"^[ \t]*from[ \t]*:", re.I | re.M),            # Outlook reply/forward header
    re.compile(r"^[ \t]*\u0645\u0646[ \t]*:", re.M),                # Arabic "From:"
    re.compile(r"^[ \t]*_{10,}[ \t\r]*$", re.M),                  # Outlook separator line
    re.compile(r"^[ \t]*on[ \t].{5,200}?wrote:", re.I | re.M | re.S),  # Gmail style
]


def new_message_text(body):
    """What the sender typed in this message: the body up to the quoted history."""
    body = s(body)
    cut = len(body)
    for rx in _QUOTE_MARKERS:
        m = rx.search(body)
        if m and m.start() < cut:
            cut = m.start()
    return body[:cut]


def looks_like_tender(subject, regex, phrases):
    """Returns the detected reference text, '' if only wording matched,
    or None if the subject does not look like tender correspondence."""
    if regex:
        try:
            m = re.search(regex, subject, re.I)
            if m:
                return m.group(0)
        except re.error:
            pass
    low = subject.lower()
    for p in phrases:
        if p and p.lower() in low:
            return ""
    return None


REV_RE = re.compile(r"(?<![A-Za-z])(?:rev(?:ision)?|r)[\s._-]*0*(\d{1,2})(?!\d)", re.I)


def detect_revision(filename, subject, body):
    for text in (filename, subject, s(body)[:500]):
        m = REV_RE.search(text or "")
        if m:
            return "Rev %02d" % int(m.group(1))
    return ""


def email_key(direction, when_str, subject):
    return "T|%s|%s|%s" % (s(direction).title(), s(when_str), s(subject))


def check_folder(path):
    """Returns 'OK', 'No path', 'Missing' or 'Not writable'."""
    if not path:
        return "No path"
    if not os.path.isdir(path):
        return "Missing"
    test = os.path.join(path, ".tem_write_test_%d" % os.getpid())
    try:
        with open(test, "w") as f:
            f.write("x")
        os.remove(test)
        return "OK"
    except Exception:
        return "Not writable"


def get_sender_smtp(item):
    try:
        if s(item.SenderEmailType).upper() == "EX":
            sender = item.Sender
            if sender is not None:
                eu = sender.GetExchangeUser()
                if eu is not None and eu.PrimarySmtpAddress:
                    return eu.PrimarySmtpAddress
                return sender.PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
        return item.SenderEmailAddress or ""
    except Exception:
        return safe(lambda: item.SenderEmailAddress, "")


def attachment_names(item):
    names = []
    try:
        atts = item.Attachments
        for i in range(1, atts.Count + 1):
            names.append(atts.Item(i).FileName)
    except Exception:
        pass
    return names


# =========================================================
# CONFIG (per user, stored in %APPDATA%)
# =========================================================
_CONFIG_LOCK = threading.Lock()


def _default_provider():
    return "outlook" if (gm_platform.IS_WINDOWS and HAS_OUTLOOK) else "imap"


class Config:
    """Per-user settings (stored outside the database)."""

    def __init__(self):
        self.db_path = ""
        self.poll_seconds = DEFAULT_POLL_SECONDS
        self.warn_days = DEFAULT_WARN_DAYS
        self.hint_regex = DEFAULT_HINT_REGEX
        self.hint_phrases = list(DEFAULT_HINT_PHRASES)
        self.toasts = True
        self.scan_root = True
        self.theme = "light"
        self.provider = _default_provider()
        self.imap = {"preset": "Gmail", "host": "imap.gmail.com", "port": 993,
                     "user": "", "security": "ssl"}
        self.name_separator = DEFAULT_NAME_SEPARATOR
        self.file_format = DEFAULT_FILE_FORMAT
        self.custom_fields = list(DEFAULT_CUSTOM_FIELDS)
        self.custom_filter = {"direction": "Both", "only_reference": False,
                              "only_attachments": False, "contains": ""}
        self.scan_enabled = True
        self.require_scan = False
        self.macro_mode = "flag"        # allow | flag | block
        self.mark_web = True
        self.auto_monitor = False
        self.check_updates = True
        self.skipped_version = ""
        self.update_snooze_until = ""     # ISO time; no automatic prompt before it
        self.last_check = {"inbox": None, "sent": None}
        self.session_password = ""      # only kept in memory
        self._load()

    _SIMPLE = ("db_path", "poll_seconds", "warn_days", "hint_regex", "hint_phrases", "toasts",
               "scan_root", "theme", "provider", "imap", "name_separator", "file_format",
               "custom_fields", "custom_filter", "scan_enabled", "require_scan", "macro_mode",
               "mark_web", "auto_monitor", "check_updates", "skipped_version", "update_snooze_until")

    def _load(self):
        d = None
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                d = json.load(f)
        except Exception:
            return
        if not isinstance(d, dict):
            return
        for key in self._SIMPLE:
            if key in d:
                cur = getattr(self, key)
                val = d[key]
                if isinstance(cur, dict) and isinstance(val, dict):
                    cur.update(val)
                else:
                    setattr(self, key, val)
        self.file_format = upgrade_format(self.file_format) or DEFAULT_FILE_FORMAT
        self.poll_seconds = int(self.poll_seconds)
        self.warn_days = int(self.warn_days)
        if self.theme not in ("light", "dark"):
            self.theme = "light"
        if self.macro_mode not in ("allow", "flag", "block"):
            self.macro_mode = "flag"
        if self.provider not in ("outlook", "imap"):
            self.provider = _default_provider()
        self.last_check["inbox"] = d.get("last_check_inbox")
        self.last_check["sent"] = d.get("last_check_sent")

    def save(self):
        d = {k: getattr(self, k) for k in self._SIMPLE}
        d["last_check_inbox"] = self.last_check["inbox"]
        d["last_check_sent"] = self.last_check["sent"]
        try:
            with _CONFIG_LOCK:
                os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
                tmp = CONFIG_FILE + ".tmp"
                with open(tmp, "w", encoding="utf-8") as f:
                    json.dump(d, f, indent=2)
                os.replace(tmp, CONFIG_FILE)  # atomic: never leaves a half-written file
        except Exception:
            pass

    def get_last_check(self, which):
        iso = self.last_check.get(which)
        if iso:
            try:
                return dt.datetime.fromisoformat(iso)
            except Exception:
                pass
        return dt.datetime.now() - dt.timedelta(hours=1)

    def set_last_check(self, which, when):
        self.last_check[which] = when.isoformat()
        self.save()


# =========================================================
# DATABASE (SQLite) + EXCEL VIEW
# =========================================================
def view_path_for(db_path):
    """The Excel view file that sits next to the database file."""
    return os.path.splitext(db_path)[0] + ".xlsx"


def _sql_value(v):
    if v is None or isinstance(v, (int, float, str)):
        return v
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, dt.date):
        return v.isoformat()
    return str(v)


class SchemaError(Exception):
    """The file is not a GrepMail database in the current format."""

    def __init__(self, kind, message):
        super().__init__(message)
        self.kind = kind   # 'exists', 'old', 'view' or 'unreadable'


def inspect_database(path):
    """'missing', 'valid', 'old' (another format), 'view' (only the Excel view file of the
    same name exists) or 'unreadable'."""
    if not os.path.exists(path):
        return "view" if os.path.exists(view_path_for(path)) else "missing"
    try:
        with open(path, "rb") as f:
            if f.read(16) != b"SQLite format 3\x00":
                return "old"
        con = sqlite3.connect(path)
        try:
            cols = {r[1] for r in con.execute('PRAGMA table_info("%s")' % SHEET_TENDERS)}
        finally:
            con.close()
        return "valid" if {"Reference", "LocalPath", "MailFolder"} <= cols else "old"
    except Exception:
        return "unreadable"


class GrepDB:
    """The GrepMail database: one SQLite file (never opened by hand, so nothing can lock it)
    plus an Excel view file that is generated from it. Thread-safe. Supports batching so a
    whole polling cycle is committed once."""

    EXPORT_DELAY = 30.0      # seconds to wait after a change before refreshing the Excel view
    RETRY_DELAY = 30.0       # seconds to wait before retrying when the view is open in Excel

    def __init__(self, path, create=False):
        self.path = path
        self.view_path = view_path_for(path)
        self.lock = threading.RLock()
        self._batch_depth = 0
        self._dirty = False
        self._export_timer = None
        self._closed = False
        self.export_error = ""
        self.con = None
        if create:
            if os.path.exists(path) or os.path.exists(self.view_path):
                raise SchemaError("exists", "A file with this name already exists.")
            os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
            self._connect()
            self._ensure_schema()
        else:
            state = inspect_database(path)
            if state == "old":
                raise SchemaError("old", "This file is not in the current GrepMail format.")
            if state != "valid":
                raise SchemaError("unreadable", "The database file could not be read.")
            self._connect()
            self._ensure_schema()
        self.request_export()

    # ---------- low level ----------
    def _connect(self):
        self.con = sqlite3.connect(self.path, check_same_thread=False, timeout=15)

    def _columns(self, table):
        return [r[1] for r in self.con.execute('PRAGMA table_info("%s")' % table) if r[1] != "id"]

    def _ensure_schema(self):
        """Creates any missing table or column (never alters data)."""
        with self.lock:
            for name, headers in SCHEMA.items():
                want = DEFAULT_CUSTOM_FIELDS if name == SHEET_CUSTOM else headers
                have = self._columns(name)
                if not have:
                    cols = "".join(', "%s"' % h for h in want)
                    self.con.execute('CREATE TABLE IF NOT EXISTS "%s" (id INTEGER PRIMARY KEY '
                                     'AUTOINCREMENT%s)' % (name, cols))
                    continue
                if name == SHEET_CUSTOM:
                    continue  # its columns follow the user's Fields page
                for h in want:
                    if h not in have:
                        self.con.execute('ALTER TABLE "%s" ADD COLUMN "%s"' % (name, h))
            for table, col in ((SHEET_EMAIL_LOG, "EmailId"), (SHEET_OFFERS, "FileHash"),
                               (SHEET_OFFERS, "Reference"), (SHEET_TENDER_LOG, "EmailId")):
                if col in self._columns(table):
                    self.con.execute('CREATE INDEX IF NOT EXISTS "ix_%s_%s" ON "%s" ("%s")'
                                     % (table, col, table, col))
            self.con.commit()

    def _read(self, fn):
        with self.lock:
            return fn(None)

    def _modify(self, fn):
        with self.lock:
            try:
                result = fn(None)
            except Exception:
                if self._batch_depth == 0:
                    self.con.rollback()
                raise
            self._dirty = True
            if self._batch_depth == 0:
                self.con.commit()
                self._dirty = False
                self.request_export()
            return result

    def flush(self):
        with self.lock:
            if self.con is not None:
                self.con.commit()

    def begin_batch(self):
        with self.lock:
            self._batch_depth += 1

    def end_batch(self):
        with self.lock:
            self._batch_depth = max(0, self._batch_depth - 1)
            if self._batch_depth == 0:
                self.con.commit()
                if self._dirty:
                    self._dirty = False
                    self.request_export()

    def _rows(self, wb, sheet):
        cur = self.con.execute('SELECT * FROM "%s" ORDER BY id' % sheet)
        names = [d[0] for d in cur.description][1:]
        out = []
        for vals in cur.fetchall():
            vals = vals[1:]
            if all(v in (None, "") for v in vals):
                continue
            out.append(dict(zip(names, vals)))
        return out

    def _append(self, wb, sheet, rowdict):
        cols = self._columns(sheet)
        items = [(k, _sql_value(v)) for k, v in rowdict.items() if k in cols]
        if not items:
            return
        self.con.execute('INSERT INTO "%s" (%s) VALUES (%s)' % (
            sheet, ", ".join('"%s"' % k for k, _ in items), ", ".join("?" * len(items))),
            [v for _, v in items])

    def _delete_where(self, wb, sheet, header, predicate):
        if header not in self._columns(sheet):
            return 0
        ids = [r[0] for r in self.con.execute('SELECT id, "%s" FROM "%s"' % (header, sheet))
               if predicate(s(r[1]))]
        for i in ids:
            self.con.execute('DELETE FROM "%s" WHERE id = ?' % sheet, (i,))
        return len(ids)

    # ---------- Excel view ----------
    def request_export(self, delay=None):
        """Schedules a refresh of the Excel view file (several changes share one refresh)."""
        with self.lock:
            if self._closed or self._export_timer is not None:
                return
            t = threading.Timer(self.EXPORT_DELAY if delay is None else delay, self._export_job)
            t.daemon = True
            self._export_timer = t
            t.start()

    def _export_job(self):
        with self.lock:
            self._export_timer = None
            if self._closed:
                return
        try:
            self.export_view()
            self.export_error = ""
        except PermissionError:
            self.export_error = "The Excel view is open. It will refresh when it is closed."
            self.request_export(self.RETRY_DELAY)
        except Exception as e:
            self.export_error = str(e)

    def export_view(self):
        """Writes the Excel view file from the database. Raises PermissionError when the
        view is open in Excel."""
        with self.lock:
            data = {}
            for name in SCHEMA:
                cur = self.con.execute('SELECT * FROM "%s" ORDER BY id' % name)
                data[name] = ([d[0] for d in cur.description][1:], [r[1:] for r in cur.fetchall()])
        wb = Workbook()
        wb.remove(wb.active)
        for name, (heads, rows) in data.items():
            ws = wb.create_sheet(name)
            ws.append(heads)
            for c in ws[1]:
                c.font = Font(bold=True)
            for r in rows:
                ws.append(list(r))
            ws.freeze_panes = "A2"
            for col in ws.columns:
                width = max(len(s(c.value)) for c in list(col)[:200]) + 2
                ws.column_dimensions[col[0].column_letter].width = max(8, min(width, 60))
        tmp = self.view_path + ".tmp"
        try:
            wb.save(tmp)
            os.replace(tmp, self.view_path)
        finally:
            if os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except OSError:
                    pass

    def close(self):
        with self.lock:
            self._closed = True
            if self._export_timer is not None:
                self._export_timer.cancel()
                self._export_timer = None
            try:
                self.con.commit()
            except Exception:
                pass
        try:
            self.export_view()
        except Exception:
            pass
        with self.lock:
            try:
                self.con.close()
            except Exception:
                pass


    # ---------- tender references ----------
    def list_tenders(self):
        def f(wb):
            return [(s(r.get("Reference")), s(r.get("LocalPath")),
                     s(r.get("MailFolder")), date_str(r.get("Deadline")))
                    for r in self._rows(wb, SHEET_TENDERS) if s(r.get("Reference"))]
        return self._read(f)

    def tender_map(self):
        return {ref.upper(): path for ref, path, _, _ in self.list_tenders()}

    def tender_scopes(self):
        """{UPPER_REF: scope} - which mail folders each reference reads."""
        def f(wb):
            return {s(r.get("Reference")).upper(): (s(r.get("Scope")) if s(r.get("Scope")) in SCOPES
                                                    else SCOPE_SUB)
                    for r in self._rows(wb, SHEET_TENDERS) if s(r.get("Reference"))}
        return self._read(f)

    def add_or_update_tender(self, ref, local_path, mail_folder, deadline="", scope=SCOPE_SUB):
        ref, local_path, deadline = ref.strip(), local_path.strip(), deadline.strip()
        mail_folder = mail_folder.strip().strip("\\")
        if scope == SCOPE_INBOX:
            mail_folder = INBOX_TOKEN
        if not ref:
            raise ValueError("Reference cannot be empty")
        if scope not in SCOPES:
            raise ValueError("Choose where to search")
        if not mail_folder:
            raise ValueError("Choose a mail folder for this reference")
        if deadline and parse_date(deadline) is None:
            raise ValueError("Deadline must be a date like 2026-12-31")
        deadline = parse_date(deadline).isoformat() if deadline else ""

        def f(wb):
            for rid, cur in self.con.execute('SELECT id, "Reference" FROM "%s"' % SHEET_TENDERS).fetchall():
                if s(cur).upper() == ref.upper():
                    self.con.execute(
                        'UPDATE "%s" SET "LocalPath"=?, "MailFolder"=?, "Scope"=?, "Deadline"=? WHERE id=?'
                        % SHEET_TENDERS, (local_path, mail_folder, scope, deadline, rid))
                    return
            self._append(wb, SHEET_TENDERS,
                         {"Reference": ref, "LocalPath": local_path, "MailFolder": mail_folder,
                          "Scope": scope, "Deadline": deadline})
        self._modify(f)

    def delete_tender(self, ref):
        self._modify(lambda wb: self._delete_where(
            wb, SHEET_TENDERS, "Reference", lambda v: v.upper() == ref.upper()))

    # ---------- exclusion words ----------
    def list_exclusions(self):
        return self._read(lambda wb: [s(r.get("Word")) for r in
                                      self._rows(wb, SHEET_EXCLUSIONS) if s(r.get("Word"))])

    def exclusion_list_lower(self):
        return [w.lower() for w in self.list_exclusions()]

    def add_exclusion(self, word):
        word = word.strip()
        if not word:
            raise ValueError("The word cannot be empty")
        if word.lower() in self.exclusion_list_lower():
            return
        self._modify(lambda wb: self._append(wb, SHEET_EXCLUSIONS, {"Word": word}))

    def delete_exclusion(self, word):
        self._modify(lambda wb: self._delete_where(
            wb, SHEET_EXCLUSIONS, "Word", lambda v: v.lower() == word.lower()))

    # ---------- custom log ----------
    def set_custom_fields(self, fields):
        """Adds any newly chosen field as a column at the right of the Custom Log.
        Columns of fields you deselect stay (old rows keep their data)."""
        def f(wb):
            have = self._columns(SHEET_CUSTOM)
            for key in fields:
                if key not in have:
                    self.con.execute('ALTER TABLE "%s" ADD COLUMN "%s"' % (SHEET_CUSTOM, key))
        self._modify(f)

    def append_custom_log(self, row):
        self._modify(lambda wb: self._append(wb, SHEET_CUSTOM, row))

    def list_custom_log(self, limit=500):
        return self._read(lambda wb: self._rows(wb, SHEET_CUSTOM))[-limit:]

    # ---------- keywords ----------
    def list_keywords(self):
        return self._read(lambda wb: [s(r.get("Keyword")) for r in
                                      self._rows(wb, SHEET_KEYWORDS) if s(r.get("Keyword"))])

    def keyword_list_lower(self):
        return [k.lower() for k in self.list_keywords()]

    def add_keyword(self, kw):
        kw = kw.strip()
        if not kw:
            raise ValueError("Keyword cannot be empty")
        if kw.lower() in self.keyword_list_lower():
            return
        self._modify(lambda wb: self._append(wb, SHEET_KEYWORDS, {"Keyword": kw}))

    def delete_keyword(self, kw):
        self._modify(lambda wb: self._delete_where(
            wb, SHEET_KEYWORDS, "Keyword", lambda v: v.lower() == kw.lower()))

    # ---------- email log ----------
    def append_email_log(self, row):
        self._modify(lambda wb: self._append(wb, SHEET_EMAIL_LOG, row))

    def _tail(self, sheet, limit):
        """The newest `limit` rows (oldest first), read with SQL instead of loading everything."""
        def f(wb):
            cur = self.con.execute('SELECT * FROM "%s" ORDER BY id DESC LIMIT ?' % sheet, (limit,))
            names = [d[0] for d in cur.description][1:]
            rows = [dict(zip(names, r[1:])) for r in cur.fetchall()]
            rows.reverse()
            return rows
        return self._read(f)

    def list_email_log(self, limit=500):
        return self._tail(SHEET_EMAIL_LOG, limit)

    def email_entry_ids(self):
        """EntryIDs plus a time/subject key, because an email's EntryID changes
        when Outlook moves it to another folder."""
        def f(wb):
            out = set()
            for eid, direction, when, subject in self.con.execute(
                    'SELECT "EmailId", "Direction", "SentOrReceived", "Subject" FROM "%s"' % SHEET_EMAIL_LOG):
                if s(eid):
                    out.add(s(eid))
                out.add(email_key(direction, when, subject))
            return out
        return self._read(f)

    # ---------- tender log ----------
    def append_tender_log(self, row):
        self._modify(lambda wb: self._append(wb, SHEET_TENDER_LOG, row))

    def list_tender_log(self, limit=500):
        return self._tail(SHEET_TENDER_LOG, limit)

    def tender_log_keys(self, actions):
        """Set of (Action, EntryID) already logged - keeps rescans idempotent."""
        return self._read(lambda wb: {(s(a), s(e)) for a, e in self.con.execute(
            'SELECT "Action", "EmailId" FROM "%s"' % SHEET_TENDER_LOG)
            if s(a) in actions and s(e)})

    # ---------- offer register ----------
    def append_offer(self, row):
        self._modify(lambda wb: self._append(wb, SHEET_OFFERS, row))

    def list_offers(self, limit=None):
        if limit:
            return self._tail(SHEET_OFFERS, limit)
        return self._read(lambda wb: self._rows(wb, SHEET_OFFERS))

    def offer_keys(self):
        def f(wb):
            out = set()
            for eid, name, received in self.con.execute(
                    'SELECT "EmailId", "OriginalName", "ReceivedDate" FROM "%s"' % SHEET_OFFERS):
                out.add((s(eid), s(name)))
                out.add((s(received), s(name)))
            return out
        return self._read(f)

    def next_offer_no(self, tender_ref, supplier):
        def f(wb):
            ids = {s(e) or s(p_) for e, p_, r, n in self.con.execute(
                'SELECT "EmailId", "FilePath", "Reference", "Name" FROM "%s"' % SHEET_OFFERS)
                if s(r).upper() == tender_ref.upper() and s(n).lower() == supplier.lower()}
            return len(ids) + 1
        return self._read(f)

    def hash_exists(self, tender_ref, file_hash):
        return self._read(lambda wb: any(
            s(r).upper() == tender_ref.upper() for (r,) in self.con.execute(
                'SELECT "Reference" FROM "%s" WHERE "FileHash" = ?' % SHEET_OFFERS, (file_hash,))))

    # ---------- unmatched ----------
    def append_unmatched(self, row):
        self._modify(lambda wb: self._append(wb, SHEET_UNMATCHED, row))

    def list_unmatched(self):
        return self._read(lambda wb: self._rows(wb, SHEET_UNMATCHED))

    def unmatched_entry_ids(self):
        return {s(r.get("EmailId")) for r in self.list_unmatched() if s(r.get("EmailId"))}

    def remove_unmatched(self, entry_ids):
        ids = set(entry_ids)
        self._modify(lambda wb: self._delete_where(
            wb, SHEET_UNMATCHED, "EmailId", lambda v: v in ids))

    # ---------- summary ----------
    def tender_summary(self, warn_days=DEFAULT_WARN_DAYS):
        tenders, files = self._read(lambda wb: (self._rows(wb, SHEET_TENDERS),
                                                self._rows(wb, SHEET_OFFERS)))
        by_ref = {}
        for o in files:
            b = by_ref.setdefault(s(o.get("Reference")).upper(),
                                  {"files": 0, "names": set(), "last": None})
            b["files"] += 1
            if s(o.get("Name")):
                b["names"].add(s(o.get("Name")).lower())
            ts = parse_ts(o.get("ReceivedDate"))
            if ts and (b["last"] is None or ts > b["last"]):
                b["last"] = ts
        rows = []
        for t_ in tenders:
            ref = s(t_.get("Reference"))
            if not ref:
                continue
            deadline = date_str(t_.get("Deadline"))
            left = days_left(deadline)
            b = by_ref.get(ref.upper())
            if left is None:
                status = "Open"
            elif left < 0:
                status = "Past Deadline"
            elif left <= warn_days:
                status = "Due Soon"
            else:
                status = "Open"
            rows.append({
                "Reference": ref, "LocalPath": s(t_.get("LocalPath")),
                "Deadline": deadline, "DaysLeft": "" if left is None else left,
                "Files": b["files"] if b else 0,
                "Names": len(b["names"]) if b else 0,
                "LastSaved": b["last"].strftime("%Y-%m-%d %H:%M") if b and b["last"] else "",
                "Status": status,
            })
        return rows

    def refresh_summary_sheet(self, warn_days=DEFAULT_WARN_DAYS):
        rows = self.tender_summary(warn_days)

        def f(wb):
            self.con.execute('DELETE FROM "%s"' % SHEET_SUMMARY)
            for r in rows:
                self._append(wb, SHEET_SUMMARY, r)
        self._modify(f)

    # ---------- weekly report ----------
    def build_weekly_report(self, path, days=7, warn_days=DEFAULT_WARN_DAYS):
        cutoff = dt.datetime.now() - dt.timedelta(days=days)
        offers, unmatched, tlog = self._read(lambda wb: (
            self._rows(wb, SHEET_OFFERS), self._rows(wb, SHEET_UNMATCHED),
            self._rows(wb, SHEET_TENDER_LOG)))
        summary = self.tender_summary(warn_days)

        def recent(rows, key):
            out = []
            for r in rows:
                ts = parse_ts(r.get(key))
                if ts and ts >= cutoff:
                    out.append(r)
            return out

        week_offers = recent(offers, "ReceivedDate")
        week_unmatched = recent(unmatched, "Timestamp")
        week_failures = [r for r in recent(tlog, "Timestamp")
                         if s(r.get("Action")) in ("Save Failed", "No Allowed Attachments",
                                                   "No Local Path", "Threat Found", "Quarantined",
                                                   "Blocked")]
        active = {s(r.get("Reference")).upper() for r in week_offers}
        quiet = [r for r in summary
                 if r["Status"] != "Past Deadline" and r["Reference"].upper() not in active]
        upcoming = [r for r in summary if r["DaysLeft"] != "" and 0 <= r["DaysLeft"] <= 14]

        wb = Workbook()
        ws = wb.active
        ws.title = "Overview"
        ws.append(["Weekly report"])
        ws.append(["Generated", now_str()])
        ws.append(["Period", "%s to %s" % (cutoff.strftime("%Y-%m-%d"),
                                           dt.datetime.now().strftime("%Y-%m-%d"))])
        ws.append([])
        ws.append(["Files saved", len(week_offers)])
        ws.append(["References with files saved this period", len(active)])
        ws.append(["References with no files saved this period", len(quiet)])
        ws.append(["Unmatched emails", len(week_unmatched)])
        ws.append(["Save problems / emails needing review", len(week_failures)])
        ws.append(["Deadlines in next 14 days", len(upcoming)])
        ws["A1"].font = Font(bold=True, size=14)

        def sheet(title, headers, rows):
            w = wb.create_sheet(title)
            w.append(headers)
            for c in w[1]:
                c.font = Font(bold=True)
            for r in rows:
                w.append([r.get(h, "") if isinstance(r, dict) else "" for h in headers])
            for col in w.columns:
                width = max(len(s(c.value)) for c in col) + 2
                w.column_dimensions[col[0].column_letter].width = min(width, 60)

        sheet("Saved Files",
              ["ReceivedDate", "Reference", "Name", "FileNo", "Revision", "Security",
               "SavedFileName", "Notes"], week_offers)
        sheet("No Files This Period",
              ["Reference", "Deadline", "DaysLeft", "Files", "LastSaved", "Status"], quiet)
        sheet("Upcoming Deadlines",
              ["Reference", "Deadline", "DaysLeft", "Files", "Names", "Status"], upcoming)
        sheet("Unmatched Emails",
              ["Timestamp", "SenderEmail", "Subject", "DetectedRef"], week_unmatched)
        sheet("Needs Attention",
              ["Timestamp", "Action", "Reference", "Subject", "Details"], week_failures)
        ws.column_dimensions["A"].width = 40
        ws.column_dimensions["B"].width = 24
        wb.save(path)
        return {"offers": len(week_offers), "quiet": len(quiet),
                "unmatched": len(week_unmatched), "attention": len(week_failures)}


ExcelDB = GrepDB     # earlier name, kept so existing references keep working


# =========================================================
# EMAIL PROCESSING (shared by live monitor and retroactive scan)
# =========================================================
class Context:
    """Snapshot of tenders/keywords taken once per cycle."""

    def __init__(self, db):
        self.tender_map = db.tender_map()
        self.keywords = db.keyword_list_lower()
        self.exclusions = db.exclusion_list_lower()


ONCE_ACTIONS = ("Not Relevant", "No Allowed Attachments", "No Local Path", "Excluded",
                "Quarantined", "Threat Found", "Blocked")


def new_known(db=None):
    if db is None:
        return {"email_ids": set(), "offer_keys": set(),
                "unmatched_ids": set(), "tlog_seen": set()}
    return {"email_ids": db.email_entry_ids(), "offer_keys": db.offer_keys(),
            "unmatched_ids": db.unmatched_entry_ids(),
            "tlog_seen": db.tender_log_keys(ONCE_ACTIONS)}


class EmailProcessor:
    def __init__(self, db, cfg, emit):
        self.db = db
        self.cfg = cfg
        self.emit = emit  # emit(kind, *args) - thread-safe hand-off to the GUI
        self.dirty_summary = False
        self.alerted = set()
        self._reported = set()
        self._scanner = None
        self.quarantine_dir = os.path.join(os.path.dirname(os.path.abspath(db.path)), "GrepMail Quarantine")
        self.reset_stats()

    def reset_stats(self):
        self.stats = {"offers": 0, "files": 0, "unmatched": 0, "problems": 0, "excluded": 0, "threats": 0}

    def log(self, msg):
        self.emit("log", msg)

    def toast(self, title, msg):
        self.emit("toast", title, msg)

    def refresh(self, *tabs):
        self.emit("refresh", tabs)

    def report_once(self, key, msg):
        if key not in self._reported:
            self._reported.add(key)
            self.log(msg)

    def _tlog(self, action, ref, subject, path, details, entry_id):
        self.db.append_tender_log({
            "Timestamp": now_str(), "Action": action, "Reference": ref,
            "Subject": subject, "LocalPath": path, "Details": details,
            "EmailId": entry_id})
        self.refresh("tenderlog")

    def _tlog_once(self, known, action, ref, subject, path, details, entry_id, problem=True):
        """Log a result row only once per email."""
        if entry_id and (action, entry_id) in known["tlog_seen"]:
            return False
        self._tlog(action, ref, subject, path, details, entry_id)
        if entry_id:
            known["tlog_seen"].add((action, entry_id))
        if problem:
            self.stats["problems"] += 1
        return True

    def scanner(self):
        if self._scanner is None:
            self._scanner = gm_security.get_scanner()
        return self._scanner

    def _write_email_rows(self, row):
        """Email Log (everything) and Custom Log (your fields and filters)."""
        self.db.append_email_log(row)
        f = self.cfg.custom_filter
        if f.get("direction", "Both") not in ("Both", row["Direction"]):
            return
        if f.get("only_reference") and not s(row.get("Reference")):
            return
        if f.get("only_attachments") and not row.get("AttachmentCount"):
            return
        if s(f.get("contains")) and s(f.get("contains")).lower() not in s(row.get("Subject")).lower():
            return
        self.db.append_custom_log({k: row.get(k, "") for k in self.cfg.custom_fields})

    def process(self, item, direction, ctx, known, retro=False,
                tender_refs=None, root_only=False, folder_name=""):
        """tender_refs: only handle emails of these references (folder-scoped scan).
        root_only: Inbox scan for unmatched mail - ignore ordinary mail, keep tender-looking."""
        is_inbox = direction == "INCOMING"
        label = direction.title()
        subject = s(safe(lambda: item.Subject))
        entry_id = s(safe(lambda: item.EntryID))
        when = safe(lambda: item.ReceivedTime if is_inbox else item.SentOn, None)
        when_str = when.strftime("%Y-%m-%d %H:%M:%S") if when else ""

        tender_ref = find_tender_reference(subject, ctx.tender_map) if is_inbox else ""
        if root_only and not tender_ref and looks_like_tender(
                subject, self.cfg.hint_regex, self.cfg.hint_phrases) is None:
            return False  # ordinary mail sitting in the Inbox itself
        if tender_ref and tender_refs is not None and tender_ref not in tender_refs:
            return False  # belongs to another reference's folder scan

        sender_email = get_sender_smtp(item)
        body = s(safe(lambda: item.Body))
        names = attachment_names(item)
        local_path = ctx.tender_map.get(tender_ref, "") if tender_ref else ""
        ekey = email_key(direction, when_str, subject)
        already = (entry_id and entry_id in known["email_ids"]) or ekey in known["email_ids"]

        def finish(result):
            """Record the email (once) with its final result."""
            if not already:
                self._write_email_rows({
                    "LoggedAt": now_str(), "Direction": label, "EmailDate": when_str,
                    "SenderName": s(safe(lambda: item.SenderName)), "SenderEmail": sender_email,
                    "To": s(safe(lambda: item.To)), "Cc": s(safe(lambda: item.CC)),
                    "Subject": subject, "MailFolder": folder_name, "Reference": tender_ref,
                    "Name": extract_name(subject, self.cfg.name_separator) if tender_ref else "",
                    "AttachmentCount": len(names), "AttachmentNames": "; ".join(names)[:250],
                    "NewText": new_message_text(body).strip()[:500], "Result": result,
                    "LocalPath": local_path, "EmailId": entry_id})
                if entry_id:
                    known["email_ids"].add(entry_id)
                known["email_ids"].add(ekey)
                self.log("[%s] %s" % (label, subject[:70]))
                self.refresh("emaillog")
            return True

        if not is_inbox:
            return finish("Sent")

        # 1) a reference in the subject?
        if not tender_ref:
            hint = looks_like_tender(subject, self.cfg.hint_regex, self.cfg.hint_phrases)
            if hint is not None and (not entry_id or entry_id not in known["unmatched_ids"]):
                self.db.append_unmatched({
                    "Timestamp": now_str(), "SenderName": s(safe(lambda: item.SenderName)),
                    "SenderEmail": sender_email, "Subject": subject,
                    "DetectedRef": hint, "EmailId": entry_id})
                if entry_id:
                    known["unmatched_ids"].add(entry_id)
                self.stats["unmatched"] += 1
                self.refresh("unmatched")
                self.log("UNMATCHED email: %s" % subject[:70])
                if not retro:
                    self.toast("Unmatched email", "No reference found for:\n" + subject[:90])
            return finish("Unmatched" if hint is not None else "No Reference")

        if entry_id and entry_id in known["unmatched_ids"]:
            self.db.remove_unmatched([entry_id])
            known["unmatched_ids"].discard(entry_id)
            self.refresh("unmatched")

        if not local_path:
            self._tlog_once(known, "No Local Path", tender_ref, subject, "",
                            "The reference has no Local Path set", entry_id)
            return finish("No Local Path")

        # 2) an exclusion word in the subject?
        word = matched_exclusion(subject, ctx.exclusions)
        if word:
            if self._tlog_once(known, "Excluded", tender_ref, subject, local_path,
                               "Subject contains the exclusion word '%s'" % word, entry_id,
                               problem=False):
                self.stats["excluded"] += 1
                self.log("Excluded ('%s'): %s" % (word, subject[:55]))
            return finish("Excluded")

        # 3) a body keyword in the sender's new text?
        keyword = matched_keyword(new_message_text(body), ctx.keywords)
        if not keyword:
            allowed = [n for n in names if is_allowed_file_type(get_extension(n))]
            if allowed:
                detail = ("Has attachment(s) but no keyword in the new message text: "
                          + ", ".join(allowed))[:170]
                older = matched_keyword(body, ctx.keywords)
                if older:
                    detail += " | '%s' appears only in the earlier thread" % older
                if self._tlog_once(known, "Not Relevant", tender_ref, subject, local_path,
                                   detail, entry_id, problem=False):
                    self.log("Not relevant (attachment, no keyword in new text): %s" % subject[:55])
            return finish("Not Relevant")

        # 4) scan and save the attachments
        result = self._save_attachments(item, subject, body, tender_ref, local_path,
                                        entry_id, sender_email, when_str, known, retro, keyword)
        return finish(result)

    def _stage_dir(self):
        d = os.path.join(gm_platform.config_dir(), "staging")
        os.makedirs(d, exist_ok=True)
        return d

    def _save_attachments(self, item, subject, body, tender_ref, local_path,
                          entry_id, sender_email, when_str, known, retro, keyword=""):
        try:
            os.makedirs(local_path, exist_ok=True)
        except Exception as e:
            self._tlog("Save Failed", tender_ref, subject, local_path,
                       "Folder error: %s" % e, entry_id)
            self.stats["problems"] += 1
            self.log("SAVE FAILED (folder): %s" % subject[:60])
            self.toast("Save failed", "Cannot use folder:\n%s" % local_path)
            return "Save Failed"

        name = extract_name(subject, self.cfg.name_separator)
        file_no = None
        saved, skipped_type, dup_count, allowed_seen = [], [], 0, 0
        threats, held, blocked, notes = [], [], [], []
        scanner = self.scanner() if self.cfg.scan_enabled else None
        scanner_ok = bool(scanner and scanner.available())
        if self.cfg.scan_enabled and not scanner_ok:
            self.report_once("noscanner", "WARNING: no virus scanner available (%s). %s" % (
                scanner.name if scanner else "none",
                "Files are held in quarantine until it is installed." if self.cfg.require_scan
                else "Files are saved without a virus scan."))
        try:
            atts = item.Attachments
            for i in range(1, atts.Count + 1):
                att = atts.Item(i)
                fname = att.FileName
                if not is_allowed_file_type(get_extension(fname)):
                    skipped_type.append(fname)
                    continue
                allowed_seen += 1
                if (entry_id, fname) in known["offer_keys"] or \
                        (when_str, fname) in known["offer_keys"]:
                    continue  # already saved in an earlier run
                stage = os.path.join(self._stage_dir(), "%s_%s" % (os.getpid(), clean_filename(fname)))
                att.SaveAsFile(stage)
                verdict = gm_security.evaluate_attachment(
                    stage, scanner if scanner_ok else None, self.cfg.macro_mode,
                    self.cfg.scan_enabled)
                action = verdict["action"]
                if action == "save" and self.cfg.scan_enabled and not scanner_ok and self.cfg.require_scan:
                    verdict.update(action="quarantine", security="Not Scanned",
                                   detail="No virus scanner available")
                    action = "quarantine"
                if action == "threat":
                    threats.append("%s (%s)" % (fname, verdict["detail"]))
                    try:
                        if os.path.exists(stage):
                            os.remove(stage)
                    except Exception:
                        pass
                    continue
                if action == "quarantine":
                    gm_security.quarantine(stage, self.quarantine_dir,
                                           "%s | %s | %s" % (subject, verdict["security"], verdict["detail"]))
                    held.append("%s (%s)" % (fname, verdict["detail"] or verdict["security"]))
                    continue
                if action == "blocked":
                    try:
                        os.remove(stage)
                    except Exception:
                        pass
                    blocked.append(fname)
                    continue
                if file_no is None:
                    file_no = self.db.next_offer_no(tender_ref, name)
                new_name = build_file_name(self.cfg.file_format, name, tender_ref,
                                           when_str[:10], fname)
                final_path = unique_file_path(local_path, new_name)
                shutil.move(stage, final_path)
                if verdict["macro"] == gm_security.MACRO_PRESENT and self.cfg.mark_web:
                    gm_security.mark_from_internet(final_path)
                if verdict["macro"] == gm_security.MACRO_PRESENT:
                    notes.append(os.path.basename(final_path))
                fhash = file_sha256(final_path)
                is_dup = self.db.hash_exists(tender_ref, fhash)
                dup_count += 1 if is_dup else 0
                self.db.append_offer({
                    "ReceivedDate": when_str, "Reference": tender_ref,
                    "Name": name, "FileNo": file_no,
                    "Revision": detect_revision(fname, subject, body),
                    "SavedFileName": os.path.basename(final_path),
                    "OriginalName": fname, "FilePath": final_path,
                    "SenderEmail": sender_email, "FileHash": fhash,
                    "Security": verdict["security"],
                    "Notes": "Duplicate of an earlier file" if is_dup else "",
                    "EmailId": entry_id, "MatchedKeyword": keyword})
                known["offer_keys"].add((entry_id, fname))
                known["offer_keys"].add((when_str, fname))
                saved.append(os.path.basename(final_path))
        except Exception as e:
            self._tlog("Save Failed", tender_ref, subject, local_path, str(e), entry_id)
            self.stats["problems"] += 1
            self.log("SAVE FAILED: %s - %s" % (subject[:50], e))
            self.toast("Save failed", "%s\n%s" % (subject[:70], e))
            return "Save Failed"

        if threats:
            if self._tlog_once(known, "Threat Found", tender_ref, subject, local_path,
                               ("Not saved: " + "; ".join(threats))[:300], entry_id):
                self.stats["threats"] += 1
                self.log("THREAT FOUND - not saved: %s" % "; ".join(threats)[:120])
                self.toast("Threat found", "%s\nNot saved: %s" % (subject[:60], "; ".join(threats)[:120]))
        if held:
            if self._tlog_once(known, "Quarantined", tender_ref, subject, self.quarantine_dir,
                               ("Held in quarantine: " + "; ".join(held))[:300], entry_id):
                self.log("QUARANTINED: %s" % "; ".join(held)[:120])
                self.toast("File quarantined", "%s\n%s" % (subject[:60], "; ".join(held)[:120]))
        if blocked:
            if self._tlog_once(known, "Blocked", tender_ref, subject, local_path,
                               "Macro-enabled file(s) blocked: " + ", ".join(blocked), entry_id):
                self.log("Blocked macro-enabled file(s): %s" % ", ".join(blocked)[:100])

        if saved:
            details = "keyword '%s' | %d file(s): %s" % (keyword, len(saved), ", ".join(saved))
            if skipped_type:
                details += " | skipped: " + ", ".join(skipped_type)
            if dup_count:
                details += " | %d duplicate(s)" % dup_count
            if notes:
                details += " | macro-enabled: " + ", ".join(notes)
            self._tlog("Saved", tender_ref, subject, local_path, details[:300], entry_id)
            self.stats["offers"] += 1
            self.stats["files"] += len(saved)
            self.dirty_summary = True
            self.log("SAVED (%d file(s)): %s" % (len(saved), subject[:55]))
            self.refresh("offers", "summary")
            if not retro:
                note = "%s\n%d file(s) saved" % (tender_ref, len(saved))
                if dup_count:
                    note += " (%d duplicate)" % dup_count
                if notes:
                    note += "\nMacro-enabled file saved - do not enable macros unless you trust it"
                self.toast("Saved - " + (name or "email"), note)
        elif skipped_type and allowed_seen == 0:
            if self._tlog_once(known, "No Allowed Attachments", tender_ref, subject,
                               local_path, "Skipped: " + ", ".join(skipped_type), entry_id):
                self.log("No allowed attachment types: %s" % subject[:60])

        if threats:
            return "Threat Found"
        if held:
            return "Quarantined"
        if saved:
            return "Saved"
        if blocked:
            return "Blocked"
        if allowed_seen == 0 and skipped_type:
            return "No Allowed Attachments"
        if allowed_seen == 0:
            return "No Attachments"
        return "Already Saved"

    # ---------- deadlines & folders ----------
    def check_deadlines(self):
        today = dt.date.today()
        counts = None
        for ref, _path, _outlook, deadline in self.db.list_tenders():
            left = days_left(deadline)
            if left is None or not (0 <= left <= self.cfg.warn_days):
                continue
            key = (ref.upper(), today.isoformat())
            if key in self.alerted:
                continue
            self.alerted.add(key)
            if counts is None:
                counts = {r["Reference"].upper(): r["Files"]
                          for r in self.db.tender_summary(self.cfg.warn_days)}
            n = counts.get(ref.upper(), 0)
            when = "today" if left == 0 else "in %d day(s)" % left
            msg = "%s closes %s (%s)\n%d file(s) saved so far" % (ref, when, deadline, n)
            self.log("DEADLINE: " + msg.replace("\n", " - "))
            self.toast("Deadline approaching", msg)

    def check_folders_report(self):
        status, problems = {}, []
        for ref, path, _outlook, _dl in self.db.list_tenders():
            st = check_folder(path)
            status[ref.upper()] = st
            if st != "OK":
                problems.append("%s: %s" % (ref, st))
        self.emit("folders", status)
        if problems:
            self.log("Folder check - %d problem(s): %s" % (len(problems), "; ".join(problems[:6])))
            self.toast("Folder check", "%d tender folder(s) need attention" % len(problems))
        else:
            self.log("Folder check - all tender folders OK")


# =========================================================
# MAIL CONNECTION AND MONITOR THREADS
# =========================================================
class MailConnection:
    """What the scanning code needs from a mailbox (Outlook or IMAP)."""

    def __init__(self, inbox, sent, root=None, closer=None, name=""):
        self.inbox, self.sent, self.root, self._closer, self.name = inbox, sent, root, closer, name

    def close(self):
        try:
            if self._closer:
                self._closer()
        except Exception:
            pass


def connect_outlook():
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    return MailConnection(ns.GetDefaultFolder(6), ns.GetDefaultFolder(5), None, None, "Outlook")


def imap_password(cfg):
    im = cfg.imap
    return gm_mail.load_password(im.get("user", ""), im.get("host", "")) or cfg.session_password


def connect_mail(cfg):
    """Connects with the provider chosen in Settings."""
    if cfg.provider == "outlook":
        if not HAS_OUTLOOK:
            raise RuntimeError("Outlook access needs Windows with pywin32 installed.")
        return connect_outlook()
    im = cfg.imap
    if not im.get("host") or not im.get("user"):
        raise RuntimeError("Enter your mail account in Settings first.")
    pw = imap_password(cfg)
    if not pw:
        raise RuntimeError("No password stored for this mail account. Enter it in Settings.")
    mb = gm_mail.connect_imap(im, pw)
    return MailConnection(mb.inbox(), mb.sent(), mb.root, mb.close, im.get("user", ""))


def provider_ready(cfg):
    """(True, '') or (False, why)."""
    if cfg.provider == "outlook" and not HAS_OUTLOOK:
        return False, ("Outlook access needs Windows with pywin32 installed.\n"
                       "Choose an IMAP mail account in Settings instead.")
    if cfg.provider == "imap" and not (cfg.imap.get("host") and cfg.imap.get("user")):
        return False, "Enter your mail account in Settings first."
    return True, ""


def _walk_subfolders(folder, depth=0):
    """Yields every subfolder below `folder` (depth-limited as a safety net)."""
    if depth > 6:
        return
    for sub in folder.Folders:
        yield sub
        for deeper in _walk_subfolders(sub, depth + 1):
            yield deeper


def _split_path(relative):
    return [seg for seg in re.split(r"[\\/]", relative) if seg]


def _folder_key(folder):
    return str(safe(lambda: folder.EntryID, id(folder)))


def resolve_tender_folders(conn, proc):
    """Maps each mail folder to scan onto the references that use it, following
    each reference's scope: Subfolders (default), Subfolders And Inbox, Inbox Only."""
    grouped = {}
    scopes = proc.db.tender_scopes()
    inbox = conn.inbox

    def add(folder, ref):
        grouped.setdefault(_folder_key(folder), (folder, set()))[1].add(ref)

    for ref, _path, relative, _dl in proc.db.list_tenders():
        if not relative:
            proc.report_once("missing:" + ref, "No mail folder configured for reference %s; "
                                               "update it in References." % ref)
            continue
        scope = scopes.get(ref.upper(), SCOPE_SUB)
        if scope == SCOPE_INBOX:
            add(inbox, ref.upper())
            continue
        try:
            if relative == INBOX_TOKEN:
                base, include_base = inbox, False
            else:
                base = conn.root if relative.startswith("/") else inbox
                if base is None:
                    raise KeyError("absolute folder paths need an IMAP account")
                for seg in _split_path(relative):
                    base = base.Folders.Item(seg)
                include_base = True
        except Exception as e:
            proc.report_once("unavailable:" + relative.casefold(),
                             "Mail folder '%s' is unavailable: %s" % (relative, e))
            continue
        try:
            if include_base:
                add(base, ref.upper())
            for sub in _walk_subfolders(base):
                add(sub, ref.upper())
        except Exception as e:
            proc.report_once("subfolders:" + relative.casefold(),
                             "Could not read subfolders of '%s': %s" % (relative, e))
        if scope == SCOPE_SUB_INBOX:
            add(inbox, ref.upper())
    return grouped


def scan_inbox_side(conn, since, proc, ctx, known, stop_event, retro=False):
    """Scans every folder that the references point to, and (optionally) the Inbox
    itself for tender-looking mail that matches no reference. Returns the newest time seen."""
    newest = since
    for folder, refs in resolve_tender_folders(conn, proc).values():
        try:
            n, _ = scan_folder(folder, since, "INCOMING", proc, ctx, known, stop_event,
                               retro, tender_refs=refs)
            newest = max(newest, n)
        except Exception as e:
            proc.log("Mail folder scan error: %s" % e)
    if proc.cfg.scan_root:
        try:
            n, _ = scan_folder(conn.inbox, since, "INCOMING", proc, ctx, known, stop_event,
                               retro, root_only=True)
            newest = max(newest, n)
        except Exception as e:
            proc.log("Inbox scan error: %s" % e)
    return newest


def scan_folder(folder, since, direction, proc, ctx, known, stop_event, retro=False,
                tender_refs=None, root_only=False):
    is_inbox = direction == "INCOMING"
    items = folder.Items
    items.Sort("[ReceivedTime]" if is_inbox else "[SentOn]", True)
    if hasattr(items, "set_since"):
        items.set_since(since)
    folder_name = s(safe(lambda: folder.FolderPath, "")).lstrip("\\")
    newest, count = since, 0
    pending = []
    for item in items:  # newest first: stop as soon as we reach old mail
        if stop_event.is_set():
            return newest, count
        try:
            if item.Class != 43:  # olMail
                continue
            when_dt = to_naive(item.ReceivedTime if is_inbox else item.SentOn)
            if when_dt <= since:
                break
            if when_dt > newest:
                newest = when_dt
            pending.append(item)
        except Exception as e:
            proc.log("Item error: %s" % e)
    # Process OLDEST first so file numbers / duplicate flags follow real order
    for item in reversed(pending):
        if stop_event.is_set():
            break
        try:
            if proc.process(item, direction, ctx, known, retro, tender_refs, root_only, folder_name):
                count += 1
            if retro and count and count % 25 == 0:
                proc.log("Rescan: %d %s emails processed..." % (count, direction.lower()))
                try:
                    proc.db.flush()
                except Exception:
                    pass
        except Exception as e:
            proc.log("Item error: %s" % e)
    return newest, count


class MonitorThread(threading.Thread):
    def __init__(self, db, cfg, emit, stop_event):
        super().__init__(daemon=True)
        self.db, self.cfg, self.stop_event = db, cfg, stop_event
        self.proc = EmailProcessor(db, cfg, emit)

    def run(self):
        proc = self.proc
        ok, why = provider_ready(self.cfg)
        if not ok:
            proc.log(why.replace("\n", " "))
            return
        if self.cfg.provider == "outlook":
            pythoncom.CoInitialize()
        conn = None
        last_in = self.cfg.get_last_check("inbox")
        last_sent = self.cfg.get_last_check("sent")
        try:
            known = new_known(self.db)  # what is already logged / saved
        except Exception as e:
            proc.log("Could not read existing history (%s) - continuing without it." % e)
            known = new_known()
        try:
            while not self.stop_event.is_set():
                try:
                    if conn is None:
                        proc.log("Connecting to %s..." % ("Outlook" if self.cfg.provider == "outlook"
                                                          else self.cfg.imap.get("host", "mail server")))
                        conn = connect_mail(self.cfg)
                        proc.log("Connected. Monitoring started.")
                        proc.check_folders_report()
                    ctx = Context(self.db)
                    self.db.begin_batch()
                    try:
                        new_in = scan_inbox_side(conn, last_in, proc, ctx, known, self.stop_event)
                        new_sent = last_sent
                        if conn.sent is not None:
                            new_sent, _ = scan_folder(conn.sent, last_sent, "OUTGOING", proc, ctx,
                                                      known, self.stop_event)
                        else:
                            proc.report_once("nosent", "No Sent folder was found on this account; "
                                                       "outgoing mail is not logged.")
                        proc.check_deadlines()
                        if proc.dirty_summary:
                            self.db.refresh_summary_sheet(self.cfg.warn_days)
                            proc.dirty_summary = False
                    finally:
                        try:
                            self.db.end_batch()
                        except Exception as e:
                            proc.log("Database save failed: %s" % e)
                    if new_in > last_in:
                        last_in = new_in
                        self.cfg.set_last_check("inbox", last_in)
                    if new_sent > last_sent:
                        last_sent = new_sent
                        self.cfg.set_last_check("sent", last_sent)
                except Exception as e:
                    proc.log("Monitor error: %s (will reconnect)" % e)
                    if conn is not None:
                        conn.close()
                    conn = None
                self.stop_event.wait(max(5, self.cfg.poll_seconds))
        finally:
            if conn is not None:
                conn.close()
            if self.cfg.provider == "outlook":
                pythoncom.CoUninitialize()
            proc.log("Monitoring stopped.")


class RescanThread(threading.Thread):
    def __init__(self, db, cfg, emit, since, stop_event):
        super().__init__(daemon=True)
        self.db, self.cfg, self.since, self.stop_event = db, cfg, since, stop_event
        self.emit = emit
        self.proc = EmailProcessor(db, cfg, emit)

    def run(self):
        proc = self.proc
        ok, why = provider_ready(self.cfg)
        if not ok:
            proc.log(why.replace("\n", " "))
            self.emit("scan_done")
            return
        if self.cfg.provider == "outlook":
            pythoncom.CoInitialize()
        conn = None
        try:
            proc.log("Rescan from %s - connecting..." % self.since.strftime("%Y-%m-%d"))
            conn = connect_mail(self.cfg)
            ctx = Context(self.db)
            known = new_known(self.db)
            proc.reset_stats()
            self.db.begin_batch()
            try:
                scan_inbox_side(conn, self.since, proc, ctx, known, self.stop_event, retro=True)
                if conn.sent is not None:
                    scan_folder(conn.sent, self.since, "OUTGOING", proc, ctx, known,
                                self.stop_event, retro=True)
                if proc.dirty_summary:
                    self.db.refresh_summary_sheet(self.cfg.warn_days)
            finally:
                try:
                    self.db.end_batch()
                except Exception as e:
                    proc.log("Database save failed: %s" % e)
            st = proc.stats
            state = "cancelled" if self.stop_event.is_set() else "complete"
            msg = ("Rescan %s: %d email(s) saved files, %d file(s) saved, %d unmatched, "
                   "%d excluded, %d threat(s), %d needing attention" % (
                       state, st["offers"], st["files"], st["unmatched"], st["excluded"],
                       st["threats"], st["problems"]))
            proc.log(msg)
            proc.toast("Rescan " + state, msg.split(": ", 1)[1])
            proc.refresh("emaillog", "tenderlog", "offers", "unmatched", "summary")
        except Exception as e:
            proc.log("Rescan error: %s" % e)
        finally:
            if conn is not None:
                conn.close()
            if self.cfg.provider == "outlook":
                pythoncom.CoUninitialize()
            self.emit("scan_done")


# =========================================================
# GUI
# =========================================================
FONT = ("Segoe UI", 10)
FONT_BOLD = ("Segoe UI", 10, "bold")
FONT_TITLE = ("Segoe UI", 14, "bold")
FONT_GROUP = ("Segoe UI", 9, "bold")
FONT_STAT = ("Segoe UI", 24, "bold")
FONT_MONO = ("Consolas", 10)

PALETTES = {
    "light": dict(
        bg="#F8F9FA", surface="#FFFFFF", primary="#2B579A", fg="#212529", muted="#6C757D",
        success="#198754", warning="#FD7E14", danger="#DC3545", alt="#F1F5F9", head="#E9ECEF",
        border="#E0E0E0", nav_active="#E7EEF9", nav_hover="#F1F5F9",
        tint_ok="#E3F4E9", tint_warn="#FFF0E0", tint_bad="#FBE3E6", tint_past="#EEF0F2",
        fg_ok="#146C43", fg_warn="#B25E09", fg_bad="#B02A37", fg_past="#6C757D",
        console_bg="#1E2124", console_fg="#D7DAE0", toast_bg="#1F2937",
        banner_bg="#FFF8DB", banner_fg="#664D03"),
    "dark": dict(
        bg="#121212", surface="#1E1E1E", primary="#4D88E5", fg="#E0E0E0", muted="#A0A0A0",
        success="#28A745", warning="#FFC107", danger="#EA868F", alt="#2A2D32", head="#2A2D32",
        border="#3A3D42", nav_active="#26354D", nav_hover="#2A2D32",
        tint_ok="#1E3A2A", tint_warn="#3F3317", tint_bad="#43232A", tint_past="#26282C",
        fg_ok="#7EE0A3", fg_warn="#FFC107", fg_bad="#EA868F", fg_past="#A0A0A0",
        console_bg="#0B0C0E", console_fg="#D7DAE0", toast_bg="#2B2F36",
        banner_bg="#3A3317", banner_fg="#FFE69C"),
}

TB_COLORS = {
    "light": dict(primary="#2B579A", secondary="#6C757D", success="#198754", info="#3A86C8",
                  warning="#FD7E14", danger="#DC3545", light="#E9ECEF", dark="#343A40",
                  bg="#F8F9FA", fg="#212529", selectbg="#2B579A", selectfg="#FFFFFF",
                  border="#CED4DA", inputfg="#212529", inputbg="#FFFFFF", active="#E9ECEF"),
    "dark": dict(primary="#4D88E5", secondary="#8A8F98", success="#28A745", info="#5BC0DE",
                 warning="#FFC107", danger="#DC3545", light="#2A2D32", dark="#1E1E1E",
                 bg="#121212", fg="#E0E0E0", selectbg="#4D88E5", selectfg="#FFFFFF",
                 border="#3A3D42", inputfg="#E0E0E0", inputbg="#1E1E1E", active="#2A2D32"),
}

PAGES = [  # (group, key, label)
    ("SETUP", "tenders", "References"),
    ("SETUP", "keywords", "Body Keywords"),
    ("SETUP", "rules", "Subject Rules"),
    ("SETUP", "settings", "Settings"),
    ("OVERVIEW", "summary", "Summary"),
    ("OVERVIEW", "offers", "Saved Files"),
    ("OVERVIEW", "unmatched", "Unmatched"),
    ("LOGS", "emaillog", "Email Log"),
    ("LOGS", "customlog", "Custom Log"),
    ("LOGS", "tenderlog", "Processing Log"),
    ("LOGS", "monitor", "Monitor Log"),
]


def keyword_language(text):
    arabic = any("\u0600" <= c <= "\u06FF" or "\u0750" <= c <= "\u077F" or
                 "\uFB50" <= c <= "\uFDFF" or "\uFE70" <= c <= "\uFEFF" for c in text)
    latin = any("a" <= c.lower() <= "z" for c in text)
    if arabic and latin:
        return "Mixed"
    if arabic:
        return "Arabic"
    return "English" if latin else "Other"


def classify_log(msg):
    """Gives each monitor message a level for the colored console."""
    low = msg.lower()
    if msg.startswith("[Incoming]") or msg.startswith("[Outgoing]") or low.startswith("saved ("):
        return "PROCESSED"
    if (low.startswith(("save failed", "threat found", "quarantined")) or "error" in low or "failed" in low
            or "could not" in low or "cannot" in low):
        return "ERROR"
    if (low.startswith(("unmatched", "not relevant", "excluded", "blocked", "warning", "deadline",
                        "no allowed", "no mail folder"))
            or "unavailable" in low or ("problem(s)" in low and " 0 problem" not in low)):
        return "WARN"
    if low.startswith("rescan"):
        return "SCAN"
    return "INFO"


class DateField(ttk.Frame):
    """Date entry with a calendar button (plain entry if the picker is missing)."""

    def __init__(self, parent, width=14):
        super().__init__(parent)
        self.widget = None
        if DateEntry is not None:
            try:
                self.widget = DateEntry(self, dateformat="%Y-%m-%d", width=width,
                                        bootstyle="primary")
                self.entry = self.widget.entry
            except Exception:
                self.widget = None
        if self.widget is None:
            self.entry = ttk.Entry(self, width=width)
            self.widget = self.entry
        self.widget.pack(fill="x")
        self.set("")

    def get(self):
        return self.entry.get().strip()

    def set(self, text):
        self.entry.delete(0, "end")
        if text:
            self.entry.insert(0, text)


class TableTab(ttk.Frame):
    """Filterable, zebra-striped table. Rows are dicts; `columns` picks what is shown."""

    def __init__(self, parent, columns, widths=None, refresh_cb=None, buttons=(),
                 tag_fn=None, headings=None, on_double=None, row_height=28):
        super().__init__(parent)
        self.columns, self.rows, self.iid_map, self.tag_fn = list(columns), [], {}, tag_fn
        top = ttk.Frame(self)
        top.pack(fill="x", pady=(0, 8))
        self._refresh_cb, self._stamp = refresh_cb, ""
        if refresh_cb:
            self._refresh_btn = ttk.Button(top, text="\u27f3 Refresh", command=self._click_refresh,
                                           bootstyle="secondary-outline")
            self._refresh_btn.pack(side="left", padx=(0, 6))
        for text, cb, style in buttons:
            ttk.Button(top, text=text, command=cb, bootstyle=style).pack(side="left", padx=(0, 6))
        self.count_label = ttk.Label(top, text="", bootstyle="secondary")
        self.count_label.pack(side="right")
        self.filter_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.filter_var, width=20).pack(side="right", padx=(0, 10))
        ttk.Label(top, text="\U0001F50D", bootstyle="secondary").pack(side="right", padx=(0, 4))
        self.filter_var.trace_add("write", lambda *a: self._render())

        body = ttk.Frame(self)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(body, columns=self.columns, show="headings",
                                 selectmode="extended", height=5)
        vs = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(body, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        for i, c in enumerate(self.columns):
            self.tree.heading(c, text=(headings or {}).get(c, c), anchor="w")
            self.tree.column(c, width=(widths[i] if widths else 120), anchor="w", minwidth=50)
        if on_double:
            self.tree.bind("<Double-1>", lambda e: on_double())

    def _click_refresh(self):
        """Gives visible feedback: the button dims while refreshing and the time is shown."""
        btn = self._refresh_btn
        btn.config(state="disabled", text="Refreshing...")

        def run():
            try:
                self._refresh_cb()
                self._stamp = "Updated " + time.strftime("%H:%M:%S")
                self._render()
            finally:
                self.after(500, lambda: btn.winfo_exists() and btn.config(state="normal",
                                                                          text="\u27f3 Refresh"))
        self.after(40, run)     # lets the window repaint the dimmed button first

    def apply_palette(self, p):
        for tag, bg, fg in (("even", p["surface"], p["fg"]), ("odd", p["alt"], p["fg"]),
                            ("ok", p["tint_ok"], p["fg_ok"]), ("warn", p["tint_warn"], p["fg_warn"]),
                            ("bad", p["tint_bad"], p["fg_bad"]), ("past", p["tint_past"], p["fg_past"])):
            self.tree.tag_configure(tag, background=bg, foreground=fg)

    def set_rows(self, rows):
        self.rows = rows
        self._render()

    def _render(self):
        flt = self.filter_var.get().strip().lower()
        self.tree.delete(*self.tree.get_children())
        self.iid_map.clear()
        shown = 0
        for r in self.rows:
            vals = ["" if r.get(c) is None else str(r.get(c)) for c in self.columns]
            if flt and flt not in " ".join(vals).lower():
                continue
            tags = tuple(self.tag_fn(r)) if self.tag_fn else ()
            if not tags:
                tags = ("odd" if shown % 2 else "even",)
            iid = self.tree.insert("", "end", values=vals, tags=tags[:1])
            self.iid_map[iid] = r
            shown += 1
        self.count_label.config(text=("%d rows" % shown) + (("   \u00b7   " + self._stamp) if self._stamp else ""))

    def selected(self):
        return [self.iid_map[i] for i in self.tree.selection() if i in self.iid_map]


class App(tb.Window):
    def __init__(self):
        self.cfg = Config()
        theme = self.cfg.theme if self.cfg.theme in PALETTES else "light"
        if sys.platform.startswith("win"):  # own taskbar identity instead of "python.exe"
            try:
                import ctypes
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("GrepMail.App")
            except Exception:
                pass
        super().__init__(title=APP_TITLE, themename="flatly", size=(1180, 740),
                         minsize=(980, 600), iconphoto=None)
        try:
            self._icon_big = tk.PhotoImage(data=ICON_128_B64, master=self)
            self._icon_small = tk.PhotoImage(data=ICON_32_B64, master=self)
            self.iconphoto(True, self._icon_big, self._icon_small)
        except Exception:
            self._icon_big = self._icon_small = None
        self._apply_windows_icon()
        self.after(300, self._apply_windows_icon)     # again once the window is shown
        self.tb_style = tb.Style()
        for name, colors in TB_COLORS.items():
            self.tb_style.register_theme(ThemeDefinition("grepmail_" + name, colors, name))
        try:
            self._rowh = scale_size(self, 28)
            self._headpad = scale_size(self, 6)
        except Exception:
            self._rowh, self._headpad = 28, 6

        self.pal = PALETTES[theme]
        self.db = None
        self.monitor = None
        self.rescan = None
        self.stop_event = threading.Event()
        self.rescan_stop = threading.Event()
        self.q = queue.Queue()
        self.toasts = []
        self.dirty = set()
        self.stale = set()
        self.last_refresh = 0.0
        self.folder_status = {}
        self._skinned, self._tables = [], []
        self._form_frames = set()
        self.pages, self.nav_items, self.nav_labels = {}, {}, {}
        self.current_page = None
        self._status = "idle"

        self._build_menu()
        self._build_layout()
        self._apply_theme(theme)
        self.show_page("tenders")
        self._poll_queue()
        self.after(250, self._startup_database)

    # ---------- skinning helpers ----------
    def skin(self, widget, role):
        self._skinned.append((widget, role))
        return widget

    def _apply_theme(self, name):
        self.theme_name = name
        self.pal = PALETTES[name]
        self.tb_style.theme_use("grepmail_" + name)
        st = self.tb_style
        st.configure("Treeview", rowheight=self._rowh, font=FONT)
        st.configure("Treeview.Heading", font=FONT_BOLD, padding=(8, self._headpad),
                     background=self.pal["head"], foreground=self.pal["fg"], relief="flat")
        self._restyle()

    def _restyle(self):
        p = self.pal
        roles = {
            "bg": dict(bg=p["bg"]), "surface": dict(bg=p["surface"]),
            "card": dict(bg=p["surface"], highlightbackground=p["border"], highlightcolor=p["border"]),
            "formcard": dict(bg=p["bg"], highlightbackground=p["border"], highlightcolor=p["border"]),
            "label": dict(bg=p["surface"], fg=p["fg"]), "muted": dict(bg=p["surface"], fg=p["muted"]),
            "bg_label": dict(bg=p["bg"], fg=p["fg"]), "bg_muted": dict(bg=p["bg"], fg=p["muted"]),
            "hint": dict(bg=TB_COLORS[self.theme_name]["inputbg"], fg=p["muted"]),
            "line": dict(bg=p["border"]), "banner": dict(bg=p["banner_bg"], fg=p["banner_fg"]),
            "stat_primary": dict(bg=p["surface"], fg=p["primary"]),
            "stat_warning": dict(bg=p["surface"], fg=p["warning"]),
            "stat_danger": dict(bg=p["surface"], fg=p["danger"]),
            "console": dict(bg=p["console_bg"], fg=p["console_fg"], insertbackground=p["console_fg"]),
        }
        for w, role in self._skinned:
            try:
                w.configure(**roles[role])
            except tk.TclError:
                pass
        for t in self._tables:
            t.apply_palette(p)
        self._highlight_nav()
        self._set_status(self._status)

    # ---------- menu / database ----------
    def _build_menu(self):
        menubar = tk.Menu(self)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="New Database...", command=self._create_database_flow)
        filemenu.add_command(label="Open Database...", command=self._open_existing_flow)
        filemenu.add_command(label="Export to Excel Now", command=self._export_excel_now)
        filemenu.add_command(label="Open Excel View", command=self._open_excel_view)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=self._on_close)
        menubar.add_cascade(label="File", menu=filemenu)
        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="Check for Updates...", command=lambda: self._check_updates(manual=True))
        helpmenu.add_command(label="About GrepMail", command=self._about)
        menubar.add_cascade(label="Help", menu=helpmenu)
        self.config(menu=menubar)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _apply_windows_icon(self):
        """The Windows taskbar uses the window's .ico, not the Tk photo icon, so an .ico built
        from the embedded logo is written to the settings folder and applied to every window."""
        if not sys.platform.startswith("win"):
            return
        try:
            import base64
            import io
            from PIL import Image
            ico = os.path.join(gm_platform.config_dir(), "grepmail_icon_%s.ico" % APP_VERSION)
            if not os.path.exists(ico):
                os.makedirs(os.path.dirname(ico), exist_ok=True)
                img = Image.open(io.BytesIO(base64.b64decode(ICON_128_B64))).convert("RGBA")
                img.save(ico, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128)])
            self.iconbitmap(default=ico)
        except Exception:
            pass

    # ----- database: first run, create, open, conflicts -----
    def _startup_database(self):
        path = self.cfg.db_path
        if path and not path.lower().endswith(".db"):
            path = os.path.splitext(path)[0] + ".db"     # a path saved by an earlier build
        if path:
            state = inspect_database(path)
            if state == "valid":
                self._open_db(path)
                return
            if state in ("old", "unreadable", "view"):
                if not self._resolve_conflict(path, state):
                    return
                return
        self._welcome_dialog()

    def _choice_dialog(self, title, message, buttons):
        """Modal dialog. buttons = [(key, text, bootstyle)]. Returns the key (None if closed)."""
        win = tb.Toplevel(self)
        win.title(title)
        win.transient(self)
        win.resizable(False, False)
        result = {"key": None}
        frm = ttk.Frame(win, padding=20)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text=message, wraplength=460, justify="left").pack(anchor="w")
        bar = ttk.Frame(frm)
        bar.pack(anchor="e", pady=(18, 0))

        def pick(k):
            result["key"] = k
            win.destroy()
        for key, text, style in buttons:
            ttk.Button(bar, text=text, bootstyle=style, command=lambda k=key: pick(k)).pack(
                side="left", padx=(8, 0))
        win.protocol("WM_DELETE_WINDOW", lambda: pick(None))
        win.update_idletasks()
        x = self.winfo_rootx() + max(0, (self.winfo_width() - win.winfo_width()) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - win.winfo_height()) // 3)
        win.geometry("+%d+%d" % (x, y))
        win.grab_set()
        self.wait_window(win)
        return result["key"]

    def _welcome_dialog(self):
        while True:
            key = self._choice_dialog(
                "Welcome to GrepMail",
                "GrepMail keeps its records in a database file (GrepMail.db) and creates an "
                "Excel file (GrepMail.xlsx) next to it for viewing. The Excel file can stay open "
                "while GrepMail runs.\n\nChoose where to create them, or open an existing database.",
                [("new", "Create new database...", "primary"),
                 ("open", "Open existing database...", "secondary-outline"),
                 ("exit", "Exit", "secondary-outline")])
            if key == "new":
                if self._create_database_flow():
                    return
            elif key == "open":
                if self._open_existing_flow():
                    return
            else:
                self._on_close()
                return

    def _conflict_text(self, path, kind):
        name = os.path.basename(path)
        if kind == "exists":
            return ("A file named '%s' already exists in this location.\n\n"
                    "To create a new GrepMail database here, rename or remove the existing file, "
                    "then press Retry. GrepMail never overwrites or deletes it for you." % name)
        if kind == "old":
            return ("'%s' is not a database in the current GrepMail format (it may come from an "
                    "older version).\n\nRename or remove this file, then press Retry to create a "
                    "new database in its place. GrepMail does not change it." % name)
        if kind == "view":
            return ("A file named '%s' already exists next to the database. It may come from an "
                    "older version, and GrepMail would overwrite it.\n\nRename or remove it, then "
                    "press Retry. GrepMail never changes it for you."
                    % os.path.basename(view_path_for(path)))
        return ("'%s' could not be read. Close any program that is using it and press Retry." % name)

    def _resolve_conflict(self, path, kind):
        """Returns True when a database ends up open; False if the user chose to exit
        or picked another location (which is handled by the called flow)."""
        while True:
            key = self._choice_dialog(
                "Database file conflict", self._conflict_text(path, kind),
                [("retry", "Retry", "primary"),
                 ("other", "Choose Another Location", "secondary-outline"),
                 ("exit", "Exit", "secondary-outline")])
            if key == "retry":
                state = inspect_database(path)
                if state == "missing":
                    return self._create_at(path)
                if state == "valid" and kind != "exists":
                    self._open_db(path)
                    return True
                kind = "exists" if state == "valid" else state
                continue
            if key == "other":
                return self._create_database_flow()
            self._on_close()
            return False

    def _create_at(self, path):
        try:
            self.db = ExcelDB(path, create=True)
        except SchemaError as e:
            return self._resolve_conflict(path, "exists" if e.kind == "exists" else "unreadable")
        except Exception as e:
            messagebox.showerror("Database error", "Could not create the database:\n%s" % e)
            return False
        self.cfg.db_path = path
        self.cfg.save()
        self.db_label.config(text="Database:  " + path)
        self._refresh_all()
        self._after_db_open()
        return True

    def _create_database_flow(self):
        while True:
            path = filedialog.asksaveasfilename(
                title="Where should GrepMail create its database?",
                defaultextension=".db", confirmoverwrite=False,
                initialdir=gm_platform.default_database_dir(), initialfile="GrepMail.db",
                filetypes=[("GrepMail Database", "*.db")])
            if not path:
                return False
            if os.path.exists(path):
                return self._resolve_conflict(path, "exists")
            if os.path.exists(view_path_for(path)):
                return self._resolve_conflict(path, "view")
            return self._create_at(path)

    def _open_existing_flow(self):
        path = filedialog.askopenfilename(
            title="Open a GrepMail database", initialdir=gm_platform.default_database_dir(),
            filetypes=[("GrepMail Database", "*.db")])
        if not path:
            return False
        state = inspect_database(path)
        if state == "valid":
            self._open_db(path)
            return bool(self.db)
        return self._resolve_conflict(path, state)

    def _open_db(self, path):
        try:
            self.db = ExcelDB(path)
        except SchemaError as e:
            self._resolve_conflict(path, e.kind)
            return
        except Exception as e:
            messagebox.showerror("Database error",
                                 "Could not open the database:\n%s" % e)
            return
        self.cfg.db_path = path
        self.cfg.save()
        self.db_label.config(text="Database:  " + path)
        self._refresh_all()
        self._after_db_open()

    def _export_excel_now(self):
        if not self._need_db():
            return
        try:
            self.db.export_view()
            messagebox.showinfo("Export to Excel", "The Excel view was refreshed:\n%s" % self.db.view_path)
        except PermissionError:
            messagebox.showwarning("Export to Excel",
                                   "The Excel view is open. Close it in Excel and try again.\n"
                                   "GrepMail also refreshes it automatically once it is closed.")
        except Exception as e:
            messagebox.showerror("Export to Excel", "Could not refresh the Excel view:\n%s" % e)

    def _open_excel_view(self):
        if not self._need_db():
            return
        if not os.path.exists(self.db.view_path):
            self._export_excel_now()
        if os.path.exists(self.db.view_path):
            gm_platform.open_path(self.db.view_path)

    def _after_db_open(self):
        try:
            self.db.set_custom_fields(self.cfg.custom_fields)
        except Exception:
            pass
        self._build_custom_table()
        if self.cfg.check_updates and not getattr(self, "_update_checked", False):
            self._update_checked = True
            self.after(5000, lambda: self._check_updates(manual=False))
        if self.cfg.auto_monitor and not (self.monitor and self.monitor.is_alive()):
            self.after(1500, self._start_monitor)

    # ----- updates -----
    def _about(self):
        messagebox.showinfo("About GrepMail", "GrepMail %s\n\nSaves tender attachments from your mailbox "
                            "into one folder per reference.\nLicense: AGPL-3.0-or-later\n"
                            "https://github.com/%s" % (APP_VERSION, gm_update.REPO))

    def _check_updates(self, manual=False):
        """Looks for a newer release in the background. An automatic check stays silent unless
        there is something new that was not skipped or postponed."""
        if not manual:
            if not self.cfg.check_updates:
                return
            if self.cfg.update_snooze_until:
                try:
                    if dt.datetime.now() < dt.datetime.fromisoformat(self.cfg.update_snooze_until):
                        return
                except ValueError:
                    pass

        def work():
            try:
                info = gm_update.check_latest(APP_VERSION)
                err = ""
            except Exception as e:
                info, err = None, str(e)
            self.after(0, lambda: self._update_result(info, err, manual))
        threading.Thread(target=work, daemon=True).start()

    def _update_result(self, info, err, manual):
        if err:
            if manual:
                messagebox.showwarning("Check for Updates", "Could not check for updates:\n%s\n\n"
                                       "You can also look at https://github.com/%s/releases" % (err, gm_update.REPO))
            return
        if not info:
            if manual:
                messagebox.showinfo("Check for Updates", "GrepMail %s is the latest version." % APP_VERSION)
            return
        if not manual and info["version"] == self.cfg.skipped_version:
            return
        self._update_dialog(info)

    def _update_dialog(self, info):
        win = tb.Toplevel(self)
        win.title("Update available")
        win.transient(self)
        frm = ttk.Frame(win, padding=18)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="GrepMail %s is available (you have %s)." % (info["version"], APP_VERSION),
                  font=FONT_BOLD).pack(anchor="w")
        ttk.Label(frm, text="What is new:").pack(anchor="w", pady=(10, 4))
        box = tk.Text(frm, width=70, height=12, wrap="word", relief="solid", bd=1)
        box.insert("1.0", info["notes"] or "No release notes were provided.")
        box.config(state="disabled")
        box.pack(fill="both", expand=True)
        status = ttk.Label(frm, text="" if info["asset"] else
                           "No installer for this system is attached to the release. Use the releases page.")
        status.pack(anchor="w", pady=(8, 0))
        bar = ttk.Frame(frm)
        bar.pack(anchor="e", pady=(12, 0))

        def later():
            self.cfg.update_snooze_until = (dt.datetime.now() + dt.timedelta(days=1)).isoformat()
            self.cfg.save()
            win.destroy()

        def skip():
            self.cfg.skipped_version = info["version"]
            self.cfg.save()
            win.destroy()

        def go():
            if not info["asset"]:
                gm_platform.open_path(info["page"])
                return
            now_btn.config(state="disabled")
            status.config(text="Downloading...")
            cancel = threading.Event()
            win.protocol("WM_DELETE_WINDOW", lambda: (cancel.set(), win.destroy()))

            def progress(done, total):
                if total:
                    self.after(0, lambda: status.winfo_exists() and status.config(
                        text="Downloading... %d%%" % (100 * done // total)))

            def work():
                try:
                    path = gm_update.download(info["asset"], progress=progress,
                                              app_version=APP_VERSION, cancel=cancel)
                    self.after(0, lambda: finish(path, ""))
                except Exception as e:
                    self.after(0, lambda: finish("", str(e)))
            threading.Thread(target=work, daemon=True).start()

        def finish(path, err):
            if not win.winfo_exists():
                return
            if err:
                status.config(text="The update failed: %s" % err)
                now_btn.config(state="normal")
                return
            win.destroy()
            if sys.platform.startswith("win"):
                if messagebox.askyesno("Install update", "The update was downloaded and checked.\n\n"
                                       "GrepMail will close and start the installer. Continue?"):
                    gm_platform.open_path(path)
                    self._on_close()
            else:
                messagebox.showinfo("Install update", "The update was downloaded to:\n%s\n\n"
                                    "To install it, open a terminal and run:\n"
                                    "sudo apt install \"%s\"" % (path, path))
                gm_platform.open_path(os.path.dirname(path))
        now_btn = ttk.Button(bar, text="Update Now", bootstyle="primary", command=go)
        now_btn.pack(side="left", padx=(0, 8))
        ttk.Button(bar, text="Remind Me Later", bootstyle="secondary-outline", command=later).pack(side="left", padx=(0, 8))
        ttk.Button(bar, text="Skip This Version", bootstyle="secondary-outline", command=skip).pack(side="left")
        win.protocol("WM_DELETE_WINDOW", later)
        win.update_idletasks()
        x = self.winfo_rootx() + max(0, (self.winfo_width() - win.winfo_width()) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - win.winfo_height()) // 3)
        win.geometry("+%d+%d" % (x, y))

    def _need_db(self):
        if not self.db:
            messagebox.showwarning("No database", "Create or open a database first (File menu).")
            return False
        return True

    # ---------- layout ----------
    def _build_layout(self):
        top = self.skin(tk.Frame(self, height=52), "surface")
        top.pack(fill="x")
        top.pack_propagate(False)
        if self._icon_small is not None:
            self.skin(tk.Label(top, image=self._icon_small, bd=0), "surface").pack(side="left", padx=(14, 0))
        self.skin(tk.Label(top, text="GrepMail", font=("Segoe UI", 12, "bold"), fg="#2B579A"),
                  "label").pack(side="left", padx=(8, 0))
        self.skin(tk.Label(top, text="v" + APP_VERSION, font=("Segoe UI", 9)), "muted").pack(
            side="left", padx=(6, 14), pady=(3, 0))
        self.db_label = self.skin(tk.Label(top, text="Database:  (none selected)", font=FONT,
                                           anchor="w"), "muted")
        self.db_label.pack(side="left")
        self.stop_btn = ttk.Button(top, text="\u25a0 Stop Monitoring", command=self._stop_monitor,
                                   state="disabled", bootstyle="danger-outline")
        self.stop_btn.pack(side="right", padx=(6, 16))
        self.start_btn = ttk.Button(top, text="\u25b6 Start Monitoring", command=self._start_monitor,
                                    bootstyle="primary")
        self.start_btn.pack(side="right", padx=6)
        self.status_badge = tk.Label(top, text="\u25cf Idle", font=FONT_BOLD, fg="#FFFFFF",
                                     padx=12, pady=3)
        self.status_badge.pack(side="right", padx=10)
        self.skin(tk.Frame(self, height=1), "line").pack(fill="x")

        body = self.skin(tk.Frame(self), "bg")
        body.pack(fill="both", expand=True)

        side = self.skin(tk.Frame(body, width=220), "surface")
        side.pack(side="left", fill="y")
        side.pack_propagate(False)
        self.skin(tk.Frame(body, width=1), "line").pack(side="left", fill="y")
        group = None
        for g, key, label in PAGES:
            if g != group:
                group = g
                self.skin(tk.Label(side, text=g, font=FONT_GROUP, anchor="w"), "muted").pack(
                    fill="x", padx=18, pady=(16, 4))
            self._nav_item(side, key, label)

        main = self.skin(tk.Frame(body), "bg")
        main.pack(side="left", fill="both", expand=True)
        self.page_title = self.skin(tk.Label(main, text="", font=FONT_TITLE, anchor="w"), "bg_label")
        self.page_title.pack(fill="x", padx=20, pady=(14, 6))
        self.content = self.skin(tk.Frame(main), "bg")
        self.content.pack(fill="both", expand=True, padx=20, pady=(0, 16))
        self.content.rowconfigure(0, weight=1)
        self.content.columnconfigure(0, weight=1)

        builders = {"tenders": self._page_tenders, "keywords": self._page_keywords,
                    "settings": self._page_settings, "summary": self._page_summary,
                    "offers": self._page_offers, "unmatched": self._page_unmatched,
                    "emaillog": self._page_emaillog, "tenderlog": self._page_tenderlog,
                    "monitor": self._page_monitor, "rules": self._page_rules,
                    "customlog": self._page_customlog}
        for _g, key, _label in PAGES:
            page = self.skin(tk.Frame(self.content), "bg")
            page.grid(row=0, column=0, sticky="nsew")
            builders[key](page)
            self.pages[key] = page

    def _nav_item(self, parent, key, label):
        row = tk.Frame(parent, cursor="hand2")
        accent = tk.Frame(row, width=4)
        accent.pack(side="left", fill="y")
        lbl = tk.Label(row, text=label, font=FONT, anchor="w", padx=14, pady=7, cursor="hand2")
        lbl.pack(side="left", fill="x", expand=True)
        row.pack(fill="x")
        for w in (row, lbl):
            w.bind("<Button-1>", lambda e, k=key: self.show_page(k))
            w.bind("<Enter>", lambda e, k=key: self._nav_hover(k, True))
            w.bind("<Leave>", lambda e, k=key: self._nav_hover(k, False))
        self.nav_items[key] = (row, accent, lbl)
        self.nav_labels[key] = lbl

    def _nav_hover(self, key, on):
        if key != self.current_page:
            row, _acc, lbl = self.nav_items[key]
            color = self.pal["nav_hover"] if on else self.pal["surface"]
            row.configure(bg=color)
            lbl.configure(bg=color)

    def _highlight_nav(self):
        p = self.pal
        for key, (row, accent, lbl) in self.nav_items.items():
            active = key == self.current_page
            bg = p["nav_active"] if active else p["surface"]
            row.configure(bg=bg)
            accent.configure(bg=p["primary"] if active else bg)
            lbl.configure(bg=bg, fg=p["primary"] if active else p["fg"],
                          font=FONT_BOLD if active else FONT)

    def show_page(self, key):
        self.current_page = key
        label = next(l for _g, k, l in PAGES if k == key)
        self.page_title.config(text=label)
        self.title("%s - %s" % (APP_TITLE, label))
        self.pages[key].tkraise()
        self._highlight_nav()
        if self.db:
            self._refresh_named({key})

    def _card(self, parent):
        return self.skin(tk.Frame(parent, highlightthickness=1, bd=0), "card")

    def _form_card(self, parent):
        """Bordered card whose inside matches the page background, so ttk widgets blend in."""
        card = self.skin(tk.Frame(parent, highlightthickness=1, bd=0), "formcard")
        inner = self.skin(tk.Frame(card), "bg")
        self._form_frames.add(inner)
        return card, inner

    def _label(self, parent, text, role="label", font=FONT, **kw):
        if parent in self._form_frames:
            role = {"label": "bg_label", "muted": "bg_muted"}.get(role, role)
        return self.skin(tk.Label(parent, text=text, font=font, anchor="w", **kw), role)

    def _hinted(self, entry, var, text):
        """Grey placeholder inside an entry; disappears as soon as the field has text."""
        lbl = self.skin(tk.Label(entry, text=text, font=("Segoe UI", 9, "italic"), anchor="w",
                                 cursor="xterm"), "hint")
        lbl.bind("<Button-1>", lambda e: entry.focus_set())

        def upd(*_):
            if var.get():
                lbl.place_forget()
            else:
                lbl.place(x=8, rely=0.5, anchor="w")
        var.trace_add("write", upd)
        upd()
        return entry

    def _table(self, parent, *a, **kw):
        t = TableTab(parent, *a, **kw)
        self._tables.append(t)
        return t

    def _set_status(self, state):
        self._status = state
        color, text = {"idle": (self.pal["muted"], "\u25cf Idle"),
                       "running": (self.pal["success"], "\u25cf Running"),
                       "stopped": (self.pal["danger"], "\u25cf Stopped")}[state]
        self.status_badge.config(bg=color, text=text, fg="#FFFFFF")

    # ---------- page: references ----------
    def _page_tenders(self, page):
        self.tender_ref_var = tk.StringVar()
        self.outlook_folder_var = tk.StringVar()
        self.tender_path_var = tk.StringVar()
        self.scope_var = tk.StringVar(value=SCOPE_SUB)
        card, form = self._form_card(page)
        card.pack(fill="x", pady=(0, 12))
        form.pack(fill="x", padx=16, pady=14)
        form.columnconfigure(0, minsize=140)
        form.columnconfigure(1, weight=1)
        self._label(form, "Reference").grid(row=0, column=0, sticky="w", pady=5)
        e = ttk.Entry(form, textvariable=self.tender_ref_var)
        e.grid(row=0, column=1, columnspan=2, sticky="ew")
        self._hinted(e, self.tender_ref_var, "Text that appears in the email subject, e.g. a tender, PO or project code")
        self._label(form, "Search").grid(row=1, column=0, sticky="w", pady=5)
        rb = self.skin(tk.Frame(form), "bg")
        rb.grid(row=1, column=1, columnspan=2, sticky="w")
        for value, text in ((SCOPE_SUB, "Subfolders only (default)"),
                            (SCOPE_SUB_INBOX, "Subfolders and Inbox"),
                            (SCOPE_INBOX, "Inbox only")):
            ttk.Radiobutton(rb, text=text, value=value, variable=self.scope_var,
                            command=self._scope_changed).pack(side="left", padx=(0, 18))
        self._label(form, "Mail Folder").grid(row=2, column=0, sticky="w", pady=5)
        e = ttk.Entry(form, textvariable=self.outlook_folder_var, state="readonly")
        e.grid(row=2, column=1, sticky="ew")
        self._hinted(e, self.outlook_folder_var, "The folder to watch (it is read together with the folders inside it)")
        self.choose_folder_btn = ttk.Button(form, text="Choose Folder...", command=self._browse_mail_folder,
                                            bootstyle="secondary-outline")
        self.choose_folder_btn.grid(row=2, column=2, padx=(8, 0), sticky="ew")
        self._label(form, "Local Path").grid(row=3, column=0, sticky="w", pady=5)
        e = ttk.Entry(form, textvariable=self.tender_path_var)
        e.grid(row=3, column=1, sticky="ew")
        self._hinted(e, self.tender_path_var, "Folder on this computer where the attachments are saved")
        ttk.Button(form, text="Browse...", command=self._browse_tender_path,
                   bootstyle="secondary-outline").grid(row=3, column=2, padx=(8, 0), sticky="ew")
        self._label(form, "Deadline").grid(row=4, column=0, sticky="w", pady=5)
        dl = self.skin(tk.Frame(form), "bg")
        dl.grid(row=4, column=1, columnspan=2, sticky="w")
        self.deadline_field = DateField(dl)
        self.deadline_field.pack(side="left")
        self._label(dl, "Optional. Used for the due-soon and past-deadline warnings.",
                    role="muted", font=("Segoe UI", 9)).pack(side="left", padx=(14, 0))
        btns = self.skin(tk.Frame(form), "bg")
        btns.grid(row=5, column=0, columnspan=3, sticky="w", pady=(10, 0))
        ttk.Button(btns, text="Add / Update", command=self._add_update_tender,
                   bootstyle="primary").pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Clear", command=self._clear_tender_form,
                   bootstyle="secondary-outline").pack(side="left")
        self._label(form, "Subfolders only: the chosen folder and the folders inside it.   "
                          "Subfolders and Inbox: the same, plus mail sitting directly in the Inbox.   "
                          "Inbox only: just the Inbox (no folder to choose).",
                    role="muted", font=("Segoe UI", 9), wraplength=860, justify="left").grid(
            row=6, column=0, columnspan=3, sticky="w", pady=(10, 0))

        self.tenders_table = self._table(
            page, ["Reference", "MailFolder", "Scope", "LocalPath", "Deadline", "DaysLeft", "Folder"],
            widths=[115, 140, 130, 170, 85, 85, 150], refresh_cb=self._refresh_tenders,
            headings={"MailFolder": "Mail Folder", "Scope": "Search", "LocalPath": "Local Path",
                      "DaysLeft": "Days Left", "Folder": "Local Path Status"},
            buttons=[("Delete selected", self._delete_tender, "danger-outline"),
                     ("Check folders", self._check_folders, "secondary-outline"),
                     ("Create missing folders", self._create_missing_folders, "secondary-outline")],
            tag_fn=self._tender_tags)
        self.tenders_table.pack(fill="both", expand=True)
        self.tenders_table.tree.bind("<<TreeviewSelect>>", self._on_tender_select)

    # ---------- page: keywords ----------
    def _page_keywords(self, page):
        self.kw_var = tk.StringVar()
        card, form = self._form_card(page)
        card.pack(fill="x", pady=(0, 12))
        form.pack(fill="x", padx=16, pady=14)
        form.columnconfigure(0, minsize=110)
        form.columnconfigure(1, weight=1)
        self._label(form, "Keyword").grid(row=0, column=0, sticky="w")
        e = ttk.Entry(form, textvariable=self.kw_var)
        e.grid(row=0, column=1, sticky="ew")
        self._hinted(e, self.kw_var, "A word or phrase in the sender's new text that marks the email as relevant, e.g. quotation")
        e.bind("<Return>", lambda ev: self._add_keyword())
        ttk.Button(form, text="Add Keyword", command=self._add_keyword, bootstyle="primary").grid(
            row=0, column=2, padx=(8, 0))
        self._label(form, "English or Arabic. An email is relevant when the sender's "
                          "new text (not the quoted earlier messages) contains any keyword.", role="muted", font=("Segoe UI", 9)).grid(
            row=1, column=1, sticky="w", pady=(6, 0))
        self.kw_table = self._table(
            page, ["Keyword", "Language"], widths=[420, 140], refresh_cb=self._refresh_keywords,
            buttons=[("Delete Selected", self._delete_keyword, "danger-outline")])
        self.kw_table.pack(fill="both", expand=True)

    # ---------- page: settings ----------
    def _page_settings(self, page):
        c = self.cfg
        self.set_poll = tk.StringVar(value=str(c.poll_seconds))
        self.set_warn = tk.StringVar(value=str(c.warn_days))
        self.set_regex = tk.StringVar(value=c.hint_regex)
        self.set_phrases = tk.StringVar(value=", ".join(c.hint_phrases))
        self.set_toasts = tk.BooleanVar(value=c.toasts)
        self.set_root = tk.BooleanVar(value=c.scan_root)
        self.set_theme = tk.StringVar(value=c.theme.capitalize())
        self.set_provider = tk.StringVar(value=self._provider_label(c.provider))
        self.set_preset = tk.StringVar(value=c.imap.get("preset", "Gmail"))
        self.set_host = tk.StringVar(value=c.imap.get("host", ""))
        self.set_port = tk.StringVar(value=str(c.imap.get("port", 993)))
        self.set_security = tk.StringVar(value=c.imap.get("security", "ssl"))
        self.set_user = tk.StringVar(value=c.imap.get("user", ""))
        self.set_pass = tk.StringVar(value="")
        self.set_autostart = tk.BooleanVar(value=gm_platform.autostart_enabled())
        self.set_automon = tk.BooleanVar(value=c.auto_monitor)
        self.set_updates = tk.BooleanVar(value=c.check_updates)
        self.set_scan = tk.BooleanVar(value=c.scan_enabled)
        self.set_require = tk.BooleanVar(value=c.require_scan)
        self.set_macro = tk.StringVar(value={"allow": "Allow", "flag": "Allow and Flag",
                                             "block": "Block"}[c.macro_mode])
        self.set_markweb = tk.BooleanVar(value=c.mark_web)

        nb = ttk.Notebook(page)
        nb.pack(fill="both", expand=True)

        def tab(title):
            fr = self.skin(tk.Frame(nb), "bg")
            nb.add(fr, text="  %s  " % title)
            inner = self.skin(tk.Frame(fr), "bg")
            self._form_frames.add(inner)
            inner.pack(fill="x", padx=16, pady=14)
            inner.columnconfigure(0, minsize=210)
            inner.columnconfigure(1, weight=1)
            return inner

        def note(parent, text, row):
            self._label(parent, text, role="muted", font=("Segoe UI", 9), wraplength=780,
                        justify="left").grid(row=row, column=0, columnspan=3, sticky="w", pady=(10, 0))

        # --- General
        f = tab("General")
        r = 0
        for text, widget, hint in (
                ("Poll Interval (seconds)", ttk.Spinbox(f, from_=5, to=3600, increment=5,
                                                        textvariable=self.set_poll, width=10), None),
                ("Warning Days", ttk.Spinbox(f, from_=0, to=365, textvariable=self.set_warn, width=10), None),
                ("Reference Pattern (regex)", ttk.Entry(f, textvariable=self.set_regex),
                 (self.set_regex, "Optional. Pattern that spots reference-like codes, e.g. [A-Z]{2,5}-\\d{3,}")),
                ("Wording (comma separated)", ttk.Entry(f, textvariable=self.set_phrases),
                 (self.set_phrases, "Optional. Phrases that mark a project email, e.g. invitation to tender")),
                ("Theme", ttk.Combobox(f, textvariable=self.set_theme, values=["Light", "Dark"],
                                       state="readonly", width=10), None)):
            self._label(f, text).grid(row=r, column=0, sticky="w", pady=6, padx=(0, 20))
            sticky = "ew" if isinstance(widget, ttk.Entry) and not isinstance(widget, ttk.Spinbox) else "w"
            widget.grid(row=r, column=1, sticky=sticky)
            if hint:
                self._hinted(widget, hint[0], hint[1])
            r += 1
        for text, var in (("Pop-up Alerts", self.set_toasts), ("Check Inbox for unmatched mail", self.set_root)):
            self._label(f, text).grid(row=r, column=0, sticky="w", pady=6)
            ttk.Checkbutton(f, variable=var, bootstyle="round-toggle").grid(row=r, column=1, sticky="w")
            r += 1
        note(f, "The pattern and wording only spot emails that look like project mail but match no "
                "reference (Unmatched). \"Check Inbox for unmatched mail\" also reads such mail in the Inbox. "
                "Settings are stored per user, outside the database.", r)

        # --- Mail account
        f = tab("Mail Account")
        r = 0
        provs = self._provider_choices()
        self._label(f, "Read mail from").grid(row=r, column=0, sticky="w", pady=6)
        cb = ttk.Combobox(f, textvariable=self.set_provider, values=provs, state="readonly", width=34)
        cb.grid(row=r, column=1, sticky="w")
        cb.bind("<<ComboboxSelected>>", lambda e: self._provider_changed())
        r += 1
        self.imap_frame = self.skin(tk.Frame(f), "bg")
        self._form_frames.add(self.imap_frame)
        self.imap_frame.grid(row=r, column=0, columnspan=3, sticky="ew")
        self.imap_frame.columnconfigure(0, minsize=210)
        self.imap_frame.columnconfigure(1, weight=1)
        g = self.imap_frame
        self._label(g, "Provider").grid(row=0, column=0, sticky="w", pady=6)
        pc = ttk.Combobox(g, textvariable=self.set_preset, values=list(gm_mail.PRESETS), state="readonly", width=34)
        pc.grid(row=0, column=1, sticky="w")
        pc.bind("<<ComboboxSelected>>", lambda e: self._preset_changed())
        self._label(g, "Server and Port").grid(row=1, column=0, sticky="w", pady=6)
        hp = self.skin(tk.Frame(g), "bg")
        hp.grid(row=1, column=1, sticky="w")
        ttk.Entry(hp, textvariable=self.set_host, width=30).pack(side="left")
        ttk.Entry(hp, textvariable=self.set_port, width=7).pack(side="left", padx=(8, 8))
        ttk.Combobox(hp, textvariable=self.set_security, values=["ssl", "starttls", "none"],
                     state="readonly", width=9).pack(side="left")
        self._label(g, "Email Address / Login").grid(row=2, column=0, sticky="w", pady=6)
        e = ttk.Entry(g, textvariable=self.set_user, width=40)
        e.grid(row=2, column=1, sticky="w")
        self._hinted(e, self.set_user, "The address or username of the mailbox")
        self._label(g, "Password").grid(row=3, column=0, sticky="w", pady=6)
        e = ttk.Entry(g, textvariable=self.set_pass, show="*", width=40)
        e.grid(row=3, column=1, sticky="w")
        self._hinted(e, self.set_pass, "App password (stored in the system's secure store)")
        self.preset_note = self._label(g, "", role="muted", font=("Segoe UI", 9), wraplength=780, justify="left")
        self.preset_note.grid(row=4, column=0, columnspan=3, sticky="w", pady=(6, 0))
        r += 1
        bar = self.skin(tk.Frame(f), "bg")
        bar.grid(row=r, column=0, columnspan=3, sticky="w", pady=(12, 0))
        ttk.Button(bar, text="Test connection", command=self._test_mail,
                   bootstyle="secondary-outline").pack(side="left")
        self._preset_changed(keep=True)
        self._provider_changed()

        # --- Startup
        f = tab("Startup")
        for r, (text, var) in enumerate((("Start GrepMail when I sign in", self.set_autostart),
                                         ("Start monitoring automatically", self.set_automon),
                                         ("Check for updates automatically", self.set_updates))):
            self._label(f, text).grid(row=r, column=0, sticky="w", pady=6)
            ttk.Checkbutton(f, variable=var, bootstyle="round-toggle").grid(row=r, column=1, sticky="w")
        note(f, "Starting with the system affects only your own user account and can be turned off here "
                "at any time. \"Start monitoring automatically\" begins watching your mail as soon as "
                "the app opens. The update check contacts GitHub once a day and sends only the program "
                "name and version.", 3)

        # --- Security
        f = tab("Security")
        scanner = gm_security.get_scanner()
        self._label(f, "Virus scanner").grid(row=0, column=0, sticky="w", pady=6)
        self.scan_status_label = self._label(f, "%s" % scanner.name)
        self.scan_status_label.grid(row=0, column=1, sticky="w")
        for r, (text, var) in enumerate((("Scan attachments before saving", self.set_scan),
                                         ("Hold files if no scanner is available", self.set_require)), start=1):
            self._label(f, text).grid(row=r, column=0, sticky="w", pady=6)
            ttk.Checkbutton(f, variable=var, bootstyle="round-toggle").grid(row=r, column=1, sticky="w")
        self._label(f, "Macro-enabled files").grid(row=3, column=0, sticky="w", pady=6)
        ttk.Combobox(f, textvariable=self.set_macro, values=["Allow", "Allow and Flag", "Block"],
                     state="readonly", width=18).grid(row=3, column=1, sticky="w")
        self._label(f, "Mark macro files as from the internet").grid(row=4, column=0, sticky="w", pady=6)
        ttk.Checkbutton(f, variable=self.set_markweb, bootstyle="round-toggle").grid(row=4, column=1, sticky="w")
        bar = self.skin(tk.Frame(f), "bg")
        bar.grid(row=5, column=0, columnspan=3, sticky="w", pady=(12, 0))
        ttk.Button(bar, text="Check status", command=self._check_scanner,
                   bootstyle="secondary-outline").pack(side="left")
        ttk.Button(bar, text="Update virus definitions", command=self._update_definitions,
                   bootstyle="secondary-outline").pack(side="left", padx=8)
        note(f, "Every attachment is saved to a temporary place, scanned, and only then moved to its Local "
                "Path. A file with a threat is never saved. Suspicious macros are held in the quarantine "
                "folder next to the database. 'Allow and Flag' saves other macro-enabled files and marks "
                "them in the Saved Files register. No scan can guarantee a file is safe: do not enable "
                "macros in files you do not trust.", 6)

        ttk.Button(page, text="Save settings", command=self._save_settings, bootstyle="primary").pack(
            anchor="w", pady=(12, 0))

    @staticmethod
    def _provider_label(provider):
        return "Outlook (Windows desktop)" if provider == "outlook" else "IMAP mail account"

    def _provider_choices(self):
        return (["Outlook (Windows desktop)", "IMAP mail account"] if gm_platform.IS_WINDOWS
                else ["IMAP mail account"])

    def _provider_changed(self):
        outlook = self.set_provider.get().startswith("Outlook")
        for child in self.imap_frame.winfo_children():
            try:
                child.configure(state="disabled" if outlook else "normal")
            except tk.TclError:
                pass
        if outlook:
            self.preset_note.config(text="Outlook is read directly through the installed Outlook program. "
                                         "No password is needed.")
        else:
            self._preset_changed(keep=True)

    def _preset_changed(self, keep=False):
        pre = gm_mail.PRESETS.get(self.set_preset.get())
        if not pre:
            return
        if not keep:
            self.set_host.set(pre["host"])
            self.set_port.set(str(pre["port"]))
            self.set_security.set(pre["security"])
        self.preset_note.config(text=pre["note"])

    # ---------- page: summary ----------
    def _page_summary(self, page):
        cards = self.skin(tk.Frame(page), "bg")
        cards.pack(fill="x", pady=(0, 12))
        self.stat_labels = {}
        for i, (key, caption, role) in enumerate((("offers", "Saved Files", "stat_primary"),
                                                  ("due", "Due Soon", "stat_warning"),
                                                  ("unmatched", "Unmatched emails", "stat_danger"))):
            cards.columnconfigure(i, weight=1, uniform="stat")
            c = self._card(cards)
            c.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 8, 0))
            self._label(c, caption, role="muted").pack(anchor="w", padx=16, pady=(12, 0))
            v = self.skin(tk.Label(c, text="0", font=FONT_STAT, anchor="w"), role)
            v.pack(anchor="w", padx=16, pady=(0, 10))
            self.stat_labels[key] = v
        self.summary_table = self._table(
            page, ["Reference", "Deadline", "DaysLeft", "Files", "Names", "LastSaved", "Status"],
            widths=[170, 110, 90, 90, 90, 150, 130], refresh_cb=self._refresh_summary,
            headings={"DaysLeft": "Days Left", "Files": "Saved Files", "LastSaved": "Last Saved"},
            buttons=[("Update Summary sheet", self._update_summary_sheet, "secondary-outline"),
                     ("Weekly report...", self._weekly_report, "secondary-outline")],
            tag_fn=lambda r: ("bad",) if r["Status"] == "Past Deadline" and r["Files"] == 0
            else ("warn",) if r["Status"] == "Due Soon"
            else ("past",) if r["Status"] == "Past Deadline" else ())
        self.summary_table.pack(fill="both", expand=True)

    # ---------- page: saved files ----------
    def _page_offers(self, page):
        self.offers_table = self._table(
            page, ["ReceivedDate", "Reference", "Name", "FileNo", "Revision", "Security",
                   "MatchedKeyword", "SavedFileName", "Notes"],
            widths=[125, 120, 105, 60, 70, 150, 80, 150, 100], refresh_cb=self._refresh_offers,
            headings={"ReceivedDate": "Received Date", "FileNo": "File No.", "MatchedKeyword": "Keyword",
                      "SavedFileName": "Saved File Name"},
            buttons=[("\u24d8 Details", self._show_offer_details, "secondary-outline"),
                     ("\U0001F4C4 Open File", self._open_offer_file, "secondary-outline"),
                     ("\U0001F4C1 Open Folder", self._open_offer_folder, "secondary-outline")],
            tag_fn=lambda r: ("warn",) if (s(r.get("Notes")) or "Macro" in s(r.get("Security"))) else (),
            on_double=self._show_offer_details)
        self.offers_table.pack(fill="both", expand=True)

    # ---------- page: unmatched ----------
    def _page_unmatched(self, page):
        self.skin(tk.Label(page, text="  Emails that look like tender mail but match no reference. "
                                      "Add the reference, then rescan to recover their attachments.",
                           font=FONT, anchor="w", pady=8), "banner").pack(fill="x", pady=(0, 12))
        self.unmatched_table = self._table(
            page, ["Timestamp", "SenderEmail", "Subject", "DetectedRef"],
            widths=[165, 190, 380, 150], refresh_cb=self._refresh_unmatched,
            headings={"Timestamp": "Time", "SenderEmail": "Sender Email", "DetectedRef": "Detected Ref."},
            buttons=[("Add selected as tender...", self._unmatched_to_tender, "primary"),
                     ("Remove selected", self._remove_unmatched, "secondary-outline"),
                     ("Rescan from oldest", self._rescan_from_oldest, "secondary-outline")])
        self.unmatched_table.pack(fill="both", expand=True)

    # ---------- pages: logs ----------
    def _page_emaillog(self, page):
        labels = dict(FIELD_CATALOG)
        cols = ["EmailDate", "Direction", "SenderEmail", "Subject", "Reference", "Result",
                "AttachmentCount", "MailFolder"]
        self.email_table = self._table(
            page, cols, widths=[130, 105, 160, 240, 110, 120, 85, 130],
            refresh_cb=self._refresh_email_log,
            headings={c: labels[c] for c in cols},
            tag_fn=lambda r: {"Saved": ("ok",), "Excluded": ("warn",), "Not Relevant": ("past",),
                              "Threat Found": ("bad",), "Quarantined": ("bad",),
                              "Save Failed": ("bad",), "Unmatched": ("warn",)}.get(s(r.get("Result")), ()))
        self.email_table.pack(fill="both", expand=True)

    def _page_tenderlog(self, page):
        self.tlog_table = self._table(
            page, ["Timestamp", "Action", "Reference", "Subject", "Details"],
            widths=[150, 170, 120, 250, 260], refresh_cb=self._refresh_tender_log,
            tag_fn=lambda r: {"Saved": ("ok",), "Not Relevant": ("past",), "Excluded": ("warn",),
                              "No Allowed Attachments": ("warn",), "Blocked": ("warn",),
                              "Save Failed": ("bad",), "No Local Path": ("bad",),
                              "Threat Found": ("bad",), "Quarantined": ("bad",)}.get(r.get("_action"), ()))
        self.tlog_table.pack(fill="both", expand=True)

    # ---------- page: subject rules ----------
    def _page_rules(self, page):
        self.sep_var = tk.StringVar(value={" - ": "Dash ( - )", " || ": "Double bar ( || )"}.get(
            self.cfg.name_separator, "Custom"))
        self.sep_custom = tk.StringVar(value=self.cfg.name_separator
                                       if self.cfg.name_separator not in (" - ", " || ") else "")
        self.fmt_var = tk.StringVar(value=self.cfg.file_format)
        self.preview_var = tk.StringVar()
        self.excl_var = tk.StringVar()
        card, f = self._form_card(page)
        card.pack(fill="x", pady=(0, 12))
        f.pack(fill="x", padx=16, pady=14)
        f.columnconfigure(0, minsize=160)
        f.columnconfigure(1, weight=1)
        self._label(f, "Subject layout").grid(row=0, column=0, sticky="w", pady=5)
        lay = self.skin(tk.Frame(f), "bg")
        lay.grid(row=0, column=1, sticky="w")
        self._label(lay, "Sender  -  Reference  rest of the subject", font=FONT_BOLD).pack(anchor="w")
        self._label(lay, "Example:  ABC Contracting - PRJ-001 Quotation     "
                         "(Sender = ABC Contracting, Reference = PRJ-001)",
                    role="muted", font=("Segoe UI", 9)).pack(anchor="w")
        self._label(f, "Separator").grid(row=1, column=0, sticky="w", pady=5)
        sp = self.skin(tk.Frame(f), "bg")
        sp.grid(row=1, column=1, sticky="w")
        cb = ttk.Combobox(sp, textvariable=self.sep_var, values=["Dash ( - )", "Double bar ( || )", "Custom"],
                          state="readonly", width=18)
        cb.pack(side="left")
        ce = ttk.Entry(sp, textvariable=self.sep_custom, width=12)
        ce.pack(side="left", padx=(8, 0))
        self._hinted(ce, self.sep_custom, "custom separator")
        self._label(f, "File name format").grid(row=2, column=0, sticky="w", pady=5)
        e = ttk.Entry(f, textvariable=self.fmt_var)
        e.grid(row=2, column=1, sticky="ew")
        self._hinted(e, self.fmt_var, "Use {Sender} {Reference} {Date} {Attachment Name}, e.g. {Sender} {Attachment Name}")
        self._label(f, "Preview").grid(row=3, column=0, sticky="w", pady=5)
        self._label(f, "", font=FONT_BOLD, textvariable=self.preview_var).grid(row=3, column=1, sticky="w")
        for v in (self.fmt_var, self.sep_var, self.sep_custom):
            v.trace_add("write", lambda *a: self._update_preview())
        ttk.Button(f, text="Save rules", command=self._save_rules, bootstyle="primary").grid(
            row=4, column=0, sticky="w", pady=(10, 0))
        self._label(f, "How an email is judged, in this order:  1) its subject contains a Reference   "
                       "2) its subject contains none of the exclusion words below   3) the sender's new "
                       "text contains a Body Keyword   4) the attachments are scanned and saved, named by "
                       "the format above. Sender = the text before the separator in the subject, Reference = the "
                       "matching reference, Date = the email date, Attachment Name = the attachment's "
                       "own file name (it comes from the attachment, not the subject). The extension "
                       "is always kept.", role="muted", font=("Segoe UI", 9),
                    wraplength=860, justify="left").grid(row=5, column=0, columnspan=2, sticky="w", pady=(10, 0))

        card2, g = self._form_card(page)
        card2.pack(fill="x", pady=(0, 12))
        g.pack(fill="x", padx=16, pady=14)
        g.columnconfigure(1, weight=1)
        self._label(g, "Exclusion word").grid(row=0, column=0, sticky="w", padx=(0, 20))
        e = ttk.Entry(g, textvariable=self.excl_var)
        e.grid(row=0, column=1, sticky="ew")
        e.bind("<Return>", lambda ev: self._add_exclusion())
        self._hinted(e, self.excl_var, "Emails whose subject contains this are never saved, e.g. clarification")
        ttk.Button(g, text="Add Word", command=self._add_exclusion, bootstyle="primary").grid(
            row=0, column=2, padx=(8, 0))
        self.excl_table = self._table(
            page, ["Word", "Language"], widths=[420, 140], refresh_cb=self._refresh_exclusions,
            buttons=[("Delete Selected", self._delete_exclusion, "danger-outline")])
        self.excl_table.pack(fill="both", expand=True)
        self._update_preview()

    def _current_separator(self):
        choice = self.sep_var.get()
        if choice.startswith("Dash"):
            return " - "
        if choice.startswith("Double"):
            return " || "
        custom = self.sep_custom.get()
        return custom if custom.strip() else " - "

    def _update_preview(self):
        name = extract_name("RE: ABC Contracting%sPRJ-001 Quotation" % self._current_separator(),
                            self._current_separator())
        self.preview_var.set(build_file_name(self.fmt_var.get(), name, "PRJ-001", "2026-11-03",
                                             "Quotation.pdf"))

    def _save_rules(self):
        fmt = upgrade_format(self.fmt_var.get().strip()) or DEFAULT_FILE_FORMAT
        self.cfg.name_separator = self._current_separator()
        self.cfg.file_format = fmt
        self.cfg.save()
        self._log_line("Subject rules saved.")
        self._show_toast("Rules saved", "Separator and file name format updated.")

    def _add_exclusion(self):
        if not self._need_db():
            return
        try:
            self.db.add_exclusion(self.excl_var.get())
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        self.excl_var.set("")
        self._refresh_exclusions()

    def _delete_exclusion(self):
        sel = self.excl_table.selected()
        if not sel or not self._need_db():
            return
        if messagebox.askyesno("Confirm", "Delete %d word(s)?" % len(sel)):
            for r in sel:
                self.db.delete_exclusion(r["Word"])
            self._refresh_exclusions()

    def _refresh_exclusions(self):
        if self.db:
            self.excl_table.set_rows([{"Word": w, "Language": keyword_language(w)}
                                      for w in self.db.list_exclusions()])

    # ---------- page: custom log ----------
    def _page_customlog(self, page):
        self.page_custom = page
        self.field_vars = {k: tk.BooleanVar(value=k in self.cfg.custom_fields) for k, _ in FIELD_CATALOG}
        flt = self.cfg.custom_filter
        self.cf_dir = tk.StringVar(value=flt.get("direction", "Both"))
        self.cf_ref = tk.BooleanVar(value=bool(flt.get("only_reference")))
        self.cf_att = tk.BooleanVar(value=bool(flt.get("only_attachments")))
        self.cf_contains = tk.StringVar(value=flt.get("contains", ""))
        card, f = self._form_card(page)
        card.pack(fill="x", pady=(0, 12))
        f.pack(fill="x", padx=16, pady=12)
        self._label(f, "Fields to keep", font=FONT_BOLD).grid(row=0, column=0, columnspan=4, sticky="w")
        for i, (key, label) in enumerate(FIELD_CATALOG):
            ttk.Checkbutton(f, text=label, variable=self.field_vars[key]).grid(
                row=1 + i // 4, column=i % 4, sticky="w", padx=(0, 22), pady=2)
        r = 2 + (len(FIELD_CATALOG) - 1) // 4
        self._label(f, "Only keep emails that match", font=FONT_BOLD).grid(
            row=r, column=0, columnspan=4, sticky="w", pady=(10, 2))
        fl = self.skin(tk.Frame(f), "bg")
        fl.grid(row=r + 1, column=0, columnspan=4, sticky="w")
        self._label(fl, "Direction").pack(side="left", padx=(0, 6))
        ttk.Combobox(fl, textvariable=self.cf_dir, values=["Both", "Incoming", "Outgoing"],
                     state="readonly", width=10).pack(side="left", padx=(0, 16))
        ttk.Checkbutton(fl, text="Has a Reference", variable=self.cf_ref).pack(side="left", padx=(0, 16))
        ttk.Checkbutton(fl, text="Has attachments", variable=self.cf_att).pack(side="left", padx=(0, 16))
        self._label(fl, "Subject contains").pack(side="left", padx=(0, 6))
        e = ttk.Entry(fl, textvariable=self.cf_contains, width=24)
        e.pack(side="left")
        self._hinted(e, self.cf_contains, "optional text")
        ttk.Button(f, text="Apply", command=self._apply_custom, bootstyle="primary").grid(
            row=r + 2, column=0, sticky="w", pady=(10, 0))
        self._label(f, "Every new email is also written to the Custom Log sheet, with only the fields "
                       "and filters chosen here. Fields you add later become new columns at the right.",
                    role="muted", font=("Segoe UI", 9), wraplength=860, justify="left").grid(
            row=r + 3, column=0, columnspan=4, sticky="w", pady=(8, 0))
        self.custom_holder = self.skin(tk.Frame(page), "bg")
        self.custom_holder.pack(fill="both", expand=True)
        self.custom_table = None

    def _build_custom_table(self):
        """The table follows the chosen fields, so it is rebuilt when they change."""
        if self.custom_table is not None:
            try:
                self._tables.remove(self.custom_table)
            except ValueError:
                pass
            self.custom_table.destroy()
        labels = dict(FIELD_CATALOG)
        cols = [k for k in self.cfg.custom_fields if k in labels] or ["Subject"]
        self.custom_table = self._table(
            self.custom_holder, cols, widths=[140] * len(cols),
            refresh_cb=self._refresh_custom_log, headings={c: labels[c] for c in cols})
        self.custom_table.pack(fill="both", expand=True)
        self.custom_table.apply_palette(self.pal)
        self._refresh_custom_log()

    def _apply_custom(self):
        if not self._need_db():
            return
        fields = [k for k, _ in FIELD_CATALOG if self.field_vars[k].get()]
        if not fields:
            messagebox.showwarning("Fields", "Choose at least one field.")
            return
        self.cfg.custom_fields = fields
        self.cfg.custom_filter = {"direction": self.cf_dir.get(), "only_reference": self.cf_ref.get(),
                                  "only_attachments": self.cf_att.get(),
                                  "contains": self.cf_contains.get().strip()}
        self.cfg.save()
        try:
            self.db.set_custom_fields(fields)
        except Exception as e:
            messagebox.showerror("Error", "Could not update the Custom Log sheet:\n%s" % e)
            return
        self._build_custom_table()
        self._log_line("Custom Log fields and filters saved.")
        self._show_toast("Custom Log", "Fields and filters applied to new emails.")

    def _refresh_custom_log(self):
        if self.db and self.custom_table is not None:
            self.custom_table.set_rows(list(reversed(self.db.list_custom_log())))

    def _page_monitor(self, page):
        card, f = self._form_card(page)
        card.pack(fill="x", pady=(0, 12))
        f.pack(fill="x", padx=16, pady=12)
        f.columnconfigure(4, weight=1)
        self._label(f, "Retroactive scan", font=FONT_BOLD).grid(row=0, column=0, columnspan=4, sticky="w")
        self._label(f, "Scan Date").grid(row=1, column=0, sticky="w", pady=(8, 0), padx=(0, 12))
        self.scan_date = DateField(f)
        self.scan_date.set((dt.date.today() - dt.timedelta(days=7)).isoformat())
        self.scan_date.grid(row=1, column=1, sticky="w", pady=(8, 0))
        self.scan_btn = ttk.Button(f, text="Run rescan", command=self._start_rescan, bootstyle="primary")
        self.scan_btn.grid(row=1, column=2, padx=(12, 6), pady=(8, 0))
        self.cancel_scan_btn = ttk.Button(f, text="Cancel", state="disabled", bootstyle="secondary-outline",
                                          command=lambda: self.rescan_stop.set())
        self.cancel_scan_btn.grid(row=1, column=3, pady=(8, 0))
        self._label(f, "Re-processes your mail from that date. Emails and files that were "
                       "already logged or saved are skipped, so it is safe to run repeatedly.",
                    role="muted", font=("Segoe UI", 9), wraplength=820, justify="left").grid(
            row=2, column=0, columnspan=5, sticky="w", pady=(8, 0))
        bar = self.skin(tk.Frame(page), "bg")
        bar.pack(fill="x", pady=(0, 6))
        self._label(bar, "Live log", role="bg_label", font=FONT_BOLD).pack(side="left")
        ttk.Button(bar, text="Clear log", command=self._clear_log, bootstyle="secondary-outline").pack(side="right")
        wrap = self.skin(tk.Frame(page), "bg")
        wrap.pack(fill="both", expand=True)
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        self.log_text = self.skin(tk.Text(wrap, wrap="none", state="disabled", font=FONT_MONO,
                                          bd=0, padx=12, pady=10, height=10), "console")
        vs = ttk.Scrollbar(wrap, orient="vertical", command=self.log_text.yview)
        hs = ttk.Scrollbar(wrap, orient="horizontal", command=self.log_text.xview)
        self.log_text.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        for tag, color in (("INFO", "#D7DAE0"), ("DEBUG", "#8A8F98"), ("WARN", "#FFB454"),
                           ("ERROR", "#FF6B6B"), ("SCAN", "#5CC8FF"), ("PROCESSED", "#7EE787")):
            self.log_text.tag_configure(tag, foreground=color)

    # ---------- refreshers ----------
    def _refresh_all(self):
        for fn in (self._refresh_tenders, self._refresh_keywords, self._refresh_exclusions,
                   self._refresh_custom_log, self._refresh_summary,
                   self._refresh_offers, self._refresh_unmatched, self._refresh_email_log,
                   self._refresh_tender_log):
            try:
                fn()
            except Exception as e:
                self._log_line("Refresh error: %s" % e)

    def _refresh_named(self, names):
        table = {"tenders": self._refresh_tenders, "keywords": self._refresh_keywords,
                 "emaillog": self._refresh_email_log, "tenderlog": self._refresh_tender_log,
                 "offers": self._refresh_offers, "unmatched": self._refresh_unmatched,
                 "summary": self._refresh_summary, "customlog": self._refresh_custom_log,
                 "rules": self._refresh_exclusions}
        current = getattr(self, "current_page", None)
        for n in names:
            if n not in table:
                continue
            if n not in (current, "unmatched"):
                self.stale.add(n)       # refreshed when that page is opened
                continue
            try:
                table[n]()
            except Exception:
                pass

    def _tender_tags(self, r):
        if r.get("Folder") not in ("OK", "", "Not checked"):
            return ("bad",)
        left = r.get("DaysLeft")
        if left != "":
            if left < 0:
                return ("past",)
            if left <= self.cfg.warn_days:
                return ("warn",)
        return ()

    def _refresh_tenders(self):
        if not self.db:
            return
        rows = []
        scopes = self.db.tender_scopes()
        for ref, path, outlook_folder, deadline in self.db.list_tenders():
            left = days_left(deadline)
            rows.append({"Reference": ref, "MailFolder": outlook_folder,
                         "Scope": scopes.get(ref.upper(), SCOPE_SUB), "LocalPath": path,
                         "Deadline": deadline, "DaysLeft": "" if left is None else left,
                         "Folder": self.folder_status.get(ref.upper(), "Not checked")})
        self.tenders_table.set_rows(rows)

    def _refresh_keywords(self):
        if self.db:
            self.kw_table.set_rows([{"Keyword": k, "Language": keyword_language(k)}
                                    for k in self.db.list_keywords()])

    def _refresh_summary(self):
        if not self.db:
            return
        rows = self.db.tender_summary(self.cfg.warn_days)
        self.summary_table.set_rows(rows)
        self.stat_labels["offers"].config(text=str(sum(r["Files"] for r in rows)))
        self.stat_labels["due"].config(text=str(sum(1 for r in rows if r["Status"] == "Due Soon")))
        self.stat_labels["unmatched"].config(text=str(len(self.db.list_unmatched())))

    def _refresh_offers(self):
        if self.db:
            self.offers_table.set_rows(list(reversed(self.db.list_offers(limit=3000))))

    def _refresh_unmatched(self):
        if not self.db:
            return
        raw = list(reversed(self.db.list_unmatched()))
        self.unmatched_table.set_rows([dict(r, DetectedRef=s(r.get("DetectedRef")) or "N/A") for r in raw])
        n = len(raw)
        self.nav_labels["unmatched"].config(text="Unmatched  (%d)" % n if n else "Unmatched")

    def _refresh_email_log(self):
        if self.db:
            arrow = {"Incoming": "\u25bc Incoming", "Outgoing": "\u25b2 Outgoing"}
            self.email_table.set_rows([dict(r, Direction=arrow.get(s(r.get("Direction")), s(r.get("Direction"))))
                                       for r in reversed(self.db.list_email_log())])

    def _refresh_tender_log(self):
        if self.db:
            self.tlog_table.set_rows([dict(r, _action=s(r.get("Action")), Action="\u25cf " + s(r.get("Action")))
                                      for r in reversed(self.db.list_tender_log())])

    # ---------- tender actions ----------
    def _browse_tender_path(self):
        p = filedialog.askdirectory(title="Select the folder where attachments are saved")
        if p:
            self.tender_path_var.set(os.path.normpath(p))

    def _scope_changed(self):
        inbox_only = self.scope_var.get() == SCOPE_INBOX
        self.choose_folder_btn.config(state="disabled" if inbox_only else "normal")
        if inbox_only:
            self.outlook_folder_var.set(INBOX_TOKEN)
        elif self.outlook_folder_var.get() == INBOX_TOKEN and self.scope_var.get() == SCOPE_SUB:
            pass  # (Inbox) with Subfolders means every folder inside the Inbox

    def _browse_mail_folder(self):
        if self.cfg.provider == "outlook":
            self._browse_outlook_folder()
        else:
            self._browse_imap_folder()

    def _browse_outlook_folder(self):
        if not HAS_OUTLOOK:
            messagebox.showerror("pywin32 missing", "Run: pip install pywin32")
            return
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)  # olFolderInbox
            folder = namespace.PickFolder()
            if folder is None:
                return
            inbox_path = str(inbox.FolderPath).rstrip("\\")
            selected_path = str(folder.FolderPath)
            inbox_prefix = inbox_path + "\\"
            if selected_path.rstrip("\\").casefold() == inbox_path.casefold():
                self.outlook_folder_var.set(INBOX_TOKEN)
                return
            if not selected_path.casefold().startswith(inbox_prefix.casefold()):
                messagebox.showwarning("Invalid mail folder",
                                       "Choose the Inbox or a folder inside your default Inbox.")
                return
            self.outlook_folder_var.set(selected_path[len(inbox_prefix):])
        except Exception as e:
            messagebox.showerror("Mail folder error", str(e))
        finally:
            pythoncom.CoUninitialize()

    def _browse_imap_folder(self):
        ok, why = provider_ready(self.cfg)
        if not ok:
            messagebox.showwarning("Mail account", why)
            return
        try:
            self.config(cursor="watch")
            self.update_idletasks()
            pw = imap_password(self.cfg)
            if not pw:
                raise RuntimeError("No password stored for this mail account. Enter it in Settings.")
            mb = gm_mail.connect_imap(self.cfg.imap, pw)
            try:
                choices = mb.relative_paths()
            finally:
                mb.close()
        except Exception as e:
            messagebox.showerror("Mail folder error", str(e))
            return
        finally:
            self.config(cursor="")
        win = tb.Toplevel(self)
        win.title("Choose a mail folder")
        win.geometry("420x480")
        win.transient(self)
        lb = tk.Listbox(win, activestyle="none", font=FONT)
        lb.pack(fill="both", expand=True, padx=12, pady=(12, 6))
        for _store, disp in choices:
            lb.insert("end", disp)

        def choose(_e=None):
            sel = lb.curselection()
            if sel:
                self.outlook_folder_var.set(choices[sel[0]][0])
                win.destroy()
        lb.bind("<Double-1>", choose)
        ttk.Button(win, text="Choose", command=choose, bootstyle="primary").pack(pady=(0, 12))
        win.grab_set()

    def _clear_tender_form(self):
        self.tender_ref_var.set("")
        self.outlook_folder_var.set("")
        self.scope_var.set(SCOPE_SUB)
        self._scope_changed()
        self.tender_path_var.set("")
        self.deadline_field.set("")

    def _add_update_tender(self):
        if not self._need_db():
            return
        try:
            self.db.add_or_update_tender(
                self.tender_ref_var.get(), self.tender_path_var.get(),
                self.outlook_folder_var.get(), self.deadline_field.get(),
                scope=self.scope_var.get())
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        self._clear_tender_form()
        self._refresh_tenders()
        self._refresh_summary()

    def _delete_tender(self):
        sel = self.tenders_table.selected()
        if not sel or not self._need_db():
            return
        if messagebox.askyesno("Confirm", "Delete %d reference(s)?" % len(sel)):
            for r in sel:
                self.db.delete_tender(r["Reference"])
            self._refresh_tenders()
            self._refresh_summary()

    def _on_tender_select(self, event):
        sel = self.tenders_table.selected()
        if sel:
            self.tender_ref_var.set(sel[0]["Reference"])
            self.outlook_folder_var.set(sel[0]["MailFolder"])
            self.scope_var.set(sel[0].get("Scope") or SCOPE_SUB)
            self._scope_changed()
            self.outlook_folder_var.set(sel[0]["MailFolder"])
            self.tender_path_var.set(sel[0]["LocalPath"])
            self.deadline_field.set(sel[0]["Deadline"])

    def _check_folders(self):
        if not self._need_db():
            return
        tenders = self.db.list_tenders()
        self._log_line("Checking %d tender folder(s)..." % len(tenders))

        def work():
            status = {ref.upper(): check_folder(path) for ref, path, _, _ in tenders}
            bad = sum(1 for v in status.values() if v != "OK")
            self.emit("folders", status)
            self.emit("log", "Folder check finished - %d problem(s)." % bad)
        threading.Thread(target=work, daemon=True).start()

    def _create_missing_folders(self):
        if not self._need_db():
            return
        missing = [(ref, p) for ref, p, _, _ in self.db.list_tenders() if p and not os.path.isdir(p)]
        if not missing:
            messagebox.showinfo("Folders", "No missing folders found (run 'Check folders' first if unsure).")
            return
        if not messagebox.askyesno("Create folders", "Create %d missing folder(s)?" % len(missing)):
            return
        failed = []
        for ref, p in missing:
            try:
                os.makedirs(p, exist_ok=True)
            except Exception as e:
                failed.append("%s: %s" % (ref, e))
        if failed:
            messagebox.showwarning("Some folders failed", "\n".join(failed[:8]))
        self._check_folders()

    # ---------- keyword actions ----------
    def _add_keyword(self):
        if not self._need_db():
            return
        try:
            self.db.add_keyword(self.kw_var.get())
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        self.kw_var.set("")
        self._refresh_keywords()

    def _delete_keyword(self):
        sel = self.kw_table.selected()
        if not sel or not self._need_db():
            return
        if messagebox.askyesno("Confirm", "Delete %d keyword(s)?" % len(sel)):
            for r in sel:
                self.db.delete_keyword(r["Keyword"])
            self._refresh_keywords()

    # ---------- offer register actions ----------
    def _open_path(self, path):
        if path and os.path.exists(path):
            try:
                gm_platform.open_path(path)
            except Exception as e:
                messagebox.showerror("Open", str(e))
        else:
            messagebox.showinfo("Not found", "Path not found:\n%s" % path)

    def _open_offer_file(self):
        sel = self.offers_table.selected()
        if sel:
            self._open_path(s(sel[0].get("FilePath")))

    def _open_offer_folder(self):
        sel = self.offers_table.selected()
        if sel:
            self._open_path(os.path.dirname(s(sel[0].get("FilePath"))))

    def _show_offer_details(self):
        sel = self.offers_table.selected()
        if not sel:
            messagebox.showinfo("Details", "Select a saved file in the table first.")
            return
        r = sel[0]
        win = tb.Toplevel(self)
        win.title("Saved file details")
        win.geometry("700x500")
        win.transient(self)
        frm = ttk.Frame(win, padding=16)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)
        fields = (("Received Date", "ReceivedDate"), ("Reference", "Reference"),
                  ("Name", "Name"), ("File No.", "FileNo"), ("Revision", "Revision"),
                  ("Matched Keyword", "MatchedKeyword"), ("Security", "Security"),
                  ("Saved File Name", "SavedFileName"), ("Original Name", "OriginalName"),
                  ("File Path", "FilePath"), ("Sender Email", "SenderEmail"),
                  ("File Hash", "FileHash"), ("Notes", "Notes"))
        for i, (label, key) in enumerate(fields):
            ttk.Label(frm, text=label, bootstyle="secondary").grid(row=i, column=0, sticky="w", pady=4, padx=(0, 16))
            var = tk.StringVar(value=s(r.get(key)))
            ttk.Entry(frm, textvariable=var, state="readonly").grid(row=i, column=1, sticky="ew")
        bar = ttk.Frame(frm)
        bar.grid(row=len(fields), column=0, columnspan=2, sticky="e", pady=(16, 0))
        ttk.Button(bar, text="Open File", bootstyle="secondary-outline",
                   command=lambda: self._open_path(s(r.get("FilePath")))).pack(side="left", padx=4)
        ttk.Button(bar, text="Open Folder", bootstyle="secondary-outline",
                   command=lambda: self._open_path(os.path.dirname(s(r.get("FilePath"))))).pack(side="left", padx=4)
        ttk.Button(bar, text="Close", bootstyle="primary", command=win.destroy).pack(side="left", padx=4)

    # ---------- unmatched actions ----------
    def _unmatched_to_tender(self):
        sel = self.unmatched_table.selected()
        if not sel:
            return
        ref = s(sel[0].get("DetectedRef"))
        self._clear_tender_form()
        self.tender_ref_var.set("" if ref == "N/A" else ref)
        self.show_page("tenders")
        self._log_line("Choose the mail folder and save folder, press 'Add / Update', then rescan.")

    def _remove_unmatched(self):
        sel = self.unmatched_table.selected()
        if not sel or not self._need_db():
            return
        self.db.remove_unmatched([s(r.get("EmailId")) for r in sel])
        self._refresh_unmatched()
        self._refresh_summary()

    def _rescan_from_oldest(self):
        times = [parse_ts(r.get("Timestamp")) for r in self.unmatched_table.rows]
        times = [t for t in times if t]
        if not times:
            messagebox.showinfo("Rescan", "There are no unmatched emails to rescan.")
            return
        self.scan_date.set((min(times) - dt.timedelta(days=1)).date().isoformat())
        self.show_page("monitor")
        self._start_rescan()

    # ---------- summary / report ----------
    def _update_summary_sheet(self):
        if not self._need_db():
            return
        try:
            self.db.refresh_summary_sheet(self.cfg.warn_days)
            self._log_line("Summary sheet updated in the database.")
            self._show_toast("Summary updated", "Summary sheet written to the database.")
        except Exception as e:
            messagebox.showerror("Error", "Could not update the sheet:\n%s" % e)

    def _weekly_report(self):
        if not self._need_db():
            return
        default = "Weekly_Report_%s.xlsx" % dt.date.today().isoformat()
        path = filedialog.asksaveasfilename(
            title="Save weekly report", initialfile=default, defaultextension=".xlsx",
            initialdir=os.path.dirname(self.cfg.db_path), filetypes=[("Excel Workbook", "*.xlsx")])
        if not path:
            return
        try:
            r = self.db.build_weekly_report(path, 7, self.cfg.warn_days)
        except Exception as e:
            messagebox.showerror("Error", "Could not create the report:\n%s" % e)
            return
        self._log_line("Weekly report saved: %s (%d files, %d references quiet, %d unmatched)" % (
            path, r["offers"], r["quiet"], r["unmatched"]))
        try:
            gm_platform.open_path(path)
        except Exception:
            pass

    # ---------- settings ----------
    def _save_settings(self):
        try:
            poll = int(self.set_poll.get())
            warn = int(self.set_warn.get())
            if poll < 5 or warn < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid value", "Seconds must be 5 or more; days must be 0 or more.")
            return
        rx = self.set_regex.get().strip()
        if rx:
            try:
                re.compile(rx)
            except re.error as e:
                messagebox.showerror("Invalid pattern", "The regular expression is not valid:\n%s" % e)
                return
        c = self.cfg
        c.poll_seconds, c.warn_days = poll, warn
        c.hint_regex = rx
        c.hint_phrases = [p.strip() for p in self.set_phrases.get().split(",") if p.strip()]
        c.toasts = self.set_toasts.get()
        c.scan_root = self.set_root.get()
        # mail account
        c.provider = "outlook" if self.set_provider.get().startswith("Outlook") else "imap"
        try:
            port = int(self.set_port.get() or 993)
        except ValueError:
            messagebox.showerror("Invalid value", "The mail server port must be a number.")
            return
        c.imap = {"preset": self.set_preset.get(), "host": self.set_host.get().strip(), "port": port,
                  "user": self.set_user.get().strip(), "security": self.set_security.get()}
        pw = self.set_pass.get()
        if pw and c.provider == "imap":
            if gm_mail.save_password(c.imap["user"], c.imap["host"], pw):
                c.session_password = ""
            else:
                c.session_password = pw
                self._log_line("Password kept for this session only (install 'keyring' to store it securely).")
            self.set_pass.set("")
        # startup
        ok, msg = gm_platform.set_autostart(self.set_autostart.get())
        if not ok:
            messagebox.showwarning("Start with system", "Could not change the start-up setting:\n%s" % msg)
            self.set_autostart.set(gm_platform.autostart_enabled())
        c.auto_monitor = self.set_automon.get()
        c.check_updates = self.set_updates.get()
        # security
        c.scan_enabled = self.set_scan.get()
        c.require_scan = self.set_require.get()
        c.macro_mode = {"Allow": "allow", "Allow and Flag": "flag", "Block": "block"}[self.set_macro.get()]
        c.mark_web = self.set_markweb.get()
        theme = self.set_theme.get().lower()
        if theme in PALETTES and theme != self.theme_name:
            c.theme = theme
            self._apply_theme(theme)
        c.save()
        self._log_line("Settings saved.")
        self._show_toast("Settings Saved", "Configuration applied.\nPoll interval updated to %ds." % poll)
        self._refresh_all()

    def _test_mail(self):
        ok, why = provider_ready(self.cfg) if self.set_provider.get().startswith("Outlook") == (self.cfg.provider == "outlook") else (True, "")
        # test with the values currently on screen, without saving them
        probe = Config.__new__(Config)
        probe.__dict__.update(self.cfg.__dict__)
        probe.provider = "outlook" if self.set_provider.get().startswith("Outlook") else "imap"
        try:
            port = int(self.set_port.get() or 993)
        except ValueError:
            messagebox.showerror("Invalid value", "The mail server port must be a number.")
            return
        probe.imap = {"preset": self.set_preset.get(), "host": self.set_host.get().strip(), "port": port,
                      "user": self.set_user.get().strip(), "security": self.set_security.get()}
        if self.set_pass.get():
            probe.session_password = self.set_pass.get()
        ok, why = provider_ready(probe)
        if not ok:
            messagebox.showwarning("Test connection", why)
            return
        self.config(cursor="watch")
        self.update_idletasks()
        try:
            if probe.provider == "outlook":
                pythoncom.CoInitialize()
                try:
                    conn = connect_outlook()
                    detail = "Outlook connected. Inbox: %s" % safe(lambda: conn.inbox.Name, "Inbox")
                finally:
                    pythoncom.CoUninitialize()
            else:
                conn = connect_mail(probe)
                n = len(safe(lambda: [f for f in conn.inbox.Folders], []))
                detail = "Connected to %s. Inbox has %d subfolder(s). Sent folder: %s." % (
                    probe.imap["host"], n, "found" if conn.sent is not None else "not found")
                conn.close()
            messagebox.showinfo("Test connection", detail)
        except Exception as e:
            messagebox.showerror("Test connection", str(e))
        finally:
            self.config(cursor="")

    def _check_scanner(self):
        self.scan_status_label.config(text="Checking...")

        def work():
            sc = gm_security.get_scanner()
            info = sc.status()
            if not info["available"]:
                text = "%s: %s" % (info["engine"], info["note"])
            else:
                when = info["definitions_date"]
                text = "%s - definitions %s%s%s" % (
                    info["engine"],
                    when.strftime("%Y-%m-%d") if when else "date unknown",
                    " (%d day(s) old - update recommended)" % info["age_days"]
                    if info["age_days"] is not None and info["age_days"] > 3 else "",
                    "" if info.get("realtime") is None else
                    ", real-time protection %s" % ("on" if info["realtime"] else "OFF"))
                if info["note"]:
                    text += ". " + info["note"]
            self.after(0, lambda: self.scan_status_label.config(text=text))
        threading.Thread(target=work, daemon=True).start()

    def _update_definitions(self):
        self.scan_status_label.config(text="Updating virus definitions...")

        def work():
            ok, msg = gm_security.get_scanner().update_definitions()
            def done():
                self.scan_status_label.config(text=msg.splitlines()[0] if msg else "")
                (messagebox.showinfo if ok else messagebox.showwarning)("Virus definitions", msg)
            self.after(0, done)
        threading.Thread(target=work, daemon=True).start()

    # ---------- thread -> GUI hand-off ----------
    def emit(self, kind, *args):
        self.q.put((kind, args))

    def _poll_queue(self):
        try:
            while True:
                kind, args = self.q.get_nowait()
                if kind == "log":
                    self._log_line(args[0])
                elif kind == "toast":
                    if self.cfg.toasts:
                        self._show_toast(args[0], args[1])
                elif kind == "refresh":
                    self.dirty.update(args[0])
                elif kind == "folders":
                    self.folder_status = args[0]
                    self._refresh_tenders()
                elif kind == "scan_done":
                    self.scan_btn.config(state="normal")
                    self.cancel_scan_btn.config(state="disabled")
        except queue.Empty:
            pass
        if self.dirty and time.time() - self.last_refresh > 2.0 and self.db:
            names, self.dirty = set(self.dirty), set()
            self.last_refresh = time.time()
            self._refresh_named(names)
        self.after(500, self._poll_queue)

    # ---------- monitor controls ----------
    def _start_monitor(self):
        if not self._need_db():
            return
        ok, why = provider_ready(self.cfg)
        if not ok:
            messagebox.showerror("Mail account", why)
            return
        if self.monitor and self.monitor.is_alive():
            return
        self.stop_event.clear()
        self.monitor = MonitorThread(self.db, self.cfg, self.emit, self.stop_event)
        self.monitor.start()
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self._set_status("running")

    def _stop_monitor(self):
        self.stop_event.set()
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self._set_status("stopped")

    def _start_rescan(self):
        if not self._need_db():
            return
        ok, why = provider_ready(self.cfg)
        if not ok:
            messagebox.showerror("Mail account", why)
            return
        d = parse_date(self.scan_date.get())
        if d is None:
            messagebox.showerror("Invalid date", "Enter the date as YYYY-MM-DD.")
            return
        if self.rescan and self.rescan.is_alive():
            return
        self.rescan_stop.clear()
        self.rescan = RescanThread(self.db, self.cfg, self.emit,
                                   dt.datetime.combine(d, dt.time.min), self.rescan_stop)
        self.rescan.start()
        self.scan_btn.config(state="disabled")
        self.cancel_scan_btn.config(state="normal")

    def _clear_log(self):
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

    def _log_line(self, msg):
        level = classify_log(msg)
        self.log_text.config(state="normal")
        self.log_text.insert("end", "[%s] %s - %s\n" % (level, dt.datetime.now().strftime("%H:%M:%S"), msg), level)
        if int(self.log_text.index("end-1c").split(".")[0]) > 2000:
            self.log_text.delete("1.0", "200.0")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    # ---------- pop-up alerts (no extra libraries needed) ----------
    def _show_toast(self, title, message):
        try:
            low = title.lower()
            accent = (self.pal["danger"] if "fail" in low else
                      self.pal["warning"] if any(k in low for k in ("unmatched", "deadline", "folder")) else
                      self.pal["success"] if any(k in low for k in ("offer", "saved", "complete", "updated")) else
                      self.pal["primary"])
            win = tk.Toplevel(self)
            win.overrideredirect(True)
            win.attributes("-topmost", True)
            win.configure(bg=self.pal["toast_bg"])
            tk.Frame(win, bg=accent, width=5).pack(side="left", fill="y")
            frame = tk.Frame(win, bg=self.pal["toast_bg"], padx=12, pady=10)
            frame.pack(side="left")
            tk.Label(frame, text=title, bg=self.pal["toast_bg"], fg="#FFFFFF",
                     font=("Segoe UI", 10, "bold"), anchor="w").pack(fill="x")
            tk.Label(frame, text=message, bg=self.pal["toast_bg"], fg="#D1D5DB", justify="left",
                     anchor="w", wraplength=290, font=("Segoe UI", 9)).pack(fill="x")
            win.update_idletasks()
            w, h = max(320, win.winfo_width()), win.winfo_height()
            self.toasts = [t for t in self.toasts if t.winfo_exists()]
            offset = sum(t.winfo_height() + 8 for t in self.toasts)
            x = self.winfo_screenwidth() - w - 20
            y = self.winfo_screenheight() - h - 60 - offset
            win.geometry("%dx%d+%d+%d" % (w, h, x, max(20, y)))
            self.toasts.append(win)
            for widget in (win, frame):
                widget.bind("<Button-1>", lambda e: win.destroy())
            win.after(9000, lambda: win.winfo_exists() and win.destroy())
        except Exception:
            pass

    def _on_close(self):
        self.stop_event.set()
        self.rescan_stop.set()
        for t in (self.monitor, self.rescan):
            if t and t.is_alive():
                t.join(timeout=5)
        if self.db:
            try:
                self.db.close()
            except Exception:
                pass
        self.destroy()


# =========================================================
# EMBEDDED ICONS (PNG, base64) - generated from App_Icon.jpg
# =========================================================
ICON_128_B64 = (
    "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"
    "AAAAAElFTkSuQmCC"
)

ICON_32_B64 = (
    "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"
)


if __name__ == "__main__":
    App().mainloop()
