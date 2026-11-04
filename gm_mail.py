"""
Mail access for GrepMail on any system: IMAP.

The messages and folders returned here deliberately look like Outlook's objects
(Subject, Body, Attachments.Item(i).SaveAsFile(...), Folders.Item(name) ...), so
the processing code works the same for Outlook and for IMAP mailboxes.

Supported through IMAP: Gmail (app password), Proton Mail (via the Proton Mail
Bridge desktop app), Microsoft 365 / Outlook.com (only if IMAP is enabled for the
account) and any other IMAP server.
"""
import re
import ssl
import email
import imaplib
import datetime as dt
from email import policy
from html.parser import HTMLParser

PRESETS = {
    "Gmail": dict(host="imap.gmail.com", port=993, security="ssl",
                  note="Use a Google 'app password' (the account needs 2-step verification)."),
    "Proton Mail (Bridge)": dict(host="127.0.0.1", port=1143, security="starttls",
                                 note="Install and sign in to Proton Mail Bridge first, then use the "
                                      "username and password that Bridge shows (not your Proton password)."),
    "Microsoft 365 / Outlook.com": dict(host="outlook.office365.com", port=993, security="ssl",
                                        note="Works only if your organization allows IMAP for your account."),
    "Other IMAP": dict(host="", port=993, security="ssl",
                       note="Enter the IMAP server, port and login given by your mail provider."),
}

KEYRING_SERVICE = "GrepMail"
try:
    import keyring  # optional: stores the password in the system's secure store
    HAS_KEYRING = True
except Exception:  # pragma: no cover
    keyring = None
    HAS_KEYRING = False


class MailError(Exception):
    pass


# ---------------------------------------------------------
# password storage
# ---------------------------------------------------------
def _account_key(user, host):
    return "%s@%s" % (user, host)


def save_password(user, host, password):
    """Returns True when stored in the system's secure store."""
    if not HAS_KEYRING:
        return False
    try:
        keyring.set_password(KEYRING_SERVICE, _account_key(user, host), password)
        return True
    except Exception:
        return False


def load_password(user, host):
    if not HAS_KEYRING:
        return ""
    try:
        return keyring.get_password(KEYRING_SERVICE, _account_key(user, host)) or ""
    except Exception:
        return ""


# ---------------------------------------------------------
# helpers
# ---------------------------------------------------------
def decode_mutf7(s):
    """IMAP folder names use modified UTF-7."""
    def repl(m):
        chunk = m.group(1)
        if chunk == "":
            return "&"
        b64 = chunk.replace(",", "/")
        b64 += "=" * (-len(b64) % 4)
        try:
            import base64
            return base64.b64decode(b64).decode("utf-16-be")
        except Exception:
            return m.group(0)
    return re.sub(r"&([^-]*)-", repl, s)


class _Strip(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts = []
        self._skip = 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self._skip += 1
        if tag in ("br", "p", "div", "tr", "li"):
            self.parts.append("\n")

    def handle_endtag(self, tag):
        if tag in ("script", "style") and self._skip:
            self._skip -= 1

    def handle_data(self, data):
        if not self._skip:
            self.parts.append(data)


def html_to_text(html):
    p = _Strip()
    try:
        p.feed(html)
    except Exception:
        return re.sub(r"<[^>]+>", " ", html)
    return "".join(p.parts)


# ---------------------------------------------------------
# Outlook-like objects
# ---------------------------------------------------------
class _Attachment:
    def __init__(self, name, data):
        self.FileName = name
        self._data = data

    def SaveAsFile(self, path):
        with open(path, "wb") as f:
            f.write(self._data)


class _Attachments:
    def __init__(self, items):
        self._items = items
        self.Count = len(items)

    def Item(self, i):  # 1-based, like Outlook
        return self._items[i - 1]


class ImapMessage:
    Class = 43  # olMail
    SenderEmailType = "SMTP"

    def __init__(self, mailbox, folder, uid, when):
        self._mb, self._folder, self._uid = mailbox, folder, uid
        self._when = when
        self._msg = None
        self._att = None
        self._body = None

    # time, known without downloading the message
    @property
    def ReceivedTime(self):
        return self._when

    @property
    def SentOn(self):
        return self._when

    @property
    def EntryID(self):
        return "imap|%s|%s|%s" % (self._folder.FolderPath, self._mb.uidvalidity(self._folder), self._uid)

    def _load(self):
        if self._msg is None:
            raw = self._mb.fetch_raw(self._folder, self._uid)
            self._msg = email.message_from_bytes(raw, policy=policy.default)
        return self._msg

    @property
    def Subject(self):
        return str(self._load().get("Subject", "") or "")

    @property
    def SenderName(self):
        return email.utils.parseaddr(str(self._load().get("From", "")))[0]

    @property
    def SenderEmailAddress(self):
        return email.utils.parseaddr(str(self._load().get("From", "")))[1]

    @property
    def To(self):
        return str(self._load().get("To", "") or "")

    @property
    def CC(self):
        return str(self._load().get("Cc", "") or "")

    @property
    def Body(self):
        if self._body is None:
            msg = self._load()
            text = None
            try:
                part = msg.get_body(preferencelist=("plain",))
                if part is not None:
                    text = part.get_content()
                else:
                    part = msg.get_body(preferencelist=("html",))
                    if part is not None:
                        text = html_to_text(part.get_content())
            except Exception:
                text = None
            self._body = (text or "").replace("\r\n", "\n")
        return self._body

    @property
    def Attachments(self):
        if self._att is None:
            items = []
            for part in self._load().iter_attachments():
                name = part.get_filename()
                if not name:
                    continue
                try:
                    data = part.get_payload(decode=True) or b""
                except Exception:
                    data = b""
                items.append(_Attachment(str(name), data))
            self._att = _Attachments(items)
        return self._att


class _Items:
    """Iterates a folder's mail newest first. set_since() narrows the server search."""

    def __init__(self, mailbox, folder):
        self._mb, self._folder = mailbox, folder
        self._since = None

    def Sort(self, *a, **k):  # always newest first
        return None

    def set_since(self, when):
        self._since = when

    def __iter__(self):
        for uid, when in self._mb.list_messages(self._folder, self._since):
            yield ImapMessage(self._mb, self._folder, uid, when)


class _FolderList:
    def __init__(self, folders):
        self._folders = folders

    def __iter__(self):
        return iter(self._folders)

    def Item(self, name):
        for f in self._folders:
            if f.Name.casefold() == str(name).casefold():
                return f
        raise KeyError(name)


class ImapFolder:
    def __init__(self, mailbox, raw_name, delim, flags):
        self._mb = mailbox
        self.raw_name = raw_name
        self.delim = delim or "/"
        self.flags = [f.lower() for f in flags]
        self.FolderPath = decode_mutf7(raw_name)
        self.Name = self.FolderPath.split(self.delim)[-1]
        self.EntryID = raw_name
        self.children = []
        self.Items = _Items(mailbox, self)

    @property
    def Folders(self):
        return _FolderList(self.children)


class _Root:
    """Stands for the account's top level (used for folder paths that start with '/')."""

    def __init__(self, mailbox):
        self._mb = mailbox

    @property
    def Folders(self):
        return _FolderList(self._mb.top)


# ---------------------------------------------------------
# the mailbox
# ---------------------------------------------------------
class ImapMailbox:
    def __init__(self, host, port, user, password, security="ssl"):
        self.host, self.port, self.user = host, int(port), user
        self._password, self.security = password, security
        self._conn = None
        self._selected = None
        self._uidv = {}
        self.folders = []        # flat list
        self.top = []            # tree roots
        self.root = _Root(self)

    # ----- connection -----
    def connect(self):
        try:
            local = self.host in ("127.0.0.1", "localhost", "::1")
            ctx = ssl._create_unverified_context() if local else ssl.create_default_context()
            if self.security == "ssl":
                self._conn = imaplib.IMAP4_SSL(self.host, self.port, ssl_context=ctx)
            else:
                self._conn = imaplib.IMAP4(self.host, self.port)
                if self.security == "starttls":
                    self._conn.starttls(ssl_context=ctx)
            self._conn.login(self.user, self._password)
        except Exception as e:
            raise MailError("Could not connect to %s: %s" % (self.host, e))
        self._load_folders()
        return self

    def close(self):
        try:
            if self._conn:
                self._conn.logout()
        except Exception:
            pass
        self._conn = None

    # ----- folders -----
    _LIST_RE = re.compile(r'\((?P<flags>[^)]*)\)\s+(?P<delim>"[^"]*"|NIL)\s+(?P<name>.+)$')

    def _load_folders(self):
        typ, data = self._conn.list()
        if typ != "OK":
            raise MailError("The server did not list its folders.")
        self.folders = []
        for line in data:
            if not line:
                continue
            if isinstance(line, tuple):
                line = line[0]
            text = line.decode("utf-8", "replace") if isinstance(line, bytes) else str(line)
            m = self._LIST_RE.match(text)
            if not m:
                continue
            flags = m.group("flags").split()
            if "\\noselect" in [f.lower() for f in flags]:
                selectable = False
            else:
                selectable = True
            delim = m.group("delim").strip('"') if m.group("delim") != "NIL" else "/"
            name = m.group("name").strip()
            if name.startswith('"') and name.endswith('"'):
                name = name[1:-1].replace('\\"', '"')
            f = ImapFolder(self, name, delim, flags)
            f.selectable = selectable
            self.folders.append(f)
        by_path = {f.raw_name: f for f in self.folders}
        # some servers list a child without its parent: add the missing parents
        for f in list(self.folders):
            name = f.raw_name
            while f.delim and f.delim in name:
                name = name.rsplit(f.delim, 1)[0]
                if name not in by_path:
                    ph = ImapFolder(self, name, f.delim, [])
                    ph.selectable = False
                    by_path[name] = ph
                    self.folders.append(ph)
        self.top = []
        for f in self.folders:
            parent = None
            if f.delim and f.delim in f.raw_name:
                parent = by_path.get(f.raw_name.rsplit(f.delim, 1)[0])
            (parent.children if parent else self.top).append(f)

    def inbox(self):
        for f in self.folders:
            if f.raw_name.upper() == "INBOX":
                return f
        raise MailError("No INBOX folder was found on this account.")

    def sent(self):
        for f in self.folders:
            if "\\sent" in f.flags:
                return f
        names = ("sent", "sent items", "sent mail", "sent messages", "[gmail]/sent mail",
                 "inbox.sent", "inbox/sent")
        for f in self.folders:
            if f.FolderPath.casefold() in names:
                return f
        return None

    def relative_paths(self):
        """[(stored value, display)] for the folder picker, inbox first.
        Folders below the inbox are stored relative to it ('Tenders/MEP');
        folders elsewhere (for example Gmail labels) start with '/'."""
        inbox = self.inbox()
        out = [("(Inbox)", "Inbox")]

        def walk(folder, parts, disp, absolute):
            for c in folder.children:
                p = parts + [c.Name]
                store = ("/" if absolute else "") + "/".join(p)
                out.append((store, " / ".join(disp + [c.Name])))
                walk(c, p, disp + [c.Name], absolute)

        walk(inbox, [], ["Inbox"], False)
        for f in self.top:
            if f is inbox:
                continue
            out.append(("/" + f.Name, f.Name))
            walk(f, [f.Name], [f.Name], True)
        return out

    # ----- messages -----
    def _select(self, folder):
        if self._selected != folder.raw_name:
            q = '"%s"' % folder.raw_name.replace('"', '\\"')
            typ, data = self._conn.select(q, readonly=True)
            if typ != "OK":
                raise MailError("Could not open folder %s" % folder.FolderPath)
            self._selected = folder.raw_name
            try:
                self._uidv[folder.raw_name] = self._conn.response("UIDVALIDITY")[1][0].decode()
            except Exception:
                self._uidv[folder.raw_name] = "0"

    def uidvalidity(self, folder):
        return self._uidv.get(folder.raw_name, "0")

    def list_messages(self, folder, since=None):
        """[(uid, naive local datetime)] newest first."""
        if not getattr(folder, "selectable", True):
            return []
        self._select(folder)
        if since is not None:
            crit = ("SINCE", (since - dt.timedelta(days=1)).strftime("%d-%b-%Y"))
        else:
            crit = ("ALL",)
        typ, data = self._conn.uid("SEARCH", None, *crit)
        if typ != "OK" or not data or not data[0]:
            return []
        uids = data[0].split()
        out = []
        for i in range(0, len(uids), 200):
            chunk = b",".join(uids[i:i + 200]).decode()
            typ, resp = self._conn.uid("FETCH", chunk, "(INTERNALDATE)")
            if typ != "OK":
                continue
            for item in resp:
                if isinstance(item, tuple):
                    item = item[0]
                if not isinstance(item, (bytes, bytearray)):
                    continue
                mu = re.search(rb"UID (\d+)", item)
                tm = imaplib.Internaldate2tuple(item)
                if not mu or tm is None:
                    continue
                import time as _t
                when = dt.datetime.fromtimestamp(_t.mktime(tm))
                out.append((mu.group(1).decode(), when))
        out.sort(key=lambda x: x[1], reverse=True)
        return out

    def fetch_raw(self, folder, uid):
        self._select(folder)
        typ, resp = self._conn.uid("FETCH", str(uid), "(BODY.PEEK[])")
        if typ != "OK":
            raise MailError("Could not read message %s" % uid)
        for item in resp:
            if isinstance(item, tuple):
                return item[1]
        raise MailError("Message %s came back empty" % uid)


def connect_imap(cfg_imap, password):
    mb = ImapMailbox(cfg_imap.get("host", ""), cfg_imap.get("port", 993),
                     cfg_imap.get("user", ""), password, cfg_imap.get("security", "ssl"))
    return mb.connect()
