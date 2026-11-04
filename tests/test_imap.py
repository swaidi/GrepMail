"""IMAP layer test with a simulated server (no network). Run: python tests/test_imap.py"""
import os
import sys
import tempfile
import datetime as dt
import email.message
import email.utils
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import gm_mail
import gm_platform
import grepmail as g

TMP = tempfile.mkdtemp()
g.CONFIG_FILE = os.path.join(TMP, "cfg.json")
gm_platform.config_dir = lambda: os.path.join(TMP, "appdata")


def make_raw(subject, body, att=None, frm="Supplier Co <sales@supplier.com>"):
    m = email.message.EmailMessage()
    m["From"], m["To"], m["Subject"] = frm, "me@x.com", subject
    m["Date"] = email.utils.format_datetime(dt.datetime(2026, 11, 3, 9, 0).astimezone())
    m.set_content(body)
    if att:
        m.add_attachment(att[1], maintype="application", subtype="pdf", filename=att[0])
    return m.as_bytes()


class FakeIMAP:
    """Just enough of imaplib.IMAP4 for GrepMail."""
    boxes = {
        "INBOX": [], "INBOX/Tenders": [], "INBOX/Tenders/MEP": [], "Sent": [], "[Gmail]/Important": [],
    }
    flags = {"Sent": "\\Sent"}
    logged = None

    def __init__(self, *a, **k): pass
    def starttls(self, **k): pass
    def login(self, u, p):
        if p != "secret":
            raise gm_mail.imaplib.IMAP4.error("bad login")
        FakeIMAP.logged = u
    def logout(self): pass
    def list(self):
        out = []
        for name in self.boxes:
            fl = self.flags.get(name, "")
            out.append(('(%s) "/" "%s"' % (fl, name)).encode())
        return "OK", out
    def select(self, name, readonly=True):
        self.cur = name.strip('"')
        return "OK", [b"1"]
    def response(self, key):
        return "OK", [b"777"]
    def uid(self, cmd, *args):
        box = self.boxes[self.cur]
        if cmd == "SEARCH":
            return "OK", [b" ".join(str(i + 1).encode() for i in range(len(box)))]
        if cmd == "FETCH":
            uids, what = args
            if "INTERNALDATE" in what:
                out = []
                for u in uids.split(","):
                    out.append(('%s (UID %s INTERNALDATE "03-Nov-2026 09:%02d:00 +0000")' % (u, u, int(u))).encode())
                return "OK", out
            u = int(uids)
            return "OK", [(b"%d (UID %d BODY[] {1}" % (u, u), box[u - 1]), b")"]


gm_mail.imaplib.IMAP4_SSL = FakeIMAP
gm_mail.imaplib.IMAP4 = type("IMAP4", (FakeIMAP,), {"error": Exception})
gm_mail.imaplib.IMAP4.error = Exception
gm_mail.HAS_KEYRING = False


def test_connect_and_folders():
    mb = gm_mail.connect_imap({"host": "h", "port": 993, "user": "me@x.com", "security": "ssl"}, "secret")
    assert mb.inbox().Name == "INBOX" and mb.sent().Name == "Sent"
    choices = dict(mb.relative_paths())
    assert "(Inbox)" in choices and "Tenders/MEP" in choices and "/[Gmail]/Important" in choices, choices
    inbox = mb.inbox()
    assert inbox.Folders.Item("Tenders").Folders.Item("MEP").Name == "MEP"
    try:
        gm_mail.connect_imap({"host": "h", "port": 993, "user": "me@x.com"}, "wrong"); raise SystemExit("bad login must fail")
    except gm_mail.MailError as e:
        assert "Could not connect" in str(e)
    print("imap connect/folders ok")


def test_end_to_end_save():
    FakeIMAP.boxes["INBOX/Tenders/MEP"] = [
        make_raw("ABC Co - PRJ-001 Quotation", "Please find our offer.", ("quote.pdf", b"%PDF-1.4 hello")),
        make_raw("ABC Co - PRJ-001 Clarification", "our offer", ("x.pdf", b"%PDF-1.4 other")),
        make_raw("Newsletter", "hi"),
    ]
    FakeIMAP.boxes["Sent"] = [make_raw("RE: ABC Co - PRJ-001 Quotation", "thanks")]
    d = tempfile.mkdtemp(dir=TMP)
    db = g.ExcelDB(os.path.join(d, "t.xlsx"), create=True)
    save = os.path.join(d, "save")
    db.add_or_update_tender("PRJ-001", save, "Tenders/MEP", "")
    db.add_keyword("offer"); db.add_exclusion("clarification")
    cfg = g.Config()
    cfg.provider = "imap"
    cfg.imap = {"host": "h", "port": 993, "user": "me@x.com", "security": "ssl"}
    cfg.session_password = "secret"
    cfg.scan_enabled = False
    conn = g.connect_mail(cfg)
    proc = g.EmailProcessor(db, cfg, lambda *a: None)
    ctx, known = g.Context(db), g.new_known(db)
    import threading
    stop = threading.Event()
    newest = g.scan_inbox_side(conn, dt.datetime(2026, 11, 1), proc, ctx, known, stop)
    g.scan_folder(conn.sent, dt.datetime(2026, 11, 1), "OUTGOING", proc, ctx, known, stop)
    assert os.listdir(save) == ["ABC Co quote.pdf"], os.listdir(save)
    res = sorted(r["Result"] for r in db.list_email_log())
    assert res == ["Excluded", "No Reference", "Saved", "Sent"], res
    with open(os.path.join(save, "ABC Co quote.pdf"), "rb") as f:
        assert f.read() == b"%PDF-1.4 hello"
    assert newest > dt.datetime(2026, 11, 1)
    conn.close()
    print("imap end-to-end ok")


if __name__ == "__main__":
    test_connect_and_folders()
    test_end_to_end_save()
    print("ALL IMAP TESTS PASSED")
