"""Core behavior tests (no Outlook, no real mail server). Run: python tests/test_core.py"""
import os
import json
import sys
import shutil
import tempfile
import datetime as dt
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import openpyxl
import grepmail as g
import gm_security

TMP = tempfile.mkdtemp()
g.CONFIG_FILE = os.path.join(TMP, "cfg.json")
import gm_platform
gm_platform.config_dir = lambda: os.path.join(TMP, "appdata")


# ---------------- fake mail objects ----------------
class Att:
    def __init__(self, name, data=b"%PDF-1.4 test"):
        self.FileName, self._d = name, data

    def SaveAsFile(self, path):
        with open(path, "wb") as f:
            f.write(self._d)


class Atts:
    def __init__(self, items):
        self._i = items
        self.Count = len(items)

    def Item(self, n):
        return self._i[n - 1]


class Mail:
    Class = 43

    def __init__(self, subject, body, atts=(), sender="a@x.com", when=None, eid="E1"):
        self.Subject, self.Body, self.SenderName = subject, body, "Sender"
        self.SenderEmailAddress, self.SenderEmailType = sender, "SMTP"
        self.To, self.CC, self.EntryID = "me@x.com", "", eid
        self.ReceivedTime = self.SentOn = when or dt.datetime(2026, 11, 3, 10, 0, 0)
        self.Attachments = Atts(list(atts))


class FakeScanner:
    name = "Fake"

    def available(self): return True

    def scan(self, path):
        data = open(path, "rb").read()
        return (gm_security.THREAT, "EICAR-Test") if b"EICAR" in data else (gm_security.CLEAN, "")

    def status(self): return {}


def make_env(**cfgkw):
    d = tempfile.mkdtemp(dir=TMP)
    db = g.ExcelDB(os.path.join(d, "t.db"), create=True)
    save = os.path.join(d, "save")
    db.add_or_update_tender("PRJ-001", save, "Tenders/MEP", "")
    db.add_keyword("offer")
    db.add_keyword("quotation")
    db.add_exclusion("clarification")
    cfg = g.Config()
    for k, v in cfgkw.items():
        setattr(cfg, k, v)
    msgs = []
    proc = g.EmailProcessor(db, cfg, lambda kind, *a: msgs.append((kind, a)))
    proc._scanner = FakeScanner()
    return db, cfg, proc, save, msgs


def run(proc, db, mail, direction="INCOMING"):
    ctx = g.Context(db)
    known = g.new_known(db)
    proc.process(mail, direction, ctx, known, False, None, False, "Inbox")
    return known


def results(db):
    return [r["Result"] for r in db.list_email_log()]


# ---------------- tests ----------------
def test_helpers():
    assert g.extract_name("RE: FW: ABC Co - PRJ-001 Quote") == "ABC Co"
    assert g.extract_name("RE: FW: ABC Co - PRJ-001 Quote", ["-"], "PRJ-001") == "ABC Co"
    assert g.extract_name("ABC Co || PRJ-001", ["||"]) == "ABC Co"
    assert g.extract_name("ABC Co ## PRJ-001", " ## ") == "ABC Co"
    assert g.build_file_name("{Sender} {Attachment Name}", "ABC", "PRJ-001", "2026-11-03", "Q.pdf") == "ABC Q.pdf"
    assert g.build_file_name("{Date} {Reference} {Sender}", "ABC", "PRJ-001", "2026-11-03", "Q.pdf") == "2026-11-03 PRJ-001 ABC.pdf"
    assert g.build_file_name("", "", "R", "d", "Q.pdf") == "Q.pdf"
    assert g.matched_exclusion("RE: Clarification on price", ["clarification"]) == "clarification"
    assert g.matched_exclusion("Offer", ["clarification"]) == ""
    print("helpers ok")


def test_flow_saved():
    db, cfg, proc, save, _ = make_env()
    run(proc, db, Mail("ABC Co - PRJ-001 Quote", "Please see our offer.", [Att("q.pdf")]))
    files = os.listdir(save)
    assert files == ["ABC Co q.pdf"], files
    assert results(db) == ["Saved"]
    saved = db.list_offers()[0]
    assert saved["Name"] == "ABC Co" and saved["FileNo"] == 1 and saved["Security"] == "Clean"
    assert db.list_tender_log()[0]["Action"] == "Saved"
    assert db.list_custom_log()[0]["Result"] == "Saved"
    print("saved flow ok")


def test_exclusion_beats_reference_and_keyword():
    db, cfg, proc, save, _ = make_env()
    run(proc, db, Mail("ABC Co - PRJ-001 Clarification on price", "our offer attached", [Att("q.pdf")]))
    assert not os.path.exists(save) or os.listdir(save) == []
    assert results(db) == ["Excluded"]
    log = db.list_tender_log()[0]
    assert log["Action"] == "Excluded" and "clarification" in log["Details"]
    print("exclusion ok")


def test_not_relevant_and_unmatched():
    db, cfg, proc, save, _ = make_env()
    run(proc, db, Mail("ABC Co - PRJ-001 hello", "no magic words", [Att("q.pdf")], eid="E2"))
    assert results(db) == ["Not Relevant"]
    assert db.list_tender_log()[0]["Action"] == "Not Relevant"
    cfg.hint_phrases = ["invitation to tender"]
    run(proc, db, Mail("Invitation to tender for X", "body", eid="E3"))
    assert "Unmatched" in results(db) and len(db.list_unmatched()) == 1
    print("not relevant / unmatched ok")


def test_quoted_keyword_ignored():
    db, cfg, proc, save, _ = make_env()
    body = "Thanks.\n\nFrom: Boss\nSent: yesterday\nplease send your offer"
    run(proc, db, Mail("ABC - PRJ-001 re", body, [Att("q.pdf")], eid="E4"))
    assert results(db) == ["Not Relevant"]
    print("quoted text ignored ok")


def test_threat_not_saved():
    db, cfg, proc, save, msgs = make_env()
    run(proc, db, Mail("ABC - PRJ-001 offer", "offer", [Att("bad.pdf", b"EICAR payload")], eid="E5"))
    assert not os.path.exists(save) or os.listdir(save) == []
    assert results(db) == ["Threat Found"]
    assert db.list_tender_log()[0]["Action"] == "Threat Found"
    assert any(k == "toast" and "Threat" in a[0] for k, a in msgs)
    assert db.list_offers() == []
    print("threat blocked ok")


def test_macro_modes():
    orig = gm_security.inspect_macros
    try:
        gm_security.inspect_macros = lambda p: (gm_security.MACRO_PRESENT, "runs automatically")
        for mode, expect_saved in (("allow", True), ("flag", True), ("block", False)):
            db, cfg, proc, save, _ = make_env(macro_mode=mode)
            run(proc, db, Mail("A - PRJ-001 offer", "offer", [Att("m.xlsm", b"data")], eid="M" + mode))
            saved = os.path.isdir(save) and os.listdir(save)
            assert bool(saved) == expect_saved, mode
            if mode == "flag":
                assert "Macro-Enabled" in db.list_offers()[0]["Security"]
            if mode == "block":
                assert results(db) == ["Blocked"]
        gm_security.inspect_macros = lambda p: (gm_security.MACRO_SUSPICIOUS, "Shell, AutoOpen")
        db, cfg, proc, save, _ = make_env()
        run(proc, db, Mail("A - PRJ-001 offer", "offer", [Att("m.xlsm", b"data")], eid="S1"))
        assert results(db) == ["Quarantined"] and (not os.path.exists(save) or not os.listdir(save))
        q = os.path.join(os.path.dirname(db.path), "GrepMail Quarantine")
        assert any(f.endswith("m.xlsm") for f in os.listdir(q))
    finally:
        gm_security.inspect_macros = orig
    print("macro modes ok")


def test_require_scan_without_scanner():
    db, cfg, proc, save, _ = make_env(require_scan=True)

    class NoScan(FakeScanner):
        name = "None"
        def available(self): return False
    proc._scanner = NoScan()
    run(proc, db, Mail("A - PRJ-001 offer", "offer", [Att("q.pdf")], eid="N1"))
    assert results(db) == ["Quarantined"]
    db2, cfg2, proc2, save2, _ = make_env(require_scan=False)
    proc2._scanner = NoScan()
    run(proc2, db2, Mail("A - PRJ-001 offer", "offer", [Att("q.pdf")], eid="N2"))
    assert results(db2) == ["Saved"] and db2.list_offers()[0]["Security"] == "Not Scanned"
    print("scanner availability ok")


def test_separators():
    x = g.extract_name
    # spaces around the separator are ignored, the reference may contain a hyphen
    for subj in ("ABC Co - PRJ-001 Q", "ABC Co-PRJ-001 Q", "ABC Co  -  PRJ-001 Q", "ABC Co -PRJ-001 Q"):
        assert x(subj, ["-"], "PRJ-001") == "ABC Co", subj
    # a hyphen inside the company name is kept when the reference is known
    assert x("Al-Habib Co - PRJ-001 Q", ["-"], "PRJ-001") == "Al-Habib Co"
    # other separators only when enabled
    assert x("ABC Co||PRJ-001", ["-", "||"], "PRJ-001") == "ABC Co"
    assert x("ABC Co / PRJ-001", ["/"], "PRJ-001") == "ABC Co"
    assert x("ABC Co | PRJ-001", ["|", "||"], "PRJ-001") == "ABC Co"
    assert x("ABC Co ## PRJ-001", ["-", "##"], "PRJ-001") == "ABC Co"
    # no known reference: first enabled separator
    assert x("ABC Co || Offer", ["||"]) == "ABC Co"
    assert x("ABC Co _ Offer", ["-", "_"]) == "ABC Co"
    assert x("ABC Co Offer", ["-"]) == "ABC Co Offer"
    # case-insensitive reference
    assert x("ABC Co - prj-001 Q", ["-"], "PRJ-001") == "ABC Co"
    assert g.normalize_separators([" || ", "|", "", "|"]) == ["||", "|"]
    c = g.Config()
    c.separators, c.custom_separator = ["-", "/"], " ## "
    assert c.active_separators() == ["##", "-", "/"]
    print("separators ok")


def test_old_separator_setting_upgraded():
    d = tempfile.mkdtemp(dir=TMP)
    path = os.path.join(d, "c.json")
    for old, seps, custom in ((" || ", ["||"], ""), (" - ", ["-"], ""), (" ## ", [], "##")):
        json.dump({"name_separator": old}, open(path, "w"))
        saved, g.CONFIG_FILE = g.CONFIG_FILE, path
        try:
            c = g.Config()
        finally:
            g.CONFIG_FILE = saved
        assert c.separators == (seps or ["-"]) and c.custom_separator == custom, (old, c.separators)
    print("old separator upgrade ok")


def test_file_format_setting():
    db, cfg, proc, save, _ = make_env(file_format="{Date} {Reference} {Sender} {Attachment Name}", separators=["-", "||"])
    run(proc, db, Mail("ABC Co || PRJ-001 offer", "offer", [Att("q.pdf")], eid="F1"))
    assert os.listdir(save) == ["2026-11-03 PRJ-001 ABC Co q.pdf"], os.listdir(save)
    print("file format ok")


def test_idempotent_rescan():
    db, cfg, proc, save, _ = make_env()
    m = Mail("A - PRJ-001 offer", "offer", [Att("q.pdf")], eid="R1")
    run(proc, db, m)
    run(proc, db, m)
    assert len(os.listdir(save)) == 1 and len(db.list_email_log()) == 1
    print("rescan idempotent ok")


def test_custom_log_filters():
    db, cfg, proc, save, _ = make_env(custom_fields=["Subject", "Result"],
                                      custom_filter={"direction": "Incoming", "only_reference": True,
                                                     "only_attachments": False, "contains": ""})
    db.set_custom_fields(cfg.custom_fields)
    run(proc, db, Mail("A - PRJ-001 offer", "offer", [Att("q.pdf")], eid="C1"))
    run(proc, db, Mail("random mail", "x", eid="C2"))
    run(proc, db, Mail("A - PRJ-001 out", "x", eid="C3"), "OUTGOING")
    rows = db.list_custom_log()
    assert [r["Subject"] for r in rows] == ["A - PRJ-001 offer"], rows
    db.export_view()
    wb = openpyxl.load_workbook(db.view_path)
    heads = [c.value for c in wb[g.SHEET_CUSTOM][1]]
    assert heads[:6] == g.DEFAULT_CUSTOM_FIELDS and "Subject" in heads
    db.set_custom_fields(["Subject", "Name", "MailFolder"])
    db.export_view()
    heads2 = [c.value for c in openpyxl.load_workbook(db.view_path)[g.SHEET_CUSTOM][1]]
    assert heads2[:len(heads)] == heads and heads2[-2:] == ["Name", "MailFolder"], heads2
    print("custom log ok")


def test_old_format_tokens_upgrade():
    assert g.build_file_name("{Name} {Original}", "ABC", "PRJ-001", "2026-11-03", "Q.pdf") == "ABC Q.pdf"
    assert g.upgrade_format("{Name} {Reference} {Original}") == "{Sender} {Reference} {Attachment Name}"
    print("old tokens upgrade ok")


def test_database_states():
    d = tempfile.mkdtemp(dir=TMP)
    p = os.path.join(d, "new.db")
    assert g.inspect_database(p) == "missing"
    db = g.GrepDB(p, create=True)
    assert g.inspect_database(p) == "valid"
    try:
        g.GrepDB(p, create=True); raise SystemExit("must refuse existing file")
    except g.SchemaError as e:
        assert e.kind == "exists"
    # an Excel file of the same name (older versions) blocks creation and is never touched
    d2 = tempfile.mkdtemp(dir=TMP)
    p2 = os.path.join(d2, "GrepMail.db")
    wb = openpyxl.Workbook(); wb.active.title = "TenderMap"; wb.save(os.path.join(d2, "GrepMail.xlsx"))
    assert g.inspect_database(p2) == "view"
    try:
        g.GrepDB(p2, create=True); raise SystemExit("must refuse when the view file exists")
    except g.SchemaError as e:
        assert e.kind == "exists"
    assert openpyxl.load_workbook(os.path.join(d2, "GrepMail.xlsx")).sheetnames == ["TenderMap"]
    # a non-database file is 'old'
    assert g.inspect_database(os.path.join(d2, "GrepMail.xlsx")) == "old"
    try:
        g.GrepDB(os.path.join(d2, "GrepMail.xlsx")); raise SystemExit("must refuse old file")
    except g.SchemaError as e:
        assert e.kind == "old"
    # reopen keeps data; a missing table is recreated
    db.add_keyword("x"); db.close()
    import sqlite3
    con = sqlite3.connect(p); con.execute('DROP TABLE "%s"' % g.SHEET_EXCLUSIONS); con.commit(); con.close()
    db2 = g.GrepDB(p)
    assert db2.list_keywords() == ["x"] and db2.list_exclusions() == []
    db2.close()
    print("database states ok")


def test_excel_view():
    d = tempfile.mkdtemp(dir=TMP)
    p = os.path.join(d, "GrepMail.db")
    db = g.GrepDB(p, create=True)
    db.add_or_update_tender("PRJ-1", os.path.join(d, "x"), "Tenders", "2026-12-20")
    db.add_keyword("offer")
    db.add_exclusion("clarification")
    db.export_view()
    wb = openpyxl.load_workbook(db.view_path)
    assert set(g.SCHEMA) <= set(wb.sheetnames)
    assert [c.value for c in wb[g.SHEET_TENDERS][2]][:3] == ["PRJ-1", os.path.join(d, "x"), "Tenders"]
    assert wb[g.SHEET_KEYWORDS]["A2"].value == "offer"
    # the database keeps working while the view file is 'open' (simulated by a failing replace)
    real = os.replace
    def locked(a, b):
        raise PermissionError("in use")
    os.replace = locked
    try:
        try:
            db.export_view(); raise SystemExit("must report the locked view")
        except PermissionError:
            pass
        db.add_keyword("confirm")      # unaffected
        assert db.keyword_list_lower() == ["offer", "confirm"]
        assert not os.path.exists(db.view_path + ".tmp")
    finally:
        os.replace = real
    db.export_view()
    assert openpyxl.load_workbook(db.view_path)[g.SHEET_KEYWORDS]["A3"].value == "confirm"
    db.close()
    print("excel view ok")


class F:
    def __init__(self, name, subs=()):
        self.Name, self.EntryID = name, "id-" + name
        self._s = {x.Name: x for x in subs}
        self.Folders = self

    def Item(self, n): return self._s[n]
    def __iter__(self): return iter(self._s.values())


def test_folder_scope():
    inbox = F("Inbox", [F("Tenders", [F("MEP", [F("Archive")]), F("Civil")]), F("Projects")])
    conn = g.MailConnection(inbox, None)
    db = g.ExcelDB(os.path.join(tempfile.mkdtemp(dir=TMP), "s.db"), create=True)
    db.add_or_update_tender("A", "x", "Tenders/MEP", "", g.SCOPE_SUB)
    db.add_or_update_tender("B", "x", "Tenders/MEP", "", g.SCOPE_SUB_INBOX)
    db.add_or_update_tender("C", "x", "", "", g.SCOPE_INBOX)
    db.add_or_update_tender("D", "x", g.INBOX_TOKEN, "", g.SCOPE_SUB)

    class P:
        def __init__(s): s.db = db; s.m = []
        def report_once(s, k, m): s.m.append(m)
    res = g.resolve_tender_folders(conn, P())
    by_name = {f.Name: sorted(r) for f, r in res.values()}
    assert by_name["MEP"] == ["A", "B", "D"] or by_name["MEP"] == ["A", "B", "D"], by_name
    assert by_name["Archive"] == ["A", "B", "D"]
    assert by_name["Civil"] == ["D"] and by_name["Tenders"] == ["D"] and by_name["Projects"] == ["D"]
    assert by_name["Inbox"] == ["B", "C"], by_name   # Inbox only + Subfolders and Inbox; D excludes the root
    print("folder scope ok")


def test_macro_detection_without_oletools():
    import zipfile
    p = os.path.join(TMP, "m.docm")
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("word/vbaProject.bin", b"x")
        z.writestr("[Content_Types].xml", "<x/>")
    assert gm_security.has_macro_container(p)
    q = os.path.join(TMP, "n.docx")
    with zipfile.ZipFile(q, "w") as z:
        z.writestr("[Content_Types].xml", "<x/>")
    assert not gm_security.has_macro_container(q)
    assert gm_security.inspect_macros(q)[0] == gm_security.MACRO_NONE
    print("macro container detection ok")


def test_oletools_classification_mocked():
    class FakeParser:
        def __init__(self, path): pass
        def detect_vba_macros(self): return True
        def analyze_macros(self):
            return [("AutoExec", "Workbook_Open", ""), ("Suspicious", "Shell", "")]
        def close(self): pass
    orig, flag = gm_security.VBA_Parser, gm_security.HAS_OLETOOLS
    gm_security.VBA_Parser, gm_security.HAS_OLETOOLS = FakeParser, True
    try:
        p = os.path.join(TMP, "z.xlsm")
        open(p, "wb").write(b"PK")
        st, d = gm_security.inspect_macros(p)
        assert st == gm_security.MACRO_SUSPICIOUS and "Shell" in d
        FakeParser.analyze_macros = lambda self: [("AutoExec", "Workbook_Open", "")]
        assert gm_security.inspect_macros(p)[0] == gm_security.MACRO_PRESENT
        FakeParser.analyze_macros = lambda self: []
        assert gm_security.inspect_macros(p)[0] == gm_security.MACRO_PRESENT
    finally:
        gm_security.VBA_Parser, gm_security.HAS_OLETOOLS = orig, flag
    print("oletools classification ok")


def test_autostart_linux():
    if not gm_platform.IS_LINUX:
        return
    os.environ["XDG_CONFIG_HOME"] = os.path.join(TMP, "xdg")
    assert not gm_platform.autostart_enabled()
    ok, _ = gm_platform.set_autostart(True)
    assert ok and gm_platform.autostart_enabled()
    txt = open(os.path.join(TMP, "xdg", "autostart", "grepmail.desktop")).read()
    assert "Exec=" in txt
    gm_platform.set_autostart(False)
    assert not gm_platform.autostart_enabled()
    print("autostart ok")


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
    print("ALL CORE TESTS PASSED")

