# SPDX-License-Identifier: AGPL-3.0-or-later
import os
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import openpyxl
import gm_import as gi
import grepmail as g

TMP = tempfile.mkdtemp()


def xlsx(name, sheets):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for title, rows in sheets.items():
        ws = wb.create_sheet(title)
        for r in rows:
            ws.append(r)
    path = os.path.join(TMP, name)
    wb.save(path)
    return path


def test_headerless_workbook():
    p = xlsx("a.xlsx", {"Tenders": [["MEP055", "D:/x/a"], ["MEP055", "D:/x/a2"], ["STR081", "D:/x/b", "Tenders\\S", "2026-12-01"]],
                        "Keywords": [["offer"], ["Offer"], ["عرض"]]})
    d = gi.read_source(p)
    assert [r["Reference"] for r in d["references"]] == ["MEP055", "STR081"]
    assert d["references"][0]["LocalPath"] == "D:/x/a2"          # last duplicate wins
    assert d["references"][1]["MailFolder"] == "Tenders\\S"
    assert d["keywords"] == ["offer", "عرض"]
    print("headerless workbook ok")


def test_headers_and_other_sheet_names():
    p = xlsx("b.xlsx", {"TenderMap": [["TenderRef", "LocalPath", "OutlookFolder", "Deadline"], ["MEP-1", "D:/p", "A\\B", "2026-01-02"]],
                        "Keywords": [["Keyword"], ["quot"]]})
    d = gi.read_source(p)
    assert d["references"][0]["MailFolder"] == "A\\B" and d["keywords"] == ["quot"]
    p = xlsx("c.xlsx", {"Sheet1": [["Tender Reference", "Path"], ["X-1", "D:/q"]], "Words": [["bid"], ["price"]]})
    d = gi.read_source(p)
    assert d["references"][0]["Reference"] == "X-1" and d["keywords"] == ["bid", "price"]
    print("headers and unknown sheet names ok")


def test_mixed_sheet_names():
    p = xlsx("mixed1.xlsx", {"Sheet1": [["MEP-9", "D:/z"], ["MEP-8", "D:/y"]], "Keywords": [["offer"]]})
    d = gi.read_source(p)
    assert len(d["references"]) == 2 and d["keywords"] == ["offer"], d
    p = xlsx("mixed2.xlsx", {"Data": [["MEP-9", "D:/z"]], "Words": [["offer"], ["bid"]]})
    d = gi.read_source(p)
    assert len(d["references"]) == 1 and d["keywords"] == ["offer", "bid"], d
    print("mixed sheet names ok")


def test_team_database_layout():
    p = xlsx("team1.xlsx", {
        "Database": [["something", "else"], ["a", "b"]],
        "Keywords": [["Keyword"], ["offer"], ["quotation"]],
        "TenderMap": [["Reference", "LocalPath", "MailFolder", "Scope", "Deadline"],
                      ["mep085", "D:/a/Attachment", "OutlookFolder", None, None],
                      ["MEP057", "D:\\a\\b", "HMG_Tender\\Riyadh\\MEP057", None, None],
                      ["MEP085", "D:/a/Dammam", "HMG_Tender\\Dammam\\MEP_085", None, None]],
        "EmailLog": [["Timestamp", "Subject"], ["2026-01-01", "x"]],
        "TenderLog": [["Timestamp", "TenderRef", "Result"], ["2026-01-01", "MEP057", "Saved"]]})
    d = gi.read_source(p)
    refs = {r["Reference"].upper(): r for r in d["references"]}
    assert set(refs) == {"MEP085", "MEP057"}, refs
    assert refs["MEP085"]["LocalPath"] == "D:/a/Dammam" and refs["MEP085"]["MailFolder"].endswith("MEP_085")
    assert d["keywords"] == ["offer", "quotation"]
    d2 = gi.read_source(xlsx("team2.xlsx", {"TenderMap": [["mep1", "D:/q", "OutlookFolder"]]}))
    assert d2["references"][0]["MailFolder"] == ""
    print("team database layout ok")


def test_csv_and_errors():
    p = os.path.join(TMP, "tender_map.csv")
    open(p, "w", encoding="utf-8").write("ref,localpath,outlookfolder\nA1,D:/a,Inbox\\X\n")
    assert gi.read_source(p)["references"][0]["Reference"] == "A1"
    k = os.path.join(TMP, "offers_keywords.csv")
    open(k, "w", encoding="utf-8").write("offer\nquote\n")
    assert gi.read_source(k)["keywords"] == ["offer", "quote"]
    assert gi.read_source(k, "references")["references"][0]["Reference"] == "offer"
    for bad, msg in (("x.xls", ".xlsx"), ("x.txt", "not supported")):
        bp = os.path.join(TMP, bad)
        open(bp, "w").write("data")
        try:
            gi.read_source(bp)
            raise AssertionError("expected error")
        except gi.ImportError_ as e:
            assert msg in str(e)
    try:
        gi.read_source(os.path.join(TMP, "missing.xlsx"))
        raise AssertionError("expected error")
    except gi.ImportError_:
        pass
    print("csv and errors ok")


def test_sqlite_source_and_plan():
    src = g.GrepDB(os.path.join(TMP, "s.db"), create=True)
    src.add_or_update_tender("R-1", "D:/r", "In\\Sub", "2027-02-03", g.SCOPE_SUB_INBOX)
    src.add_keyword("tender"); src.add_exclusion("clarification"); src.close()
    d = gi.read_source(os.path.join(TMP, "s.db"))
    assert d["references"][0]["Scope"] == g.SCOPE_SUB_INBOX and d["exclusions"] == ["clarification"]
    plan = g.plan_reference_import
    assert plan({"Reference": "A", "LocalPath": "D:/a"})[2:5] == (g.INBOX_TOKEN, "", g.SCOPE_INBOX)
    assert plan({"Reference": "A", "LocalPath": "D:/a", "MailFolder": "X\\Y"})[4] == g.SCOPE_SUB
    assert plan({"Reference": "A", "LocalPath": "D:/a", "MailFolder": "X", "Deadline": "31/12/2026"})[3] == "2026-12-31"
    assert plan({"Reference": "A", "LocalPath": "D:/a", "MailFolder": "X", "Deadline": "soon"})[5] == "Bad deadline"
    assert plan({"Reference": "", "LocalPath": "D:/a"})[5] and plan({"Reference": "A", "LocalPath": ""})[5]
    print("sqlite source and plan ok")


def test_forced_export_is_quick():
    import time
    db = g.GrepDB(os.path.join(TMP, "e.db"), create=True)
    db.add_keyword("offer")                      # starts the normal 30 s timer
    assert db._export_timer is not None
    db.request_export(0.5, force=True)
    time.sleep(2.0)
    assert db._export_timer is None and os.path.exists(db.view_path)
    names = openpyxl.load_workbook(db.view_path).sheetnames
    assert "Keywords" in names
    db.close()
    print("forced export ok")


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
    print("ALL IMPORT TESTS PASSED")
