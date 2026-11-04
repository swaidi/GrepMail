"""Opens the real window (needs a display) and walks through the first-run flow.
Run on Linux:  xvfb-run -a python tests/gui_smoke.py [screenshot_dir]"""
import os
import sys
import tempfile
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import openpyxl
import grepmail as g
import gm_platform
from tkinter import filedialog

shots = sys.argv[1] if len(sys.argv) > 1 else None
tmp = tempfile.mkdtemp()
g.CONFIG_FILE = os.path.join(tmp, "cfg.json")
gm_platform.config_dir = lambda: os.path.join(tmp, "appdata")
app = g.App()
app.update()
if app.winfo_exists():
    pass


def shot(name):
    if not shots:
        return
    from PIL import ImageGrab
    app.update_idletasks(); app.update()
    x, y, w, h = app.winfo_rootx(), app.winfo_rooty(), app.winfo_width(), app.winfo_height()
    ImageGrab.grab(bbox=(x, y, x + w, y + h), xdisplay=os.environ["DISPLAY"]).save(os.path.join(shots, name + ".png"))


def script():
    # the app opened a welcome dialog on its own; close it by answering "exit" would quit,
    # so the test drives the flows directly instead
    for w in app.winfo_children():
        if w.winfo_class() == "Toplevel":
            w.destroy()
    # ---- 1. creating over an existing file -> conflict, retry after removing it
    target = os.path.join(tmp, "GrepMail.db")
    open(target, "wb").write(b"not a database")
    answers = iter(["retry", "retry"])
    seen = []

    def fake_choice(title, message, buttons):
        seen.append((title, message))
        key = next(answers)
        if len(seen) == 2:
            os.remove(target)       # the user removes the file before the 2nd Retry
        return key
    app._choice_dialog = fake_choice
    filedialog.asksaveasfilename = lambda **k: target
    assert app._create_database_flow() is True
    assert len(seen) == 2 and "already exists" in seen[0][1], seen
    assert g.inspect_database(target) == "valid" and app.db is not None
    app.db.export_view()
    assert g.SHEET_CUSTOM in openpyxl.load_workbook(app.db.view_path).sheetnames
    print("conflict flow ok")

    # ---- 2. old-format file opened -> conflict message mentions format
    old = os.path.join(tmp, "old.xlsx")
    wb = openpyxl.Workbook(); wb.active.title = "TenderMap"; wb.save(old)
    seen.clear(); answers = iter(["exit"])
    app._on_close = lambda: seen.append(("closed", ""))
    filedialog.askopenfilename = lambda **k: old
    app.db = None
    app._open_existing_flow()
    assert "current GrepMail format" in seen[0][1] and seen[-1][0] == "closed", seen
    print("old format flow ok")
    app.db = g.ExcelDB(target)

    # ---- 3. screens
    app.geometry("1180x760+0+0")
    app.db.add_or_update_tender("PRJ-001", os.path.join(tmp, "save"), "Tenders\\MEP", "2026-12-20")
    app.db.add_keyword("offer")
    app.db.add_exclusion("clarification")
    app._refresh_all()
    for key in ("tenders", "rules", "settings", "customlog", "offers", "tenderlog", "emaillog", "summary", "keywords"):
        app.show_page(key); app.update(); shot("g_" + key)
    # security tab and mail account tab
    app.show_page("settings"); app.update()
    nb = [w for w in app.pages["settings"].winfo_children() if w.winfo_class() == "TNotebook"][0]
    for i, name in ((1, "mail"), (2, "startup"), (3, "security")):
        nb.select(i); app.update(); shot("g_settings_" + name)
    # scope radio disables folder button
    app.show_page("tenders")
    app.scope_var.set(g.SCOPE_INBOX); app._scope_changed()
    assert app.outlook_folder_var.get() == g.INBOX_TOKEN
    app.update(); shot("g_tenders_inbox_only")
    # ---- 4. File > Import References and Keywords (headerless workbook)
    src = os.path.join(tmp, "teamfile.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Tenders"
    ws.append(["MEP055", os.path.join(tmp, "p55")]); ws.append(["PRJ-001", os.path.join(tmp, "save2")])
    ws.append(["X-9", ""])
    kw = wb.create_sheet("Keywords"); kw.append(["offer"]); kw.append(["pricing"]); kw.append(["عرض"])
    wb.save(src)
    filedialog.askopenfilename = lambda **k: src
    app._import_flow(); app.update()
    dlg = [w for w in app.winfo_children() if w.winfo_class() == "Toplevel"][-1]

    def find(widget, cls, text=None):
        for c in widget.winfo_children():
            if c.winfo_class() == cls and (text is None or str(c.cget("text")) == text):
                return c
            r = find(c, cls, text)
            if r is not None:
                return r
        return None
    import tkinter as tk
    texts = []
    def all_widgets(w):
        for c in w.winfo_children():
            yield c
            yield from all_widgets(c)
    for lbl in [x for x in all_widgets(dlg) if x.winfo_class() == "TLabel"]:
        var = str(lbl.cget("textvariable"))
        texts.append(app.getvar(var) if var else str(lbl.cget("text")))
    assert any("Will import: 2 reference(s), 2 new keyword(s)" in x and "Skipped: 1 row(s)" in x for x in texts), texts
    shot("g_import_preview")
    find(dlg, "TButton", "Import").invoke(); app.update()
    refs = {r[0]: r for r in app.db.list_tenders()}
    assert set(refs) == {"PRJ-001", "MEP055"} and refs["MEP055"][2] == g.INBOX_TOKEN, refs
    assert refs["PRJ-001"][1] == os.path.join(tmp, "save2")        # existing reference updated
    assert app.db.keyword_list_lower() == ["offer", "pricing", "عرض"], app.db.keyword_list_lower()
    print("import flow ok")
    app.destroy()
    print("GUI SMOKE PASSED")


def guarded():
    try:
        script()
    except BaseException:
        import traceback
        traceback.print_exc()
        os._exit(1)


app.after(500, guarded)
app.mainloop()
