# SPDX-License-Identifier: AGPL-3.0-or-later
"""Reads References, Keywords and Exclusion words from another file so they can be imported
into GrepMail.db. No GUI and no dependence on grepmail.py.

Supported sources:
    .xlsx / .xlsm  Excel workbooks: GrepMail 1.x / 2.x, the older Tender Manager "Database"
                   files, and workbooks WITHOUT a header row
    .csv           one list per file (references or keywords)
    .db            another GrepMail.db (SQLite)

Without a header row the columns are read by position:
    references:  Reference, Local Path, Mail Folder, Deadline
    keywords:    one keyword per row (first column)
"""
import csv
import os
import sqlite3

SQLITE_MAGIC = b"SQLite format 3\x00"

_REF_COLS = {
    "reference": ("reference", "tenderref", "tenderreference", "ref"),
    "localpath": ("localpath", "savepath", "path", "folderpath"),
    "mailfolder": ("mailfolder", "outlookfolder"),
    "deadline": ("deadline", "duedate"),
    "scope": ("scope",),
}
_HEADER_WORDS = {n for names in _REF_COLS.values() for n in names} | {"keyword", "keywords", "word", "words"}


_SKIP_SHEETS = ("log", "summary", "offer", "unmatched", "option", "dashboard", "email")
_PLACEHOLDERS = {"outlookfolder", "mailfolder", "localpath", "reference", "scope", "deadline"}


class ImportError_(Exception):
    """A problem the user can understand (wrong file type, unreadable file...)."""


def clean(v):
    return "" if v is None else str(v).strip()


def _key(v):
    return clean(v).lower().replace(" ", "").replace("_", "")


def _is_header(row):
    """A header row starts with a column title, or has at least two titles. A data row that
    merely contains a title-like word (e.g. a leftover 'OutlookFolder') is not a header."""
    hits = [_key(c) in _HEADER_WORDS for c in row]
    return bool(hits) and (hits[0] or sum(hits) >= 2)


def _rows_with_data(rows):
    return [list(r) for r in rows if any(clean(c) for c in r)]


def refs_from_rows(rows):
    """Reference rows -> dicts. A header row is used when present, otherwise the column order."""
    rows = _rows_with_data(rows)
    if not rows:
        return []
    idx = {"reference": 0, "localpath": 1, "mailfolder": 2, "deadline": 3, "scope": None}
    if _is_header(rows[0]):
        head = [_key(h) for h in rows[0]]
        rows = rows[1:]
        for field, names in _REF_COLS.items():
            idx[field] = next((head.index(n) for n in names if n in head), None)
        if idx["reference"] is None:
            raise ImportError_("No Reference column was found in the header row.")
    out = []
    for r in rows:
        def cell(field):
            i = idx[field]
            return r[i] if i is not None and i < len(r) else ""
        if clean(cell("reference")):
            folder = clean(cell("mailfolder"))
            if _key(folder) in _PLACEHOLDERS:          # a leftover column title typed as a value
                folder = ""
            out.append({"Reference": clean(cell("reference")), "LocalPath": clean(cell("localpath")),
                        "MailFolder": folder, "Deadline": cell("deadline"),
                        "Scope": clean(cell("scope"))})
    return out


def words_from_rows(rows):
    rows = _rows_with_data(rows)
    if rows and _is_header(rows[0]):
        rows = rows[1:]
    return [clean(r[0]) for r in rows if clean(r[0])]


def _dedupe_refs(refs):
    seen, out = {}, []
    for r in refs:                       # the last occurrence of a reference wins
        k = r["Reference"].upper()
        if k in seen:
            out[seen[k]] = r
        else:
            seen[k] = len(out)
            out.append(r)
    return out


def _dedupe_words(words):
    seen, out = set(), []
    for w in words:
        if w.lower() not in seen:
            seen.add(w.lower())
            out.append(w)
    return out


def _result(refs, kws, excl):
    return {"references": _dedupe_refs(refs), "keywords": _dedupe_words(kws),
            "exclusions": _dedupe_words(excl)}


def read_sqlite(path):
    con = sqlite3.connect("file:%s?mode=ro" % os.path.abspath(path).replace("\\", "/"), uri=True)
    con.row_factory = sqlite3.Row
    try:
        def table(name):
            try:
                return [dict(r) for r in con.execute('SELECT * FROM "%s"' % name)]
            except sqlite3.Error:
                return []
        refs = [{"Reference": clean(r.get("Reference")), "LocalPath": clean(r.get("LocalPath")),
                 "MailFolder": clean(r.get("MailFolder")), "Deadline": r.get("Deadline"),
                 "Scope": clean(r.get("Scope"))} for r in table("References") if clean(r.get("Reference"))]
        kws = [clean(r.get("Keyword")) for r in table("Keywords") if clean(r.get("Keyword"))]
        excl = [clean(r.get("Word")) for r in table("Exclusions") if clean(r.get("Word"))]
    finally:
        con.close()
    return _result(refs, kws, excl)


def read_excel(path):
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as e:
        raise ImportError_("The Excel file could not be opened (%s). Save it as .xlsx and try again." % e)
    refs, kws, excl, unknown = [], [], [], []
    try:
        for ws in wb.worksheets:
            name = ws.title.lower()
            rows = _rows_with_data(ws.iter_rows(values_only=True))
            if not rows:
                continue
            if any(w in name for w in _SKIP_SHEETS):
                continue                     # logs, summaries and registers are not setup data
            if "keyword" in name:
                kws += words_from_rows(rows)
            elif "exclu" in name:
                excl += words_from_rows(rows)
            elif any(w in name for w in ("tender", "reference", "map")):
                refs += refs_from_rows(rows)
            else:
                unknown.append(rows)
        for rows in unknown:                 # sheets with other names: decide by their shape
            width = max(sum(1 for c in r if clean(c)) for r in rows)
            if width >= 2 and not refs:
                refs += refs_from_rows(rows)
            elif width == 1 and not kws:
                kws += words_from_rows(rows)
    finally:
        wb.close()
    return _result(refs, kws, excl)


def read_csv(path, kind="references"):
    for enc in ("utf-8-sig", "cp1256", "latin-1"):
        try:
            with open(path, newline="", encoding=enc) as f:
                rows = list(csv.reader(f))
            break
        except UnicodeDecodeError:
            continue
    if kind == "keywords":
        return _result([], words_from_rows(rows), [])
    return _result(refs_from_rows(rows), [], [])


def guess_csv_kind(path):
    return "keywords" if "keyword" in os.path.basename(path).lower() else "references"


def read_source(path, csv_kind=None):
    """Reads any supported file. csv_kind: 'references' or 'keywords' (CSV only)."""
    if not os.path.isfile(path):
        raise ImportError_("The file was not found.")
    ext = os.path.splitext(path)[1].lower()
    with open(path, "rb") as f:
        head = f.read(16)
    if head == SQLITE_MAGIC:
        return read_sqlite(path)
    if ext in (".xlsx", ".xlsm"):
        return read_excel(path)
    if ext == ".xls":
        raise ImportError_("Old .xls files are not supported. Open it in Excel, save it as .xlsx and try again.")
    if ext == ".csv":
        return read_csv(path, csv_kind or guess_csv_kind(path))
    raise ImportError_("This file type is not supported. Choose an Excel (.xlsx), CSV or GrepMail.db file.")
