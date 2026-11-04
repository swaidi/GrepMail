# SPDX-License-Identifier: AGPL-3.0-or-later
"""Command-line version of File > Import References and Keywords.

Usage (run from the project folder, with GrepMail closed):
    python tools/import_data.py <source> <GrepMail.db> [--exclusions] [--dry-run] [--keywords]

<source> is an Excel workbook (.xlsx, with or without a header row), a CSV file
(--keywords when it holds keywords) or another GrepMail.db. See gm_import.py for the layout.
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import grepmail as g  # noqa: E402
import gm_import  # noqa: E402


def main(argv):
    args = [a for a in argv if not a.startswith("--")]
    if len(args) != 2:
        print(__doc__)
        return 2
    src, dst = args
    if g.inspect_database(dst) != "valid":
        print("Destination is not a valid GrepMail.db: %s" % dst)
        return 1
    if os.path.abspath(src) == os.path.abspath(dst):
        print("Source and destination are the same file.")
        return 1
    try:
        data = gm_import.read_source(src, "keywords" if "--keywords" in argv else
                                     ("references" if src.lower().endswith(".csv") else None))
    except gm_import.ImportError_ as e:
        print(e)
        return 1
    plans = [g.plan_reference_import(r) for r in data["references"]]
    if "--dry-run" in argv:
        print("DRY RUN - nothing is written.")
        for ref, lp, folder, dl, scope, err in plans:
            print("  Reference=%r | Local Path=%r | Folder=%r | Scope=%s | Deadline=%r %s" % (
                ref, lp, folder, scope, dl, ("<< SKIPPED: " + err) if err else ""))
        for k in data["keywords"]:
            print("  Keyword=%r" % k)
        return 0
    db = g.GrepDB(dst)
    n_r = n_k = n_e = 0
    try:
        for ref, lp, folder, dl, scope, err in plans:
            if err:
                print("Skipped %s: %s" % (ref or "(blank)", err))
                continue
            db.add_or_update_tender(ref, lp, folder, dl, scope)
            n_r += 1
        have = set(db.keyword_list_lower())
        for k in data["keywords"]:
            if k.lower() not in have:
                db.add_keyword(k)
                n_k += 1
        if "--exclusions" in argv:
            have = set(db.exclusion_list_lower())
            for w in data["exclusions"]:
                if w.lower() not in have:
                    db.add_exclusion(w)
                    n_e += 1
    finally:
        db.close()
    print("References imported/updated: %d" % n_r)
    print("New keywords added: %d" % n_k)
    if "--exclusions" in argv:
        print("New exclusion words added: %d" % n_e)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
