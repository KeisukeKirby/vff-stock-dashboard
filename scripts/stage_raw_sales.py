"""Stage the raw sales workbooks under data/raw/ with the upload names the sales ETL expects.

The sales-1st-half-26-dashboard ETL scripts reference their inputs as
'<8-hex>-<Name_with_underscores>.xlsx' (Claude upload names). The same files
live on this PC under Downloads with their original names, so this script
matches each referenced name to a local file (case/punctuation-insensitive)
and copies it into data/raw/ under the referenced name. data/raw/ is
git-ignored: the workbooks are confidential.

Usage: python scripts/stage_raw_sales.py [--dry-run]
"""
import os, re, sys, glob, shutil, difflib

HERE = os.path.dirname(os.path.abspath(__file__))
ETL_DIR = os.path.join(HERE, "sales_etl")
RAW = os.path.join(HERE, "..", "data", "raw")
ROOTS = [r"C:\Users\Barefootinc TH\Downloads\Sales Records", r"C:\Users\Barefootinc TH\Downloads"]
ETL_FILES = ["etl.py", "etl_jul2026.py", "etl_2025.py", "etl_2024.py"]


def norm(s):
    s = re.sub(r"^[0-9a-f]{8}-", "", s)
    s = os.path.splitext(s)[0].lower()
    return re.sub(r"[_,()\s\-]+", " ", s).strip()


def main(dry=False):
    refs = set()
    for f in ETL_FILES:
        txt = open(os.path.join(ETL_DIR, f), encoding="utf-8").read()
        refs.update(re.findall(r"SRC \+ '([^']+)'", txt))
    local = set()
    for root in ROOTS:
        for pat in ("*.xlsx", "*.csv", os.path.join("**", "*.xlsx"), os.path.join("**", "*.csv")):
            local.update(glob.glob(os.path.join(root, pat), recursive=True))
    local = sorted(local)
    loc_norm = {p: norm(os.path.basename(p)) for p in local}
    os.makedirs(RAW, exist_ok=True)
    bad = 0
    for ref in sorted(refs):
        n = norm(ref)
        best = max(local, key=lambda p: difflib.SequenceMatcher(None, n, loc_norm[p]).ratio())
        score = difflib.SequenceMatcher(None, n, loc_norm[best]).ratio()
        dst = os.path.join(RAW, ref)
        mark = "ok " if score >= 0.9 else "?? "
        if score < 0.9:
            bad += 1
        print(f"{mark}{ref:75s} <- {best.replace(ROOTS[1] + os.sep, '')} ({score:.2f})")
        if not dry and score >= 0.9:
            shutil.copy2(best, dst)
    print(f"\n{len(refs)} referenced, {bad} unmatched, staged in {os.path.abspath(RAW)}")
    return bad


if __name__ == "__main__":
    sys.exit(1 if main("--dry-run" in sys.argv) else 0)
