"""Make the copied sales ETL scripts read/write via env vars instead of the cloud session's hard-coded paths.

SALES_RAW_DIR  -> where the staged workbooks live (data/raw/)
SALES_OUT_DIR  -> where records*.json are written (data/)
Only the two path lines per script change; the ETL logic stays byte-identical
to sales-1st-half-26-dashboard so the numbers match that dashboard.
"""
import os, re

HERE = os.path.dirname(os.path.abspath(__file__))
ETL_DIR = os.path.join(HERE, "sales_etl")
for fn in ["etl.py", "etl_jul2026.py", "etl_2025.py", "etl_2024.py"]:
    path = os.path.join(ETL_DIR, fn)
    src = open(path, encoding="utf-8").read()
    n_src = len(re.findall(r'^SRC = "/root/\.claude/uploads/[^"]+"$', src, flags=re.M))
    # lambda replacements: no backslash processing on the replacement text
    src, n1 = re.subn(r'^SRC = "/root/\.claude/uploads/[^"]+"$',
                      lambda m: 'SRC = os.path.join(os.environ["SALES_RAW_DIR"], "")  # patched: was the cloud upload dir', src, flags=re.M)
    src, n2 = re.subn(r"open\('/tmp/claude-0/[^']*/scratchpad/(records[^']*\.json)', 'w'",
                      lambda m: f"open(os.path.join(os.environ['SALES_OUT_DIR'], '{m.group(1)}'), 'w'", src)
    if "import os" not in src.split("\n# ===")[0]:
        src = src.replace("import openpyxl", "import os\nimport openpyxl", 1)
    open(path, "w", encoding="utf-8").write(src)
    print(f"{fn}: SRC lines patched={n1} (found {n_src}), output paths patched={n2}")
