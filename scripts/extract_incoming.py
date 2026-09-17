"""Extract per-SKU incoming quantities from a Vibram packing list (.xls) into data/incoming.json.

The supplier sends legacy .xls (OLE2) files, which the Python standard library
cannot read, so the sheet is exported to TSV through Excel itself (COM,
read-only) by a short PowerShell step -- no extra Python package needed.

Layout (Guangzhou Vibram Rubber packing list): header lines with 'Invoice No.:',
'Date:', 'ETD:', 'Shipped Total: <n>Cartons (<n> prs)'; then one block per
style/colour: 'PO#:', 'Style:', 'Style#:' (e.g. 26W3201 -> gender W),
'Color:', a 'CARTON#' header whose next row holds the size labels, carton rows
'<from> - <to> | CTNS | PAIRS | qty per size... | PRS/CTN', and a TOTAL row.
Every carton row, block TOTAL and the grand total are cross-checked; any
mismatch aborts without writing.

Re-running with another packing list adds that PO to data/incoming.json (an
existing entry with the same PO is replaced). When a shipment has been received
and appears in the month-end stock sheet, remove it with --remove <PO>.

Usage: python scripts/extract_incoming.py "<packing list .xls>"
       python scripts/extract_incoming.py --remove CN26015C
"""
import datetime, json, os, re, subprocess, sys, tempfile
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "data", "incoming.json")

PS_EXPORT = r"""
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($env:PL_SRC, 0, $true)
  $ws = $null
  foreach ($s in $wb.Worksheets) { if ($s.Visible -eq -1) { $ws = $s; break } }
  $ur = $ws.UsedRange; $rows = $ur.Rows.Count; $cols = [Math]::Min($ur.Columns.Count, 40)
  $vals = $ws.Range($ws.Cells(1, 1), $ws.Cells($ur.Row + $rows - 1, $cols)).Value2
  $sb = New-Object System.Text.StringBuilder
  for ($r = 1; $r -le $ur.Row + $rows - 1; $r++) {
    $line = for ($c = 1; $c -le $cols; $c++) { ("" + $vals[$r, $c]) -replace "[`t`r`n]", ' ' }
    [void]$sb.AppendLine($line -join "`t")
  }
  [System.IO.File]::WriteAllText($env:PL_OUT, $sb.ToString(), (New-Object System.Text.UTF8Encoding $false))
  $wb.Close($false)
} finally { $xl.Quit(); [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) }
"""


def export_tsv(path):
    fd, tsv = tempfile.mkstemp(suffix=".tsv"); os.close(fd)
    env = dict(os.environ, PL_SRC=os.path.abspath(path), PL_OUT=tsv)
    subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", PS_EXPORT], env=env, check=True)
    rows = [line.rstrip("\r\n").split("\t") for line in open(tsv, encoding="utf-8")]
    os.remove(tsv)
    return rows


def num(v):
    try:
        f = float(v); return int(f) if f.is_integer() else f
    except (TypeError, ValueError):
        return None


def parse_date(text, fmts):
    for f in fmts:
        try:
            return datetime.datetime.strptime(text.strip(), f).date().isoformat()
        except ValueError:
            pass
    return None


def parse(rows):
    joined = ["\t".join(r) for r in rows]
    head = "\n".join(joined[:20])
    shipment = {}
    m = re.search(r"Invoice No\.?:\s*([\w-]+)", head); shipment["invoice"] = m.group(1) if m else None
    m = re.search(r"Date:\s*([A-Za-z]+ \d{1,2},\s*\d{4})", head); shipment["doc_date"] = parse_date(m.group(1), ["%B %d, %Y"]) if m else None
    m = re.search(r"ETD:\s*(\d{1,2})\s*([A-Za-z]{3})[A-Za-z]*\s*,?\s*(\d{4})", head)
    shipment["etd"] = parse_date(f"{m.group(1)} {m.group(2)} {m.group(3)}", ["%d %b %Y"]) if m else None
    m = re.search(r"Shipped Total:\s*(\d+)\s*Cartons?\s*\((\d+)\s*prs\)", head, re.I)
    declared = (int(m.group(1)), int(m.group(2))) if m else None

    blocks, cur, size_cols, prs_col = [], None, {}, None
    for i, r in enumerate(rows):
        a = [c.strip() for c in r] + [""] * 3
        if a[0] == "PO#:":
            cur = dict(po=a[2], lines=defaultdict(int), cartons=0, pairs=0); blocks.append(cur)
        elif cur is not None and a[0] == "Style:":
            cur["style"] = a[2]
        elif cur is not None and a[0] == "Style#:":
            cur["style_no"] = a[2]
        elif cur is not None and a[0] == "Color:":
            cur["color"] = a[2]
        elif cur is not None and a[0].upper().startswith("CARTON#"):
            prs_col = next(j for j, c in enumerate(a) if c.upper().startswith("PRS/CTN"))
            nxt = rows[i + 1]
            size_cols = {j: int(float(c)) for j, c in enumerate(nxt) if j < prs_col and num(c) is not None}
        elif cur is not None and num(a[0]) is not None and a[1] == "-":
            ctns, pairs, per = num(a[3]), num(a[4]), num(a[prs_col])
            run = {size_cols[j]: num(a[j]) for j in size_cols if num(a[j])}
            where = f"{cur.get('style')} {cur.get('color')} carton {a[0]}-{a[2]}"
            assert sum(run.values()) == per, f"{where}: sizes sum {sum(run.values())} != PRS/CTN {per}"
            assert ctns * per == pairs and num(a[2]) - num(a[0]) + 1 == ctns, f"{where}: {ctns} ctns x {per} != {pairs}"
            for s, q in run.items():
                cur["lines"][s] += q * ctns
            cur["cartons"] += ctns; cur["pairs"] += pairs
        elif cur is not None and a[0] == "TOTAL":
            total_run = {size_cols[j]: num(a[j]) for j in size_cols if num(a[j])}
            where = f"{cur.get('style')} {cur.get('color')}"
            assert (num(a[3]), num(a[4])) == (cur["cartons"], cur["pairs"]), f"{where}: TOTAL row {a[3]}/{a[4]} != cartons {cur['cartons']}/{cur['pairs']}"
            assert total_run == dict(cur["lines"]), f"{where}: TOTAL size run {total_run} != {dict(cur['lines'])}"

    cartons, pairs = sum(b["cartons"] for b in blocks), sum(b["pairs"] for b in blocks)
    if declared:
        assert declared == (cartons, pairs), f"Shipped Total {declared} != parsed {(cartons, pairs)}"
    pos = sorted({b["po"] for b in blocks})
    assert len(pos) == 1, f"expected one PO per packing list, got {pos}"
    lines = []
    for b in blocks:
        g = re.search(r"\d{2}([MWU])\d", b.get("style_no", ""))
        for s, q in sorted(b["lines"].items()):
            lines.append(dict(style=b["style"], style_no=b.get("style_no"), gender=g.group(1) if g else "U", color=b["color"], size=s, qty=q))
    shipment.update(po=pos[0], cartons=cartons, pairs=pairs, lines=lines)
    return shipment


def load():
    return json.load(open(OUT, encoding="utf-8")) if os.path.exists(OUT) else {"shipments": []}


def save(d):
    d["shipments"].sort(key=lambda s: (s.get("etd") or "", s["po"]))
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    json.dump(d, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)


def main(argv):
    d = load()
    if argv[:1] == ["--remove"]:
        before = len(d["shipments"]); d["shipments"] = [s for s in d["shipments"] if s["po"] != argv[1]]
        save(d); print(f"removed {before - len(d['shipments'])} shipment(s) with PO {argv[1]}"); return
    path = argv[0]
    sh = parse(export_tsv(path))
    invoice = sh.pop("invoice")  # printed for checking only: the repo is public, like the stock workbook's invoice columns
    sh["source_file"] = os.path.basename(path)
    d["shipments"] = [s for s in d["shipments"] if s["po"] != sh["po"]] + [sh]
    save(d)
    by_item = defaultdict(int)
    for l in sh["lines"]:
        by_item[(l["style"], l["gender"], l["color"])] += l["qty"]
    print(f"PO {sh['po']}  invoice {invoice}  date {sh['doc_date']}  ETD {sh['etd']}  {sh['cartons']} ctns / {sh['pairs']} prs  ({len(sh['lines'])} SKU lines)")
    for (st, g, c), q in by_item.items():
        print(f"  {st:16s} {g} {c:18s} {q:4d}")
    print("wrote", os.path.abspath(OUT))


if __name__ == "__main__":
    main(sys.argv[1:])
