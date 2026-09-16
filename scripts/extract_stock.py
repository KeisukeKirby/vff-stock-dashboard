"""Extract SKU-level stock from the 'VFF Stock <dd-mm-yy>.xlsx' sheet into data/stock.json.

Stdlib only (zipfile + xml) so it runs on a bare Python install.

Sheet layout (sheet 'Net (2)'), 0-based column index:
    0 No. | 1 ITEMS (merged down per model+color) | 2 Size (M39 / W38 / 36)
    3..43  one column per import invoice (IVyymmnnn) plus 'Defect'/'Influ' adjustments
    45..48 pending purchase orders (header 'Import'; not yet in Import Total)
    50 Import Total | 51..54 Sales 2023/2024/2025/2026 | 60 Sales Total
    61 Balance (= Import Total - Sales Total)  <- used as the stock figure
    62 Order Import (= sum of 45..48)
    65 'Office' physical count (blank for July-2026 arrivals, so Balance is the reliable one)
    67..80 per-store stock columns (all empty in the 31-08-26 file)

Usage: python scripts/extract_stock.py "<path to xlsx>"  -> writes data/stock.json
"""
import json, os, re, sys, zipfile
import xml.etree.ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "data", "stock.json")

# Excel model name -> model key used by sales-1st-half-26-dashboard's
# data/dashboard_data.json (vff_shoes.model_monthly). Trailope WP and
# Groundsplay LS are variants that the sales file does not separate.
MODEL_MAP = {
    "V-Run": "VFF V-Run", "V-Soul": "VFF V-Soul", "KSO EVO": "VFF Kso Evo", "KSO": "VFF KSO",
    "Breezandal": "VFF Breezandal", "Trailope": "VFF Trailope", "Trailope WP": "VFF Trailope",
    "V-Alpha": "VFF V-Alpha", "Spidrwalk": "VFF Spidrwalk", "Graspifier": "VFF Graspifier",
    "V-Train 2.0": "VFF V-Train 2.0", "V-Trek": "VFF V-Trek", "Scramkey": "VFF Scramkey",
    "KMD Evo": "VFF KMD Evo", "Groundsplay": "VFF Groundsplay", "Groundsplay LS": "VFF Groundsplay",
}
COL = dict(no=0, item=1, size=2, import_total=50, s2023=51, s2024=52, s2025=53, s2026=54,
           sales_total=60, balance=61, order_import=62, office=65)
INVOICE_COLS = range(3, 44)
ORDER_COLS = range(45, 49)


def col_idx(ref):
    n = 0
    for ch in re.match(r"[A-Z]+", ref).group(0):
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def read_sheet(path):
    z = zipfile.ZipFile(path)
    shared = ["".join(t.text or "" for t in si.iter("{%s}t" % NS["m"]))
              for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", NS)]
    sheet = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
    rows = []
    for row in sheet.find("m:sheetData", NS).findall("m:row", NS):
        cells = {}
        for c in row.findall("m:c", NS):
            t, v = c.get("t"), c.find("m:v", NS)
            if t == "s" and v is not None:
                val = shared[int(v.text)]
            elif t == "inlineStr":
                val = "".join(x.text or "" for x in c.iter("{%s}t" % NS["m"]))
            elif v is not None:
                try:
                    f = float(v.text); val = int(f) if f.is_integer() else f
                except ValueError:
                    val = v.text
            else:
                continue
            cells[col_idx(c.get("r"))] = val
        rows.append((int(row.get("r")), cells))
    return rows


def parse_item(text):
    """'V-Run M\\n(Black/Yellow)' -> ('V-Run', 'M', 'Black/Yellow'); 'Groundsplay LS (Black Lime)' -> gender None."""
    s = " ".join(text.split())
    m = re.match(r"^(?P<model>.+?)(?:\s+(?P<gender>[MWU]))?\s*\((?P<color>[^()]*)\)\s*$", s)
    if not m:
        return s, None, None
    return m.group("model").strip(), m.group("gender"), m.group("color").strip()


def main(path):
    rows = read_sheet(path)
    hdr = rows[2][1]  # row 3 holds the per-column headers
    invoice_month = {}
    for i in INVOICE_COLS:
        h = hdr.get(i)
        mm = re.match(r"IV(\d{2})(\d{2})\d{3}$", str(h or ""))
        if mm:
            invoice_month[i] = f"20{mm.group(1)}-{mm.group(2)}"

    skus, cur_item, skipped = [], None, []
    for rnum, c in rows[3:]:
        item, size = c.get(COL["item"]), c.get(COL["size"])
        if item and size is None and c.get(COL["no"]) is None:
            continue  # section banner such as 'Vibram'
        if item:
            cur_item = item
        if size is None or size == "TOTAL":
            continue
        size = str(size).strip()
        if not re.match(r"^[MW]?\d{2}$", size):
            skipped.append((rnum, cur_item, size)); continue

        def num(i):
            v = c.get(i); return v if isinstance(v, (int, float)) else 0

        model, gender, color = parse_item(cur_item)
        if gender is None:
            gender = size[0] if size[0] in "MW" else "U"
        size_num = int(size[-2:])
        first_import = min((invoice_month[i] for i in INVOICE_COLS if i in invoice_month and num(i) > 0), default=None)
        skus.append(dict(
            row=rnum, item=" ".join(cur_item.split()), model=model, gender=gender, color=color,
            size=size, size_num=size_num, sales_model=MODEL_MAP.get(model),
            import_total=num(COL["import_total"]), s2023=num(COL["s2023"]), s2024=num(COL["s2024"]),
            s2025=num(COL["s2025"]), s2026=num(COL["s2026"]), sales_total=num(COL["sales_total"]),
            stock=num(COL["balance"]), on_order=num(COL["order_import"]),
            office=c.get(COL["office"]), first_import=first_import,
        ))

    as_of = re.search(r"(\d{2})-(\d{2})-(\d{2})", os.path.basename(path))
    as_of = f"20{as_of.group(3)}-{as_of.group(2)}-{as_of.group(1)}" if as_of else None
    # Excel's own TOTAL row, for the reconciliation note
    total_row = next((c for _, c in rows if c.get(COL["size"]) == "TOTAL"), {})
    out = dict(
        as_of=as_of, source_file=os.path.basename(path), sku_count=len(skus),
        excel_total_row={k: total_row.get(i) for k, i in COL.items() if i in total_row},
        skus=skus,
    )
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)

    tot = lambda k: sum(s[k] for s in skus)
    print(f"as_of={as_of} skus={len(skus)} models={len({s['model'] for s in skus})} items={len({s['item'] for s in skus})}")
    print(f"stock={tot('stock')} on_order={tot('on_order')} s2026={tot('s2026')} sales_total={tot('sales_total')} import_total={tot('import_total')}")
    print("excel TOTAL row:", out["excel_total_row"])
    if skipped:
        print("skipped rows:", skipped)
    unmapped = sorted({s["model"] for s in skus if not s["sales_model"]})
    if unmapped:
        print("MODELS WITHOUT SALES MAPPING:", unmapped)
    print("wrote", os.path.abspath(OUT))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\Barefootinc TH\Downloads\VFF Stock 31-08-26.xlsx new.xlsx")
