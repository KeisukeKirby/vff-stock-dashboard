"""Build index.html: src/template.html + data/stock.json + data/sales_vff.json -> index.html.

Everything the page needs is embedded as one JSON blob (window.DATA), so the
result is a single self-contained file that Vercel serves statically.
The only derived numbers computed here (rather than in the browser) are the
per-model June-2026 export estimates, because they need both sources at once
and are easier to audit as a printed table.

Usage: python scripts/build.py
"""
import json, os, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..")
p = lambda *a: os.path.join(ROOT, *a)

stock = json.load(open(p("data", "stock.json"), encoding="utf-8"))
sales = json.load(open(p("data", "sales_vff.json"), encoding="utf-8"))
template = open(p("src", "template.html"), encoding="utf-8").read()

# ---- June 2026 export estimate --------------------------------------------
# The sales file counts a one-off 948-pair export in 2026-06 (store 'Export');
# the stock sheet's 2026 sales column does not. Per model, the gap between the
# two sources over Jan-Aug is therefore the export share of that model. Only
# gaps >= 50 pairs are treated as export (smaller gaps are ordinary timing /
# consignment differences), and the estimate is capped at the model's June qty.
mm = sales["vff_shoes"]["model_monthly"]
months26 = [m for m in sales["months"] if m.startswith("2026")]
excel26 = {}
for s in stock["skus"]:
    if s["sales_model"]:
        excel26[s["sales_model"]] = excel26.get(s["sales_model"], 0) + s["s2026"]
export_month = "2026-06"
export_total = int(sales["vff_shoes"]["store_monthly"].get("Export", {}).get(export_month, {}).get("qty", 0) or 0)
export_adjust, recon = {}, []
for model, xl in sorted(excel26.items()):
    js = sum(int(mm.get(model, {}).get(m, {}).get("qty", 0) or 0) for m in months26)
    jun = int(mm.get(model, {}).get(export_month, {}).get("qty", 0) or 0)
    gap = js - xl
    est = min(gap, jun) if gap >= 50 else 0
    if est:
        export_adjust[model] = est
    recon.append(dict(model=model, excel_2026=xl, json_2026=js, gap=gap, june=jun, export_est=est))
print(f"{'model':18s} {'excel26':>7s} {'json26':>7s} {'gap':>5s} {'june':>5s} {'export':>6s}")
for r in recon:
    print(f"{r['model']:18s} {r['excel_2026']:7d} {r['json_2026']:7d} {r['gap']:5d} {r['june']:5d} {r['export_est']:6d}")
print(f"export estimate total: {sum(export_adjust.values())} (sales file Export store in {export_month}: {export_total})")

sold_without_stock = sorted(
    (m for m in mm if m not in excel26 and sum(int(x.get("qty", 0) or 0) for k, x in mm[m].items() if k.startswith("2026")) > 0),
    key=lambda m: -sum(int(x.get("qty", 0) or 0) for k, x in mm[m].items() if k.startswith("2026")))

data = dict(
    built_at=datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
    stock=stock, sales=sales,
    export_month=export_month, export_total=export_total, export_adjust=export_adjust,
    reconciliation=recon,
    sold_without_stock=[dict(model=m, qty_2026=sum(int(x.get("qty", 0) or 0) for k, x in mm[m].items() if k.startswith("2026")))
                        for m in sold_without_stock],
)
blob = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
html = template.replace("/*__DATA__*/null", blob)
assert html != template, "template is missing the /*__DATA__*/null placeholder"
open(p("index.html"), "w", encoding="utf-8").write(html)
print("wrote index.html", os.path.getsize(p("index.html")) // 1024, "KB")
