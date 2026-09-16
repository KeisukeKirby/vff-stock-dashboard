"""Build index.html: src/template.html + data/stock.json + data/sales_vff.json -> index.html.

Everything the page needs is embedded as one JSON blob (window.DATA), so the
result is a single self-contained file that Vercel serves statically. Only the
slice of the sales file the page uses (model-level monthly qty for 2026) is
embedded, to keep the page small.

Usage: python scripts/build.py
"""
import json, os, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..")
p = lambda *a: os.path.join(ROOT, *a)

stock = json.load(open(p("data", "stock.json"), encoding="utf-8"))
sales = json.load(open(p("data", "sales_vff.json"), encoding="utf-8"))
template = open(p("src", "template.html"), encoding="utf-8").read()

months = [m for m in sales["months"] if m.startswith("2026")]
mm = sales["vff_shoes"]["model_monthly"]
used_models = {s["sales_model"] for s in stock["skus"] if s["sales_model"]}
model_monthly_qty = {m: {mo: int(mm.get(m, {}).get(mo, {}).get("qty", 0) or 0) for mo in months} for m in sorted(used_models)}

# The one-off export that inflates June in the sales file is not in the stock
# sheet's sales column, so per model the Jan-Aug gap between the two sources
# is that model's export share. Gaps under 50 pairs are ordinary timing
# differences and are left alone; the estimate is capped at the June qty.
export_month = "2026-06"
export_total = int(sales["vff_shoes"]["store_monthly"].get("Export", {}).get(export_month, {}).get("qty", 0) or 0)
excel26 = {}
for s in stock["skus"]:
    if s["sales_model"]:
        excel26[s["sales_model"]] = excel26.get(s["sales_model"], 0) + s["s2026"]
export_adjust = {}
for m, xl in excel26.items():
    gap = sum(model_monthly_qty[m].values()) - xl
    if gap >= 50:
        export_adjust[m] = min(gap, model_monthly_qty[m][export_month])
print("export estimate per model:", export_adjust, "total", sum(export_adjust.values()), "/ sales file", export_total)

SKU_FIELDS = ["row", "item", "model", "gender", "color", "size", "size_num", "sales_model",
              "stock", "on_order", "s2026", "s2025", "s2024", "first_import"]
data = dict(
    built_at=datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
    as_of=stock["as_of"], source_file=stock["source_file"],
    skus=[{k: s[k] for k in SKU_FIELDS} for s in stock["skus"]],
    sales=dict(source_repo=sales["source_repo"], source_commit=sales["source_commit"],
               months=months, recent=months[-3:], model_monthly_qty=model_monthly_qty),
    export_month=export_month, export_total=export_total, export_adjust=export_adjust,
)
blob = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
html = template.replace("/*__DATA__*/null", blob)
assert html != template, "template is missing the /*__DATA__*/null placeholder"
open(p("index.html"), "w", encoding="utf-8").write(html)
print(f"skus={len(data['skus'])} models_in_sales={len(model_monthly_qty)} months={months[0]}..{months[-1]} export_{export_month}={export_total}")
print("wrote index.html", os.path.getsize(p("index.html")) // 1024, "KB")
