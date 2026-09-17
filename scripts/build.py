"""Build index.html: src/template.html + data/stock.json + data/sku_sales.json (+ data/sales_vff.json for reconciliation) -> index.html.

Everything the page needs is embedded as one JSON blob (window.DATA), so the
result is a single self-contained file that Vercel serves statically.

Sales figures now come from data/sku_sales.json, i.e. the sales dashboard's own
ETL scripts run locally (scripts/sales_etl/) and aggregated per SKU x month by
scripts/sku_sales.py. data/sales_vff.json (the dashboard's published JSON) is
only used to print/embed a reconciliation of the two.

Usage: python scripts/build.py
"""
import json, os, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..")
p = lambda *a: os.path.join(ROOT, *a)

stock = json.load(open(p("data", "stock.json"), encoding="utf-8"))
sku_sales = json.load(open(p("data", "sku_sales.json"), encoding="utf-8"))
sales = json.load(open(p("data", "sales_vff.json"), encoding="utf-8"))
template = open(p("src", "template.html"), encoding="utf-8").read()

months = sku_sales["months"]
export_month = "2026-06"
export_by_model = {m: q for m, q in sku_sales["export_by_model"].items() if q}
export_total = sum(export_by_model.values())

# ---- reconciliation: local ETL run vs the published sales JSON (model-level 2026 qty) ----
mm_json = sales["vff_shoes"]["model_monthly"]
local_mm = sku_sales["model_monthly"]
def upper_key(d):
    out = {}
    for k, v in d.items():
        out[k.upper()] = out.get(k.upper(), 0) + v
    return out
loc_tot = upper_key({m: sum(v.values()) for m, v in local_mm.items()})
js_tot = upper_key({m: sum(int(x.get("qty", 0) or 0) for mo, x in v.items() if mo.startswith("2026")) for m, v in mm_json.items()})
recon = []
for k in sorted(set(loc_tot) | set(js_tot), key=lambda k: -max(loc_tot.get(k, 0), js_tot.get(k, 0))):
    if max(loc_tot.get(k, 0), js_tot.get(k, 0)) > 0:
        recon.append(dict(model=k.title().replace("Vff ", "VFF ").replace("Kso", "KSO").replace("Kmd", "KMD").replace("El-X", "EL-X").replace("Cvt", "CVT"),
                          local=round(loc_tot.get(k, 0)), json=js_tot.get(k, 0)))
print(f"{'model':24s} {'local':>6s} {'json':>6s} {'diff':>5s}")
for r in recon:
    print(f"{r['model']:24s} {r['local']:6d} {r['json']:6d} {r['local'] - r['json']:5d}")
print(f"export {export_month}: {export_by_model} total {export_total}")
print(f"matched to stock SKUs: {sku_sales['matched_qty']} pairs, unmatched: {sku_sales['unmatched_qty']}")

# per-SKU monthly (domestic, from the ETL records) attached to each stock row
by_row = sku_sales["skus"]
SKU_FIELDS = ["row", "item", "model", "gender", "color", "size", "size_num", "sales_model",
              "stock", "on_order", "s2026", "s2025", "s2024", "first_import"]
skus = []
for s in stock["skus"]:
    d = {k: s[k] for k in SKU_FIELDS}
    d["monthly"] = by_row.get(str(s["row"]), {})
    d["incoming"] = 0
    skus.append(d)

# ---- incoming shipments (packing lists, scripts/extract_incoming.py) matched to stock SKUs ----
# Packing-list colour spellings differ only in punctuation ('Ivory-Deep Lake' vs 'Ivory Deep Lake'),
# so model/colour are compared with non-alphanumerics stripped.
norm = lambda t: "".join(ch for ch in str(t or "").lower() if ch.isalnum())
incoming_path = p("data", "incoming.json")
shipments = json.load(open(incoming_path, encoding="utf-8"))["shipments"] if os.path.exists(incoming_path) else []
sku_by_key = {(norm(s["model"]), s["gender"], norm(s["color"]), s["size_num"]): s for s in skus}
incoming_unmatched = []
for sh in shipments:
    sh["matched_pairs"] = 0
    for l in sh["lines"]:
        s = sku_by_key.get((norm(l["style"]), l["gender"], norm(l["color"]), l["size"]))
        if s:
            s["incoming"] += l["qty"]; sh["matched_pairs"] += l["qty"]
        else:
            incoming_unmatched.append(dict(po=sh["po"], **l))
incoming_vs_order = [dict(model=s["model"], gender=s["gender"], color=s["color"], size=s["size"], incoming=s["incoming"], on_order=s["on_order"])
                     for s in skus if s["incoming"] != s["on_order"]]
for sh in shipments:
    print(f"incoming PO {sh['po']} ETD {sh['etd']}: {sh['pairs']} prs, matched to stock SKUs {sh['matched_pairs']}")
if incoming_unmatched:
    print("WARNING incoming lines not in the stock sheet:", incoming_unmatched)
print(f"SKUs where packing-list incoming != stock sheet Order Import: {len(incoming_vs_order)}", incoming_vs_order[:10])

stocked = {s["sales_model"].upper() for s in stock["skus"] if s["sales_model"]}
unmatched_stocked = [u for u in sku_sales["unmatched"] if u["model"].upper() in stocked]

data = dict(
    built_at=datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
    as_of=stock["as_of"], source_file=stock["source_file"],
    skus=skus,
    sales=dict(source_repo=sales["source_repo"], source_commit=sales["source_commit"],
               months=months, recent=months[-3:],
               model_monthly_qty={m: {mo: v.get(mo, 0) for mo in months} for m, v in local_mm.items()},
               excluded_stores=sku_sales["excluded_stores"]),
    export_month=export_month, export_total=export_total, export_by_model=export_by_model,
    reconciliation=recon,
    unmatched_stocked=unmatched_stocked, unmatched_stocked_qty=round(sum(u["qty"] for u in unmatched_stocked)),
    unmatched_other_qty=round(sku_sales["unmatched_qty"] - sum(u["qty"] for u in unmatched_stocked)),
    # invoice/container numbers stay out of the public page; PO, dates and totals identify the shipment
    shipments=[{k: sh.get(k) for k in ("po", "doc_date", "etd", "cartons", "pairs", "matched_pairs")} for sh in shipments],
    incoming_unmatched=incoming_unmatched, incoming_vs_order=incoming_vs_order,
)
blob = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
html = template.replace("/*__DATA__*/null", blob)
assert html != template, "template is missing the /*__DATA__*/null placeholder"
open(p("index.html"), "w", encoding="utf-8").write(html)
print(f"skus={len(skus)} months={months[0]}..{months[-1]}")
print("wrote index.html", os.path.getsize(p("index.html")) // 1024, "KB")
