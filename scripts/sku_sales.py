"""Aggregate the sales ETL records into VFF-shoe SKU x month quantities and match them to the stock sheet's SKUs.

Inputs : data/records.json + data/records_jul2026.json (from scripts/sales_etl/*.py), data/stock.json
Output : data/sku_sales.json  {months, skus: {<stock row>: {month: qty}}, model_monthly, unmatched, ...}

Rules copied from sales-1st-half-26-dashboard's aggregate.py so the model-level
totals reproduce that dashboard exactly: the same excluded stores and the same
model-name canonicalisation (casing/brand-prefix variants collapsed to the
highest-revenue spelling). Store 'Export' (the one-off June 2026 export) is
kept in model_monthly (for the reconciliation check against the sales JSON)
but excluded from the per-SKU figures, which are meant to be domestic pace.

Usage: python scripts/sku_sales.py
"""
import json, os, re
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "..", "data")
p = lambda *a: os.path.join(DATA, *a)

EXCLUDED_STORES_ALL_YEARS = {'Central Ladprao 3F (Coollabo)', 'VFF Cart LP', 'Thaniya', 'K Village', 'EDV Consignment'}
EXPORT_STORE = 'Export'
GENDER = {'Women': 'W', 'Men': 'M', 'Unisex': 'U'}

# Excel colour spelling -> ETL colour spelling, for the few that differ beyond punctuation/case.
COLOR_ALIAS = {
    'deeplakebalck': 'deeplakeblack',      # Scramkey typo in the stock sheet
    'militarydarkgy': 'militarydarkgray',  # abbreviated in some online receipts
}


def norm_color(c):
    c = re.sub(r'[^a-z0-9]', '', str(c or '').lower())
    return COLOR_ALIAS.get(c, c)


def load_records(names):
    recs = []
    for n in names:
        recs += json.load(open(p(n), encoding="utf-8"))
    return [r for r in recs if r['store'] not in EXCLUDED_STORES_ALL_YEARS]


def build_canon(records):
    """Same casing/brand-prefix collapse as aggregate.py (Oleno aliases omitted: VFF only here)."""
    def strip_brand(brand, model):
        return model.strip() if not brand else re.sub(r'^\s*' + re.escape(brand) + r'\s+', '', model.strip(), flags=re.I)
    casing = defaultdict(lambda: defaultdict(float))
    for r in records:
        if r['model']:
            casing[(r['brand'], strip_brand(r['brand'], r['model']).upper())][r['model']] += r['amount']
    def pick(brand, variants_amount):
        variants = list(variants_amount)
        bu = (brand or '').strip().upper()
        pool = [v for v in variants if v.strip().upper().startswith(bu + ' ') or v.strip().upper() == bu] or variants
        pool = [v for v in pool if v != v.upper()] or pool
        return max(pool, key=lambda v: variants_amount[v])
    canon = {(brand, v.upper()): pick(brand, va) for (brand, _k), va in casing.items() for v in va}
    return lambda brand, m: canon.get((brand, m.upper()), m) if m else m


def main():
    records = load_records(["records.json", "records_jul2026.json"])
    # canonical spellings are chosen over all years' records, as aggregate.py does
    canon = build_canon(records + load_records(["records_2025.json", "records_2024.json"]))
    stock = json.load(open(p("stock.json"), encoding="utf-8"))
    months = sorted({r['month'] for r in records if r['month'].startswith('2026')})

    model_monthly = defaultdict(lambda: defaultdict(float))        # incl. Export (matches sales JSON)
    export_by_model = defaultdict(float)
    combo = defaultdict(lambda: defaultdict(float))                # (model, gender, color, size) -> month -> qty (domestic)
    combo_raw = {}
    for r in records:
        if r['brand'] != 'VFF' or not r.get('is_vff_shoe') or not r['month'].startswith('2026'):
            continue
        m = canon(r['brand'], r['model'])
        model_monthly[m][r['month']] += r['qty']
        if r['store'] == EXPORT_STORE:
            export_by_model[m] += r['qty']
            continue
        g = GENDER.get(r.get('gender') or 'Unisex', 'U')
        size = int(r['size']) if r.get('size') and str(r['size']).isdigit() else None
        key = (m, g, norm_color(r.get('color')), size)
        combo[key][r['month']] += r['qty']
        combo_raw.setdefault(key, (r['model'], r.get('gender'), r.get('color'), r.get('size')))

    # match stock SKUs
    sku_index = {}
    for s in stock['skus']:
        if s['sales_model']:
            # model names compared case-insensitively: the canonical spelling can differ between runs ('VFF KSO' / 'VFF Kso')
            sku_index.setdefault((s['sales_model'].upper(), s['gender'], norm_color(s['color']), s['size_num']), []).append(s['row'])
    stocked_models = {s['sales_model'].upper() for s in stock['skus'] if s['sales_model']}
    skus, unmatched = {}, []
    matched_qty = 0.0
    dup = 0
    for key, by_month in combo.items():
        rows = sku_index.get((key[0].upper(),) + key[1:])
        if rows:
            if len(rows) > 1:
                dup += 1
            row = rows[0]
            skus.setdefault(row, {})
            for mo, q in by_month.items():
                skus[row][mo] = skus[row].get(mo, 0) + q
            matched_qty += sum(by_month.values())
        else:
            unmatched.append(dict(model=key[0], gender=key[1], color=combo_raw[key][2], size=combo_raw[key][3], qty=sum(by_month.values())))
    unmatched.sort(key=lambda x: -x['qty'])
    unmatched_qty = sum(u['qty'] for u in unmatched)

    # reconciliation against the sales JSON already in data/
    sales = json.load(open(p("sales_vff.json"), encoding="utf-8"))
    mm = sales['vff_shoes']['model_monthly']
    print(f"{'model':22s} {'local2026':>9s} {'json2026':>8s} {'diff':>5s} {'export':>6s}")
    for m in sorted(model_monthly, key=lambda m: -sum(model_monthly[m].values())):
        loc = sum(model_monthly[m].values())
        js = sum(int(x.get('qty', 0) or 0) for k, x in mm.get(m, {}).items() if k.startswith('2026'))
        flag = '' if abs(loc - js) < 0.5 else '  <-- MISMATCH'
        print(f"{m:22s} {loc:9.0f} {js:8d} {loc - js:5.0f} {export_by_model.get(m, 0):6.0f}{flag}")
    print(f"\nVFF shoe records 2026 (domestic): matched to stock SKUs {matched_qty:.0f}, unmatched {unmatched_qty:.0f} "
          f"({matched_qty / max(matched_qty + unmatched_qty, 1):.1%} matched); duplicate stock rows for a key: {dup}")
    print("unmatched within models that ARE in the stock sheet (model, gender, color, size, qty):")
    for u in [u for u in unmatched if u['model'].upper() in stocked_models][:40]:
        print(f"  {u['model']:20s} {u['gender']} {str(u['color']):28s} {str(u['size']):4s} {u['qty']:5.0f}")
    by_model_unmatched = defaultdict(float)
    for u in unmatched:
        by_model_unmatched[u['model']] += u['qty']
    print("unmatched by model:", {k: round(v) for k, v in sorted(by_model_unmatched.items(), key=lambda kv: -kv[1])})

    out = dict(
        months=months, export_store=EXPORT_STORE, excluded_stores=sorted(EXCLUDED_STORES_ALL_YEARS),
        source_files=["records.json", "records_jul2026.json"],
        skus={str(row): {mo: round(q, 1) for mo, q in sorted(v.items())} for row, v in skus.items()},
        model_monthly={m: {mo: round(q, 1) for mo, q in sorted(v.items())} for m, v in model_monthly.items()},
        export_by_model={m: round(q, 1) for m, q in export_by_model.items()},
        matched_qty=round(matched_qty, 1), unmatched_qty=round(unmatched_qty, 1), unmatched=unmatched,
    )
    json.dump(out, open(p("sku_sales.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("wrote", os.path.abspath(p("sku_sales.json")))


if __name__ == "__main__":
    main()
