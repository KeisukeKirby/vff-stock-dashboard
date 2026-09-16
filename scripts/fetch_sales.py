"""Copy the VFF-shoes slice of sales-1st-half-26-dashboard's data/dashboard_data.json
into data/sales_vff.json, recording which commit it came from.

Usage: python scripts/fetch_sales.py <path to dashboard_data.json> [<git commit>]
"""
import json, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "data", "sales_vff.json")
KEEP = ["note", "monthly", "monthly_prev", "monthly_2024h1", "store_monthly", "model_monthly",
        "gender_monthly", "color_monthly", "gender_size_monthly"]


def main(src, commit=None):
    d = json.load(open(src, encoding="utf-8"))
    v = d["vff_shoes"]
    out = dict(
        source_repo="keisukekirby/sales-1st-half-26-dashboard",
        source_branch="claude/sales-analytics-dashboard-oxxvy1",
        source_commit=commit, source_path="data/dashboard_data.json",
        months=d["months"], prev_months=d["prev_months"],
        stores=[{"store": s["store"], "group": s["group"]} for s in d["stores"]],
        group_label=d["group_label"],
        vff_shoes={k: v[k] for k in KEEP if k in v},
    )
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print("months:", out["months"], "models:", len(out["vff_shoes"]["model_monthly"]),
          "size(KB):", os.path.getsize(OUT) // 1024)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
