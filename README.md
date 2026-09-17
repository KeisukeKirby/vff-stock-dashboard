# VFF 在庫一覧

Vibram FiveFingers(VFF)シューズの月末在庫表を **SKU(モデル × 性別 × カラー × サイズ)一覧** にし、今年の月平均販売・在庫月数・ステータスを付けた静的ページです。まずはシンプルな一覧から始め、段階的に精度を上げていきます。

- 在庫: 月末の `VFF Stock <dd-mm-yy>.xlsx`(Balance 列 = 輸入累計 − 販売累計)、入荷予定は Order Import 列
- 販売実績: [sales-1st-half-26-dashboard](https://github.com/keisukekirby/sales-1st-half-26-dashboard) の ETL スクリプト(`scripts/sales_etl/`、同リポジトリからの複製)を、同じ生データ(店舗・EC・委託・百貨店の明細 Excel)に対してローカルで実行し、VFF シューズを SKU × 月で集計(`scripts/sku_sales.py`)。公開 JSON(`data/sales_vff.json`)は突合の参照用

## 画面

- **在庫一覧タブ** — SKU(モデル × カラー × サイズ × 性別)をアルファベット順に一覧。列は 在庫数 / 入荷予定 / 今年の月平均 / 直近3ヶ月の月平均 / 在庫月数 / ステータス のみ。ステータス・モデル・性別・カラー検索で絞り込み、列見出しで並び替え、CSV書き出し
- **詳細タブ** — KPI(総在庫、月平均販売、在庫月数、欠品中 / 要発注 / 在庫過剰の SKU 数)、リードタイム設定とステータス定義、モデル別サマリー(月別実績・2026年計)、データの注記

## 計算(意図的に単純)

| 項目 | 定義 |
|---|---|
| 今年の月平均 | SKU の 2026年1〜8月販売足数(明細ベース、輸出を除く)÷ 販売月数(初入荷または初販売が2026年ならその月〜8月、それ以前なら 8) |
| 直近3ヶ月の月平均 | SKU の 6〜8月販売足数 ÷ 3(同じ明細、輸出を除く) |
| 入荷予定 | 仕入先の梱包明細(Packing List .xls)の SKU 別足数。複数便は合算。在庫月数・ステータスの計算には含めない |
| 在庫月数 | 在庫 ÷ 今年の月平均 |
| 欠品中 | 在庫 0 かつ 今年の販売あり |
| 要発注 | 在庫月数 < リードタイム(既定 3ヶ月、画面で 2〜4 を選択) |
| 適正 | リードタイム 〜 12ヶ月 |
| 在庫過剰 | 在庫月数 > 12ヶ月 |
| 動きなし | 今年の販売 0 で在庫あり |
| 入荷待ち | 在庫 0・販売 0・発注済み(新色/新モデル) |

6月の一括輸出(948足、店舗 `Export`)は SKU 別の数字から除いています(在庫表の販売列にも含まれていないため)。明細のカラー・サイズ表記と在庫表の照合率は 94.6%(未照合は在庫表に無い旧カラー・旧モデル)。公開 JSON との突合は店舗×月で一致(イベント5月 +3足のみ差)。

## 構成

| パス | 役割 |
|---|---|
| `index.html` | 完成ページ(生成物。データ埋め込み済みの単一ファイル) |
| `src/template.html` | ページ本体のテンプレート。`/*__DATA__*/null` に build.py がデータを埋め込む |
| `scripts/extract_stock.py` | 在庫表 xlsx → `data/stock.json`(標準ライブラリのみ) |
| `scripts/extract_incoming.py` | 梱包明細 .xls → `data/incoming.json`(Excel 経由で読み取り専用に開く。箱・品番・全体の合計を検算し、合わなければ書き出さない。PO 単位で追加/置換、入荷後は `--remove <PO>`) |
| `scripts/stage_raw_sales.py` | Downloads にある販売明細 Excel を、ETL が期待するファイル名で `data/raw/`(git 管理外)へ複製 |
| `scripts/sales_etl/etl*.py` | 販売ダッシュボードの ETL の複製。`patch_sales_etl.py` で入出力パスを環境変数化(それ以外は同一)。7月単月ローダーは二重計上防止のため無効化 |
| `scripts/sku_sales.py` | ETL の明細 → VFF シューズ SKU × 月 → 在庫表の SKU に照合 → `data/sku_sales.json` |
| `scripts/fetch_sales.py` | 販売ダッシュボードの公開 JSON を `data/sales_vff.json` へ(突合用) |
| `scripts/build.py` | stock.json + sku_sales.json(+ sales_vff.json)をテンプレートに埋め込み `index.html` を生成 |

## 更新手順

初回のみ: `python -m venv .venv` → `.venv\Scripts\python -m pip install openpyxl`(ETL が使う唯一の外部ライブラリ)

```bash
python scripts/extract_stock.py "C:\path\to\VFF Stock 30-09-26.xlsx"
python scripts/extract_incoming.py "C:\path\to\<Packing List>.xls"
python scripts/stage_raw_sales.py
set SALES_RAW_DIR=data\raw
set SALES_OUT_DIR=data
.venv\Scripts\python scripts\sales_etl\etl.py
.venv\Scripts\python scripts\sales_etl\etl_jul2026.py
.venv\Scripts\python scripts\sales_etl\etl_2025.py
.venv\Scripts\python scripts\sales_etl\etl_2024.py
python scripts/sku_sales.py
python scripts/build.py
```

新しい月の販売ファイルが増えたら、販売ダッシュボード側の ETL 更新を `scripts/sales_etl/` に取り込み(`patch_sales_etl.py` を再適用)、`stage_raw_sales.py` の対応表で新ファイルが拾えていることを確認してください。`index.html` を main に push すると Vercel が自動デプロイします。ローカル確認は `python -m http.server 8765`。

> 元の xlsx(在庫表・販売明細)はリポジトリに含めていません(`.gitignore`)。
> 予測・配分ロジック付きの v1(5タブ版)は git 履歴 `c724f91` にあります。
