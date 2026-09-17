# vff-stock-dashboard
## 概要
VFF(Vibram FiveFingers)シューズの月末在庫表 xlsx を SKU(モデル×性別×カラー×サイズ)一覧にし、今年の月平均販売・在庫月数・ステータス(欠品中/要発注/適正/在庫過剰など)を付ける静的ページ。段階的に精度を上げる方針(v1 の予測付き版は git 履歴 c724f91 にある)。
## 技術スタック
Python 3(在庫表は標準ライブラリで読む。販売 ETL だけ openpyxl → `.venv`)/ 単一 HTML、外部ライブラリなし
## コマンド(順番どおり。README に詳細)
- 在庫: `python scripts/extract_stock.py "<xlsx>"` → data/stock.json
- 入荷予定: `python scripts/extract_incoming.py "<Packing List .xls>"` → data/incoming.json(PO 単位で追加/置換。月末在庫表に入荷が反映されたら `--remove <PO>`)
- 販売明細: `python scripts/stage_raw_sales.py` → data/raw/ に複製 → `SALES_RAW_DIR`/`SALES_OUT_DIR` を設定して `.venv\Scripts\python scripts\sales_etl\etl*.py`(4本)→ data/records*.json
- SKU 集計: `python scripts/sku_sales.py` → data/sku_sales.json(照合率と JSON との突合を表示)
- ビルド: `python scripts/build.py` → index.html
- ローカル確認: `python -m http.server 8765` して http://localhost:8765/
## 触ってはいけない領域
- index.html は生成物。直接編集せず src/template.html を直して build する
- data/*.json は生成物。数値を手で直さない(直すなら元ファイル側)
- scripts/sales_etl/etl*.py は販売ダッシュボードの ETL の複製。ロジックを変えない(パス行と7月単月ローダーの無効化だけが差分)。数字を直すなら販売ダッシュボード側で直して取り込む
- data/raw/・data/records*.json・.venv は git 管理外(明細は社外秘)
## 固有の規約
- 画面は「在庫一覧」タブ(SKU 表だけ、アルファベット順。列は 在庫数 / 入荷予定 / 今年の月平均 / 直近3ヶ月の月平均 / 在庫月数 / ステータス)と「詳細」タブ(KPI・設定・モデル別サマリー・注記・突合表)。Keisuke の指示なしに一覧へ列を足さない
- 入荷予定は梱包明細の SKU 別足数(在庫表の Order Import 列と照合)。在庫月数・ステータスには含めない。公開ページなので invoice 番号・コンテナ番号は埋め込まない(PO・日付・箱数・足数のみ)
- SKU 行の販売は ETL 明細の SKU × 月集計(輸出店舗 Export を除く国内分)。今年の月平均 = 2026年計 ÷ 販売月数(初入荷/初販売から)。直近3ヶ月 = 6〜8月 ÷ 3
- 在庫は在庫表の Balance 列(Office 列は7月入荷分が未記入のため使わない)
- 明細の7月は「7〜8月ファイル」(etl.py)が正。etl_jul2026.py の7月単月ローダーは二重計上になるので無効化済み
- 公開 JSON との突合: 店舗×月で一致(イベント5月 +3 のみ)。モデル別の±数足は JSON 側の7月内訳が近似だったため
