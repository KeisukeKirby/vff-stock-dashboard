# vff-stock-dashboard
## 概要
VFF(Vibram FiveFingers)シューズの月末在庫表 xlsx を SKU(モデル×性別×カラー×サイズ)一覧にし、今年の月平均販売・在庫月数・ステータス(欠品中/要発注/適正/在庫過剰など)を付ける静的ページ。段階的に精度を上げる方針(v1 の予測付き版は git 履歴 c724f91 にある)。
## 技術スタック
Python 3 標準ライブラリのみ(xlsx は zipfile+xml で読む) / 単一 HTML、外部ライブラリなし
## コマンド
- 在庫更新: `python scripts/extract_stock.py "<xlsx path>"` → data/stock.json
- 販売更新: `python scripts/fetch_sales.py "<dashboard_data.json path>" <commit>` → data/sales_vff.json
- ビルド: `python scripts/build.py` → index.html(data を埋め込み)
- ローカル確認: `python -m http.server 8765` して http://localhost:8765/
## 触ってはいけない領域
- index.html は生成物。直接編集せず src/template.html を直して build する
- data/*.json は生成物。数値を手で直さない(直すなら元ファイル側)
## 固有の規約
- 画面は「在庫一覧」タブ(SKU 表だけ、アルファベット順)と「詳細」タブ(KPI・設定・モデル別サマリー・注記)。一覧に情報を足さない
- SKU 行の今年の月平均は在庫表の 2026 年販売列 ÷ 販売月数。直近3ヶ月はモデル別(JSON)を SKU の構成比で按分した推定(SKU 別月次は JSON に無い)
- 在庫は在庫表の Balance 列(Office 列は7月入荷分が未記入のため使わない)
- 販売実績 JSON の 6月には輸出948足が入っている(在庫表の販売列には無い)。build.py がモデル別に推定し、直近3ヶ月平均から除く
