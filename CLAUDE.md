# vff-stock-dashboard
## 概要
VFF(Vibram FiveFingers)シューズの在庫(月末在庫表 xlsx)と販売実績(sales-1st-half-26-dashboard の集計JSON)を突き合わせ、在庫推移予測と発注アラートを出す静的ダッシュボード。
## 技術スタック
Python 3 標準ライブラリのみ(xlsx は zipfile+xml で読む) / 単一 HTML + inline SVG(外部ライブラリなし)
## コマンド
- 在庫更新: `python scripts/extract_stock.py "<xlsx path>"` → data/stock.json
- 販売更新: `python scripts/fetch_sales.py "<dashboard_data.json path>" <commit>` → data/sales_vff.json
- ビルド: `python scripts/build.py` → index.html(data を埋め込み)
- ローカル確認: `python -m http.server 8765` して http://localhost:8765/
## 触ってはいけない領域
- index.html は生成物。直接編集せず src/template.html を直して build する
- data/*.json は生成物。数値を手で直さない(直すなら元ファイル側)
## 固有の規約
- 販売実績の正は sales_vff.json(モデル別月次)。在庫表の販売列はSKUへの配分比率にのみ使う
- 在庫表の Balance 列を在庫とする(Office 列は7月入荷分が未記入のため使わない)
- 2026-06 の輸出948足は build.py がモデル別に推定し、UI で除外できる
