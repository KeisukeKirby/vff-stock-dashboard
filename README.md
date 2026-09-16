# VFF 在庫一覧

Vibram FiveFingers(VFF)シューズの月末在庫表を **SKU(モデル × 性別 × カラー × サイズ)一覧** にし、今年の月平均販売・在庫月数・ステータスを付けた静的ページです。まずはシンプルな一覧から始め、段階的に精度を上げていきます。

- 在庫: 月末の `VFF Stock <dd-mm-yy>.xlsx`(Balance 列 = 輸入累計 − 販売累計)、入荷予定は Order Import 列
- 販売実績: [sales-1st-half-26-dashboard](https://github.com/keisukekirby/sales-1st-half-26-dashboard) の `data/dashboard_data.json`(`vff_shoes.model_monthly`)。モデル別サマリーの直近3ヶ月に使用

## 画面

- **在庫一覧タブ** — SKU(モデル × カラー × サイズ × 性別)をアルファベット順に一覧。列は 在庫数 / 今年の月平均 / 直近3ヶ月の月平均(推定) / 在庫月数 / ステータス のみ。ステータス・モデル・性別・カラー検索で絞り込み、列見出しで並び替え、CSV書き出し
- **詳細タブ** — KPI(総在庫、月平均販売、在庫月数、欠品中 / 要発注 / 在庫過剰の SKU 数)、リードタイム設定とステータス定義、モデル別サマリー(月別実績・2026年計)、データの注記

## 計算(意図的に単純)

| 項目 | 定義 |
|---|---|
| 今年の月平均 | SKU の 2026年販売 ÷ 販売月数(初入荷が2026年なら入荷月〜8月、それ以前なら 8) |
| 直近3ヶ月の月平均(推定) | モデルの直近3ヶ月(6〜8月)の月平均(販売実績JSON、6月の輸出推定分を除く)× モデル内でのその SKU の 2026年販売構成比 |
| 在庫月数 | 在庫 ÷ 今年の月平均 |
| 欠品中 | 在庫 0 かつ 今年の販売あり |
| 要発注 | 在庫月数 < リードタイム(既定 3ヶ月、画面で 2〜4 を選択) |
| 適正 | リードタイム 〜 12ヶ月 |
| 在庫過剰 | 在庫月数 > 12ヶ月 |
| 動きなし | 今年の販売 0 で在庫あり |
| 入荷待ち | 在庫 0・販売 0・発注済み(新色/新モデル) |

SKU 別の月次販売は販売実績 JSON に無いため、直近3ヶ月の SKU 値は構成比による推定です(SKU 別に出すには販売明細が必要 — 次の精度向上項目)。6月の一括輸出(948足)は在庫表の販売列に無いため、モデル別に推定して直近3ヶ月の平均から除いています(`scripts/build.py`)。

## 構成

| パス | 役割 |
|---|---|
| `index.html` | 完成ページ(生成物。データ埋め込み済みの単一ファイル) |
| `src/template.html` | ページ本体のテンプレート。`/*__DATA__*/null` に build.py がデータを埋め込む |
| `scripts/extract_stock.py` | xlsx → `data/stock.json`(標準ライブラリのみ) |
| `scripts/fetch_sales.py` | 販売ダッシュボードの JSON から VFF シューズ部分を `data/sales_vff.json` へ |
| `scripts/build.py` | 2つの JSON をテンプレートに埋め込み `index.html` を生成 |

## 更新手順

```bash
python scripts/extract_stock.py "C:\path\to\VFF Stock 30-09-26.xlsx"
python scripts/fetch_sales.py "C:\path\to\dashboard_data.json" <commit>
python scripts/build.py
```

`index.html` を main に push すると Vercel が自動デプロイします。ローカル確認は `python -m http.server 8765`。

> 元の xlsx(輸入インボイス番号を含む)はリポジトリに含めていません(`.gitignore`)。
> 予測・配分ロジック付きの v1(5タブ版)は git 履歴 `c724f91` にあります。
