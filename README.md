# VFF 在庫一覧

Vibram FiveFingers(VFF)シューズの月末在庫表を **SKU(モデル × 性別 × カラー × サイズ)一覧** にし、今年の月平均販売・在庫月数・ステータスを付けた静的ページです。まずはシンプルな一覧から始め、段階的に精度を上げていきます。

- 在庫: 月末の `VFF Stock <dd-mm-yy>.xlsx`(Balance 列 = 輸入累計 − 販売累計)、入荷予定は Order Import 列
- 販売実績: [sales-1st-half-26-dashboard](https://github.com/keisukekirby/sales-1st-half-26-dashboard) の `data/dashboard_data.json`(`vff_shoes.model_monthly`)。モデル別サマリーの直近3ヶ月に使用

## 画面

- **KPI** — 総在庫、月平均販売(2026年・在庫表)、在庫月数、欠品中 / 要発注 / 在庫過剰の SKU 数
- **モデル別サマリー** — 在庫、入荷予定、2026年販売、月平均、在庫月数、ステータス、直近3ヶ月の実績(販売実績JSON・モデル計)
- **SKU 在庫一覧** — ステータス / モデル / 性別 / カラー検索で絞り込み、列見出しで並び替え、CSV書き出し

## 計算(意図的に単純)

| 項目 | 定義 |
|---|---|
| 月平均 | SKU の 2026年販売 ÷ 販売月数(初入荷が2026年なら入荷月〜8月、それ以前なら 8) |
| 在庫月数 | 在庫 ÷ 月平均 |
| 欠品中 | 在庫 0 かつ 今年の販売あり |
| 要発注 | 在庫月数 < リードタイム(既定 3ヶ月、画面で 2〜4 を選択) |
| 適正 | リードタイム 〜 12ヶ月 |
| 在庫過剰 | 在庫月数 > 12ヶ月 |
| 動きなし | 今年の販売 0 で在庫あり |
| 入荷待ち | 在庫 0・販売 0・発注済み(新色/新モデル) |

SKU 別の月次販売は販売実績 JSON に無いため、直近3ヶ月はモデル単位でのみ表示しています(SKU 別に出すには販売明細が必要 — 次の精度向上項目)。

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
