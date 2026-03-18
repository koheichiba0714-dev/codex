# r6kouchin-dashboard-sapporo

令和6年度 就労継続支援B型の工賃実績を、北海道全体比較と札幌市・小樽市の WAM 詳細をあわせて一覧・検索・集計するための静的 dashboard である。

主な機能:

- クイックプリセットで `札幌市` `札幌市・小樽市詳細` `高工賃異常` `改善優先` を即時切替
- `伸ばす候補 / 改善優先 / 確認優先` のアクションボード
- `工賃 × 稼働率` と `工賃 × 人員` の散布図
- 1事業所の深掘りパネルで、平均比、人員体制、送迎、食事加算、主活動、注視ポイントを確認
- 絞り込み結果の CSV 出力
- app 配信用データは manifest + records chunk に分割し、初回ロードを軽くしている

## データ更新

```bash
python3 /Users/chibakohei/Documents/Playground/scripts/extract_hokkaido_shuro_b_pdf.py
python3 /Users/chibakohei/Documents/Playground/scripts/fetch_wam_hokkaido_sapporo_otaru_b_details.py
python3 /Users/chibakohei/Documents/Playground/scripts/build_r6kouchinjissekib_dashboard_dataset.py \
  --dashboard-input /Users/chibakohei/Documents/Playground/data/exports/hokkaido_shuro_b/normalized/shuro_b_dashboard.json \
  --wam-details-input /Users/chibakohei/Documents/Playground/data/exports/wam/hokkaido_sapporo_otaru_shuro_b/details/hokkaido_sapporo_otaru_shuro_b_details.json \
  --wam-match-override-input /Users/chibakohei/Documents/Playground/data/inputs/wam/hokkaido_sapporo_otaru_shuro_b_match_overrides.csv \
  --link-enrichment-input /Users/chibakohei/Documents/Playground/data/inputs/web_links/hokkaido_sapporo_otaru_shuro_b_office_links.csv \
  --integrated-output /Users/chibakohei/Documents/Playground/artifacts/hokkaido_dashboard_integrated.json \
  --records-output-json /Users/chibakohei/Documents/Playground/artifacts/hokkaido_dashboard_records.json \
  --records-output-csv /Users/chibakohei/Documents/Playground/artifacts/hokkaido_dashboard_records.csv \
  --match-report-output /Users/chibakohei/Documents/Playground/artifacts/hokkaido_dashboard_match_report.csv \
  --app-output /Users/chibakohei/Documents/Playground/apps/r6kouchin-dashboard-sapporo/data/dashboard-data.json \
  --wam-focus-municipalities '札幌市,小樽市' \
  --wam-focus-label '札幌市・小樽市'
```

WAM 取得はキャッシュ付きで、`data/exports/wam/hokkaido_sapporo_otaru_shuro_b/details/by_office/` に事業所単位 JSON を保存する。

途中検証だけ行う時は以下で件数を絞れる。

```bash
python3 /Users/chibakohei/Documents/Playground/scripts/fetch_wam_hokkaido_sapporo_otaru_b_details.py --limit 20
```

## ローカル確認

```bash
cd /Users/chibakohei/Documents/Playground/apps/r6kouchin-dashboard-sapporo
python3 -m http.server 4173
```

ブラウザで `http://localhost:4173` を開く。

## Vercel 配備

```bash
cd /Users/chibakohei/Documents/Playground/apps/r6kouchin-dashboard-sapporo
npx vercel
```

本番反映する時は `npx vercel --prod` を使う。別ドメイン運用にするため、先にこのコピーを新しい Vercel project に relink してから deploy する。
