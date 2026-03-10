# r6kouchin-dashboard

令和6年度 就労継続支援B型の工賃実績を、WAM の大阪市B型詳細と統合して一覧・検索・集計するための静的 dashboard である。

主な機能:

- クイックプリセットで `大阪市` `大阪市 × WAM一致` `高工賃異常` `改善優先` を即時切替
- `伸ばす候補 / 改善優先 / 確認優先` のアクションボード
- `工賃 × 稼働率` と `工賃 × 人員` の散布図
- 1事業所の深掘りパネルで、平均比、人員体制、送迎、食事加算、主活動、注視ポイントを確認
- 絞り込み結果の CSV 出力
- app 配信用データは manifest + records chunk に分割し、初回ロードを軽くしている

## データ更新

```bash
python3 /Users/chibakohei/Documents/Playground/scripts/extract_r6kouchinjissekib.py
python3 /Users/chibakohei/Documents/Playground/scripts/fetch_wam_osakashi_b_details.py
python3 /Users/chibakohei/Documents/Playground/scripts/build_r6kouchinjissekib_dashboard_dataset.py
```

WAM 取得はキャッシュ付きで、`data/exports/wam/osakashi_shuro_b/details/by_office/` に事業所単位 JSON を保存する。

途中検証だけ行う時は以下で件数を絞れる。

```bash
python3 /Users/chibakohei/Documents/Playground/scripts/fetch_wam_osakashi_b_details.py --limit 20
```

## ローカル確認

```bash
cd /Users/chibakohei/Documents/Playground/apps/r6kouchin-dashboard
python3 -m http.server 4173
```

ブラウザで `http://localhost:4173` を開く。

## Vercel 配備

```bash
cd /Users/chibakohei/Documents/Playground/apps/r6kouchin-dashboard
npx vercel
```

本番反映する時は `npx vercel --prod` を使う。
