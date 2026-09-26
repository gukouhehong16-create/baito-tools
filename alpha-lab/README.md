# Alpha Lab データ更新

米国株リサーチ用ページ「Alpha Lab」のデータを GitHub Actions で取得する仕組みです。サイト本体（main）には含まれません。

- `update.py` が Yahoo Finance・FRED・CNN・Nasdaq からデータを取り、`alpha-lab-data` ブランチに JSON で保存します。
- ページの更新ボタンが `request.json` を書き換えると、`.github/workflows/alpha-lab.yml` が動きます。
- `mode` は `quotes`（株価・チャート・マーケット指標・ニュース・経済データ・予定）、`score`、`analyst`、`all` のいずれかです。
