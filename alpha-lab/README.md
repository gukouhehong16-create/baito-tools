# Alpha Lab データ更新

米国株リサーチ用ページ「Alpha Lab」のデータを、Claude を使わずに取得する仕組みです。サイト本体（main）には含まれません。

## Google Apps Script（ページが使う方式）

`apps-script/Code.gs` を自分の Google アカウントの Apps Script に貼り付けて動かします。

1. https://script.google.com で新しいプロジェクトを作り、`Code.gs` の中身を全部貼り付けて保存する
2. 関数「setup」を実行し、権限を許可する

以後は1分ごとに自動で動き、次のことをします。

- ページの更新ボタンが Google ドライブの「Alpha Lab データ」フォルダに置いた依頼ファイル（`alpha-lab-request.json`）を読み、その種類の更新を実行する
- 火〜土の朝（日本時間 6:40〜7:00）に、株価・スコア・アナリスト評価を自動で更新する
- 結果を同じフォルダの `data/` に gzip＋base64 の JSON で保存し、`alpha-lab-status.json` に一覧とハッシュを書く

ページは Google Drive コネクタで `alpha-lab-status.json` を読み、変わったデータだけをページのデータベースへ写します。

アナリスト評価の更新では、決算日・EPS予想の修正・決算サプライズ（`meta/earn`）と、今後2週間の決算と米国の経済指標の予定（`meta/cal`）も Yahoo Finance から取ります。

メール通知: ページの「メール通知」の設定とウォッチリストは `alpha-lab-config.json` としてドライブに置かれ、Apps Script が読みます。
毎朝の更新のあとと、米国市場の取引時間中は15分ごとに、価格アラート・大きな値動き・52週高値/安値・格上げ/格下げ・目標株価の変化・決算の前日・EPS予想の修正を判定し、まとめて1通のメールで知らせます。
依頼ファイルの `add` に書かれた銘柄（ページに新しく加わった銘柄）は、以後の更新対象に加わります。

## GitHub Actions（予備）

- `update.py` が Yahoo Finance・FRED・CNN・Nasdaq からデータを取り、`alpha-lab-data` ブランチに JSON で保存します（Apps Script 版と同じ形式）。
- `request.json` を書き換えると `.github/workflows/alpha-lab.yml` が動きます。
- `mode` は `quotes`（株価・チャート・マーケット指標・ニュース・経済データ・予定）、`score`、`analyst`、`all` のいずれかです。
- Apps Script は初回だけ、このブランチの `quotes.json` を銘柄一覧として読みます。
