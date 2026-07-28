# バイトの計算ツール

学生アルバイト向けの計算ツールをまとめた静的サイトです。ビルドなし、フレームワークなし、単一 HTML ファイルの集合。

## 構成

```
baito-tools.pages.dev/
├─ index.html          ハブ（ツール一覧）
├─ favicon.ico / favicon.svg / apple-touch-icon.png
├─ robots.txt / sitemap.xml
├─ kyuryo/
│   └─ index.html      バイト給料計算機
└─ nenshu-kabe/
    └─ index.html      年収の壁 判定ツール
```

| ツール | URL | 内容 |
|---|---|---|
| バイト給料計算機 | `/kyuryo/` | 深夜手当・残業手当こみで支給額を計算 |
| 年収の壁 判定ツール | `/nenshu-kabe/` | 扶養を外れないラインと、あと働ける額 |

## 旧サイトからの移行

このリポジトリは、別々のドメインで公開していた2本のツールを統合したものです。旧サイトは削除せず、`_redirects` で 301 転送しています。

| 旧 URL | 転送先 |
|---|---|
| `baito-pay-calculator.pages.dev` | `/kyuryo/` |
| `nenshu-kabe-checker.pages.dev` | `/nenshu-kabe/` |

旧 Cloudflare Pages プロジェクトと旧 Search Console プロパティは**削除しないこと**。削除すると転送が消え、旧 URL が積んだ評価が失われます。

## 3本目以降の追加手順

1. `新ツール名/index.html` を作る
2. canonical と og:url を `https://baito-tools.pages.dev/新ツール名/` にする
3. ハブの `index.html` にカードを1枚追加する（`.tools` の中に `<a class="tool">` を足すだけ）
4. `sitemap.xml` に URL を1件追加する
5. 関連するツールから相互リンクを張る
6. Search Console でその1URLだけインデックス登録をリクエスト

GitHub リポジトリの作成も Cloudflare の設定も不要です。

## デプロイ

```
npx wrangler pages deploy . --project-name baito-tools --branch main
```

Cloudflare Pages の設定はフレームワークプリセット「なし」、ビルドコマンド空欄、ビルド出力ディレクトリ空欄。フレームワークを選ぶとビルドが失敗します。

## ライセンス

MIT
