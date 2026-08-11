# ポートフォリオ一式

副業・フリーランス受注用。本業の成果物は守秘義務で出せないため、
**実務と同じ工程で架空題材を作り直したサンプル**を実弾にしている。

## 構成

```
ポートフォリオ/
├── index.html                          ポートフォリオサイト本体
├── works/                              サイトに載せる成果物
│   ├── rotaria-01/03/08/11/13.png      営業資料サンプルのページ画像
│   └── Rotaria_サービス紹介資料_サンプル.pdf
├── 営業資料サンプル/
│   ├── build.js                        15ページを生成するスクリプト
│   ├── node_modules -> lms-pptx生成のものをsymlink
│   └── out/                            pptx / pdf / 各ページpng
└── ランサーズ提案文_営業資料デザイン.txt   応募文（【　】を埋めて使う）
```

公開URL: https://mmtrk0606.github.io/

## 残りの未記入箇所

`index.html` の中で黄色マーカー（`class="todo"`）が付いている箇所。
実績の下段3枠のみ未記入で、受注実績ができ次第そこに入れる。

画像は `works/` に置き、`<span>IMAGE 01</span>` を
`<img src="works/xxx.png" alt="">` に置き換え、`.todo` の span を外す。

## 営業資料サンプルの作り直し

```bash
cd ~/Desktop/ポートフォリオ/営業資料サンプル
node build.js

# PDF化
cd out
/Applications/LibreOffice.app/Contents/MacOS/soffice --headless \
  --convert-to pdf --outdir . "Rotaria_サービス紹介資料_サンプル.pptx"

# ページ画像（サイト用は 90dpi）
pdftoppm -png -r 90 "Rotaria_サービス紹介資料_サンプル.pdf" p
```

フォントは `游ゴシック`（Mac / Windows の両方にあるため崩れにくい）。
LibreOffice でのPDF変換時にスライド背景が落ちることがあるため、
`bg()` で全面の矩形を敷いている。**この矩形を消すと表紙が白抜けする。**

## 表示確認

納品・公開前に3幅（1440 / 768 / 390）で必ず実機確認する。
確認用スクリプトは Playwright を使用（`lms-featuredoc/node_modules` を参照）。

## 掲載時の原則

- 現職の成果物（LP・pptx・試験書・スクリーンショット）は**一切載せない**
- 架空題材のサンプルには「実在しないサービス」「数値は架空」と明記する
- 受注案件を載せる場合は、クライアントの掲載許可を取ってから
