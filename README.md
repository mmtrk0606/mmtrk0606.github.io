# ポートフォリオ一式

副業・フリーランス受注用。本業の成果物は守秘義務で出せないため、
**実務と同じ工程で架空題材を作り直したサンプル**を実弾にしている。

## 構成

```
ポートフォリオ/
├── index.html                          ポートフォリオサイト本体
├── works/                              サイトに載せる成果物
│   ├── rotaria-01/03/08/11/13.png      営業資料サンプルのページ画像
│   ├── rezerva-cover/list/views/screen.png  不具合レポートのページ画像
│   ├── work-testplan/manual/site.png   実績カードのサムネイル（16:10）
│   └── *.pdf                           各サンプルの配布用PDF
├── 営業資料サンプル/
│   ├── build.js                        15ページを生成するスクリプト
│   ├── node_modules -> lms-pptx生成のものをsymlink
│   └── out/                            pptx / pdf / 各ページpng
├── 検証サンプル/
│   ├── mock/app.html                   架空システム Rezerva のモック画面
│   ├── build_report.py                 不具合レポート（10件）を生成
│   ├── build_testplan.py               テスト設計書（72ケース）を生成
│   ├── build_manual.py                 操作マニュアル（11ページ）を生成
│   ├── shots/                          不具合の証跡スクリーンショット
│   ├── manual_shots/                   マニュアル用の画面キャプチャ
│   └── out/                            xlsx / docx / pdf
└── ランサーズ提案文_営業資料デザイン.txt   応募文（【　】を埋めて使う）
```

公開URL: https://mmtrk0606.github.io/

## 実績セクションの構成

主役2本（`.feature`）＋下段カード3枚（`.work`）。
下段カードは**サムネイル＋1〜2行＋リンク先**だけに留める。
ページを縦に伸ばさないための制約なので、ここに長文を足さない。

受注実績ができたらカードを差し替える。未記入のまま置くときは
`class="todo"`（黄色マーカー）を付けておくと、公開前に気付ける。

## 検証サンプルの作り直し

```bash
cd ~/Desktop/ポートフォリオ/検証サンプル
python3 build_report.py      # 不具合レポート  → out/*.xlsx
python3 build_testplan.py    # テスト設計書    → out/*.xlsx
python3 build_manual.py      # 操作マニュアル  → out/*.docx

cd out
/Applications/LibreOffice.app/Contents/MacOS/soffice --headless \
  --convert-to pdf --outdir . *.xlsx *.docx
```

`mock/app.html` はクエリで画面を切り替える。
`?v=list` 系は**不具合の再現用**（重複予約や 16:00〜14:00 など、わざと壊した状態）、
`?v=*_ok` と `?v=login` `?v=done` が**マニュアル用の正常な状態**。
マニュアルに再現用の画面を使わないこと。

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
