// ポートフォリオ用 営業資料サンプル — 架空のBtoB SaaS「Rotaria」／全15スライド
// ※実在しないサービスです。数値・事例はすべてサンプルとして作成した架空のものです。
const P = require("pptxgenjs");
const prs = new P();
prs.defineLayout({ name: "WIDE169", width: 13.33, height: 7.5 });
prs.layout = "WIDE169";
prs.author = "河村";
prs.title = "Rotaria サービス紹介資料（制作サンプル）";

const FONT = "游ゴシック";
const C = {
  main:    "0F4C5C",  // ディープティール
  dark:    "093642",
  accent:  "E8763C",  // オレンジ
  accentBg:"FBE8DC",
  ink:     "16262B",
  muted:   "7A8C92",
  line:    "DCE6E6",
  bgSub:   "F3F7F7",
  white:   "FFFFFF",
  green:   "2E8B70",
  red:     "C4564B",
  amber:   "D9A441",
  ghost:   "AFC3C7",
};
const SW = 13.33, SH = 7.5;
const ML = 0.78, MR = 0.78;          // 左右マージン
const CW = SW - ML - MR;             // 本文幅 11.77
const FOOT_Y = 6.86;

let PAGE = 0;
const TOTAL = 15;

/* ============================================================
   共通パーツ
   ============================================================ */
// 背景。LibreOffice でのPDF変換時に slide background が落ちることがあるため、
// 全面の矩形を先に敷いて確実に塗る。
function bg(s, color) {
  const col = color || C.white;
  s.background = { color: col };
  s.addShape("rect", { x: 0, y: 0, w: SW, h: SH, fill: { color: col }, line: { type: "none" } });
}

// 中面の見出し（ラベル＋1メッセージ）
function head(s, label, message, sub) {
  s.addShape("rect", { x: ML, y: 0.62, w: 0.055, h: 0.28, fill: { color: C.accent }, line: { type: "none" } });
  s.addText(label, {
    x: ML + 0.17, y: 0.58, w: 8, h: 0.36, fontFace: FONT, fontSize: 11.5, bold: true,
    color: C.accent, charSpacing: 1.4, valign: "middle",
  });
  s.addText(message, {
    x: ML, y: 1.06, w: CW, h: 0.72, fontFace: FONT, fontSize: 25, bold: true,
    color: C.main, valign: "middle", lineSpacing: 32,
  });
  if (sub) {
    s.addText(sub, {
      x: ML, y: 1.82, w: CW, h: 0.34, fontFace: FONT, fontSize: 13, color: C.muted, valign: "middle",
    });
  }
}

function footer(s, opts) {
  PAGE++;
  const o = opts || {};
  const col = o.onDark ? C.ghost : C.muted;
  if (!o.onDark) {
    s.addShape("rect", { x: ML, y: FOOT_Y, w: CW, h: 0.012, fill: { color: C.line }, line: { type: "none" } });
  }
  s.addText("Rotaria ／ サービス紹介資料", {
    x: ML, y: FOOT_Y + 0.06, w: 7, h: 0.32, fontFace: FONT, fontSize: 9.5, color: col, valign: "middle",
  });
  s.addText(String(PAGE).padStart(2, "0"), {
    x: SW - MR - 2, y: FOOT_Y + 0.06, w: 2, h: 0.32, fontFace: FONT, fontSize: 9.5,
    color: col, align: "right", valign: "middle",
  });
}

// カード（枠＋塗り）
function card(s, x, y, w, h, opts) {
  const o = opts || {};
  s.addShape("roundRect", {
    x, y, w, h, rectRadius: o.radius === undefined ? 0.06 : o.radius,
    fill: { color: o.fill || C.bgSub },
    line: o.line === false ? { type: "none" } : { color: o.lineColor || C.line, width: 1 },
  });
}

/* ============================================================
   01 表紙
   ============================================================ */
{
  const s = prs.addSlide();
  bg(s, C.main);
  // 右側の色面 ＋ シフト表を抽象化したグリッド
  s.addShape("rect", { x: 8.05, y: 0, w: SW - 8.05, h: SH, fill: { color: C.dark }, line: { type: "none" } });
  s.addShape("rect", { x: 0, y: 0, w: 0.14, h: SH, fill: { color: C.accent }, line: { type: "none" } });
  {
    const gx = 8.75, gy = 1.62, cw = 0.62, ch = 0.62, cols = 6, rows = 6, gap = 0.09;
    const filled = { "1-1": 1, "2-3": 1, "3-0": 1, "0-4": 1, "4-2": 1, "5-5": 1, "2-4": 2, "4-5": 2, "1-2": 2 };
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const key = `${r}-${c}`;
        const kind = filled[key];
        s.addShape("rect", {
          x: gx + c * (cw + gap), y: gy + r * (ch + gap), w: cw, h: ch,
          fill: kind === 1 ? { color: C.main } : kind === 2 ? { color: C.accent } : { type: "none" },
          line: kind ? { type: "none" } : { color: "17515F", width: 0.75 },
        });
      }
    }
  }

  s.addText("Rotaria", {
    x: 1.05, y: 2.02, w: 8, h: 0.98, fontFace: FONT, fontSize: 46, bold: true, color: C.white, charSpacing: 1,
  });
  s.addShape("rect", { x: 1.08, y: 3.06, w: 1.5, h: 0.045, fill: { color: C.accent }, line: { type: "none" } });
  s.addText("多店舗のシフト作成を、\n30分から3分へ。", {
    x: 1.05, y: 3.34, w: 6.6, h: 1.34, fontFace: FONT, fontSize: 26, bold: true, color: C.white, lineSpacing: 40,
  });
  s.addText("シフト管理クラウド ／ サービス紹介資料", {
    x: 1.05, y: 4.86, w: 6.6, h: 0.36, fontFace: FONT, fontSize: 14, color: C.ghost,
  });
  s.addText("2026.08　株式会社ロタリア", {
    x: 1.05, y: 5.9, w: 6.6, h: 0.32, fontFace: FONT, fontSize: 11.5, color: C.ghost,
  });
  s.addText("※本資料は制作サンプルです。Rotaria は実在しないサービスであり、\n　掲載の数値・事例はすべて架空のものです。", {
    x: 1.05, y: 6.52, w: 6.6, h: 0.56, fontFace: FONT, fontSize: 9, color: "6E8B93", lineSpacing: 14,
  });
  PAGE++;
}

/* ============================================================
   02 目次
   ============================================================ */
{
  const s = prs.addSlide();
  bg(s);
  head(s, "CONTENTS", "本日ご説明する内容");

  const items = [
    ["01", "シフト作成の現状", "店長が抱えている作業と、その原因"],
    ["02", "Rotaria のご紹介", "解決する範囲と、4つの機能"],
    ["03", "導入した効果", "作業時間・人件費・定着率の変化"],
    ["04", "料金と導入の流れ", "プラン、導入までのスケジュール"],
  ];
  let y = 2.42;
  items.forEach(([no, t, d]) => {
    s.addShape("rect", { x: ML, y: y + 0.02, w: CW, h: 0.012, fill: { color: C.line }, line: { type: "none" } });
    s.addText(no, {
      x: ML, y: y + 0.16, w: 0.9, h: 0.62, fontFace: FONT, fontSize: 22, bold: true, color: C.accent, valign: "middle",
    });
    s.addText(t, {
      x: ML + 0.95, y: y + 0.16, w: 4.4, h: 0.62, fontFace: FONT, fontSize: 17, bold: true, color: C.ink, valign: "middle",
    });
    s.addText(d, {
      x: ML + 5.4, y: y + 0.16, w: CW - 5.4, h: 0.62, fontFace: FONT, fontSize: 12.5, color: C.muted, valign: "middle",
    });
    y += 1.02;
  });
  s.addShape("rect", { x: ML, y: y + 0.02, w: CW, h: 0.012, fill: { color: C.line }, line: { type: "none" } });
  footer(s);
}

/* ============================================================
   03 課題
   ============================================================ */
{
  const s = prs.addSlide();
  bg(s);
  head(s, "01　シフト作成の現状", "シフト作成は、店長の「見えない残業」になっています。",
       "複数店舗を運営する企業の店長が、シフト業務に費やしている時間はひと月あたり平均9.5時間。");

  const cols = [
    ["希望の回収が終わらない", "LINE・紙・口頭がばらばらに届き、\n締切を過ぎてからの変更も入る。", "回収に平均 3.2h"],
    ["組み替えが連鎖する", "1人の変更が全体に波及し、\n完成間際にやり直しになる。", "調整に平均 4.1h"],
    ["共有後も問い合わせが続く", "「自分は何時から？」の連絡に\n個別対応が発生する。", "共有後に平均 2.2h"],
  ];
  const cw = (CW - 0.5) / 3;
  cols.forEach(([t, d, m], i) => {
    const x = ML + i * (cw + 0.25);
    card(s, x, 2.62, cw, 3.32);
    s.addShape("rect", { x, y: 2.62, w: cw, h: 0.05, fill: { color: C.accent }, line: { type: "none" } });
    s.addText(String(i + 1).padStart(2, "0"), {
      x: x + 0.34, y: 2.9, w: 1.2, h: 0.42, fontFace: FONT, fontSize: 15, bold: true, color: C.accent, valign: "middle",
    });
    s.addText(t, {
      x: x + 0.34, y: 3.36, w: cw - 0.68, h: 0.5, fontFace: FONT, fontSize: 16, bold: true, color: C.ink, valign: "top",
    });
    s.addText(d, {
      x: x + 0.34, y: 3.94, w: cw - 0.68, h: 1.0, fontFace: FONT, fontSize: 12.5, color: C.muted, lineSpacing: 21,
    });
    s.addShape("rect", { x: x + 0.34, y: 5.06, w: cw - 0.68, h: 0.012, fill: { color: C.line }, line: { type: "none" } });
    s.addText(m, {
      x: x + 0.34, y: 5.2, w: cw - 0.68, h: 0.5, fontFace: FONT, fontSize: 14, bold: true, color: C.main, valign: "middle",
    });
  });
  s.addText("＊自社調べ（架空）／ 3店舗以上を運営する企業の店長 120名を対象にした想定値", {
    x: ML, y: 6.12, w: CW, h: 0.3, fontFace: FONT, fontSize: 9.5, color: C.muted,
  });
  footer(s);
}

/* ============================================================
   04 原因の構造
   ============================================================ */
{
  const s = prs.addSlide();
  bg(s);
  head(s, "01　シフト作成の現状", "原因は、3つの工程が分断していることにあります。",
       "使う道具が工程ごとに違うため、情報が引き継がれず、毎回まとめ直す作業が発生しています。");

  // 3工程 → 分断のビジュアル
  const steps = [
    ["希望を集める", "LINE・紙・口頭", "AB2C33"],
    ["シフトを組む", "Excel（店長の手作業）", "AB2C33"],
    ["共有する", "紙の掲示・個別連絡", "AB2C33"],
  ];
  const bw = 3.1, gap = 1.28;
  const startX = ML + 0.35;
  steps.forEach(([t, tool], i) => {
    const x = startX + i * (bw + gap);
    card(s, x, 2.72, bw, 1.72, { fill: C.white, lineColor: C.line });
    s.addText(t, {
      x: x + 0.2, y: 2.92, w: bw - 0.4, h: 0.46, fontFace: FONT, fontSize: 16, bold: true, color: C.ink, align: "center", valign: "middle",
    });
    s.addText(tool, {
      x: x + 0.2, y: 3.44, w: bw - 0.4, h: 0.78, fontFace: FONT, fontSize: 12, color: C.muted, align: "center", valign: "top",
    });
    // 分断マーク（工程間の×）
    if (i < 2) {
      const mx = x + bw + gap / 2;
      s.addShape("line", { x: x + bw + 0.16, y: 3.58, w: gap - 0.32, h: 0, line: { color: C.line, width: 1.5, dashType: "dash" } });
      s.addShape("ellipse", { x: mx - 0.24, y: 3.34, w: 0.48, h: 0.48, fill: { color: C.white }, line: { color: C.red, width: 1.5 } });
      s.addText("×", {
        x: mx - 0.24, y: 3.34, w: 0.48, h: 0.48, fontFace: FONT, fontSize: 15, bold: true, color: C.red, align: "center", valign: "middle",
      });
      s.addText("転記", {
        x: mx - 0.5, y: 3.86, w: 1.0, h: 0.28, fontFace: FONT, fontSize: 10.5, color: C.red, align: "center", valign: "middle",
      });
    }
  });

  // 帰結
  s.addShape("rect", { x: ML, y: 4.98, w: CW, h: 0.012, fill: { color: C.line }, line: { type: "none" } });
  card(s, ML, 5.24, CW, 1.04, { fill: C.accentBg, line: false });
  s.addText("工程をまたぐたびに人が転記しているため、店舗が増えるほど作業時間が比例して増えていきます。", {
    x: ML + 0.42, y: 5.24, w: CW - 0.84, h: 1.04, fontFace: FONT, fontSize: 15, bold: true, color: "9A4A1E", valign: "middle",
  });
  footer(s);
}

/* ============================================================
   05 サービス概要
   ============================================================ */
{
  const s = prs.addSlide();
  bg(s);
  head(s, "02　Rotaria のご紹介", "Rotaria は、3つの工程をひとつにまとめるサービスです。",
       "希望の回収から共有までを1つの画面で完結させ、店長の作業を「確認するだけ」に変えます。");

  // 中央帯：3工程が1本につながる
  const barY = 2.78, barH = 1.5;
  card(s, ML, barY, CW, barH, { fill: C.main, line: false, radius: 0.08 });
  const seg = ["希望を集める", "シフトを組む", "共有する"];
  const segW = CW / 3;
  seg.forEach((t, i) => {
    const x = ML + i * segW;
    if (i > 0) {
      s.addShape("rect", { x, y: barY + 0.3, w: 0.012, h: barH - 0.6, fill: { color: "2C6675" }, line: { type: "none" } });
    }
    s.addText(t, {
      x, y: barY, w: segW, h: barH, fontFace: FONT, fontSize: 17, bold: true, color: C.white, align: "center", valign: "middle",
    });
  });
  s.addShape("rect", { x: ML, y: barY, w: CW, h: 0.06, fill: { color: C.accent }, line: { type: "none" } });
  s.addText("Rotaria がまとめて担当する範囲", {
    x: ML, y: barY + barH + 0.12, w: CW, h: 0.32, fontFace: FONT, fontSize: 11.5, color: C.accent, align: "center", bold: true,
  });

  // 4機能
  const fns = [
    ["01", "希望のオンライン収集", "スマホから提出、自動で集計"],
    ["02", "シフトの自動作成", "条件を満たす案を1クリックで"],
    ["03", "人件費の可視化", "組みながら予算と照合"],
    ["04", "法令チェック", "違反する組み方を自動で検知"],
  ];
  const fw = (CW - 0.66) / 4;
  fns.forEach(([no, t, d], i) => {
    const x = ML + i * (fw + 0.22);
    card(s, x, 4.92, fw, 1.42, { fill: C.white });
    s.addText(no, { x: x + 0.22, y: 5.06, w: 1, h: 0.3, fontFace: FONT, fontSize: 11, bold: true, color: C.accent, valign: "middle" });
    s.addText(t, { x: x + 0.22, y: 5.36, w: fw - 0.44, h: 0.4, fontFace: FONT, fontSize: 13.5, bold: true, color: C.ink, valign: "middle" });
    s.addText(d, { x: x + 0.22, y: 5.76, w: fw - 0.44, h: 0.44, fontFace: FONT, fontSize: 11, color: C.muted, valign: "top" });
  });
  footer(s);
}

/* ============================================================
   06 Before / After
   ============================================================ */
{
  const s = prs.addSlide();
  bg(s);
  head(s, "02　Rotaria のご紹介", "店長の仕事は、「組む」から「確認する」に変わります。");

  const colW = (CW - 1.1) / 2;
  // Before
  card(s, ML, 2.42, colW, 3.9, { fill: C.bgSub, line: false });
  s.addText("導入前", { x: ML + 0.34, y: 2.62, w: 2.4, h: 0.44, fontFace: FONT, fontSize: 13, bold: true, color: C.muted, valign: "middle" });
  s.addText("月 9.5 時間", {
    x: ML + colW - 3.14, y: 2.62, w: 2.8, h: 0.44, fontFace: FONT, fontSize: 18, bold: true,
    color: C.muted, align: "right", valign: "middle",
  });
  s.addShape("rect", { x: ML + 0.34, y: 3.2, w: colW - 0.68, h: 0.012, fill: { color: "D3DCDC" }, line: { type: "none" } });
  const before = [
    "LINE・紙で届いた希望を1件ずつ転記する",
    "Excelで組み、崩れるたびに組み直す",
    "人件費は月末の集計まで分からない",
    "労働時間の上限は目視で確認する",
    "完成後、個別の問い合わせに対応する",
  ];
  before.forEach((t, i) => {
    const y = 3.42 + i * 0.58;
    s.addShape("rect", { x: ML + 0.34, y: y + 0.19, w: 0.16, h: 0.02, fill: { color: C.muted }, line: { type: "none" } });
    s.addText(t, { x: ML + 0.64, y, w: colW - 1.0, h: 0.44, fontFace: FONT, fontSize: 12.5, color: C.ink, valign: "middle" });
  });

  // 矢印
  const ax = ML + colW + 0.18;
  s.addShape("rightArrow", { x: ax, y: 4.06, w: 0.74, h: 0.56, fill: { color: C.accent }, line: { type: "none" } });

  // After
  const rx = ML + colW + 1.1;
  card(s, rx, 2.42, colW, 3.9, { fill: C.white, lineColor: C.main });
  s.addShape("rect", { x: rx, y: 2.42, w: colW, h: 0.06, fill: { color: C.main }, line: { type: "none" } });
  s.addText("Rotaria 導入後", { x: rx + 0.34, y: 2.62, w: 2.8, h: 0.44, fontFace: FONT, fontSize: 13, bold: true, color: C.main, valign: "middle" });
  s.addText("月 2.6 時間", {
    x: rx + colW - 3.14, y: 2.62, w: 2.8, h: 0.44, fontFace: FONT, fontSize: 18, bold: true,
    color: C.accent, align: "right", valign: "middle",
  });
  s.addShape("rect", { x: rx + 0.34, y: 3.2, w: colW - 0.68, h: 0.012, fill: { color: C.line }, line: { type: "none" } });
  const after = [
    "希望はスマホから提出、集計は自動",
    "条件を満たす案が1クリックで出る",
    "人件費は組みながらリアルタイムに表示",
    "上限を超える組み方は自動で警告",
    "確定と同時に全員のスマホへ配信",
  ];
  after.forEach((t, i) => {
    const y = 3.42 + i * 0.58;
    s.addText("✓", { x: rx + 0.3, y, w: 0.34, h: 0.44, fontFace: FONT, fontSize: 12.5, bold: true, color: C.green, valign: "middle" });
    s.addText(t, { x: rx + 0.64, y, w: colW - 1.0, h: 0.44, fontFace: FONT, fontSize: 12.5, color: C.ink, valign: "middle" });
  });
  footer(s);
}

/* ============================================================
   07 機能01 希望収集
   ============================================================ */
{
  const s = prs.addSlide();
  bg(s);
  head(s, "機能 01 ／ 希望のオンライン収集", "希望はスマホから。集計は自動で終わります。",
       "スタッフは空き時間を選んで送信するだけ。締切と未提出者の管理も画面上で完結します。");

  // 左：スマホ枠
  const px = ML + 0.5, py = 2.66, pw = 2.5, ph = 3.6;
  card(s, px, py, pw, ph, { fill: C.white, lineColor: C.main, radius: 0.16 });
  s.addShape("rect", { x: px + 0.85, y: py + 0.14, w: 0.8, h: 0.07, fill: { color: C.line }, line: { type: "none" } });
  s.addText("8月の希望を送る", {
    x: px + 0.2, y: py + 0.36, w: pw - 0.4, h: 0.34, fontFace: FONT, fontSize: 12, bold: true, color: C.main, valign: "middle",
  });
  const days = [["8/1（金）", "10:00-15:00", true], ["8/2（土）", "終日OK", true], ["8/3（日）", "不可", false], ["8/4（月）", "17:00-22:00", true]];
  days.forEach(([d, t, ok], i) => {
    const y = py + 0.82 + i * 0.6;
    card(s, px + 0.2, y, pw - 0.4, 0.5, { fill: ok ? C.bgSub : C.white, lineColor: C.line, radius: 0.04 });
    s.addText(d, { x: px + 0.32, y, w: 1.0, h: 0.5, fontFace: FONT, fontSize: 9.5, color: C.ink, valign: "middle" });
    s.addText(t, { x: px + 1.25, y, w: 1.0, h: 0.5, fontFace: FONT, fontSize: 9.5, color: ok ? C.green : C.muted, align: "right", valign: "middle" });
  });
  card(s, px + 0.2, py + 3.06, pw - 0.4, 0.42, { fill: C.accent, line: false, radius: 0.04 });
  s.addText("送信する", {
    x: px + 0.2, y: py + 3.06, w: pw - 0.4, h: 0.42, fontFace: FONT, fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle",
  });
  s.addText("スタッフの画面", { x: px, y: py + ph + 0.1, w: pw, h: 0.3, fontFace: FONT, fontSize: 10.5, color: C.muted, align: "center" });

  // 矢印
  s.addShape("rightArrow", { x: px + pw + 0.42, y: 4.2, w: 0.62, h: 0.5, fill: { color: C.accent }, line: { type: "none" } });

  // 右：集計ボード
  const bx = px + pw + 1.28, bw = SW - MR - bx, by = 2.66;
  card(s, bx, by, bw, 3.6, { fill: C.white, lineColor: C.line });
  s.addShape("rect", { x: bx, y: by, w: bw, h: 0.52, fill: { color: C.main }, line: { type: "none" } });
  s.addText("提出状況　8月シフト（締切 7/25）", {
    x: bx + 0.28, y: by, w: bw - 0.56, h: 0.52, fontFace: FONT, fontSize: 12.5, bold: true, color: C.white, valign: "middle",
  });
  const rows = [
    ["田中", "提出済", true], ["佐藤", "提出済", true], ["鈴木", "未提出", false],
    ["高橋", "提出済", true], ["伊藤", "提出済", true], ["渡辺", "未提出", false],
  ];
  const colw = (bw - 0.56) / 2;
  rows.forEach(([n, st, ok], i) => {
    const cx = bx + 0.28 + (i % 2) * colw;
    const y = by + 0.78 + Math.floor(i / 2) * 0.58;
    s.addText(n, { x: cx, y, w: 1.4, h: 0.44, fontFace: FONT, fontSize: 12, color: C.ink, valign: "middle" });
    card(s, cx + 1.4, y + 0.05, 1.16, 0.34, { fill: ok ? "E6F2ED" : C.accentBg, line: false, radius: 0.03 });
    s.addText(st, {
      x: cx + 1.4, y: y + 0.05, w: 1.16, h: 0.34, fontFace: FONT, fontSize: 10, bold: true,
      color: ok ? C.green : C.accent, align: "center", valign: "middle",
    });
  });
  s.addShape("rect", { x: bx + 0.28, y: by + 2.58, w: bw - 0.56, h: 0.012, fill: { color: C.line }, line: { type: "none" } });
  s.addText("未提出の2名へ、ワンクリックでまとめて催促できます。", {
    x: bx + 0.28, y: by + 2.74, w: bw - 0.56, h: 0.68, fontFace: FONT, fontSize: 12.5, color: C.ink, valign: "middle",
  });
  s.addText("店長の画面", { x: bx, y: by + 3.7, w: bw, h: 0.3, fontFace: FONT, fontSize: 10.5, color: C.muted, align: "center" });
  footer(s);
}

/* ============================================================
   08 機能02 自動作成
   ============================================================ */
{
  const s = prs.addSlide();
  bg(s);
  head(s, "機能 02 ／ シフトの自動作成", "1クリックで、条件を満たしたシフト案が出ます。",
       "必要人数・スキル・希望・法令の4条件を同時に満たす案を自動で作成。手直しもドラッグで行えます。");

  // 条件チップ
  const conds = ["必要人数", "スキル要件", "本人の希望", "労働時間の上限"];
  const chw = 2.16;
  conds.forEach((t, i) => {
    const x = ML + i * (chw + 0.2);
    card(s, x, 2.56, chw, 0.5, { fill: C.bgSub, line: false, radius: 0.25 });
    s.addText(t, { x, y: 2.56, w: chw, h: 0.5, fontFace: FONT, fontSize: 11.5, bold: true, color: C.main, align: "center", valign: "middle" });
  });
  s.addText("を同時に満たす", {
    x: ML + 4 * (chw + 0.2), y: 2.56, w: 2.4, h: 0.5, fontFace: FONT, fontSize: 11.5, color: C.muted, valign: "middle",
  });

  // シフト表
  const tx = ML, ty = 3.3, tw = CW;
  const names = ["田中", "佐藤", "鈴木", "高橋"];
  const dayLabels = ["月", "火", "水", "木", "金", "土", "日"];
  const nameW = 1.05, cellW = (tw - nameW) / 7, cellH = 0.62;

  // ヘッダ行
  s.addShape("rect", { x: tx, y: ty, w: tw, h: 0.44, fill: { color: C.main }, line: { type: "none" } });
  dayLabels.forEach((d, i) => {
    s.addText(d, {
      x: tx + nameW + i * cellW, y: ty, w: cellW, h: 0.44, fontFace: FONT, fontSize: 11.5, bold: true,
      color: i >= 5 ? "FFD9C2" : C.white, align: "center", valign: "middle",
    });
  });
  // 本体
  const plan = [
    ["A", "A", "", "B", "B", "A", ""],
    ["B", "", "B", "A", "A", "", "A"],
    ["", "B", "A", "", "B", "B", "A"],
    ["A", "A", "B", "B", "", "A", "B"],
  ];
  names.forEach((n, r) => {
    const y = ty + 0.44 + r * cellH;
    s.addShape("rect", { x: tx, y, w: tw, h: cellH, fill: { color: r % 2 ? C.bgSub : C.white }, line: { color: C.line, width: 0.75 } });
    s.addText(n, { x: tx + 0.18, y, w: nameW, h: cellH, fontFace: FONT, fontSize: 11.5, color: C.ink, valign: "middle" });
    plan[r].forEach((v, c) => {
      const cx = tx + nameW + c * cellW;
      s.addShape("rect", { x: cx, y, w: cellW, h: cellH, fill: { type: "none" }, line: { color: C.line, width: 0.75 } });
      if (v) {
        const isA = v === "A";
        card(s, cx + 0.12, y + 0.11, cellW - 0.24, cellH - 0.22, {
          fill: isA ? "DDEAEC" : C.accentBg, line: false, radius: 0.03,
        });
        s.addText(isA ? "早番" : "遅番", {
          x: cx + 0.12, y: y + 0.11, w: cellW - 0.24, h: cellH - 0.22, fontFace: FONT, fontSize: 10,
          bold: true, color: isA ? C.main : "A85224", align: "center", valign: "middle",
        });
      }
    });
  });
  s.addText("希望と必要人数を満たした案。手直しはセルのドラッグで反映されます。", {
    x: ML, y: ty + 0.44 + 4 * cellH + 0.16, w: CW, h: 0.36, fontFace: FONT, fontSize: 11.5, color: C.muted,
  });
  footer(s);
}

/* ============================================================
   09 機能03 人件費
   ============================================================ */
{
  const s = prs.addSlide();
  bg(s);
  head(s, "機能 03 ／ 人件費の可視化", "人件費は、組みながらリアルタイムに分かります。",
       "月末の集計を待たずに予算との差が見えるため、超過する前に調整できます。");

  // 棒グラフ（図形で描画）
  const gx = ML + 0.3, gy = 2.72, gh = 2.9, gw = 7.2;
  const budget = 0.78; // 予算ラインの高さ比
  const bars = [
    ["第1週", 0.62, false], ["第2週", 0.71, false], ["第3週", 0.86, true], ["第4週", 0.68, false],
  ];
  // 軸
  s.addShape("rect", { x: gx, y: gy + gh, w: gw, h: 0.012, fill: { color: C.line }, line: { type: "none" } });
  // 予算ライン
  const by = gy + gh - gh * budget;
  s.addShape("line", { x: gx, y: by, w: gw, h: 0, line: { color: C.red, width: 1.5, dashType: "dash" } });
  s.addText("予算ライン", {
    x: gx + gw - 1.7, y: by - 0.34, w: 1.7, h: 0.3, fontFace: FONT, fontSize: 10.5, bold: true, color: C.red, align: "right", valign: "middle",
  });
  const barW = 1.02, step = gw / 4;
  bars.forEach(([label, ratio, over], i) => {
    const h = gh * ratio;
    const x = gx + i * step + (step - barW) / 2;
    s.addShape("rect", {
      x, y: gy + gh - h, w: barW, h, fill: { color: over ? C.red : C.main }, line: { type: "none" },
    });
    s.addText(label, {
      x: x - 0.3, y: gy + gh + 0.08, w: barW + 0.6, h: 0.3, fontFace: FONT, fontSize: 11, color: C.muted, align: "center",
    });
    if (over) {
      card(s, x - 0.36, gy + gh - h - 0.52, barW + 0.72, 0.42, { fill: C.red, line: false, radius: 0.04 });
      s.addText("超過 +4.2%", {
        x: x - 0.36, y: gy + gh - h - 0.52, w: barW + 0.72, h: 0.42, fontFace: FONT, fontSize: 10.5,
        bold: true, color: C.white, align: "center", valign: "middle",
      });
    }
  });

  // 右：説明
  const rx = ML + 8.1, rw = SW - MR - rx;
  card(s, rx, 2.72, rw, 3.3, { fill: C.bgSub, line: false });
  s.addText("超過する前に気づけます", {
    x: rx + 0.34, y: 3.0, w: rw - 0.68, h: 0.44, fontFace: FONT, fontSize: 15, bold: true, color: C.main, valign: "middle",
  });
  const pts = [
    "シフトを組んだ時点で当月の人件費を試算",
    "予算を超える週は画面上で警告",
    "店舗別・月別に実績を蓄積し、翌月の計画に反映",
  ];
  pts.forEach((t, i) => {
    const y = 3.62 + i * 0.78;
    s.addShape("rect", { x: rx + 0.34, y: y + 0.21, w: 0.14, h: 0.02, fill: { color: C.accent }, line: { type: "none" } });
    s.addText(t, { x: rx + 0.6, y, w: rw - 0.94, h: 0.66, fontFace: FONT, fontSize: 12, color: C.ink, valign: "middle", lineSpacing: 19 });
  });
  footer(s);
}

/* ============================================================
   10 機能04 法令チェック
   ============================================================ */
{
  const s = prs.addSlide();
  bg(s);
  head(s, "機能 04 ／ 法令チェック", "法令に触れる組み方は、その場で止まります。",
       "確定前に自動で検査します。ルールは企業ごとの就業規則に合わせて設定できます。");

  // 警告カード
  card(s, ML, 2.62, CW, 1.16, { fill: "FCEEEC", lineColor: C.red });
  s.addShape("rect", { x: ML, y: 2.62, w: 0.06, h: 1.16, fill: { color: C.red }, line: { type: "none" } });
  s.addText("！", {
    x: ML + 0.3, y: 2.62, w: 0.5, h: 1.16, fontFace: FONT, fontSize: 22, bold: true, color: C.red, align: "center", valign: "middle",
  });
  s.addText("鈴木さんの週の労働時間が 41.5 時間になります（上限 40 時間）", {
    x: ML + 0.92, y: 2.78, w: CW - 3.6, h: 0.42, fontFace: FONT, fontSize: 15, bold: true, color: "8E322A", valign: "middle",
  });
  s.addText("8/22（金）の遅番を高橋さんへ振り替えると解消します。", {
    x: ML + 0.92, y: 3.2, w: CW - 3.6, h: 0.38, fontFace: FONT, fontSize: 12, color: "9C4A42", valign: "middle",
  });
  card(s, SW - MR - 2.3, 2.94, 1.9, 0.52, { fill: C.red, line: false, radius: 0.04 });
  s.addText("振り替える", {
    x: SW - MR - 2.3, y: 2.94, w: 1.9, h: 0.52, fontFace: FONT, fontSize: 12, bold: true, color: C.white, align: "center", valign: "middle",
  });

  // チェック項目
  s.addText("確定前に自動で検査する項目", {
    x: ML, y: 4.06, w: CW, h: 0.36, fontFace: FONT, fontSize: 12.5, bold: true, color: C.main, valign: "middle",
  });
  const checks = [
    ["週の労働時間の上限", "40時間／法定・変形労働にも対応"],
    ["連続勤務日数", "6日を超える連勤を検知"],
    ["勤務間インターバル", "終業から次の始業まで11時間"],
    ["休憩時間の付与", "6時間超で45分、8時間超で60分"],
    ["深夜帯の配置", "18歳未満のスタッフを自動で除外"],
    ["月間の所定超過", "契約時間を超えるスタッフを警告"],
  ];
  const cw2 = (CW - 0.4) / 3;
  checks.forEach(([t, d], i) => {
    const x = ML + (i % 3) * (cw2 + 0.2);
    const y = 4.54 + Math.floor(i / 3) * 1.02;
    card(s, x, y, cw2, 0.86, { fill: C.white });
    s.addText("✓", { x: x + 0.22, y, w: 0.34, h: 0.86, fontFace: FONT, fontSize: 12.5, bold: true, color: C.green, valign: "middle" });
    s.addText(t, { x: x + 0.58, y: y + 0.12, w: cw2 - 0.78, h: 0.34, fontFace: FONT, fontSize: 12.5, bold: true, color: C.ink, valign: "middle" });
    s.addText(d, { x: x + 0.58, y: y + 0.44, w: cw2 - 0.78, h: 0.32, fontFace: FONT, fontSize: 10.5, color: C.muted, valign: "middle" });
  });
  footer(s);
}

/* ============================================================
   11 導入効果
   ============================================================ */
{
  const s = prs.addSlide();
  bg(s);
  head(s, "03　導入した効果", "導入した店舗では、作業時間が平均 72% 減りました。",
       "導入から3か月経過した42店舗の平均値です。");

  const kpis = [
    ["－72", "%", "シフト作成にかかる時間", "月 9.5h → 2.6h"],
    ["－4.1", "%", "人件費（同時期比）", "予算超過の事前検知による"],
    ["＋18", "pt", "スタッフの継続率", "希望の反映率が上がったため"],
  ];
  const kw = (CW - 0.5) / 3;
  kpis.forEach(([num, unit, label, note], i) => {
    const x = ML + i * (kw + 0.25);
    card(s, x, 2.62, kw, 2.5, { fill: i === 0 ? C.main : C.white, lineColor: C.line, line: i !== 0 });
    const numCol = i === 0 ? C.white : C.main;
    const labCol = i === 0 ? C.ghost : C.muted;
    s.addText([
      { text: num, options: { fontSize: 52, bold: true, color: numCol } },
      { text: " " + unit, options: { fontSize: 20, bold: true, color: i === 0 ? C.accent : C.accent } },
    ], {
      x: x + 0.34, y: 2.92, w: kw - 0.68, h: 1.0, fontFace: FONT, valign: "middle",
    });
    s.addShape("rect", { x: x + 0.36, y: 4.06, w: 0.9, h: 0.035, fill: { color: C.accent }, line: { type: "none" } });
    s.addText(label, {
      x: x + 0.34, y: 4.24, w: kw - 0.68, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: i === 0 ? C.white : C.ink, valign: "middle",
    });
    s.addText(note, {
      x: x + 0.34, y: 4.62, w: kw - 0.68, h: 0.36, fontFace: FONT, fontSize: 11, color: labCol, valign: "middle",
    });
  });

  card(s, ML, 5.42, CW, 0.94, { fill: C.bgSub, line: false });
  s.addText("削減できた時間は、店長が売場とスタッフ育成に使う時間へ振り替えられています。", {
    x: ML + 0.42, y: 5.42, w: CW - 0.84, h: 0.94, fontFace: FONT, fontSize: 14, bold: true, color: C.main, valign: "middle",
  });
  s.addText("＊本資料は制作サンプルのため、記載の数値はすべて架空のものです。", {
    x: ML, y: 6.44, w: CW, h: 0.28, fontFace: FONT, fontSize: 9, color: C.muted,
  });
  footer(s);
}

/* ============================================================
   12 導入事例
   ============================================================ */
{
  const s = prs.addSlide();
  bg(s);
  head(s, "03　導入した効果", "「シフトの日に休日出勤しなくてよくなりました」",
       "飲食チェーン A社様（12店舗・スタッフ 180名）");

  // 引用
  card(s, ML, 2.66, 7.0, 3.5, { fill: C.bgSub, line: false });
  s.addText("“", {
    x: ML + 0.28, y: 2.72, w: 0.8, h: 0.8, fontFace: FONT, fontSize: 44, bold: true, color: C.accent, valign: "top",
  });
  s.addText(
    "これまでは月末の休日に店に出て、半日かけてシフトを組んでいました。\n\n" +
    "導入後は、集まった希望をもとに出てきた案を確認して、2〜3か所を直すだけです。" +
    "人件費が予算を超えそうな週もその場で分かるので、月末に慌てることがなくなりました。",
    { x: ML + 0.5, y: 3.44, w: 6.1, h: 2.1, fontFace: FONT, fontSize: 13.5, color: C.ink, lineSpacing: 24, valign: "top" }
  );
  s.addText("A社 エリアマネージャー 様", {
    x: ML + 0.5, y: 5.62, w: 6.1, h: 0.34, fontFace: FONT, fontSize: 11.5, color: C.muted,
  });

  // 右：数値
  const rx = ML + 7.4, rw = SW - MR - rx;
  const facts = [
    ["導入前", "月 11.0 時間／店", C.muted],
    ["導入後", "月 2.4 時間／店", C.accent],
  ];
  facts.forEach(([k, v, col], i) => {
    const y = 2.66 + i * 1.16;
    card(s, rx, y, rw, 1.0, { fill: C.white });
    s.addText(k, { x: rx + 0.3, y: y + 0.12, w: rw - 0.6, h: 0.3, fontFace: FONT, fontSize: 11, color: C.muted, valign: "middle" });
    s.addText(v, { x: rx + 0.3, y: y + 0.42, w: rw - 0.6, h: 0.44, fontFace: FONT, fontSize: 18, bold: true, color: col, valign: "middle" });
  });
  card(s, rx, 4.98, rw, 1.18, { fill: C.main, line: false });
  s.addText("12店舗あわせて\n月 103 時間の削減", {
    x: rx + 0.3, y: 4.98, w: rw - 0.6, h: 1.18, fontFace: FONT, fontSize: 15, bold: true, color: C.white, valign: "middle", lineSpacing: 24,
  });
  footer(s);
}

/* ============================================================
   13 料金
   ============================================================ */
{
  const s = prs.addSlide();
  bg(s);
  head(s, "04　料金と導入の流れ", "料金は、利用するスタッフ数で決まります。",
       "店舗数による追加費用はかかりません。初期費用・サポート費用も含まれています。");

  const plans = [
    ["スターター", "300", "円 / 人・月", ["〜30名", "希望収集・自動作成", "メールサポート"], false],
    ["スタンダード", "250", "円 / 人・月", ["31〜200名", "全機能を利用可能", "電話・チャットサポート", "他システムとのCSV連携"], true],
    ["エンタープライズ", "個別", "お見積り", ["201名〜", "全機能を利用可能", "専任担当者による支援", "API連携・SSO"], false],
  ];
  const pw = (CW - 0.5) / 3;
  plans.forEach(([name, price, unit, feats, rec], i) => {
    const x = ML + i * (pw + 0.25);
    const y = rec ? 2.5 : 2.66;
    const h = rec ? 3.72 : 3.4;
    card(s, x, y, pw, h, { fill: rec ? C.white : C.bgSub, lineColor: rec ? C.main : C.line });
    if (rec) {
      s.addShape("rect", { x, y, w: pw, h: 0.06, fill: { color: C.accent }, line: { type: "none" } });
      card(s, x + pw - 1.5, y - 0.0, 1.3, 0.38, { fill: C.accent, line: false, radius: 0.03 });
      s.addText("おすすめ", {
        x: x + pw - 1.5, y, w: 1.3, h: 0.38, fontFace: FONT, fontSize: 10, bold: true, color: C.white, align: "center", valign: "middle",
      });
    }
    s.addText(name, {
      x: x + 0.32, y: y + 0.44, w: pw - 0.64, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: C.main, valign: "middle",
    });
    s.addText([
      { text: price, options: { fontSize: price === "個別" ? 26 : 34, bold: true, color: C.ink } },
      { text: "　" + unit, options: { fontSize: 11.5, color: C.muted } },
    ], {
      x: x + 0.32, y: y + 0.92, w: pw - 0.64, h: 0.68, fontFace: FONT, valign: "middle",
    });
    s.addShape("rect", { x: x + 0.32, y: y + 1.68, w: pw - 0.64, h: 0.012, fill: { color: C.line }, line: { type: "none" } });
    feats.forEach((f, j) => {
      const fy = y + 1.84 + j * 0.44;
      s.addText("✓", { x: x + 0.32, y: fy, w: 0.3, h: 0.36, fontFace: FONT, fontSize: 11, bold: true, color: C.green, valign: "middle" });
      s.addText(f, { x: x + 0.62, y: fy, w: pw - 0.94, h: 0.36, fontFace: FONT, fontSize: 11.5, color: C.ink, valign: "middle" });
    });
  });
  s.addText("＊すべて税抜表示です。30日間の無料トライアルをご用意しています。", {
    x: ML, y: 6.42, w: CW, h: 0.3, fontFace: FONT, fontSize: 10, color: C.muted,
  });
  footer(s);
}

/* ============================================================
   14 導入の流れ
   ============================================================ */
{
  const s = prs.addSlide();
  bg(s);
  head(s, "04　料金と導入の流れ", "お申し込みから運用開始まで、最短 2 週間です。",
       "スタッフ情報の登録と就業ルールの設定は、当社が代行します。");

  const steps = [
    ["01", "お打ち合わせ", "現在の運用と\nご要望を確認", "1日"],
    ["02", "無料トライアル", "1店舗で\n実際に運用", "14日間"],
    ["03", "初期設定", "スタッフ登録・\n就業ルール設定", "3日"],
    ["04", "操作説明会", "店長向け\nオンライン60分", "1日"],
    ["05", "運用開始", "全店舗へ\n順次展開", "—"],
  ];
  const sw2 = (CW - 0.8) / 5;
  steps.forEach(([no, t, d, term], i) => {
    const x = ML + i * (sw2 + 0.2);
    const isLast = i === steps.length - 1;
    card(s, x, 2.7, sw2, 2.7, { fill: isLast ? C.main : C.white, lineColor: isLast ? C.main : C.line });
    s.addText(no, {
      x: x + 0.24, y: 2.9, w: sw2 - 0.48, h: 0.4, fontFace: FONT, fontSize: 13, bold: true,
      color: isLast ? C.accent : C.accent, valign: "middle",
    });
    s.addText(t, {
      x: x + 0.24, y: 3.34, w: sw2 - 0.48, h: 0.46, fontFace: FONT, fontSize: 14, bold: true,
      color: isLast ? C.white : C.ink, valign: "middle",
    });
    s.addText(d, {
      x: x + 0.24, y: 3.86, w: sw2 - 0.48, h: 0.86, fontFace: FONT, fontSize: 11, lineSpacing: 18,
      color: isLast ? C.ghost : C.muted, valign: "top",
    });
    s.addShape("rect", { x: x + 0.24, y: 4.82, w: sw2 - 0.48, h: 0.012, fill: { color: isLast ? "2C6675" : C.line }, line: { type: "none" } });
    s.addText(term, {
      x: x + 0.24, y: 4.94, w: sw2 - 0.48, h: 0.34, fontFace: FONT, fontSize: 11.5, bold: true,
      color: isLast ? C.white : C.main, valign: "middle",
    });
    if (i < 4) {
      s.addShape("rightArrow", { x: x + sw2 + 0.02, y: 3.94, w: 0.16, h: 0.22, fill: { color: C.ghost }, line: { type: "none" } });
    }
  });

  card(s, ML, 5.66, CW, 0.76, { fill: C.accentBg, line: false });
  s.addText("トライアル中の設定内容は、そのまま本番でご利用いただけます。作り直しは発生しません。", {
    x: ML + 0.42, y: 5.66, w: CW - 0.84, h: 0.76, fontFace: FONT, fontSize: 12.5, bold: true, color: "9A4A1E", valign: "middle",
  });
  footer(s);
}

/* ============================================================
   15 クロージング
   ============================================================ */
{
  const s = prs.addSlide();
  bg(s, C.main);
  s.addShape("rect", { x: 8.05, y: 0, w: SW - 8.05, h: SH, fill: { color: C.dark }, line: { type: "none" } });
  s.addShape("rect", { x: 0, y: 0, w: 0.14, h: SH, fill: { color: C.accent }, line: { type: "none" } });
  // 右面：チェックの入ったシフト表（＝完成した状態）
  {
    const gx = 8.75, gy = 2.28, cw = 0.62, ch = 0.62, cols = 6, rows = 4, gap = 0.09;
    const filled = { "0-1": 1, "1-3": 1, "2-0": 1, "3-4": 1, "0-4": 2, "2-2": 2, "3-1": 2, "1-5": 2 };
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const kind = filled[`${r}-${c}`];
        s.addShape("rect", {
          x: gx + c * (cw + gap), y: gy + r * (ch + gap), w: cw, h: ch,
          fill: kind === 1 ? { color: C.main } : kind === 2 ? { color: C.accent } : { type: "none" },
          line: kind ? { type: "none" } : { color: "17515F", width: 0.75 },
        });
      }
    }
  }

  s.addText("まずは1店舗、\n30日間の無料トライアルから。", {
    x: 1.05, y: 2.06, w: 6.6, h: 1.5, fontFace: FONT, fontSize: 27, bold: true, color: C.white, lineSpacing: 44,
  });
  s.addText("現在お使いのシフト表をお送りいただければ、\nそのまま移行した状態でご覧いただけます。", {
    x: 1.05, y: 3.66, w: 6.6, h: 0.7, fontFace: FONT, fontSize: 13.5, color: C.ghost, lineSpacing: 22,
  });
  s.addShape("rect", { x: 1.08, y: 4.56, w: 1.5, h: 0.045, fill: { color: C.accent }, line: { type: "none" } });

  const infos = [["株式会社ロタリア", ""], ["お問い合わせ", "sample@example.com"], ["Web", "example.com"]];
  infos.forEach(([k, v], i) => {
    const y = 4.9 + i * 0.5;
    s.addText(k, { x: 1.05, y, w: 2.4, h: 0.4, fontFace: FONT, fontSize: 12, color: C.ghost, valign: "middle" });
    s.addText(v, { x: 3.45, y, w: 4.2, h: 0.4, fontFace: FONT, fontSize: 12.5, bold: true, color: C.white, valign: "middle" });
  });
  s.addText("※本資料は制作サンプルです。Rotaria は実在しないサービスであり、\n　掲載の数値・事例はすべて架空のものです。", {
    x: 1.05, y: 6.52, w: 6.6, h: 0.56, fontFace: FONT, fontSize: 9, color: "6E8B93", lineSpacing: 14,
  });
  PAGE++;
}

/* ============================================================ */
prs.writeFile({ fileName: "out/Rotaria_サービス紹介資料_サンプル.pptx" })
  .then(f => console.log(`生成完了: ${f}　（全${PAGE}スライド）`))
  .catch(e => { console.error(e); process.exit(1); });
