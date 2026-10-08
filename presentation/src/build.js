// RIVET-WM seminar deck — rebuilt in the GAT seminar style
// (white slides, black 24pt top-left titles, diagrams carry the color, equations as images)
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");
const { applyTheme } = require("/mnt/skills/public/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || path.join(__dirname, "RIVET-WM_seminar.pptx");
const MEDIA = path.join(__dirname, "media");
const NOTES = JSON.parse(fs.readFileSync(path.join(__dirname, "notes.json"), "utf8"));

const THEME = {
  name: "RIVET-WM Seminar",
  headFontFace: "맑은 고딕",
  bodyFontFace: "맑은 고딕",
  colors: {
    dk1: "000000", lt1: "FFFFFF", dk2: "808080", lt2: "D9D9D9",
    accent1: "156082", // teal: edges, outlines
    accent2: "00B0F0", // highlight blue
    accent3: "FF7B7A", // coral: key nodes
    accent4: "EBDB88", // gold: secondary nodes
    accent5: "92D050", // green
    accent6: "FF0000", // alert red
    hlink: "156082", folHlink: "808080",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.author = "탁혜원 · 최재원 · 허동진";
pres.title = "RIVET-WM";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const S = pres.shapes;

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "TITLE",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0, y: 2.2, w: 13.33, h: 1.0, fontSize: 54, fontFace: "Georgia", align: "center", valign: "middle", color: C.text1, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "subtitle", type: "body", x: 0, y: 3.3, w: 13.33, h: 0.6, fontSize: 28, align: "center", valign: "middle", color: C.text1, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "author", type: "body", x: 0, y: 4.3, w: 13.33, h: 0.48, fontSize: 20, bold: true, align: "center", valign: "middle", color: C.text1, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "date", type: "body", x: 0, y: 4.86, w: 13.33, h: 0.48, fontSize: 18, align: "center", valign: "middle", color: C.text1, margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "TOC",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.22, y: 0.23, w: 4.0, h: 0.64, fontSize: 32, bold: true, align: "left", valign: "middle", color: C.text1, margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.38, y: 0.48, w: 12.5, h: 0.5, fontSize: 24, bold: true, align: "left", valign: "middle", color: C.text1, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "desc", type: "body", x: 0.38, y: 1.11, w: 12.5, h: 0.4, fontSize: 14, bold: true, align: "left", valign: "middle", color: C.text1, margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CLOSING",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0, y: 2.7, w: 13.33, h: 1.0, fontSize: 54, align: "center", valign: "middle", color: C.text1, margin: 0 }, text: "" } },
  ],
});

// ---------- helpers ----------
const IMG = {
  image7: [2600, 860], image10: [2390, 217], image11: [2600, 364], image12: [2600, 315],
  image13: [2600, 453], image14: [2600, 719], image16: [2600, 397], image17: [2600, 378],
  image18: [2329, 1383], image19: [2600, 199], image20: [2600, 1054], image21: [2600, 1527],
  image22: [2600, 1054], image25: [2000, 1241], image27: [2600, 280], image28: [2600, 1134],
  image34: [500, 500],
};
function img(slide, name, x, y, o) {
  const [iw, ih] = IMG[name];
  let w = o.w, h = o.h;
  if (w && !h) h = (w * ih) / iw;
  if (h && !w) w = (h * iw) / ih;
  slide.addImage({ path: path.join(MEDIA, name + ".png"), x, y, w, h, altText: o.alt || name, objectName: o.name || name });
  return { w, h };
}

let section = "";
function content(title, desc, noteKey) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  if (desc) s.addText(desc, { placeholder: "desc", bullet: false });
  if (noteKey && NOTES[noteKey]) s.addNotes(NOTES[noteKey]);
  return s;
}

function T(slide, text, o) {
  slide.addText(text, Object.assign({ isTextBox: true, margin: 0, color: C.text1, fontSize: 14, valign: "top", align: "left" }, o));
}
// "Bold term : explanation"
function D(term, body, size = 14, last = true) {
  return [
    { text: term + " : ", options: { bold: true, fontSize: size } },
    { text: body, options: { fontSize: size, breakLine: !last } },
  ];
}

function arrow(slide, x1, y1, x2, y2, o = {}) {
  const line = { color: o.color || C.text1, width: o.width || 1.25 };
  if (o.dash) line.dashType = o.dash;
  if (!o.noHead) line.endArrowType = "triangle";
  slide.addShape(S.LINE, {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
    flipH: x2 < x1, flipV: y2 < y1, line, objectName: o.name,
  });
}
// polyline: arrowhead on last segment only
function path2(slide, pts, o = {}) {
  for (let i = 0; i < pts.length - 1; i++) {
    const last = i === pts.length - 2;
    arrow(slide, pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1], Object.assign({}, o, { noHead: o.noHead || !last }));
  }
}

function box(slide, x, y, w, h, title, sub, o = {}) {
  const runs = [{ text: title, options: { bold: true, fontSize: o.size || 14, breakLine: !!sub } }];
  if (sub) runs.push({ text: sub, options: { fontSize: o.subSize || 12 } });
  const line = { color: o.line || C.text1, width: o.lw || 1 };
  if (o.dash) line.dashType = o.dash;
  slide.addText(runs, {
    shape: S.RECTANGLE, x, y, w, h,
    fill: { color: o.fill || C.background1, transparency: o.transp || 0 },
    line, align: "center", valign: "middle", margin: 3, color: C.text1, objectName: o.name,
  });
}

function node(slide, cx, cy, d, label, o = {}) {
  const runs = typeof label === "string" ? [{ text: label, options: {} }] : label;
  slide.addText(runs, {
    shape: S.OVAL, x: cx - d / 2, y: cy - d / 2, w: d, h: d,
    fill: { color: o.fill || C.background1, transparency: o.transp || 0 },
    line: { color: o.line || C.text1, width: o.lw || 1.5 },
    align: "center", valign: "middle", margin: 0, fontSize: o.size || 16, bold: !!o.bold,
    italic: !!o.italic, fontFace: o.font, color: C.text1, objectName: o.name,
  });
}
// edge between two circle centers, trimmed to their rims
function edge(slide, a, b, ra, rb, o = {}) {
  const dx = b[0] - a[0], dy = b[1] - a[1], L = Math.hypot(dx, dy), ux = dx / L, uy = dy / L;
  arrow(slide, a[0] + ux * ra, a[1] + uy * ra, b[0] - ux * rb, b[1] - uy * rb, o);
}
// math label like g_t
function sym(base, sub) {
  const r = [{ text: base, options: { italic: true, fontFace: "Times New Roman", fontSize: 20 } }];
  if (sub) r.push({ text: sub, options: { italic: true, fontFace: "Times New Roman", fontSize: 20, subscript: true } });
  return r;
}

// booktabs-style table: top / mid / bottom rules only
const RULE = (pt) => ({ type: "solid", pt, color: HEX.dk1 });
const NONE = { type: "none" };
function table(slide, header, rows, o) {
  const fs_ = o.fontSize || 14;
  const hdr = header.map((h, j) => ({
    text: h,
    options: { bold: true, fontSize: fs_, color: C.text1, align: (o.align && o.align[j]) || "left", border: [RULE(1.5), NONE, RULE(0.75), NONE] },
  }));
  const body = rows.map((r, i) => r.map((c, j) => {
    const last = i === rows.length - 1;
    const bold = o.boldRow === i;
    return {
      text: String(c),
      options: {
        bold, fontSize: fs_, color: C.text1, align: (o.align && o.align[j]) || "left",
        fill: bold ? { color: C.background2, transparency: 50 } : undefined,
        border: [NONE, NONE, last ? RULE(1.5) : { type: "solid", pt: 0.5, color: HEX.lt2 }, NONE],
      },
    };
  }));
  slide.addTable([hdr, ...body], {
    x: o.x, y: o.y, w: o.w, colW: o.colW, rowH: o.rowH || 0.38, margin: [0.04, 0.08, 0.04, 0.08],
    valign: "middle", objectName: o.name,
  });
}

function bar(slide, labels, series, o) {
  slide.addChart(pres.charts.BAR, series.map((s) => ({ name: s.name, labels, values: s.values })), {
    x: o.x, y: o.y, w: o.w, h: o.h, barDir: "col", barGapWidthPct: o.gap || 70,
    chartColors: o.colors,
    showTitle: true, title: o.title, titleFontFace: "+mn-lt", titleFontSize: 14, titleColor: HEX.dk1,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: o.fmt || "0.0",
    dataLabelFontFace: "+mn-lt", dataLabelFontSize: 12, dataLabelColor: HEX.dk1,
    catAxisLabelFontFace: "+mn-lt", catAxisLabelFontSize: 12, catAxisLabelColor: HEX.dk1,
    catAxisLineColor: HEX.dk1, catGridLine: { style: "none" },
    valAxisHidden: true, valGridLine: { style: "none" },
    valAxisMinVal: 0, valAxisMaxVal: o.max,
    showLegend: series.length > 1, legendPos: "b", legendFontFace: "+mn-lt", legendFontSize: 12, legendColor: HEX.dk1,
    objectName: o.name,
  });
}


// ======================================================================
// 1. Title
// ======================================================================
section = "Introduction";
pres.addSection({ title: section });
{
  const s = pres.addSlide({ masterName: "TITLE", sectionTitle: section });
  s.addText("RIVET-WM", { placeholder: "title" });
  s.addText("전류로 운전상태를 읽어 진동 이상을 판정하는 유압펌프 감시 모델", { placeholder: "subtitle", bullet: false, fontSize: 26 });
  s.addText("탁혜원 · 최재원 · 허동진", { placeholder: "author", bullet: false });
  s.addText("KAMP 2026 · 과제 3", { placeholder: "date", bullet: false });
  T(s, "Team 진짜멋진제조에이아이를만들어서세상을놀래킬거야", { x: 0, y: 5.5, w: 13.33, h: 0.4, fontSize: 14, align: "center", valign: "middle" });
  s.addNotes(NOTES[1]);
}

// ======================================================================
// 2. TOC
// ======================================================================
{
  const s = pres.addSlide({ masterName: "TOC", sectionTitle: section });
  s.addText("목차", { placeholder: "title" });
  const items = [
    "1. 문제 : 부하가 바뀌면 정상도 경보가 된다",
    "2. 데이터 진단",
    "3. 유형 1 : 두 전문가가 이상을 판정",
    "4. 실험 결과",
    "5. 유형 2 : 판정을 설명하고 사람을 부르기",
    "6. 현장 적용 및 결론",
  ];
  items.forEach((it, i) => T(s, it, { x: 3.3, y: 1.72 + i * 0.68, w: 8.5, h: 0.5, fontSize: 24, valign: "middle" }));
  s.addNotes("발표 순서입니다. 문제 정의와 데이터 진단에서 출발해, 유형 1의 두 전문가 판정 모델과 실험 결과를 보고, 유형 2의 교정 루프와 근거 기반 설명, 마지막으로 현장 적용과 결론 순으로 말씀드리겠습니다.");
}

// ======================================================================
// 3. Core contributions
// ======================================================================
{
  const s = content("핵심 기여", "부하를 읽고 → 같은 부하의 정상과 비교해 판정하고 → 근거로 설명하고 → 작업자 확인으로 교정한다", 2);
  const cols = [0.5, 3.6, 6.7, 9.8];
  const fills = [C.accent3, C.accent4, C.accent2, C.accent5];
  const data = [
    ["부하를 알고 비교한다", "55.5% → 8.6%", "처음 보는 정상 운전을 이상으로 오인\n외부 시험대 오경보 19.0→10.4%"],
    ["경보는 적게, 상태는 자세히", "11.7건 / 시간", "10개 방법 중 가장 적은 경보\n5가지 상태로 나눠 알림"],
    ["고장은 지우지 않는 교정", "92.3% → 64.6%", "작업자 확인 후 운전변화 알림\n고장 섞인 교정 11/16건 거절"],
    ["AI는 설명만, 호출은 규칙표", "11% → 100%", "답변의 근거 인용률\n규칙표 10,206개 조합 위반 0"],
  ];
  const d = 0.59, cy = 2.05 + d / 2;
  for (let i = 0; i < 3; i++) arrow(s, cols[i] + d, cy, cols[i + 1], cy, { color: C.accent1, width: 1.5 });
  data.forEach(([h, m, t], i) => {
    node(s, cols[i] + d / 2, cy, d, String(i + 1), { fill: fills[i], bold: true });
    T(s, h, { x: cols[i], y: 2.9, w: 2.9, h: 0.75, fontSize: 16, bold: true });
    T(s, m, { x: cols[i], y: 3.7, w: 2.9, h: 0.55, fontSize: 24, bold: true, valign: "middle" });
    T(s, t, { x: cols[i], y: 4.35, w: 3.0, h: 0.9, fontSize: 13, paraSpaceAfter: 2 });
  });
  T(s, [
    { text: "최종 결합(A+B), 공식 라벨 시간순 평가", options: { bold: true, fontSize: 16, breakLine: true } },
    { text: "이상 구간 5개 모두 탐지 · 재현율 0.993 · F1 0.915 · 정상 행 오경보 0.6% (24 / 4,090행, 한 구간)", options: { fontSize: 16 } },
  ], { x: 0.5, y: 5.75, w: 12.3, h: 0.85, paraSpaceAfter: 4 });
}

// ======================================================================
// 4. Problem definition
// ======================================================================
{
  const s = content("문제 정의", "프레스가 무거운 일을 하면 진동도 커진다 — 고정 기준은 이 정상 변화를 이상으로 착각한다", 3);
  const bx = [0.5, 3.0, 5.5], bw = 2.0, by = 1.95, bh = 0.7;
  box(s, bx[0], by, bw, bh, "유압 펌프 이상", null, { fill: C.accent3, size: 15 });
  box(s, bx[1], by, bw, bh, "압력 불안정", null, { size: 15 });
  box(s, bx[2], by, bw, bh, "불량 · 가동 중단", null, { size: 15 });
  arrow(s, bx[0] + bw, by + bh / 2, bx[1], by + bh / 2);
  arrow(s, bx[1] + bw, by + bh / 2, bx[2], by + bh / 2);
  T(s, "이상 유형 : 마모 · 정렬 · 캐비테이션 · 밸브 충격 · 계측 이상", { x: 0.5, y: 2.8, w: 7.0, h: 0.35, fontSize: 12, color: C.text2 });
  T(s, [
    ...D("지금의 감시", "주기 점검과 작업자 경험(노하우)에 의존", 16, false),
    ...D("어려운 점", "약한 이상 진동이 정상적인 부하 변화에 묻힘", 16, false),
    ...D("결과", "오경보가 반복되면 작업자가 경보를 믿지 않게 됨", 16),
  ], { x: 0.5, y: 3.4, w: 7.4, h: 1.5, paraSpaceAfter: 10 });
  T(s, "22.3%", { x: 8.6, y: 1.85, w: 4.3, h: 1.0, fontSize: 60, bold: true, valign: "middle" });
  T(s, "고정 기준 : 가장 무거운 부하에서 정상인데 경보가 난 비율", { x: 8.6, y: 2.9, w: 4.3, h: 0.6, fontSize: 14 });
  T(s, "3.7–5.7%", { x: 8.6, y: 3.65, w: 4.3, h: 0.6, fontSize: 32, bold: true, valign: "middle" });
  T(s, "부하를 반영한 기준이면", { x: 8.6, y: 4.3, w: 4.3, h: 0.35, fontSize: 14 });
  T(s, "The Problem", { x: 0.5, y: 5.2, w: 6, h: 0.45, fontSize: 20, bold: true, valign: "middle" });
  T(s, [
    { text: "P1. ", options: { bold: true, fontSize: 18 } },
    { text: "진짜 이상은 놓치지 않는다", options: { fontSize: 18, breakLine: true } },
    { text: "P2. ", options: { bold: true, fontSize: 18 } },
    { text: "부하 변화는 이상으로 착각하지 않는다", options: { fontSize: 18 } },
  ], { x: 0.5, y: 5.75, w: 9.0, h: 0.95, paraSpaceAfter: 6 });
}

// ======================================================================
// 5-6. Data diagnosis
// ======================================================================
section = "데이터 진단";
pres.addSection({ title: section });
const n4 = NOTES[4];
const n4split = n4.indexOf("정상·이상이 다른 날짜");
{
  const s = content("데이터 진단 ① 신호 들여다보기", "정상 전류는 매끈한 사인파 — 이상은 파형 모양이 깨진다");
  if (n4split > 0) s.addNotes(n4.slice(0, n4split).trim()); else s.addNotes(n4);
  img(s, "image7", 0.65, 1.7, { w: 12.0, alt: "데이터 진단: (a) 전류 파형 (b) 파고율 분포 (c) 평면 이탈도 분포" });
  T(s, "(a) 전류 파형   (b) crest factor 분포   (c) plane deviation 분포 — 개발 분할 20행 이상 정상 287 · 이상 8구간", { x: 0.65, y: 5.72, w: 12.0, h: 0.3, fontSize: 12, color: C.text2 });
  T(s, [
    ...D("데이터", "정상 : 이상 = 20,000 : 600행 (불균형) · 이상 구간의 진동 표준편차 6.1배, 그러나 일부 이상은 정상 진폭 범위", 14, false),
    ...D("측정 한계", "0.1초 간격이라 5 Hz까지만 보임 → 주파수 대신 파형 모양(crest factor 등)을 본다", 14),
  ], { x: 0.65, y: 6.12, w: 12.2, h: 0.85, paraSpaceAfter: 4 });
}
{
  const s = content("데이터 진단 ② 발견 → 설계", "데이터에서 본 6가지 특징이 각각 하나의 설계 결정이 됨");
  if (n4split > 0) s.addNotes(n4.slice(n4split).trim());
  const rows = [
    ["구간이 짧다 (최대 5초)", "구간 단위로, 과거 정보만 써서 판정"],
    ["정상 전류는 사인파", "A : 전류 진폭으로 부하 추정,  B : 파형 모양 특징"],
    ["부하가 진동 변동의 92.6%를 설명", "A : 같은 부하의 정상 진동과 비교"],
    ["진폭은 정상인데 이상인 경우", "B : 파형 모양 + A : 같은 부하 대비 잔차"],
    ["관측 재개 직후 오경보가 몰림", "3번 중 2번 이어질 때만 경보 + 판단 보류"],
    ["라벨 = 날짜 = 저장 형식", "시간순 3분할 + 형식 무관 점검 + 가상 고장 주입"],
  ];
  T(s, "발견", { x: 0.9, y: 1.75, w: 4.6, h: 0.35, fontSize: 14, bold: true, color: C.text2 });
  T(s, "설계 대응", { x: 6.9, y: 1.75, w: 5.8, h: 0.35, fontSize: 14, bold: true, color: C.text2 });
  rows.forEach(([a, b], i) => {
    const y = 2.2 + i * 0.5;
    node(s, 0.6, y + 0.18, 0.32, String(i + 1), { size: 11, lw: 1 });
    T(s, a, { x: 0.9, y, w: 4.9, h: 0.36, fontSize: 16, bold: true, valign: "middle" });
    arrow(s, 5.9, y + 0.18, 6.7, y + 0.18, { color: C.accent1, width: 1.5 });
    T(s, b, { x: 6.9, y, w: 6.0, h: 0.36, fontSize: 16, valign: "middle" });
  });
  T(s, [
    ...D("주의", "정상과 이상이 서로 다른 날짜 · 저장 형식으로 기록 → 형식만 봐도 라벨이 갈림 (ROC-AUC 1.000)", 14, false),
    { text: "형식 차이를 없애도 분리가 유지됨을 확인 (AP 0.968–0.993) · 개발 · 후보 · 최종을 시간순으로 나누고 최종은 1회만 평가", options: { fontSize: 14, breakLine: true } },
    { text: "→ 이 데이터에 잘 맞는 것만으로는 부족하다 — 처음 보는 운전 상황에서도 버티는 구조가 필요", options: { fontSize: 16, bold: true } },
  ], { x: 0.5, y: 5.4, w: 12.4, h: 1.4, paraSpaceAfter: 6 });
}

// ======================================================================
// 7-14. Type 1 — two experts
// ======================================================================
section = "유형 1 판정";
pres.addSection({ title: section });
{
  const s = content("전체 파이프라인", "설비 신호를 읽기만 하고, 판정 → 설명 → 사람 확인 → 교정으로 돈다 — 설명이 느려도 감시는 멈추지 않음", 6);
  const xs = [0.55, 3.1, 5.65, 8.2, 10.75], bw = 2.05, by = 1.85, bh = 0.95;
  const st = [
    ["읽기 전용 수집", "0.1초마다 1행", {}],
    ["유형 1 판정", "전문가 A · B", { fill: C.accent3, transp: 20 }],
    ["판정 기록 보관", "최근 4,096개", {}],
    ["유형 2 설명", "질문 · 상태 변화 · 요청", { fill: C.accent4 }],
    ["출력", "화면 · 정비 요청 · 일지", {}],
  ];
  st.forEach(([a, b, o], i) => {
    box(s, xs[i], by, bw, bh, a, b, Object.assign({ size: 15 }, o));
    if (i < 4) arrow(s, xs[i] + bw, by + bh / 2, xs[i + 1], by + bh / 2);
  });
  const ly = 3.25, lh = 0.6;
  box(s, xs[1], ly, xs[3] + bw - xs[1], lh, "교정 루프", "작업자 확인 → canary test → 반영", { dash: "dash", size: 14 });
  path2(s, [[xs[4] + bw / 2, by + bh], [xs[4] + bw / 2, ly + lh / 2], [xs[3] + bw, ly + lh / 2]], { dash: "dash" });
  arrow(s, xs[1] + bw / 2, ly, xs[1] + bw / 2, by + bh, { dash: "dash" });
  T(s, "얼마나 빠른가", { x: 0.55, y: 4.2, w: 5, h: 0.35, fontSize: 16, bold: true });
  table(s, ["구성", "주기 · 지연"], [
    ["전문가 A", "행마다 약 0.2 ms"],
    ["전문가 B", "행마다 약 18 ms"],
    ["지식그래프 검색 · 검사", "수 ms"],
    ["언어모델 답변", "질문당 평균 약 31초"],
  ], { x: 0.55, y: 4.65, w: 5.8, colW: [2.9, 2.9], rowH: 0.36 });
  T(s, "안전하게 붙이기", { x: 7.0, y: 4.2, w: 5, h: 0.35, fontSize: 16, bold: true });
  T(s, [
    ...D("읽기 전용", "설비로 쓰는 경로가 없음 — 값을 읽기만 함", 14, false),
    ...D("판정과 설명 분리", "유형 1은 유형 2를 기다리지 않음", 14, false),
    ...D("구현 검증", "기준 구현과 20,802 프레임에서 판정 100% 일치", 14, false),
    ...D("검사 통과", "유형 2 28,901건 · 보안 103건 · 종단 32건", 14),
  ], { x: 7.0, y: 4.7, w: 5.9, h: 1.9, paraSpaceAfter: 8 });
}
{
  const s = content("유형 1 판정 : 두 전문가 모델", "A는 '같은 부하의 정상과 다른가', B는 '전류 파형 모양이 깨졌나'를 본다 — 둘 중 하나라도 경보면 경보", 5);
  const inX = 0.5, inW = 1.6, rowA = 2.15, rowB = 3.95, bh = 0.8;
  const midY = (rowA + rowB + bh) / 2;
  box(s, inX, midY - 0.55, inW, 1.1, "전류 · 진동", "0.1초 간격", { size: 15 });
  T(s, "전문가 A : 같은 부하의 정상과 비교 (정상 기록만 학습)", { x: 2.75, y: rowA - 0.4, w: 8, h: 0.35, fontSize: 14, bold: true });
  T(s, "전문가 B : 전류 파형 모양 판정 (라벨로 학습)", { x: 2.75, y: rowB - 0.4, w: 8, h: 0.35, fontSize: 14, bold: true });
  const aX = [2.75, 4.75, 6.75, 8.75], aW = 1.65;
  [["부하 추정", "전류 진폭"], ["정상과 비교", "같은 부하"], ["건강 점수", "편차 누적"], ["보정 · 판정", "경보 예산 1%"]].forEach(([a, b], i) => {
    box(s, aX[i], rowA, aW, bh, a, b, { fill: C.accent2, transp: 75 });
    if (i < 3) arrow(s, aX[i] + aW, rowA + bh / 2, aX[i + 1], rowA + bh / 2);
  });
  const bX = [2.75, 5.45, 8.15], bW = 2.25;
  [["모양 특징", "crest · shape"], ["파형 분류", "MiniRocket / GBDT"], ["확률 보정", "Platt scaling"]].forEach(([a, b], i) => {
    box(s, bX[i], rowB, bW, bh, a, b, { fill: C.accent4 });
    if (i < 2) arrow(s, bX[i] + bW, rowB + bh / 2, bX[i + 1], rowB + bh / 2);
  });
  const cX = 10.95, cW = 1.95;
  box(s, cX, rowA, cW, rowB + bh - rowA, "결합 판정", "A 또는 B\n5가지 상태", { fill: C.accent3, size: 16, subSize: 13 });
  const sx = 2.4;
  arrow(s, inX + inW, midY, sx, midY, { noHead: true });
  arrow(s, sx, rowA + bh / 2, sx, rowB + bh / 2, { noHead: true });
  arrow(s, sx, rowA + bh / 2, aX[0], rowA + bh / 2);
  arrow(s, sx, rowB + bh / 2, bX[0], rowB + bh / 2);
  arrow(s, aX[3] + aW, rowA + bh / 2, cX, rowA + bh / 2);
  arrow(s, bX[2] + bW, rowB + bh / 2, cX, rowB + bh / 2);
  const fy = 5.15;
  path2(s, [[cX + cW / 2, rowB + bh], [cX + cW / 2, fy], [inX + inW / 2, fy], [inX + inW / 2, midY + 0.55]], { dash: "dash", color: C.accent1 });
  T(s, "작업자 확인 → 교정 루프", { x: 4.5, y: fy + 0.07, w: 4.3, h: 0.3, fontSize: 12, align: "center", color: C.text2 });
  T(s, [
    ...D("전문가 A", "부하가 바뀐 정상과 진짜 이상을 가른다 → 5가지 상태로 판정", 14, false),
    ...D("전문가 B", "진폭이 정상이어도 파형 모양이 깨진 이상을 잡는다 → 확률 0.40 이상이면 경보", 14, false),
    ...D("결합", "둘 중 하나라도 경보면 경보, 0.1초(1행)마다 판정", 14),
  ], { x: 0.5, y: 5.7, w: 12.4, h: 1.2, paraSpaceAfter: 6 });
}
{
  const s = content("전문가 A ① 진동을 '부하 몫'과 '건강 몫'으로 나누기", "부하로 설명되지 않는 진동 상승만 이상 신호로 남긴다", 7);
  const d = 0.72, r = d / 2;
  const P = {
    g0: [1.1, 2.3], g1: [3.5, 2.3], q: [5.6, 2.3],
    x: [2.3, 3.75], y: [4.55, 3.75],
    h0: [1.1, 5.2], h1: [3.5, 5.2],
  };
  const E = [["g0", "g1"], ["h0", "h1"], ["g1", "x"], ["g1", "y"], ["q", "y"], ["h1", "y"]];
  E.forEach(([a, b]) => edge(s, P[a], P[b], r, r, { color: C.accent1, width: 1.5 }));
  node(s, ...P.g0, d, sym("g", "t−1"));
  node(s, ...P.g1, d, sym("g", "t"));
  node(s, ...P.q, d, sym("q", "t"));
  node(s, ...P.h0, d, sym("h", "t−1"));
  node(s, ...P.h1, d, sym("h", "t"));
  node(s, ...P.x, d, sym("x", "t"), { fill: C.background2 });
  node(s, ...P.y, d, sym("y", "t"), { fill: C.background2 });
  T(s, "부하는 이어짐", { x: 1.45, y: 1.75, w: 1.7, h: 0.3, fontSize: 12, align: "center", color: C.text2 });
  T(s, "천천히 변함", { x: 1.6, y: 4.86, w: 1.4, h: 0.3, fontSize: 12, align: "center", color: C.text2 });
  node(s, 0.62, 6.2, 0.26, "", { lw: 1 });
  T(s, "보이지 않는 상태 — g : 부하 구간(4단계), q : 운전 맥락, h : 건강 편차", { x: 0.9, y: 6.05, w: 5.6, h: 0.3, fontSize: 12, valign: "middle" });
  node(s, 0.62, 6.6, 0.26, "", { lw: 1, fill: C.background2 });
  T(s, "측정값 — x : 전류 파형, y : 진동 크기(로그 RMS)", { x: 0.9, y: 6.45, w: 5.6, h: 0.3, fontSize: 12, valign: "middle" });
  const ex = 6.9, ew = 5.9;
  img(s, "image10", ex, 1.95, { w: 5.2, alt: "관측 분해 식" });
  T(s, "진동 = 같은 부하에서 원래 나는 진동 μ + 건강 이상으로 늘어난 진동 h + 구간 간 변동 u + 잡음 v", { x: ex, y: 2.55, w: ew, h: 0.6, fontSize: 14 });
  img(s, "image11", ex, 3.3, { w: 5.6, alt: "건강 상태의 OU process 식" });
  T(s, "h는 칼만 필터로 추적 · 측정이 끊기면 300초 시간 상수로 0으로 돌아감 (오래된 증거는 잊음)", { x: ex, y: 4.18, w: ew, h: 0.6, fontSize: 14 });
  T(s, "93% · 58%", { x: ex, y: 4.95, w: ew, h: 0.6, fontSize: 32, bold: true, valign: "middle" });
  T(s, "하부 · 상부 진동 변동 중 부하 구간이 설명하는 몫", { x: ex, y: 5.6, w: ew, h: 0.35, fontSize: 14 });
  T(s, D("가벼운 모델", "신경망 없이 파라미터 약 320개, 정상 기록만 학습 · 한 행 약 0.2 ms", 14), { x: ex, y: 6.1, w: ew, h: 0.65 });
}
{
  const s = content("전문가 A ② 전류로 부하 읽기", "전류는 한 구간 안에서 하나의 사인파 — 진폭이 클수록 무거운 부하", 8);
  T(s, D("① 전류 = 사인파", "진폭이 부하 구간을 결정", 16), { x: 0.5, y: 1.8, w: 7.6, h: 0.4 });
  img(s, "image12", 0.7, 2.25, { w: 6.9, alt: "전류 방출 모델 식" });
  T(s, D("② 시작 위상은 몰라도 됨", "72개 위상 후보를 평균 (phase-marginalized)", 16), { x: 0.5, y: 3.25, w: 7.8, h: 0.4 });
  img(s, "image13", 0.7, 3.72, { w: 6.4, alt: "phase-marginalized likelihood 식" });
  T(s, [
    ...D("③ 부하 구간 4개", "진폭 약 119 / 146 / 202 / 245 (BIC로 개수 선택)", 16, false),
    ...D("④ 앞 구간에서 이어받기", "직전 구간의 부하로 다음 부하를 예측, 18초 넘게 끊기면 새로 시작", 16),
  ], { x: 0.5, y: 5.05, w: 7.8, h: 1.0, paraSpaceAfter: 8 });
  bar(s, ["3행", "8행", "10행"], [{ name: "오분류율", values: [42.0, 8.4, 3.0] }], {
    x: 8.6, y: 1.75, w: 4.3, h: 3.9, colors: [HEX.accent1], title: "관측 행 수별 오분류율 (%)", max: 50, name: "오분류율 차트",
  });
  T(s, [
    { text: "1초(10행)면 97% 정확", options: { bold: true, fontSize: 16, breakLine: true } },
    { text: "확신 0.9 이상 또는 10행이면 부하 확정, 그 전에는 '판단 보류'", options: { fontSize: 14 } },
  ], { x: 8.75, y: 5.8, w: 4.15, h: 0.95, paraSpaceAfter: 4 });
}
{
  const s = content("전문가 A ③ 경보 예산과 판정 규칙", "'정상에서 100번 중 1번 나올 만큼 이상한가'로 판정 — 이 1%를 세 경로가 나눠 쓴다", 9);
  const bx = 0.5, bw = 3.9, by = 2.3, bh = 0.4;
  T(s, "경보 예산 α = 0.01 (1%)", { x: bx, y: 1.8, w: bw, h: 0.4, fontSize: 16, bold: true });
  const parts = [
    ["건강 모델", "0.8 α", 0.8, C.accent1, "같은 부하 대비 진동 상승"],
    ["sentinel", "0.1 α", 0.1, C.accent3, "전류 센서가 죽어도 잡는 안전망"],
    ["데이터 품질", "0.1 α", 0.1, C.accent4, "전류 파형이 깨졌는지"],
  ];
  let cx = bx;
  parts.forEach(([n, v, f, c]) => {
    s.addShape(S.RECTANGLE, { x: cx, y: by, w: bw * f, h: bh, fill: { color: c }, line: { color: C.background1, width: 1.5 }, objectName: "budget " + n });
    cx += bw * f;
  });
  parts.forEach(([n, v, f, c, why], i) => {
    const y = 2.92 + i * 0.52;
    s.addShape(S.RECTANGLE, { x: bx, y: y + 0.07, w: 0.2, h: 0.2, fill: { color: c }, line: { color: c, width: 0.5 } });
    T(s, [
      { text: n + "  ", options: { bold: true, fontSize: 14 } },
      { text: v, options: { fontSize: 14, breakLine: true } },
      { text: why, options: { fontSize: 12, color: C.text2 } },
    ], { x: bx + 0.32, y, w: 3.6, h: 0.5 });
  });
  img(s, "image14", bx, 4.55, { w: 3.9, alt: "conformal p-value와 경보 예산 분배 식" });
  T(s, "p값 = 정상 기록 중 이보다 더 이상한 비율", { x: bx, y: 5.7, w: 3.9, h: 0.3, fontSize: 12, color: C.text2 });
  const rx = 4.95, rw = 3.75;
  T(s, "판정 순서 (위에서부터)", { x: rx, y: 1.8, w: rw, h: 0.4, fontSize: 16, bold: true });
  const rules = [
    ["1  전류 파형이 깨짐", "→ 센서 점검"],
    ["2  처음 보는 운전 상태", "→ 물리 예측과 맞으면 운전변화 알림"],
    ["3  부하 확정 전", "→ 판단 보류 가능"],
    ["4  부하 확정 후", "→ 예산을 넘으면 경보 후보"],
  ];
  rules.forEach(([a, b], i) => {
    const y = 2.3 + i * 0.78;
    s.addText([{ text: a, options: { bold: true, fontSize: 14, breakLine: true } }, { text: b, options: { fontSize: 12 } }], {
      shape: S.RECTANGLE, x: rx, y, w: rw, h: 0.62, fill: { color: C.background1 }, line: { color: C.text1, width: 1 },
      align: "left", valign: "middle", margin: [2, 8, 2, 8], color: C.text1,
    });
    if (i < 3) arrow(s, rx + rw / 2, y + 0.62, rx + rw / 2, y + 0.78);
  });
  T(s, [
    { text: "3행 중 2행일 때만 경보", options: { bold: true, fontSize: 18, breakLine: true } },
    { text: "한 번 튄 값은 무시, 30초 안의 경보는 한 건으로 묶음", options: { fontSize: 14 } },
  ], { x: rx, y: 5.5, w: rw + 0.3, h: 0.9, paraSpaceAfter: 4 });
  const ox = 9.35;
  T(s, "출력 5가지 상태", { x: ox, y: 1.8, w: 3.5, h: 0.4, fontSize: 16, bold: true });
  const outs = [
    ["정상", "", C.background1], ["운전변화 알림", "부하만 바뀜", C.accent2], ["센서 점검", "측정 문제", C.accent4], ["경보", "이상 의심", C.accent3], ["판단 보류", "정보 부족", C.background2],
  ];
  outs.forEach(([n, why, c], i) => {
    const y = 2.55 + i * 0.62;
    node(s, ox + 0.27, y, 0.5, "", { fill: c, lw: 1.25 });
    T(s, [{ text: n, options: { fontSize: 16, bold: true } }, { text: why ? "  " + why : "", options: { fontSize: 12, color: C.text2 } }], { x: ox + 0.7, y: y - 0.2, w: 2.95, h: 0.4, valign: "middle" });
  });
  T(s, D("정상 보정 구간의 경보 후보 비율", "평균 0.42% (목표 1% 이하)", 14), { x: 0.5, y: 6.45, w: 12.4, h: 0.4 });
}
{
  const s = content("전문가 B ① 전류 파형 모양 보기", "진폭이 정상이어도 파형 모양이 깨진 이상을 잡는다 — 전류 한 주기(약 1.65초)가 쌓이면 모양 전체를 비교", 10);
  const bh = 0.8, mainY = 2.55, topY = 1.85, botY = 3.25;
  const X = { hist: 0.5, len: 2.55, br: 4.65, platt: 7.25, pers: 9.3, alarm: 11.35 };
  box(s, X.hist, mainY, 1.6, bh, "구간 이력", "현재 + 과거 행");
  box(s, X.len, mainY, 1.6, bh, "길이 분기", "16관측 ≈ 1.65초");
  box(s, X.br, topY, 2.1, bh, "MiniRocket", "길 때 : 파형 전체 모양", { fill: C.accent4 });
  box(s, X.br, botY, 2.1, bh, "GBDT", "짧을 때 : 요약 특징 33개", { fill: C.accent4 });
  box(s, X.platt, mainY, 1.6, bh, "확률 보정", "Platt scaling");
  box(s, X.pers, mainY, 1.6, bh, "3회 중 2회", "이어질 때만");
  box(s, X.alarm, mainY, 1.4, bh, "경보", null, { fill: C.accent3, size: 16 });
  const my = mainY + bh / 2;
  arrow(s, X.hist + 1.6, my, X.len, my);
  const sx = 4.35;
  arrow(s, X.len + 1.6, my, sx, my, { noHead: true });
  arrow(s, sx, topY + bh / 2, sx, botY + bh / 2, { noHead: true });
  arrow(s, sx, topY + bh / 2, X.br, topY + bh / 2);
  arrow(s, sx, botY + bh / 2, X.br, botY + bh / 2);
  const mx = 7.0;
  arrow(s, X.br + 2.1, topY + bh / 2, mx, topY + bh / 2, { noHead: true });
  arrow(s, X.br + 2.1, botY + bh / 2, mx, botY + bh / 2, { noHead: true });
  arrow(s, mx, topY + bh / 2, mx, botY + bh / 2, { noHead: true });
  arrow(s, mx, my, X.platt, my);
  arrow(s, X.platt + 1.6, my, X.pers, my);
  arrow(s, X.pers + 1.6, my, X.alarm, my);
  img(s, "image16", 0.5, 4.5, { w: 6.0, alt: "crest factor 식" });
  T(s, [
    ...D("crest factor", "최댓값 ÷ RMS → 매끈한 사인파면 √2 ≈ 1.414", 14, false),
    { text: "정상 1.32–1.54,  모양이 깨진 이상 1.5–3.5", options: { fontSize: 14 } },
  ], { x: 0.5, y: 5.5, w: 6.5, h: 0.7, paraSpaceAfter: 2 });
  T(s, D("MiniRocket", "9,996개 특징으로 '매끈하게 반복하는가'를 봄", 14), { x: 0.5, y: 6.3, w: 6.5, h: 0.65 });
  T(s, "모양 정보를 더할수록 정확해짐 (개발 검증 AP)", { x: 7.4, y: 4.55, w: 5.5, h: 0.4, fontSize: 16, bold: true });
  const steps = [["요약 통계", "0.950"], ["+ crest factor", "0.980"], ["+ 파형 모양", "0.992"]];
  steps.forEach(([n, v], i) => {
    const x = 7.4 + i * 1.9;
    T(s, v, { x, y: 5.05, w: 1.6, h: 0.6, fontSize: 28, bold: true, valign: "middle" });
    T(s, n, { x, y: 5.7, w: 1.7, h: 0.35, fontSize: 14 });
    if (i < 2) arrow(s, x + 1.25, 5.35, x + 1.8, 5.35, { color: C.accent1, width: 1.5 });
  });
}
{
  const s = content("전문가 B ② 확률 보정과 모델 선택", "점수를 믿을 수 있는 확률로 바꾸고, 경보 기준 0.40은 세 가지 근거에서 같은 값이 나온다", 11);
  T(s, D("확률 보정", "Platt scaling — 확률 오차(Brier)가 가장 작음", 16), { x: 0.5, y: 1.8, w: 5.9, h: 0.4 });
  T(s, "Brier (낮을수록 좋음) : 보정 전 0.00434 → Platt 0.00338", { x: 0.5, y: 2.25, w: 5.9, h: 0.35, fontSize: 13 });
  img(s, "image17", 0.5, 2.8, { w: 5.9, alt: "Platt scaling과 임계값 식" });
  T(s, "경보 기준 0.40, 세 근거가 일치", { x: 0.5, y: 3.95, w: 5.9, h: 0.4, fontSize: 16, bold: true });
  T(s, [
    ...D("비용", "놓침 비용을 오경보의 1.5배로 보면 1 / (1 + 1.5) = 0.40", 14, false),
    ...D("F1 최적", "이론값 0.469 (F1이 평평한 구간 안)", 14, false),
    ...D("정상 구간 1% 보장", "0.354", 14, false),
    { text: "정답을 보고 고른 최적값 0.248과도 F1 차이 0.003", options: { fontSize: 13, color: C.text2 } },
  ], { x: 0.5, y: 4.4, w: 5.9, h: 1.55, paraSpaceAfter: 5 });
  T(s, "7계열 20개 구성 비교 → 후보 확인 구간에서 선정", { x: 6.9, y: 1.8, w: 6.0, h: 0.4, fontSize: 16, bold: true });
  table(s, ["후보", "AP", "경보/100구간"], [
    ["하이브리드 (crest factor 제외)", "0.859", "13.5"],
    ["정상 전용 신경망", "0.597", "30.2"],
    ["하이브리드 (crest factor 포함)", "0.909", "9.4"],
    ["기준 GBDT", "0.711", "13.5"],
  ], { x: 6.9, y: 2.3, w: 6.0, colW: [3.3, 0.9, 1.8], align: ["left", "right", "right"], boldRow: 2 });
  T(s, D("선정 규칙 (결과 보기 전에 문서로 고정)", "후보 ≤ 3개, 경보 ≤ 기준의 1.5배 중 AP 최고", 14), { x: 6.9, y: 4.4, w: 6.0, h: 0.7 });
  T(s, [
    { text: "최종 구간 (1회만 평가) : ", options: { bold: true, fontSize: 18 } },
    { text: "정밀도 1.000 · 오경보 0 / 4,090행 · 재현율 0.726 · F1 0.841", options: { fontSize: 18, breakLine: true } },
    { text: "정상의 70%가 처음 보는 저부하 운전이었지만, 정상의 최대 확률은 0.245로 기준 0.40에 한참 못 미침", options: { fontSize: 14 } },
  ], { x: 0.5, y: 6.05, w: 12.4, h: 0.85, paraSpaceAfter: 4 });
}
{
  const s = content("두 전문가 결합 : 서로의 빈틈을 메운다", "A는 놓치지 않는 쪽(재현율), B는 틀리지 않는 쪽(정밀도) — 합치면 이상 구간 5개 모두 탐지", 13);
  img(s, "image19", 0.5, 1.8, { w: 7.0, alt: "OR rule 결합 식" });
  T(s, "계산 전에 정한 규칙 : A 경보 또는 B 확률 ≥ 0.40", { x: 7.8, y: 1.85, w: 5.1, h: 0.45, fontSize: 12, color: C.text2, valign: "middle" });
  table(s, ["시스템", "F1", "정밀도", "재현율", "이상 구간", "정상 오경보"], [
    ["전문가 A", "0.911", "0.847", "0.985", "5/5", "24행"],
    ["전문가 B", "0.841", "1.000", "0.726", "3/5", "0행"],
    ["A+B 결합", "0.915", "0.848", "0.993", "5/5", "24행"],
  ], { x: 0.5, y: 2.65, w: 6.6, colW: [1.35, 0.95, 1.0, 1.0, 1.1, 1.2], align: ["left", "right", "right", "right", "right", "right"], boldRow: 2 });
  T(s, [
    ...D("A", "B가 전혀 못 본 저진폭 이상 구간 2곳을 잡음", 14, false),
    ...D("B", "오경보 0행 — A의 오경보 24행이 몰린 정상 구간에서도 조용", 14, false),
    ...D("현장 선택", "오경보가 부담이면 B 단독, 놓침이 치명적이면 A+B", 14, false),
    ...D("운영 중 개선", "B만 재학습 → F1 0.841 → 0.911, 구간 3 → 5", 14),
  ], { x: 0.5, y: 4.35, w: 6.6, h: 2.3, paraSpaceAfter: 7 });
  bar(s, ["둘 다 경보", "A만 경보", "B만 경보", "둘 다 놓침"], [{ name: "행 수", values: [97, 36, 1, 1] }], {
    x: 7.6, y: 2.55, w: 5.3, h: 4.0, colors: [HEX.accent1], title: "이상 135행을 누가 잡았나 (행)", fmt: "0", max: 110, name: "판정 분담 차트",
  });
}

// ======================================================================
// 15-19. Results
// ======================================================================
section = "실험 결과";
pres.addSection({ title: section });
{
  const s = content("결과 ① 같은 부하끼리 비교하면", "부하 변화가 더 이상 오경보로 이어지지 않는다", 12);
  img(s, "image18", 0.4, 1.75, { h: 4.75, alt: "부하 구간별 하부 진동과 고정 기준 오경보율" });
  const x = 8.75, w = 4.15;
  T(s, "92.6%", { x, y: 1.85, w, h: 0.6, fontSize: 32, bold: true, valign: "middle" });
  T(s, "하부 진동 변동 중 부하 구간이 설명하는 몫", { x, y: 2.48, w, h: 0.6, fontSize: 14 });
  T(s, "22.3% → 3.7–5.7%", { x, y: 3.3, w, h: 0.6, fontSize: 28, bold: true, valign: "middle" });
  T(s, "고부하 정상 구간의 오경보 (고정 기준 → 부하 조건)", { x, y: 3.93, w, h: 0.6, fontSize: 14 });
  T(s, "55.5% → 8.6%", { x, y: 4.75, w, h: 0.6, fontSize: 28, bold: true, valign: "middle" });
  T(s, "처음 보는 정상 운전을 이상으로 오인한 비율", { x, y: 5.38, w, h: 0.6, fontSize: 14 });
  T(s, "기준 : 재사용 검증 구간 교차검증 · 10개 방법 동일 조건 비교", { x: 0.4, y: 6.65, w: 12.5, h: 0.3, fontSize: 12, color: C.text2 });
}
{
  const s = content("결과 ② 공식 라벨 최종 평가", "시간순으로 나눈 마지막 구간을 단 한 번 평가 (점 보정 없음) — 정상 4,090행 · 이상 135행 · 이상 구간 5개", 14);
  img(s, "image20", 0.9, 1.7, { w: 11.5, alt: "공식 라벨 시간순 최종 평가: 모델별 F1·정밀도·재현율과 구간별 검출" });
  T(s, [
    { text: "A+B : 이상 구간 5개 모두 탐지 — A가 B의 놓침 36행을, B가 A의 놓침 1행을 메움 (재현율 0.993, F1 0.915)", options: { bold: true, fontSize: 16, breakLine: true } },
    { text: "이상 시작부터 첫 경보까지 최대 0.1초 (중앙값 0초) · 비교 : 기준 GBDT는 F1 0.660, 구간 4/5", options: { fontSize: 14 } },
  ], { x: 0.9, y: 6.5, w: 12.0, h: 0.75, paraSpaceAfter: 4 });
}
{
  const s = content("결과 ③ 10개 방법 비교 : 경보는 가장 적게", "같은 데이터 · 같은 경보 예산에서 비교 — 경보 부담은 가장 적고, 약한 고장 탐지는 상위권", 15);
  img(s, "image21", 0.4, 1.7, { h: 4.75, alt: "10개 방법의 경보 부담과 순탐지 비교" });
  const x = 8.75, w = 4.15;
  [
    ["경보가 가장 적음", "정상 운전 중 경보 11.7건/시간\n(같은 전류를 쓰는 연속 회귀 33.6건)"],
    ["약한 고장도 잡음", "가상 고장 순탐지 0.256, 상위권\n(시계열 기반 모델 Chronos-2 0.039)"],
    ["정상 오경보 1%대", "전체 1.14% · 가동 구간 1.35%"],
    ["상태를 나눠 알림", "운전변화 · 판단 보류는 경보와 따로 셈"],
  ].forEach(([h, b], i) => {
    T(s, [{ text: h, options: { bold: true, fontSize: 16, breakLine: true } }, { text: b, options: { fontSize: 14 } }], { x, y: 1.8 + i * 1.15, w, h: 1.0, paraSpaceAfter: 3 });
  });
  T(s, "순탐지 = 고장 탐지율 − 정상 오표시율 · 비교 : LightGBM, Isolation Forest, 오토인코더, 마할라노비스 거리, 연속 공변량 회귀, Chronos-2(zero-shot), 상태조건 신경망 등", { x: 0.4, y: 6.55, w: 12.5, h: 0.55, fontSize: 12, color: C.text2 });
}
{
  const s = content("결과 ④ 오경보와 놓침은 어디서 생기나", "오경보는 관측 재개 직후에, 놓침은 짧은 이력에 몰린다 → 조건마다 설계로 막음", 16);
  img(s, "image22", 1.45, 1.65, { w: 10.4, alt: "오경보는 구간 시작에, 미탐은 짧은 이력에 집중" });
  T(s, [
    ...D("재개 직후 오경보", "첫 관측 3.9% → 3번째부터 0% · 3번 중 2번 규칙으로 즉시 경보 60.8건/h → 0", 14, false),
    ...D("고부하 오경보", "같은 부하끼리 비교 → 처음 보는 최고 부하에서 52.1% → 1.82%", 14, false),
    ...D("짧은 이력 놓침", "B가 놓친 37행 중 34행이 16관측 미만 → 이 구간은 A가 5개 모두 잡음", 14),
  ], { x: 0.5, y: 5.95, w: 12.4, h: 1.15, paraSpaceAfter: 3 });
}
{
  const s = content("결과 ⑤ 다른 설비에서도 통할까", "외부 유압 시험대에서도 부하 조건화가 오경보를 19.0% → 10.4%로 줄였다", 22);
  img(s, "image28", 1.65, 1.62, { w: 10.0, alt: "외부 유압 시험대에서의 오경보율 비교와 전이 결과" });
  T(s, [
    ...D("실험", "UCI 유압 시험대 (CC BY 4.0) · 외부→대회, 대회→외부, 공동 학습 · 비교 Isolation Forest 34.7%, 고정 임계값 65.9%", 13, false),
    ...D("한계", "다른 장비에서 학습한 모델을 그대로 쓰면 오경보 ≈ 100%", 13, false),
    ...D("그래서 현장에서는", "현장 정상 데이터로 기준을 다시 맞춘다 (대회 자료로 재학습한 축소 모델은 이상 16/17 버스트 탐지)", 13),
  ], { x: 0.5, y: 6.05, w: 12.4, h: 1.1, paraSpaceAfter: 3 });
}

// ======================================================================
// 20-24. Type 2 — correction loop & grounded explanation
// ======================================================================
section = "유형 2";
pres.addSection({ title: section });
{
  const s = content("교정 루프 : 배우되, 고장은 지우지 않는다", "더 민감해질 수는 있어도, 물리 기준보다 덜 민감해지지는 않는다", 17);
  const cx = 4.4, cy = 4.25, R = 1.8, d = 0.62;
  const labels = [
    ["운영 · 판정 · 설명", "right"],
    ["작업자 확인", "right"],
    ["교정 제안", "right"],
    ["canary test", "left"],
    ["적용 · 되돌리기", "left"],
    ["주기 재보정", "left"],
  ];
  const pts = labels.map((_, i) => {
    const a = (-90 + i * 60) * Math.PI / 180;
    return [cx + R * Math.cos(a), cy + R * Math.sin(a)];
  });
  pts.forEach((p, i) => edge(s, p, pts[(i + 1) % 6], d / 2 + 0.03, d / 2 + 0.03, { color: C.accent1, width: 1.75 }));
  pts.forEach((p, i) => {
    const isCanary = i === 3;
    node(s, p[0], p[1], d, String(i + 1), { fill: isCanary ? C.accent3 : C.background1, bold: true });
    const [txt, side] = labels[i];
    if (side === "right") T(s, txt, { x: p[0] + d / 2 + 0.15, y: p[1] - 0.2, w: 2.4, h: 0.4, fontSize: 16, bold: isCanary, valign: "middle" });
    else T(s, txt, { x: p[0] - d / 2 - 0.15 - 2.0, y: p[1] - 0.2, w: 2.0, h: 0.4, fontSize: 16, bold: isCanary, valign: "middle", align: "right" });
  });
  const cp = pts[3];
  arrow(s, cp[0] + d / 2, cp[1], cp[0] + 1.35, cp[1], { color: C.accent6, dash: "dash", width: 1.5 });
  T(s, [
    { text: "하나라도 실패하면 거절", options: { fontSize: 13, bold: true, breakLine: true } },
    { text: "엔지니어 검토로", options: { fontSize: 12, color: C.text2 } },
  ], { x: cp[0] + 1.45, y: cp[1] - 0.3, w: 2.3, h: 0.6, valign: "middle" });
  const x = 8.4, w = 4.5;
  T(s, "92.3% → 64.6%", { x, y: 1.9, w, h: 0.6, fontSize: 28, bold: true, valign: "middle" });
  T(s, "작업자 확인 반영 후 운전변화 알림 (운영 시험)", { x, y: 2.52, w, h: 0.6, fontSize: 14 });
  T(s, "11 / 16", { x, y: 3.3, w, h: 0.6, fontSize: 28, bold: true, valign: "middle" });
  T(s, "고장이 섞인 교정 제안을 거절", { x, y: 3.92, w, h: 0.35, fontSize: 14 });
  T(s, [
    ...D("canary test", "가짜 고장(진동 +30% · +50%)을 넣어 보고, 탐지가 떨어지면 거절", 14, false),
    ...D("시험 3종", "새 운전 · 기존 운전 · 정상 확인", 14, false),
    ...D("제안 조건", "확정 구간 ≥ 6개 (각 ≥ 12행)", 14),
  ], { x, y: 4.65, w, h: 1.9, paraSpaceAfter: 8 });
}
{
  const s = content("유형 2 : 판정을 사람의 말로 설명", "AI는 근거를 설명하고, 사람을 부를지는 규칙표가 정한다", 18);
  const bw = 1.85, bh = 0.85, gap = 0.35, x0 = 0.5, r1 = 1.85, r2 = 3.15;
  const xs = [x0, x0 + bw + gap, x0 + 2 * (bw + gap)];
  box(s, xs[0], r1, bw, bh, "깨우기", "변화 · 요청 · 질문");
  box(s, xs[1], r1, bw, bh, "맥락 모으기", "읽기 전용 도구 11개");
  box(s, xs[2], r1, bw, bh, "지식그래프", "근거 검색");
  box(s, xs[2], r2, bw, bh, "답변 작성", "언어모델", { fill: C.accent4 });
  box(s, xs[1], r2, bw, bh, "validator", "17개 검사");
  box(s, xs[0], r2, bw, bh, "답변 카드", "작업자 화면", { fill: C.accent3 });
  arrow(s, xs[0] + bw, r1 + bh / 2, xs[1], r1 + bh / 2);
  arrow(s, xs[1] + bw, r1 + bh / 2, xs[2], r1 + bh / 2);
  arrow(s, xs[2] + bw / 2, r1 + bh, xs[2] + bw / 2, r2);
  arrow(s, xs[2], r2 + bh / 2, xs[1] + bw, r2 + bh / 2);
  arrow(s, xs[1], r2 + bh / 2, xs[0] + bw, r2 + bh / 2);
  const ry = 4.45;
  box(s, x0, ry, 3 * bw + 2 * gap, 0.7, "호출 규칙표", "13개 규칙이 등급을 정함 — AI가 바꿀 수 없음");
  arrow(s, xs[0] + bw / 2, ry, xs[0] + bw / 2, r2 + bh, { dash: "dash", color: C.accent6, width: 1.5 });
  T(s, [
    ...D("답변 5요소", "어디서 · 무엇이 변했나 · 원인 후보 · 먼저 점검 · 긴급도", 14, false),
    ...D("예시", "“방금 알림은 왜 났어?” → 운전변화 알림, 등급 '관찰 강화'", 14, false),
    ...D("속도", "평균 30.7초 → 초 단위 보호가 아니라 운영자 지원", 14),
  ], { x: 0.5, y: 5.45, w: 6.6, h: 1.4, paraSpaceAfter: 5 });
  const im = img(s, "image25", 7.45, 1.8, { w: 5.45, alt: "실제 답변 화면: 운전변화 알림, 화면 시점 121초, 관찰 강화 등급" });
  s.addShape(S.RECTANGLE, { x: 7.45, y: 1.8, w: im.w, h: im.h, fill: { type: "none" }, line: { color: C.background2, width: 0.75 } });
  T(s, "재생 자료의 실제 답변 화면 (화면 121초 기준 · 심사 시 시연 가능)", { x: 7.45, y: 1.8 + im.h + 0.08, w: 5.45, h: 0.3, fontSize: 12, color: C.text2 });
  T(s, D("평가", "질문 20쌍에서 근거 인용률 11% → 100%", 14), { x: 7.45, y: 5.75, w: 5.45, h: 0.7 });
}
{
  const s = content("유형 2 ① 지식그래프에서 근거 찾기", "임베딩 없이 정해진 규칙으로 찾는다 — 같은 질문과 상태에는 항상 같은 근거", 19);
  const d = 0.5, r = d / 2;
  const N = {
    s: [1.9, 3.6],
    a: [3.3, 2.5], b: [3.5, 3.85], c: [2.5, 5.0], hub: [0.9, 2.4],
    a2: [4.9, 2.0], b2: [5.2, 3.2], b3: [5.0, 4.6], c2: [3.9, 5.5],
    o1: [0.7, 4.9], o2: [0.4, 3.5],
  };
  const hop1 = [["s", "a"], ["s", "b"], ["s", "c"], ["s", "hub"]];
  const hop2 = [["a", "a2"], ["b", "b2"], ["b", "b3"], ["c", "c2"]];
  const other = [["hub", "o1"], ["hub", "o2"], ["hub", "a"], ["o1", "c"]];
  other.forEach(([u, v]) => edge(s, N[u], N[v], r, r, { noHead: true, color: C.background2, width: 1.25 }));
  hop2.forEach(([u, v]) => edge(s, N[u], N[v], r, r, { noHead: true, color: C.accent1, width: 1.5 }));
  hop1.forEach(([u, v]) => edge(s, N[u], N[v], r, r, { noHead: true, color: C.accent2, width: 2.5 }));
  node(s, ...N.s, 0.62, "s", { fill: C.accent3, italic: true, font: "Times New Roman", size: 18 });
  ["a", "b", "c"].forEach((k) => node(s, ...N[k], d, "", { fill: C.accent4 }));
  node(s, ...N.hub, 0.62, "hub", { fill: C.accent4, size: 11 });
  ["a2", "b2", "b3", "c2"].forEach((k) => node(s, ...N[k], d, "", { fill: C.accent2, transp: 60 }));
  ["o1", "o2"].forEach((k) => node(s, ...N[k], d, "", { line: C.text2, lw: 1 }));
  const ly = 6.2;
  [["출발점", C.accent3, 0], ["1칸", C.accent4, 0], ["2칸", C.accent2, 60], ["선택 안 됨", C.background1, 0]].forEach(([n, c, tr], i) => {
    node(s, 0.62 + i * 1.5, ly, 0.24, "", { fill: c, transp: tr, lw: 1 });
    T(s, n, { x: 0.82 + i * 1.5, y: ly - 0.16, w: 1.2, h: 0.32, fontSize: 12, valign: "middle" });
  });
  T(s, "hub : 연결이 많은 노드 → 점수를 깎아 한쪽으로 쏠리지 않게 함", { x: 0.5, y: 6.55, w: 6.2, h: 0.3, fontSize: 12, color: C.text2 });
  const x = 6.9, w = 6.0;
  [
    ["출발점 정하기", "질문 속 단어(한 · 영 별칭)와 현재 판정 상태에서 시작 노드를 고름"],
    ["두 칸까지 퍼뜨리기", "연결을 따라 점수 전파, 한 칸마다 ×0.55 · '점검', '원인', '호출' 등 질문 의도별 가중"],
    ["상위 근거 고르기", "노드 8 · 경로 5 · 출처 6 · 주제 요약 2를 모델에 전달"],
  ].forEach(([h, b], i) => {
    const y = 1.85 + i * 1.05;
    node(s, x + 0.22, y + 0.22, 0.44, String(i + 1), { size: 14, bold: true, lw: 1.25 });
    T(s, [{ text: h, options: { bold: true, fontSize: 16, breakLine: true } }, { text: b, options: { fontSize: 14 } }], { x: x + 0.65, y, w: w - 0.65, h: 0.95, paraSpaceAfter: 2 });
  });
  img(s, "image27", x, 5.05, { w: 5.6, alt: "2-hop 점수 전파 식" });
  T(s, [
    { text: "205 · 320 · 76", options: { bold: true, fontSize: 16 } },
    { text: "  노드 · 연결 · 출처 — 설비 부위, 원인 후보, 신호, 판정 상태, 점검 순서, 문헌 근거 (모두 출처 · 근거 등급 표기)", options: { fontSize: 13 } },
  ], { x, y: 5.85, w, h: 0.7 });
}
{
  const s = content("유형 2 ② AI 답변 검사 (validator)", "AI의 답은 17개 검사를 모두 통과해야 화면에 나간다 — 검사는 AI가 아니라 규칙(정규식 · 구조 비교)", 20);
  T(s, "언어모델이 일하는 방식", { x: 0.5, y: 1.8, w: 5.4, h: 0.4, fontSize: 18, bold: true });
  T(s, [
    ...D("도구", "읽기 전용 조회 11개 + 답안 제출 1개 — 데이터 대신 도구 결과만 봄", 14, false),
    ...D("속도", "첫 턴에 도구 6개를 한꺼번에 호출 → 요청 5–8회 → 2회 안팎", 14, false),
    ...D("한도", "최대 9턴 · 서버 기한 80초", 14),
  ], { x: 0.5, y: 2.3, w: 5.4, h: 2.0, paraSpaceAfter: 8 });
  T(s, "검사에 실패하면", { x: 0.5, y: 4.5, w: 5.4, h: 0.4, fontSize: 16, bold: true });
  const fx = [0.5, 2.35, 4.2], fw = 1.55, fy = 5.0, fh = 0.7;
  box(s, fx[0], fy, fw, fh, "실패", "위반 목록 전달", { fill: C.accent3 });
  box(s, fx[1], fy, fw, fh, "고쳐 쓰기", "1회");
  box(s, fx[2], fy, fw, fh, "대체 답변", "규칙 기반");
  arrow(s, fx[0] + fw, fy + fh / 2, fx[1], fy + fh / 2);
  arrow(s, fx[1] + fw, fy + fh / 2, fx[2], fy + fh / 2);
  T(s, [
    { text: "대체 답변도 같은 검사를 거침", options: { breakLine: true } },
    { text: "통과 못 한 문장은 화면에 안 나감" },
  ], { x: 0.5, y: 5.9, w: 5.4, h: 0.65, fontSize: 14, bold: true, paraSpaceAfter: 2 });
  T(s, "17개 검사 — AI가 하면 안 되는 것", { x: 6.5, y: 1.8, w: 6.4, h: 0.4, fontSize: 18, bold: true });
  table(s, ["묶음", "막는 것"], [
    ["숫자 · 단위", "도구에 없는 숫자 · 단위 만들기"],
    ["단정 금지", "원인 확정, 남은 수명 예측"],
    ["제어 금지", "정지 · 재가동 · 설정 변경 권고"],
    ["근거 연결", "근거 없는 가설, 없는 출처 인용"],
    ["신선도", "120초 넘은 데이터로 답하기"],
    ["주입 방어", "자료 속 지시문 따르기"],
    ["운영자 말투", "p값 · 도구 이름 같은 내부 용어"],
    ["호출 등급", "규칙표와 다른 등급 말하기"],
  ], { x: 6.5, y: 2.3, w: 6.4, colW: [1.9, 4.5], rowH: 0.45 });
}
{
  const s = content("유형 2 ③ 사람을 부르는 규칙표", "누구를 언제 부를지는 AI가 아니라 13개 규칙이 정한다 — 같은 입력이면 항상 같은 등급", 21);
  T(s, "4단계 호출 등급 (위에서부터 처음 맞는 규칙)", { x: 0.5, y: 1.8, w: 7, h: 0.4, fontSize: 16, bold: true });
  table(s, ["등급", "언제", "다시 확인"], [
    ["즉시 현장 확인 후보", "경보 + 센서 정상 + 높음 지속", "120초"],
    ["담당자 호출", "경보지만 센서 의심 · 운전 변화와 겹침", "300초"],
    ["관찰 강화", "판단 보류, 운전변화 알림만", "600초"],
    ["기록 유지", "정상", "1,800초"],
  ], { x: 0.5, y: 2.3, w: 6.9, colW: [2.2, 3.5, 1.2], align: ["left", "left", "right"], rowH: 0.48 });
  T(s, [
    { text: "'즉시 현장 확인 후보'도 정지 판단이 아님 (기록 · 보류 · 알림만)", options: { fontSize: 14, breakLine: true } },
    { text: "경보는 어떤 입력에서도 '담당자 호출' 아래로 내려가지 않음", options: { fontSize: 14, breakLine: true } },
    { text: "입력 조합 10,206개 전수 검사 → 위반 0", options: { fontSize: 16, bold: true } },
  ], { x: 0.5, y: 4.85, w: 6.9, h: 1.3, paraSpaceAfter: 6 });
  bar(s, ["근거 인용", "규칙표와 등급 일치"], [
    { name: "도입 전", values: [11, 60] },
    { name: "도입 후 (지식그래프 + 규칙표)", values: [100, 100] },
  ], { x: 7.8, y: 1.75, w: 5.1, h: 4.2, colors: ["A6A6A6", HEX.accent1], title: "도입 전 → 후 (질문 20쌍, %)", fmt: "0", max: 115, gap: 60, name: "도입 전후 차트" });
  T(s, "질문 20쌍(실제 재생 12 + 시뮬레이션 8) : 7개 채점 기준 충족 99% → 98% (동등) · 설비 제어 지시 0건", { x: 0.5, y: 6.4, w: 12.4, h: 0.4, fontSize: 14 });
}

// ======================================================================
// 25-28. Field deployment & conclusion
// ======================================================================
section = "현장 적용 및 결론";
pres.addSection({ title: section });
{
  const s = content("현장 적용 : 누가 무엇을 하나", "알림은 3단계 + 데이터 점검 — 읽기 전용 병행 운영으로 시작해 현장 시험을 통과하면 전환", 23);
  T(s, "현장 알림과 담당자 조치", { x: 0.5, y: 1.8, w: 5.6, h: 0.4, fontSize: 16, bold: true });
  [
    ["특이 없음", "← 정상", "주기 점검 유지", C.background1],
    ["관찰", "← 운전변화 알림 · 판단 보류", "운전원이 작업 기록과 대조해 피드백", C.accent2],
    ["점검 권고", "← 경보", "정비 담당이 근거 확인 후 점검", C.accent3],
    ["데이터 점검", "← 센서 점검 알림", "계측 담당이 센서 · 케이블 · 수집 장치 확인", C.accent4],
  ].forEach(([n, from, act, c], i) => {
    const y = 2.6 + i * 0.95;
    node(s, 0.8, y, 0.55, "", { fill: c, lw: 1.25 });
    T(s, [
      { text: n, options: { bold: true, fontSize: 16 } },
      { text: "  " + from, options: { fontSize: 13, color: C.text2, breakLine: true } },
      { text: "→ " + act, options: { fontSize: 14 } },
    ], { x: 1.3, y: y - 0.36, w: 5.1, h: 0.75, valign: "middle", paraSpaceAfter: 2 });
  });
  T(s, [
    { text: "설비와는 읽기 전용으로만 연동", options: { breakLine: true } },
    { text: "호출 규칙 13개, 10,206개 조합에서 위반 0" },
  ], { x: 0.5, y: 6.2, w: 5.8, h: 0.65, fontSize: 13, paraSpaceAfter: 2 });
  const x = 6.9;
  T(s, "단계적 도입", { x, y: 1.8, w: 5, h: 0.4, fontSize: 16, bold: true });
  const steps = [
    ["설치 점검", "진동 채널(AI0/AI1) 매핑 · 전류 0 확인"],
    ["읽기 전용 병행 운영", "현장 정상 기록으로 기준 설정"],
    ["모델 · 임계값 동결", "시험 중에는 바꾸지 않음"],
    ["다른 날짜 현장 시험", "정상 오경보 95% 상한 ≤ 1%, 경보 ≤ 6건/시간"],
    ["운영 전환", "기준을 통과하면 전환"],
  ];
  const d = 0.5, y0 = 2.6, dy = 0.88;
  arrow(s, x + d / 2, y0 + d / 2, x + d / 2, y0 + 4 * dy, { noHead: true, color: C.accent1, width: 1.5 });
  steps.forEach(([h, b], i) => {
    const y = y0 + i * dy;
    node(s, x + d / 2, y, d, String(i + 1), { size: 14, bold: true, fill: i === 4 ? C.accent5 : C.background1 });
    T(s, [{ text: h, options: { bold: true, fontSize: 16, breakLine: true } }, { text: b, options: { fontSize: 13 } }], { x: x + 0.75, y: y - 0.3, w: 5.25, h: 0.75 });
  });
}
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText("시연", { placeholder: "title" });
  s.addText("심사 때 현장에서 실제 시뮬레이션 화면으로 시연", { placeholder: "desc", bullet: false });
  img(s, "image34", 4.92, 1.9, { w: 3.5, alt: "시연 페이지 QR 코드" });
  T(s, "QR 코드로 시뮬레이션 화면 접속", { x: 4.0, y: 5.6, w: 5.33, h: 0.4, fontSize: 16, align: "center" });
  s.addNotes("심사 시 현장 시연이 가능합니다. QR 코드로 실제 시뮬레이션 화면에 접속할 수 있습니다.");
}
{
  const s = content("결론 및 향후 과제", "작업자 경험에 좌우되던 판단에서, 공정과 사람이 대화하는 운영으로", 25);
  T(s, "기여", { x: 0.5, y: 1.8, w: 6, h: 0.4, fontSize: 16, bold: true });
  const fills = [C.accent3, C.accent4, C.accent2, C.accent5];
  const items = [
    ["부하를 알고 비교한다", "처음 보는 정상 운전 오인 55.5% → 8.6%"],
    ["경보는 적게, 상태는 자세히", "경보 11.7건/시간, 10개 방법 중 최소"],
    ["고장은 지우지 않는 교정", "운전변화 알림 92.3% → 64.6%, 고장 섞인 교정 11/16 거절"],
    ["AI는 설명만, 호출은 규칙표", "근거 인용 11% → 100%, 규칙표 위반 0"],
  ];
  const d = 0.55, y0 = 2.55, dy = 0.95, x = 0.5;
  arrow(s, x + d / 2, y0, x + d / 2, y0 + 3 * dy, { noHead: true, color: C.accent1, width: 1.5 });
  items.forEach(([h, b], i) => {
    const y = y0 + i * dy;
    node(s, x + d / 2, y, d, String(i + 1), { fill: fills[i], bold: true, size: 14 });
    T(s, [{ text: h, options: { bold: true, fontSize: 16, breakLine: true } }, { text: b, options: { fontSize: 13 } }], { x: x + 0.8, y: y - 0.3, w: 6.5, h: 0.75 });
  });
  const nx = 8.0, nw = 4.6;
  T(s, "향후 과제", { x: nx, y: 1.8, w: nw, h: 0.4, fontSize: 16, bold: true });
  const nexts = ["현장 정상 데이터로 재보정", "다른 날짜 현장 시험", "고장 사례 축적 · 지식그래프 확장"];
  nexts.forEach((n, i) => {
    const y = 2.35 + i * 1.0;
    box(s, nx, y, nw, 0.62, n, null, { size: 15 });
    if (i < 2) arrow(s, nx + nw / 2, y + 0.62, nx + nw / 2, y + 1.0);
  });
  T(s, [
    { text: "최종 결합(A+B) : ", options: { bold: true, fontSize: 16 } },
    { text: "이상 구간 5/5 · 재현율 0.993 · F1 0.915 · 정상 행 오경보 0.6%", options: { fontSize: 16 } },
  ], { x: 0.5, y: 6.35, w: 12.4, h: 0.45, valign: "middle" });
}
{
  const s = pres.addSlide({ masterName: "CLOSING", sectionTitle: section });
  s.addText("감사합니다", { placeholder: "title" });
  T(s, "Q & A", { x: 0, y: 3.8, w: 13.33, h: 0.6, fontSize: 28, fontFace: "Georgia", align: "center", valign: "middle" });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
