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
  s.addText([
    { text: "고장 증거를 지키며 적응하는 산업용 ", options: {} },
    { text: "World Model", options: { fontFace: "Georgia" } },
  ], { placeholder: "subtitle", bullet: false });
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
    "1. 연구 배경 및 문제 정의",
    "2. 데이터 진단",
    "3. 유형 1 판정 : 두 전문가 모델",
    "4. 실험 결과",
    "5. 유형 2 : 교정 루프와 근거 기반 설명",
    "6. 현장 적용 및 결론",
  ];
  items.forEach((it, i) => T(s, it, { x: 3.9, y: 1.72 + i * 0.68, w: 7.2, h: 0.5, fontSize: 24, valign: "middle" }));
  s.addNotes("발표 순서입니다. 문제 정의와 데이터 진단에서 출발해, 유형 1의 두 전문가 판정 모델과 실험 결과를 보고, 유형 2의 교정 루프와 근거 기반 설명, 마지막으로 현장 적용과 결론 순으로 말씀드리겠습니다.");
}

// ======================================================================
// 3. Core contributions
// ======================================================================
{
  const s = content("핵심 기여", "판정부터 설명 · 교정까지 네 가지 기여가 하나의 파이프라인으로 이어짐", 2);
  const cols = [0.5, 3.6, 6.7, 9.8];
  const fills = [C.accent3, C.accent4, C.accent2, C.accent5];
  const transp = [0, 0, 0, 0];
  const data = [
    ["같은 부하 정상 대비 판정", "55.5% → 8.6%", "처음 보는 정상 운전상태 오표시율\n외부 시험대 오경보 19.0% → 10.4%"],
    ["하나의 경보 예산, 5가지 상태", "11.7건 / 시간", "경보 부담, 10개 방법 중 최소\n정상 오경보 1.14% / 1.35%"],
    ["고장 증거를 지키는 교정 루프", "92.3% → 64.6%", "semi-synthetic 정상 알림 비율\n고장 섞인 조건 16개 중 11개 기각"],
    ["근거 설명 + 호출 규칙표", "205 · 320 · 76", "지식그래프 노드 · 연결 · 출처\n호출 규칙 13개, 조합 10,206개 위반 0"],
  ];
  const d = 0.59, cy = 2.05 + d / 2;
  for (let i = 0; i < 3; i++) arrow(s, cols[i] + d, cy, cols[i + 1], cy, { color: C.accent1, width: 1.5 });
  data.forEach(([h, m, t], i) => {
    node(s, cols[i] + d / 2, cy, d, String(i + 1), { fill: fills[i], transp: transp[i], bold: true });
    T(s, h, { x: cols[i], y: 2.9, w: 2.9, h: 0.75, fontSize: 16, bold: true });
    T(s, m, { x: cols[i], y: 3.7, w: 2.9, h: 0.55, fontSize: 24, bold: true, valign: "middle" });
    T(s, t, { x: cols[i], y: 4.35, w: 3.0, h: 0.9, fontSize: 13, paraSpaceAfter: 2 });
  });
  T(s, [
    { text: "최종 결합(A+B), 공식 라벨 시간순 최종 분할", options: { bold: true, fontSize: 16, breakLine: true } },
    { text: "F1 0.915 · 재현율 0.993 · 이상 구간 5/5 · 정상 오경보 24 / 4,090행 (한 구간)", options: { fontSize: 16 } },
  ], { x: 0.5, y: 5.75, w: 12.3, h: 0.85, paraSpaceAfter: 4 });
}

// ======================================================================
// 4. Problem definition
// ======================================================================
{
  const s = content("문제 정의", "정상 부하 변화와 이상을 구분하지 못하면 오경보가 쌓이고, 경보는 신뢰를 잃는다", 3);
  // causal chain
  const bx = [0.5, 3.0, 5.5], bw = 2.0, by = 1.95, bh = 0.7;
  box(s, bx[0], by, bw, bh, "유압 펌프 이상", null, { fill: C.accent3, size: 15 });
  box(s, bx[1], by, bw, bh, "압력 불안정", null, { size: 15 });
  box(s, bx[2], by, bw, bh, "불량 · 가동 중단", null, { size: 15 });
  arrow(s, bx[0] + bw, by + bh / 2, bx[1], by + bh / 2);
  arrow(s, bx[1] + bw, by + bh / 2, bx[2], by + bh / 2);
  T(s, "이상 유형 : 마모 · 정렬 · 캐비테이션 · 밸브 충격 · 계측 이상", { x: 0.5, y: 2.8, w: 7.0, h: 0.35, fontSize: 12, color: C.text2 });
  T(s, [
    ...D("현재 대응", "주기 점검과 작업자 노하우(내재화된 지식)에 의존", 16, false),
    ...D("구분의 어려움", "약한 진동 변화가 정상 운전 변화와 겹침", 16, false),
    ...D("오경보 불신", "정상 운전 변화를 이상으로 오인 → 경보를 믿지 않게 됨", 16),
  ], { x: 0.5, y: 3.4, w: 7.4, h: 1.5, paraSpaceAfter: 10 });
  // stat
  T(s, "22.3%", { x: 8.6, y: 1.85, w: 4.3, h: 1.0, fontSize: 60, bold: true, valign: "middle" });
  T(s, "고정 기준 : 최고 부하 구간 정상 창 중 경보 비율", { x: 8.6, y: 2.9, w: 4.3, h: 0.35, fontSize: 14 });
  T(s, "3.7–5.7%", { x: 8.6, y: 3.45, w: 4.3, h: 0.6, fontSize: 32, bold: true, valign: "middle" });
  T(s, "부하를 반영한 기준", { x: 8.6, y: 4.1, w: 4.3, h: 0.35, fontSize: 14 });
  // problem statement
  T(s, "The Problem", { x: 0.5, y: 5.2, w: 6, h: 0.45, fontSize: 20, bold: true, valign: "middle" });
  T(s, [
    { text: "P1. ", options: { bold: true, fontSize: 18 } },
    { text: "이상을 놓치지 않고 감지할 것", options: { fontSize: 18, breakLine: true } },
    { text: "P2. ", options: { bold: true, fontSize: 18 } },
    { text: "정상 부하 변화를 이상으로 착각하지 않을 것", options: { fontSize: 18 } },
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
  const s = content("데이터 진단 ① 신호 특성", "정상 전류는 매끈한 정현파, 이상은 모양이 깨지고 두 분포가 분리됨");
  if (n4split > 0) s.addNotes(n4.slice(0, n4split).trim()); else s.addNotes(n4);
  img(s, "image7", 0.4, 1.75, { w: 12.5, alt: "데이터 진단: (a) 전류 파형 (b) 파고율 분포 (c) 평면 이탈도 분포" });
  T(s, "(a) 전류 파형   (b) crest factor 분포   (c) plane deviation 분포 — 개발 분할 20행 이상 정상 287 · 이상 8구간", { x: 0.4, y: 5.95, w: 12.5, h: 0.35, fontSize: 12, color: C.text2 });
  T(s, D("샘플링 제약", "0.1초 표본(≤ 5 Hz) → 스펙트럼 대신 시간 영역 모양 특징(crest factor 등) 사용", 16), { x: 0.4, y: 6.45, w: 12.5, h: 0.4 });
}
{
  const s = content("데이터 진단 ② 발견과 설계 대응", "데이터에서 찾은 6가지 특성을 각각 설계 결정으로 연결");
  if (n4split > 0) s.addNotes(n4.slice(n4split).trim());
  const rows = [
    ["5초 이하로 끊긴 구간", "구간 단위 · 인과적 처리"],
    ["정상 전류 = 정현파", "A : 부하 추정 · 품질 경로,  B : 파형 모양 특징"],
    ["부하가 진동 변동의 92.6% 설명", "A : 같은 부하끼리 비교"],
    ["진폭이 정상인 이상", "B : 파형 모양 + A : 상태 조건 잔차"],
    ["관측 재개 직후 오경보", "3회 중 2회 지속 규칙 · 판단 보류"],
    ["라벨 = 날짜 = 저장 형식", "시간순 3분할 + semi-synthetic 고장"],
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
    ...D("주의", "정상 · 이상이 서로 다른 날짜와 저장 형식으로 기록 → 형식만으로도 라벨 구분 가능 (ROC-AUC 1.000)", 14, false),
    { text: "형식 변화 민감도 점검 (정상 전류 양자화 · 직류 성분 제거 · 진동만 학습) : AP 0.968–0.993", options: { fontSize: 14, breakLine: true } },
    { text: "→ 데이터가 단순한 만큼, 고정된 단일 모델은 처음 보는 상황에 범용적으로 쓰기 어렵다", options: { fontSize: 16, bold: true } },
  ], { x: 0.5, y: 5.4, w: 12.4, h: 1.3, paraSpaceAfter: 6 });
}

// ======================================================================
// 7-14. Type 1 — two experts
// ======================================================================
section = "유형 1 판정";
pres.addSection({ title: section });
{
  const s = content("전체 파이프라인", "설비 수집부터 운영자 화면까지 — 유형 2가 느려도 유형 1 판정은 멈추지 않음", 6);
  const xs = [0.55, 3.1, 5.65, 8.2, 10.75], bw = 2.05, by = 1.85, bh = 0.95;
  const st = [
    ["읽기 전용 수집", "0.1초 행", {}],
    ["유형 1 판정", "전문가 A · B", { fill: C.accent3, transp: 20 }],
    ["스냅샷 링", "번호순 보관 (4,096개)", {}],
    ["유형 2 설명", "질문 · 점검 · 요청", { fill: C.accent4 }],
    ["출력", "화면 · 초안 · 일지", {}],
  ];
  st.forEach(([a, b, o], i) => {
    box(s, xs[i], by, bw, bh, a, b, Object.assign({ size: 15 }, o));
    if (i < 4) arrow(s, xs[i] + bw, by + bh / 2, xs[i + 1], by + bh / 2);
  });
  // correction loop
  const ly = 3.25, lh = 0.6;
  box(s, xs[1], ly, xs[3] + bw - xs[1], lh, "교정 루프", "작업자 확인 → canary test → 반영", { dash: "dash", size: 14 });
  path2(s, [[xs[4] + bw / 2, by + bh], [xs[4] + bw / 2, ly + lh / 2], [xs[3] + bw, ly + lh / 2]], { dash: "dash" });
  arrow(s, xs[1] + bw / 2, ly, xs[1] + bw / 2, by + bh, { dash: "dash" });
  // latency table + server facts
  T(s, "구성별 주기 · 지연", { x: 0.55, y: 4.25, w: 5, h: 0.35, fontSize: 16, bold: true });
  table(s, ["구성", "주기 · 지연"], [
    ["전문가 A", "행마다 약 0.2 ms"],
    ["지식그래프 검색", "수 ms"],
    ["모델 호출", "질문당 약 31초"],
    ["validator · 정책표", "수 ms"],
  ], { x: 0.55, y: 4.7, w: 5.8, colW: [2.9, 2.9], rowH: 0.36 });
  T(s, "에지 서버", { x: 7.0, y: 4.25, w: 5, h: 0.35, fontSize: 16, bold: true });
  T(s, [
    ...D("구성", "Worker 1개 + 상태 보관 객체 2개", 14, false),
    ...D("판정 모델", "전문가 A 탑재, 모델 호출은 별도 요청", 14, false),
    ...D("쓰기 경로", "설비로 쓰는 경로 없음 (읽기 전용)", 14, false),
    ...D("배포 검증", "유형 1 20,802 프레임 · 유형 2 28,901건 불일치 0", 14),
  ], { x: 7.0, y: 4.75, w: 5.9, h: 1.8, paraSpaceAfter: 8 });
}
{
  const s = content("유형 1 판정 : 두 전문가 모델", "서로 다른 근거로 서로 다른 이상을 잡는 두 전문가를 OR rule로 결합", 5);
  const inX = 0.5, inW = 1.6, rowA = 2.15, rowB = 3.95, bh = 0.8;
  const midY = (rowA + rowB + bh) / 2;
  box(s, inX, midY - 0.55, inW, 1.1, "전류 · 진동", "0.1초 간격", { size: 15 });
  T(s, "전문가 A : 부하 조건부 world model", { x: 2.75, y: rowA - 0.4, w: 7, h: 0.35, fontSize: 14, bold: true });
  T(s, "전문가 B : 전류 파형 판정기", { x: 2.75, y: rowB - 0.4, w: 7, h: 0.35, fontSize: 14, bold: true });
  const aX = [2.75, 4.75, 6.75, 8.75], aW = 1.65;
  [["부하 추정", "전류 기반"], ["정상과 비교", "같은 부하"], ["건강 점수", "칼만 누적"], ["보정 · 판정", "conformal"]].forEach(([a, b], i) => {
    box(s, aX[i], rowA, aW, bh, a, b, { fill: C.accent2, transp: 75 });
    if (i < 3) arrow(s, aX[i] + aW, rowA + bh / 2, aX[i + 1], rowA + bh / 2);
  });
  const bX = [2.75, 5.45, 8.15], bW = 2.25;
  [["crest · shape", "특징 추출"], ["MiniRocket / GBDT", "이력 길이로 분기"], ["Platt scaling", "확률 보정"]].forEach(([a, b], i) => {
    box(s, bX[i], rowB, bW, bh, a, b, { fill: C.accent4 });
    if (i < 2) arrow(s, bX[i] + bW, rowB + bh / 2, bX[i + 1], rowB + bh / 2);
  });
  const cX = 10.95, cW = 1.95;
  box(s, cX, rowA, cW, rowB + bh - rowA, "결합 판정", "OR rule\n5가지 상태", { fill: C.accent3, size: 16, subSize: 13 });
  // fan-out from input
  const sx = 2.4;
  arrow(s, inX + inW, midY, sx, midY, { noHead: true });
  arrow(s, sx, rowA + bh / 2, sx, rowB + bh / 2, { noHead: true });
  arrow(s, sx, rowA + bh / 2, aX[0], rowA + bh / 2);
  arrow(s, sx, rowB + bh / 2, bX[0], rowB + bh / 2);
  // fan-in
  arrow(s, aX[3] + aW, rowA + bh / 2, cX, rowA + bh / 2);
  arrow(s, bX[2] + bW, rowB + bh / 2, cX, rowB + bh / 2);
  // feedback loop
  const fy = 5.15;
  path2(s, [[cX + cW / 2, rowB + bh], [cX + cW / 2, fy], [inX + inW / 2, fy], [inX + inW / 2, midY + 0.55]], { dash: "dash", color: C.accent1 });
  T(s, "작업자 확인 → 교정 루프", { x: 4.5, y: fy + 0.07, w: 4.3, h: 0.3, fontSize: 12, align: "center", color: C.text2 });
  T(s, [
    ...D("A 부하 조건부 건강 모델", "정상 기록만 학습 · 경보 예산 α = 0.01 · 5가지 상태 출력", 14, false),
    ...D("B 전류 파형 판정기", "라벨 기반 파형 모델 · 확률 임계값 0.40", 14, false),
    ...D("결합 경보", "A 경보 OR B 경보 · 기록 1행(0.1초)마다 판정", 14),
  ], { x: 0.5, y: 5.7, w: 12.4, h: 1.2, paraSpaceAfter: 6 });
}
{
  const s = content("전문가 A ① 확률 world model", "부하가 진동 수준을 정하고, 건강 상태는 그 위에 얹힌 느린 편차", 7);
  // graphical model
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
  T(s, "전이", { x: 1.6, y: 1.75, w: 1.4, h: 0.3, fontSize: 12, align: "center", color: C.text2 });
  T(s, "OU 감쇠", { x: 1.6, y: 5.45, w: 1.4, h: 0.3, fontSize: 12, align: "center", color: C.text2 });
  // legend
  node(s, 0.62, 6.2, 0.26, "", { lw: 1 });
  T(s, "숨은 상태 — g : 부하 구간(4수준), q : 운전 맥락(2상태), h : 건강 편차", { x: 0.9, y: 6.05, w: 5.6, h: 0.3, fontSize: 12, valign: "middle" });
  node(s, 0.62, 6.6, 0.26, "", { lw: 1, fill: C.background2 });
  T(s, "관측 — x : 전류 파형, y : 진동 로그 RMS", { x: 0.9, y: 6.45, w: 5.6, h: 0.3, fontSize: 12, valign: "middle" });
  // equations
  const ex = 6.9, ew = 5.9;
  img(s, "image10", ex, 1.95, { w: 5.2, alt: "관측 분해 식" });
  T(s, "진동 = 같은 부하 · 맥락의 정상 기대치 μ + 건강 편차 h + 버스트 간 변동 u + 잡음 v", { x: ex, y: 2.55, w: ew, h: 0.6, fontSize: 14 });
  img(s, "image11", ex, 3.3, { w: 5.6, alt: "건강 상태의 OU process 식" });
  T(s, [
    { text: "건강 편차는 공백 동안 0으로 감쇠 (τ" },
    { text: "h", options: { subscript: true } },
    { text: " = 300 s), 버스트마다 Kalman update 1회" },
  ], { x: ex, y: 4.18, w: ew, h: 0.6, fontSize: 14 });
  T(s, "93% · 58%", { x: ex, y: 4.95, w: ew, h: 0.6, fontSize: 32, bold: true, valign: "middle" });
  T(s, "부하 구간이 설명하는 버스트 간 진동 변동 (하부 · 상부)", { x: ex, y: 5.6, w: ew, h: 0.35, fontSize: 14 });
  T(s, D("결론", "부하로 설명되지 않는 진동 상승만 건강 편차 h로 남음 · 신경망 없이 약 320개 값, 행당 약 0.2 ms", 14), { x: ex, y: 6.1, w: ew, h: 0.65 });
}
{
  const s = content("전문가 A ② 부하 상태 추정", "전류 진폭이 부하 구간을 결정 — 10행이면 부하 구간 확정, 그 전 행은 판단 보류", 8);
  T(s, D("① 전류 방출 모델", "진폭이 부하 구간을 결정", 16), { x: 0.5, y: 1.8, w: 7.6, h: 0.4 });
  img(s, "image12", 0.7, 2.25, { w: 6.9, alt: "전류 방출 모델 식" });
  T(s, D("② phase-marginalized likelihood", "위상 72점 격자 평균", 16), { x: 0.5, y: 3.25, w: 7.6, h: 0.4 });
  img(s, "image13", 0.7, 3.72, { w: 6.4, alt: "phase-marginalized likelihood 식" });
  T(s, [
    ...D("③ 부하 구간 4개", "BIC로 1–6 중 선택, 진폭 119 / 146 / 202 / 245", 16, false),
    ...D("④ 공백 인지 전이", "직전 구간에서 prior 생성, 18초 넘는 공백은 연결 끊음", 16),
  ], { x: 0.5, y: 5.05, w: 7.8, h: 1.0, paraSpaceAfter: 8 });
  bar(s, ["3행", "8행", "10행"], [{ name: "오분류율", values: [42.0, 8.4, 3.0] }], {
    x: 8.6, y: 1.75, w: 4.3, h: 3.9, colors: [HEX.accent1], title: "버스트 길이별 오분류율 (%)", max: 50, name: "오분류율 차트",
  });
  T(s, [
    { text: "readiness rule", options: { bold: true, fontSize: 16, breakLine: true } },
    { text: "확신 ≥ 0.9 또는 10행이면 확정, 그 전에는 판단 보류가 흡수", options: { fontSize: 14 } },
  ], { x: 8.75, y: 5.8, w: 4.15, h: 0.95, paraSpaceAfter: 4 });
}
{
  const s = content("전문가 A ③ 보정 · 경보 예산 · 판정 규칙", "경보 예산 α = 0.01 하나로 세 경로를 묶고, 부하 변화는 '운전변화 알림'으로 분리", 9);
  // budget bar
  const bx = 0.5, bw = 3.9, by = 2.3, bh = 0.4;
  T(s, "경보 예산 α = 0.01", { x: bx, y: 1.8, w: bw, h: 0.4, fontSize: 16, bold: true });
  const parts = [["건강 모델", "0.8 α", 0.8, C.accent1], ["sentinel", "0.1 α", 0.1, C.accent3], ["데이터 품질", "0.1 α", 0.1, C.accent4]];
  let cx = bx;
  parts.forEach(([n, v, f, c]) => {
    s.addShape(S.RECTANGLE, { x: cx, y: by, w: bw * f, h: bh, fill: { color: c }, line: { color: C.background1, width: 1.5 }, objectName: "budget " + n });
    cx += bw * f;
  });
  parts.forEach(([n, v, f, c], i) => {
    s.addShape(S.RECTANGLE, { x: bx, y: 2.95 + i * 0.38 + 0.07, w: 0.2, h: 0.2, fill: { color: c }, line: { color: c, width: 0.5 } });
    T(s, [{ text: n + "  ", options: { bold: true } }, { text: v }], { x: bx + 0.32, y: 2.95 + i * 0.38, w: 3.4, h: 0.34, fontSize: 14, valign: "middle" });
  });
  img(s, "image14", bx, 4.25, { w: 3.9, alt: "conformal p-value와 경보 예산 분배 식" });
  T(s, "행 위치 구간별 split-conformal p-value", { x: bx, y: 5.4, w: 3.9, h: 0.3, fontSize: 12, color: C.text2 });
  // decision order
  const rx = 4.95, rw = 3.75;
  T(s, "판정 순서 (위에서부터)", { x: rx, y: 1.8, w: rw, h: 0.4, fontSize: 16, bold: true });
  const rules = [
    ["1  센서 품질 실패", "→ 센서 점검"],
    ["2  새 운전 상태", "→ 물리 일치 검사"],
    ["3  readiness 전", "→ 판단 보류 포함"],
    ["4  readiness 후", "→ 예산 초과 시 경보 후보"],
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
    { text: "30초 안의 경보는 하나의 alarm episode로 묶음", options: { fontSize: 14 } },
  ], { x: rx, y: 5.5, w: rw + 0.3, h: 0.9, paraSpaceAfter: 4 });
  // five outputs
  const ox = 9.35;
  T(s, "출력 5가지 상태", { x: ox, y: 1.8, w: 3.5, h: 0.4, fontSize: 16, bold: true });
  const outs = [
    ["정상", C.background1, 0], ["운전변화 알림", C.accent2, 0], ["센서 점검", C.accent4, 0], ["경보", C.accent3, 0], ["판단 보류", C.background2, 0],
  ];
  outs.forEach(([n, c, tr], i) => {
    const y = 2.55 + i * 0.62;
    node(s, ox + 0.27, y, 0.5, "", { fill: c, transp: tr, lw: 1.25 });
    T(s, n, { x: ox + 0.7, y: y - 0.2, w: 2.9, h: 0.4, fontSize: 16, valign: "middle" });
  });
  T(s, D("정상 보정 구간 플래그율", "평균 0.42% (상한 약 1.02% 이하)", 14), { x: 0.5, y: 6.45, w: 12.4, h: 0.4 });
}
{
  const s = content("전문가 B ① 전류 파형 판정기", "이력 길이 16관측을 기준으로 두 경로 — 파형 모양으로 진폭이 정상인 이상도 포착", 10);
  const bh = 0.8, mainY = 2.55, topY = 1.85, botY = 3.25;
  const X = { hist: 0.5, len: 2.55, br: 4.65, platt: 7.25, pers: 9.3, alarm: 11.35 };
  box(s, X.hist, mainY, 1.6, bh, "구간 이력", "최근 기록");
  box(s, X.len, mainY, 1.6, bh, "길이 분기", "16관측 기준");
  box(s, X.br, topY, 2.1, bh, "MiniRocket", "16관측 이상", { fill: C.accent4 });
  box(s, X.br, botY, 2.1, bh, "GBDT", "16관측 미만 · 33개 특징", { fill: C.accent4 });
  box(s, X.platt, mainY, 1.6, bh, "Platt scaling", "확률 보정");
  box(s, X.pers, mainY, 1.6, bh, "3회 중 2회", "지속 규칙");
  box(s, X.alarm, mainY, 1.4, bh, "경보", null, { fill: C.accent3, size: 16 });
  const my = mainY + bh / 2;
  arrow(s, X.hist + 1.6, my, X.len, my);
  // split
  const sx = 4.35;
  arrow(s, X.len + 1.6, my, sx, my, { noHead: true });
  arrow(s, sx, topY + bh / 2, sx, botY + bh / 2, { noHead: true });
  arrow(s, sx, topY + bh / 2, X.br, topY + bh / 2);
  arrow(s, sx, botY + bh / 2, X.br, botY + bh / 2);
  // merge
  const mx = 7.0;
  arrow(s, X.br + 2.1, topY + bh / 2, mx, topY + bh / 2, { noHead: true });
  arrow(s, X.br + 2.1, botY + bh / 2, mx, botY + bh / 2, { noHead: true });
  arrow(s, mx, topY + bh / 2, mx, botY + bh / 2, { noHead: true });
  arrow(s, mx, my, X.platt, my);
  arrow(s, X.platt + 1.6, my, X.pers, my);
  arrow(s, X.pers + 1.6, my, X.alarm, my);
  // crest factor
  img(s, "image16", 0.5, 4.5, { w: 6.0, alt: "crest factor 식" });
  T(s, [
    { text: "정상 전류는 정현파 → crest ≈ √2 ≈ 1.414", options: { breakLine: true } },
    { text: "정상 1.32–1.54,  이상 1.5–3.5" },
  ], { x: 0.5, y: 5.5, w: 6.4, h: 0.65, fontSize: 14, paraSpaceAfter: 2 });
  T(s, [
    ...D("특징", "요약 통계 27 + crest factor 6 = 33개 (GBDT)", 14, false),
    { text: "MiniRocket 경로 9,996개", options: { fontSize: 14 } },
  ], { x: 0.5, y: 6.25, w: 6.4, h: 0.65, paraSpaceAfter: 2 });
  // AP progression
  T(s, "개발 시간순 검증 AP", { x: 7.4, y: 4.55, w: 5.5, h: 0.4, fontSize: 16, bold: true });
  const steps = [["요약 통계", "0.950"], ["+ crest factor", "0.980"], ["+ 파형 경로", "0.992"]];
  steps.forEach(([n, v], i) => {
    const x = 7.4 + i * 1.9;
    T(s, v, { x, y: 5.05, w: 1.6, h: 0.6, fontSize: 28, bold: true, valign: "middle" });
    T(s, n, { x, y: 5.7, w: 1.7, h: 0.35, fontSize: 14 });
    if (i < 2) arrow(s, x + 1.25, 5.35, x + 1.8, 5.35, { color: C.accent1, width: 1.5 });
  });
}
{
  const s = content("전문가 B ② 보정과 모델 선택", "확률은 Platt로 보정, 임계값 0.40은 세 원리가 일치 — 선정 규칙은 후보 확인 전에 문서로 고정", 11);
  T(s, D("확률 보정", "Platt (CV Brier 최저)", 16), { x: 0.5, y: 1.8, w: 5.9, h: 0.4 });
  T(s, "Brier : 보정 전 0.00434 · isotonic 0.00355 · Platt 0.00338", { x: 0.5, y: 2.25, w: 5.9, h: 0.35, fontSize: 14 });
  img(s, "image17", 0.5, 2.8, { w: 5.9, alt: "Platt scaling과 임계값 식" });
  T(s, "임계값 0.40, 세 원리가 일치", { x: 0.5, y: 3.95, w: 5.9, h: 0.4, fontSize: 16, bold: true });
  T(s, [
    ...D("비용비", "오경보 : 미탐 = 1.5", 14, false),
    ...D("F1 최적 정리", "이론값 0.469", 14, false),
    ...D("구간 conformal 보장 1%", "0.354", 14),
  ], { x: 0.5, y: 4.4, w: 5.9, h: 1.2, paraSpaceAfter: 6 });
  T(s, "후보 확인 구간 : 7계열 20개 구성에서 선정", { x: 6.9, y: 1.8, w: 6.0, h: 0.4, fontSize: 16, bold: true });
  table(s, ["후보", "AP", "경보/100구간"], [
    ["하이브리드 (crest factor 제외)", "0.859", "13.5"],
    ["정상 전용 신경망", "0.597", "30.2"],
    ["하이브리드 (crest factor 포함)", "0.909", "9.4"],
    ["기준 GBDT", "0.711", "13.5"],
  ], { x: 6.9, y: 2.3, w: 6.0, colW: [3.3, 0.9, 1.8], align: ["left", "right", "right"], boldRow: 2 });
  T(s, D("선정 규칙", "후보 ≤ 3개, 경보 ≤ 기준 모델의 1.5배 (100구간당 20.3건) 중 AP 최고", 14), { x: 6.9, y: 4.4, w: 6.0, h: 0.7 });
  T(s, [
    { text: "최종 구간 (1회 평가) : ", options: { bold: true, fontSize: 18 } },
    { text: "F1 0.841 · 정밀도 1.000 · 재현율 0.726 · 오경보 0 / 4,090행", options: { fontSize: 18 } },
  ], { x: 0.5, y: 6.1, w: 12.4, h: 0.5, valign: "middle" });
}
{
  const s = content("두 전문가 결합 : OR 규칙과 상보성", "재현율 전문가 A + 정밀도 전문가 B → 이상 구간 5/5, 재현율 0.993", 13);
  img(s, "image19", 0.5, 1.8, { w: 7.0, alt: "OR rule 결합 식" });
  T(s, "계산 전에 선언한 결합 규칙", { x: 7.8, y: 1.85, w: 4, h: 0.45, fontSize: 12, color: C.text2, valign: "middle" });
  table(s, ["시스템", "F1", "정밀도", "재현율", "이상 구간", "정상 오경보"], [
    ["전문가 A", "0.911", "0.847", "0.985", "5/5", "24행"],
    ["전문가 B", "0.841", "1.000", "0.726", "3/5", "0행"],
    ["A+B 결합", "0.915", "0.848", "0.993", "5/5", "24행"],
  ], { x: 0.5, y: 2.65, w: 6.6, colW: [1.35, 0.95, 1.0, 1.0, 1.1, 1.2], align: ["left", "right", "right", "right", "right", "right"], boldRow: 2 });
  T(s, [
    ...D("A", "B가 놓친 진폭 정상 구간 2개를 포착 (재현율 담당)", 14, false),
    ...D("B", "정밀도 1.000, 오경보 0행 (정밀도 담당)", 14, false),
    { text: "A의 오경보 24행은 정상 한 구간에 집중, B는 그 구간에서 경보 없음", options: { fontSize: 14, breakLine: true } },
    ...D("운영 중 개선", "B만 재학습 → F1 0.841 → 0.911, 구간 3 → 5", 14),
  ], { x: 0.5, y: 4.4, w: 6.6, h: 2.2, paraSpaceAfter: 8 });
  bar(s, ["둘 다 경보", "A만 경보", "B만 경보", "둘 다 놓침"], [{ name: "행 수", values: [97, 36, 1, 1] }], {
    x: 7.6, y: 2.55, w: 5.3, h: 4.0, colors: [HEX.accent1], title: "최종 구간 이상 135행의 판정 분담 (행)", fmt: "0", max: 110, name: "판정 분담 차트",
  });
}

// ======================================================================
// 15-19. Results
// ======================================================================
section = "실험 결과";
pres.addSection({ title: section });
{
  const s = content("결과 ① 부하 조건화의 효과", "같은 부하의 정상과 비교하면 부하 변화가 오경보로 이어지지 않는다", 12);
  img(s, "image18", 0.4, 1.75, { h: 4.75, alt: "부하 구간별 하부 진동과 고정 기준 오경보율" });
  const x = 8.75, w = 4.15;
  T(s, "92.6%", { x, y: 1.85, w, h: 0.6, fontSize: 32, bold: true, valign: "middle" });
  T(s, "부하 구간이 설명하는 하부 진동 변동", { x, y: 2.48, w, h: 0.35, fontSize: 14 });
  T(s, "22.3% → 3.7–5.7%", { x, y: 3.3, w, h: 0.6, fontSize: 28, bold: true, valign: "middle" });
  T(s, "정상 구간 오경보 (고정 기준 → 부하 조건)", { x, y: 3.93, w, h: 0.35, fontSize: 14 });
  T(s, "55.5% → 8.6%", { x, y: 4.75, w, h: 0.6, fontSize: 28, bold: true, valign: "middle" });
  T(s, "처음 보는 정상 운전상태 오표시율 (버스트 기준)", { x, y: 5.38, w, h: 0.6, fontSize: 14 });
  T(s, "기준 : 재사용 검증 구간 교차검증 · 10개 방법 동일 조건 비교", { x: 0.4, y: 6.65, w: 12.5, h: 0.3, fontSize: 12, color: C.text2 });
}
{
  const s = content("결과 ② 공식 라벨 시간순 최종 평가", "최종 구간 1회 평가, point-adjust 없음 — 정상 4,090행 · 이상 135행 · 이상 구간 5개", 14);
  img(s, "image20", 0.9, 1.7, { w: 11.5, alt: "공식 라벨 시간순 최종 평가: 모델별 F1·정밀도·재현율과 구간별 검출" });
  T(s, [
    { text: "A가 B의 미탐 36행, B가 A의 미탐 1행을 보완 → 결합 F1 0.915, 재현율 0.993", options: { bold: true, fontSize: 16, breakLine: true } },
    { text: "각 구간 첫 관측부터 첫 경보까지 지연 최대 0.1초, 중앙값 0초 · B에도 지속 규칙 적용 시 결합 F1 0.911", options: { fontSize: 14 } },
  ], { x: 0.9, y: 6.5, w: 12.0, h: 0.75, paraSpaceAfter: 4 });
}
{
  const s = content("결과 ③ 탐지와 경보 부담 : 10개 방법 비교", "하나의 경보 예산 아래 경보 부담은 최소, 약한 고장 탐지는 상위권", 15);
  img(s, "image21", 0.4, 1.7, { h: 4.75, alt: "10개 방법의 경보 부담과 순탐지 비교" });
  const x = 8.75, w = 4.15;
  [
    ["경보 부담 최소", "정상 스트림 alarm episode 11.7건/시간\n10개 방법 중 최소"],
    ["약한 고장 탐지 상위권", "semi-synthetic net detection 0.256"],
    ["5가지 상태 출력", "정상 · 운전변화 알림 · 센서 점검 · 경보 · 판단 보류"],
    ["정상 오경보 1%대", "전체 1.14% · 가동 구간 1.35%"],
  ].forEach(([h, b], i) => {
    T(s, [{ text: h, options: { bold: true, fontSize: 16, breakLine: true } }, { text: b, options: { fontSize: 14 } }], { x, y: 1.8 + i * 1.15, w, h: 1.0, paraSpaceAfter: 3 });
  });
  T(s, "비교 : LightGBM, Isolation Forest, 오토인코더, 마할라노비스 거리, 연속 공변량 회귀, Chronos-2(zero-shot), 상태조건 신경망 등 · 새 운전상태의 운전변화 알림 80.9–89.4%는 경보와 별도 집계", { x: 0.4, y: 6.55, w: 12.5, h: 0.55, fontSize: 12, color: C.text2 });
}
{
  const s = content("결과 ④ 오경보 · 미탐 조건 분석", "오경보는 구간 시작에, 미탐은 짧은 이력에 집중 → 조건마다 설계로 대응", 16);
  img(s, "image22", 1.45, 1.65, { w: 10.4, alt: "오경보는 구간 시작에, 미탐은 짧은 이력에 집중" });
  T(s, [
    ...D("구간 시작 오경보", "판단 보류 + 3회 중 2회 지속 규칙 → 즉시 경보 60.8건/h → 0 (B 개발 분할)", 14, false),
    ...D("부하 오경보", "같은 부하 비교 → 고정 기준 22.3%, 처음 보는 최고 부하 52.1% → 1.82%", 14, false),
    ...D("짧은 이력 미탐", "16관측 미만 구간은 A가 5개 모두 포착 · 진동 단독 상승 = 기계 변화 후보", 14),
  ], { x: 0.5, y: 5.95, w: 12.4, h: 1.15, paraSpaceAfter: 3 });
}
{
  const s = content("결과 ⑤ 외부 장비 전이 : 외부 유압 시험대", "부하 조건화가 처음 보는 운전 조건의 정상 오경보를 19.0% → 10.4%로 줄임", 22);
  img(s, "image28", 1.65, 1.62, { w: 10.0, alt: "외부 유압 시험대에서의 오경보율 비교와 전이 결과" });
  T(s, [
    ...D("실험", "UCI 유압 시험대 (CC BY 4.0) · 외부→대회, 대회→외부, 공동 학습 · 비교 Isolation Forest 34.7%, 고정 임계값 65.9%", 13, false),
    ...D("외부 데이터 활용", "외부 모터 전류 · 엔진 진동 (파형 모양 설계 확인) · MetroPT-3 (첫 경보 시점 방법 시연)", 13, false),
    ...D("현장 적용", "다른 장비 모델을 그대로 쓰면 오경보 ≈ 100% → 현장 정상 데이터로 기준 재보정 필수", 13),
  ], { x: 0.5, y: 6.05, w: 12.4, h: 1.1, paraSpaceAfter: 3 });
}

// ======================================================================
// 20-24. Type 2 — correction loop & grounded explanation
// ======================================================================
section = "유형 2";
pres.addSection({ title: section });
{
  const s = content("교정 루프 : canary test로 교정 후보 선별", "운영 경험을 반영하되, 고장 증거를 지우는 교정은 통과시키지 않는다", 17);
  const cx = 4.0, cy = 4.25, R = 1.8, d = 0.62;
  const labels = [
    ["운영 · 판정 · 설명", "right"],
    ["작업자 피드백", "right"],
    ["교정 후보 갱신", "right"],
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
  // rejection branch from canary
  const cp = pts[3];
  arrow(s, cp[0] + d / 2, cp[1], cp[0] + 2.4, cp[1], { color: C.accent6, dash: "dash", width: 1.5 });
  T(s, "고장 증거를 지우면 기각", { x: cp[0] + 2.5, y: cp[1] - 0.18, w: 2.3, h: 0.36, fontSize: 13, valign: "middle" });
  const x = 8.4, w = 4.5;
  T(s, "92.3% → 64.6%", { x, y: 1.9, w, h: 0.6, fontSize: 28, bold: true, valign: "middle" });
  T(s, "semi-synthetic 시험 : 정상 알림 비율", { x, y: 2.52, w, h: 0.35, fontSize: 14 });
  T(s, "11 / 16", { x, y: 3.25, w, h: 0.6, fontSize: 28, bold: true, valign: "middle" });
  T(s, "고장이 섞인 확인 조건을 기각", { x, y: 3.87, w, h: 0.35, fontSize: 14 });
  T(s, [
    ...D("canary test", "약한 고장을 섞은 시험을 통과한 후보만 반영", 14, false),
    ...D("결과", "운영 경험이 고장 증거를 지우지 않는 범위에서 공정 기준으로 축적", 14),
  ], { x, y: 4.65, w, h: 1.6, paraSpaceAfter: 8 });
}
{
  const s = content("유형 2 : 근거 기반 설명과 실제 답변", "등급은 규칙표가 정하고, AI는 근거를 설명한다 — 어디서 · 무엇이 · 왜 · 무엇을 점검", 18);
  const bw = 1.85, bh = 0.85, gap = 0.35, x0 = 0.5, r1 = 1.85, r2 = 3.15;
  const xs = [x0, x0 + bw + gap, x0 + 2 * (bw + gap)];
  box(s, xs[0], r1, bw, bh, "계기", "변화 · 요청 · 질문");
  box(s, xs[1], r1, bw, bh, "context", "assembly");
  box(s, xs[2], r1, bw, bh, "지식그래프", "2-hop 검색");
  box(s, xs[2], r2, bw, bh, "답변 생성", "언어모델", { fill: C.accent4 });
  box(s, xs[1], r2, bw, bh, "validator", "17개 검사");
  box(s, xs[0], r2, bw, bh, "답변 카드", "작업자 화면", { fill: C.accent3 });
  arrow(s, xs[0] + bw, r1 + bh / 2, xs[1], r1 + bh / 2);
  arrow(s, xs[1] + bw, r1 + bh / 2, xs[2], r1 + bh / 2);
  arrow(s, xs[2] + bw / 2, r1 + bh, xs[2] + bw / 2, r2);
  arrow(s, xs[2], r2 + bh / 2, xs[1] + bw, r2 + bh / 2);
  arrow(s, xs[1], r2 + bh / 2, xs[0] + bw, r2 + bh / 2);
  const ry = 4.45;
  box(s, x0, ry, 3 * bw + 2 * gap, 0.7, "호출 규칙표", "고정 규칙 13개가 점검 등급 결정 — AI가 바꿀 수 없음");
  arrow(s, xs[0] + bw / 2, ry, xs[0] + bw / 2, r2 + bh, { dash: "dash", color: C.accent6, width: 1.5 });
  T(s, [
    ...D("질문", "“방금 알림은 왜 났어?” (화면 121초 기준)", 14, false),
    ...D("판정 상태", "운전변화 알림 → 등급 : 관찰 강화 (규칙표 결정)", 14, false),
    ...D("지식 근거", "6건 · 그래프 노드 205 · 연결 320 · 출처 76", 14),
  ], { x: 0.5, y: 5.45, w: 6.6, h: 1.3, paraSpaceAfter: 5 });
  const im = img(s, "image25", 7.45, 1.8, { w: 5.45, alt: "실제 답변 화면: 운전변화 알림, 화면 시점 121초, 관찰 강화 등급" });
  s.addShape(S.RECTANGLE, { x: 7.45, y: 1.8, w: im.w, h: im.h, fill: { type: "none" }, line: { color: C.background2, width: 0.75 } });
  T(s, "재생 자료의 실제 답변 화면 (심사 시 시연 가능)", { x: 7.45, y: 1.8 + im.h + 0.08, w: 5.45, h: 0.3, fontSize: 12, color: C.text2 });
  T(s, D("평가", "재생 12쌍 · 시뮬레이션 8쌍에서 citation compliance 11% → 100%", 14), { x: 7.45, y: 5.75, w: 5.45, h: 0.7 });
}
{
  const s = content("유형 2 ① 지식그래프 근거 검색", "임베딩 없이 deterministic 검색 — 같은 질문에는 항상 같은 근거", 19);
  // abstract 2-hop retrieval graph
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
  // legend
  const ly = 6.2;
  [["seed", C.accent3, 0], ["1-hop", C.accent4, 0], ["2-hop", C.accent2, 60], ["선택 안 됨", C.background1, 0]].forEach(([n, c, tr], i) => {
    node(s, 0.62 + i * 1.5, ly, 0.24, "", { fill: c, transp: tr, lw: 1 });
    T(s, n, { x: 0.82 + i * 1.5, y: ly - 0.16, w: 1.2, h: 0.32, fontSize: 12, valign: "middle" });
  });
  T(s, "hub : 이웃이 많은 노드 → hub damping으로 점수 감쇠", { x: 0.5, y: 6.55, w: 6.0, h: 0.3, fontSize: 12, color: C.text2 });
  // steps
  const x = 6.9, w = 6.0;
  [
    ["entity linking", "질문의 한 · 영 별칭 + 현재 판정 상태 · 경로 · 같은 부하 대비 신호에서 seed 생성"],
    ["2-hop 점수 전파", "홉마다 ×0.55, 관계 가중(0.7–1.2) · hub damping · 의도 배율"],
    ["상위 근거 선택", "노드 8 · 경로 5 · 출처 6 · 주제 요약 2를 모델에 전달"],
  ].forEach(([h, b], i) => {
    const y = 1.85 + i * 1.05;
    node(s, x + 0.22, y + 0.22, 0.44, String(i + 1), { size: 14, bold: true, lw: 1.25 });
    T(s, [{ text: h, options: { bold: true, fontSize: 16, breakLine: true } }, { text: b, options: { fontSize: 14 } }], { x: x + 0.65, y, w: w - 0.65, h: 0.95, paraSpaceAfter: 2 });
  });
  img(s, "image27", x, 5.05, { w: 5.6, alt: "2-hop 점수 전파 식" });
  T(s, [
    { text: "205 · 320 · 76", options: { bold: true, fontSize: 16 } },
    { text: "  노드 · 간선 · 출처 (노드 유형 11, 간선 유형 19, 모든 노드 · 간선에 출처와 근거 등급)", options: { fontSize: 13 } },
  ], { x, y: 5.85, w, h: 0.7 });
}
{
  const s = content("유형 2 ② tool calling LLM과 validator", "AI의 답은 17개 검사를 통과해야 화면에 나간다", 20);
  T(s, "tool calling LLM", { x: 0.5, y: 1.8, w: 5.4, h: 0.4, fontSize: 18, bold: true });
  T(s, [
    ...D("도구", "읽기 전용 조회 11개 + 답안 제출 1개", 14, false),
    ...D("첫 턴", "상태 · 같은 부하 비교 · 정책 · 지식그래프 · 공정 지식 · 점검 초안 6개 병렬 호출", 14, false),
    ...D("요청 수", "질문당 5–8회 → 2회 안팎 (최대 9턴, 기한 80초)", 14),
  ], { x: 0.5, y: 2.3, w: 5.4, h: 2.0, paraSpaceAfter: 8 });
  T(s, "검증 실패 시", { x: 0.5, y: 4.5, w: 5.4, h: 0.4, fontSize: 16, bold: true });
  const fx = [0.5, 2.35, 4.2], fw = 1.55, fy = 5.0, fh = 0.7;
  box(s, fx[0], fy, fw, fh, "실패", "위반 목록 반환", { fill: C.accent3 });
  box(s, fx[1], fy, fw, fh, "수리 1회", "다시 작성");
  box(s, fx[2], fy, fw, fh, "대체 답변", "deterministic");
  arrow(s, fx[0] + fw, fy + fh / 2, fx[1], fy + fh / 2);
  arrow(s, fx[1] + fw, fy + fh / 2, fx[2], fy + fh / 2);
  T(s, "통과 못 한 문장은 화면에 나가지 않음", { x: 0.5, y: 5.9, w: 5.4, h: 0.4, fontSize: 14, bold: true });
  T(s, "validator 17개 검사 (대표 6가지)", { x: 6.5, y: 1.8, w: 6.4, h: 0.4, fontSize: 18, bold: true });
  table(s, ["검사", "강제하는 것"], [
    ["숫자 · 단위", "도구 값만, 단위 없음"],
    ["단정 · 수명", "원인 확정 · 남은 수명 금지"],
    ["제어 지시", "정지 · 설정 변경 권고 금지"],
    ["근거 · 신선도", "가설마다 근거, 120초 이내"],
    ["지시문 방어", "도구 결과 속 지시문 무시 (패턴 13개)"],
    ["등급 일치", "답의 등급 = 정책표 등급"],
  ], { x: 6.5, y: 2.35, w: 6.4, colW: [2.0, 4.4], rowH: 0.48 });
}
{
  const s = content("유형 2 ③ deterministic escalation과 평가", "호출 등급은 모델이 아니라 13개 규칙표가 정한다 — 모델은 등급을 바꿀 수 없음", 21);
  T(s, "규칙 13개 → 등급 4개 (위에서부터 처음 맞는 규칙)", { x: 0.5, y: 1.8, w: 7, h: 0.4, fontSize: 16, bold: true });
  table(s, ["등급", "대표 조건", "재확인"], [
    ["즉시 현장 확인 후보", "경보 + 높음 지속", "120초"],
    ["담당자 호출", "센서 · 수집 의심, 운전 변화 겹침", "300초"],
    ["관찰 강화", "판단 보류, 운전변화 알림", "600초"],
    ["기록 유지", "정상", "1,800초"],
  ], { x: 0.5, y: 2.3, w: 6.9, colW: [2.2, 3.5, 1.2], align: ["left", "left", "right"], rowH: 0.48 });
  T(s, [
    { text: "경보는 어떤 입력에서도 '담당자 호출' 아래로 내려가지 않음", options: { fontSize: 14, breakLine: true } },
    { text: "입력 조합 10,206칸 전수 검사 → 일관성 위반 0", options: { fontSize: 16, bold: true } },
  ], { x: 0.5, y: 4.95, w: 6.9, h: 0.9, paraSpaceAfter: 6 });
  bar(s, ["근거 표기 (citation)", "정책 등급 일치"], [
    { name: "도입 전", values: [11, 60] },
    { name: "도입 후 (지식그래프 + 정책표)", values: [100, 100] },
  ], { x: 7.8, y: 1.75, w: 5.1, h: 4.2, colors: ["A6A6A6", HEX.accent1], title: "질문 20쌍 평가 결과 (%)", fmt: "0", max: 115, gap: 60, name: "도입 전후 차트" });
  T(s, "질문 20쌍(실제 재생 12 + 시뮬레이션 8) : 7개 기준 충족 99% → 98% (동등) · 평균 지연 28.9초 → 30.7초 · 설비 제어 지시 0건", { x: 0.5, y: 6.35, w: 12.4, h: 0.4, fontSize: 14 });
}

// ======================================================================
// 25-28. Field deployment & conclusion
// ======================================================================
section = "현장 적용 및 결론";
pres.addSection({ title: section });
{
  const s = content("현장 적용 : 3단계 대응과 단계적 도입", "읽기 전용 병행 운영으로 시작해, 현장 시험 기준을 통과하면 운영으로 전환", 23);
  T(s, "현장 대응 : 3단계 + 별도 데이터 점검", { x: 0.5, y: 1.8, w: 5.4, h: 0.4, fontSize: 16, bold: true });
  [
    ["정상", "특이 없음", C.background1, 0],
    ["관찰", "← 운전변화 알림 · 판단 보류", C.accent2, 0],
    ["점검 권고", "← 경보", C.accent3, 0],
    ["데이터 점검", "← 센서 점검 알림 (별도 경로)", C.accent4, 0],
  ].forEach(([n, b, c, tr], i) => {
    const y = 2.6 + i * 0.85;
    node(s, 0.8, y, 0.55, "", { fill: c, transp: tr, lw: 1.25 });
    T(s, n, { x: 1.25, y: y - 0.2, w: 1.5, h: 0.4, fontSize: 16, bold: true, valign: "middle" });
    T(s, b, { x: 2.75, y: y - 0.2, w: 3.4, h: 0.4, fontSize: 14, valign: "middle" });
  });
  T(s, "설비로 쓰는 경로 없음 · 13개 호출 규칙은 10,206개 입력 조합에서 위반 0", { x: 0.5, y: 6.0, w: 5.6, h: 0.7, fontSize: 14 });
  const x = 6.9;
  T(s, "단계적 도입", { x, y: 1.8, w: 5, h: 0.4, fontSize: 16, bold: true });
  const steps = [
    ["설치 점검", "진동 채널 매핑 확인 · 전류 0 확인"],
    ["읽기 전용 병행 운영", "현장 정상 기록으로 기준 설정"],
    ["모델 · 임계값 동결", "동결한 구성으로 시험 준비"],
    ["다른 날짜 현장 시험", "정상 오경보 95% 신뢰구간 상한 ≤ 1%, ≤ 6건/시간"],
    ["운영 전환", "기준 통과 후 전환"],
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
  s.addText("심사 시 현장 시연 — 실제 시뮬레이션 화면", { placeholder: "desc", bullet: false });
  img(s, "image34", 4.92, 1.9, { w: 3.5, alt: "시연 페이지 QR 코드" });
  T(s, "QR 코드로 시뮬레이션 화면 접속", { x: 4.0, y: 5.6, w: 5.33, h: 0.4, fontSize: 16, align: "center" });
  s.addNotes("심사 시 현장 시연이 가능합니다. QR 코드로 실제 시뮬레이션 화면에 접속할 수 있습니다.");
}
{
  const s = content("결론 및 향후 과제", "작업자 경험에 좌우되던 판단에서, 공정과 사람이 대화하는 운영으로", 25);
  T(s, "기여", { x: 0.5, y: 1.8, w: 6, h: 0.4, fontSize: 16, bold: true });
  const fills = [[C.accent3, 0], [C.accent4, 0], [C.accent2, 0], [C.accent5, 0]];
  const items = [
    ["같은 부하 정상 대비 판정", "처음 보는 정상 운전상태 오표시 55.5% → 8.6%"],
    ["하나의 경보 예산, 5가지 상태", "경보 11.7건/시간, 10개 방법 중 최소"],
    ["고장 증거를 지키는 교정 루프", "정상 알림 92.3% → 64.6%, 고장 섞인 조건 11/16 기각"],
    ["근거 설명 + 호출 규칙표", "citation compliance 11% → 100%, 위반 0"],
  ];
  const d = 0.55, y0 = 2.55, dy = 0.95, x = 0.5;
  arrow(s, x + d / 2, y0, x + d / 2, y0 + 3 * dy, { noHead: true, color: C.accent1, width: 1.5 });
  items.forEach(([h, b], i) => {
    const y = y0 + i * dy;
    node(s, x + d / 2, y, d, String(i + 1), { fill: fills[i][0], transp: fills[i][1], bold: true, size: 14 });
    T(s, [{ text: h, options: { bold: true, fontSize: 16, breakLine: true } }, { text: b, options: { fontSize: 13 } }], { x: x + 0.8, y: y - 0.3, w: 6.3, h: 0.75 });
  });
  const nx = 8.0, nw = 4.6;
  T(s, "향후 과제", { x: nx, y: 1.8, w: nw, h: 0.4, fontSize: 16, bold: true });
  const nexts = ["현장 정상 데이터 재보정", "현장 시험", "고장 사례 축적 · 지식그래프 확장"];
  nexts.forEach((n, i) => {
    const y = 2.35 + i * 1.0;
    box(s, nx, y, nw, 0.62, n, null, { size: 15 });
    if (i < 2) arrow(s, nx + nw / 2, y + 0.62, nx + nw / 2, y + 1.0);
  });
  T(s, [
    { text: "최종 결합(A+B) : ", options: { bold: true, fontSize: 16 } },
    { text: "F1 0.915 · 재현율 0.993 · 이상 구간 5/5 · 정상 오경보 24 / 4,090행", options: { fontSize: 16 } },
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
