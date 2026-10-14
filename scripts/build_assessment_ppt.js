const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require("/root/.claude/skills/synced/58492139-558f-4c31-bab3-9821f6393e30_b3cf6d08-3c01-4a4e-ad7a-826a6a0e3731/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";

// 藍(남색) + 朱(주홍) — 일본 전통색에서 착안한 팔레트
const THEME = {
  name: "Japanese Assessment",
  headFontFace: "맑은 고딕",
  bodyFontFace: "맑은 고딕",
  colors: {
    dk1: "1E1E24", lt1: "FFFFFF", dk2: "1F2D4F", lt2: "F3F5FA",
    accent1: "D9472B", accent2: "1F2D4F", accent3: "F2B544", accent4: "3A8F6B",
    accent5: "6B7A99", accent6: "FBE7E2", hlink: "D9472B", folHlink: "6B7A99",
  },
};
const JP = "Meiryo"; // 일본어 문장용

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.title = "일본어 수행평가 안내";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.5, w: 11.7, h: 1.5, fontSize: 54, bold: true, color: C.background1, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.1, w: 11.7, h: 1.0, fontSize: 30, color: C.accent3, align: "left", valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 12.1, h: 1.0, fontSize: 40, bold: true, color: C.text2, align: "left", valign: "middle", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 12.3, y: 6.95, w: 0.6, h: 0.35, fontSize: 14, color: C.accent5, align: "right" },
});

(async () => {
  const icUser = await icon(fa.FaUserAlt, "FFFFFF");
  const icSchool = await icon(fa.FaSchool, "FFFFFF");
  const icHeart = await icon(fa.FaHeart, "FFFFFF");
  const icBow = await icon(fa.FaHandsHelping, "FFFFFF");
  const icPen = await icon(fa.FaPenFancy, "FFFFFF");
  const icComment = await icon(fa.FaCommentDots, "FFFFFF");
  const icMic = await icon(fa.FaMicrophone, "FFFFFF");
  const icRedo = await icon(fa.FaRedoAlt, "FFFFFF");
  const icCheck = await icon(fa.FaCheck, "3A8F6B");
  const icWarn = await icon(fa.FaExclamation, "D9472B");
  const icStar = await icon(fa.FaStar, "F2B544");

  // 1. 표지
  pres.addSection({ title: "표지" });
  let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "표지" });
  s.addText("2학기 일본어 수행평가", { x: 0.8, y: 1.5, w: 8, h: 0.7, fontSize: 28, color: C.accent6, margin: 0, isTextBox: true, objectName: "Kicker" });
  s.addText("나를 소개해요!", { placeholder: "title" });
  s.addText([{ text: "じこしょうかい", options: { fontFace: JP } }, { text: "  말하기 평가" }], { placeholder: "body" });
  s.addShape(pres.shapes.OVAL, { x: 10.2, y: 0.9, w: 2.4, h: 2.4, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "Sun" });
  s.addText("2학년 일본어", { x: 0.8, y: 6.3, w: 6, h: 0.5, fontSize: 22, color: C.background2, margin: 0, isTextBox: true, objectName: "Footer" });
  s.addNotes("수행평가 전체 안내. 평가 일정과 기준은 학교 계획에 맞게 수정하세요.");

  // 2. 한눈에 보기
  pres.addSection({ title: "평가 개요" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "평가 개요" });
  s.addText("한눈에 보기", { placeholder: "title" });
  const stats = [
    ["20%", "성적 반영 비율"],
    ["1분", "발표 시간"],
    ["11월", "평가 시기"],
  ];
  stats.forEach(([big, label], i) => {
    const x = 0.6 + i * 4.2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.9, w: 3.9, h: 3.4, rectRadius: 0.2, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Stat card ${i + 1}` });
    s.addText(big, { x, y: 2.3, w: 3.9, h: 1.6, fontSize: 66, bold: true, color: C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `Stat value ${i + 1}` });
    s.addText(label, { x, y: 4.1, w: 3.9, h: 0.8, fontSize: 28, color: C.text2, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `Stat label ${i + 1}` });
  });
  s.addText("일본어로 1분 동안 자기소개를 하고, 원고 없이 발표해요", { x: 0.6, y: 5.75, w: 12.1, h: 0.8, fontSize: 28, bold: true, color: C.text1, align: "left", valign: "middle", margin: 0, isTextBox: true, objectName: "Summary" });

  // 3. 무엇을 말하나요
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "평가 개요" });
  s.addText("발표에 꼭 들어가야 할 4가지", { placeholder: "title" });
  const items = [
    [icUser, "이름", "はじめまして。〇〇です。"],
    [icSchool, "학년", "中学2年生です。"],
    [icHeart, "좋아하는 것", "〇〇が好きです。"],
    [icBow, "마무리 인사", "よろしくお願いします。"],
  ];
  items.forEach(([ic, head, jp], i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.6 + col * 6.2, y = 1.85 + row * 2.45;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 5.9, h: 2.15, rectRadius: 0.15, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Item card ${i + 1}` });
    s.addShape(pres.shapes.OVAL, { x: x + 0.35, y: y + 0.5, w: 1.15, h: 1.15, fill: { color: C.accent1 }, line: { type: "none" }, objectName: `Item circle ${i + 1}` });
    s.addImage({ data: ic, x: x + 0.6, y: y + 0.75, w: 0.65, h: 0.65, objectName: `Item icon ${i + 1}` });
    s.addText(head, { x: x + 1.8, y: y + 0.3, w: 3.9, h: 0.75, fontSize: 32, bold: true, color: C.text2, margin: 0, valign: "middle", isTextBox: true, objectName: `Item head ${i + 1}` });
    s.addText(jp, { x: x + 1.8, y: y + 1.1, w: 3.95, h: 0.75, fontSize: 22, fontFace: JP, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: `Item jp ${i + 1}` });
  });

  // 4. 진행 순서
  pres.addSection({ title: "진행 방법" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "진행 방법" });
  s.addText("이렇게 진행해요", { placeholder: "title" });
  const steps = [
    [icPen, "원고 쓰기", "수업 시간에\n일본어로 작성"],
    [icComment, "피드백", "선생님께\n원고 점검받기"],
    [icRedo, "연습하기", "외워서\n소리 내어 연습"],
    [icMic, "발표하기", "앞에 나와서\n1분 발표"],
  ];
  steps.forEach(([ic, head, desc], i) => {
    const x = 0.6 + i * 3.15;
    s.addShape(pres.shapes.OVAL, { x: x + 0.8, y: 2.2, w: 1.5, h: 1.5, fill: { color: i === 3 ? C.accent1 : C.text2 }, line: { type: "none" }, objectName: `Step circle ${i + 1}` });
    s.addImage({ data: ic, x: x + 1.15, y: 2.55, w: 0.8, h: 0.8, objectName: `Step icon ${i + 1}` });
    if (i < 3) s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: x + 2.75, y: 2.75, w: 0.4, h: 0.4, rotate: 90, fill: { color: C.accent5 }, line: { type: "none" }, objectName: `Arrow ${i + 1}` });
    s.addText(`${i + 1}단계`, { x, y: 3.95, w: 3.1, h: 0.5, fontSize: 22, bold: true, color: C.accent1, align: "center", margin: 0, isTextBox: true, objectName: `Step num ${i + 1}` });
    s.addText(head, { x, y: 4.45, w: 3.1, h: 0.75, fontSize: 32, bold: true, color: C.text2, align: "center", margin: 0, isTextBox: true, objectName: `Step head ${i + 1}` });
    s.addText(desc, { x, y: 5.3, w: 3.1, h: 1.3, fontSize: 24, color: C.text1, align: "center", valign: "top", margin: 0, isTextBox: true, objectName: `Step desc ${i + 1}` });
  });

  // 5. 평가 기준
  pres.addSection({ title: "평가 기준" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "평가 기준" });
  s.addText("평가 기준 (20점 만점)", { placeholder: "title" });
  const hdr = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, align: "center" } });
  const cell = (t, o = {}) => ({ text: t, options: { color: C.text1, ...o } });
  const rows = [
    [hdr("영역"), hdr("이런 점을 봐요"), hdr("배점")],
    [cell("내용", { bold: true, align: "center" }), cell("4가지 내용이 모두 들어 있나요?"), cell("8점", { bold: true, color: C.accent1, align: "center" })],
    [cell("발음 · 유창성", { bold: true, align: "center" }), cell("정확한 발음으로 자연스럽게 말하나요?"), cell("6점", { bold: true, color: C.accent1, align: "center" })],
    [cell("태도", { bold: true, align: "center" }), cell("원고 없이, 큰 목소리로 발표하나요?"), cell("4점", { bold: true, color: C.accent1, align: "center" })],
    [cell("시간", { bold: true, align: "center" }), cell("1분(±10초) 안에 발표를 마치나요?"), cell("2점", { bold: true, color: C.accent1, align: "center" })],
  ];
  s.addTable(rows, {
    x: 0.6, y: 1.8, w: 12.1, colW: [3.0, 7.1, 2.0], rowH: 0.95,
    fontSize: 26, valign: "middle", margin: [0, 0.2, 0, 0.2],
    border: { type: "solid", pt: 1, color: "D5DAE6" }, fill: { color: "FFFFFF" },
    objectName: "Rubric table",
  });

  // 6. 예시 원고
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "평가 기준" });
  s.addText("예시 원고", { placeholder: "title" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.75, w: 12.1, h: 4.95, rectRadius: 0.15, fill: { color: C.background2 }, line: { type: "none" }, objectName: "Script card" });
  const lines = [
    ["はじめまして。キム・ミンジュです。", "처음 뵙겠습니다. 김민주입니다."],
    ["中学2年生です。14さいです。", "중학교 2학년입니다. 14살입니다."],
    ["わたしは ねこが 好きです。", "저는 고양이를 좋아합니다."],
    ["どうぞ よろしく おねがいします。", "잘 부탁드립니다."],
  ];
  lines.forEach(([jp, ko], i) => {
    const y = 1.95 + i * 1.18;
    s.addText(jp, { x: 1.0, y, w: 11.3, h: 0.65, fontSize: 30, bold: true, fontFace: JP, color: C.text2, margin: 0, valign: "middle", isTextBox: true, objectName: `Script jp ${i + 1}` });
    s.addText(ko, { x: 1.0, y: y + 0.6, w: 11.3, h: 0.45, fontSize: 20, color: C.accent5, margin: 0, valign: "middle", isTextBox: true, objectName: `Script ko ${i + 1}` });
  });

  // 7. 꿀팁 & 유의사항
  pres.addSection({ title: "유의사항" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "유의사항" });
  s.addText("꼭 기억하세요", { placeholder: "title" });
  const colBox = (x, title, ic, color, bullets, tag) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.75, w: 5.9, h: 4.6, rectRadius: 0.15, fill: { color: C.background2 }, line: { type: "none" }, objectName: `${tag} card` });
    s.addShape(pres.shapes.OVAL, { x: x + 0.4, y: 2.05, w: 0.8, h: 0.8, fill: { color: C.background1 }, line: { type: "none" }, objectName: `${tag} circle` });
    s.addImage({ data: ic, x: x + 0.6, y: 2.25, w: 0.4, h: 0.4, objectName: `${tag} icon` });
    s.addText(title, { x: x + 1.4, y: 2.05, w: 4.2, h: 0.8, fontSize: 30, bold: true, color, margin: 0, valign: "middle", isTextBox: true, objectName: `${tag} title` });
    s.addText(bullets.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < bullets.length - 1 } })), {
      x: x + 0.4, y: 3.15, w: 5.2, h: 3.3, fontSize: 28, color: C.text1, valign: "top", paraSpaceAfter: 20, margin: 0, isTextBox: true, objectName: `${tag} list`,
    });
  };
  colBox(0.6, "이렇게 하면 좋아요", icCheck, C.accent4, [
    "천천히, 또박또박 말하기",
    "친구들 얼굴 보며 말하기",
    "집에서 거울 보고 연습하기",
  ], "Tips");
  colBox(6.8, "주의하세요", icWarn, C.accent1, [
    "원고를 보면 태도 점수 감점",
    "결석 시 추후 재평가",
    "번역기로 쓴 원고는 0점",
  ], "Warn");

  // 8. 마무리
  pres.addSection({ title: "마무리" });
  s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "마무리" });
  s.addText("がんばってください！", { placeholder: "title", fontFace: JP });
  s.addText("궁금한 점은 언제든지 선생님께 물어보세요", { placeholder: "body" });
  s.addImage({ data: icStar, x: 11.0, y: 1.0, w: 1.4, h: 1.4, objectName: "Star" });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
