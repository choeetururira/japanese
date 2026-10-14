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
pres.title = "일본어회화 2학기 수행평가 안내";
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
  const W = "FFFFFF";
  const ic = {};
  for (const [k, comp, col] of [
    ["place", fa.FaMapMarkerAlt, W], ["food", fa.FaUtensils, W], ["see", fa.FaCamera, W],
    ["shop", fa.FaShoppingBag, W], ["free", fa.FaStar, W],
    ["group", fa.FaUsers, W], ["pen", fa.FaPenFancy, W], ["mic", fa.FaMicrophone, W], ["redo", fa.FaRedoAlt, W],
    ["polite", fa.FaUserTie, W], ["ink", fa.FaPen, W], ["kanji", fa.FaLanguage, W], ["mean", fa.FaBookOpen, W],
    ["ban", fa.FaExclamation, W], ["arrow", fa.FaArrowRight, "D9472B"], ["plane", fa.FaPlane, "F2B544"],
  ]) ic[k] = await icon(comp, col);

  const txt = (sl, t, o) => sl.addText(t, { margin: 0, isTextBox: true, ...o });
  const card = (sl, x, y, w, h, name, color = C.background2) =>
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.15, fill: { color }, line: { type: "none" }, objectName: name });
  const circleIcon = (sl, x, y, d, data, name, color = C.accent1) => {
    sl.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color }, line: { type: "none" }, objectName: `${name} circle` });
    sl.addImage({ data, x: x + d * 0.25, y: y + d * 0.25, w: d * 0.5, h: d * 0.5, objectName: `${name} icon` });
  };

  // 1. 표지
  pres.addSection({ title: "표지" });
  let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "표지" });
  txt(s, "송호고등학교 2학년 일본어회화 · 2학기 수행평가", { x: 0.8, y: 1.5, w: 9.2, h: 0.7, fontSize: 26, color: C.accent6, objectName: "Kicker" });
  s.addText("일본 여행 버킷리스트", { placeholder: "title" });
  s.addText([{ text: "〜たいです", options: { fontFace: JP, bold: true } }, { text: "  로 쓰고, 말하기" }], { placeholder: "body" });
  s.addShape(pres.shapes.OVAL, { x: 10.2, y: 0.9, w: 2.4, h: 2.4, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "Sun" });
  s.addImage({ data: ic.plane, x: 10.8, y: 1.5, w: 1.2, h: 1.2, objectName: "Plane" });
  s.addNotes("2학기 수행평가: 희망표현 〜たいです를 활용한 일본 여행 버킷리스트 작성(7점) + 인터뷰 말하기(5점).");

  // 2. 한눈에 보기
  pres.addSection({ title: "평가 개요" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "평가 개요" });
  s.addText("한눈에 보기", { placeholder: "title" });
  [["5문장", "버킷리스트 작성"], ["7점", "쓰기"], ["5점", "말하기 (인터뷰)"]].forEach(([big, label], i) => {
    const x = 0.6 + i * 4.2;
    card(s, x, 1.9, 3.9, 3.4, `Stat card ${i + 1}`);
    txt(s, big, { x, y: 2.3, w: 3.9, h: 1.6, fontSize: 66, bold: true, color: C.accent1, align: "center", valign: "middle", objectName: `Stat value ${i + 1}` });
    txt(s, label, { x, y: 4.1, w: 3.9, h: 0.8, fontSize: 28, color: C.text2, align: "center", valign: "middle", objectName: `Stat label ${i + 1}` });
  });
  txt(s, [{ text: "일본 여행지에서 해 보고 싶은 것 5가지를 " }, { text: "〜たいです", options: { fontFace: JP, color: C.accent1 } }, { text: "로 쓰고 말해요" }],
    { x: 0.6, y: 5.75, w: 12.1, h: 0.8, fontSize: 28, bold: true, color: C.text1, valign: "middle", objectName: "Summary" });

  // 3. 문법: 〜たいです
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "평가 개요" });
  s.addText([{ text: "희망표현 " }, { text: "〜たいです", options: { fontFace: JP } }], { placeholder: "title" });
  card(s, 0.6, 1.7, 12.1, 1.3, "Formula card", C.text2);
  txt(s, [
    { text: "동사 ます형  −  ", options: { color: C.background1 } },
    { text: "ます", options: { color: C.accent3, fontFace: JP, strike: "sngStrike" } },
    { text: "  +  ", options: { color: C.background1 } },
    { text: "たいです", options: { color: C.accent3, fontFace: JP } },
    { text: "   = ~하고 싶습니다", options: { color: C.background1 } },
  ], { x: 0.6, y: 1.7, w: 12.1, h: 1.3, fontSize: 34, bold: true, align: "center", valign: "middle", objectName: "Formula" });
  [
    ["行（い）きます", "行（い）きたいです", "가고 싶습니다"],
    ["食（た）べます", "食（た）べたいです", "먹고 싶습니다"],
    ["買（か）います", "買（か）いたいです", "사고 싶습니다"],
  ].forEach(([a, b, ko], i) => {
    const y = 3.35 + i * 1.15;
    card(s, 0.6, y, 12.1, 0.95, `Example row ${i + 1}`);
    txt(s, a, { x: 0.9, y, w: 3.6, h: 0.95, fontSize: 28, fontFace: JP, color: C.text1, valign: "middle", objectName: `Ex verb ${i + 1}` });
    s.addImage({ data: ic.arrow, x: 4.55, y: y + 0.27, w: 0.42, h: 0.42, objectName: `Ex arrow ${i + 1}` });
    txt(s, b, { x: 5.2, y, w: 4.3, h: 0.95, fontSize: 28, bold: true, fontFace: JP, color: C.accent1, valign: "middle", objectName: `Ex wish ${i + 1}` });
    txt(s, ko, { x: 9.6, y, w: 2.9, h: 0.95, fontSize: 26, color: C.text2, valign: "middle", objectName: `Ex ko ${i + 1}` });
  });

  // 4. 5가지 조건
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "평가 개요" });
  s.addText("5문장은 이렇게 구성해요", { placeholder: "title" });
  [
    [ic.place, "장소", "東京（とうきょう）タワーに 行（い）きたいです。"],
    [ic.food, "먹거리", "たこやきを 食（た）べたいです。"],
    [ic.see, "볼거리", "富士山（ふじさん）を 見（み）たいです。"],
    [ic.shop, "살 거리", "おみやげを 買（か）いたいです。"],
    [ic.free, "자유 문장", "着物（きもの）を 着（き）たいです。"],
  ].forEach(([icData, head, jp], i) => {
    const y = 1.6 + i * 1.0;
    card(s, 0.6, y, 12.1, 0.85, `Cond row ${i + 1}`);
    circleIcon(s, 0.8, y + 0.1, 0.65, icData, `Cond ${i + 1}`);
    txt(s, head, { x: 1.7, y, w: 2.4, h: 0.85, fontSize: 28, bold: true, color: C.text2, valign: "middle", objectName: `Cond head ${i + 1}` });
    txt(s, jp, { x: 4.1, y, w: 8.4, h: 0.85, fontSize: 26, fontFace: JP, color: C.text1, valign: "middle", objectName: `Cond jp ${i + 1}` });
  });
  txt(s, "※ 모둠활동에서 조사한 여행지 내용으로 써요 (위 문장은 예시)", { x: 0.6, y: 6.6, w: 11.3, h: 0.45, fontSize: 20, color: C.accent5, valign: "middle", objectName: "Cond note" });

  // 5. 작성 규칙
  pres.addSection({ title: "작성 규칙" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "작성 규칙" });
  s.addText("꼭 지켜야 할 규칙", { placeholder: "title" });
  [
    [ic.group, "모둠 조사 내용", "우리 모둠이 조사한\n여행지로 쓰기"],
    [ic.polite, "정중형 쓰기", "〜たいです로\n쓰고 말하기"],
    [ic.ink, "검정·파란 펜", "다른 색 펜,\n연필은 안 돼요"],
    [ic.kanji, "한자엔 히라가나", "한자에는 히라가나를\n함께 써야 인정"],
    [ic.mean, "뜻 쓰기", "문장마다 아래에\n한국어 뜻 쓰기"],
    [ic.ban, "부분 점수 없음", "한 군데라도 틀리면\n그 문장은 0점"],
  ].forEach(([icData, head, desc], i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.6 + col * 4.1, y = 1.7 + row * 2.6;
    const hot = i === 5;
    card(s, x, y, 3.9, 2.35, `Rule card ${i + 1}`, hot ? C.accent1 : C.background2);
    circleIcon(s, x + 0.3, y + 0.3, 0.75, icData, `Rule ${i + 1}`, hot ? C.text2 : C.accent1);
    txt(s, head, { x: x + 1.2, y: y + 0.3, w: 2.6, h: 0.75, fontSize: 24, bold: true, color: hot ? C.background1 : C.text2, valign: "middle", objectName: `Rule head ${i + 1}` });
    txt(s, desc, { x: x + 0.3, y: y + 1.2, w: 3.4, h: 1.0, fontSize: 22, color: hot ? C.background1 : C.text1, valign: "top", objectName: `Rule desc ${i + 1}` });
  });

  // 6. 진행 순서
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "작성 규칙" });
  s.addText("이렇게 진행해요", { placeholder: "title" });
  [
    [ic.group, "모둠 조사", "여행지의 장소·음식\n볼거리·쇼핑 조사"],
    [ic.pen, "5문장 쓰기", "활동지에 일본어\n문장과 뜻 작성"],
    [ic.redo, "말하기 연습", "내 문장을 외워서\n소리 내어 연습"],
    [ic.mic, "인터뷰 평가", "선생님 질문에\n내 문장으로 답하기"],
  ].forEach(([icData, head, desc], i) => {
    const x = 0.6 + i * 3.15;
    circleIcon(s, x + 0.8, 2.2, 1.5, icData, `Step ${i + 1}`, i === 3 ? C.accent1 : C.text2);
    if (i < 3) s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: x + 2.75, y: 2.75, w: 0.4, h: 0.4, rotate: 90, fill: { color: C.accent5 }, line: { type: "none" }, objectName: `Arrow ${i + 1}` });
    txt(s, `${i + 1}단계`, { x, y: 3.95, w: 3.1, h: 0.5, fontSize: 22, bold: true, color: C.accent1, align: "center", objectName: `Step num ${i + 1}` });
    txt(s, head, { x, y: 4.45, w: 3.1, h: 0.75, fontSize: 32, bold: true, color: C.text2, align: "center", objectName: `Step head ${i + 1}` });
    txt(s, desc, { x, y: 5.3, w: 3.1, h: 1.3, fontSize: 24, color: C.text1, align: "center", valign: "top", objectName: `Step desc ${i + 1}` });
  });

  // 7. 쓰기 채점기준
  pres.addSection({ title: "채점 기준" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "채점 기준" });
  s.addText("채점 기준 ① 쓰기 (7점)", { placeholder: "title" });
  card(s, 0.6, 1.7, 3.6, 5.0, "Write callout", C.text2);
  txt(s, "7점", { x: 0.6, y: 2.0, w: 3.6, h: 1.4, fontSize: 66, bold: true, color: C.accent3, align: "center", valign: "middle", objectName: "Write big" });
  txt(s, "오류 없이 쓴\n문장 수로\n점수가 정해져요", { x: 0.6, y: 3.6, w: 3.6, h: 2.2, fontSize: 26, color: C.background1, align: "center", valign: "top", objectName: "Write desc" });
  const hdr = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, align: "center" } });
  const cell = (t, o = {}) => ({ text: t, options: { color: C.text1, ...o } });
  const pt = (t) => cell(t, { bold: true, color: C.accent1, align: "center" });
  s.addTable([
    [hdr("조건에 맞게 오류 없이 쓴 문장"), hdr("점수")],
    [cell("5가지 조건 모두"), pt("7점")],
    [cell("4문장"), pt("6점")],
    [cell("3문장"), pt("5점")],
    [cell("2문장"), pt("4점")],
    [cell("1문장"), pt("3점")],
    [cell("참여했으나 미제출 (백지 제출 등)"), pt("2점")],
  ], {
    x: 4.5, y: 1.7, w: 8.2, colW: [6.4, 1.8], rowH: 0.71,
    fontSize: 22, valign: "middle", margin: [0, 0.15, 0, 0.2],
    border: { type: "solid", pt: 1, color: "D5DAE6" }, fill: { color: "FFFFFF" }, objectName: "Write rubric",
  });

  // 8. 말하기 채점기준
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "채점 기준" });
  s.addText("채점 기준 ② 인터뷰 말하기 (5점)", { placeholder: "title" });
  [
    ["5점", "내 문장 3개", "질문 정확히 이해\n발음·문법 정확\n유창하게 답변"],
    ["4점", "내 문장 2개", "질문 정확히 이해\n올바르게 활용\n답변"],
    ["3점", "내 문장 1개", "질문에 대해\n올바르게 활용\n답변"],
    ["2점", "기본 점수", "참여했으나\n미제출\n(백지 제출 등)"],
  ].forEach(([score, head, desc], i) => {
    const x = 0.6 + i * 3.08;
    const top = i === 0;
    card(s, x, 1.7, 2.85, 5.0, `Speak card ${i + 1}`, top ? C.text2 : C.background2);
    txt(s, score, { x, y: 1.95, w: 2.85, h: 1.2, fontSize: 60, bold: true, color: top ? C.accent3 : C.accent1, align: "center", valign: "middle", objectName: `Speak score ${i + 1}` });
    txt(s, head, { x, y: 3.3, w: 2.85, h: 0.7, fontSize: 28, bold: true, color: top ? C.background1 : C.text2, align: "center", valign: "middle", objectName: `Speak head ${i + 1}` });
    txt(s, desc, { x: x + 0.15, y: 4.2, w: 2.55, h: 2.2, fontSize: 22, color: top ? C.background1 : C.text1, align: "center", valign: "top", objectName: `Speak desc ${i + 1}` });
  });

  // 9. 인터뷰 예시
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "채점 기준" });
  s.addText("인터뷰는 이런 느낌이에요", { placeholder: "title" });
  [
    [true, "日本（にほん）で 何（なに）を したいですか。", "일본에서 무엇을 하고 싶어요?"],
    [false, "大阪（おおさか）で たこやきを 食（た）べたいです。", "오사카에서 다코야키를 먹고 싶습니다."],
    [true, "何（なに）を 買（か）いたいですか。", "무엇을 사고 싶어요?"],
    [false, "おみやげを 買（か）いたいです。", "기념품을 사고 싶습니다."],
  ].forEach(([teacher, jp, ko], i) => {
    const y = 1.65 + i * 1.22;
    const x = teacher ? 0.6 : 2.1;
    card(s, x, y, 10.6, 1.08, `Bubble ${i + 1}`, teacher ? C.background2 : C.accent6);
    txt(s, teacher ? "선생님" : "나", { x: x + 0.25, y, w: 1.3, h: 1.08, fontSize: 22, bold: true, color: teacher ? C.text2 : C.accent1, valign: "middle", objectName: `Speaker ${i + 1}` });
    txt(s, jp, { x: x + 1.55, y: y + 0.08, w: 8.8, h: 0.55, fontSize: 24, bold: true, fontFace: JP, color: C.text1, valign: "middle", objectName: `Bubble jp ${i + 1}` });
    txt(s, ko, { x: x + 1.55, y: y + 0.6, w: 8.8, h: 0.4, fontSize: 18, color: C.accent5, valign: "middle", objectName: `Bubble ko ${i + 1}` });
  });
  txt(s, "※ 예시 대화예요. 실제 질문은 달라질 수 있어요", { x: 0.6, y: 6.6, w: 11.3, h: 0.45, fontSize: 20, color: C.accent5, valign: "middle", objectName: "Interview note" });

  // 10. 마무리
  pres.addSection({ title: "마무리" });
  s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "마무리" });
  s.addText("がんばってください！", { placeholder: "title", fontFace: JP });
  s.addText("궁금한 점은 언제든지 선생님께 물어보세요", { placeholder: "body" });
  s.addImage({ data: ic.plane, x: 10.8, y: 1.0, w: 1.4, h: 1.4, objectName: "Plane end" });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
