// 일본 문화광장(ぶんかのひろば) 1~4 수업용 슬라이드 생성기
// 사용법: node scripts/build_culture_ppt.js 일본문화광장_수업.pptx
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const gi = require("react-icons/gi");
const { applyTheme } = require("/root/.claude/skills/synced/58492139-558f-4c31-bab3-9821f6393e30_b3cf6d08-3c01-4a4e-ad7a-826a6a0e3731/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";

// 한글: Jua(주아) – 둥글고 굵어 멀리서도 잘 보이는 귀여운 글꼴 (Google Fonts, 무료)
const KO = "Jua";

// 교재의 청록 차양 + 주홍 포인트 + 겨자색
const THEME = {
  name: "Bunka no Hiroba",
  headFontFace: KO,
  bodyFontFace: KO,
  colors: {
    dk1: "1E2426", lt1: "FFFFFF", dk2: "0B5E5C", lt2: "EAF6F5",
    accent1: "E4572E", accent2: "0B5E5C", accent3: "F2B544", accent4: "1AAFA5",
    accent5: "5F7472", accent6: "FDE6DC", hlink: "E4572E", folHlink: "5F7472",
  },
};
const JP = "Meiryo"; // 일본어용 글꼴

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.title = "일본 문화광장 – ぶんかのひろば";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------- 후처리: 한글 글꼴 지정 + 클릭 시 정답 등장 애니메이션 ----------
const JSZip = require("jszip");

// 한 개 도형이 클릭할 때 '밝기 변화(페이드)'로 나타나는 시퀀스
function revealTiming(spids) {
  let id = 3;
  const clicks = spids.map((spid) => {
    const a = id, b = id + 1, c = id + 2, d = id + 3, e = id + 4;
    id += 5;
    return `<p:par><p:cTn id="${a}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>`
      + `<p:par><p:cTn id="${b}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>`
      + `<p:par><p:cTn id="${c}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>`
      + `<p:set><p:cBhvr><p:cTn id="${d}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>`
      + `<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${e}" dur="400"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:animEffect>`
      + `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`;
  }).join("");
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>`
    + `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${clicks}</p:childTnLst></p:cTn>`
    + `<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>`
    + `<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>`
    + `</p:childTnLst></p:cTn></p:par></p:tnLst>`
    + `<p:bldLst>${spids.map((sp) => `<p:bldP spid="${sp}" grpId="0" animBg="1"/>`).join("")}</p:bldLst></p:timing>`;
}

async function postProcess(file, reveals) {
  const fs = require("fs");
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  // 테마: 한글(동아시아) 글꼴도 Jua로
  const tp = "ppt/theme/theme1.xml";
  let theme = await zip.file(tp).async("string");
  theme = theme.replace(/<a:ea typeface="[^"]*"\/>/g, `<a:ea typeface="${KO}"/>`)
    .replace(/<a:font script="Hang" typeface="[^"]*"\/>/g, `<a:font script="Hang" typeface="${KO}"/>`);
  zip.file(tp, theme);
  // Jua는 원래 굵은 글꼴이라 '굵게'를 쓰면 뭉개짐 → 일본어(Meiryo)가 아닌 글자의 굵게 해제
  const unbold = (x) => x.replace(/<a:(rPr|defRPr)\b([^>]*?)\sb="1"([^>]*?)(\/>|>([\s\S]*?)<\/a:\1>)/g,
    (m, tag, pre, post, end, inner) => (inner && inner.includes("Meiryo")) ? m : `<a:${tag}${pre}${post}${end}`);
  for (const f of Object.keys(zip.files).filter((n) => /^ppt\/(slides|slideLayouts)\/[^/]+\.xml$/.test(n))) {
    zip.file(f, unbold(await zip.file(f).async("string")));
  }
  for (const { slide, names } of reveals) {
    const sp = `ppt/slides/slide${slide}.xml`;
    let xml = await zip.file(sp).async("string");
    const spids = names.map((n) => {
      const m = xml.match(new RegExp(`<p:cNvPr id="(\\d+)" name="${n}"`));
      if (!m) throw new Error(`shape not found: ${n} in ${sp}`);
      return m[1];
    });
    xml = xml.replace("</p:sld>", revealTiming(spids) + "</p:sld>");
    zip.file(sp, xml);
  }
  // 폴더 항목 없이 파일만 다시 담기 (클라우드 열기 호환)
  const out = new JSZip();
  for (const f of Object.values(zip.files)) if (!f.dir) out.file(f.name, await f.async("nodebuffer"), { createFolders: false });
  fs.writeFileSync(file, await out.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.6, w: 11.7, h: 1.5, fontSize: 60, bold: true, color: C.background1, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.2, w: 11.7, h: 1.0, fontSize: 32, color: C.accent3, align: "left", valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "SECTION",
  background: { color: C.accent4 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.55, w: 11.7, h: 1.4, fontSize: 60, bold: true, color: C.background1, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.75, w: 11.7, h: 1.4, fontSize: 30, color: C.text1, align: "left", valign: "top", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 12.3, y: 6.95, w: 0.6, h: 0.35, fontSize: 14, color: C.background1, align: "right" },
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 1.0, fontSize: 40, bold: true, color: C.text2, align: "left", valign: "middle", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 12.3, y: 6.95, w: 0.6, h: 0.35, fontSize: 14, color: C.accent5, align: "right" },
});

(async () => {
  const W = "FFFFFF";
  const ic = {};
  for (const [k, comp, col] of [
    ["bow", fa.FaHandshake, W], ["gift", fa.FaGift, W], ["torii", fa.FaToriiGate, W], ["fire", gi.GiFireworkRocket, W],
    ["card", fa.FaIdCard, W], ["hand", fa.FaHandHolding, W], ["eye", fa.FaEye, W],
    ["home", fa.FaHome, W], ["cake", fa.FaBirthdayCake, W], ["sun", fa.FaSun, W], ["snow", fa.FaSnowflake, W],
    ["redo", fa.FaRedoAlt, W], ["cookie", fa.FaCookieBite, W], ["pen", fa.FaPen, W],
    ["grad", fa.FaGraduationCap, W], ["heart", fa.FaHeartbeat, W], ["car", fa.FaCar, W],
    ["laptop", fa.FaLaptop, W], ["flask", fa.FaFlask, W], ["palette", fa.FaPalette, W],
    ["cert", fa.FaCertificate, W], ["univ", fa.FaUniversity, W], ["users", fa.FaUsers, W],
    ["q", fa.FaQuestion, W], ["check", fa.FaCheck, W], ["bulb", fa.FaLightbulb, "F2B544"],
    ["arrow", fa.FaArrowRight, "E4572E"], ["arrowDown", fa.FaArrowDown, "E4572E"],
    ["bowBig", fa.FaHandshake, "F2B544"], ["giftBig", fa.FaGift, "F2B544"], ["toriiBig", fa.FaToriiGate, "F2B544"], ["fireBig", gi.GiFireworkRocket, "F2B544"],
  ]) ic[k] = await icon(comp, col);

  const clickReveal = []; // [{ slide: 번호, names: [objectName...] }]
  const txt = (sl, t, o) => sl.addText(t, { margin: 0, isTextBox: true, ...o });
  const card = (sl, x, y, w, h, name, color = C.background2) =>
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.15, fill: { color }, line: { type: "none" }, objectName: name });
  const circleIcon = (sl, x, y, d, data, name, color = C.accent1) => {
    sl.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color }, line: { type: "none" }, objectName: `${name} circle` });
    sl.addImage({ data, x: x + d * 0.25, y: y + d * 0.25, w: d * 0.5, h: d * 0.5, objectName: `${name} icon` });
  };
  const jp = (t, o = {}) => ({ text: t, options: { fontFace: JP, ...o } });
  const ko = (t, o = {}) => ({ text: t, options: o });

  // 생각해 보기 / 자기 의견 쓰기 박스 (빈칸 + 정답 공개용 노트)
  const thinkBox = (sl, x, y, w, h, label, runs, name) => {
    card(sl, x, y, w, h, `${name} box`, C.accent6);
    sl.addImage({ data: ic.bulb, x: x + 0.3, y: y + (h - 0.6) / 2, w: 0.6, h: 0.6, objectName: `${name} bulb` });
    txt(sl, label, { x: x + 1.1, y: y + 0.15, w: w - 1.3, h: 0.45, fontSize: 20, bold: true, color: C.accent1, valign: "middle", objectName: `${name} label` });
    txt(sl, runs, { x: x + 1.1, y: y + 0.6, w: w - 1.3, h: h - 0.75, fontSize: 26, color: C.text1, valign: "middle", objectName: `${name} text` });
  };

  // ---------- 퀴즈 ----------
  // items: [질문 runs, 정답 runs]
  // 정답은 슬라이드 쇼에서 클릭(Enter)할 때마다 하나씩 나타납니다 (addClickReveal 참고)
  const quiz = (sectionTitle, num, items) => {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle });
    s.addText(`확인 퀴즈 ${num}`, { placeholder: "title" });
    const names = [];
    items.forEach(([q, a], i) => {
      const y = 1.55 + i * 1.1;
      card(s, 0.6, y, 12.1, 0.95, `Q${i + 1} row`, C.background2);
      s.addShape(pres.shapes.OVAL, { x: 0.8, y: y + 0.15, w: 0.65, h: 0.65, fill: { color: C.text2 }, line: { type: "none" }, objectName: `Q${i + 1} num circle` });
      txt(s, String(i + 1), { x: 0.8, y: y + 0.15, w: 0.65, h: 0.65, fontSize: 24, color: C.background1, align: "center", valign: "middle", objectName: `Q${i + 1} num` });
      txt(s, q, { x: 1.65, y, w: 6.45, h: 0.95, fontSize: 24, color: C.text1, valign: "middle", objectName: `Q${i + 1} question` });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 8.4, y: y + 0.12, w: 4.1, h: 0.71, rectRadius: 0.1,
        fill: { color: C.background1 }, line: { color: C.accent4, width: 2, dashType: "dash" },
        objectName: `Q${i + 1} answer blank`,
      });
      const name = `Q${i + 1} answer`;
      s.addText(a, {
        shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, x: 8.37, y: y + 0.09, w: 4.16, h: 0.77, margin: 0,
        fill: { color: C.accent1 }, line: { type: "none" }, fontSize: 26, bold: true, color: C.background1,
        align: "center", valign: "middle", objectName: name,
      });
      names.push(name);
    });
    clickReveal.push({ slide: pres.slides.length, names });
    s.addNotes("학생들이 공책이나 활동지에 답을 적게 한 뒤, Enter(또는 클릭)를 누를 때마다 1번부터 정답이 하나씩 나타납니다. 틀린 문항은 앞 슬라이드로 돌아가 다시 짚어 주세요.");
  };

  // ================= 표지 =================
  pres.addSection({ title: "표지" });
  let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "표지" });
  txt(s, "日本語会話 · 문화 수업", { x: 0.8, y: 1.6, w: 8.5, h: 0.7, fontSize: 28, color: C.accent6, objectName: "Kicker" });
  s.addText([jp("ぶんかの ひろば", { bold: true })], { placeholder: "title" });
  s.addText("일본 문화광장 1 ~ 4", { placeholder: "body" });
  [ic.bowBig, ic.giftBig, ic.toriiBig, ic.fireBig].forEach((d, i) =>
    s.addImage({ data: d, x: 1.0 + i * 1.5, y: 5.6, w: 0.9, h: 0.9, objectName: `Cover icon ${i + 1}` }));

  // ================= 오늘 배울 내용 =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "표지" });
  s.addText("오늘 만날 4개의 문화광장", { placeholder: "title" });
  [
    [ic.bow, "1", "인사 문화", "えしゃく\nけいれい\nさいけいれい"],
    [ic.gift, "2", "선물 문화", "おみやげ\nおちゅうげん\nおせいぼ\nおかえし"],
    [ic.torii, "3", "오미쿠지와\n오마모리", "えんぎもの\nおみくじ\nおまもり"],
    [ic.fire, "4", "불꽃 축제와\n관련 직업", "はなび\nはなびし"],
  ].forEach(([d, n, head, kw], i) => {
    const x = 0.6 + i * 3.08;
    card(s, x, 1.65, 2.85, 5.0, `Overview card ${i + 1}`, C.background2);
    circleIcon(s, x + 0.85, 1.95, 1.15, d, `Overview ${i + 1}`, i % 2 ? C.accent1 : C.text2);
    txt(s, `문화광장 ${n}`, { x, y: 3.3, w: 2.85, h: 0.5, fontSize: 20, bold: true, color: C.accent1, align: "center", valign: "middle", objectName: `Overview num ${i + 1}` });
    txt(s, head, { x, y: 3.85, w: 2.85, h: 1.2, fontSize: 28, bold: true, color: C.text2, align: "center", valign: "middle", objectName: `Overview head ${i + 1}` });
    txt(s, kw, { x: x + 0.1, y: 5.05, w: 2.65, h: 1.5, fontSize: 20, fontFace: JP, color: C.text1, align: "center", valign: "middle", objectName: `Overview kw ${i + 1}` });
  });
  s.addNotes("각 문화광장마다 핵심 키워드 → 확인 퀴즈 순서로 진행합니다. 학습지의 빈칸(네모 칸)이 핵심 키워드입니다.");

  // ================= 문화광장 1: 인사 =================
  let sec = "문화광장 1 인사";
  pres.addSection({ title: sec });
  s = pres.addSlide({ masterName: "SECTION", sectionTitle: sec });
  txt(s, "문화광장 1", { x: 0.8, y: 1.7, w: 6, h: 0.8, fontSize: 32, bold: true, color: C.text2, objectName: "Section kicker" });
  s.addText("일본의 인사 문화", { placeholder: "title" });
  s.addText([ko("키워드  "), jp("えしゃく · けいれい · さいけいれい", { bold: true })], { placeholder: "body" });
  s.addImage({ data: ic.bowBig, x: 10.4, y: 0.9, w: 1.8, h: 1.8, objectName: "Section icon" });

  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
  s.addText("고개를 숙이는 각도로 구분해요", { placeholder: "title" });
  [
    ["15°", "えしゃく", "가벼운 인사", "가장 일반적인\n인사법"],
    ["30°", "けいれい", "경례", "처음 만났을 때,\n자기소개 할 때"],
    ["45~90°", "さいけいれい", "경례보다 공손한 인사", "아주 높은 사람,\n존경하는 사람에게"],
  ].forEach(([deg, word, kname, desc], i) => {
    const x = 0.6 + i * 4.1;
    const hl = i === 1;
    card(s, x, 1.6, 3.85, 5.1, `Bow card ${i + 1}`, hl ? C.text2 : C.background2);
    txt(s, deg, { x, y: 1.8, w: 3.85, h: 1.2, fontSize: 60, bold: true, color: hl ? C.accent3 : C.accent1, align: "center", valign: "middle", objectName: `Bow deg ${i + 1}` });
    txt(s, word, { x, y: 3.15, w: 3.85, h: 0.85, fontSize: 36, bold: true, fontFace: JP, color: hl ? C.background1 : C.text2, align: "center", valign: "middle", objectName: `Bow word ${i + 1}` });
    txt(s, kname, { x, y: 4.1, w: 3.85, h: 0.6, fontSize: 24, bold: true, color: hl ? C.accent6 : C.accent5, align: "center", valign: "middle", objectName: `Bow kname ${i + 1}` });
    txt(s, desc, { x: x + 0.2, y: 4.9, w: 3.45, h: 1.5, fontSize: 24, color: hl ? C.background1 : C.text1, align: "center", valign: "middle", objectName: `Bow desc ${i + 1}` });
  });
  s.addNotes("만남·감사·사과·부탁할 때 인사하는 것은 세계 공통이지만, 일본은 상황에 따라 고개를 숙이는 각도가 다릅니다. 학생들과 15°, 30°, 45° 인사를 직접 따라 해 보세요.");

  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
  s.addText("인사하며 자기소개 하기", { placeholder: "title" });
  const steps = [
    ["경례", "30°  けいれい"], ["첫 만남의 인사", "はじめまして。"], ["이름과 소속 소개", "〇〇です。"],
    ["나를 소개하는 코멘트", "취미 · 특기"], ["마무리 인사", "よろしく おねがいします。"], ["경례", "30°  けいれい"],
  ];
  steps.forEach(([head, sub], i) => {
    const row = Math.floor(i / 3), col = i % 3;
    const x = 0.6 + col * 4.15, y = 1.55 + row * 1.95;
    const hl = i === 4;
    card(s, x, y, 3.6, 1.6, `Step ${i + 1}`, hl ? C.accent6 : C.background2);
    s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: y + 0.2, w: 0.55, h: 0.55, fill: { color: C.accent1 }, line: { type: "none" }, objectName: `Step ${i + 1} num circle` });
    txt(s, String(i + 1), { x: x + 0.2, y: y + 0.2, w: 0.55, h: 0.55, fontSize: 22, bold: true, color: C.background1, align: "center", valign: "middle", objectName: `Step ${i + 1} num` });
    txt(s, head, { x: x + 0.9, y: y + 0.15, w: 2.6, h: 0.65, fontSize: 22, bold: true, color: C.text2, valign: "middle", objectName: `Step ${i + 1} head` });
    txt(s, sub, { x: x + 0.2, y: y + 0.85, w: 3.3, h: 0.6, fontSize: hl ? 20 : 24, bold: true, fontFace: JP, color: C.accent1, valign: "middle", objectName: `Step ${i + 1} sub` });
    if (col < 2) s.addImage({ data: ic.arrow, x: x + 3.7, y: y + 0.6, w: 0.4, h: 0.4, objectName: `Step ${i + 1} arrow` });
  });
  thinkBox(s, 0.6, 5.55, 12.1, 1.3, "생각해 보기",
    [ko("경례(けいれい)는 처음 만났을 때 또는 "), ko("(        )", { bold: true, color: C.accent1 }), ko(" 할 때 하는 인사!")], "Think 1");
  s.addNotes("정답: 자기소개. ⑤ 마무리 인사 よろしく おねがいします(잘 부탁드립니다)가 학습지 빈칸입니다. 짝과 함께 경례 → 자기소개 → 경례 순서로 연습시켜 보세요.");

  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
  s.addText([ko("비즈니스에서는 명함("), jp("めいし"), ko(") 교환")], { placeholder: "title" });
  [
    [ic.card, "미리 준비하기", "명함은 만나기 전에 미리 준비해요"],
    [ic.hand, "오른손으로 건네고, 왼손으로 받기", "주로 오른손으로 건네고 왼손으로 받아요"],
    [ic.eye, "천천히, 또박또박", "상대방의 눈을 보며 말해요"],
  ].forEach(([d, head, desc], i) => {
    const y = 1.65 + i * 1.75;
    circleIcon(s, 0.8, y + 0.1, 1.2, d, `Card manner ${i + 1}`, i === 1 ? C.accent1 : C.text2);
    txt(s, head, { x: 2.4, y, w: 10.2, h: 0.8, fontSize: 32, bold: true, color: C.text2, valign: "middle", objectName: `Card manner head ${i + 1}` });
    txt(s, desc, { x: 2.4, y: y + 0.8, w: 10.2, h: 0.6, fontSize: 24, color: C.accent5, valign: "middle", objectName: `Card manner desc ${i + 1}` });
  });
  s.addNotes("탐구 활동: 세계의 다양한 인사 문화를 조사하여 발표해 봅시다.");

  quiz(sec, 1, [
    [[ko("고개를 "), ko("15°", { bold: true }), ko(" 숙이는 가장 일반적인 가벼운 인사는?")], [jp("えしゃく")]],
    [[ko("처음 만났을 때·자기소개 할 때 "), ko("30°", { bold: true }), ko(" 숙이는 인사는?")], [jp("けいれい")]],
    [[ko("아주 높은 사람에게 "), ko("45°~90°", { bold: true }), ko(" 숙이는 공손한 인사는?")], [jp("さいけいれい")]],
    [[ko("자기소개를 마칠 때 하는 인사말 ‘잘 부탁드립니다’는?")], [jp("よろしく おねがいします", { fontSize: 20 })]],
    [[ko("명함은 주로 (   )손으로 건네고, 왼손으로 받는다")], [ko("오른손")]],
  ]);

  // ================= 문화광장 2: 선물 =================
  sec = "문화광장 2 선물";
  pres.addSection({ title: sec });
  s = pres.addSlide({ masterName: "SECTION", sectionTitle: sec });
  txt(s, "문화광장 2", { x: 0.8, y: 1.7, w: 6, h: 0.8, fontSize: 32, bold: true, color: C.text2, objectName: "Section kicker" });
  s.addText("일본의 선물 문화", { placeholder: "title" });
  s.addText([ko("키워드  "), jp("おみやげ · プレゼント · おちゅうげん · おせいぼ · おかえし", { bold: true })], { placeholder: "body" });
  s.addImage({ data: ic.giftBig, x: 10.4, y: 0.9, w: 1.8, h: 1.8, objectName: "Section icon" });

  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
  s.addText("선물의 용도에 따라 이름이 달라요", { placeholder: "title" });
  [
    [ic.home, "おみやげ", "오미야게", "남의 집 방문할 때 준비하는 선물\n여행지에서 사 온 선물", C.text2],
    [ic.cake, "プレゼント", "프레젠토", "생일, 크리스마스 등\n특별한 날에 주는 선물", C.accent1],
  ].forEach(([d, word, read, desc, col], i) => {
    const x = 0.6 + i * 6.2;
    card(s, x, 1.6, 5.9, 5.1, `Gift card ${i + 1}`, C.background2);
    circleIcon(s, x + 2.35, 1.85, 1.2, d, `Gift ${i + 1}`, col);
    txt(s, word, { x, y: 3.2, w: 5.9, h: 1.0, fontSize: 48, bold: true, fontFace: JP, color: col, align: "center", valign: "middle", objectName: `Gift word ${i + 1}` });
    txt(s, read, { x, y: 4.2, w: 5.9, h: 0.55, fontSize: 24, color: C.accent5, align: "center", valign: "middle", objectName: `Gift read ${i + 1}` });
    txt(s, desc, { x: x + 0.3, y: 4.95, w: 5.3, h: 1.4, fontSize: 26, color: C.text1, align: "center", valign: "middle", objectName: `Gift desc ${i + 1}` });
  });
  s.addNotes("로보타로: 일본에서는 선물의 용도에 따라 이름이 다릅니다.");

  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
  s.addText("계절마다 감사의 마음을 전하는 선물", { placeholder: "title" });
  txt(s, "평소 신세를 진 지인, 직장 상사, 친척 등에게 보내요", { x: 0.6, y: 1.45, w: 12.1, h: 0.6, fontSize: 26, color: C.accent5, valign: "middle", objectName: "Season lead" });
  [
    [ic.sun, "おちゅうげん", "오추겐", "7월 초 ~ 8월 중순", "여름 선물", C.accent1],
    [ic.snow, "おせいぼ", "오세보", "연말 (12월)", "연말 선물", C.text2],
  ].forEach(([d, word, read, when, tag, col], i) => {
    const x = 0.6 + i * 6.2;
    card(s, x, 2.3, 5.9, 4.4, `Season card ${i + 1}`, C.background2);
    circleIcon(s, x + 0.4, 2.6, 1.1, d, `Season ${i + 1}`, col);
    txt(s, tag, { x: x + 1.75, y: 2.6, w: 3.9, h: 1.1, fontSize: 28, bold: true, color: col, valign: "middle", objectName: `Season tag ${i + 1}` });
    txt(s, word, { x, y: 3.9, w: 5.9, h: 1.0, fontSize: 48, bold: true, fontFace: JP, color: C.text1, align: "center", valign: "middle", objectName: `Season word ${i + 1}` });
    txt(s, read, { x, y: 4.85, w: 5.9, h: 0.5, fontSize: 22, color: C.accent5, align: "center", valign: "middle", objectName: `Season read ${i + 1}` });
    txt(s, when, { x, y: 5.5, w: 5.9, h: 0.8, fontSize: 32, bold: true, color: col, align: "center", valign: "middle", objectName: `Season when ${i + 1}` });
  });
  s.addNotes("지역마다 시기는 조금씩 차이가 있습니다. 오추겐은 7월 초~8월 중순, 오세보는 연말에 보내는 선물입니다.");

  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
  s.addText([ko("받으면 꼭 답례하기 – "), jp("おかえし")], { placeholder: "title" });
  [
    [ic.redo, [jp("おかえし", { bold: true, color: C.accent1 }), ko("  오카에시")], "선물을 받으면 가까운 시일 안에 답례하거나 감사의 말을 전해요"],
    [ic.cookie, [ko("어떤 선물을 줄까?")], "실용적이면서 부담스럽지 않은 금액의 과자류 등을 준비해요"],
  ].forEach(([d, head, desc], i) => {
    const y = 1.6 + i * 1.8;
    circleIcon(s, 0.8, y + 0.1, 1.2, d, `Okaeshi ${i + 1}`, i ? C.text2 : C.accent1);
    txt(s, head, { x: 2.4, y, w: 10.2, h: 0.75, fontSize: 32, bold: true, color: C.text2, valign: "middle", objectName: `Okaeshi head ${i + 1}` });
    txt(s, desc, { x: 2.4, y: y + 0.75, w: 10.2, h: 0.75, fontSize: 24, color: C.text1, valign: "middle", objectName: `Okaeshi desc ${i + 1}` });
  });
  thinkBox(s, 0.6, 5.3, 12.1, 1.5, "생각해 보기",
    [ko("일본 친구에게 한국에서 가져온 과자나 특산물을 주면 "), ko("(        )", { bold: true, color: C.accent1 }), ko("가 되겠네!")], "Think 2");
  s.addNotes("정답: おみやげ(오미야게). 탐구 활동: 일본인이 한국에서 많이 사는 오미야게는 무엇이 있는지 조사하여 발표해 봅시다.");

  quiz(sec, 2, [
    [[ko("남의 집을 방문할 때 준비하거나, 여행지에서 사 온 선물은?")], [jp("おみやげ")]],
    [[ko("생일·크리스마스 등 특별한 날에 주는 선물은?")], [jp("プレゼント")]],
    [[ko("7월 초 ~ 8월 중순에 감사의 마음을 담아 보내는 선물은?")], [jp("おちゅうげん")]],
    [[ko("연말에 감사의 마음을 담아 보내는 선물은?")], [jp("おせいぼ")]],
    [[ko("선물을 받은 뒤 보답으로 하는 답례를 무엇이라고 할까?")], [jp("おかえし")]],
  ]);

  // ================= 문화광장 3: 오미쿠지·오마모리 =================
  sec = "문화광장 3 오미쿠지";
  pres.addSection({ title: sec });
  s = pres.addSlide({ masterName: "SECTION", sectionTitle: sec });
  txt(s, "문화광장 3", { x: 0.8, y: 1.7, w: 6, h: 0.8, fontSize: 32, bold: true, color: C.text2, objectName: "Section kicker" });
  s.addText("일본의 오미쿠지와 오마모리", { placeholder: "title" });
  s.addText([ko("키워드  "), jp("えんぎもの · おみくじ · おまもり", { bold: true })], { placeholder: "body" });
  s.addImage({ data: ic.toriiBig, x: 10.4, y: 0.9, w: 1.8, h: 1.8, objectName: "Section icon" });

  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
  s.addText([ko("행운을 비는 물건 – "), jp("えんぎもの")], { placeholder: "title" });
  card(s, 0.6, 1.6, 12.1, 1.6, "Engimono card", C.text2);
  txt(s, [jp("えんぎもの", { bold: true, color: C.accent3 }), ko("  엔기모노")], { x: 1.0, y: 1.7, w: 11.3, h: 0.7, fontSize: 34, color: C.background1, valign: "middle", objectName: "Engimono word" });
  txt(s, "행운을 부르거나 길흉을 점치기 위한 물건 – 일본은 운수를 중요하게 여겨요", { x: 1.0, y: 2.45, w: 11.3, h: 0.6, fontSize: 24, color: C.background1, valign: "middle", objectName: "Engimono desc" });
  txt(s, [jp("おみくじ", { bold: true, color: C.accent1 }), ko("  오미쿠지는 엔기모노 중 하나!")], { x: 0.6, y: 3.45, w: 12.1, h: 0.75, fontSize: 32, color: C.text2, valign: "middle", objectName: "Omikuji head" });
  [
    [[ko("어디서?")], [ko("신사("), jp("じんじゃ"), ko(")나 절("), jp("おてら"), ko(")에서")]],
    [[ko("무엇?")], [ko("길흉을 점치는 제비(뽑기)")]],
    [[ko("얼마?")], [ko("보통 100엔 정도")]],
  ].forEach(([head, body], i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 4.45, 3.85, 2.3, `Omikuji fact ${i + 1}`, C.background2);
    txt(s, head, { x: x + 0.3, y: 4.6, w: 3.25, h: 0.6, fontSize: 24, bold: true, color: C.accent1, valign: "middle", objectName: `Omikuji fact head ${i + 1}` });
    txt(s, body, { x: x + 0.3, y: 5.2, w: 3.25, h: 1.4, fontSize: 26, bold: true, color: C.text1, valign: "middle", objectName: `Omikuji fact body ${i + 1}` });
  });
  s.addNotes("오미쿠지 통을 흔들어 뽑은 막대의 숫자에 해당하는 오미쿠지를 확인합니다.");

  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
  s.addText("오미쿠지 결과는 대길부터 흉까지", { placeholder: "title" });
  [
    ["大吉", "だいきち", "대길", "제일 좋은 운세", C.accent1],
    ["吉", "きち", "길", "좋은 운세", C.text2],
    ["凶", "きょう", "흉", "좋지 않은 운세", C.accent5],
  ].forEach(([kanji, read, kname, desc, col], i) => {
    const x = 0.6 + i * 2.55;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.6, w: 2.25, h: 5.1, fill: { color: C.background1 }, line: { color: col, width: 3 }, objectName: `Fortune slip ${i + 1}` });
    txt(s, kanji, { x, y: 1.75, w: 2.25, h: 1.8, fontSize: kanji.length > 1 ? 60 : 80, bold: true, fontFace: JP, color: col, align: "center", valign: "middle", objectName: `Fortune kanji ${i + 1}` });
    txt(s, read, { x, y: 3.65, w: 2.25, h: 0.6, fontSize: 24, fontFace: JP, color: C.text1, align: "center", valign: "middle", objectName: `Fortune read ${i + 1}` });
    txt(s, kname, { x, y: 4.35, w: 2.25, h: 0.8, fontSize: 36, bold: true, color: col, align: "center", valign: "middle", objectName: `Fortune kname ${i + 1}` });
    txt(s, desc, { x: x + 0.1, y: 5.4, w: 2.05, h: 1.0, fontSize: 20, color: C.text1, align: "center", valign: "middle", objectName: `Fortune desc ${i + 1}` });
  });
  card(s, 8.5, 1.6, 4.2, 5.1, "Tie card", C.accent6);
  txt(s, "‘흉’이 나오면?", { x: 8.8, y: 1.85, w: 3.6, h: 0.7, fontSize: 28, bold: true, color: C.accent1, valign: "middle", objectName: "Tie head" });
  txt(s, "신사에 마련된 장소에 묶어 두고 가요", { x: 8.8, y: 2.6, w: 3.6, h: 1.4, fontSize: 26, bold: true, color: C.text1, valign: "top", objectName: "Tie body" });
  txt(s, "‘길’ 이상이면 가지고 가기도 하고, 길흉에 상관없이 모두 묶어 두기도 해요", { x: 8.8, y: 4.3, w: 3.6, h: 2.1, fontSize: 22, color: C.accent5, valign: "top", objectName: "Tie note" });

  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
  s.addText([ko("부적 같은 작은 주머니 – "), jp("おまもり")], { placeholder: "title" });
  txt(s, "신사나 절에서 파는 작은 주머니. 안에는 신의 이름이나 주문이 적혀 있어요", { x: 0.6, y: 1.5, w: 12.1, h: 0.9, fontSize: 26, color: C.text1, valign: "middle", objectName: "Omamori lead" });
  txt(s, "몸에 지니거나 휴대폰·자동차 등에 매달고 다녀요", { x: 0.6, y: 2.35, w: 12.1, h: 0.6, fontSize: 24, color: C.accent5, valign: "middle", objectName: "Omamori lead 2" });
  [
    [ic.grad, "학업 성취", "시험 합격"],
    [ic.heart, "건강", "아프지 않게"],
    [ic.car, "교통안전", "안전한 이동"],
  ].forEach(([d, head, sub], i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 3.2, 3.85, 1.6, `Omamori kind ${i + 1}`, C.background2);
    circleIcon(s, x + 0.25, 3.4, 1.2, d, `Omamori kind ${i + 1}`, i === 0 ? C.accent1 : C.text2);
    txt(s, head, { x: x + 1.6, y: 3.4, w: 2.15, h: 0.65, fontSize: 28, bold: true, color: C.text2, valign: "middle", objectName: `Omamori kind head ${i + 1}` });
    txt(s, sub, { x: x + 1.6, y: 4.05, w: 2.15, h: 0.55, fontSize: 20, color: C.accent5, valign: "middle", objectName: `Omamori kind sub ${i + 1}` });
  });
  thinkBox(s, 0.6, 5.2, 12.1, 1.6, "자기 의견 쓰기",
    [ko("나는 "), ko("(              )", { bold: true, color: C.accent1 }), ko("을/를 위해 오마모리를 사고 싶어.")], "Think 3");
  s.addNotes("학생 각자 오마모리를 사고 싶은 이유를 학습지에 쓰고 발표하게 합니다. 탐구 활동: 일본의 다양한 엔기모노(えんぎもの)에 대해 조사해서 발표해 봅시다.");

  quiz(sec, 3, [
    [[ko("행운을 부르거나 길흉을 점치기 위한 물건은?")], [jp("えんぎもの")]],
    [[ko("신사나 절에서 파는, 길흉을 점치는 제비(뽑기)는?")], [jp("おみくじ")]],
    [[ko("오미쿠지에서 가장 좋은 운세는?")], [jp("大吉（だいきち）", { fontSize: 22 })]],
    [[ko("신사나 절에서 파는, 부적 같은 작은 주머니는?")], [jp("おまもり")]],
    [[ko("‘흉’이 나오면 오미쿠지를 신사에 (   )")], [ko("묶어 둔다")]],
  ]);

  // ================= 문화광장 4: 불꽃 축제 =================
  sec = "문화광장 4 불꽃 축제";
  pres.addSection({ title: sec });
  s = pres.addSlide({ masterName: "SECTION", sectionTitle: sec });
  txt(s, "문화광장 4", { x: 0.8, y: 1.7, w: 6, h: 0.8, fontSize: 32, bold: true, color: C.text2, objectName: "Section kicker" });
  s.addText("불꽃 축제와 관련된 직업", { placeholder: "title" });
  s.addText([ko("키워드  "), jp("はなび · はなびし", { bold: true })], { placeholder: "body" });
  s.addImage({ data: ic.fireBig, x: 10.4, y: 0.9, w: 1.8, h: 1.8, objectName: "Section icon" });

  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
  s.addText([ko("불꽃을 기획하고 준비하는 사람 – "), jp("はなびし")], { placeholder: "title" });
  card(s, 0.6, 1.55, 12.1, 1.55, "Hanabishi card", C.text2);
  txt(s, [jp("はなび", { bold: true, color: C.accent3 }), ko(" 불꽃  +  "), jp("し", { bold: true, color: C.accent3 }), ko(" 사람  =  "), jp("はなびし", { bold: true, color: C.accent3 }), ko(" 하나비시")],
    { x: 1.0, y: 1.6, w: 11.3, h: 0.8, fontSize: 32, color: C.background1, valign: "middle", objectName: "Hanabishi formula" });
  txt(s, "축제 콘셉트에 맞게 불꽃 축제를 기획·연출하는 ‘불꽃 연출가’, ‘불꽃 디자이너’", { x: 1.0, y: 2.4, w: 11.3, h: 0.6, fontSize: 22, color: C.background1, valign: "middle", objectName: "Hanabishi desc" });
  txt(s, "하나비시가 되려면 무엇이 필요할까?", { x: 0.6, y: 3.3, w: 12.1, h: 0.65, fontSize: 28, bold: true, color: C.text2, valign: "middle", objectName: "Need head" });
  [
    [ic.laptop, "컴퓨터 운영 능력"], [ic.flask, "화약에 대한 전문 지식"], [ic.palette, "예술적 감각"],
  ].forEach(([d, head], i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 4.1, 3.85, 1.35, `Need ${i + 1}`, C.background2);
    circleIcon(s, x + 0.2, 4.2, 1.15, d, `Need ${i + 1}`, i === 1 ? C.accent1 : C.text2);
    txt(s, head, { x: x + 1.5, y: 4.1, w: 2.25, h: 1.35, fontSize: 24, bold: true, color: C.text2, valign: "middle", objectName: `Need head ${i + 1}` });
  });
  circleIcon(s, 0.6, 5.75, 0.95, ic.cert, "Need cert", C.accent4);
  txt(s, "‘폭발물 취급/제조 보안 책임자’ 자격 + 안전 세미나로 ‘연기 및 화재 소비 안전 핸드북’ 취득", { x: 1.8, y: 5.65, w: 10.9, h: 1.15, fontSize: 22, color: C.text1, valign: "middle", objectName: "Need cert text" });

  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
  s.addText("한국에도 하나비시 같은 직업이 있을까?", { placeholder: "title" });
  card(s, 0.6, 1.55, 3.85, 3.4, "Korea stat card", C.text2);
  txt(s, "약 20명", { x: 0.6, y: 1.8, w: 3.85, h: 1.4, fontSize: 60, bold: true, color: C.accent3, align: "center", valign: "middle", objectName: "Korea stat" });
  txt(s, "불꽃 축제를 총괄하는\n불꽃 연출가", { x: 0.8, y: 3.3, w: 3.45, h: 1.3, fontSize: 24, color: C.background1, align: "center", valign: "middle", objectName: "Korea stat label" });
  [
    [ic.cert, "화약 관련 자격증", "불꽃 연출가가 되려면 꼭 취득해야 해요"],
    [ic.univ, "대학 강의 ‘불꽃 축제 특수 효과’", "기획·연출·방사 등 모든 과정을 배워요"],
  ].forEach(([d, head, desc], i) => {
    const y = 1.55 + i * 1.75;
    card(s, 4.75, y, 7.95, 1.55, `Korea fact ${i + 1}`, C.background2);
    circleIcon(s, 4.95, y + 0.2, 1.15, d, `Korea fact ${i + 1}`, i ? C.text2 : C.accent1);
    txt(s, head, { x: 6.35, y: y + 0.1, w: 6.2, h: 0.7, fontSize: 26, bold: true, color: C.text2, valign: "middle", objectName: `Korea fact head ${i + 1}` });
    txt(s, desc, { x: 6.35, y: y + 0.8, w: 6.2, h: 0.6, fontSize: 22, color: C.text1, valign: "middle", objectName: `Korea fact desc ${i + 1}` });
  });
  thinkBox(s, 0.6, 5.25, 12.1, 1.6, "자기 의견 쓰기",
    [ko("내가 하나비시("), jp("はなびし"), ko(")라면 "), ko("(           )", { bold: true, color: C.accent1 }), ko(" 모양의 불꽃을 연출하고 싶어.")], "Think 4");
  s.addNotes("학생들이 연출하고 싶은 불꽃 모양을 쓰고 발표하게 합니다. 탐구 활동: 일본의 행사와 관련된 다른 직업에 대해 조사해서 발표해 봅시다.");

  quiz(sec, 4, [
    [[ko("일본어로 ‘불꽃’은?")], [jp("はなび")]],
    [[ko("불꽃 축제를 기획·연출하는 불꽃 연출가는?")], [jp("はなびし")]],
    [[ko("하나비시에게 필요한 것: 컴퓨터 운영 능력, (   )에 대한 전문 지식, 예술적 감각")], [ko("화약")]],
    [[ko("한국에서 활동하는 불꽃 연출가는 약 몇 명?")], [ko("약 20명")]],
    [[ko("한국에서 불꽃 연출가가 되려면 (   ) 관련 자격증을 취득해야 한다")], [ko("화약")]],
  ]);

  // ================= 마무리: 키워드 총정리 =================
  pres.addSection({ title: "마무리" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "마무리" });
  s.addText("키워드 총정리", { placeholder: "title" });
  [
    ["인사", "えしゃく · けいれい\nさいけいれい\nよろしく おねがいします"],
    ["선물", "おみやげ · プレゼント\nおちゅうげん · おせいぼ\nおかえし"],
    ["오미쿠지", "えんぎもの · おみくじ\nおまもり\n大吉 · 吉 · 凶"],
    ["불꽃 축제", "はなび · はなびし"],
  ].forEach(([head, kws], i) => {
    const x = 0.6 + (i % 2) * 6.2, y = 1.55 + Math.floor(i / 2) * 2.65;
    const alt = i === 1 || i === 2;
    card(s, x, y, 5.9, 2.45, `Summary card ${i + 1}`, alt ? C.accent6 : C.background2);
    txt(s, `${i + 1}. ${head}`, { x: x + 0.3, y: y + 0.2, w: 5.3, h: 0.65, fontSize: 28, bold: true, color: alt ? C.accent1 : C.text2, valign: "middle", objectName: `Summary head ${i + 1}` });
    txt(s, kws, { x: x + 0.3, y: y + 0.9, w: 5.3, h: 1.45, fontSize: 24, fontFace: JP, color: C.text1, valign: "top", objectName: `Summary list ${i + 1}` });
  });

  s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "마무리" });
  s.addText([jp("おつかれさまでした！", { bold: true })], { placeholder: "title" });
  s.addText("학습지 빈칸과 자기 의견 쓰기를 완성하고 확인 스탬프를 받으세요", { placeholder: "body", fontSize: 26 });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  await postProcess(OUT, clickReveal);
  console.log("wrote", OUT);
})();
