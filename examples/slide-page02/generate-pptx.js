const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

// ─── Config ────────────────────────────────────────────────────────
const C = {
  accent: "007BFF", dark: "333333", warm: "F59E0B", light: "F8F9FA",
  white: "FFFFFF", gray: "9CA3AF", grayLight: "E5E7EB", grayDark: "6B7280",
  green: "10B981", purple: "8B5CF6", red: "EF4444",
  blueBg: "EBF5FF", yellowBg: "FEF3C7", greenBg: "ECFDF5", purpleBg: "F3E8FF", redBg: "FEE2E2",
};
const FONT = "Meiryo";
const FONT_EN = "Calibri";
const SW = 10, SH = 5.625; // slide width/height in inches
const MX = 0.7; // margin x

// ─── Icon Rendering ────────────────────────────────────────────────
function renderSvg(Icon, color, size = 256) {
  return ReactDOMServer.renderToStaticMarkup(React.createElement(Icon, { color, size: String(size) }));
}
async function toBase64(Icon, color, size = 256) {
  const svg = renderSvg(Icon, color, size);
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ─── Helpers ───────────────────────────────────────────────────────
function shadow() { return { type: "outer", color: "000000", blur: 4, offset: 1, angle: 135, opacity: 0.08 }; }

function addHeader(slide, enTitle, jpTitle) {
  // Blue accent bar
  slide.addShape("rect", { x: MX, y: 0.45, w: 0.06, h: 0.45, fill: { color: C.accent } });
  // EN label
  slide.addText(enTitle, { x: MX + 0.2, y: 0.42, w: 5, h: 0.22, fontSize: 8, fontFace: FONT_EN, color: C.gray, charSpacing: 3, bold: false, margin: 0 });
  // JP title
  slide.addText(jpTitle, { x: MX + 0.2, y: 0.62, w: 6, h: 0.4, fontSize: 22, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
  // Divider line
  slide.addShape("line", { x: MX, y: 1.05, w: SW - MX * 2, h: 0, line: { color: C.grayLight, width: 0.75 } });
  // Logo right
  slide.addText("AI RENTAL CAR", { x: 7.2, y: 0.7, w: 2.5, h: 0.25, fontSize: 8, fontFace: FONT_EN, color: C.gray, align: "right", charSpacing: 2, margin: 0 });
}

function addFooter(slide, pageNum) {
  slide.addShape("line", { x: MX, y: 5.25, w: SW - MX * 2, h: 0, line: { color: "F3F4F6", width: 0.5 } });
  slide.addText("AIレンタカー - Confidential", { x: MX, y: 5.3, w: 4, h: 0.2, fontSize: 7, fontFace: FONT, color: C.gray, margin: 0 });
  slide.addText([
    { text: "Page ", options: { fontSize: 7, fontFace: FONT_EN, color: C.gray } },
    { text: String(pageNum).padStart(2, "0"), options: { fontSize: 9, fontFace: FONT_EN, color: C.accent, bold: true } },
  ], { x: 8, y: 5.3, w: 1.6, h: 0.2, align: "right", margin: 0 });
}

function addDecoCircle(slide, x, y, w, h) {
  slide.addShape("ellipse", { x, y, w, h, fill: { color: C.accent, transparency: 95 } });
}

function addCard(slide, x, y, w, h, opts = {}) {
  const o = { x, y, w, h, fill: { color: opts.fill || C.white }, shadow: shadow() };
  if (opts.borderColor) { o.line = { color: opts.borderColor, width: 0.5 }; }
  slide.addShape("rect", o);
  // Top accent line
  if (opts.topColor) { slide.addShape("rect", { x, y, w, h: 0.04, fill: { color: opts.topColor } }); }
  // Left accent line
  if (opts.leftColor) { slide.addShape("rect", { x, y, w: 0.04, h, fill: { color: opts.leftColor } }); }
}

function addIconCircle(slide, x, y, size, bgColor, iconData) {
  slide.addShape("ellipse", { x, y, w: size, h: size, fill: { color: bgColor } });
  if (iconData) {
    const pad = size * 0.25;
    slide.addImage({ data: iconData, x: x + pad, y: y + pad, w: size - pad * 2, h: size - pad * 2 });
  }
}

// ─── Main ──────────────────────────────────────────────────────────
async function main() {
  console.log("Generating icons...");

  // Pre-render all icons
  const icons = {};
  const iconMap = {
    carSide: fa.FaCarSide, calendar: fa.FaCalendar, mapMarker: fa.FaMapMarkerAlt,
    user: fa.FaUser, briefcase: fa.FaBriefcase, chartPie: fa.FaChartPie,
    cogs: fa.FaCogs, yenSign: fa.FaYenSign, umbrellaBeach: fa.FaUmbrellaBeach,
    star: fa.FaStar, handshake: fa.FaHandshake, robot: fa.FaRobot,
    chartLine: fa.FaChartLine, tags: fa.FaTags, calendarCheck: fa.FaCalendarCheck,
    users: fa.FaUsers, mountain: fa.FaMountain, trophy: fa.FaTrophy,
    home: fa.FaHome, child: fa.FaChild, campground: fa.FaCampground,
    car: fa.FaCar, arrowUp: fa.FaArrowUp, building: fa.FaBuilding,
    lightbulb: fa.FaLightbulb, globe: fa.FaGlobe, search: fa.FaSearch,
    mapMarkedAlt: fa.FaMapMarkedAlt, instagram: fa.FaInstagram,
    userTie: fa.FaUserTie, clipboardCheck: fa.FaClipboardCheck,
    chartBar: fa.FaChartBar, broom: fa.FaBroom, tools: fa.FaTools,
    handsHelping: fa.FaHandsHelping, wrench: fa.FaWrench, shieldAlt: fa.FaShieldAlt,
    box: fa.FaBox, laptop: fa.FaLaptop, handHoldingUsd: fa.FaHandHoldingUsd,
    calculator: fa.FaCalculator, wallet: fa.FaWallet, coins: fa.FaCoins,
    carCrash: fa.FaCarCrash, checkCircle: fa.FaCheckCircle, conciergeBell: fa.FaConciergeBell,
    bullhorn: fa.FaBullhorn, userCog: fa.FaUserCog, bullseye: fa.FaBullseye,
    tasks: fa.FaTasks, envelope: fa.FaEnvelope, chevronRight: fa.FaChevronRight,
    chevronDown: fa.FaChevronDown, minus: fa.FaMinus,
  };

  const colorVariants = {
    blue: "#007BFF", white: "#FFFFFF", warm: "#F59E0B", dark: "#333333",
    green: "#10B981", purple: "#8B5CF6", red: "#EF4444", gray: "#9CA3AF",
    grayDark: "#6B7280",
  };

  // Generate icons in needed colors
  const needed = [
    ["carSide","blue"],["carSide","white"],["calendar","blue"],["mapMarker","blue"],
    ["user","blue"],["briefcase","white"],["chartPie","white"],["cogs","white"],
    ["yenSign","white"],["umbrellaBeach","blue"],["umbrellaBeach","warm"],
    ["star","warm"],["star","blue"],["handshake","green"],["handshake","warm"],
    ["robot","blue"],["robot","purple"],["robot","white"],
    ["chartLine","blue"],["tags","blue"],["calendarCheck","blue"],
    ["users","blue"],["users","white"],["mountain","warm"],["trophy","blue"],["trophy","white"],
    ["home","blue"],["child","warm"],["campground","green"],
    ["car","blue"],["car","purple"],["car","white"],["car","grayDark"],
    ["arrowUp","white"],["arrowUp","blue"],["building","blue"],
    ["lightbulb","white"],["globe","blue"],["search","blue"],
    ["mapMarkedAlt","blue"],["instagram","blue"],
    ["userTie","blue"],["clipboardCheck","blue"],["chartBar","blue"],
    ["broom","grayDark"],["tools","grayDark"],["handsHelping","grayDark"],
    ["wrench","grayDark"],["shieldAlt","warm"],["shieldAlt","green"],
    ["box","green"],["laptop","purple"],["handHoldingUsd","warm"],
    ["calculator","blue"],["wallet","warm"],["wallet","purple"],["wallet","white"],
    ["coins","warm"],["coins","white"],
    ["carCrash","red"],["checkCircle","blue"],["checkCircle","warm"],["checkCircle","green"],
    ["conciergeBell","blue"],["bullhorn","warm"],["userCog","blue"],
    ["bullseye","white"],["tasks","blue"],["envelope","blue"],
    ["chevronRight","gray"],["chevronDown","gray"],["minus","red"],
  ];

  for (const [name, color] of needed) {
    const key = `${name}_${color}`;
    if (!icons[key]) {
      icons[key] = await toBase64(iconMap[name], colorVariants[color]);
    }
  }
  console.log(`Generated ${Object.keys(icons).length} icon variants`);

  // ─── Create Presentation ───────────────────────────────────────
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "AIレンタカー";
  pres.title = "AIレンタカー 事業計画書";

  // ═══════════════════════════════════════════════════════════════
  // SLIDE 01: Cover
  // ═══════════════════════════════════════════════════════════════
  console.log("Slide 01: Cover");
  let s = pres.addSlide();
  addDecoCircle(s, 6, -1.5, 5, 5);
  addDecoCircle(s, -1, 3, 4, 4);
  s.addText("BUSINESS PLAN 2026", { x: 0, y: 1.0, w: SW, h: 0.3, fontSize: 10, fontFace: FONT_EN, color: C.accent, align: "center", charSpacing: 6, margin: 0 });
  s.addText([
    { text: "AI", options: { fontSize: 44, fontFace: FONT, color: C.accent, bold: true } },
    { text: "レンタカー", options: { fontSize: 44, fontFace: FONT, color: C.warm, bold: true } },
  ], { x: 0, y: 1.5, w: SW, h: 0.8, align: "center", margin: 0 });
  s.addText("普通自動車レンタカー事業計画書", { x: 0, y: 2.3, w: SW, h: 0.35, fontSize: 14, fontFace: FONT, color: C.grayDark, align: "center", margin: 0 });
  s.addShape("rect", { x: 4.5, y: 2.8, w: 1, h: 0.05, fill: { color: C.accent } });
  s.addText("「レジャー・週末利用に最適化した、AI活用型レンタカーサービス」", { x: 1, y: 3.0, w: 8, h: 0.35, fontSize: 13, fontFace: FONT, color: C.dark, align: "center", bold: true, margin: 0 });
  // Bottom 3 items
  const coverItems = [
    [icons.calendar_blue, "2026", 2.5],
    [icons.mapMarker_blue, "千葉県", 4.2],
    [icons.user_blue, "個人事業主（副業スタート）", 5.6],
  ];
  for (const [icon, label, x] of coverItems) {
    addIconCircle(s, x, 4.0, 0.35, C.blueBg.replace("#", ""), icon);
    s.addText(label, { x: x + 0.45, y: 4.02, w: 2, h: 0.3, fontSize: 9, fontFace: FONT, color: C.grayDark, valign: "middle", margin: 0 });
  }

  // ═══════════════════════════════════════════════════════════════
  // SLIDE 02: Agenda
  // ═══════════════════════════════════════════════════════════════
  console.log("Slide 02: Agenda");
  s = pres.addSlide();
  addDecoCircle(s, 7, -1, 3, 3);
  addHeader(s, "AGENDA", "目次");
  addFooter(s, 2);
  const agendaItems = [
    ["01", "事業概要 & 提供価値", "事業の全体像と4つの提供価値"],
    ["02", "市場・ターゲット・競合", "市場環境、顧客像、競合との差別化"],
    ["03", "サービス・チャネル・オペレーション", "サービス内容、集客戦略、運営体制"],
    ["04", "財務計画・リスク・成長ロードマップ", "投資計画、収支モデル、リスク対策、成長戦略"],
  ];
  agendaItems.forEach(([num, title, sub], i) => {
    const y = 1.3 + i * 0.9;
    addCard(s, 1.2, y, 7.6, 0.75, { fill: C.light, leftColor: C.accent });
    s.addText(num, { x: 1.4, y: y + 0.05, w: 0.6, h: 0.65, fontSize: 22, fontFace: FONT_EN, color: "B3D8FF", bold: true, valign: "middle", margin: 0 });
    s.addText(title, { x: 2.1, y: y + 0.1, w: 5, h: 0.3, fontSize: 13, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
    s.addText(sub, { x: 2.1, y: y + 0.4, w: 5, h: 0.25, fontSize: 9, fontFace: FONT, color: C.grayDark, margin: 0 });
  });

  // ═══════════════════════════════════════════════════════════════
  // SLIDE 03: Section Divider 1
  // ═══════════════════════════════════════════════════════════════
  console.log("Slide 03: Section 1 Divider");
  s = pres.addSlide();
  // Left dark panel
  s.addShape("rect", { x: 0, y: 0, w: 3.5, h: SH, fill: { color: C.dark } });
  s.addShape("ellipse", { x: 2.5, y: -0.5, w: 1.8, h: 1.8, fill: { color: C.accent, transparency: 80 } });
  addIconCircle(s, 0.6, 0.8, 0.45, C.accent, icons.briefcase_white);
  s.addText("Section 01", { x: 0.6, y: 1.5, w: 2.5, h: 0.25, fontSize: 10, fontFace: FONT_EN, color: C.warm, charSpacing: 3, margin: 0 });
  s.addText([
    { text: "事業概要 &\n提供価値", options: { fontSize: 28, fontFace: FONT, color: C.white, bold: true, breakLine: true } },
  ], { x: 0.6, y: 1.9, w: 2.6, h: 1.2, margin: 0 });
  s.addText("AIレンタカーの事業全体像と、お客様に提供する4つの価値をご紹介します。", { x: 0.6, y: 3.2, w: 2.4, h: 0.6, fontSize: 9, fontFace: FONT, color: C.gray, margin: 0 });
  // Right content
  const divItems1 = [
    [icons.carSide_blue, "Business Model", "レンタカー事業（日貸・時間貸）", "EBF5FF"],
    [icons.mapMarker_blue, "Location", "千葉県（首都圏・房総エリア）", C.yellowBg.replace("#","")],
    [icons.robot_blue, "Future Vision", "AI活用（需要予測・価格最適化）", "EBF5FF"],
  ];
  divItems1.forEach(([icon, en, jp, bg], i) => {
    const y = 1.2 + i * 1.2;
    addIconCircle(s, 4.2, y + 0.1, 0.5, bg, icon);
    s.addText(en, { x: 4.9, y: y, w: 4, h: 0.2, fontSize: 8, fontFace: FONT_EN, color: C.gray, charSpacing: 2, margin: 0 });
    s.addText(jp, { x: 4.9, y: y + 0.25, w: 4.5, h: 0.3, fontSize: 14, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
  });

  // ═══════════════════════════════════════════════════════════════
  // SLIDE 04: Business Overview
  // ═══════════════════════════════════════════════════════════════
  console.log("Slide 04: Business Overview");
  s = pres.addSlide();
  addDecoCircle(s, 7, 3, 2.5, 2.5);
  addHeader(s, "BUSINESS OVERVIEW", "事業概要と提供価値");
  addFooter(s, 4);
  // Left column header
  s.addShape("ellipse", { x: MX, y: 1.25, w: 0.6, h: 0.22, fill: { color: C.accent } });
  s.addText("Overview", { x: MX, y: 1.25, w: 0.6, h: 0.22, fontSize: 7, fontFace: FONT_EN, color: C.white, align: "center", valign: "middle", bold: true, margin: 0 });
  s.addText("事業概要", { x: MX + 0.7, y: 1.22, w: 2, h: 0.28, fontSize: 13, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
  const overviewItems = [
    ["事業名", "AIレンタカー"],
    ["事業形態", "個人事業主（副業スタート）"],
    ["事業内容", "普通自動車のレンタカー事業（日貸・時間貸）"],
    ["開業予定地", "千葉県（首都圏・房総エリアへのアクセス拠点）"],
  ];
  overviewItems.forEach(([label, val], i) => {
    const y = 1.65 + i * 0.65;
    addCard(s, MX, y, 4.0, 0.55, { fill: C.light, leftColor: C.accent });
    s.addText(label, { x: MX + 0.15, y: y + 0.05, w: 1.5, h: 0.18, fontSize: 7, fontFace: FONT_EN, color: C.gray, margin: 0 });
    s.addText(val, { x: MX + 0.15, y: y + 0.23, w: 3.7, h: 0.25, fontSize: 10, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
  });
  // Right column header
  s.addShape("ellipse", { x: 5.2, y: 1.25, w: 0.5, h: 0.22, fill: { color: C.warm } });
  s.addText("Value", { x: 5.2, y: 1.25, w: 0.5, h: 0.22, fontSize: 7, fontFace: FONT_EN, color: C.white, align: "center", valign: "middle", bold: true, margin: 0 });
  s.addText("提供価値", { x: 5.8, y: 1.22, w: 2, h: 0.28, fontSize: 13, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
  const valueItems = [
    [icons.umbrellaBeach_warm, "レジャー・週末利用に最適化", "観光・アウトドアに特化した車両ラインナップ"],
    [icons.star_warm, "清潔・高品質な車両管理", "徹底した清掃と品質管理で安心のサービス"],
    [icons.handshake_warm, "柔軟な受渡・対応", "個人運営ならではの柔軟なサービス提供"],
    [icons.robot_blue, "AI活用による高稼働運営", "需要予測・価格最適化・予約最適化"],
  ];
  valueItems.forEach(([icon, title, sub], i) => {
    const y = 1.65 + i * 0.65;
    addCard(s, 5.2, y, 4.1, 0.55, { fill: C.light, leftColor: C.warm });
    s.addImage({ data: icon, x: 5.4, y: y + 0.12, w: 0.3, h: 0.3 });
    s.addText(title, { x: 5.8, y: y + 0.05, w: 3.2, h: 0.22, fontSize: 10, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
    s.addText(sub, { x: 5.8, y: y + 0.28, w: 3.2, h: 0.2, fontSize: 8, fontFace: FONT, color: C.grayDark, margin: 0 });
  });

  // ═══════════════════════════════════════════════════════════════
  // SLIDE 05: Value Proposition
  // ═══════════════════════════════════════════════════════════════
  console.log("Slide 05: Value Proposition");
  s = pres.addSlide();
  addDecoCircle(s, -0.5, -1, 3, 3);
  addHeader(s, "VALUE PROPOSITION", "4つの提供価値");
  addFooter(s, 5);
  const vpCards = [
    [C.accent, icons.umbrellaBeach_blue, "EBF5FF", "レジャー最適化", "観光・アウトドアに特化した車両ラインナップ", "週末・連休に強い車種構成", C.accent],
    [C.warm, icons.star_warm, C.yellowBg.replace("#",""), "高品質管理", "徹底した清掃と品質管理で安心・快適なドライブ体験", "徹底した清掃・点検体制", C.warm],
    [C.green, icons.handshake_green, C.greenBg.replace("#",""), "柔軟な運営", "お客様一人ひとりのニーズに迅速に対応", "お客様のニーズに即対応", C.green],
  ];
  vpCards.forEach(([topC, icon, iconBg, title, desc, tag, tagC], i) => {
    const x = MX + i * 2.95;
    addCard(s, x, 1.25, 2.7, 2.0, { topColor: topC, borderColor: C.grayLight });
    addIconCircle(s, x + 0.9, 1.45, 0.45, iconBg, icon);
    s.addText(title, { x: x + 0.15, y: 2.0, w: 2.4, h: 0.25, fontSize: 12, fontFace: FONT, color: C.dark, bold: true, align: "center", margin: 0 });
    s.addText(desc, { x: x + 0.15, y: 2.3, w: 2.4, h: 0.45, fontSize: 8, fontFace: FONT, color: C.grayDark, align: "center", margin: 0 });
    s.addText(tag, { x: x + 0.15, y: 2.8, w: 2.4, h: 0.2, fontSize: 7, fontFace: FONT, color: tagC, bold: true, align: "center", margin: 0 });
  });
  // Dark AI bar
  s.addShape("rect", { x: MX, y: 3.5, w: SW - MX * 2, h: 1.3, fill: { color: C.dark }, shadow: shadow() });
  addIconCircle(s, MX + 0.3, 3.75, 0.55, C.accent, icons.robot_white);
  s.addText("将来のAI活用による高稼働運営", { x: MX + 1.1, y: 3.7, w: 5, h: 0.3, fontSize: 14, fontFace: FONT, color: C.white, bold: true, margin: 0 });
  s.addText("需要予測・価格最適化・予約最適化により、車両稼働率を最大化し、収益性の高い運営を実現します。", { x: MX + 1.1, y: 4.05, w: 5, h: 0.35, fontSize: 8, fontFace: FONT, color: C.gray, margin: 0 });
  const aiIcons = [
    [icons.chartLine_blue, "需要予測", 7.2],
    [icons.tags_blue, "価格最適化", 8.0],
    [icons.calendarCheck_blue, "予約最適化", 8.8],
  ];
  aiIcons.forEach(([icon, label, x]) => {
    addIconCircle(s, x, 3.75, 0.4, "FFFFFF1A".substring(0, 6), icon);
    s.addText(label, { x: x - 0.15, y: 4.2, w: 0.7, h: 0.2, fontSize: 7, fontFace: FONT, color: C.gray, align: "center", margin: 0 });
  });

  // ═══════════════════════════════════════════════════════════════
  // SLIDE 06: Section Divider 2
  // ═══════════════════════════════════════════════════════════════
  console.log("Slide 06: Section 2 Divider");
  s = pres.addSlide();
  s.addShape("rect", { x: 0, y: 0, w: 3.5, h: SH, fill: { color: C.dark } });
  s.addShape("ellipse", { x: 2.5, y: -0.5, w: 1.8, h: 1.8, fill: { color: C.accent, transparency: 80 } });
  addIconCircle(s, 0.6, 0.8, 0.45, C.accent, icons.chartPie_white);
  s.addText("Section 02", { x: 0.6, y: 1.5, w: 2.5, h: 0.25, fontSize: 10, fontFace: FONT_EN, color: C.warm, charSpacing: 3, margin: 0 });
  s.addText("市場環境 &\nターゲット", { x: 0.6, y: 1.9, w: 2.6, h: 1.2, fontSize: 28, fontFace: FONT, color: C.white, bold: true, margin: 0 });
  s.addText("市場動向・競合環境の分析と、ターゲット顧客像を明確にします。", { x: 0.6, y: 3.2, w: 2.4, h: 0.6, fontSize: 9, fontFace: FONT, color: C.gray, margin: 0 });
  const divItems2 = [
    [icons.users_blue, "Target Customer", "首都圏在住の個人・ファミリー", "EBF5FF"],
    [icons.mountain_warm, "Market Trend", "観光回復・アウトドア需要増加", C.yellowBg.replace("#","")],
    [icons.trophy_blue, "Differentiation", "用途特化・品質重視・柔軟運営", "EBF5FF"],
  ];
  divItems2.forEach(([icon, en, jp, bg], i) => {
    const y = 1.2 + i * 1.2;
    addIconCircle(s, 4.2, y + 0.1, 0.5, bg, icon);
    s.addText(en, { x: 4.9, y: y, w: 4, h: 0.2, fontSize: 8, fontFace: FONT_EN, color: C.gray, charSpacing: 2, margin: 0 });
    s.addText(jp, { x: 4.9, y: y + 0.25, w: 4.5, h: 0.3, fontSize: 14, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
  });

  // ═══════════════════════════════════════════════════════════════
  // SLIDE 07: Target Customer
  // ═══════════════════════════════════════════════════════════════
  console.log("Slide 07: Target Customer");
  s = pres.addSlide();
  addDecoCircle(s, 7, -1.5, 4, 4);
  addHeader(s, "TARGET CUSTOMER", "ターゲット顧客");
  addFooter(s, 7);
  const tgtCards = [
    [icons.home_blue, "EBF5FF", "首都圏在住の個人", "東京・千葉・埼玉・神奈川"],
    [icons.child_warm, C.yellowBg.replace("#",""), "ファミリー層", "子育て世帯の週末・連休利用"],
    [icons.campground_green, C.greenBg.replace("#",""), "アウトドア利用者", "週末・連休の観光・キャンプ"],
    [icons.car_purple, C.purpleBg.replace("#",""), "マイカー非保有層", "都市部で車を持たない個人"],
  ];
  tgtCards.forEach(([icon, bg, title, sub], i) => {
    const x = MX + i * 2.2;
    addCard(s, x, 1.25, 2.0, 1.6, { borderColor: C.grayLight });
    addIconCircle(s, x + 0.65, 1.4, 0.45, bg, icon);
    s.addText(title, { x: x + 0.1, y: 1.95, w: 1.8, h: 0.25, fontSize: 10, fontFace: FONT, color: C.dark, bold: true, align: "center", margin: 0 });
    s.addText(sub, { x: x + 0.1, y: 2.2, w: 1.8, h: 0.35, fontSize: 7, fontFace: FONT, color: C.grayDark, align: "center", margin: 0 });
  });
  // Bottom left: Market
  addCard(s, MX, 3.1, 4.1, 1.8, { fill: C.light });
  s.addImage({ data: icons.chartLine_blue, x: MX + 0.15, y: 3.22, w: 0.2, h: 0.2 });
  s.addText("市場環境", { x: MX + 0.4, y: 3.2, w: 2, h: 0.25, fontSize: 11, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
  const mktItems = [
    ["観光回復トレンド", "国内観光市場の回復に伴い、レンタカー需要が安定成長"],
    ["アウトドア需要増加", "キャンプ・アウトドアブームの定着で車両ニーズが拡大"],
  ];
  mktItems.forEach(([t, d], i) => {
    const y = 3.6 + i * 0.55;
    addIconCircle(s, MX + 0.15, y, 0.25, C.accent, icons.arrowUp_white);
    s.addText(t, { x: MX + 0.5, y: y - 0.02, w: 3, h: 0.2, fontSize: 9, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
    s.addText(d, { x: MX + 0.5, y: y + 0.18, w: 3.3, h: 0.2, fontSize: 7, fontFace: FONT, color: C.grayDark, margin: 0 });
  });
  // Bottom right: Future
  s.addShape("rect", { x: 5.2, y: 3.1, w: 4.1, h: 1.8, fill: { color: C.dark }, shadow: shadow() });
  s.addText("FUTURE TARGET", { x: 5.4, y: 3.25, w: 3, h: 0.18, fontSize: 8, fontFace: FONT_EN, color: C.gray, charSpacing: 2, margin: 0 });
  s.addText("将来のターゲット拡大", { x: 5.4, y: 3.5, w: 3, h: 0.3, fontSize: 13, fontFace: FONT, color: C.white, bold: true, margin: 0 });
  addIconCircle(s, 5.5, 4.0, 0.3, "FFFFFF", icons.building_blue);
  s.addText("法人利用・長期レンタル需要", { x: 5.9, y: 4.02, w: 3, h: 0.25, fontSize: 9, fontFace: FONT, color: C.gray, margin: 0 });
  addIconCircle(s, 5.5, 4.4, 0.3, "FFFFFF", icons.car_blue);
  s.addText("キャンピングカー特化市場", { x: 5.9, y: 4.42, w: 3, h: 0.25, fontSize: 9, fontFace: FONT, color: C.gray, margin: 0 });

  // ═══════════════════════════════════════════════════════════════
  // SLIDE 08: Competitive Analysis
  // ═══════════════════════════════════════════════════════════════
  console.log("Slide 08: Competitive Analysis");
  s = pres.addSlide();
  addDecoCircle(s, -1, 3, 5, 5);
  addHeader(s, "COMPETITIVE ANALYSIS", "競合比較");
  addFooter(s, 8);
  // Grid table
  const colW = [1.6, 2.2, 2.2, 2.6];
  const colX = [MX, MX + colW[0], MX + colW[0] + colW[1], MX + colW[0] + colW[1] + colW[2]];
  const tableW = colW.reduce((a, b) => a + b, 0);
  // Header row
  s.addShape("rect", { x: MX, y: 1.2, w: tableW, h: 0.35, fill: { color: C.accent } });
  ["比較項目", "大手レンタカー", "地場中小", "AIレンタカー"].forEach((t, i) => {
    const bgC = i === 3 ? C.warm : C.accent;
    if (i === 3) s.addShape("rect", { x: colX[i], y: 1.2, w: colW[i], h: 0.35, fill: { color: C.warm } });
    s.addText(t, { x: colX[i] + 0.1, y: 1.2, w: colW[i] - 0.2, h: 0.35, fontSize: 9, fontFace: FONT, color: C.white, bold: true, valign: "middle", margin: 0 });
  });
  const rows = [
    ["価格帯", "高価格（画一的）", "低〜中価格", "適正価格（用途最適化）"],
    ["車両品質", "標準的", "ばらつきあり", "清潔・高品質管理"],
    ["サービス", "画一的な対応", "個人差あり", "柔軟な個別対応"],
    ["用途特化", "汎用的", "限定的", "レジャー・週末特化"],
    ["受渡し", "店舗固定", "店舗固定", "柔軟な受渡対応"],
  ];
  rows.forEach((row, ri) => {
    const y = 1.55 + ri * 0.38;
    const bg = ri % 2 === 0 ? C.white : C.light;
    s.addShape("rect", { x: MX, y, w: tableW, h: 0.38, fill: { color: bg } });
    s.addShape("rect", { x: colX[3], y, w: colW[3], h: 0.38, fill: { color: "EBF5FF" } });
    row.forEach((cell, ci) => {
      const opts = { x: colX[ci] + 0.1, y, w: colW[ci] - 0.2, h: 0.38, fontSize: 8, fontFace: FONT, valign: "middle", margin: 0 };
      opts.color = ci === 0 ? C.dark : ci === 3 ? C.accent : C.grayDark;
      opts.bold = ci === 0 || ci === 3;
      s.addText(cell, opts);
    });
  });
  // Bottom dark insight
  s.addShape("rect", { x: MX, y: 3.55, w: SW - MX * 2, h: 0.9, fill: { color: C.dark }, shadow: shadow() });
  addIconCircle(s, MX + 0.25, 3.7, 0.45, C.accent, icons.lightbulb_white);
  s.addText("ポジショニング", { x: MX + 0.9, y: 3.6, w: 3, h: 0.25, fontSize: 11, fontFace: FONT, color: C.white, bold: true, margin: 0 });
  s.addText("大手の画一性と地場中小の品質ばらつきの間に「用途特化 × 高品質 × 柔軟運営」のポジションを確立。", { x: MX + 0.9, y: 3.9, w: 7, h: 0.3, fontSize: 8, fontFace: FONT, color: C.gray, margin: 0 });

  // ═══════════════════════════════════════════════════════════════
  // SLIDE 09: Differentiation
  // ═══════════════════════════════════════════════════════════════
  console.log("Slide 09: Differentiation");
  s = pres.addSlide();
  addDecoCircle(s, 7, -1, 5, 5);
  addHeader(s, "DIFFERENTIATION", "差別化ポイント");
  addFooter(s, 9);
  const diffCards = [
    [C.accent, icons.umbrellaBeach_blue, "EBF5FF", "用途特化戦略", "レジャー・週末利用に完全特化。観光・アウトドアに最適な車両ラインナップとプランを設計。", "週末・連休プラン最適化", C.accent],
    [C.warm, icons.star_warm, C.yellowBg.replace("#",""), "品質訴求", "徹底した清掃・点検体制で「選ばれる品質」を実現。口コミ評価を向上。", "徹底した品質管理体制", C.warm],
    [C.green, icons.handshake_green, C.greenBg.replace("#",""), "柔軟な個人運営", "受渡場所・時間の調整、急な変更への対応など、大手にはない柔軟なサービス。", "顧客ニーズに即座に対応", C.green],
    [C.purple, icons.robot_purple, C.purpleBg.replace("#",""), "AI活用（将来構想）", "需要予測・価格最適化・予約最適化をAIで実現し、車両稼働率を最大化。", "データドリブン経営", C.purple],
  ];
  diffCards.forEach(([topC, icon, iconBg, title, desc, tag, tagC], i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = MX + col * 4.4, y = 1.2 + row * 1.95;
    addCard(s, x, y, 4.15, 1.75, { topColor: topC, borderColor: C.grayLight });
    addIconCircle(s, x + 0.15, y + 0.2, 0.4, iconBg, icon);
    s.addText(title, { x: x + 0.65, y: y + 0.2, w: 3, h: 0.25, fontSize: 12, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
    s.addText(desc, { x: x + 0.15, y: y + 0.65, w: 3.8, h: 0.6, fontSize: 8, fontFace: FONT, color: C.grayDark, margin: 0 });
    s.addText(tag, { x: x + 0.15, y: y + 1.35, w: 3, h: 0.2, fontSize: 8, fontFace: FONT, color: tagC, bold: true, margin: 0 });
  });

  // ═══════════════════════════════════════════════════════════════
  // SLIDE 10: Section Divider 3
  // ═══════════════════════════════════════════════════════════════
  console.log("Slide 10: Section 3 Divider");
  s = pres.addSlide();
  s.addShape("rect", { x: 0, y: 0, w: 3.5, h: SH, fill: { color: C.dark } });
  s.addShape("ellipse", { x: 2.5, y: -0.5, w: 1.8, h: 1.8, fill: { color: C.accent, transparency: 80 } });
  addIconCircle(s, 0.6, 0.8, 0.45, C.accent, icons.cogs_white);
  s.addText("Section 03", { x: 0.6, y: 1.5, w: 2.5, h: 0.25, fontSize: 10, fontFace: FONT_EN, color: C.warm, charSpacing: 3, margin: 0 });
  s.addText("サービス &\n戦略", { x: 0.6, y: 1.9, w: 2.6, h: 1.2, fontSize: 28, fontFace: FONT, color: C.white, bold: true, margin: 0 });
  s.addText("サービスの詳細設計、集客チャネル、オペレーション体制を解説します。", { x: 0.6, y: 3.2, w: 2.4, h: 0.6, fontSize: 9, fontFace: FONT, color: C.gray, margin: 0 });
  const divItems3 = [
    [icons.conciergeBell_blue, "Service Design", "レンタカー（日貸・時間貸）+ オプション", "EBF5FF"],
    [icons.bullhorn_warm, "Sales Channel", "Web + 比較サイト + Google + SNS", C.yellowBg.replace("#","")],
    [icons.userCog_blue, "Operations", "代表者1名による効率運営", "EBF5FF"],
  ];
  divItems3.forEach(([icon, en, jp, bg], i) => {
    const y = 1.2 + i * 1.2;
    addIconCircle(s, 4.2, y + 0.1, 0.5, bg, icon);
    s.addText(en, { x: 4.9, y: y, w: 4, h: 0.2, fontSize: 8, fontFace: FONT_EN, color: C.gray, charSpacing: 2, margin: 0 });
    s.addText(jp, { x: 4.9, y: y + 0.25, w: 4.5, h: 0.3, fontSize: 14, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
  });

  // ═══════════════════════════════════════════════════════════════
  // SLIDE 11: Service Details
  // ═══════════════════════════════════════════════════════════════
  console.log("Slide 11: Service Details");
  s = pres.addSlide();
  addDecoCircle(s, 7, -2, 6, 6);
  addHeader(s, "SERVICE DETAILS", "サービス内容");
  addFooter(s, 11);
  const svcItems = [
    ["01", C.accent, "レンタカー（日貸・時間貸）", "普通自動車を日単位・時間単位でレンタル。個人・ファミリーの多様なニーズに対応します。"],
    ["02", C.accent, "週末・連休向けプラン", "週末やGW・お盆・年末年始など繁忙期に合わせた特別プランを用意。"],
    ["03", C.warm, "オプションサービス", "チャイルドシート、ETCカード、アウトドア用品など、旅をより快適にするオプション。"],
    ["04", C.warm, "将来構想", "キャンピングカーの導入、法人契約、長期レンタルへの展開を計画。"],
  ];
  svcItems.forEach(([num, color, title, desc], i) => {
    const y = 1.2 + i * 0.9;
    addCard(s, 1.2, y, 7.6, 0.75, { borderColor: C.grayLight });
    s.addShape("ellipse", { x: 1.35, y: y + 0.15, w: 0.4, h: 0.4, fill: { color } });
    s.addText(num, { x: 1.35, y: y + 0.15, w: 0.4, h: 0.4, fontSize: 10, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(title, { x: 1.95, y: y + 0.1, w: 6, h: 0.25, fontSize: 12, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
    s.addText(desc, { x: 1.95, y: y + 0.38, w: 6.5, h: 0.25, fontSize: 8, fontFace: FONT, color: C.grayDark, margin: 0 });
  });

  // ═══════════════════════════════════════════════════════════════
  // SLIDE 12: Sales Channels
  // ═══════════════════════════════════════════════════════════════
  console.log("Slide 12: Sales Channels");
  s = pres.addSlide();
  addDecoCircle(s, 7, -2, 6, 6);
  addHeader(s, "SALES CHANNELS", "集客・販売チャネル");
  addFooter(s, 12);
  const channels = [
    [icons.globe_blue, "Step 1", "自社Webサイト", "予約・問い合わせの窓口。\nSEO対策で検索流入を獲得"],
    [icons.search_blue, "Step 2", "比較サイト", "レンタカー比較サイトへの掲載で新規顧客にリーチ"],
    [icons.mapMarkedAlt_blue, "Step 3", "Googleマップ", "口コミ・レビュー獲得で地域検索からの集客を強化"],
    [icons.instagram_blue, "Step 4", "SNS (Instagram)", "レジャー・アウトドア写真で認知拡大とファン獲得"],
  ];
  channels.forEach(([icon, step, title, desc], i) => {
    const x = MX + i * 2.25;
    addCard(s, x, 1.3, 1.95, 2.8, { topColor: C.accent, borderColor: C.grayLight });
    addIconCircle(s, x + 0.6, 1.55, 0.45, "EBF5FF", icon);
    s.addShape("ellipse", { x: x + 0.5, y: 2.15, w: 0.7, h: 0.22, fill: { color: "EBF5FF" } });
    s.addText(step, { x: x + 0.5, y: 2.15, w: 0.7, h: 0.22, fontSize: 7, fontFace: FONT_EN, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(title, { x: x + 0.1, y: 2.5, w: 1.75, h: 0.25, fontSize: 11, fontFace: FONT, color: C.dark, bold: true, align: "center", margin: 0 });
    s.addText(desc, { x: x + 0.1, y: 2.85, w: 1.75, h: 0.8, fontSize: 7, fontFace: FONT, color: C.grayDark, align: "center", margin: 0 });
    if (i < 3) {
      s.addImage({ data: icons.chevronRight_gray, x: x + 2.0, y: 2.2, w: 0.2, h: 0.2 });
    }
  });

  // ═══════════════════════════════════════════════════════════════
  // SLIDE 13: Operations
  // ═══════════════════════════════════════════════════════════════
  console.log("Slide 13: Operations");
  s = pres.addSlide();
  addDecoCircle(s, 7, -2, 6, 6);
  addHeader(s, "OPERATIONS", "オペレーション体制");
  addFooter(s, 13);
  // Layer 1: Core
  s.addShape("rect", { x: MX, y: 1.2, w: SW - MX * 2, h: 1.15, fill: { color: "EBF5FF" }, line: { color: C.accent, width: 1.5 } });
  s.addShape("ellipse", { x: MX + 0.15, y: 1.3, w: 0.45, h: 0.2, fill: { color: C.accent } });
  s.addText("Core", { x: MX + 0.15, y: 1.3, w: 0.45, h: 0.2, fontSize: 7, fontFace: FONT_EN, color: C.white, align: "center", valign: "middle", bold: true, margin: 0 });
  s.addText("経営・管理", { x: MX + 0.7, y: 1.28, w: 2, h: 0.25, fontSize: 10, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
  const coreItems = [
    [icons.userTie_blue, "代表者", "受付・車両管理・事務を兼任"],
    [icons.clipboardCheck_blue, "品質管理", "清掃・点検の徹底"],
    [icons.chartBar_blue, "経営分析", "稼働率・収益のモニタリング"],
  ];
  coreItems.forEach(([icon, t, d], i) => {
    const x = MX + 0.2 + i * 2.85;
    s.addImage({ data: icon, x, y: 1.65, w: 0.25, h: 0.25 });
    s.addText(t, { x: x + 0.35, y: 1.63, w: 2, h: 0.18, fontSize: 9, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
    s.addText(d, { x: x + 0.35, y: 1.82, w: 2.2, h: 0.18, fontSize: 7, fontFace: FONT, color: C.grayDark, margin: 0 });
  });
  // Arrow
  s.addImage({ data: icons.chevronDown_gray, x: 4.85, y: 2.4, w: 0.2, h: 0.2 });
  // Layer 2: Daily
  s.addShape("rect", { x: MX, y: 2.7, w: SW - MX * 2, h: 1.05, fill: { color: C.light }, line: { color: C.grayLight, width: 0.75 } });
  s.addShape("ellipse", { x: MX + 0.15, y: 2.8, w: 0.45, h: 0.2, fill: { color: C.grayDark } });
  s.addText("Daily", { x: MX + 0.15, y: 2.8, w: 0.45, h: 0.2, fontSize: 7, fontFace: FONT_EN, color: C.white, align: "center", valign: "middle", bold: true, margin: 0 });
  s.addText("日常オペレーション", { x: MX + 0.7, y: 2.78, w: 2.5, h: 0.25, fontSize: 10, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
  const dailyItems = [
    [icons.car_grayDark, "車両受渡", "お客様への車両引渡・返却"],
    [icons.broom_grayDark, "車両清掃", "返却後の清掃・消毒作業"],
    [icons.tools_grayDark, "車両点検", "定期的な整備・安全確認"],
  ];
  dailyItems.forEach(([icon, t, d], i) => {
    const x = MX + 0.2 + i * 2.85;
    s.addImage({ data: icon, x, y: 3.15, w: 0.25, h: 0.25 });
    s.addText(t, { x: x + 0.35, y: 3.13, w: 2, h: 0.18, fontSize: 9, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
    s.addText(d, { x: x + 0.35, y: 3.32, w: 2.2, h: 0.18, fontSize: 7, fontFace: FONT, color: C.grayDark, margin: 0 });
  });
  // Arrow
  s.addImage({ data: icons.chevronDown_gray, x: 4.85, y: 3.8, w: 0.2, h: 0.2 });
  // Layer 3: External
  s.addShape("rect", { x: MX, y: 4.1, w: SW - MX * 2, h: 0.95, fill: { color: C.light }, line: { color: C.grayLight, width: 0.75 } });
  s.addShape("ellipse", { x: MX + 0.15, y: 4.2, w: 0.6, h: 0.2, fill: { color: C.gray } });
  s.addText("External", { x: MX + 0.15, y: 4.2, w: 0.6, h: 0.2, fontSize: 7, fontFace: FONT_EN, color: C.white, align: "center", valign: "middle", bold: true, margin: 0 });
  s.addText("外部委託（繁忙期）", { x: MX + 0.85, y: 4.18, w: 2.5, h: 0.25, fontSize: 10, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
  const extItems = [
    [icons.handsHelping_grayDark, "清掃外注", "繁忙期の清掃・点検業務を外注"],
    [icons.wrench_grayDark, "整備委託", "専門的な整備・修理は外部委託"],
  ];
  extItems.forEach(([icon, t, d], i) => {
    const x = MX + 0.2 + i * 4.0;
    s.addImage({ data: icon, x, y: 4.55, w: 0.25, h: 0.25 });
    s.addText(t, { x: x + 0.35, y: 4.53, w: 2, h: 0.18, fontSize: 9, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
    s.addText(d, { x: x + 0.35, y: 4.72, w: 3, h: 0.18, fontSize: 7, fontFace: FONT, color: C.grayDark, margin: 0 });
  });

  // ═══════════════════════════════════════════════════════════════
  // SLIDE 14: Section Divider 4
  // ═══════════════════════════════════════════════════════════════
  console.log("Slide 14: Section 4 Divider");
  s = pres.addSlide();
  s.addShape("rect", { x: 0, y: 0, w: 3.5, h: SH, fill: { color: C.dark } });
  s.addShape("ellipse", { x: 2.5, y: -0.5, w: 1.8, h: 1.8, fill: { color: C.accent, transparency: 80 } });
  addIconCircle(s, 0.6, 0.8, 0.45, C.accent, icons.yenSign_white);
  s.addText("Section 04", { x: 0.6, y: 1.5, w: 2.5, h: 0.25, fontSize: 10, fontFace: FONT_EN, color: C.warm, charSpacing: 3, margin: 0 });
  s.addText("財務計画・\nリスク・成長", { x: 0.6, y: 1.9, w: 2.6, h: 1.2, fontSize: 28, fontFace: FONT, color: C.white, bold: true, margin: 0 });
  s.addText("投資計画、収支モデル、リスク対策、成長戦略をご説明します。", { x: 0.6, y: 3.2, w: 2.4, h: 0.6, fontSize: 9, fontFace: FONT, color: C.gray, margin: 0 });
  const divItems4 = [
    [icons.car_blue, "Investment", "総額700〜1,000万円", "EBF5FF"],
    [icons.coins_warm, "Revenue", "月商約18万円・月間利益2〜5万円", C.yellowBg.replace("#","")],
    [icons.chartLine_blue, "Growth", "1台→2台→キャンピングカー特化", "EBF5FF"],
  ];
  divItems4.forEach(([icon, en, jp, bg], i) => {
    const y = 1.2 + i * 1.2;
    addIconCircle(s, 4.2, y + 0.1, 0.5, bg, icon);
    s.addText(en, { x: 4.9, y: y, w: 4, h: 0.2, fontSize: 8, fontFace: FONT_EN, color: C.gray, charSpacing: 2, margin: 0 });
    s.addText(jp, { x: 4.9, y: y + 0.25, w: 4.5, h: 0.3, fontSize: 14, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
  });

  // ═══════════════════════════════════════════════════════════════
  // SLIDE 15: Initial Investment
  // ═══════════════════════════════════════════════════════════════
  console.log("Slide 15: Initial Investment");
  s = pres.addSlide();
  addDecoCircle(s, 7, -2, 6, 6);
  addHeader(s, "INITIAL INVESTMENT", "初期投資の内訳");
  addFooter(s, 15);
  const investCards = [
    [C.accent, icons.car_blue, "車両購入", "600-900万円", "中古車1台"],
    [C.warm, icons.shieldAlt_warm, "登録・整備・保険", "50万円", ""],
    [C.green, icons.box_green, "設備・備品", "20万円", ""],
    [C.purple, icons.laptop_purple, "Web・広告", "10万円", ""],
  ];
  investCards.forEach(([color, icon, label, amount, sub], i) => {
    const x = MX + i * 2.2;
    addCard(s, x, 1.2, 2.0, 1.15, { borderColor: C.grayLight });
    s.addShape("rect", { x, y: 1.2, w: 0.04, h: 1.15, fill: { color } });
    s.addText(label, { x: x + 0.15, y: 1.28, w: 1.5, h: 0.18, fontSize: 7, fontFace: FONT, color: C.grayDark, margin: 0 });
    s.addImage({ data: icon, x: x + 1.55, y: 1.28, w: 0.2, h: 0.2 });
    s.addText(amount, { x: x + 0.15, y: 1.55, w: 1.7, h: 0.35, fontSize: 18, fontFace: FONT_EN, color: C.dark, bold: true, margin: 0 });
    if (sub) s.addText(sub, { x: x + 0.15, y: 1.95, w: 1.5, h: 0.18, fontSize: 7, fontFace: FONT, color: C.grayDark, margin: 0 });
  });
  // Bar chart
  addCard(s, MX, 2.55, 4.1, 2.4, { borderColor: C.grayLight });
  s.addText("投資内訳チャート", { x: MX + 0.15, y: 2.65, w: 3, h: 0.25, fontSize: 10, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
  const bars = [
    ["車両", "600-900万円", 0.85, C.accent],
    ["登録等", "50万円", 0.07, C.warm],
    ["設備", "20万円", 0.03, C.green],
    ["広告", "10万円", 0.015, C.purple],
  ];
  bars.forEach(([label, val, pct, color], i) => {
    const y = 3.05 + i * 0.4;
    s.addText(label, { x: MX + 0.15, y: y, w: 0.6, h: 0.18, fontSize: 7, fontFace: FONT, color: C.grayDark, margin: 0 });
    s.addText(val, { x: MX + 2.8, y: y, w: 1, h: 0.18, fontSize: 7, fontFace: FONT_EN, color: C.dark, bold: true, align: "right", margin: 0 });
    s.addShape("rect", { x: MX + 0.15, y: y + 0.2, w: 3.7, h: 0.12, fill: { color: "F3F4F6" } });
    s.addShape("rect", { x: MX + 0.15, y: y + 0.2, w: Math.max(3.7 * pct, 0.12), h: 0.12, fill: { color } });
  });
  // Total summary
  s.addShape("rect", { x: 5.2, y: 2.55, w: 4.1, h: 2.4, fill: { color: C.dark }, shadow: shadow() });
  s.addText("TOTAL INVESTMENT", { x: 5.4, y: 2.75, w: 3, h: 0.18, fontSize: 8, fontFace: FONT_EN, color: C.gray, charSpacing: 2, margin: 0 });
  s.addText("700〜1,000万円", { x: 5.4, y: 3.1, w: 3.5, h: 0.5, fontSize: 28, fontFace: FONT_EN, color: C.white, bold: true, margin: 0 });
  s.addShape("rect", { x: 5.4, y: 3.7, w: 0.8, h: 0.04, fill: { color: C.accent } });
  s.addText("車両購入費が全体の85%以上を占めます。中古車を活用し初期コストを抑制。段階的な投資で事業リスクを最小化する戦略です。", { x: 5.4, y: 3.9, w: 3.5, h: 0.8, fontSize: 8, fontFace: FONT, color: C.gray, margin: 0 });

  // ═══════════════════════════════════════════════════════════════
  // SLIDE 16: Revenue & Funding
  // ═══════════════════════════════════════════════════════════════
  console.log("Slide 16: Revenue & Funding");
  s = pres.addSlide();
  addDecoCircle(s, 7, -1.5, 5, 5);
  addHeader(s, "REVENUE MODEL & FUNDING", "収支モデル・資金調達");
  addFooter(s, 16);
  // Left: Revenue
  s.addImage({ data: icons.calculator_blue, x: MX, y: 1.2, w: 0.2, h: 0.2 });
  s.addText("月次収支モデル（標準）", { x: MX + 0.3, y: 1.18, w: 3, h: 0.25, fontSize: 10, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
  // KPI pair
  const revKpis = [["稼働日数", "12日/月"], ["平均単価", "15,000円/日"]];
  revKpis.forEach(([label, val], i) => {
    const x = MX + i * 2.05;
    addCard(s, x, 1.55, 1.85, 0.7, { borderColor: C.grayLight });
    s.addText(label, { x: x + 0.1, y: 1.6, w: 1.5, h: 0.18, fontSize: 7, fontFace: FONT, color: C.grayDark, margin: 0 });
    s.addText(val, { x: x + 0.1, y: 1.82, w: 1.6, h: 0.35, fontSize: 16, fontFace: FONT_EN, color: C.dark, bold: true, margin: 0 });
  });
  // Breakdown
  addCard(s, MX, 2.4, 4.1, 2.4, { fill: C.light });
  const revRows = [
    [icons.arrowUp_blue, "月商", "約18万円", C.accent, true],
    [null, "固定費", "10〜14万円", C.grayDark, false],
    [null, "変動費", "2万円", C.grayDark, false],
    [icons.coins_warm, "月間利益", "2〜5万円", C.accent, true],
  ];
  revRows.forEach(([icon, label, val, color, isBold], i) => {
    const y = 2.55 + i * 0.5;
    if (i === 3) s.addShape("line", { x: MX + 0.15, y: y - 0.05, w: 3.8, h: 0, line: { color: C.accent, width: 1.5 } });
    else if (i > 0) s.addShape("line", { x: MX + 0.15, y: y - 0.05, w: 3.8, h: 0, line: { color: C.grayLight, width: 0.5 } });
    if (icon) s.addImage({ data: icon, x: MX + 0.15, y: y + 0.05, w: 0.18, h: 0.18 });
    s.addText(label, { x: MX + 0.4, y: y + 0.02, w: 1.5, h: 0.22, fontSize: 9, fontFace: FONT, color: C.dark, bold: isBold, margin: 0 });
    s.addText(val, { x: MX + 2.5, y: y, w: 1.4, h: 0.25, fontSize: i === 3 ? 16 : 12, fontFace: FONT_EN, color, bold: isBold, align: "right", margin: 0 });
  });
  // Right: Funding
  s.addImage({ data: icons.handHoldingUsd_warm, x: 5.2, y: 1.2, w: 0.2, h: 0.2 });
  s.addText("資金調達計画", { x: 5.5, y: 1.18, w: 3, h: 0.25, fontSize: 10, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
  const fundKpis = [["自己資金", "200万円", C.dark], ["借入希望", "800-1,000万円", C.warm]];
  fundKpis.forEach(([label, val, color], i) => {
    const x = 5.2 + i * 2.05;
    addCard(s, x, 1.55, 1.85, 0.7, { borderColor: i === 1 ? C.warm : C.grayLight });
    s.addText(label, { x: x + 0.1, y: 1.6, w: 1.5, h: 0.18, fontSize: 7, fontFace: FONT, color: C.grayDark, margin: 0 });
    s.addText(val, { x: x + 0.1, y: 1.82, w: 1.6, h: 0.35, fontSize: 14, fontFace: FONT_EN, color, bold: true, margin: 0 });
  });
  // Funding bar
  addCard(s, 5.2, 2.4, 4.1, 0.8, { fill: C.light });
  s.addText("資金構成", { x: 5.35, y: 2.48, w: 2, h: 0.18, fontSize: 8, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
  s.addShape("rect", { x: 5.35, y: 2.72, w: 0.74, h: 0.2, fill: { color: C.accent } });
  s.addText("自己資金", { x: 5.35, y: 2.72, w: 0.74, h: 0.2, fontSize: 6, fontFace: FONT, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addShape("rect", { x: 6.09, y: 2.72, w: 3.06, h: 0.2, fill: { color: C.warm } });
  s.addText("借入", { x: 6.09, y: 2.72, w: 3.06, h: 0.2, fontSize: 6, fontFace: FONT, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("200万円（20%）", { x: 5.35, y: 2.96, w: 1.5, h: 0.15, fontSize: 6, fontFace: FONT, color: C.grayDark, margin: 0 });
  s.addText("800〜1,000万円（80%）", { x: 7.5, y: 2.96, w: 1.8, h: 0.15, fontSize: 6, fontFace: FONT, color: C.grayDark, align: "right", margin: 0 });
  // Bottom dark box
  s.addShape("rect", { x: 5.2, y: 3.35, w: 4.1, h: 1.45, fill: { color: C.dark }, shadow: shadow() });
  s.addText("FUNDING PURPOSE", { x: 5.4, y: 3.5, w: 3, h: 0.18, fontSize: 8, fontFace: FONT_EN, color: C.gray, charSpacing: 2, margin: 0 });
  s.addText("借入用途", { x: 5.4, y: 3.75, w: 2, h: 0.25, fontSize: 12, fontFace: FONT, color: C.white, bold: true, margin: 0 });
  addIconCircle(s, 5.5, 4.15, 0.25, C.accent, icons.car_white);
  s.addText("車両購入費", { x: 5.85, y: 4.17, w: 2, h: 0.2, fontSize: 9, fontFace: FONT, color: C.gray, margin: 0 });
  addIconCircle(s, 5.5, 4.5, 0.25, C.warm, icons.wallet_white);
  s.addText("運転資金", { x: 5.85, y: 4.52, w: 2, h: 0.2, fontSize: 9, fontFace: FONT, color: C.gray, margin: 0 });

  // ═══════════════════════════════════════════════════════════════
  // SLIDE 17: Risk Management
  // ═══════════════════════════════════════════════════════════════
  console.log("Slide 17: Risk Management");
  s = pres.addSlide();
  addDecoCircle(s, 7, 3, 5, 5);
  addHeader(s, "RISK MANAGEMENT", "リスクと対策");
  addFooter(s, 17);
  const risks = [
    [C.yellowBg.replace("#",""), icons.chartBar_blue, "Risk 01", "稼働率低下", "繁忙期と閑散期の需要差が大きく、平均稼働率が想定を下回るリスク", "副業スタートで固定費を抑制。本業収入でリスクヘッジ"],
    [C.redBg.replace("#",""), icons.carCrash_red, "Risk 02", "事故・故障", "車両の事故や故障により営業停止、修理費用が発生するリスク", "任意保険加入・定期点検の徹底で予防と補償を確保"],
    ["DBEAFE", icons.tags_blue, "Risk 03", "価格競争", "大手や新規参入者との価格競争で収益性が圧迫されるリスク", "レジャー特化・品質訴求で価格以外の価値で差別化"],
    [C.purpleBg.replace("#",""), icons.wallet_purple, "Risk 04", "資金繰り", "借入返済と運営費のバランスが崩れ、資金繰りが悪化するリスク", "慎重な台数拡大。1台で安定運営確認してから次の投資へ"],
  ];
  risks.forEach(([bg, icon, riskNum, title, risk, solution], i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = MX + col * 4.4, y = 1.2 + row * 1.95;
    addCard(s, x, y, 4.15, 1.75, { borderColor: C.grayLight });
    addIconCircle(s, x + 0.15, y + 0.2, 0.35, bg, icon);
    s.addText(riskNum, { x: x + 0.6, y: y + 0.15, w: 1.5, h: 0.15, fontSize: 7, fontFace: FONT_EN, color: C.gray, charSpacing: 1, margin: 0 });
    s.addText(title, { x: x + 0.6, y: y + 0.3, w: 3, h: 0.22, fontSize: 11, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
    // Risk description
    s.addShape("rect", { x: x + 0.15, y: y + 0.65, w: 0.03, h: 0.35, fill: { color: C.red } });
    s.addText(risk, { x: x + 0.3, y: y + 0.65, w: 3.6, h: 0.35, fontSize: 7, fontFace: FONT, color: C.grayDark, margin: 0 });
    // Solution
    s.addShape("rect", { x: x + 0.15, y: y + 1.15, w: 0.03, h: 0.35, fill: { color: C.green } });
    s.addImage({ data: icons.shieldAlt_green, x: x + 0.25, y: y + 1.18, w: 0.15, h: 0.15 });
    s.addText("対策", { x: x + 0.45, y: y + 1.15, w: 0.5, h: 0.18, fontSize: 7, fontFace: FONT, color: C.green, bold: true, margin: 0 });
    s.addText(solution, { x: x + 0.3, y: y + 1.33, w: 3.6, h: 0.25, fontSize: 7, fontFace: FONT, color: C.grayDark, margin: 0 });
  });

  // ═══════════════════════════════════════════════════════════════
  // SLIDE 18: Growth Roadmap
  // ═══════════════════════════════════════════════════════════════
  console.log("Slide 18: Growth Roadmap");
  s = pres.addSlide();
  addDecoCircle(s, -1, -2, 5, 5);
  addHeader(s, "GROWTH ROADMAP", "成長ロードマップ");
  addFooter(s, 18);
  const phases = [
    [C.accent, "Y1", "1年目", "基盤構築フェーズ", ["中古車1台で開業", "安定運営体制の確立", "稼働率の向上", "口コミ・リピーター獲得"], "車両 1台 / 稼働率安定化"],
    [C.warm, "Y2-3", "2〜3年目", "拡大フェーズ", ["2台目の車両導入", "車種バリエーション拡充", "集客チャネルの多角化", "清掃・点検の一部外注化"], "車両 2台 / 売上倍増"],
    [C.green, "Y4+", "将来", "発展フェーズ", ["キャンピングカー特化", "AI活用運営の実装", "法人契約・長期レンタル", "エリア拡大"], "AI活用 / キャンピングカー特化"],
  ];
  phases.forEach(([color, badge, title, phase, items, goal], i) => {
    const x = MX + i * 3.0;
    // Timeline dot + line
    s.addShape("ellipse", { x: x + 0.35, y: 1.25, w: 0.45, h: 0.45, fill: { color } });
    s.addText(badge, { x: x + 0.35, y: 1.25, w: 0.45, h: 0.45, fontSize: 9, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    if (i < 2) s.addShape("rect", { x: x + 0.85, y: 1.44, w: 2.2, h: 0.04, fill: { color } });
    // Card
    addCard(s, x, 1.85, 2.75, 3.1, { topColor: color, borderColor: C.grayLight });
    s.addText(title, { x: x + 0.15, y: 1.97, w: 2, h: 0.25, fontSize: 12, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
    s.addText(phase, { x: x + 0.15, y: 2.22, w: 2, h: 0.2, fontSize: 8, fontFace: FONT, color, bold: true, margin: 0 });
    items.forEach((item, j) => {
      s.addImage({ data: color === C.accent ? icons.checkCircle_blue : color === C.warm ? icons.checkCircle_warm : icons.checkCircle_green, x: x + 0.15, y: 2.55 + j * 0.3, w: 0.15, h: 0.15 });
      s.addText(item, { x: x + 0.38, y: 2.53 + j * 0.3, w: 2.2, h: 0.2, fontSize: 8, fontFace: FONT, color: C.grayDark, margin: 0 });
    });
    // Goal
    s.addShape("line", { x: x + 0.15, y: 3.85, w: 2.4, h: 0, line: { color: "F3F4F6", width: 0.5 } });
    s.addText("目標", { x: x + 0.15, y: 3.92, w: 0.5, h: 0.15, fontSize: 7, fontFace: FONT, color: C.grayDark, margin: 0 });
    s.addText(goal, { x: x + 0.15, y: 4.1, w: 2.4, h: 0.35, fontSize: 9, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
  });

  // ═══════════════════════════════════════════════════════════════
  // SLIDE 19: Summary
  // ═══════════════════════════════════════════════════════════════
  console.log("Slide 19: Summary");
  s = pres.addSlide();
  addDecoCircle(s, 7, -1.5, 6, 6);
  addHeader(s, "SUMMARY & NEXT STEPS", "まとめ・Next Steps");
  addFooter(s, 19);
  // Left: Summary
  s.addShape("rect", { x: MX, y: 1.2, w: 4.1, h: 1.6, fill: { color: C.accent }, shadow: shadow() });
  s.addImage({ data: icons.bullseye_white, x: MX + 0.15, y: 1.3, w: 0.2, h: 0.2 });
  s.addText("事業の要約", { x: MX + 0.45, y: 1.28, w: 2, h: 0.25, fontSize: 11, fontFace: FONT, color: C.white, bold: true, margin: 0 });
  const summaryItems = [
    [icons.car_white, "事業モデル", "千葉県拠点のレンタカー事業（個人事業主・副業）"],
    [icons.users_white, "ターゲット", "首都圏ファミリー・マイカー非保有層"],
    [icons.trophy_white, "差別化", "レジャー特化 × 高品質 × 柔軟運営 × AI活用"],
  ];
  summaryItems.forEach(([icon, label, val], i) => {
    const y = 1.62 + i * 0.35;
    s.addImage({ data: icon, x: MX + 0.2, y: y + 0.02, w: 0.18, h: 0.18 });
    s.addText(label, { x: MX + 0.5, y: y - 0.02, w: 0.8, h: 0.14, fontSize: 6, fontFace: FONT, color: "FFFFFF90".substring(0,6), margin: 0 });
    s.addText(val, { x: MX + 0.5, y: y + 0.12, w: 3.2, h: 0.18, fontSize: 8, fontFace: FONT, color: C.white, bold: true, margin: 0 });
  });
  // Numbers
  s.addShape("rect", { x: MX, y: 2.95, w: 4.1, h: 1.3, fill: { color: C.dark }, shadow: shadow() });
  s.addImage({ data: icons.chartLine_blue, x: MX + 0.15, y: 3.05, w: 0.18, h: 0.18 });
  s.addText("数値サマリー", { x: MX + 0.4, y: 3.03, w: 2, h: 0.22, fontSize: 9, fontFace: FONT, color: C.white, bold: true, margin: 0 });
  const numItems = [
    ["初期投資", "700-1,000万円", C.white], ["月間利益", "2-5万円", C.white],
    ["借入希望", "800-1,000万円", C.warm], ["自己資金", "200万円", C.white],
  ];
  numItems.forEach(([label, val, color], i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = MX + 0.2 + col * 2.0, y = 3.35 + row * 0.5;
    s.addText(label, { x, y, w: 1.5, h: 0.15, fontSize: 7, fontFace: FONT, color: C.gray, margin: 0 });
    s.addText(val, { x, y: y + 0.17, w: 1.8, h: 0.25, fontSize: 14, fontFace: FONT_EN, color, bold: true, margin: 0 });
  });
  // Right: Next Steps
  s.addImage({ data: icons.tasks_blue, x: 5.2, y: 1.2, w: 0.2, h: 0.2 });
  s.addText("Next Steps", { x: 5.5, y: 1.18, w: 2, h: 0.25, fontSize: 10, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
  const steps = [
    [C.accent, "1", "車両取得・開業準備", "中古車購入、レンタカー事業許可取得、保険加入"],
    [C.warm, "2", "Webサイト・集客基盤構築", "予約サイト構築、Googleマップ登録、SNS開設"],
    [C.green, "3", "営業開始・稼働率向上", "テスト運営、口コミ蓄積、月間稼働12日以上を目標"],
  ];
  steps.forEach(([color, num, title, desc], i) => {
    const y = 1.55 + i * 0.7;
    addCard(s, 5.2, y, 4.1, 0.6, { fill: C.light, leftColor: color });
    s.addShape("ellipse", { x: 5.4, y: y + 0.1, w: 0.3, h: 0.3, fill: { color } });
    s.addText(num, { x: 5.4, y: y + 0.1, w: 0.3, h: 0.3, fontSize: 10, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(title, { x: 5.85, y: y + 0.08, w: 3, h: 0.2, fontSize: 9, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
    s.addText(desc, { x: 5.85, y: y + 0.3, w: 3.2, h: 0.2, fontSize: 7, fontFace: FONT, color: C.grayDark, margin: 0 });
  });
  // Funding ask
  s.addShape("rect", { x: 5.2, y: 3.8, w: 4.1, h: 1.0, fill: { color: C.yellowBg.replace("#","") }, line: { color: C.warm, width: 1.5 }, shadow: shadow() });
  s.addImage({ data: icons.handHoldingUsd_warm, x: 5.35, y: 3.9, w: 0.2, h: 0.2 });
  s.addText("資金調達のお願い", { x: 5.6, y: 3.88, w: 2, h: 0.25, fontSize: 9, fontFace: FONT, color: C.dark, bold: true, margin: 0 });
  s.addText("調達額", { x: 5.35, y: 4.22, w: 1, h: 0.13, fontSize: 6, fontFace: FONT, color: C.grayDark, margin: 0 });
  s.addText("800-1,000万円", { x: 5.35, y: 4.37, w: 1.8, h: 0.25, fontSize: 14, fontFace: FONT_EN, color: C.accent, bold: true, margin: 0 });
  s.addText("用途", { x: 7.2, y: 4.22, w: 1, h: 0.13, fontSize: 6, fontFace: FONT, color: C.grayDark, margin: 0 });
  s.addText("車両購入＋運転資金", { x: 7.2, y: 4.37, w: 2, h: 0.25, fontSize: 10, fontFace: FONT, color: C.dark, bold: true, margin: 0 });

  // ═══════════════════════════════════════════════════════════════
  // SLIDE 20: Thank You
  // ═══════════════════════════════════════════════════════════════
  console.log("Slide 20: Thank You");
  s = pres.addSlide();
  addDecoCircle(s, 6, -2, 6, 6);
  addDecoCircle(s, -2, 3, 5, 5);
  s.addShape("ellipse", { x: 4.5, y: 1.0, w: 0.7, h: 0.7, fill: { color: C.accent }, shadow: shadow() });
  s.addImage({ data: icons.carSide_white, x: 4.62, y: 1.13, w: 0.45, h: 0.45 });
  s.addText("THANK YOU FOR YOUR TIME", { x: 0, y: 1.95, w: SW, h: 0.2, fontSize: 8, fontFace: FONT_EN, color: C.gray, align: "center", charSpacing: 4, margin: 0 });
  s.addText("ご清聴ありがとうございました", { x: 0, y: 2.3, w: SW, h: 0.6, fontSize: 30, fontFace: FONT, color: C.dark, align: "center", bold: true, margin: 0 });
  s.addShape("rect", { x: 4.5, y: 3.05, w: 1, h: 0.06, fill: { color: C.accent } });
  s.addText("AI RENTAL CAR", { x: 0, y: 3.3, w: SW, h: 0.35, fontSize: 18, fontFace: FONT_EN, color: C.accent, align: "center", bold: true, charSpacing: 3, margin: 0 });
  s.addText("AIレンタカー事業計画書", { x: 0, y: 3.65, w: SW, h: 0.25, fontSize: 11, fontFace: FONT, color: C.grayDark, align: "center", margin: 0 });
  // Contact
  s.addImage({ data: icons.mapMarker_blue, x: 3.3, y: 4.2, w: 0.18, h: 0.18 });
  s.addText("千葉県", { x: 3.55, y: 4.2, w: 1, h: 0.2, fontSize: 9, fontFace: FONT, color: C.grayDark, margin: 0 });
  s.addImage({ data: icons.envelope_blue, x: 5.5, y: 4.2, w: 0.18, h: 0.18 });
  s.addText("info@ai-rentalcar.jp", { x: 5.75, y: 4.2, w: 2, h: 0.2, fontSize: 9, fontFace: FONT_EN, color: C.grayDark, margin: 0 });
  // Footer
  addFooter(s, 20);

  // ═══════════════════════════════════════════════════════════════
  // Save
  // ═══════════════════════════════════════════════════════════════
  const outPath = "/Users/nogataka/dev/slide-test/output/slide-page02/AIレンタカー事業計画書.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log(`\nPPTX saved to: ${outPath}`);
}

main().catch(err => { console.error(err); process.exit(1); });
