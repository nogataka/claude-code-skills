const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const {
  FaCar, FaCalendarAlt, FaMapMarkerAlt, FaBriefcase, FaRobot, FaUsers, FaMountain, FaBuilding,
  FaChartLine, FaChartArea, FaConciergeBell, FaBullhorn, FaGlobe, FaSearch, FaMapMarkedAlt,
  FaInstagram, FaCogs, FaUserCog, FaYenSign, FaCalculator, FaPhoneAlt, FaBroom, FaFileAlt,
  FaHandsHelping, FaUserTie, FaTag, FaCalendarCheck, FaChartBar, FaMinusCircle, FaWallet,
  FaUniversity, FaPiggyBank, FaShieldAlt, FaRoad, FaRocket, FaExclamationTriangle, FaTags,
  FaCoins, FaSeedling, FaExpandArrowsAlt, FaEnvelope, FaCheckCircle, FaTimesCircle, FaStar,
  FaFileContract, FaToolbox, FaAd, FaUmbrellaBeach, FaHandshake, FaFlag, FaArrowUp,
  FaChevronRight, FaChevronDown, FaClock, FaChartPie, FaStore
} = require("react-icons/fa");

// ---- COLORS ----
const C = {
  dark: "0F172A",
  darkBg: "1E293B",
  accent: "3B82F6",
  warm: "F97316",
  white: "FFFFFF",
  gray50: "F9FAFB",
  gray100: "F3F4F6",
  gray200: "E5E7EB",
  gray300: "D1D5DB",
  gray400: "9CA3AF",
  gray500: "6B7280",
  gray600: "4B5563",
  gray700: "374151",
  gray800: "1F2937",
  blue50: "EFF6FF",
  orange50: "FFF7ED",
  red50: "FEF2F2",
  red400: "F87171",
  red500: "EF4444",
  yellow50: "FEFCE8",
  yellow600: "CA8A04",
  green50: "F0FDF4",
  green500: "22C55E",
  green700: "15803D",
};

// ---- ICON HELPER ----
function renderIconSvg(IconComponent, color, size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}

async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = renderIconSvg(IconComponent, "#" + color, size);
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

// ---- SHADOW FACTORY ----
const mkShadow = (opts = {}) => ({
  type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.12, ...opts
});

// ---- MAIN ----
async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10" x 5.625"
  pres.author = "AI Rental Car";
  pres.title = "AIレンタカー 事業計画書";

  const W = 10, H = 5.625;

  // Pre-render icons
  const icons = {};
  const iconList = [
    ["car", FaCar, C.white], ["carAccent", FaCar, C.accent], ["carDark", FaCar, C.dark],
    ["calendarAlt", FaCalendarAlt, C.accent], ["mapMarker", FaMapMarkerAlt, C.accent],
    ["briefcase", FaBriefcase, C.accent], ["robot", FaRobot, C.accent],
    ["mapMarkerWarm", FaMapMarkerAlt, C.warm], ["users", FaUsers, C.accent],
    ["mountain", FaMountain, C.warm], ["building", FaBuilding, C.gray500],
    ["chartLine", FaChartLine, C.accent], ["chartArea", FaChartArea, C.accent],
    ["conciergeBell", FaConciergeBell, C.warm], ["bullhorn", FaBullhorn, C.accent],
    ["globe", FaGlobe, C.accent], ["search", FaSearch, C.accent],
    ["mapMarked", FaMapMarkedAlt, C.warm], ["instagram", FaInstagram, C.warm],
    ["cogs", FaCogs, C.white], ["userCog", FaUserCog, C.warm],
    ["yenSign", FaYenSign, C.accent], ["calculator", FaCalculator, C.warm],
    ["phoneAlt", FaPhoneAlt, C.accent], ["broom", FaBroom, C.accent],
    ["fileAlt", FaFileAlt, C.accent], ["handsHelping", FaHandsHelping, C.warm],
    ["robotWarm", FaRobot, C.warm], ["userTie", FaUserTie, C.accent],
    ["tag", FaTag, C.accent], ["calendarCheck", FaCalendarCheck, C.accent],
    ["chartBar", FaChartBar, C.accent], ["minusCircle", FaMinusCircle, C.red400],
    ["wallet", FaWallet, C.accent], ["university", FaUniversity, C.warm],
    ["piggyBank", FaPiggyBank, C.warm], ["shieldAlt", FaShieldAlt, C.accent],
    ["shieldRed", FaShieldAlt, "EF4444"], ["road", FaRoad, C.accent],
    ["rocket", FaRocket, C.white], ["rocketDark", FaRocket, C.dark],
    ["excTriangle", FaExclamationTriangle, C.red500], ["excTriangleYellow", FaExclamationTriangle, C.yellow600],
    ["tags", FaTags, C.warm], ["coins", FaCoins, C.accent],
    ["seedling", FaSeedling, C.white], ["expand", FaExpandArrowsAlt, C.white],
    ["envelope", FaEnvelope, C.accent], ["checkCircle", FaCheckCircle, C.accent],
    ["checkCircleWarm", FaCheckCircle, C.warm], ["checkCircleDark", FaCheckCircle, C.dark],
    ["timesCircle", FaTimesCircle, C.red400], ["star", FaStar, C.accent],
    ["fileContract", FaFileContract, C.warm], ["toolbox", FaToolbox, C.accent],
    ["ad", FaAd, C.warm], ["umbrellaBeach", FaUmbrellaBeach, C.warm],
    ["handshake", FaHandshake, C.warm], ["flag", FaFlag, C.white],
    ["arrowUp", FaArrowUp, C.white], ["chevRight", FaChevronRight, C.gray300],
    ["chevDown", FaChevronDown, C.gray300], ["clock", FaClock, C.gray400],
    ["chartPie", FaChartPie, C.accent], ["chartPieWarm", FaChartPie, C.warm],
    ["store", FaStore, C.yellow600], ["carWhite", FaCar, C.white],
    ["chartLineYellow", FaChartLine, C.yellow600], ["carAccentIcon", FaCar, C.accent],
    ["carWarm", FaCar, C.warm],
  ];
  for (const [name, comp, color] of iconList) {
    icons[name] = await iconToBase64Png(comp, color);
  }

  // ========== HELPER FUNCTIONS ==========

  function addHeader(slide, enLabel, jpTitle, pageNum, accentColor = C.accent) {
    // Accent bar
    slide.addShape(pres.shapes.RECTANGLE, { x: 1.25, y: 0.7, w: 0.04, h: 0.55, fill: { color: accentColor } });
    // English label
    slide.addText(enLabel.toUpperCase(), {
      x: 1.4, y: 0.68, w: 5, h: 0.25, fontSize: 8, fontFace: "Inter",
      color: C.gray400, charSpacing: 3, margin: 0
    });
    // Japanese title
    slide.addText(jpTitle, {
      x: 1.4, y: 0.88, w: 6, h: 0.45, fontSize: 22, fontFace: "Noto Sans JP",
      color: C.dark, bold: true, margin: 0
    });
    // Brand icon + text (right)
    slide.addImage({ data: icons.carDark, x: 8.2, y: 0.78, w: 0.25, h: 0.25 });
    slide.addText("AI RENTAL CAR", {
      x: 8.5, y: 0.78, w: 1.3, h: 0.25, fontSize: 7, fontFace: "Inter",
      color: C.gray400, bold: true, charSpacing: 2, margin: 0
    });
    // Header line
    slide.addShape(pres.shapes.LINE, { x: 1.25, y: 1.35, w: 7.5, h: 0, line: { color: C.gray200, width: 0.75 } });
    // Footer
    addFooter(slide, pageNum);
  }

  function addFooter(slide, pageNum) {
    slide.addShape(pres.shapes.LINE, { x: 0, y: 5.15, w: 10, h: 0, line: { color: C.gray100, width: 0.5 } });
    slide.addText("Confidential", {
      x: 1.25, y: 5.2, w: 3, h: 0.3, fontSize: 8, fontFace: "Noto Sans JP", color: C.gray400, margin: 0
    });
    slide.addText([
      { text: "Page ", options: { fontSize: 8, color: C.gray400, fontFace: "Inter" } },
      { text: String(pageNum).padStart(2, "0"), options: { fontSize: 10, color: C.accent, bold: true, fontFace: "Inter" } }
    ], { x: 8, y: 5.2, w: 1, h: 0.3, align: "right", margin: 0 });
  }

  function addSectionDivider(slide, sectionNum, enLabel, jpTitle, subtitle, items, icon, pageNum, accentColor = C.accent) {
    // Left panel
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 3.33, h: H, fill: { color: C.dark } });
    // Decorative circles
    slide.addShape(pres.shapes.OVAL, { x: 2.1, y: -0.5, w: 2.2, h: 2.2, fill: { color: accentColor, transparency: 80 } });
    slide.addShape(pres.shapes.OVAL, { x: -0.5, y: 3.5, w: 1.8, h: 1.8, fill: { color: C.warm, transparency: 90 } });
    // Icon box
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.75, y: 0.8, w: 0.6, h: 0.6, rectRadius: 0.08, fill: { color: accentColor } });
    slide.addImage({ data: icon, x: 0.88, y: 0.93, w: 0.35, h: 0.35 });
    // Section label
    slide.addText("SECTION " + String(sectionNum).padStart(2, "0"), {
      x: 1.55, y: 0.88, w: 1.5, h: 0.35, fontSize: 9, fontFace: "Inter",
      color: accentColor, bold: true, charSpacing: 3, margin: 0
    });
    // Title
    slide.addText(jpTitle, {
      x: 0.75, y: 1.7, w: 2.3, h: 1, fontSize: 28, fontFace: "Noto Sans JP",
      color: C.white, bold: true, margin: 0
    });
    // Subtitle
    slide.addText(subtitle, {
      x: 0.75, y: 2.8, w: 2.3, h: 0.6, fontSize: 10, fontFace: "Noto Sans JP",
      color: C.gray400, margin: 0
    });
    // Bottom items
    slide.addText(items.join("  |  "), {
      x: 0.75, y: 4.4, w: 2.3, h: 0.3, fontSize: 8, fontFace: "Noto Sans JP", color: C.gray500, margin: 0
    });
    slide.addText("Confidential", {
      x: 0.75, y: 4.85, w: 2, h: 0.2, fontSize: 8, fontFace: "Noto Sans JP", color: C.gray600, margin: 0
    });
    return slide;
  }

  function addSectionRightItem(slide, icon, enLabel, jpTitle, desc, y, labelColor) {
    slide.addShape(pres.shapes.OVAL, { x: 3.9, y: y, w: 0.6, h: 0.6, fill: { color: labelColor === C.warm ? C.orange50 : C.blue50 } });
    slide.addImage({ data: icon, x: 4.05, y: y + 0.12, w: 0.3, h: 0.3 });
    slide.addText(enLabel.toUpperCase(), {
      x: 4.7, y: y - 0.05, w: 3, h: 0.2, fontSize: 8, fontFace: "Inter",
      color: labelColor, bold: true, charSpacing: 2, margin: 0
    });
    slide.addText(jpTitle, {
      x: 4.7, y: y + 0.15, w: 4, h: 0.3, fontSize: 14, fontFace: "Noto Sans JP",
      color: C.dark, bold: true, margin: 0
    });
    slide.addText(desc, {
      x: 4.7, y: y + 0.45, w: 4.5, h: 0.25, fontSize: 10, fontFace: "Noto Sans JP",
      color: C.gray600, margin: 0
    });
  }

  // ========== SLIDE 1: COVER ==========
  {
    const slide = pres.addSlide();
    // Gradient bg via dark fill
    slide.background = { color: C.dark };
    // Darker overlay shapes for gradient effect
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H, fill: { color: C.darkBg, transparency: 30 } });
    // Accent bar left
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.06, h: H, fill: { color: C.accent } });
    // Decorative circles
    slide.addShape(pres.shapes.OVAL, { x: 7.5, y: 0.3, w: 2.8, h: 2.8, fill: { color: C.accent, transparency: 90 } });
    slide.addShape(pres.shapes.OVAL, { x: 1, y: 3.5, w: 2.2, h: 2.2, fill: { color: C.warm, transparency: 95 } });
    // Icon box
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 4.4, y: 0.8, w: 1.2, h: 1.2, rectRadius: 0.15, fill: { color: C.accent }, shadow: mkShadow() });
    slide.addImage({ data: icons.carWhite, x: 4.7, y: 1.05, w: 0.6, h: 0.6 });
    // Business Plan label
    slide.addText("BUSINESS PLAN", {
      x: 2.5, y: 2.2, w: 5, h: 0.3, fontSize: 10, fontFace: "Inter",
      color: C.accent, bold: true, charSpacing: 5, align: "center", margin: 0
    });
    // Title
    slide.addText("AIレンタカー", {
      x: 1, y: 2.5, w: 8, h: 0.8, fontSize: 42, fontFace: "Noto Sans JP",
      color: C.white, bold: true, align: "center", margin: 0
    });
    slide.addText("事業計画書", {
      x: 1, y: 3.2, w: 8, h: 0.5, fontSize: 20, fontFace: "Noto Sans JP",
      color: C.gray300, align: "center", margin: 0
    });
    // Accent line
    slide.addShape(pres.shapes.RECTANGLE, { x: 4.6, y: 3.8, w: 0.8, h: 0.04, fill: { color: C.accent } });
    // Tagline
    slide.addText("AI活用で実現する次世代レンタカー事業", {
      x: 1.5, y: 3.95, w: 7, h: 0.35, fontSize: 11, fontFace: "Noto Sans JP",
      color: C.gray400, align: "center", margin: 0
    });
    // Date & location
    slide.addImage({ data: icons.calendarAlt, x: 3.4, y: 4.7, w: 0.2, h: 0.2 });
    slide.addText("2026 / 02", {
      x: 3.65, y: 4.7, w: 1.2, h: 0.2, fontSize: 9, fontFace: "Inter", color: C.gray500, margin: 0
    });
    slide.addImage({ data: icons.mapMarker, x: 5.3, y: 4.7, w: 0.2, h: 0.2 });
    slide.addText("千葉県", {
      x: 5.55, y: 4.7, w: 0.8, h: 0.2, fontSize: 9, fontFace: "Noto Sans JP", color: C.gray500, margin: 0
    });
    // Footer
    addFooter(slide, 1);
  }

  // ========== SLIDE 2: AGENDA ==========
  {
    const slide = pres.addSlide();
    slide.background = { color: C.white };
    addHeader(slide, "Agenda", "目次", 2);

    const sections = [
      { num: "01", en: "Overview", jp: "事業概要・提供価値・ターゲット", color: C.accent },
      { num: "02", en: "Market & Service", jp: "市場環境・サービス・集客", color: C.accent },
      { num: "03", en: "Operations & Finance", jp: "オペレーション・投資・収支", color: C.warm },
      { num: "04", en: "Growth Strategy", jp: "資金調達・リスク・成長戦略", color: C.warm },
    ];

    sections.forEach((s, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = 1.25 + col * 4.2;
      const y = 1.7 + row * 1.4;
      // Number badge
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 0.65, h: 0.65, rectRadius: 0.1, fill: { color: s.color }, shadow: mkShadow() });
      slide.addText(s.num, { x, y, w: 0.65, h: 0.65, fontSize: 14, fontFace: "Inter", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
      // Label
      slide.addText(s.en.toUpperCase(), {
        x: x + 0.85, y: y + 0.02, w: 3, h: 0.2, fontSize: 8, fontFace: "Inter",
        color: s.color, bold: true, charSpacing: 2, margin: 0
      });
      slide.addText(s.jp, {
        x: x + 0.85, y: y + 0.25, w: 3.2, h: 0.35, fontSize: 13, fontFace: "Noto Sans JP",
        color: C.dark, bold: true, margin: 0
      });
    });
  }

  // ========== SLIDE 3: SECTION 1 DIVIDER ==========
  {
    const slide = pres.addSlide();
    slide.background = { color: C.gray50 };
    addSectionDivider(slide, 1, "Section 01", "事業概要", "AIレンタカー事業の\n全体像をご紹介します",
      ["Overview", "提供価値", "ターゲット"], icons.carWhite, 3, C.accent);

    addSectionRightItem(slide, icons.briefcase, "Business Model", "個人事業主による副業スタート",
      "普通自動車のレンタカー事業（日貸・時間貸）", 1.3, C.accent);
    addSectionRightItem(slide, icons.mapMarkerWarm, "Location", "千葉県",
      "首都圏・房総エリアへのアクセス拠点", 2.5, C.warm);
    addSectionRightItem(slide, icons.robot, "Future Vision", "AI活用による高稼働運営",
      "需要予測・価格最適化・予約最適化で競合と差別化", 3.7, C.accent);
  }

  // ========== SLIDE 4: 事業概要・提供価値 ==========
  {
    const slide = pres.addSlide();
    slide.background = { color: C.white };
    addHeader(slide, "Value Proposition", "事業概要・提供価値", 4);

    // Left column: 事業概要
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.25, y: 1.5, w: 1.1, h: 0.3, rectRadius: 0.05, fill: { color: C.accent } });
    slide.addText("OVERVIEW", { x: 1.25, y: 1.5, w: 1.1, h: 0.3, fontSize: 8, fontFace: "Inter", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    slide.addText("事業概要", { x: 2.45, y: 1.5, w: 1.5, h: 0.3, fontSize: 13, fontFace: "Noto Sans JP", color: C.dark, bold: true, margin: 0 });

    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.25, y: 1.95, w: 3.5, h: 2.9, rectRadius: 0.1, fill: { color: C.gray50 }, line: { color: C.gray100, width: 0.5 } });

    const leftItems = [
      { icon: icons.tag, label: "事業名", val: "AIレンタカー" },
      { icon: icons.briefcase, label: "事業形態", val: "個人事業主（副業スタート）" },
      { icon: icons.carAccent, label: "事業内容", val: "普通自動車のレンタカー事業\n（日貸・時間貸）" },
      { icon: icons.mapMarker, label: "開業予定地", val: "千葉県（首都圏・房総エリア）" },
    ];
    leftItems.forEach((item, i) => {
      const y = 2.15 + i * 0.65;
      slide.addImage({ data: item.icon, x: 1.5, y: y, w: 0.2, h: 0.2 });
      slide.addText(item.label, { x: 1.85, y: y - 0.05, w: 2, h: 0.2, fontSize: 10, fontFace: "Noto Sans JP", color: C.dark, bold: true, margin: 0 });
      slide.addText(item.val, { x: 1.85, y: y + 0.15, w: 2.7, h: 0.35, fontSize: 8, fontFace: "Noto Sans JP", color: C.gray600, margin: 0 });
    });

    // Right column: 提供価値
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.25, y: 1.5, w: 0.85, h: 0.3, rectRadius: 0.05, fill: { color: C.warm } });
    slide.addText("VALUE", { x: 5.25, y: 1.5, w: 0.85, h: 0.3, fontSize: 8, fontFace: "Inter", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    slide.addText("提供価値", { x: 6.2, y: 1.5, w: 1.5, h: 0.3, fontSize: 13, fontFace: "Noto Sans JP", color: C.dark, bold: true, margin: 0 });

    const rightItems = [
      { icon: icons.umbrellaBeach, label: "レジャー最適化", desc: "週末利用に特化した車両ラインナップ", bg: C.orange50 },
      { icon: icons.checkCircle, label: "高品質な車両管理", desc: "清潔・高品質な車両を常に提供", bg: C.blue50 },
      { icon: icons.handshake, label: "柔軟な対応", desc: "個人運営ならではの柔軟な受渡・対応", bg: C.orange50 },
      { icon: icons.robot, label: "AI活用運営", desc: "需要予測・価格最適化・予約最適化", bg: C.blue50 },
    ];
    rightItems.forEach((item, i) => {
      const y = 1.95 + i * 0.7;
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.25, y, w: 3.5, h: 0.6, rectRadius: 0.08, fill: { color: C.white }, line: { color: C.gray200, width: 0.5 }, shadow: mkShadow() });
      slide.addShape(pres.shapes.OVAL, { x: 5.4, y: y + 0.08, w: 0.44, h: 0.44, fill: { color: item.bg } });
      slide.addImage({ data: item.icon, x: 5.5, y: y + 0.15, w: 0.22, h: 0.22 });
      slide.addText(item.label, { x: 5.95, y: y + 0.05, w: 2.5, h: 0.2, fontSize: 10, fontFace: "Noto Sans JP", color: C.dark, bold: true, margin: 0 });
      slide.addText(item.desc, { x: 5.95, y: y + 0.28, w: 2.6, h: 0.2, fontSize: 8, fontFace: "Noto Sans JP", color: C.gray600, margin: 0 });
    });
  }

  // ========== SLIDE 5: ターゲット顧客 ==========
  {
    const slide = pres.addSlide();
    slide.background = { color: C.white };
    addHeader(slide, "Target Customer", "ターゲット顧客", 5);

    const cards = [
      { icon: icons.users, title: "首都圏の個人・\nファミリー", desc: "首都圏在住でマイカーを持たない層。週末や連休にレジャーで車が必要な個人・家族。", badge: "Primary Target", badgeIcon: icons.chartPie, borderColor: C.accent, iconBg: C.blue50 },
      { icon: icons.mountain, title: "アウトドア・\n観光利用者", desc: "週末・連休に房総エリアでの観光やアウトドアを楽しむアクティブ層。", badge: "Primary Target", badgeIcon: icons.chartPieWarm, borderColor: C.warm, iconBg: C.orange50 },
      { icon: icons.building, title: "法人・長期\nレンタル需要", desc: "将来的には法人利用や長期レンタル需要にも対応し、安定収益を確保する。", badge: "Future Target", badgeIcon: icons.clock, borderColor: "64748B", iconBg: C.gray100 },
    ];

    cards.forEach((c, i) => {
      const x = 1.25 + i * 2.7;
      // Card bg
      slide.addShape(pres.shapes.RECTANGLE, { x, y: 1.55, w: 2.4, h: 3.2, fill: { color: C.white }, line: { color: C.gray200, width: 0.5 }, shadow: mkShadow() });
      // Top border
      slide.addShape(pres.shapes.RECTANGLE, { x, y: 1.55, w: 2.4, h: 0.06, fill: { color: c.borderColor } });
      // Icon circle
      slide.addShape(pres.shapes.OVAL, { x: x + 0.2, y: 1.85, w: 0.6, h: 0.6, fill: { color: c.iconBg } });
      slide.addImage({ data: c.icon, x: x + 0.35, y: 1.97, w: 0.3, h: 0.3 });
      // Title
      slide.addText(c.title, { x: x + 0.15, y: 2.55, w: 2.1, h: 0.55, fontSize: 12, fontFace: "Noto Sans JP", color: C.dark, bold: true, margin: 0 });
      // Description
      slide.addText(c.desc, { x: x + 0.15, y: 3.15, w: 2.1, h: 0.9, fontSize: 9, fontFace: "Noto Sans JP", color: C.gray600, margin: 0 });
      // Footer badge
      slide.addShape(pres.shapes.LINE, { x: x + 0.15, y: 4.2, w: 2.1, h: 0, line: { color: C.gray100, width: 0.5 } });
      slide.addImage({ data: c.badgeIcon, x: x + 0.2, y: 4.35, w: 0.15, h: 0.15 });
      slide.addText(c.badge, { x: x + 0.42, y: 4.35, w: 1.5, h: 0.15, fontSize: 8, fontFace: "Inter", color: C.gray500, margin: 0 });
    });
  }

  // ========== SLIDE 6: SECTION 2 DIVIDER ==========
  {
    const slide = pres.addSlide();
    slide.background = { color: C.gray50 };
    addSectionDivider(slide, 2, "Section 02", "市場環境\nサービス", "レンタカー市場の動向と\n当社のサービス設計",
      ["市場分析", "サービス", "集客"], icons.chartLine, 6, C.accent);

    addSectionRightItem(slide, icons.chartArea, "Market", "市場・競合環境",
      "観光回復とアウトドア需要の拡大で安定成長", 1.3, C.accent);
    addSectionRightItem(slide, icons.conciergeBell, "Service", "サービス内容",
      "レジャー特化プランとオプションサービス", 2.5, C.warm);
    addSectionRightItem(slide, icons.bullhorn, "Channel", "集客・販売チャネル",
      "Web・SNS・口コミによるマルチチャネル展開", 3.7, C.accent);
  }

  // ========== SLIDE 7: 市場・競合環境 ==========
  {
    const slide = pres.addSlide();
    slide.background = { color: C.white };
    addHeader(slide, "Market Analysis", "市場・競合環境", 7);

    const gridCards = [
      {
        title: "市場動向", sub: "Market Trend", badge: "成長", badgeColor: C.green700, badgeBg: C.green50,
        icon: icons.chartLine, iconBg: C.blue50,
        items: ["観光回復による需要増加", "アウトドアブームの継続", "レンタカー市場は安定成長基調"],
        itemIcon: icons.checkCircle
      },
      {
        title: "大手レンタカー", sub: "Major Players", badge: "競合", badgeColor: "DC2626", badgeBg: C.red50,
        icon: icons.excTriangle, iconBg: C.red50,
        items: ["画一的なサービス", "高価格設定", "柔軟性に欠ける対応"],
        itemIcon: icons.timesCircle
      },
      {
        title: "地場中小レンタカー", sub: "Local Players", badge: "競合", badgeColor: C.yellow600, badgeBg: C.yellow50,
        icon: icons.store, iconBg: C.yellow50,
        items: ["品質にばらつきがある", "デジタル化の遅れ", "集客力が弱い"],
        itemIcon: icons.excTriangleYellow
      },
      {
        title: "AIレンタカーの強み", sub: "Our Advantage", badge: "差別化", badgeColor: C.white, badgeBg: C.accent,
        icon: icons.star, iconBg: C.blue50, highlighted: true,
        items: ["用途特化（レジャー・アウトドア）", "品質重視の車両管理", "柔軟運営 + AI活用"],
        itemIcon: icons.checkCircle
      },
    ];

    gridCards.forEach((c, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = 1.25 + col * 3.85;
      const y = 1.55 + row * 1.7;

      // Card
      const lineOpts = c.highlighted ? { color: C.accent, width: 1.5 } : { color: C.gray200, width: 0.5 };
      slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.55, h: 1.5, fill: { color: C.white }, line: lineOpts, shadow: mkShadow() });
      // Badge
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 2.6, y: y + 0.1, w: 0.7, h: 0.25, rectRadius: 0.04, fill: { color: c.badgeBg } });
      slide.addText(c.badge, { x: x + 2.6, y: y + 0.1, w: 0.7, h: 0.25, fontSize: 7, fontFace: "Noto Sans JP", color: c.badgeColor, bold: true, align: "center", valign: "middle", margin: 0 });
      // Icon
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.15, y: y + 0.15, w: 0.5, h: 0.5, rectRadius: 0.06, fill: { color: c.iconBg } });
      slide.addImage({ data: c.icon, x: x + 0.27, y: y + 0.27, w: 0.25, h: 0.25 });
      // Title
      slide.addText(c.title, { x: x + 0.8, y: y + 0.15, w: 1.8, h: 0.22, fontSize: 10, fontFace: "Noto Sans JP", color: C.dark, bold: true, margin: 0 });
      slide.addText(c.sub, { x: x + 0.8, y: y + 0.37, w: 1.5, h: 0.15, fontSize: 7, fontFace: "Inter", color: C.gray500, margin: 0 });
      // Items
      c.items.forEach((item, j) => {
        const iy = y + 0.7 + j * 0.22;
        slide.addImage({ data: c.itemIcon, x: x + 0.2, y: iy, w: 0.15, h: 0.15 });
        slide.addText(item, { x: x + 0.45, y: iy, w: 2.8, h: 0.18, fontSize: 8, fontFace: "Noto Sans JP", color: c.highlighted ? C.dark : C.gray600, bold: c.highlighted, margin: 0 });
      });
    });
  }

  // ========== SLIDE 8: サービス内容 ==========
  {
    const slide = pres.addSlide();
    slide.background = { color: C.white };
    addHeader(slide, "Service Details", "サービス内容", 8);

    const services = [
      { num: "01", title: "レンタカー（日貸・時間貸）", desc: "普通自動車の日貸・時間貸サービス。利用者のニーズに合わせた柔軟な貸出期間を提供します。", badge: "メイン", badgeColor: C.accent, numBg: C.accent },
      { num: "02", title: "週末・連休向けプラン", desc: "レジャー需要にフォーカスした週末・連休パッケージプラン。お得な料金設定で稼働率を向上。", badge: "メイン", badgeColor: C.accent, numBg: C.accent },
      { num: "03", title: "オプションサービス", desc: "チャイルドシート、ETC車載器、アウトドア用品などのオプション。ファミリー層やアウトドア利用者の利便性を高めます。", badge: "付加価値", badgeColor: C.warm, numBg: C.warm },
      { num: "04", title: "将来構想：キャンピングカー・法人契約・長期レンタル", desc: "事業拡大フェーズではキャンピングカー特化、法人向け契約、長期レンタルサービスを展開し収益の多角化を図ります。", badge: "将来", badgeColor: "64748B", numBg: "64748B" },
    ];

    services.forEach((s, i) => {
      const y = 1.55 + i * 0.82;
      slide.addShape(pres.shapes.RECTANGLE, { x: 1.25, y, w: 7.5, h: 0.7, fill: { color: C.white }, line: { color: C.gray200, width: 0.5 }, shadow: mkShadow() });
      // Number circle
      slide.addShape(pres.shapes.OVAL, { x: 1.45, y: y + 0.12, w: 0.45, h: 0.45, fill: { color: s.numBg } });
      slide.addText(s.num, { x: 1.45, y: y + 0.12, w: 0.45, h: 0.45, fontSize: 9, fontFace: "Inter", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
      // Title
      slide.addText(s.title, { x: 2.1, y: y + 0.08, w: 5, h: 0.22, fontSize: 10, fontFace: "Noto Sans JP", color: C.dark, bold: true, margin: 0 });
      // Description
      slide.addText(s.desc, { x: 2.1, y: y + 0.32, w: 5.5, h: 0.3, fontSize: 8, fontFace: "Noto Sans JP", color: C.gray500, margin: 0 });
      // Badge
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.9, y: y + 0.18, w: 0.7, h: 0.25, rectRadius: 0.04, fill: { color: s.badgeColor } });
      slide.addText(s.badge, { x: 7.9, y: y + 0.18, w: 0.7, h: 0.25, fontSize: 7, fontFace: "Noto Sans JP", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    });
  }

  // ========== SLIDE 9: 集客・販売チャネル ==========
  {
    const slide = pres.addSlide();
    slide.background = { color: C.white };
    addHeader(slide, "Sales Channel", "集客・販売チャネル", 9);

    const channels = [
      { num: "Ch.1", icon: icons.globe, iconBg: C.blue50, title: "自社Webサイト", desc: "予約・問い合わせの中心チャネル。SEO対策で検索流入を獲得。", footer: "予約・問い合わせ", footerColor: C.accent, borderColor: C.accent },
      { num: "Ch.2", icon: icons.search, iconBg: C.blue50, title: "比較サイト", desc: "レンタカー比較サイトへの掲載で新規顧客のリーチを拡大。", footer: "新規獲得", footerColor: C.accent, borderColor: C.accent },
      { num: "Ch.3", icon: icons.mapMarked, iconBg: C.orange50, title: "Google マップ", desc: "Googleマップ・口コミで地域の信頼性を構築。近隣顧客を獲得。", footer: "信頼構築", footerColor: C.warm, borderColor: C.warm },
      { num: "Ch.4", icon: icons.instagram, iconBg: C.orange50, title: "SNS", desc: "Instagram等のSNSで車両・レジャー情報を発信。ブランド認知を向上。", footer: "ブランド認知", footerColor: C.warm, borderColor: C.warm },
    ];

    channels.forEach((ch, i) => {
      const x = 1.25 + i * 2.05;
      slide.addShape(pres.shapes.RECTANGLE, { x, y: 1.55, w: 1.8, h: 3.2, fill: { color: C.white }, line: { color: C.gray200, width: 0.5 }, shadow: mkShadow() });
      slide.addShape(pres.shapes.RECTANGLE, { x, y: 1.55, w: 1.8, h: 0.06, fill: { color: ch.borderColor } });
      // Icon
      slide.addShape(pres.shapes.OVAL, { x: x + 0.15, y: 1.8, w: 0.5, h: 0.5, fill: { color: ch.iconBg } });
      slide.addImage({ data: ch.icon, x: x + 0.27, y: 1.92, w: 0.25, h: 0.25 });
      // Badge
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 1.05, y: 1.85, w: 0.55, h: 0.22, rectRadius: 0.04, fill: { color: ch.borderColor } });
      slide.addText(ch.num, { x: x + 1.05, y: 1.85, w: 0.55, h: 0.22, fontSize: 7, fontFace: "Inter", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
      // Title
      slide.addText(ch.title, { x: x + 0.1, y: 2.45, w: 1.6, h: 0.35, fontSize: 12, fontFace: "Noto Sans JP", color: C.dark, bold: true, margin: 0 });
      // Desc
      slide.addText(ch.desc, { x: x + 0.1, y: 2.85, w: 1.6, h: 0.9, fontSize: 8, fontFace: "Noto Sans JP", color: C.gray500, margin: 0 });
      // Footer
      slide.addShape(pres.shapes.LINE, { x: x + 0.1, y: 4.1, w: 1.6, h: 0, line: { color: C.gray100, width: 0.5 } });
      slide.addText(ch.footer, { x: x + 0.1, y: 4.2, w: 1.6, h: 0.2, fontSize: 9, fontFace: "Noto Sans JP", color: ch.footerColor, bold: true, margin: 0 });
      // Arrow (except last)
      if (i < 3) {
        slide.addImage({ data: icons.chevRight, x: x + 1.82, y: 2.9, w: 0.2, h: 0.2 });
      }
    });
  }

  // ========== SLIDE 10: SECTION 3 DIVIDER ==========
  {
    const slide = pres.addSlide();
    slide.background = { color: C.gray50 };
    addSectionDivider(slide, 3, "Section 03", "オペレーション\n財務計画", "運営体制・初期投資・\n収支モデルについて",
      ["運営体制", "初期投資", "収支"], icons.cogs, 10, C.warm);

    addSectionRightItem(slide, icons.userCog, "Operation", "オペレーション体制",
      "代表者1名による効率的な運営設計", 1.3, C.warm);
    addSectionRightItem(slide, icons.yenSign, "Investment", "初期投資計画",
      "700〜1,000万円の初期投資内訳", 2.5, C.accent);
    addSectionRightItem(slide, icons.calculator, "Revenue Model", "月次収支モデル",
      "月商約18万円の収支シミュレーション", 3.7, C.warm);
  }

  // ========== SLIDE 11: オペレーション体制 ==========
  {
    const slide = pres.addSlide();
    slide.background = { color: C.white };
    addHeader(slide, "Operations", "オペレーション体制", 11, C.warm);

    // Layer 1: Core
    slide.addShape(pres.shapes.RECTANGLE, { x: 1.25, y: 1.55, w: 7.5, h: 0.85, fill: { color: C.blue50 }, line: { color: C.accent, width: 1.2 } });
    slide.addText("CORE", { x: 1.4, y: 1.6, w: 0.7, h: 0.15, fontSize: 7, fontFace: "Inter", color: C.accent, bold: true, charSpacing: 2, margin: 0 });
    slide.addText("運営体制", { x: 1.4, y: 1.75, w: 0.8, h: 0.2, fontSize: 10, fontFace: "Noto Sans JP", color: C.dark, bold: true, margin: 0 });
    slide.addImage({ data: icons.userTie, x: 2.5, y: 1.7, w: 0.25, h: 0.25 });
    slide.addText("代表者 1名", { x: 2.85, y: 1.65, w: 2, h: 0.2, fontSize: 11, fontFace: "Noto Sans JP", color: C.dark, bold: true, margin: 0 });
    slide.addText("受付・車両管理・清掃・事務を兼任", { x: 2.85, y: 1.85, w: 3, h: 0.2, fontSize: 8, fontFace: "Noto Sans JP", color: C.gray500, margin: 0 });

    // Arrow
    slide.addImage({ data: icons.chevDown, x: 4.9, y: 2.5, w: 0.2, h: 0.2 });

    // Layer 2: Daily
    slide.addShape(pres.shapes.RECTANGLE, { x: 1.25, y: 2.8, w: 7.5, h: 0.9, fill: { color: C.gray50 }, line: { color: C.gray200, width: 0.5 } });
    slide.addText("DAILY", { x: 1.4, y: 2.85, w: 0.7, h: 0.15, fontSize: 7, fontFace: "Inter", color: C.accent, bold: true, charSpacing: 2, margin: 0 });
    slide.addText("日常業務", { x: 1.4, y: 3.0, w: 0.8, h: 0.2, fontSize: 10, fontFace: "Noto Sans JP", color: C.dark, bold: true, margin: 0 });

    const dailyItems = [
      { icon: icons.phoneAlt, label: "受付対応" },
      { icon: icons.carAccent, label: "車両管理" },
      { icon: icons.broom, label: "清掃" },
      { icon: icons.fileAlt, label: "事務処理" },
    ];
    dailyItems.forEach((item, i) => {
      const x = 2.5 + i * 1.55;
      slide.addShape(pres.shapes.RECTANGLE, { x, y: 2.9, w: 1.3, h: 0.65, fill: { color: C.white }, line: { color: C.gray100, width: 0.5 } });
      slide.addImage({ data: item.icon, x: x + 0.45, y: 2.95, w: 0.3, h: 0.3 });
      slide.addText(item.label, { x, y: 3.28, w: 1.3, h: 0.2, fontSize: 8, fontFace: "Noto Sans JP", color: C.dark, bold: true, align: "center", margin: 0 });
    });

    // Arrow
    slide.addImage({ data: icons.chevDown, x: 4.9, y: 3.8, w: 0.2, h: 0.2 });

    // Layer 3: Peak
    slide.addShape(pres.shapes.RECTANGLE, { x: 1.25, y: 4.1, w: 7.5, h: 0.85, fill: { color: C.gray50 }, line: { color: C.gray200, width: 0.5 } });
    slide.addText("PEAK", { x: 1.4, y: 4.15, w: 0.7, h: 0.15, fontSize: 7, fontFace: "Inter", color: C.warm, bold: true, charSpacing: 2, margin: 0 });
    slide.addText("繁忙期対応", { x: 1.4, y: 4.3, w: 1, h: 0.2, fontSize: 10, fontFace: "Noto Sans JP", color: C.dark, bold: true, margin: 0 });
    // External help
    slide.addImage({ data: icons.handsHelping, x: 2.6, y: 4.25, w: 0.25, h: 0.25 });
    slide.addText("外注活用", { x: 2.95, y: 4.2, w: 1.5, h: 0.2, fontSize: 10, fontFace: "Noto Sans JP", color: C.dark, bold: true, margin: 0 });
    slide.addText("清掃・点検業務の外部委託", { x: 2.95, y: 4.4, w: 2.5, h: 0.2, fontSize: 8, fontFace: "Noto Sans JP", color: C.gray500, margin: 0 });
    // AI
    slide.addImage({ data: icons.robotWarm, x: 5.6, y: 4.25, w: 0.25, h: 0.25 });
    slide.addText("AI効率化", { x: 5.95, y: 4.2, w: 1.5, h: 0.2, fontSize: 10, fontFace: "Noto Sans JP", color: C.dark, bold: true, margin: 0 });
    slide.addText("予約最適化で業務負荷を軽減", { x: 5.95, y: 4.4, w: 2.5, h: 0.2, fontSize: 8, fontFace: "Noto Sans JP", color: C.gray500, margin: 0 });
  }

  // ========== SLIDE 12: 初期投資 ==========
  {
    const slide = pres.addSlide();
    slide.background = { color: C.white };
    addHeader(slide, "Initial Investment", "初期投資", 12, C.warm);

    const kpis = [
      { label: "VEHICLE", value: "600-900", unit: "万円", desc: "車両購入（中古1台）", icon: icons.carAccent, iconBg: C.blue50, barColor: C.accent },
      { label: "REGISTRATION", value: "50", unit: "万円", desc: "登録・整備・保険", icon: icons.fileContract, iconBg: C.orange50, barColor: C.warm },
      { label: "EQUIPMENT", value: "20", unit: "万円", desc: "設備・備品", icon: icons.toolbox, iconBg: C.blue50, barColor: C.accent },
      { label: "MARKETING", value: "10", unit: "万円", desc: "Web・広告", icon: icons.ad, iconBg: C.orange50, barColor: C.warm },
    ];

    kpis.forEach((k, i) => {
      const x = 1.25 + i * 2.05;
      slide.addShape(pres.shapes.RECTANGLE, { x, y: 1.55, w: 1.8, h: 1.2, fill: { color: C.white }, line: { color: C.gray200, width: 0.5 }, shadow: mkShadow() });
      slide.addShape(pres.shapes.RECTANGLE, { x, y: 1.55, w: 0.04, h: 1.2, fill: { color: k.barColor } });
      slide.addText(k.label, { x: x + 0.15, y: 1.6, w: 1, h: 0.15, fontSize: 7, fontFace: "Inter", color: C.gray500, bold: true, charSpacing: 1, margin: 0 });
      slide.addText([
        { text: k.value, options: { fontSize: 18, fontFace: "Inter", color: C.dark, bold: true } },
        { text: k.unit, options: { fontSize: 9, color: C.gray500 } }
      ], { x: x + 0.15, y: 1.78, w: 1.2, h: 0.35, margin: 0 });
      slide.addShape(pres.shapes.OVAL, { x: x + 1.2, y: 1.65, w: 0.4, h: 0.4, fill: { color: k.iconBg } });
      slide.addImage({ data: k.icon, x: x + 1.3, y: 1.75, w: 0.2, h: 0.2 });
      slide.addText(k.desc, { x: x + 0.15, y: 2.2, w: 1.5, h: 0.2, fontSize: 8, fontFace: "Noto Sans JP", color: C.gray500, margin: 0 });
    });

    // Bar chart
    slide.addShape(pres.shapes.RECTANGLE, { x: 1.25, y: 3.0, w: 3.55, h: 1.85, fill: { color: C.white }, line: { color: C.gray200, width: 0.5 }, shadow: mkShadow() });
    slide.addText("INVESTMENT BREAKDOWN", { x: 1.45, y: 3.1, w: 3, h: 0.2, fontSize: 8, fontFace: "Inter", color: C.accent, bold: true, charSpacing: 1, margin: 0 });

    const bars = [
      { label: "車両購入", pct: 85, val: "600-900万", color: C.accent },
      { label: "登録・保険", pct: 15, val: "50万", color: C.warm },
      { label: "設備・備品", pct: 8, val: "20万", color: C.accent },
      { label: "Web・広告", pct: 5, val: "10万", color: C.warm },
    ];
    bars.forEach((b, i) => {
      const y = 3.45 + i * 0.33;
      slide.addText(b.label, { x: 1.45, y, w: 0.8, h: 0.2, fontSize: 8, fontFace: "Noto Sans JP", color: C.gray500, align: "right", margin: 0 });
      slide.addShape(pres.shapes.RECTANGLE, { x: 2.35, y: y + 0.02, w: 2.2, h: 0.18, fill: { color: C.gray100 } });
      slide.addShape(pres.shapes.RECTANGLE, { x: 2.35, y: y + 0.02, w: 2.2 * b.pct / 100, h: 0.18, fill: { color: b.color } });
      slide.addText(b.val, { x: 2.35, y: y, w: 2.2 * b.pct / 100, h: 0.2, fontSize: 6, fontFace: "Inter", color: C.white, bold: true, align: "right", margin: [0, 2, 0, 0] });
    });

    // Donut
    slide.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 3.0, w: 3.55, h: 1.85, fill: { color: C.white }, line: { color: C.gray200, width: 0.5 }, shadow: mkShadow() });
    slide.addText("TOTAL INVESTMENT", { x: 5.4, y: 3.1, w: 3, h: 0.2, fontSize: 8, fontFace: "Inter", color: C.accent, bold: true, charSpacing: 1, margin: 0 });

    // Use pie chart for total
    slide.addChart(pres.charts.DOUGHNUT, [{
      name: "投資", labels: ["車両・設備", "登録・広告"], values: [82, 18]
    }], {
      x: 5.6, y: 3.3, w: 1.8, h: 1.4,
      chartColors: [C.accent, C.warm],
      showLegend: false, showTitle: false,
      showPercent: false, showValue: false
    });
    slide.addText([
      { text: "TOTAL\n", options: { fontSize: 7, fontFace: "Inter", color: C.gray500 } },
      { text: "700-1,000\n", options: { fontSize: 14, fontFace: "Inter", color: C.dark, bold: true, breakLine: true } },
      { text: "万円", options: { fontSize: 8, color: C.gray500 } }
    ], { x: 5.95, y: 3.65, w: 1.1, h: 0.75, align: "center", valign: "middle", margin: 0 });

    // Legend
    slide.addShape(pres.shapes.OVAL, { x: 7.6, y: 3.8, w: 0.15, h: 0.15, fill: { color: C.accent } });
    slide.addText("車両・設備", { x: 7.8, y: 3.8, w: 0.8, h: 0.15, fontSize: 8, fontFace: "Noto Sans JP", color: C.gray600, margin: 0 });
    slide.addShape(pres.shapes.OVAL, { x: 7.6, y: 4.05, w: 0.15, h: 0.15, fill: { color: C.warm } });
    slide.addText("登録・広告", { x: 7.8, y: 4.05, w: 0.8, h: 0.15, fontSize: 8, fontFace: "Noto Sans JP", color: C.gray600, margin: 0 });
  }

  // ========== SLIDE 13: 収支モデル ==========
  {
    const slide = pres.addSlide();
    slide.background = { color: C.white };
    addHeader(slide, "Revenue Model", "収支モデル（月次・標準）", 13, C.warm);

    // Table
    const tableRows = [
      ["項目", "数値", "単位", "備考"],
      ["稼働日数", "12", "日 / 月", "週末+祝日中心"],
      ["平均単価", "15,000", "円 / 日", "日貸ベース"],
      ["月商", "180,000", "円 / 月", "12日 x 15,000円"],
      ["固定費", "100,000-140,000", "円 / 月", "保険・駐車場・ローン等"],
      ["変動費", "20,000", "円 / 月", "ガソリン・消耗品等"],
    ];

    const tableData = tableRows.map((row, ri) => {
      if (ri === 0) {
        return row.map(cell => ({ text: cell, options: { fill: { color: C.dark }, color: C.white, bold: true, fontSize: 8, fontFace: "Noto Sans JP", align: "center" } }));
      }
      const isRevenue = ri === 3;
      return row.map((cell, ci) => ({
        text: cell,
        options: {
          fill: { color: isRevenue ? C.blue50 : (ci === 0 ? C.gray50 : C.white) },
          color: isRevenue && ci <= 2 ? C.accent : (ci === 0 ? C.dark : C.gray600),
          bold: ci === 0 || ci === 1 || isRevenue,
          fontSize: isRevenue && ci === 1 ? 11 : 8,
          fontFace: ci === 1 ? "Inter" : "Noto Sans JP",
          align: ci === 0 ? "left" : "center"
        }
      }));
    });

    slide.addTable(tableData, {
      x: 1.25, y: 1.5, w: 7.5,
      colW: [1.8, 2, 1.4, 2.3],
      border: { pt: 0.5, color: C.gray200 },
      rowH: [0.35, 0.35, 0.35, 0.4, 0.35, 0.35]
    });

    // Revenue flow
    const flowItems = [
      { label: "REVENUE", value: "¥180,000", color: C.accent },
      { label: "COST", value: "¥120,000-160,000", color: C.red400 },
      { label: "PROFIT", value: "¥20,000-50,000", color: C.green500, bg: C.green50, textColor: C.green700 },
    ];

    slide.addShape(pres.shapes.RECTANGLE, { x: 1.25, y: 4.1, w: 2.2, h: 0.7, fill: { color: C.white }, shadow: mkShadow() });
    slide.addShape(pres.shapes.RECTANGLE, { x: 1.25, y: 4.1, w: 0.06, h: 0.7, fill: { color: C.accent } });
    slide.addText("REVENUE", { x: 1.45, y: 4.15, w: 1.5, h: 0.15, fontSize: 7, fontFace: "Inter", color: C.gray500, margin: 0 });
    slide.addText("¥180,000", { x: 1.45, y: 4.35, w: 1.8, h: 0.3, fontSize: 16, fontFace: "Inter", color: C.dark, bold: true, margin: 0 });

    slide.addText("−", { x: 3.55, y: 4.2, w: 0.4, h: 0.5, fontSize: 20, color: C.gray300, align: "center", valign: "middle", margin: 0 });

    slide.addShape(pres.shapes.RECTANGLE, { x: 4.0, y: 4.1, w: 2.4, h: 0.7, fill: { color: C.white }, shadow: mkShadow() });
    slide.addShape(pres.shapes.RECTANGLE, { x: 4.0, y: 4.1, w: 0.06, h: 0.7, fill: { color: C.red400 } });
    slide.addText("COST", { x: 4.2, y: 4.15, w: 1.5, h: 0.15, fontSize: 7, fontFace: "Inter", color: C.gray500, margin: 0 });
    slide.addText("¥120,000-160,000", { x: 4.2, y: 4.35, w: 2, h: 0.3, fontSize: 14, fontFace: "Inter", color: C.dark, bold: true, margin: 0 });

    slide.addText("＝", { x: 6.5, y: 4.2, w: 0.4, h: 0.5, fontSize: 20, color: C.gray300, align: "center", valign: "middle", margin: 0 });

    slide.addShape(pres.shapes.RECTANGLE, { x: 6.95, y: 4.1, w: 1.8, h: 0.7, fill: { color: C.green50 }, shadow: mkShadow() });
    slide.addShape(pres.shapes.RECTANGLE, { x: 6.95, y: 4.1, w: 0.06, h: 0.7, fill: { color: C.green500 } });
    slide.addText("PROFIT", { x: 7.15, y: 4.15, w: 1.5, h: 0.15, fontSize: 7, fontFace: "Inter", color: C.gray500, margin: 0 });
    slide.addText("¥20,000-50,000", { x: 7.15, y: 4.35, w: 1.5, h: 0.3, fontSize: 15, fontFace: "Inter", color: C.green700, bold: true, margin: 0 });
  }

  // ========== SLIDE 14: SECTION 4 DIVIDER ==========
  {
    const slide = pres.addSlide();
    slide.background = { color: C.gray50 };
    addSectionDivider(slide, 4, "Section 04", "成長戦略", "資金調達・リスク管理・\n中長期的な成長計画",
      ["資金調達", "リスク", "ロードマップ"], icons.rocket, 14, C.warm);

    addSectionRightItem(slide, icons.piggyBank, "Funding", "資金調達計画",
      "自己資金200万円 + 借入800〜1,000万円", 1.3, C.warm);
    addSectionRightItem(slide, icons.shieldRed, "Risk Management", "リスクと対策",
      "稼働率・事故・価格競争・資金繰りへの対策", 2.5, "EF4444");
    addSectionRightItem(slide, icons.road, "Roadmap", "成長ロードマップ",
      "1台 → 2台 → キャンピングカー特化・AI活用", 3.7, C.accent);
  }

  // ========== SLIDE 15: 資金調達計画 ==========
  {
    const slide = pres.addSlide();
    slide.background = { color: C.white };
    addHeader(slide, "Funding Plan", "資金調達計画", 15, C.warm);

    // Left: Self fund
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.25, y: 1.5, w: 1.1, h: 0.28, rectRadius: 0.05, fill: { color: C.accent } });
    slide.addText("SELF FUND", { x: 1.25, y: 1.5, w: 1.1, h: 0.28, fontSize: 7, fontFace: "Inter", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    slide.addText("自己資金", { x: 2.45, y: 1.5, w: 1.2, h: 0.28, fontSize: 12, fontFace: "Noto Sans JP", color: C.dark, bold: true, margin: 0 });

    slide.addShape(pres.shapes.RECTANGLE, { x: 1.25, y: 1.9, w: 3.5, h: 1.8, fill: { color: C.blue50 }, line: { color: "BFDBFE", width: 0.5 } });
    slide.addShape(pres.shapes.OVAL, { x: 1.5, y: 2.1, w: 0.55, h: 0.55, fill: { color: C.white }, shadow: mkShadow() });
    slide.addImage({ data: icons.wallet, x: 1.63, y: 2.22, w: 0.3, h: 0.3 });
    slide.addText([
      { text: "200", options: { fontSize: 24, fontFace: "Inter", color: C.dark, bold: true } },
      { text: "万円", options: { fontSize: 10, color: C.gray500 } }
    ], { x: 2.2, y: 2.15, w: 2, h: 0.4, margin: 0 });
    slide.addImage({ data: icons.checkCircle, x: 1.55, y: 2.85, w: 0.15, h: 0.15 });
    slide.addText("初期資金として準備済み", { x: 1.8, y: 2.85, w: 2.5, h: 0.2, fontSize: 9, fontFace: "Noto Sans JP", color: C.gray700, margin: 0 });
    slide.addImage({ data: icons.checkCircle, x: 1.55, y: 3.15, w: 0.15, h: 0.15 });
    slide.addText("副業収入からの積立", { x: 1.8, y: 3.15, w: 2.5, h: 0.2, fontSize: 9, fontFace: "Noto Sans JP", color: C.gray700, margin: 0 });

    // Right: Borrowing
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.25, y: 1.5, w: 1.2, h: 0.28, rectRadius: 0.05, fill: { color: C.warm } });
    slide.addText("BORROWING", { x: 5.25, y: 1.5, w: 1.2, h: 0.28, fontSize: 7, fontFace: "Inter", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    slide.addText("借入", { x: 6.55, y: 1.5, w: 1.2, h: 0.28, fontSize: 12, fontFace: "Noto Sans JP", color: C.dark, bold: true, margin: 0 });

    slide.addShape(pres.shapes.RECTANGLE, { x: 5.25, y: 1.9, w: 3.5, h: 1.8, fill: { color: C.orange50 }, line: { color: "FED7AA", width: 0.5 } });
    slide.addShape(pres.shapes.OVAL, { x: 5.5, y: 2.1, w: 0.55, h: 0.55, fill: { color: C.white }, shadow: mkShadow() });
    slide.addImage({ data: icons.university, x: 5.63, y: 2.22, w: 0.3, h: 0.3 });
    slide.addText([
      { text: "800-1,000", options: { fontSize: 22, fontFace: "Inter", color: C.dark, bold: true } },
      { text: "万円", options: { fontSize: 10, color: C.gray500 } }
    ], { x: 6.2, y: 2.15, w: 2.2, h: 0.4, margin: 0 });
    slide.addImage({ data: icons.checkCircleWarm, x: 5.55, y: 2.85, w: 0.15, h: 0.15 });
    slide.addText("車両購入資金", { x: 5.8, y: 2.85, w: 2.5, h: 0.2, fontSize: 9, fontFace: "Noto Sans JP", color: C.gray700, margin: 0 });
    slide.addImage({ data: icons.checkCircleWarm, x: 5.55, y: 3.15, w: 0.15, h: 0.15 });
    slide.addText("運転資金", { x: 5.8, y: 3.15, w: 2.5, h: 0.2, fontSize: 9, fontFace: "Noto Sans JP", color: C.gray700, margin: 0 });

    // Bottom bar
    slide.addShape(pres.shapes.RECTANGLE, { x: 1.25, y: 4.0, w: 7.5, h: 0.85, fill: { color: C.dark } });
    slide.addText("TOTAL FUNDING", { x: 1.5, y: 4.1, w: 1.5, h: 0.2, fontSize: 8, fontFace: "Inter", color: C.accent, bold: true, margin: 0 });
    slide.addText("資金調達総額", { x: 1.5, y: 4.3, w: 1.2, h: 0.2, fontSize: 8, fontFace: "Noto Sans JP", color: C.gray400, margin: 0 });
    slide.addShape(pres.shapes.LINE, { x: 3.2, y: 4.1, w: 0, h: 0.65, line: { color: C.gray700, width: 0.5 } });

    const fundItems = [
      { val: "1,000-1,200", unit: "万円", label: "調達総額" },
      { val: "17-20", unit: "%", label: "自己資金比率" },
      { val: "80-83", unit: "%", label: "借入比率" },
    ];
    fundItems.forEach((f, i) => {
      const x = 3.7 + i * 1.8;
      slide.addText([
        { text: f.val, options: { fontSize: 17, fontFace: "Inter", color: C.white, bold: true } },
        { text: f.unit, options: { fontSize: 9, color: C.white } }
      ], { x, y: 4.12, w: 1.6, h: 0.3, align: "center", margin: 0 });
      slide.addText(f.label, { x, y: 4.45, w: 1.6, h: 0.2, fontSize: 8, fontFace: "Noto Sans JP", color: C.gray400, align: "center", margin: 0 });
    });
  }

  // ========== SLIDE 16: リスクと対策 ==========
  {
    const slide = pres.addSlide();
    slide.background = { color: C.white };
    addHeader(slide, "Risk Management", "リスクと対策", 16, C.warm);

    const risks = [
      { title: "稼働率低下", sub: "需要不足による低稼働", badge: "中", badgeColor: C.yellow600, badgeBg: C.yellow50, icon: icons.chartLineYellow, iconBg: C.yellow50, mitigation: "副業スタートで固定費を抑制し、リスクを最小化" },
      { title: "事故・故障", sub: "車両損害・修理費用", badge: "高", badgeColor: "DC2626", badgeBg: C.red50, icon: icons.excTriangle, iconBg: C.red50, mitigation: "任意保険加入・定期点検の徹底" },
      { title: "価格競争", sub: "大手との価格競争激化", badge: "中", badgeColor: C.yellow600, badgeBg: C.yellow50, icon: icons.tags, iconBg: C.orange50, mitigation: "レジャー特化・品質訴求で差別化" },
      { title: "資金繰り", sub: "キャッシュフロー悪化", badge: "低", badgeColor: C.green700, badgeBg: C.green50, icon: icons.coins, iconBg: C.blue50, mitigation: "慎重な台数拡大で財務を安定化" },
    ];

    risks.forEach((r, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = 1.25 + col * 3.85;
      const y = 1.55 + row * 1.55;

      slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.55, h: 1.35, fill: { color: C.white }, line: { color: C.gray200, width: 0.5 }, shadow: mkShadow() });
      // Badge
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 2.8, y: y + 0.1, w: 0.5, h: 0.22, rectRadius: 0.04, fill: { color: r.badgeBg } });
      slide.addText(r.badge, { x: x + 2.8, y: y + 0.1, w: 0.5, h: 0.22, fontSize: 7, fontFace: "Noto Sans JP", color: r.badgeColor, bold: true, align: "center", valign: "middle", margin: 0 });
      // Icon
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.15, y: y + 0.15, w: 0.45, h: 0.45, rectRadius: 0.06, fill: { color: r.iconBg } });
      slide.addImage({ data: r.icon, x: x + 0.25, y: y + 0.25, w: 0.25, h: 0.25 });
      // Title
      slide.addText(r.title, { x: x + 0.75, y: y + 0.12, w: 2, h: 0.2, fontSize: 10, fontFace: "Noto Sans JP", color: C.dark, bold: true, margin: 0 });
      slide.addText(r.sub, { x: x + 0.75, y: y + 0.32, w: 2, h: 0.15, fontSize: 7, fontFace: "Noto Sans JP", color: C.gray500, margin: 0 });
      // Mitigation
      slide.addShape(pres.shapes.RECTANGLE, { x: x + 0.12, y: y + 0.7, w: 3.3, h: 0.5, fill: { color: C.green50 } });
      slide.addImage({ data: icons.shieldAlt, x: x + 0.22, y: y + 0.77, w: 0.15, h: 0.15 });
      slide.addText("軽減策", { x: x + 0.45, y: y + 0.75, w: 0.6, h: 0.18, fontSize: 7, fontFace: "Noto Sans JP", color: C.accent, bold: true, margin: 0 });
      slide.addText(r.mitigation, { x: x + 0.22, y: y + 0.97, w: 3.1, h: 0.18, fontSize: 8, fontFace: "Noto Sans JP", color: C.gray600, margin: 0 });
    });

    // Summary bar
    slide.addShape(pres.shapes.RECTANGLE, { x: 1.25, y: 4.75, w: 7.5, h: 0.3, fill: { color: C.gray50 }, line: { color: C.gray200, width: 0.5 } });
    slide.addImage({ data: icons.shieldAlt, x: 1.4, y: 4.8, w: 0.15, h: 0.15 });
    slide.addText("総合リスク評価", { x: 1.65, y: 4.78, w: 1.2, h: 0.2, fontSize: 8, fontFace: "Noto Sans JP", color: C.dark, bold: true, margin: 0 });
    // Risk dots
    slide.addShape(pres.shapes.OVAL, { x: 6.8, y: 4.83, w: 0.1, h: 0.1, fill: { color: "DC2626" } });
    slide.addText("高: 1", { x: 6.95, y: 4.8, w: 0.5, h: 0.15, fontSize: 7, fontFace: "Noto Sans JP", color: C.gray500, margin: 0 });
    slide.addShape(pres.shapes.OVAL, { x: 7.5, y: 4.83, w: 0.1, h: 0.1, fill: { color: "EAB308" } });
    slide.addText("中: 2", { x: 7.65, y: 4.8, w: 0.5, h: 0.15, fontSize: 7, fontFace: "Noto Sans JP", color: C.gray500, margin: 0 });
    slide.addShape(pres.shapes.OVAL, { x: 8.2, y: 4.83, w: 0.1, h: 0.1, fill: { color: C.green500 } });
    slide.addText("低: 1", { x: 8.35, y: 4.8, w: 0.5, h: 0.15, fontSize: 7, fontFace: "Noto Sans JP", color: C.gray500, margin: 0 });
  }

  // ========== SLIDE 17: 成長ロードマップ ==========
  {
    const slide = pres.addSlide();
    slide.background = { color: C.white };
    addHeader(slide, "Growth Roadmap", "成長ロードマップ", 17, C.warm);

    // Timeline bar
    slide.addShape(pres.shapes.LINE, { x: 1.5, y: 1.85, w: 7, h: 0, line: { color: C.gray200, width: 2 } });
    // Timeline dots
    const phases = [
      { x: 2.5, label: "Year 1", color: C.accent, icon: icons.flag },
      { x: 5, label: "Year 2-3", color: C.warm, icon: icons.arrowUp },
      { x: 7.5, label: "Future", color: C.dark, icon: icons.rocket },
    ];
    phases.forEach(p => {
      slide.addShape(pres.shapes.OVAL, { x: p.x - 0.25, y: 1.6, w: 0.5, h: 0.5, fill: { color: p.color }, shadow: mkShadow() });
      slide.addImage({ data: p.icon, x: p.x - 0.12, y: 1.73, w: 0.24, h: 0.24 });
      slide.addText(p.label, { x: p.x - 0.5, y: 2.15, w: 1, h: 0.2, fontSize: 8, fontFace: "Inter", color: p.color, bold: true, align: "center", margin: 0 });
    });

    // Phase cards
    const phaseCards = [
      {
        title: "安定運営の確立", phase: "Phase 1", phaseLabel: "1年目", color: C.accent,
        icon: icons.seedling, phaseBg: C.blue50,
        items: ["1台で安定運営", "稼働率の向上", "口コミ・リピーター獲得", "オペレーション最適化"],
        footer: "1台体制", footerIcon: icons.carAccent
      },
      {
        title: "事業拡大", phase: "Phase 2", phaseLabel: "2〜3年目", color: C.warm,
        icon: icons.expand, phaseBg: C.orange50,
        items: ["2台目導入", "車種ラインナップ拡充", "法人契約の開拓", "収益基盤の強化"],
        footer: "2台体制", footerIcon: icons.carWarm
      },
      {
        title: "特化・AI活用", phase: "Phase 3", phaseLabel: "将来", color: C.dark,
        icon: icons.rocketDark, phaseBg: C.gray100,
        items: ["キャンピングカー特化", "AI需要予測・価格最適化", "予約最適化システム", "高稼働・高収益モデル"],
        footer: "AI駆動型運営", footerIcon: icons.robot
      },
    ];

    phaseCards.forEach((pc, i) => {
      const x = 1.25 + i * 2.7;
      const y = 2.5;
      slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.4, h: 2.55, fill: { color: C.white }, line: { color: C.gray100, width: 0.5 }, shadow: mkShadow() });
      slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.4, h: 0.06, fill: { color: pc.color } });
      // Phase icon
      slide.addShape(pres.shapes.OVAL, { x: x + 0.15, y: y + 0.2, w: 0.5, h: 0.5, fill: { color: pc.color } });
      slide.addImage({ data: pc.icon, x: x + 0.27, y: y + 0.32, w: 0.25, h: 0.25 });
      // Phase badge
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.8, y: y + 0.2, w: 0.7, h: 0.22, rectRadius: 0.04, fill: { color: pc.phaseBg } });
      slide.addText(pc.phase, { x: x + 0.8, y: y + 0.2, w: 0.7, h: 0.22, fontSize: 7, fontFace: "Inter", color: pc.color, bold: true, align: "center", valign: "middle", margin: 0 });
      slide.addText(pc.phaseLabel, { x: x + 0.8, y: y + 0.45, w: 1, h: 0.15, fontSize: 7, fontFace: "Noto Sans JP", color: C.gray400, margin: 0 });
      // Title
      slide.addText(pc.title, { x: x + 0.15, y: y + 0.8, w: 2.1, h: 0.25, fontSize: 11, fontFace: "Noto Sans JP", color: C.dark, bold: true, margin: 0 });
      // Items
      pc.items.forEach((item, j) => {
        const iy = y + 1.15 + j * 0.25;
        const checkIcon = pc.color === C.accent ? icons.checkCircle : (pc.color === C.warm ? icons.checkCircleWarm : icons.checkCircleDark);
        slide.addImage({ data: checkIcon, x: x + 0.2, y: iy + 0.02, w: 0.13, h: 0.13 });
        slide.addText(item, { x: x + 0.42, y: iy, w: 1.8, h: 0.2, fontSize: 8, fontFace: "Noto Sans JP", color: C.gray600, margin: 0 });
      });
      // Footer
      slide.addShape(pres.shapes.LINE, { x: x + 0.15, y: y + 2.2, w: 2.1, h: 0, line: { color: C.gray100, width: 0.5 } });
      slide.addImage({ data: pc.footerIcon, x: x + 0.2, y: y + 2.3, w: 0.15, h: 0.15 });
      slide.addText(pc.footer, { x: x + 0.42, y: y + 2.3, w: 1.5, h: 0.15, fontSize: 7, fontFace: "Noto Sans JP", color: C.gray400, margin: 0 });
    });
  }

  // ========== SLIDE 18: CLOSING ==========
  {
    const slide = pres.addSlide();
    slide.background = { color: C.dark };
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H, fill: { color: C.darkBg, transparency: 30 } });
    // Accent bar
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.06, h: H, fill: { color: C.accent } });
    // Decorative
    slide.addShape(pres.shapes.OVAL, { x: 1, y: 0.8, w: 2.2, h: 2.2, fill: { color: C.accent, transparency: 90 } });
    slide.addShape(pres.shapes.OVAL, { x: 7, y: 3.5, w: 2.8, h: 2.8, fill: { color: C.warm, transparency: 95 } });

    // Icon
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 4.25, y: 0.6, w: 1.5, h: 1.5, rectRadius: 0.18, fill: { color: C.accent }, shadow: mkShadow() });
    slide.addImage({ data: icons.carWhite, x: 4.65, y: 0.9, w: 0.7, h: 0.7 });

    // Thank you
    slide.addText("THANK YOU", {
      x: 1, y: 2.2, w: 8, h: 0.3, fontSize: 10, fontFace: "Inter",
      color: C.accent, bold: true, charSpacing: 5, align: "center", margin: 0
    });
    slide.addText("ご清聴ありがとうございました", {
      x: 0.5, y: 2.5, w: 9, h: 0.7, fontSize: 30, fontFace: "Noto Sans JP",
      color: C.white, bold: true, align: "center", margin: 0
    });
    // Line
    slide.addShape(pres.shapes.RECTANGLE, { x: 4.6, y: 3.3, w: 0.8, h: 0.04, fill: { color: C.accent } });
    // Subtitle
    slide.addText("AIレンタカー 事業計画書", {
      x: 1, y: 3.45, w: 8, h: 0.35, fontSize: 14, fontFace: "Noto Sans JP",
      color: C.gray300, align: "center", margin: 0
    });

    // Contact box
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 2.8, y: 4.0, w: 4.4, h: 0.9, rectRadius: 0.1, fill: { color: C.white, transparency: 90 } });
    slide.addText("CONTACT INFORMATION", {
      x: 2.8, y: 4.05, w: 4.4, h: 0.2, fontSize: 8, fontFace: "Inter",
      color: C.gray400, charSpacing: 2, align: "center", margin: 0
    });
    slide.addImage({ data: icons.envelope, x: 3.6, y: 4.4, w: 0.2, h: 0.2 });
    slide.addText("info@ai-rental-car.jp", {
      x: 3.85, y: 4.4, w: 1.8, h: 0.2, fontSize: 9, fontFace: "Inter", color: C.gray300, margin: 0
    });
    slide.addImage({ data: icons.mapMarker, x: 5.9, y: 4.4, w: 0.2, h: 0.2 });
    slide.addText("千葉県", {
      x: 6.15, y: 4.4, w: 0.8, h: 0.2, fontSize: 9, fontFace: "Noto Sans JP", color: C.gray300, margin: 0
    });

    // Footer
    slide.addText("Confidential", {
      x: 1.25, y: 5.2, w: 3, h: 0.3, fontSize: 8, fontFace: "Noto Sans JP", color: C.gray600, margin: 0
    });
    slide.addText([
      { text: "Page ", options: { fontSize: 8, color: C.gray600, fontFace: "Inter" } },
      { text: "18", options: { fontSize: 10, color: C.accent, bold: true, fontFace: "Inter" } }
    ], { x: 8, y: 5.2, w: 1, h: 0.3, align: "right", margin: 0 });
  }

  // ========== WRITE FILE ==========
  await pres.writeFile({ fileName: "/Users/nogataka/dev/slide-test/output/slide-page01/presentation.pptx" });
  console.log("PPTX generated successfully!");
}

main().catch(err => { console.error(err); process.exit(1); });
