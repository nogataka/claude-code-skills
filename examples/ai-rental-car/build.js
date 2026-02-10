const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

// Icon imports
const {
  FaCar, FaRobot, FaGem, FaUsers, FaChartLine, FaCog, FaGlobe,
  FaWrench, FaYenSign, FaCalculator, FaHandshake, FaExclamationTriangle,
  FaRoad, FaCheckCircle, FaStar, FaMapMarkerAlt, FaCalendarAlt,
  FaChild, FaCarSide, FaCampground, FaInstagram, FaGoogle, FaLaptop,
  FaUserTie, FaBroom, FaToolbox, FaMoneyBillWave, FaPiggyBank,
  FaUniversity, FaShieldAlt, FaSearchDollar, FaBalanceScale, FaArrowUp,
  FaFlag, FaRocket, FaBrain, FaLightbulb
} = require("react-icons/fa");

function renderIconSvg(Icon, color = "#000000", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(Icon, { color, size: String(size) })
  );
}
async function icon64(Icon, color, size = 256) {
  const svg = renderIconSvg(Icon, color, size);
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// Colors (no # for pptxgenjs)
const C = {
  dark: "1A1A2E",
  dark2: "16213E",
  teal: "16C79A",
  amber: "F5A623",
  white: "FFFFFF",
  offWhite: "F8F9FA",
  gray100: "F1F3F5",
  gray200: "DEE2E6",
  gray300: "CED4DA",
  gray400: "ADB5BD",
  gray500: "6C757D",
  gray600: "495057",
  gray700: "343A40",
  tealLight: "E6FAF3",
  tealDark: "0F9B74",
  amberLight: "FFF8EC",
  amberDark: "D48E1A",
  red50: "FFF5F5",
  red500: "E53E3E",
  blue50: "EBF5FF",
  blue500: "3B82F6",
};
const FONT_H = "Georgia";
const FONT_B = "Calibri";

const makeShadow = () => ({
  type: "outer", color: "000000", blur: 5, offset: 2, angle: 135, opacity: 0.1,
});

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10" x 5.625"
  pres.author = "AIレンタカー";
  pres.title = "AIレンタカー｜事業計画書";

  // Pre-render icons
  const I = {};
  const defs = [
    ["car_w", FaCar, "#FFFFFF"], ["car_teal", FaCar, "#16C79A"],
    ["robot_w", FaRobot, "#FFFFFF"], ["robot_teal", FaRobot, "#16C79A"],
    ["gem_teal", FaGem, "#16C79A"], ["gem_amber", FaGem, "#F5A623"],
    ["users_teal", FaUsers, "#16C79A"], ["users_amber", FaUsers, "#F5A623"],
    ["chart_teal", FaChartLine, "#16C79A"], ["chart_amber", FaChartLine, "#F5A623"],
    ["cog_teal", FaCog, "#16C79A"], ["globe_teal", FaGlobe, "#16C79A"],
    ["wrench_teal", FaWrench, "#16C79A"], ["yen_teal", FaYenSign, "#16C79A"],
    ["calc_teal", FaCalculator, "#16C79A"], ["hand_teal", FaHandshake, "#16C79A"],
    ["warn_amber", FaExclamationTriangle, "#F5A623"],
    ["road_teal", FaRoad, "#16C79A"], ["check_teal", FaCheckCircle, "#16C79A"],
    ["check_w", FaCheckCircle, "#FFFFFF"],
    ["star_amber", FaStar, "#F5A623"], ["star_w", FaStar, "#FFFFFF"],
    ["map_teal", FaMapMarkerAlt, "#16C79A"],
    ["calendar_teal", FaCalendarAlt, "#16C79A"],
    ["child_gray", FaChild, "#6C757D"],
    ["carside_teal", FaCarSide, "#16C79A"],
    ["camp_amber", FaCampground, "#F5A623"],
    ["insta_teal", FaInstagram, "#16C79A"],
    ["google_teal", FaGoogle, "#16C79A"],
    ["laptop_teal", FaLaptop, "#16C79A"],
    ["usertie_teal", FaUserTie, "#16C79A"],
    ["broom_gray", FaBroom, "#6C757D"],
    ["toolbox_gray", FaToolbox, "#6C757D"],
    ["money_teal", FaMoneyBillWave, "#16C79A"],
    ["money_amber", FaMoneyBillWave, "#F5A623"],
    ["piggy_teal", FaPiggyBank, "#16C79A"],
    ["univ_teal", FaUniversity, "#16C79A"],
    ["shield_teal", FaShieldAlt, "#16C79A"],
    ["shield_amber", FaShieldAlt, "#F5A623"],
    ["search_teal", FaSearchDollar, "#16C79A"],
    ["balance_amber", FaBalanceScale, "#F5A623"],
    ["arrowup_teal", FaArrowUp, "#16C79A"],
    ["flag_w", FaFlag, "#FFFFFF"], ["flag_amber", FaFlag, "#F5A623"],
    ["rocket_teal", FaRocket, "#16C79A"], ["rocket_w", FaRocket, "#FFFFFF"],
    ["brain_teal", FaBrain, "#16C79A"],
    ["light_amber", FaLightbulb, "#F5A623"],
    ["car_dark", FaCar, "#1A1A2E"],
  ];
  console.log("Rendering icons...");
  for (const [k, comp, c] of defs) I[k] = await icon64(comp, c);
  console.log("Icons ready.");

  // ============================================
  // Helper: footer for content slides (2-11)
  // ============================================
  function footer(s, num) {
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 5.275, w: 10, h: 0.35, fill: { color: C.white },
    });
    s.addShape(pres.shapes.LINE, {
      x: 0.5, y: 5.275, w: 9, h: 0, line: { color: C.gray200, width: 0.5 },
    });
    s.addImage({ data: I.car_dark, x: 0.5, y: 5.33, w: 0.18, h: 0.18 });
    s.addText("AIレンタカー｜事業計画書", {
      x: 0.75, y: 5.3, w: 3, h: 0.25,
      fontFace: FONT_B, fontSize: 7, color: C.gray400, valign: "middle", margin: 0,
    });
    s.addText(String(num), {
      x: 8.8, y: 5.3, w: 0.7, h: 0.26,
      fontFace: FONT_H, fontSize: 10, bold: true, color: C.teal, align: "right", valign: "middle", margin: 0,
    });
  }

  // Helper: slide header
  function header(s, enLabel, jpTitle, iconData) {
    if (iconData) {
      s.addShape(pres.shapes.OVAL, {
        x: 0.5, y: 0.35, w: 0.55, h: 0.55, fill: { color: C.tealLight },
      });
      s.addImage({ data: iconData, x: 0.58, y: 0.43, w: 0.38, h: 0.38 });
    }
    const tx = iconData ? 1.2 : 0.5;
    s.addText(enLabel, {
      x: tx, y: 0.32, w: 5, h: 0.2,
      fontFace: FONT_B, fontSize: 7, color: C.teal, charSpacing: 3, margin: 0,
    });
    s.addText(jpTitle, {
      x: tx, y: 0.52, w: 6, h: 0.4,
      fontFace: FONT_H, fontSize: 20, bold: true, color: C.dark, margin: 0,
    });
    s.addShape(pres.shapes.LINE, {
      x: 0.5, y: 1.05, w: 9, h: 0, line: { color: C.gray200, width: 0.5 },
    });
  }

  // ============================================
  // SLIDE 1: Cover
  // ============================================
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };

    // Decorative shapes
    s.addShape(pres.shapes.OVAL, { x: 7.5, y: -1, w: 5, h: 5, fill: { color: C.dark2 } });
    s.addShape(pres.shapes.OVAL, { x: -1.5, y: 3.5, w: 4, h: 4, fill: { color: C.teal, transparency: 90 } });

    // Left accent bar
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.08, h: 5.625, fill: { color: C.teal } });

    // Car icon circle
    s.addShape(pres.shapes.OVAL, { x: 0.7, y: 1.2, w: 1.0, h: 1.0, fill: { color: C.teal } });
    s.addImage({ data: I.car_w, x: 0.88, y: 1.38, w: 0.65, h: 0.65 });

    // Title
    s.addText("AI レンタカー", {
      x: 0.7, y: 2.4, w: 6, h: 0.7,
      fontFace: FONT_H, fontSize: 42, bold: true, color: C.white, margin: 0,
    });
    s.addText("事業計画書", {
      x: 0.7, y: 3.05, w: 5, h: 0.55,
      fontFace: FONT_H, fontSize: 28, color: C.teal, margin: 0,
    });

    // Tagline
    s.addText("個人事業 × AI活用で実現する、新しいレンタカー体験", {
      x: 0.7, y: 3.75, w: 6, h: 0.3,
      fontFace: FONT_B, fontSize: 11, color: C.gray400, margin: 0,
    });

    // Bottom info
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 5.0, w: 10, h: 0.625,
      fill: { color: "000000", transparency: 60 },
    });
    s.addText("開業予定地：千葉県（首都圏・房総エリア）", {
      x: 0.7, y: 5.1, w: 4, h: 0.3,
      fontFace: FONT_B, fontSize: 9, color: C.gray400, margin: 0,
    });
    s.addText("Confidential", {
      x: 7, y: 5.1, w: 2.5, h: 0.3,
      fontFace: FONT_B, fontSize: 9, color: C.gray500, align: "right", margin: 0,
    });
  }

  // ============================================
  // SLIDE 2: 事業概要
  // ============================================
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };
    header(s, "BUSINESS OVERVIEW", "1. 事業概要", I.car_teal);

    // 4 info cards in 2x2
    const cards = [
      { label: "事業名", value: "AIレンタカー", icon: I.robot_teal },
      { label: "事業形態", value: "個人事業主\n（副業スタート）", icon: I.usertie_teal },
      { label: "事業内容", value: "普通自動車のレンタカー事業\n（日貸・時間貸）", icon: I.carside_teal },
      { label: "開業予定地", value: "千葉県\n（首都圏・房総エリア）", icon: I.map_teal },
    ];
    const cw = 4.1, ch = 1.4, gx = 0.3, gy = 0.25;
    const sx = 0.55, sy = 1.2;
    cards.forEach((c, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const cx = sx + col * (cw + gx), cy = sy + row * (ch + gy);
      s.addShape(pres.shapes.RECTANGLE, {
        x: cx, y: cy, w: cw, h: ch, fill: { color: C.white }, shadow: makeShadow(),
      });
      s.addShape(pres.shapes.RECTANGLE, {
        x: cx, y: cy, w: 0.06, h: ch, fill: { color: C.teal },
      });
      s.addShape(pres.shapes.OVAL, {
        x: cx + 0.25, y: cy + 0.3, w: 0.55, h: 0.55, fill: { color: C.tealLight },
      });
      s.addImage({ data: c.icon, x: cx + 0.35, y: cy + 0.4, w: 0.35, h: 0.35 });
      s.addText(c.label, {
        x: cx + 0.95, y: cy + 0.2, w: 2.8, h: 0.25,
        fontFace: FONT_B, fontSize: 8, bold: true, color: C.teal, margin: 0,
      });
      s.addText(c.value, {
        x: cx + 0.95, y: cy + 0.5, w: 2.9, h: 1.0,
        fontFace: FONT_H, fontSize: 14, bold: true, color: C.dark, margin: 0,
      });
    });

    // Bottom callout
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.55, y: 4.65, w: 8.5, h: 0.45,
      fill: { color: C.dark },
    });
    s.addImage({ data: I.camp_amber, x: 0.75, y: 4.72, w: 0.28, h: 0.28 });
    s.addText("将来構想：キャンピングカーを含む車両ラインナップの拡大", {
      x: 1.1, y: 4.65, w: 7, h: 0.45,
      fontFace: FONT_B, fontSize: 10, color: C.amber, valign: "middle", margin: 0,
    });
    footer(s, 2);
  }

  // ============================================
  // SLIDE 3: 提供価値
  // ============================================
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };
    header(s, "VALUE PROPOSITION", "2. 提供価値", I.gem_teal);

    const items = [
      { icon: I.star_amber, title: "レジャー最適化", desc: "週末・レジャー利用に最適化した車両ラインナップを提供" },
      { icon: I.check_teal, title: "清潔・高品質", desc: "徹底した車両管理で清潔・高品質な状態を維持" },
      { icon: I.usertie_teal, title: "柔軟な対応", desc: "個人運営ならではの柔軟な受渡・きめ細かい対応" },
      { icon: I.brain_teal, title: "AI活用運営", desc: "需要予測・価格最適化・予約最適化による高稼働運営（将来）" },
    ];
    items.forEach((item, i) => {
      const cy = 1.2 + i * 0.95;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: cy, w: 8.8, h: 0.8, fill: { color: C.white }, shadow: makeShadow(),
      });
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: cy, w: 0.06, h: 0.8, fill: { color: i === 3 ? C.amber : C.teal },
      });
      s.addShape(pres.shapes.OVAL, {
        x: 0.85, y: cy + 0.15, w: 0.5, h: 0.5, fill: { color: i === 3 ? C.amberLight : C.tealLight },
      });
      s.addImage({ data: item.icon, x: 0.95, y: cy + 0.25, w: 0.3, h: 0.3 });
      s.addText(item.title, {
        x: 1.55, y: cy + 0.1, w: 3, h: 0.3,
        fontFace: FONT_H, fontSize: 14, bold: true, color: C.dark, margin: 0,
      });
      s.addText(item.desc, {
        x: 1.55, y: cy + 0.42, w: 7.5, h: 0.3,
        fontFace: FONT_B, fontSize: 10, color: C.gray600, margin: 0,
      });
    });
    footer(s, 3);
  }

  // ============================================
  // SLIDE 4: ターゲット顧客
  // ============================================
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };
    header(s, "TARGET CUSTOMERS", "3. ターゲット顧客", I.users_teal);

    // Primary targets (left 2 cards) — expanded to fill vertical space
    const primaries = [
      { icon: I.users_amber, title: "首都圏在住の個人・ファミリー", desc: "車を持たない都心部居住者が主要顧客層。週末の外出・旅行需要が高い" },
      { icon: I.calendar_teal, title: "週末・連休の観光・アウトドア利用者", desc: "房総エリアへのドライブ、キャンプ、海水浴など季節イベントの需要" },
    ];
    primaries.forEach((p, i) => {
      const cy = 1.2 + i * 1.85;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: cy, w: 5.2, h: 1.65, fill: { color: C.white }, shadow: makeShadow(),
      });
      s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: cy, w: 0.06, h: 1.65, fill: { color: C.teal } });
      s.addShape(pres.shapes.OVAL, { x: 0.85, y: cy + 0.35, w: 0.65, h: 0.65, fill: { color: i === 0 ? C.amberLight : C.tealLight } });
      s.addImage({ data: p.icon, x: 0.97, y: cy + 0.47, w: 0.4, h: 0.4 });
      s.addText(p.title, {
        x: 1.7, y: cy + 0.25, w: 3.8, h: 0.35,
        fontFace: FONT_H, fontSize: 14, bold: true, color: C.dark, margin: 0,
      });
      s.addText(p.desc, {
        x: 1.7, y: cy + 0.7, w: 3.95, h: 0.6,
        fontFace: FONT_B, fontSize: 10, color: C.gray600, margin: 0,
      });
    });

    // Secondary targets (right column) — expanded to match
    s.addShape(pres.shapes.RECTANGLE, {
      x: 6.0, y: 1.2, w: 3.4, h: 3.7, fill: { color: C.dark },
    });
    s.addText("SECONDARY", {
      x: 6.2, y: 1.4, w: 2, h: 0.2,
      fontFace: FONT_B, fontSize: 7, color: C.teal, charSpacing: 3, margin: 0,
    });
    s.addText("将来ターゲット", {
      x: 6.2, y: 1.65, w: 3, h: 0.35,
      fontFace: FONT_H, fontSize: 15, bold: true, color: C.white, margin: 0,
    });

    const secs = [
      { label: "マイカー非保有層", desc: "維持費を避けつつ必要時のみ利用" },
      { label: "法人・長期レンタル", desc: "代車・業務用途での安定需要" },
    ];
    secs.forEach((sec, i) => {
      const sy = 2.3 + i * 1.05;
      s.addImage({ data: I.check_w, x: 6.2, y: sy + 0.05, w: 0.25, h: 0.25 });
      s.addText(sec.label, {
        x: 6.55, y: sy, w: 2.5, h: 0.3,
        fontFace: FONT_B, fontSize: 11, bold: true, color: C.white, margin: 0,
      });
      s.addText(sec.desc, {
        x: 6.55, y: sy + 0.35, w: 2.7, h: 0.3,
        fontFace: FONT_B, fontSize: 9, color: C.gray400, margin: 0,
      });
    });

    footer(s, 4);
  }

  // ============================================
  // SLIDE 5: 市場・競合環境
  // ============================================
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };
    header(s, "MARKET & COMPETITION", "4. 市場・競合環境", I.chart_teal);

    // Market overview (top)
    const mY = 1.2;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: mY, w: 8.8, h: 0.85, fill: { color: C.tealLight },
    });
    s.addImage({ data: I.chart_amber, x: 0.8, y: mY + 0.17, w: 0.4, h: 0.4 });
    s.addText("市場環境", {
      x: 1.35, y: mY + 0.08, w: 2, h: 0.3,
      fontFace: FONT_H, fontSize: 14, bold: true, color: C.dark, margin: 0,
    });
    s.addText("観光回復・アウトドア需要増加により、レンタカー市場は安定成長中", {
      x: 1.35, y: mY + 0.42, w: 7.5, h: 0.3,
      fontFace: FONT_B, fontSize: 10, color: C.gray600, margin: 0,
    });

    // Competition comparison (3 columns)
    const compY = 2.3;
    s.addText("競合比較", {
      x: 0.6, y: compY, w: 3, h: 0.3,
      fontFace: FONT_H, fontSize: 14, bold: true, color: C.dark, margin: 0,
    });

    const comps = [
      { label: "大手レンタカー", traits: "画一的なサービス\n高価格帯\n手続きが煩雑", color: C.gray500, bg: C.gray100 },
      { label: "地場中小", traits: "品質にばらつき\n知名度が低い\n予約がしにくい", color: C.gray500, bg: C.gray100 },
      { label: "AIレンタカー（当社）", traits: "用途特化\n品質重視\n柔軟な運営", color: C.tealDark, bg: C.tealLight },
    ];
    comps.forEach((comp, i) => {
      const cx = 0.6 + i * 3.1;
      const isUs = i === 2;
      s.addShape(pres.shapes.RECTANGLE, {
        x: cx, y: compY + 0.45, w: 2.85, h: 2.2,
        fill: { color: isUs ? C.tealLight : C.white },
        shadow: makeShadow(),
      });
      s.addShape(pres.shapes.RECTANGLE, {
        x: cx, y: compY + 0.45, w: 2.85, h: 0.05,
        fill: { color: isUs ? C.teal : C.gray300 },
      });
      if (isUs) {
        s.addImage({ data: I.star_amber, x: cx + 2.45, y: compY + 0.6, w: 0.22, h: 0.22 });
      }
      s.addText(comp.label, {
        x: cx + 0.15, y: compY + 0.6, w: 2.3, h: 0.3,
        fontFace: FONT_H, fontSize: 11, bold: true, color: isUs ? C.tealDark : C.dark, margin: 0,
      });
      s.addText(comp.traits, {
        x: cx + 0.15, y: compY + 1.0, w: 2.5, h: 1.4,
        fontFace: FONT_B, fontSize: 10, color: comp.color, margin: 0,
      });
    });

    footer(s, 5);
  }

  // ============================================
  // SLIDE 6: サービス内容
  // ============================================
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };
    header(s, "SERVICE DETAILS", "5. サービス内容", I.cog_teal);

    // Main services (left) — expanded to fill vertical space
    s.addText("基本サービス", {
      x: 0.6, y: 1.2, w: 3, h: 0.25,
      fontFace: FONT_B, fontSize: 8, bold: true, color: C.teal, charSpacing: 2, margin: 0,
    });
    const services = [
      { icon: I.car_teal, text: "レンタカー（日貸・時間貸）" },
      { icon: I.calendar_teal, text: "週末・連休向けプラン" },
    ];
    services.forEach((svc, i) => {
      const sy = 1.55 + i * 0.85;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: sy, w: 4.3, h: 0.7, fill: { color: C.white }, shadow: makeShadow(),
      });
      s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: sy, w: 0.06, h: 0.7, fill: { color: C.teal } });
      s.addImage({ data: svc.icon, x: 0.85, y: sy + 0.2, w: 0.3, h: 0.3 });
      s.addText(svc.text, {
        x: 1.3, y: sy, w: 3.3, h: 0.7,
        fontFace: FONT_B, fontSize: 11, bold: true, color: C.dark, valign: "middle", margin: 0,
      });
    });

    // Options
    s.addText("オプション", {
      x: 0.6, y: 3.5, w: 3, h: 0.25,
      fontFace: FONT_B, fontSize: 8, bold: true, color: C.amber, charSpacing: 2, margin: 0,
    });
    const opts = ["チャイルドシート", "ETC車載器", "アウトドア用品"];
    opts.forEach((opt, i) => {
      const ox = 0.6 + i * 1.5;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: ox, y: 3.85, w: 1.35, h: 0.5, rectRadius: 0.06,
        fill: { color: C.amberLight },
      });
      s.addText(opt, {
        x: ox, y: 3.85, w: 1.35, h: 0.5,
        fontFace: FONT_B, fontSize: 8, color: C.amberDark, align: "center", valign: "middle", margin: 0,
      });
    });

    // Future (right panel) — expanded height
    s.addShape(pres.shapes.RECTANGLE, {
      x: 5.2, y: 1.2, w: 4.2, h: 3.85, fill: { color: C.dark },
    });
    s.addText("FUTURE PLAN", {
      x: 5.4, y: 1.4, w: 2, h: 0.2,
      fontFace: FONT_B, fontSize: 7, color: C.teal, charSpacing: 3, margin: 0,
    });
    s.addText("将来構想", {
      x: 5.4, y: 1.65, w: 3, h: 0.4,
      fontFace: FONT_H, fontSize: 17, bold: true, color: C.white, margin: 0,
    });

    const futures = [
      { icon: I.camp_amber, text: "キャンピングカー\nラインナップ追加" },
      { icon: I.hand_teal, text: "法人契約の獲得" },
      { icon: I.calendar_teal, text: "長期レンタルプラン" },
    ];
    futures.forEach((f, i) => {
      const fy = 2.3 + i * 0.8;
      s.addImage({ data: f.icon, x: 5.4, y: fy + 0.08, w: 0.32, h: 0.32 });
      s.addText(f.text, {
        x: 5.85, y: fy, w: 3.2, h: 0.55,
        fontFace: FONT_B, fontSize: 11, color: C.white, margin: 0,
      });
    });

    footer(s, 6);
  }

  // ============================================
  // SLIDE 7: 集客・販売チャネル
  // ============================================
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };
    header(s, "SALES CHANNELS", "6. 集客・販売チャネル", I.globe_teal);

    const channels = [
      { icon: I.laptop_teal, title: "自社Webサイト", desc: "予約・問い合わせの受付\nオンライン完結で24h対応", bg: C.tealLight },
      { icon: I.search_teal, title: "レンタカー比較サイト", desc: "大手比較サイトへの掲載\n新規顧客の流入経路", bg: C.white },
      { icon: I.google_teal, title: "Googleマップ・口コミ", desc: "ローカルSEO対策\n実績と信頼の可視化", bg: C.white },
      { icon: I.insta_teal, title: "SNS（Instagram等）", desc: "車両・利用シーンの発信\nブランド認知度向上", bg: C.white },
    ];
    const chW = 2.0, chH = 3.8, chGap = 0.2;
    const chStartX = 0.7;
    channels.forEach((ch, i) => {
      const cx = chStartX + i * (chW + chGap);
      s.addShape(pres.shapes.RECTANGLE, {
        x: cx, y: 1.2, w: chW, h: chH, fill: { color: ch.bg }, shadow: makeShadow(),
      });
      s.addShape(pres.shapes.RECTANGLE, {
        x: cx, y: 1.2, w: chW, h: 0.05, fill: { color: C.teal },
      });
      // Number badge
      s.addShape(pres.shapes.OVAL, {
        x: cx + 0.15, y: 1.5, w: 0.45, h: 0.45, fill: { color: C.teal },
      });
      s.addText(String(i + 1), {
        x: cx + 0.15, y: 1.5, w: 0.45, h: 0.45,
        fontFace: FONT_H, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle", margin: 0,
      });
      // Icon
      s.addShape(pres.shapes.OVAL, {
        x: cx + (chW - 0.75) / 2, y: 2.25, w: 0.75, h: 0.75, fill: { color: C.tealLight },
      });
      s.addImage({ data: ch.icon, x: cx + (chW - 0.45) / 2, y: 2.4, w: 0.45, h: 0.45 });
      // Text
      s.addText(ch.title, {
        x: cx + 0.1, y: 3.2, w: chW - 0.2, h: 0.5,
        fontFace: FONT_H, fontSize: 12, bold: true, color: C.dark, align: "center", valign: "middle", margin: 0,
      });
      s.addText(ch.desc, {
        x: cx + 0.1, y: 3.8, w: chW - 0.2, h: 0.65,
        fontFace: FONT_B, fontSize: 9, color: C.gray500, align: "center", margin: 0,
      });
    });

    footer(s, 7);
  }

  // ============================================
  // SLIDE 8: オペレーション
  // ============================================
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };
    header(s, "OPERATIONS", "7. オペレーション", I.wrench_teal);

    // Center: org/flow
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 1.2, w: 5.5, h: 3.7, fill: { color: C.white }, shadow: makeShadow(),
    });
    s.addText("運営体制", {
      x: 0.8, y: 1.35, w: 2, h: 0.3,
      fontFace: FONT_H, fontSize: 14, bold: true, color: C.dark, margin: 0,
    });

    // Big "1名" callout
    s.addShape(pres.shapes.OVAL, {
      x: 1.5, y: 1.9, w: 1.5, h: 1.5, fill: { color: C.tealLight },
    });
    s.addText("1名", {
      x: 1.5, y: 1.9, w: 1.5, h: 1.1,
      fontFace: FONT_H, fontSize: 36, bold: true, color: C.tealDark, align: "center", valign: "middle", margin: 0,
    });
    s.addText("代表者", {
      x: 1.5, y: 2.85, w: 1.5, h: 0.35,
      fontFace: FONT_B, fontSize: 9, color: C.tealDark, align: "center", valign: "middle", margin: 0,
    });

    // Tasks
    const tasks = [
      { icon: I.usertie_teal, label: "受付対応" },
      { icon: I.car_teal, label: "車両管理" },
      { icon: I.broom_gray, label: "清掃" },
      { icon: I.toolbox_gray, label: "事務" },
    ];
    tasks.forEach((t, i) => {
      const tx = 3.5, ty = 1.9 + i * 0.7;
      s.addImage({ data: t.icon, x: tx, y: ty + 0.05, w: 0.25, h: 0.25 });
      s.addText(t.label, {
        x: tx + 0.35, y: ty, w: 1.8, h: 0.35,
        fontFace: FONT_B, fontSize: 10, color: C.dark, valign: "middle", margin: 0,
      });
    });

    // Note
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.8, y: 4.1, w: 5.1, h: 0.55, fill: { color: C.amberLight },
    });
    s.addImage({ data: I.light_amber, x: 0.95, y: 4.2, w: 0.25, h: 0.25 });
    s.addText("繁忙期は清掃・点検の外注を検討", {
      x: 1.3, y: 4.1, w: 4, h: 0.55,
      fontFace: FONT_B, fontSize: 9, color: C.amberDark, valign: "middle", margin: 0,
    });

    // Right: key principle
    s.addShape(pres.shapes.RECTANGLE, {
      x: 6.4, y: 1.2, w: 3.0, h: 3.7, fill: { color: C.dark },
    });
    s.addText("運営方針", {
      x: 6.6, y: 1.45, w: 2.5, h: 0.3,
      fontFace: FONT_H, fontSize: 14, bold: true, color: C.white, margin: 0,
    });
    s.addText("副業からスタートし固定費を最小限に抑え、段階的に事業を拡大", {
      x: 6.6, y: 1.9, w: 2.6, h: 0.8,
      fontFace: FONT_B, fontSize: 10, color: C.gray300, margin: 0,
    });
    s.addShape(pres.shapes.LINE, {
      x: 6.6, y: 2.85, w: 2.4, h: 0, line: { color: C.teal, width: 1 },
    });
    const principles = ["低リスク・段階拡大", "品質優先の管理体制", "柔軟なコスト構造"];
    principles.forEach((p, i) => {
      const py = 3.05 + i * 0.45;
      s.addImage({ data: I.check_w, x: 6.6, y: py + 0.02, w: 0.2, h: 0.2 });
      s.addText(p, {
        x: 6.9, y: py, w: 2.2, h: 0.28,
        fontFace: FONT_B, fontSize: 10, color: C.white, valign: "middle", margin: 0,
      });
    });

    footer(s, 8);
  }

  // ============================================
  // SLIDE 9: 初期投資
  // ============================================
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };
    header(s, "INITIAL INVESTMENT", "8. 初期投資", I.yen_teal);

    // Investment breakdown
    const items = [
      { label: "車両購入（中古1台）", amount: "600〜900", pct: 85, color: C.teal },
      { label: "登録・整備・保険", amount: "50", pct: 7, color: C.tealDark },
      { label: "設備・備品", amount: "20", pct: 3, color: C.amber },
      { label: "Web・広告", amount: "10", pct: 2, color: C.amberDark },
    ];
    items.forEach((item, i) => {
      const iy = 1.2 + i * 0.75;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: iy, w: 5.5, h: 0.6, fill: { color: C.white }, shadow: makeShadow(),
      });
      s.addText(item.label, {
        x: 0.8, y: iy, w: 2.2, h: 0.6,
        fontFace: FONT_B, fontSize: 10, bold: true, color: C.dark, valign: "middle", margin: 0,
      });
      // Bar
      const barMaxW = 1.5;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 3.1, y: iy + 0.2, w: barMaxW, h: 0.2, rectRadius: 0.04,
        fill: { color: C.gray100 },
      });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 3.1, y: iy + 0.2, w: barMaxW * item.pct / 100, h: 0.2, rectRadius: 0.04,
        fill: { color: item.color },
      });
      s.addText(item.amount + "万円", {
        x: 4.7, y: iy, w: 1.3, h: 0.6,
        fontFace: FONT_H, fontSize: 11, bold: true, color: C.dark, align: "right", valign: "middle", margin: 0,
      });
    });

    // Total callout
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 4.3, w: 5.5, h: 0.7, fill: { color: C.dark },
    });
    s.addText("合計", {
      x: 0.8, y: 4.3, w: 1.5, h: 0.7,
      fontFace: FONT_H, fontSize: 14, bold: true, color: C.white, valign: "middle", margin: 0,
    });
    s.addText("700〜1,000万円", {
      x: 3.0, y: 4.3, w: 3, h: 0.7,
      fontFace: FONT_H, fontSize: 22, bold: true, color: C.teal, align: "right", valign: "middle", margin: 0,
    });

    // Right: pie chart
    s.addChart(pres.charts.DOUGHNUT, [{
      name: "投資内訳",
      labels: ["車両購入", "登録等", "設備", "Web"],
      values: [86, 7, 4, 3],
    }], {
      x: 6.3, y: 1.2, w: 3.2, h: 3.2,
      showPercent: true,
      showTitle: false,
      showLegend: true,
      legendPos: "b",
      legendFontSize: 8,
      legendColor: C.gray600,
      chartColors: [C.teal, C.tealDark, C.amber, C.amberDark],
      dataLabelColor: C.white,
      dataLabelFontSize: 9,
    });

    footer(s, 9);
  }

  // ============================================
  // SLIDE 10: 収支モデル
  // ============================================
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };
    header(s, "REVENUE MODEL", "9. 収支モデル（月次・標準）", I.calc_teal);

    // KPI cards (top row)
    const kpis = [
      { label: "稼働日数", value: "12", unit: "日/月", color: C.teal },
      { label: "平均単価", value: "15,000", unit: "円/日", color: C.tealDark },
      { label: "月商", value: "約18", unit: "万円", color: C.amber },
    ];
    kpis.forEach((k, i) => {
      const kx = 0.6 + i * 3.1;
      s.addShape(pres.shapes.RECTANGLE, {
        x: kx, y: 1.2, w: 2.85, h: 1.3, fill: { color: C.white }, shadow: makeShadow(),
      });
      s.addShape(pres.shapes.RECTANGLE, {
        x: kx, y: 1.2, w: 2.85, h: 0.05, fill: { color: k.color },
      });
      s.addText(k.label, {
        x: kx + 0.15, y: 1.35, w: 2.5, h: 0.25,
        fontFace: FONT_B, fontSize: 8, color: C.gray500, margin: 0,
      });
      s.addText(k.value, {
        x: kx + 0.15, y: 1.6, w: 2, h: 0.5,
        fontFace: FONT_H, fontSize: 28, bold: true, color: C.dark, margin: 0,
      });
      s.addText(k.unit, {
        x: kx + 0.15, y: 2.12, w: 2, h: 0.2,
        fontFace: FONT_B, fontSize: 9, color: C.gray500, margin: 0,
      });
    });

    // PL breakdown
    const plY = 2.75;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: plY, w: 5.5, h: 2.15, fill: { color: C.white }, shadow: makeShadow(),
    });
    s.addText("月次損益", {
      x: 0.8, y: plY + 0.1, w: 2, h: 0.3,
      fontFace: FONT_H, fontSize: 13, bold: true, color: C.dark, margin: 0,
    });

    const plItems = [
      { label: "月商", val: "約18万円", bar: 100, color: C.teal },
      { label: "固定費", val: "10〜14万円", bar: 72, color: C.gray400 },
      { label: "変動費", val: "2万円", bar: 11, color: C.gray300 },
      { label: "月間利益", val: "2〜5万円", bar: 22, color: C.amber },
    ];
    plItems.forEach((pl, i) => {
      const py = plY + 0.5 + i * 0.4;
      s.addText(pl.label, {
        x: 0.8, y: py, w: 1.2, h: 0.3,
        fontFace: FONT_B, fontSize: 9, bold: true, color: C.dark, margin: 0,
      });
      const barMax = 2.2;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 2.1, y: py + 0.06, w: barMax, h: 0.18, rectRadius: 0.04, fill: { color: C.gray100 },
      });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 2.1, y: py + 0.06, w: barMax * pl.bar / 100, h: 0.18, rectRadius: 0.04, fill: { color: pl.color },
      });
      s.addText(pl.val, {
        x: 4.4, y: py, w: 1.5, h: 0.3,
        fontFace: FONT_B, fontSize: 9, bold: true, color: C.dark, align: "right", margin: 0,
      });
    });

    // Right: summary callout
    s.addShape(pres.shapes.RECTANGLE, {
      x: 6.4, y: plY, w: 3.0, h: 2.15, fill: { color: C.dark },
    });
    s.addText("PROFIT", {
      x: 6.6, y: plY + 0.2, w: 2, h: 0.2,
      fontFace: FONT_B, fontSize: 7, color: C.teal, charSpacing: 3, margin: 0,
    });
    s.addText("月間利益", {
      x: 6.6, y: plY + 0.45, w: 2, h: 0.3,
      fontFace: FONT_H, fontSize: 13, bold: true, color: C.white, margin: 0,
    });
    s.addText("2〜5", {
      x: 6.6, y: plY + 0.85, w: 2.5, h: 0.6,
      fontFace: FONT_H, fontSize: 40, bold: true, color: C.teal, margin: 0,
    });
    s.addText("万円 / 月", {
      x: 6.6, y: plY + 1.45, w: 2.5, h: 0.3,
      fontFace: FONT_B, fontSize: 12, color: C.gray400, margin: 0,
    });
    s.addText("副業として安定した副収入を確保", {
      x: 6.6, y: plY + 1.8, w: 2.5, h: 0.25,
      fontFace: FONT_B, fontSize: 8, color: C.gray500, margin: 0,
    });

    footer(s, 10);
  }

  // ============================================
  // SLIDE 11: 資金調達計画 + リスクと対策
  // ============================================
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };
    header(s, "FUNDING & RISK", "10-11. 資金調達計画 / リスクと対策", I.piggy_teal);

    // Left: Funding
    const fY = 1.2;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: fY, w: 4.3, h: 3.55, fill: { color: C.white }, shadow: makeShadow(),
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: fY, w: 4.3, h: 0.05, fill: { color: C.teal } });
    s.addText("資金調達計画", {
      x: 0.8, y: fY + 0.15, w: 3, h: 0.35,
      fontFace: FONT_H, fontSize: 14, bold: true, color: C.dark, margin: 0,
    });

    // Funding items
    const funds = [
      { label: "自己資金", value: "200万円", icon: I.piggy_teal, barPct: 20, color: C.teal },
      { label: "借入希望", value: "800〜1,000万円", icon: I.univ_teal, barPct: 83, color: C.tealDark },
    ];
    funds.forEach((f, i) => {
      const fy = fY + 0.65 + i * 1.05;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.8, y: fy, w: 3.9, h: 0.85, fill: { color: C.tealLight },
      });
      s.addImage({ data: f.icon, x: 1.0, y: fy + 0.15, w: 0.35, h: 0.35 });
      s.addText(f.label, {
        x: 1.5, y: fy + 0.08, w: 2, h: 0.25,
        fontFace: FONT_B, fontSize: 9, bold: true, color: C.tealDark, margin: 0,
      });
      s.addText(f.value, {
        x: 1.5, y: fy + 0.35, w: 2.5, h: 0.35,
        fontFace: FONT_H, fontSize: 16, bold: true, color: C.dark, margin: 0,
      });
    });

    s.addText("借入用途：車両購入・運転資金", {
      x: 0.8, y: fY + 2.85, w: 3.5, h: 0.3,
      fontFace: FONT_B, fontSize: 9, color: C.gray500, margin: 0,
    });

    // Stacked bar
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.8, y: fY + 3.2, w: 3.9 * 0.2, h: 0.2, fill: { color: C.teal },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.8 + 3.9 * 0.2, y: fY + 3.2, w: 3.9 * 0.8, h: 0.2, fill: { color: C.tealDark },
    });
    s.addText("20%", {
      x: 0.8, y: fY + 3.2, w: 3.9 * 0.2, h: 0.2,
      fontFace: FONT_B, fontSize: 7, bold: true, color: C.white, align: "center", valign: "middle", margin: 0,
    });
    s.addText("80%", {
      x: 0.8 + 3.9 * 0.2, y: fY + 3.2, w: 3.9 * 0.8, h: 0.2,
      fontFace: FONT_B, fontSize: 7, bold: true, color: C.white, align: "center", valign: "middle", margin: 0,
    });

    // Right: Risks
    s.addShape(pres.shapes.RECTANGLE, {
      x: 5.2, y: fY, w: 4.2, h: 3.55, fill: { color: C.white }, shadow: makeShadow(),
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: fY, w: 4.2, h: 0.05, fill: { color: C.amber } });
    s.addText("リスクと対策", {
      x: 5.4, y: fY + 0.15, w: 3, h: 0.35,
      fontFace: FONT_H, fontSize: 14, bold: true, color: C.dark, margin: 0,
    });

    const risks = [
      { risk: "稼働率低下", action: "副業スタートで固定費抑制" },
      { risk: "事故・故障", action: "任意保険加入・定期点検" },
      { risk: "価格競争", action: "レジャー特化・品質訴求" },
      { risk: "資金繰り", action: "慎重な台数拡大" },
    ];
    risks.forEach((r, i) => {
      const ry = fY + 0.65 + i * 0.7;
      s.addImage({ data: I.warn_amber, x: 5.4, y: ry + 0.02, w: 0.22, h: 0.22 });
      s.addText(r.risk, {
        x: 5.7, y: ry - 0.02, w: 1.5, h: 0.25,
        fontFace: FONT_B, fontSize: 9, bold: true, color: C.dark, margin: 0,
      });
      s.addText(r.action, {
        x: 5.7, y: ry + 0.25, w: 3.4, h: 0.25,
        fontFace: FONT_B, fontSize: 9, color: C.gray600, margin: 0,
      });
      if (i < 3) {
        s.addShape(pres.shapes.LINE, {
          x: 5.4, y: ry + 0.58, w: 3.7, h: 0, line: { color: C.gray200, width: 0.5 },
        });
      }
    });

    footer(s, 11);
  }

  // ============================================
  // SLIDE 12: 成長ロードマップ (Closing)
  // ============================================
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };

    // Decorative
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.04, fill: { color: C.teal } });
    s.addShape(pres.shapes.OVAL, { x: 8, y: 3.5, w: 3, h: 3, fill: { color: C.dark2 } });

    // Header
    s.addText("GROWTH ROADMAP", {
      x: 0.7, y: 0.35, w: 5, h: 0.2,
      fontFace: FONT_B, fontSize: 8, color: C.teal, charSpacing: 4, margin: 0,
    });
    s.addText("12. 成長ロードマップ", {
      x: 0.7, y: 0.6, w: 6, h: 0.45,
      fontFace: FONT_H, fontSize: 22, bold: true, color: C.white, margin: 0,
    });

    // Timeline
    // Horizontal line
    s.addShape(pres.shapes.LINE, {
      x: 0.7, y: 2.3, w: 8.6, h: 0, line: { color: C.teal, width: 2 },
    });

    const phases = [
      { year: "1年目", title: "安定運営", desc: "1台で安定運営\n稼働率向上に注力", icon: I.flag_amber, x: 1.2 },
      { year: "2〜3年目", title: "事業拡大", desc: "2台目導入\n売上基盤の拡大", icon: I.arrowup_teal, x: 4.2 },
      { year: "将来", title: "AI × キャンピングカー", desc: "キャンピングカー特化\nAI活用で高効率運営", icon: I.rocket_teal, x: 7.2 },
    ];
    phases.forEach((p) => {
      // Node circle
      s.addShape(pres.shapes.OVAL, {
        x: p.x, y: 2.05, w: 0.5, h: 0.5, fill: { color: C.teal },
      });
      s.addImage({ data: p.icon, x: p.x + 0.1, y: 2.15, w: 0.3, h: 0.3 });
      // Year
      s.addText(p.year, {
        x: p.x - 0.3, y: 1.5, w: 1.2, h: 0.35,
        fontFace: FONT_B, fontSize: 9, bold: true, color: C.teal, align: "center", margin: 0,
      });
      // Title
      s.addText(p.title, {
        x: p.x - 0.5, y: 2.7, w: 1.8, h: 0.35,
        fontFace: FONT_H, fontSize: 13, bold: true, color: C.white, align: "center", margin: 0,
      });
      // Desc
      s.addText(p.desc, {
        x: p.x - 0.6, y: 3.1, w: 2.0, h: 0.6,
        fontFace: FONT_B, fontSize: 9, color: C.gray400, align: "center", margin: 0,
      });
    });

    // Bottom CTA
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 4.4, w: 10, h: 1.225,
      fill: { color: "000000", transparency: 50 },
    });
    s.addImage({ data: I.car_w, x: 0.7, y: 4.65, w: 0.4, h: 0.4 });
    s.addText("AI レンタカー", {
      x: 1.2, y: 4.6, w: 3, h: 0.5,
      fontFace: FONT_H, fontSize: 20, bold: true, color: C.white, valign: "middle", margin: 0,
    });
    s.addText("個人の挑戦から、新しいモビリティ体験を。", {
      x: 1.2, y: 5.05, w: 5, h: 0.3,
      fontFace: FONT_B, fontSize: 10, color: C.gray400, margin: 0,
    });
    s.addText("Thank you", {
      x: 7, y: 4.6, w: 2.5, h: 0.5,
      fontFace: FONT_H, fontSize: 18, italic: true, color: C.teal, align: "right", valign: "middle", margin: 0,
    });
  }

  // Write
  const out = "/Users/nogataka/dev/slide-test/output/ai-rental-car/AIレンタカー_事業計画書.pptx";
  await pres.writeFile({ fileName: out });
  console.log("Created: " + out);
}

main().catch(console.error);
