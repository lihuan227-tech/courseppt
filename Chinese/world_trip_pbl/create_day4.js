/**
 * Day 4: 美洲 Americas (6/11) — Global Explorer Camp 环球探索沉浸式夏令营
 * ~45 slides — 3 countries (美国 USA, 墨西哥 Mexico, 巴西 Brazil)
 * Each country: expanded with dedicated topic slides + large images
 * Run: node create_day4.js
 */
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const pptx = new pptxgen();
pptx.defineLayout({ name: "LAYOUT_16x9", width: 10.0, height: 5.625 });
pptx.layout = "LAYOUT_16x9";
pptx.author = "谷雨中文 GR EDU";
pptx.title = "Global Explorer Camp · Day 4: 美洲 Americas";

// ── Colors (NO # prefix) ──
const C = {
  primary:    "2E7D32",
  secondary:  "E8F5E9",
  accent:     "FF6F00",
  dark:       "1B5E20",
  white:      "FFFFFF",
  black:      "212121",
  gray:       "616161",
  gold:       "FFD54F",
  darkGold:   "FFA000",
  lightGreen: "C8E6C9",
  lightAmber: "FFE0B2",
  bgBlue:     "E8EAF6",
  bgGreen:    "E8F5E9",
  bgOrange:   "FFF3E0",
  bgPink:     "FCE4EC",
  usa:        "1565C0",
  mexico:     "E65100",
  brazil:     "2E7D32",
  usaBg:      "E3F2FD",
  mexicoBg:   "FFF3E0",
  brazilBg:   "E8F5E9",
  contAsia:   "D32F2F",
  contAfrica: "FF8F00",
  contEurope: "1565C0",
  contNA:     "2E7D32",
  contSA:     "7CB342",
  contOceania:"00897B",
  contAntarc: "90A4AE",
  teal:       "00897B",
  purple:     "7B1FA2",
  green:      "388E3C",
  quizGreen:  "2E7D32",
  quizBg:     "E8F5E9",
};

// ── Local images ──
const IMG_DIR = path.join(__dirname, "images");
function imgPath(filename) {
  return path.join(IMG_DIR, filename);
}
function imgExists(filename) {
  return fs.existsSync(imgPath(filename));
}

// ── Helpers ──
function goldBars(slide) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: 9.8, h: 0.06, fill: { color: C.gold },
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 5.44, w: 9.8, h: 0.06, fill: { color: C.gold },
  });
}

function headerBar(slide, text, color) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: 9.8, h: 0.75, fill: { color: color || C.primary },
  });
  slide.addText(text, {
    x: 0.4, y: 0.05, w: 9.0, h: 0.65,
    fontSize: 28, fontFace: "Georgia", color: C.white, bold: true,
    align: "left", valign: "middle",
  });
}

function slideNum(slide) {
  slide.slideNumber = { x: "95%", y: "93%", fontSize: 8, color: C.gray };
}

function safeImage(slide, filename, label, x, y, w, h) {
  if (imgExists(filename)) {
    slide.addImage({ path: imgPath(filename), x: x, y: y, w: w, h: h, sizing: { type: "cover", w: w, h: h } });
  } else {
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: x, y: y, w: w, h: h, rectRadius: 0.12,
      fill: { color: C.bgBlue }, line: { color: C.gray, width: 1 },
    });
    slide.addText(label || "Photo", {
      x: x, y: y, w: w, h: h,
      fontSize: 13, fontFace: "Calibri", color: C.gray, align: "center", valign: "middle",
    });
  }
}

function card(slide, x, y, w, h, bgColor, borderColor) {
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: w, h: h, rectRadius: 0.15,
    fill: { color: bgColor },
    line: borderColor ? { color: borderColor, width: 1.2 } : undefined,
    shadow: { type: "outer", blur: 4, offset: 2, color: "BDBDBD", opacity: 0.25 },
  });
}

function footerText(slide, text, color) {
  slide.addText(text, {
    x: 0.5, y: 5.05, w: 9.0, h: 0.35,
    fontSize: 13, fontFace: "Calibri", color: color || C.primary, bold: true, align: "center",
  });
}

function quizSlides(title, allQs, color) {
  [0, 1].forEach((page) => {
    const s = pptx.addSlide();
    s.background = { fill: C.quizBg };
    slideNum(s);
    headerBar(s, title + " (" + (page + 1) + "/2)", C.quizGreen);

    var pageQs = allQs.slice(page * 5, page * 5 + 5);
    pageQs.forEach((item, i) => {
      var num = page * 5 + i + 1;
      var yy = 1.0 + i * 0.85;
      var rowBg = i % 2 === 0 ? C.white : C.quizBg;

      s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
        x: 0.3, y: yy, w: 9.2, h: 0.72, rectRadius: 0.1,
        fill: { color: rowBg },
        line: { color: C.quizGreen, width: 1 },
      });
      s.addShape(pptx.shapes.OVAL, {
        x: 0.45, y: yy + 0.13, w: 0.45, h: 0.45,
        fill: { color: C.quizGreen },
      });
      s.addText("" + num, {
        x: 0.45, y: yy + 0.13, w: 0.45, h: 0.45,
        fontSize: 14, fontFace: "Georgia", color: C.white,
        bold: true, align: "center", valign: "middle",
      });
      s.addText(item.q, {
        x: 1.05, y: yy, w: 4.45, h: 0.72,
        fontSize: 12, fontFace: "Georgia", color: C.dark,
        bold: true, align: "left", valign: "middle",
      });
      s.addText("\u2192 " + item.a, {
        x: 5.6, y: yy, w: 3.75, h: 0.72,
        fontSize: 11, fontFace: "Calibri", color: C.quizGreen,
        bold: true, align: "left", valign: "middle",
      });
    });
  });
}

let slideCount = 0;
function countSlide() { slideCount++; }


// ══════════════════════════════════════════════════════════════
// SLIDE 1 — Boarding Time
// ══════════════════════════════════════════════════════════════
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.dark };
  goldBars(s);
  slideNum(s);

  s.addShape(pptx.shapes.OVAL, {
    x: 0.3, y: 1.5, w: 0.9, h: 0.9, fill: { color: C.primary, transparency: 60 },
  });
  s.addShape(pptx.shapes.OVAL, {
    x: 8.5, y: 3.8, w: 0.7, h: 0.7, fill: { color: C.accent, transparency: 55 },
  });

  s.addText([
    { text: "Global Explorer Camp", options: { fontSize: 36, fontFace: "Georgia", color: C.gold, bold: true, breakLine: true } },
    { text: "环球探索沉浸式夏令营", options: { fontSize: 22, fontFace: "Georgia", color: C.white, breakLine: true } },
  ], { x: 0.5, y: 0.5, w: 9.0, h: 1.3, align: "center" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 2.5, y: 1.85, w: 5.0, h: 0.04, fill: { color: C.gold },
  });

  card(s, 2.0, 2.1, 6.0, 2.8, C.secondary);
  s.addText([
    { text: "BOARDING PASS  登机牌", options: { fontSize: 22, fontFace: "Georgia", color: C.primary, bold: true, breakLine: true } },
  ], { x: 2.3, y: 2.2, w: 5.4, h: 0.55, align: "center" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 2.3, y: 2.75, w: 5.4, h: 0.02, fill: { color: C.accent },
  });

  s.addText([
    { text: "Flight 航班: GR-004", options: { fontSize: 16, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "Destination 目的地:  美洲 AMERICAS", options: { fontSize: 16, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
    { text: "Date 日期: June 11, 2025  (6/11)", options: { fontSize: 14, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "Gate 登机口: Room 101", options: { fontSize: 14, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "Passenger 旅客: ________________", options: { fontSize: 14, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 2.5, y: 2.85, w: 5.0, h: 2.0, align: "left", valign: "top" });

  s.addText("Fasten seatbelts!  系好安全带！", {
    x: 0.5, y: 5.0, w: 9.0, h: 0.35,
    fontSize: 13, fontFace: "Calibri", color: C.gold, align: "center",
  });
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 2 — 护照进度 Passport Progress
// ══════════════════════════════════════════════════════════════
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🪪 护照进度  Passport Progress", C.dark);

  s.addText("集齐4个签证章就完成环球之旅！Collect all 4 stamps!", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.45,
    fontSize: 16, fontFace: "Georgia", color: C.dark, bold: true, align: "center",
  });

  const stamps = [
    { label: "亚洲\nAsia", color: C.contAsia, done: true, day: "Day 1" },
    { label: "非洲\nAfrica", color: C.contAfrica, done: true, day: "Day 2" },
    { label: "欧洲\nEurope", color: C.contEurope, done: true, day: "Day 3" },
    { label: "美洲\nAmericas", color: C.contNA, done: false, day: "Day 4" },
  ];

  stamps.forEach((st, i) => {
    const x = 0.5 + i * 2.35;
    s.addShape(pptx.shapes.OVAL, {
      x: x, y: 1.6, w: 2.1, h: 2.1,
      fill: { color: st.done ? st.color : C.white },
      line: { color: st.color, width: 3, dashType: st.done ? "solid" : "dash" },
    });
    s.addText(st.label, {
      x: x, y: 1.7, w: 2.1, h: 1.0,
      fontSize: 16, fontFace: "Georgia", color: st.done ? C.white : st.color, bold: true, align: "center", valign: "middle",
    });
    s.addText(st.done ? "\u2713" : "?", {
      x: x, y: 2.7, w: 2.1, h: 0.6,
      fontSize: 28, fontFace: "Georgia", color: st.done ? C.gold : C.gray, bold: true, align: "center",
    });
    s.addText(st.day, {
      x: x, y: 3.8, w: 2.1, h: 0.4,
      fontSize: 13, fontFace: "Calibri", color: st.color, bold: true, align: "center",
    });
  });

  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 1.5, y: 4.35, w: 7.0, h: 0.6, rectRadius: 0.15,
    fill: { color: C.dark },
  });
  s.addText("最后一个章！Last stamp today! Let's go!", {
    x: 1.5, y: 4.35, w: 7.0, h: 0.6,
    fontSize: 16, fontFace: "Georgia", color: C.gold, bold: true, align: "center", valign: "middle",
  });
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 3 — 今天目标 Today's Goals
// ══════════════════════════════════════════════════════════════
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🎯 今天目标  Today's Goals", C.primary);

  s.addText("深入了解美洲3个国家", {
    x: 0.5, y: 0.95, w: 9.0, h: 0.5,
    fontSize: 22, fontFace: "Georgia", color: C.dark, bold: true, align: "center",
  });

  const countries = [
    { flag: "🇺🇸", name: "美国 USA", color: C.usa, bg: C.usaBg },
    { flag: "🇲🇽", name: "墨西哥 Mexico", color: C.mexico, bg: C.mexicoBg },
    { flag: "🇧🇷", name: "巴西 Brazil", color: C.brazil, bg: C.brazilBg },
  ];

  countries.forEach((ct, i) => {
    const x = 0.5 + i * 3.1;
    card(s, x, 1.65, 2.8, 1.6, ct.bg, ct.color);
    s.addText(ct.flag, {
      x: x, y: 1.7, w: 2.8, h: 0.7,
      fontSize: 36, fontFace: "Calibri", align: "center", valign: "middle",
    });
    s.addText(ct.name, {
      x: x, y: 2.4, w: 2.8, h: 0.7,
      fontSize: 18, fontFace: "Georgia", color: ct.color, bold: true, align: "center",
    });
  });

  card(s, 0.5, 3.5, 9.0, 1.6, C.white, C.dark);
  s.addText([
    { text: "学习目标 Learning Targets:", options: { fontSize: 14, fontFace: "Georgia", color: C.dark, bold: true, breakLine: true } },
    { text: "1. 每个国家的基本信息（国旗、首都、人口、语言）", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "2. 文化特色与旅行礼节", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "3. 代表美食与饮食习惯", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "4. 比较三个国家的异同", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
  ], { x: 0.8, y: 3.55, w: 8.4, h: 1.5, valign: "top" });
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 4 — 认识美洲 About the Americas
// ══════════════════════════════════════════════════════════════
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🌎 认识美洲  About the Americas", C.contNA);

  const stats = [
    { icon: "🌎", label: "北美洲 + 南美洲", val: "Two Continents!" },
    { icon: "🌐", label: "35个国家", val: "35 Countries" },
    { icon: "🌿", label: "亚马逊雨林", val: "Amazon Rainforest" },
    { icon: "⛰️", label: "安第斯山脉", val: "Andes Mountains (7,000km)" },
    { icon: "📏", label: "从北极到南极", val: "Arctic to Antarctic" },
    { icon: "🗣️", label: "英/西/葡语为主", val: "English/Spanish/Portuguese" },
  ];

  stats.forEach((st, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.3 + col * 3.15;
    const y = 1.0 + row * 2.1;

    card(s, x, y, 2.9, 1.8, C.white, C.contNA);
    s.addText(st.icon, {
      x: x, y: y + 0.1, w: 2.9, h: 0.6,
      fontSize: 30, fontFace: "Calibri", align: "center",
    });
    s.addText(st.label, {
      x: x, y: y + 0.7, w: 2.9, h: 0.5,
      fontSize: 16, fontFace: "Georgia", color: C.contNA, bold: true, align: "center",
    });
    s.addText(st.val, {
      x: x, y: y + 1.2, w: 2.9, h: 0.45,
      fontSize: 12, fontFace: "Calibri", color: C.gray, align: "center",
    });
  });

  footerText(s, "美洲横跨南北半球，拥有丰富的自然和文化多样性！", C.contNA);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 5 — 美洲地图 Americas Map
// ══════════════════════════════════════════════════════════════
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: "E8F0FE" };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🗺️ 美洲地图  Americas Map", C.contNA);

  const regions = [
    { label: "🇺🇸 美国\nUSA", x: 0.8, y: 1.2, w: 3.0, h: 1.8, color: C.usa },
    { label: "🇲🇽 墨西哥\nMexico", x: 1.2, y: 3.1, w: 2.2, h: 1.2, color: C.mexico },
    { label: "🇧🇷 巴西\nBrazil", x: 4.5, y: 3.0, w: 2.8, h: 2.0, color: C.brazil },
    { label: "加拿大\nCanada", x: 0.8, y: 0.85, w: 2.5, h: 0.6, color: C.contAntarc },
    { label: "中美洲\nC. America", x: 2.0, y: 3.4, w: 1.5, h: 0.8, color: C.teal },
    { label: "安第斯\nAndes", x: 3.5, y: 3.5, w: 1.0, h: 1.8, color: C.contSA },
  ];

  regions.forEach((r) => {
    const isTarget = r.label.includes("美国") || r.label.includes("墨西哥") || r.label.includes("巴西");
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: r.x, y: r.y, w: r.w, h: r.h, rectRadius: 0.2,
      fill: { color: r.color, transparency: isTarget ? 15 : 60 },
      line: isTarget ? { color: r.color, width: 2.5 } : undefined,
    });
    s.addText(r.label, {
      x: r.x, y: r.y, w: r.w, h: r.h,
      fontSize: isTarget ? 14 : 11, fontFace: "Georgia",
      color: C.white, bold: true,
      align: "center", valign: "middle",
    });
  });

  // Legend on the right
  card(s, 7.5, 1.2, 2.2, 3.5, C.white, C.contNA);
  s.addText("今天去这3个国家!", {
    x: 7.5, y: 1.3, w: 2.2, h: 0.4,
    fontSize: 12, fontFace: "Georgia", color: C.dark, bold: true, align: "center",
  });

  const legend = [
    { flag: "🇺🇸", name: "美国", color: C.usa },
    { flag: "🇲🇽", name: "墨西哥", color: C.mexico },
    { flag: "🇧🇷", name: "巴西", color: C.brazil },
  ];
  legend.forEach((l, i) => {
    s.addText(l.flag + " " + l.name, {
      x: 7.6, y: 1.85 + i * 0.7, w: 2.0, h: 0.5,
      fontSize: 14, fontFace: "Georgia", color: l.color, bold: true, align: "center", valign: "middle",
    });
  });

  footerText(s, "今天我们去这三个国家！Today we visit these 3 countries!", C.contNA);
})();


// ══════════════════════════════════════════════════════════════
// ══════════════ USA SECTION (8 slides) ═══════════════════════
// ══════════════════════════════════════════════════════════════

// SLIDE 6 — 🇺🇸 美国概览 USA Overview
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🇺🇸 美国概览  USA Overview", C.usa);

  // Liberty photo LEFT
  safeImage(s, "americas_liberty.jpg", "Statue of Liberty 自由女神", 0.3, 0.95, 5.0, 3.5);

  // Info card RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.usa);

  const items = [
    { text: "🏴 国旗：星条旗 50星13条", y: 1.1 },
    { text: "👥 人口：约3.3亿", y: 1.6 },
    { text: "🗣️ 语言：无官方语言！", y: 2.1 },
    { text: "     (英语是事实上的通用语)", y: 2.45 },
    { text: "🏛️ 首都：华盛顿DC Washington DC", y: 2.85 },
    { text: "📏 面积：世界第三大国", y: 3.35 },
    { text: "🗽 象征：自由女神像", y: 3.85 },
  ];

  items.forEach((it) => {
    s.addText(it.text, {
      x: 5.85, y: it.y, w: 3.5, h: 0.42,
      fontSize: 13, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "美国是世界上最多元文化的国家之一！", C.usa);
})();


// SLIDE 7 — 🏛️ 首都：华盛顿DC
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🏛️ 首都：华盛顿DC  Capital: Washington DC", C.usa);

  // DC photo LEFT
  safeImage(s, "americas_dc.jpg", "Washington DC 华盛顿", 0.3, 0.95, 5.0, 3.5);

  // Info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.usa);
  s.addText([
    { text: "首都不是纽约！", options: { fontSize: 14, fontFace: "Georgia", color: C.usa, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "白宫 White House", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "  总统的家和办公室", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "国会大厦 Capitol Building", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "  美国国会开会的地方", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "林肯纪念堂 Lincoln Memorial", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "  纪念第16任总统", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "史密森尼博物馆 Smithsonian", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "  世界最大的博物馆群(免费!)", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 3.8, valign: "top" });

  footerText(s, "DC不属于任何一个州，是特别行政区！", C.usa);
})();


// SLIDE 8 — 🗽 自由女神像 Statue of Liberty
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🗽 自由女神像  Statue of Liberty", C.usa);

  // Liberty photo LEFT
  safeImage(s, "americas_liberty.jpg", "Statue of Liberty 自由女神", 0.3, 0.95, 5.0, 3.5);

  // Info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.usa);
  s.addText([
    { text: "法国1886年送的礼物！", options: { fontSize: 14, fontFace: "Georgia", color: C.usa, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "高93米（约30层楼高）", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "代表自由和民主", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "每年400万游客参观", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "位于纽约港口的小岛上", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "右手举火炬，左手拿书", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 2.8, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 4.0, w: 3.5, h: 0.02, fill: { color: C.usa },
  });

  s.addText("「给我你的疲惫和贫穷的人民」\n——女神底座上的诗句", {
    x: 5.85, y: 4.1, w: 3.5, h: 0.7,
    fontSize: 11, fontFace: "Georgia", color: C.usa, bold: true, align: "left",
  });
})();


// SLIDE 9 — 🌈 美国多元文化 American Diversity
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🌈 美国多元文化  American Diversity", C.usa);

  // NYC photo LEFT
  safeImage(s, "americas_nyc.jpg", "New York City 纽约", 0.3, 0.95, 5.0, 3.5);

  // Info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.usaBg, C.usa);
  s.addText([
    { text: "移民国家 Nation of Immigrants", options: { fontSize: 14, fontFace: "Georgia", color: C.usa, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "唐人街 Chinatown", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "  纽约、旧金山都有中国城", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "华人修铁路历史", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "  19世纪华工建造太平洋铁路", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "好莱坞 Hollywood", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "  全球电影中心", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "NASA 美国宇航局", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "  1969年人类首次登月！", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 5.85, y: 1.05, w: 3.5, h: 3.5, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 4.55, w: 3.5, h: 0.02, fill: { color: C.usa },
  });
  s.addText("「你知道美国有哪些华人名人吗？」", {
    x: 5.85, y: 4.6, w: 3.5, h: 0.35,
    fontSize: 11, fontFace: "Georgia", color: C.usa, bold: true, align: "left",
  });

  footerText(s, "美国是一个由来自世界各地的移民建立的国家！", C.usa);
})();


// SLIDE 10 — 🍔 美国美食 American Food
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🍔 美国美食  American Food", C.usa);

  // Hamburger photo LEFT
  safeImage(s, "americas_hamburger.jpg", "Hamburger 汉堡", 0.3, 0.95, 5.0, 3.5);

  // Food list RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.usa);
  s.addText([
    { text: "🍔 汉堡 Hamburger", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    最具代表性的美式食物", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🌭 热狗 Hot Dog", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    看棒球赛必吃！", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🥧 苹果派 Apple Pie", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    「像苹果派一样美国」", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🦃 感恩节火鸡 Thanksgiving Turkey", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    每年11月全家团聚", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🍖 BBQ 烤肉", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    德州BBQ最有名！", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 5.85, y: 1.05, w: 3.5, h: 3.1, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 4.2, w: 3.5, h: 0.02, fill: { color: C.usa },
  });
  s.addText("可口可乐 Coca-Cola 1886年诞生于亚特兰大！", {
    x: 5.85, y: 4.3, w: 3.5, h: 0.5,
    fontSize: 11, fontFace: "Calibri", color: C.usa, bold: true, align: "left",
  });

  footerText(s, "美国饮食深受世界各地移民影响！", C.usa);
})();


// SLIDE 11 — ⚠️ 美国礼节 USA Etiquette
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "⚠️ 美国礼节  USA Etiquette", C.usa);

  card(s, 0.3, 1.0, 9.3, 4.2, C.usaBg, C.usa);
  s.addText("🇺🇸 在美国旅行要注意：", {
    x: 0.55, y: 1.1, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.usa, bold: true,
  });

  const tips = [
    { icon: "🤝", text: "握手微笑——美国人见面最常见的方式" },
    { icon: "💰", text: "给小费15-20%！（餐厅、出租车都要给）" },
    { icon: "📏", text: "注意个人空间——保持一定距离" },
    { icon: "👋", text: "直呼名字OK——老师也可以叫名字" },
    { icon: "👟", text: "鞋子可以穿进屋（和亚洲不同！）" },
    { icon: "🗣️", text: "说话很直接——Yes就是Yes，No就是No" },
    { icon: "⏰", text: "准时很重要——迟到被认为不礼貌" },
  ];

  tips.forEach((t, i) => {
    const y = 1.65 + i * 0.48;
    s.addText(t.icon + "  " + t.text, {
      x: 0.8, y: y, w: 8.5, h: 0.42,
      fontSize: 14, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "小费文化是美国最独特的文化之一！", C.usa);
})();


// SLIDE 12-13 — ✅ 美国 Check Understanding (1/2, 2/2)
(() => {
  quizSlides("✅ 美国 Check Understanding", [
    { q: "美国的首都是哪里？", a: "华盛顿DC (不是纽约!)" },
    { q: "美国国旗有几颗星？", a: "50颗 (代表50个州)" },
    { q: "自由女神像是谁送的？", a: "法国 (1886年)" },
    { q: "在美国吃饭要给多少小费？", a: "15-20%" },
    { q: "美国有官方语言吗？", a: "没有！(英语是通用语)" },
    { q: "自由女神像有多高？", a: "93米 (约30层楼)" },
    { q: "华人在美国建造了什么？", a: "太平洋铁路" },
    { q: "感恩节吃什么？", a: "火鸡 Turkey" },
    { q: "在美国可以穿鞋进屋吗？", a: "可以！" },
    { q: "可口可乐诞生于哪一年？", a: "1886年" },
  ], C.usa);
  countSlide(); countSlide();
})();


// ══════════════════════════════════════════════════════════════
// ══════════════ MEXICO SECTION (7 slides) ════════════════════
// ══════════════════════════════════════════════════════════════

// SLIDE 14 — 🇲🇽 墨西哥概览 Mexico Overview
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🇲🇽 墨西哥概览  Mexico Overview", C.mexico);

  // Chichen Itza LEFT
  safeImage(s, "americas_chichenitza.jpg", "Chichen Itza 奇琴伊察", 0.3, 0.95, 5.0, 3.5);

  // Info card RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.mexico);

  const items = [
    { text: "🏴 国旗：绿白红 + 鹰蛇仙人掌", y: 1.1 },
    { text: "👥 人口：约1.3亿", y: 1.6 },
    { text: "🗣️ 语言：西班牙语 Spanish", y: 2.1 },
    { text: "🏛️ 首都：墨西哥城 Mexico City", y: 2.6 },
    { text: "🌵 仙人掌是国家的象征", y: 3.1 },
    { text: "🏗️ 古文明：玛雅 + 阿兹特克", y: 3.6 },
    { text: "🌶️ 世界上辣椒种类最多！", y: 4.1 },
  ];

  items.forEach((it) => {
    s.addText(it.text, {
      x: 5.85, y: it.y, w: 3.5, h: 0.42,
      fontSize: 13, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "墨西哥国旗上的鹰叼蛇站在仙人掌上——来自阿兹特克传说！", C.mexico);
})();


// SLIDE 15 — 🏛️ 玛雅与阿兹特克文明
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🏛️ 玛雅与阿兹特克文明  Maya & Aztec", C.mexico);

  // Chichen Itza photo LEFT
  safeImage(s, "americas_chichenitza.jpg", "Chichen Itza 奇琴伊察金字塔", 0.3, 0.95, 5.0, 3.5);

  // Info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.mexico);
  s.addText([
    { text: "🏛️ 奇琴伊察金字塔", options: { fontSize: 14, fontFace: "Georgia", color: C.mexico, bold: true, breakLine: true } },
    { text: "  世界新七大奇迹之一！", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "🔢 玛雅人发明了数字0", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "  和印度一样独立发明的！", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "☀️ 阿兹特克太阳石", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "  精确的天文历法", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "🏙️ 墨西哥城建在湖上！", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "  阿兹特克首都在湖中的岛上", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 3.8, valign: "top" });

  footerText(s, "玛雅和阿兹特克是美洲最伟大的古文明！", C.mexico);
})();


// SLIDE 16 — 💀 亡灵节 Day of the Dead
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "💀 亡灵节  Day of the Dead", C.mexico);

  // Day of Dead photo LEFT
  safeImage(s, "americas_dayofdead.jpg", "Day of the Dead 亡灵节", 0.3, 0.95, 5.0, 3.5);

  // Info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.mexico);
  s.addText([
    { text: "不是恐怖的！Not scary!", options: { fontSize: 14, fontFace: "Georgia", color: C.mexico, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "纪念去世的亲人", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "彩色骷髅 Colorful skulls", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "万寿菊花 Marigold flowers", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "搭建祭坛 Build altars", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "放上去世亲人喜欢的食物", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "11月1-2日庆祝", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 2.8, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 4.0, w: 3.5, h: 0.02, fill: { color: C.mexico },
  });

  s.addText("迪士尼电影「Coco 寻梦环游记」\n就是关于亡灵节的！", {
    x: 5.85, y: 4.1, w: 3.5, h: 0.7,
    fontSize: 12, fontFace: "Georgia", color: C.mexico, bold: true, align: "left",
  });
})();


// SLIDE 17 — 🌮 墨西哥美食 Mexican Food
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🌮 墨西哥美食  Mexican Food", C.mexico);

  // Taco photo LEFT
  safeImage(s, "americas_taco.jpg", "Tacos 墨西哥卷饼", 0.3, 0.95, 5.0, 3.5);

  // Food list RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.mexico);
  s.addText([
    { text: "🌮 Taco 墨西哥卷饼", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    最具代表性的墨西哥食物", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🥑 Guacamole 牛油果酱", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    美国超级碗必备!", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🌯 Burrito 大卷饼", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    包着米饭豆子肉和酱", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🌽 玉米有59种颜色！", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    包括蓝色和黑色的玉米", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 5.85, y: 1.05, w: 3.5, h: 2.8, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 3.9, w: 3.5, h: 0.02, fill: { color: C.mexico },
  });

  s.addText([
    { text: "🍫 巧克力是阿兹特克人发明的！", options: { fontSize: 12, fontFace: "Georgia", color: C.mexico, bold: true, breakLine: true } },
    { text: "  原来是苦的，加了糖才变甜", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 5.85, y: 4.0, w: 3.5, h: 0.8, valign: "top" });
})();


// SLIDE 18 — ⚠️ 墨西哥礼节 Mexico Etiquette
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "⚠️ 墨西哥礼节  Mexico Etiquette", C.mexico);

  card(s, 0.3, 1.0, 9.3, 4.2, C.mexicoBg, C.mexico);
  s.addText("🇲🇽 在墨西哥旅行要注意：", {
    x: 0.55, y: 1.1, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.mexico, bold: true,
  });

  const tips = [
    { icon: "🤗", text: "Hola! 拥抱 + 亲脸颊——热情的问候方式" },
    { icon: "⏰", text: "时间比较随意——迟到15-30分钟很正常" },
    { icon: "🍽️", text: "午餐是最重要的一餐（下午2-4点！）" },
    { icon: "😊", text: "非常热情友好——陌生人也会打招呼" },
    { icon: "🗣️", text: "说一点西班牙语会让墨西哥人很开心" },
    { icon: "🌶️", text: "不怕辣？墨西哥辣椒可能超出想象！" },
  ];

  tips.forEach((t, i) => {
    const y = 1.65 + i * 0.53;
    s.addText(t.icon + "  " + t.text, {
      x: 0.8, y: y, w: 8.5, h: 0.45,
      fontSize: 14, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "墨西哥人是世界上最热情好客的民族之一！", C.mexico);
})();


// SLIDE 19-20 — ✅ 墨西哥 Check Understanding (1/2, 2/2)
(() => {
  quizSlides("✅ 墨西哥 Check Understanding", [
    { q: "墨西哥国旗上有什么动物？", a: "鹰叼蛇站在仙人掌上" },
    { q: "墨西哥人说什么语言？", a: "西班牙语 Spanish" },
    { q: "奇琴伊察是什么？", a: "玛雅金字塔（世界七大奇迹之一）" },
    { q: "玛雅人发明了什么数字？", a: "数字 0 ！" },
    { q: "亡灵节是恐怖的节日吗？", a: "不是！是纪念去世亲人" },
    { q: "巧克力是谁发明的？", a: "阿兹特克人" },
    { q: "墨西哥人怎么打招呼？", a: "拥抱 + 亲脸颊" },
    { q: "玉米有多少种颜色？", a: "59种！" },
    { q: "「Coco寻梦环游记」讲的是什么节日？", a: "亡灵节 Day of the Dead" },
    { q: "墨西哥一天中最重要的一餐是？", a: "午餐（下午2-4点）" },
  ], C.mexico);
  countSlide(); countSlide();
})();


// ══════════════════════════════════════════════════════════════
// ══════════════ BRAZIL SECTION (7 slides) ════════════════════
// ══════════════════════════════════════════════════════════════

// SLIDE 21 — 🇧🇷 巴西概览 Brazil Overview
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🇧🇷 巴西概览  Brazil Overview", C.brazil);

  // Rio photo LEFT
  safeImage(s, "americas_rio.jpg", "Rio de Janeiro 里约热内卢", 0.3, 0.95, 5.0, 3.5);

  // Info card RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.brazil);

  const items = [
    { text: "🏴 国旗：绿黄蓝 + 星空图案", y: 1.1 },
    { text: "👥 人口：约2.1亿", y: 1.6 },
    { text: "🗣️ 语言：葡萄牙语", y: 2.1 },
    { text: "     (不是西班牙语！)", y: 2.45 },
    { text: "🏛️ 首都：巴西利亚 Brasilia", y: 2.85 },
    { text: "     (不是里约！)", y: 3.15 },
    { text: "⚽ 足球王国——5次世界杯冠军！", y: 3.55 },
    { text: "📏 南美洲面积最大的国家", y: 4.05 },
  ];

  items.forEach((it) => {
    s.addText(it.text, {
      x: 5.85, y: it.y, w: 3.5, h: 0.38,
      fontSize: 12, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "巴西是南美洲最大的国家，也是足球的王国！", C.brazil);
})();


// SLIDE 22 — 🌿 亚马逊雨林 Amazon Rainforest
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🌿 亚马逊雨林  Amazon Rainforest", C.brazil);

  // Amazon photo LEFT
  safeImage(s, "americas_amazon.jpg", "Amazon Rainforest 亚马逊雨林", 0.3, 0.95, 5.0, 3.5);

  // Info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.brazil);
  s.addText([
    { text: "地球之肺 Lungs of Earth", options: { fontSize: 14, fontFace: "Georgia", color: C.brazil, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "产生全球20%的氧气！", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "世界上最大的热带雨林", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "拥有世界10%的物种", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "亚马逊河是世界最大的河流", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "  (水量是第二名的10倍!)", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "面积等于25个日本！", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 2.8, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 4.0, w: 3.5, h: 0.02, fill: { color: C.brazil },
  });

  s.addText("如果亚马逊消失，地球气候将剧变！", {
    x: 5.85, y: 4.1, w: 3.5, h: 0.7,
    fontSize: 12, fontFace: "Calibri", color: C.brazil, bold: true, align: "left",
  });
})();


// SLIDE 23 — 🎭 嘉年华 Carnival
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🎭 嘉年华 Carnival  世界最大的派对", C.brazil);

  // Carnival photo LEFT
  safeImage(s, "americas_carnival.jpg", "Rio Carnival 里约嘉年华", 0.3, 0.95, 5.0, 3.5);

  // Info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.brazil);
  s.addText([
    { text: "世界最大的派对！", options: { fontSize: 14, fontFace: "Georgia", color: C.brazil, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "在里约热内卢举行", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "桑巴舞 Samba dancing", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "华丽的服装和花车", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "持续5天不停歇", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "200万人上街跳舞！", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "在天主教大斋节前举行", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 2.8, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 4.0, w: 3.5, h: 0.02, fill: { color: C.brazil },
  });

  s.addText("嘉年华是巴西人一年中最期待的活动！", {
    x: 5.85, y: 4.1, w: 3.5, h: 0.7,
    fontSize: 12, fontFace: "Calibri", color: C.brazil, bold: true, align: "left",
  });
})();


// SLIDE 24 — 🍖 巴西美食 Brazilian Food
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🍖 巴西美食  Brazilian Food", C.brazil);

  // Text-focused slide with card layout
  card(s, 0.3, 0.95, 5.8, 4.3, C.brazilBg, C.brazil);
  s.addText([
    { text: "🍖 Churrasco 巴西烤肉", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    服务员拿大串肉走来走去!", options: { fontSize: 12, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🫘 Feijoada 黑豆饭", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    巴西的国民菜", options: { fontSize: 12, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🧀 Pao de Queijo 奶酪面包", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    圆圆的Q弹奶酪面包球", options: { fontSize: 12, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🫐 Acai Bowl 巴西莓碗", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    超级食物，从亚马逊来的！", options: { fontSize: 12, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 0.55, y: 1.1, w: 5.3, h: 2.8, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.55, y: 4.0, w: 5.3, h: 0.02, fill: { color: C.brazil },
  });
  s.addText("☕ 巴西是世界第一咖啡产国！", {
    x: 0.55, y: 4.1, w: 5.3, h: 0.5,
    fontSize: 13, fontFace: "Georgia", color: C.brazil, bold: true,
  });

  // Side facts card RIGHT
  card(s, 6.35, 0.95, 3.3, 4.3, C.white, C.brazil);
  s.addText("🍖 巴西饮食特色", {
    x: 6.55, y: 1.05, w: 2.9, h: 0.4,
    fontSize: 14, fontFace: "Georgia", color: C.brazil, bold: true, align: "center",
  });

  const factItems = [
    "巴西烤肉餐厅\n吃到饱(rodizio)!\n绿牌=继续上肉\n红牌=我吃饱了",
    "巴西是世界上\n最大的咖啡出口国\n全球1/3咖啡豆\n来自巴西！",
    "巴西莓(Acai)\n在巴西冲浪者中\n最受欢迎的\n能量食物",
  ];

  factItems.forEach((txt, i) => {
    s.addText(txt, {
      x: 6.55, y: 1.55 + i * 0.95, w: 2.9, h: 0.85,
      fontSize: 10, fontFace: "Calibri", color: C.black, align: "center", valign: "middle",
    });
    if (i < factItems.length - 1) {
      s.addShape(pptx.shapes.RECTANGLE, {
        x: 6.85, y: 2.4 + i * 0.95, w: 2.3, h: 0.01, fill: { color: C.brazil },
      });
    }
  });
})();


// SLIDE 25 — ⚠️ 巴西礼节 Brazil Etiquette
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "⚠️ 巴西礼节  Brazil Etiquette", C.brazil);

  card(s, 0.3, 1.0, 9.3, 4.2, C.brazilBg, C.brazil);
  s.addText("🇧🇷 在巴西旅行要注意：", {
    x: 0.55, y: 1.1, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.brazil, bold: true,
  });

  const tips = [
    { icon: "😘", text: "亲两次脸颊——见面和告别都这样" },
    { icon: "🤗", text: "非常热情外向——巴西人很爱聊天" },
    { icon: "📏", text: "站得很近——这是友好的表示" },
    { icon: "👍", text: "竖大拇指 = OK/好的（常用手势）" },
    { icon: "⚽", text: "千万别说足球不好——足球是巴西人的生命！" },
    { icon: "⏰", text: "迟到很正常——巴西时间比较灵活" },
  ];

  tips.forEach((t, i) => {
    const y = 1.65 + i * 0.53;
    s.addText(t.icon + "  " + t.text, {
      x: 0.8, y: y, w: 8.5, h: 0.45,
      fontSize: 14, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "巴西人是世界上最热情开朗的民族之一！", C.brazil);
})();


// SLIDE 26-27 — ✅ 巴西 Check Understanding (1/2, 2/2)
(() => {
  quizSlides("✅ 巴西 Check Understanding", [
    { q: "巴西人说什么语言？", a: "葡萄牙语 (不是西班牙语!)" },
    { q: "巴西的首都是哪里？", a: "巴西利亚 (不是里约!)" },
    { q: "亚马逊雨林产生多少氧气？", a: "全球20%！" },
    { q: "嘉年华持续几天？", a: "5天" },
    { q: "巴西赢了几次世界杯？", a: "5次！最多的国家" },
    { q: "巴西烤肉餐厅怎么叫停？", a: "翻红牌(我吃饱了)" },
    { q: "在巴西见面怎么打招呼？", a: "亲两次脸颊" },
    { q: "巴西是世界第几大咖啡产国？", a: "第一！" },
    { q: "亚马逊河水量是第二名的几倍？", a: "10倍！" },
    { q: "在巴西千万不能说什么不好？", a: "足球！" },
  ], C.brazil);
  countSlide(); countSlide();
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 28 — 🌽 美洲改变世界的食物
// ══════════════════════════════════════════════════════════════
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🌽 美洲改变世界的食物  Foods that Changed the World", C.dark);

  s.addText("这些食物都来自美洲！Without the Americas, no...", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.45,
    fontSize: 14, fontFace: "Georgia", color: C.dark, bold: true, align: "center",
  });

  const foods = [
    { icon: "🍅", name: "番茄 Tomato", impact: "没有番茄就没有\n披萨和炒鸡蛋！", color: "E53935" },
    { icon: "🥔", name: "土豆 Potato", impact: "没有土豆就没有\n薯条和薯片！", color: C.accent },
    { icon: "🌽", name: "玉米 Corn", impact: "现在全球主食\n养活几十亿人", color: C.darkGold },
    { icon: "🍫", name: "巧克力 Chocolate", impact: "阿兹特克人的\n神圣饮品", color: "795548" },
    { icon: "🌶️", name: "辣椒 Chili Pepper", impact: "没有辣椒就没有\n四川火锅！", color: "D32F2F" },
  ];

  foods.forEach((f, i) => {
    const x = 0.2 + i * 1.9;
    card(s, x, 1.55, 1.75, 3.4, C.white, f.color);
    s.addText(f.icon, {
      x: x, y: 1.65, w: 1.75, h: 0.7,
      fontSize: 36, fontFace: "Calibri", align: "center",
    });
    s.addText(f.name, {
      x: x, y: 2.35, w: 1.75, h: 0.5,
      fontSize: 12, fontFace: "Georgia", color: f.color, bold: true, align: "center",
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: x + 0.2, y: 2.9, w: 1.35, h: 0.02, fill: { color: f.color },
    });
    s.addText(f.impact, {
      x: x + 0.05, y: 3.0, w: 1.65, h: 1.6,
      fontSize: 11, fontFace: "Calibri", color: C.black, align: "center", valign: "top",
    });
  });

  footerText(s, "美洲的食物改变了全世界的饮食习惯！", C.dark);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 29 — 🎭 Mini Role Play
// ══════════════════════════════════════════════════════════════
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🎭 Mini Role Play  打招呼练习 (3-5 min)", C.purple);

  s.addText("站起来和旁边的同学练习！Stand up and practice!", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.45,
    fontSize: 16, fontFace: "Georgia", color: C.purple, bold: true, align: "center",
  });

  const greetings = [
    { flag: "🇺🇸", country: "美国 USA", action: "握手 + 「Hi!」\n微笑", bg: C.usaBg, color: C.usa },
    { flag: "🇲🇽", country: "墨西哥 Mexico", action: "拥抱 + 「Hola!」\n亲脸颊", bg: C.mexicoBg, color: C.mexico },
    { flag: "🇧🇷", country: "巴西 Brazil", action: "亲两次脸颊 +\n「Oi!」", bg: C.brazilBg, color: C.brazil },
  ];

  greetings.forEach((g, i) => {
    const x = 0.3 + i * 3.15;
    card(s, x, 1.55, 2.9, 3.0, g.bg, g.color);
    s.addText(g.flag, {
      x: x, y: 1.65, w: 2.9, h: 0.7,
      fontSize: 40, fontFace: "Calibri", align: "center",
    });
    s.addText(g.country, {
      x: x, y: 2.35, w: 2.9, h: 0.5,
      fontSize: 16, fontFace: "Georgia", color: g.color, bold: true, align: "center",
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: x + 0.3, y: 2.9, w: 2.3, h: 0.02, fill: { color: g.color },
    });
    s.addText(g.action, {
      x: x, y: 3.05, w: 2.9, h: 1.2,
      fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, align: "center", valign: "middle",
    });
  });

  footerText(s, "尊重每个文化的问候方式！Respect every culture's greeting!", C.purple);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 30 — 🏆 上午竞赛 + Project 提醒
// ══════════════════════════════════════════════════════════════
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.dark };
  goldBars(s);
  slideNum(s);

  s.addText("🏆 上午竞赛  Morning Challenge", {
    x: 0.5, y: 0.4, w: 9.0, h: 0.6,
    fontSize: 28, fontFace: "Georgia", color: C.gold, bold: true, align: "center",
  });

  card(s, 1.0, 1.3, 7.8, 2.5, C.secondary);
  s.addText([
    { text: "分组比赛！Team Challenge!", options: { fontSize: 20, fontFace: "Georgia", color: C.primary, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Round 1: 说出3个国家的首都", options: { fontSize: 14, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "Round 2: 每个国家的打招呼方式", options: { fontSize: 14, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "Round 3: 说出每个国家的代表食物", options: { fontSize: 14, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "Round 4: 说出美洲改变世界的食物", options: { fontSize: 14, fontFace: "Calibri", color: C.black, breakLine: true } },
  ], { x: 1.3, y: 1.4, w: 7.2, h: 2.3, align: "center", valign: "middle" });

  card(s, 1.0, 4.0, 7.8, 1.2, C.accent);
  s.addText([
    { text: "🧩 Project 提醒：今天是最后一页！", options: { fontSize: 16, fontFace: "Georgia", color: C.white, bold: true, breakLine: true } },
    { text: "明天文化展览，准备好你的护照！", options: { fontSize: 14, fontFace: "Calibri", color: C.gold, breakLine: true } },
  ], { x: 1.3, y: 4.05, w: 7.2, h: 1.1, align: "center", valign: "middle" });
})();


// ══════════════════════════════════════════════════════════════
// ══════════════ AFTERNOON SESSION ════════════════════════════
// ══════════════════════════════════════════════════════════════

// SLIDE 31 — ☀️ 下午开始 Afternoon Session
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.accent };
  slideNum(s);

  s.addShape(pptx.shapes.OVAL, {
    x: 7.5, y: 0.5, w: 1.5, h: 1.5, fill: { color: C.gold, transparency: 40 },
  });

  s.addText([
    { text: "☀️ 下午开始", options: { fontSize: 34, fontFace: "Georgia", color: C.white, bold: true, breakLine: true } },
    { text: "Afternoon Session", options: { fontSize: 20, fontFace: "Georgia", color: C.secondary, breakLine: true } },
  ], { x: 0.5, y: 1.2, w: 9.0, h: 1.5, align: "center" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 3.0, y: 2.8, w: 4.0, h: 0.04, fill: { color: C.white },
  });

  s.addText([
    { text: "复习 → 对比总结 → 语言学习 → Project Time!", options: { fontSize: 16, fontFace: "Calibri", color: C.white, bold: true, breakLine: true } },
  ], { x: 0.5, y: 3.1, w: 9.0, h: 0.6, align: "center" });
})();


// SLIDE 32 — 📖 快速复习 Quick Review
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "📖 快速复习  Quick Review", C.dark);

  s.addText("说出3个国家 + 每个国家一个特别的礼节", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.5,
    fontSize: 16, fontFace: "Georgia", color: C.primary, bold: true, align: "center",
  });

  const reviews = [
    { flag: "🇺🇸", country: "美国", hint: "___?", color: C.usa, bg: C.usaBg },
    { flag: "🇲🇽", country: "墨西哥", hint: "___?", color: C.mexico, bg: C.mexicoBg },
    { flag: "🇧🇷", country: "巴西", hint: "___?", color: C.brazil, bg: C.brazilBg },
  ];

  reviews.forEach((r, i) => {
    const x = 0.3 + i * 3.15;
    card(s, x, 1.6, 2.9, 3.2, r.bg, r.color);
    s.addText(r.flag, {
      x: x, y: 1.7, w: 2.9, h: 0.8,
      fontSize: 42, fontFace: "Calibri", align: "center",
    });
    s.addText(r.country, {
      x: x, y: 2.5, w: 2.9, h: 0.5,
      fontSize: 20, fontFace: "Georgia", color: r.color, bold: true, align: "center",
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: x + 0.3, y: 3.1, w: 2.3, h: 0.02, fill: { color: r.color },
    });
    s.addText("特别的礼节：", {
      x: x, y: 3.25, w: 2.9, h: 0.4,
      fontSize: 13, fontFace: "Calibri", color: C.gray, align: "center",
    });
    s.addText(r.hint, {
      x: x, y: 3.7, w: 2.9, h: 0.8,
      fontSize: 22, fontFace: "Georgia", color: r.color, bold: true, align: "center",
    });
  });

  footerText(s, "「谁能第一个说出来？」Who can answer first?", C.dark);
})();


// SLIDE 33 — 🌎 美洲三国文化对比 Comparison Table
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🌎 美洲三国文化对比  Comparison", C.dark);

  const cols = [1.8, 2.4, 2.4, 2.4];
  const startX = 0.4;
  const startY = 0.95;
  const rowH = 0.63;

  const headers = ["", "🇺🇸 美国", "🇲🇽 墨西哥", "🇧🇷 巴西"];
  const headerColors = [C.dark, C.usa, C.mexico, C.brazil];

  let cx = startX;
  headers.forEach((h, i) => {
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cx, y: startY, w: cols[i], h: rowH,
      fill: { color: headerColors[i] },
      line: { color: C.white, width: 1 },
    });
    s.addText(h, {
      x: cx, y: startY, w: cols[i], h: rowH,
      fontSize: 13, fontFace: "Georgia", color: C.white, bold: true, align: "center", valign: "middle",
    });
    cx += cols[i];
  });

  const rows = [
    ["打招呼", "握手 Hi!", "拥抱 Hola!", "亲脸颊 Oi!"],
    ["语言", "英语(非官方)", "西班牙语", "葡萄牙语"],
    ["代表食物", "汉堡 BBQ", "Taco 玉米饼", "烤肉 Churrasco"],
    ["重要节日", "感恩节", "亡灵节", "嘉年华"],
    ["守时", "很准时!", "比较灵活", "也比较灵活"],
    ["特别的", "给小费!", "59种玉米", "足球=生命"],
  ];

  rows.forEach((row, ri) => {
    const ry = startY + rowH * (ri + 1);
    const bgColor = ri % 2 === 0 ? C.white : C.bgBlue;
    let cx2 = startX;
    row.forEach((cell, ci) => {
      s.addShape(pptx.shapes.RECTANGLE, {
        x: cx2, y: ry, w: cols[ci], h: rowH,
        fill: { color: ci === 0 ? C.dark : bgColor },
        line: { color: "E0E0E0", width: 0.5 },
      });
      s.addText(cell, {
        x: cx2, y: ry, w: cols[ci], h: rowH,
        fontSize: 12, fontFace: "Calibri", color: ci === 0 ? C.white : C.black,
        bold: ci === 0, align: "center", valign: "middle",
      });
      cx2 += cols[ci];
    });
  });

  footerText(s, "三个国家各有特色，你最想去哪个？", C.dark);
})();


// SLIDE 34 — 🤔 共同点与不同 Similarities & Differences
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🤔 共同点与不同  Similarities & Differences", C.dark);

  // Similarities LEFT
  card(s, 0.3, 1.0, 4.5, 4.0, C.bgGreen, C.green);
  s.addText("✅ 共同点 Similarities", {
    x: 0.5, y: 1.1, w: 4.1, h: 0.45,
    fontSize: 16, fontFace: "Georgia", color: C.green, bold: true, align: "center",
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.55, w: 3.5, h: 0.02, fill: { color: C.green },
  });
  s.addText([
    { text: "都在美洲大陆", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "都曾是欧洲殖民地", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "都是多元文化国家", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "都热爱体育运动", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "打招呼都很热情友好", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 0.6, y: 1.7, w: 3.9, h: 2.8, valign: "top" });

  // Differences RIGHT
  card(s, 5.1, 1.0, 4.5, 4.0, C.bgOrange, C.accent);
  s.addText("❌ 不同点 Differences", {
    x: 5.3, y: 1.1, w: 4.1, h: 0.45,
    fontSize: 16, fontFace: "Georgia", color: C.accent, bold: true, align: "center",
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.6, y: 1.55, w: 3.5, h: 0.02, fill: { color: C.accent },
  });
  s.addText([
    { text: "语言不同（英/西/葡）", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "打招呼亲密程度不同", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "对时间的态度不同", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "    美国准时/墨巴灵活", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "小费文化只有美国有", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.4, y: 1.7, w: 3.9, h: 2.8, valign: "top" });
})();


// SLIDE 35 — 🧳 旅行小贴士 Travel Tips Summary
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🧳 旅行小贴士  Travel Tips Summary", C.dark);

  s.addText("「如果你去美洲旅行，要记住：」", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.45,
    fontSize: 16, fontFace: "Georgia", color: C.dark, bold: true, align: "center",
  });

  const tips = [
    { flag: "🇺🇸", country: "在美国", tips: "记得给小费！\n握手微笑说Hi\n准时很重要", color: C.usa, bg: C.usaBg },
    { flag: "🇲🇽", country: "在墨西哥", tips: "拥抱说Hola！\n午餐最重要\n迟到没关系", color: C.mexico, bg: C.mexicoBg },
    { flag: "🇧🇷", country: "在巴西", tips: "亲脸颊说Oi！\n别说足球不好\n竖大拇指=OK", color: C.brazil, bg: C.brazilBg },
  ];

  tips.forEach((t, i) => {
    const x = 0.3 + i * 3.15;
    card(s, x, 1.55, 2.9, 2.8, t.bg, t.color);
    s.addText(t.flag + " " + t.country, {
      x: x, y: 1.65, w: 2.9, h: 0.55,
      fontSize: 16, fontFace: "Georgia", color: t.color, bold: true, align: "center",
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: x + 0.3, y: 2.2, w: 2.3, h: 0.02, fill: { color: t.color },
    });
    s.addText(t.tips, {
      x: x + 0.15, y: 2.35, w: 2.6, h: 1.8,
      fontSize: 13, fontFace: "Calibri", color: C.black, align: "center", valign: "top",
    });
  });

  card(s, 1.5, 4.5, 7.0, 0.7, C.dark);
  s.addText("尊重每个国家的文化是最重要的！Respect every culture!", {
    x: 1.5, y: 4.5, w: 7.0, h: 0.7,
    fontSize: 15, fontFace: "Georgia", color: C.gold, bold: true, align: "center", valign: "middle",
  });
})();


// SLIDE 36 — 📝 生词卡 Vocabulary Cards
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "📝 生词卡  Vocabulary Cards", C.dark);

  const words = [
    { zh: "汉堡", py: "han bao", en: "hamburger" },
    { zh: "火鸡", py: "huo ji", en: "turkey" },
    { zh: "金字塔", py: "jin zi ta", en: "pyramid" },
    { zh: "亡灵节", py: "wang ling jie", en: "Day of Dead" },
    { zh: "嘉年华", py: "jia nian hua", en: "carnival" },
    { zh: "烤肉", py: "kao rou", en: "BBQ" },
    { zh: "雨林", py: "yu lin", en: "rainforest" },
    { zh: "移民", py: "yi min", en: "immigrant" },
    { zh: "小费", py: "xiao fei", en: "tip" },
    { zh: "足球", py: "zu qiu", en: "football" },
  ];

  const colW = [1.4, 2.0, 1.4, 2.0, 1.4, 1.4];
  const tableX = 0.3;
  const tableY = 0.95;
  const rowH = 0.42;

  const thLabels = ["汉字", "拼音 Pinyin", "English", "汉字", "拼音 Pinyin", "English"];
  let hx = tableX;
  thLabels.forEach((lbl, ci) => {
    s.addShape(pptx.shapes.RECTANGLE, {
      x: hx, y: tableY, w: colW[ci], h: rowH,
      fill: { color: C.dark }, line: { color: C.white, width: 0.5 },
    });
    s.addText(lbl, {
      x: hx, y: tableY, w: colW[ci], h: rowH,
      fontSize: 11, fontFace: "Georgia", color: C.white, bold: true, align: "center", valign: "middle",
    });
    hx += colW[ci];
  });

  for (let r = 0; r < 5; r++) {
    const ry = tableY + rowH * (r + 1);
    const bg = r % 2 === 0 ? C.white : C.bgBlue;
    const left = words[r];
    const right = words[r + 5];

    const rowData = [left.zh, left.py, left.en, right.zh, right.py, right.en];
    let rx = tableX;
    rowData.forEach((cell, ci) => {
      const isChinese = ci === 0 || ci === 3;
      s.addShape(pptx.shapes.RECTANGLE, {
        x: rx, y: ry, w: colW[ci], h: rowH,
        fill: { color: bg }, line: { color: "E0E0E0", width: 0.5 },
      });
      s.addText(cell, {
        x: rx, y: ry, w: colW[ci], h: rowH,
        fontSize: isChinese ? 14 : 11, fontFace: isChinese ? "Georgia" : "Calibri",
        color: isChinese ? C.primary : C.black,
        bold: isChinese, align: "center", valign: "middle",
      });
      rx += colW[ci];
    });
  }

  footerText(s, "跟老师一起读！Read with the teacher!", C.dark);
})();


// SLIDE 37 — 💬 句型练习 Sentence Patterns
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "💬 句型练习  Sentence Patterns", C.dark);

  const patterns = [
    { pattern: "「我喜欢吃___」", eng: "I like to eat ___", example: "我喜欢吃汉堡/taco/巴西烤肉", color: C.usa },
    { pattern: "「在___，人们用___打招呼」", eng: "In ___, people greet by ___", example: "在巴西，人们亲两次脸颊说Oi", color: C.brazil },
    { pattern: "「去___旅行要注意___」", eng: "When traveling to ___, be careful about ___", example: "去美国旅行要注意给15-20%小费", color: C.mexico },
  ];

  patterns.forEach((p, i) => {
    const y = 1.0 + i * 1.45;
    card(s, 0.3, y, 9.3, 1.25, C.white, p.color);
    s.addText(p.pattern, {
      x: 0.5, y: y + 0.05, w: 5.5, h: 0.45,
      fontSize: 16, fontFace: "Georgia", color: p.color, bold: true, align: "left", valign: "middle",
    });
    s.addText(p.eng, {
      x: 6.0, y: y + 0.05, w: 3.4, h: 0.45,
      fontSize: 11, fontFace: "Calibri", color: C.gray, align: "right", valign: "middle",
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 0.5, y: y + 0.55, w: 8.9, h: 0.01, fill: { color: p.color },
    });
    s.addText("例: " + p.example, {
      x: 0.5, y: y + 0.6, w: 8.9, h: 0.5,
      fontSize: 13, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "和同伴一起练习！Practice with a partner!", C.dark);
})();


// SLIDE 38 — 🎭 Role Play 华人移民 (10-15min)
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🎭 Role Play  角色扮演：华人移民 (10-15 min)", C.purple);

  s.addText("「你是一个华人移民，到美洲三个国家生活！」", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.4,
    fontSize: 15, fontFace: "Georgia", color: C.purple, bold: true, align: "center",
  });

  const rounds = [
    { round: "Round 1", scene: "到美国", action: "握手说Hi\n去唐人街吃中餐\n记得给小费！", color: C.usa, bg: C.usaBg },
    { round: "Round 2", scene: "到墨西哥", action: "拥抱说Hola\n吃taco和guacamole\n下午2点吃午餐", color: C.mexico, bg: C.mexicoBg },
    { round: "Round 3", scene: "到巴西", action: "亲脸颊说Oi\n去吃巴西烤肉\n一起看足球赛！", color: C.brazil, bg: C.brazilBg },
  ];

  rounds.forEach((r, i) => {
    const x = 0.3 + i * 3.15;
    card(s, x, 1.5, 2.9, 1.8, r.bg, r.color);
    s.addText(r.round + ": " + r.scene, {
      x: x, y: 1.55, w: 2.9, h: 0.4,
      fontSize: 13, fontFace: "Georgia", color: r.color, bold: true, align: "center",
    });
    s.addText(r.action, {
      x: x + 0.1, y: 2.0, w: 2.7, h: 1.1,
      fontSize: 12, fontFace: "Calibri", color: C.black, align: "center", valign: "middle",
    });
  });

  card(s, 0.3, 3.55, 9.3, 1.6, C.white, C.purple);
  s.addText("分层 Differentiation", {
    x: 0.5, y: 3.6, w: 3.0, h: 0.35,
    fontSize: 13, fontFace: "Georgia", color: C.purple, bold: true,
  });

  const levels = [
    { level: "🟢 零基础", desc: "模仿动作 + 说国家名", x: 0.5, color: C.green },
    { level: "🔵 Level 2-3", desc: "「你好！我想吃___」+ 做动作", x: 3.5, color: "1565C0" },
    { level: "🟣 Level 4", desc: "完整对话 + 介绍文化礼节 + 华人移民经历", x: 6.5, color: C.purple },
  ];

  levels.forEach((l) => {
    s.addText(l.level, {
      x: l.x, y: 4.0, w: 2.8, h: 0.35,
      fontSize: 12, fontFace: "Georgia", color: l.color, bold: true,
    });
    s.addText(l.desc, {
      x: l.x, y: 4.35, w: 2.8, h: 0.6,
      fontSize: 11, fontFace: "Calibri", color: C.black,
    });
  });
})();


// SLIDE 39 — 🎨 Project Time! Passport Template
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🎨 Project Time!  护照美洲页（最后一页!）", C.teal);

  s.addText("Passport 美洲页 Template — 这是护照的最后一页！", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.teal, bold: true, align: "center",
  });

  const fields = [
    { icon: "🌎", label: "Where am I?", value: "美洲 (美国/墨西哥/巴西)" },
    { icon: "👀", label: "What did I see?", value: "自由女神像 / 奇琴伊察 / 亚马逊雨林" },
    { icon: "🍜", label: "What did I eat?", value: "汉堡 / taco / 巴西烤肉" },
    { icon: "🎭", label: "Cultural Discovery", value: "小费文化 / 亡灵节 / 嘉年华" },
    { icon: "💬", label: "My Sentence", value: "我喜欢吃___。在___，人们..." },
  ];

  fields.forEach((f, i) => {
    const y = 1.45 + i * 0.75;
    card(s, 0.5, y, 9.0, 0.62, C.white, C.teal);
    s.addText(f.icon + " " + f.label, {
      x: 0.7, y: y + 0.02, w: 3.5, h: 0.28,
      fontSize: 13, fontFace: "Georgia", color: C.teal, bold: true,
    });
    s.addText("\u2192 " + f.value, {
      x: 0.7, y: y + 0.3, w: 8.5, h: 0.28,
      fontSize: 12, fontFace: "Calibri", color: C.black,
    });
  });

  footerText(s, "用彩色笔画出你最喜欢的！Draw your favorite!", C.teal);
})();


// SLIDE 40 — 📊 Project 分层
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "📊 Project 分层  Differentiated Tasks", C.teal);

  card(s, 0.3, 1.0, 9.3, 1.2, C.bgGreen, C.green);
  s.addText("🟢 零基础 Beginners", {
    x: 0.5, y: 1.05, w: 3.5, h: 0.35,
    fontSize: 14, fontFace: "Georgia", color: C.green, bold: true,
  });
  s.addText("画图 + 写词  (汉堡 / taco / 烤肉)", {
    x: 0.5, y: 1.45, w: 8.8, h: 0.55,
    fontSize: 13, fontFace: "Calibri", color: C.black,
  });

  card(s, 0.3, 2.4, 9.3, 1.2, C.bgBlue, "1565C0");
  s.addText("🔵 Level 2-3 Intermediate", {
    x: 0.5, y: 2.45, w: 4.0, h: 0.35,
    fontSize: 14, fontFace: "Georgia", color: "1565C0", bold: true,
  });
  s.addText("「I see ___. I eat ___. In Brazil, people kiss cheeks.」", {
    x: 0.5, y: 2.85, w: 8.8, h: 0.55,
    fontSize: 13, fontFace: "Calibri", color: C.black,
  });

  card(s, 0.3, 3.8, 9.3, 1.4, "F3E5F5", C.purple);
  s.addText("🟣 Level 4 Advanced", {
    x: 0.5, y: 3.85, w: 3.5, h: 0.35,
    fontSize: 14, fontFace: "Georgia", color: C.purple, bold: true,
  });
  s.addText("「In the USA, I saw the Statue of Liberty. I learned about tipping 15-20%.\nIn Mexico, Day of the Dead is not scary — it celebrates family love.」", {
    x: 0.5, y: 4.25, w: 8.8, h: 0.75,
    fontSize: 12, fontFace: "Calibri", color: C.black,
  });
})();


// SLIDE 41 — 📄 准备展览 Prepare for Exhibition
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "📄 准备展览  Prepare for Exhibition", C.teal);

  s.addText("明天是文化展览日！Tomorrow is Exhibition Day!", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.45,
    fontSize: 18, fontFace: "Georgia", color: C.teal, bold: true, align: "center",
  });

  card(s, 0.5, 1.5, 9.0, 3.5, C.white, C.teal);
  s.addText([
    { text: "展览准备清单 Exhibition Checklist:", options: { fontSize: 16, fontFace: "Georgia", color: C.dark, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "1. 检查护照4页都完成了吗？", options: { fontSize: 14, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "   亚洲 + 非洲 + 欧洲 + 美洲", options: { fontSize: 12, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "2. 每一页都有图画和文字吗？", options: { fontSize: 14, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "3. 准备好向同学和家长介绍", options: { fontSize: 14, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "   用中文说：我去了___，看到了___，吃了___", options: { fontSize: 12, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "4. 想一想：四天旅行中你最喜欢哪个国家？为什么？", options: { fontSize: 14, fontFace: "Calibri", color: C.black, breakLine: true } },
  ], { x: 0.8, y: 1.6, w: 8.4, h: 3.2, valign: "top" });
})();


// SLIDE 42 — 🗣️ 分享时间 Sharing Time
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🗣️ 分享时间  Sharing Time", C.dark);

  s.addText("和同伴分享你的美洲页！\nShare your Americas page with a partner!", {
    x: 0.5, y: 1.0, w: 9.0, h: 0.8,
    fontSize: 18, fontFace: "Georgia", color: C.dark, bold: true, align: "center",
  });

  card(s, 0.5, 2.0, 9.0, 2.8, C.white, C.dark);
  s.addText("句型提示 Sentence Starters:", {
    x: 0.7, y: 2.1, w: 8.6, h: 0.4,
    fontSize: 14, fontFace: "Georgia", color: C.dark, bold: true,
  });

  const starters = [
    "「我去了___，看到了___。」",
    "「我最喜欢吃___，因为___。」",
    "「在___旅行，要注意___。」",
    "「我觉得最有趣的是___。」",
  ];

  starters.forEach((st, i) => {
    s.addText(st, {
      x: 0.7, y: 2.6 + i * 0.5, w: 8.6, h: 0.45,
      fontSize: 14, fontFace: "Calibri", color: C.primary, bold: true,
    });
  });
})();


// SLIDE 43 — 🪪 美洲签证章
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.dark };
  goldBars(s);
  slideNum(s);

  s.addText("🪪 美洲签证章  Americas Visa Stamp", {
    x: 0.5, y: 0.4, w: 9.0, h: 0.6,
    fontSize: 26, fontFace: "Georgia", color: C.gold, bold: true, align: "center",
  });

  s.addShape(pptx.shapes.OVAL, {
    x: 3.0, y: 1.2, w: 4.0, h: 3.5,
    fill: { color: C.dark },
    line: { color: C.contNA, width: 4, dashType: "dash" },
  });

  s.addText([
    { text: "AMERICAS", options: { fontSize: 30, fontFace: "Georgia", color: C.contNA, bold: true, breakLine: true } },
    { text: "美洲", options: { fontSize: 24, fontFace: "Georgia", color: C.gold, breakLine: true } },
    { text: "\u2713 VISITED", options: { fontSize: 20, fontFace: "Georgia", color: C.quizGreen, bold: true, breakLine: true } },
    { text: "6/11/2025", options: { fontSize: 14, fontFace: "Calibri", color: C.white, breakLine: true } },
    { text: "美国 \u00B7 墨西哥 \u00B7 巴西", options: { fontSize: 14, fontFace: "Calibri", color: C.lightAmber, breakLine: true } },
  ], { x: 3.0, y: 1.4, w: 4.0, h: 3.1, align: "center", valign: "middle" });

  // Show all 4 stamps collected
  s.addText("集齐4个签证章！All 4 stamps collected!", {
    x: 0.5, y: 4.75, w: 9.0, h: 0.3,
    fontSize: 13, fontFace: "Calibri", color: C.gold, bold: true, align: "center",
  });

  const allStamps = [
    { name: "亚洲", color: C.contAsia },
    { name: "非洲", color: C.contAfrica },
    { name: "欧洲", color: C.contEurope },
    { name: "美洲", color: C.contNA },
  ];
  allStamps.forEach((st, i) => {
    const x = 0.8 + i * 0.7;
    s.addShape(pptx.shapes.OVAL, {
      x: x, y: 1.8, w: 0.5, h: 0.5,
      fill: { color: st.color },
    });
    s.addText("\u2713", {
      x: x, y: 1.8, w: 0.5, h: 0.5,
      fontSize: 12, fontFace: "Georgia", color: C.white, bold: true, align: "center", valign: "middle",
    });
  });
})();


// SLIDE 44 — ✈️ 明天文化展
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.dark };
  goldBars(s);
  slideNum(s);

  s.addText([
    { text: "✈️ 明天：文化展览日！", options: { fontSize: 30, fontFace: "Georgia", color: C.gold, bold: true, breakLine: true } },
    { text: "Tomorrow: Culture Exhibition Day!", options: { fontSize: 18, fontFace: "Georgia", color: C.white, breakLine: true } },
  ], { x: 0.5, y: 0.5, w: 9.0, h: 1.2, align: "center" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 3.0, y: 1.7, w: 4.0, h: 0.04, fill: { color: C.gold },
  });

  card(s, 1.5, 2.0, 7.0, 2.8, C.secondary);
  s.addText([
    { text: "Flight 航班: GR-005 (Final!)", options: { fontSize: 18, fontFace: "Georgia", color: C.teal, bold: true, breakLine: true } },
    { text: "Destination: 文化展览 EXHIBITION", options: { fontSize: 18, fontFace: "Georgia", color: C.teal, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "带上你的护照，向大家展示你的环球之旅！", options: { fontSize: 16, fontFace: "Georgia", color: C.dark, bold: true, breakLine: true } },
    { text: "Show your passport to everyone!", options: { fontSize: 14, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "准备好用中文介绍你最喜欢的国家！", options: { fontSize: 14, fontFace: "Georgia", color: C.dark, breakLine: true } },
  ], { x: 1.8, y: 2.1, w: 6.4, h: 2.6, align: "center", valign: "middle" });

  s.addText("See you tomorrow, explorers!  明天见，小探险家们！", {
    x: 0.5, y: 4.95, w: 9.0, h: 0.35,
    fontSize: 14, fontFace: "Calibri", color: C.gold, bold: true, align: "center",
  });
})();


// ── Save ──
const outPath = path.join(__dirname, "day4_americas.pptx");
pptx.writeFile({ fileName: outPath }).then(() => {
  console.log("Created: " + outPath);
  console.log("Total slides: " + slideCount);
}).catch((err) => {
  console.error("Error:", err);
});
