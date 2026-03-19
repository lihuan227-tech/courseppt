/**
 * Day 1: 亚洲 Asia (6/8) — Global Explorer Camp 环球探索沉浸式夏令营
 * ~47 slides — 3 countries (中国 China, 日本 Japan, 印度 India)
 * Each country: expanded with dedicated topic slides + large images
 * Run: node create_day1.js
 */
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const pptx = new pptxgen();
pptx.defineLayout({ name: "LAYOUT_16x9", width: 10.0, height: 5.625 });
pptx.layout = "LAYOUT_16x9";
pptx.author = "谷雨中文 GR EDU";
pptx.title = "Global Explorer Camp · Day 1: 亚洲 Asia";

// ── Colors (NO # prefix) ──
const C = {
  primary:    "C62828",
  secondary:  "FFF8E1",
  accent:     "FF8F00",
  dark:       "1A237E",
  white:      "FFFFFF",
  black:      "212121",
  gray:       "616161",
  gold:       "FFD54F",
  darkGold:   "FFA000",
  lightRed:   "FFCDD2",
  lightAmber: "FFE0B2",
  bgBlue:     "E8EAF6",
  bgGreen:    "E8F5E9",
  bgOrange:   "FFF3E0",
  bgPink:     "FCE4EC",
  china:      "E53935",
  japan:      "D32F2F",
  india:      "FF9800",
  contAsia:   "D32F2F",
  contAfrica: "FF8F00",
  contEurope: "1565C0",
  contNA:     "2E7D32",
  contSA:     "7CB342",
  contOceania:"00897B",
  contAntarc: "90A4AE",
  mapLand:    "C8E6C9",
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
    { text: "Flight 航班: GR-001", options: { fontSize: 16, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "Destination 目的地:  亚洲 ASIA", options: { fontSize: 16, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
    { text: "Date 日期: June 8, 2025  (6/8)", options: { fontSize: 14, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "Gate 登机口: Room 101", options: { fontSize: 14, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "Passenger 旅客: ________________", options: { fontSize: 14, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 2.5, y: 2.85, w: 5.0, h: 2.0, align: "left", valign: "top" });

  s.addText("Fasten seatbelts!  系好安全带！", {
    x: 0.5, y: 5.0, w: 9.0, h: 0.35,
    fontSize: 13, fontFace: "Calibri", color: C.gold, align: "center",
  });
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 2 — 本周行程 Weekly Itinerary
// ══════════════════════════════════════════════════════════════
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "📅 本周行程  Weekly Itinerary", C.dark);

  const days = [
    { day: "Day 1", label: "亚洲\nAsia", color: C.contAsia, active: true },
    { day: "Day 2", label: "非洲\nAfrica", color: C.contAfrica, active: false },
    { day: "Day 3", label: "欧洲\nEurope", color: C.contEurope, active: false },
    { day: "Day 4", label: "美洲\nAmericas", color: C.contNA, active: false },
    { day: "Day 5", label: "展览\nExhibition", color: C.teal, active: false },
  ];

  days.forEach((d, i) => {
    const x = 0.4 + i * 1.85;
    const bw = 1.7;
    card(s, x, 1.2, bw, 3.5, d.active ? d.color : C.white, d.active ? undefined : d.color);
    s.addText(d.day, {
      x: x, y: 1.3, w: bw, h: 0.45,
      fontSize: 16, fontFace: "Georgia", color: d.active ? C.white : d.color, bold: true, align: "center",
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: x + 0.15, y: 1.78, w: bw - 0.3, h: 0.03, fill: { color: d.active ? C.gold : d.color },
    });
    s.addText(d.label, {
      x: x, y: 2.0, w: bw, h: 1.2,
      fontSize: 20, fontFace: "Georgia", color: d.active ? C.white : d.color, bold: true, align: "center", valign: "middle",
    });
    if (d.active) {
      s.addText("TODAY!", {
        x: x, y: 3.8, w: bw, h: 0.5,
        fontSize: 18, fontFace: "Georgia", color: C.gold, bold: true, align: "center",
      });
    }
  });

  footerText(s, "6/8 - 6/12  五天环游世界！", C.dark);
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

  s.addText("深入了解亚洲3个国家", {
    x: 0.5, y: 0.95, w: 9.0, h: 0.5,
    fontSize: 22, fontFace: "Georgia", color: C.dark, bold: true, align: "center",
  });

  const countries = [
    { flag: "🇨🇳", name: "中国 China", color: C.china, bg: C.lightRed },
    { flag: "🇯🇵", name: "日本 Japan", color: C.japan, bg: C.bgPink },
    { flag: "🇮🇳", name: "印度 India", color: C.india, bg: C.bgOrange },
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
// SLIDE 4 — 认识七大洲
// ══════════════════════════════════════════════════════════════
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🌍 认识七大洲  Seven Continents", C.dark);

  const conts = [
    { name: "亚洲\nAsia", size: "最大!", color: C.contAsia },
    { name: "非洲\nAfrica", size: "#2", color: C.contAfrica },
    { name: "欧洲\nEurope", size: "#6", color: C.contEurope },
    { name: "北美洲\nN. America", size: "#3", color: C.contNA },
    { name: "南美洲\nS. America", size: "#4", color: C.contSA },
    { name: "大洋洲\nOceania", size: "#7", color: C.contOceania },
    { name: "南极洲\nAntarctica", size: "#5", color: C.contAntarc },
  ];

  conts.forEach((ct, i) => {
    const col = i % 4;
    const row = i < 4 ? 0 : 1;
    const x = 0.3 + col * 2.4;
    const y = 1.0 + row * 2.1;
    const isAsia = i === 0;
    const bw = 2.1;
    const bh = 1.8;

    card(s, x, y, bw, bh, isAsia ? ct.color : C.white, ct.color);
    s.addText(ct.name, {
      x: x, y: y + 0.15, w: bw, h: 0.9,
      fontSize: 14, fontFace: "Georgia", color: isAsia ? C.white : ct.color, bold: true, align: "center", valign: "middle",
    });
    s.addText(ct.size, {
      x: x, y: y + 1.1, w: bw, h: 0.5,
      fontSize: isAsia ? 20 : 14, fontFace: "Georgia", color: isAsia ? C.gold : C.gray, bold: true, align: "center",
    });
  });
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 5 — 世界地图 World Map
// ══════════════════════════════════════════════════════════════
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: "E8F0FE" };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🗺️ 世界地图  World Map", C.dark);

  const blobs = [
    { label: "北美洲", x: 0.5, y: 1.3, w: 2.2, h: 1.8, color: C.contNA },
    { label: "南美洲", x: 1.3, y: 3.1, w: 1.4, h: 1.8, color: C.contSA },
    { label: "欧洲", x: 3.8, y: 1.1, w: 1.5, h: 1.2, color: C.contEurope },
    { label: "非洲", x: 3.8, y: 2.4, w: 1.8, h: 2.2, color: C.contAfrica },
    { label: "亚洲", x: 5.8, y: 1.0, w: 3.2, h: 2.6, color: C.contAsia },
    { label: "大洋洲", x: 7.2, y: 3.8, w: 1.8, h: 1.2, color: C.contOceania },
    { label: "南极洲", x: 3.0, y: 4.8, w: 4.0, h: 0.6, color: C.contAntarc },
  ];

  blobs.forEach((b) => {
    const isAsia = b.label === "亚洲";
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: b.x, y: b.y, w: b.w, h: b.h, rectRadius: 0.3,
      fill: { color: b.color, transparency: isAsia ? 10 : 55 },
      line: isAsia ? { color: C.primary, width: 2.5 } : undefined,
    });
    s.addText(b.label, {
      x: b.x, y: b.y, w: b.w, h: b.h,
      fontSize: isAsia ? 18 : 12, fontFace: "Georgia",
      color: C.white, bold: true,
      align: "center", valign: "middle",
    });
  });

  s.addText("We are here!  我们在这里！ →", {
    x: 5.0, y: 3.7, w: 3.5, h: 0.4,
    fontSize: 13, fontFace: "Calibri", color: C.primary, bold: true, align: "center",
  });
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 6 — 认识亚洲 About Asia
// ══════════════════════════════════════════════════════════════
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🌏 认识亚洲  About Asia", C.contAsia);

  const stats = [
    { icon: "🏔️", label: "世界最大的洲", val: "Largest Continent" },
    { icon: "🌐", label: "48个国家", val: "48 Countries" },
    { icon: "👥", label: "46亿人口", val: "4.6 Billion People" },
    { icon: "⛰️", label: "珠穆朗玛峰", val: "Mt. Everest (8,849m)" },
    { icon: "📏", label: "面积4,458万km\u00B2", val: "Covers 30% of Earth" },
    { icon: "🗣️", label: "2,300+种语言", val: "Most Diverse!" },
  ];

  stats.forEach((st, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.3 + col * 3.15;
    const y = 1.0 + row * 2.1;

    card(s, x, y, 2.9, 1.8, C.white, C.contAsia);
    s.addText(st.icon, {
      x: x, y: y + 0.1, w: 2.9, h: 0.6,
      fontSize: 30, fontFace: "Calibri", align: "center",
    });
    s.addText(st.label, {
      x: x, y: y + 0.7, w: 2.9, h: 0.5,
      fontSize: 16, fontFace: "Georgia", color: C.contAsia, bold: true, align: "center",
    });
    s.addText(st.val, {
      x: x, y: y + 1.2, w: 2.9, h: 0.45,
      fontSize: 12, fontFace: "Calibri", color: C.gray, align: "center",
    });
  });

  footerText(s, "亚洲是世界上人口最多、面积最大的洲！", C.contAsia);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 7 — 亚洲地图 Asia Map
// ══════════════════════════════════════════════════════════════
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: "E8F0FE" };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🗺️ 亚洲地图  Asia Map", C.contAsia);

  // Simplified Asia map with country blobs
  const countries = [
    { label: "🇨🇳 中国\nChina", x: 4.5, y: 1.5, w: 3.0, h: 2.0, color: C.china },
    { label: "🇯🇵 日本\nJapan", x: 7.8, y: 1.5, w: 1.5, h: 1.2, color: C.japan },
    { label: "🇮🇳 印度\nIndia", x: 3.5, y: 3.0, w: 2.0, h: 2.0, color: C.india },
    { label: "东南亚\nSE Asia", x: 5.8, y: 3.5, w: 1.8, h: 1.5, color: C.teal },
    { label: "中亚\nC. Asia", x: 2.0, y: 1.5, w: 2.0, h: 1.5, color: C.contAntarc },
    { label: "西亚\nW. Asia", x: 0.5, y: 2.2, w: 2.0, h: 1.8, color: C.accent },
  ];

  countries.forEach((c) => {
    const isTarget = ["🇨🇳 中国\nChina", "🇯🇵 日本\nJapan", "🇮🇳 印度\nIndia"].includes(c.label);
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: c.x, y: c.y, w: c.w, h: c.h, rectRadius: 0.2,
      fill: { color: c.color, transparency: isTarget ? 15 : 60 },
      line: isTarget ? { color: c.color, width: 2.5 } : undefined,
    });
    s.addText(c.label, {
      x: c.x, y: c.y, w: c.w, h: c.h,
      fontSize: isTarget ? 14 : 11, fontFace: "Georgia",
      color: C.white, bold: true,
      align: "center", valign: "middle",
    });
  });

  footerText(s, "今天我们去这三个国家！Today we visit these 3 countries!", C.contAsia);
})();


// ══════════════════════════════════════════════════════════════
// ══════════════ CHINA SECTION (8 slides) ═════════════════════
// ══════════════════════════════════════════════════════════════

// SLIDE 8 — 🇨🇳 中国概览 China Overview
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🇨🇳 中国概览  China Overview", C.china);

  // Great Wall photo LEFT
  safeImage(s, "china_great_wall.jpg", "Great Wall 长城", 0.3, 0.95, 5.0, 3.5);

  // Info card RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.china);

  const items = [
    { text: "🏴 国旗：红色，五颗黄星", y: 1.1 },
    { text: "👥 人口：约14亿（世界第一！）", y: 1.6 },
    { text: "🗣️ 语言：中文（普通话）", y: 2.1 },
    { text: "🏛️ 首都：北京 Beijing", y: 2.6 },
    { text: "📏 面积：世界第三大", y: 3.1 },
    { text: "🐼 国宝：大熊猫 Giant Panda", y: 3.6 },
    { text: "🏗️ 四大文明古国之一", y: 4.1 },
  ];

  items.forEach((it) => {
    s.addText(it.text, {
      x: 5.85, y: it.y, w: 3.5, h: 0.42,
      fontSize: 13, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "中国有五千年的历史！China has 5,000 years of history!", C.china);
})();


// SLIDE 9 — 🏛️ 首都：北京 Capital: Beijing
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🏛️ 首都：北京  Capital: Beijing", C.china);

  // Beijing photo LEFT
  safeImage(s, "china_beijing.jpg", "Beijing 北京", 0.3, 0.95, 5.0, 3.5);

  // Info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.china);
  s.addText([
    { text: "北京是中国的首都", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "有3000多年历史", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "2200万人口", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "故宫、天安门、长城都在北京附近", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "2008年和2022年奥运会城市", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 2.8, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 4.0, w: 3.5, h: 0.02, fill: { color: C.china },
  });

  s.addText("北京在中国的北方，这就是为什么叫「北」京！", {
    x: 5.85, y: 4.1, w: 3.5, h: 0.7,
    fontSize: 12, fontFace: "Calibri", color: C.china, bold: true, align: "left",
  });

  // Forbidden city small image bottom-left
  safeImage(s, "china_forbidden_city.jpg", "Forbidden City 故宫", 0.3, 4.6, 2.0, 0.8);
})();


// SLIDE 10 — 🏯 长城 The Great Wall
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🏯 长城  The Great Wall", C.china);

  // Great Wall photo LEFT (large)
  safeImage(s, "china_great_wall.jpg", "Great Wall 长城", 0.3, 0.95, 5.0, 4.0);

  // Info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.china);
  s.addText([
    { text: "长城有21,196公里长！", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "有2000多年历史", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "从太空看不到长城", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "  （这是一个误解！）", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "是世界七大奇迹之一", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "每年有1000万人参观", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "用来防御北方游牧民族", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 3.2, valign: "top" });

  footerText(s, "万里长城万里长！The Great Wall stretches 21,196 km!", C.china);
})();


// SLIDE 11 — 💡 四大发明 Four Great Inventions
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "💡 四大发明  Four Great Inventions", C.china);

  // Compass image LEFT
  safeImage(s, "china_inventions.jpg", "Compass 指南针", 0.3, 0.95, 5.0, 3.5);

  // Four inventions RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.china);

  const inventions = [
    { name: "📜 造纸术 Paper", detail: "105 AD — 蔡伦发明", y: 1.1 },
    { name: "📖 印刷术 Printing", detail: "868 AD — 最早的印刷书", y: 1.85 },
    { name: "💥 火药 Gunpowder", detail: "9th century — 原来是炼丹用的！", y: 2.6 },
    { name: "🧭 指南针 Compass", detail: "11th century — 帮助航海", y: 3.35 },
  ];

  inventions.forEach((inv) => {
    s.addText(inv.name, {
      x: 5.85, y: inv.y, w: 3.5, h: 0.35,
      fontSize: 13, fontFace: "Georgia", color: C.china, bold: true, align: "left",
    });
    s.addText(inv.detail, {
      x: 5.85, y: inv.y + 0.35, w: 3.5, h: 0.3,
      fontSize: 11, fontFace: "Calibri", color: C.gray, align: "left",
    });
  });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 4.2, w: 3.5, h: 0.02, fill: { color: C.china },
  });
  s.addText("这四大发明改变了整个世界！", {
    x: 5.85, y: 4.3, w: 3.5, h: 0.5,
    fontSize: 12, fontFace: "Calibri", color: C.china, bold: true, align: "left",
  });
})();


// SLIDE 12 — 🎆 春节 Chinese New Year
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🎆 春节  Chinese New Year", C.china);

  // Spring Festival photo LEFT
  safeImage(s, "china_spring_festival.jpg", "Chinese New Year 春节", 0.3, 0.95, 5.0, 3.5);

  // Info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.china);
  s.addText([
    { text: "春节是中国最重要的节日", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "全家人一起吃年夜饭", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "吃饺子、放鞭炮、贴春联", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "给孩子红包（压岁钱）", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "舞龙舞狮", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "持续15天！", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 2.8, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 4.0, w: 3.5, h: 0.02, fill: { color: C.china },
  });

  s.addText("「你们过春节吗？你收到过红包吗？」", {
    x: 5.85, y: 4.1, w: 3.5, h: 0.7,
    fontSize: 12, fontFace: "Georgia", color: C.china, bold: true, align: "left",
  });
})();


// SLIDE 13 — 🥢 中国饮食文化 Chinese Food Culture
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🥢 中国饮食文化  Chinese Food Culture", C.china);

  // Dumplings photo LEFT
  safeImage(s, "china_dumplings.jpg", "Dumplings 饺子", 0.3, 0.95, 5.0, 3.5);

  // Food info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.china);
  s.addText([
    { text: "🥟 八大菜系 Eight Cuisines", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "🥟 饺子 面条 炒饭 烤鸭", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "🌶️ 味道各不同:", options: { fontSize: 13, fontFace: "Calibri", color: C.china, bold: true, breakLine: true } },
    { text: "南甜 北咸 东辣 西酸", options: { fontSize: 14, fontFace: "Georgia", color: C.black, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "🥢 筷子有3000年历史！", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "🍵 茶文化——中国是茶的故乡", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
  ], { x: 5.85, y: 1.05, w: 3.5, h: 3.0, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 4.15, w: 3.5, h: 0.02, fill: { color: C.china },
  });
  s.addText("「你家过年吃什么？」", {
    x: 5.85, y: 4.25, w: 3.5, h: 0.5,
    fontSize: 13, fontFace: "Georgia", color: C.china, bold: true, align: "left",
  });

  footerText(s, "中国菜是世界上种类最多的菜系之一！", C.china);
})();


// SLIDE 14 — ⚠️ 中国旅行礼节 China Travel Etiquette
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "⚠️ 中国旅行礼节  China Travel Etiquette", C.china);

  // Text-focused slide
  card(s, 0.3, 1.0, 9.3, 4.2, C.lightRed, C.china);
  s.addText("🇨🇳 在中国旅行要注意：", {
    x: 0.55, y: 1.1, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.china, bold: true,
  });

  const tips = [
    { icon: "🤝", text: "用双手递东西表示尊重" },
    { icon: "🥢", text: "吃饭时不要把筷子插在饭里（不吉利）" },
    { icon: "🎁", text: "收到礼物不当面打开" },
    { icon: "👋", text: "见面握手，不拥抱" },
    { icon: "🔢", text: "数字4不吉利（发音像「死」）" },
    { icon: "🎨", text: "送礼不要用白色包装（白色=丧事）" },
    { icon: "🍵", text: "别人给你倒茶，轻敲桌面表示感谢" },
  ];

  tips.forEach((t, i) => {
    const y = 1.65 + i * 0.48;
    s.addText(t.icon + "  " + t.text, {
      x: 0.8, y: y, w: 8.5, h: 0.42,
      fontSize: 14, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "尊重当地文化是旅行者最重要的礼节！", C.china);
})();


// SLIDE 15-16 — ✅ 中国 Check Understanding (1/2, 2/2)
(() => {
  quizSlides("✅ 中国 Check Understanding", [
    { q: "中国的首都是哪里？", a: "北京 Beijing" },
    { q: "中国人用什么吃饭？", a: "筷子 Chopsticks" },
    { q: "中国有几大发明？", a: "四大发明" },
    { q: "吃饭时筷子不能怎么放？", a: "不能插在饭里" },
    { q: "中国有多少人口？", a: "约14亿" },
    { q: "中国的四大发明是什么？", a: "造纸、印刷、火药、指南针" },
    { q: "中国人过年吃什么？", a: "饺子" },
    { q: "长城有多长？", a: "21,196公里" },
    { q: "北京举办过几次奥运会？", a: "两次（2008和2022）" },
    { q: "收到礼物时，中国人一般怎么做？", a: "不当面打开" },
  ], C.china);
  countSlide(); countSlide();
})();


// ══════════════════════════════════════════════════════════════
// ══════════════ JAPAN SECTION (8 slides) ═════════════════════
// ══════════════════════════════════════════════════════════════

// SLIDE 17 — 🇯🇵 日本概览 Japan Overview
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🇯🇵 日本概览  Japan Overview", C.japan);

  // Mount Fuji LEFT
  safeImage(s, "japan_fuji.jpg", "Mt. Fuji 富士山", 0.3, 0.95, 5.0, 3.5);

  // Info card RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.japan);

  const items = [
    { text: "🏴 国旗：白底红圆（太阳）", y: 1.1 },
    { text: "👥 人口：约1.25亿", y: 1.6 },
    { text: "🗣️ 语言：日语", y: 2.1 },
    { text: "🏛️ 首都：东京 Tokyo", y: 2.6 },
    { text: "🏝️ 由6,852个岛屿组成！", y: 3.1 },
    { text: "🗻 富士山：日本最高峰", y: 3.6 },
    { text: "🚄 新干线：时速320km!", y: 4.1 },
  ];

  items.forEach((it) => {
    s.addText(it.text, {
      x: 5.85, y: it.y, w: 3.5, h: 0.42,
      fontSize: 13, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "日本是世界上最安全、最干净的国家之一！", C.japan);
})();


// SLIDE 18 — 🏙️ 首都：东京 Capital: Tokyo
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🏙️ 首都：东京  Capital: Tokyo", C.japan);

  // Tokyo skyline LEFT
  safeImage(s, "japan_tokyo.jpg", "Tokyo 东京", 0.3, 0.95, 5.0, 3.5);

  // Info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.japan);
  s.addText([
    { text: "东京是日本的首都", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "世界上最大的城市之一", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "3700万人口（大都市圈）", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "2021年举办了奥运会", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "新宿站每天350万人通过", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "有世界上最复杂的地铁系统", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 2.8, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 4.0, w: 3.5, h: 0.02, fill: { color: C.japan },
  });

  s.addText("「东京」的意思是「东方的首都」！", {
    x: 5.85, y: 4.1, w: 3.5, h: 0.7,
    fontSize: 12, fontFace: "Calibri", color: C.japan, bold: true, align: "left",
  });
})();


// SLIDE 19 — 🌸 樱花与富士山 Cherry Blossoms & Mt. Fuji
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🌸 樱花与富士山  Cherry Blossoms & Mt. Fuji", C.japan);

  // Cherry blossom photo LEFT
  safeImage(s, "japan_cherry_blossom.jpg", "Cherry Blossoms 樱花", 0.3, 0.95, 5.0, 3.5);

  // Info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.japan);
  s.addText([
    { text: "🌸 樱花 Cherry Blossoms", options: { fontSize: 14, fontFace: "Georgia", color: C.japan, bold: true, breakLine: true } },
    { text: "樱花(sakura)是日本的象征", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "每年3-4月盛开", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "「花见」——赏花野餐", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "🗻 富士山 Mt. Fuji", options: { fontSize: 14, fontFace: "Georgia", color: C.japan, bold: true, breakLine: true } },
    { text: "高3,776米，日本最高峰", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "活火山！最后喷发1707年", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "联合国世界文化遗产", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 3.8, valign: "top" });

  footerText(s, "樱花只开一周，象征生命的短暂与美丽！", C.japan);
})();


// SLIDE 20 — 🎮 动漫与流行文化 Anime & Pop Culture
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🎮 动漫与流行文化  Anime & Pop Culture", C.japan);

  // Pokemon center photo LEFT
  safeImage(s, "japan_pokemon.jpg", "Pokemon Center 宝可梦中心", 0.3, 0.95, 5.0, 3.5);

  // Pop culture info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.bgPink, C.japan);
  s.addText([
    { text: "⚡ Pokemon宝可梦", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "  全世界孩子都爱！", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🐉 Dragon Ball龙珠", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "  Naruto火影忍者", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🎵 Nintendo任天堂", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "  Mario马里奥的家！", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🎬 宫崎骏 Studio Ghibli", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "  龙猫、千与千寻", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🚄 新干线时速320km！", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "  电车准时到秒！", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 5.85, y: 1.05, w: 3.5, h: 3.5, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 4.5, w: 3.5, h: 0.02, fill: { color: C.japan },
  });
  s.addText("「你最喜欢哪个日本动漫？」", {
    x: 5.85, y: 4.55, w: 3.5, h: 0.4,
    fontSize: 12, fontFace: "Georgia", color: C.japan, bold: true, align: "left",
  });
})();


// SLIDE 21 — 🍣 日本美食 Japanese Food
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🍣 日本美食  Japanese Food", C.japan);

  // Sushi photo LEFT
  safeImage(s, "japan_sushi.jpg", "Sushi 寿司", 0.3, 0.95, 5.0, 3.5);

  // Food list RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.japan);
  s.addText([
    { text: "🍣 寿司 sushi", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    全球最受欢迎的日本料理", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🍜 拉面 ramen", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    日本有5万多家拉面店！", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🍤 天妇罗 tempura", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    炸虾炸蔬菜，酥脆美味", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🍵 抹茶 matcha", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    日本茶道文化", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🍱 便当 bento", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    午餐像艺术品！", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 5.85, y: 1.05, w: 3.5, h: 3.1, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 4.2, w: 3.5, h: 0.02, fill: { color: C.japan },
  });

  s.addText("特别：吃面条时发出声音是礼貌的！", {
    x: 5.85, y: 4.3, w: 3.5, h: 0.5,
    fontSize: 12, fontFace: "Calibri", color: C.japan, bold: true, align: "left",
  });

  footerText(s, "日本料理2013年入选联合国非物质文化遗产！", C.japan);
})();


// SLIDE 22 — ⚠️ 日本旅行礼节 Japan Travel Etiquette
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "⚠️ 日本旅行礼节  Japan Travel Etiquette", C.japan);

  card(s, 0.3, 1.0, 9.3, 4.2, C.bgPink, C.japan);
  s.addText("🇯🇵 在日本旅行要注意：", {
    x: 0.55, y: 1.1, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.japan, bold: true,
  });

  const tips = [
    { icon: "🙇", text: "见面鞠躬（不握手！）弯腰越深越尊重" },
    { icon: "👟", text: "进屋必须脱鞋" },
    { icon: "🙏", text: "吃饭前说「いただきます」(我开动了)" },
    { icon: "🤫", text: "不要在公共交通上大声说话" },
    { icon: "💰", text: "不要给小费（被认为不礼貌）" },
    { icon: "🚶", text: "不要边走边吃" },
    { icon: "📱", text: "电车上手机要调静音" },
  ];

  tips.forEach((t, i) => {
    const y = 1.65 + i * 0.48;
    s.addText(t.icon + "  " + t.text, {
      x: 0.8, y: y, w: 8.5, h: 0.42,
      fontSize: 14, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "日本人非常注重礼节和规矩！", C.japan);
})();


// SLIDE 23-24 — ✅ 日本 Check Understanding (1/2, 2/2)
(() => {
  quizSlides("✅ 日本 Check Understanding", [
    { q: "日本由多少个岛屿组成？", a: "6,852个！" },
    { q: "在日本见面应该怎么做？", a: "鞠躬 Bow" },
    { q: "进别人家要做什么？", a: "脱鞋 Take off shoes" },
    { q: "在日本吃拉面可以发出声音吗？", a: "可以！这是礼貌的" },
    { q: "富士山有多高？", a: "3,776米" },
    { q: "Pokemon(宝可梦)来自哪个国家？", a: "日本" },
    { q: "东京有多少人口（大都市圈）？", a: "3700万" },
    { q: "在日本可以给服务员小费吗？", a: "不可以，被认为不礼貌" },
    { q: "日本的樱花在什么季节开放？", a: "春天（3-4月）" },
    { q: "新干线的时速是多少？", a: "320公里！" },
  ], C.japan);
  countSlide(); countSlide();
})();


// ══════════════════════════════════════════════════════════════
// ══════════════ INDIA SECTION (7 slides) ═════════════════════
// ══════════════════════════════════════════════════════════════

// SLIDE 25 — 🇮🇳 印度概览 India Overview
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🇮🇳 印度概览  India Overview", C.india);

  // Taj Mahal LEFT
  safeImage(s, "india_taj_mahal.jpg", "Taj Mahal 泰姬陵", 0.3, 0.95, 5.0, 3.5);

  // Info card RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.india);

  const items = [
    { text: "🏴 国旗：橙白绿三色+蓝色法轮", y: 1.1 },
    { text: "👥 人口：约14亿（世界第一！）", y: 1.55 },
    { text: "🗣️ 语言：印地语+英语", y: 2.0 },
    { text: "        （22种官方语言！）", y: 2.35 },
    { text: "🏛️ 首都：新德里 New Delhi", y: 2.7 },
    { text: "💡 发明了数字「0」！", y: 3.15 },
    { text: "🐘 国家动物：孟加拉虎", y: 3.6 },
    { text: "🎬 宝莱坞：世界最大电影业", y: 4.05 },
  ];

  items.forEach((it) => {
    s.addText(it.text, {
      x: 5.85, y: it.y, w: 3.5, h: 0.38,
      fontSize: 12, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "印度是世界上人口最多的民主国家！", C.india);
})();


// SLIDE 26 — 🏞️ 恒河 The Ganges River
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🏞️ 恒河  The Ganges River", C.india);

  // Ganges photo LEFT
  safeImage(s, "india_ganges.jpg", "Ganges River 恒河", 0.3, 0.95, 5.0, 3.5);

  // Info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.india);
  s.addText([
    { text: "恒河是印度最神圣的河流", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "全长2,525公里", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "印度人在恒河里沐浴祈祷", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "相信恒河水能洗去罪孽", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "瓦拉纳西是恒河边最神圣的城市", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "每天傍晚有恒河祭典仪式", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 2.8, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 4.0, w: 3.5, h: 0.02, fill: { color: C.india },
  });

  s.addText("恒河对印度人就像黄河对中国人一样重要！", {
    x: 5.85, y: 4.1, w: 3.5, h: 0.7,
    fontSize: 12, fontFace: "Calibri", color: C.india, bold: true, align: "left",
  });
})();


// SLIDE 27 — 🎨 洒红节与排灯节 Holi & Diwali
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🎨 洒红节与排灯节  Holi & Diwali", C.india);

  // Holi photo LEFT
  safeImage(s, "india_holi.jpg", "Holi Festival 洒红节", 0.3, 0.95, 5.0, 3.5);

  // Festival info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.india);
  s.addText([
    { text: "🎨 洒红节 Holi", options: { fontSize: 14, fontFace: "Georgia", color: C.india, bold: true, breakLine: true } },
    { text: "颜色节！互相泼彩色粉末", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "庆祝春天到来", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "全世界最多彩的节日", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "🪔 排灯节 Diwali", options: { fontSize: 14, fontFace: "Georgia", color: C.india, bold: true, breakLine: true } },
    { text: "灯光节，像中国的春节！", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "家家户户点灯", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "放烟花，吃甜点", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "庆祝光明战胜黑暗", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 3.8, valign: "top" });

  footerText(s, "印度节日多彩多姿，一年有上百个节日！", C.india);
})();


// SLIDE 28 — 🍛 印度美食与饮食禁忌 Indian Food & Taboos
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🍛 印度美食与饮食禁忌  Indian Food & Taboos", C.india);

  // Food info LEFT
  card(s, 0.3, 0.95, 5.8, 4.3, C.bgOrange, C.india);
  s.addText([
    { text: "🍛 咖喱 curry", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    印度最有名的食物，种类超多！", options: { fontSize: 12, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🫓 飞饼 naan", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    用来蘸咖喱吃", options: { fontSize: 12, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "☕ 奶茶 chai", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    印度的国民饮料！", options: { fontSize: 12, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🌶️ 香料 spices", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    咖喱粉、姜黄、小茴香...", options: { fontSize: 12, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 0.55, y: 1.1, w: 5.3, h: 2.6, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.55, y: 3.8, w: 5.3, h: 0.02, fill: { color: C.india },
  });
  s.addText([
    { text: "⚠️ 饮食禁忌:", options: { fontSize: 13, fontFace: "Georgia", color: C.india, bold: true, breakLine: true } },
    { text: "🐄 印度教徒不吃牛肉——牛是神圣的！", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "🐷 穆斯林不吃猪肉", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "🥬 很多印度人吃素(vegetarian)!", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
  ], { x: 0.55, y: 3.85, w: 5.3, h: 1.3, valign: "top" });

  // Side facts card RIGHT
  card(s, 6.35, 0.95, 3.3, 4.3, C.white, C.india);
  s.addText("🌶️ 印度饮食特色", {
    x: 6.55, y: 1.05, w: 2.9, h: 0.4,
    fontSize: 14, fontFace: "Georgia", color: C.india, bold: true, align: "center",
  });

  const factItems = [
    "印度是世界上\n素食人口最多的国家",
    "印度有超过30种\n不同的咖喱！",
    "印度的香料出口\n占全球70%",
    "印度奶茶(chai)\n每天消费超过\n10亿杯！",
  ];

  factItems.forEach((txt, i) => {
    s.addText(txt, {
      x: 6.55, y: 1.55 + i * 0.85, w: 2.9, h: 0.75,
      fontSize: 11, fontFace: "Calibri", color: C.black, align: "center", valign: "middle",
    });
    if (i < factItems.length - 1) {
      s.addShape(pptx.shapes.RECTANGLE, {
        x: 6.85, y: 2.3 + i * 0.85, w: 2.3, h: 0.01, fill: { color: C.india },
      });
    }
  });
})();


// SLIDE 29 — ⚠️ 印度旅行礼节 India Travel Etiquette
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "⚠️ 印度旅行礼节  India Travel Etiquette", C.india);

  card(s, 0.3, 1.0, 9.3, 4.2, C.lightAmber, C.india);
  s.addText("🇮🇳 在印度旅行要注意：", {
    x: 0.55, y: 1.1, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.india, bold: true,
  });

  const tips = [
    { icon: "🙏", text: "双手合十说「Namaste」(你好)" },
    { icon: "🤚", text: "用右手吃饭和递东西（左手被认为不干净）" },
    { icon: "🚫", text: "不要摸别人的头（头是神圣的）" },
    { icon: "👟", text: "进寺庙要脱鞋" },
    { icon: "🐄", text: "很多人不吃牛肉（宗教原因）" },
    { icon: "👗", text: "穿着要保守，特别是去寺庙" },
    { icon: "🤝", text: "男性不要主动和女性握手" },
  ];

  tips.forEach((t, i) => {
    const y = 1.65 + i * 0.48;
    s.addText(t.icon + "  " + t.text, {
      x: 0.8, y: y, w: 8.5, h: 0.42,
      fontSize: 14, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "印度是一个多宗教、多文化的国家！", C.india);
})();


// SLIDE 30-31 — ✅ 印度 Check Understanding (1/2, 2/2)
(() => {
  quizSlides("✅ 印度 Check Understanding", [
    { q: "印度有多少种官方语言？", a: "22种！" },
    { q: "印度人见面怎么打招呼？", a: "双手合十说Namaste" },
    { q: "为什么很多印度人不吃牛肉？", a: "牛是神圣的" },
    { q: "印度人发明了什么数字？", a: "数字 0 ！" },
    { q: "印度最神圣的河流叫什么？", a: "恒河 Ganges" },
    { q: "洒红节(Holi)是什么节日？", a: "颜色节，互相泼彩色粉末" },
    { q: "为什么很多印度人吃素？", a: "宗教信仰" },
    { q: "排灯节(Diwali)像中国的什么节日？", a: "春节" },
    { q: "在印度用哪只手吃饭？", a: "右手" },
    { q: "在印度能用左手递东西给别人吗？", a: "不能，左手被认为不干净" },
  ], C.india);
  countSlide(); countSlide();
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 32 — 🎭 Mini Role Play
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
    { flag: "🇨🇳", country: "中国 China", action: "握手 + 「你好！」", bg: C.lightRed, color: C.china },
    { flag: "🇯🇵", country: "日本 Japan", action: "鞠躬 + 「こんにちは」", bg: C.bgPink, color: C.japan },
    { flag: "🇮🇳", country: "印度 India", action: "合十 + 「Namaste」", bg: C.bgOrange, color: C.india },
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
// SLIDE 33 — 🏆 上午竞赛
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

  card(s, 1.0, 1.3, 7.8, 3.5, C.secondary);
  s.addText([
    { text: "分组比赛！Team Challenge!", options: { fontSize: 20, fontFace: "Georgia", color: C.primary, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "Round 1: 说出3个国家的首都", options: { fontSize: 14, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "Round 2: 每个国家的打招呼方式", options: { fontSize: 14, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "Round 3: 说出每个国家不能做的事", options: { fontSize: 14, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "Round 4: 用中文介绍一个国家的食物", options: { fontSize: 14, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "最快最准确的组得分！", options: { fontSize: 16, fontFace: "Georgia", color: C.primary, bold: true, breakLine: true } },
  ], { x: 1.3, y: 1.5, w: 7.2, h: 3.0, align: "center", valign: "middle" });
})();


// ══════════════════════════════════════════════════════════════
// ══════════════ AFTERNOON SESSION ════════════════════════════
// ══════════════════════════════════════════════════════════════

// SLIDE 34 — ☀️ 下午开始 Afternoon Session
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


// SLIDE 35 — 📖 快速复习 Quick Review
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
    { flag: "🇨🇳", country: "中国", hint: "___?", color: C.china, bg: C.lightRed },
    { flag: "🇯🇵", country: "日本", hint: "___?", color: C.japan, bg: C.bgPink },
    { flag: "🇮🇳", country: "印度", hint: "___?", color: C.india, bg: C.bgOrange },
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


// SLIDE 36 — 🌏 亚洲三国文化对比 Comparison Table
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🌏 亚洲三国文化对比  Comparison", C.dark);

  const cols = [1.8, 2.4, 2.4, 2.4];
  const startX = 0.4;
  const startY = 0.95;
  const rowH = 0.63;

  const headers = ["", "🇨🇳 中国", "🇯🇵 日本", "🇮🇳 印度"];
  const headerColors = [C.dark, C.china, C.japan, C.india];

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
    ["打招呼", "握手", "鞠躬", "合十 Namaste"],
    ["吃饭工具", "筷子", "筷子", "手/右手"],
    ["重要节日", "春节", "女儿节", "排灯节 Diwali"],
    ["代表食物", "饺子", "寿司", "咖喱"],
    ["不能做的事", "筷子插饭里", "大声说话", "左手递东西"],
    ["特别的动物", "熊猫 🐼", "鹤 🦢", "牛 🐄 (神圣)"],
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


// SLIDE 37 — 🤔 共同点与不同 Similarities & Differences
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
    { text: "都有悠久的历史和文化", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "都用米饭做主食", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "都重视礼节和尊重长辈", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "都有丰富的节日文化", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
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
    { text: "打招呼方式不同（握手/鞠躬/合十）", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "吃饭工具不同（筷子 vs 手）", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "宗教信仰不同", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "食物口味不同", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "    中国多样 / 日本清淡 / 印度香料", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 5.4, y: 1.7, w: 3.9, h: 2.8, valign: "top" });
})();


// SLIDE 38 — 🧳 旅行小贴士 Travel Tips Summary
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🧳 旅行小贴士  Travel Tips Summary", C.dark);

  s.addText("「如果你去亚洲旅行，要记住：」", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.45,
    fontSize: 16, fontFace: "Georgia", color: C.dark, bold: true, align: "center",
  });

  const tips = [
    { flag: "🇨🇳", country: "在中国", tips: "学会用筷子！\n不要把筷子插饭里！", color: C.china, bg: C.lightRed },
    { flag: "🇯🇵", country: "在日本", tips: "记得鞠躬！进屋脱鞋！\n吃面可以出声！", color: C.japan, bg: C.bgPink },
    { flag: "🇮🇳", country: "在印度", tips: "用右手！\n双手合十说Namaste！\n不要摸头！", color: C.india, bg: C.bgOrange },
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


// SLIDE 39 — 📝 生词卡 Vocabulary Cards
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "📝 生词卡  Vocabulary Cards", C.dark);

  const words = [
    { zh: "饺子", py: "jiao zi", en: "dumplings" },
    { zh: "面条", py: "mian tiao", en: "noodles" },
    { zh: "寿司", py: "shou si", en: "sushi" },
    { zh: "拉面", py: "la mian", en: "ramen" },
    { zh: "咖喱", py: "ga li", en: "curry" },
    { zh: "筷子", py: "kuai zi", en: "chopsticks" },
    { zh: "鞠躬", py: "ju gong", en: "bow" },
    { zh: "合十", py: "he shi", en: "palms together" },
    { zh: "国旗", py: "guo qi", en: "national flag" },
    { zh: "礼节", py: "li jie", en: "etiquette" },
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


// SLIDE 40 — 💬 句型练习 Sentence Patterns
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "💬 句型练习  Sentence Patterns", C.dark);

  const patterns = [
    { pattern: "「我喜欢吃___」", eng: "I like to eat ___", example: "我喜欢吃饺子/寿司/咖喱", color: C.china },
    { pattern: "「在___，人们用___打招呼」", eng: "In ___, people greet by ___", example: "在日本，人们用鞠躬打招呼", color: C.japan },
    { pattern: "「去___旅行要注意___」", eng: "When traveling to ___, be careful about ___", example: "去印度旅行要注意用右手吃饭", color: C.india },
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


// SLIDE 41 — 🎭 Role Play (10-15min)
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🎭 Role Play  角色扮演 (10-15 min)", C.purple);

  s.addText("「你到了亚洲三个国家旅行！」", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.4,
    fontSize: 15, fontFace: "Georgia", color: C.purple, bold: true, align: "center",
  });

  const rounds = [
    { round: "Round 1", scene: "到中国", action: "用筷子吃饺子\n和人握手说「你好」", color: C.china, bg: C.lightRed },
    { round: "Round 2", scene: "到日本", action: "鞠躬，脱鞋进屋\n吃寿司说「いただきます」", color: C.japan, bg: C.bgPink },
    { round: "Round 3", scene: "到印度", action: "合十说「Namaste」\n用右手吃咖喱", color: C.india, bg: C.bgOrange },
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
    { level: "🟣 Level 4", desc: "完整对话 + 介绍文化礼节", x: 6.5, color: C.purple },
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


// SLIDE 42 — 🎨 Project Time! Passport Template
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🎨 Project Time!  护照亚洲页", C.teal);

  s.addText("Passport 亚洲页 Template", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.teal, bold: true, align: "center",
  });

  const fields = [
    { icon: "🌎", label: "Where am I?", value: "亚洲 (中国/日本/印度)" },
    { icon: "👀", label: "What did I see?", value: "长城 / 富士山 / 泰姬陵" },
    { icon: "🍜", label: "What did I eat?", value: "饺子 / 寿司 / 咖喱" },
    { icon: "🎭", label: "Cultural Discovery", value: "鞠躬 / 合十 / 筷子文化" },
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


// SLIDE 43 — 📊 Project 分层
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
  s.addText("画图 + 写词  (饺子 / 寿司 / 咖喱)", {
    x: 0.5, y: 1.45, w: 8.8, h: 0.55,
    fontSize: 13, fontFace: "Calibri", color: C.black,
  });

  card(s, 0.3, 2.4, 9.3, 1.2, C.bgBlue, "1565C0");
  s.addText("🔵 Level 2-3 Intermediate", {
    x: 0.5, y: 2.45, w: 4.0, h: 0.35,
    fontSize: 14, fontFace: "Georgia", color: "1565C0", bold: true,
  });
  s.addText("「I see ___. I eat ___. In Japan, people bow.」", {
    x: 0.5, y: 2.85, w: 8.8, h: 0.55,
    fontSize: 13, fontFace: "Calibri", color: C.black,
  });

  card(s, 0.3, 3.8, 9.3, 1.4, "F3E5F5", C.purple);
  s.addText("🟣 Level 4 Advanced", {
    x: 0.5, y: 3.85, w: 3.5, h: 0.35,
    fontSize: 14, fontFace: "Georgia", color: C.purple, bold: true,
  });
  s.addText("「In China, I saw the Great Wall. I ate dumplings with chopsticks.\nWhen traveling there, don't stick chopsticks in rice because...」", {
    x: 0.5, y: 4.25, w: 8.8, h: 0.75,
    fontSize: 12, fontFace: "Calibri", color: C.black,
  });
})();


// SLIDE 44 — 📄 Project 示范 Example
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "📄 Project 示范  Example Page", C.teal);

  card(s, 1.0, 1.0, 8.0, 4.2, C.white, C.teal);
  s.addText("🌏 My Asia Page  我的亚洲页", {
    x: 1.2, y: 1.1, w: 7.6, h: 0.5,
    fontSize: 18, fontFace: "Georgia", color: C.teal, bold: true, align: "center",
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 2.0, y: 1.6, w: 6.0, h: 0.02, fill: { color: C.teal },
  });

  const entries = [
    "🌎 I visited: 亚洲 Asia \u2014 中国, 日本, 印度",
    "👀 I saw: 长城 Great Wall, 富士山 Mt. Fuji, 泰姬陵 Taj Mahal",
    "🍜 I ate: 饺子 dumplings with 筷子 chopsticks!",
    "🎭 I learned: 在日本要鞠躬，在印度说Namaste",
    "💬 My sentence: 我最喜欢吃寿司。在中国，人们用筷子吃饭。",
    "🎨 [Draw your favorite food or landmark here!]",
  ];

  entries.forEach((e, i) => {
    s.addText(e, {
      x: 1.4, y: 1.75 + i * 0.52, w: 7.2, h: 0.45,
      fontSize: 13, fontFace: "Calibri", color: i === 5 ? C.gray : C.black,
      align: "left", valign: "middle",
    });
  });
})();


// SLIDE 45 — 🗣️ 分享时间 Sharing Time
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🗣️ 分享时间  Sharing Time", C.dark);

  s.addText("和同伴分享你的亚洲页！\nShare your Asia page with a partner!", {
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


// SLIDE 46 — 🪪 亚洲签证章
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.dark };
  goldBars(s);
  slideNum(s);

  s.addText("🪪 亚洲签证章  Asia Visa Stamp", {
    x: 0.5, y: 0.4, w: 9.0, h: 0.6,
    fontSize: 26, fontFace: "Georgia", color: C.gold, bold: true, align: "center",
  });

  s.addShape(pptx.shapes.OVAL, {
    x: 3.0, y: 1.2, w: 4.0, h: 3.5,
    fill: { color: C.dark },
    line: { color: C.contAsia, width: 4, dashType: "dash" },
  });

  s.addText([
    { text: "ASIA", options: { fontSize: 34, fontFace: "Georgia", color: C.contAsia, bold: true, breakLine: true } },
    { text: "亚洲", options: { fontSize: 24, fontFace: "Georgia", color: C.gold, breakLine: true } },
    { text: "\u2713 VISITED", options: { fontSize: 20, fontFace: "Georgia", color: C.quizGreen, bold: true, breakLine: true } },
    { text: "6/8/2025", options: { fontSize: 14, fontFace: "Calibri", color: C.white, breakLine: true } },
    { text: "中国 \u00B7 日本 \u00B7 印度", options: { fontSize: 14, fontFace: "Calibri", color: C.lightAmber, breakLine: true } },
  ], { x: 3.0, y: 1.4, w: 4.0, h: 3.1, align: "center", valign: "middle" });

  s.addText("恭喜你完成亚洲之旅！Congratulations!", {
    x: 0.5, y: 4.85, w: 9.0, h: 0.4,
    fontSize: 14, fontFace: "Calibri", color: C.gold, bold: true, align: "center",
  });
})();


// SLIDE 47 — ✈️ 明天航班 Tomorrow's Flight
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.dark };
  goldBars(s);
  slideNum(s);

  s.addText([
    { text: "✈️ 明天航班", options: { fontSize: 30, fontFace: "Georgia", color: C.gold, bold: true, breakLine: true } },
    { text: "Tomorrow's Flight", options: { fontSize: 18, fontFace: "Georgia", color: C.white, breakLine: true } },
  ], { x: 0.5, y: 0.5, w: 9.0, h: 1.2, align: "center" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 3.0, y: 1.7, w: 4.0, h: 0.04, fill: { color: C.gold },
  });

  card(s, 1.5, 2.0, 7.0, 2.6, C.secondary);
  s.addText([
    { text: "Flight 航班: GR-002", options: { fontSize: 18, fontFace: "Georgia", color: C.contAfrica, bold: true, breakLine: true } },
    { text: "Destination 目的地: 非洲 AFRICA", options: { fontSize: 18, fontFace: "Georgia", color: C.contAfrica, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "明天我们去非洲！", options: { fontSize: 16, fontFace: "Georgia", color: C.dark, bold: true, breakLine: true } },
    { text: "那里的人怎么打招呼？", options: { fontSize: 16, fontFace: "Georgia", color: C.dark, breakLine: true } },
    { text: "Tomorrow we fly to Africa! How do people greet there?", options: { fontSize: 12, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 1.8, y: 2.1, w: 6.4, h: 2.4, align: "center", valign: "middle" });

  s.addText("See you tomorrow, explorers!  明天见，小探险家们！", {
    x: 0.5, y: 4.85, w: 9.0, h: 0.4,
    fontSize: 14, fontFace: "Calibri", color: C.gold, bold: true, align: "center",
  });
})();


// ── Save ──
const outPath = path.join(__dirname, "day1_asia.pptx");
pptx.writeFile({ fileName: outPath }).then(() => {
  console.log("Created: " + outPath);
  console.log("Total slides: " + slideCount);
}).catch((err) => {
  console.error("Error:", err);
});
