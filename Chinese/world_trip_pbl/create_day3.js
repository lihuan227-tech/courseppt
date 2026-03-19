/**
 * Day 3: 欧洲 Europe (6/10) — Global Explorer Camp 环球探索沉浸式夏令营
 * ~44 slides — 3 countries (法国 France, 意大利 Italy, 英国 UK)
 * Each country: expanded with dedicated topic slides + large images
 * Run: node create_day3.js
 */
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const pptx = new pptxgen();
pptx.defineLayout({ name: "LAYOUT_16x9", width: 10.0, height: 5.625 });
pptx.layout = "LAYOUT_16x9";
pptx.author = "谷雨中文 GR EDU";
pptx.title = "Global Explorer Camp · Day 3: 欧洲 Europe";

// ── Colors (NO # prefix) ──
const C = {
  primary:    "1565C0",
  secondary:  "E3F2FD",
  accent:     "FFC107",
  dark:       "0D47A1",
  white:      "FFFFFF",
  black:      "212121",
  gray:       "616161",
  gold:       "FFD54F",
  darkGold:   "FFA000",
  lightBlue:  "BBDEFB",
  lightAmber: "FFE0B2",
  bgBlue:     "E3F2FD",
  bgGreen:    "E8F5E9",
  bgOrange:   "FFF3E0",
  bgPink:     "FCE4EC",
  bgRed:      "FFEBEE",
  france:     "1565C0",
  italy:      "388E3C",
  uk:         "C62828",
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
  franceBg:   "E3F2FD",
  italyBg:    "E8F5E9",
  ukBg:       "FFEBEE",
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
    { text: "Flight 航班: GR-003", options: { fontSize: 16, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "Destination 目的地:  欧洲 EUROPE", options: { fontSize: 16, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
    { text: "Date 日期: June 10, 2025  (6/10)", options: { fontSize: 14, fontFace: "Calibri", color: C.gray, breakLine: true } },
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

  const stamps = [
    { label: "Day 1\n亚洲 Asia", color: C.contAsia, done: true },
    { label: "Day 2\n非洲 Africa", color: C.contAfrica, done: true },
    { label: "Day 3\n欧洲 Europe", color: C.contEurope, done: false },
    { label: "Day 4\n美洲 Americas", color: C.contNA, done: false },
    { label: "Day 5\n展览 Exhibition", color: C.teal, done: false },
  ];

  stamps.forEach((st, i) => {
    const x = 0.4 + i * 1.85;
    const bw = 1.7;

    s.addShape(pptx.shapes.OVAL, {
      x: x + 0.1, y: 1.3, w: 1.5, h: 1.5,
      fill: { color: st.done ? st.color : C.white },
      line: { color: st.color, width: st.done ? 3 : 2, dashType: st.done ? "solid" : "dash" },
    });
    s.addText(st.label, {
      x: x + 0.1, y: 1.4, w: 1.5, h: 1.1,
      fontSize: 12, fontFace: "Georgia", color: st.done ? C.white : st.color, bold: true, align: "center", valign: "middle",
    });
    if (st.done) {
      s.addText("\u2713", {
        x: x + 0.1, y: 2.2, w: 1.5, h: 0.5,
        fontSize: 20, fontFace: "Georgia", color: C.white, bold: true, align: "center",
      });
    }
  });

  card(s, 0.5, 3.2, 9.0, 2.0, C.white, C.contEurope);
  s.addText([
    { text: "已完成: 亚洲 \u2713  非洲 \u2713", options: { fontSize: 16, fontFace: "Georgia", color: C.dark, bold: true, breakLine: true } },
    { text: "今天: 欧洲 \u2014 法国、意大利、英国!", options: { fontSize: 18, fontFace: "Georgia", color: C.contEurope, bold: true, breakLine: true } },
    { text: "2 stamps collected, 3 more to go!", options: { fontSize: 13, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 0.8, y: 3.35, w: 8.4, h: 1.7, align: "center", valign: "middle" });
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

  s.addText("深入了解欧洲3个国家", {
    x: 0.5, y: 0.95, w: 9.0, h: 0.5,
    fontSize: 22, fontFace: "Georgia", color: C.dark, bold: true, align: "center",
  });

  const countries = [
    { flag: "🇫🇷", name: "法国 France", color: C.france, bg: C.franceBg },
    { flag: "🇮🇹", name: "意大利 Italy", color: C.italy, bg: C.italyBg },
    { flag: "🇬🇧", name: "英国 UK", color: C.uk, bg: C.ukBg },
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
    { text: "4. 比较三个国家的异同 + 丝绸之路连接", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
  ], { x: 0.8, y: 3.55, w: 8.4, h: 1.5, valign: "top" });
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 4 — 认识欧洲 About Europe
// ══════════════════════════════════════════════════════════════
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🌍 认识欧洲  About Europe", C.contEurope);

  const stats = [
    { icon: "🏛️", label: "44个国家", val: "44 Countries" },
    { icon: "🎨", label: "文艺复兴发源地", val: "Birthplace of Renaissance" },
    { icon: "💡", label: "工业革命起源", val: "Industrial Revolution" },
    { icon: "👥", label: "7.5亿人口", val: "750 Million People" },
    { icon: "🗣️", label: "200+种语言", val: "200+ Languages" },
    { icon: "🏰", label: "古堡与教堂", val: "Castles & Cathedrals" },
  ];

  stats.forEach((st, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.3 + col * 3.15;
    const y = 1.0 + row * 2.1;

    card(s, x, y, 2.9, 1.8, C.white, C.contEurope);
    s.addText(st.icon, {
      x: x, y: y + 0.1, w: 2.9, h: 0.6,
      fontSize: 30, fontFace: "Calibri", align: "center",
    });
    s.addText(st.label, {
      x: x, y: y + 0.7, w: 2.9, h: 0.5,
      fontSize: 16, fontFace: "Georgia", color: C.contEurope, bold: true, align: "center",
    });
    s.addText(st.val, {
      x: x, y: y + 1.2, w: 2.9, h: 0.45,
      fontSize: 12, fontFace: "Calibri", color: C.gray, align: "center",
    });
  });

  footerText(s, "欧洲是现代文明的重要发源地！", C.contEurope);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 5 — 欧洲地图 Europe Map
// ══════════════════════════════════════════════════════════════
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: "E8F0FE" };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🗺️ 欧洲地图  Europe Map", C.contEurope);

  const regions = [
    { label: "🇬🇧 英国\nUK", x: 1.5, y: 1.2, w: 1.6, h: 1.4, color: C.uk },
    { label: "🇫🇷 法国\nFrance", x: 2.5, y: 2.5, w: 2.0, h: 1.6, color: C.france },
    { label: "🇩🇪 德国\nGermany", x: 4.2, y: 1.5, w: 1.5, h: 1.3, color: "616161" },
    { label: "🇮🇹 意大利\nItaly", x: 4.5, y: 2.8, w: 1.3, h: 2.0, color: C.italy },
    { label: "🇪🇸 西班牙\nSpain", x: 1.0, y: 3.5, w: 1.8, h: 1.3, color: "FF8F00" },
    { label: "北欧\nNordic", x: 4.5, y: 0.9, w: 2.5, h: 1.0, color: "00897B" },
    { label: "东欧\nE. Europe", x: 6.5, y: 1.5, w: 2.5, h: 2.5, color: "90A4AE" },
    { label: "🇬🇷 希腊\nGreece", x: 6.0, y: 3.5, w: 1.5, h: 1.2, color: "1976D2" },
  ];

  regions.forEach((r) => {
    const isTarget = r.label.includes("法国") || r.label.includes("意大利") || r.label.includes("英国");
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: r.x, y: r.y, w: r.w, h: r.h, rectRadius: 0.2,
      fill: { color: r.color, transparency: isTarget ? 15 : 60 },
      line: isTarget ? { color: r.color, width: 2.5 } : undefined,
    });
    s.addText(r.label, {
      x: r.x, y: r.y, w: r.w, h: r.h,
      fontSize: isTarget ? 13 : 10, fontFace: "Georgia",
      color: C.white, bold: true,
      align: "center", valign: "middle",
    });
  });

  footerText(s, "今天我们去这三个国家！Today we visit these 3 countries!", C.contEurope);
})();


// ══════════════════════════════════════════════════════════════
// ══════════════ FRANCE SECTION (8 slides) ════════════════════
// ══════════════════════════════════════════════════════════════

// SLIDE 6 — 🇫🇷 法国概览 France Overview
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🇫🇷 法国概览  France Overview", C.france);

  safeImage(s, "europe_eiffel.jpg", "Eiffel Tower 埃菲尔铁塔", 0.3, 0.95, 5.0, 3.5);

  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.france);

  const items = [
    { text: "🏴 国旗：蓝白红 三色旗", y: 1.1 },
    { text: "👥 人口：约6700万", y: 1.6 },
    { text: "🗣️ 语言：法语 French", y: 2.1 },
    { text: "🏛️ 首都：巴黎 Paris", y: 2.6 },
    { text: "🍷 世界美食之都", y: 3.1 },
    { text: "🎨 艺术与时尚中心", y: 3.6 },
    { text: "💡 启蒙运动发源地", y: 4.1 },
  ];

  items.forEach((it) => {
    s.addText(it.text, {
      x: 5.85, y: it.y, w: 3.5, h: 0.42,
      fontSize: 13, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "法国是世界上最受欢迎的旅游国家！France is the most visited country!", C.france);
})();


// SLIDE 7 — 🏛️ 首都：巴黎 Capital: Paris
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🏛️ 首都：巴黎  Capital: Paris", C.france);

  safeImage(s, "europe_paris.jpg", "Paris 巴黎", 0.3, 0.95, 5.0, 3.5);

  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.france);
  s.addText([
    { text: "巴黎被称为「浪漫之都」", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "塞纳河穿过城市中心", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "每年2200万游客来访", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "世界时尚之都", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "2024年举办了奥运会", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 2.8, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 4.0, w: 3.5, h: 0.02, fill: { color: C.france },
  });

  s.addText("巴黎的名字来自古代高卢部落「Parisii」！", {
    x: 5.85, y: 4.1, w: 3.5, h: 0.7,
    fontSize: 12, fontFace: "Calibri", color: C.france, bold: true, align: "left",
  });
})();


// SLIDE 8 — 🗼 埃菲尔铁塔 Eiffel Tower
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🗼 埃菲尔铁塔  Eiffel Tower", C.france);

  safeImage(s, "europe_eiffel.jpg", "Eiffel Tower 埃菲尔铁塔", 0.3, 0.95, 5.0, 4.0);

  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.france);
  s.addText([
    { text: "1889年建成", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "高324米（约100层楼高！）", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "原计划20年后拆除", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "  （因为广播天线被保留下来了）", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "每年700万人参观", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "晚上整点会闪灯5分钟", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "用了750万公斤的铁", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 3.5, valign: "top" });

  footerText(s, "铁塔本来只是临时建筑，现在是法国的标志！", C.france);
})();


// SLIDE 9 — 🖼️ 卢浮宫与蒙娜丽莎 Louvre & Mona Lisa
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🖼️ 卢浮宫与蒙娜丽莎  Louvre & Mona Lisa", C.france);

  safeImage(s, "europe_louvre.jpg", "Louvre 卢浮宫", 0.3, 0.95, 5.0, 3.5);

  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.france);
  s.addText([
    { text: "世界最大的博物馆", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "拥有38万件藏品", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "玻璃金字塔入口是标志", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "最出名的是蒙娜丽莎", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "蒙娜丽莎很小！", options: { fontSize: 13, fontFace: "Calibri", color: C.france, bold: true, bullet: true, breakLine: true } },
    { text: "  只有 77 x 53 cm!", options: { fontSize: 12, fontFace: "Calibri", color: C.france, bold: true, breakLine: true } },
    { text: "全部看完要走约15公里", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 3.5, valign: "top" });

  footerText(s, "卢浮宫每年有约1000万游客！The most visited museum!", C.france);
})();


// SLIDE 10 — 🥐 法国美食 French Food
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🥐 法国美食  French Food", C.france);

  safeImage(s, "europe_croissant.jpg", "Croissant 可颂", 0.3, 0.95, 5.0, 3.5);

  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.france);

  const foods = [
    { name: "🥐 可颂 Croissant", detail: "法国人每天早上的最爱", y: 1.1 },
    { name: "🥖 法棍 Baguette", detail: "法国人每天吃3000万根！", y: 1.85 },
    { name: "🧀 奶酪 Cheese", detail: "法国有400多种奶酪", y: 2.6 },
    { name: "🐌 蜗牛 Escargot", detail: "法国的特色菜！", y: 3.35 },
    { name: "🍬 马卡龙 Macaron", detail: "五颜六色的甜点", y: 4.1 },
  ];

  foods.forEach((f) => {
    s.addText(f.name, {
      x: 5.85, y: f.y, w: 3.5, h: 0.35,
      fontSize: 13, fontFace: "Georgia", color: C.france, bold: true, align: "left",
    });
    s.addText(f.detail, {
      x: 5.85, y: f.y + 0.35, w: 3.5, h: 0.3,
      fontSize: 11, fontFace: "Calibri", color: C.gray, align: "left",
    });
  });
})();


// SLIDE 11 — ⚠️ 法国礼节 French Etiquette
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "⚠️ 法国旅行礼节  French Etiquette", C.france);

  card(s, 0.3, 1.0, 9.3, 4.2, C.franceBg, C.france);
  s.addText("🇫🇷 在法国旅行要注意：", {
    x: 0.55, y: 1.1, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.france, bold: true,
  });

  const tips = [
    { icon: "💋", text: "见面亲脸颊「la bise」（左右各一次或多次）" },
    { icon: "🍽️", text: "不要催服务员（被认为很不礼貌！）" },
    { icon: "🥖", text: "面包放在桌上，不放在盘子里" },
    { icon: "⏰", text: "午餐通常1-2个小时（法国人很重视吃饭）" },
    { icon: "🗣️", text: "打招呼先说「Bonjour!」（你好）" },
    { icon: "🏪", text: "商店周日经常不营业" },
    { icon: "💐", text: "送花不要送菊花（只在葬礼用）" },
  ];

  tips.forEach((t, i) => {
    const y = 1.65 + i * 0.48;
    s.addText(t.icon + "  " + t.text, {
      x: 0.8, y: y, w: 8.5, h: 0.42,
      fontSize: 14, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "在法国，慢慢享受生活是一种艺术！", C.france);
})();


// SLIDE 12-13 — ✅ 法国 Check Understanding (1/2, 2/2)
(() => {
  quizSlides("✅ 法国 Check Understanding", [
    { q: "法国的国旗是什么颜色？", a: "蓝、白、红" },
    { q: "法国的首都是哪里？", a: "巴黎 Paris" },
    { q: "埃菲尔铁塔有多高？", a: "324米" },
    { q: "埃菲尔铁塔原来计划多少年后拆？", a: "20年" },
    { q: "卢浮宫最有名的画是什么？", a: "蒙娜丽莎 Mona Lisa" },
    { q: "蒙娜丽莎有多大？", a: "只有77 x 53 cm！" },
    { q: "法国有多少种奶酪？", a: "400多种" },
    { q: "法国人见面怎么打招呼？", a: "亲脸颊 la bise" },
    { q: "法国人每天吃多少根法棍？", a: "3000万根" },
    { q: "在法国为什么不能催服务员？", a: "被认为很不礼貌" },
  ], C.france);
  countSlide(); countSlide();
})();


// ══════════════════════════════════════════════════════════════
// ══════════════ ITALY SECTION (7 slides) ═════════════════════
// ══════════════════════════════════════════════════════════════

// SLIDE 14 — 🇮🇹 意大利概览 Italy Overview
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🇮🇹 意大利概览  Italy Overview", C.italy);

  safeImage(s, "europe_colosseum.jpg", "Colosseum 斗兽场", 0.3, 0.95, 5.0, 3.5);

  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.italy);

  const items = [
    { text: "🏴 国旗：绿白红 三色旗", y: 1.1 },
    { text: "👥 人口：约5900万", y: 1.6 },
    { text: "🗣️ 语言：意大利语 Italian", y: 2.1 },
    { text: "🏛️ 首都：罗马 Rome", y: 2.6 },
    { text: "👢 国土形状像一只靴子！", y: 3.1 },
    { text: "🎨 文艺复兴发源地", y: 3.6 },
    { text: "🏛️ 罗马帝国的故乡", y: 4.1 },
  ];

  items.forEach((it) => {
    s.addText(it.text, {
      x: 5.85, y: it.y, w: 3.5, h: 0.42,
      fontSize: 13, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "意大利的国土像一只高跟靴子，你看到了吗？", C.italy);
})();


// SLIDE 15 — 🏛️ 首都：罗马 Capital: Rome
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🏛️ 首都：罗马  Capital: Rome", C.italy);

  safeImage(s, "europe_colosseum.jpg", "Colosseum 斗兽场", 0.3, 0.95, 5.0, 3.5);

  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.italy);
  s.addText([
    { text: "罗马有2700年历史", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "斗兽场能容纳5万人", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "许愿池每天收到3000欧元硬币", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "万神殿有2000年历史", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "梵蒂冈在罗马城内", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "  （世界上最小的国家！）", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 3.2, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 4.2, w: 3.5, h: 0.02, fill: { color: C.italy },
  });

  s.addText("「条条大路通罗马」All roads lead to Rome!", {
    x: 5.85, y: 4.3, w: 3.5, h: 0.6,
    fontSize: 12, fontFace: "Calibri", color: C.italy, bold: true, align: "left",
  });
})();


// SLIDE 16 — 🏗️ 比萨斜塔与威尼斯 Leaning Tower & Venice
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🏗️ 比萨斜塔与威尼斯  Pisa & Venice", C.italy);

  safeImage(s, "europe_venice.jpg", "Venice 威尼斯", 0.3, 0.95, 5.0, 4.0);

  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.italy);

  s.addText("🏗️ 比萨斜塔", {
    x: 5.85, y: 1.05, w: 3.5, h: 0.4,
    fontSize: 14, fontFace: "Georgia", color: C.italy, bold: true,
  });
  s.addText([
    { text: "倾斜约4度", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "建了200年才完工", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "地基太软导致倾斜", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.45, w: 3.5, h: 1.2, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 2.7, w: 3.5, h: 0.02, fill: { color: C.italy },
  });

  s.addText("🚣 威尼斯水城", {
    x: 5.85, y: 2.8, w: 3.5, h: 0.4,
    fontSize: 14, fontFace: "Georgia", color: C.italy, bold: true,
  });
  s.addText([
    { text: "由118个小岛组成", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "没有汽车，只有船！", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "贡多拉(Gondola)是传统小船", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "城市正在慢慢下沉", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 3.2, w: 3.5, h: 1.6, valign: "top" });

  footerText(s, "威尼斯水城 \u2014 世界上最独特的城市之一！", C.italy);
})();


// SLIDE 17 — 🍕 意大利美食 Italian Food
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🍕 意大利美食  Italian Food", C.italy);

  safeImage(s, "europe_pizza.jpg", "Pizza 披萨", 0.3, 0.95, 5.0, 3.5);

  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.italy);

  const foods = [
    { name: "🍕 披萨 Pizza", detail: "Margherita是最经典的", y: 1.1 },
    { name: "🍝 意面 Pasta", detail: "有350多种形状！", y: 1.85 },
    { name: "🍦 冰淇淋 Gelato", detail: "比普通冰淇淋更顺滑", y: 2.6 },
    { name: "☕ 浓缩咖啡 Espresso", detail: "意大利人的每日必备", y: 3.35 },
    { name: "🍰 提拉米苏 Tiramisu", detail: "意思是「带我走」！", y: 4.1 },
  ];

  foods.forEach((f) => {
    s.addText(f.name, {
      x: 5.85, y: f.y, w: 3.5, h: 0.35,
      fontSize: 13, fontFace: "Georgia", color: C.italy, bold: true, align: "left",
    });
    s.addText(f.detail, {
      x: 5.85, y: f.y + 0.35, w: 3.5, h: 0.3,
      fontSize: 11, fontFace: "Calibri", color: C.gray, align: "left",
    });
  });
})();


// SLIDE 18 — ⚠️ 意大利礼节 Italian Etiquette
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "⚠️ 意大利旅行礼节  Italian Etiquette", C.italy);

  card(s, 0.3, 1.0, 9.3, 4.2, C.italyBg, C.italy);
  s.addText("🇮🇹 在意大利旅行要注意：", {
    x: 0.55, y: 1.1, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.italy, bold: true,
  });

  const tips = [
    { icon: "👋", text: "见面说「Ciao!」+ 喜欢用手势说话" },
    { icon: "🍕", text: "不要加菠萝在披萨上！（意大利人会生气）" },
    { icon: "☕", text: "早上以后不要在咖啡里加牛奶（cappuccino只在早上喝）" },
    { icon: "⛪", text: "进教堂要穿长袖长裤（表示尊重）" },
    { icon: "🍝", text: "吃意面不要用勺子卷（用叉子就好）" },
    { icon: "🧈", text: "面包不蘸橄榄油（那是美国人的做法）" },
    { icon: "💰", text: "给小费不是必须的（和美国不一样）" },
  ];

  tips.forEach((t, i) => {
    const y = 1.65 + i * 0.48;
    s.addText(t.icon + "  " + t.text, {
      x: 0.8, y: y, w: 8.5, h: 0.42,
      fontSize: 14, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "意大利人对食物非常认真和骄傲！", C.italy);
})();


// SLIDE 19-20 — ✅ 意大利 Check Understanding (1/2, 2/2)
(() => {
  quizSlides("✅ 意大利 Check Understanding", [
    { q: "意大利的国土像什么形状？", a: "一只靴子！" },
    { q: "意大利的首都在哪里？", a: "罗马 Rome" },
    { q: "斗兽场能容纳多少人？", a: "5万人" },
    { q: "比萨斜塔倾斜多少度？", a: "约4度" },
    { q: "威尼斯有多少个小岛？", a: "118个" },
    { q: "意大利面有多少种形状？", a: "350多种" },
    { q: "「Tiramisu」是什么意思？", a: "「带我走」" },
    { q: "为什么不能加菠萝在披萨上？", a: "意大利人觉得不正宗" },
    { q: "Cappuccino什么时候喝？", a: "只在早上喝" },
    { q: "世界上最小的国家在哪个城市里？", a: "梵蒂冈在罗马城内" },
  ], C.italy);
  countSlide(); countSlide();
})();


// ══════════════════════════════════════════════════════════════
// ══════════════ UK SECTION (7 slides) ════════════════════════
// ══════════════════════════════════════════════════════════════

// SLIDE 21 — 🇬🇧 英国概览 UK Overview
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🇬🇧 英国概览  UK Overview", C.uk);

  safeImage(s, "europe_bigben.jpg", "Big Ben 大本钟", 0.3, 0.95, 5.0, 3.5);

  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.uk);

  const items = [
    { text: "🏴 国旗：Union Jack 米字旗", y: 1.1 },
    { text: "👥 人口：约6700万", y: 1.6 },
    { text: "🗣️ 语言：英语 English", y: 2.1 },
    { text: "🏛️ 首都：伦敦 London", y: 2.6 },
    { text: "👑 有国王！King Charles III", y: 3.1 },
    { text: "🏭 工业革命发源地", y: 3.6 },
    { text: "⚽ 现代足球的故乡", y: 4.1 },
  ];

  items.forEach((it) => {
    s.addText(it.text, {
      x: 5.85, y: it.y, w: 3.5, h: 0.42,
      fontSize: 13, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "英国的全名是「大不列颠及北爱尔兰联合王国」！", C.uk);
})();


// SLIDE 22 — 🏙️ 首都：伦敦 Capital: London
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🏙️ 首都：伦敦  Capital: London", C.uk);

  safeImage(s, "europe_london.jpg", "London 伦敦", 0.3, 0.95, 5.0, 3.5);

  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.uk);
  s.addText([
    { text: "大本钟 Big Ben \u2014 地标建筑", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "白金汉宫 \u2014 国王住在这里", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "伦敦塔桥 \u2014 可以打开让船通过", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "双层巴士 \u2014 红色的，很有名", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "红色电话亭 \u2014 英国的标志", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "泰晤士河穿过城市中心", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 3.2, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 4.2, w: 3.5, h: 0.02, fill: { color: C.uk },
  });

  s.addText("伦敦有300多个博物馆，大部分免费！", {
    x: 5.85, y: 4.3, w: 3.5, h: 0.6,
    fontSize: 12, fontFace: "Calibri", color: C.uk, bold: true, align: "left",
  });
})();


// SLIDE 23 — ☕ 下午茶文化 Afternoon Tea
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "☕ 下午茶文化  Afternoon Tea Culture", C.uk);

  card(s, 0.3, 0.95, 4.5, 4.2, C.white, C.uk);
  s.addText("🇬🇧 英国下午茶", {
    x: 0.5, y: 1.05, w: 4.1, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.uk, bold: true, align: "center",
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.6, y: 1.5, w: 3.9, h: 0.02, fill: { color: C.uk },
  });
  s.addText([
    { text: "英国人每天喝1.65亿杯茶！", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "下午茶时间：3-5点", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "吃scones(司康饼)和三明治", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "先倒茶再加牛奶", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "三层点心架是经典", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 0.6, y: 1.65, w: 3.9, h: 2.5, valign: "top" });

  card(s, 5.1, 0.95, 4.5, 4.2, C.white, "FF8F00");
  s.addText("🇨🇳 中国茶文化对比", {
    x: 5.3, y: 1.05, w: 4.1, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: "FF8F00", bold: true, align: "center",
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.4, y: 1.5, w: 3.9, h: 0.02, fill: { color: "FF8F00" },
  });
  s.addText([
    { text: "中国是茶的故乡（5000年历史）", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "绿茶、红茶、乌龙茶等种类多", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "功夫茶 \u2014 讲究泡茶技艺", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "茶是待客之道", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "中国茶传到英国变成下午茶！", options: { fontSize: 13, fontFace: "Calibri", color: "FF8F00", bold: true, bullet: true, breakLine: true } },
  ], { x: 5.4, y: 1.65, w: 3.9, h: 2.5, valign: "top" });

  footerText(s, "茶从中国传到英国，变成了不一样的文化！", C.uk);
})();


// SLIDE 24 — 🍟 英国美食 British Food
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🍟 英国美食  British Food", C.uk);

  safeImage(s, "europe_fishandchips.jpg", "Fish & Chips 炸鱼薯条", 0.3, 0.95, 5.0, 3.5);

  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.uk);

  const foods = [
    { name: "🐟 炸鱼薯条 Fish & Chips", detail: "英国最有名的食物！", y: 1.1 },
    { name: "🍳 英式早餐 Full English", detail: "鸡蛋+培根+香肠+豆子+吐司", y: 1.85 },
    { name: "🥧 牧羊人派 Shepherd's Pie", detail: "肉馅加土豆泥", y: 2.6 },
    { name: "🍖 Sunday Roast", detail: "周日全家一起吃烤肉大餐", y: 3.35 },
    { name: "☕ 茶 Tea", detail: "英国人的灵魂饮品", y: 4.1 },
  ];

  foods.forEach((f) => {
    s.addText(f.name, {
      x: 5.85, y: f.y, w: 3.5, h: 0.35,
      fontSize: 13, fontFace: "Georgia", color: C.uk, bold: true, align: "left",
    });
    s.addText(f.detail, {
      x: 5.85, y: f.y + 0.35, w: 3.5, h: 0.3,
      fontSize: 11, fontFace: "Calibri", color: C.gray, align: "left",
    });
  });
})();


// SLIDE 25 — ⚠️ 英国礼节 British Etiquette
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "⚠️ 英国旅行礼节  British Etiquette", C.uk);

  card(s, 0.3, 1.0, 9.3, 4.2, C.ukBg, C.uk);
  s.addText("🇬🇧 在英国旅行要注意：", {
    x: 0.55, y: 1.1, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.uk, bold: true,
  });

  const tips = [
    { icon: "🚶", text: "一定要排队！（英国人最讨厌插队）" },
    { icon: "🙏", text: "多说 Please 和 Thank you" },
    { icon: "💰", text: "不要问别人收入多少（很不礼貌）" },
    { icon: "🌧️", text: "聊天气是最安全的话题（英国经常下雨）" },
    { icon: "🚗", text: "靠左开车！（过马路先看右边）" },
    { icon: "🤫", text: "公共场所说话声音要小" },
    { icon: "👑", text: "对皇室要表示尊重" },
  ];

  tips.forEach((t, i) => {
    const y = 1.65 + i * 0.48;
    s.addText(t.icon + "  " + t.text, {
      x: 0.8, y: y, w: 8.5, h: 0.42,
      fontSize: 14, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "英国人非常有礼貌，排队是一种美德！", C.uk);
})();


// SLIDE 26-27 — ✅ 英国 Check Understanding (1/2, 2/2)
(() => {
  quizSlides("✅ 英国 Check Understanding", [
    { q: "英国的国旗叫什么名字？", a: "Union Jack 米字旗" },
    { q: "英国现在的国王是谁？", a: "King Charles III" },
    { q: "英国人每天喝多少杯茶？", a: "1.65亿杯！" },
    { q: "英国人最讨厌什么？", a: "插队！" },
    { q: "英国开车靠哪边？", a: "靠左边" },
    { q: "英国最有名的食物是什么？", a: "炸鱼薯条 Fish & Chips" },
    { q: "下午茶的时间是几点？", a: "3-5点" },
    { q: "英国人最喜欢聊什么话题？", a: "天气！" },
    { q: "白金汉宫里住着谁？", a: "国王" },
    { q: "伦敦有多少个博物馆？", a: "300多个" },
  ], C.uk);
  countSlide(); countSlide();
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 28 — 🐪 丝绸之路 Silk Road
// ══════════════════════════════════════════════════════════════
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🐪 丝绸之路  The Silk Road", C.dark);

  safeImage(s, "europe_silkroad.jpg", "Silk Road 丝绸之路", 0.3, 0.95, 5.0, 3.5);

  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.dark);
  s.addText([
    { text: "连接中国和欧洲的贸易之路", options: { fontSize: 13, fontFace: "Calibri", color: C.dark, bold: true, bullet: true, breakLine: true } },
    { text: "长约6400公里", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "中国出口：丝绸、茶叶、瓷器", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "欧洲出口：金银、宝石、玻璃", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "马可波罗 Marco Polo", options: { fontSize: 13, fontFace: "Calibri", color: C.dark, bold: true, bullet: true, breakLine: true } },
    { text: "  来自意大利，到中国旅行了24年", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "丝绸之路不只运货，还传播了文化、宗教和技术", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 3.5, valign: "top" });

  footerText(s, "丝绸之路把亚洲和欧洲连在了一起！", C.dark);
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
    { flag: "🇫🇷", country: "法国 France", action: "亲脸颊 +「Bonjour!」", bg: C.franceBg, color: C.france },
    { flag: "🇮🇹", country: "意大利 Italy", action: "手势 +「Ciao!」", bg: C.italyBg, color: C.italy },
    { flag: "🇬🇧", country: "英国 UK", action: "握手 +「Hello!」\n+ 排队！", bg: C.ukBg, color: C.uk },
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

  footerText(s, "每个国家的问候方式都不一样！Every culture greets differently!", C.purple);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 30 — 🏆 上午竞赛
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
    { text: "Round 4: 丝绸之路运了什么货物？", options: { fontSize: 14, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "最快最准确的组得分！", options: { fontSize: 16, fontFace: "Georgia", color: C.primary, bold: true, breakLine: true } },
  ], { x: 1.3, y: 1.5, w: 7.2, h: 3.0, align: "center", valign: "middle" });
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
    { text: "复习 \u2192 对比总结 \u2192 语言学习 \u2192 Project Time!", options: { fontSize: 16, fontFace: "Calibri", color: C.white, bold: true, breakLine: true } },
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
    { flag: "🇫🇷", country: "法国", hint: "___?", color: C.france, bg: C.franceBg },
    { flag: "🇮🇹", country: "意大利", hint: "___?", color: C.italy, bg: C.italyBg },
    { flag: "🇬🇧", country: "英国", hint: "___?", color: C.uk, bg: C.ukBg },
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


// SLIDE 33 — 🌍 欧洲三国文化对比 Comparison Table
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🌍 欧洲三国文化对比  Comparison", C.dark);

  const cols = [1.8, 2.4, 2.4, 2.4];
  const startX = 0.4;
  const startY = 0.95;
  const rowH = 0.63;

  const headers = ["", "🇫🇷 法国", "🇮🇹 意大利", "🇬🇧 英国"];
  const headerColors = [C.dark, C.france, C.italy, C.uk];

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
    ["打招呼", "亲脸颊 la bise", "Ciao + 手势", "握手 + Hello"],
    ["代表食物", "可颂/法棍", "披萨/意面", "炸鱼薯条"],
    ["著名地标", "埃菲尔铁塔", "斗兽场", "大本钟"],
    ["重要文化", "时尚/美食", "文艺复兴", "下午茶"],
    ["不能做的事", "催服务员", "菠萝配披萨", "插队"],
    ["特别的传统", "午餐1-2小时", "早上后不喝cappuccino", "聊天气"],
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

  card(s, 0.3, 1.0, 4.5, 4.0, C.bgGreen, C.green);
  s.addText("✅ 共同点 Similarities", {
    x: 0.5, y: 1.1, w: 4.1, h: 0.45,
    fontSize: 16, fontFace: "Georgia", color: C.green, bold: true, align: "center",
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.55, w: 3.5, h: 0.02, fill: { color: C.green },
  });
  s.addText([
    { text: "都在欧洲，地理位置很近", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "都有悠久的历史和丰富的文化遗产", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "都非常重视美食文化", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "都有世界著名的博物馆", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "都通过丝绸之路和中国有联系", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 0.6, y: 1.7, w: 3.9, h: 2.8, valign: "top" });

  card(s, 5.1, 1.0, 4.5, 4.0, C.bgOrange, C.accent);
  s.addText("❌ 不同点 Differences", {
    x: 5.3, y: 1.1, w: 4.1, h: 0.45,
    fontSize: 16, fontFace: "Georgia", color: C.accent, bold: true, align: "center",
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.6, y: 1.55, w: 3.5, h: 0.02, fill: { color: C.accent },
  });
  s.addText([
    { text: "打招呼方式不同（亲脸/Ciao/握手）", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "开车方向不同（英国靠左）", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "政治制度不同（英国有国王）", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "食物风格不同", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "    法国精致/意大利热情/英国传统", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 5.4, y: 1.7, w: 3.9, h: 2.8, valign: "top" });
})();


// SLIDE 35 — 🧳 旅行小贴士 Travel Tips Summary
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🧳 旅行小贴士  Travel Tips Summary", C.dark);

  s.addText("「如果你去欧洲旅行，要记住：」", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.45,
    fontSize: 16, fontFace: "Georgia", color: C.dark, bold: true, align: "center",
  });

  const tips = [
    { flag: "🇫🇷", country: "在法国", tips: "说Bonjour！\n不要催服务员！\n慢慢享受午餐！", color: C.france, bg: C.franceBg },
    { flag: "🇮🇹", country: "在意大利", tips: "说Ciao！\n不加菠萝在披萨上！\n进教堂穿长袖！", color: C.italy, bg: C.italyBg },
    { flag: "🇬🇧", country: "在英国", tips: "记得排队！\n多说Please!\n过马路先看右！", color: C.uk, bg: C.ukBg },
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
    { zh: "铁塔", py: "tie ta", en: "tower" },
    { zh: "博物馆", py: "bo wu guan", en: "museum" },
    { zh: "国王", py: "guo wang", en: "king" },
    { zh: "披萨", py: "pi sa", en: "pizza" },
    { zh: "奶酪", py: "nai lao", en: "cheese" },
    { zh: "排队", py: "pai dui", en: "queue/line up" },
    { zh: "丝绸", py: "si chou", en: "silk" },
    { zh: "教堂", py: "jiao tang", en: "church" },
    { zh: "下午茶", py: "xia wu cha", en: "afternoon tea" },
    { zh: "礼貌", py: "li mao", en: "polite" },
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
    { pattern: "「我喜欢吃___」", eng: "I like to eat ___", example: "我喜欢吃可颂/披萨/炸鱼薯条", color: C.france },
    { pattern: "「在___，人们用___打招呼」", eng: "In ___, people greet by ___", example: "在法国，人们用亲脸颊打招呼", color: C.italy },
    { pattern: "「去___旅行要注意___」", eng: "When traveling to ___, be careful about ___", example: "去英国旅行要注意排队和靠左开车", color: C.uk },
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


// SLIDE 38 — 🎭 Role Play (10-15min)
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🎭 Role Play  角色扮演 (10-15 min)", C.purple);

  s.addText("「你到了欧洲三个国家旅行！」", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.4,
    fontSize: 15, fontFace: "Georgia", color: C.purple, bold: true, align: "center",
  });

  const rounds = [
    { round: "Round 1", scene: "到法国", action: "亲脸颊说「Bonjour」\n慢慢吃可颂和法棍", color: C.france, bg: C.franceBg },
    { round: "Round 2", scene: "到意大利", action: "手势+「Ciao」\n吃披萨（不加菠萝！）", color: C.italy, bg: C.italyBg },
    { round: "Round 3", scene: "到英国", action: "握手说「Hello」\n排队买炸鱼薯条", color: C.uk, bg: C.ukBg },
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
    { level: "🔵 Level 2-3", desc: "「Bonjour! 我想吃___」+ 做动作", x: 3.5, color: "1565C0" },
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


// SLIDE 39 — 🎨 Project Time! Passport Template
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🎨 Project Time!  护照欧洲页", C.teal);

  s.addText("Passport 欧洲页 Template", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.teal, bold: true, align: "center",
  });

  const fields = [
    { icon: "🌍", label: "Where am I?", value: "欧洲 (法国/意大利/英国)" },
    { icon: "👀", label: "What did I see?", value: "埃菲尔铁塔 / 斗兽场 / 大本钟" },
    { icon: "🍽️", label: "What did I eat?", value: "可颂 / 披萨 / 炸鱼薯条" },
    { icon: "🎭", label: "Cultural Discovery", value: "亲脸颊 / Ciao手势 / 排队文化" },
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
  s.addText("画图 + 写词  (可颂 / 披萨 / 炸鱼薯条)", {
    x: 0.5, y: 1.45, w: 8.8, h: 0.55,
    fontSize: 13, fontFace: "Calibri", color: C.black,
  });

  card(s, 0.3, 2.4, 9.3, 1.2, C.bgBlue, "1565C0");
  s.addText("🔵 Level 2-3 Intermediate", {
    x: 0.5, y: 2.45, w: 4.0, h: 0.35,
    fontSize: 14, fontFace: "Georgia", color: "1565C0", bold: true,
  });
  s.addText("「I see the Eiffel Tower. I eat pizza. In the UK, people queue.」", {
    x: 0.5, y: 2.85, w: 8.8, h: 0.55,
    fontSize: 13, fontFace: "Calibri", color: C.black,
  });

  card(s, 0.3, 3.8, 9.3, 1.4, "F3E5F5", C.purple);
  s.addText("🟣 Level 4 Advanced", {
    x: 0.5, y: 3.85, w: 3.5, h: 0.35,
    fontSize: 14, fontFace: "Georgia", color: C.purple, bold: true,
  });
  s.addText("「In France, I saw the Louvre. I ate croissants and baguettes.\nWhen traveling there, don't rush the waiter because the French value slow dining.」", {
    x: 0.5, y: 4.25, w: 8.8, h: 0.75,
    fontSize: 12, fontFace: "Calibri", color: C.black,
  });
})();


// SLIDE 41 — 📄 Project 示范 Example
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "📄 Project 示范  Example Page", C.teal);

  card(s, 1.0, 1.0, 8.0, 4.2, C.white, C.teal);
  s.addText("🌍 My Europe Page  我的欧洲页", {
    x: 1.2, y: 1.1, w: 7.6, h: 0.5,
    fontSize: 18, fontFace: "Georgia", color: C.teal, bold: true, align: "center",
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 2.0, y: 1.6, w: 6.0, h: 0.02, fill: { color: C.teal },
  });

  const entries = [
    "🌍 I visited: 欧洲 Europe \u2014 法国, 意大利, 英国",
    "👀 I saw: 埃菲尔铁塔, 斗兽场, 大本钟",
    "🍽️ I ate: 可颂 croissants, 披萨 pizza, 炸鱼薯条 fish & chips!",
    "🎭 I learned: 在法国亲脸颊，在意大利说Ciao，在英国排队",
    "💬 My sentence: 我最喜欢吃披萨。在英国，人们排队很重要。",
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


// SLIDE 42 — 🗣️ 分享时间 Sharing Time
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🗣️ 分享时间  Sharing Time", C.dark);

  s.addText("和同伴分享你的欧洲页！\nShare your Europe page with a partner!", {
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


// SLIDE 43 — 🪪 欧洲签证章
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.dark };
  goldBars(s);
  slideNum(s);

  s.addText("🪪 欧洲签证章  Europe Visa Stamp", {
    x: 0.5, y: 0.4, w: 9.0, h: 0.6,
    fontSize: 26, fontFace: "Georgia", color: C.gold, bold: true, align: "center",
  });

  s.addShape(pptx.shapes.OVAL, {
    x: 3.0, y: 1.2, w: 4.0, h: 3.5,
    fill: { color: C.dark },
    line: { color: C.contEurope, width: 4, dashType: "dash" },
  });

  s.addText([
    { text: "EUROPE", options: { fontSize: 34, fontFace: "Georgia", color: C.contEurope, bold: true, breakLine: true } },
    { text: "欧洲", options: { fontSize: 24, fontFace: "Georgia", color: C.gold, breakLine: true } },
    { text: "\u2713 VISITED", options: { fontSize: 20, fontFace: "Georgia", color: C.quizGreen, bold: true, breakLine: true } },
    { text: "6/10/2025", options: { fontSize: 14, fontFace: "Calibri", color: C.white, breakLine: true } },
    { text: "法国 \u00B7 意大利 \u00B7 英国", options: { fontSize: 14, fontFace: "Calibri", color: C.lightAmber, breakLine: true } },
  ], { x: 3.0, y: 1.4, w: 4.0, h: 3.1, align: "center", valign: "middle" });

  s.addText("恭喜你完成欧洲之旅！Congratulations!", {
    x: 0.5, y: 4.85, w: 9.0, h: 0.4,
    fontSize: 14, fontFace: "Calibri", color: C.gold, bold: true, align: "center",
  });
})();


// SLIDE 44 — ✈️ 明天航班 Tomorrow's Flight
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
    { text: "Flight 航班: GR-004", options: { fontSize: 18, fontFace: "Georgia", color: C.contNA, bold: true, breakLine: true } },
    { text: "Destination 目的地: 美洲 AMERICAS", options: { fontSize: 18, fontFace: "Georgia", color: C.contNA, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "明天我们去美洲！", options: { fontSize: 16, fontFace: "Georgia", color: C.dark, bold: true, breakLine: true } },
    { text: "那里有什么有名的地方？", options: { fontSize: 16, fontFace: "Georgia", color: C.dark, breakLine: true } },
    { text: "Tomorrow we fly to the Americas! What famous places are there?", options: { fontSize: 12, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 1.8, y: 2.1, w: 6.4, h: 2.4, align: "center", valign: "middle" });

  s.addText("See you tomorrow, explorers!  明天见，小探险家们！", {
    x: 0.5, y: 4.85, w: 9.0, h: 0.4,
    fontSize: 14, fontFace: "Calibri", color: C.gold, bold: true, align: "center",
  });
})();


// ── Save ──
const outPath = path.join(__dirname, "day3_europe.pptx");
pptx.writeFile({ fileName: outPath }).then(() => {
  console.log("Created: " + outPath);
  console.log("Total slides: " + slideCount);
}).catch((err) => {
  console.error("Error:", err);
});
