/**
 * Day 2: 非洲 Africa (6/9) — Global Explorer Camp 环球探索沉浸式夏令营
 * ~44 slides — 3 countries (埃及 Egypt, 肯尼亚 Kenya, 南非 South Africa)
 * Each country: expanded with dedicated topic slides + large images
 * Run: node create_day2.js
 */
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const pptx = new pptxgen();
pptx.defineLayout({ name: "LAYOUT_16x9", width: 10.0, height: 5.625 });
pptx.layout = "LAYOUT_16x9";
pptx.author = "谷雨中文 GR EDU";
pptx.title = "Global Explorer Camp · Day 2: 非洲 Africa";

// ── Colors (NO # prefix) ──
const C = {
  primary:    "FF8F00",
  secondary:  "FFF3E0",
  accent:     "4E342E",
  dark:       "3E2723",
  white:      "FFFFFF",
  black:      "212121",
  gray:       "616161",
  gold:       "FFD54F",
  darkGold:   "FFA000",
  lightAmber: "FFE0B2",
  bgBlue:     "E8EAF6",
  bgGreen:    "E8F5E9",
  bgOrange:   "FFF3E0",
  bgPink:     "FCE4EC",
  bgBrown:    "EFEBE9",
  egypt:      "C62828",
  kenya:      "2E7D32",
  southAfrica:"1565C0",
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
  lightRed:   "FFCDD2",
  lightGreen: "C8E6C9",
  lightBlue:  "BBDEFB",
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
      fill: { color: C.bgOrange }, line: { color: C.gray, width: 1 },
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
    { text: "Flight 航班: GR-002", options: { fontSize: 16, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "Destination 目的地:  非洲 AFRICA", options: { fontSize: 16, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
    { text: "Date 日期: June 9, 2025  (6/9)", options: { fontSize: 14, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "Gate 登机口: Room 101", options: { fontSize: 14, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "Passenger 旅客: ________________", options: { fontSize: 14, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 2.5, y: 2.85, w: 5.0, h: 2.0, align: "left", valign: "top" });

  s.addText("Fasten seatbelts!  系好安全带！", {
    x: 0.5, y: 5.0, w: 9.0, h: 0.35,
    fontSize: 13, fontFace: "Calibri", color: C.gold, align: "center",
  });
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 2 — 📅 护照进度 Passport Progress
// ══════════════════════════════════════════════════════════════
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "📅 护照进度  Passport Progress", C.dark);

  const days = [
    { day: "Day 1", label: "亚洲\nAsia", color: C.contAsia, active: false, done: true },
    { day: "Day 2", label: "非洲\nAfrica", color: C.contAfrica, active: true, done: false },
    { day: "Day 3", label: "欧洲\nEurope", color: C.contEurope, active: false, done: false },
    { day: "Day 4", label: "美洲\nAmericas", color: C.contNA, active: false, done: false },
    { day: "Day 5", label: "展览\nExhibition", color: C.teal, active: false, done: false },
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
    if (d.done) {
      s.addText("\u2713", {
        x: x, y: 3.8, w: bw, h: 0.5,
        fontSize: 28, fontFace: "Georgia", color: C.quizGreen, bold: true, align: "center",
      });
    }
  });

  footerText(s, "1 stamp done!  已盖1个章！继续加油！", C.dark);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 3 — 🎯 今天目标 Today's Goals
// ══════════════════════════════════════════════════════════════
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🎯 今天目标  Today's Goals", C.primary);

  s.addText("深入了解非洲3个国家", {
    x: 0.5, y: 0.95, w: 9.0, h: 0.5,
    fontSize: 22, fontFace: "Georgia", color: C.dark, bold: true, align: "center",
  });

  const countries = [
    { flag: "🇪🇬", name: "埃及 Egypt", color: C.egypt, bg: C.lightRed },
    { flag: "🇰🇪", name: "肯尼亚 Kenya", color: C.kenya, bg: C.lightGreen },
    { flag: "🇿🇦", name: "南非 S. Africa", color: C.southAfrica, bg: C.lightBlue },
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
    { text: "2. 非洲独特的自然景观与野生动物", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "3. 文化特色与旅行礼节", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "4. 比较三个国家的异同", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
  ], { x: 0.8, y: 3.55, w: 8.4, h: 1.5, valign: "top" });
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 4 — 🌍 认识非洲 About Africa
// ══════════════════════════════════════════════════════════════
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🌍 认识非洲  About Africa", C.contAfrica);

  const stats = [
    { icon: "🌍", label: "54个国家", val: "54 Countries" },
    { icon: "👥", label: "14亿人口", val: "1.4 Billion People" },
    { icon: "🏜️", label: "撒哈拉沙漠", val: "Sahara \u2014 世界最大沙漠" },
    { icon: "🏞️", label: "尼罗河", val: "Nile \u2014 世界最长河流" },
    { icon: "🗣️", label: "2,000+种语言", val: "Most Linguistically Diverse!" },
    { icon: "🦁", label: "野生动物天堂", val: "Big Five + Great Migration" },
  ];

  stats.forEach((st, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.3 + col * 3.15;
    const y = 1.0 + row * 2.1;

    card(s, x, y, 2.9, 1.8, C.white, C.contAfrica);
    s.addText(st.icon, {
      x: x, y: y + 0.1, w: 2.9, h: 0.6,
      fontSize: 30, fontFace: "Calibri", align: "center",
    });
    s.addText(st.label, {
      x: x, y: y + 0.7, w: 2.9, h: 0.5,
      fontSize: 16, fontFace: "Georgia", color: C.contAfrica, bold: true, align: "center",
    });
    s.addText(st.val, {
      x: x, y: y + 1.2, w: 2.9, h: 0.45,
      fontSize: 12, fontFace: "Calibri", color: C.gray, align: "center",
    });
  });

  footerText(s, "非洲是人类文明的发源地！Africa is the cradle of humanity!", C.contAfrica);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 5 — 🗺️ 非洲地图 Africa Map
// ══════════════════════════════════════════════════════════════
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: "E8F0FE" };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🗺️ 非洲地图  Africa Map", C.contAfrica);

  // Simplified Africa map with country blobs
  const regions = [
    { label: "🇪🇬 埃及\nEgypt", x: 4.8, y: 1.0, w: 2.0, h: 1.2, color: C.egypt },
    { label: "🇰🇪 肯尼亚\nKenya", x: 5.5, y: 2.8, w: 1.8, h: 1.2, color: C.kenya },
    { label: "🇿🇦 南非\nS. Africa", x: 4.0, y: 4.0, w: 2.2, h: 1.0, color: C.southAfrica },
    { label: "北非\nN. Africa", x: 2.0, y: 1.0, w: 2.5, h: 1.2, color: C.contAntarc },
    { label: "西非\nW. Africa", x: 1.5, y: 2.3, w: 2.0, h: 1.5, color: "8D6E63" },
    { label: "中非\nC. Africa", x: 3.5, y: 2.3, w: 2.0, h: 1.5, color: "A1887F" },
    { label: "东非\nE. Africa", x: 5.5, y: 2.3, w: 1.0, h: 0.5, color: C.teal },
  ];

  regions.forEach((c) => {
    const isTarget = c.label.includes("埃及") || c.label.includes("肯尼亚") || c.label.includes("南非");
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

  footerText(s, "今天我们去这三个国家！Today we visit these 3 countries!", C.contAfrica);
})();


// ══════════════════════════════════════════════════════════════
// ══════════════ EGYPT SECTION (8 slides) ═════════════════════
// ══════════════════════════════════════════════════════════════

// SLIDE 6 — 🇪🇬 埃及概览 Egypt Overview
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🇪🇬 埃及概览  Egypt Overview", C.egypt);

  // Pyramids photo LEFT
  safeImage(s, "africa_pyramids.jpg", "Pyramids 金字塔", 0.3, 0.95, 5.0, 3.5);

  // Info card RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.egypt);

  const items = [
    { text: "🏴 国旗：红白黑三色 + 金鹰", y: 1.1 },
    { text: "👥 人口：约1.04亿", y: 1.6 },
    { text: "🗣️ 语言：阿拉伯语", y: 2.1 },
    { text: "🏛️ 首都：开罗 Cairo", y: 2.6 },
    { text: "📏 地跨非洲和亚洲", y: 3.1 },
    { text: "🏺 四大文明古国之一", y: 3.6 },
    { text: "🌍 尼罗河流经全境", y: 4.1 },
  ];

  items.forEach((it) => {
    s.addText(it.text, {
      x: 5.85, y: it.y, w: 3.5, h: 0.42,
      fontSize: 13, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "埃及有5000年的文明史！Egypt has 5,000 years of civilization!", C.egypt);
})();


// SLIDE 7 — 🏛️ 首都：开罗 Capital: Cairo
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🏛️ 首都：开罗  Capital: Cairo", C.egypt);

  // Cairo photo LEFT
  safeImage(s, "africa_cairo.jpg", "Cairo 开罗", 0.3, 0.95, 5.0, 3.5);

  // Info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.egypt);
  s.addText([
    { text: "开罗是埃及的首都", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "超过2000万人口", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "非洲最大的城市之一", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "位于尼罗河畔", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "阿拉伯文化中心", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "金字塔就在城市边上！", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 2.8, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 4.0, w: 3.5, h: 0.02, fill: { color: C.egypt },
  });

  s.addText("「开罗」在阿拉伯语里意思是「胜利之城」！", {
    x: 5.85, y: 4.1, w: 3.5, h: 0.7,
    fontSize: 12, fontFace: "Calibri", color: C.egypt, bold: true, align: "left",
  });
})();


// SLIDE 8 — 🏺 金字塔与狮身人面像 Pyramids & Sphinx
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🏺 金字塔与狮身人面像  Pyramids & Sphinx", C.egypt);

  // Sphinx photo LEFT (large)
  safeImage(s, "africa_sphinx.jpg", "Sphinx 狮身人面像", 0.3, 0.95, 5.0, 4.0);

  // Info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.egypt);
  s.addText([
    { text: "有4500年历史！", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "古代世界七大奇迹之一", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "胡夫金字塔高147米", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "用230万块石头建成！", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "每块石头重2.5吨", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "狮身人面像长73米", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "至今仍是未解之谜", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 3.2, valign: "top" });

  footerText(s, "金字塔是人类最伟大的建筑之一！One of humanity's greatest achievements!", C.egypt);
})();


// SLIDE 9 — 🏞️ 尼罗河 The Nile
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🏞️ 尼罗河  The Nile", C.egypt);

  // Nile photo LEFT
  safeImage(s, "africa_nile.jpg", "Nile River 尼罗河", 0.3, 0.95, 5.0, 3.5);

  // Info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.egypt);
  s.addText([
    { text: "世界最长河流：6,650公里！", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "古埃及文明的摇篮", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "每年洪水带来肥沃土壤", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "现在有阿斯旺大坝控制洪水", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "流经11个非洲国家", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "从南向北流入地中海", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 2.8, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 4.0, w: 3.5, h: 0.02, fill: { color: C.egypt },
  });

  s.addText("没有尼罗河就没有埃及文明！No Nile, no Egypt!", {
    x: 5.85, y: 4.1, w: 3.5, h: 0.7,
    fontSize: 12, fontFace: "Calibri", color: C.egypt, bold: true, align: "left",
  });
})();


// SLIDE 10 — 🍽️ 埃及美食 Egyptian Food
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🍽️ 埃及美食  Egyptian Food", C.egypt);

  // Food info — full width card
  card(s, 0.3, 0.95, 9.3, 4.3, C.white, C.egypt);
  s.addText([
    { text: "🥣 鹰嘴豆泥 Hummus", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    中东最受欢迎的酱料，用大饼蘸着吃", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🍲 库莎丽 Koshari", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    埃及国菜！米饭+意面+扁豆+番茄酱", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🧆 法拉费尔 Falafel", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    炸鹰嘴豆丸子，外酥里嫩", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🫓 大饼 Pita", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    每餐必备的主食", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🍵 甜茶 Sweet Tea", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    埃及人特别爱喝加糖的红茶", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 0.6, y: 1.1, w: 8.7, h: 3.5, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.6, y: 4.55, w: 8.7, h: 0.02, fill: { color: C.egypt },
  });
  s.addText("「你吃过鹰嘴豆泥吗？你觉得好吃吗？」", {
    x: 0.6, y: 4.6, w: 8.7, h: 0.5,
    fontSize: 13, fontFace: "Georgia", color: C.egypt, bold: true, align: "center",
  });
})();


// SLIDE 11 — ⚠️ 埃及旅行礼节 Egypt Travel Etiquette
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "⚠️ 埃及旅行礼节  Egypt Travel Etiquette", C.egypt);

  card(s, 0.3, 1.0, 9.3, 4.2, C.lightRed, C.egypt);
  s.addText("🇪🇬 在埃及旅行要注意：", {
    x: 0.55, y: 1.1, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.egypt, bold: true,
  });

  const tips = [
    { icon: "🤚", text: "用右手吃饭和递东西（左手被认为不干净）" },
    { icon: "🕌", text: "进清真寺要脱鞋" },
    { icon: "🌙", text: "斋月期间白天不要在公开场合吃东西" },
    { icon: "💋", text: "亲脸颊是常见的打招呼方式" },
    { icon: "👗", text: "穿着要保守，尤其是女性" },
    { icon: "📸", text: "拍照前要先问许可" },
    { icon: "🙏", text: "「As-salamu alaykum」= 你好（愿和平与你同在）" },
  ];

  tips.forEach((t, i) => {
    const y = 1.65 + i * 0.48;
    s.addText(t.icon + "  " + t.text, {
      x: 0.8, y: y, w: 8.5, h: 0.42,
      fontSize: 14, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "尊重伊斯兰文化是在埃及旅行的基本礼节！", C.egypt);
})();


// SLIDE 12-13 — ✅ 埃及 Check Understanding (1/2, 2/2)
(() => {
  quizSlides("✅ 埃及 Check Understanding", [
    { q: "埃及的首都是哪里？", a: "开罗 Cairo" },
    { q: "世界最长的河流是什么？", a: "尼罗河 Nile (6,650km)" },
    { q: "金字塔有多少年历史？", a: "约4,500年" },
    { q: "胡夫金字塔用了多少块石头？", a: "230万块" },
    { q: "埃及人说什么语言？", a: "阿拉伯语" },
    { q: "埃及的国菜叫什么？", a: "库莎丽 Koshari" },
    { q: "进清真寺要注意什么？", a: "要脱鞋" },
    { q: "用哪只手吃饭和递东西？", a: "右手" },
    { q: "狮身人面像有多长？", a: "73米" },
    { q: "尼罗河流经几个国家？", a: "11个" },
  ], C.egypt);
  countSlide(); countSlide();
})();


// ══════════════════════════════════════════════════════════════
// ══════════════ KENYA SECTION (7 slides) ═════════════════════
// ══════════════════════════════════════════════════════════════

// SLIDE 14 — 🇰🇪 肯尼亚概览 Kenya Overview
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🇰🇪 肯尼亚概览  Kenya Overview", C.kenya);

  // Kenya photo LEFT
  safeImage(s, "africa_kenya.jpg", "Kenya 肯尼亚", 0.3, 0.95, 5.0, 3.5);

  // Info card RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.kenya);

  const items = [
    { text: "🏴 国旗：黑红绿+马赛盾牌", y: 1.1 },
    { text: "👥 人口：约5,500万", y: 1.6 },
    { text: "🗣️ 斯瓦希里语 + 英语", y: 2.1 },
    { text: "🏛️ 首都：内罗毕 Nairobi", y: 2.6 },
    { text: "🌍 赤道穿过肯尼亚！", y: 3.1 },
    { text: "🦁 Safari（野生动物之旅）圣地", y: 3.6 },
    { text: "🏃 世界长跑冠军的摇篮", y: 4.1 },
  ];

  items.forEach((it) => {
    s.addText(it.text, {
      x: 5.85, y: it.y, w: 3.5, h: 0.42,
      fontSize: 13, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "肯尼亚被称为「Safari之都」！Kenya is the Safari Capital!", C.kenya);
})();


// SLIDE 15 — 🦁 非洲五大动物 Big Five
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🦁 非洲五大动物  The Big Five", C.kenya);

  // Lion photo LEFT (large)
  safeImage(s, "africa_lion.jpg", "Lion 狮子", 0.3, 0.95, 5.0, 4.0);

  // Big Five info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.kenya);

  const bigFive = [
    { name: "🦁 狮子 Lion", detail: "百兽之王，群居动物", y: 1.1 },
    { name: "🐘 大象 Elephant", detail: "陆地最大的动物", y: 1.85 },
    { name: "🦏 犀牛 Rhinoceros", detail: "濒危动物，需要保护", y: 2.6 },
    { name: "🐆 花豹 Leopard", detail: "非洲最神秘的猫科动物", y: 3.35 },
    { name: "🐃 水牛 Buffalo", detail: "非洲最危险的动物之一", y: 4.1 },
  ];

  bigFive.forEach((bf) => {
    s.addText(bf.name, {
      x: 5.85, y: bf.y, w: 3.5, h: 0.35,
      fontSize: 13, fontFace: "Georgia", color: C.kenya, bold: true, align: "left",
    });
    s.addText(bf.detail, {
      x: 5.85, y: bf.y + 0.35, w: 3.5, h: 0.3,
      fontSize: 11, fontFace: "Calibri", color: C.gray, align: "left",
    });
  });

  footerText(s, "在Safari中看到Big Five是最大的愿望！", C.kenya);
})();


// SLIDE 16 — 🐘 动物大迁徙 Great Migration
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🐘 动物大迁徙  The Great Migration", C.kenya);

  // Elephant/safari photo LEFT
  safeImage(s, "africa_elephant.jpg", "Elephants 大象", 0.3, 0.95, 5.0, 3.5);

  // Info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.kenya);
  s.addText([
    { text: "世界上最壮观的自然景象！", options: { fontSize: 14, fontFace: "Georgia", color: C.kenya, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "200万只动物参与迁徙", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "包括角马、斑马、瞪羚", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "从坦桑尼亚到肯尼亚", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "马赛马拉是最佳观赏地", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "每年7-10月最壮观", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "跨越鳄鱼出没的马拉河！", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 3.5, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 4.5, w: 3.5, h: 0.02, fill: { color: C.kenya },
  });
  s.addText("这是地球上最大的动物迁徙！", {
    x: 5.85, y: 4.55, w: 3.5, h: 0.4,
    fontSize: 12, fontFace: "Calibri", color: C.kenya, bold: true, align: "left",
  });
})();


// SLIDE 17 — 🍽️ 肯尼亚美食与文化 Kenya Food & Culture
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🍽️ 肯尼亚美食与文化  Kenya Food & Culture", C.kenya);

  // Food + culture info full width
  card(s, 0.3, 0.95, 5.8, 4.3, C.lightGreen, C.kenya);
  s.addText([
    { text: "🍚 乌伽黎 Ugali", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    玉米面做的主食，像年糕一样", options: { fontSize: 12, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🥩 烤肉 Nyama Choma", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    肯尼亚最受欢迎的菜！", options: { fontSize: 12, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🫓 恰帕提 Chapati", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    印度传来的薄饼，非常好吃", options: { fontSize: 12, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🍵 肯尼亚红茶 Kenyan Tea", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    世界第三大茶叶出口国！", options: { fontSize: 12, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 0.55, y: 1.1, w: 5.3, h: 2.8, valign: "top" });

  // Culture card RIGHT
  card(s, 6.35, 0.95, 3.3, 4.3, C.white, C.kenya);
  s.addText("🎭 马赛族文化", {
    x: 6.55, y: 1.05, w: 2.9, h: 0.4,
    fontSize: 14, fontFace: "Georgia", color: C.kenya, bold: true, align: "center",
  });

  const cultureItems = [
    "马赛族是肯尼亚\n最著名的部族",
    "穿红色斗篷\n佩戴彩色珠子",
    "跳跃舞：跳得越高\n越受尊敬！",
    "半游牧生活\n与野生动物共存",
  ];

  cultureItems.forEach((txt, i) => {
    s.addText(txt, {
      x: 6.55, y: 1.55 + i * 0.85, w: 2.9, h: 0.75,
      fontSize: 11, fontFace: "Calibri", color: C.black, align: "center", valign: "middle",
    });
    if (i < cultureItems.length - 1) {
      s.addShape(pptx.shapes.RECTANGLE, {
        x: 6.85, y: 2.3 + i * 0.85, w: 2.3, h: 0.01, fill: { color: C.kenya },
      });
    }
  });
})();


// SLIDE 18 — ⚠️ 肯尼亚旅行礼节 Kenya Travel Etiquette
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "⚠️ 肯尼亚旅行礼节  Kenya Travel Etiquette", C.kenya);

  card(s, 0.3, 1.0, 9.3, 4.2, C.lightGreen, C.kenya);
  s.addText("🇰🇪 在肯尼亚旅行要注意：", {
    x: 0.55, y: 1.1, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.kenya, bold: true,
  });

  const tips = [
    { icon: "👋", text: "Jambo! (你好) \u2014 斯瓦希里语的问候" },
    { icon: "🤝", text: "长握手表示尊重，不要急着松手" },
    { icon: "☝️", text: "不要用手指指着别人，用整只手" },
    { icon: "🦁", text: "Safari时保持距离，不要下车！" },
    { icon: "📸", text: "拍当地人之前要先问" },
    { icon: "🙏", text: "Hakuna Matata! = 没有烦恼！" },
    { icon: "👴", text: "尊重长者，让长者先吃饭" },
  ];

  tips.forEach((t, i) => {
    const y = 1.65 + i * 0.48;
    s.addText(t.icon + "  " + t.text, {
      x: 0.8, y: y, w: 8.5, h: 0.42,
      fontSize: 14, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "Hakuna Matata! 在肯尼亚，每天都是美好的一天！", C.kenya);
})();


// SLIDE 19-20 — ✅ 肯尼亚 Check Understanding (1/2, 2/2)
(() => {
  quizSlides("✅ 肯尼亚 Check Understanding", [
    { q: "肯尼亚的首都是哪里？", a: "内罗毕 Nairobi" },
    { q: "非洲五大动物是哪五种？", a: "狮子、大象、犀牛、花豹、水牛" },
    { q: "动物大迁徙有多少只动物？", a: "200万只" },
    { q: "Jambo是什么意思？", a: "你好！(斯瓦希里语)" },
    { q: "肯尼亚人说哪两种语言？", a: "斯瓦希里语和英语" },
    { q: "Ugali是用什么做的？", a: "玉米面" },
    { q: "马赛族的跳跃舞有什么意义？", a: "跳得越高越受尊敬" },
    { q: "Safari时最重要的规矩是什么？", a: "保持距离，不要下车" },
    { q: "Hakuna Matata是什么意思？", a: "没有烦恼！" },
    { q: "什么线穿过肯尼亚？", a: "赤道" },
  ], C.kenya);
  countSlide(); countSlide();
})();


// ══════════════════════════════════════════════════════════════
// ══════════════ SOUTH AFRICA SECTION (7 slides) ══════════════
// ══════════════════════════════════════════════════════════════

// SLIDE 21 — 🇿🇦 南非概览 South Africa Overview
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🇿🇦 南非概览  South Africa Overview", C.southAfrica);

  // Table Mountain photo LEFT
  safeImage(s, "africa_table_mountain.jpg", "Table Mountain 桌山", 0.3, 0.95, 5.0, 3.5);

  // Info card RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.southAfrica);

  const items = [
    { text: "🏴 国旗：6种颜色！世界最多", y: 1.1 },
    { text: "👥 人口：约6,000万", y: 1.55 },
    { text: "🗣️ 11种官方语言！", y: 2.0 },
    { text: "🏛️ 3个首都！", y: 2.45 },
    { text: "    (行政/立法/司法各一个)", y: 2.75 },
    { text: "🌈 彩虹之国 Rainbow Nation", y: 3.15 },
    { text: "🏔️ 桌山：世界新七大自然奇观", y: 3.6 },
    { text: "⛵ 好望角在这里！", y: 4.05 },
  ];

  items.forEach((it) => {
    s.addText(it.text, {
      x: 5.85, y: it.y, w: 3.5, h: 0.38,
      fontSize: 12, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "南非被称为「彩虹之国」因为文化非常多元！", C.southAfrica);
})();


// SLIDE 22 — 🏔️ 桌山与自然 Table Mountain & Nature
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🏔️ 桌山与自然  Table Mountain & Nature", C.southAfrica);

  // Kilimanjaro/nature photo LEFT
  safeImage(s, "africa_kilimanjaro.jpg", "Nature 自然风光", 0.3, 0.95, 5.0, 3.5);

  // Info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.southAfrica);
  s.addText([
    { text: "🏔️ 桌山 Table Mountain", options: { fontSize: 14, fontFace: "Georgia", color: C.southAfrica, bold: true, breakLine: true } },
    { text: "山顶像桌子一样平！", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "世界新七大自然奇观之一", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "可以坐缆车上去", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "⛵ 好望角 Cape of Good Hope", options: { fontSize: 14, fontFace: "Georgia", color: C.southAfrica, bold: true, breakLine: true } },
    { text: "非洲大陆的最南端附近", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "两大洋在这里相遇", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "🦁 克鲁格国家公园", options: { fontSize: 14, fontFace: "Georgia", color: C.southAfrica, bold: true, breakLine: true } },
    { text: "南非最大的野生动物保护区", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 3.8, valign: "top" });

  footerText(s, "南非的自然风光令人叹为观止！Breathtaking nature!", C.southAfrica);
})();


// SLIDE 23 — ✊ 曼德拉与彩虹之国 Mandela & Rainbow Nation
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "✊ 曼德拉与彩虹之国  Mandela & Rainbow Nation", C.southAfrica);

  // Mandela photo LEFT
  safeImage(s, "africa_mandela.jpg", "Nelson Mandela 曼德拉", 0.3, 0.95, 5.0, 3.5);

  // Info RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.white, C.southAfrica);
  s.addText([
    { text: "纳尔逊·曼德拉的故事", options: { fontSize: 14, fontFace: "Georgia", color: C.southAfrica, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "为了种族平等奋斗一生", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "被关在监狱27年！", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "出狱后当选南非总统", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "主张和解，不是报复", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "获得诺贝尔和平奖", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 2.8, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.85, y: 3.9, w: 3.5, h: 0.02, fill: { color: C.southAfrica },
  });

  s.addText([
    { text: "为什么叫「彩虹之国」？", options: { fontSize: 13, fontFace: "Georgia", color: C.southAfrica, bold: true, breakLine: true } },
    { text: "11种语言 = 多元文化共存", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "不同肤色的人和平生活在一起", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
  ], { x: 5.85, y: 3.95, w: 3.5, h: 1.2, valign: "top" });

  footerText(s, "曼德拉说：「教育是改变世界最有力的武器。」", C.southAfrica);
})();


// SLIDE 24 — 🍽️ 南非美食 South African Food
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🍽️ 南非美食  South African Food", C.southAfrica);

  card(s, 0.3, 0.95, 9.3, 4.3, C.white, C.southAfrica);
  s.addText([
    { text: "🔥 烤肉 Braai", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    南非人的骄傲！比BBQ更神圣，是社交活动", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🥩 干肉条 Biltong", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    风干的调味肉条，南非最受欢迎的零食", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🍚 玉米粥 Pap", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    和肯尼亚的Ugali类似，南非的主食", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🍛 咖喱肉派 Bobotie", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    南非国菜！咖喱肉+蛋奶冻烤出来", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "🍵 如意宝茶 Rooibos Tea", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    { text: "    只在南非才有的红茶，不含咖啡因！", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 0.6, y: 1.1, w: 8.7, h: 3.5, valign: "top" });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.6, y: 4.55, w: 8.7, h: 0.02, fill: { color: C.southAfrica },
  });
  s.addText("「在南非，Braai不只是烤肉，是一种生活方式！」", {
    x: 0.6, y: 4.6, w: 8.7, h: 0.5,
    fontSize: 13, fontFace: "Georgia", color: C.southAfrica, bold: true, align: "center",
  });
})();


// SLIDE 25 — ⚠️ 南非旅行礼节 South Africa Travel Etiquette
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "⚠️ 南非旅行礼节  S. Africa Travel Etiquette", C.southAfrica);

  card(s, 0.3, 1.0, 9.3, 4.2, C.lightBlue, C.southAfrica);
  s.addText("🇿🇦 在南非旅行要注意：", {
    x: 0.55, y: 1.1, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.southAfrica, bold: true,
  });

  const tips = [
    { icon: "👋", text: "Sawubona! (我看见你了) \u2014 祖鲁语的问候" },
    { icon: "🤝", text: "三步握手：握手→扣拇指→再握手，表示友谊" },
    { icon: "🌈", text: "尊重多元文化，南非有11种语言" },
    { icon: "🔥", text: "被邀请去Braai是很大的荣幸！" },
    { icon: "✊", text: "Ubuntu精神：「我因为你而存在」" },
    { icon: "😊", text: "南非人非常热情友好" },
    { icon: "🏞️", text: "尊重自然，不要随意接近野生动物" },
  ];

  tips.forEach((t, i) => {
    const y = 1.65 + i * 0.48;
    s.addText(t.icon + "  " + t.text, {
      x: 0.8, y: y, w: 8.5, h: 0.42,
      fontSize: 14, fontFace: "Calibri", color: C.black, align: "left", valign: "middle",
    });
  });

  footerText(s, "Ubuntu: I am because we are! 我因为你而存在！", C.southAfrica);
})();


// SLIDE 26-27 — ✅ 南非 Check Understanding (1/2, 2/2)
(() => {
  quizSlides("✅ 南非 Check Understanding", [
    { q: "南非有几种官方语言？", a: "11种！" },
    { q: "南非为什么叫「彩虹之国」？", a: "多元文化共存" },
    { q: "曼德拉被关了多少年？", a: "27年" },
    { q: "南非有几个首都？", a: "3个！" },
    { q: "南非国旗有几种颜色？", a: "6种" },
    { q: "Braai是什么？", a: "南非烤肉/社交活动" },
    { q: "桌山的特点是什么？", a: "山顶像桌子一样平" },
    { q: "Sawubona是什么意思？", a: "我看见你了（祖鲁语问候）" },
    { q: "Ubuntu是什么意思？", a: "我因为你而存在" },
    { q: "只在南非才有的茶叫什么？", a: "如意宝茶 Rooibos" },
  ], C.southAfrica);
  countSlide(); countSlide();
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 28 — ☕ 埃塞俄比亚咖啡的故事 Ethiopian Coffee Story
// ══════════════════════════════════════════════════════════════
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "☕ 埃塞俄比亚咖啡的故事  Ethiopian Coffee", C.accent);

  // Coffee photo LEFT
  safeImage(s, "africa_coffee.jpg", "Coffee 咖啡", 0.3, 0.95, 5.0, 3.5);

  // Story RIGHT
  card(s, 5.6, 0.95, 4.0, 4.3, C.bgBrown, C.accent);
  s.addText([
    { text: "☕ 咖啡的发源地！", options: { fontSize: 14, fontFace: "Georgia", color: C.accent, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "传说一个牧羊人发现", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "山羊吃了红色果子后", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "变得精力充沛、又蹦又跳！", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "☕ 咖啡仪式 Coffee Ceremony", options: { fontSize: 13, fontFace: "Georgia", color: C.accent, bold: true, breakLine: true } },
    { text: "生豆现烤现磨现煮", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "要喝三杯才算礼貌", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "是社交和待客的重要仪式", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.85, y: 1.1, w: 3.5, h: 3.8, valign: "top" });

  footerText(s, "全世界的咖啡都来自非洲！Coffee comes from Africa!", C.accent);
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
    { flag: "🇪🇬", country: "埃及 Egypt", action: "亲脸颊 +\n「As-salamu alaykum」", bg: C.lightRed, color: C.egypt },
    { flag: "🇰🇪", country: "肯尼亚 Kenya", action: "长握手 +\n「Jambo!」", bg: C.lightGreen, color: C.kenya },
    { flag: "🇿🇦", country: "南非 S. Africa", action: "三步握手 +\n「Sawubona!」", bg: C.lightBlue, color: C.southAfrica },
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
    { text: "Round 1: 说出3个非洲国家的首都", options: { fontSize: 14, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "Round 2: 每个国家的打招呼方式", options: { fontSize: 14, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "Round 3: 说出非洲五大动物", options: { fontSize: 14, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "Round 4: 用中文介绍一个国家的食物", options: { fontSize: 14, fontFace: "Calibri", color: C.black, breakLine: true } },
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
  s.background = { fill: C.primary };
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
    { flag: "🇪🇬", country: "埃及", hint: "___?", color: C.egypt, bg: C.lightRed },
    { flag: "🇰🇪", country: "肯尼亚", hint: "___?", color: C.kenya, bg: C.lightGreen },
    { flag: "🇿🇦", country: "南非", hint: "___?", color: C.southAfrica, bg: C.lightBlue },
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


// SLIDE 33 — 🌍 非洲三国对比表 Comparison Table
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  slideNum(s);
  headerBar(s, "🌍 非洲三国文化对比  Comparison", C.dark);

  const cols = [1.8, 2.4, 2.4, 2.4];
  const startX = 0.4;
  const startY = 0.95;
  const rowH = 0.63;

  const headers = ["", "🇪🇬 埃及", "🇰🇪 肯尼亚", "🇿🇦 南非"];
  const headerColors = [C.dark, C.egypt, C.kenya, C.southAfrica];

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
    ["打招呼", "亲脸颊", "Jambo!", "Sawubona!"],
    ["语言", "阿拉伯语", "斯瓦希里语+英语", "11种语言"],
    ["代表食物", "库莎丽 Koshari", "烤肉 Nyama Choma", "烤肉 Braai"],
    ["特色景点", "金字塔", "马赛马拉", "桌山"],
    ["有趣的事", "狮身人面像", "动物大迁徙", "彩虹之国"],
    ["特别的动物", "骆驼 🐪", "狮子 🦁", "企鹅 🐧"],
  ];

  rows.forEach((row, ri) => {
    const ry = startY + rowH * (ri + 1);
    const bgColor = ri % 2 === 0 ? C.white : C.bgOrange;
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
    { text: "都在非洲大陆", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "都有丰富的野生动物", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "都重视待客之道", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "都有独特的美食文化", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "都有悠久的历史传统", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 0.6, y: 1.7, w: 3.9, h: 2.8, valign: "top" });

  // Differences RIGHT
  card(s, 5.1, 1.0, 4.5, 4.0, C.bgOrange, C.primary);
  s.addText("❌ 不同点 Differences", {
    x: 5.3, y: 1.1, w: 4.1, h: 0.45,
    fontSize: 16, fontFace: "Georgia", color: C.primary, bold: true, align: "center",
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.6, y: 1.55, w: 3.5, h: 0.02, fill: { color: C.primary },
  });
  s.addText([
    { text: "地理位置不同（北/东/南）", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "语言完全不同", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "宗教信仰不同", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "气候差异很大", options: { fontSize: 13, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "    埃及沙漠/肯尼亚草原/南非温带", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 5.4, y: 1.7, w: 3.9, h: 2.8, valign: "top" });
})();


// SLIDE 35 — 🧳 旅行小贴士 Travel Tips Summary
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🧳 旅行小贴士  Travel Tips Summary", C.dark);

  s.addText("「如果你去非洲旅行，要记住：」", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.45,
    fontSize: 16, fontFace: "Georgia", color: C.dark, bold: true, align: "center",
  });

  const tips = [
    { flag: "🇪🇬", country: "在埃及", tips: "用右手吃饭！\n进清真寺脱鞋！\n斋月不公开吃东西！", color: C.egypt, bg: C.lightRed },
    { flag: "🇰🇪", country: "在肯尼亚", tips: "说Jambo!\nSafari不要下车！\n长握手表尊重！", color: C.kenya, bg: C.lightGreen },
    { flag: "🇿🇦", country: "在南非", tips: "三步握手！\n尊重多元文化！\nUbuntu精神！", color: C.southAfrica, bg: C.lightBlue },
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
    { zh: "金字塔", py: "jin zi ta", en: "pyramid" },
    { zh: "沙漠", py: "sha mo", en: "desert" },
    { zh: "狮子", py: "shi zi", en: "lion" },
    { zh: "大象", py: "da xiang", en: "elephant" },
    { zh: "尼罗河", py: "ni luo he", en: "Nile River" },
    { zh: "迁徙", py: "qian xi", en: "migration" },
    { zh: "咖啡", py: "ka fei", en: "coffee" },
    { zh: "彩虹", py: "cai hong", en: "rainbow" },
    { zh: "和平", py: "he ping", en: "peace" },
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
    const bg = r % 2 === 0 ? C.white : C.bgOrange;
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
    { pattern: "「我想去___看___」", eng: "I want to go to ___ to see ___", example: "我想去埃及看金字塔", color: C.egypt },
    { pattern: "「在___，人们说___打招呼」", eng: "In ___, people say ___ to greet", example: "在肯尼亚，人们说Jambo打招呼", color: C.kenya },
    { pattern: "「___最有名的是___」", eng: "___ is most famous for ___", example: "南非最有名的是彩虹之国", color: C.southAfrica },
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


// SLIDE 38 — 🎭 Role Play (Safari + 餐厅)
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.secondary };
  goldBars(s);
  slideNum(s);
  headerBar(s, "🎭 Role Play  角色扮演 (10-15 min)", C.purple);

  s.addText("「你到了非洲三个国家旅行！」", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.4,
    fontSize: 15, fontFace: "Georgia", color: C.purple, bold: true, align: "center",
  });

  const rounds = [
    { round: "Round 1", scene: "Safari导游", action: "介绍Big Five\n「看！那是___！」", color: C.kenya, bg: C.lightGreen },
    { round: "Round 2", scene: "埃及餐厅", action: "点菜：「我想吃___」\n用右手吃大饼", color: C.egypt, bg: C.lightRed },
    { round: "Round 3", scene: "南非Braai", action: "三步握手 + 说Sawubona\n「这是___，很好吃！」", color: C.southAfrica, bg: C.lightBlue },
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
    { level: "🟢 零基础", desc: "模仿动作 + 说动物名", x: 0.5, color: C.green },
    { level: "🔵 Level 2-3", desc: "「你好！我想看___」+ 做动作", x: 3.5, color: "1565C0" },
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
  headerBar(s, "🎨 Project Time!  护照非洲页", C.teal);

  s.addText("Passport 非洲页 Template", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.teal, bold: true, align: "center",
  });

  const fields = [
    { icon: "🌎", label: "Where am I?", value: "非洲 (埃及/肯尼亚/南非)" },
    { icon: "👀", label: "What did I see?", value: "金字塔 / Safari / 桌山" },
    { icon: "🍜", label: "What did I eat?", value: "库莎丽 / 烤肉 / Braai" },
    { icon: "🎭", label: "Cultural Discovery", value: "三步握手 / Jambo / Ubuntu" },
    { icon: "💬", label: "My Sentence", value: "我想去___看___。___最有名的是..." },
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
  s.addText("画图 + 写词  (金字塔 / 狮子 / 大象)", {
    x: 0.5, y: 1.45, w: 8.8, h: 0.55,
    fontSize: 13, fontFace: "Calibri", color: C.black,
  });

  card(s, 0.3, 2.4, 9.3, 1.2, C.bgBlue, "1565C0");
  s.addText("🔵 Level 2-3 Intermediate", {
    x: 0.5, y: 2.45, w: 4.0, h: 0.35,
    fontSize: 14, fontFace: "Georgia", color: "1565C0", bold: true,
  });
  s.addText("「I see ___. I eat ___. In Kenya, people say Jambo.」", {
    x: 0.5, y: 2.85, w: 8.8, h: 0.55,
    fontSize: 13, fontFace: "Calibri", color: C.black,
  });

  card(s, 0.3, 3.8, 9.3, 1.4, "F3E5F5", C.purple);
  s.addText("🟣 Level 4 Advanced", {
    x: 0.5, y: 3.85, w: 3.5, h: 0.35,
    fontSize: 14, fontFace: "Georgia", color: C.purple, bold: true,
  });
  s.addText("「In Egypt, I saw the Pyramids. They are 4,500 years old!\nMandela was in prison for 27 years. He taught us about peace...」", {
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
  s.addText("🌍 My Africa Page  我的非洲页", {
    x: 1.2, y: 1.1, w: 7.6, h: 0.5,
    fontSize: 18, fontFace: "Georgia", color: C.teal, bold: true, align: "center",
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 2.0, y: 1.6, w: 6.0, h: 0.02, fill: { color: C.teal },
  });

  const entries = [
    "🌎 I visited: 非洲 Africa \u2014 埃及, 肯尼亚, 南非",
    "👀 I saw: 金字塔 Pyramids, 狮子 Lions on Safari, 桌山 Table Mountain",
    "🍜 I ate: 库莎丽 Koshari, 烤肉 Nyama Choma, Braai!",
    "🎭 I learned: 在肯尼亚说Jambo, 在南非说Sawubona, Ubuntu精神",
    "💬 My sentence: 我想去肯尼亚看狮子。南非最有名的是彩虹之国。",
    "🎨 [Draw your favorite animal or landmark here!]",
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

  s.addText("和同伴分享你的非洲页！\nShare your Africa page with a partner!", {
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
    "「我最想去___，因为___。」",
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


// SLIDE 43 — 🪪 非洲签证章
(() => {
  const s = pptx.addSlide(); countSlide();
  s.background = { fill: C.dark };
  goldBars(s);
  slideNum(s);

  s.addText("🪪 非洲签证章  Africa Visa Stamp", {
    x: 0.5, y: 0.4, w: 9.0, h: 0.6,
    fontSize: 26, fontFace: "Georgia", color: C.gold, bold: true, align: "center",
  });

  s.addShape(pptx.shapes.OVAL, {
    x: 3.0, y: 1.2, w: 4.0, h: 3.5,
    fill: { color: C.dark },
    line: { color: C.contAfrica, width: 4, dashType: "dash" },
  });

  s.addText([
    { text: "AFRICA", options: { fontSize: 34, fontFace: "Georgia", color: C.contAfrica, bold: true, breakLine: true } },
    { text: "非洲", options: { fontSize: 24, fontFace: "Georgia", color: C.gold, breakLine: true } },
    { text: "\u2713 VISITED", options: { fontSize: 20, fontFace: "Georgia", color: C.quizGreen, bold: true, breakLine: true } },
    { text: "6/9/2025", options: { fontSize: 14, fontFace: "Calibri", color: C.white, breakLine: true } },
    { text: "埃及 \u00B7 肯尼亚 \u00B7 南非", options: { fontSize: 14, fontFace: "Calibri", color: C.lightAmber, breakLine: true } },
  ], { x: 3.0, y: 1.4, w: 4.0, h: 3.1, align: "center", valign: "middle" });

  s.addText("恭喜你完成非洲之旅！Congratulations!", {
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
    { text: "Flight 航班: GR-003", options: { fontSize: 18, fontFace: "Georgia", color: C.contEurope, bold: true, breakLine: true } },
    { text: "Destination 目的地: 欧洲 EUROPE", options: { fontSize: 18, fontFace: "Georgia", color: C.contEurope, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "明天我们去欧洲！", options: { fontSize: 16, fontFace: "Georgia", color: C.dark, bold: true, breakLine: true } },
    { text: "那里有什么著名的建筑？", options: { fontSize: 16, fontFace: "Georgia", color: C.dark, breakLine: true } },
    { text: "Tomorrow we fly to Europe! What famous buildings are there?", options: { fontSize: 12, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 1.8, y: 2.1, w: 6.4, h: 2.4, align: "center", valign: "middle" });

  s.addText("See you tomorrow, explorers!  明天见，小探险家们！", {
    x: 0.5, y: 4.85, w: 9.0, h: 0.4,
    fontSize: 14, fontFace: "Calibri", color: C.gold, bold: true, align: "center",
  });
})();


// ── Save ──
const outPath = path.join(__dirname, "day2_africa.pptx");
pptx.writeFile({ fileName: outPath }).then(() => {
  console.log("Created: " + outPath);
  console.log("Total slides: " + slideCount);
}).catch((err) => {
  console.error("Error:", err);
});
