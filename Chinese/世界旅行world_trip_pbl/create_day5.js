/**
 * Day 5: 国际文化展 Exhibition (6/12) — 环游世界之旅 PBL Summer Camp
 * FINAL version with continent comparison focus
 * Run: node create_day5.js
 */
const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10.0" x 5.625"
pres.author = "谷雨中文 GR EDU";
pres.title = "Global Explorer Camp · Day 5: 国际文化展";

// ── Color palette (no # prefix) ──
const C = {
  pri:    "6A1B9A",
  sec:    "F3E5F5",
  accent: "FFD54F",
  dark:   "4A148C",
  white:  "FFFFFF",
  black:  "212121",
  midPur: "9C27B0",
  ltPur:  "CE93D8",
  palePur:"EDE7F6",
  gray:   "616161",
  asiaRed:    "C62828",
  africaAmber:"FF8F00",
  euroBlue:   "1565C0",
  ameriGreen: "2E7D32",
  green4: "4CAF50",
  blue4:  "1976D2",
  gold:   "F9A825",
};

// ── Helpers ──
function topBar(s, clr) {
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 9.8, h: 0.07, fill: { color: clr || C.accent },
  });
}
function bottomBar(s, clr) {
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.43, w: 9.8, h: 0.07, fill: { color: clr || C.accent },
  });
}

function hdr(s, text, opts) {
  s.addText(text, Object.assign({
    x: 0.4, y: 0.15, w: 9.0, h: 0.65,
    fontFace: "Georgia", fontSize: 28, color: C.dark, bold: true,
  }, opts || {}));
}

function accentLine(s, x, y, w) {
  s.addShape(pres.shapes.RECTANGLE, {
    x: x || 0.4, y: y || 0.8, w: w || 4.0, h: 0.05, fill: { color: C.accent },
  });
}

function card(s, x, y, w, h, fill, line) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: w, h: h, rectRadius: 0.12,
    fill: { color: fill || C.white },
    line: line ? { color: line, width: 2 } : undefined,
    shadow: { type: "outer", blur: 4, offset: 2, color: "999999", opacity: 0.25 },
  });
}

function goldStar(s, x, y, w, rot) {
  s.addShape(pres.shapes.STAR_5_POINT, {
    x: x, y: y, w: w, h: w, fill: { color: C.accent }, rotate: rot || 0,
  });
}

function stampCircles(s, xStart, y, size) {
  var colors = [C.asiaRed, C.africaAmber, C.euroBlue, C.ameriGreen];
  var labels = ["亚", "非", "欧", "美"];
  colors.forEach(function(clr, i) {
    var xp = xStart + i * (size + 0.25);
    s.addShape(pres.shapes.OVAL, {
      x: xp, y: y, w: size, h: size,
      fill: { color: clr }, line: { color: C.accent, width: 2 },
    });
    s.addText(labels[i], {
      x: xp, y: y, w: size, h: size,
      fontSize: Math.round(size * 18), fontFace: "Calibri", color: C.white,
      bold: true, align: "center", valign: "middle",
    });
  });
}

function tableRow(s, x, y, w, h, cells, isHeader) {
  var colW = w / cells.length;
  cells.forEach(function(cell, i) {
    var xp = x + i * colW;
    s.addShape(pres.shapes.RECTANGLE, {
      x: xp, y: y, w: colW, h: h,
      fill: { color: isHeader ? C.dark : (i % 2 === 0 ? C.sec : C.white) },
      line: { color: C.ltPur, width: 1 },
    });
    s.addText(cell, {
      x: xp, y: y, w: colW, h: h,
      fontSize: isHeader ? 12 : 11, fontFace: "Calibri",
      color: isHeader ? C.white : C.black,
      bold: isHeader, align: "center", valign: "middle", breakLine: true,
    });
  });
}

// ============================================================
// Build all slides
// ============================================================
function build() {

  // ============================================================
  // SLIDE 1 — Title: Global Explorer Camp 最终站
  // ============================================================
  var s1 = pres.addSlide();
  s1.background = { color: C.dark };
  goldStar(s1, 0.2, 0.2, 0.5, 15);
  goldStar(s1, 8.8, 0.15, 0.4, -10);
  goldStar(s1, 8.2, 4.6, 0.45, 20);
  goldStar(s1, 0.3, 4.5, 0.35, -5);
  goldStar(s1, 4.6, 0.1, 0.3, 25);

  accentLine(s1, 1.5, 0.35, 6.2);

  s1.addText("Global Explorer Camp", {
    x: 0.5, y: 0.5, w: 8.5, h: 0.55,
    fontSize: 22, fontFace: "Georgia", color: C.accent,
    italic: true, align: "center",
  });
  s1.addText("最终站：国际文化展", {
    x: 0.5, y: 1.1, w: 8.5, h: 0.8,
    fontSize: 36, fontFace: "Georgia", color: C.white,
    bold: true, align: "center",
  });
  s1.addText("International Culture Exhibition", {
    x: 0.5, y: 1.85, w: 8.5, h: 0.45,
    fontSize: 20, fontFace: "Calibri", color: C.ltPur,
    italic: true, align: "center",
  });

  accentLine(s1, 2.5, 2.4, 4.2);

  card(s1, 2.0, 2.65, 5.2, 1.2, C.pri, C.accent);
  s1.addText("ARRIVED", {
    x: 2.0, y: 2.6, w: 5.2, h: 0.5,
    fontSize: 28, fontFace: "Georgia", color: C.accent,
    bold: true, align: "center", valign: "middle",
  });
  stampCircles(s1, 2.55, 3.15, 0.55);

  s1.addText("6/12 周五 Friday  |  谷雨中文 GR EDU", {
    x: 0.5, y: 4.1, w: 8.5, h: 0.4,
    fontSize: 16, fontFace: "Calibri", color: C.accent, align: "center",
  });
  accentLine(s1, 1.5, 4.7, 6.2);

  // ============================================================
  // SLIDE 2 — 我们的旅程 All 4 stamps
  // ============================================================
  var s2 = pres.addSlide();
  s2.background = { color: C.sec };
  topBar(s2);

  hdr(s2, "我们的旅程 Our Journey", { align: "center" });

  var journeyData = [
    { label: "亚洲\nAsia", color: C.asiaRed, emoji: "🌏" },
    { label: "非洲\nAfrica", color: C.africaAmber, emoji: "🌍" },
    { label: "欧洲\nEurope", color: C.euroBlue, emoji: "🌍" },
    { label: "美洲\nAmericas", color: C.ameriGreen, emoji: "🌎" },
    { label: "文化展\nExhibition", color: C.accent, emoji: "🏆" },
  ];

  // Connecting line
  s2.addShape(pres.shapes.RECTANGLE, {
    x: 1.2, y: 2.05, w: 7.0, h: 0.06, fill: { color: C.pri },
  });

  journeyData.forEach(function(item, i) {
    var xp = 0.8 + i * 1.7;
    s2.addShape(pres.shapes.OVAL, {
      x: xp, y: 1.55, w: 1.1, h: 1.1,
      fill: { color: item.color }, line: { color: C.dark, width: 2 },
    });
    s2.addText(item.emoji, {
      x: xp, y: 1.55, w: 1.1, h: 0.6,
      fontSize: 24, align: "center", valign: "middle",
    });
    s2.addText(item.label, {
      x: xp - 0.2, y: 2.8, w: 1.5, h: 0.7,
      fontSize: 13, fontFace: "Calibri", color: C.black,
      bold: true, align: "center", valign: "top", breakLine: true,
    });
  });

  card(s2, 1.5, 3.8, 6.2, 0.7, C.accent);
  s2.addText("4个签证章全部集齐! All 4 visa stamps collected!", {
    x: 1.5, y: 3.8, w: 6.2, h: 0.7,
    fontSize: 18, fontFace: "Calibri", color: C.dark,
    bold: true, align: "center", valign: "middle",
  });
  bottomBar(s2);

  // ============================================================
  // SLIDE 3 — 今天日程
  // ============================================================
  var s3 = pres.addSlide();
  s3.background = { color: C.white };
  topBar(s3);

  hdr(s3, "今天日程 Today's Schedule", { align: "center" });

  card(s3, 0.4, 0.85, 9.0, 0.4, C.pri);
  s3.addText("上午 Morning", {
    x: 0.4, y: 0.85, w: 9.0, h: 0.4,
    fontSize: 15, fontFace: "Georgia", color: C.white,
    bold: true, align: "center", valign: "middle",
  });

  var morningItems = [
    { time: "15 min", label: "四大洲知识回顾+对比 Continent Review & Compare", emoji: "🌍" },
    { time: "15 min", label: "护照封面设计 Passport Cover Design", emoji: "🎨" },
    { time: "10 min", label: "展示练习 Presentation Prep", emoji: "🎤" },
    { time: "5 min", label: "展览规则 + Gallery Walk", emoji: "🚶" },
    { time: "10 min", label: "展览+知识大挑战 Exhibition + Quiz", emoji: "🏆" },
  ];
  morningItems.forEach(function(item, i) {
    var yp = 1.35 + i * 0.45;
    card(s3, 1.0, yp, 7.2, 0.38, C.sec);
    s3.addText(item.emoji + "  " + item.label + "  (" + item.time + ")", {
      x: 1.0, y: yp, w: 7.2, h: 0.38,
      fontSize: 12, fontFace: "Calibri", color: C.pri,
      bold: true, align: "center", valign: "middle",
    });
  });

  card(s3, 0.4, 3.65, 9.0, 0.4, C.dark);
  s3.addText("下午 Afternoon", {
    x: 0.4, y: 3.65, w: 9.0, h: 0.4,
    fontSize: 15, fontFace: "Georgia", color: C.accent,
    bold: true, align: "center", valign: "middle",
  });

  var afternoonItems = [
    { time: "15 min", label: "Final Page: 我最喜欢的国家", emoji: "📖" },
    { time: "15 min", label: "Show & Tell 展示", emoji: "🎤" },
    { time: "10 min", label: "投票+颁奖 Voting+Awards", emoji: "🏆" },
    { time: "10 min", label: "结营仪式 Closing Ceremony", emoji: "🎓" },
  ];
  afternoonItems.forEach(function(item, i) {
    var yp = 4.15 + i * 0.32;
    card(s3, 1.0, yp, 7.2, 0.28, C.sec);
    s3.addText(item.emoji + "  " + item.label + "  (" + item.time + ")", {
      x: 1.0, y: yp, w: 7.2, h: 0.28,
      fontSize: 11, fontFace: "Calibri", color: C.dark,
      bold: true, align: "center", valign: "middle",
    });
  });
  bottomBar(s3);

  // ============================================================
  // SLIDE 4 — 🌏 亚洲回顾 (3-column cards, red accent)
  // ============================================================
  var s4 = pres.addSlide();
  s4.background = { color: C.white };
  topBar(s4, C.asiaRed);

  hdr(s4, "🌏 亚洲回顾 Asia Review", { color: C.asiaRed });
  accentLine(s4, 0.4, 0.8, 3.0);

  var asiaCountries = [
    {
      name: "🇨🇳 中国 China",
      facts: [
        "👋 握手 Handshake",
        "🥢 筷子 Chopsticks",
        "🥟 饺子 Dumplings",
        "⚠️ 不插筷子在饭里",
        "Don't stick chopsticks\nupright in rice",
      ],
    },
    {
      name: "🇯🇵 日本 Japan",
      facts: [
        "🙇 鞠躬 Bow",
        "👟 脱鞋进屋 Remove shoes",
        "🍣 寿司 Sushi",
        "🍜 吃面出声OK!",
        "Slurping noodles\nis polite!",
      ],
    },
    {
      name: "🇮🇳 印度 India",
      facts: [
        "🙏 合十 Namaste",
        "🤚 用右手 Right hand",
        "🍛 咖喱 Curry",
        "⚠️ 不摸别人的头",
        "Don't touch\nsomeone's head",
      ],
    },
  ];

  asiaCountries.forEach(function(country, i) {
    var xp = 0.3 + i * 3.1;
    card(s4, xp, 0.9, 2.95, 0.45, C.asiaRed);
    s4.addText(country.name, {
      x: xp, y: 0.9, w: 2.95, h: 0.45,
      fontSize: 13, fontFace: "Georgia", color: C.white,
      bold: true, align: "center", valign: "middle",
    });
    card(s4, xp, 1.4, 2.95, 3.7, C.white, C.asiaRed);
    s4.addText(country.facts.join("\n"), {
      x: xp + 0.1, y: 1.5, w: 2.75, h: 3.5,
      fontSize: 12, fontFace: "Calibri", color: C.black,
      align: "center", valign: "top", breakLine: true, lineSpacing: 17,
    });
  });
  bottomBar(s4, C.asiaRed);

  // ============================================================
  // SLIDE 5 — 🌍 非洲回顾 (3-column cards, amber accent)
  // ============================================================
  var s5 = pres.addSlide();
  s5.background = { color: C.white };
  topBar(s5, C.africaAmber);

  hdr(s5, "🌍 非洲回顾 Africa Review", { color: C.africaAmber });
  accentLine(s5, 0.4, 0.8, 3.0);

  var africaCountries = [
    {
      name: "🇪🇬 埃及 Egypt",
      facts: [
        "💋 亲脸颊 Cheek kiss",
        "🤚 用右手 Right hand",
        "🧆 鹰嘴豆泥 Hummus",
        "🏛️ 金字塔 Pyramids",
        "Ancient wonder\nof the world",
      ],
    },
    {
      name: "🇰🇪 肯尼亚 Kenya",
      facts: [
        "👋 Jambo! Hello!",
        "👆 不用手指指人",
        "Don't point at people",
        "🍚 Ugali 乌咖力",
        "🦁 Safari 大草原",
      ],
    },
    {
      name: "🇿🇦 南非 S. Africa",
      facts: [
        "🤝 三步握手",
        "Three-step handshake",
        "🗣️ 11种官方语言!",
        "11 official languages!",
        "🥩 Braai 南非烤肉",
      ],
    },
  ];

  africaCountries.forEach(function(country, i) {
    var xp = 0.3 + i * 3.1;
    card(s5, xp, 0.9, 2.95, 0.45, C.africaAmber);
    s5.addText(country.name, {
      x: xp, y: 0.9, w: 2.95, h: 0.45,
      fontSize: 13, fontFace: "Georgia", color: C.white,
      bold: true, align: "center", valign: "middle",
    });
    card(s5, xp, 1.4, 2.95, 3.7, C.white, C.africaAmber);
    s5.addText(country.facts.join("\n"), {
      x: xp + 0.1, y: 1.5, w: 2.75, h: 3.5,
      fontSize: 12, fontFace: "Calibri", color: C.black,
      align: "center", valign: "top", breakLine: true, lineSpacing: 17,
    });
  });
  bottomBar(s5, C.africaAmber);

  // ============================================================
  // SLIDE 6 — 🌍 欧洲回顾 (3-column cards, blue accent)
  // ============================================================
  var s6 = pres.addSlide();
  s6.background = { color: C.white };
  topBar(s6, C.euroBlue);

  hdr(s6, "🌍 欧洲回顾 Europe Review", { color: C.euroBlue });
  accentLine(s6, 0.4, 0.8, 3.0);

  var euroCountries = [
    {
      name: "🇫🇷 法国 France",
      facts: [
        "💋 亲脸颊 Bisou",
        "🚫 不催服务员",
        "Don't rush waiters",
        "🥐 可颂 Croissant",
        "🗼 埃菲尔铁塔",
      ],
    },
    {
      name: "🇮🇹 意大利 Italy",
      facts: [
        "👋 Ciao! + 手势",
        "Hand gestures = Italian!",
        "🍕 不加菠萝在披萨上!",
        "No pineapple on pizza!",
        "🏛️ 罗马 Rome",
      ],
    },
    {
      name: "🇬🇧 英国 UK",
      facts: [
        "🚶 排队! Queue up!",
        "🙏 Please & Thank you",
        "Very polite culture",
        "🐟 炸鱼薯条",
        "Fish and Chips",
      ],
    },
  ];

  euroCountries.forEach(function(country, i) {
    var xp = 0.3 + i * 3.1;
    card(s6, xp, 0.9, 2.95, 0.45, C.euroBlue);
    s6.addText(country.name, {
      x: xp, y: 0.9, w: 2.95, h: 0.45,
      fontSize: 13, fontFace: "Georgia", color: C.white,
      bold: true, align: "center", valign: "middle",
    });
    card(s6, xp, 1.4, 2.95, 3.7, C.white, C.euroBlue);
    s6.addText(country.facts.join("\n"), {
      x: xp + 0.1, y: 1.5, w: 2.75, h: 3.5,
      fontSize: 12, fontFace: "Calibri", color: C.black,
      align: "center", valign: "top", breakLine: true, lineSpacing: 17,
    });
  });
  bottomBar(s6, C.euroBlue);

  // ============================================================
  // SLIDE 7 — 🌎 美洲回顾 (3-column cards, green accent)
  // ============================================================
  var s7 = pres.addSlide();
  s7.background = { color: C.white };
  topBar(s7, C.ameriGreen);

  hdr(s7, "🌎 美洲回顾 Americas Review", { color: C.ameriGreen });
  accentLine(s7, 0.4, 0.8, 3.0);

  var ameriCountries = [
    {
      name: "🇺🇸 美国 USA",
      facts: [
        "🤝 握手 Handshake",
        "💵 给小费 Tip 15-20%",
        "🍔 汉堡 Hamburger",
        "🗽 自由女神像",
        "Statue of Liberty",
      ],
    },
    {
      name: "🇲🇽 墨西哥 Mexico",
      facts: [
        "🤗 拥抱 Hug + Abrazo",
        "👋 Hola! 你好!",
        "🌮 Taco 玉米饼",
        "🍫 巧克力发源地!",
        "Birthplace of chocolate!",
      ],
    },
    {
      name: "🇧🇷 巴西 Brazil",
      facts: [
        "💋 亲两次 Two kisses",
        "⚽ 别说足球不好!",
        "Don't dis soccer!",
        "🥩 Churrasco 烤肉",
        "🎭 狂欢节 Carnival",
      ],
    },
  ];

  ameriCountries.forEach(function(country, i) {
    var xp = 0.3 + i * 3.1;
    card(s7, xp, 0.9, 2.95, 0.45, C.ameriGreen);
    s7.addText(country.name, {
      x: xp, y: 0.9, w: 2.95, h: 0.45,
      fontSize: 13, fontFace: "Georgia", color: C.white,
      bold: true, align: "center", valign: "middle",
    });
    card(s7, xp, 1.4, 2.95, 3.7, C.white, C.ameriGreen);
    s7.addText(country.facts.join("\n"), {
      x: xp + 0.1, y: 1.5, w: 2.75, h: 3.5,
      fontSize: 12, fontFace: "Calibri", color: C.black,
      align: "center", valign: "top", breakLine: true, lineSpacing: 17,
    });
  });
  bottomBar(s7, C.ameriGreen);

  // ============================================================
  // SLIDE 8 — 🌍 四大洲打招呼方式大对比
  // ============================================================
  var s8 = pres.addSlide();
  s8.background = { color: C.white };
  topBar(s8);

  hdr(s8, "🌍 四大洲打招呼方式大对比", { fontSize: 26 });
  s8.addText("How People Say Hello Around the World", {
    x: 0.4, y: 0.65, w: 9.0, h: 0.3,
    fontSize: 14, fontFace: "Calibri", color: C.pri,
    italic: true, align: "left",
  });

  // Table header
  var tblX = 0.3, tblW = 9.2, rowH = 0.55;
  tableRow(s8, tblX, 1.05, tblW, 0.45, ["大洲 Continent", "国家 Countries", "打招呼方式 Greeting"], true);

  var greetingRows = [
    ["🌏 亚洲 Asia", "中国 / 日本 / 印度", "握手 / 鞠躬 / 合十Namaste"],
    ["🌍 非洲 Africa", "埃及 / 肯尼亚 / 南非", "亲脸颊 / Jambo! / 三步握手"],
    ["🌍 欧洲 Europe", "法国 / 意大利 / 英国", "亲脸颊 / Ciao!+手势 / Hello+握手"],
    ["🌎 美洲 Americas", "美国 / 墨西哥 / 巴西", "握手 / 拥抱Hola / 亲两次"],
  ];

  var rowColors = [C.asiaRed, C.africaAmber, C.euroBlue, C.ameriGreen];
  greetingRows.forEach(function(row, i) {
    var yp = 1.5 + i * rowH;
    // Colored left edge
    s8.addShape(pres.shapes.RECTANGLE, {
      x: tblX, y: yp, w: 0.08, h: rowH,
      fill: { color: rowColors[i] },
    });
    var colWidths = [2.6, 3.3, 3.3];
    var xCurr = tblX;
    row.forEach(function(cell, ci) {
      s8.addShape(pres.shapes.RECTANGLE, {
        x: xCurr, y: yp, w: colWidths[ci], h: rowH,
        fill: { color: i % 2 === 0 ? C.sec : C.white },
        line: { color: C.ltPur, width: 0.5 },
      });
      s8.addText(cell, {
        x: xCurr + 0.1, y: yp, w: colWidths[ci] - 0.2, h: rowH,
        fontSize: 12, fontFace: "Calibri", color: C.black,
        align: "center", valign: "middle", breakLine: true,
      });
      xCurr += colWidths[ci];
    });
  });

  // Bottom insight
  card(s8, 1.0, 3.85, 7.4, 0.75, C.accent);
  s8.addText("全世界的人都在说「你好」，只是方式不同!\nEveryone says hello — just in different ways!", {
    x: 1.0, y: 3.85, w: 7.4, h: 0.75,
    fontSize: 15, fontFace: "Calibri", color: C.dark,
    bold: true, align: "center", valign: "middle", breakLine: true,
  });

  // Visual: greeting emojis
  var greetEmojis = ["🤝", "🙇", "🙏", "💋", "🤗"];
  greetEmojis.forEach(function(em, i) {
    s8.addText(em, {
      x: 1.2 + i * 1.5, y: 4.75, w: 1.0, h: 0.5,
      fontSize: 28, align: "center", valign: "middle",
    });
  });
  bottomBar(s8);

  // ============================================================
  // SLIDE 9 — 🍽️ 四大洲饮食文化对比
  // ============================================================
  var s9 = pres.addSlide();
  s9.background = { color: C.white };
  topBar(s9);

  hdr(s9, "🍽️ 四大洲饮食文化对比", { fontSize: 26 });
  s9.addText("Eating Cultures Around the World", {
    x: 0.4, y: 0.65, w: 9.0, h: 0.3,
    fontSize: 14, fontFace: "Calibri", color: C.pri,
    italic: true, align: "left",
  });

  // Utensils comparison cards
  var utensilCards = [
    {
      emoji: "🥢", label: "用筷子 Chopsticks",
      detail: "中国、日本\nChina, Japan",
      color: C.asiaRed,
    },
    {
      emoji: "🤚", label: "用手 By Hand",
      detail: "印度、肯尼亚\n埃及 (右手!)",
      color: C.africaAmber,
    },
    {
      emoji: "🍴", label: "用刀叉 Fork & Knife",
      detail: "法国、意大利\n英国、美国",
      color: C.euroBlue,
    },
  ];

  utensilCards.forEach(function(uc, i) {
    var xp = 0.3 + i * 3.1;
    card(s9, xp, 1.0, 2.95, 2.0, C.sec, uc.color);
    s9.addText(uc.emoji, {
      x: xp, y: 1.05, w: 2.95, h: 0.55,
      fontSize: 30, align: "center", valign: "middle",
    });
    s9.addText(uc.label, {
      x: xp, y: 1.55, w: 2.95, h: 0.4,
      fontSize: 14, fontFace: "Calibri", color: uc.color,
      bold: true, align: "center", valign: "middle",
    });
    s9.addText(uc.detail, {
      x: xp + 0.1, y: 2.0, w: 2.75, h: 0.85,
      fontSize: 12, fontFace: "Calibri", color: C.black,
      align: "center", valign: "top", breakLine: true, lineSpacing: 16,
    });
  });

  // BBQ comparison
  card(s9, 0.3, 3.2, 9.2, 1.0, C.white, C.accent);
  s9.addText("🔥 每个地方都有BBQ/烤肉!  Everyone loves BBQ!", {
    x: 0.3, y: 3.2, w: 9.2, h: 0.4,
    fontSize: 14, fontFace: "Georgia", color: C.dark,
    bold: true, align: "center", valign: "middle",
  });

  var bbqList = [
    { flag: "🇨🇳", label: "中国烧烤\nChinese BBQ" },
    { flag: "🇰🇪", label: "Nyama Choma\n肯尼亚烤肉" },
    { flag: "🇿🇦", label: "Braai\n南非烤肉" },
    { flag: "🇺🇸", label: "American BBQ\n美式烤肉" },
    { flag: "🇧🇷", label: "Churrasco\n巴西烤肉" },
  ];

  bbqList.forEach(function(bb, i) {
    var xp = 0.5 + i * 1.8;
    s9.addText(bb.flag + "\n" + bb.label, {
      x: xp, y: 3.6, w: 1.65, h: 0.55,
      fontSize: 9, fontFace: "Calibri", color: C.black,
      align: "center", valign: "top", breakLine: true,
    });
  });

  card(s9, 1.5, 4.4, 6.4, 0.55, C.accent);
  s9.addText("全世界的人都爱烤肉!\nEveryone around the world loves grilled meat!", {
    x: 1.5, y: 4.4, w: 6.4, h: 0.55,
    fontSize: 13, fontFace: "Calibri", color: C.dark,
    bold: true, align: "center", valign: "middle", breakLine: true,
  });
  bottomBar(s9);

  // ============================================================
  // SLIDE 10 — 🎨 护照封面设计 (15min)
  // ============================================================
  var s10 = pres.addSlide();
  s10.background = { color: C.sec };
  topBar(s10);

  hdr(s10, "🎨 护照封面设计 Passport Cover Design", { align: "center", fontSize: 26 });
  s10.addText("15 min", {
    x: 7.8, y: 0.15, w: 1.5, h: 0.4,
    fontSize: 14, fontFace: "Calibri", color: C.pri, bold: true, align: "center",
  });

  // Passport mockup (left)
  card(s10, 0.6, 0.9, 3.8, 4.2, C.dark, C.accent);
  s10.addText("PASSPORT", {
    x: 0.6, y: 1.05, w: 3.8, h: 0.4,
    fontSize: 18, fontFace: "Georgia", color: C.accent,
    bold: true, align: "center",
  });
  s10.addText("护 照", {
    x: 0.6, y: 1.4, w: 3.8, h: 0.4,
    fontSize: 20, fontFace: "Georgia", color: C.accent,
    bold: true, align: "center",
  });

  s10.addShape(pres.shapes.OVAL, {
    x: 1.6, y: 2.0, w: 1.6, h: 1.6,
    fill: { color: C.pri }, line: { color: C.accent, width: 2 },
  });
  s10.addText("🌍", {
    x: 1.6, y: 2.0, w: 1.6, h: 1.6,
    fontSize: 48, align: "center", valign: "middle",
  });

  s10.addText("My Travel Passport", {
    x: 0.6, y: 3.7, w: 3.8, h: 0.35,
    fontSize: 14, fontFace: "Calibri", color: C.accent,
    italic: true, align: "center",
  });
  stampCircles(s10, 0.95, 4.2, 0.5);

  // Instructions (right)
  card(s10, 4.8, 0.9, 4.6, 4.2, C.white, C.pri);
  s10.addText("封面设计要素 Cover Elements", {
    x: 4.8, y: 0.95, w: 4.6, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.dark,
    bold: true, align: "center",
  });

  var coverSteps = [
    "1. 写上你的名字（中文 + 英文）",
    "   Write your name in Chinese & English",
    "2. 写「My Travel Passport」",
    "3. 画4个签证章（亚/非/欧/美）",
    "   Draw 4 visa stamps",
    "4. 装饰封面（画画/贴纸/颜色）",
    "   Decorate your cover",
    "5. 可以加国旗或地标!",
    "   Add flags or landmarks!",
  ];
  s10.addText(coverSteps.join("\n"), {
    x: 5.0, y: 1.4, w: 4.2, h: 3.5,
    fontSize: 12, fontFace: "Calibri", color: C.black,
    align: "left", valign: "top", breakLine: true, lineSpacing: 17,
  });
  bottomBar(s10);

  // ============================================================
  // SLIDE 11 — 🎤 展示练习 3 levels
  // ============================================================
  var s11 = pres.addSlide();
  s11.background = { color: C.white };
  topBar(s11);

  hdr(s11, "🎤 展示练习 Presentation Practice", { align: "center", fontSize: 26 });

  var levels = [
    {
      level: "🟢 零基础 Beginner",
      color: C.green4, bg: "E8F5E9",
      content: "「我喜欢___」\n(point to pictures)\n指着图片说一个词\n画你最喜欢的食物",
    },
    {
      level: "🔵 Level 2-3",
      color: C.blue4, bg: "E3F2FD",
      content: "「我最喜欢___」\n「他们用___打招呼」\n「我想吃___」\n说3个句子",
    },
    {
      level: "🟣 Level 4+",
      color: C.pri, bg: C.sec,
      content: "2分钟展示 2-min talk\n比较两个大洲的礼节\nCompare greeting styles\n讲饮食文化差异",
    },
  ];

  levels.forEach(function(lv, i) {
    var xp = 0.3 + i * 3.1;
    card(s11, xp, 0.85, 2.95, 0.5, lv.color);
    s11.addText(lv.level, {
      x: xp, y: 0.85, w: 2.95, h: 0.5,
      fontSize: 13, fontFace: "Georgia", color: C.white,
      bold: true, align: "center", valign: "middle",
    });
    card(s11, xp, 1.4, 2.95, 3.5, lv.bg, lv.color);
    s11.addText(lv.content, {
      x: xp + 0.15, y: 1.55, w: 2.65, h: 3.2,
      fontSize: 13, fontFace: "Calibri", color: C.black,
      align: "center", valign: "top", breakLine: true, lineSpacing: 18,
    });
  });

  s11.addText("练习1-2遍，然后向同学展示! Practice 1-2 times, then present!", {
    x: 0.5, y: 5.0, w: 8.5, h: 0.35,
    fontSize: 13, fontFace: "Calibri", color: C.pri,
    bold: true, align: "center",
  });
  bottomBar(s11);

  // ============================================================
  // SLIDE 12 — 展览规则 + Gallery Walk 指南
  // ============================================================
  var s12 = pres.addSlide();
  s12.background = { color: C.sec };
  topBar(s12);

  hdr(s12, "展览规则 + Gallery Walk 指南", { align: "center", fontSize: 26 });

  // Rules (left column)
  card(s12, 0.3, 0.85, 4.4, 4.2, C.white, C.pri);
  s12.addText("展览规则 Exhibition Rules", {
    x: 0.3, y: 0.9, w: 4.4, h: 0.4,
    fontSize: 15, fontFace: "Georgia", color: C.dark,
    bold: true, align: "center",
  });

  var rules = [
    { emoji: "🗣️", rule: "用中文/英文说!" },
    { emoji: "👂", rule: "认真听同学展示" },
    { emoji: "❓", rule: "问一个问题" },
    { emoji: "👏", rule: "鼓掌表示赞赏" },
  ];
  rules.forEach(function(r, i) {
    var yp = 1.45 + i * 0.8;
    s12.addText(r.emoji + "  " + r.rule, {
      x: 0.6, y: yp, w: 3.8, h: 0.65,
      fontSize: 15, fontFace: "Calibri", color: C.black,
      bold: true, align: "left", valign: "middle",
    });
  });

  // Gallery Walk steps (right column)
  card(s12, 4.95, 0.85, 4.5, 4.2, C.white, C.accent);
  s12.addText("Gallery Walk 步骤", {
    x: 4.95, y: 0.9, w: 4.5, h: 0.4,
    fontSize: 15, fontFace: "Georgia", color: C.dark,
    bold: true, align: "center",
  });

  var walkSteps = [
    { step: "Step 1", label: "参观3个同学的展位", color: C.asiaRed },
    { step: "Step 2", label: "问一个问题", color: C.africaAmber },
    { step: "Step 3", label: "写一个优点", color: C.euroBlue },
    { step: "Step 4", label: "回到你的展位", color: C.ameriGreen },
  ];
  walkSteps.forEach(function(ws, i) {
    var yp = 1.45 + i * 0.8;
    s12.addShape(pres.shapes.OVAL, {
      x: 5.2, y: yp + 0.08, w: 0.5, h: 0.5,
      fill: { color: ws.color },
    });
    s12.addText(ws.step, {
      x: 5.2, y: yp + 0.08, w: 0.5, h: 0.5,
      fontSize: 8, fontFace: "Georgia", color: C.white,
      bold: true, align: "center", valign: "middle",
    });
    s12.addText(ws.label, {
      x: 5.85, y: yp, w: 3.4, h: 0.65,
      fontSize: 14, fontFace: "Calibri", color: C.black,
      bold: true, align: "left", valign: "middle",
    });
  });
  bottomBar(s12);

  // ============================================================
  // SLIDE 13 — 展览开始! GO!
  // ============================================================
  var s13 = pres.addSlide();
  s13.background = { color: C.dark };

  goldStar(s13, 0.5, 0.5, 0.6, 10);
  goldStar(s13, 8.5, 0.4, 0.5, -15);
  goldStar(s13, 0.3, 4.2, 0.45, 20);
  goldStar(s13, 8.6, 4.3, 0.55, -8);
  goldStar(s13, 2.0, 4.5, 0.3, 30);
  goldStar(s13, 7.0, 4.6, 0.35, -20);

  s13.addShape(pres.shapes.OVAL, {
    x: 3.0, y: 0.8, w: 3.5, h: 3.5,
    fill: { color: C.accent }, line: { color: C.white, width: 4 },
  });
  s13.addText("GO!", {
    x: 3.0, y: 0.8, w: 3.5, h: 3.5,
    fontSize: 72, fontFace: "Georgia", color: C.dark,
    bold: true, align: "center", valign: "middle",
  });

  s13.addText("展览时间开始! Exhibition Time!", {
    x: 0.5, y: 4.4, w: 8.5, h: 0.5,
    fontSize: 22, fontFace: "Calibri", color: C.accent,
    bold: true, align: "center",
  });

  // ============================================================
  // SLIDE 14 — 🏆 世界知识大挑战 12题
  // ============================================================
  var s14 = pres.addSlide();
  s14.background = { color: C.white };
  topBar(s14);

  hdr(s14, "🏆 世界知识大挑战 World Knowledge Quiz", { fontSize: 26, align: "center" });
  s14.addText("12 Questions  每洲3题  礼节+食物+地理", {
    x: 0.5, y: 0.65, w: 8.5, h: 0.3,
    fontSize: 13, fontFace: "Calibri", color: C.pri,
    italic: true, align: "center",
  });

  var quiz = [
    "1. 在日本，见面时应该? → 鞠躬",
    "2. 筷子不能插在哪里? → 饭里",
    "3. 印度人用什么手吃饭? → 右手",
    "4. 肯尼亚人说什么打招呼? → Jambo",
    "5. 金字塔在哪个国家? → 埃及",
    "6. 南非有多少官方语言? → 11种",
    "7. 意大利人不在披萨上放? → 菠萝",
    "8. 英国人最重要的礼节? → 排队",
    "9. 法国的经典面包是? → 可颂",
    "10. 在美国要给服务员? → 小费",
    "11. 巧克力发源于哪国? → 墨西哥",
    "12. 巴西人见面亲几次? → 两次",
  ];

  var leftQ = quiz.slice(0, 6);
  var rightQ = quiz.slice(6);

  card(s14, 0.3, 1.0, 4.5, 3.8, C.sec, C.pri);
  s14.addText("亚洲 + 非洲", {
    x: 0.3, y: 1.0, w: 4.5, h: 0.35,
    fontSize: 12, fontFace: "Georgia", color: C.pri,
    bold: true, align: "center",
  });
  s14.addText(leftQ.join("\n"), {
    x: 0.5, y: 1.35, w: 4.1, h: 3.35,
    fontSize: 11, fontFace: "Calibri", color: C.black,
    align: "left", valign: "top", breakLine: true, lineSpacing: 19,
  });

  card(s14, 5.0, 1.0, 4.5, 3.8, C.sec, C.pri);
  s14.addText("欧洲 + 美洲", {
    x: 5.0, y: 1.0, w: 4.5, h: 0.35,
    fontSize: 12, fontFace: "Georgia", color: C.pri,
    bold: true, align: "center",
  });
  s14.addText(rightQ.join("\n"), {
    x: 5.2, y: 1.35, w: 4.1, h: 3.35,
    fontSize: 11, fontFace: "Calibri", color: C.black,
    align: "left", valign: "top", breakLine: true, lineSpacing: 19,
  });

  s14.addText("抢答! Fastest correct answer wins!", {
    x: 0.5, y: 4.95, w: 8.5, h: 0.35,
    fontSize: 13, fontFace: "Calibri", color: C.pri,
    bold: true, align: "center",
  });
  bottomBar(s14);

  // ============================================================
  // SLIDE 15 — 答案揭晓
  // ============================================================
  var s15 = pres.addSlide();
  s15.background = { color: C.sec };
  topBar(s15);

  hdr(s15, "答案揭晓 Answer Reveal", { align: "center" });

  var answers = [
    { q: "1. 日本见面 →", a: " 鞠躬 Bow", color: C.asiaRed },
    { q: "2. 筷子不插 →", a: " 饭里 Rice", color: C.asiaRed },
    { q: "3. 印度吃饭 →", a: " 右手 Right hand", color: C.asiaRed },
    { q: "4. 肯尼亚 →", a: " Jambo!", color: C.africaAmber },
    { q: "5. 金字塔 →", a: " 埃及 Egypt", color: C.africaAmber },
    { q: "6. 南非语言 →", a: " 11种!", color: C.africaAmber },
    { q: "7. 披萨不放 →", a: " 菠萝 Pineapple", color: C.euroBlue },
    { q: "8. 英国礼节 →", a: " 排队 Queue", color: C.euroBlue },
    { q: "9. 法国面包 →", a: " 可颂 Croissant", color: C.euroBlue },
    { q: "10. 美国 →", a: " 小费 Tip", color: C.ameriGreen },
    { q: "11. 巧克力 →", a: " 墨西哥 Mexico", color: C.ameriGreen },
    { q: "12. 巴西亲 →", a: " 两次 Twice", color: C.ameriGreen },
  ];

  answers.forEach(function(ans, i) {
    var col = i < 6 ? 0 : 1;
    var row = i < 6 ? i : i - 6;
    var xp = 0.3 + col * 4.8;
    var yp = 0.85 + row * 0.72;
    card(s15, xp, yp, 4.4, 0.6, C.white, ans.color);
    s15.addText(ans.q + ans.a, {
      x: xp + 0.15, y: yp, w: 4.1, h: 0.6,
      fontSize: 12, fontFace: "Calibri", color: C.black,
      bold: true, align: "left", valign: "middle",
    });
  });

  s15.addText("你答对了几题? How many did you get right?", {
    x: 0.5, y: 5.05, w: 8.5, h: 0.3,
    fontSize: 14, fontFace: "Calibri", color: C.pri,
    bold: true, align: "center",
  });
  bottomBar(s15);

  // ============================================================
  // SLIDE 16 — 下午开始 Section Divider
  // ============================================================
  var s16 = pres.addSlide();
  s16.background = { color: C.dark };

  accentLine(s16, 1.5, 1.2, 6.2);
  accentLine(s16, 1.5, 3.2, 6.2);
  goldStar(s16, 0.3, 0.3, 0.45, 12);
  goldStar(s16, 8.5, 0.25, 0.4, -15);
  goldStar(s16, 0.25, 4.5, 0.35, 20);
  goldStar(s16, 8.6, 4.4, 0.4, -8);

  s16.addText("下午开始", {
    x: 0.75, y: 1.4, w: 7.7, h: 0.75,
    fontSize: 42, fontFace: "Georgia", color: C.accent,
    bold: true, align: "center", valign: "middle",
  });
  s16.addText("Afternoon Session", {
    x: 0.75, y: 2.2, w: 7.7, h: 0.45,
    fontSize: 22, fontFace: "Calibri", color: C.white,
    align: "center", italic: true,
  });
  s16.addText("Final Page + Show & Tell + Ceremony", {
    x: 0.75, y: 2.7, w: 7.7, h: 0.4,
    fontSize: 18, fontFace: "Calibri", color: C.ltPur,
    align: "center",
  });

  // ============================================================
  // SLIDE 17 — 🎨 Final Page: 我最喜欢的国家
  // ============================================================
  var s17 = pres.addSlide();
  s17.background = { color: C.sec };
  topBar(s17);

  hdr(s17, "🎨 Final Page: 我最喜欢的国家", { align: "center", fontSize: 26 });
  s17.addText("选择你这周学的12个国家中最喜欢的一个", {
    x: 0.3, y: 0.65, w: 9.0, h: 0.3,
    fontSize: 13, fontFace: "Calibri", color: C.pri,
    italic: true, align: "center",
  });

  // Page mockup
  card(s17, 0.5, 1.0, 8.6, 4.1, C.white, C.accent);

  var fpItems = [
    { emoji: "🌎", label: "Which country?", desc: "我最喜欢哪个国家?" },
    { emoji: "🏴", label: "Draw the flag", desc: "画国旗 Draw it!" },
    { emoji: "👋", label: "How do they say hello?", desc: "他们怎么打招呼?" },
    { emoji: "🍽️", label: "What would I eat?", desc: "我会吃什么食物?" },
    { emoji: "⚠️", label: "One thing to remember", desc: "去那里要注意什么礼节?" },
    { emoji: "💬", label: "Write a sentence", desc: "「My favorite country is ___ because ___」" },
  ];

  fpItems.forEach(function(item, i) {
    var yp = 1.1 + i * 0.6;
    card(s17, 0.8, yp, 8.0, 0.5, i % 2 === 0 ? C.sec : C.palePur);
    s17.addText(item.emoji + "  " + item.label, {
      x: 0.9, y: yp, w: 3.5, h: 0.5,
      fontSize: 12, fontFace: "Calibri", color: C.dark,
      bold: true, align: "left", valign: "middle",
    });
    s17.addText(item.desc, {
      x: 4.4, y: yp, w: 4.2, h: 0.5,
      fontSize: 11, fontFace: "Calibri", color: C.black,
      align: "left", valign: "middle",
    });
  });
  bottomBar(s17);

  // ============================================================
  // SLIDE 18 — Final Page 分层
  // ============================================================
  var s18 = pres.addSlide();
  s18.background = { color: C.white };
  topBar(s18);

  hdr(s18, "Final Page 分层 Differentiated", { align: "center" });

  var fpLevels = [
    {
      level: "🟢 画 + 词\nDraw + Words",
      color: C.green4, bg: "E8F5E9",
      content: "画你最喜欢的国家\nDraw your favorite country\n写2-3个词\nWrite 2-3 words\n例: 日本 寿司 鞠躬",
    },
    {
      level: "🔵 Sentences\n句子",
      color: C.blue4, bg: "E3F2FD",
      content: "写2-3个句子\n「我最喜欢___」\n「他们用___打招呼」\n「我想吃___」\n「去那里要___」",
    },
    {
      level: "🟣 Paragraph\n段落 + 礼节细节",
      color: C.pri, bg: C.sec,
      content: "写一段话 Write a paragraph\n包含: 国家/打招呼方式\n食物/要注意的礼节\nInclude etiquette details\n比较和你文化的不同",
    },
  ];

  fpLevels.forEach(function(lv, i) {
    var xp = 0.3 + i * 3.1;
    card(s18, xp, 0.85, 2.95, 0.55, lv.color);
    s18.addText(lv.level, {
      x: xp, y: 0.85, w: 2.95, h: 0.55,
      fontSize: 12, fontFace: "Georgia", color: C.white,
      bold: true, align: "center", valign: "middle", breakLine: true,
    });
    card(s18, xp, 1.45, 2.95, 3.6, lv.bg, lv.color);
    s18.addText(lv.content, {
      x: xp + 0.15, y: 1.6, w: 2.65, h: 3.3,
      fontSize: 12, fontFace: "Calibri", color: C.black,
      align: "center", valign: "top", breakLine: true, lineSpacing: 18,
    });
  });
  bottomBar(s18);

  // ============================================================
  // SLIDE 19 — 🎤 Show & Tell
  // ============================================================
  var s19 = pres.addSlide();
  s19.background = { color: C.sec };
  topBar(s19);

  hdr(s19, "🎤 Show and Tell 展示时间", { align: "center", fontSize: 28 });

  card(s19, 0.5, 0.9, 8.6, 1.8, C.white, C.pri);
  s19.addText("每位同学展示你的护照 (1-2 min)\nEach student presents their passport", {
    x: 0.7, y: 0.95, w: 8.2, h: 0.7,
    fontSize: 18, fontFace: "Calibri", color: C.dark,
    bold: true, align: "center", valign: "middle", breakLine: true,
  });

  var showSteps = [
    "1. 举起护照给大家看 Hold up your passport",
    "2. 说你最喜欢的国家 Share your favorite country",
    "3. 说那里的打招呼方式和食物 Greeting & food",
    "4. 说为什么喜欢 Tell us why",
  ];
  s19.addText(showSteps.join("\n"), {
    x: 1.5, y: 1.65, w: 6.5, h: 0.95,
    fontSize: 13, fontFace: "Calibri", color: C.black,
    align: "left", valign: "top", breakLine: true, lineSpacing: 17,
  });

  card(s19, 0.5, 2.95, 8.6, 2.2, C.white, C.accent);
  s19.addText("观众 Audience — 问一个问题!", {
    x: 0.5, y: 3.0, w: 8.6, h: 0.45,
    fontSize: 16, fontFace: "Georgia", color: C.pri,
    bold: true, align: "center",
  });

  var audienceQ = [
    "「你最喜欢什么食物?」What food did you like best?",
    "「他们怎么打招呼?」How do they say hello?",
    "「去那里要注意什么?」What etiquette to remember?",
    "「我也喜欢___!」I also like ___!",
  ];
  s19.addText(audienceQ.join("\n"), {
    x: 1.2, y: 3.5, w: 7.2, h: 1.5,
    fontSize: 13, fontFace: "Calibri", color: C.black,
    align: "left", valign: "top", breakLine: true, lineSpacing: 18,
  });
  bottomBar(s19);

  // ============================================================
  // SLIDE 20 — 参观提示卡
  // ============================================================
  var s20 = pres.addSlide();
  s20.background = { color: C.white };
  topBar(s20);

  hdr(s20, "参观提示卡 Visitor Prompt Cards", { align: "center" });

  card(s20, 0.4, 0.85, 4.3, 4.2, C.sec, C.pri);
  s20.addText("可以问的问题 Questions to Ask", {
    x: 0.4, y: 0.9, w: 4.3, h: 0.45,
    fontSize: 14, fontFace: "Georgia", color: C.dark,
    bold: true, align: "center",
  });

  var questions = [
    "「你最喜欢哪个国家?」",
    "Which country is your favorite?",
    "",
    "「他们怎么打招呼?」",
    "How do they greet each other?",
    "",
    "「那里的人吃什么?」",
    "What do people eat there?",
    "",
    "「去那里要注意什么礼节?」",
    "What etiquette to remember?",
  ];
  s20.addText(questions.join("\n"), {
    x: 0.6, y: 1.4, w: 3.9, h: 3.5,
    fontSize: 12, fontFace: "Calibri", color: C.black,
    align: "left", valign: "top", breakLine: true, lineSpacing: 15,
  });

  card(s20, 5.0, 0.85, 4.3, 4.2, C.sec, C.accent);
  s20.addText("反馈句型 Feedback Starters", {
    x: 5.0, y: 0.9, w: 4.3, h: 0.45,
    fontSize: 14, fontFace: "Georgia", color: C.dark,
    bold: true, align: "center",
  });

  var feedback = [
    "「你的护照很漂亮!」",
    "Your passport is beautiful!",
    "",
    "「我喜欢你画的___」",
    "I like your drawing of ___",
    "",
    "「你说得很好!」",
    "You spoke really well!",
    "",
    "「我也想去___!」",
    "I want to visit ___ too!",
  ];
  s20.addText(feedback.join("\n"), {
    x: 5.2, y: 1.4, w: 3.9, h: 3.5,
    fontSize: 12, fontFace: "Calibri", color: C.black,
    align: "left", valign: "top", breakLine: true, lineSpacing: 15,
  });
  bottomBar(s20);

  // ============================================================
  // SLIDE 21 — 🗳️ 投票
  // ============================================================
  var s21 = pres.addSlide();
  s21.background = { color: C.sec };
  topBar(s21);

  hdr(s21, "🗳️ 投票时间 Voting Time", { align: "center", fontSize: 30 });

  var voteCategories = [
    { emoji: "🎨", label: "Best Design\n最佳设计", color: C.asiaRed },
    { emoji: "🗣️", label: "Best Expression\n最佳表达", color: C.euroBlue },
    { emoji: "💡", label: "Most Creative\n最有创意", color: C.ameriGreen },
    { emoji: "🧠", label: "Best Knowledge\n最佳知识", color: C.africaAmber },
  ];

  voteCategories.forEach(function(cat, i) {
    var xp = 0.4 + i * 2.3;
    card(s21, xp, 0.95, 2.1, 3.0, C.white, cat.color);
    s21.addText(cat.emoji, {
      x: xp, y: 1.1, w: 2.1, h: 0.8,
      fontSize: 40, align: "center", valign: "middle",
    });
    card(s21, xp + 0.15, 1.95, 1.8, 0.5, cat.color);
    s21.addText(cat.label, {
      x: xp + 0.15, y: 1.95, w: 1.8, h: 0.5,
      fontSize: 12, fontFace: "Georgia", color: C.white,
      bold: true, align: "center", valign: "middle", breakLine: true,
    });
    s21.addText("投票给:\n___________", {
      x: xp + 0.15, y: 2.6, w: 1.8, h: 1.0,
      fontSize: 12, fontFace: "Calibri", color: C.black,
      align: "center", valign: "middle", breakLine: true,
    });
  });

  s21.addText("每个同学投4票，每类选一位! Each student votes once per category!", {
    x: 0.5, y: 4.2, w: 8.5, h: 0.4,
    fontSize: 14, fontFace: "Calibri", color: C.pri,
    bold: true, align: "center",
  });
  bottomBar(s21);

  // ============================================================
  // SLIDE 22 — 🏆 颁奖
  // ============================================================
  var s22 = pres.addSlide();
  s22.background = { color: C.dark };

  goldStar(s22, 0.3, 0.3, 0.5, 10);
  goldStar(s22, 8.7, 0.2, 0.45, -12);
  goldStar(s22, 0.2, 4.5, 0.4, 20);
  goldStar(s22, 8.8, 4.6, 0.35, -18);

  hdr(s22, "🏆 颁奖典礼 Awards Ceremony", {
    fontSize: 32, color: C.accent, align: "center",
  });

  var awards = [
    { emoji: "🎨", label: "Best Design 最佳设计" },
    { emoji: "🗣️", label: "Best Expression 最佳表达" },
    { emoji: "💡", label: "Most Creative 最有创意" },
    { emoji: "🧠", label: "Best Knowledge 最佳知识" },
  ];

  awards.forEach(function(aw, i) {
    var yp = 1.05 + i * 0.65;
    card(s22, 1.5, yp, 6.2, 0.55, C.pri, C.accent);
    s22.addText(aw.emoji + "  " + aw.label + "  →  ___________", {
      x: 1.5, y: yp, w: 6.2, h: 0.55,
      fontSize: 16, fontFace: "Calibri", color: C.white,
      bold: true, align: "center", valign: "middle",
    });
  });

  card(s22, 1.0, 3.8, 7.2, 0.9, C.accent);
  s22.addText("Global Explorer 环球探索者\nALL STUDENTS! 每一位同学!", {
    x: 1.0, y: 3.8, w: 7.2, h: 0.9,
    fontSize: 18, fontFace: "Georgia", color: C.dark,
    bold: true, align: "center", valign: "middle", breakLine: true,
  });

  s22.addText("恭喜大家! Congratulations to everyone!", {
    x: 0.5, y: 4.85, w: 8.5, h: 0.4,
    fontSize: 16, fontFace: "Calibri", color: C.accent,
    bold: true, align: "center",
  });

  // ============================================================
  // SLIDE 23 — 本周回顾
  // ============================================================
  var s23 = pres.addSlide();
  s23.background = { color: C.white };
  topBar(s23);

  hdr(s23, "本周回顾 Week in Review", { align: "center" });

  var summaryCards = [
    { emoji: "🌍", stat: "4", label: "大洲\nContinents", color: C.pri },
    { emoji: "🏳️", stat: "12", label: "国家\nCountries", color: C.midPur },
    { emoji: "👋", stat: "12", label: "种打招呼\nGreetings", color: C.asiaRed },
    { emoji: "📝", stat: "30+", label: "词汇\nVocabulary", color: C.euroBlue },
    { emoji: "🎭", stat: "✓", label: "文化礼节\nEtiquette", color: C.africaAmber },
  ];

  summaryCards.forEach(function(sc, i) {
    var xp = 0.3 + i * 1.85;
    card(s23, xp, 0.85, 1.7, 2.2, C.sec, sc.color);
    s23.addText(sc.emoji, {
      x: xp, y: 0.9, w: 1.7, h: 0.5,
      fontSize: 28, align: "center", valign: "middle",
    });
    s23.addText(sc.stat, {
      x: xp, y: 1.35, w: 1.7, h: 0.6,
      fontSize: 32, fontFace: "Georgia", color: sc.color,
      bold: true, align: "center", valign: "middle",
    });
    s23.addText(sc.label, {
      x: xp, y: 1.95, w: 1.7, h: 0.6,
      fontSize: 11, fontFace: "Calibri", color: C.black,
      bold: true, align: "center", valign: "top", breakLine: true,
    });
  });

  var daySummary = [
    { day: "Day 1", cont: "亚洲 Asia", color: C.asiaRed },
    { day: "Day 2", cont: "非洲 Africa", color: C.africaAmber },
    { day: "Day 3", cont: "欧洲 Europe", color: C.euroBlue },
    { day: "Day 4", cont: "美洲 Americas", color: C.ameriGreen },
    { day: "Day 5", cont: "文化展 Exhibition", color: C.pri },
  ];

  daySummary.forEach(function(ds, i) {
    var xp = 0.4 + i * 1.8;
    card(s23, xp, 3.3, 1.65, 0.65, ds.color);
    s23.addText(ds.day + "\n" + ds.cont, {
      x: xp, y: 3.3, w: 1.65, h: 0.65,
      fontSize: 10, fontFace: "Calibri", color: C.white,
      bold: true, align: "center", valign: "middle", breakLine: true,
    });
  });

  s23.addText("了不起的一周! What an amazing week!", {
    x: 0.5, y: 4.2, w: 8.5, h: 0.4,
    fontSize: 16, fontFace: "Calibri", color: C.pri,
    bold: true, align: "center",
  });
  bottomBar(s23);

  // ============================================================
  // SLIDE 24 — 句型回顾
  // ============================================================
  var s24 = pres.addSlide();
  s24.background = { color: C.sec };
  topBar(s24);

  hdr(s24, "句型回顾 Sentence Patterns Review", { align: "center" });

  var patterns = [
    {
      day: "Day 1 亚洲", color: C.asiaRed,
      pattern: "「我喜欢吃___」I like to eat ___\n「在___人们用___打招呼」In ___ people greet by ___",
    },
    {
      day: "Day 2 非洲", color: C.africaAmber,
      pattern: "「这是___，它来自___」This is ___, it comes from ___",
    },
    {
      day: "Day 3 欧洲", color: C.euroBlue,
      pattern: "「我想去___看___」I want to go to ___ to see ___",
    },
    {
      day: "Day 4 美洲", color: C.ameriGreen,
      pattern: "「如果没有___就没有___」Without ___, there would be no ___",
    },
  ];

  patterns.forEach(function(pt, i) {
    var yp = 0.9 + i * 1.1;
    card(s24, 0.5, yp, 2.2, 0.85, pt.color);
    s24.addText(pt.day, {
      x: 0.5, y: yp, w: 2.2, h: 0.85,
      fontSize: 14, fontFace: "Georgia", color: C.white,
      bold: true, align: "center", valign: "middle",
    });
    card(s24, 2.9, yp, 6.3, 0.85, C.white, pt.color);
    s24.addText(pt.pattern, {
      x: 3.1, y: yp, w: 5.9, h: 0.85,
      fontSize: 13, fontFace: "Calibri", color: C.black,
      bold: true, align: "left", valign: "middle", breakLine: true,
    });
  });
  bottomBar(s24);

  // ============================================================
  // SLIDE 25 — 🎓 Certificate (gold border)
  // ============================================================
  var s25 = pres.addSlide();
  s25.background = { color: C.white };

  // Gold border
  s25.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.3, y: 0.2, w: 9.0, h: 5.0,
    fill: { color: C.white }, rectRadius: 0.2,
    line: { color: C.accent, width: 4 },
  });
  s25.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.5, y: 0.4, w: 8.6, h: 4.6,
    fill: { color: C.white }, rectRadius: 0.15,
    line: { color: C.dark, width: 1.5, dashType: "dash" },
  });

  goldStar(s25, 0.5, 0.4, 0.4, 10);
  goldStar(s25, 8.7, 0.4, 0.4, -10);
  goldStar(s25, 0.5, 4.6, 0.4, 15);
  goldStar(s25, 8.7, 4.6, 0.4, -15);

  s25.addText("🎓", {
    x: 4.2, y: 0.5, w: 1.0, h: 0.7,
    fontSize: 36, align: "center", valign: "middle",
  });

  s25.addText("环球探索者证书", {
    x: 1.0, y: 1.1, w: 7.4, h: 0.6,
    fontSize: 32, fontFace: "Georgia", color: C.dark,
    bold: true, align: "center",
  });
  s25.addText("Global Explorer Certificate", {
    x: 1.0, y: 1.6, w: 7.4, h: 0.4,
    fontSize: 18, fontFace: "Georgia", color: C.pri,
    italic: true, align: "center",
  });

  accentLine(s25, 2.5, 2.05, 4.4);

  s25.addText("___________________________", {
    x: 2.0, y: 2.2, w: 5.4, h: 0.5,
    fontSize: 20, fontFace: "Georgia", color: C.black,
    align: "center", valign: "middle",
  });
  s25.addText("Student Name 学生姓名", {
    x: 2.0, y: 2.65, w: 5.4, h: 0.3,
    fontSize: 11, fontFace: "Calibri", color: C.pri,
    align: "center",
  });

  s25.addText("has explored 4 continents, 12 countries\nlearned greetings, food, and cultural etiquette\n成功探索了4个大洲12个国家，学会了打招呼、美食和文化礼节!", {
    x: 0.8, y: 2.95, w: 7.8, h: 0.8,
    fontSize: 12, fontFace: "Calibri", color: C.black,
    align: "center", valign: "middle", breakLine: true, lineSpacing: 15,
  });

  var certLabels = ["亚洲", "非洲", "欧洲", "美洲"];
  var certColors = [C.asiaRed, C.africaAmber, C.euroBlue, C.ameriGreen];
  certColors.forEach(function(clr, i) {
    var xp = 2.4 + i * 1.3;
    s25.addShape(pres.shapes.OVAL, {
      x: xp, y: 3.85, w: 0.7, h: 0.7,
      fill: { color: clr }, line: { color: C.accent, width: 2 },
    });
    s25.addText("\u2713", {
      x: xp, y: 3.75, w: 0.7, h: 0.5,
      fontSize: 18, color: C.white, bold: true, align: "center", valign: "middle",
    });
    s25.addText(certLabels[i], {
      x: xp - 0.15, y: 4.55, w: 1.0, h: 0.25,
      fontSize: 10, fontFace: "Calibri", color: C.black,
      align: "center",
    });
  });

  s25.addText("谷雨中文 GR EDU  ·  6/12/2025", {
    x: 1.0, y: 4.8, w: 7.4, h: 0.3,
    fontSize: 12, fontFace: "Calibri", color: C.pri,
    align: "center",
  });

  // ============================================================
  // SLIDE 26 — 📸 回忆墙
  // ============================================================
  var s26 = pres.addSlide();
  s26.background = { color: C.dark };

  s26.addText("📸 回忆墙 Memory Wall", {
    x: 0.3, y: 0.2, w: 9.0, h: 0.6,
    fontSize: 32, fontFace: "Georgia", color: C.accent,
    bold: true, align: "center",
  });

  var photoFrames = [
    { x: 0.5, y: 1.0, w: 2.6, h: 1.8 },
    { x: 3.4, y: 1.0, w: 2.6, h: 1.8 },
    { x: 6.3, y: 1.0, w: 2.6, h: 1.8 },
    { x: 1.8, y: 3.0, w: 2.6, h: 1.8 },
    { x: 4.8, y: 3.0, w: 2.6, h: 1.8 },
  ];
  var photoLabels = ["Day 1 亚洲", "Day 2 非洲", "Day 3 欧洲", "Day 4 美洲", "Day 5 文化展"];

  photoFrames.forEach(function(frame, i) {
    s26.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: frame.x, y: frame.y, w: frame.w, h: frame.h,
      fill: { color: C.white, transparency: 90 }, rectRadius: 0.1,
      line: { color: C.accent, width: 2 },
    });
    s26.addText("📷", {
      x: frame.x, y: frame.y, w: frame.w, h: frame.h * 0.7,
      fontSize: 28, align: "center", valign: "middle",
    });
    s26.addText(photoLabels[i], {
      x: frame.x, y: frame.y + frame.h * 0.65, w: frame.w, h: frame.h * 0.35,
      fontSize: 11, fontFace: "Calibri", color: C.white,
      bold: true, align: "center", valign: "middle",
    });
  });

  s26.addText("Global Explorer Camp 2025", {
    x: 0.5, y: 4.95, w: 8.5, h: 0.35,
    fontSize: 14, fontFace: "Calibri", color: C.accent,
    italic: true, align: "center",
  });

  // ============================================================
  // SLIDE 27 — 🧳 暑假旅行挑战
  // ============================================================
  var s27 = pres.addSlide();
  s27.background = { color: C.sec };
  topBar(s27);

  hdr(s27, "🧳 暑假旅行挑战 Summer Travel Challenge", { align: "center", fontSize: 26 });
  s27.addText("用这周学的知识，和爸妈旅行时注意这些礼节!", {
    x: 0.4, y: 0.65, w: 9.0, h: 0.3,
    fontSize: 13, fontFace: "Calibri", color: C.pri,
    italic: true, align: "center",
  });

  var challenges = [
    {
      emoji: "👋",
      label: "学一种新的打招呼方式",
      desc: "Learn a new greeting",
      color: C.asiaRed,
    },
    {
      emoji: "🍳",
      label: "做一道外国菜",
      desc: "Cook a foreign dish (sushi/taco/pizza)",
      color: C.africaAmber,
    },
    {
      emoji: "📖",
      label: "读一本关于其他国家的书",
      desc: "Read about another country",
      color: C.euroBlue,
    },
    {
      emoji: "🗺️",
      label: "旅行时观察当地礼节",
      desc: "Observe local etiquette when traveling",
      color: C.ameriGreen,
    },
    {
      emoji: "👨‍👩‍👧",
      label: "教爸妈说一种外语「你好」",
      desc: "Teach your parents hello in another language",
      color: C.pri,
    },
  ];

  challenges.forEach(function(ch, i) {
    var yp = 1.05 + i * 0.8;
    card(s27, 0.5, yp, 8.6, 0.7, C.white, ch.color);
    s27.addText(ch.emoji, {
      x: 0.6, y: yp, w: 0.7, h: 0.7,
      fontSize: 24, align: "center", valign: "middle",
    });
    s27.addText(ch.label, {
      x: 1.4, y: yp, w: 3.5, h: 0.7,
      fontSize: 14, fontFace: "Calibri", color: C.dark,
      bold: true, align: "left", valign: "middle",
    });
    s27.addText(ch.desc, {
      x: 5.0, y: yp, w: 3.8, h: 0.7,
      fontSize: 12, fontFace: "Calibri", color: C.gray,
      italic: true, align: "left", valign: "middle",
    });
  });

  s27.addText("你好 / Bonjour / Hola / Jambo / Namaste / Ciao / Oi", {
    x: 0.5, y: 5.05, w: 8.5, h: 0.3,
    fontSize: 13, fontFace: "Calibri", color: C.pri,
    bold: true, align: "center",
  });
  bottomBar(s27);

  // ============================================================
  // SLIDE 28 — ✈️ 旅程结束，探索不止!
  // ============================================================
  var s28 = pres.addSlide();
  s28.background = { color: C.dark };

  goldStar(s28, 0.3, 0.3, 0.5, 12);
  goldStar(s28, 8.7, 0.2, 0.45, -10);
  goldStar(s28, 0.2, 4.5, 0.4, 18);
  goldStar(s28, 8.8, 4.4, 0.5, -15);
  goldStar(s28, 4.5, 0.1, 0.35, 25);

  s28.addText("✈️", {
    x: 3.8, y: 0.4, w: 1.5, h: 0.8,
    fontSize: 42, align: "center", valign: "middle",
  });

  s28.addText("旅程结束，探索不止!", {
    x: 0.3, y: 1.2, w: 9.0, h: 0.7,
    fontSize: 34, fontFace: "Georgia", color: C.accent,
    bold: true, align: "center",
  });
  s28.addText("The journey ends, but exploration never stops!", {
    x: 0.3, y: 1.9, w: 9.0, h: 0.4,
    fontSize: 18, fontFace: "Calibri", color: C.white,
    italic: true, align: "center",
  });

  accentLine(s28, 2.0, 2.5, 5.2);

  // 4 continent farewell stamps
  var farewellStamps = [
    { emoji: "🌏", label: "亚洲\n再见!", color: C.asiaRed },
    { emoji: "🌍", label: "非洲\nKwaheri!", color: C.africaAmber },
    { emoji: "🌍", label: "欧洲\nAu revoir!", color: C.euroBlue },
    { emoji: "🌎", label: "美洲\nAdios!", color: C.ameriGreen },
  ];

  farewellStamps.forEach(function(fs, i) {
    var xp = 1.0 + i * 2.0;
    s28.addShape(pres.shapes.OVAL, {
      x: xp, y: 2.8, w: 1.5, h: 1.5,
      fill: { color: fs.color }, line: { color: C.accent, width: 2 },
    });
    s28.addText(fs.emoji + "\n" + fs.label, {
      x: xp, y: 2.8, w: 1.5, h: 1.5,
      fontSize: 13, fontFace: "Calibri", color: C.white,
      bold: true, align: "center", valign: "middle", breakLine: true,
    });
  });

  s28.addText("谷雨中文 GR EDU  ·  Global Explorer Camp 2025  ·  See you!", {
    x: 0.5, y: 4.6, w: 8.5, h: 0.35,
    fontSize: 14, fontFace: "Calibri", color: C.accent,
    italic: true, align: "center",
  });
  accentLine(s28, 2.0, 5.1, 5.2);

  // ============================================================
  // SAVE
  // ============================================================
  return pres.writeFile({ fileName: "/Users/Huan/projects/summercourse/Chinese/world_trip_pbl/day5_exhibition.pptx" });
}

build().then(function() {
  console.log("Created: day5_exhibition.pptx (28 slides)");
}).catch(function(err) {
  console.error("Build failed:", err);
  process.exit(1);
});
