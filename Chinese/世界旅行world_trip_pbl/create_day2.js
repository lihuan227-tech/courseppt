/**
 * Day 2: 非洲 Africa (6/9) — Global Explorer Camp 环球探索沉浸式夏令营
 * 33 slides — 3 countries deep dive: 埃及 Egypt, 肯尼亚 Kenya, 南非 South Africa
 * Run: node create_day2.js
 */
const pptxgen = require("pptxgenjs");
const https = require("https");
const http = require("http");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "谷雨中文 GR EDU";
pres.title = "Global Explorer Camp · Day 2: 非洲 Africa";

// ── Color palette (no # prefix) ──
const C = {
  primary:   "FF8F00",
  secondary: "FFF3E0",
  accent:    "4E342E",
  dark:      "3E2723",
  white:     "FFFFFF",
  gold:      "FFD54F",
  lightGold: "FFE082",
  warmGray:  "5D4037",
  orange:    "FF6F00",
  cream:     "FFF8E1",
  green:     "388E3C",
  lightGreen:"C8E6C9",
  blue:      "1565C0",
  lightBlue: "BBDEFB",
  red:       "C62828",
  lightRed:  "FFCDD2",
  purple:    "7B1FA2",
  lightPurple:"E1BEE7",
  gray:      "9E9E9E",
  darkText:  "212121",
  overlay:   "000000",

  egyptRed:   "CE1126",
  egyptWhite: "FFFFFF",
  egyptBlack: "000000",
  kenyaRed:   "BB0000",
  kenyaGreen: "006600",
  kenyaBlack: "000000",
  saRed:      "DE3831",
  saBlue:     "002395",
  saGreen:    "007A4D",
  saGold:     "FFB612",
};

// ── Image URLs (Wikimedia Commons) ──
const IMG = {
  pyramids:      "https://upload.wikimedia.org/wikipedia/commons/thumb/e/e3/Kheops-Pyramid.jpg/1280px-Kheops-Pyramid.jpg",
  lion:          "https://upload.wikimedia.org/wikipedia/commons/thumb/7/73/Lion_waiting_in_Namibia.jpg/1280px-Lion_waiting_in_Namibia.jpg",
  kilimanjaro:   "https://upload.wikimedia.org/wikipedia/commons/thumb/6/6b/Mt._Kilimanjaro_12.2008.JPG/1280px-Mt._Kilimanjaro_12.2008.JPG",
  tableMountain: "https://upload.wikimedia.org/wikipedia/commons/thumb/4/4b/Table_Mountain_DanieVDM.jpg/1280px-Table_Mountain_DanieVDM.jpg",
  victoriaFalls: "https://upload.wikimedia.org/wikipedia/commons/thumb/5/57/Victoria_Falls_from_the_Air.jpg/1280px-Victoria_Falls_from_the_Air.jpg",
};

// ── Fetch image as base64 with redirect support ──
function fetchImageBase64(url, maxRedirects) {
  if (maxRedirects === undefined) maxRedirects = 5;
  return new Promise(function(resolve, reject) {
    if (maxRedirects <= 0) return reject(new Error("Too many redirects"));
    var mod = url.startsWith("https") ? https : http;
    mod.get(url, { headers: { "User-Agent": "Mozilla/5.0" } }, function(res) {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return fetchImageBase64(res.headers.location, maxRedirects - 1).then(resolve).catch(reject);
      }
      if (res.statusCode !== 200) return reject(new Error("HTTP " + res.statusCode));
      var chunks = [];
      res.on("data", function(chunk) { chunks.push(chunk); });
      res.on("end", function() {
        var buf = Buffer.concat(chunks);
        var mime = res.headers["content-type"] || "image/jpeg";
        resolve("image/" + mime.split("/")[1] + ";base64," + buf.toString("base64"));
      });
      res.on("error", reject);
    }).on("error", reject);
  });
}

// ── Helpers ──
function goldBars(s) {
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 9.8, h: 0.06, fill: { color: C.primary },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.44, w: 9.8, h: 0.06, fill: { color: C.primary },
  });
}

function titleBar(s, text, barColor, txtColor) {
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 9.8, h: 0.72,
    fill: { color: barColor || C.accent },
  });
  s.addText(text, {
    x: 0.35, y: 0.05, w: 9.1, h: 0.62,
    fontSize: 28, fontFace: "Georgia",
    color: txtColor || C.gold, bold: true,
  });
}

function footer(s, text) {
  s.addText(text || "谷雨中文 GR EDU  |  Global Explorer Camp  |  Day 2 非洲", {
    x: 0.3, y: 5.1, w: 9.2, h: 0.3,
    fontSize: 10, fontFace: "Calibri",
    color: C.warmGray, align: "center",
  });
}

function addImageWithFallback(s, url, x, y, w, h, fallbackColor, label) {
  s.addShape(pres.shapes.RECTANGLE, {
    x: x, y: y, w: w, h: h,
    fill: { color: fallbackColor || C.lightGold },
  });
  s.addText(label || "Photo", {
    x: x, y: y + h * 0.38, w: w, h: h * 0.24,
    fontSize: 13, fontFace: "Calibri",
    color: C.warmGray, align: "center", italic: true,
  });
  try {
    s.addImage({ path: url, x: x, y: y, w: w, h: h });
  } catch (e) {
    // fallback rectangle already in place
  }
}

function checkSlide(s, title, questions, bgColor) {
  s.background = { fill: bgColor || C.secondary };
  titleBar(s, title);

  questions.forEach(function(q, i) {
    var yy = 1.0 + i * 1.35;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.5, y: yy, w: 8.8, h: 1.15,
      rectRadius: 0.12,
      fill: { color: i % 2 === 0 ? C.cream : C.white },
      line: { color: C.primary, width: 1.5 },
    });
    s.addText([
      { text: q.q, options: { fontSize: 16, fontFace: "Calibri", color: C.accent, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 4, breakLine: true } },
      { text: q.a, options: { fontSize: 14, fontFace: "Calibri", color: C.primary, breakLine: true } },
    ], { x: 0.7, y: yy + 0.05, w: 8.4, h: 1.05, valign: "middle" });
  });
  footer(s);
}


// ══════════════════════════════════════════════════════════════
// SLIDE 1 — Boarding Time (GR-002, brown bg)
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.dark };
  goldBars(s);

  s.addText([
    { text: "Global Explorer Camp", options: { fontSize: 34, fontFace: "Georgia", color: C.gold, bold: true, breakLine: true } },
    { text: "环球探索沉浸式夏令营", options: { fontSize: 22, fontFace: "Georgia", color: C.lightGold, breakLine: true } },
  ], { x: 0.5, y: 0.4, w: 9.0, h: 1.1, align: "center" });

  // Boarding pass card
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 2.0, y: 1.8, w: 5.8, h: 2.3,
    fill: { color: C.accent }, rectRadius: 0.18,
    shadow: { type: "outer", blur: 8, offset: 3, color: "000000", opacity: 0.4 },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 2.0, y: 1.8, w: 5.8, h: 0.38,
    fill: { color: C.primary },
  });
  s.addText("BOARDING PASS \u00B7 登机牌", {
    x: 2.0, y: 1.8, w: 5.8, h: 0.38,
    fontSize: 14, fontFace: "Georgia", color: C.dark, align: "center", bold: true,
  });

  s.addText([
    { text: "航班 GR-002", options: { fontSize: 22, fontFace: "Georgia", color: C.gold, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "From: 亚洲 ASIA", options: { fontSize: 16, fontFace: "Calibri", color: C.lightGold, breakLine: true } },
    { text: "To: 非洲 AFRICA", options: { fontSize: 18, fontFace: "Calibri", color: C.white, bold: true, breakLine: true } },
  ], { x: 2.3, y: 2.3, w: 3.6, h: 1.5 });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.2, y: 2.25, w: 0.04, h: 1.5,
    fill: { color: C.gold, transparency: 50 },
  });

  s.addText([
    { text: "6/9", options: { fontSize: 30, fontFace: "Georgia", color: C.gold, bold: true, breakLine: true } },
    { text: "周二 Tuesday", options: { fontSize: 13, fontFace: "Calibri", color: C.lightGold, breakLine: true } },
  ], { x: 6.4, y: 2.3, w: 1.2, h: 1.4, align: "center" });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 2.5, y: 4.3, w: 4.8, h: 0.5,
    fill: { color: C.primary }, rectRadius: 0.15,
  });
  s.addText("请出示你的护照！Show your passport!", {
    x: 2.5, y: 4.3, w: 4.8, h: 0.5,
    fontSize: 17, fontFace: "Calibri", color: C.dark, align: "center", bold: true,
  });
  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 2 — 护照进度 (1 stamp done)
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "护照签证章进度 Passport Stamps");

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.9, y: 2.55, w: 8.0, h: 0.06, fill: { color: C.primary },
  });

  var stops = [
    { label: "亚洲\nAsia",    x: 1.5,  status: "done" },
    { label: "非洲\nAfrica",   x: 3.5,  status: "current" },
    { label: "欧洲\nEurope",   x: 5.5,  status: "future" },
    { label: "美洲\nAmericas", x: 7.2,  status: "future" },
    { label: "文化展\nExpo",   x: 8.6,  status: "future" },
  ];

  stops.forEach(function(st) {
    var dotColor = st.status === "done" ? C.primary : st.status === "current" ? C.orange : "BDBDBD";
    var dotSize = st.status === "current" ? 0.65 : 0.48;
    var offy = (0.65 - dotSize) / 2;

    s.addShape(pres.shapes.OVAL, {
      x: st.x - dotSize / 2, y: 2.26 + offy, w: dotSize, h: dotSize,
      fill: { color: dotColor },
    });

    if (st.status === "done") {
      s.addText("\u2714", {
        x: st.x - 0.25, y: 2.26 + offy, w: 0.5, h: dotSize,
        fontSize: 18, fontFace: "Calibri", color: C.white, align: "center", valign: "middle",
      });
    }
    if (st.status === "current") {
      s.addText("\u25B6", {
        x: st.x - 0.25, y: 2.26 + offy, w: 0.5, h: dotSize,
        fontSize: 16, fontFace: "Calibri", color: C.white, align: "center", valign: "middle",
      });
    }

    s.addText(st.label, {
      x: st.x - 0.7, y: 2.95, w: 1.4, h: 0.65,
      fontSize: 14, fontFace: "Calibri",
      color: st.status === "done" ? C.primary : st.status === "current" ? C.orange : C.gray,
      align: "center", bold: st.status !== "future",
    });
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 1.2, y: 3.85, w: 7.4, h: 0.7,
    fill: { color: C.cream }, rectRadius: 0.15,
    line: { color: C.primary, width: 2 },
  });
  s.addText("亚洲 \u2714 > 非洲（今天!）> 欧洲 > 美洲 > 文化展", {
    x: 1.3, y: 3.85, w: 7.2, h: 0.7,
    fontSize: 16, fontFace: "Calibri", color: C.accent, align: "center", bold: true, valign: "middle",
  });

  s.addText("集满5个签证章，成为环球探索家！", {
    x: 1.2, y: 4.7, w: 7.4, h: 0.35,
    fontSize: 13, fontFace: "Calibri", color: C.warmGray, align: "center", italic: true,
  });
  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 3 — 今天目标
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "今天的目标 Today\u2019s Goals");

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 1.5, y: 1.1, w: 6.8, h: 3.5,
    fill: { color: C.white }, rectRadius: 0.15,
    line: { color: C.primary, width: 2.5 },
    shadow: { type: "outer", blur: 5, offset: 2, color: "BDBDBD", opacity: 0.3 },
  });

  s.addText([
    { text: "深入了解非洲3个国家", options: { fontSize: 26, fontFace: "Georgia", color: C.primary, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 12, breakLine: true } },
    { text: "埃及 Egypt", options: { fontSize: 20, fontFace: "Georgia", color: C.accent, bold: true, breakLine: true } },
    { text: "肯尼亚 Kenya", options: { fontSize: 20, fontFace: "Georgia", color: C.accent, bold: true, breakLine: true } },
    { text: "南非 South Africa", options: { fontSize: 20, fontFace: "Georgia", color: C.accent, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "每个国家: 概览 > 文化与礼节 > 美食 > 小测验", options: { fontSize: 14, fontFace: "Calibri", color: C.warmGray, breakLine: true } },
  ], { x: 1.8, y: 1.3, w: 6.2, h: 3.1, align: "center", valign: "middle" });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 2.0, y: 4.75, w: 5.8, h: 0.45,
    fill: { color: C.primary }, rectRadius: 0.12,
  });
  s.addText("目标：获得非洲签证章！", {
    x: 2.0, y: 4.75, w: 5.8, h: 0.45,
    fontSize: 16, fontFace: "Calibri", color: C.dark, align: "center", bold: true, valign: "middle",
  });
  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 4 — 认识非洲
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "认识非洲 Meet Africa");

  var stats = [
    { num: "54", label: "个国家\ncountries", sub: "第二大洲", bg: C.lightGold },
    { num: "撒哈拉", label: "沙漠\nSahara", sub: "世界最大沙漠", bg: C.lightRed },
    { num: "尼罗河", label: "Nile\nRiver", sub: "世界最长河流", bg: C.lightBlue },
    { num: "2000+", label: "种语言\nlanguages", sub: "多样性惊人", bg: C.lightPurple },
  ];

  stats.forEach(function(st, i) {
    var x = 0.4 + i * 2.35;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x, y: 1.0, w: 2.15, h: 3.2,
      fill: { color: st.bg }, rectRadius: 0.12,
      shadow: { type: "outer", blur: 3, offset: 2, color: "BDBDBD", opacity: 0.25 },
    });
    s.addText(st.num, {
      x: x + 0.1, y: 1.15, w: 1.95, h: 0.7,
      fontSize: 24, fontFace: "Georgia", color: C.accent, align: "center", bold: true,
    });
    s.addText(st.label, {
      x: x + 0.1, y: 1.9, w: 1.95, h: 0.8,
      fontSize: 14, fontFace: "Calibri", color: C.warmGray, align: "center",
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: x + 0.3, y: 2.8, w: 1.55, h: 0.02, fill: { color: C.primary, transparency: 40 },
    });
    s.addText(st.sub, {
      x: x + 0.1, y: 2.95, w: 1.95, h: 0.55,
      fontSize: 12, fontFace: "Calibri", color: C.accent, align: "center", italic: true,
    });
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 1.0, y: 4.4, w: 7.8, h: 0.55,
    fill: { color: C.cream }, rectRadius: 0.12,
    line: { color: C.primary, width: 1.5 },
  });
  s.addText("非洲不是一个国家，是54个国家！每个都不一样！", {
    x: 1.0, y: 4.4, w: 7.8, h: 0.55,
    fontSize: 15, fontFace: "Calibri", color: C.accent, align: "center", bold: true, valign: "middle",
  });
  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 5 — 埃及概览
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "\uD83C\uDDEA\uD83C\uDDEC 埃及概览 Egypt Overview");

  // Pyramid photo left
  addImageWithFallback(s, IMG.pyramids, 0.3, 0.85, 5.2, 3.5, C.lightGold, "Pyramids of Giza \u2014 金字塔");

  // Info cards right
  var facts = [
    { icon: "\uD83C\uDFF3\uFE0F", label: "国旗", value: "红白黑三色 + 金鹰" },
    { icon: "\uD83D\uDC65", label: "人口", value: "1.04亿" },
    { icon: "\uD83D\uDDE3", label: "语言", value: "阿拉伯语" },
    { icon: "\uD83C\uDFDB", label: "首都", value: "开罗 Cairo" },
    { icon: "\uD83C\uDF0A", label: "地理", value: "尼罗河畔的文明古国" },
  ];

  facts.forEach(function(f, i) {
    var yy = 0.9 + i * 0.72;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.7, y: yy, w: 3.8, h: 0.62,
      fill: { color: i % 2 === 0 ? C.cream : C.white }, rectRadius: 0.08,
    });
    s.addText([
      { text: f.label + "  ", options: { fontSize: 12, fontFace: "Calibri", color: C.warmGray, bold: true } },
      { text: f.value, options: { fontSize: 13, fontFace: "Calibri", color: C.accent } },
    ], { x: 5.85, y: yy, w: 3.5, h: 0.62, valign: "middle" });
  });

  // Flag stripe hint
  s.addShape(pres.shapes.RECTANGLE, { x: 5.7, y: 4.55, w: 1.2, h: 0.22, fill: { color: C.egyptRed } });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.9, y: 4.55, w: 1.3, h: 0.22, fill: { color: C.egyptWhite } });
  s.addShape(pres.shapes.RECTANGLE, { x: 8.2, y: 4.55, w: 1.3, h: 0.22, fill: { color: C.egyptBlack } });

  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 6 — 埃及文化与礼节
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "\uD83C\uDDEA\uD83C\uDDEC 埃及文化与礼节 Culture & Etiquette");

  // Culture section (left)
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.3, y: 0.85, w: 4.6, h: 2.1,
    fill: { color: C.cream }, rectRadius: 0.12,
    line: { color: C.primary, width: 1.5 },
  });
  s.addText([
    { text: "文化 Culture", options: { fontSize: 18, fontFace: "Georgia", color: C.primary, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "金字塔有4500年历史！", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "象形文字 \u2014 古埃及人的文字", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "纸莎草纸 \u2014 最早的「纸」来自埃及", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
  ], { x: 0.5, y: 0.9, w: 4.2, h: 2.0, valign: "top" });

  // Etiquette section (right)
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.1, y: 0.85, w: 4.5, h: 2.1,
    fill: { color: C.lightRed }, rectRadius: 0.12,
    line: { color: C.red, width: 1.5 },
  });
  s.addText([
    { text: "\u26A0\uFE0F 礼节 Etiquette", options: { fontSize: 18, fontFace: "Georgia", color: C.red, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "用右手吃饭、递东西", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "进清真寺要脱鞋", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "不要用左手！", options: { fontSize: 13, fontFace: "Calibri", color: C.red, bold: true, bullet: true, breakLine: true } },
  ], { x: 5.3, y: 0.9, w: 4.1, h: 2.0, valign: "top" });

  // More etiquette at bottom
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.3, y: 3.15, w: 9.3, h: 1.6,
    fill: { color: C.white }, rectRadius: 0.12,
    line: { color: C.primary, width: 1 },
  });
  s.addText([
    { text: "更多礼节 More Etiquette", options: { fontSize: 16, fontFace: "Georgia", color: C.primary, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "见面握手 + 亲脸颊（男性之间也是）", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "斋月期间白天不在公共场合吃东西", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "对长辈说话要特别礼貌", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
  ], { x: 0.5, y: 3.2, w: 8.9, h: 1.5, valign: "top" });

  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 7 — 埃及美食
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "\uD83C\uDDEA\uD83C\uDDEC 埃及美食 Egyptian Food");

  var foods = [
    { zh: "鹰嘴豆泥", en: "Hummus", desc: "用鹰嘴豆做的糊", bg: C.lightGold },
    { zh: "库莎丽", en: "Koshari", desc: "米+面+豆+番茄酱", bg: C.lightRed },
    { zh: "烤肉", en: "Kebab", desc: "串烤羊肉或鸡肉", bg: C.lightBlue },
    { zh: "法拉费尔", en: "Falafel", desc: "炸鹰嘴豆丸子", bg: C.lightGreen },
    { zh: "甜茶", en: "Sweet Tea", desc: "加很多糖的红茶", bg: C.lightPurple },
  ];

  foods.forEach(function(f, i) {
    var x = 0.3 + (i % 3) * 3.15;
    var y = i < 3 ? 0.9 : 3.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x, y: y, w: 2.95, h: 1.8,
      fill: { color: f.bg }, rectRadius: 0.12,
      shadow: { type: "outer", blur: 3, offset: 2, color: "BDBDBD", opacity: 0.2 },
    });
    s.addText(f.zh, {
      x: x + 0.1, y: y + 0.15, w: 2.75, h: 0.5,
      fontSize: 20, fontFace: "Georgia", color: C.accent, bold: true, align: "center",
    });
    s.addText(f.en, {
      x: x + 0.1, y: y + 0.65, w: 2.75, h: 0.4,
      fontSize: 14, fontFace: "Calibri", color: C.primary, align: "center", bold: true,
    });
    s.addText(f.desc, {
      x: x + 0.1, y: y + 1.1, w: 2.75, h: 0.5,
      fontSize: 12, fontFace: "Calibri", color: C.warmGray, align: "center",
    });
  });

  // Eating tip
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 3.45, y: 3.0, w: 6.1, h: 1.8,
    fill: { color: C.cream }, rectRadius: 0.12,
    line: { color: C.primary, width: 2 },
  });
  s.addText([
    { text: "吃饭方式", options: { fontSize: 18, fontFace: "Georgia", color: C.primary, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "用大饼（pita）蘸着吃！", options: { fontSize: 14, fontFace: "Calibri", color: C.accent, breakLine: true } },
    { text: "一定要用右手！", options: { fontSize: 14, fontFace: "Calibri", color: C.red, bold: true, breakLine: true } },
    { text: "和中国用筷子一样，是文化习惯", options: { fontSize: 12, fontFace: "Calibri", color: C.warmGray, italic: true, breakLine: true } },
  ], { x: 3.65, y: 3.05, w: 5.7, h: 1.7, valign: "middle" });

  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 8-9 — 埃及 Check Understanding (2 slides)
// ══════════════════════════════════════════════════════════════
(function() {
  var allQs = [
    { q: "金字塔有多少年历史？", a: "4500年！" },
    { q: "在埃及用什么手吃饭？", a: "右手！不能用左手" },
    { q: "古埃及人写什么文字？", a: "象形文字 Hieroglyphs" },
    { q: "埃及最著名的古迹是什么？", a: "金字塔和狮身人面像" },
    { q: "尼罗河有多长？", a: "6,650公里" },
    { q: "古埃及人用什么文字？", a: "象形文字" },
    { q: "在埃及斋月期间白天能在公共场合吃东西吗？", a: "不能" },
    { q: "埃及人见面除了握手还做什么？", a: "亲脸颊" },
    { q: "纸莎草纸是谁发明的？", a: "古埃及人" },
    { q: "鹰嘴豆泥用什么蘸着吃？", a: "大饼pita" },
  ];

  [0, 1].forEach(function(page) {
    var s = pres.addSlide();
    s.background = { fill: C.secondary };
    titleBar(s, "\u2705 埃及小测验 Egypt Check (" + (page + 1) + "/2)");

    var pageQs = allQs.slice(page * 5, page * 5 + 5);
    pageQs.forEach(function(item, i) {
      var num = page * 5 + i + 1;
      var yy = 1.0 + i * 0.85;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.3, y: yy, w: 9.2, h: 0.72, rectRadius: 0.1,
        fill: { color: i % 2 === 0 ? C.cream : C.white },
        line: { color: C.primary, width: 1.5 },
      });
      s.addShape(pres.shapes.OVAL, {
        x: 0.45, y: yy + 0.13, w: 0.45, h: 0.45,
        fill: { color: C.primary },
      });
      s.addText("" + num, {
        x: 0.45, y: yy + 0.13, w: 0.45, h: 0.45,
        fontSize: 14, fontFace: "Georgia", color: C.gold,
        bold: true, align: "center", valign: "middle",
      });
      s.addText(item.q, {
        x: 1.05, y: yy, w: 4.45, h: 0.72,
        fontSize: 12, fontFace: "Calibri", color: C.accent,
        bold: true, align: "left", valign: "middle",
      });
      s.addText("\u2192 " + item.a, {
        x: 5.6, y: yy, w: 3.75, h: 0.72,
        fontSize: 11, fontFace: "Calibri", color: C.primary,
        bold: true, align: "left", valign: "middle",
      });
    });
    footer(s);
  });
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 9 — 肯尼亚概览
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "\uD83C\uDDF0\uD83C\uDDEA 肯尼亚概览 Kenya Overview");

  // Lion photo left
  addImageWithFallback(s, IMG.lion, 0.3, 0.85, 5.2, 3.5, C.lightGold, "African Lion \u2014 非洲狮子");

  var facts = [
    { label: "国旗", value: "黑红绿 + 盾和矛" },
    { label: "人口", value: "5500万" },
    { label: "语言", value: "斯瓦希里语 + 英语" },
    { label: "首都", value: "内罗毕 Nairobi" },
    { label: "地理", value: "赤道穿过的国家！" },
  ];

  facts.forEach(function(f, i) {
    var yy = 0.9 + i * 0.72;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.7, y: yy, w: 3.8, h: 0.62,
      fill: { color: i % 2 === 0 ? C.cream : C.white }, rectRadius: 0.08,
    });
    s.addText([
      { text: f.label + "  ", options: { fontSize: 12, fontFace: "Calibri", color: C.warmGray, bold: true } },
      { text: f.value, options: { fontSize: 13, fontFace: "Calibri", color: C.accent } },
    ], { x: 5.85, y: yy, w: 3.5, h: 0.62, valign: "middle" });
  });

  // Flag stripe
  s.addShape(pres.shapes.RECTANGLE, { x: 5.7, y: 4.55, w: 1.2, h: 0.22, fill: { color: C.kenyaBlack } });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.9, y: 4.55, w: 1.3, h: 0.22, fill: { color: C.kenyaRed } });
  s.addShape(pres.shapes.RECTANGLE, { x: 8.2, y: 4.55, w: 1.3, h: 0.22, fill: { color: C.kenyaGreen } });

  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 10 — 肯尼亚文化与礼节
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "\uD83C\uDDF0\uD83C\uDDEA 肯尼亚文化与礼节 Culture & Etiquette");

  // Culture left
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.3, y: 0.85, w: 4.6, h: 2.1,
    fill: { color: C.cream }, rectRadius: 0.12,
    line: { color: C.primary, width: 1.5 },
  });
  s.addText([
    { text: "文化 Culture", options: { fontSize: 18, fontFace: "Georgia", color: C.primary, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "动物大迁徙 \u2014 百万动物穿越草原", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "马赛族 Maasai \u2014 著名的跳跃舞", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "非洲五大动物: 狮子/象/犀牛/豹/水牛", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
  ], { x: 0.5, y: 0.9, w: 4.2, h: 2.0, valign: "top" });

  // Etiquette right
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.1, y: 0.85, w: 4.5, h: 2.1,
    fill: { color: C.lightRed }, rectRadius: 0.12,
    line: { color: C.red, width: 1.5 },
  });
  s.addText([
    { text: "\u26A0\uFE0F 礼节 Etiquette", options: { fontSize: 18, fontFace: "Georgia", color: C.red, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "打招呼说 Jambo!", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "握手很重要，时间长=尊重", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "指人用整只手，不用手指", options: { fontSize: 13, fontFace: "Calibri", color: C.red, bold: true, bullet: true, breakLine: true } },
  ], { x: 5.3, y: 0.9, w: 4.1, h: 2.0, valign: "top" });

  // More etiquette bottom
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.3, y: 3.15, w: 9.3, h: 1.6,
    fill: { color: C.white }, rectRadius: 0.12,
    line: { color: C.primary, width: 1 },
  });
  s.addText([
    { text: "更多礼节 More Etiquette", options: { fontSize: 16, fontFace: "Georgia", color: C.primary, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "对长辈特别尊敬 \u2014 要用双手接东西", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "拍动物照片要小心距离 \u2014 Safari安全第一！", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "马赛族的红色是勇气的象征", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
  ], { x: 0.5, y: 3.2, w: 8.9, h: 1.5, valign: "top" });

  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 11 — 肯尼亚美食
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "\uD83C\uDDF0\uD83C\uDDEA 肯尼亚美食 Kenyan Food");

  var foods = [
    { zh: "乌伽黎", en: "Ugali", desc: "玉米糊，主食", bg: C.lightGold },
    { zh: "烤肉", en: "Nyama Choma", desc: "炭烤肉，最爱！", bg: C.lightRed },
    { zh: "恰帕提", en: "Chapati", desc: "印度传来的煎饼", bg: C.lightBlue },
    { zh: "肯尼亚茶", en: "Kenyan Tea", desc: "加奶的红茶", bg: C.lightGreen },
  ];

  foods.forEach(function(f, i) {
    var x = 0.3 + i * 2.35;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x, y: 0.9, w: 2.15, h: 2.0,
      fill: { color: f.bg }, rectRadius: 0.12,
      shadow: { type: "outer", blur: 3, offset: 2, color: "BDBDBD", opacity: 0.2 },
    });
    s.addText(f.zh, {
      x: x + 0.1, y: 0.95, w: 1.95, h: 0.55,
      fontSize: 20, fontFace: "Georgia", color: C.accent, bold: true, align: "center",
    });
    s.addText(f.en, {
      x: x + 0.1, y: 1.5, w: 1.95, h: 0.4,
      fontSize: 14, fontFace: "Calibri", color: C.primary, align: "center", bold: true,
    });
    s.addText(f.desc, {
      x: x + 0.1, y: 1.95, w: 1.95, h: 0.55,
      fontSize: 12, fontFace: "Calibri", color: C.warmGray, align: "center",
    });
  });

  // Eating method
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.5, y: 3.2, w: 9.0, h: 1.65,
    fill: { color: C.cream }, rectRadius: 0.12,
    line: { color: C.primary, width: 2 },
  });
  s.addText([
    { text: "怎么吃 Ugali？", options: { fontSize: 18, fontFace: "Georgia", color: C.primary, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "1. 用右手揪一小团 ugali", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, breakLine: true } },
    { text: "2. 用手指捏成小碗形状", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, breakLine: true } },
    { text: "3. 舀起菜或肉一起吃！", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, breakLine: true } },
    { text: "和中国人用馒头夹菜有点像！", options: { fontSize: 13, fontFace: "Calibri", color: C.primary, bold: true, italic: true, breakLine: true } },
  ], { x: 0.7, y: 3.25, w: 8.6, h: 1.55, valign: "top" });

  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 12-13 — 肯尼亚 Check (2 slides)
// ══════════════════════════════════════════════════════════════
(function() {
  var allQs = [
    { q: "肯尼亚说什么语言？", a: "斯瓦希里语 + 英语" },
    { q: "非洲五大动物是哪五大？", a: "狮子、大象、犀牛、豹子、水牛" },
    { q: "Jambo 是什么意思？", a: "你好！Hello!（斯瓦希里语）" },
    { q: "动物大迁徙发生在哪里？", a: "肯尼亚和坦桑尼亚" },
    { q: "Jambo是什么语言？", a: "斯瓦希里语" },
    { q: "肯尼亚的赤道穿过吗？", a: "是的！" },
    { q: "马赛族以什么舞蹈出名？", a: "跳跃舞" },
    { q: "用手指指人在肯尼亚礼貌吗？", a: "不礼貌，要用整只手" },
    { q: "乌伽黎(ugali)是用什么做的？", a: "玉米" },
    { q: "动物大迁徙有多少只动物？", a: "约200万" },
  ];

  [0, 1].forEach(function(page) {
    var s = pres.addSlide();
    s.background = { fill: C.secondary };
    titleBar(s, "\u2705 肯尼亚小测验 Kenya Check (" + (page + 1) + "/2)");

    var pageQs = allQs.slice(page * 5, page * 5 + 5);
    pageQs.forEach(function(item, i) {
      var num = page * 5 + i + 1;
      var yy = 1.0 + i * 0.85;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.3, y: yy, w: 9.2, h: 0.72, rectRadius: 0.1,
        fill: { color: i % 2 === 0 ? C.cream : C.white },
        line: { color: C.primary, width: 1.5 },
      });
      s.addShape(pres.shapes.OVAL, {
        x: 0.45, y: yy + 0.13, w: 0.45, h: 0.45,
        fill: { color: C.primary },
      });
      s.addText("" + num, {
        x: 0.45, y: yy + 0.13, w: 0.45, h: 0.45,
        fontSize: 14, fontFace: "Georgia", color: C.gold,
        bold: true, align: "center", valign: "middle",
      });
      s.addText(item.q, {
        x: 1.05, y: yy, w: 4.45, h: 0.72,
        fontSize: 12, fontFace: "Calibri", color: C.accent,
        bold: true, align: "left", valign: "middle",
      });
      s.addText("\u2192 " + item.a, {
        x: 5.6, y: yy, w: 3.75, h: 0.72,
        fontSize: 11, fontFace: "Calibri", color: C.primary,
        bold: true, align: "left", valign: "middle",
      });
    });
    footer(s);
  });
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 13 — 南非概览
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "\uD83C\uDDFF\uD83C\uDDE6 南非概览 South Africa Overview");

  // Table Mountain photo left
  addImageWithFallback(s, IMG.tableMountain, 0.3, 0.85, 5.2, 3.5, C.lightGreen, "Table Mountain \u2014 桌山");

  var facts = [
    { label: "国旗", value: "六色彩虹旗！" },
    { label: "人口", value: "6000万" },
    { label: "语言", value: "11种官方语言!" },
    { label: "首都", value: "3个首都！" },
    { label: "别称", value: "彩虹之国 Rainbow Nation" },
  ];

  facts.forEach(function(f, i) {
    var yy = 0.9 + i * 0.72;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.7, y: yy, w: 3.8, h: 0.62,
      fill: { color: i % 2 === 0 ? C.cream : C.white }, rectRadius: 0.08,
    });
    s.addText([
      { text: f.label + "  ", options: { fontSize: 12, fontFace: "Calibri", color: C.warmGray, bold: true } },
      { text: f.value, options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bold: i === 2 || i === 3 } },
    ], { x: 5.85, y: yy, w: 3.5, h: 0.62, valign: "middle" });
  });

  // SA flag colors
  var flagColors = [C.saRed, C.saBlue, C.saGreen, C.saGold, C.egyptBlack, C.egyptWhite];
  flagColors.forEach(function(clr, i) {
    s.addShape(pres.shapes.RECTANGLE, {
      x: 5.7 + i * 0.63, y: 4.55, w: 0.63, h: 0.22, fill: { color: clr },
    });
  });

  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 14 — 南非文化与礼节
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "\uD83C\uDDFF\uD83C\uDDE6 南非文化与礼节 Culture & Etiquette");

  // Culture left
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.3, y: 0.85, w: 4.6, h: 2.1,
    fill: { color: C.cream }, rectRadius: 0.12,
    line: { color: C.primary, width: 1.5 },
  });
  s.addText([
    { text: "文化 Culture", options: { fontSize: 18, fontFace: "Georgia", color: C.primary, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "曼德拉 Mandela \u2014 种族和解英雄", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "11种语言 = 多元文化共存", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "3个首都: 行政/立法/司法", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
  ], { x: 0.5, y: 0.9, w: 4.2, h: 2.0, valign: "top" });

  // Etiquette right
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.1, y: 0.85, w: 4.5, h: 2.1,
    fill: { color: C.lightRed }, rectRadius: 0.12,
    line: { color: C.red, width: 1.5 },
  });
  s.addText([
    { text: "\u26A0\uFE0F 礼节 Etiquette", options: { fontSize: 18, fontFace: "Georgia", color: C.red, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "祖鲁语说 Sawubona（我看见你了）", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "特别的三步握手！", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "尊重不同文化背景", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
  ], { x: 5.3, y: 0.9, w: 4.1, h: 2.0, valign: "top" });

  // Bottom
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.3, y: 3.15, w: 9.3, h: 1.6,
    fill: { color: C.white }, rectRadius: 0.12,
    line: { color: C.primary, width: 1 },
  });
  s.addText([
    { text: "更多礼节 More Etiquette", options: { fontSize: 16, fontFace: "Georgia", color: C.primary, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "Braai（烤肉）是最重要的社交活动", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "南非人非常友善好客", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "Ubuntu精神: 我的存在因为有你", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
  ], { x: 0.5, y: 3.2, w: 8.9, h: 1.5, valign: "top" });

  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 15 — 南非美食
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "\uD83C\uDDFF\uD83C\uDDE6 南非美食 South African Food");

  var foods = [
    { zh: "烤肉", en: "Braai", desc: "南非国民烧烤", bg: C.lightGold },
    { zh: "干肉条", en: "Biltong", desc: "风干腌制肉条", bg: C.lightRed },
    { zh: "玉米粥", en: "Pap", desc: "类似ugali的主食", bg: C.lightBlue },
    { zh: "波波提", en: "Bobotie", desc: "咖喱肉派，国菜！", bg: C.lightGreen },
    { zh: "路易波士茶", en: "Rooibos", desc: "南非特有红灌木茶", bg: C.lightPurple },
  ];

  foods.forEach(function(f, i) {
    var x = 0.3 + (i % 3) * 3.15;
    var y = i < 3 ? 0.9 : 3.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x, y: y, w: 2.95, h: 1.8,
      fill: { color: f.bg }, rectRadius: 0.12,
      shadow: { type: "outer", blur: 3, offset: 2, color: "BDBDBD", opacity: 0.2 },
    });
    s.addText(f.zh, {
      x: x + 0.1, y: y + 0.15, w: 2.75, h: 0.5,
      fontSize: 20, fontFace: "Georgia", color: C.accent, bold: true, align: "center",
    });
    s.addText(f.en, {
      x: x + 0.1, y: y + 0.65, w: 2.75, h: 0.4,
      fontSize: 14, fontFace: "Calibri", color: C.primary, align: "center", bold: true,
    });
    s.addText(f.desc, {
      x: x + 0.1, y: y + 1.1, w: 2.75, h: 0.5,
      fontSize: 12, fontFace: "Calibri", color: C.warmGray, align: "center",
    });
  });

  // Braai culture
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 3.45, y: 3.0, w: 6.1, h: 1.8,
    fill: { color: C.cream }, rectRadius: 0.12,
    line: { color: C.primary, width: 2 },
  });
  s.addText([
    { text: "Braai 文化", options: { fontSize: 18, fontFace: "Georgia", color: C.primary, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "Braai 就像中国的聚餐BBQ！", options: { fontSize: 14, fontFace: "Calibri", color: C.accent, breakLine: true } },
    { text: "朋友家人围在一起烤肉聊天", options: { fontSize: 13, fontFace: "Calibri", color: C.warmGray, breakLine: true } },
    { text: "南非甚至有「National Braai Day」！", options: { fontSize: 13, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
  ], { x: 3.65, y: 3.05, w: 5.7, h: 1.7, valign: "middle" });

  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 16-17 — 南非 Check (2 slides)
// ══════════════════════════════════════════════════════════════
(function() {
  var allQs = [
    { q: "南非有几种官方语言？", a: "11种！" },
    { q: "南非有几个首都？", a: "3个！（行政/立法/司法）" },
    { q: "Sawubona 是什么意思？", a: "「我看见你了」I see you!（祖鲁语）" },
    { q: "南非最有名的山叫什么？", a: "桌山 Table Mountain" },
    { q: "南非的国旗有几种颜色？", a: "6种" },
    { q: "Nelson Mandela为什么出名？", a: "反对种族歧视/争取平等" },
    { q: "Sawubona是什么意思？", a: "我看见你了" },
    { q: "南非braai和什么活动很像？", a: "BBQ烧烤聚会" },
    { q: "南非为什么叫彩虹之国？", a: "因为有很多不同文化的人" },
    { q: "路易波士茶(rooibos)来自哪里？", a: "南非" },
  ];

  [0, 1].forEach(function(page) {
    var s = pres.addSlide();
    s.background = { fill: C.secondary };
    titleBar(s, "\u2705 南非小测验 South Africa Check (" + (page + 1) + "/2)");

    var pageQs = allQs.slice(page * 5, page * 5 + 5);
    pageQs.forEach(function(item, i) {
      var num = page * 5 + i + 1;
      var yy = 1.0 + i * 0.85;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.3, y: yy, w: 9.2, h: 0.72, rectRadius: 0.1,
        fill: { color: i % 2 === 0 ? C.cream : C.white },
        line: { color: C.primary, width: 1.5 },
      });
      s.addShape(pres.shapes.OVAL, {
        x: 0.45, y: yy + 0.13, w: 0.45, h: 0.45,
        fill: { color: C.primary },
      });
      s.addText("" + num, {
        x: 0.45, y: yy + 0.13, w: 0.45, h: 0.45,
        fontSize: 14, fontFace: "Georgia", color: C.gold,
        bold: true, align: "center", valign: "middle",
      });
      s.addText(item.q, {
        x: 1.05, y: yy, w: 4.45, h: 0.72,
        fontSize: 12, fontFace: "Calibri", color: C.accent,
        bold: true, align: "left", valign: "middle",
      });
      s.addText("\u2192 " + item.a, {
        x: 5.6, y: yy, w: 3.75, h: 0.72,
        fontSize: 11, fontFace: "Calibri", color: C.primary,
        bold: true, align: "left", valign: "middle",
      });
    });
    footer(s);
  });
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 17 — Mini Role Play 三种打招呼
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "\uD83C\uDFAD Mini Role Play \u2014 三种打招呼");

  var greetings = [
    {
      country: "埃及 Egypt",
      how: "握手 + 亲脸颊",
      detail: "先握手，然后\n轻碰双颊",
      bg: C.lightGold,
      lineClr: C.primary,
    },
    {
      country: "肯尼亚 Kenya",
      how: "Jambo! + 长握手",
      detail: "说 Jambo!\n握手时间长=尊重",
      bg: C.lightGreen,
      lineClr: C.green,
    },
    {
      country: "南非 South Africa",
      how: "Sawubona + 三步握手",
      detail: "说 Sawubona!\n普通握 > 拇指扣 > 普通握",
      bg: C.lightBlue,
      lineClr: C.blue,
    },
  ];

  greetings.forEach(function(g, i) {
    var x = 0.3 + i * 3.15;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x, y: 0.85, w: 2.95, h: 3.5,
      fill: { color: g.bg }, rectRadius: 0.12,
      line: { color: g.lineClr, width: 2 },
      shadow: { type: "outer", blur: 4, offset: 2, color: "BDBDBD", opacity: 0.25 },
    });
    s.addText(g.country, {
      x: x + 0.1, y: 0.95, w: 2.75, h: 0.5,
      fontSize: 18, fontFace: "Georgia", color: C.accent, bold: true, align: "center",
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: x + 0.4, y: 1.5, w: 2.15, h: 0.02, fill: { color: g.lineClr },
    });
    s.addText(g.how, {
      x: x + 0.1, y: 1.6, w: 2.75, h: 0.7,
      fontSize: 16, fontFace: "Calibri", color: C.primary, bold: true, align: "center",
    });
    s.addText(g.detail, {
      x: x + 0.15, y: 2.4, w: 2.65, h: 1.5,
      fontSize: 13, fontFace: "Calibri", color: C.warmGray, align: "center",
    });
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 1.0, y: 4.5, w: 7.8, h: 0.5,
    fill: { color: C.primary }, rectRadius: 0.12,
  });
  s.addText("两人一组，练习3种打招呼方式！", {
    x: 1.0, y: 4.5, w: 7.8, h: 0.5,
    fontSize: 16, fontFace: "Calibri", color: C.dark, align: "center", bold: true, valign: "middle",
  });
  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 18 — 上午竞赛
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "\uD83C\uDFC6 上午竞赛 Morning Quiz");

  var questions = [
    { q: "Q1: 金字塔有多少年历史？", a: "4500年" },
    { q: "Q2: 在埃及和肯尼亚吃饭用什么手？", a: "右手" },
    { q: "Q3: 肯尼亚语说「你好」怎么说？", a: "Jambo!" },
    { q: "Q4: 南非有几种官方语言？", a: "11种" },
    { q: "Q5: Sawubona 是什么意思？", a: "我看见你了" },
    { q: "Q6: 哪个国家有3个首都？", a: "南非" },
  ];

  questions.forEach(function(item, i) {
    var col = i < 3 ? 0 : 1;
    var row = i % 3;
    var x = 0.4 + col * 4.8;
    var y = 0.95 + row * 1.45;

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x, y: y, w: 4.5, h: 1.25,
      fill: { color: i % 2 === 0 ? C.cream : C.white }, rectRadius: 0.1,
      line: { color: C.primary, width: 1 },
    });
    s.addText([
      { text: item.q, options: { fontSize: 14, fontFace: "Calibri", color: C.accent, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 4, breakLine: true } },
      { text: "答：" + item.a, options: { fontSize: 13, fontFace: "Calibri", color: C.primary, breakLine: true } },
    ], { x: x + 0.15, y: y + 0.05, w: 4.2, h: 1.15, valign: "middle" });
  });

  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 19 — Project 提醒
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "\uD83E\uDDE9 Project 提醒 Reminder");

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 1.0, y: 1.0, w: 7.8, h: 3.5,
    fill: { color: C.cream }, rectRadius: 0.18,
    line: { color: C.primary, width: 2.5 },
    shadow: { type: "outer", blur: 6, offset: 3, color: "BDBDBD", opacity: 0.3 },
  });

  s.addText([
    { text: "下午任务预告", options: { fontSize: 22, fontFace: "Georgia", color: C.accent, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "1. 三国文化对比 \u2014 深入比较", options: { fontSize: 15, fontFace: "Calibri", color: C.accent, breakLine: true } },
    { text: "2. 生词和句型练习", options: { fontSize: 15, fontFace: "Calibri", color: C.accent, breakLine: true } },
    { text: "3. 角色扮演: Safari + 非洲餐厅", options: { fontSize: 15, fontFace: "Calibri", color: C.accent, breakLine: true } },
    { text: "4. Project Time: 完成护照非洲页", options: { fontSize: 15, fontFace: "Calibri", color: C.accent, breakLine: true } },
    { text: "5. 分享 + 获得签证章！", options: { fontSize: 15, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
  ], { x: 1.3, y: 1.2, w: 7.2, h: 3.1, valign: "middle" });

  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 20 — 下午开始
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.dark };
  goldBars(s);

  s.addText("下午开始", {
    x: 0.5, y: 1.4, w: 9.0, h: 0.85,
    fontSize: 36, fontFace: "Georgia", color: C.gold, align: "center", bold: true,
  });
  s.addText("Afternoon Session", {
    x: 0.5, y: 2.3, w: 9.0, h: 0.55,
    fontSize: 22, fontFace: "Calibri", color: C.lightGold, align: "center",
  });
  s.addText("深入比较 + 句型 + 角色扮演 + Project", {
    x: 1.5, y: 3.2, w: 6.8, h: 0.45,
    fontSize: 16, fontFace: "Calibri", color: C.warmGray, align: "center",
  });
  footer(s, "谷雨中文 GR EDU  |  Global Explorer Camp");
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 21 — 快速复习
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "快速复习 Quick Review");

  var reviewQ = [
    "埃及人用什么手吃饭？",
    "肯尼亚用斯瓦希里语怎么说「你好」？",
    "南非有几种官方语言？",
    "Sawubona 是什么意思？",
    "哪个国家有金字塔？",
  ];

  reviewQ.forEach(function(q, i) {
    var y = 0.95 + i * 0.82;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 1.0, y: y, w: 7.8, h: 0.68,
      fill: { color: i % 2 === 0 ? C.lightGold : C.white }, rectRadius: 0.1,
    });
    s.addText((i + 1) + ". " + q, {
      x: 1.2, y: y, w: 7.4, h: 0.68,
      fontSize: 17, fontFace: "Calibri", color: C.accent, bold: true, valign: "middle",
    });
  });

  s.addText("举手回答，答对加分！", {
    x: 1.0, y: 5.0, w: 7.8, h: 0.3,
    fontSize: 12, fontFace: "Calibri", color: C.warmGray, align: "center", italic: true,
  });
  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 22 — 非洲三国文化对比表
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "\uD83C\uDF0D 非洲三国文化对比表");

  var tableRows = [
    { cat: "", c1: "埃及", c2: "肯尼亚", c3: "南非", isHeader: true },
    { cat: "打招呼", c1: "握手+亲脸颊", c2: "Jambo!+长握手", c3: "Sawubona+三步握手" },
    { cat: "吃饭方式", c1: "用饼蘸+右手", c2: "用手吃ugali", c3: "braai烤肉聚餐" },
    { cat: "重要文化", c1: "金字塔/尼罗河", c2: "动物大迁徙/马赛族", c3: "彩虹之国/曼德拉" },
    { cat: "代表食物", c1: "鹰嘴豆泥", c2: "乌伽黎", c3: "烤肉braai" },
    { cat: "注意事项", c1: "用右手！", c2: "别用手指指人", c3: "尊重多元文化" },
  ];

  var colWidths = [1.6, 2.4, 2.4, 2.4];
  var startX = 0.5;
  var startY = 0.85;
  var rowH = 0.72;

  tableRows.forEach(function(row, ri) {
    var y = startY + ri * rowH;
    var cells = [row.cat, row.c1, row.c2, row.c3];
    var cx = startX;

    cells.forEach(function(cell, ci) {
      var bgClr;
      if (ri === 0) bgClr = C.accent;
      else if (ci === 0) bgClr = C.lightGold;
      else bgClr = ri % 2 === 0 ? C.cream : C.white;

      s.addShape(pres.shapes.RECTANGLE, {
        x: cx, y: y, w: colWidths[ci], h: rowH,
        fill: { color: bgClr },
        line: { color: C.primary, width: 0.5 },
      });

      var txtColor = ri === 0 ? C.gold : ci === 0 ? C.accent : C.darkText;
      var isBold = ri === 0 || ci === 0;
      s.addText(cell, {
        x: cx + 0.08, y: y, w: colWidths[ci] - 0.16, h: rowH,
        fontSize: ri === 0 ? 14 : 12, fontFace: ri === 0 ? "Georgia" : "Calibri",
        color: txtColor, bold: isBold, align: "center", valign: "middle",
      });
      cx += colWidths[ci];
    });
  });

  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 23 — 共同点与不同
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "共同点与不同 Similarities & Differences");

  // Similarities
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.3, y: 0.85, w: 4.5, h: 2.8,
    fill: { color: C.lightGreen }, rectRadius: 0.12,
    line: { color: C.green, width: 2 },
  });
  s.addText([
    { text: "共同点 Similarities", options: { fontSize: 18, fontFace: "Georgia", color: C.green, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "都很重视社交和分享食物", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "都对客人很热情好客", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "握手是重要礼节", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "都有丰富的历史", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
  ], { x: 0.5, y: 0.9, w: 4.1, h: 2.7, valign: "top" });

  // Differences
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.0, y: 0.85, w: 4.5, h: 2.8,
    fill: { color: C.lightRed }, rectRadius: 0.12,
    line: { color: C.red, width: 2 },
  });
  s.addText([
    { text: "不同点 Differences", options: { fontSize: 18, fontFace: "Georgia", color: C.red, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "语言完全不同", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "食物风格不同", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "地理环境差异大", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, bullet: true, breakLine: true } },
    { text: "（沙漠 vs 草原 vs 海边）", options: { fontSize: 12, fontFace: "Calibri", color: C.warmGray, italic: true, breakLine: true } },
  ], { x: 5.2, y: 0.9, w: 4.1, h: 2.7, valign: "top" });

  // Key message
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.8, y: 3.85, w: 8.2, h: 0.95,
    fill: { color: C.cream }, rectRadius: 0.15,
    line: { color: C.primary, width: 2.5 },
  });
  s.addText([
    { text: "非洲不是一个国家，是54个国家！", options: { fontSize: 18, fontFace: "Georgia", color: C.primary, bold: true, breakLine: true } },
    { text: "每个都不一样！Each one is unique!", options: { fontSize: 14, fontFace: "Calibri", color: C.warmGray, breakLine: true } },
  ], { x: 1.0, y: 3.9, w: 7.8, h: 0.85, align: "center", valign: "middle" });

  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 24 — 旅行小贴士
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "\uD83E\uDDF3 旅行小贴士 Travel Tips");

  var tips = [
    {
      country: "去埃及",
      tips: "用右手！\n尊重伊斯兰文化\n进清真寺脱鞋",
      bg: C.lightGold, lineClr: C.primary,
    },
    {
      country: "去肯尼亚",
      tips: "Safari保持距离！\n说Jambo打招呼\n长握手表尊重",
      bg: C.lightGreen, lineClr: C.green,
    },
    {
      country: "去南非",
      tips: "尊重多元文化\n说Sawubona\n享受braai烤肉！",
      bg: C.lightBlue, lineClr: C.blue,
    },
  ];

  tips.forEach(function(t, i) {
    var x = 0.3 + i * 3.15;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x, y: 0.85, w: 2.95, h: 3.5,
      fill: { color: t.bg }, rectRadius: 0.12,
      line: { color: t.lineClr, width: 2 },
      shadow: { type: "outer", blur: 4, offset: 2, color: "BDBDBD", opacity: 0.25 },
    });
    s.addText(t.country, {
      x: x + 0.1, y: 0.95, w: 2.75, h: 0.55,
      fontSize: 20, fontFace: "Georgia", color: C.accent, bold: true, align: "center",
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: x + 0.4, y: 1.55, w: 2.15, h: 0.02, fill: { color: t.lineClr },
    });
    s.addText(t.tips, {
      x: x + 0.2, y: 1.7, w: 2.55, h: 2.2,
      fontSize: 14, fontFace: "Calibri", color: C.accent, align: "center", valign: "top",
    });
  });

  s.addText("尊重当地文化是最重要的旅行礼节！", {
    x: 0.5, y: 4.55, w: 9.0, h: 0.4,
    fontSize: 14, fontFace: "Calibri", color: C.warmGray, align: "center", italic: true,
  });
  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 25 — 生词卡
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "生词卡 Vocabulary Cards");

  var vocab = [
    { zh: "金字塔", py: "jinzita",      en: "Pyramid" },
    { zh: "沙漠",   py: "shamo",        en: "Desert" },
    { zh: "草原",   py: "caoyuan",      en: "Grassland" },
    { zh: "狮子",   py: "shizi",        en: "Lion" },
    { zh: "大象",   py: "daxiang",      en: "Elephant" },
    { zh: "烤肉",   py: "kaoru",        en: "BBQ / Braai" },
    { zh: "握手",   py: "woshou",       en: "Handshake" },
    { zh: "礼节",   py: "lijie",        en: "Etiquette" },
    { zh: "打招呼", py: "da zhaohu",    en: "Greet" },
    { zh: "尊重",   py: "zunzhong",     en: "Respect" },
  ];

  var bgColors = [C.lightGold, C.lightGreen, C.lightRed, C.lightBlue, C.lightPurple];

  vocab.forEach(function(v, i) {
    var col = i < 5 ? 0 : 1;
    var row = i % 5;
    var x = 0.4 + col * 4.8;
    var y = 0.88 + row * 0.88;

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x, y: y, w: 4.5, h: 0.78,
      fill: { color: bgColors[row] }, rectRadius: 0.1,
    });
    s.addText(v.zh, {
      x: x + 0.15, y: y, w: 1.3, h: 0.78,
      fontSize: 22, fontFace: "Georgia", color: C.accent, bold: true, valign: "middle",
    });
    s.addText(v.py, {
      x: x + 1.5, y: y, w: 1.5, h: 0.78,
      fontSize: 12, fontFace: "Calibri", color: C.warmGray, valign: "middle", italic: true,
    });
    s.addText(v.en, {
      x: x + 3.0, y: y, w: 1.3, h: 0.78,
      fontSize: 14, fontFace: "Calibri", color: C.primary, valign: "middle", bold: true,
    });
  });

  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 26 — 句型练习
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "句型练习 Sentence Patterns");

  // Pattern 1
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 0.9, w: 9.0, h: 1.3,
    fill: { color: C.lightGold }, rectRadius: 0.12,
  });
  s.addText([
    { text: "句型1：「在___，人们用___打招呼」", options: { fontSize: 18, fontFace: "Georgia", color: C.accent, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "例：「在肯尼亚，人们用Jambo打招呼」", options: { fontSize: 14, fontFace: "Calibri", color: C.primary, breakLine: true } },
    { text: "Example: In Kenya, people greet each other with Jambo", options: { fontSize: 11, fontFace: "Calibri", color: C.warmGray, italic: true, breakLine: true } },
  ], { x: 0.6, y: 0.95, w: 8.6, h: 1.2, valign: "middle" });

  // Pattern 2
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 2.4, w: 9.0, h: 1.3,
    fill: { color: C.lightGreen }, rectRadius: 0.12,
  });
  s.addText([
    { text: "句型2：「去___旅行要注意___」", options: { fontSize: 18, fontFace: "Georgia", color: C.accent, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "例：「去埃及旅行要注意用右手吃饭」", options: { fontSize: 14, fontFace: "Calibri", color: C.green, breakLine: true } },
    { text: "Example: When traveling to Egypt, remember to eat with your right hand", options: { fontSize: 11, fontFace: "Calibri", color: C.warmGray, italic: true, breakLine: true } },
  ], { x: 0.6, y: 2.45, w: 8.6, h: 1.2, valign: "middle" });

  // Fill in
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 3.9, w: 9.0, h: 1.1,
    fill: { color: C.cream }, rectRadius: 0.12,
    line: { color: C.primary, width: 1.5 },
  });
  s.addText([
    { text: "你来试试！Your turn!", options: { fontSize: 16, fontFace: "Georgia", color: C.accent, bold: true, breakLine: true } },
    { text: "「在___，人们用___打招呼」    「去___旅行要注意___」", options: { fontSize: 14, fontFace: "Calibri", color: C.primary, breakLine: true } },
  ], { x: 0.6, y: 3.95, w: 8.6, h: 1.0, align: "center", valign: "middle" });

  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 27 — Role Play Safari + 餐厅
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "\uD83C\uDFAD Role Play \u2014 Safari + 非洲餐厅");

  // Safari scenario
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.3, y: 0.85, w: 4.5, h: 3.3,
    fill: { color: C.lightGold }, rectRadius: 0.12,
    shadow: { type: "outer", blur: 4, offset: 2, color: "BDBDBD", opacity: 0.25 },
  });
  s.addText("Safari \u573A\u666F", {
    x: 0.4, y: 0.9, w: 4.3, h: 0.45,
    fontSize: 18, fontFace: "Georgia", color: C.accent, bold: true, align: "center",
  });
  s.addText([
    { text: "A = 游客 Tourist", options: { fontSize: 13, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
    { text: "B = 导游 Guide", options: { fontSize: 13, fontFace: "Calibri", color: C.green, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "A：「Jambo! 这是什么动物？」", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, breakLine: true } },
    { text: "B：「这是狮子，它住在大草原」", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, breakLine: true } },
    { text: "A：「狮子危险吗？」", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, breakLine: true } },
    { text: "B：「要保持距离！」", options: { fontSize: 13, fontFace: "Calibri", color: C.red, bold: true, breakLine: true } },
  ], { x: 0.5, y: 1.4, w: 4.1, h: 2.6, valign: "top" });

  // Restaurant scenario
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.0, y: 0.85, w: 4.5, h: 3.3,
    fill: { color: C.lightGreen }, rectRadius: 0.12,
    shadow: { type: "outer", blur: 4, offset: 2, color: "BDBDBD", opacity: 0.25 },
  });
  s.addText("\u975E\u6D32\u9910\u5385 Restaurant", {
    x: 5.1, y: 0.9, w: 4.3, h: 0.45,
    fontSize: 18, fontFace: "Georgia", color: C.accent, bold: true, align: "center",
  });
  s.addText([
    { text: "A = 客人 Guest", options: { fontSize: 13, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
    { text: "B = 服务员 Waiter", options: { fontSize: 13, fontFace: "Calibri", color: C.green, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "A：「我想吃___」", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, breakLine: true } },
    { text: "B：「好的，要用手吃哦」", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, breakLine: true } },
    { text: "A：「用右手对吗？」", options: { fontSize: 13, fontFace: "Calibri", color: C.accent, breakLine: true } },
    { text: "B：「对！右手！」", options: { fontSize: 13, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
  ], { x: 5.2, y: 1.4, w: 4.1, h: 2.6, valign: "top" });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 1.0, y: 4.35, w: 7.8, h: 0.55,
    fill: { color: C.cream }, rectRadius: 0.1,
    line: { color: C.primary, width: 1 },
  });
  s.addText("两人一组，选一个场景练习！", {
    x: 1.0, y: 4.35, w: 7.8, h: 0.55,
    fontSize: 14, fontFace: "Calibri", color: C.accent, align: "center", bold: true, valign: "middle",
  });
  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 28 — Project Time
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "Project Time! \u62A4\u7167\u975E\u6D32\u9875");

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.5, y: 0.85, w: 9.0, h: 4.1,
    fill: { color: C.cream }, rectRadius: 0.15,
    line: { color: C.primary, width: 2.5 },
    shadow: { type: "outer", blur: 6, offset: 3, color: "BDBDBD", opacity: 0.3 },
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 0.85, w: 9.0, h: 0.52,
    fill: { color: C.primary },
  });
  s.addText("AFRICA \u975E\u6D32 \u2014 My Passport Page \u6211\u7684\u62A4\u7167\u9875", {
    x: 0.6, y: 0.85, w: 8.8, h: 0.52,
    fontSize: 18, fontFace: "Georgia", color: C.dark, bold: true, align: "center", valign: "middle",
  });

  var fields = [
    { label: "\u6211\u53BB\u4E86\u54EA\u91CC Where I went", y: 1.5 },
    { label: "\u6211\u770B\u5230\u4E86\u4EC0\u4E48 What I saw", y: 2.15 },
    { label: "\u6211\u5403\u4E86\u4EC0\u4E48 What I ate", y: 2.8 },
    { label: "\u6253\u62DB\u547C\u65B9\u5F0F How I greeted", y: 3.45 },
    { label: "\u6211\u7684\u53E5\u5B50 My sentence", y: 4.1 },
  ];

  fields.forEach(function(f) {
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.8, y: f.y, w: 8.4, h: 0.5,
      fill: { color: C.white },
      line: { color: C.primary, width: 0.75, dashType: "dash" },
    });
    s.addText(f.label, {
      x: 0.9, y: f.y, w: 8.2, h: 0.5,
      fontSize: 13, fontFace: "Calibri", color: C.accent, valign: "middle",
    });
  });

  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 29 — Project 分层
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "Project \u5206\u5C42 Differentiation");

  var levels = [
    {
      title: "Beginner \u96F6\u57FA\u7840",
      desc: "\u753B\u753B + \u5199\u8BCD\u8BED\uFF08\u72EE\u5B50\u3001\u91D1\u5B57\u5854\u3001\u70E4\u8089\uFF09",
      en: "Draw pictures + write words",
      bg: C.lightGreen, clr: C.green,
    },
    {
      title: "Intermediate Level 2-3",
      desc: "\u5199\u53E5\u5B50\uFF1A\u300C\u5728___\uFF0C\u4EBA\u4EEC\u7528___\u6253\u62DB\u547C\u300D",
      en: "Write sentences using today's patterns",
      bg: C.lightBlue, clr: C.blue,
    },
    {
      title: "Advanced Level 4+",
      desc: "\u5199\u6BB5\u843D\uFF1A\u6BD4\u8F83\u4E09\u4E2A\u56FD\u5BB6\u7684\u6587\u5316\u548C\u7F8E\u98DF",
      en: "Write a paragraph comparing 3 countries",
      bg: C.lightPurple, clr: C.purple,
    },
  ];

  levels.forEach(function(lv, i) {
    var y = 0.95 + i * 1.4;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.5, y: y, w: 9.0, h: 1.2,
      fill: { color: lv.bg }, rectRadius: 0.12,
    });
    s.addText([
      { text: lv.title, options: { fontSize: 18, fontFace: "Georgia", color: lv.clr, bold: true, breakLine: true } },
      { text: lv.desc, options: { fontSize: 14, fontFace: "Calibri", color: C.accent, breakLine: true } },
      { text: lv.en, options: { fontSize: 12, fontFace: "Calibri", color: C.warmGray, italic: true, breakLine: true } },
    ], { x: 0.7, y: y, w: 8.6, h: 1.2, valign: "middle" });
  });

  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 30 — 分享时间
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "\u5206\u4EAB\u65F6\u95F4 Sharing Time");

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 1.0, y: 1.0, w: 7.8, h: 3.3,
    fill: { color: C.cream }, rectRadius: 0.15,
    line: { color: C.primary, width: 2 },
    shadow: { type: "outer", blur: 5, offset: 3, color: "BDBDBD", opacity: 0.3 },
  });

  s.addText([
    { text: "Partner Share \u548C\u540C\u4F34\u5206\u4EAB", options: { fontSize: 22, fontFace: "Georgia", color: C.accent, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "Step 1: \u627E\u4E00\u4E2A\u540C\u4F34 Find a partner", options: { fontSize: 15, fontFace: "Calibri", color: C.accent, breakLine: true } },
    { text: "Step 2: \u7ED9\u540C\u4F34\u770B\u4F60\u7684\u62A4\u7167\u975E\u6D32\u9875", options: { fontSize: 15, fontFace: "Calibri", color: C.accent, breakLine: true } },
    { text: "Step 3: \u8BFB\u4F60\u5199\u7684\u53E5\u5B50\u7ED9\u540C\u4F34\u542C", options: { fontSize: 15, fontFace: "Calibri", color: C.accent, breakLine: true } },
    { text: "Step 4: \u7528\u4E09\u79CD\u6253\u62DB\u547C\u65B9\u5F0F\u4E92\u76F8\u95EE\u597D\uFF01", options: { fontSize: 15, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
  ], { x: 1.3, y: 1.1, w: 7.2, h: 3.1, valign: "middle" });

  s.addText("\u6BCF\u4E2A\u4EBA\u67092\u5206\u949F\u5206\u4EAB\u65F6\u95F4\uFF01", {
    x: 1.0, y: 4.5, w: 7.8, h: 0.4,
    fontSize: 15, fontFace: "Calibri", color: C.primary, align: "center", bold: true,
  });
  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 31 — 非洲签证章
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.secondary };
  titleBar(s, "\u975E\u6D32\u7B7E\u8BC1\u7AE0 Africa Visa Stamp");

  s.addShape(pres.shapes.OVAL, {
    x: 3.0, y: 1.0, w: 3.8, h: 3.0,
    fill: { color: C.white },
    line: { color: C.primary, width: 4 },
    shadow: { type: "outer", blur: 6, offset: 3, color: "BDBDBD", opacity: 0.35 },
  });

  s.addText([
    { text: "AFRICA", options: { fontSize: 28, fontFace: "Georgia", color: C.primary, bold: true, breakLine: true } },
    { text: "\u975E\u6D32", options: { fontSize: 22, fontFace: "Georgia", color: C.accent, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "APPROVED", options: { fontSize: 20, fontFace: "Georgia", color: C.green, bold: true, breakLine: true } },
    { text: "6/9", options: { fontSize: 18, fontFace: "Calibri", color: C.warmGray, breakLine: true } },
  ], { x: 3.2, y: 1.3, w: 3.4, h: 2.5, align: "center", valign: "middle" });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 1.5, y: 4.2, w: 6.8, h: 0.7,
    fill: { color: C.lightGold }, rectRadius: 0.12,
  });
  s.addText("\u5DF2\u67092\u4E2A\u7B7E\u8BC1\u7AE0\uFF01 \u4E9A\u6D32 \u2714  \u975E\u6D32 \u2714  \u8FD8\u67093\u4E2A\uFF01", {
    x: 1.5, y: 4.2, w: 6.8, h: 0.7,
    fontSize: 17, fontFace: "Calibri", color: C.accent, align: "center", bold: true, valign: "middle",
  });
  footer(s);
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 32 — 明天预告 (GR-003 Europe)
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.dark };
  goldBars(s);

  s.addText("\u660E\u5929\u9884\u544A", {
    x: 0.5, y: 1.1, w: 9.0, h: 0.7,
    fontSize: 34, fontFace: "Georgia", color: C.gold, align: "center", bold: true,
  });
  s.addText("Tomorrow\u2019s Preview", {
    x: 0.5, y: 1.8, w: 9.0, h: 0.5,
    fontSize: 20, fontFace: "Calibri", color: C.lightGold, align: "center",
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 2.2, y: 2.6, w: 5.4, h: 1.8,
    fill: { color: C.accent }, rectRadius: 0.15,
    shadow: { type: "outer", blur: 6, offset: 3, color: "000000", opacity: 0.35 },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 2.2, y: 2.6, w: 5.4, h: 0.35,
    fill: { color: C.primary },
  });
  s.addText("NEXT FLIGHT", {
    x: 2.2, y: 2.6, w: 5.4, h: 0.35,
    fontSize: 12, fontFace: "Georgia", color: C.dark, bold: true, align: "center", valign: "middle",
  });
  s.addText([
    { text: "\u822A\u73ED GR-003 \u00B7 \u6B27\u6D32 EUROPE", options: { fontSize: 22, fontFace: "Georgia", color: C.gold, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "\u51C6\u5907\u597D\u4F60\u7684\u62A4\u7167\uFF0C\u660E\u5929\u98DE\u5F80\u6B27\u6D32\uFF01", options: { fontSize: 16, fontFace: "Calibri", color: C.lightGold, breakLine: true } },
    { text: "\u6CD5\u56FD / \u610F\u5927\u5229 / \u897F\u73ED\u7259 \u7B49\u4F60\u6765\u63A2\u7D22\uFF01", options: { fontSize: 14, fontFace: "Calibri", color: C.warmGray, breakLine: true } },
  ], { x: 2.4, y: 3.0, w: 5.0, h: 1.3, align: "center" });

  s.addText("Global Explorer Camp  \u00B7  \u8C37\u96E8\u4E2D\u6587 GR EDU", {
    x: 0.5, y: 4.8, w: 9.0, h: 0.4,
    fontSize: 12, fontFace: "Calibri", color: C.warmGray, align: "center",
  });
})();


// ══════════════════════════════════════════════════════════════
// SLIDE 33 — End Card
// ══════════════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { fill: C.dark };
  goldBars(s);

  s.addText([
    { text: "Day 2 Complete!", options: { fontSize: 34, fontFace: "Georgia", color: C.gold, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "\u975E\u6D32\u63A2\u9669\u7ED3\u675F", options: { fontSize: 24, fontFace: "Georgia", color: C.lightGold, breakLine: true } },
    { text: "", options: { fontSize: 14, breakLine: true } },
    { text: "\u57C3\u53CA \u00B7 \u80AF\u5C3C\u4E9A \u00B7 \u5357\u975E", options: { fontSize: 20, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "See you tomorrow for Europe!", options: { fontSize: 18, fontFace: "Calibri", color: C.warmGray, breakLine: true } },
    { text: "\u660E\u5929\u6B27\u6D32\u89C1\uFF01", options: { fontSize: 18, fontFace: "Calibri", color: C.warmGray, breakLine: true } },
  ], { x: 1.0, y: 1.0, w: 8.0, h: 3.2, align: "center", valign: "middle" });

  s.addText("Global Explorer Camp  \u00B7  \u8C37\u96E8\u4E2D\u6587 GR EDU", {
    x: 0.5, y: 4.8, w: 9.0, h: 0.4,
    fontSize: 12, fontFace: "Calibri", color: C.warmGray, align: "center",
  });
})();


// ══════════════════════════════════════════════════════════════
// SAVE
// ══════════════════════════════════════════════════════════════
var outPath = "/Users/Huan/projects/summercourse/Chinese/world_trip_pbl/day2_africa.pptx";
pres.writeFile({ fileName: outPath })
  .then(function() { console.log("Created " + pres.slides.length + " slides: " + outPath); })
  .catch(function(err) { console.error("Error:", err); });
