/**
 * Day 3: 欧洲 Europe (6/10) — Global Explorer Camp 环球探索沉浸式夏令营
 * 33 slides: 3 countries (France, Italy, UK), each with 4 slides
 * Run: node create_day3.js
 */
const pptxgen = require("pptxgenjs");
const path = require("path");

const pres = new pptxgen();
pres.defineLayout({ name: "CUSTOM_16x9", width: 10.0, height: 5.625 });
pres.layout = "CUSTOM_16x9";
pres.author = "谷雨中文 GR EDU";
pres.title = "Global Explorer Camp · Day 3: 欧洲 Europe";

// ── Color palette ──
const C = {
  primary:   "1565C0",
  secondary: "E3F2FD",
  accent:    "FFC107",
  dark:      "0D47A1",
  white:     "FFFFFF",
  black:     "333333",
  lightBlue: "BBDEFB",
  lightGold: "FFF8E1",
  france:    "0055A4",
  italy:     "008C45",
  uk:        "003478",
  gray:      "616161",
  green:     "2E7D32",
  red:       "C62828",
};

// ── Image URLs ──
const IMG = {
  eiffel:    "https://upload.wikimedia.org/wikipedia/commons/thumb/a/a8/Tour_Eiffel_Wikimedia_Commons.jpg/800px-Tour_Eiffel_Wikimedia_Commons.jpg",
  colosseum: "https://upload.wikimedia.org/wikipedia/commons/thumb/d/de/Colosseo_2020.jpg/1280px-Colosseo_2020.jpg",
  pizza:     "https://upload.wikimedia.org/wikipedia/commons/thumb/a/a3/Eq_it-na_pizza-margherita_sep2005_sml.jpg/1024px-Eq_it-na_pizza-margherita_sep2005_sml.jpg",
  bigBen:    "https://upload.wikimedia.org/wikipedia/commons/thumb/9/93/Clock_Tower_-_Palace_of_Westminster%2C_London_-_May_2007.jpg/800px-Clock_Tower_-_Palace_of_Westminster%2C_London_-_May_2007.jpg",
};

// ── Helpers ──
function goldBars(slide) {
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 9.8, h: 0.06, fill: { color: C.accent },
  });
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 5.44, w: 9.8, h: 0.06, fill: { color: C.accent },
  });
}

function headerBar(slide, text, barColor) {
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 9.8, h: 0.72,
    fill: { color: barColor || C.primary },
  });
  slide.addText(text, {
    x: 0.4, y: 0.06, w: 9.0, h: 0.6,
    fontSize: 28, fontFace: "Georgia", color: C.white, bold: true,
  });
}

function footer(slide, text) {
  slide.addText(text || "谷雨中文 GR EDU  |  Global Explorer Camp", {
    x: 0.4, y: 5.1, w: 9.0, h: 0.32,
    fontSize: 11, fontFace: "Calibri", color: C.gray, align: "center",
  });
}

function imgWithFallback(slide, url, x, y, w, h, label) {
  slide.addShape(pres.ShapeType.roundRect, {
    x: x, y: y, w: w, h: h, rectRadius: 0.12,
    fill: { color: C.lightBlue },
  });
  slide.addText(label || "Photo", {
    x: x, y: y + h * 0.4, w: w, h: h * 0.2,
    fontSize: 13, fontFace: "Calibri", color: C.gray, align: "center", italic: true,
  });
  slide.addImage({ path: url, x: x, y: y, w: w, h: h });
}


// ══════════════════════════════════════════════════
// Slide 1 — Boarding Time (GR-003, blue bg)
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  goldBars(s);

  s.addText([
    { text: "Global Explorer Camp", options: { fontSize: 34, fontFace: "Georgia", color: C.accent, bold: true, breakLine: true } },
    { text: "环球探索沉浸式夏令营", options: { fontSize: 20, fontFace: "Georgia", color: C.white, breakLine: true } },
  ], { x: 0.6, y: 0.5, w: 8.6, h: 1.1, align: "center" });

  s.addShape(pres.ShapeType.rect, {
    x: 2.8, y: 1.7, w: 4.2, h: 0.04, fill: { color: C.accent },
  });

  // Boarding pass card
  s.addShape(pres.ShapeType.roundRect, {
    x: 2.1, y: 2.0, w: 5.6, h: 2.1, rectRadius: 0.15,
    fill: { color: C.lightGold },
    line: { color: C.accent, width: 2.5 },
  });

  // Dashed divider
  s.addShape(pres.ShapeType.line, {
    x: 5.9, y: 2.1, w: 0, h: 1.9,
    line: { color: C.accent, width: 1.5, dashType: "dash" },
  });

  s.addText("BOARDING PASS", {
    x: 2.2, y: 2.05, w: 3.6, h: 0.32,
    fontSize: 12, fontFace: "Georgia", bold: true, color: C.dark,
  });

  s.addText([
    { text: "航班 GR-003", options: { fontSize: 14, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "From: Africa 非洲", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "To: EUROPE 欧洲", options: { fontSize: 14, fontFace: "Calibri", color: C.dark, bold: true, breakLine: true } },
    { text: "Date: 6/10 Wednesday", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
  ], { x: 2.3, y: 2.45, w: 3.4, h: 1.5, valign: "top" });

  s.addText([
    { text: "EUROPE", options: { fontSize: 22, fontFace: "Georgia", bold: true, color: C.dark, breakLine: true } },
    { text: "Day 3", options: { fontSize: 16, fontFace: "Georgia", bold: true, color: C.primary, breakLine: true } },
  ], { x: 6.1, y: 2.4, w: 1.4, h: 1.2, align: "center", valign: "top" });

  // CTA
  s.addShape(pres.ShapeType.roundRect, {
    x: 2.5, y: 4.3, w: 4.8, h: 0.5,
    fill: { color: C.accent }, rectRadius: 0.15,
  });
  s.addText("请出示你的护照！Show your passport!", {
    x: 2.5, y: 4.3, w: 4.8, h: 0.5,
    fontSize: 17, fontFace: "Calibri", color: C.dark, align: "center", bold: true,
  });

  footer(s);
}


// ══════════════════════════════════════════════════
// Slide 2 — 护照进度 (2 stamps done)
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "护照签证章进度 Passport Stamps");

  // Progress line
  s.addShape(pres.ShapeType.rect, {
    x: 0.9, y: 2.55, w: 8.0, h: 0.06,
    fill: { color: C.primary },
  });

  const stops = [
    { label: "亚洲\nAsia",    x: 1.5,  status: "done" },
    { label: "非洲\nAfrica",   x: 3.5,  status: "done" },
    { label: "欧洲\nEurope",   x: 5.5,  status: "current" },
    { label: "美洲\nAmericas", x: 7.2,  status: "future" },
    { label: "文化展\nExpo",   x: 8.6,  status: "future" },
  ];

  stops.forEach(function(st) {
    var dotColor = st.status === "done" ? C.primary : st.status === "current" ? C.accent : "BDBDBD";
    var dotSize = st.status === "current" ? 0.65 : 0.48;
    var offy = (0.65 - dotSize) / 2;

    s.addShape(pres.ShapeType.ellipse, {
      x: st.x - dotSize / 2, y: 2.26 + offy, w: dotSize, h: dotSize,
      fill: { color: dotColor },
    });

    var icon = st.status === "done" ? "V" : st.status === "current" ? ">" : "";
    if (icon) {
      s.addText(icon, {
        x: st.x - 0.3, y: 1.7, w: 0.6, h: 0.45,
        fontSize: 22, fontFace: "Georgia", color: C.primary, align: "center", bold: true,
      });
    }

    s.addText(st.label, {
      x: st.x - 0.7, y: 2.95, w: 1.4, h: 0.65,
      fontSize: 15, fontFace: "Calibri",
      color: st.status === "done" ? C.primary : st.status === "current" ? C.accent : C.gray,
      align: "center", bold: st.status !== "future",
    });
  });

  // Summary
  s.addShape(pres.ShapeType.roundRect, {
    x: 1.2, y: 3.85, w: 7.4, h: 0.75,
    fill: { color: C.lightGold }, rectRadius: 0.15,
    line: { color: C.accent, width: 2 },
  });
  s.addText("亚洲 V  非洲 V  欧洲（今天!）>  美洲  >  文化展", {
    x: 1.3, y: 3.85, w: 7.2, h: 0.75,
    fontSize: 16, fontFace: "Calibri",
    color: C.dark, align: "center", bold: true, valign: "middle",
  });

  s.addText("已有2个签证章！今天拿第3个！", {
    x: 1.2, y: 4.75, w: 7.4, h: 0.35,
    fontSize: 14, fontFace: "Calibri", color: C.gray, align: "center", italic: true,
  });
}


// ══════════════════════════════════════════════════
// Slide 3 — 今天目标
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "今天的目标 Today\u2019s Goals");

  // Morning card
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.5, y: 1.0, w: 4.3, h: 2.5, rectRadius: 0.12,
    fill: { color: C.white }, line: { color: C.primary, width: 2 },
  });
  s.addText("Morning 上午", {
    x: 0.7, y: 1.1, w: 3.9, h: 0.4,
    fontSize: 20, fontFace: "Georgia", color: C.primary, bold: true,
  });
  s.addText([
    { text: "认识欧洲3个国家（法/意/英）", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "了解每国文化礼节和美食", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "学习打招呼方式和趣味知识", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "完成上午竞赛", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 0.8, y: 1.6, w: 3.8, h: 1.7, valign: "top" });

  // Afternoon card
  s.addShape(pres.ShapeType.roundRect, {
    x: 5.1, y: 1.0, w: 4.3, h: 2.5, rectRadius: 0.12,
    fill: { color: C.white }, line: { color: C.accent, width: 2 },
  });
  s.addText("Afternoon 下午", {
    x: 5.3, y: 1.1, w: 3.9, h: 0.4,
    fontSize: 20, fontFace: "Georgia", color: C.accent, bold: true,
  });
  s.addText([
    { text: "三国对比 + 共同点与不同", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "旅行小贴士 + 丝绸之路", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "角色扮演：欧洲小厨师", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "完成护照欧洲页", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 5.2, y: 1.6, w: 3.9, h: 1.7, valign: "top" });

  // Bottom
  s.addShape(pres.ShapeType.roundRect, {
    x: 1.5, y: 3.8, w: 6.8, h: 0.65,
    fill: { color: C.lightGold }, rectRadius: 0.12,
    line: { color: C.accent, width: 1.5 },
  });
  s.addText("今天的目标：获得欧洲签证章！", {
    x: 1.5, y: 3.8, w: 6.8, h: 0.65,
    fontSize: 18, fontFace: "Calibri", color: C.dark, align: "center", bold: true, valign: "middle",
  });
  footer(s);
}


// ══════════════════════════════════════════════════
// Slide 4 — 认识欧洲
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "认识欧洲 Meet Europe!");

  const stats = [
    { title: "44个国家", sub: "44 countries", icon: "44", color: C.primary },
    { title: "第二小的洲", sub: "2nd smallest", icon: "2nd", color: C.dark },
    { title: "古罗马文明", sub: "Ancient Rome", icon: "2700yr", color: C.italy },
    { title: "文艺复兴", sub: "Renaissance", icon: "Art", color: C.france },
    { title: "很多发明", sub: "Many inventions", icon: "Inv", color: C.accent },
  ];

  stats.forEach(function(st, i) {
    var xPos = 0.3 + i * 1.88;
    s.addShape(pres.ShapeType.roundRect, {
      x: xPos, y: 1.0, w: 1.7, h: 2.3, rectRadius: 0.12,
      fill: { color: C.white },
      shadow: { type: "outer", blur: 4, offset: 2, color: "000000", opacity: 0.15 },
    });
    s.addShape(pres.ShapeType.ellipse, {
      x: xPos + 0.45, y: 1.1, w: 0.8, h: 0.8,
      fill: { color: st.color, transparency: 20 },
    });
    s.addText(st.icon, {
      x: xPos + 0.45, y: 1.1, w: 0.8, h: 0.8,
      fontSize: 14, fontFace: "Georgia", color: C.white, bold: true, align: "center",
    });
    s.addText(st.title, {
      x: xPos, y: 2.0, w: 1.7, h: 0.4,
      fontSize: 14, fontFace: "Georgia", color: C.dark, bold: true, align: "center",
    });
    s.addText(st.sub, {
      x: xPos, y: 2.4, w: 1.7, h: 0.35,
      fontSize: 11, fontFace: "Calibri", color: C.gray, align: "center",
    });
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: 0.5, y: 3.6, w: 8.8, h: 1.1, rectRadius: 0.15,
    fill: { color: C.lightGold }, line: { color: C.accent, width: 1.5 },
  });
  s.addText([
    { text: "从古罗马到文艺复兴，欧洲创造了无数改变世界的发明！", options: { fontSize: 15, fontFace: "Calibri", color: C.dark, bold: true, breakLine: true } },
    { text: "印刷术/蒸汽机/汽车/电话...很多都源自欧洲", options: { fontSize: 13, fontFace: "Calibri", color: C.primary, breakLine: true } },
    { text: "今天去3个国家：法国 / 意大利 / 英国", options: { fontSize: 14, fontFace: "Calibri", color: C.dark, bold: true, breakLine: true } },
  ], { x: 0.7, y: 3.65, w: 8.4, h: 1.0, align: "center" });
}


// ══════════════════════════════════════════════════
// Slide 5 — 法国概览
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "法国 France — 概览", C.france);

  // Eiffel photo left
  imgWithFallback(s, IMG.eiffel, 0.3, 0.85, 4.6, 3.5, "Eiffel Tower");

  // Info card right
  s.addShape(pres.ShapeType.roundRect, {
    x: 5.1, y: 0.85, w: 4.5, h: 3.5, rectRadius: 0.15,
    fill: { color: C.white },
    shadow: { type: "outer", blur: 4, offset: 2, color: "000000", opacity: 0.15 },
  });

  s.addText([
    { text: "法国 France", options: { fontSize: 26, fontFace: "Georgia", color: C.france, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "国旗：蓝白红三色 Blue-White-Red", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "人口：6700万", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "语言：法语 French", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "首都：巴黎 Paris", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Bonjour = 你好！", options: { fontSize: 18, fontFace: "Georgia", color: C.france, bold: true, breakLine: true } },
    { text: "「浪漫之都」「时尚之都」", options: { fontSize: 14, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
  ], { x: 5.3, y: 0.95, w: 4.1, h: 3.3, valign: "top" });

  footer(s, "首都巴黎 Paris — 世界时尚与美食之都");
}


// ══════════════════════════════════════════════════
// Slide 6 — 法国文化与礼节
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "法国 France — 文化与礼节", C.france);

  // Left: culture
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.3, y: 0.85, w: 4.6, h: 1.8, rectRadius: 0.12,
    fill: { color: C.white }, line: { color: C.france, width: 1.5 },
  });
  s.addText([
    { text: "文化亮点 Culture", options: { fontSize: 16, fontFace: "Georgia", color: C.france, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 3, breakLine: true } },
    { text: "埃菲尔铁塔 Eiffel Tower — 324米高！", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "卢浮宫 Louvre — 蒙娜丽莎在这里！", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "时尚之都 — 世界时装中心", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
  ], { x: 0.5, y: 0.9, w: 4.2, h: 1.7, valign: "top" });

  // Right: etiquette
  s.addShape(pres.ShapeType.roundRect, {
    x: 5.1, y: 0.85, w: 4.5, h: 3.6, rectRadius: 0.12,
    fill: { color: C.lightGold }, line: { color: C.accent, width: 2 },
  });
  s.addText([
    { text: "礼节 Etiquette", options: { fontSize: 16, fontFace: "Georgia", color: C.dark, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 3, breakLine: true } },
    { text: "见面亲脸颊（la bise）不握手！", options: { fontSize: 12, fontFace: "Calibri", color: C.red, bold: true, breakLine: true } },
    { text: "说 Bonjour 很重要!", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "吃饭是艺术，午餐1-2小时", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "不要催服务员!", options: { fontSize: 12, fontFace: "Calibri", color: C.red, bold: true, breakLine: true } },
    { text: "面包放桌上，不放盘子里", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "给小费不是必须的", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
  ], { x: 5.3, y: 0.9, w: 4.1, h: 3.5, valign: "top" });

  // Bottom: fun fact
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.3, y: 2.85, w: 4.6, h: 1.6, rectRadius: 0.12,
    fill: { color: C.white }, line: { color: C.france, width: 1.5 },
  });
  s.addText([
    { text: "你知道吗？", options: { fontSize: 14, fontFace: "Georgia", color: C.france, bold: true, breakLine: true } },
    { text: "法国人见面时，根据地区", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "要亲2-4次脸颊！", options: { fontSize: 13, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
  ], { x: 0.5, y: 2.95, w: 4.2, h: 1.4, valign: "top" });

  footer(s);
}


// ══════════════════════════════════════════════════
// Slide 7 — 法国美食
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "法国 France — 美食", C.france);

  const foods = [
    { zh: "可颂 croissant", desc: "又香又酥的法式早餐", color: C.france },
    { zh: "法棍 baguette", desc: "每天卖3000万根!", color: C.dark },
    { zh: "奶酪 cheese", desc: "有400多种!", color: C.accent },
    { zh: "蜗牛 escargot", desc: "法国人的美味！", color: C.primary },
    { zh: "马卡龙 macaron", desc: "彩色的甜蜜小饼", color: C.france },
  ];

  foods.forEach(function(f, i) {
    var col = i % 3;
    var row = Math.floor(i / 3);
    var xPos = 0.3 + col * 3.15;
    var yPos = 0.85 + row * 1.7;
    var cardW = 2.95;
    s.addShape(pres.ShapeType.roundRect, {
      x: xPos, y: yPos, w: cardW, h: 1.45, rectRadius: 0.12,
      fill: { color: C.white }, line: { color: f.color, width: 1.5 },
    });
    s.addText(f.zh, {
      x: xPos + 0.15, y: yPos + 0.1, w: cardW - 0.3, h: 0.5,
      fontSize: 16, fontFace: "Georgia", color: f.color, bold: true,
    });
    s.addText(f.desc, {
      x: xPos + 0.15, y: yPos + 0.65, w: cardW - 0.3, h: 0.6,
      fontSize: 13, fontFace: "Calibri", color: C.black,
    });
  });

  // Bottom quote
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.5, y: 4.3, w: 9.0, h: 0.7, rectRadius: 0.12,
    fill: { color: C.lightGold }, line: { color: C.accent, width: 1.5 },
  });
  s.addText("「法国人觉得吃饭是生活中最重要的事之一」", {
    x: 0.5, y: 4.3, w: 9.0, h: 0.7,
    fontSize: 16, fontFace: "Calibri", color: C.dark, bold: true, align: "center",
  });
}


// ══════════════════════════════════════════════════
// Slide 8-9 — 法国 Check Understanding (2 slides)
// ══════════════════════════════════════════════════
{
  const allQs = [
    { q: "法国人见面怎么打招呼？", a: "亲脸颊 (la bise)!" },
    { q: "法国有多少种奶酪？", a: "400多种!" },
    { q: "面包放哪里？", a: "放桌上，不放盘子里!" },
    { q: "法国最著名的建筑是什么？", a: "埃菲尔铁塔" },
    { q: "法国有多少种奶酪？", a: "400多种" },
    { q: "法国人午餐一般吃多久？", a: "1-2小时" },
    { q: "蒙娜丽莎在哪个博物馆？", a: "卢浮宫" },
    { q: "在法国需要给小费吗？", a: "不是必须的" },
    { q: "法国每天卖多少根法棍面包？", a: "3000万根" },
    { q: "面包应该放在哪里？", a: "桌上，不放盘子里" },
  ];

  [0, 1].forEach(function(page) {
    const s = pres.addSlide();
    s.background = { color: C.secondary };
    headerBar(s, "Check Understanding \u2014 法国 France (" + (page + 1) + "/2)", C.france);

    var pageQs = allQs.slice(page * 5, page * 5 + 5);
    pageQs.forEach(function(item, i) {
      var num = page * 5 + i + 1;
      var yy = 1.0 + i * 0.85;
      var rowBg = i % 2 === 0 ? C.white : C.lightBlue;

      s.addShape(pres.ShapeType.roundRect, {
        x: 0.3, y: yy, w: 9.2, h: 0.72, rectRadius: 0.1,
        fill: { color: rowBg },
        line: { color: C.france, width: 1 },
      });
      s.addShape(pres.ShapeType.ellipse, {
        x: 0.45, y: yy + 0.13, w: 0.45, h: 0.45,
        fill: { color: C.france },
      });
      s.addText("" + num, {
        x: 0.45, y: yy + 0.13, w: 0.45, h: 0.45,
        fontSize: 14, fontFace: "Georgia", color: C.white,
        bold: true, align: "center", valign: "middle",
      });
      s.addText(item.q, {
        x: 1.05, y: yy, w: 4.45, h: 0.72,
        fontSize: 12, fontFace: "Calibri", color: C.dark,
        bold: true, align: "left", valign: "middle",
      });
      s.addText("\u2192 " + item.a, {
        x: 5.6, y: yy, w: 3.75, h: 0.72,
        fontSize: 11, fontFace: "Calibri", color: C.france,
        bold: true, align: "left", valign: "middle",
      });
    });
    footer(s, "法国 France  \u2713  完成！");
  });
}


// ══════════════════════════════════════════════════
// Slide 9 — 意大利概览
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "意大利 Italy — 概览", C.italy);

  // Colosseum photo left
  imgWithFallback(s, IMG.colosseum, 0.3, 0.85, 4.6, 3.5, "Colosseum");

  // Info card right
  s.addShape(pres.ShapeType.roundRect, {
    x: 5.1, y: 0.85, w: 4.5, h: 3.5, rectRadius: 0.15,
    fill: { color: C.white },
    shadow: { type: "outer", blur: 4, offset: 2, color: "000000", opacity: 0.15 },
  });

  s.addText([
    { text: "意大利 Italy", options: { fontSize: 26, fontFace: "Georgia", color: C.italy, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "国旗：绿白红 Green-White-Red", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "人口：5900万", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "语言：意大利语 Italian", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "首都：罗马 Rome", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "形状像靴子 Boot-shaped!", options: { fontSize: 14, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "Ciao = 你好！", options: { fontSize: 18, fontFace: "Georgia", color: C.italy, bold: true, breakLine: true } },
  ], { x: 5.3, y: 0.95, w: 4.1, h: 3.3, valign: "top" });

  footer(s, "首都罗马 Rome — 形状像靴子!");
}


// ══════════════════════════════════════════════════
// Slide 10 — 意大利文化与礼节
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "意大利 Italy — 文化与礼节", C.italy);

  // Left: culture
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.3, y: 0.85, w: 4.6, h: 1.8, rectRadius: 0.12,
    fill: { color: C.white }, line: { color: C.italy, width: 1.5 },
  });
  s.addText([
    { text: "文化亮点 Culture", options: { fontSize: 16, fontFace: "Georgia", color: C.italy, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 3, breakLine: true } },
    { text: "斗兽场 Colosseum — 能坐5万人!", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "比萨斜塔 — 歪了快4度!", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "文艺复兴起源地 — 达芬奇/米开朗基罗", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
  ], { x: 0.5, y: 0.9, w: 4.2, h: 1.7, valign: "top" });

  // Right: etiquette
  s.addShape(pres.ShapeType.roundRect, {
    x: 5.1, y: 0.85, w: 4.5, h: 3.6, rectRadius: 0.12,
    fill: { color: C.lightGold }, line: { color: C.accent, width: 2 },
  });
  s.addText([
    { text: "礼节 Etiquette", options: { fontSize: 16, fontFace: "Georgia", color: C.dark, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 3, breakLine: true } },
    { text: "说 Ciao! 打招呼", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "意大利人说话爱用手势!", options: { fontSize: 12, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
    { text: "吃饭时间很固定（午餐1-3pm）", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "不要在咖啡里加奶（早上除外!）", options: { fontSize: 12, fontFace: "Calibri", color: C.red, bold: true, breakLine: true } },
    { text: "不要在披萨上加菠萝（会生气!）", options: { fontSize: 12, fontFace: "Calibri", color: C.red, bold: true, breakLine: true } },
    { text: "进教堂要穿长袖长裤", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
  ], { x: 5.3, y: 0.9, w: 4.1, h: 3.5, valign: "top" });

  // Bottom
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.3, y: 2.85, w: 4.6, h: 1.6, rectRadius: 0.12,
    fill: { color: C.white }, line: { color: C.italy, width: 1.5 },
  });
  s.addText([
    { text: "你知道吗？", options: { fontSize: 14, fontFace: "Georgia", color: C.italy, bold: true, breakLine: true } },
    { text: "意大利人用手势表达感情", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "超过250种常用手势！", options: { fontSize: 13, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
  ], { x: 0.5, y: 2.95, w: 4.2, h: 1.4, valign: "top" });

  footer(s);
}


// ══════════════════════════════════════════════════
// Slide 11 — 意大利美食
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "意大利 Italy — 美食", C.italy);

  // Pizza photo left
  imgWithFallback(s, IMG.pizza, 0.3, 0.85, 4.2, 3.2, "Pizza Margherita");

  // Foods right
  s.addShape(pres.ShapeType.roundRect, {
    x: 4.7, y: 0.85, w: 4.9, h: 3.2, rectRadius: 0.15,
    fill: { color: C.white },
    shadow: { type: "outer", blur: 4, offset: 2, color: "000000", opacity: 0.15 },
  });
  s.addText([
    { text: "意大利美食", options: { fontSize: 22, fontFace: "Georgia", color: C.italy, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "披萨 Pizza Margherita", options: { fontSize: 14, fontFace: "Calibri", color: C.dark, bold: true, breakLine: true } },
    { text: "  发源地是那不勒斯 Naples!", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "意面 Pasta", options: { fontSize: 14, fontFace: "Calibri", color: C.dark, bold: true, breakLine: true } },
    { text: "  有350多种形状!", options: { fontSize: 12, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
    { text: "冰淇淋 Gelato", options: { fontSize: 14, fontFace: "Calibri", color: C.dark, bold: true, breakLine: true } },
    { text: "  比普通冰淇淋更浓更滑", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "浓缩咖啡 Espresso", options: { fontSize: 14, fontFace: "Calibri", color: C.dark, bold: true, breakLine: true } },
    { text: "提拉米苏 Tiramisu", options: { fontSize: 14, fontFace: "Calibri", color: C.dark, bold: true, breakLine: true } },
  ], { x: 4.9, y: 0.95, w: 4.5, h: 3.0, valign: "top" });

  // Bottom
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.5, y: 4.3, w: 9.0, h: 0.7, rectRadius: 0.12,
    fill: { color: C.lightGold }, line: { color: C.accent, width: 1.5 },
  });
  s.addText("「披萨发源地是那不勒斯 Naples」", {
    x: 0.5, y: 4.3, w: 9.0, h: 0.7,
    fontSize: 16, fontFace: "Calibri", color: C.dark, bold: true, align: "center",
  });
}


// ══════════════════════════════════════════════════
// Slide 12-13 — 意大利 Check Understanding (2 slides)
// ══════════════════════════════════════════════════
{
  const allQs = [
    { q: "意大利形状像什么？", a: "靴子 Boot!" },
    { q: "意面有多少种形状？", a: "350多种!" },
    { q: "什么时候不要在咖啡里加奶？", a: "早上以后!" },
    { q: "意大利最有名的食物是什么？", a: "披萨和意面" },
    { q: "意面有多少种形状？", a: "350多种" },
    { q: "什么时候不能在咖啡里加牛奶？", a: "早上以后" },
    { q: "罗马斗兽场能容纳多少人？", a: "5万人" },
    { q: "比萨斜塔倾斜了多少度？", a: "约4度" },
    { q: "文艺复兴从哪个国家开始？", a: "意大利" },
    { q: "能在披萨上加菠萝吗？", a: "意大利人会生气" },
  ];

  [0, 1].forEach(function(page) {
    const s = pres.addSlide();
    s.background = { color: C.secondary };
    headerBar(s, "Check Understanding \u2014 意大利 Italy (" + (page + 1) + "/2)", C.italy);

    var pageQs = allQs.slice(page * 5, page * 5 + 5);
    pageQs.forEach(function(item, i) {
      var num = page * 5 + i + 1;
      var yy = 1.0 + i * 0.85;
      var rowBg = i % 2 === 0 ? C.white : C.lightBlue;

      s.addShape(pres.ShapeType.roundRect, {
        x: 0.3, y: yy, w: 9.2, h: 0.72, rectRadius: 0.1,
        fill: { color: rowBg },
        line: { color: C.italy, width: 1 },
      });
      s.addShape(pres.ShapeType.ellipse, {
        x: 0.45, y: yy + 0.13, w: 0.45, h: 0.45,
        fill: { color: C.italy },
      });
      s.addText("" + num, {
        x: 0.45, y: yy + 0.13, w: 0.45, h: 0.45,
        fontSize: 14, fontFace: "Georgia", color: C.white,
        bold: true, align: "center", valign: "middle",
      });
      s.addText(item.q, {
        x: 1.05, y: yy, w: 4.45, h: 0.72,
        fontSize: 12, fontFace: "Calibri", color: C.dark,
        bold: true, align: "left", valign: "middle",
      });
      s.addText("\u2192 " + item.a, {
        x: 5.6, y: yy, w: 3.75, h: 0.72,
        fontSize: 11, fontFace: "Calibri", color: C.italy,
        bold: true, align: "left", valign: "middle",
      });
    });
    footer(s, "意大利 Italy  \u2713  完成！");
  });
}


// ══════════════════════════════════════════════════
// Slide 13 — 英国概览
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "英国 United Kingdom — 概览", C.uk);

  // Big Ben photo left
  imgWithFallback(s, IMG.bigBen, 0.3, 0.85, 4.6, 3.5, "Big Ben");

  // Info card right
  s.addShape(pres.ShapeType.roundRect, {
    x: 5.1, y: 0.85, w: 4.5, h: 3.5, rectRadius: 0.15,
    fill: { color: C.white },
    shadow: { type: "outer", blur: 4, offset: 2, color: "000000", opacity: 0.15 },
  });

  s.addText([
    { text: "英国 United Kingdom", options: { fontSize: 24, fontFace: "Georgia", color: C.uk, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "国旗：米字旗 Union Jack", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "人口：6700万", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "语言：英语 English", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "首都：伦敦 London", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "有国王！Has a King!", options: { fontSize: 14, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "Hello = 你好！", options: { fontSize: 18, fontFace: "Georgia", color: C.uk, bold: true, breakLine: true } },
  ], { x: 5.3, y: 0.95, w: 4.1, h: 3.3, valign: "top" });

  footer(s, "首都伦敦 London — 有国王的国家!");
}


// ══════════════════════════════════════════════════
// Slide 14 — 英国文化与礼节
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "英国 UK — 文化与礼节", C.uk);

  // Left: culture
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.3, y: 0.85, w: 4.6, h: 1.8, rectRadius: 0.12,
    fill: { color: C.white }, line: { color: C.uk, width: 1.5 },
  });
  s.addText([
    { text: "文化亮点 Culture", options: { fontSize: 16, fontFace: "Georgia", color: C.uk, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 3, breakLine: true } },
    { text: "大本钟 Big Ben — 伦敦地标", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "白金汉宫 Buckingham Palace", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "哈利波特 Harry Potter!", options: { fontSize: 12, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
    { text: "下午茶文化 vs 中国茶文化", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
  ], { x: 0.5, y: 0.9, w: 4.2, h: 1.7, valign: "top" });

  // Right: etiquette
  s.addShape(pres.ShapeType.roundRect, {
    x: 5.1, y: 0.85, w: 4.5, h: 3.6, rectRadius: 0.12,
    fill: { color: C.lightGold }, line: { color: C.accent, width: 2 },
  });
  s.addText([
    { text: "礼节 Etiquette", options: { fontSize: 16, fontFace: "Georgia", color: C.dark, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 3, breakLine: true } },
    { text: "排队！英国人最爱排队!", options: { fontSize: 12, fontFace: "Calibri", color: C.red, bold: true, breakLine: true } },
    { text: "说 Please 和 Thank you 非常重要", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "不要问别人赚多少钱!", options: { fontSize: 12, fontFace: "Calibri", color: C.red, bold: true, breakLine: true } },
    { text: "聊天气是最安全的话题", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "靠左行走和开车!", options: { fontSize: 12, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
  ], { x: 5.3, y: 0.9, w: 4.1, h: 3.5, valign: "top" });

  // Bottom
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.3, y: 2.85, w: 4.6, h: 1.6, rectRadius: 0.12,
    fill: { color: C.white }, line: { color: C.uk, width: 1.5 },
  });
  s.addText([
    { text: "英式下午茶 vs 中国茶", options: { fontSize: 14, fontFace: "Georgia", color: C.uk, bold: true, breakLine: true } },
    { text: "英国: scones + 三明治 + 红茶", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "中国: 茶叶 + 茶点 + 功夫茶", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
  ], { x: 0.5, y: 2.95, w: 4.2, h: 1.4, valign: "top" });

  footer(s);
}


// ══════════════════════════════════════════════════
// Slide 15 — 英国美食
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "英国 UK — 美食", C.uk);

  const foods = [
    { zh: "炸鱼薯条 Fish & Chips", desc: "英国最经典的食物!", color: C.uk },
    { zh: "英式早餐 Full English", desc: "鸡蛋+培根+香肠+豆子", color: C.dark },
    { zh: "牧羊人派 Shepherd\u2019s Pie", desc: "土豆泥+肉馅，冬天最爱", color: C.primary },
    { zh: "烤牛肉 Roast Beef", desc: "周日家庭传统大餐", color: C.accent },
    { zh: "下午茶 Scones+三明治", desc: "英国人每天喝1.65亿杯茶!", color: C.uk },
  ];

  foods.forEach(function(f, i) {
    var col = i % 3;
    var row = Math.floor(i / 3);
    var xPos = 0.3 + col * 3.15;
    var yPos = 0.85 + row * 1.7;
    var cardW = 2.95;
    s.addShape(pres.ShapeType.roundRect, {
      x: xPos, y: yPos, w: cardW, h: 1.45, rectRadius: 0.12,
      fill: { color: C.white }, line: { color: f.color, width: 1.5 },
    });
    s.addText(f.zh, {
      x: xPos + 0.15, y: yPos + 0.1, w: cardW - 0.3, h: 0.5,
      fontSize: 14, fontFace: "Georgia", color: f.color, bold: true,
    });
    s.addText(f.desc, {
      x: xPos + 0.15, y: yPos + 0.65, w: cardW - 0.3, h: 0.6,
      fontSize: 12, fontFace: "Calibri", color: C.black,
    });
  });

  // Bottom quote
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.5, y: 4.3, w: 9.0, h: 0.7, rectRadius: 0.12,
    fill: { color: C.lightGold }, line: { color: C.accent, width: 1.5 },
  });
  s.addText("「英国人每天喝1.65亿杯茶！」", {
    x: 0.5, y: 4.3, w: 9.0, h: 0.7,
    fontSize: 16, fontFace: "Calibri", color: C.dark, bold: true, align: "center",
  });
}


// ══════════════════════════════════════════════════
// Slide 16-17 — 英国 Check Understanding (2 slides)
// ══════════════════════════════════════════════════
{
  const allQs = [
    { q: "英国人最爱做什么？", a: "排队!" },
    { q: "英国人每天喝多少杯茶？", a: "1.65亿杯!" },
    { q: "在英国开车靠哪边？", a: "左边!" },
    { q: "英国最有名的钟叫什么？", a: "大本钟 Big Ben" },
    { q: "英国人每天喝多少杯茶？", a: "1.65亿杯" },
    { q: "哈利波特的作者是哪国人？", a: "英国人" },
    { q: "在英国开车靠哪边？", a: "左边" },
    { q: "在英国最安全的聊天话题是什么？", a: "天气" },
    { q: "英国有国王还是总统？", a: "国王" },
    { q: "在英国什么行为最重要？", a: "排队！说Please和Thank you" },
  ];

  [0, 1].forEach(function(page) {
    const s = pres.addSlide();
    s.background = { color: C.secondary };
    headerBar(s, "Check Understanding \u2014 英国 UK (" + (page + 1) + "/2)", C.uk);

    var pageQs = allQs.slice(page * 5, page * 5 + 5);
    pageQs.forEach(function(item, i) {
      var num = page * 5 + i + 1;
      var yy = 1.0 + i * 0.85;
      var rowBg = i % 2 === 0 ? C.white : C.lightBlue;

      s.addShape(pres.ShapeType.roundRect, {
        x: 0.3, y: yy, w: 9.2, h: 0.72, rectRadius: 0.1,
        fill: { color: rowBg },
        line: { color: C.uk, width: 1 },
      });
      s.addShape(pres.ShapeType.ellipse, {
        x: 0.45, y: yy + 0.13, w: 0.45, h: 0.45,
        fill: { color: C.uk },
      });
      s.addText("" + num, {
        x: 0.45, y: yy + 0.13, w: 0.45, h: 0.45,
        fontSize: 14, fontFace: "Georgia", color: C.white,
        bold: true, align: "center", valign: "middle",
      });
      s.addText(item.q, {
        x: 1.05, y: yy, w: 4.45, h: 0.72,
        fontSize: 12, fontFace: "Calibri", color: C.dark,
        bold: true, align: "left", valign: "middle",
      });
      s.addText("\u2192 " + item.a, {
        x: 5.6, y: yy, w: 3.75, h: 0.72,
        fontSize: 11, fontFace: "Calibri", color: C.uk,
        bold: true, align: "left", valign: "middle",
      });
    });
    footer(s, "英国 UK  \u2713  完成！");
  });
}


// ══════════════════════════════════════════════════
// Slide 17 — Mini Role Play
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "Mini Role Play — 三国打招呼 (5 min)", C.primary);

  const scenes = [
    { country: "法国 France", action: "亲脸颊 + Bonjour!", desc: "假装跟同学亲脸颊\n（不是真亲哦！）说 Bonjour!", color: C.france },
    { country: "意大利 Italy", action: "Ciao! + 手势", desc: "大声说 Ciao!\n用手势表达「好吃」", color: C.italy },
    { country: "英国 UK", action: "排队 + Please/Thank you", desc: "排好队\n说 Please 和 Thank you", color: C.uk },
  ];

  scenes.forEach(function(sc, i) {
    var xPos = 0.3 + i * 3.2;
    s.addShape(pres.ShapeType.roundRect, {
      x: xPos, y: 0.85, w: 2.95, h: 3.0, rectRadius: 0.15,
      fill: { color: C.white }, line: { color: sc.color, width: 2 },
    });
    s.addText(sc.country, {
      x: xPos, y: 0.95, w: 2.95, h: 0.4,
      fontSize: 16, fontFace: "Georgia", color: sc.color, bold: true, align: "center",
    });
    s.addShape(pres.ShapeType.rect, {
      x: xPos + 0.3, y: 1.4, w: 2.35, h: 0.03,
      fill: { color: sc.color, transparency: 40 },
    });
    s.addText(sc.action, {
      x: xPos, y: 1.5, w: 2.95, h: 0.55,
      fontSize: 15, fontFace: "Calibri", color: C.dark, bold: true, align: "center",
    });
    s.addText(sc.desc, {
      x: xPos + 0.15, y: 2.15, w: 2.65, h: 1.5,
      fontSize: 12, fontFace: "Calibri", color: C.black, align: "center",
    });
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: 0.5, y: 4.1, w: 8.8, h: 0.7, rectRadius: 0.12,
    fill: { color: C.lightGold }, line: { color: C.accent, width: 1.5 },
  });
  s.addText("跟着老师一起做！全班分成3组，每组代表一个国家！", {
    x: 0.5, y: 4.1, w: 8.8, h: 0.7,
    fontSize: 16, fontFace: "Calibri", color: C.dark, bold: true, align: "center",
  });
}


// ══════════════════════════════════════════════════
// Slide 18 — 上午竞赛
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "上午竞赛 Morning Quiz Challenge!", C.accent);

  const questions = [
    { q: "Q1: 法国人见面亲几次脸颊？", a: "2-4次!", color: C.france },
    { q: "Q2: 意大利的形状像什么？", a: "靴子!", color: C.italy },
    { q: "Q3: 英国人开车靠哪边？", a: "左边!", color: C.uk },
    { q: "Q4: 不要在意大利披萨上加什么？", a: "菠萝!", color: C.italy },
  ];

  questions.forEach(function(item, i) {
    var yPos = 0.9 + i * 1.0;
    s.addShape(pres.ShapeType.roundRect, {
      x: 0.4, y: yPos, w: 5.2, h: 0.8, rectRadius: 0.12,
      fill: { color: C.white }, line: { color: item.color, width: 2 },
    });
    s.addText(item.q, {
      x: 0.6, y: yPos, w: 4.8, h: 0.8,
      fontSize: 15, fontFace: "Calibri", color: C.dark, bold: true, valign: "middle",
    });
    s.addShape(pres.ShapeType.roundRect, {
      x: 5.8, y: yPos, w: 3.6, h: 0.8, rectRadius: 0.12,
      fill: { color: item.color, transparency: 15 },
      line: { color: item.color, width: 1.5 },
    });
    s.addText(item.a, {
      x: 5.8, y: yPos, w: 3.6, h: 0.8,
      fontSize: 16, fontFace: "Calibri", color: C.dark, bold: true, align: "center", valign: "middle",
    });
  });

  footer(s, "答对最多的小组获得加分！");
}


// ══════════════════════════════════════════════════
// Slide 19 — Project 提醒
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "Project 提醒 — 下午完成欧洲页!", C.primary);

  s.addShape(pres.ShapeType.roundRect, {
    x: 0.5, y: 0.9, w: 8.8, h: 3.4, rectRadius: 0.15,
    fill: { color: C.white },
    shadow: { type: "outer", blur: 4, offset: 2, color: "000000", opacity: 0.15 },
  });

  s.addText([
    { text: "下午你需要完成：", options: { fontSize: 20, fontFace: "Georgia", color: C.dark, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "画出你最喜欢的欧洲地标", options: { fontSize: 15, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "写下你想吃的欧洲美食", options: { fontSize: 15, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "记录三国不同的打招呼方式", options: { fontSize: 15, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "写一个你学到的文化礼节", options: { fontSize: 15, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "用句型写一句话", options: { fontSize: 15, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 0.8, y: 1.0, w: 8.2, h: 3.2, valign: "top" });

  s.addShape(pres.ShapeType.roundRect, {
    x: 1.5, y: 4.5, w: 6.8, h: 0.6, rectRadius: 0.12,
    fill: { color: C.lightGold }, line: { color: C.accent, width: 2 },
  });
  s.addText("记住今天看到的地标和文化！", {
    x: 1.5, y: 4.5, w: 6.8, h: 0.6,
    fontSize: 16, fontFace: "Calibri", color: C.dark, bold: true, align: "center",
  });
}


// ══════════════════════════════════════════════════
// Slide 20 — 下午开始
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  goldBars(s);

  s.addText("下午开始 Afternoon Session", {
    x: 0.5, y: 2.0, w: 9.0, h: 0.8,
    fontSize: 34, fontFace: "Georgia", color: C.accent, bold: true, align: "center",
  });

  s.addText("Section B: 深入探索与Project时间", {
    x: 0.5, y: 2.85, w: 9.0, h: 0.55,
    fontSize: 20, fontFace: "Calibri", color: C.white, align: "center",
  });

  s.addText("Deep Dive & Passport Project", {
    x: 0.5, y: 3.35, w: 9.0, h: 0.5,
    fontSize: 16, fontFace: "Calibri", color: C.accent, italic: true, align: "center",
  });
}


// ══════════════════════════════════════════════════
// Slide 21 — 快速复习
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "快速复习 Quick Review");

  s.addText("上午我们去了哪3个国家？", {
    x: 0.5, y: 0.85, w: 9.0, h: 0.45,
    fontSize: 20, fontFace: "Calibri", color: C.dark, bold: true, align: "center",
  });

  const review = [
    { name: "法国", key: "埃菲尔铁塔\n可颂/奶酪\nBonjour!", color: C.france },
    { name: "意大利", key: "斗兽场\n披萨/意面\nCiao!", color: C.italy },
    { name: "英国", key: "大本钟\n炸鱼薯条\nHello!", color: C.uk },
  ];

  review.forEach(function(r, i) {
    var xPos = 0.6 + i * 3.1;
    s.addShape(pres.ShapeType.roundRect, {
      x: xPos, y: 1.45, w: 2.7, h: 2.5, rectRadius: 0.15,
      fill: { color: C.white }, line: { color: r.color, width: 2 },
    });
    s.addText(r.name, {
      x: xPos, y: 1.55, w: 2.7, h: 0.5,
      fontSize: 18, fontFace: "Georgia", color: r.color, bold: true, align: "center",
    });
    s.addText(r.key, {
      x: xPos + 0.1, y: 2.15, w: 2.5, h: 1.6,
      fontSize: 13, fontFace: "Calibri", color: C.black, align: "center",
    });
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: 0.5, y: 4.2, w: 8.8, h: 0.65, rectRadius: 0.12,
    fill: { color: C.lightGold },
  });
  s.addText("Oral Review: 告诉你的同桌，你记住了哪些？", {
    x: 0.5, y: 4.2, w: 8.8, h: 0.65,
    fontSize: 17, fontFace: "Calibri", color: C.dark, bold: true, align: "center",
  });
}


// ══════════════════════════════════════════════════
// Slide 22 — 欧洲三国对比表
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "欧洲三国对比表 Comparison Table");

  // Table header row
  var tblY = 0.85;
  var rowH = 0.55;
  var col0W = 1.8;
  var colW = 2.6;
  var tblX = 0.3;

  // Header cells
  var headers = ["", "法国 France", "意大利 Italy", "英国 UK"];
  var headerColors = [C.primary, C.france, C.italy, C.uk];

  headers.forEach(function(h, i) {
    var xPos = tblX + (i === 0 ? 0 : col0W + (i - 1) * colW);
    var w = i === 0 ? col0W : colW;
    s.addShape(pres.ShapeType.rect, {
      x: xPos, y: tblY, w: w, h: rowH,
      fill: { color: headerColors[i] },
    });
    s.addText(h, {
      x: xPos, y: tblY, w: w, h: rowH,
      fontSize: 13, fontFace: "Georgia", color: C.white, bold: true, align: "center", valign: "middle",
    });
  });

  var rows = [
    ["打招呼", "亲脸颊+Bonjour", "Ciao!+手势", "握手+Hello"],
    ["吃饭特点", "午餐1-2小时", "不加菠萝在披萨!", "下午茶"],
    ["代表食物", "可颂/奶酪", "披萨/意面", "炸鱼薯条"],
    ["注意事项", "不催服务员", "不加奶在咖啡", "要排队!"],
    ["特别文化", "时尚/卢浮宫", "文艺复兴", "哈利波特/王室"],
  ];

  rows.forEach(function(row, ri) {
    var yPos = tblY + rowH + ri * rowH;
    var bgColor = ri % 2 === 0 ? C.white : C.lightGold;

    row.forEach(function(cell, ci) {
      var xPos = tblX + (ci === 0 ? 0 : col0W + (ci - 1) * colW);
      var w = ci === 0 ? col0W : colW;
      s.addShape(pres.ShapeType.rect, {
        x: xPos, y: yPos, w: w, h: rowH,
        fill: { color: ci === 0 ? C.primary : bgColor },
        line: { color: "BDBDBD", width: 0.5 },
      });
      s.addText(cell, {
        x: xPos, y: yPos, w: w, h: rowH,
        fontSize: 11, fontFace: "Calibri",
        color: ci === 0 ? C.white : C.dark,
        bold: ci === 0, align: "center", valign: "middle",
      });
    });
  });

  footer(s, "三个国家，三种风格，都很精彩！");
}


// ══════════════════════════════════════════════════
// Slide 23 — 共同点与不同
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "共同点与不同 Similarities & Differences");

  // Similarities card
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.3, y: 0.85, w: 4.5, h: 3.0, rectRadius: 0.15,
    fill: { color: C.white }, line: { color: C.green, width: 2 },
  });
  s.addText([
    { text: "共同点 Similarities", options: { fontSize: 18, fontFace: "Georgia", color: C.green, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "都重视用餐礼仪", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "都有悠久历史", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "都有世界级地标", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "都有独特的美食文化", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
  ], { x: 0.5, y: 0.95, w: 4.1, h: 2.8, valign: "top" });

  // Differences card
  s.addShape(pres.ShapeType.roundRect, {
    x: 5.0, y: 0.85, w: 4.6, h: 3.0, rectRadius: 0.15,
    fill: { color: C.white }, line: { color: C.primary, width: 2 },
  });
  s.addText([
    { text: "不同 Differences", options: { fontSize: 18, fontFace: "Georgia", color: C.primary, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "打招呼方式不同", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "  法:亲脸颊 / 意:手势 / 英:握手", options: { fontSize: 12, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "饮食风格不同", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "  法:精致 / 意:热情 / 英:传统", options: { fontSize: 12, fontFace: "Calibri", color: C.gray, breakLine: true } },
    { text: "性格特点不同", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "  法:浪漫 / 意:热情 / 英:绅士", options: { fontSize: 12, fontFace: "Calibri", color: C.gray, breakLine: true } },
  ], { x: 5.2, y: 0.95, w: 4.2, h: 2.8, valign: "top" });

  // Bottom
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.5, y: 4.1, w: 8.8, h: 0.7, rectRadius: 0.12,
    fill: { color: C.lightGold }, line: { color: C.accent, width: 1.5 },
  });
  s.addText("不同文化，不同方式，都很精彩！Different cultures, all wonderful!", {
    x: 0.5, y: 4.1, w: 8.8, h: 0.7,
    fontSize: 15, fontFace: "Calibri", color: C.dark, bold: true, align: "center",
  });
}


// ══════════════════════════════════════════════════
// Slide 24 — 旅行小贴士 + 丝绸之路简介
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "旅行小贴士 + 丝绸之路 Silk Road");

  // Travel tips left
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.3, y: 0.85, w: 4.5, h: 3.5, rectRadius: 0.12,
    fill: { color: C.white }, line: { color: C.primary, width: 2 },
  });
  s.addText([
    { text: "旅行小贴士 Travel Tips", options: { fontSize: 16, fontFace: "Georgia", color: C.primary, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "在法国说 Bonjour 再开口说话", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "在意大利不要催服务员", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "在英国一定要排队", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "欧洲用欧元(法/意)和英镑(英)", options: { fontSize: 12, fontFace: "Calibri", color: C.black, bullet: true, breakLine: true } },
    { text: "尊重当地文化，入乡随俗！", options: { fontSize: 13, fontFace: "Calibri", color: C.primary, bold: true, bullet: true, breakLine: true } },
  ], { x: 0.5, y: 0.95, w: 4.1, h: 3.3, valign: "top" });

  // Silk Road right
  s.addShape(pres.ShapeType.roundRect, {
    x: 5.0, y: 0.85, w: 4.6, h: 3.5, rectRadius: 0.12,
    fill: { color: C.lightGold }, line: { color: C.accent, width: 2 },
  });
  s.addText([
    { text: "丝绸之路 Silk Road", options: { fontSize: 16, fontFace: "Georgia", color: C.dark, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "连接中国和欧洲的古代贸易之路", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "中国 -> 丝绸、茶叶、瓷器 -> 欧洲", options: { fontSize: 13, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
    { text: "欧洲 -> 玻璃、宝石、香料 -> 中国", options: { fontSize: 13, fontFace: "Calibri", color: C.dark, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "马可波罗沿丝绸之路来到中国", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "在中国住了17年！", options: { fontSize: 13, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
    { text: "茶叶从中国传到英国!", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
  ], { x: 5.2, y: 0.95, w: 4.2, h: 3.3, valign: "top" });

  footer(s, "丝绸之路连接了中国和欧洲两千多年！");
}


// ══════════════════════════════════════════════════
// Slide 25 — 生词卡 Vocabulary
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "生词卡 Vocabulary Cards");

  const vocab = [
    { zh: "披萨", py: "pi sa", en: "pizza" },
    { zh: "面包", py: "mian bao", en: "bread" },
    { zh: "奶酪", py: "nai lao", en: "cheese" },
    { zh: "冰淇淋", py: "bing qi lin", en: "ice cream" },
    { zh: "铁塔", py: "tie ta", en: "tower" },
    { zh: "城堡", py: "cheng bao", en: "castle" },
    { zh: "排队", py: "pai dui", en: "queue" },
    { zh: "礼节", py: "li jie", en: "etiquette" },
  ];

  vocab.forEach(function(v, i) {
    var col = i % 4;
    var row = Math.floor(i / 4);
    var xPos = 0.3 + col * 2.35;
    var yPos = 0.85 + row * 2.1;

    s.addShape(pres.ShapeType.roundRect, {
      x: xPos, y: yPos, w: 2.15, h: 1.85, rectRadius: 0.12,
      fill: { color: C.white },
      shadow: { type: "outer", blur: 4, offset: 2, color: "000000", opacity: 0.12 },
    });
    s.addText(v.zh, {
      x: xPos, y: yPos + 0.15, w: 2.15, h: 0.5,
      fontSize: 22, fontFace: "Georgia", color: C.dark, bold: true, align: "center",
    });
    s.addText(v.py, {
      x: xPos, y: yPos + 0.7, w: 2.15, h: 0.35,
      fontSize: 12, fontFace: "Calibri", color: C.primary, italic: true, align: "center",
    });
    s.addText(v.en, {
      x: xPos, y: yPos + 1.1, w: 2.15, h: 0.35,
      fontSize: 13, fontFace: "Calibri", color: C.gray, align: "center",
    });
  });
}


// ══════════════════════════════════════════════════
// Slide 26 — 句型练习 Sentence Patterns
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "句型练习 Sentence Patterns");

  // Pattern 1
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.5, y: 0.9, w: 8.8, h: 1.9, rectRadius: 0.15,
    fill: { color: C.white }, line: { color: C.primary, width: 2 },
  });
  s.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 0.9, w: 8.8, h: 0.4, fill: { color: C.primary },
  });
  s.addText("句型 1:「我想去___看___」", {
    x: 0.5, y: 0.9, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.white, bold: true, align: "center",
  });
  s.addText([
    { text: "我想去法国看埃菲尔铁塔。", options: { fontSize: 15, fontFace: "Calibri", color: C.dark, bold: true, breakLine: true } },
    { text: "我想去意大利看斗兽场。", options: { fontSize: 15, fontFace: "Calibri", color: C.dark, bold: true, breakLine: true } },
    { text: "我想去英国看大本钟。", options: { fontSize: 15, fontFace: "Calibri", color: C.dark, bold: true, breakLine: true } },
  ], { x: 1.0, y: 1.4, w: 7.8, h: 1.3, valign: "top" });

  // Pattern 2
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.5, y: 3.0, w: 8.8, h: 1.9, rectRadius: 0.15,
    fill: { color: C.white }, line: { color: C.accent, width: 2 },
  });
  s.addShape(pres.ShapeType.rect, {
    x: 0.5, y: 3.0, w: 8.8, h: 0.4, fill: { color: C.accent },
  });
  s.addText("句型 2:「___最有名的是___」", {
    x: 0.5, y: 3.0, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: "Georgia", color: C.dark, bold: true, align: "center",
  });
  s.addText([
    { text: "法国最有名的是埃菲尔铁塔。", options: { fontSize: 15, fontFace: "Calibri", color: C.dark, bold: true, breakLine: true } },
    { text: "意大利最有名的是披萨。", options: { fontSize: 15, fontFace: "Calibri", color: C.dark, bold: true, breakLine: true } },
    { text: "英国最有名的是下午茶。", options: { fontSize: 15, fontFace: "Calibri", color: C.dark, bold: true, breakLine: true } },
  ], { x: 1.0, y: 3.5, w: 7.8, h: 1.3, valign: "top" });
}


// ══════════════════════════════════════════════════
// Slide 27 — Role Play: 欧洲小厨师
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "Role Play:「欧洲小厨师」(10-15 min)", C.accent);

  // Role A
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.3, y: 0.85, w: 4.5, h: 2.4, rectRadius: 0.15,
    fill: { color: C.white }, line: { color: C.primary, width: 2 },
  });
  s.addShape(pres.ShapeType.rect, {
    x: 0.3, y: 0.85, w: 4.5, h: 0.4, fill: { color: C.primary },
  });
  s.addText("A = 厨师 Chef", {
    x: 0.3, y: 0.85, w: 4.5, h: 0.4,
    fontSize: 14, fontFace: "Georgia", color: C.white, bold: true, align: "center",
  });
  s.addText([
    { text: "介绍你的菜：", options: { fontSize: 14, fontFace: "Calibri", color: C.dark, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 3, breakLine: true } },
    { text: "「这是法国的可颂面包，", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "  很好吃！又香又甜！」", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "「这是意大利的披萨，", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "  上面有奶酪和番茄！」", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
  ], { x: 0.5, y: 1.35, w: 4.1, h: 1.8, valign: "top" });

  // Role B
  s.addShape(pres.ShapeType.roundRect, {
    x: 5.0, y: 0.85, w: 4.6, h: 2.4, rectRadius: 0.15,
    fill: { color: C.white }, line: { color: C.accent, width: 2 },
  });
  s.addShape(pres.ShapeType.rect, {
    x: 5.0, y: 0.85, w: 4.6, h: 0.4, fill: { color: C.accent },
  });
  s.addText("B = 美食评论家 Critic", {
    x: 5.0, y: 0.85, w: 4.6, h: 0.4,
    fontSize: 14, fontFace: "Georgia", color: C.dark, bold: true, align: "center",
  });
  s.addText([
    { text: "评价这道菜：", options: { fontSize: 14, fontFace: "Calibri", color: C.dark, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 3, breakLine: true } },
    { text: "「我觉得可颂很好吃！」", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "「我想吃更多披萨！」", options: { fontSize: 12, fontFace: "Calibri", color: C.black, breakLine: true } },
    { text: "「___最有名的是___」", options: { fontSize: 12, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
    { text: "「我想去___看___」", options: { fontSize: 12, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
  ], { x: 5.2, y: 1.35, w: 4.2, h: 1.8, valign: "top" });

  // Cultural experience
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.5, y: 3.5, w: 8.8, h: 1.3, rectRadius: 0.12,
    fill: { color: C.lightGold }, line: { color: C.accent, width: 2 },
  });
  s.addText([
    { text: "文化体验 Cultural Experience", options: { fontSize: 16, fontFace: "Georgia", color: C.dark, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "模拟用刀叉吃饭  vs  用筷子吃饭", options: { fontSize: 14, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
    { text: "哪个更方便？你更喜欢哪个？和同学讨论！", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
  ], { x: 0.7, y: 3.55, w: 8.4, h: 1.2, align: "center" });
}


// ══════════════════════════════════════════════════
// Slide 28 — Project Time! 护照欧洲页
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "Project Time! 护照欧洲页", C.accent);

  const sections = [
    { title: "Where I went", desc: "我去了___\n(法国/意大利/英国)", color: C.primary },
    { title: "What I saw", desc: "我看到了___\n(铁塔/斗兽场/大本钟)", color: C.dark },
    { title: "What I ate", desc: "我吃了___\n(披萨/可颂/炸鱼薯条)", color: C.accent },
    { title: "Etiquette", desc: "礼节：\n(排队/亲脸颊/手势)", color: C.france },
    { title: "My Sentence", desc: "用句型写一句话：\n「我想去___看___」", color: C.green },
  ];

  sections.forEach(function(sec, i) {
    var col = i < 3 ? i : i - 3;
    var row = i < 3 ? 0 : 1;
    var xPos = row === 0 ? 0.3 + col * 3.15 : 1.5 + col * 3.4;
    var yPos = 0.9 + row * 2.05;
    var cardW = row === 0 ? 2.95 : 3.2;

    s.addShape(pres.ShapeType.roundRect, {
      x: xPos, y: yPos, w: cardW, h: 1.8, rectRadius: 0.12,
      fill: { color: C.white }, line: { color: sec.color, width: 2 },
    });
    s.addShape(pres.ShapeType.ellipse, {
      x: xPos + (cardW / 2 - 0.3), y: yPos + 0.1, w: 0.6, h: 0.6,
      fill: { color: sec.color, transparency: 25 },
    });
    s.addText(sec.title, {
      x: xPos, y: yPos + 0.7, w: cardW, h: 0.3,
      fontSize: 13, fontFace: "Georgia", color: sec.color, bold: true, align: "center",
    });
    s.addText(sec.desc, {
      x: xPos + 0.1, y: yPos + 1.05, w: cardW - 0.2, h: 0.65,
      fontSize: 11, fontFace: "Calibri", color: C.black, align: "center",
    });
  });

  s.addText("15-20 minutes  /  用中文写和画！Write & draw in Chinese!", {
    x: 0.5, y: 5.0, w: 9.0, h: 0.32,
    fontSize: 13, fontFace: "Calibri", color: C.primary, bold: true, align: "center",
  });
}


// ══════════════════════════════════════════════════
// Slide 29 — Project 分层
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "Project 分层 Differentiated Levels");

  // Level 1
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.4, y: 0.85, w: 8.8, h: 1.2, rectRadius: 0.12,
    fill: { color: C.white }, line: { color: C.green, width: 2 },
  });
  s.addText([
    { text: "Level 1（K-1年级）", options: { fontSize: 16, fontFace: "Georgia", color: C.green, bold: true, breakLine: true } },
    { text: "画3个国家的地标 + 写国家名字 + 写一种食物", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
  ], { x: 0.6, y: 0.95, w: 8.4, h: 1.0, valign: "top" });

  // Level 2
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.4, y: 2.2, w: 8.8, h: 1.2, rectRadius: 0.12,
    fill: { color: C.white }, line: { color: C.primary, width: 2 },
  });
  s.addText([
    { text: "Level 2（2-3年级）", options: { fontSize: 16, fontFace: "Georgia", color: C.primary, bold: true, breakLine: true } },
    { text: "Level 1 + 用句型写2句话 + 写一个文化发现", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
  ], { x: 0.6, y: 2.3, w: 8.4, h: 1.0, valign: "top" });

  // Level 3
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.4, y: 3.55, w: 8.8, h: 1.2, rectRadius: 0.12,
    fill: { color: C.white }, line: { color: C.accent, width: 2 },
  });
  s.addText([
    { text: "Level 3（4-5年级）", options: { fontSize: 16, fontFace: "Georgia", color: C.accent, bold: true, breakLine: true } },
    { text: "Level 2 + 对比三国文化差异 + 写丝绸之路小短文", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
  ], { x: 0.6, y: 3.65, w: 8.4, h: 1.0, valign: "top" });
}


// ══════════════════════════════════════════════════
// Slide 30 — 分享 Share Time
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.secondary };
  headerBar(s, "分享时间 Share Time!");

  s.addShape(pres.ShapeType.roundRect, {
    x: 0.5, y: 0.9, w: 8.8, h: 3.2, rectRadius: 0.15,
    fill: { color: C.white },
    shadow: { type: "outer", blur: 4, offset: 2, color: "000000", opacity: 0.15 },
  });

  s.addText([
    { text: "向全班展示你的护照欧洲页！", options: { fontSize: 20, fontFace: "Georgia", color: C.dark, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "用中文说:", options: { fontSize: 16, fontFace: "Calibri", color: C.primary, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "「我去了___，看到了___」", options: { fontSize: 16, fontFace: "Calibri", color: C.dark, breakLine: true } },
    { text: "「___最有名的是___」", options: { fontSize: 16, fontFace: "Calibri", color: C.dark, breakLine: true } },
    { text: "「我学到了一个礼节：___」", options: { fontSize: 16, fontFace: "Calibri", color: C.dark, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "每人30秒-1分钟！大声说！", options: { fontSize: 14, fontFace: "Calibri", color: C.gray, italic: true, breakLine: true } },
  ], { x: 0.8, y: 1.0, w: 8.2, h: 3.0, align: "center" });

  footer(s);
}


// ══════════════════════════════════════════════════
// Slide 31 — 签证章 Stamp
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  goldBars(s);

  s.addText("恭喜！获得欧洲签证章！", {
    x: 0.5, y: 0.3, w: 9.0, h: 0.65,
    fontSize: 28, fontFace: "Georgia", color: C.accent, bold: true, align: "center",
  });

  // Stamp circle
  s.addShape(pres.ShapeType.ellipse, {
    x: 3.2, y: 1.15, w: 3.4, h: 3.4,
    fill: { color: C.primary, transparency: 30 },
    line: { color: C.accent, width: 4 },
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: 3.4, y: 1.35, w: 3.0, h: 3.0,
    line: { color: C.accent, width: 2, dashType: "dash" },
  });
  s.addText([
    { text: "EUROPE", options: { fontSize: 28, fontFace: "Georgia", color: C.accent, bold: true, breakLine: true } },
    { text: "欧洲", options: { fontSize: 20, fontFace: "Georgia", color: C.white, bold: true, breakLine: true } },
    { text: "6/10", options: { fontSize: 18, fontFace: "Georgia", color: C.accent, breakLine: true } },
    { text: "STAMP 3 of 5", options: { fontSize: 12, fontFace: "Calibri", color: C.white, breakLine: true } },
  ], { x: 3.2, y: 1.55, w: 3.4, h: 2.6, align: "center" });

  s.addShape(pres.ShapeType.roundRect, {
    x: 1.5, y: 4.7, w: 6.8, h: 0.55, rectRadius: 0.15,
    fill: { color: C.accent },
  });
  s.addText("3个签证章！还有2个！ Asia / Africa / Europe > Americas > Showcase", {
    x: 1.5, y: 4.7, w: 6.8, h: 0.55,
    fontSize: 13, fontFace: "Calibri", color: C.dark, bold: true, align: "center",
  });
}


// ══════════════════════════════════════════════════
// Slide 32 — 明天预告
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  goldBars(s);

  s.addText("明天预告 Coming Tomorrow...", {
    x: 0.5, y: 0.3, w: 9.0, h: 0.65,
    fontSize: 30, fontFace: "Georgia", color: C.accent, bold: true, align: "center",
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: 1.8, y: 1.2, w: 6.2, h: 2.5, rectRadius: 0.2,
    fill: { color: C.primary, transparency: 30 },
    line: { color: C.accent, width: 2 },
  });

  s.addText([
    { text: "GR-004", options: { fontSize: 18, fontFace: "Calibri", color: C.accent, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "目的地：美洲 AMERICAS", options: { fontSize: 28, fontFace: "Georgia", color: C.white, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "从北美到南美，新大陆等你探索！", options: { fontSize: 17, fontFace: "Calibri", color: C.white, breakLine: true } },
    { text: "汉堡包 / 墨西哥卷饼 / 玉米...", options: { fontSize: 15, fontFace: "Calibri", color: C.accent, breakLine: true } },
  ], { x: 2.0, y: 1.4, w: 5.8, h: 2.1, align: "center" });

  s.addShape(pres.ShapeType.roundRect, {
    x: 2.5, y: 4.0, w: 4.8, h: 0.55, rectRadius: 0.15,
    fill: { color: C.accent },
  });
  s.addText("准备好你的护照，我们继续出发！", {
    x: 2.5, y: 4.0, w: 4.8, h: 0.55,
    fontSize: 17, fontFace: "Calibri", color: C.dark, bold: true, align: "center",
  });

  footer(s, "GR EDU  |  Global Explorer Camp 2025");
}


// ══════════════════════════════════════════════════
// Slide 33 — Thank You / End
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  goldBars(s);

  s.addText([
    { text: "Day 3 Complete!", options: { fontSize: 32, fontFace: "Georgia", color: C.accent, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "第三天完成！", options: { fontSize: 24, fontFace: "Georgia", color: C.white, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "法国 + 意大利 + 英国", options: { fontSize: 20, fontFace: "Calibri", color: C.accent, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "3/5 签证章已收集！", options: { fontSize: 18, fontFace: "Calibri", color: C.white, breakLine: true } },
  ], { x: 0.5, y: 1.0, w: 9.0, h: 3.5, align: "center" });

  footer(s, "谷雨中文 GR EDU  |  Global Explorer Camp 2025");
}


// ── Save ──
const outPath = path.join(__dirname, "day3_europe.pptx");
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
