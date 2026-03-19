/**
 * Day 4: 美洲 Americas (6/11) — Global Explorer Camp 环球探索沉浸式夏令营
 * 3 countries: USA, Mexico, Brazil — each with Overview/Culture/Food/Check
 * Run: node create_day4.js
 */
const pptxgen = require("pptxgenjs");
const https = require("https");
const http = require("http");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_16x9"; // 10.0" x 5.625"
pptx.author = "谷雨中文 GR EDU";
pptx.title = "Global Explorer Camp · Day 4: 美洲 Americas";

// ── Color palette (no # prefix) ──
const C = {
  pri:     "2E7D32",
  sec:     "E8F5E9",
  accent:  "FF6F00",
  dark:    "1B5E20",
  white:   "FFFFFF",
  black:   "333333",
  midGrn:  "4CAF50",
  paleGrn: "F1F8E9",
  ltOrg:   "FFF3E0",
  gray:    "616161",
  usa:     "3C3B6E",
  mex:     "006847",
  bra:     "009C3B",
  quiz:    "F57F17",
  stamp:   "BF360C",
};

// ── Image URLs ──
const IMG_URLS = {
  liberty:     "https://upload.wikimedia.org/wikipedia/commons/thumb/a/a1/Statue_of_Liberty_7.jpg/800px-Statue_of_Liberty_7.jpg",
  machuPicchu: "https://upload.wikimedia.org/wikipedia/commons/thumb/e/eb/Machu_Picchu%2C_Peru.jpg/1280px-Machu_Picchu%2C_Peru.jpg",
  chichenItza: "https://upload.wikimedia.org/wikipedia/commons/thumb/4/4e/Chichen_Itza_3.jpg/1280px-Chichen_Itza_3.jpg",
};

// ── Download image to base64, follow redirects ──
function fetchImageBase64(url, maxRedirects) {
  if (maxRedirects === undefined) maxRedirects = 5;
  return new Promise((resolve) => {
    if (maxRedirects <= 0) { resolve(null); return; }
    const proto = url.startsWith("https") ? https : http;
    proto.get(url, { headers: { "User-Agent": "Mozilla/5.0" } }, (res) => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        resolve(fetchImageBase64(res.headers.location, maxRedirects - 1));
        return;
      }
      if (res.statusCode !== 200) {
        console.warn("  WARN: HTTP " + res.statusCode + " for " + url.slice(0, 80));
        res.resume();
        resolve(null);
        return;
      }
      const chunks = [];
      res.on("data", (c) => chunks.push(c));
      res.on("end", () => {
        const buf = Buffer.concat(chunks);
        const ct = res.headers["content-type"] || "image/jpeg";
        const mime = ct.split(";")[0].trim();
        resolve("data:" + mime + ";base64," + buf.toString("base64"));
      });
      res.on("error", () => resolve(null));
    }).on("error", () => resolve(null));
  });
}

// ── Helpers ──
function greenBg(s) { s.background = { color: C.sec }; }
function darkBg(s)  { s.background = { color: C.dark }; }
function slideNum(s) { s.slideNumber = { x: "95%", y: "95%", fontSize: 8, color: C.gray }; }

function topBar(s, clr) {
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: 9.8, h: 0.07, fill: { color: clr || C.accent },
  });
}
function bottomBar(s, clr) {
  s.addShape(pptx.shapes.RECTANGLE, {
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
  s.addShape(pptx.shapes.RECTANGLE, {
    x: x || 0.4, y: y || 0.8, w: w || 4.0, h: 0.05, fill: { color: C.accent },
  });
}

function card(s, x, y, w, h, fill, line) {
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: w, h: h, rectRadius: 0.12,
    fill: { color: fill || C.white },
    line: line ? { color: line, width: 2 } : undefined,
    shadow: { type: "outer", blur: 4, offset: 2, color: "999999", opacity: 0.25 },
  });
}

function cardHeader(s, x, y, w, h, fill, text) {
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: w, h: h, rectRadius: 0.1,
    fill: { color: fill },
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: x, y: y + h * 0.5, w: w, h: h * 0.5, fill: { color: fill },
  });
  s.addText(text, {
    x: x, y: y, w: w, h: h,
    fontFace: "Georgia", fontSize: 16, color: C.white,
    bold: true, align: "center", valign: "middle",
  });
}

function body(s, text, opts) {
  s.addText(text, Object.assign({
    fontFace: "Calibri", fontSize: 13, color: C.black,
    valign: "top", breakLine: true, lineSpacing: 19,
  }, opts));
}

function interactiveBar(s, text, y) {
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.8, y: y || 4.85, w: 8.2, h: 0.42, fill: { color: C.dark }, rectRadius: 0.08,
  });
  s.addText(text, {
    x: 0.8, y: y || 4.85, w: 8.2, h: 0.42,
    fontFace: "Calibri", fontSize: 13, color: C.accent,
    bold: true, align: "center", valign: "middle",
  });
}

function addPhoto(s, imgData, x, y, w, h, fallbackColor, fallbackLabel) {
  if (imgData) {
    s.addImage({ data: imgData, x: x, y: y, w: w, h: h, rounding: true });
  } else {
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: x, y: y, w: w, h: h, rectRadius: 0.15,
      fill: { color: fallbackColor || C.midGrn },
    });
    s.addText(fallbackLabel || "Photo", {
      x: x, y: y, w: w, h: h,
      fontFace: "Georgia", fontSize: 18, color: C.white,
      bold: true, align: "center", valign: "middle",
    });
  }
}

// ══════════════════════════════════════════════════════════════
// Main
// ══════════════════════════════════════════════════════════════
async function buildPresentation() {
  console.log("Downloading images...");
  const keys = Object.keys(IMG_URLS);
  const results = await Promise.all(
    keys.map((k) => {
      console.log("  Fetching: " + k);
      return fetchImageBase64(IMG_URLS[k]);
    })
  );
  const IMG = {};
  keys.forEach((k, i) => {
    IMG[k] = results[i];
    if (!results[i]) console.warn("  FALLBACK: " + k);
    else console.log("  OK: " + k + " (" + Math.round(results[i].length / 1024) + " KB)");
  });

  // ══════════════════════════════════════════════════════════════
  // SLIDE 1 — Boarding Time (GR-004, green bg)
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    darkBg(s);
    topBar(s);
    bottomBar(s);

    s.addText("Global Explorer Camp", {
      x: 0.5, y: 0.3, w: 9.0, h: 0.7,
      fontFace: "Georgia", fontSize: 22, color: C.accent, italic: true, align: "center",
    });
    s.addText("环球探索沉浸式夏令营", {
      x: 0.5, y: 0.95, w: 9.0, h: 0.7,
      fontFace: "Georgia", fontSize: 34, color: C.white, bold: true, align: "center",
    });

    s.addShape(pptx.shapes.RECTANGLE, {
      x: 2.5, y: 1.75, w: 5.0, h: 0.04, fill: { color: C.accent },
    });

    s.addText([
      { text: "航班 GR-004", options: { fontSize: 20, fontFace: "Calibri", color: C.accent, bold: true, breakLine: true } },
      { text: "美洲 AMERICAS \u00B7 6/11 周四", options: { fontSize: 18, fontFace: "Calibri", color: C.white, breakLine: true } },
    ], { x: 1.0, y: 1.9, w: 8.0, h: 0.8, align: "center" });

    card(s, 2.0, 2.8, 6.0, 1.6, C.white);
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 2.0, y: 2.8, w: 6.0, h: 0.42, rectRadius: 0.12,
      fill: { color: C.accent },
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 2.0, y: 3.0, w: 6.0, h: 0.22, fill: { color: C.accent },
    });
    s.addText("BOARDING PASS  登机牌", {
      x: 2.0, y: 2.8, w: 6.0, h: 0.42,
      fontFace: "Georgia", fontSize: 14, color: C.white,
      bold: true, align: "center", valign: "middle",
    });

    s.addText([
      { text: "Flight 航班", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
      { text: "GR-004", options: { fontSize: 14, fontFace: "Calibri", color: C.black, bold: true, breakLine: true } },
    ], { x: 2.3, y: 3.3, w: 1.4, h: 0.9, align: "center", valign: "top" });

    s.addText([
      { text: "From 出发", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
      { text: "欧洲 Europe", options: { fontSize: 13, fontFace: "Calibri", color: C.black, breakLine: true } },
    ], { x: 3.9, y: 3.3, w: 1.6, h: 0.9, align: "center", valign: "top" });

    s.addText("✈️", { x: 5.5, y: 3.4, w: 0.6, h: 0.5, fontSize: 20, align: "center" });

    s.addText([
      { text: "To 目的地", options: { fontSize: 11, fontFace: "Calibri", color: C.gray, breakLine: true } },
      { text: "美洲 Americas", options: { fontSize: 14, fontFace: "Calibri", color: C.accent, bold: true, breakLine: true } },
    ], { x: 6.1, y: 3.3, w: 1.6, h: 0.9, align: "center", valign: "top" });

    s.addText("🇺🇸  🇲🇽  🇧🇷", {
      x: 2.5, y: 4.8, w: 5.0, h: 0.35,
      fontSize: 18, align: "center",
    });
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 2 — 护照进度 (3 stamps, last one!)
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "🗺️ 护照进度  Passport Progress");
    accentLine(s);

    const items = [
      { lbl: "✅ 亚洲\nAsia",     fill: C.pri },
      { lbl: "✅ 非洲\nAfrica",   fill: C.pri },
      { lbl: "✅ 欧洲\nEurope",   fill: C.pri },
      { lbl: "📍 美洲\nAmericas", fill: C.accent },
      { lbl: "🎉 文化展\nExhibition", fill: "BDBDBD" },
    ];
    const bw = 1.6, gap = 0.2, sx = 0.5, by = 1.4;
    items.forEach((it, i) => {
      const bx = sx + i * (bw + gap);
      s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
        x: bx, y: by, w: bw, h: 1.15, rectRadius: 0.12,
        fill: { color: it.fill },
        shadow: { type: "outer", blur: 3, offset: 2, color: "999999", opacity: 0.25 },
      });
      s.addText(it.lbl, {
        x: bx, y: by, w: bw, h: 1.15,
        fontFace: "Calibri", fontSize: 14, color: C.white,
        bold: true, align: "center", valign: "middle", breakLine: true,
      });
      if (i < items.length - 1) {
        s.addText("→", {
          x: bx + bw - 0.05, y: by + 0.3, w: gap + 0.1, h: 0.5,
          fontSize: 20, color: C.dark, align: "center", bold: true,
        });
      }
    });

    card(s, 1.0, 3.1, 8.0, 0.9, C.white, C.accent);
    s.addText([
      { text: "🎒 3个签证章！最后一个！", options: { fontSize: 18, fontFace: "Calibri", color: C.dark, bold: true, breakLine: true } },
      { text: "「集齐签证，明天展示你的护照！」", options: { fontSize: 16, fontFace: "Calibri", color: C.accent, breakLine: true } },
    ], { x: 1.0, y: 3.1, w: 8.0, h: 0.9, align: "center", valign: "middle" });

    s.addText("💪 今天拿到最后一个签证章，明天就是文化展！", {
      x: 1.0, y: 4.2, w: 8.0, h: 0.5,
      fontFace: "Calibri", fontSize: 15, color: C.pri, align: "center", italic: true,
    });
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 3 — 今天目标
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "📋 今天目标  Today's Goals");
    accentLine(s);

    card(s, 0.4, 1.1, 4.4, 2.0, C.white, C.pri);
    cardHeader(s, 0.4, 1.1, 4.4, 0.5, C.pri, "☀️ 上午 Morning");
    body(s, [
      "1. 认识美洲大陆 (北美+南美)",
      "2. 深入了解美国、墨西哥、巴西",
      "3. 每国: 概览→文化→美食→检查",
      "4. 迷你角色扮演 + 竞赛",
    ].join("\n"), { x: 0.6, y: 1.7, w: 4.0, h: 1.3 });

    card(s, 5.2, 1.1, 4.4, 2.0, C.white, C.accent);
    cardHeader(s, 5.2, 1.1, 4.4, 0.5, C.accent, "🌙 下午 Afternoon");
    body(s, [
      "1. 三国对比表 + 共同点/不同",
      "2. 生词/句型/Role Play",
      "3. 完成护照美洲页",
      "4. 准备明天展览 + 签证章",
    ].join("\n"), { x: 5.4, y: 1.7, w: 4.0, h: 1.3 });

    card(s, 1.5, 3.5, 7.0, 0.8, C.dark);
    s.addText("🌎 这是最后一站！完成后，明天就展示你的完整护照！", {
      x: 1.5, y: 3.5, w: 7.0, h: 0.8,
      fontFace: "Calibri", fontSize: 18, color: C.accent,
      bold: true, align: "center", valign: "middle",
    });
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 4 — 认识美洲
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "🌎 认识美洲  Meet the Americas");
    accentLine(s);

    card(s, 0.4, 1.0, 4.4, 1.8, C.white, C.pri);
    cardHeader(s, 0.4, 1.0, 4.4, 0.45, C.pri, "🌎 北美洲 North America");
    body(s, [
      "• 23个国家和地区",
      "• 美国、墨西哥、加拿大...",
      "• 世界最大经济体在这里",
    ].join("\n"), { x: 0.6, y: 1.55, w: 4.0, h: 1.15 });

    card(s, 5.2, 1.0, 4.4, 1.8, C.white, C.accent);
    cardHeader(s, 5.2, 1.0, 4.4, 0.45, C.accent, "🌎 南美洲 South America");
    body(s, [
      "• 12个国家",
      "• 巴西、阿根廷、秘鲁...",
      "• 亚马逊雨林 + 安第斯山脉",
    ].join("\n"), { x: 5.4, y: 1.55, w: 4.0, h: 1.15 });

    card(s, 0.4, 3.1, 9.2, 1.5, C.paleGrn, C.midGrn);
    body(s, [
      "🔑 美洲总共约35个国家",
      "🌳 亚马逊雨林 = 世界最大热带雨林",
      "⛰️ 安第斯山脉 = 世界最长山脉",
      "🌽 番茄、土豆、玉米、巧克力、辣椒 — 全部来自美洲！",
    ].join("\n"), { x: 0.6, y: 3.2, w: 8.8, h: 1.3, fontSize: 14, lineSpacing: 22 });

    interactiveBar(s, "🙋 你知道哪些食物是从美洲传到中国的？", 4.85);
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 5 — 🇺🇸 美国概览
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "🇺🇸 美国概览  USA Overview");
    accentLine(s);

    addPhoto(s, IMG.liberty, 0.3, 1.0, 4.6, 3.4, C.usa, "🗽 自由女神像\nStatue of Liberty");

    card(s, 5.1, 1.0, 4.5, 3.4, C.white, C.usa);
    cardHeader(s, 5.1, 1.0, 4.5, 0.45, C.usa, "🇺🇸 基本信息 Key Facts");
    body(s, [
      "🏳️ 国旗：星条旗",
      "   50颗星 = 50个州",
      "   13条纹 = 最初13个殖民地",
      "",
      "👥 人口：3.3亿",
      "🗣️ 语言：英语（没有官方语言！）",
      "🏛️ 首都：华盛顿DC",
      "   （不是纽约！）",
      "📍 50个州",
    ].join("\n"), { x: 5.3, y: 1.55, w: 4.1, h: 2.75, fontSize: 12, lineSpacing: 16 });

    interactiveBar(s, "🙋 美国居然没有官方语言！你知道吗？", 4.85);
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 6 — 🇺🇸 美国文化与礼节
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "🇺🇸 美国文化与礼节  USA Culture & Etiquette");
    accentLine(s);

    card(s, 0.3, 1.0, 4.6, 3.6, C.white, C.usa);
    cardHeader(s, 0.3, 1.0, 4.6, 0.45, C.usa, "🌟 文化亮点 Culture Highlights");
    body(s, [
      "🗽 自由女神是法国送的！",
      "🎬 好莱坞 Hollywood",
      "🚀 NASA太空探索",
      "🌍 移民国家 = 多元文化",
      "🏮 唐人街 Chinatown",
    ].join("\n"), { x: 0.5, y: 1.55, w: 4.2, h: 2.95, fontSize: 13, lineSpacing: 20 });

    card(s, 5.1, 1.0, 4.5, 3.6, C.ltOrg, C.accent);
    cardHeader(s, 5.1, 1.0, 4.5, 0.45, C.accent, "⚠️ 礼节 Etiquette");
    body(s, [
      "🤝 见面握手 + 微笑",
      "📏 保持个人空间（别站太近）",
      "💰 给小费！餐厅15-20%",
      "👋 直呼名字OK",
      "   （不像中国要叫「老师」「阿姨」）",
      "👟 鞋子可以穿进屋",
      "   （和日本相反！）",
    ].join("\n"), { x: 5.3, y: 1.55, w: 4.1, h: 2.95, fontSize: 12, lineSpacing: 17 });

    interactiveBar(s, "🙋 在美国餐厅不给小费会怎样？服务员会很不高兴！", 4.85);
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 7 — 🇺🇸 美国美食
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "🇺🇸 美国美食  USA Food");
    accentLine(s);

    const foods = [
      { emoji: "🍔", cn: "汉堡 hamburger", x: 0.3, y: 1.0 },
      { emoji: "🌭", cn: "热狗 hot dog", x: 2.5, y: 1.0 },
      { emoji: "🥧", cn: "苹果派 apple pie", x: 4.7, y: 1.0 },
      { emoji: "🦃", cn: "烤火鸡 roast turkey\n（感恩节）", x: 0.3, y: 2.3 },
      { emoji: "🥩", cn: "BBQ烧烤", x: 2.5, y: 2.3 },
    ];

    foods.forEach((f) => {
      card(s, f.x, f.y, 2.0, 1.1, C.white, C.midGrn);
      s.addText(f.emoji, {
        x: f.x, y: f.y + 0.05, w: 2.0, h: 0.5,
        fontSize: 28, align: "center",
      });
      s.addText(f.cn, {
        x: f.x + 0.1, y: f.y + 0.55, w: 1.8, h: 0.5,
        fontFace: "Calibri", fontSize: 11, color: C.dark, bold: true, align: "center", breakLine: true,
      });
    });

    card(s, 4.7, 2.3, 5.0, 1.1, C.ltOrg, C.accent);
    s.addText([
      { text: "🥤 美国发明了可口可乐！", options: { fontSize: 14, fontFace: "Calibri", color: C.dark, bold: true, breakLine: true } },
      { text: "「美国是移民国家，所以有全世界的食物！\n中餐在美国也超受欢迎」", options: { fontSize: 12, fontFace: "Calibri", color: C.pri, breakLine: true } },
    ], { x: 4.9, y: 2.35, w: 4.6, h: 1.0, valign: "middle" });

    card(s, 0.3, 3.6, 9.3, 0.95, C.dark);
    s.addText("「美国是一个移民大熔炉 melting pot — 你能在美国吃到全世界的食物！」", {
      x: 0.5, y: 3.65, w: 8.9, h: 0.85,
      fontFace: "Georgia", fontSize: 16, color: C.accent,
      bold: true, align: "center", valign: "middle",
    });

    interactiveBar(s, "🙋 你在美国最喜欢吃什么？", 4.85);
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 8-9 — ✅ 美国 Check Understanding (2 slides)
  // ══════════════════════════════════════════════════════════════
  (() => {
    const allQs = [
      { q: "美国首都是哪里？", a: "华盛顿DC（不是纽约！）" },
      { q: "餐厅要给多少小费？", a: "15-20%" },
      { q: "自由女神是谁送的？", a: "法国！" },
      { q: "美国有官方语言吗？", a: "没有！" },
      { q: "自由女神是谁送的？", a: "法国" },
      { q: "美国有官方语言吗？", a: "没有！" },
      { q: "在餐厅要给多少小费？", a: "15-20%" },
      { q: "美国有多少个州？", a: "50个" },
      { q: "唐人街(Chinatown)说明了什么？", a: "美国是移民国家" },
      { q: "在美国可以穿鞋进别人家吗？", a: "一般可以（和日本相反！）" },
    ];

    [0, 1].forEach((page) => {
      const s = pptx.addSlide();
      greenBg(s);
      topBar(s);

      hdr(s, "✅ 美国 Check Understanding (" + (page + 1) + "/2)");
      accentLine(s);

      var pageQs = allQs.slice(page * 5, page * 5 + 5);
      pageQs.forEach((item, i) => {
        var num = page * 5 + i + 1;
        var yy = 1.0 + i * 0.85;
        var rowBg = i % 2 === 0 ? C.white : C.paleGrn;

        s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
          x: 0.3, y: yy, w: 9.2, h: 0.72, rectRadius: 0.1,
          fill: { color: rowBg },
          line: { color: C.usa, width: 1 },
        });
        s.addShape(pptx.shapes.OVAL, {
          x: 0.45, y: yy + 0.13, w: 0.45, h: 0.45,
          fill: { color: C.usa },
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
          fontSize: 11, fontFace: "Calibri", color: C.pri,
          bold: true, align: "left", valign: "middle",
        });
      });

      interactiveBar(s, "🏆 全对的同学举手！All correct? Raise your hand!", 5.1);
      slideNum(s);
    });
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 9 — 🇲🇽 墨西哥概览
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "🇲🇽 墨西哥概览  Mexico Overview");
    accentLine(s);

    addPhoto(s, IMG.chichenItza, 0.3, 1.0, 4.6, 3.4, C.mex, "🏛️ 奇琴伊察\nChichen Itza");

    card(s, 5.1, 1.0, 4.5, 3.4, C.white, C.mex);
    cardHeader(s, 5.1, 1.0, 4.5, 0.45, C.mex, "🇲🇽 基本信息 Key Facts");
    body(s, [
      "🏳️ 国旗：绿白红",
      "   中间有鹰蛇仙人掌图案！",
      "",
      "👥 人口：1.3亿",
      "🗣️ 语言：西班牙语",
      "🏛️ 首都：墨西哥城",
      "",
      "🏗️ 玛雅 Maya 和",
      "   阿兹特克 Aztec 古文明",
    ].join("\n"), { x: 5.3, y: 1.55, w: 4.1, h: 2.75, fontSize: 12, lineSpacing: 16 });

    interactiveBar(s, "🙋 玛雅人发明了数字「0」！Maya invented zero!", 4.85);
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 10 — 🇲🇽 墨西哥文化与礼节
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "🇲🇽 墨西哥文化与礼节  Mexico Culture & Etiquette");
    accentLine(s);

    card(s, 0.3, 1.0, 4.6, 3.6, C.white, C.mex);
    cardHeader(s, 0.3, 1.0, 4.6, 0.45, C.mex, "🌟 文化亮点 Culture");
    body(s, [
      "💀 亡灵节 Day of the Dead",
      "   纪念去世亲人",
      "   （不是恐怖节日！）",
      "🎩 墨西哥帽 sombrero",
      "🌈 彩色街道 colorful streets",
    ].join("\n"), { x: 0.5, y: 1.55, w: 4.2, h: 2.95, fontSize: 13, lineSpacing: 20 });

    card(s, 5.1, 1.0, 4.5, 3.6, C.ltOrg, C.accent);
    cardHeader(s, 5.1, 1.0, 4.5, 0.45, C.accent, "⚠️ 礼节 Etiquette");
    body(s, [
      "👋 说 Hola!",
      "🤗 见面拥抱 + 亲脸颊",
      "⏰ 时间观念比较随意",
      "   （约9点可能9:30到）",
      "🍽️ 午餐是一天最重要的一餐",
      "   （2-4pm）",
      "😊 陌生人之间也很热情友好",
    ].join("\n"), { x: 5.3, y: 1.55, w: 4.1, h: 2.95, fontSize: 12, lineSpacing: 17 });

    interactiveBar(s, "🙋 亡灵节不是万圣节！是温暖地纪念亲人的节日", 4.85);
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 11 — 🇲🇽 墨西哥美食
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "🇲🇽 墨西哥美食  Mexico Food");
    accentLine(s);

    const foods = [
      { emoji: "🌮", cn: "玉米饼 taco", x: 0.3, y: 1.0 },
      { emoji: "🥑", cn: "牛油果酱\nguacamole", x: 2.5, y: 1.0 },
      { emoji: "🌯", cn: "墨西哥卷饼\nburrito", x: 4.7, y: 1.0 },
      { emoji: "🌶️", cn: "辣椒 sauce", x: 0.3, y: 2.3 },
      { emoji: "🍫", cn: "巧克力\n（阿兹特克人发明的！）", x: 2.5, y: 2.3 },
    ];

    foods.forEach((f) => {
      card(s, f.x, f.y, 2.0, 1.1, C.white, C.midGrn);
      s.addText(f.emoji, {
        x: f.x, y: f.y + 0.05, w: 2.0, h: 0.5,
        fontSize: 28, align: "center",
      });
      s.addText(f.cn, {
        x: f.x + 0.1, y: f.y + 0.55, w: 1.8, h: 0.5,
        fontFace: "Calibri", fontSize: 11, color: C.dark, bold: true, align: "center", breakLine: true,
      });
    });

    card(s, 4.7, 2.3, 5.0, 1.1, C.ltOrg, C.accent);
    s.addText([
      { text: "🌽 59种颜色的玉米！", options: { fontSize: 14, fontFace: "Calibri", color: C.dark, bold: true, breakLine: true } },
      { text: "「没有墨西哥就没有巧克力！」", options: { fontSize: 13, fontFace: "Georgia", color: C.accent, bold: true, breakLine: true } },
    ], { x: 4.9, y: 2.35, w: 4.6, h: 1.0, valign: "middle" });

    card(s, 0.3, 3.6, 9.3, 0.95, C.dark);
    s.addText("「墨西哥有59种颜色的玉米 — 不只是黄色！还有蓝色、红色、黑色...」", {
      x: 0.5, y: 3.65, w: 8.9, h: 0.85,
      fontFace: "Georgia", fontSize: 16, color: C.accent,
      bold: true, align: "center", valign: "middle",
    });

    interactiveBar(s, "🙋 你吃过墨西哥菜吗？最喜欢什么？", 4.85);
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 12-13 — ✅ 墨西哥 Check Understanding (2 slides)
  // ══════════════════════════════════════════════════════════════
  (() => {
    const allQs = [
      { q: "巧克力是谁发明的？", a: "阿兹特克人！" },
      { q: "墨西哥有多少种颜色的玉米？", a: "59种！" },
      { q: "亡灵节是恐怖的吗？", a: "不是！是纪念亲人的温暖节日" },
      { q: "墨西哥最有名的古迹是什么？", a: "奇琴伊察金字塔" },
      { q: "巧克力是谁发明的？", a: "阿兹特克人" },
      { q: "墨西哥有几种颜色的玉米？", a: "59种" },
      { q: "亡灵节是恐怖的节日吗？", a: "不是！是纪念亲人" },
      { q: "墨西哥人午餐是几点？", a: "下午2-4点" },
      { q: "墨西哥城建在什么上面？", a: "古代湖泊" },
      { q: "在墨西哥约了9点，几点到？", a: "可能9:30(时间观念随意)" },
    ];

    [0, 1].forEach((page) => {
      const s = pptx.addSlide();
      greenBg(s);
      topBar(s);

      hdr(s, "✅ 墨西哥 Check Understanding (" + (page + 1) + "/2)");
      accentLine(s);

      var pageQs = allQs.slice(page * 5, page * 5 + 5);
      pageQs.forEach((item, i) => {
        var num = page * 5 + i + 1;
        var yy = 1.0 + i * 0.85;
        var rowBg = i % 2 === 0 ? C.white : C.paleGrn;

        s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
          x: 0.3, y: yy, w: 9.2, h: 0.72, rectRadius: 0.1,
          fill: { color: rowBg },
          line: { color: C.mex, width: 1 },
        });
        s.addShape(pptx.shapes.OVAL, {
          x: 0.45, y: yy + 0.13, w: 0.45, h: 0.45,
          fill: { color: C.mex },
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
          fontSize: 11, fontFace: "Calibri", color: C.pri,
          bold: true, align: "left", valign: "middle",
        });
      });

      interactiveBar(s, "🏆 墨西哥小达人！Mexico expert!", 5.1);
      slideNum(s);
    });
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 13 — 🇧🇷 巴西概览
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "🇧🇷 巴西概览  Brazil Overview");
    accentLine(s);

    addPhoto(s, IMG.machuPicchu, 0.3, 1.0, 4.6, 3.4, C.bra, "🌳 南美洲风光\nSouth America");

    card(s, 5.1, 1.0, 4.5, 3.4, C.white, C.bra);
    cardHeader(s, 5.1, 1.0, 4.5, 0.45, C.bra, "🇧🇷 基本信息 Key Facts");
    body(s, [
      "🏳️ 国旗：绿底黄菱形",
      "   + 蓝色地球仪",
      "",
      "👥 人口：2.1亿",
      "🗣️ 语言：葡萄牙语",
      "   （不是西班牙语！）",
      "🏛️ 首都：巴西利亚",
      "   （不是里约！）",
      "🌳 亚马逊雨林 = 地球之肺",
    ].join("\n"), { x: 5.3, y: 1.55, w: 4.1, h: 2.75, fontSize: 12, lineSpacing: 16 });

    interactiveBar(s, "🙋 巴西说葡萄牙语！因为被葡萄牙殖民过", 4.85);
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 14 — 🇧🇷 巴西文化与礼节
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "🇧🇷 巴西文化与礼节  Brazil Culture & Etiquette");
    accentLine(s);

    card(s, 0.3, 1.0, 4.6, 3.6, C.white, C.bra);
    cardHeader(s, 0.3, 1.0, 4.6, 0.45, C.bra, "🌟 文化亮点 Culture");
    body(s, [
      "🎭 嘉年华 Carnival",
      "   世界最大派对！",
      "⚽ 足球王国",
      "   贝利 Pelé / 内马尔 Neymar",
      "🌳 亚马逊雨林产生20%氧气",
      "💃 桑巴舞 samba",
    ].join("\n"), { x: 0.5, y: 1.55, w: 4.2, h: 2.95, fontSize: 13, lineSpacing: 20 });

    card(s, 5.1, 1.0, 4.5, 3.6, C.ltOrg, C.accent);
    cardHeader(s, 5.1, 1.0, 4.5, 0.45, C.accent, "⚠️ 礼节 Etiquette");
    body(s, [
      "😘 见面亲两次脸颊！",
      "🤗 巴西人非常热情外向",
      "📏 说话时站得很近",
      "👍 竖大拇指 = OK",
      "   （不像有些地方不礼貌）",
      "⏰ 巴西时间 = 迟到很正常",
      "⚽ 足球是信仰",
      "   不要说足球不好！",
    ].join("\n"), { x: 5.3, y: 1.55, w: 4.1, h: 2.95, fontSize: 12, lineSpacing: 16 });

    interactiveBar(s, "⚽ 巴西赢了5次世界杯冠军！足球就是他们的信仰！", 4.85);
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 15 — 🇧🇷 巴西美食
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "🇧🇷 巴西美食  Brazil Food");
    accentLine(s);

    const foods = [
      { emoji: "🥩", cn: "烤肉 churrasco\n（服务员拿大串肉\n走来走去！）", x: 0.3, y: 1.0 },
      { emoji: "🫘", cn: "黑豆饭\nfeijoada", x: 2.5, y: 1.0 },
      { emoji: "🧀", cn: "奶酪面包\npão de queijo", x: 4.7, y: 1.0 },
      { emoji: "🫐", cn: "巴西莓碗\naçaí bowl", x: 0.3, y: 2.3 },
      { emoji: "☕", cn: "咖啡\n（世界第一咖啡产国）", x: 2.5, y: 2.3 },
    ];

    foods.forEach((f) => {
      card(s, f.x, f.y, 2.0, 1.1, C.white, C.midGrn);
      s.addText(f.emoji, {
        x: f.x, y: f.y + 0.05, w: 2.0, h: 0.45,
        fontSize: 26, align: "center",
      });
      s.addText(f.cn, {
        x: f.x + 0.1, y: f.y + 0.45, w: 1.8, h: 0.6,
        fontFace: "Calibri", fontSize: 10, color: C.dark, bold: true, align: "center", breakLine: true,
      });
    });

    card(s, 4.7, 2.3, 5.0, 1.1, C.ltOrg, C.accent);
    s.addText("「巴西烤肉可以一直吃到你说停！\n服务员会一直拿肉来，直到你翻红牌」", {
      x: 4.9, y: 2.35, w: 4.6, h: 1.0,
      fontFace: "Calibri", fontSize: 13, color: C.dark, bold: true, valign: "middle", breakLine: true,
    });

    card(s, 0.3, 3.6, 9.3, 0.95, C.dark);
    s.addText("「巴西是世界第一咖啡产国 — 全世界的咖啡有1/3来自巴西！」", {
      x: 0.5, y: 3.65, w: 8.9, h: 0.85,
      fontFace: "Georgia", fontSize: 16, color: C.accent,
      bold: true, align: "center", valign: "middle",
    });

    interactiveBar(s, "🙋 你想试试巴西烤肉吗？可以一直吃到饱！", 4.85);
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 16-17 — ✅ 巴西 Check Understanding (2 slides)
  // ══════════════════════════════════════════════════════════════
  (() => {
    const allQs = [
      { q: "巴西说什么语言？", a: "葡萄牙语（不是西班牙语！）" },
      { q: "巴西首都是哪里？", a: "巴西利亚（不是里约！）" },
      { q: "亚马逊雨林产多少氧气？", a: "20%！地球之肺" },
      { q: "巴西最有名的节日是什么？", a: "狂欢节 Carnival" },
      { q: "巴西说什么语言？", a: "葡萄牙语(不是西班牙语!)" },
      { q: "巴西首都是哪里？", a: "巴西利亚(不是里约!)" },
      { q: "亚马逊雨林产生多少氧气？", a: "20%" },
      { q: "巴西人见面亲几次脸颊？", a: "两次" },
      { q: "在巴西竖大拇指是什么意思？", a: "OK/好的" },
      { q: "在巴西千万不能说什么不好？", a: "足球！" },
    ];

    [0, 1].forEach((page) => {
      const s = pptx.addSlide();
      greenBg(s);
      topBar(s);

      hdr(s, "✅ 巴西 Check Understanding (" + (page + 1) + "/2)");
      accentLine(s);

      var pageQs = allQs.slice(page * 5, page * 5 + 5);
      pageQs.forEach((item, i) => {
        var num = page * 5 + i + 1;
        var yy = 1.0 + i * 0.85;
        var rowBg = i % 2 === 0 ? C.white : C.paleGrn;

        s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
          x: 0.3, y: yy, w: 9.2, h: 0.72, rectRadius: 0.1,
          fill: { color: rowBg },
          line: { color: C.bra, width: 1 },
        });
        s.addShape(pptx.shapes.OVAL, {
          x: 0.45, y: yy + 0.13, w: 0.45, h: 0.45,
          fill: { color: C.bra },
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
          fontSize: 11, fontFace: "Calibri", color: C.pri,
          bold: true, align: "left", valign: "middle",
        });
      });

      interactiveBar(s, "🏆 巴西小达人！Brazil expert!", 5.1);
      slideNum(s);
    });
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 17 — 🌽 美洲改变世界的食物
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "🌽 美洲改变世界的食物  Foods That Changed the World");
    accentLine(s);

    const foods = [
      { emoji: "🍅", name: "番茄 Tomato", arrow: "→ 披萨 / 番茄炒蛋", x: 0.4, y: 1.1 },
      { emoji: "🥔", name: "土豆 Potato",  arrow: "→ 薯条 French fries",  x: 5.2, y: 1.1 },
      { emoji: "🌽", name: "玉米 Corn",    arrow: "→ 全球主食",       x: 0.4, y: 2.4 },
      { emoji: "🍫", name: "巧克力 Chocolate", arrow: "→ 全世界都爱", x: 5.2, y: 2.4 },
      { emoji: "🌶️", name: "辣椒 Chili",   arrow: "→ 四川火锅！",    x: 0.4, y: 3.7 },
    ];

    foods.forEach((f) => {
      card(s, f.x, f.y, 4.4, 1.05, C.white, C.midGrn);
      s.addText(f.emoji, {
        x: f.x + 0.15, y: f.y + 0.1, w: 0.8, h: 0.85,
        fontSize: 34, align: "center", valign: "middle",
      });
      s.addText(f.name, {
        x: f.x + 1.0, y: f.y + 0.1, w: 1.8, h: 0.45,
        fontFace: "Calibri", fontSize: 14, color: C.dark, bold: true,
      });
      s.addText(f.arrow, {
        x: f.x + 1.0, y: f.y + 0.5, w: 3.0, h: 0.4,
        fontFace: "Calibri", fontSize: 13, color: C.pri,
      });
    });

    card(s, 5.2, 3.7, 4.4, 1.05, C.dark);
    s.addText("「如果没有美洲，\n世界食物会很不同！」", {
      x: 5.2, y: 3.7, w: 4.4, h: 1.05,
      fontFace: "Georgia", fontSize: 16, color: C.accent,
      bold: true, align: "center", valign: "middle", breakLine: true, lineSpacing: 24,
    });
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 18 — 🎭 Mini Role Play
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "🎭 迷你角色扮演  Mini Role Play");
    accentLine(s);

    const scenes = [
      { flag: "🇺🇸", country: "美国", greeting: "握手 + Hi! Nice to meet you!", fill: C.usa },
      { flag: "🇲🇽", country: "墨西哥", greeting: "拥抱 + Hola! ¿Cómo estás?", fill: C.mex },
      { flag: "🇧🇷", country: "巴西", greeting: "亲脸颊 + Oi! Tudo bem?", fill: C.bra },
    ];

    scenes.forEach((sc, i) => {
      const cy = 1.1 + i * 1.15;
      card(s, 0.4, cy, 9.2, 0.95, C.white, sc.fill);

      s.addShape(pptx.shapes.OVAL, {
        x: 0.6, y: cy + 0.1, w: 0.75, h: 0.75,
        fill: { color: sc.fill },
      });
      s.addText(sc.flag, {
        x: 0.6, y: cy + 0.1, w: 0.75, h: 0.75,
        fontSize: 26, align: "center", valign: "middle",
      });

      s.addText(sc.country, {
        x: 1.5, y: cy + 0.05, w: 2.0, h: 0.4,
        fontFace: "Georgia", fontSize: 18, color: C.dark, bold: true,
      });
      body(s, sc.greeting, {
        x: 1.5, y: cy + 0.45, w: 7.8, h: 0.4, fontSize: 14, lineSpacing: 16,
      });
    });

    card(s, 1.5, 4.6, 7.0, 0.65, C.dark);
    s.addText("🎭 两人一组，选一个国家，练习打招呼！", {
      x: 1.5, y: 4.6, w: 7.0, h: 0.65,
      fontFace: "Calibri", fontSize: 16, color: C.accent,
      bold: true, align: "center", valign: "middle",
    });
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 19 — 🏆 竞赛 + 🧩 Project提醒
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "🏆 竞赛 + 🧩 Project 提醒");
    accentLine(s);

    card(s, 0.4, 1.0, 4.4, 2.8, C.white, C.pri);
    cardHeader(s, 0.4, 1.0, 4.4, 0.45, C.pri, "🏆 上午快问快答");
    body(s, [
      "1. 美国首都？华盛顿DC",
      "2. 巧克力来自？墨西哥",
      "3. 巴西说什么语言？葡萄牙语",
      "4. 自由女神是谁送的？法国",
      "5. 亚马逊产多少氧气？20%",
      "6. 墨西哥有几种颜色的玉米？59",
    ].join("\n"), { x: 0.6, y: 1.55, w: 4.0, h: 2.1, fontSize: 12, lineSpacing: 18 });

    card(s, 5.2, 1.0, 4.4, 2.8, C.ltOrg, C.accent);
    cardHeader(s, 5.2, 1.0, 4.4, 0.45, C.accent, "🧩 Project 提醒");
    body(s, [
      "✏️ 下午完成护照美洲页",
      "   — 这是最后一页！",
      "",
      "🎨 装饰你的护照封面",
      "",
      "🎤 练习你的展示",
      "",
      "🎉 明天文化展！",
    ].join("\n"), { x: 5.4, y: 1.55, w: 4.0, h: 2.1, fontSize: 12, lineSpacing: 17 });

    card(s, 1.0, 4.1, 8.0, 0.65, C.dark);
    s.addText("「集齐4个签证章，成为环球小旅行家！」", {
      x: 1.0, y: 4.1, w: 8.0, h: 0.65,
      fontFace: "Georgia", fontSize: 18, color: C.accent,
      bold: true, align: "center", valign: "middle",
    });
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 20 — 下午开始 (Section Divider)
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    darkBg(s);
    topBar(s);
    bottomBar(s);

    s.addText("🌙", {
      x: 3.8, y: 0.5, w: 2.0, h: 1.0,
      fontSize: 52, align: "center",
    });

    s.addText("下午开始", {
      x: 1.0, y: 1.5, w: 8.0, h: 0.8,
      fontFace: "Georgia", fontSize: 34, color: C.white, bold: true, align: "center",
    });
    s.addText("Afternoon Session", {
      x: 1.0, y: 2.2, w: 8.0, h: 0.6,
      fontFace: "Georgia", fontSize: 22, color: C.accent, italic: true, align: "center",
    });

    s.addShape(pptx.shapes.RECTANGLE, {
      x: 3.0, y: 3.0, w: 4.0, h: 0.04, fill: { color: C.accent },
    });

    s.addText("对比三国 + 生词句型 + 完成护照 + 准备展览", {
      x: 1.0, y: 3.3, w: 8.0, h: 0.5,
      fontFace: "Calibri", fontSize: 18, color: C.white, align: "center",
    });

    s.addText("🇺🇸  🇲🇽  🇧🇷", {
      x: 2.0, y: 4.2, w: 6.0, h: 0.5,
      fontSize: 28, align: "center",
    });
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 21 — 快速复习
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "🔄 快速复习  Quick Review");
    accentLine(s);

    s.addText("上午学了什么？What did we learn this morning?", {
      x: 0.4, y: 0.95, w: 9.0, h: 0.4,
      fontFace: "Calibri", fontSize: 16, color: C.pri, italic: true,
    });

    const reviews = [
      "🇺🇸 美国首都是 ___？(华盛顿DC)",
      "🇺🇸 餐厅要给 ___% 小费？(15-20%)",
      "🇲🇽 巧克力是 ___ 人发明的？(阿兹特克)",
      "🇲🇽 墨西哥有 ___ 种颜色的玉米？(59)",
      "🇧🇷 巴西说 ___ 语？(葡萄牙语)",
      "🇧🇷 亚马逊雨林产生 ___% 的氧气？(20%)",
    ];

    reviews.forEach((r, i) => {
      const ry = 1.5 + i * 0.6;
      card(s, 0.6, ry, 8.6, 0.48, i % 2 === 0 ? C.white : C.paleGrn, C.midGrn);
      s.addText(r, {
        x: 0.8, y: ry, w: 8.2, h: 0.48,
        fontFace: "Calibri", fontSize: 14, color: C.dark, valign: "middle",
      });
    });

    interactiveBar(s, "🗣️ 大声说出答案！Shout out the answers!", 5.0);
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 22 — 美洲三国对比表
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "📊 美洲三国对比表  Three Countries Comparison");
    accentLine(s);

    // Table headers
    const cols = [
      { label: "", w: 1.5, x: 0.3 },
      { label: "🇺🇸 美国", w: 2.5, x: 1.8 },
      { label: "🇲🇽 墨西哥", w: 2.5, x: 4.3 },
      { label: "🇧🇷 巴西", w: 2.5, x: 6.8 },
    ];
    const headerY = 0.95;
    const headerH = 0.42;
    const headerColors = [C.gray, C.usa, C.mex, C.bra];

    cols.forEach((col, i) => {
      if (i > 0) {
        s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
          x: col.x, y: headerY, w: col.w, h: headerH, rectRadius: 0.08,
          fill: { color: headerColors[i] },
        });
        s.addText(col.label, {
          x: col.x, y: headerY, w: col.w, h: headerH,
          fontFace: "Georgia", fontSize: 14, color: C.white,
          bold: true, align: "center", valign: "middle",
        });
      }
    });

    // Table rows
    const rows = [
      { label: "打招呼", vals: ["握手 + 微笑", "拥抱 + 亲脸颊", "亲两次脸颊"] },
      { label: "时间观念", vals: ["准时", "随意", "迟到正常"] },
      { label: "代表食物", vals: ["汉堡 / BBQ", "taco / 巧克力", "烤肉 / 咖啡"] },
      { label: "注意事项", vals: ["给小费!\n直呼名字", "午餐最重要", "别说足球不好"] },
      { label: "有趣文化", vals: ["移民大熔炉", "亡灵节", "嘉年华"] },
    ];

    rows.forEach((row, ri) => {
      const ry = 1.45 + ri * 0.75;
      const rowBg = ri % 2 === 0 ? C.white : C.paleGrn;

      // Row label
      s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
        x: 0.3, y: ry, w: 1.5, h: 0.65, rectRadius: 0.06,
        fill: { color: C.pri },
      });
      s.addText(row.label, {
        x: 0.3, y: ry, w: 1.5, h: 0.65,
        fontFace: "Calibri", fontSize: 12, color: C.white,
        bold: true, align: "center", valign: "middle",
      });

      // Row values
      row.vals.forEach((val, ci) => {
        const cx = 1.8 + ci * 2.5;
        s.addShape(pptx.shapes.RECTANGLE, {
          x: cx, y: ry, w: 2.5, h: 0.65,
          fill: { color: rowBg },
          line: { color: C.midGrn, width: 0.5 },
        });
        s.addText(val, {
          x: cx + 0.05, y: ry, w: 2.4, h: 0.65,
          fontFace: "Calibri", fontSize: 11, color: C.dark,
          align: "center", valign: "middle", breakLine: true,
        });
      });
    });

    interactiveBar(s, "🙋 哪个国家的礼节和中国最像？Which is most like China?", 5.1);
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 23 — 共同点与不同
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "🤝 共同点与不同  Similarities & Differences");
    accentLine(s);

    card(s, 0.3, 1.0, 4.5, 3.4, C.white, C.pri);
    cardHeader(s, 0.3, 1.0, 4.5, 0.45, C.pri, "✅ 共同点 Similarities");
    body(s, [
      "😊 都很热情友好",
      "   All are warm and friendly",
      "",
      "🥩 都爱BBQ烤肉",
      "   All love BBQ / grilled meat",
      "",
      "🌍 都有丰富的移民和混合文化",
      "   Rich immigrant / mixed cultures",
    ].join("\n"), { x: 0.5, y: 1.55, w: 4.1, h: 2.75, fontSize: 12, lineSpacing: 17 });

    card(s, 5.0, 1.0, 4.6, 3.4, C.ltOrg, C.accent);
    cardHeader(s, 5.0, 1.0, 4.6, 0.45, C.accent, "❌ 不同点 Differences");
    body(s, [
      "🗣️ 语言不同",
      "   英语 / 西班牙语 / 葡萄牙语",
      "",
      "🤝 打招呼方式不同",
      "   握手 / 拥抱 / 亲脸颊",
      "",
      "⏰ 时间观念不同",
      "   准时 / 随意 / 迟到正常",
      "",
      "🍽️ 食物风格不同",
    ].join("\n"), { x: 5.2, y: 1.55, w: 4.2, h: 2.75, fontSize: 12, lineSpacing: 16 });

    interactiveBar(s, "🙋 你觉得哪个国家最有趣？Which country interests you most?", 4.85);
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 24 — 旅行小贴士
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "📌 旅行小贴士  Travel Tips");
    accentLine(s);

    const tips = [
      { flag: "🇺🇸", country: "美国", tips: "💰 给小费 15-20%\n📏 保持个人空间\n👋 直呼名字OK", fill: C.usa },
      { flag: "🇲🇽", country: "墨西哥", tips: "🐢 慢节奏，别急\n🤗 热情拥抱\n🍽️ 午餐最重要", fill: C.mex },
      { flag: "🇧🇷", country: "巴西", tips: "🤗 超级热情\n⚽ 尊重足球\n⏰ 迟到正常", fill: C.bra },
    ];

    tips.forEach((t, i) => {
      const tx = 0.3 + i * 3.2;
      card(s, tx, 1.0, 3.0, 3.4, C.white, t.fill);

      s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
        x: tx, y: 1.0, w: 3.0, h: 0.7, rectRadius: 0.1,
        fill: { color: t.fill },
      });
      s.addShape(pptx.shapes.RECTANGLE, {
        x: tx, y: 1.5, w: 3.0, h: 0.2, fill: { color: t.fill },
      });
      s.addText(t.flag + " " + t.country, {
        x: tx, y: 1.0, w: 3.0, h: 0.7,
        fontFace: "Georgia", fontSize: 18, color: C.white,
        bold: true, align: "center", valign: "middle",
      });

      body(s, t.tips, {
        x: tx + 0.2, y: 1.85, w: 2.6, h: 2.4, fontSize: 13, lineSpacing: 22, breakLine: true,
      });
    });

    interactiveBar(s, "📌 记住这些小贴士，做一个有礼貌的旅行者！", 4.85);
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 25 — 📝 生词卡
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "📝 生词卡  Vocabulary Cards");
    accentLine(s);

    const vocab = [
      { cn: "汉堡包", py: "hànbǎo bāo", en: "hamburger" },
      { cn: "玉米",   py: "yùmǐ",       en: "corn" },
      { cn: "巧克力", py: "qiǎokèlì",   en: "chocolate" },
      { cn: "辣椒",   py: "làjiāo",      en: "chili pepper" },
      { cn: "烤肉",   py: "kǎoròu",      en: "BBQ / grilled meat" },
      { cn: "雨林",   py: "yǔlín",       en: "rainforest" },
      { cn: "小费",   py: "xiǎofèi",     en: "tip (gratuity)" },
      { cn: "移民",   py: "yímín",       en: "immigrant" },
    ];

    vocab.forEach((v, i) => {
      const col = i % 4;
      const row = Math.floor(i / 4);
      const vx = 0.3 + col * 2.4;
      const vy = 1.0 + row * 2.0;

      card(s, vx, vy, 2.2, 1.7, C.white, C.pri);
      s.addText(v.cn, {
        x: vx, y: vy + 0.1, w: 2.2, h: 0.55,
        fontFace: "Georgia", fontSize: 22, color: C.dark,
        bold: true, align: "center", valign: "middle",
      });
      s.addText(v.py, {
        x: vx, y: vy + 0.65, w: 2.2, h: 0.4,
        fontFace: "Calibri", fontSize: 12, color: C.pri,
        italic: true, align: "center", valign: "middle",
      });
      s.addText(v.en, {
        x: vx, y: vy + 1.1, w: 2.2, h: 0.4,
        fontFace: "Calibri", fontSize: 13, color: C.accent,
        bold: true, align: "center", valign: "middle",
      });
    });
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 26 — 💬 句型练习
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "💬 句型练习  Sentence Patterns");
    accentLine(s);

    card(s, 0.4, 1.1, 9.2, 1.6, C.white, C.pri);
    cardHeader(s, 0.4, 1.1, 9.2, 0.45, C.pri, "句型一  Pattern 1");
    s.addText("「___原来是从美洲来的！」", {
      x: 0.6, y: 1.65, w: 8.8, h: 0.45,
      fontFace: "Georgia", fontSize: 20, color: C.accent, bold: true, align: "center",
    });
    body(s, "例: 番茄原来是从美洲来的！/ 巧克力原来是从美洲来的！\nEx: Tomatoes actually came from the Americas!", {
      x: 0.6, y: 2.1, w: 8.8, h: 0.5, fontSize: 12,
    });

    card(s, 0.4, 2.95, 9.2, 1.6, C.white, C.accent);
    cardHeader(s, 0.4, 2.95, 9.2, 0.45, C.accent, "句型二  Pattern 2");
    s.addText("「如果没有___，就没有___」", {
      x: 0.6, y: 3.5, w: 8.8, h: 0.45,
      fontFace: "Georgia", fontSize: 20, color: C.pri, bold: true, align: "center",
    });
    body(s, "例: 如果没有番茄，就没有番茄炒蛋！/ 如果没有辣椒，就没有四川火锅！\nEx: Without tomatoes, no scrambled eggs! Without chili, no Sichuan hotpot!", {
      x: 0.6, y: 3.95, w: 8.8, h: 0.5, fontSize: 12,
    });

    interactiveBar(s, "🗣️ 用句型造一个句子！Make a sentence using the patterns!", 4.85);
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 27 — 🎭 Role Play — 华人到美洲
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "🎭 角色扮演 — 华人到美洲  Role Play");
    accentLine(s);

    card(s, 0.4, 1.0, 4.4, 3.5, C.white, C.pri);
    cardHeader(s, 0.4, 1.0, 4.4, 0.5, C.pri, "📖 场景 Scenes");
    body(s, [
      "1️⃣ 来到美国，握手打招呼",
      "   Hi! Nice to meet you!",
      "",
      "2️⃣ 在墨西哥，拥抱新朋友",
      "   Hola! 尝试taco!",
      "",
      "3️⃣ 在巴西，亲脸颊",
      "   Oi! 看足球比赛!",
      "",
      "4️⃣ 介绍中国食物给新朋友",
    ].join("\n"), { x: 0.6, y: 1.6, w: 4.0, h: 2.8, fontSize: 12, lineSpacing: 16 });

    card(s, 5.2, 1.0, 4.4, 3.5, C.ltOrg, C.accent);
    cardHeader(s, 5.2, 1.0, 4.4, 0.5, C.accent, "🎯 练习要点 Key Points");
    body(s, [
      "🤝 注意打招呼方式不同！",
      "   美国=握手 墨西哥=拥抱 巴西=亲脸颊",
      "",
      "🗣️ 用中文介绍自己",
      "   「我从中国来」",
      "   「这是中国的饺子」",
      "",
      "💡 观察文化差异",
      "   小费/时间/热情程度",
    ].join("\n"), { x: 5.4, y: 1.6, w: 4.0, h: 2.8, fontSize: 12, lineSpacing: 16 });

    interactiveBar(s, "🎭 选一个场景，和同学一起演！Pick a scene and act it out!", 4.85);
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 28 — 🎨 Project Time — 护照美洲页
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "🎨 Project Time!  护照美洲页");
    accentLine(s);

    s.addText("Passport Page 4: Americas 美洲 — 最后一页！", {
      x: 0.4, y: 0.95, w: 9.0, h: 0.35,
      fontFace: "Calibri", fontSize: 15, color: C.pri, bold: true, italic: true,
    });

    const sections = [
      { title: "🌎 Where I Went\n我去了哪里", detail: "美国/墨西哥/巴西 + 国旗", fill: C.pri },
      { title: "👀 What I Saw\n我看到了什么", detail: "自由女神/奇琴伊察/亚马逊", fill: C.midGrn },
      { title: "🍽️ What I Ate\n我吃了什么", detail: "汉堡/taco/巴西烤肉/巧克力", fill: C.accent },
      { title: "💡 Cultural Discovery\n文化发现", detail: "打招呼方式/小费/时间观念", fill: C.dark },
      { title: "✏️ My Sentence\n我的句子", detail: "「___原来是从美洲来的！」", fill: C.stamp },
    ];

    sections.forEach((sec, i) => {
      const sy = 1.4 + i * 0.75;
      card(s, 0.4, sy, 9.2, 0.65, C.white, sec.fill);

      s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
        x: 0.4, y: sy, w: 3.2, h: 0.65, rectRadius: 0.1,
        fill: { color: sec.fill },
      });
      s.addText(sec.title, {
        x: 0.4, y: sy, w: 3.2, h: 0.65,
        fontFace: "Calibri", fontSize: 11, color: C.white,
        bold: true, align: "center", valign: "middle", breakLine: true,
      });

      body(s, sec.detail, {
        x: 3.8, y: sy, w: 5.6, h: 0.65, fontSize: 12, valign: "middle", breakLine: true,
      });
    });
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 29 — Project 分层
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "📊 Project 分层  Differentiation Levels");
    accentLine(s);

    // Level 1
    card(s, 0.4, 1.0, 9.2, 1.15, C.white, C.midGrn);
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 0.4, y: 1.0, w: 1.8, h: 1.15, rectRadius: 0.1,
      fill: { color: C.midGrn },
    });
    s.addText("🟢 Level 1\n画 + 词", {
      x: 0.4, y: 1.0, w: 1.8, h: 1.15,
      fontFace: "Calibri", fontSize: 14, color: C.white,
      bold: true, align: "center", valign: "middle", breakLine: true,
    });
    body(s, "画出美洲的标志 + 写关键词\n例: 🗽 + 自由女神 / 🌮 + 玉米饼 / ⚽ + 足球", {
      x: 2.4, y: 1.05, w: 7.0, h: 1.0, fontSize: 13, lineSpacing: 18,
    });

    // Level 2
    card(s, 0.4, 2.4, 9.2, 1.15, C.white, "1565C0");
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 0.4, y: 2.4, w: 1.8, h: 1.15, rectRadius: 0.1,
      fill: { color: "1565C0" },
    });
    s.addText("🔵 Level 2\n简单句子", {
      x: 0.4, y: 2.4, w: 1.8, h: 1.15,
      fontFace: "Calibri", fontSize: 14, color: C.white,
      bold: true, align: "center", valign: "middle", breakLine: true,
    });
    body(s, "用简单句子描述\n例: 美国人握手打招呼。巴西烤肉很好吃。", {
      x: 2.4, y: 2.45, w: 7.0, h: 1.0, fontSize: 13, lineSpacing: 18,
    });

    // Level 3
    card(s, 0.4, 3.8, 9.2, 1.15, C.white, "7B1FA2");
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 0.4, y: 3.8, w: 1.8, h: 1.15, rectRadius: 0.1,
      fill: { color: "7B1FA2" },
    });
    s.addText("🟣 Level 3\n完整段落", {
      x: 0.4, y: 3.8, w: 1.8, h: 1.15,
      fontFace: "Calibri", fontSize: 14, color: C.white,
      bold: true, align: "center", valign: "middle", breakLine: true,
    });
    body(s, "写完整的句子和段落 + 使用句型\n例: 巧克力原来是从美洲来的！如果没有墨西哥，就没有巧克力！", {
      x: 2.4, y: 3.85, w: 7.0, h: 1.0, fontSize: 13, lineSpacing: 18,
    });
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 30 — 准备明天展览
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "🎪 准备明天展览  Prepare for Exhibition");
    accentLine(s);

    const steps = [
      { n: "1", title: "检查护照 4 页", detail: "亚洲 + 非洲 + 欧洲 + 美洲\nReview all 4 continent pages", fill: C.pri },
      { n: "2", title: "装饰护照封面", detail: "画你最喜欢的标志\nDecorate your passport cover", fill: C.accent },
      { n: "3", title: "练习展示", detail: "练习介绍你最喜欢的国家\nPractice presenting your favorite country", fill: C.midGrn },
    ];

    steps.forEach((st, i) => {
      const sy = 1.0 + i * 1.35;
      card(s, 0.4, sy, 9.2, 1.15, C.white, st.fill);

      s.addShape(pptx.shapes.OVAL, {
        x: 0.6, y: sy + 0.15, w: 0.85, h: 0.85,
        fill: { color: st.fill },
      });
      s.addText(st.n, {
        x: 0.6, y: sy + 0.15, w: 0.85, h: 0.85,
        fontFace: "Georgia", fontSize: 28, color: C.white,
        bold: true, align: "center", valign: "middle",
      });

      s.addText(st.title, {
        x: 1.7, y: sy + 0.05, w: 7.5, h: 0.45,
        fontFace: "Georgia", fontSize: 18, color: C.dark, bold: true,
      });
      body(s, st.detail, {
        x: 1.7, y: sy + 0.5, w: 7.5, h: 0.55, fontSize: 13, breakLine: true,
      });
    });

    interactiveBar(s, "📋 Check: 你的护照有4页了吗？Do you have all 4 pages?", 4.85);
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 31 — 分享时间
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    greenBg(s);
    topBar(s);

    hdr(s, "🤝 分享时间  Partner Share");
    accentLine(s);

    card(s, 0.4, 1.0, 9.2, 3.5, C.white, C.pri);

    s.addText([
      { text: "和同伴分享你的护照！", options: { fontSize: 20, fontFace: "Georgia", color: C.dark, bold: true, breakLine: true } },
      { text: "Share your passport with a partner!", options: { fontSize: 16, fontFace: "Georgia", color: C.pri, italic: true, breakLine: true } },
    ], { x: 0.6, y: 1.1, w: 8.8, h: 0.8, align: "center" });

    const shareSteps = [
      "1️⃣ 展示你最喜欢的一页 Show your favorite page",
      "2️⃣ 用今天的句型说一句话 Say a sentence using today's patterns",
      "3️⃣ 告诉同伴一个有趣的事实 Tell a fun fact",
      "4️⃣ 听同伴的分享，问一个问题 Listen and ask a question",
    ];

    shareSteps.forEach((step, i) => {
      const sy = 2.0 + i * 0.55;
      s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
        x: 1.0, y: sy, w: 8.0, h: 0.45,
        fill: { color: i % 2 === 0 ? C.paleGrn : C.ltOrg }, rectRadius: 0.08,
      });
      s.addText(step, {
        x: 1.2, y: sy, w: 7.6, h: 0.45,
        fontFace: "Calibri", fontSize: 13, color: C.dark, valign: "middle",
      });
    });

    interactiveBar(s, "👂 Good listeners ask good questions! 好的听众会问好问题！", 4.85);
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 32 — 🪪 美洲签证章 (集齐4个!)
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    darkBg(s);
    topBar(s);
    bottomBar(s);

    s.addText("🪪 签证章时间！  Visa Stamp Time!", {
      x: 0.5, y: 0.3, w: 9.0, h: 0.6,
      fontFace: "Georgia", fontSize: 30, color: C.white, bold: true, align: "center",
    });

    s.addShape(pptx.shapes.OVAL, {
      x: 3.0, y: 1.2, w: 3.8, h: 3.0,
      line: { color: C.accent, width: 4 },
      fill: { color: C.dark },
    });
    s.addShape(pptx.shapes.OVAL, {
      x: 3.15, y: 1.35, w: 3.5, h: 2.7,
      line: { color: C.accent, width: 2, dashType: "dash" },
    });
    s.addText("AMERICAS", {
      x: 3.2, y: 1.6, w: 3.4, h: 0.6,
      fontFace: "Georgia", fontSize: 22, color: C.accent,
      bold: true, align: "center",
    });
    s.addText("美洲 ✓", {
      x: 3.2, y: 2.15, w: 3.4, h: 0.5,
      fontFace: "Georgia", fontSize: 20, color: C.white,
      bold: true, align: "center",
    });
    s.addText("6/11", {
      x: 3.2, y: 2.6, w: 3.4, h: 0.45,
      fontFace: "Calibri", fontSize: 16, color: C.accent, align: "center",
    });
    s.addText("APPROVED", {
      x: 3.2, y: 3.05, w: 3.4, h: 0.45,
      fontFace: "Georgia", fontSize: 14, color: C.accent,
      italic: true, align: "center",
    });

    s.addText("🎉 集齐4个签证章！明天展览！", {
      x: 1.0, y: 4.4, w: 8.0, h: 0.5,
      fontFace: "Georgia", fontSize: 22, color: C.accent,
      bold: true, align: "center",
    });

    s.addText("🌏 亚洲 ✓  🌍 非洲 ✓  🌍 欧洲 ✓  🌎 美洲 ✓", {
      x: 1.0, y: 4.9, w: 8.0, h: 0.4,
      fontFace: "Calibri", fontSize: 14, color: C.white, align: "center",
    });
    slideNum(s);
  })();

  // ══════════════════════════════════════════════════════════════
  // SLIDE 33 — ✈️ 明天文化展
  // ══════════════════════════════════════════════════════════════
  (() => {
    const s = pptx.addSlide();
    darkBg(s);
    topBar(s);
    bottomBar(s);

    s.addText("✈️ 明天预告  Coming Tomorrow!", {
      x: 0.5, y: 0.3, w: 9.0, h: 0.6,
      fontFace: "Georgia", fontSize: 30, color: C.white, bold: true, align: "center",
    });

    s.addShape(pptx.shapes.RECTANGLE, {
      x: 3.0, y: 1.0, w: 4.0, h: 0.04, fill: { color: C.accent },
    });

    s.addText("🎉", {
      x: 3.5, y: 1.2, w: 3.0, h: 0.8,
      fontSize: 48, align: "center",
    });

    card(s, 1.5, 2.1, 7.0, 2.0, C.pri);
    s.addText("国际文化展", {
      x: 1.5, y: 2.2, w: 7.0, h: 0.7,
      fontFace: "Georgia", fontSize: 32, color: C.accent, bold: true, align: "center",
    });
    s.addText("International Culture Exhibition", {
      x: 1.5, y: 2.8, w: 7.0, h: 0.5,
      fontFace: "Georgia", fontSize: 18, color: C.white, italic: true, align: "center",
    });
    s.addText([
      { text: "📖 带上你的护照！", options: { fontSize: 14, fontFace: "Calibri", color: C.white, breakLine: true } },
      { text: "🎨 准备好你的展示！", options: { fontSize: 14, fontFace: "Calibri", color: C.white, breakLine: true } },
      { text: "🌏 这是我们旅程的终点站！", options: { fontSize: 14, fontFace: "Calibri", color: C.white, breakLine: true } },
    ], { x: 1.5, y: 3.3, w: 7.0, h: 0.7, align: "center" });

    s.addText("🌏🌍🌎  环游世界的旅程即将圆满完成！  🌎🌍🌏", {
      x: 0.5, y: 4.4, w: 9.0, h: 0.45,
      fontFace: "Calibri", fontSize: 18, color: C.accent,
      align: "center", bold: true,
    });

    s.addText("谷雨中文 GR EDU  |  Global Explorer Camp 2025", {
      x: 1.5, y: 5.0, w: 7.0, h: 0.3,
      fontFace: "Calibri", fontSize: 11, color: C.white, align: "center",
    });
    slideNum(s);
  })();

  // ─── Save ───
  const outPath = __dirname + "/day4_americas.pptx";
  pptx.writeFile({ fileName: outPath })
    .then(() => console.log("Created: " + outPath))
    .catch((err) => console.error(err));
}

buildPresentation().catch(console.error);
