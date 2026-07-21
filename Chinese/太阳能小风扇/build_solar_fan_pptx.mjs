import path from "node:path";
import fs from "node:fs/promises";
import { fileURLToPath } from "node:url";
import {
  Presentation,
  PresentationFile,
} from "/Users/Huan/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUT = path.join(__dirname, "太阳能小风扇_project.pptx");
const ASSET = (name) => path.join(__dirname, "assets", name);
const W = 1280;
const H = 720;

const C = {
  ink: "#19324a",
  muted: "#617184",
  paper: "#fffaf0",
  sky: "#dff3ff",
  sun: "#ffb627",
  orange: "#ff8c00",
  green: "#43a047",
  teal: "#00897b",
  blue: "#1976d2",
  red: "#c62828",
  line: "#d8ccb6",
  dark: "#1f5d54",
  panel: "#18314f",
  white: "#ffffff",
};

const presentation = Presentation.create({ slideSize: { width: W, height: H } });

function rect(slide, x, y, width, height, fill = C.white, line = "transparent", radius = "roundRect") {
  return slide.shapes.add({
    geometry: radius,
    position: { left: x, top: y, width, height },
    fill,
    line: { fill: line, width: line === "transparent" ? 0 : 2, style: "solid" },
  });
}

function ellipse(slide, x, y, width, height, fill, line = "transparent", lineWidth = 0) {
  return slide.shapes.add({
    geometry: "ellipse",
    position: { left: x, top: y, width, height },
    fill,
    line: { fill: line, width: lineWidth, style: "solid" },
  });
}

function text(slide, value, x, y, width, height, opts = {}) {
  const shape = rect(slide, x, y, width, height, opts.fill ?? "transparent", "transparent", "rect");
  shape.text = value;
  shape.text.fontSize = opts.size ?? 28;
  shape.text.color = opts.color ?? C.ink;
  shape.text.bold = Boolean(opts.bold);
  shape.text.typeface = opts.face ?? "Noto Sans SC";
  shape.text.alignment = opts.align ?? "left";
  shape.text.verticalAlignment = opts.valign ?? "top";
  shape.text.insets = opts.insets ?? { left: 0, right: 0, top: 0, bottom: 0 };
  return shape;
}

function line(slide, x, y, width, height, color = C.ink, lineWidth = 4) {
  return slide.shapes.add({
    geometry: "line",
    position: { left: x, top: y, width, height },
    fill: "transparent",
    line: { fill: color, width: lineWidth, style: "solid" },
  });
}

async function photo(slide, fileName, x, y, width, height, opts = {}) {
  const bytes = await fs.readFile(ASSET(fileName));
  const img = slide.images.add({
    dataUrl: `data:image/png;base64,${bytes.toString("base64")}`,
    fit: opts.fit ?? "cover",
    alt: opts.alt ?? fileName,
  });
  img.position = { left: x, top: y, width, height };
  return img;
}

function base(slide, kicker, title, subtitle) {
  rect(slide, 0, 0, W, H, C.paper, "transparent", "rect");
  ellipse(slide, -120, -150, 420, 420, "#ffe3a1");
  ellipse(slide, 1040, -90, 320, 320, "#ccefe8");
  text(slide, kicker, 64, 38, 440, 34, { size: 18, color: C.teal, bold: true, valign: "middle" });
  text(slide, title, 64, 78, 720, 72, { size: 44, color: C.dark, bold: true });
  if (subtitle) text(slide, subtitle, 66, 143, 750, 46, { size: 19, color: C.muted });
  line(slide, 64, 196, 340, 0, "#9dd6cc", 4);
}

function drawSun(slide, cx, cy, r = 44) {
  ellipse(slide, cx - r, cy - r, r * 2, r * 2, C.sun);
  for (let i = 0; i < 8; i += 1) {
    const a = (Math.PI * 2 * i) / 8;
    const x1 = cx + Math.cos(a) * (r + 10);
    const y1 = cy + Math.sin(a) * (r + 10);
    const x2 = cx + Math.cos(a) * (r + 38);
    const y2 = cy + Math.sin(a) * (r + 38);
    line(slide, x1, y1, x2 - x1, y2 - y1, "#ffd46f", 7);
  }
}

function drawSolarPanel(slide, x, y, w = 230, h = 130) {
  rect(slide, x, y, w, h, C.panel, "#17345b", "roundRect");
  for (let i = 1; i < 5; i += 1) line(slide, x + (w * i) / 5, y, 0, h, "#78b9ff", 2);
  for (let i = 1; i < 3; i += 1) line(slide, x, y + (h * i) / 3, w, 0, "#78b9ff", 2);
}

function drawFan(slide, cx, cy, scale = 1) {
  const r = 78 * scale;
  ellipse(slide, cx - r, cy - r, r * 2, r * 2, "#f7fcff", C.dark, 6);
  ellipse(slide, cx - 18 * scale, cy - 74 * scale, 36 * scale, 76 * scale, C.green);
  ellipse(slide, cx - 74 * scale, cy - 18 * scale, 76 * scale, 36 * scale, C.blue);
  ellipse(slide, cx - 18 * scale, cy - 2 * scale, 36 * scale, 76 * scale, C.sun);
  ellipse(slide, cx - 2 * scale, cy - 18 * scale, 76 * scale, 36 * scale, C.orange);
  ellipse(slide, cx - 18 * scale, cy - 18 * scale, 36 * scale, 36 * scale, C.orange);
  ellipse(slide, cx - 7 * scale, cy - 7 * scale, 14 * scale, 14 * scale, C.white);
}

function bullet(slide, value, x, y, width, color = C.ink) {
  text(slide, `• ${value}`, x, y, width, 36, { size: 24, color });
}

function card(slide, x, y, w, h, title, body, color = C.dark) {
  rect(slide, x, y, w, h, C.white, C.line, "roundRect");
  text(slide, title, x + 20, y + 18, w - 40, 34, { size: 24, bold: true, color });
  text(slide, body, x + 20, y + 62, w - 40, h - 76, { size: 18, color: C.muted });
}

async function slide01() {
  const s = presentation.slides.add();
  rect(s, 0, 0, W, H, "#fff7df", "transparent", "rect");
  await photo(s, "solar_fan_car_hero.png", 668, 0, 612, 720, { alt: "realistic DIY solar fan car project in sunlight" });
  rect(s, 606, 0, 120, 720, "linear(90deg, #fff7df, rgba(255,247,223,0))", "transparent", "rect");
  text(s, "新能源 STEM 项目", 82, 64, 360, 34, { size: 20, color: C.teal, bold: true });
  text(s, "太阳能小风扇", 82, 150, 650, 96, { size: 64, color: C.dark, bold: true });
  text(s, "Solar Mini Fan Project", 86, 248, 560, 42, { size: 28, color: C.muted });
  text(s, "光能 → 电能 → 动能", 90, 332, 470, 48, { size: 32, color: C.orange, bold: true });
  text(s, "像真实科技小制作一样搭建、测试、展示。", 90, 420, 480, 62, { size: 25, color: C.ink, bold: true });
  return s;
}

async function slide02() {
  const s = presentation.slides.add();
  base(s, "ENGAGE", "今天的问题", "Can sunlight make wind?");
  text(s, "没有电池，风扇能不能转起来？", 110, 260, 760, 74, { size: 46, color: C.dark, bold: true });
  text(s, "Students predict first, then test with light, angle, and circuit connections.", 114, 346, 760, 62, { size: 24, color: C.muted });
  await photo(s, "solar_fan_testing_sunlight.png", 842, 214, 344, 324, { alt: "solar fan car being tested in sunlight" });
  rect(s, 842, 214, 344, 324, "transparent", C.line, "roundRect");
  return s;
}

async function slide03() {
  const s = presentation.slides.add();
  base(s, "SCIENCE", "能量旅行路线", "Energy conversion in one simple system");
  await photo(s, "solar_fan_circuit_closeup.png", 710, 218, 456, 310, { alt: "close-up solar panel wires motor and propeller" });
  card(s, 86, 248, 260, 130, "1 光能", "太阳光照到太阳能板", C.orange);
  card(s, 390, 248, 260, 130, "2 电能", "太阳能板产生电流", C.blue);
  card(s, 86, 414, 260, 130, "3 动能", "马达带动扇叶转动", C.green);
  card(s, 390, 414, 260, 130, "4 风", "空气被推动", C.teal);
  text(s, "句型: 我发现光越强，风扇转得越 ____，因为 ____。", 116, 590, 940, 50, { size: 28, color: C.ink, bold: true });
  return s;
}

async function slide04() {
  const s = presentation.slides.add();
  base(s, "MATERIALS", "材料准备", "Each team checks parts before building");
  await photo(s, "solar_fan_parts_flatlay.png", 694, 226, 444, 360, { alt: "flat lay of solar fan project parts" });
  const items = [
    ["太阳能板", "Solar panel", C.panel],
    ["小马达", "Mini motor", C.blue],
    ["塑料扇叶", "Fan blade", C.green],
    ["红黑导线", "Wires", C.red],
    ["硬纸板底座", "Cardboard base", C.orange],
    ["支架材料", "Straws / craft sticks", C.teal],
  ];
  items.forEach((it, i) => {
    const x = 88 + (i % 2) * 286;
    const y = 238 + Math.floor(i / 2) * 116;
    rect(s, x, y, 244, 82, C.white, C.line, "roundRect");
    ellipse(s, x + 22, y + 26, 52, 52, it[2]);
    text(s, it[0], x + 88, y + 18, 140, 28, { size: 21, bold: true, color: C.ink });
    text(s, it[1], x + 88, y + 50, 140, 24, { size: 15, color: C.muted });
  });
  return s;
}

async function slide05() {
  const s = presentation.slides.add();
  base(s, "CIRCUIT", "先让马达转起来", "Closed circuit before decoration");
  await photo(s, "solar_fan_circuit_closeup.png", 82, 236, 560, 354, { alt: "real solar circuit close-up with motor and propeller" });
  card(s, 710, 248, 410, 112, "1 夹稳导线", "红线、黑线分别连接太阳能板和马达两端。", C.red);
  card(s, 710, 390, 410, 112, "2 照光测试", "把太阳能板面向强光，观察马达是否启动。", C.orange);
  card(s, 710, 532, 410, 96, "3 排查问题", "不转时先检查连接，再调整角度。", C.teal);
  return s;
}

async function slide06() {
  const s = presentation.slides.add();
  base(s, "BUILD", "制作步骤", "Build, test, then reinforce");
  await photo(s, "solar_fan_building_classroom.png", 742, 224, 420, 336, { alt: "students assembling solar fan car kit" });
  const steps = [
    ["1", "认识零件", "观察 panel, motor, blade, wires"],
    ["2", "连接电路", "红线和黑线分别连接两端"],
    ["3", "安装扇叶", "轻轻压上马达轴，不要太紧"],
    ["4", "固定底座", "保持扇叶离开桌面"],
    ["5", "调整角度", "寻找最快启动的位置"],
  ];
  steps.forEach((st, i) => {
    const y = 238 + i * 78;
    ellipse(s, 110, y, 46, 46, i % 2 ? C.teal : C.orange);
    text(s, st[0], 110, y + 3, 46, 34, { size: 24, color: C.white, bold: true, align: "center", valign: "middle" });
    text(s, st[1], 180, y - 2, 240, 34, { size: 27, color: C.dark, bold: true });
    text(s, st[2], 430, y + 2, 620, 32, { size: 21, color: C.muted });
  });
  return s;
}

async function slide07() {
  const s = presentation.slides.add();
  base(s, "TEST", "测试挑战", "Change one variable at a time");
  await photo(s, "solar_fan_testing_sunlight.png", 830, 232, 330, 260, { alt: "completed solar fan car testing in sunlight" });
  const cols = [90, 330, 570];
  const heads = ["测试条件", "预测", "结果"];
  heads.forEach((h, i) => {
    rect(s, cols[i], 238, 210, 54, "#e9f7f3", C.line, "rect");
    text(s, h, cols[i] + 14, 250, 180, 28, { size: 22, bold: true, color: C.dark });
  });
  ["正对强光", "斜着对光", "在阴影里", "改变支架高度"].forEach((row, r) => {
    const y = 292 + r * 70;
    cols.forEach((x) => rect(s, x, y, 210, 70, C.white, C.line, "rect"));
    text(s, row, 104, y + 20, 210, 28, { size: 20, color: C.ink });
  });
  text(s, "挑战: 怎样让风扇在 5 秒内启动？", 94, 606, 760, 42, { size: 30, bold: true, color: C.orange });
  return s;
}

async function slide08() {
  const s = presentation.slides.add();
  base(s, "VOCABULARY", "中文关键词", "Use these words in the final explanation");
  const words = [
    ["太阳能板", "solar panel"],
    ["导线", "wire"],
    ["马达", "motor"],
    ["扇叶", "fan blade"],
    ["电路", "circuit"],
    ["能量转换", "energy conversion"],
    ["光能", "light energy"],
    ["动能", "motion energy"],
  ];
  words.forEach((w, i) => {
    const x = 90 + (i % 4) * 292;
    const y = 248 + Math.floor(i / 4) * 148;
    rect(s, x, y, 250, 104, C.white, C.line, "roundRect");
    text(s, w[0], x + 20, y + 22, 210, 32, { size: 25, bold: true, color: C.dark, align: "center" });
    text(s, w[1], x + 20, y + 60, 210, 26, { size: 17, color: C.muted, align: "center" });
  });
  return s;
}

async function slide09() {
  const s = presentation.slides.add();
  base(s, "SHOWCASE", "小小工程师展示", "Explain what worked, what changed, and why");
  await photo(s, "solar_fan_car_hero.png", 790, 216, 360, 330, { alt: "finished solar STEM model ready for showcase" });
  card(s, 88, 238, 300, 174, "我的作品能...", "在光照下启动\n支架稳定\n扇叶可以安全转动", C.green);
  card(s, 430, 238, 300, 174, "我测试了...", "光照强弱\n太阳能板角度\n支架高度或扇叶方向", C.blue);
  card(s, 88, 444, 642, 108, "我发现...", "光越强，转得越快；角度和结构会影响启动速度。", C.orange);
  text(s, "收尾句: 太阳能是一种清洁能源，它可以帮助我们减少对普通电池的依赖。", 90, 600, 1080, 58, { size: 27, color: C.dark, bold: true, align: "center" });
  return s;
}

for (const fn of [
  slide01,
  slide02,
  slide03,
  slide04,
  slide05,
  slide06,
  slide07,
  slide08,
  slide09,
]) {
  await fn();
}

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(OUT);
console.log(OUT);
