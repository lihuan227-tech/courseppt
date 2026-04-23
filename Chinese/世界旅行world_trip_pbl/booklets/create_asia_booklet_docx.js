const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, BorderStyle, WidthType,
        ShadingType, PageBreak, LevelFormat, PageNumber, VerticalAlign } = require("docx");

// ─── Constants ───
const W = 9360; // US Letter content width (1" margins)
const RED = "C62828";
const ACCENT = "FF8F00";
const BG_WARM = "FFF8E1";
const CHINA_RED = "DE2910";
const JAPAN_RED = "BC002D";
const INDIA_GREEN = "138808";
const GRAY = "888888";
const LGRAY = "CCCCCC";
const border = { style: BorderStyle.SINGLE, size: 1, color: LGRAY };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const cm = { top: 80, bottom: 80, left: 120, right: 120 };
const dash = { style: BorderStyle.DASHED, size: 2, color: "BBBBBB" };
const dashB = { top: dash, bottom: dash, left: dash, right: dash };

// ─── Helpers ───
function h(text, color = RED, sz = 28) {
  return new Paragraph({ spacing: { before: 80, after: 80 },
    children: [new TextRun({ text, bold: true, size: sz*2, font: "Arial", color })] });
}
function sub(text, color = ACCENT) {
  return new Paragraph({ spacing: { before: 60, after: 60 },
    shading: { fill: color, type: ShadingType.CLEAR },
    children: [new TextRun({ text: "  " + text, bold: true, size: 24, font: "Arial", color: "FFFFFF" })] });
}
function p(cn, en) {
  const c = [new TextRun({ text: cn, bold: true, size: 24, font: "Arial" })];
  if (en) c.push(new TextRun({ text: "  " + en, size: 20, font: "Arial", color: "666666", italics: true }));
  return new Paragraph({ spacing: { before: 80, after: 40 }, children: c });
}
function t(str, sz = 22, color = "333333", bold = false) {
  return new Paragraph({ spacing: { before: 30, after: 30 },
    children: [new TextRun({ text: str, size: sz, font: "Arial", color, bold })] });
}
function sp(pts = 80) { return new Paragraph({ spacing: { before: pts, after: 0 }, children: [] }); }
function pb() { return new Paragraph({ children: [new PageBreak()] }); }
function bl() {
  return new Paragraph({ spacing: { before: 30, after: 30 },
    border: { bottom: { style: BorderStyle.DOTTED, size: 2, color: "999999", space: 1 } },
    children: [new TextRun({ text: " ", size: 28 })] });
}
function wa(n = 3) {
  const rows = [];
  for (let i = 0; i < n; i++) rows.push(new Paragraph({
    border: { bottom: { style: BorderStyle.DOTTED, size: 1, color: "DDDDDD", space: 1 } },
    children: [new TextRun({ text: " ", size: 28 })] }));
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({ children: [new TableCell({ borders, width: { size: W, type: WidthType.DXA },
      margins: { top: 60, bottom: 60, left: 100, right: 100 }, children: rows })] })] });
}
function db(label, n = 5) {
  const lines = [new Paragraph({ children: [new TextRun({ text: label, size: 18, font: "Arial", color: "BBBBBB", italics: true })] })];
  for (let i = 0; i < n; i++) lines.push(new Paragraph({ children: [new TextRun({ text: " ", size: 24 })] }));
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({ children: [new TableCell({ borders: dashB, width: { size: W, type: WidthType.DXA },
      margins: cm, children: lines })] })] });
}
function tc(chars, sz = 48) {
  const cw = Math.floor(W / chars.length);
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: Array(chars.length).fill(cw),
    rows: [new TableRow({ children: chars.map(ch => new TableCell({ borders: dashB,
      width: { size: cw, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
      margins: { top: 100, bottom: 100, left: 40, right: 40 },
      children: [new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: ch, size: sz*2, font: "KaiTi", color: "DDDDDD" })] })] })) })] });
}
function ff(emoji, str) {
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({ children: [new TableCell({
      borders: { top: noBorder, bottom: noBorder, right: noBorder,
        left: { style: BorderStyle.SINGLE, size: 8, color: ACCENT } },
      width: { size: W, type: WidthType.DXA },
      shading: { fill: BG_WARM, type: ShadingType.CLEAR }, margins: cm,
      children: [new Paragraph({ children: [
        new TextRun({ text: emoji + " ", size: 24 }),
        new TextRun({ text: str, size: 22, font: "Arial" }) ] })] })] })] });
}

// ─── Build 8-page booklet ───
const doc = new Document({
  styles: { default: { document: { run: { font: "Arial", size: 22 } } } },
  sections: [{
    properties: {
      page: { size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
    },
    headers: { default: new Header({ children: [new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: "Global Explorer Camp \u00b7 \u8c37\u96e8\u4e2d\u6587 GR EDU", size: 16, color: GRAY })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "\u63a2\u7d22\u4e9a\u6d32 Explore Asia \u00b7 Page ", size: 16, color: GRAY }),
        new TextRun({ children: [PageNumber.CURRENT], size: 16, color: GRAY })] })] }) },
    children: [

      // ════════════════════════════════════
      // COVER PAGE: 封面
      // ════════════════════════════════════
      sp(300),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "\ud83c\udf0f", size: 72 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 },
        children: [new TextRun({ text: "\u63a2\u7d22\u4e9a\u6d32", bold: true, size: 64, font: "Arial", color: RED })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
        children: [new TextRun({ text: "Explore: Asia", bold: true, size: 48, font: "Arial", color: ACCENT })] }),
      sp(60),
      new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "\u2708\ufe0f Global Explorer Camp \u5168\u7403\u63a2\u9669\u8425", bold: true, size: 28, font: "Arial", color: RED })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
        children: [new TextRun({ text: "\ud83c\udde8\ud83c\uddf3 \ud83c\uddef\ud83c\uddf5 \ud83c\uddee\ud83c\uddf3", size: 48 })] }),
      sp(60),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 },
        children: [new TextRun({ text: "\ud83c\udfaf \u6559\u5b66\u76ee\u6807 Learning Objectives", bold: true, size: 24, font: "Arial", color: RED })] }),
      new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "\ud83d\udc41 \u6211\u4f1a\u8ba4\uff1a\u4e9a\u6d32  \u4e2d\u56fd  \u65e5\u672c  \u5370\u5ea6  \u9996\u90fd", size: 24, font: "Arial" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
        children: [new TextRun({ text: "\u270d\ufe0f \u6211\u4f1a\u5199\uff1a\u4e2d\u56fd  \u65e5\u672c  \u4e9a\u6d32  \u9996\u90fd", size: 24, font: "Arial" })] }),
      sp(100),
      new Paragraph({ spacing: { after: 80 },
        children: [
          new TextRun({ text: "\u5c0f\u63a2\u9669\u5bb6 Explorer:  ", bold: true, size: 26, font: "Arial" }),
          new TextRun({ text: "________________________________", size: 26 }) ] }),
      new Paragraph({ spacing: { after: 100 },
        children: [
          new TextRun({ text: "\u65e5\u671f Date:  ", bold: true, size: 26, font: "Arial" }),
          new TextRun({ text: "________________________________", size: 26 }) ] }),
      sp(100),
      new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "\ud83c\udfef \ud83d\uddfb \ud83d\udc3c \ud83d\udc09 \ud83c\udf8b \ud83d\udd4c \ud83c\udf8e \ud83e\udde7", size: 40 })] }),

      pb(),

      // ════════════════════════════════════
      // PAGE 2: 亚洲 (K-1 低级别)
      // ════════════════════════════════════
      h("\ud83c\udf0f \u8ba4\u8bc6\u4e9a\u6d32 About Asia (K-1)", RED),
      sub("\u6211\u4f1a\u8ba4 I Can Read \u2014 \u8fde\u4e00\u8fde Match Words to Pictures"),
      // Matching activity: words left, pictures right (scrambled)
      new Table({
        width: { size: W, type: WidthType.DXA }, columnWidths: [2800, 3760, 2800],
        rows: [
          // Header
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 2800, type: WidthType.DXA },
              shading: { fill: ACCENT, type: ShadingType.CLEAR }, margins: cm,
              children: [new Paragraph({ alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: "\u8bcd\u8bed Words", bold: true, size: 22, color: "FFFFFF" })] })] }),
            new TableCell({ borders, width: { size: 3760, type: WidthType.DXA },
              shading: { fill: "EEEEEE", type: ShadingType.CLEAR }, margins: cm,
              children: [new Paragraph({ alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: "\u2190 \u8fde\u7ebf Draw lines \u2192", size: 20, color: GRAY, italics: true })] })] }),
            new TableCell({ borders, width: { size: 2800, type: WidthType.DXA },
              shading: { fill: ACCENT, type: ShadingType.CLEAR }, margins: cm,
              children: [new Paragraph({ alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: "\u56fe\u7247 Pictures", bold: true, size: 22, color: "FFFFFF" })] })] }),
          ] }),
          // Row 1: 亚洲 → 🏛️ capital building (scrambled to match 首都)
          ...[
            ["\u4e9a\u6d32", "\ud83c\uddef\ud83c\uddf5 \u65e5\u51fa\u4e4b\u56fd Land of Rising Sun"],
            ["\u4e2d\u56fd", "\ud83c\udf0f \u6700\u5927\u7684\u6d32 Biggest continent"],
            ["\u65e5\u672c", "\ud83d\ude4f Namaste \u5408\u5341"],
            ["\u5370\u5ea6", "\ud83d\udc3c \u718a\u732b Panda"],
            ["\u9996\u90fd", "\ud83c\udfe2 \u5317\u4eac Beijing / \u4e1c\u4eac Tokyo"],
          ].map(([word, pic], idx) => new TableRow({ children: [
            new TableCell({ borders, width: { size: 2800, type: WidthType.DXA },
              shading: { fill: idx % 2 === 0 ? BG_WARM : "FFFFFF", type: ShadingType.CLEAR }, margins: cm,
              verticalAlign: VerticalAlign.CENTER,
              children: [new Paragraph({ alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: word, bold: true, size: 28, font: "Arial" })] })] }),
            new TableCell({ borders, width: { size: 3760, type: WidthType.DXA },
              shading: { fill: "FFFFFF", type: ShadingType.CLEAR }, margins: cm,
              children: [new Paragraph({ children: [new TextRun({ text: " ", size: 22 })] })] }),
            new TableCell({ borders, width: { size: 2800, type: WidthType.DXA },
              shading: { fill: idx % 2 === 0 ? "E3F2FD" : "FFFFFF", type: ShadingType.CLEAR }, margins: cm,
              verticalAlign: VerticalAlign.CENTER,
              children: [new Paragraph({ alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: pic, size: 20, font: "Arial" })] })] }),
          ] })),
        ]
      }),
      sp(40),
      sub("\u63cf\u4e00\u63cf Trace: \u4e9a\u6d32 \u9996\u90fd"),
      tc(["\u4e9a", "\u6d32", "\u9996", "\u90fd"]),
      sp(40),
      sub("\u2b55 \u5708\u4e00\u5708 Circle the Correct Answer"),
      t("\u4e9a\u6d32\u6709\u591a\u5c11\u4e2a\u56fd\u5bb6\uff1f How many countries?", 22, "333333", true),
      t("   A. 20     B. 48     C. 100", 22),
      sp(20),
      t("\u4e9a\u6d32\u662f\u4e16\u754c\u4e0a\u6700______\u7684\u6d32\u3002 Asia is the ___ continent.", 22, "333333", true),
      t("   A. \u5c0f small     B. \u51b7 cold     C. \u5927 big", 22),
      sp(20),
      t("\u4e16\u754c\u4e0a\u6700\u9ad8\u7684\u5c71\u662f\uff1f The tallest mountain is:", 22, "333333", true),
      t("   A. \u5bcc\u58eb\u5c71 Mt. Fuji     B. \u73e0\u7a46\u6717\u739b\u5cf0 Mt. Everest     C. \u9ec4\u5c71 Mt. Huang", 22),
      sp(40),
      sub("\ud83c\udf0d \u6d82\u8272 Colour the Continents"),
      p("\u5728\u4e0b\u9762\u7684\u4e16\u754c\u5730\u56fe\u4e0a\uff0c\u7ed9\u4e03\u5927\u6d32\u6d82\u4e0a\u4e0d\u540c\u7684\u989c\u8272", "Colour each continent a different colour on the map below"),
      t("\ud83d\udd34 \u4e9a\u6d32 Asia = \u7ea2\u8272 Red    \ud83d\udfe1 \u975e\u6d32 Africa = \u9ec4\u8272 Yellow    \ud83d\udd35 \u6b27\u6d32 Europe = \u84dd\u8272 Blue", 18),
      t("\ud83d\udfe2 \u5317\u7f8e\u6d32 N.America = \u7eff\u8272 Green    \ud83d\udfe0 \u5357\u7f8e\u6d32 S.America = \u6a59\u8272 Orange    \ud83d\udfe3 \u5927\u6d0b\u6d32 Oceania = \u7d2b\u8272 Purple", 18),
      db("\ud83c\udf0d \u4e16\u754c\u5730\u56fe World Map \u2014 \u7528\u4e0d\u540c\u989c\u8272\u6d82\u51fa\u4e03\u5927\u6d32\uff01", 5),

      pb(),

      // ════════════════════════════════════
      // PAGE 2: 亚洲 (Level 2+ 高级别)
      // ════════════════════════════════════
      h("\ud83c\udf0f \u8ba4\u8bc6\u4e9a\u6d32 About Asia (Level 2+)", RED),
      sub("\u6211\u4f1a\u5199 I Can Write"),
      t("\u6bcf\u4e2a\u8bcd\u5199\u4e24\u904d Write each word twice:", 20, GRAY),
      tc(["\u4e9a", "\u6d32"], 40), tc(["\u4e9a", "\u6d32"], 40),
      tc(["\u9996", "\u90fd"], 40), tc(["\u9996", "\u90fd"], 40),
      sp(40),
      sub("\u586b\u7a7a Fill in the Blanks"),
      p("\u4e9a\u6d32\u6709 ______ \u4e2a\u56fd\u5bb6\u3002", "Asia has ___ countries."),
      p("\u4e9a\u6d32\u662f\u4e16\u754c\u4e0a\u6700 ______ \u7684\u6d32\u3002", "Asia is the ___ continent."),
      p("\u4e16\u754c\u4e0a\u6700\u9ad8\u7684\u5c71\u662f ____________\u3002", "The tallest mountain is ___."),
      sp(40),
      sub("\u7528\u53e5\u5b50\u56de\u7b54 Answer in Full Sentences"),
      p("\u5173\u4e8e\u4e9a\u6d32\uff0c\u4f60\u5b66\u5230\u4e86\u4ec0\u4e48\uff1f", "What did you learn about Asia?"),
      wa(4),

      pb(),

      // ════════════════════════════════════
      // PAGE 3: 中国 (K-1)
      // ════════════════════════════════════
      h("\ud83c\udde8\ud83c\uddf3 \u4e2d\u56fd China (K-1)", CHINA_RED),
      sub("\u6d82\u4e00\u6d82\u56fd\u65d7 Colour the Flag", CHINA_RED),
      t("\ud83d\udd34 \u7ea2\u8272 Red + \u2b50 \u9ec4\u8272 Yellow\uff08\u4e94\u9897\u661f\uff09"),
      db("\u5728\u8fd9\u91cc\u753b\u4e2d\u56fd\u56fd\u65d7 Draw the Chinese flag here", 3),
      sp(40),
      sub("\u63cf\u4e00\u63cf Trace: \u4e2d\u56fd", CHINA_RED),
      tc(["\u4e2d", "\u56fd", "\u4e2d", "\u56fd"]),
      sp(40),
      sub("\u2b55 \u5708\u4e00\u5708 Circle the Correct Answer", CHINA_RED),
      t("\u4e2d\u56fd\u7684\u9996\u90fd\u662f\uff1f Capital of China?", 22, "333333", true),
      t("   A. \u4e0a\u6d77 Shanghai     B. \u5317\u4eac Beijing     C. \u4e1c\u4eac Tokyo", 22),
      sp(20),
      t("\u4e2d\u56fd\u4eba\u7528\u4ec0\u4e48\u5403\u996d\uff1f Eat with?", 22, "333333", true),
      t("   A. \u53c9\u5b50 Fork     B. \u624b Hand     C. \u7b77\u5b50 Chopsticks", 22),
      sp(20),
      t("\u4e2d\u56fd\u7684\u56fd\u5b9d\u662f\uff1f National treasure?", 22, "333333", true),
      t("   A. \u8001\u864e Tiger     B. \u718a\u732b Panda     C. \u9f99 Dragon", 22),
      sp(40),
      sub("\u753b\u4e00\u753b Draw", CHINA_RED),
      p("\u753b\u4e00\u753b\u4f60\u6700\u559c\u6b22\u7684\u4e2d\u56fd\u98df\u7269\u6216\u5730\u65b9", "Draw your favorite Chinese food or place"),
      db("\u997a\u5b50? \u957f\u57ce? \u718a\u732b? Dumplings? Great Wall? Panda?", 4),

      pb(),

      // ════════════════════════════════════
      // PAGE 4: 中国 (Level 2+)
      // ════════════════════════════════════
      h("\ud83c\udde8\ud83c\uddf3 \u4e2d\u56fd China (Level 2+)", CHINA_RED),
      sub("\u6211\u4f1a\u5199 I Can Write", CHINA_RED),
      t("\u6bcf\u4e2a\u8bcd\u5199\u4e24\u904d Write each word twice:", 20, GRAY),
      tc(["\u4e2d", "\u56fd"], 40), tc(["\u4e2d", "\u56fd"], 40),
      sp(40),
      sub("\u586b\u7a7a Fill in the Blanks", CHINA_RED),
      p("\u4e2d\u56fd\u7684\u9996\u90fd\u662f ____________\u3002", "The capital of China is ___."),
      p("\u4e2d\u56fd\u4eba\u7528 ____________ \u5403\u996d\u3002", "Chinese people eat with ___."),
      p("\u4e2d\u56fd\u7684\u56fd\u5b9d\u662f ____________\u3002", "China's national treasure is ___."),
      sp(40),
      sub("\u7528\u53e5\u5b50\u56de\u7b54 Answer in Full Sentences", CHINA_RED),
      p("\u5728\u4e2d\u56fd\uff0c\u7b77\u5b50\u4e0d\u80fd\u600e\u4e48\u653e\uff1f\u4e3a\u4ec0\u4e48\uff1f", "How should you NOT place chopsticks? Why?"),
      wa(3),
      p("\u4f60\u6700\u559c\u6b22\u5403\u4ec0\u4e48\u4e2d\u56fd\u83dc\uff1f\u4e3a\u4ec0\u4e48\uff1f", "What Chinese food do you like most? Why?"),
      wa(3),

      pb(),

      // ════════════════════════════════════
      // PAGE 5: 日本 (K-1)
      // ════════════════════════════════════
      h("\ud83c\uddef\ud83c\uddf5 \u65e5\u672c Japan (K-1)", JAPAN_RED),
      sub("\u6d82\u4e00\u6d82\u56fd\u65d7 Colour the Flag", JAPAN_RED),
      t("\u2b1c \u767d\u8272 White + \ud83d\udd34 \u7ea2\u8272 Red circle"),
      db("\u5728\u8fd9\u91cc\u753b\u65e5\u672c\u56fd\u65d7 Draw the Japanese flag here", 3),
      sp(40),
      sub("\u63cf\u4e00\u63cf Trace: \u65e5\u672c", JAPAN_RED),
      tc(["\u65e5", "\u672c", "\u65e5", "\u672c"]),
      sp(40),
      sub("\u2b55 \u5708\u4e00\u5708 Circle the Correct Answer", JAPAN_RED),
      t("\u5728\u65e5\u672c\u89c1\u9762\u5e94\u8be5\u600e\u4e48\u505a\uff1f How to greet?", 22, "333333", true),
      t("   A. \u62e5\u62b1 Hug     B. \u63e1\u624b Shake hands     C. \u97a0\u8eac Bow", 22),
      sp(20),
      t("\u8fdb\u522b\u4eba\u5bb6\u8981\u505a\u4ec0\u4e48\uff1f Entering a house?", 22, "333333", true),
      t("   A. \u8131\u978b Take off shoes     B. \u6d17\u624b Wash hands     C. \u62cd\u624b Clap", 22),
      sp(20),
      t("\u5728\u65e5\u672c\u5403\u62c9\u9762\u53d1\u51fa\u58f0\u97f3\u662f\uff1f Slurping ramen is:", 22, "333333", true),
      t("   A. \u4e0d\u793c\u8c8c Rude     B. \u793c\u8c8c Polite     C. \u5947\u602a Weird", 22),
      sp(40),
      sub("\u753b\u4e00\u753b Draw", JAPAN_RED),
      p("\u753b\u4e00\u753b\u5bcc\u58eb\u5c71\u6216\u6a31\u82b1", "Draw Mt. Fuji or cherry blossoms"),
      db("\u5bcc\u58eb\u5c71? \u6a31\u82b1? \u5bff\u53f8? Mt. Fuji? Cherry blossoms? Sushi?", 4),

      pb(),

      // ════════════════════════════════════
      // PAGE 6: 日本 (Level 2+)
      // ════════════════════════════════════
      h("\ud83c\uddef\ud83c\uddf5 \u65e5\u672c Japan (Level 2+)", JAPAN_RED),
      sub("\u6211\u4f1a\u5199 I Can Write", JAPAN_RED),
      t("\u6bcf\u4e2a\u8bcd\u5199\u4e24\u904d Write each word twice:", 20, GRAY),
      tc(["\u65e5", "\u672c"], 40), tc(["\u65e5", "\u672c"], 40),
      sp(40),
      sub("\u586b\u7a7a Fill in the Blanks", JAPAN_RED),
      p("\u65e5\u672c\u7684\u9996\u90fd\u662f ____________\u3002", "The capital of Japan is ___."),
      p("\u65e5\u672c\u7531 ____________ \u4e2a\u5c9b\u5c7f\u7ec4\u6210\u3002", "Japan has ___ islands."),
      p("\u65e5\u672c\u6700\u9ad8\u7684\u5c71\u662f ____________\u3002", "Japan's tallest mountain is ___."),
      sp(40),
      sub("\u7528\u53e5\u5b50\u56de\u7b54 Answer in Full Sentences", JAPAN_RED),
      p("\u5728\u65e5\u672c\uff0c\u89c1\u9762\u65f6\u5e94\u8be5\u600e\u4e48\u505a\uff1f", "How do you greet people in Japan?"),
      wa(3),
      p("\u5728\u65e5\u672c\u65c5\u884c\uff0c\u4e0d\u80fd\u505a\u4ec0\u4e48\uff1f\u8bf4\u4e24\u4e2a\u4f8b\u5b50\u3002", "What should you NOT do in Japan? Two examples."),
      wa(3),

      pb(),

      // ════════════════════════════════════
      // PAGE 7: 印度 (K-1)
      // ════════════════════════════════════
      h("\ud83c\uddee\ud83c\uddf3 \u5370\u5ea6 India (K-1)", INDIA_GREEN),
      sub("\u6d82\u4e00\u6d82\u56fd\u65d7 Colour the Flag", INDIA_GREEN),
      t("\ud83d\udfe0 \u6a59\u8272 Saffron + \u2b1c \u767d\u8272 White + \ud83d\udfe2 \u7eff\u8272 Green + \ud83d\udd35 Chakra"),
      db("\u5728\u8fd9\u91cc\u753b\u5370\u5ea6\u56fd\u65d7 Draw the Indian flag here", 3),
      sp(40),
      sub("\u63cf\u4e00\u63cf Trace: \u5370\u5ea6", INDIA_GREEN),
      tc(["\u5370", "\u5ea6", "\u5370", "\u5ea6"]),
      sp(40),
      sub("\u2b55 \u5708\u4e00\u5708 Circle the Correct Answer", INDIA_GREEN),
      t("\u5370\u5ea6\u4eba\u89c1\u9762\u600e\u4e48\u6253\u62db\u547c\uff1f How do Indians greet?", 22, "333333", true),
      t("   A. \u63e1\u624b Shake hands     B. \u5408\u5341 Namaste     C. \u97a0\u8eac Bow", 22),
      sp(20),
      t("\u5370\u5ea6\u4eba\u7528\u54ea\u53ea\u624b\u5403\u996d\uff1f Which hand?", 22, "333333", true),
      t("   A. \u5de6\u624b Left     B. \u53f3\u624b Right     C. \u4e24\u53ea\u624b Both", 22),
      sp(20),
      t("\u5370\u5ea6\u53d1\u660e\u4e86\u4ec0\u4e48\u6570\u5b57\uff1f What number?", 22, "333333", true),
      t("   A. 1     B. 0     C. 10", 22),
      sp(40),
      sub("\u753b\u4e00\u753b Draw", INDIA_GREEN),
      p("\u753b\u4e00\u753b\u53cc\u624b\u5408\u5341(Namaste)\u6216\u6cf0\u59ec\u9675", "Draw the Namaste gesture or Taj Mahal"),
      db("\ud83d\ude4f Namaste! / Taj Mahal / \u5496\u55b1 Curry", 4),

      pb(),

      // ════════════════════════════════════
      // PAGE 8: 印度 (Level 2+)
      // ════════════════════════════════════
      h("\ud83c\uddee\ud83c\uddf3 \u5370\u5ea6 India (Level 2+)", INDIA_GREEN),
      sub("\u6211\u4f1a\u5199 I Can Write", INDIA_GREEN),
      t("\u6bcf\u4e2a\u8bcd\u5199\u4e24\u904d Write each word twice:", 20, GRAY),
      tc(["\u5370", "\u5ea6"], 40), tc(["\u5370", "\u5ea6"], 40),
      sp(40),
      sub("\u586b\u7a7a Fill in the Blanks", INDIA_GREEN),
      p("\u5370\u5ea6\u7684\u9996\u90fd\u662f ____________\u3002", "The capital of India is ___."),
      p("\u5370\u5ea6\u6709 ____________ \u79cd\u5b98\u65b9\u8bed\u8a00\u3002", "India has ___ official languages."),
      p("\u5370\u5ea6\u6559\u5f92\u4e0d\u5403 ____________\uff0c\u56e0\u4e3a\u725b\u662f ____________ \u7684\u3002", "Hindus don't eat ___ because cows are ___."),
      sp(40),
      sub("\u7528\u53e5\u5b50\u56de\u7b54 Answer in Full Sentences", INDIA_GREEN),
      p("\u6392\u706f\u8282(Diwali)\u662f\u4ec0\u4e48\u8282\u65e5\uff1f\u50cf\u4e2d\u56fd\u7684\u4ec0\u4e48\u8282\uff1f", "What is Diwali? Like which Chinese holiday?"),
      wa(3),
      p("\u5728\u5370\u5ea6\u65c5\u884c\uff0c\u8981\u6ce8\u610f\u54ea\u4e9b\u793c\u8282\uff1f", "What etiquette should you follow in India?"),
      wa(3),

    ]
  }]
});

const out = "/Users/Huan/projects/summercourse/Chinese/\u4e16\u754c\u65c5\u884cworld_trip_pbl/booklets/asia_booklet.docx";
Packer.toBuffer(doc).then(buf => { fs.writeFileSync(out, buf); console.log("Created: " + out); });
