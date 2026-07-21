// Build day1–day4 booklets for the 零废弃 Zero Waste unit.
// Structure mirrors 我的职业梦想/booklets: Cover → §1 Question Bank (8 Qs, circle answer)
// → §2 Apply (pick + draw + sentence frame + write lines) → §3 连一连 Match → §4 描一描写一写.
// Run: node build_booklets.js   (uses repo-root node_modules/docx)
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
} = require('docx');

// ===== shared colors =====
const DARK = '2C2C2C', GRAY = '888888', LGRAY = 'D8D8D8';
const C_PICK = '1565C0', C_MATCH = 'E08A1E', C_WRITE = '6A1B9A';

// ===== page geometry (US Letter, 0.75" margins) =====
const PAGE = { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } };
const CW = 12240 - 1080 - 1080;

// ===== helpers (ported from reference) =====
const border = (color = 'CCCCCC', size = 4) => ({ style: BorderStyle.SINGLE, size, color });
const noBorder = () => ({ style: BorderStyle.NONE, size: 0, color: 'FFFFFF' });
const allBorders = (b) => ({ top: b, bottom: b, left: b, right: b, insideHorizontal: b, insideVertical: b });

function shadedBar(text, colorHex, size = 24) {
  return new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: [CW], borders: allBorders(noBorder()),
    rows: [new TableRow({ children: [new TableCell({
      width: { size: CW, type: WidthType.DXA },
      shading: { fill: colorHex, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 200, right: 200 },
      children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: 'FFFFFF', size })] })],
    })] })],
  });
}
const sectionBreak = () => new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] });

function drawBox(label, height = 3200, colorHex = 'BBBBBB') {
  return new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: [CW],
    borders: allBorders({ style: BorderStyle.SINGLE, size: 6, color: colorHex }),
    rows: [new TableRow({ height: { value: height, rule: 'atLeast' }, children: [new TableCell({
      width: { size: CW, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, verticalAlign: 'center',
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: label, italics: true, color: LGRAY, size: 18 })] })],
    })] })],
  });
}

function writeLines(n = 4) {
  return new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: [CW],
    borders: { top: noBorder(), left: noBorder(), right: noBorder(),
      bottom: { style: BorderStyle.SINGLE, size: 8, color: '666666' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 8, color: '666666' }, insideVertical: noBorder() },
    rows: Array.from({ length: n }, () => new TableRow({ height: { value: 560, rule: 'atLeast' },
      children: [new TableCell({ width: { size: CW, type: WidthType.DXA },
        borders: { top: noBorder(), left: noBorder(), right: noBorder(), bottom: { style: BorderStyle.SINGLE, size: 8, color: '666666' } },
        margins: { top: 60, bottom: 60, left: 0, right: 0 }, children: [new Paragraph({ children: [new TextRun({ text: '', size: 22 })] })] })] })),
  });
}

// ===== Cover =====
function cover(day) {
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 1000, after: 160 }, children: [new TextRun({ text: day.coverEmoji, size: 120 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 160, after: 80 }, children: [new TextRun({ text: '零废弃与可持续发展', bold: true, size: 56, color: day.accent })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 500 }, children: [new TextRun({ text: 'Zero Waste & Sustainability', bold: true, size: 32, color: day.accent })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 160, after: 80 }, children: [new TextRun({ text: `Day ${day.day} · ${day.topicCn}`, bold: true, size: 44, color: DARK })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 160 }, children: [new TextRun({ text: day.topicEn, italics: true, size: 26, color: GRAY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 500 }, children: [new TextRun({ text: day.emojiRow, size: 30 })] }),
    ...['姓名 Name', '班级 Class', '日期 Date'].map((lab, i) => new Paragraph({
      spacing: { before: i === 0 ? 500 : 260, after: 260 },
      children: [new TextRun({ text: `${lab}: `, bold: true, size: 26 }), new TextRun({ text: '____________________________', size: 26, color: GRAY })],
    })),
    new Paragraph({ spacing: { before: 800, after: 0 }, children: [new TextRun({ text: '' })] }),
    new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [CW],
      borders: allBorders({ style: BorderStyle.SINGLE, size: 10, color: day.accent }),
      rows: [new TableRow({ height: { value: 2100, rule: 'atLeast' }, children: [new TableCell({
        width: { size: CW, type: WidthType.DXA }, shading: { fill: 'F7F8F0', type: ShadingType.CLEAR },
        margins: { top: 200, bottom: 200, left: 200, right: 200 }, verticalAlign: 'center',
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '🤔 今天的大问题  Today’s Big Question', bold: true, size: 24, color: day.accent })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 180 }, children: [new TextRun({ text: day.topicCn, bold: true, size: 32, color: DARK })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60 }, children: [new TextRun({ text: day.topicEn, italics: true, size: 18, color: GRAY })] }),
        ],
      })] })],
    }),
  ];
}

// ===== §1 Question Bank — 2-column grid (4 rows × 2), fits one page =====
const QB_COLORS = ['C5283C', '6A1B9A', '1565C0', '2E7D32', 'E08A1E', '00897B', '7B5E3F', 'D81B60'];
const QB_COLW = Math.floor((CW - 200) / 2);
function qbCell(num, q, color) {
  const opts = q.opts.map((opt, idx) => new Paragraph({
    spacing: { before: 30, after: 0 }, indent: { left: 340 },
    children: [
      new TextRun({ text: '☐ ', bold: true, size: 22, color: DARK }),
      new TextRun({ text: `${String.fromCharCode(65 + idx)}. `, bold: true, size: 22, color: DARK }),
      new TextRun({ text: opt.cn, size: 24, color: DARK }),
      new TextRun({ text: `  ${opt.en}`, italics: true, size: 13, color: GRAY }),
    ],
  }));
  return new TableCell({
    width: { size: QB_COLW, type: WidthType.DXA }, borders: allBorders({ style: BorderStyle.SINGLE, size: 6, color }),
    margins: { top: 120, bottom: 120, left: 170, right: 150 }, verticalAlign: 'center',
    children: [
      new Paragraph({ spacing: { before: 0, after: 24 }, children: [
        new TextRun({ text: `${num} `, bold: true, size: 24, color }),
        new TextRun({ text: `${q.em} `, size: 22 }),
        new TextRun({ text: q.cn, bold: true, size: 24, color: DARK }),
      ] }),
      new Paragraph({ spacing: { before: 0, after: 40 }, indent: { left: 340 }, children: [new TextRun({ text: q.en, italics: true, size: 13, color: GRAY })] }),
      ...opts,
    ],
  });
}
function gapCell() {
  return new TableCell({ width: { size: 200, type: WidthType.DXA }, borders: allBorders(noBorder()), children: [new Paragraph({ children: [new TextRun({ text: '' })] })] });
}
function questionBank(day) {
  const rows = [];
  for (let i = 0; i < day.questions.length; i += 2) {
    rows.push(new TableRow({ cantSplit: true, children: [
      qbCell(i + 1, day.questions[i], QB_COLORS[i % QB_COLORS.length]),
      gapCell(),
      qbCell(i + 2, day.questions[i + 1], QB_COLORS[(i + 1) % QB_COLORS.length]),
    ] }));
  }
  return [
    sectionBreak(),
    shadedBar('一、题库 · 8 个思考题 / Question Bank · 8 Questions  (圈出正确答案)', day.accent, 24),
    new Paragraph({ spacing: { before: 110, after: 110 }, children: [new TextRun({ text: '👉 读一读, 圈出对的答案。', size: 24, italics: true, color: GRAY }), new TextRun({ text: ' / Read & circle the right answer.', size: 16, italics: true, color: GRAY })] }),
    new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [QB_COLW, 200, QB_COLW], borders: allBorders(noBorder()), rows }),
    new Paragraph({ spacing: { before: 160, after: 0 }, alignment: AlignmentType.CENTER, children: [
      new TextRun({ text: '🏆 ', size: 20 }),
      new TextRun({ text: '答对 8 题 = “Zero Waste 小达人” 徽章! ', bold: true, size: 24, color: day.accent }),
      new TextRun({ text: 'All 8 right = Zero Waste Hero badge!', italics: true, size: 14, color: GRAY }),
    ] }),
  ];
}

// ===== §2 Apply (pick + draw + sentence frame + write lines) =====
function pickCell(opt) {
  return new TableCell({
    width: { size: Math.floor(CW / 2), type: WidthType.DXA }, borders: allBorders(noBorder()),
    margins: { top: 40, bottom: 40, left: 200, right: 100 },
    children: [new Paragraph({ children: [
      new TextRun({ text: '☐  ', size: 26, bold: true }),
      new TextRun({ text: `${opt.em}  `, size: 24 }),
      new TextRun({ text: `${opt.cn}  `, size: 24, bold: true, color: DARK }),
      new TextRun({ text: opt.en, size: 16, color: GRAY }),
    ] })],
  });
}
function applyActivity(day) {
  const a = day.apply;
  const rows = [];
  for (let i = 0; i < a.options.length; i += 2) {
    rows.push(new TableRow({ children: [pickCell(a.options[i]), a.options[i + 1] ? pickCell(a.options[i + 1]) : new TableCell({ borders: allBorders(noBorder()), children: [new Paragraph({ children: [new TextRun({ text: '' })] })] })] }));
  }
  return [
    sectionBreak(),
    shadedBar(`二、${a.titleCn} / ${a.titleEn}`, C_PICK, 24),
    new Paragraph({ spacing: { before: 100, after: 60 }, children: [new TextRun({ text: a.pickCn, size: 24, bold: true, color: DARK })] }),
    new Paragraph({ spacing: { before: 20, after: 100 }, children: [new TextRun({ text: a.pickEn, size: 18, italics: true, color: GRAY })] }),
    new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [Math.floor(CW / 2), Math.floor(CW / 2)], borders: allBorders(noBorder()), rows }),
    new Paragraph({ spacing: { before: 200, after: 80 }, children: [
      new TextRun({ text: '🎨 画一画 ', size: 24, bold: true, color: C_PICK }),
      new TextRun({ text: a.drawCn, size: 22, bold: true, color: DARK }),
      new TextRun({ text: '  Draw', size: 16, italics: true, color: GRAY }),
    ] }),
    drawBox('✏️  在 这 里 画 / Draw here', 3800),
    new Paragraph({ spacing: { before: 160, after: 60 }, children: [new TextRun({ text: '✏️ 写一写 ', size: 22, bold: true, color: C_PICK }), new TextRun({ text: 'Write your answer:', size: 15, bold: true, italics: true, color: C_PICK })] }),
    new Paragraph({ spacing: { before: 0, after: 40 }, children: [
      new TextRun({ text: '💬 提示 ', size: 22, bold: true, color: C_PICK }),
      new TextRun({ text: a.frameZh, size: 26, bold: true, color: DARK }),
    ] }),
    new Paragraph({ spacing: { before: 0, after: 100 }, children: [new TextRun({ text: a.frameEn, size: 14, italics: true, color: GRAY })] }),
    writeLines(5),
  ];
}

// ===== §3 连一连 Match (我会认 ↔ emoji) =====
const PERM5 = [2, 0, 4, 1, 3];
function matchSection(day) {
  const words = day.match;
  const shuffled = PERM5.map((j) => words[j]);
  const colW = Math.floor(CW / 2);
  const rows = words.map((w, i) => {
    const right = shuffled[i];
    return new TableRow({ height: { value: 1900, rule: 'atLeast' }, children: [
      new TableCell({ width: { size: colW, type: WidthType.DXA }, borders: allBorders(noBorder()), margins: { top: 200, bottom: 200, left: 240, right: 240 }, verticalAlign: 'center',
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: w.cn, bold: true, size: 60, color: DARK })] })] }),
      new TableCell({ width: { size: colW, type: WidthType.DXA }, borders: allBorders(noBorder()), margins: { top: 200, bottom: 200, left: 240, right: 240 }, verticalAlign: 'center',
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: right.em, size: 64 })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: right.en, bold: true, size: 26, color: DARK })] }),
        ] }),
    ] });
  });
  return [
    sectionBreak(),
    shadedBar('三、连一连 / Match  (用线连起来)', C_MATCH, 24),
    new Paragraph({ spacing: { before: 120, after: 60 }, children: [new TextRun({ text: '👉 把词语和图画用线连起来。', size: 24, italics: true, color: GRAY }), new TextRun({ text: ' / Draw a line from each word to its picture.', size: 15, italics: true, color: GRAY })] }),
    new Paragraph({ spacing: { before: 200, after: 0 }, children: [new TextRun({ text: '' })] }),
    new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [colW, colW], borders: allBorders(noBorder()), rows }),
  ];
}

// ===== §4 描一描, 写一写 =====
function traceWrite(day) {
  return [
    sectionBreak(),
    shadedBar(`四、描一描, 写一写 / Trace and Write  (${day.writeChars})`, C_WRITE, 24),
    new Paragraph({ spacing: { before: 200, after: 80 }, children: [new TextRun({ text: `👉 在下面贴上写字纸, 写一写今天学到的字: ${day.writeChars}。`, size: 24, italics: true, color: GRAY })] }),
    new Paragraph({ spacing: { before: 40, after: 200 }, children: [new TextRun({ text: `Insert your writing paper below and practice: ${day.writeChars}.`, size: 16, italics: true, color: GRAY })] }),
    new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [CW], borders: allBorders(border(C_WRITE, 12)),
      rows: [new TableRow({ height: { value: 11000, rule: 'atLeast' }, children: [new TableCell({
        width: { size: CW, type: WidthType.DXA }, shading: { fill: 'FFFFFF', type: ShadingType.CLEAR }, margins: { top: 200, bottom: 200, left: 200, right: 200 }, verticalAlign: 'center',
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '📄  在这里贴上写字纸 / Insert your writing paper here', italics: true, color: LGRAY, size: 22 })] })],
      })] })] }),
  ];
}

// ===================================================================== DATA
const DAYS = require('./booklet_content.js');

DAYS.forEach((day) => {
  const doc = new Document({
    styles: { default: { document: { run: { font: 'Microsoft YaHei', size: 22 } } } },
    numbering: { config: [] },
    sections: [{ properties: { page: PAGE }, children: [
      ...cover(day), ...questionBank(day), ...applyActivity(day), ...matchSection(day), ...traceWrite(day),
    ] }],
  });
  Packer.toBuffer(doc).then((buf) => {
    const out = path.join(__dirname, `day${day.day}_booklet.docx`);
    fs.writeFileSync(out, buf);
    console.log('Created', out);
  });
});
