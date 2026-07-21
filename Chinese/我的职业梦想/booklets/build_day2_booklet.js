// Build day2_booklet.docx — 我的职业梦想 Unit · Day 2: Problem Solver Day
// (科学家 + 工程师 + 发明家)
// Structure mirrors build_day1_booklet.js — light gray borders, B&W-print friendly.
// Run: node build_day2_booklet.js

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
} = require('docx');

const OUT = path.join(__dirname, 'day2_booklet.docx');

// ===== Palette (Day 2 — Science/Engineer: navy + sky) =====
const ACCENT = '1565C0';   // science blue — primary
const SKY    = '42A5F5';   // light blue
const CORAL  = 'E53935';   // engineering red
const PURPLE = '6A1B9A';
const DARK   = '2C2C2C';
const GRAY   = '888888';
const LGRAY  = 'D8D8D8';

// ===== Page geometry =====
const PAGE = {
  size: { width: 12240, height: 15840 },
  margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
};
const CW = 12240 - 1080 - 1080;

// ===== Helpers =====
function border(color = 'CCCCCC', size = 4) {
  return { style: BorderStyle.SINGLE, size, color };
}
function noBorder() { return { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }; }
function allBorders(b) { return { top: b, bottom: b, left: b, right: b, insideHorizontal: b, insideVertical: b }; }

function shadedBar(text, colorHex, size = 22) {
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders(noBorder()),
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        shading: { fill: colorHex, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 200, right: 200 },
        children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: 'FFFFFF', size })] })],
      })],
    })],
  });
}

// ===== Cover =====
const coverChildren = [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 1200, after: 200 },
    children: [new TextRun({ text: '🔬', size: 120 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text: '我的职业梦想', bold: true, size: 60, color: ACCENT })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 600 },
    children: [new TextRun({ text: 'My Career Dream', bold: true, size: 36, color: ACCENT })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text: 'Day 2 · 小小问题解决家', bold: true, size: 44, color: DARK })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 200 },
    children: [new TextRun({ text: 'Problem-Solver Day · Scientists, Engineers, Inventors',
      italics: true, size: 26, color: GRAY })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 800 },
    children: [new TextRun({ text: '🔬  👷  💡  🧪  🏗️', size: 30, color: CORAL })],
  }),
  new Paragraph({
    spacing: { before: 1000, after: 200 },
    children: [
      new TextRun({ text: '姓名 Name: ', bold: true, size: 26 }),
      new TextRun({ text: '____________________________', size: 26, color: GRAY }),
    ],
  }),
  new Paragraph({
    spacing: { before: 200, after: 200 },
    children: [
      new TextRun({ text: '班级 Class: ', bold: true, size: 26 }),
      new TextRun({ text: '____________________________', size: 26, color: GRAY }),
    ],
  }),
  new Paragraph({
    spacing: { before: 200, after: 200 },
    children: [
      new TextRun({ text: '日期 Date: ', bold: true, size: 26 }),
      new TextRun({ text: '____________________________', size: 26, color: GRAY }),
    ],
  }),
];

// ===== §1 题库 · 8 道选择题 / Question Bank =====
const qbQuestions = [
  {
    em: '🔬',
    cn: '科 学 家 最 喜 欢 问 什 么?',
    en: 'What do scientists like to ask?',
    opts: [
      { cn: '多 少 钱?', en: 'How much does it cost?' },
      { cn: '为 什 么?', en: 'Why?' },
      { cn: '几 点 了?', en: 'What time is it?' },
    ],
    correct: 1,
  },
  {
    em: '👷',
    cn: '工 程 师 最 喜 欢 想 什 么?',
    en: 'What do engineers like to think about?',
    opts: [
      { cn: '怎 么 做?', en: 'How can we do it?' },
      { cn: '今 天 吃 什 么?', en: 'What should we eat today?' },
      { cn: '谁 跑 得 快?', en: 'Who can run faster?' },
    ],
    correct: 0,
  },
  {
    em: '💡',
    cn: '哪 一 位 发 明 家 发 明 了 实 用 电 灯?',
    en: 'Which inventor created the practical light bulb?',
    opts: [
      { cn: '贝 尔', en: 'Bell' },
      { cn: '爱 迪 生', en: 'Edison' },
      { cn: '牛 顿', en: 'Newton' },
    ],
    correct: 1,
  },
  {
    em: '☎️',
    cn: '贝 尔 最 有 名 的 发 明 是 什 么?',
    en: "What is Bell's most famous invention?",
    opts: [
      { cn: '飞 机', en: 'Airplane' },
      { cn: '电 话', en: 'Telephone' },
      { cn: '火 车', en: 'Train' },
    ],
    correct: 1,
  },
  {
    em: '🤔',
    cn: '科 学 家 和 工 程 师 都 有 的 特 点 是 什 么?',
    en: 'What trait do scientists and engineers share?',
    opts: [
      { cn: '爱 睡 觉', en: 'They like sleeping.' },
      { cn: '爱 玩 游 戏', en: 'They like playing games.' },
      { cn: '好 奇、爱 问 问 题', en: 'They are curious and ask questions.' },
    ],
    correct: 2,
  },
  {
    em: '🥤',
    cn: '做 纸 杯 电 话 的 时 候, 线 应 该 是 什 么 样?',
    en: 'What should the string be like when using a cup telephone?',
    opts: [
      { cn: '松 松 的', en: 'Loose.' },
      { cn: '拉 直 拉 紧', en: 'Straight and tight.' },
      { cn: '打 结', en: 'Tied in knots.' },
    ],
    correct: 1,
  },
  {
    em: '🃏',
    cn: '为 什 么 纸 牌 可 以 搭 成 「房 子」?',
    en: 'Why can cards be used to build a "house"?',
    opts: [
      { cn: '因 为 纸 牌 会 发 光', en: 'Because the cards glow.' },
      { cn: '因 为 折 起 来 以 后 更 稳, 更 能 撑 住',
        en: 'Because folded cards are stronger and more stable.' },
      { cn: '因 为 纸 牌 会 自 己 站 起 来', en: 'Because the cards stand up by themselves.' },
    ],
    correct: 1,
  },
  {
    em: '🧪',
    cn: '如 果 实 验 失 败 了, 应 该 怎 么 办?',
    en: 'What should you do if an experiment fails?',
    opts: [
      { cn: '马 上 放 弃', en: 'Give up right away.' },
      { cn: '再 试 一 次, 一 直 改 进', en: 'Try again and keep improving.' },
      { cn: '把 材 料 全 扔 掉', en: 'Throw the materials away.' },
    ],
    correct: 1,
  },
];

function qbCell(num, q) {
  const optionParas = q.opts.map((opt, idx) => {
    const letter = String.fromCharCode(65 + idx);
    return new Paragraph({
      spacing: { before: 60, after: 0 },
      indent: { left: 600 },
      children: [
        new TextRun({ text: '☐  ', bold: true, size: 20, color: DARK }),
        new TextRun({ text: `${letter}.  `, bold: true, size: 20, color: DARK }),
        new TextRun({ text: opt.cn, size: 20, color: DARK }),
        new TextRun({ text: `  ·  ${opt.en}`, italics: true, size: 15, color: GRAY }),
      ],
    });
  });
  const lightBorder = { style: BorderStyle.SINGLE, size: 4, color: 'BBBBBB' };
  return new TableCell({
    width: { size: CW, type: WidthType.DXA },
    borders: allBorders(lightBorder),
    margins: { top: 200, bottom: 200, left: 240, right: 240 },
    verticalAlign: 'center',
    children: [
      new Paragraph({
        spacing: { before: 0, after: 40 },
        children: [
          new TextRun({ text: `${num}.  `, bold: true, size: 26, color: DARK }),
          new TextRun({ text: `${q.em}  `, size: 26 }),
          new TextRun({ text: q.cn, bold: true, size: 22, color: DARK }),
        ],
      }),
      new Paragraph({
        spacing: { before: 0, after: 120 },
        indent: { left: 600 },
        children: [new TextRun({ text: q.en, italics: true, size: 16, color: GRAY })],
      }),
      ...optionParas,
    ],
  });
}

const qbRows = qbQuestions.map((q, i) => new TableRow({
  cantSplit: true,
  children: [qbCell(i + 1, q)],
}));

const section1Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('一、选 择 题 / Multiple Choice  (圈 出 正 确 答 案)', ACCENT, 24),
  new Paragraph({
    spacing: { before: 160, after: 160 },
    children: [new TextRun({
      text: '👉 读 一 读, 圈 出 对 的 答 案。/ Read each question and circle the right answer.',
      size: 22, italics: true, color: GRAY,
    })],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders(noBorder()),
    rows: qbRows,
  }),
  new Paragraph({
    spacing: { before: 240, after: 0 },
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({ text: '🏆 ', size: 22 }),
      new TextRun({ text: '答 对 8 题 = 「小 小 科 学 家」 徽 章! ',
        bold: true, size: 20, color: ACCENT }),
      new TextRun({ text: 'All 8 right = Junior Scientist badge!',
        italics: true, size: 16, color: GRAY }),
    ],
  }),
];

// ===== §2 纸桥挑战实验记录 / Paper Bridge Challenge =====
const lightCellBorder = { style: BorderStyle.SINGLE, size: 4, color: 'BBBBBB' };

// 2x2 grid of 4 bridge shape options with checkbox + photo placeholder
const bridgeShapes = [
  { em: '📦', cn: '中 空 形',     en: 'Hollow Shape' },
  { em: '〰️', cn: 'W 形',        en: 'W Shape' },
  { em: '⊓',  cn: '凹 凸 形',     en: 'Corrugated Shape' },
  { em: '━',  cn: '平 面 纸 桥',  en: 'Flat Paper Bridge' },
];

function bridgePredictionCell(s) {
  return new TableCell({
    width: { size: Math.floor(CW / 2), type: WidthType.DXA },
    borders: allBorders(lightCellBorder),
    margins: { top: 180, bottom: 180, left: 240, right: 240 },
    verticalAlign: 'center',
    children: [
      new Paragraph({
        spacing: { before: 0, after: 60 },
        children: [
          new TextRun({ text: '☐  ', bold: true, size: 26, color: DARK }),
          new TextRun({ text: `${s.em}  `, size: 24 }),
          new TextRun({ text: s.cn, bold: true, size: 22, color: DARK }),
        ],
      }),
      new Paragraph({
        spacing: { before: 0, after: 100 },
        indent: { left: 600 },
        children: [new TextRun({ text: s.en, italics: true, size: 15, color: GRAY })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 80, after: 0 },
        children: [new TextRun({ text: '🖼️  插 入 照 片 / Insert photo', size: 14, italics: true, color: LGRAY })],
      }),
    ],
  });
}

// Experiment record table (4 rows × 2 cols)
const recordRows = [
  new TableRow({
    children: [
      new TableCell({
        width: { size: Math.floor(CW * 0.55), type: WidthType.DXA },
        borders: allBorders(lightCellBorder),
        shading: { fill: 'F0F0F0', type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 240, right: 240 },
        children: [new Paragraph({
          children: [new TextRun({ text: '形 状  Shape', bold: true, size: 22, color: DARK })],
        })],
      }),
      new TableCell({
        width: { size: Math.floor(CW * 0.45), type: WidthType.DXA },
        borders: allBorders(lightCellBorder),
        shading: { fill: 'F0F0F0', type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 240, right: 240 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: '承 载 硬 币 数  # of Coins', bold: true, size: 22, color: DARK })],
        })],
      }),
    ],
  }),
  ...bridgeShapes.map(s => new TableRow({
    height: { value: 700, rule: 'atLeast' },
    children: [
      new TableCell({
        width: { size: Math.floor(CW * 0.55), type: WidthType.DXA },
        borders: allBorders(lightCellBorder),
        margins: { top: 140, bottom: 140, left: 240, right: 240 },
        verticalAlign: 'center',
        children: [new Paragraph({
          children: [
            new TextRun({ text: `${s.em}  `, size: 22 }),
            new TextRun({ text: s.cn, bold: true, size: 22, color: DARK }),
            new TextRun({ text: `  ·  ${s.en}`, italics: true, size: 14, color: GRAY }),
          ],
        })],
      }),
      new TableCell({
        width: { size: Math.floor(CW * 0.45), type: WidthType.DXA },
        borders: allBorders(lightCellBorder),
        margins: { top: 140, bottom: 140, left: 240, right: 240 },
        verticalAlign: 'center',
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: '_____________', size: 22, color: GRAY })],
        })],
      }),
    ],
  })),
];

// "Why stronger" checkboxes
const whyOptions = [
  { cn: '更 厚', en: 'Thicker' },
  { cn: '有 更 多 支 撑', en: 'More Support' },
  { cn: '不 容 易 弯', en: 'Harder to Bend' },
];

function answerLine() {
  return new Paragraph({
    spacing: { before: 200, after: 0 },
    border: { bottom: { color: '888888', size: 8, space: 1, style: BorderStyle.SINGLE } },
    children: [new TextRun({ text: '', size: 22 })],
  });
}

const section2Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('二、🏗️ 纸 桥 挑 战 实 验 记 录 表 / Paper Bridge Challenge', SKY, 22),

  // --- 1. Prediction ---
  new Paragraph({
    spacing: { before: 200, after: 60 },
    children: [
      new TextRun({ text: '1.  我 的 预 测  My Prediction', bold: true, size: 24, color: ACCENT }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 40 },
    children: [
      new TextRun({ text: '👉 你 觉 得 哪 一 种 桥 最 坚 固?', bold: true, size: 22, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 160 },
    children: [
      new TextRun({ text: 'Which bridge do you think will be the strongest?',
        italics: true, size: 16, color: GRAY }),
    ],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [Math.floor(CW / 2), Math.floor(CW / 2)],
    borders: allBorders(noBorder()),
    rows: [
      new TableRow({ children: [bridgePredictionCell(bridgeShapes[0]), bridgePredictionCell(bridgeShapes[1])] }),
      new TableRow({ children: [bridgePredictionCell(bridgeShapes[2]), bridgePredictionCell(bridgeShapes[3])] }),
    ],
  }),
  new Paragraph({
    spacing: { before: 240, after: 60 },
    children: [
      new TextRun({ text: '为 什 么?  Why?', bold: true, size: 20, color: ACCENT }),
    ],
  }),
  answerLine(),
  answerLine(),
  answerLine(),

  // --- 2. Experiment Record ---
  new Paragraph({
    spacing: { before: 320, after: 60 },
    children: [
      new TextRun({ text: '2.  实 验 记 录  Experiment Results', bold: true, size: 24, color: ACCENT }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 40 },
    children: [
      new TextRun({ text: '👉 试 一 试 — 每 一 种 桥 最 多 能 压 上 几 个 硬 币?',
        bold: true, size: 22, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 160 },
    children: [
      new TextRun({ text: 'Test how many coins each bridge can hold.',
        italics: true, size: 16, color: GRAY }),
    ],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [Math.floor(CW * 0.55), Math.floor(CW * 0.45)],
    borders: allBorders(lightCellBorder),
    rows: recordRows,
  }),

  // --- 3. Result ---
  new Paragraph({
    spacing: { before: 320, after: 60 },
    children: [
      new TextRun({ text: '3.  实 验 结 果  Results', bold: true, size: 24, color: ACCENT }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 60 },
    children: [
      new TextRun({ text: '👉 哪 一 种 桥 最 坚 固?', bold: true, size: 22, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 120 },
    children: [
      new TextRun({ text: 'Which bridge was the strongest?', italics: true, size: 16, color: GRAY }),
    ],
  }),
  ...bridgeShapes.map(s => new Paragraph({
    spacing: { before: 50, after: 0 },
    indent: { left: 600 },
    children: [
      new TextRun({ text: '☐  ', bold: true, size: 22, color: DARK }),
      new TextRun({ text: `${s.em}  `, size: 20 }),
      new TextRun({ text: s.cn, bold: true, size: 20, color: DARK }),
      new TextRun({ text: `  ·  ${s.en}`, italics: true, size: 14, color: GRAY }),
    ],
  })),
  new Paragraph({
    spacing: { before: 200, after: 40 },
    children: [
      new TextRun({ text: '它 一 共 压 了 ', size: 22, color: DARK }),
      new TextRun({ text: '______', bold: true, size: 24, color: ACCENT }),
      new TextRun({ text: ' 个 硬 币。', size: 22, color: DARK }),
      new TextRun({ text: '  ·  It held ______ coins.', italics: true, size: 14, color: GRAY }),
    ],
  }),

  // --- 4. My Discovery ---
  new Paragraph({
    spacing: { before: 320, after: 60 },
    children: [
      new TextRun({ text: '4.  我 的 发 现  My Discovery', bold: true, size: 24, color: ACCENT }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 60 },
    children: [
      new TextRun({ text: '👉 它 为 什 么 更 坚 固?', bold: true, size: 22, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 120 },
    children: [
      new TextRun({ text: 'Why was this bridge stronger?', italics: true, size: 16, color: GRAY }),
    ],
  }),
  ...whyOptions.map(o => new Paragraph({
    spacing: { before: 50, after: 0 },
    indent: { left: 600 },
    children: [
      new TextRun({ text: '☐  ', bold: true, size: 22, color: DARK }),
      new TextRun({ text: o.cn, bold: true, size: 20, color: DARK }),
      new TextRun({ text: `  ·  ${o.en}`, italics: true, size: 14, color: GRAY }),
    ],
  })),
  new Paragraph({
    spacing: { before: 50, after: 0 },
    indent: { left: 600 },
    children: [
      new TextRun({ text: '☐  ', bold: true, size: 22, color: DARK }),
      new TextRun({ text: '其 他  Other: ', bold: true, size: 20, color: DARK }),
      new TextRun({ text: '______________________________', size: 20, color: GRAY }),
    ],
  }),

  // Final badge
  new Paragraph({
    spacing: { before: 320, after: 0 },
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({ text: '⭐ ', size: 24 }),
      new TextRun({ text: '我 是 小 小 工 程 师!  ', bold: true, size: 22, color: CORAL }),
      new TextRun({ text: "I Am a Young Engineer!", italics: true, size: 16, color: GRAY }),
    ],
  }),
];

// ===== §3 连一连 / Match — no borders =====
const matchWords = [
  { char: '科 学 家', em: '🔬', en: 'scientist' },
  { char: '工 程 师', em: '👷', en: 'engineer' },
  { char: '发 明 家', em: '💡', en: 'inventor' },
  { char: '电 话',    em: '☎️', en: 'telephone' },
  { char: '实 验',    em: '🧪', en: 'experiment' },
];

const matchShuffled = [matchWords[2], matchWords[0], matchWords[4], matchWords[1], matchWords[3]];
const matchRows = matchWords.map((w, i) => {
  const right = matchShuffled[i];
  const colW = Math.floor(CW / 2);
  return new TableRow({
    height: { value: 1100, rule: 'atLeast' },
    children: [
      new TableCell({
        width: { size: colW, type: WidthType.DXA },
        borders: allBorders(noBorder()),
        margins: { top: 200, bottom: 200, left: 240, right: 240 },
        verticalAlign: 'center',
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: w.char, bold: true, size: 48, color: DARK })],
          }),
        ],
      }),
      new TableCell({
        width: { size: colW, type: WidthType.DXA },
        borders: allBorders(noBorder()),
        margins: { top: 200, bottom: 200, left: 240, right: 240 },
        verticalAlign: 'center',
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: right.em, size: 56 })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: right.en, bold: true, size: 26, color: DARK })],
          }),
        ],
      }),
    ],
  });
});

const section3Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('三、连一连 / Match  (用线连起来)', CORAL, 24),
  new Paragraph({
    spacing: { before: 300, after: 0 },
    children: [new TextRun({ text: '' })],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [Math.floor(CW / 2), Math.floor(CW / 2)],
    borders: allBorders(noBorder()),
    rows: matchRows,
  }),
];

// ===== §4 描一描, 写一写 =====
const section4Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('四、描一描, 写一写 / Trace and Write  (科 学 · 实 验)', PURPLE, 24),
  new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({
      text: '👉 在 下 面 贴 上 写 字 纸, 写 一 写 今 天 学 到 的 字: 科 学 · 实 验。',
      size: 22, italics: true, color: GRAY,
    })],
  }),
  new Paragraph({
    spacing: { before: 60, after: 200 },
    children: [new TextRun({
      text: "Insert your writing paper below and practice today's characters: 科 学 · 实 验.",
      size: 20, italics: true, color: GRAY,
    })],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders({ style: BorderStyle.SINGLE, size: 6, color: 'BBBBBB' }),
    rows: [new TableRow({
      height: { value: 11000, rule: 'atLeast' },
      children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        margins: { top: 200, bottom: 200, left: 200, right: 200 },
        verticalAlign: 'center',
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: '📄  在 这 里 贴 上 写 字 纸 / Insert your writing paper here',
            italics: true, color: LGRAY, size: 22,
          })],
        })],
      })],
    })],
  }),
];

// ===== Build =====
const doc = new Document({
  styles: { default: { document: { run: { font: 'Microsoft YaHei', size: 22 } } } },
  numbering: { config: [] },
  sections: [{
    properties: { page: PAGE },
    children: [
      ...coverChildren,
      ...section1Children,
      ...section2Children,
      ...section3Children,
      ...section4Children,
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log(`Created ${OUT}`);
});
