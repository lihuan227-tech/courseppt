// Build day4_booklet.docx — 我的职业梦想 Unit · Day 4: 社区小帮手 (Community Helpers)
// Structure mirrors build_day2_booklet.js — light gray borders, B&W-print friendly.
// Run: node build_day4_booklet.js

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
} = require('docx');

const OUT = path.join(__dirname, 'day4_booklet.docx');

// ===== Palette (Day 4 — Community Helpers: warm orange + helper red) =====
const ACCENT = 'D84315';   // helper-red — primary
const SKY    = 'FB8C00';   // warm orange
const CORAL  = 'C62828';   // emergency red
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
    children: [new TextRun({ text: '🚒', size: 120 })],
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
    children: [new TextRun({ text: 'Day 4 · 社区小帮手', bold: true, size: 44, color: DARK })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 200 },
    children: [new TextRun({ text: 'Community Helpers in Our Neighborhood',
      italics: true, size: 26, color: GRAY })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 800 },
    children: [new TextRun({ text: '👮  🚒  🩺  📚  📮', size: 30, color: CORAL })],
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

// ===== §1 选择题 / Multiple Choice =====
const qbQuestions = [
  {
    em: '💡',
    cn: '学 校 突 然 停 电 了, 你 最 希 望 谁 来 帮 忙?',
    en: 'The school suddenly lost power. Who would you call?',
    opts: [
      { cn: '电 工', en: 'Electrician' },
      { cn: '厨 师', en: 'Chef' },
      { cn: '兽 医', en: 'Veterinarian' },
    ],
    correct: 0,
  },
  {
    em: '💧',
    cn: '你 家 水 管 一 直 漏 水, 谁 最 能 帮 你?',
    en: 'Your sink is leaking. Who can help?',
    opts: [
      { cn: '水 管 工', en: 'Plumber' },
      { cn: '消 防 员', en: 'Firefighter' },
      { cn: '老 师', en: 'Teacher' },
    ],
    correct: 0,
  },
  {
    em: '👮',
    cn: '你 在 公 园 里 迷 路 了, 应 该 找 谁 帮 忙?',
    en: 'You are lost in a park. Who should you ask for help?',
    opts: [
      { cn: '警 察', en: 'Police Officer' },
      { cn: '厨 师', en: 'Chef' },
      { cn: '农 夫', en: 'Farmer' },
    ],
    correct: 0,
  },
  {
    em: '🍽️',
    cn: '餐 厅 来 了 很 多 客 人, 谁 最 忙?',
    en: 'A restaurant is full of customers. Who is probably the busiest?',
    opts: [
      { cn: '厨 师', en: 'Chef' },
      { cn: '飞 行 员', en: 'Pilot' },
      { cn: '邮 递 员', en: 'Mail Carrier' },
    ],
    correct: 0,
  },
  {
    em: '🚒',
    cn: '学 校 附 近 着 了 小 火, 谁 会 最 先 赶 来 帮 大 家?',
    en: 'A small fire starts near the school. Who will arrive to help first?',
    opts: [
      { cn: '消 防 员', en: 'Firefighter' },
      { cn: '医 生', en: 'Doctor' },
      { cn: '图 书 管 理 员', en: 'Librarian' },
    ],
    correct: 0,
  },
  {
    em: '😴',
    cn: '要 是 有 一 天, 所 有 社 区 工 作 者 都 放 假 了, 会 怎 么 样?',
    en: 'If all community helpers took a day off, what might happen?',
    opts: [
      { cn: '社 区 会 出 很 多 问 题', en: 'The community would face many problems.' },
      { cn: '什 么 都 不 会 变', en: 'Nothing would change.' },
      { cn: '大 家 都 会 变 成 老 师', en: 'Everyone would become a teacher.' },
    ],
    correct: 0,
  },
  {
    em: '🌟',
    cn: '社 区 工 作 者 为 什 么 都 很 重 要?',
    en: 'Why are all community helpers important?',
    opts: [
      { cn: '因 为 他 们 都 帮 助 社 区 变 得 更 好',
        en: 'Because they all help make the community better.' },
      { cn: '因 为 他 们 都 穿 制 服',
        en: 'Because they all wear uniforms.' },
      { cn: '因 为 他 们 都 开 汽 车',
        en: 'Because they all drive cars.' },
    ],
    correct: 0,
  },
  {
    em: '💌',
    cn: '你 想 感 谢 社 区 工 作 者, 可 以 说 什 么?',
    en: 'What can you say to thank a community helper?',
    opts: [
      { cn: '「谢 谢 你 帮 助 大 家!」', en: '"Thank you for helping everyone!"' },
      { cn: '「我 不 需 要 帮 助。」', en: "\"I don't need help.\"" },
      { cn: '「快 点 工 作 吧。」', en: '"Hurry up and work."' },
    ],
    correct: 0,
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
      new TextRun({ text: '答 对 8 题 = 「社 区 小 帮 手」 徽 章! ',
        bold: true, size: 20, color: ACCENT }),
      new TextRun({ text: 'All 8 right = Community Helper badge!',
        italics: true, size: 16, color: GRAY }),
    ],
  }),
];

// ===== §2 问答题 · 感谢社区小帮手 / Thank a Community Helper =====
const helperOptions = [
  { em: '👮',  cn: '警 察',     en: 'Police Officer' },
  { em: '🚒',  cn: '消 防 员', en: 'Firefighter' },
  { em: '🩺',  cn: '医 生',     en: 'Doctor' },
  { em: '📚',  cn: '老 师',     en: 'Teacher' },
  { em: '👨‍🍳', cn: '厨 师',    en: 'Chef' },
  { em: '📮',  cn: '邮 递 员', en: 'Mail Carrier' },
  { em: '🧹',  cn: '清 洁 工', en: 'Cleaner' },
];

function helperCheckCell(opt) {
  return new TableCell({
    width: { size: Math.floor(CW / 2), type: WidthType.DXA },
    borders: allBorders(noBorder()),
    margins: { top: 40, bottom: 40, left: 200, right: 100 },
    children: [new Paragraph({
      children: [
        new TextRun({ text: '☐  ', size: 24, bold: true }),
        new TextRun({ text: `${opt.em}  `, size: 22 }),
        new TextRun({ text: `${opt.cn}  `, size: 22, bold: true, color: DARK }),
        new TextRun({ text: opt.en, size: 16, color: GRAY }),
      ],
    })],
  });
}

function answerLinesTable(rowCount) {
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: {
      top: noBorder(), left: noBorder(), right: noBorder(),
      bottom: { style: BorderStyle.SINGLE, size: 8, color: '666666' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 8, color: '666666' },
      insideVertical: noBorder(),
    },
    rows: Array.from({ length: rowCount }, () => new TableRow({
      height: { value: 560, rule: 'atLeast' },
      children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        borders: {
          top: noBorder(), left: noBorder(), right: noBorder(),
          bottom: { style: BorderStyle.SINGLE, size: 8, color: '666666' },
        },
        margins: { top: 60, bottom: 60, left: 0, right: 0 },
        children: [new Paragraph({ children: [new TextRun({ text: '', size: 22 })] })],
      })],
    })),
  });
}

const otherCell = new TableCell({
  width: { size: Math.floor(CW / 2), type: WidthType.DXA },
  borders: allBorders(noBorder()),
  margins: { top: 40, bottom: 40, left: 200, right: 100 },
  children: [new Paragraph({
    children: [
      new TextRun({ text: '☐  ', size: 24, bold: true }),
      new TextRun({ text: '💫  ', size: 22 }),
      new TextRun({ text: '其 他  ', size: 22, bold: true, color: DARK }),
      new TextRun({ text: 'Other: ', size: 16, color: GRAY }),
      new TextRun({ text: '__________________', size: 20, color: GRAY }),
    ],
  })],
});

const section2Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('二、💌 感 谢 社 区 小 帮 手 / Thank a Community Helper', SKY, 22),

  // --- Pick a helper ---
  new Paragraph({
    spacing: { before: 240, after: 60 },
    children: [
      new TextRun({ text: '👉 你 最 想 感 谢 哪 一 位 社 区 工 作 者?', bold: true, size: 22, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 160 },
    children: [
      new TextRun({ text: 'Which community helper would you like to thank?',
        italics: true, size: 16, color: GRAY }),
    ],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [Math.floor(CW / 2), Math.floor(CW / 2)],
    borders: allBorders(noBorder()),
    rows: [
      new TableRow({ children: [helperCheckCell(helperOptions[0]), helperCheckCell(helperOptions[1])] }),
      new TableRow({ children: [helperCheckCell(helperOptions[2]), helperCheckCell(helperOptions[3])] }),
      new TableRow({ children: [helperCheckCell(helperOptions[4]), helperCheckCell(helperOptions[5])] }),
      new TableRow({ children: [helperCheckCell(helperOptions[6]), otherCell] }),
    ],
  }),

  // --- Drawing: express your thanks ---
  new Paragraph({
    spacing: { before: 320, after: 60 },
    children: [
      new TextRun({ text: '🎨 画 一 张 图 片, 表 示 你 的 感 谢',
        bold: true, size: 22, color: ACCENT }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 160 },
    children: [
      new TextRun({ text: 'Draw a picture to show your thanks.',
        italics: true, size: 16, color: GRAY }),
    ],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders({ style: BorderStyle.SINGLE, size: 6, color: 'BBBBBB' }),
    rows: [new TableRow({
      height: { value: 3600, rule: 'atLeast' },
      children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        verticalAlign: 'center',
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: '✏️  在 这 里 画 / Draw here', italics: true, color: LGRAY, size: 18 })],
        })],
      })],
    })],
  }),

  // --- Thank-you message ---
  new Paragraph({
    spacing: { before: 320, after: 60 },
    children: [
      new TextRun({ text: '✏️ 写 几 句 感 谢 的 话:', bold: true, size: 22, color: ACCENT }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 160 },
    children: [
      new TextRun({ text: 'Write a few sentences of thanks.',
        italics: true, size: 16, color: GRAY }),
    ],
  }),
  answerLinesTable(4),

  new Paragraph({
    spacing: { before: 360, after: 0 },
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({ text: '⭐ ', size: 24 }),
      new TextRun({ text: '谢 谢 所 有 帮 助 我 们 的 人!  ',
        bold: true, size: 22, color: CORAL }),
      new TextRun({ text: 'Thank You, Community Helpers!  🌟',
        italics: true, size: 16, color: GRAY }),
    ],
  }),
];

// ===== §3 连一连 / Match — no borders =====
const matchWords = [
  { char: '老 师',    em: '📚',  en: 'teacher' },
  { char: '学 校',    em: '🏫',  en: 'school' },
  { char: '医 院',    em: '🏥',  en: 'hospital' },
  { char: '消 防 员', em: '🚒',  en: 'firefighter' },
  { char: '厨 师',    em: '👨‍🍳', en: 'chef' },
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
  shadedBar('四、描一描, 写一写 / Trace and Write  (学 校 · 医 院)', PURPLE, 24),
  new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({
      text: '👉 在 下 面 贴 上 写 字 纸, 写 一 写 今 天 学 到 的 字: 学 校 · 医 院。',
      size: 22, italics: true, color: GRAY,
    })],
  }),
  new Paragraph({
    spacing: { before: 60, after: 200 },
    children: [new TextRun({
      text: "Insert your writing paper below and practice today's characters: 学 校 · 医 院.",
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
