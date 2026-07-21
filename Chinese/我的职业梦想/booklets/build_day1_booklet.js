// Build day1_booklet.docx — 我的职业梦想 Unit · Day 1: 认识职业世界 / Discover Careers
// Modeled on 小小艺术家/booklets/build_day1_booklet.js
// Run: node build_day1_booklet.js

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
} = require('docx');

const OUT = path.join(__dirname, 'day1_booklet.docx');

// ===== Palette (Day 1 — Career World: navy + gold) =====
const ACCENT = '1E3A5F';   // navy — primary (Interest cell)
const SKY    = '4A90E2';   // light blue
const CORAL  = 'F5A623';   // gold (career badge)
const PURPLE = '6A1B9A';
const YELLOW = 'F5C242';
const GREEN  = '43A047';   // Skill cell
const RED    = 'C5283C';   // Help cell (matches slide)
const DARK   = '2C2C2C';
const GRAY   = '888888';
const LGRAY  = 'D8D8D8';

// ===== Page geometry =====
const PAGE = {
  size: { width: 12240, height: 15840 },                         // US Letter
  margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },  // 0.75"
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

function photoBox(label, height = 1800, colorHex = LGRAY) {
  const b = border(colorHex, 8);
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders(b),
    rows: [new TableRow({
      height: { value: height, rule: 'atLeast' },
      children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        shading: { fill: 'F8F8F8', type: ShadingType.CLEAR },
        margins: { top: 200, bottom: 200, left: 160, right: 160 },
        verticalAlign: 'center',
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: `📷  ${label}`, color: GRAY, italics: true, size: 22 })],
        })],
      })],
    })],
  });
}

// ===== Cover =====
const coverChildren = [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 1200, after: 200 },
    children: [new TextRun({ text: '💼', size: 120 })],
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
    children: [new TextRun({ text: 'Day 1 · 认识职业世界', bold: true, size: 44, color: DARK })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 200 },
    children: [new TextRun({ text: 'Discover the World of Careers', italics: true, size: 28, color: GRAY })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 800 },
    children: [new TextRun({ text: '🩺 📚 👮 👨‍🍳 👷', size: 30, color: CORAL })],
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

// ===== §1 题库 · 8 个职业思考题 / Question Bank =====
// Single-column full-width cards with 3 options (A/B/C) each.
const qbCardColors = ['1565C0', '6A1B9A', '558B2F', 'F57C00', 'C5283C', '00897B', '7B5E3F', 'D81B60'];

const qbQuestions = [
  {
    em: '🌱',
    cn: '你 长 大 想 做 什 么 工 作? 怎 么 想, 才 能 想 对?',
    en: 'How can we figure out what job we want when we grow up?',
    opts: [
      { cn: '想 一 想 自 己 喜 欢 什 么、会 什 么、想 帮 谁',
        en: 'Think about what you like, what you can do, and who you want to help.' },
      { cn: '只 挑 最 能 赚 钱 的 工 作',
        en: 'Just pick the job that pays the most.' },
      { cn: '看 别 人 选 什 么, 自 己 也 跟 着 选',
        en: 'Pick whatever everyone else picks.' },
    ],
    correct: 0,
  },
  {
    em: '🐶',
    cn: '你 很 喜 欢 动 物, 还 想 照 顾 生 病 的 小 动 物 — 你 可 以 当 什 么?',
    en: 'You love animals and want to help sick ones — what could you be?',
    opts: [
      { cn: '兽 医', en: 'Vet' },
      { cn: '飞 行 员', en: 'Pilot' },
      { cn: '厨 师', en: 'Chef' },
    ],
    correct: 0,
  },
  {
    em: '👷',
    cn: '工 程 师 平 时 主 要 在 做 什 么?',
    en: 'What does an engineer mostly do?',
    opts: [
      { cn: '动 脑 筋 想 办 法, 把 问 题 解 决',
        en: 'Use their brain to solve problems.' },
      { cn: '只 在 工 地 搬 东 西',
        en: 'Just carry things at the construction site.' },
      { cn: '天 天 修 同 一 个 东 西',
        en: 'Fix the same thing every day.' },
    ],
    correct: 0,
  },
  {
    em: '💡',
    cn: '一 个 爱 好, 以 后 可 以 做 出 多 少 种 工 作?',
    en: 'One interest — how many jobs can it lead to?',
    opts: [
      { cn: '好 多 种, 都 不 一 样', en: 'Many different careers.' },
      { cn: '只 能 做 一 种', en: 'Only one career.' },
      { cn: '什 么 工 作 都 做 不 了', en: 'Cannot become a career.' },
    ],
    correct: 0,
  },
  {
    em: '🛠️',
    cn: '工 程 师 看 到 问 题 以 后, 接 下 来 会 做 什 么?',
    en: 'After an engineer spots a problem, what comes next?',
    opts: [
      { cn: '动 脑 筋, 想 办 法', en: 'Think of a way to fix it.' },
      { cn: '马 上 放 弃, 不 做 了', en: 'Give up right away.' },
      { cn: '不 再 试 了', en: 'Stop trying.' },
    ],
    correct: 0,
  },
  {
    em: '🌉',
    cn: '桥 和 大 楼, 一 般 是 谁 设 计 的?',
    en: 'Who usually designs bridges and big buildings?',
    opts: [
      { cn: '工 程 师', en: 'Engineer' },
      { cn: '兽 医', en: 'Vet' },
      { cn: '魔 术 师', en: 'Magician' },
    ],
    correct: 0,
  },
  {
    em: '🎮',
    cn: '游 戏 设 计 师 除 了 会 玩 游 戏, 还 要 会 些 什 么?',
    en: 'Besides playing games, what else does a game designer need to know?',
    opts: [
      { cn: '会 编 程, 会 画 画, 还 会 讲 故 事',
        en: 'Coding, drawing, and storytelling.' },
      { cn: '会 跑 得 很 快', en: 'Running fast.' },
      { cn: '会 做 很 多 种 菜', en: 'Cooking many dishes.' },
    ],
    correct: 0,
  },
  {
    em: '📚',
    cn: '要 是 一 个 老 师 都 没 有, 大 家 会 怎 么 样?',
    en: 'What if there were no teachers at all?',
    opts: [
      { cn: '很 多 人 想 学 东 西, 都 会 很 难',
        en: 'A lot of people would have trouble learning.' },
      { cn: '大 家 都 学 会 开 飞 机', en: 'Everyone would learn to fly planes.' },
      { cn: '学 校 全 变 成 餐 厅', en: 'Schools would all turn into restaurants.' },
    ],
    correct: 0,
  },
];

function qbCell(num, q, color) {
  const optionParas = q.opts.map((opt, idx) => {
    const letter = String.fromCharCode(65 + idx);  // A, B, C
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
  // Light gray thin border on all sides — B&W-print friendly. No background fill.
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
  children: [qbCell(i + 1, q, qbCardColors[i])],
}));

const section1Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('一、题库 · 8 个思考题 / Question Bank · 8 Reflection Questions  (圈出正确答案)', ACCENT, 24),
  new Paragraph({
    spacing: { before: 160, after: 160 },
    children: [new TextRun({
      text: '👉 读 一 读, 圈 出 对 的 答案。/ Read each question and circle the right answer.',
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
      new TextRun({ text: '答 对 8 题 = 「小 小 职 业 人」 徽 章! ', bold: true, size: 20, color: ACCENT }),
      new TextRun({ text: 'All 8 right = Future Career Hero badge!', italics: true, size: 16, color: GRAY }),
    ],
  }),
];

// ===== §2 我的最爱 — pick your dream job + draw yourself doing it =====
const favOptions = [
  { em: '🩺', cn: '医生',   en: 'Doctor' },
  { em: '📚', cn: '老师',   en: 'Teacher' },
  { em: '👮', cn: '警察',   en: 'Police' },
  { em: '👨‍🍳', cn: '厨师',   en: 'Chef' },
  { em: '👷', cn: '工程师', en: 'Engineer' },
  { em: '👩‍🔬', cn: '科学家', en: 'Scientist' },
];

function favCell(opt) {
  return new TableCell({
    width: { size: Math.floor(CW / 2), type: WidthType.DXA },
    borders: allBorders(noBorder()),
    margins: { top: 40, bottom: 40, left: 200, right: 100 },
    children: [new Paragraph({
      children: [
        new TextRun({ text: '☐  ', size: 24, bold: true }),
        new TextRun({ text: `${opt.em}  `, size: 22 }),
        new TextRun({ text: `${opt.cn}  `, size: 22, bold: true, color: DARK }),
        new TextRun({ text: opt.en, size: 18, color: GRAY }),
      ],
    })],
  });
}

const section2Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('二、我的梦想 / My Dream Job', SKY, 22),
  new Paragraph({
    spacing: { before: 100, after: 80 },
    children: [new TextRun({ text: '👉 你长大想当什么？(可以选一个或多个)', size: 24, bold: true, color: DARK })],
  }),
  new Paragraph({
    spacing: { before: 40, after: 100 },
    children: [new TextRun({ text: 'What do you want to be when you grow up? (Pick one or more)', size: 18, italics: true, color: GRAY })],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [Math.floor(CW / 2), Math.floor(CW / 2)],
    borders: allBorders(noBorder()),
    rows: [
      new TableRow({ children: [favCell(favOptions[0]), favCell(favOptions[1])] }),
      new TableRow({ children: [favCell(favOptions[2]), favCell(favOptions[3])] }),
      new TableRow({ children: [favCell(favOptions[4]), favCell(favOptions[5])] }),
    ],
  }),
  new Paragraph({
    spacing: { before: 200, after: 80 },
    children: [
      new TextRun({ text: '🎨 画 一 画  Draw yourself doing your dream job ', size: 22, bold: true, color: SKY }),
      new TextRun({ text: '— 画 工 作 的 你 (工 具 / 制 服 / 工 作 的 地 方)', size: 16, italics: true, color: GRAY }),
    ],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders({ style: BorderStyle.SINGLE, size: 6, color: 'BBBBBB' }),
    rows: [new TableRow({
      height: { value: 3200, rule: 'atLeast' },
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
  // "Write your answer" label
  new Paragraph({
    spacing: { before: 160, after: 60 },
    children: [
      new TextRun({ text: '✏️ 写 一 写 / Write your answer:', size: 16, bold: true, color: SKY }),
    ],
  }),
  // Sentence frame — placed UNDER 写一写
  new Paragraph({
    spacing: { before: 0, after: 40 },
    children: [
      new TextRun({ text: '💬 提 示  Sentence frame: ', size: 16, bold: true, color: SKY }),
      new TextRun({ text: '「我 长 大 了 想 当 ______, 因 为 ______。」', size: 18, bold: true, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 100 },
    children: [
      new TextRun({ text: '"When I grow up, I want to be ___, because ___."',
        size: 13, italics: true, color: GRAY }),
    ],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: {
      top: noBorder(), left: noBorder(), right: noBorder(),
      bottom: { style: BorderStyle.SINGLE, size: 8, color: '666666' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 8, color: '666666' },
      insideVertical: noBorder(),
    },
    rows: [1, 2, 3, 4].map(() => new TableRow({
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
  }),
];

// ===== §3 连一连 / Match — 我会认: 医生 老师 警察 厨师 工程师 =====
const matchWords = [
  { char: '医生',   py: 'yī shēng',        en: 'doctor',   em: '🩺' },
  { char: '老师',   py: 'lǎo shī',         en: 'teacher',  em: '📚' },
  { char: '警察',   py: 'jǐng chá',        en: 'police',   em: '👮' },
  { char: '厨师',   py: 'chú shī',         en: 'chef',     em: '👨‍🍳' },
  { char: '工程师', py: 'gōng chéng shī',  en: 'engineer', em: '👷' },
];

// Shuffle pairs so left-right indices don't match
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
            children: [new TextRun({ text: w.char, bold: true, size: 52, color: DARK })],
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

// ===== §4 描一描, 写一写 (blank) =====
const section4Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('四、描一描, 写一写 / Trace and Write  (医生 · 老师)', PURPLE, 24),
  new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({
      text: '👉 在下面贴上写字纸, 写一写今天学到的字: 医生 · 老师。',
      size: 22, italics: true, color: GRAY,
    })],
  }),
  new Paragraph({
    spacing: { before: 60, after: 200 },
    children: [new TextRun({
      text: 'Insert your writing paper below and practice today’s characters: 医生 · 老师.',
      size: 20, italics: true, color: GRAY,
    })],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders(border(PURPLE, 12)),
    rows: [new TableRow({
      height: { value: 11000, rule: 'atLeast' },
      children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
        margins: { top: 200, bottom: 200, left: 200, right: 200 },
        verticalAlign: 'center',
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: '📄  在这里贴上写字纸 / Insert your writing paper here',
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
