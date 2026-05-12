// Build day2_booklet.docx — 我的职业梦想 Unit · Day 2: Problem Solver Day (科学家 + 工程师)
// Modeled on 小小艺术家/booklets/build_day1_booklet.js
// Run: node build_day2_booklet.js

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
} = require('docx');

const OUT = path.join(__dirname, 'day2_booklet.docx');

// ===== Palette (Day 2 — Problem Solver: science blue + engineering red) =====
const ACCENT = '1565C0';   // science blue — primary
const SKY    = '42A5F5';   // light blue
const CORAL  = 'E53935';   // engineering red
const PURPLE = '6A1B9A';
const YELLOW = 'F9A825';
const GREEN  = '43A047';
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
    children: [new TextRun({ text: 'Day 2 · Problem Solver Day', bold: true, size: 44, color: DARK })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 200 },
    children: [new TextRun({ text: '小小科学家 + 工程师', italics: true, size: 28, color: GRAY })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 800 },
    children: [new TextRun({ text: '🔬 为什么? · 🛠️ 怎么办?', size: 26, color: CORAL })],
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

// ===== §1 看图选择 / MC — 科学家 vs 工程师 by problem type =====
const mcQuestions = [
  {
    img: '一杯脏水, 想知道为什么会脏  A glass of dirty water — why is it dirty?',
    q: '这个问题 — 谁来研究？  Who studies this question?',
    options: ['科学家 Scientist', '工程师 Engineer', '医生 Doctor'],
  },
  {
    img: '河的两边, 要造一座桥  Two riverbanks — needs a bridge',
    q: '这个问题 — 谁来解决？  Who solves this problem?',
    options: ['科学家 Scientist', '老师 Teacher', '工程师 Engineer'],
  },
  {
    img: '一个 小朋友 睡着了, 上面 一个 梦境 泡泡  A child asleep with a dream bubble above',
    q: '为什么 我们 会 做梦? — 这个问题 谁 来 研究?  Why do we dream? — Who studies this?',
    options: ['老师 Teacher', '科学家 Scientist', '厨师 Chef'],
  },
  {
    img: '大太阳 + 暖暖的 阳光 + 笑脸  Big sun + warm sunshine + smiling face',
    q: '太阳 怎么 给 我们 温暖? — 谁 来 研究?  How does the sun warm us? — Who studies this?',
    options: ['工程师 Engineer (⚠️ trick)', '科学家 Scientist', '警察 Police'],
  },
];

function mcQuestion(num, q) {
  const blocks = [];
  blocks.push(new Paragraph({
    spacing: { before: 80, after: 40 },
    children: [new TextRun({ text: `第 ${num} 题 / Q ${num}`, bold: true, size: 20, color: ACCENT })],
  }));
  blocks.push(photoBox(q.img, 1100, ACCENT));
  blocks.push(new Paragraph({
    spacing: { before: 60, after: 20 },
    children: [new TextRun({ text: q.q, bold: true, size: 20 })],
  }));
  blocks.push(new Paragraph({
    spacing: { before: 0, after: 60 },
    indent: { left: 200 },
    children: [
      new TextRun({ text: '☐  A.  ', size: 20, bold: true, color: DARK }),
      new TextRun({ text: q.options[0], size: 18 }),
      new TextRun({ text: '     ', size: 18 }),
      new TextRun({ text: '☐  B.  ', size: 20, bold: true, color: DARK }),
      new TextRun({ text: q.options[1], size: 18 }),
      new TextRun({ text: '     ', size: 18 }),
      new TextRun({ text: '☐  C.  ', size: 20, bold: true, color: DARK }),
      new TextRun({ text: q.options[2], size: 18 }),
    ],
  }));
  return blocks;
}

const section1Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('一、看图选择 / Multiple Choice  (圈出正确答案)', ACCENT, 24),
  new Paragraph({
    spacing: { before: 200, after: 200 },
    children: [new TextRun({
      text: '看一看 — 科学家(为什么?) 还是 工程师(怎么办?)？/ Scientist (WHY) or Engineer (HOW)?',
      size: 22, italics: true, color: GRAY,
    })],
  }),
];
mcQuestions.forEach((q, i) => mcQuestion(i + 1, q).forEach(b => section1Children.push(b)));

// ===== §2 我喜欢的科学家 — favorite scientist + write + draw =====
const favOptions = [
  { em: '💡',  cn: '爱迪生',     en: 'Edison · 电灯' },
  { em: '✈️',  cn: '莱特兄弟',   en: 'Wright Brothers · 飞机' },
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
        new TextRun({ text: opt.en, size: 16, color: GRAY }),
      ],
    })],
  });
}

function otherCell() {
  return new TableCell({
    width: { size: CW, type: WidthType.DXA },
    columnSpan: 2,
    borders: allBorders(noBorder()),
    margins: { top: 60, bottom: 40, left: 200, right: 100 },
    children: [new Paragraph({
      children: [
        new TextRun({ text: '☐  ', size: 24, bold: true }),
        new TextRun({ text: '🤔  ', size: 22 }),
        new TextRun({ text: '其他  ', size: 22, bold: true, color: DARK }),
        new TextRun({ text: 'Other: ', size: 16, color: GRAY }),
        new TextRun({ text: '_______________________________________________', size: 22, color: GRAY }),
      ],
    })],
  });
}

function writeLine(label, color) {
  return new Paragraph({
    spacing: { before: 80, after: 40 },
    children: [
      new TextRun({ text: label, bold: true, size: 22, color }),
      new TextRun({ text: '  _________________________________________________________', size: 22, color: GRAY }),
    ],
  });
}

const section2Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('二、我喜欢的 科学家 / My Favorite Scientist', SKY, 22),
  new Paragraph({
    spacing: { before: 120, after: 60 },
    children: [new TextRun({ text: '👉 你 喜欢的 科学家 是谁? 他们 做了 什么? 为什么 你 喜欢?', size: 22, bold: true, color: DARK })],
  }),
  new Paragraph({
    spacing: { before: 0, after: 100 },
    children: [new TextRun({ text: 'Who is your favorite scientist? What did they do? Why do you like them?', size: 16, italics: true, color: GRAY })],
  }),
  // Example pills — Edison + Wright Brothers + write-in 其他
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [Math.floor(CW / 2), Math.floor(CW / 2)],
    borders: allBorders(noBorder()),
    rows: [
      new TableRow({ children: [favCell(favOptions[0]), favCell(favOptions[1])] }),
      new TableRow({ children: [otherCell()] }),
    ],
  }),
  // 写一写 — 3 fill-in lines
  new Paragraph({
    spacing: { before: 200, after: 60 },
    children: [
      new TextRun({ text: '✏️ 写一写  Write ', size: 22, bold: true, color: SKY }),
      new TextRun({ text: '— 把你的 答案 写在 横线上', size: 16, italics: true, color: GRAY }),
    ],
  }),
  writeLine('🌟 我 喜欢的 科学家 是:', SKY),
  writeLine('🛠️ 他 / 他们 做了:', GREEN),
  writeLine('❤️ 因为:', CORAL),
  // 画一画 — draw box
  new Paragraph({
    spacing: { before: 160, after: 80 },
    children: [
      new TextRun({ text: '🎨 画一画  Draw ', size: 22, bold: true, color: SKY }),
      new TextRun({ text: '— 画 ta / ta 的 发明 / ta 的 实验', size: 16, italics: true, color: GRAY }),
    ],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders(border(SKY, 12)),
    rows: [new TableRow({
      height: { value: 2400, rule: 'atLeast' },
      children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        verticalAlign: 'center',
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: '✏️  在这里画 / Draw here', italics: true, color: LGRAY, size: 18 })],
        })],
      })],
    })],
  }),
];

// ===== §3 连一连 / Match — 我会认: 科学家 工程师 实验 发明 观察 =====
const matchWords = [
  { char: '科学家', py: 'kē xué jiā',       en: 'scientist',  em: '👩‍🔬' },
  { char: '工程师', py: 'gōng chéng shī',   en: 'engineer',   em: '👷‍♂️' },
  { char: '实验',   py: 'shí yàn',          en: 'experiment', em: '🧪' },
  { char: '发明',   py: 'fā míng',          en: 'invention',  em: '💡' },
  { char: '观察',   py: 'guān chá',         en: 'observe',    em: '👀' },
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
        borders: allBorders(border(CORAL, 8)),
        margins: { top: 200, bottom: 200, left: 240, right: 240 },
        verticalAlign: 'center',
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: w.char, bold: true, size: 52, color: DARK })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: w.py, size: 22, color: GRAY, italics: true })],
          }),
        ],
      }),
      new TableCell({
        width: { size: colW, type: WidthType.DXA },
        borders: allBorders(border(SKY, 8)),
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
    spacing: { before: 200, after: 200 },
    children: [new TextRun({
      text: '👉 把中文词语和正确的英文/表情用一根线连起来。',
      size: 22, italics: true, color: GRAY,
    })],
  }),
  new Paragraph({
    spacing: { before: 80, after: 200 },
    children: [new TextRun({
      text: 'Draw a line from each Chinese word to its matching emoji + English.',
      size: 20, italics: true, color: GRAY,
    })],
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
  shadedBar('四、描一描, 写一写 / Trace and Write  (实验 · 发明 · 观察)', PURPLE, 24),
  new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({
      text: '👉 在下面贴上写字纸, 写一写今天学到的字: 实验 · 发明 · 观察。',
      size: 22, italics: true, color: GRAY,
    })],
  }),
  new Paragraph({
    spacing: { before: 60, after: 200 },
    children: [new TextRun({
      text: 'Insert your writing paper below and practice today’s characters: 实验 · 发明 · 观察.',
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
