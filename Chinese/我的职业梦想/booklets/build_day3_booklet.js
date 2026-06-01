// Build day3_booklet.docx — 我的职业梦想 Unit · Day 3: 小小企业家 / Little Entrepreneurs
// Modeled on 小小艺术家/booklets/build_day1_booklet.js
// Run: node build_day3_booklet.js

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
} = require('docx');

const OUT = path.join(__dirname, 'day3_booklet.docx');

// ===== Palette (Day 3 — Entrepreneur: amber/gold + money green) =====
const ACCENT = 'D48E1F';   // entrepreneur amber/gold — primary
const SKY    = 'F5C242';   // idea yellow
const CORAL  = '2E7D32';   // money green
const PURPLE = '6A1B9A';
const YELLOW = 'F5A623';
const GREEN  = '388E3C';
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

function photoBox(label, height = 1800, colorHex = 'BBBBBB') {
  // Transparent — light gray border only, no background fill.
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders({ style: BorderStyle.SINGLE, size: 6, color: colorHex }),
    rows: [new TableRow({
      height: { value: height, rule: 'atLeast' },
      children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        margins: { top: 200, bottom: 200, left: 160, right: 160 },
        verticalAlign: 'center',
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: `📷  ${label}`, color: GRAY, italics: true, size: 18 })],
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
    children: [new TextRun({ text: '💡', size: 120 })],
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
    children: [new TextRun({ text: 'Day 3 · 小小企业家', bold: true, size: 44, color: DARK })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 200 },
    children: [new TextRun({ text: 'Little Entrepreneurs', italics: true, size: 28, color: GRAY })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 800 },
    children: [new TextRun({ text: '📱 乔布斯 · 🦘 Lily · 🍿 9-yr CEO', size: 24, color: CORAL })],
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

// ===== §1 看图选择 / MC — match entrepreneur to their product =====
const mcQuestions = [
  {
    img: '一部 iPhone — 全功能一体小盒子  An iPhone — all-in-one little box',
    q: '这是谁发明的产品？  Who invented this product?',
    options: ['乔布斯 Steve Jobs', 'Lily Born', '爱迪生 Edison'],
  },
  {
    img: '三只脚的小杯 (袋鼠杯)  A 3-legged cup (Kangaroo Cup)',
    q: '这是谁 8 岁时设计的？  Who designed this at age 8?',
    options: ['乔布斯 Jobs', '爱迪生 Edison', 'Lily Born'],
  },
  {
    img: '一袋好吃的爆米花  A bag of yummy popcorn',
    q: '美国一个 9 岁男孩自己开了什么公司？  A 9-year-old US boy runs a company that sells:',
    options: ['手机 phones', '爆米花 popcorn', '杯子 cups'],
  },
  {
    img: '一位老爷爷喝水, 水洒了出来  Grandpa drinks water — it spills!',
    q: 'Lily 看见了什么「需要」？  What "need" did Lily see?',
    options: ['爷爷想睡觉', '爷爷手抖, 喝水洒出来', '爷爷想吃饭'],
  },
];

function mcQuestion(num, q) {
  // Two-column layout: LEFT = question + A/B/C; RIGHT = transparent photo box.
  const leftW  = Math.floor(CW * 0.55);
  const rightW = CW - leftW;
  const lightBorder = { style: BorderStyle.SINGLE, size: 6, color: 'BBBBBB' };

  const leftCell = new TableCell({
    width: { size: leftW, type: WidthType.DXA },
    borders: allBorders(noBorder()),
    margins: { top: 200, bottom: 200, left: 100, right: 240 },
    verticalAlign: 'center',
    children: [
      new Paragraph({
        spacing: { before: 0, after: 40 },
        children: [
          new TextRun({ text: `${num}.  `, bold: true, size: 26, color: DARK }),
          new TextRun({ text: q.q, bold: true, size: 22, color: DARK }),
        ],
      }),
      new Paragraph({
        spacing: { before: 100, after: 0 },
        indent: { left: 200 },
        children: [
          new TextRun({ text: '☐  A.  ', bold: true, size: 20, color: DARK }),
          new TextRun({ text: q.options[0], size: 20, color: DARK }),
        ],
      }),
      new Paragraph({
        spacing: { before: 60, after: 0 },
        indent: { left: 200 },
        children: [
          new TextRun({ text: '☐  B.  ', bold: true, size: 20, color: DARK }),
          new TextRun({ text: q.options[1], size: 20, color: DARK }),
        ],
      }),
      new Paragraph({
        spacing: { before: 60, after: 0 },
        indent: { left: 200 },
        children: [
          new TextRun({ text: '☐  C.  ', bold: true, size: 20, color: DARK }),
          new TextRun({ text: q.options[2], size: 20, color: DARK }),
        ],
      }),
    ],
  });

  const rightCell = new TableCell({
    width: { size: rightW, type: WidthType.DXA },
    borders: allBorders(lightBorder),
    margins: { top: 200, bottom: 200, left: 160, right: 160 },
    verticalAlign: 'center',
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: `📷  ${q.img}`, color: GRAY, italics: true, size: 16 })],
    })],
  });

  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [leftW, rightW],
    borders: allBorders(noBorder()),
    rows: [new TableRow({
      cantSplit: true,
      height: { value: 2000, rule: 'atLeast' },
      children: [leftCell, rightCell],
    })],
  });
}

const section1Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('一、看图选择 / Multiple Choice  (圈出正确答案)', ACCENT, 24),
  new Paragraph({
    spacing: { before: 200, after: 200 },
    children: [new TextRun({
      text: '看一看 — 这是谁的产品？谁看到了什么「需要」？',
      size: 22, italics: true, color: GRAY,
    })],
  }),
];
mcQuestions.forEach((q, i) => {
  section1Children.push(mcQuestion(i + 1, q));
  section1Children.push(new Paragraph({ spacing: { before: 120, after: 0 }, children: [new TextRun({ text: '' })] }));
});

// ===== §2 我最喜欢的产品 — favorite product + write + draw =====
const favOptions = [
  { em: '📱', cn: 'iPhone',         en: '乔布斯 · Steve Jobs' },
  { em: '🦘', cn: '三只脚 杯子',    en: 'Kangaroo Cup · Lily Born' },
  { em: '🍿', cn: '爆米花',         en: 'Popcorn · 9-yr CEO' },
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

function emptyCell() {
  return new TableCell({
    width: { size: Math.floor(CW / 2), type: WidthType.DXA },
    borders: allBorders(noBorder()),
    children: [new Paragraph('')],
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
  shadedBar('二、我最喜欢的 产品 / My Favorite Product', SKY, 22),
  new Paragraph({
    spacing: { before: 120, after: 60 },
    children: [new TextRun({ text: '👉 你 最喜欢 哪个 产品? 为什么?', size: 22, bold: true, color: DARK })],
  }),
  new Paragraph({
    spacing: { before: 0, after: 100 },
    children: [new TextRun({ text: 'Which product do you like best? Why?', size: 16, italics: true, color: GRAY })],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [Math.floor(CW / 2), Math.floor(CW / 2)],
    borders: allBorders(noBorder()),
    rows: [
      new TableRow({ children: [favCell(favOptions[0]), favCell(favOptions[1])] }),
      new TableRow({ children: [favCell(favOptions[2]), emptyCell()] }),
      new TableRow({ children: [otherCell()] }),
    ],
  }),
  // 写一写 — 2 fill-in lines
  new Paragraph({
    spacing: { before: 200, after: 60 },
    children: [
      new TextRun({ text: '✏️ 写一写  Write ', size: 22, bold: true, color: SKY }),
      new TextRun({ text: '— 把你的 答案 写在 横线上', size: 16, italics: true, color: GRAY }),
    ],
  }),
  writeLine('🌟 我 最喜欢的 产品 是:', ACCENT),
  writeLine('❤️ 因为:', CORAL),
  // 画一画 — big draw box (fills most of page 2)
  new Paragraph({
    spacing: { before: 160, after: 80 },
    children: [
      new TextRun({ text: '🎨 画一画  Draw ', size: 22, bold: true, color: SKY }),
      new TextRun({ text: '— 画 你 喜欢的 产品', size: 16, italics: true, color: GRAY }),
    ],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders({ style: BorderStyle.SINGLE, size: 6, color: 'BBBBBB' }),
    rows: [new TableRow({
      height: { value: 5400, rule: 'atLeast' },
      children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        verticalAlign: 'center',
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: '✏️  在这里画 / Draw here', italics: true, color: LGRAY, size: 18 })],
        })],
      })],
    })],
  }),
  // 写字横线 — 4 ruled lines below the box
  new Paragraph({
    spacing: { before: 200, after: 60 },
    children: [
      new TextRun({ text: '✏️ 写几句话  Write a few sentences', size: 18, bold: true, color: SKY }),
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

// ===== §3 连一连 / Match — 我会认: 企业家 产品 顾客 买 卖 =====
const matchWords = [
  { char: '企业家', py: 'qǐ yè jiā', en: 'entrepreneur', em: '💼' },
  { char: '产品',   py: 'chǎn pǐn',  en: 'product',      em: '📦' },
  { char: '顾客',   py: 'gù kè',     en: 'customer',     em: '👥' },
  { char: '买',     py: 'mǎi',       en: 'buy',          em: '🛒' },
  { char: '卖',     py: 'mài',       en: 'sell',         em: '🏪' },
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

// ===== §4 描一描, 写一写 (blank) =====
const section4Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('四、描一描, 写一写 / Trace and Write  (产品 · 买 · 卖)', PURPLE, 24),
  new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({
      text: '👉 在下面贴上写字纸, 写一写今天学到的字: 产品 · 买 · 卖。',
      size: 22, italics: true, color: GRAY,
    })],
  }),
  new Paragraph({
    spacing: { before: 60, after: 200 },
    children: [new TextRun({
      text: 'Insert your writing paper below and practice today’s characters: 产品 · 买 · 卖.',
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
