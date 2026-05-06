// Build day1_booklet.docx — Wilderness Unit · Day 3: 方向地图
// Modeled on little_artist_pbl/booklets/build_day1_booklet.js
// Run: node build_day1_booklet.js

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
} = require('docx');

const OUT = path.join(__dirname, 'day3_booklet.docx');

// ===== Palette (Wilderness — adventure) =====
const PINE   = '2D5A3D';   // pine green — primary
const SUN    = 'E07A2C';   // sunset orange
const BROWN  = '6B4423';   // soil brown
const SKY    = '4A90D9';
const ALERT  = 'D04A3C';
const YELLOW = 'F5C242';
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
    children: [new TextRun({ text: '🧭', size: 120 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text: '野外生存与探险', bold: true, size: 60, color: PINE })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 600 },
    children: [new TextRun({ text: 'Wilderness Adventure', bold: true, size: 36, color: PINE })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text: 'Day 3 · 方向地图', bold: true, size: 44, color: DARK })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 800 },
    children: [new TextRun({ text: 'Compass & Direction', italics: true, size: 28, color: GRAY })],
  }),
  new Paragraph({
    spacing: { before: 1200, after: 200 },
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

// ===== Section 1: Multiple choice — identify environments / dangers =====
const mcQuestions = [
  {
    "img": "指南针 — 红针指北  Compass with red needle pointing north",
    "q": "指南针的红针指向哪里？  Which way does the red needle point?",
    "options": [
      "东 East",
      "北 North",
      "南 South"
    ]
  },
  {
    "img": "太阳从地平线升起  Sun rising at the horizon",
    "q": "太阳从哪个方向升起？  Where does the sun rise?",
    "options": [
      "东 East",
      "西 West",
      "北 North"
    ]
  },
  {
    "img": "一张地图 + 指南针  A map and compass",
    "q": "指南针是用来做什么的？  What is a compass for?",
    "options": [
      "做饭 Cooking",
      "找方向 Finding direction",
      "睡觉 Sleeping"
    ]
  },
  {
    "img": "太阳落山 — 红色的天空  Sunset",
    "q": "太阳从哪个方向落下？  Where does the sun set?",
    "options": [
      "东 East",
      "西 West",
      "南 South"
    ]
  }
];

function mcQuestion(num, q) {
  const blocks = [];
  blocks.push(new Paragraph({
    spacing: { before: 80, after: 40 },
    children: [new TextRun({ text: `第 ${num} 题 / Q ${num}`, bold: true, size: 20, color: PINE })],
  }));
  blocks.push(photoBox(q.img, 1100, PINE));
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
  shadedBar('一、看图选择 / Multiple Choice  (圈出正确答案)', PINE, 24),
  new Paragraph({
    spacing: { before: 200, after: 200 },
    children: [new TextRun({
      text: '看一看, 这是什么地方? 有什么危险? / Look — what place is this and what is the danger?',
      size: 22, italics: true, color: GRAY,
    })],
  }),
];
mcQuestions.forEach((q, i) => mcQuestion(i + 1, q).forEach(b => section1Children.push(b)));

// ===== Section 2: 我最爱的自然环境 — checkboxes + draw =====
const section2Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('二、没有指南针怎么找方向？ / Find Direction Without a Compass', SUN, 22),
  new Paragraph({
    spacing: { before: 100, after: 80 },
    children: [new TextRun({ text: '👉 如果你在森林里没有指南针, 你会用什么方法找到方向？写一写, 画一画。', size: 22, bold: true, color: DARK })],
  }),
  new Paragraph({
    spacing: { before: 40, after: 100 },
    children: [new TextRun({ text: 'If you have no compass in the forest, how would you find direction? Write and draw.', size: 16, italics: true, color: GRAY })],
  }),
];

section2Children.push(new Paragraph({
  spacing: { before: 200, after: 80 },
  children: [
    new TextRun({ text: '🎨 写一写 / 画一画  Write + Draw your method ', size: 22, bold: true, color: SUN }),
    new TextRun({ text: '— 例如: 看太阳、看苔藓、看星星……', size: 16, italics: true, color: GRAY }),
  ],
}));
section2Children.push(new Table({
  width: { size: CW, type: WidthType.DXA },
  columnWidths: [CW],
  borders: allBorders(border(SUN, 12)),
  rows: [new TableRow({
    height: { value: 5800, rule: 'atLeast' },
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
}));

// ===== Section 3: Match (连一连) — 5 read words =====
const matchWords = [
  {
    "char": "东",
    "py": "dōng",
    "en": "East (sunrise)",
    "em": "🌅"
  },
  {
    "char": "南",
    "py": "nán",
    "en": "South (warm)",
    "em": "🥵"
  },
  {
    "char": "西",
    "py": "xī",
    "en": "West (sunset)",
    "em": "🌇"
  },
  {
    "char": "北",
    "py": "běi",
    "en": "North (cold)",
    "em": "❄️"
  },
  {
    "char": "指南针",
    "py": "zhǐ nán zhēn",
    "en": "Compass",
    "em": "🧭"
  }
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
        borders: allBorders(border(BROWN, 8)),
        margins: { top: 200, bottom: 200, left: 240, right: 240 },
        verticalAlign: 'center',
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: w.char, bold: true, size: 56, color: DARK })] }),
          new Paragraph({ alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: w.py, size: 22, color: GRAY, italics: true })] }),
        ],
      }),
      new TableCell({
        width: { size: colW, type: WidthType.DXA },
        borders: allBorders(border(SUN, 8)),
        margins: { top: 200, bottom: 200, left: 240, right: 240 },
        verticalAlign: 'center',
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: right.em, size: 56 })] }),
          new Paragraph({ alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: right.en, bold: true, size: 26, color: DARK })] }),
        ],
      }),
    ],
  });
});

const section3Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('三、连一连 / Match  (用线连起来)', BROWN, 24),
  new Paragraph({
    spacing: { before: 200, after: 200 },
    children: [new TextRun({
      text: '👉 把中文词语和正确的图标/英文用一根线连起来。',
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

// ===== Section 4: Trace and Write — blank for writing paper =====
const section4Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('四、描一描, 写一写 / Trace and Write  (东  南  西  北)', ALERT, 24),
  new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({
      text: '👉 在下面贴上写字纸, 写一写今天学到的字: 东  南  西  北',
      size: 22, italics: true, color: GRAY,
    })],
  }),
  new Paragraph({
    spacing: { before: 60, after: 200 },
    children: [new TextRun({
      text: 'Insert your writing paper below and practice today’s characters.',
      size: 20, italics: true, color: GRAY,
    })],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders(border(ALERT, 12)),
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

// ===== Build doc =====
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
