// Build day1_booklet.docx — Wilderness Unit · Day 1: 认识自然
// Modeled on little_artist_pbl/booklets/build_day1_booklet.js
// Run: node build_day1_booklet.js

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
} = require('docx');

const OUT = path.join(__dirname, 'day1_booklet.docx');

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
    children: [new TextRun({ text: '🌲', size: 120 })],
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
    children: [new TextRun({ text: 'Day 1 · 认识自然', bold: true, size: 44, color: DARK })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 800 },
    children: [new TextRun({ text: 'Nature & Safety', italics: true, size: 28, color: GRAY })],
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

// ===== Section 1: Multiple choice — what's the danger in each place? =====
const mcQuestions = [
  {
    img: '🌾 草地 Grassland — 长长的草, 看不清地面',
    q: '草地上有什么危险？  What is the danger on the grassland?',
    options: ['虫子叮咬 Bug bites', '冻伤 Frostbite', '溺水 Drowning'],
  },
  {
    img: '🌲 森林 Forest — 树很多, 光线暗',
    q: '森林里有什么危险？  What is the danger in the forest?',
    options: ['中暑 Heatstroke', '迷路 / 野生动物 Lost / wild animals', '溺水 Drowning'],
  },
  {
    img: '🏞️ 河边 Riverside — 水流, 滑石头',
    q: '河边有什么危险？  What is the danger by the river?',
    options: ['冻伤 Frostbite', '虫子叮咬 Bug bites', '溺水 / 滑倒 Drowning / slipping'],
  },
  {
    img: '❄️ 雪地 Snow — 很冷, 地面滑',
    q: '雪地里有什么危险？  What is the danger in the snow?',
    options: ['冻伤 / 滑倒 Frostbite / slipping', '中暑 Heatstroke', '虫子叮咬 Bug bites'],
  },
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
  shadedBar('一、这个地方有什么危险？ / What\'s the Danger?  (圈一圈)', PINE, 24),
  new Paragraph({
    spacing: { before: 200, after: 200 },
    children: [new TextRun({
      text: '看一看图片 — 这个地方有什么危险？把对的圈一圈。/ Look — what is the danger here? Circle the right answer.',
      size: 22, italics: true, color: GRAY,
    })],
  }),
];
mcQuestions.forEach((q, i) => mcQuestion(i + 1, q).forEach(b => section1Children.push(b)));

// ===== Section 2: 我最爱的自然环境 — checkboxes + draw =====
const section2Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('二、我最爱的自然环境 / My Favorite Nature Place', SUN, 22),
  new Paragraph({
    spacing: { before: 100, after: 60 },
    children: [new TextRun({ text: '👉 你最喜欢哪种自然环境？(可以选一个或多个) 在这里要注意什么？', size: 22, bold: true, color: DARK })],
  }),
  new Paragraph({
    spacing: { before: 40, after: 100 },
    children: [new TextRun({ text: 'Which nature place do you like? (Pick one or more) — and what should you watch out for there?', size: 16, italics: true, color: GRAY })],
  }),
];

const favOptions = [
  { em: '🌲', cn: '森林', en: 'Forest' },
  { em: '🏔️', cn: '山地', en: 'Mountain' },
  { em: '🌾', cn: '草地', en: 'Grassland' },
  { em: '🏞️', cn: '河边', en: 'Riverside' },
  { em: '🏜️', cn: '沙漠', en: 'Desert' },
  { em: '❄️', cn: '雪地', en: 'Snow' },
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
section2Children.push(new Table({
  width: { size: CW, type: WidthType.DXA },
  columnWidths: [Math.floor(CW / 2), Math.floor(CW / 2)],
  borders: allBorders(noBorder()),
  rows: [
    new TableRow({ children: [favCell(favOptions[0]), favCell(favOptions[1])] }),
    new TableRow({ children: [favCell(favOptions[2]), favCell(favOptions[3])] }),
    new TableRow({ children: [favCell(favOptions[4]), favCell(favOptions[5])] }),
  ],
}));

section2Children.push(new Paragraph({
  spacing: { before: 200, after: 80 },
  children: [
    new TextRun({ text: '🎨 画一画 / 写一写  Draw + Write ', size: 22, bold: true, color: SUN }),
    new TextRun({ text: '— 画你喜欢的地方, 写一写要注意什么 (例如: 草地里要小心虫子)', size: 14, italics: true, color: GRAY }),
  ],
}));
section2Children.push(new Table({
  width: { size: CW, type: WidthType.DXA },
  columnWidths: [CW],
  borders: allBorders(border(SUN, 12)),
  rows: [new TableRow({
    height: { value: 3200, rule: 'atLeast' },
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
  { char: '森林', py: 'sēn lín', en: 'Forest', em: '🌲' },
  { char: '山地', py: 'shān dì', en: 'Mountain', em: '🏔️' },
  { char: '草地', py: 'cǎo dì', en: 'Grassland', em: '🌾' },
  { char: '河边', py: 'hé biān', en: 'Riverside', em: '🏞️' },
  { char: '沙漠', py: 'shā mò', en: 'Desert', em: '🏜️' },
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
  shadedBar('四、描一描, 写一写 / Trace and Write  (森林  山地  河边)', ALERT, 24),
  new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({
      text: '👉 在下面贴上写字纸, 写一写今天学到的字: 森林  山地  河边',
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
