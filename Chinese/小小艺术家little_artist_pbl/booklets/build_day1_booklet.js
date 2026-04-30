// Build day1_booklet.docx — Little Artist Unit · Day 1: 艺术是表达
// Modeled on world_trip_pbl/booklets/build_booklet_docx.js style.
// Run: node build_day1_booklet.js

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, PageBreak,
  HeadingLevel, LevelFormat, TabStopType, TabStopPosition,
} = require('docx');

const OUT = path.join(__dirname, 'day1_booklet.docx');

// ===== Palette (Little Artist — playful) =====
const ACCENT = 'D81B60';   // magenta — primary
const SKY    = '42A5F5';   // sky blue
const CORAL  = 'FF7043';   // coral
const PURPLE = '9C27B0';
const YELLOW = 'FFC107';
const GREEN  = '66BB6A';
const DARK   = '2C2C2C';
const GRAY   = '888888';
const LGRAY  = 'D8D8D8';

// ===== Page geometry =====
const PAGE = {
  size: { width: 12240, height: 15840 },                         // US Letter
  margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },  // 0.75"
};
const CW = 12240 - 1080 - 1080; // 10080 DXA content width

// ===== Helpers =====
function border(color = 'CCCCCC', size = 4) {
  return { style: BorderStyle.SINGLE, size, color };
}
function noBorder() { return { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }; }
function allBorders(b) { return { top: b, bottom: b, left: b, right: b, insideHorizontal: b, insideVertical: b }; }

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, size: 22, ...opts })],
  });
}
function pBold(text, color = DARK) { return p(text, { bold: true, color }); }

// Colored shaded title bar
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

// Photo placeholder box (colored border, light fill, label inside)
function photoBox(label, height = 1800, colorHex = LGRAY) {
  const b = border(colorHex, 8);
  const heightDxa = height; // pass DXA directly
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders(b),
    rows: [new TableRow({
      height: { value: heightDxa, rule: 'atLeast' },
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

// Single-cell card with rounded-feel — colored border + body content
function card(children, colorHex = ACCENT) {
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders(border(colorHex, 8)),
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        margins: { top: 160, bottom: 160, left: 200, right: 200 },
        children,
      })],
    })],
  });
}

// Lines for writing space (3-4 thin underlines)
function writingLines(count = 3) {
  const cells = [];
  for (let i = 0; i < count; i++) {
    cells.push(new TableRow({
      height: { value: 600, rule: 'atLeast' },
      children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        borders: { top: noBorder(), bottom: border('999999', 4), left: noBorder(), right: noBorder() },
        children: [new Paragraph('')],
      })],
    }));
  }
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders(noBorder()),
    rows: cells,
  });
}

// ===== Cover page =====
const coverChildren = [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 1200, after: 200 },
    children: [new TextRun({ text: '🎨', size: 120 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text: '小小艺术家', bold: true, size: 60, color: ACCENT })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 600 },
    children: [new TextRun({ text: 'Little Artist', bold: true, size: 36, color: ACCENT })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text: 'Day 1 · 艺术是表达', bold: true, size: 44, color: DARK })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 800 },
    children: [new TextRun({ text: 'Art is Expression', italics: true, size: 28, color: GRAY })],
  }),
  // Name + class lines
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

// ===== Section 1: Multiple choice =====
const mcQuestions = [
  {
    img: '《蒙娜丽莎》Mona Lisa',
    q: '这是什么画？  What kind of painting is this?',
    options: ['油画 Oil Painting', '素描 Pencil Sketch', '拼贴画 Collage'],
  },
  {
    img: '铅笔画 — 小猫 / 苹果  Pencil drawing of a cat / apple',
    q: '这是什么画？  What kind of painting is this?',
    options: ['水彩画 Watercolor', '素描 Pencil Sketch', '油画 Oil Painting'],
  },
  {
    img: '颜色很淡、透明的花  Soft, transparent flowers',
    q: '这是什么画？  What kind of painting is this?',
    options: ['水彩画 Watercolor', '拼贴画 Collage', '素描 Pencil Sketch'],
  },
  {
    img: '《好饿的毛毛虫》风格  The Very Hungry Caterpillar style',
    q: '这是什么画？  What kind of painting is this?',
    options: ['拼贴画 Collage', '油画 Oil Painting', '水墨画 Ink Painting'],
  },
];

function mcQuestion(num, q) {
  const blocks = [];
  blocks.push(new Paragraph({
    spacing: { before: 80, after: 40 },
    children: [
      new TextRun({ text: `第 ${num} 题 / Q ${num}`, bold: true, size: 20, color: ACCENT }),
    ],
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
      text: '看一看, 这是什么画？/ Look — what kind of painting is it?',
      size: 22, italics: true, color: GRAY,
    })],
  }),
];
mcQuestions.forEach((q, i) => mcQuestion(i + 1, q).forEach(b => section1Children.push(b)));

// ===== Section 2: Subjective — pick favorite + small draw (compact, share page with §3) =====
const section2Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('二、我的最爱 / My Favorite Art', SKY, 22),
  new Paragraph({
    spacing: { before: 100, after: 80 },
    children: [new TextRun({ text: '👉 你最喜欢什么艺术？(可以选一个或多个)', size: 24, bold: true, color: DARK })],
  }),
  new Paragraph({
    spacing: { before: 40, after: 100 },
    children: [new TextRun({ text: 'What is your favorite art form? (Pick one or more)', size: 18, italics: true, color: GRAY })],
  }),
];

const favOptions = [
  { em: '🎨', cn: '画画', en: 'Drawing' },
  { em: '🎵', cn: '音乐', en: 'Music' },
  { em: '💃', cn: '舞蹈', en: 'Dance' },
  { em: '🎭', cn: '戏剧', en: 'Drama' },
  { em: '🎬', cn: '电影', en: 'Film' },
  { em: '🗿', cn: '雕塑', en: 'Sculpture' },
];
// 6 checkboxes in a 2-column table to save vertical space
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

// Draw section — compact frame so §2+§3 fit one page
section2Children.push(new Paragraph({
  spacing: { before: 200, after: 80 },
  children: [
    new TextRun({ text: '🎨 画一画  Draw your favorite art ', size: 22, bold: true, color: SKY }),
    new TextRun({ text: '— 画画的你 / 跳舞 / 乐器 / 表演……', size: 16, italics: true, color: GRAY }),
  ],
}));
section2Children.push(new Table({
  width: { size: CW, type: WidthType.DXA },
  columnWidths: [CW],
  borders: allBorders(border(SKY, 12)),
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
// NO page break — §3 follows on same page

// (Section 3 — old trace section removed; now just an empty array)
const section3Children = [];

// ===== Section 4: Match (连一连) — 我会认 5 chars =====
const matchWords = [
  { char: '开心', py: 'kāi xīn', en: 'happy', em: '😄' },
  { char: '难过', py: 'nán guò', en: 'sad', em: '😢' },
  { char: '生气', py: 'shēng qì', en: 'angry', em: '😡' },
  { char: '喜欢', py: 'xǐ huān', en: 'like / love', em: '😍' },
  { char: '心情', py: 'xīn qíng', en: 'mood', em: '🎭' },
];

// 2-column match table: left = Chinese word, right = emoji+English (shuffled)
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
            children: [new TextRun({ text: w.char, bold: true, size: 56, color: DARK })],
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

const section4Children = [
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

// ===== Section 5: Trace and Write (blank — teacher inserts writing paper) =====
const section5Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('四、描一描, 写一写 / Trace and Write', PURPLE, 24),
  new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({
      text: '👉 在下面贴上写字纸, 写一写今天学到的字。',
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
  // Big blank framed area for teacher to insert writing paper
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

// ===== Build doc =====
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: 'Microsoft YaHei', size: 22 } },
    },
  },
  numbering: { config: [] },
  sections: [{
    properties: { page: PAGE },
    children: [
      ...coverChildren,
      ...section1Children,
      ...section2Children,
      ...section3Children,
      ...section4Children,
      ...section5Children,
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log(`Created ${OUT}`);
});
