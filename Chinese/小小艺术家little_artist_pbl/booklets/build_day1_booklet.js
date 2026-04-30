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
  new Paragraph({ children: [new PageBreak()] }),
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
  // Question header
  blocks.push(new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [
      new TextRun({ text: `第 ${num} 题 / Question ${num}`, bold: true, size: 24, color: ACCENT }),
    ],
  }));
  // Photo placeholder
  blocks.push(photoBox(q.img, 1600, ACCENT));
  blocks.push(new Paragraph({ spacing: { before: 80, after: 80 }, children: [new TextRun('')] }));
  // Question text
  blocks.push(p(q.q, { bold: true, size: 24 }));
  // Options A/B/C
  ['A', 'B', 'C'].forEach((letter, i) => {
    blocks.push(new Paragraph({
      spacing: { before: 80, after: 80 },
      indent: { left: 360 },
      children: [
        new TextRun({ text: `☐  ${letter}.  `, size: 26, bold: true, color: DARK }),
        new TextRun({ text: q.options[i], size: 24 }),
      ],
    }));
  });
  // Spacer
  blocks.push(new Paragraph({ spacing: { before: 100 }, children: [new TextRun('')] }));
  return blocks;
}

const section1Children = [
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
section1Children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== Section 2: Subjective — pick favorite + draw =====
const section2Children = [
  shadedBar('二、我的最爱 / My Favorite Art', SKY, 24),
  new Paragraph({
    spacing: { before: 200, after: 200 },
    children: [new TextRun({ text: '👉 你最喜欢什么艺术？(可以选一个或多个)', size: 26, bold: true, color: DARK })],
  }),
  new Paragraph({
    spacing: { before: 80, after: 200 },
    children: [new TextRun({ text: 'What is your favorite art form? (Pick one or more)', size: 22, italics: true, color: GRAY })],
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
favOptions.forEach(opt => {
  section2Children.push(new Paragraph({
    spacing: { before: 80, after: 80 },
    indent: { left: 360 },
    children: [
      new TextRun({ text: '☐  ', size: 30, bold: true }),
      new TextRun({ text: `${opt.em}  `, size: 24 }),
      new TextRun({ text: `${opt.cn}  `, size: 26, bold: true, color: DARK }),
      new TextRun({ text: opt.en, size: 22, color: GRAY }),
    ],
  }));
});

// Draw section
section2Children.push(new Paragraph({ spacing: { before: 300 }, children: [new TextRun('')] }));
section2Children.push(new Paragraph({
  spacing: { before: 200, after: 200 },
  children: [new TextRun({ text: '🎨 画一画  Draw your favorite art', size: 26, bold: true, color: SKY })],
}));
section2Children.push(new Paragraph({
  spacing: { before: 80, after: 200 },
  children: [new TextRun({
    text: '👉 画你喜欢的艺术 (画画的你 / 跳舞 / 乐器 / 表演……都可以)',
    size: 22, italics: true, color: GRAY,
  })],
}));
// Big drawing frame
section2Children.push(new Table({
  width: { size: CW, type: WidthType.DXA },
  columnWidths: [CW],
  borders: allBorders(border(SKY, 12)),
  rows: [new TableRow({
    height: { value: 6800, rule: 'atLeast' },
    children: [new TableCell({
      width: { size: CW, type: WidthType.DXA },
      shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
      margins: { top: 120, bottom: 120, left: 120, right: 120 },
      verticalAlign: 'center',
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: '✏️  在这里画 / Draw here', italics: true, color: LGRAY, size: 22 })],
      })],
    })],
  })],
}));
section2Children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== Section 3: Match (连一连) — 我会认 5 chars =====
const matchWords = [
  { char: '开心', py: 'kāi xīn', en: 'happy', em: '😄' },
  { char: '难过', py: 'nán guò', en: 'sad', em: '😢' },
  { char: '生气', py: 'shēng qì', en: 'angry', em: '😡' },
  { char: '喜欢', py: 'xǐ huān', en: 'like / love', em: '😍' },
  { char: '心情', py: 'xīn qíng', en: 'mood', em: '🎭' },
];

// Build a 2-column match table: left = Chinese word, right = emoji+English (shuffled)
const matchShuffled = [matchWords[2], matchWords[0], matchWords[4], matchWords[1], matchWords[3]];
const matchRows = matchWords.map((w, i) => {
  const right = matchShuffled[i];
  const colW = Math.floor(CW / 2);
  return new TableRow({
    height: { value: 1100, rule: 'atLeast' },
    children: [
      // LEFT: word + pinyin
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
      // RIGHT: emoji + English
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
  shadedBar('三、连一连 / Match  (用线连起来)', CORAL, 24),
  new Paragraph({
    spacing: { before: 200, after: 300 },
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
  new Paragraph({ children: [new PageBreak()] }),
];

// ===== Section 4: Trace (描一描) — 我会写 3 chars =====
const traceChars = [
  { char: '开心', py: 'kāi xīn', en: 'happy' },
  { char: '喜欢', py: 'xǐ huān', en: 'like / love' },
  { char: '生气', py: 'shēng qì', en: 'angry' },
];

const section4Children = [
  shadedBar('四、描一描, 写一写 / Trace and Write', PURPLE, 24),
  new Paragraph({
    spacing: { before: 200, after: 300 },
    children: [new TextRun({
      text: '👉 先描一描灰色的字, 然后自己写一遍。',
      size: 22, italics: true, color: GRAY,
    })],
  }),
];

traceChars.forEach((w, idx) => {
  // Word title row
  section4Children.push(new Paragraph({
    spacing: { before: 320, after: 120 },
    children: [
      new TextRun({ text: `${idx + 1}. `, bold: true, size: 28, color: PURPLE }),
      new TextRun({ text: w.char, bold: true, size: 36, color: DARK }),
      new TextRun({ text: `   ${w.py}   `, size: 24, italics: true, color: GRAY }),
      new TextRun({ text: w.en, size: 22, color: GRAY }),
    ],
  }));
  // Trace row — 4 trace boxes (gray characters) + 4 blank boxes for student
  const traceCellW = Math.floor(CW / 8);
  const cells = [];
  // 4 trace cells (light gray character)
  for (let i = 0; i < 4; i++) {
    cells.push(new TableCell({
      width: { size: traceCellW, type: WidthType.DXA },
      borders: allBorders(border('CCCCCC', 6)),
      margins: { top: 120, bottom: 120, left: 60, right: 60 },
      verticalAlign: 'center',
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: w.char, size: 56, color: 'D8D8D8' })],
      })],
    }));
  }
  // 4 blank cells
  for (let i = 0; i < 4; i++) {
    cells.push(new TableCell({
      width: { size: traceCellW, type: WidthType.DXA },
      borders: allBorders(border('CCCCCC', 6)),
      margins: { top: 120, bottom: 120, left: 60, right: 60 },
      verticalAlign: 'center',
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: ' ', size: 56 })],
      })],
    }));
  }
  section4Children.push(new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: Array(8).fill(traceCellW),
    rows: [new TableRow({ height: { value: 1300, rule: 'atLeast' }, children: cells })],
  }));
});

// Final encouragement page
section4Children.push(new Paragraph({ spacing: { before: 600 }, children: [new TextRun('')] }));
section4Children.push(shadedBar('🎉  恭喜你完成 Day 1 练习册! Great job, Little Artist!', GREEN, 24));

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
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log(`Created ${OUT}`);
});
