// Build day2_booklet.docx — 小小生活家 Little Life Helper · Day 2: 整理和收纳达人
// Structure mirrors the Day 1 kitchen booklet (中低 level, 4 pages):
//   p1 Cover (unit coloring cover + 姓名/班级)
//   p2 一、选一选  — bilingual multiple choice from the Day 2 deck
//   p3 二、连一连  — match word ↔ picture/English
//   p4 三、描一描, 写一写 — 我会写: 分类 · 书包 · 收拾 (trace page image)
// Content source: day2_organizing-手工活动已加.pdf (70-slide final deck)
// Run: node make_trace_page.js && node build_day2_booklet.js
const fs = require('fs');
const path = require('path');
const REPO = path.resolve(__dirname, '../../..');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
} = require(path.join(REPO, 'node_modules/docx'));

const OUT = path.join(__dirname, 'day2_booklet.docx');
const ASSETS = path.join(__dirname, 'assets');

// ===== palette (matches Day 1 booklet bars) =====
const BAR1 = 'C0451B';   // 一、选一选 — deep orange
const BAR2 = 'E8A33D';   // 二、连一连 — orange
const BAR3 = '7B2D8E';   // 三、描一描 — purple
const DARK = '1A1A1A', GRAY = '888888';

// ===== page geometry (US Letter, 0.75" margins) =====
const PAGE = { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } };
const CW = 12240 - 1080 - 1080;

const CN = { ascii: 'Times New Roman', eastAsia: 'SimSun' };
const KAI = { ascii: 'Times New Roman', eastAsia: 'KaiTi' };

const noBorder = () => ({ style: BorderStyle.NONE, size: 0, color: 'FFFFFF' });
const allBorders = (b) => ({ top: b, bottom: b, left: b, right: b, insideHorizontal: b, insideVertical: b });

function shadedBar(text, colorHex, size = 22) {
  return new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: [CW], borders: allBorders(noBorder()),
    rows: [new TableRow({ children: [new TableCell({
      width: { size: CW, type: WidthType.DXA },
      shading: { fill: colorHex, type: ShadingType.CLEAR },
      margins: { top: 70, bottom: 70, left: 160, right: 160 },
      children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: 'FFFFFF', size, font: CN })] })],
    })] })],
  });
}

const pageBreak = () => new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] });

// ===== page 1 · cover =====
const coverChildren = [
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 60 },
    children: [new ImageRun({
      data: fs.readFileSync(path.join(ASSETS, 'cover_little_life_helper.png')),
      transformation: { width: 618, height: 866 },
    })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 40, after: 0 },
    children: [
      new TextRun({ text: 'Day 2 · 整理和收纳达人  ', bold: true, size: 28, color: DARK, font: CN }),
      new TextRun({ text: 'Organizing & Storage Master', bold: true, italics: true, size: 24, color: BAR1, font: CN }),
    ],
  }),
];

// ===== page 2 · 一、选一选 =====
const QUESTIONS = [
  {
    n: '1', emoji: '🎒',
    cn: '整理书包的第一步是什么？', en: 'What is the first step to organize your backpack?',
    opts: [
      ['A.', '直接把新东西塞进去', 'Just stuff the new things in.'],
      ['B.', '先把书包里的东西全部拿出来', 'Take everything out of the backpack first.'],
      ['C.', '马上拉上拉链', 'Zip it up right away.'],
    ],
  },
  {
    n: '2', emoji: '🥤',
    cn: '水杯应该放在书包的哪里？', en: 'Where should the water bottle go?',
    opts: [
      ['A.', '侧边口袋，盖子拧紧', 'In the side pocket, with the lid tightened.'],
      ['B.', '和作业本放在一起', 'Together with the homework notebooks.'],
      ['C.', '书包底部，压在书下面', 'At the bottom, under the books.'],
    ],
  },
  {
    n: '3', emoji: '📚',
    cn: '大课本和文件夹应该放在哪里？', en: 'Where do big textbooks and folders belong?',
    opts: [
      ['A.', '前面的小口袋', 'In the small front pocket.'],
      ['B.', '大夹层，靠近背部竖着放', 'In the big compartment, upright and close to your back.'],
      ['C.', '铅笔盒里', 'Inside the pencil case.'],
    ],
  },
  {
    n: '4', emoji: '🧊',
    cn: '厚外套和被子太占地方，用哪种收纳工具最好？', en: 'Which tool works best for bulky coats and quilts?',
    opts: [
      ['A.', '真空压缩袋', 'Vacuum storage bag.'],
      ['B.', '旋转收纳盘', 'Lazy Susan turntable.'],
      ['C.', '洞洞板', 'Pegboard.'],
    ],
  },
  {
    n: '5', emoji: '🍶',
    cn: '柜子后面的小瓶子很难拿到，用哪种收纳工具？', en: 'Which tool helps you reach small bottles at the back of a cabinet?',
    opts: [
      ['A.', '洞洞板', 'Pegboard.'],
      ['B.', '旋转收纳盘', 'Lazy Susan turntable.'],
      ['C.', '真空压缩袋', 'Vacuum storage bag.'],
    ],
  },
  {
    n: '6', emoji: '👕',
    cn: '哪种方法叠衣服最省空间？', en: 'Which way of folding clothes saves the most space?',
    opts: [
      ['A.', '揉成一团塞进去', 'Squeeze it into a ball and stuff it in.'],
      ['B.', '卷卷叠衣法', 'The roll-up folding method.'],
      ['C.', '摊开平放在上面', 'Lay it flat and open on top.'],
    ],
  },
  {
    n: '7', emoji: '🧳',
    cn: '整理行李箱的正确顺序是哪一个？', en: 'What is the correct order for packing a suitcase?',
    opts: [
      ['A.', '把东西都拿出来分类 → 叠衣服 → 用小袋子收纳 → 分区摆放', 'Take all out and sort → fold → use pouches → pack by zones.'],
      ['B.', '先分区摆放 → 再拿出来 → 叠衣服 → 收纳', 'Pack by zones → take out → fold → store.'],
      ['C.', '随便塞进去 → 关上箱子 → 再打开整理', 'Stuff it in → close it → open and redo it.'],
    ],
  },
];

function questionBlock(q) {
  const out = [new Paragraph({
    spacing: { before: 80, after: 40 },
    children: [
      new TextRun({ text: `${q.n} ${q.emoji} `, bold: true, size: 22, font: CN }),
      new TextRun({ text: q.cn, bold: true, size: 22, font: CN }),
      new TextRun({ text: `  ${q.en}`, bold: true, size: 19, font: CN }),
    ],
  })];
  q.opts.forEach(([lab, cn, en]) => {
    out.push(new Paragraph({
      spacing: { before: 6, after: 6 },
      children: [
        new TextRun({ text: `☐ ${lab} `, size: 19, font: CN }),
        new TextRun({ text: `${cn}  `, size: 19, font: CN }),
        new TextRun({ text: en, bold: true, size: 18, font: CN }),
      ],
    }));
  });
  // thin rule under each question
  out.push(new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: [CW],
    borders: { ...allBorders(noBorder()), bottom: { style: BorderStyle.SINGLE, size: 4, color: 'BBBBBB' } },
    rows: [new TableRow({ children: [new TableCell({
      width: { size: CW, type: WidthType.DXA },
      borders: { top: noBorder(), left: noBorder(), right: noBorder(), bottom: { style: BorderStyle.SINGLE, size: 4, color: 'BBBBBB' } },
      children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '', size: 8 })] })],
    })] })],
  }));
  return out;
}

const section1Children = [
  pageBreak(),
  shadedBar('一、选一选。', BAR1, 22),
  ...QUESTIONS.flatMap(questionBlock),
];

// ===== page 3 · 二、连一连 =====
const MATCH_LEFT = ['乱', '分类', '收拾', '书包', '放回'];
const MATCH_RIGHT = [
  ['🎒', 'backpack'],
  ['🗂', 'sort'],
  ['🌀', 'messy'],
  ['📥', 'put back'],
  ['🧹', 'tidy up'],
];

const matchRows = MATCH_LEFT.map((word, i) => new TableRow({
  height: { value: 1950, rule: 'atLeast' },
  children: [
    new TableCell({
      width: { size: Math.floor(CW / 2), type: WidthType.DXA }, borders: allBorders(noBorder()),
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: word, size: 72, font: KAI, color: DARK })] })],
    }),
    new TableCell({
      width: { size: Math.floor(CW / 2), type: WidthType.DXA }, borders: allBorders(noBorder()),
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: MATCH_RIGHT[i][0], size: 72 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: MATCH_RIGHT[i][1], bold: true, size: 22, font: CN })] }),
      ],
    }),
  ],
}));

const section2Children = [
  pageBreak(),
  shadedBar('二、连一连', BAR2, 22),
  new Paragraph({ spacing: { before: 160, after: 160 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: '把词语和图片连起来。 Draw a line from each word to its picture.', italics: true, size: 21, color: GRAY, font: CN })] }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [Math.floor(CW / 2), Math.floor(CW / 2)],
    borders: allBorders(noBorder()),
    rows: matchRows,
  }),
];

// ===== page 4 · 三、描一描, 写一写 =====
const section3Children = [
  pageBreak(),
  shadedBar('三、描一描, 写一写 / Trace and Write  (分类 · 书包 · 收拾)', BAR3, 22),
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 80, after: 0 },
    children: [new ImageRun({
      data: fs.readFileSync(path.join(ASSETS, 'day2_trace_page.png')),
      transformation: { width: 626, height: 850 },
    })],
  }),
];

// ===== build =====
const doc = new Document({
  styles: { default: { document: { run: { font: CN, size: 22 } } } },
  sections: [{ properties: { page: PAGE }, children: [
    ...coverChildren, ...section1Children, ...section2Children, ...section3Children,
  ] }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log(`Created ${OUT}`);
});
