// Build day4_booklet.docx — 小小生活家 Little Life Helper · Day 4: 友善我先行 · 接人待物
// Same 4-page structure as the Day 1/Day 2 booklets (中低 level):
//   p1 Cover (unit coloring cover + 姓名/班级)
//   p2 一、选一选  — bilingual multiple choice from the Day 4 deck
//   p3 二、连一连  — 我会认: 友善 · 朋友 · 分享 · 帮助 · 尊重
//   p4 三、描一描, 写一写 — 我会写 朋友 (+ 描一描 友善 · 帮助)
// Content source: 0 final slides/day4_Good manners & kindness.pdf
// Run: node make_trace_page.js 4 && node build_day4_booklet.js
const fs = require('fs');
const path = require('path');
const REPO = path.resolve(__dirname, '../../..');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
} = require(path.join(REPO, 'node_modules/docx'));

const OUT = path.join(__dirname, 'day4_booklet.docx');
const ASSETS = path.join(__dirname, 'assets');

// ===== palette (same bars as Day 1/2 booklets) =====
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
      new TextRun({ text: 'Day 4 · 友善我先行 · 接人待物  ', bold: true, size: 28, color: DARK, font: CN }),
      new TextRun({ text: 'Kindness & Good Manners', bold: true, italics: true, size: 24, color: BAR1, font: CN }),
    ],
  }),
];

// ===== page 2 · 一、选一选 =====
const QUESTIONS = [
  {
    n: '1', emoji: '💛',
    cn: '友善是什么意思？', en: 'What does being kind mean?',
    opts: [
      ['A.', '只跟自己喜欢的人玩', 'Only play with the friends you like.'],
      ['B.', '好好说话、帮助别人、避免矛盾', 'Speak kindly, help others, avoid fights.'],
      ['C.', '谁最厉害就听谁的', 'Follow whoever is the strongest.'],
    ],
  },
  {
    n: '2', emoji: '✏️',
    cn: '同学帮你捡起铅笔，你应该说什么？', en: 'A classmate picks up your pencil. What do you say?',
    opts: [
      ['A.', '「谢谢你！」', '"Thank you!"'],
      ['B.', '「你早就该帮我了。」', '"You should have helped me sooner."'],
      ['C.', '什么也不说', 'Say nothing.'],
    ],
  },
  {
    n: '3', emoji: '🦶',
    cn: '你不小心踩到别人的脚，应该说什么？', en: 'You accidentally step on someone\'s foot. What do you say?',
    opts: [
      ['A.', '「你干吗站在这里？」', '"Why are you standing here?"'],
      ['B.', '「对不起，你没事吧？」', '"Sorry! Are you OK?"'],
      ['C.', '「我又没看见。」', '"Well, I didn\'t see you."'],
    ],
  },
  {
    n: '4', emoji: '🙇',
    cn: '别人对你说「对不起」，你可以说什么？', en: 'Someone says sorry to you. What can you say?',
    opts: [
      ['A.', '「没关系。」', '"It\'s OK."'],
      ['B.', '「就是你的错！」', '"It IS your fault!"'],
      ['C.', '不理他', 'Ignore them.'],
    ],
  },
  {
    n: '5', emoji: '🖊️',
    cn: '你想借同学的笔，怎么说最友善？', en: 'You want to borrow a classmate\'s pen. What is the kindest way?',
    opts: [
      ['A.', '「把笔给我！」', '"Give me the pen!"'],
      ['B.', '「请问，我可以借一下你的笔吗？」', '"Excuse me, may I borrow your pen?"'],
      ['C.', '不用说，拿了就用', 'Just take it without asking.'],
    ],
  },
  {
    n: '6', emoji: '🚪',
    cn: '家里来客人了，小主人应该怎么做？', en: 'A guest comes to your home. What should you do?',
    opts: [
      ['A.', '继续玩，不说话', 'Keep playing and say nothing.'],
      ['B.', '说「你好，欢迎来我们家！」，请客人坐下', 'Say "Hello, welcome!" and invite them to sit.'],
      ['C.', '躲进自己的房间', 'Hide in your room.'],
    ],
  },
  {
    n: '7', emoji: '💬',
    cn: '「我话语」的四个步骤是哪一个？', en: 'What are the four steps of an "I-message"?',
    opts: [
      ['A.', '客观描述 → 表达感受 → 说明原因 → 提出需求', 'Describe → say how you feel → say why → say what you hope for.'],
      ['B.', '先指责 → 再生气 → 再不理人', 'Blame → get angry → stop talking.'],
      ['C.', '大声说 → 让别人听我的', 'Speak louder → make others obey.'],
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

// ===== page 3 · 二、连一连 (我会认 5 words) =====
const MATCH_LEFT = ['尊重', '朋友', '友善', '帮助', '分享'];
const MATCH_RIGHT = [
  ['🍪', 'share'],
  ['🙏', 'respect'],
  ['🙋', 'help'],
  ['💛', 'kind'],
  ['🤝', 'friend'],
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
  shadedBar('三、描一描, 写一写 / Trace and Write  (朋友 · 友善 · 帮助)', BAR3, 22),
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 80, after: 0 },
    children: [new ImageRun({
      data: fs.readFileSync(path.join(ASSETS, 'day4_trace_page.png')),
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
