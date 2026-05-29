// Build day3_booklet.docx — 仰望星空 Unit · Day 3: 探秘 航天 / Space Exploration
// Based on day3_space_exploration航天.pptx (manually revised final deck)
// Modeled on build_day1_booklet.js (same 4-section unit structure)
// Run: node build_day3_booklet.js

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
} = require('docx');

const OUT = path.join(__dirname, 'day3_booklet.docx');

// ===== Palette (Day 3 — Space Exploration) =====
const NIGHT  = '0D1B3E';   // deep night sky
const COSMIC = '6A1B9A';
const STAR   = 'F5C242';
const GOLD   = 'FFB700';
const EARTH  = '1E88E5';
const MARS   = 'D84315';   // mars red — primary (§1 bar)
const NEBULA = '7B1FA2';
const MOON_C = 'B0BEC5';
const SKY    = '42A5F5';
const CORAL  = 'FF8F00';   // §3 bar
const PURPLE = '6A1B9A';   // §4 bar
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

// ===== Cover =====
const coverChildren = [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 1200, after: 200 },
    children: [new TextRun({ text: '🚀', size: 120 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text: '仰望星空', bold: true, size: 60, color: NIGHT })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 600 },
    children: [new TextRun({ text: 'Looking Up at the Stars', bold: true, size: 36, color: NIGHT })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text: 'Day 3 · 探秘 航天', bold: true, size: 44, color: DARK })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 200 },
    children: [new TextRun({ text: 'Space Exploration', italics: true, size: 28, color: GRAY })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 800 },
    children: [new TextRun({ text: '🚀 👨‍🚀 🌕 🔴 🛰️', size: 30, color: MARS })],
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

// ===== §1 题库 · 8 个 航天 小问题 / Question Bank =====
const qbCardColors = [MARS, MOON_C, NEBULA, SKY, GOLD, STAR, COSMIC, EARTH];

const qbQuestions = [
  { em: '🚀',  cn: '谁 坐 火箭 去 太空?',                en: 'Who flies a rocket to space?',           a: '宇航员 Astronaut',          b: '厨师 Chef' },
  { em: '🌕',  cn: '月球 上 有 没有 空气?',              en: 'Is there air on the moon?',              a: '没有 No air',               b: '有 Yes, air' },
  { em: '🔴',  cn: '火星 是 什么 颜色?',                 en: 'What color is Mars?',                    a: '红色 Red',                  b: '绿色 Green' },
  { em: '🏠',  cn: '宇航员 在 太空 的 「家」 叫?',        en: "What's the astronauts' home in space?",  a: '太空站 Space Station',      b: '学校 School' },
  { em: '🍱',  cn: '宇航员 在 太空 怎么 吃 饭?',          en: 'How do astronauts eat in space?',        a: '压缩袋装 食物 Packet food', b: '用 碗 和 筷子 Bowl + chopsticks' },
  { em: '💪',  cn: '宇航员 上 太空 前 要 做 什么?',       en: 'What do astronauts do before launch?',   a: '严格 训练 Train hard',      b: '看 电视 Watch TV' },
  { em: '🧪',  cn: '宇航员 在 太空站 做 什么 工作?',      en: "What's the astronauts' job in space?",   a: '做 实验 Experiments',       b: '玩 游戏 Play games' },
  { em: '🎒',  cn: '上 月球 一定 要 带 什么?',           en: 'What MUST you bring to the moon?',       a: '氧气罐 Oxygen tank',        b: '雨伞 Umbrella' },
];

function qbCell(num, q, color) {
  return new TableCell({
    width: { size: Math.floor(CW / 2), type: WidthType.DXA },
    borders: allBorders(border(color, 10)),
    shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
    margins: { top: 140, bottom: 140, left: 200, right: 200 },
    verticalAlign: 'center',
    children: [
      new Paragraph({
        spacing: { before: 0, after: 40 },
        children: [
          new TextRun({ text: `${num}  `, bold: true, size: 28, color }),
          new TextRun({ text: `${q.em}  `, size: 26 }),
          new TextRun({ text: q.cn, bold: true, size: 19, color: DARK }),
        ],
      }),
      new Paragraph({
        spacing: { before: 0, after: 100 },
        indent: { left: 600 },
        children: [new TextRun({ text: q.en, italics: true, size: 14, color: GRAY })],
      }),
      new Paragraph({
        spacing: { before: 0, after: 60 },
        indent: { left: 600 },
        children: [
          new TextRun({ text: '☐  A.  ', bold: true, size: 18, color: DARK }),
          new TextRun({ text: q.a, size: 18, color: DARK }),
        ],
      }),
      new Paragraph({
        spacing: { before: 0, after: 0 },
        indent: { left: 600 },
        children: [
          new TextRun({ text: '☐  B.  ', bold: true, size: 18, color: DARK }),
          new TextRun({ text: q.b, size: 18, color: DARK }),
        ],
      }),
    ],
  });
}

const qbRows = [];
for (let row = 0; row < 4; row++) {
  qbRows.push(new TableRow({
    height: { value: 1600, rule: 'atLeast' },
    children: [
      qbCell(row * 2 + 1, qbQuestions[row * 2],     qbCardColors[row * 2]),
      qbCell(row * 2 + 2, qbQuestions[row * 2 + 1], qbCardColors[row * 2 + 1]),
    ],
  }));
}

const section1Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('一、题库 · 8 个 航天 小 问题 / Question Bank · 8 Space Questions  (圈出正确答案)', MARS, 24),
  new Paragraph({
    spacing: { before: 160, after: 160 },
    children: [new TextRun({
      text: '👉 看 问题, 圈出 对的 答案。/ Read the question and circle the right answer.',
      size: 22, italics: true, color: GRAY,
    })],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [Math.floor(CW / 2), Math.floor(CW / 2)],
    borders: allBorders(noBorder()),
    rows: qbRows,
  }),
  new Paragraph({
    spacing: { before: 240, after: 0 },
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({ text: '🏆 ', size: 22 }),
      new TextRun({ text: '答对 8 题 = 「小小宇航员」徽章! ', bold: true, size: 20, color: MARS }),
      new TextRun({ text: 'All 8 right = Little Astronaut badge!', italics: true, size: 16, color: GRAY }),
    ],
  }),
];

// ===== §2 如果 你 是 宇航员 — open creative prompt + draw + write =====
const section2Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('二、如果 我 是 宇航员 / If I Were an Astronaut', EARTH, 22),
  new Paragraph({
    spacing: { before: 160, after: 80 },
    children: [new TextRun({
      text: '🚀 如果 你 是 宇航员, 你 想 去 哪里? 你 会 做 什么?',
      size: 24, bold: true, color: DARK,
    })],
  }),
  new Paragraph({
    spacing: { before: 40, after: 200 },
    children: [new TextRun({
      text: 'If you were an astronaut — where would you go? What would you do?',
      size: 18, italics: true, color: GRAY,
    })],
  }),
  new Paragraph({
    spacing: { before: 100, after: 40 },
    children: [
      new TextRun({ text: '🎨 画一画  Draw it', size: 22, bold: true, color: EARTH }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 80 },
    children: [
      new TextRun({ text: '画 穿 上 宇航服 的 你 — 在 月球 / 火星 / 太空站', size: 14, color: GRAY }),
      new TextRun({ text: '  ·  ', size: 14, color: LGRAY }),
      new TextRun({ text: 'Draw yourself in a spacesuit — on the moon / Mars / space station', size: 14, italics: true, color: GRAY }),
    ],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders(border(EARTH, 12)),
    rows: [new TableRow({
      height: { value: 4000, rule: 'atLeast' },
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
  new Paragraph({
    spacing: { before: 240, after: 40 },
    children: [
      new TextRun({ text: '✏️ 写一写  Write it', size: 22, bold: true, color: EARTH }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 80 },
    children: [
      new TextRun({ text: '完成 下面 的 句子', size: 14, color: GRAY }),
      new TextRun({ text: '  ·  ', size: 14, color: LGRAY }),
      new TextRun({ text: 'Complete the sentences below', size: 14, italics: true, color: GRAY }),
    ],
  }),
  // Sentence frame 1: 我 想 去 ___
  new Paragraph({
    spacing: { before: 120, after: 60 },
    children: [
      new TextRun({ text: '我 想 去 ', size: 24, bold: true, color: DARK }),
      new TextRun({ text: '_____________________________', size: 24, color: GRAY }),
      new TextRun({ text: ' 。', size: 24, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 160 },
    children: [
      new TextRun({ text: 'I want to go to ', size: 16, italics: true, color: GRAY }),
      new TextRun({ text: '_______________________ .', size: 16, color: GRAY }),
    ],
  }),
  // Sentence frame 2: 我 会 看到 ___
  new Paragraph({
    spacing: { before: 120, after: 60 },
    children: [
      new TextRun({ text: '我 会 看到 ', size: 24, bold: true, color: DARK }),
      new TextRun({ text: '____________________________', size: 24, color: GRAY }),
      new TextRun({ text: ' 。', size: 24, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 160 },
    children: [
      new TextRun({ text: 'I will see ', size: 16, italics: true, color: GRAY }),
      new TextRun({ text: '_________________________ .', size: 16, color: GRAY }),
    ],
  }),
  // Sentence frame 3: 我 想 做 ___
  new Paragraph({
    spacing: { before: 120, after: 60 },
    children: [
      new TextRun({ text: '我 想 做 ', size: 24, bold: true, color: DARK }),
      new TextRun({ text: '_____________________________', size: 24, color: GRAY }),
      new TextRun({ text: ' 。', size: 24, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 0 },
    children: [
      new TextRun({ text: 'I want to ', size: 16, italics: true, color: GRAY }),
      new TextRun({ text: '_______________________ .', size: 16, color: GRAY }),
    ],
  }),
];

// ===== §3 连一连 / Match — 我会认: 宇航员 火箭 月球 火星 太空站 =====
const matchWords = [
  { char: '宇航员', py: 'yǔ háng yuán', en: 'astronaut',     em: '👨‍🚀' },
  { char: '火箭',   py: 'huǒ jiàn',     en: 'rocket',        em: '🚀' },
  { char: '月球',   py: 'yuè qiú',      en: 'moon',          em: '🌕' },
  { char: '火星',   py: 'huǒ xīng',     en: 'Mars',          em: '🔴' },
  { char: '太空 站', py: 'tài kōng zhàn', en: 'space station', em: '🛰️' },
];

// Shuffle pairs so left-right indices don't match
const matchShuffled = [matchWords[3], matchWords[0], matchWords[4], matchWords[1], matchWords[2]];
const matchRows = matchWords.map((w, i) => {
  const right = matchShuffled[i];
  const colW = Math.floor(CW / 2);
  return new TableRow({
    height: { value: 1100, rule: 'atLeast' },
    children: [
      new TableCell({
        width: { size: colW, type: WidthType.DXA },
        borders: allBorders(border(MARS, 8)),
        margins: { top: 200, bottom: 200, left: 240, right: 240 },
        verticalAlign: 'center',
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: w.char, bold: true, size: 44, color: DARK })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: w.py, size: 22, color: GRAY, italics: true })],
          }),
        ],
      }),
      new TableCell({
        width: { size: colW, type: WidthType.DXA },
        borders: allBorders(border(STAR, 8)),
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
      text: '👉 把中文词语和正确的表情 + 英文连起来。',
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

// ===== §4 描一描, 写一写 — 火箭 · 月球 · 火星 =====
const section4Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('四、描一描, 写一写 / Trace and Write  (火箭 · 月球 · 火星)', PURPLE, 24),
  new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({
      text: '👉 在下面贴上田字格写字纸, 写一写今天学到的字: 火箭 · 月球 · 火星。',
      size: 22, italics: true, color: GRAY,
    })],
  }),
  new Paragraph({
    spacing: { before: 60, after: 200 },
    children: [new TextRun({
      text: 'Insert your grid-paper below and practice today’s characters: 火箭 · 月球 · 火星.',
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
            text: '📄  在这里贴上田字格写字纸 / Insert your grid paper here',
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
