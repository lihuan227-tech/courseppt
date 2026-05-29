// Build day4_booklet.docx — 仰望星空 Unit · Day 4: 外星人 是否 存在? / Do Aliens Exist?
// Based on day4_aliens- final.pdf (manually revised final deck)
// Modeled on build_day3_booklet.js (same 4-section unit structure)
// Run: node build_day4_booklet.js

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
} = require('docx');

const OUT = path.join(__dirname, 'day4_booklet.docx');

// ===== Palette (Day 4 — Aliens · Do Aliens Exist?) =====
const NIGHT  = '0D1B3E';
const COSMIC = '6A1B9A';
const STAR   = 'F5C242';
const GOLD   = 'FFB700';
const EARTH  = '1E88E5';
const MARS   = 'D84315';   // §3 bar (Zorp from Mars)
const NEBULA = '7B1FA2';   // §4 bar
const ALIEN  = '66BB6A';   // §1 bar — primary alien green
const SKY    = '42A5F5';
const PINK   = 'EC407A';
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
    children: [new TextRun({ text: '👽', size: 120 })],
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
    children: [new TextRun({ text: 'Day 4 · 外星人 是否 存在?', bold: true, size: 44, color: DARK })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 200 },
    children: [new TextRun({ text: 'Do Aliens Exist?', italics: true, size: 28, color: GRAY })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 800 },
    children: [new TextRun({ text: '👽 🛸 🪐 🌌 🔴', size: 30, color: ALIEN })],
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

// ===== §1 题库 · 8 个 外星人 小问题 / Question Bank =====
const qbCardColors = [ALIEN, MARS, GOLD, SKY, NEBULA, COSMIC, STAR, EARTH];

const qbQuestions = [
  { em: '✉️', cn: '谁 写 信 给 地球 小朋友?',          en: 'Who wrote the letter to Earth kids?',      a: '火星人 Zorp Zorp from Mars',  b: '太阳 The Sun' },
  { em: '🔴', cn: '火星 上 有 可以 呼吸 的 空气 吗?',   en: 'Is there breathable air on Mars?',         a: '没有 No, no air',              b: '有 Yes, plenty' },
  { em: '🪐', cn: '太阳系 里 最大 的 行星 是?',         en: 'Largest planet in the solar system?',      a: '木星 Jupiter',                 b: '月球 The Moon' },
  { em: '⭕', cn: '土星 外围 有 什么 特别?',            en: "What's special around Saturn?",            a: '光环 Rings',                   b: '海洋 An ocean' },
  { em: '🔵', cn: '天王星 看 起来 是 什么 颜色?',       en: 'What color does Uranus look?',             a: '蓝色 Blue',                    b: '红色 Red' },
  { em: '💭', cn: '「外星人 长 什么 样」是 我们 的?',   en: '"What aliens look like" is our…',          a: '猜想 A guess',                 b: '事实 A fact' },
  { em: '🌍', cn: 'Zorp 想 让 地球 小朋友 做 什么?',    en: 'What does Zorp want Earth kids to do?',    a: '保护 地球 Protect Earth',      b: '离开 地球 Leave Earth' },
  { em: '✨', cn: '因为 星球 不一样, 所以 外星人?',     en: 'Different planets → aliens are…',          a: '也 不一样 Also different',     b: '都 一样 All the same' },
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
  shadedBar('一、题库 · 8 个 外星人 小 问题 / Question Bank · 8 Alien Questions  (圈出正确答案)', ALIEN, 24),
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
      new TextRun({ text: '答对 8 题 = 「外星 探险家」徽章! ', bold: true, size: 20, color: ALIEN }),
      new TextRun({ text: 'All 8 right = Alien Explorer badge!', italics: true, size: 16, color: GRAY }),
    ],
  }),
];

// ===== §2 设计 你 的 外星朋友 — draw + write =====
const section2Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('二、设计 你 的 外星朋友 / Design Your Alien Friend', COSMIC, 22),
  new Paragraph({
    spacing: { before: 160, after: 80 },
    children: [new TextRun({
      text: '👽 想 一 想 — 如果 你 有 一个 外星朋友, 它 长 什么 样? 住 在 哪里?',
      size: 24, bold: true, color: DARK,
    })],
  }),
  new Paragraph({
    spacing: { before: 40, after: 200 },
    children: [new TextRun({
      text: 'Imagine your own alien friend — what do they look like? Where do they live?',
      size: 18, italics: true, color: GRAY,
    })],
  }),
  new Paragraph({
    spacing: { before: 100, after: 40 },
    children: [
      new TextRun({ text: '🎨 画一画  Draw it', size: 22, bold: true, color: COSMIC }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 80 },
    children: [
      new TextRun({ text: '画 你 的 外星朋友 — 几 只 眼睛? 几 条 腿? 什么 颜色?', size: 14, color: GRAY }),
      new TextRun({ text: '  ·  ', size: 14, color: LGRAY }),
      new TextRun({ text: 'Draw your alien — eyes, legs, colors, special features', size: 14, italics: true, color: GRAY }),
    ],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders(border(COSMIC, 12)),
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
      new TextRun({ text: '✏️ 写一写  Write it', size: 22, bold: true, color: COSMIC }),
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
  // Sentence frame 1: 我 的 外星朋友 叫 ___
  new Paragraph({
    spacing: { before: 120, after: 60 },
    children: [
      new TextRun({ text: '我 的 外星朋友 叫 ', size: 24, bold: true, color: DARK }),
      new TextRun({ text: '_____________________________', size: 24, color: GRAY }),
      new TextRun({ text: ' 。', size: 24, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 160 },
    children: [
      new TextRun({ text: 'My alien friend is called ', size: 16, italics: true, color: GRAY }),
      new TextRun({ text: '_______________________ .', size: 16, color: GRAY }),
    ],
  }),
  // Sentence frame 2: 它 住 在 ___ 星球
  new Paragraph({
    spacing: { before: 120, after: 60 },
    children: [
      new TextRun({ text: '它 住 在 ', size: 24, bold: true, color: DARK }),
      new TextRun({ text: '_________________', size: 24, color: GRAY }),
      new TextRun({ text: ' 星球 上 。', size: 24, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 160 },
    children: [
      new TextRun({ text: 'It lives on planet ', size: 16, italics: true, color: GRAY }),
      new TextRun({ text: '_________________ .', size: 16, color: GRAY }),
    ],
  }),
  // Sentence frame 3: 它 有 ___ 只 眼睛
  new Paragraph({
    spacing: { before: 120, after: 60 },
    children: [
      new TextRun({ text: '它 有 ', size: 24, bold: true, color: DARK }),
      new TextRun({ text: '_______', size: 24, color: GRAY }),
      new TextRun({ text: ' 只 眼睛 和 ', size: 24, bold: true, color: DARK }),
      new TextRun({ text: '_______', size: 24, color: GRAY }),
      new TextRun({ text: ' 条 腿 。', size: 24, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 160 },
    children: [
      new TextRun({ text: 'It has ', size: 16, italics: true, color: GRAY }),
      new TextRun({ text: '____', size: 16, color: GRAY }),
      new TextRun({ text: ' eyes and ', size: 16, italics: true, color: GRAY }),
      new TextRun({ text: '____', size: 16, color: GRAY }),
      new TextRun({ text: ' legs.', size: 16, italics: true, color: GRAY }),
    ],
  }),
  // Sentence frame 4: 它 会 ___ (special ability)
  new Paragraph({
    spacing: { before: 120, after: 60 },
    children: [
      new TextRun({ text: '它 会 ', size: 24, bold: true, color: DARK }),
      new TextRun({ text: '_____________________________', size: 24, color: GRAY }),
      new TextRun({ text: ' 。', size: 24, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 0 },
    children: [
      new TextRun({ text: 'It can ', size: 16, italics: true, color: GRAY }),
      new TextRun({ text: '___________________ (fly / jump / glow / …) .', size: 16, color: GRAY }),
    ],
  }),
];

// ===== §3 连一连 / Match — 我会认: 外星人 生命 信号 发现 猜想 =====
const matchWords = [
  { char: '外星人', py: 'wài xīng rén', en: 'alien',     em: '👽' },
  { char: '生命',   py: 'shēng mìng',   en: 'life',      em: '🌱' },
  { char: '信号',   py: 'xìn hào',      en: 'signal',    em: '📡' },
  { char: '发现',   py: 'fā xiàn',      en: 'discover',  em: '🔍' },
  { char: '猜想',   py: 'cāi xiǎng',    en: 'guess',     em: '💭' },
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
  shadedBar('三、连一连 / Match  (用线连起来)', MARS, 24),
  new Paragraph({
    spacing: { before: 200, after: 200 },
    children: [new TextRun({
      text: '👉 把 中文 词语 和 正确 的 表情 + 英文 连 起来。',
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

// ===== §4 描一描, 写一写 — 外星人 · 生命 =====
const section4Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('四、描一描, 写一写 / Trace and Write  (外星人 · 生命)', NEBULA, 24),
  new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({
      text: '👉 在下面贴上田字格写字纸, 写一写今天学到的字: 外星人 · 生命。',
      size: 22, italics: true, color: GRAY,
    })],
  }),
  new Paragraph({
    spacing: { before: 60, after: 200 },
    children: [new TextRun({
      text: 'Insert your grid-paper below and practice today’s characters: 外星人 · 生命.',
      size: 20, italics: true, color: GRAY,
    })],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders(border(NEBULA, 12)),
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
