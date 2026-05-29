// Build day2_booklet.docx — 仰望星空 Unit · Day 2: 星空与星座 / Stars & Constellations
// Based on day2_constellations星座.pptx (manually revised final deck)
// Modeled on build_day1_booklet.js (same 4-section unit structure)
// Run: node build_day2_booklet.js

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
} = require('docx');

const OUT = path.join(__dirname, 'day2_booklet.docx');

// ===== Palette (Day 2 — Mythic Night) =====
const NIGHT  = '0D1B3E';   // deep night sky
const COSMIC = '6A1B9A';   // nebula purple — primary (§1 bar)
const STAR   = 'F5C242';   // golden star
const GOLD   = 'FFB700';
const EARTH  = '1E88E5';   // earth blue (§2 bar)
const PINK   = 'EC407A';   // myth pink
const NEBULA = '7B1FA2';
const CORAL  = 'FF8F00';   // §3 bar
const PURPLE = '6A1B9A';   // §4 bar
const SILVER = 'B0BEC5';
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
    children: [new TextRun({ text: '⭐', size: 120 })],
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
    children: [new TextRun({ text: 'Day 2 · 星空 与 星座', bold: true, size: 44, color: DARK })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 200 },
    children: [new TextRun({ text: 'Stars & Constellations', italics: true, size: 28, color: GRAY })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 800 },
    children: [new TextRun({ text: '⭐ ✨ 🌌 🐻 🦁', size: 30, color: STAR })],
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

// ===== §1 题库 · 8 个 星空 小问题 / Question Bank =====
const qbCardColors = [COSMIC, STAR, NEBULA, EARTH, GOLD, PINK, NIGHT, SILVER];

const qbQuestions = [
  { em: '⭐',  cn: '什么 是 「星座」?',              en: 'What is a constellation?',                a: '古人 把 星星 连 起来 的 图案 · A pattern from connecting stars',  b: '天上 真的 有 动物 · Real animals up in the sky' },
  { em: '🥄',  cn: '北斗七星 像 什么?',              en: 'What does the Big Dipper look like?',     a: '大 勺子',                       b: '一朵 花' },
  { em: '🦁',  cn: '哪 一个 是 狮子 星座?',          en: 'Which one is the lion constellation?',    a: '狮子座 Leo',                    b: '大熊座 Big Bear' },
  { em: '🗺️', cn: '古人 没 GPS — 用 什么 找 方向?', en: 'No GPS — how did ancient people navigate?', a: '看 星星',                      b: '问 树' },
  { em: '🧭',  cn: '哪 颗 星 帮 古人 找 北方?',      en: 'Which star helped find North?',           a: '北极星 North Star',             b: '月亮 Moon' },
  { em: '🌌',  cn: '银河 看 起来 像 什么?',           en: 'What does the Milky Way look like?',      a: '一条 白色 大 河 White river',    b: '一座 红色 桥 Red bridge' },
  { em: '🦢',  cn: '哪 一个 星座 像 一只 飞 的 天鹅?', en: 'Which constellation looks like a swan?', a: '天鹅座 Cygnus',                 b: '狮子座 Leo' },
  { em: '🎨',  cn: '米罗 爷爷 画 星空 用 什么 元素?', en: 'What elements did Miró use?',             a: '点 · 线 · 面',                  b: '圆 · 三角 · 正方形' },
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
  shadedBar('一、题库 · 8 个 星空 小 问题 / Question Bank · 8 Star Questions  (圈出正确答案)', COSMIC, 24),
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
      new TextRun({ text: '答对 8 题 = 「小小星空 探险家」徽章! ', bold: true, size: 20, color: COSMIC }),
      new TextRun({ text: 'All 8 right = Little Star Explorer badge!', italics: true, size: 16, color: GRAY }),
    ],
  }),
];

// ===== §2 创造 你 的 星座 — open creative prompt + draw + write =====
const section2Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('二、创造 你 的 星座 / Create Your Own Constellation', EARTH, 22),
  new Paragraph({
    spacing: { before: 160, after: 80 },
    children: [new TextRun({
      text: '✨ 如果 你 可以 创造 一个 属于 你 自己 的 星座, 你 会 创造 什么?',
      size: 24, bold: true, color: DARK,
    })],
  }),
  new Paragraph({
    spacing: { before: 40, after: 200 },
    children: [new TextRun({
      text: 'If you could create your very own constellation — what would it look like?',
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
      new TextRun({ text: '把 几 颗 星 连 起来 — 看 像 什么?', size: 14, color: GRAY }),
      new TextRun({ text: '  ·  ', size: 14, color: LGRAY }),
      new TextRun({ text: 'Connect a few stars — what shape do you see?', size: 14, italics: true, color: GRAY }),
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
  // Sentence frame 1: 我的星座叫 ___
  new Paragraph({
    spacing: { before: 120, after: 60 },
    children: [
      new TextRun({ text: '我 的 星座 叫 ', size: 24, bold: true, color: DARK }),
      new TextRun({ text: '_____________________________', size: 24, color: GRAY }),
      new TextRun({ text: ' 。', size: 24, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 160 },
    children: [
      new TextRun({ text: 'My constellation is called ', size: 16, italics: true, color: GRAY }),
      new TextRun({ text: '_______________________ .', size: 16, color: GRAY }),
    ],
  }),
  // Sentence frame 2: 它看起来像 ___
  new Paragraph({
    spacing: { before: 120, after: 60 },
    children: [
      new TextRun({ text: '它 看 起来 像 ', size: 24, bold: true, color: DARK }),
      new TextRun({ text: '____________________________', size: 24, color: GRAY }),
      new TextRun({ text: ' 。', size: 24, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 160 },
    children: [
      new TextRun({ text: 'It looks like ', size: 16, italics: true, color: GRAY }),
      new TextRun({ text: '_________________________ .', size: 16, color: GRAY }),
    ],
  }),
  // Sentence frame 3: 它的故事是 ___
  new Paragraph({
    spacing: { before: 120, after: 60 },
    children: [
      new TextRun({ text: '它 的 故事 是 ', size: 24, bold: true, color: DARK }),
      new TextRun({ text: '__________________________', size: 24, color: GRAY }),
      new TextRun({ text: ' 。', size: 24, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 0 },
    children: [
      new TextRun({ text: 'Its story is ', size: 16, italics: true, color: GRAY }),
      new TextRun({ text: '____________________________ .', size: 16, color: GRAY }),
    ],
  }),
];

// ===== §3 连一连 / Match — 我会认: 星星 星座 银河 神话 故事 =====
const matchWords = [
  { char: '星星', py: 'xīng xing', en: 'stars',        em: '⭐' },
  { char: '星座', py: 'xīng zuò',  en: 'constellation', em: '✨' },
  { char: '银河', py: 'yín hé',    en: 'Milky Way',     em: '🌌' },
  { char: '太阳', py: 'tài yáng',  en: 'Sun',           em: '☀️' },
  { char: '月亮', py: 'yuè liang', en: 'Moon',          em: '🌙' },
];

// Shuffle pairs so left-right indices don't match (cosmic shuffle)
const matchShuffled = [matchWords[3], matchWords[0], matchWords[4], matchWords[1], matchWords[2]];
const matchRows = matchWords.map((w, i) => {
  const right = matchShuffled[i];
  const colW = Math.floor(CW / 2);
  return new TableRow({
    height: { value: 1100, rule: 'atLeast' },
    children: [
      new TableCell({
        width: { size: colW, type: WidthType.DXA },
        borders: allBorders(border(COSMIC, 8)),
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

// ===== §4 描一描, 写一写 — 星星 · 星座 =====
const section4Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('四、描一描, 写一写 / Trace and Write  (星星 · 星座)', PURPLE, 24),
  new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({
      text: '👉 在下面贴上田字格写字纸, 写一写今天学到的字: 星星 · 星座。',
      size: 22, italics: true, color: GRAY,
    })],
  }),
  new Paragraph({
    spacing: { before: 60, after: 200 },
    children: [new TextRun({
      text: 'Insert your grid-paper below and practice today’s characters: 星星 · 星座.',
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
