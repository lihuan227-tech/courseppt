// Build day1_ai_booklet.docx — 玩转 创新 科技 Unit · Day 1: 认识 AI · 做 聪明 的 AI 小主人
// Based on day1_ai.pptx (final 40-slide deck)
// Modeled on the 仰望星空 unit's build_day*_booklet.js (same 4-section structure)
// Run: node build_day1_ai_booklet.js

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
} = require('docx');

const OUT = path.join(__dirname, 'day1_ai_booklet.docx');

// ===== Palette (Day 1 — AI · Innovative Tech unit) =====
const INK       = '0F1A3A';
const AI_PURPLE = '6A1B9A';   // §1 bar — primary AI purple
const CYBER     = '2A47E0';   // §2 bar
const ORANGE    = 'FB8C00';   // §3 bar
const GREEN     = '2E7D32';   // §4 bar
const PINK      = 'D81B60';
const STAR      = 'F5C242';
const GOLD      = 'FFB700';
const SKY       = '42A5F5';
const DARK      = '2C2C2C';
const GRAY      = '888888';
const LGRAY     = 'D8D8D8';

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
    children: [new TextRun({ text: '🤖', size: 120 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text: '玩转 创新 科技', bold: true, size: 60, color: INK })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 600 },
    children: [new TextRun({ text: 'Playing with Innovative Tech', bold: true, size: 32, color: INK })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text: 'Day 1 · 认识 AI', bold: true, size: 44, color: DARK })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 200 },
    children: [new TextRun({ text: '做 聪明 的 AI 小主人', bold: true, size: 30, color: AI_PURPLE })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 200 },
    children: [new TextRun({ text: 'Becoming a Smart AI Master', italics: true, size: 24, color: GRAY })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 800 },
    children: [new TextRun({ text: '🤖 💻 🧠 ✨ 🔮', size: 30, color: AI_PURPLE })],
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

// ===== §1 题库 · 8 个 AI 小问题 / Question Bank =====
const qbCardColors = [AI_PURPLE, ORANGE, GREEN, PINK, CYBER, SKY, GOLD, AI_PURPLE];

const qbQuestions = [
  { em: '🤖', cn: '「AI」 是 什么 的 简称?',           en: "What's 'AI' short for?",                a: 'Artificial Intelligence (人工 智能)', b: 'Apple Internet' },
  { em: '🎙️', cn: 'Siri 是 AI 吗?',                    en: 'Is Siri AI?',                           a: '是 — 会 听, 会 答',                 b: '不是 — 只是 一个 app' },
  { em: '💡', cn: '普通 电灯 是 AI 吗?',                en: 'Is a regular light bulb AI?',           a: '不 是 — 只 会 「开/关」',           b: '是 — 它 会 亮' },
  { em: '🤔', cn: 'AI 会 犯错 吗?',                     en: 'Can AI make mistakes?',                 a: '会 — 它 也 会 「胡说 八道」',        b: '不会 — AI 什么 都 对' },
  { em: '🚗', cn: '自动 驾驶 车 是 AI 吗?',              en: 'Is a self-driving car AI?',             a: '是 — 车 自己 看 路, 自己 判 断!',   b: '不 是 — 只 是 普通 车' },
  { em: '🛡️', cn: '我 把 电话 告诉 AI — 对 吗?',       en: 'Should I share my phone number with AI?',a: '不 对 — 隐私 要 保护!',             b: '对 — 没关系' },
  { em: '📚', cn: 'AI 帮 我 做 数学题, 我 直接 抄 — 聪明 吗?', en: 'AI did homework, I copy it — smart?',  a: '不 聪明 — 我 没 学到 东西',         b: '聪明 — 省时间' },
  { em: '🧠', cn: 'AI 怎么 学? (像 什么?)',             en: 'How does AI learn?',                    a: '看 很多 例子, 像 「百科 全书」',     b: '出生 就 会 — 它 是 神奇 的' },
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
    height: { value: 1700, rule: 'atLeast' },
    children: [
      qbCell(row * 2 + 1, qbQuestions[row * 2],     qbCardColors[row * 2]),
      qbCell(row * 2 + 2, qbQuestions[row * 2 + 1], qbCardColors[row * 2 + 1]),
    ],
  }));
}

const section1Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('一、题库 · 8 个 AI 小 问题 / Question Bank · 8 AI Questions  (圈出 正确 答案)', AI_PURPLE, 24),
  new Paragraph({
    spacing: { before: 160, after: 160 },
    children: [new TextRun({
      text: '👉 看 问题, 圈出 对 的 答案。/ Read the question and circle the right answer.',
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
      new TextRun({ text: '答对 8 题 = 「聪明 AI 小 主人」 徽章! ', bold: true, size: 20, color: AI_PURPLE }),
      new TextRun({ text: 'All 8 right = Smart AI Master badge!', italics: true, size: 16, color: GRAY }),
    ],
  }),
];

// ===== §2 设计 你 的 AI 小帮手 — open creative + draw + write =====
const section2Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('二、设计 你 的 AI 小帮手 / Design Your AI Helper', CYBER, 22),
  new Paragraph({
    spacing: { before: 160, after: 80 },
    children: [new TextRun({
      text: '🤖 想 一 想 — 如果 你 能 创造 一 个 AI 小 帮手, 它 是 什么 样 的? 帮 谁?',
      size: 24, bold: true, color: DARK,
    })],
  }),
  new Paragraph({
    spacing: { before: 40, after: 200 },
    children: [new TextRun({
      text: 'Imagine your own AI helper — what does it look like? Who does it help?',
      size: 18, italics: true, color: GRAY,
    })],
  }),
  new Paragraph({
    spacing: { before: 100, after: 40 },
    children: [
      new TextRun({ text: '🎨 画一画  Draw it', size: 22, bold: true, color: CYBER }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 80 },
    children: [
      new TextRun({ text: '画 你 的 AI 小帮手 — 它 长 什么 样? 颜色? 形状?', size: 14, color: GRAY }),
      new TextRun({ text: '  ·  ', size: 14, color: LGRAY }),
      new TextRun({ text: 'Draw your AI — its look, color, shape', size: 14, italics: true, color: GRAY }),
    ],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders(border(CYBER, 12)),
    rows: [new TableRow({
      height: { value: 3800, rule: 'atLeast' },
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
      new TextRun({ text: '✏️ 写一写  Write it', size: 22, bold: true, color: CYBER }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 80 },
    children: [
      new TextRun({ text: '看 看 别 的 同学 怎么 写 — 然后 你 来 写!', size: 14, color: GRAY }),
      new TextRun({ text: '  ·  ', size: 14, color: LGRAY }),
      new TextRun({ text: 'See examples — then write your own!', size: 14, italics: true, color: GRAY }),
    ],
  }),
  // ===== Examples (3 mini cards showing completed responses) =====
  new Paragraph({
    spacing: { before: 120, after: 60 },
    children: [new TextRun({ text: '💡 例子  Examples:', size: 18, bold: true, color: ORANGE })],
  }),
  // Example 1
  new Paragraph({
    spacing: { before: 60, after: 0 },
    indent: { left: 300 },
    children: [
      new TextRun({ text: '① ', bold: true, size: 20, color: ORANGE }),
      new TextRun({ text: '我 的 AI 叫 「小亮」。它 帮助 爷爷 奶奶。它 会 提 醒 吃 药。', size: 18, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 100 },
    indent: { left: 600 },
    children: [
      new TextRun({ text: "My AI is 'Xiao Liang'. It helps grandparents. It can remind them to take medicine.",
                    size: 13, italics: true, color: GRAY }),
    ],
  }),
  // Example 2
  new Paragraph({
    spacing: { before: 0, after: 0 },
    indent: { left: 300 },
    children: [
      new TextRun({ text: '② ', bold: true, size: 20, color: ORANGE }),
      new TextRun({ text: '我 的 AI 叫 「学学」。它 帮助 小朋友。它 会 解释 数学题。', size: 18, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 100 },
    indent: { left: 600 },
    children: [
      new TextRun({ text: "My AI is 'Xue Xue'. It helps kids. It can explain math problems.",
                    size: 13, italics: true, color: GRAY }),
    ],
  }),
  // Example 3
  new Paragraph({
    spacing: { before: 0, after: 0 },
    indent: { left: 300 },
    children: [
      new TextRun({ text: '③ ', bold: true, size: 20, color: ORANGE }),
      new TextRun({ text: '我 的 AI 叫 「环 环」。它 帮助 地球。它 会 分类 垃圾。', size: 18, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 200 },
    indent: { left: 600 },
    children: [
      new TextRun({ text: "My AI is 'Huan Huan'. It helps Earth. It can sort trash.",
                    size: 13, italics: true, color: GRAY }),
    ],
  }),

  // ===== Now YOU write =====
  new Paragraph({
    spacing: { before: 200, after: 80 },
    children: [
      new TextRun({ text: '✨ 现在 你 来 写!  Now Your Turn!', size: 18, bold: true, color: CYBER }),
    ],
  }),
  // Frame: 聪明主人怎么用它
  new Paragraph({
    spacing: { before: 120, after: 60 },
    children: [
      new TextRun({ text: '聪明 主人 怎么 用 它? ', size: 24, bold: true, color: DARK }),
      new TextRun({ text: '_____________________________', size: 24, color: GRAY }),
      new TextRun({ text: ' 。', size: 24, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 0 },
    children: [
      new TextRun({ text: 'A smart master uses it by ', size: 16, italics: true, color: GRAY }),
      new TextRun({ text: '___________ .', size: 16, color: GRAY }),
    ],
  }),
];

// ===== §3 连一连 / Match — 我会认: 人工智能 机器人 电脑 学习 =====
const matchWords = [
  { char: '人工 智能', py: 'rén gōng zhì néng', en: 'AI',       em: '🤖' },
  { char: '机器人',    py: 'jī qì rén',         en: 'Robot',    em: '🤖' },
  { char: '电脑',      py: 'diàn nǎo',          en: 'Computer', em: '💻' },
  { char: '帮助',      py: 'bāng zhù',          en: 'Help',     em: '🤝' },
  { char: '学习',      py: 'xué xí',            en: 'Learn',    em: '📚' },
];

// Shuffle right column so pairs don't line up
const matchShuffled = [matchWords[2], matchWords[4], matchWords[0], matchWords[1], matchWords[3]];
const matchRows = matchWords.map((w, i) => {
  const right = matchShuffled[i];
  const colW = Math.floor(CW / 2);
  return new TableRow({
    height: { value: 1100, rule: 'atLeast' },
    children: [
      new TableCell({
        width: { size: colW, type: WidthType.DXA },
        borders: allBorders(border(ORANGE, 8)),
        margins: { top: 200, bottom: 200, left: 240, right: 240 },
        verticalAlign: 'center',
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: w.char, bold: true, size: 40, color: DARK })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: w.py, size: 20, color: GRAY, italics: true })],
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
  shadedBar('三、连一连 / Match  (用 线 连 起来)', ORANGE, 24),
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

// ===== §4 描一描, 写一写 — 电脑 · 学习 =====
const section4Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('四、描 一 描, 写 一 写 / Trace and Write  (电脑 · 学习)', GREEN, 24),
  new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({
      text: '👉 在 下面 贴 上 田字格 写字 纸, 写 一 写 今天 学 到 的 字: 电脑 · 学习。',
      size: 22, italics: true, color: GRAY,
    })],
  }),
  new Paragraph({
    spacing: { before: 60, after: 200 },
    children: [new TextRun({
      text: "Insert your grid-paper below and practice today's characters: 电脑 · 学习.",
      size: 20, italics: true, color: GRAY,
    })],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders(border(GREEN, 12)),
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
            text: '📄  在 这里 贴 上 田字格 写字 纸 / Insert your grid paper here',
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
