// Build day3_ml_booklet.docx — 玩转 创新 科技 Unit · Day 3: 机器 学习 · 我 是 AI 训练 师
// Based on day3_ml-final.pdf (62 student-visible slides + 1 hidden teacher slide)
// Modeled on build_day2_3dprint_booklet.js (same 4-section structure)
// Run: node build_day3_ml_booklet.js

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
} = require('docx');

const OUT = path.join(__dirname, 'day3_ml_booklet.docx');

// ===== Palette (Day 3 — 机器 学习 / Machine Learning) =====
const INK       = '0F1A3A';
const ML_GREEN  = '2E7D32';   // §1 bar — Day 3 primary (matches deck's ML_GREEN)
const CYBER     = '2A47E0';   // §2 bar — tech/data blue
const ORANGE    = 'FB8C00';   // §3 bar — match vocab
const PURPLE    = '6A1B9A';   // §4 bar — trace + write
const PINK      = 'D81B60';
const STAR      = 'F5C242';
const GOLD      = 'FFB700';
const SKY       = '42A5F5';
const RED       = 'C8253E';
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
    children: [new TextRun({ text: '🤖  🧠', size: 96 })],
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
    children: [new TextRun({ text: 'Day 3 · 机器 学习', bold: true, size: 44, color: DARK })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 200 },
    children: [new TextRun({ text: '我 是 AI 训练 师', bold: true, size: 30, color: ML_GREEN })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 200 },
    children: [new TextRun({ text: "I'm an AI Trainer · Machine Learning Lab", italics: true, size: 24, color: GRAY })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 800 },
    children: [new TextRun({ text: '🤖  📷  🧠  ✨  🎯', size: 30, color: ML_GREEN })],
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

// ===== §1 题库 · 8 个 机器 学习 小问题 =====
const qbCardColors = [ML_GREEN, CYBER, ORANGE, PINK, PURPLE, SKY, ML_GREEN, GOLD];

const qbQuestions = [
  { em: '🧠', cn: '「机器 学习」 是 什么 意思?',
    en: 'What does "machine learning" mean?',
    a: '电脑 看 很 多 例子, 学 会 找 规律',
    b: '机器 自己 会 飞 + 会 说 话' },
  { em: '🍴', cn: 'AI 「吃」 什么 长 大?',
    en: 'What does AI "eat" to grow up?',
    a: '数据 — 很 多 很 多 图片 / 例子',
    b: '电池 + 充 电 就 行' },
  { em: '📉', cn: '如果 AI 看 的 例子 太 少 — 会 怎么 样?',
    en: 'What if AI sees too few examples?',
    a: '容易 猜 错 — 它 学 得 不 够',
    b: '没 问 题 — AI 看 一 张 就 会' },
  { em: '🐱', cn: 'AI 看 什么 来 认 猫?',
    en: 'How does AI tell something is a cat?',
    a: '看 「特 征」 — 尖 耳朵, 胡须, 尾巴',
    b: '听 「喵 喵」 — 它 用 耳朵 听' },
  { em: '🐰', cn: 'AI 只 学 了 白色 的 兔子 — 看 到 白色 北极熊 会 说 什么?',
    en: 'Only trained on white rabbits — what does AI say to a polar bear?',
    a: '「兔子!」 — 因为 它 只 学 了 「白色 = 兔子」',
    b: '「北极熊!」 — AI 总 是 对 的' },
  { em: '🔧', cn: '如果 AI 犯 错 了, 怎么 办?',
    en: 'What do we do when AI makes mistakes?',
    a: '加 更 多 + 更 清 楚 的 例子, 再 训练 一 次',
    b: '把 AI 关 掉 — 再 也 不 用 了' },
  { em: '🏫', cn: 'AI 学 习 像 你 在 学校 — 哪 里 像?',
    en: 'How is AI learning like you at school?',
    a: '都 需要 好 的 学习 材料 + 多 练 习',
    b: '不 像 — AI 比 我 聪明 100 倍' },
  { em: '🏷️', cn: '「贴 标签」 是 什么 意思?',
    en: 'What does "labeling" mean?',
    a: '告诉 AI 「这 是 什么」 — 例如 「这 是 猫」',
    b: '把 sticker 贴 到 电脑 屏幕 上' },
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
  shadedBar('一、题库 · 8 个 机器 学习 小 问题  (圈出 正确 答案)', ML_GREEN, 24),
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
      new TextRun({ text: '答对 8 题 = 「AI 训练师」 徽章! ', bold: true, size: 20, color: ML_GREEN }),
      new TextRun({ text: 'All 8 right = AI Trainer badge!', italics: true, size: 16, color: GRAY }),
    ],
  }),
];

// ===== §2 机器 学习 六 步 排 序 / ML 6-Step Sequencing =====
// 6 step cards shown in shuffled order — students write 1-6 in the empty number box

// Step definitions (the correct order is steps[0]=#1, steps[1]=#2, etc.)
const mlSteps = [
  { em: '📷', cn: '收 集 例 子', en: 'Give AI lots of example photos',
    detail: '给 AI 很 多 例子 (照片 / 图片)' },
  { em: '🏷️', cn: '贴 标 签',     en: "Tell AI what each one is (e.g. 'cat')",
    detail: '告 诉 AI: 这 是 什么 (例: 「猫」)' },
  { em: '🧠', cn: 'AI 学 习',     en: 'AI looks at examples one by one',
    detail: 'AI 一 张 一 张 看 例子' },
  { em: '🔍', cn: '找 规 律',     en: "AI finds what's common (features)",
    detail: 'AI 找 共同 点 (特 征)' },
  { em: '🧪', cn: '测 试',         en: 'Show AI a NEW image to test',
    detail: '给 AI 一 张 「新」 图 测 一 测' },
  { em: '✨', cn: '做 判 断',     en: 'AI uses learned patterns to predict',
    detail: 'AI 用 学 到 的 规律 猜!' },
];

// Shuffled display order (so it's a real sorting exercise — NOT sequential)
// Order shown: step 3, step 5, step 1, step 6, step 2, step 4
const shuffleOrder = [2, 4, 0, 5, 1, 3];  // 0-indexed into mlSteps
const shuffledSteps = shuffleOrder.map(i => ({ ...mlSteps[i], correctNumber: i + 1 }));

const stepColors = [ML_GREEN, CYBER, ORANGE, PINK, PURPLE, GOLD];

function stepCell(step, color) {
  return new TableCell({
    width: { size: Math.floor(CW / 2), type: WidthType.DXA },
    borders: allBorders(border(color, 10)),
    shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
    margins: { top: 200, bottom: 200, left: 200, right: 200 },
    verticalAlign: 'center',
    children: [
      // Empty number box + emoji row
      new Paragraph({
        spacing: { before: 0, after: 40 },
        children: [
          new TextRun({ text: '☐ ', bold: true, size: 40, color }),
          new TextRun({ text: '  ' + step.em + '  ', size: 32 }),
          new TextRun({ text: step.cn, bold: true, size: 26, color: DARK }),
        ],
      }),
      // English subtitle
      new Paragraph({
        spacing: { before: 0, after: 60 },
        indent: { left: 700 },
        children: [new TextRun({ text: step.en, italics: true, size: 14, color: GRAY })],
      }),
      // Detail line
      new Paragraph({
        spacing: { before: 0, after: 0 },
        indent: { left: 700 },
        children: [new TextRun({ text: step.detail, size: 16, color: DARK })],
      }),
    ],
  });
}

const stepRows = [];
for (let row = 0; row < 3; row++) {
  stepRows.push(new TableRow({
    height: { value: 1400, rule: 'atLeast' },
    children: [
      stepCell(shuffledSteps[row * 2],     stepColors[row * 2]),
      stepCell(shuffledSteps[row * 2 + 1], stepColors[row * 2 + 1]),
    ],
  }));
}

const section2Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('二、机器 学习 六 步 排 序 / ML 6-Step Sequencing', CYBER, 22),
  new Paragraph({
    spacing: { before: 160, after: 80 },
    children: [new TextRun({
      text: '🔢 这 6 个 步骤 顺 序 乱 了! 请 在 □ 里 写 1-6, 排 出 正 确 顺 序。',
      size: 22, bold: true, color: DARK,
    })],
  }),
  new Paragraph({
    spacing: { before: 40, after: 160 },
    children: [new TextRun({
      text: 'These 6 steps are out of order! Write 1-6 in the ☐ boxes to put them in the correct order.',
      size: 16, italics: true, color: GRAY,
    })],
  }),
  // Hint line
  new Paragraph({
    spacing: { before: 0, after: 160 },
    children: [
      new TextRun({ text: '💡 提示  Hint:  ', size: 16, bold: true, color: ML_GREEN }),
      new TextRun({ text: '先 给 AI 例子, 最 后 让 它 自己 「猜」。',
                    size: 16, color: DARK }),
      new TextRun({ text: '   First give examples — last, let AI guess.',
                    size: 14, italics: true, color: GRAY }),
    ],
  }),
  // The 6 step cards (shuffled)
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [Math.floor(CW / 2), Math.floor(CW / 2)],
    borders: allBorders(noBorder()),
    rows: stepRows,
  }),
  // Bottom check prompt
  new Paragraph({
    spacing: { before: 240, after: 0 },
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({ text: '✅ ', size: 20 }),
      new TextRun({ text: '排 好 了 — 请 老师 检 查! ', bold: true, size: 18, color: CYBER }),
      new TextRun({ text: 'When you finish, ask your teacher to check.',
                    italics: true, size: 14, color: GRAY }),
    ],
  }),
];

// ===== §3 连一连 / Match — 我会认: 机器 分类 数据 训练 特征 =====
const matchWords = [
  { char: '机器',  py: 'jī qì',     en: 'Machine',  em: '🤖' },
  { char: '分类',  py: 'fēn lèi',   en: 'Classify', em: '🗂️' },
  { char: '数据',  py: 'shù jù',    en: 'Data',     em: '📊' },
  { char: '训练',  py: 'xùn liàn',  en: 'Train',    em: '🏋️' },
  { char: '特征',  py: 'tè zhēng',  en: 'Features', em: '🔍' },
];

// Shuffle right column so pairs don't line up
const matchShuffled = [matchWords[3], matchWords[0], matchWords[4], matchWords[1], matchWords[2]];
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

// ===== §4 描一描, 写一写 — 分类 · 训练 (matches 我会写 from deck) =====
const section4Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('四、描 一 描, 写 一 写 / Trace and Write  (分类 · 训练)', PURPLE, 24),
  new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({
      text: '👉 在 下面 贴 上 田字格 写字 纸, 写 一 写 今天 学 到 的 字: 分类 · 训练。',
      size: 22, italics: true, color: GRAY,
    })],
  }),
  new Paragraph({
    spacing: { before: 60, after: 200 },
    children: [new TextRun({
      text: "Insert your grid-paper below and practice today's characters: 分类 · 训练.",
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
