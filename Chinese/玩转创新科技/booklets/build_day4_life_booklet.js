// Build day4_life_booklet.docx — 玩转 创新 科技 Unit · Day 4: 科技 改变 生活 · 造纸 术
// Based on day4_technology changes life -造纸术.pptx (final 37-slide deck)
// Same 4-section structure as Day 1/2/3 booklets
// Run: node build_day4_life_booklet.js

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
} = require('docx');

const OUT = path.join(__dirname, 'day4_life_booklet.docx');

// ===== Palette (Day 4 — 科技 改变 生活 / 造纸 术) =====
const INK       = '0F1A3A';
const LIFE_TEAL = '00796B';   // §1 bar — Day 4 primary (matches deck)
const ANCIENT   = 'B85042';   // ancient China accent
const CYBER     = '2A47E0';   // §2 bar — tech timeline blue
const ORANGE    = 'FB8C00';   // §3 bar
const PURPLE    = '6A1B9A';   // §4 bar
const PINK      = 'D81B60';
const STAR      = 'F5C242';
const GOLD      = 'FFB700';
const SKY       = '42A5F5';
const ML_GREEN  = '2E7D32';
const DARK      = '2C2C2C';
const GRAY      = '888888';
const LGRAY     = 'D8D8D8';

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

// NOTE: No `shading:` parameter anywhere on body cells — keeps backgrounds truly
// transparent (Word/Pages render as page color, no grey tint).

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
    children: [new TextRun({ text: '🔥  📜  🤖', size: 96 })],
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
    children: [new TextRun({ text: 'Day 4 · 科技 改变 生活', bold: true, size: 44, color: DARK })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 200 },
    children: [new TextRun({ text: '从 古代 到 现代 · 造纸 术', bold: true, size: 30, color: LIFE_TEAL })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 200 },
    children: [new TextRun({ text: "From Ancient to Modern · The Art of Paper", italics: true, size: 24, color: GRAY })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 800 },
    children: [new TextRun({ text: '🔥  🛞  📜  🖨️  💡  🤖', size: 30, color: LIFE_TEAL })],
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

// ===== §1 题库 · 8 个 科技 改变 生活 小问题 =====
const qbCardColors = [LIFE_TEAL, ANCIENT, ORANGE, CYBER, PINK, SKY, PURPLE, LIFE_TEAL];

const qbQuestions = [
  { em: '🔥', cn: '人类 学 会 用 火, 解决 了 什么 问题?',
    en: 'What problem did fire solve for humans?',
    a: '又 冷 又 没 法 吃 熟 的 食物',
    b: '没有 手 机 玩' },
  { em: '📜', cn: '是 谁 改进 了 造纸 术?',
    en: 'Who improved the art of paper-making?',
    a: '中国 古代 的 蔡 伦 (东汉)',
    b: '美国 的 爱迪生' },
  { em: '🎋', cn: '在 纸 没 有 以 前, 古 人 写 字 用 什么?',
    en: 'Before paper, what did ancient people write on?',
    a: '石头 / 竹简 / 丝绸 — 太 重 + 太 贵',
    b: '塑料 板 / 平板 电脑' },
  { em: '🔧', cn: '科技 是 为 了 做 什么 而 发明 的?',
    en: 'Why are technologies invented?',
    a: '帮 人 解决 真实 的 问题',
    b: '让 大家 都 没 事 做' },
  { em: '🏛️', cn: '造纸 术 是 中国 「___ 大 发明」 之 一?',
    en: "Paper-making is one of China's ___ Great Inventions?",
    a: '四 大 发明 (火药 / 指南针 / 印刷 / 造纸)',
    b: '十 大 发明' },
  { em: '📚', cn: '有 了 纸 + 印刷 以 后 — 会 发生 什么?',
    en: 'After paper + printing came, what happened?',
    a: '更 多 书, 知识 传 得 快 + 更 多 人 学 习',
    b: '只 有 国 王 一 个 人 可以 看 书' },
  { em: '⏰', cn: '过去 没 有 电脑, 现在 我们 用 电脑 — 这 说 明 什么?',
    en: 'Past: no computers. Now: yes. What does this show?',
    a: '科技 一 直 在 进 步, 改变 我们 的 生活',
    b: '过去 的 人 比 我们 聪明' },
  { em: '🌱', cn: '如果 没 有 纸 — 我们 今天 的 学校 会 怎么 样?',
    en: 'Without paper, what would school look like today?',
    a: '没 课本, 没 作业 本, 知识 难 传 下 去',
    b: '上 学 会 更 轻 松' },
];

function qbCell(num, q, color) {
  return new TableCell({
    width: { size: Math.floor(CW / 2), type: WidthType.DXA },
    borders: allBorders(border(color, 10)),
    // no shading — transparent background
    margins: { top: 180, bottom: 180, left: 200, right: 200 },
    verticalAlign: 'center',
    children: [
      // Question line — bigger font (24 half-points = 12pt)
      new Paragraph({
        spacing: { before: 0, after: 60 },
        children: [
          new TextRun({ text: `${num}  `, bold: true, size: 32, color }),
          new TextRun({ text: `${q.em}  `, size: 30 }),
          new TextRun({ text: q.cn, bold: true, size: 24, color: DARK }),
        ],
      }),
      // EN translation — also bigger
      new Paragraph({
        spacing: { before: 0, after: 120 },
        indent: { left: 600 },
        children: [new TextRun({ text: q.en, italics: true, size: 18, color: GRAY })],
      }),
      // Options
      new Paragraph({
        spacing: { before: 0, after: 80 },
        indent: { left: 600 },
        children: [
          new TextRun({ text: '☐  A.  ', bold: true, size: 22, color: DARK }),
          new TextRun({ text: q.a, size: 22, color: DARK }),
        ],
      }),
      new Paragraph({
        spacing: { before: 0, after: 0 },
        indent: { left: 600 },
        children: [
          new TextRun({ text: '☐  B.  ', bold: true, size: 22, color: DARK }),
          new TextRun({ text: q.b, size: 22, color: DARK }),
        ],
      }),
    ],
  });
}

const qbRows = [];
for (let row = 0; row < 4; row++) {
  qbRows.push(new TableRow({
    height: { value: 2100, rule: 'atLeast' },
    children: [
      qbCell(row * 2 + 1, qbQuestions[row * 2],     qbCardColors[row * 2]),
      qbCell(row * 2 + 2, qbQuestions[row * 2 + 1], qbCardColors[row * 2 + 1]),
    ],
  }));
}

const section1Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('一、题库 · 8 个 科技 改变 生活 小 问题  (圈出 正确 答案)', LIFE_TEAL, 24),
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
      new TextRun({ text: '答对 8 题 = 「科技 小 工程师」 徽章! ', bold: true, size: 20, color: LIFE_TEAL }),
      new TextRun({ text: 'All 8 right = Tech Engineer badge!', italics: true, size: 16, color: GRAY }),
    ],
  }),
];

// ===== §2 科技 时间 排 序 / Technology Timeline Sequencing =====
// 6 inventions shown in shuffled order — students write 1-6 in the correct order.

const techSteps = [
  { em: '🔥', cn: '火',     en: 'Fire',     detail: '~150万 年 前 / 远古 人类' },
  { em: '🛞', cn: '轮 子',  en: 'Wheel',    detail: '~公元 前 3500 年 / 苏美 尔 人' },
  { em: '📜', cn: '纸',     en: 'Paper',    detail: '公元 105 年 / 中国 蔡 伦' },
  { em: '☎️', cn: '电 话',  en: 'Phone',    detail: '1876 年 / 美国 贝 尔' },
  { em: '💻', cn: '电 脑',  en: 'Computer', detail: '1946 年 / 美国 ENIAC' },
  { em: '🤖', cn: 'AI',     en: 'AI',       detail: '1956 年 / 美国 麦 卡 锡' },
];

// Shuffled display order so kids must think (NOT chronological)
const shuffleOrder = [2, 5, 0, 4, 1, 3];  // 0-indexed into techSteps
const shuffledTech = shuffleOrder.map(i => ({ ...techSteps[i], correctNumber: i + 1 }));
const stepColors = [LIFE_TEAL, CYBER, ORANGE, PINK, PURPLE, ML_GREEN];

function techCell(step, color) {
  return new TableCell({
    width: { size: Math.floor(CW / 2), type: WidthType.DXA },
    borders: allBorders(border(color, 10)),
    margins: { top: 220, bottom: 220, left: 200, right: 200 },
    verticalAlign: 'center',
    children: [
      new Paragraph({
        spacing: { before: 0, after: 60 },
        children: [
          new TextRun({ text: '☐ ', bold: true, size: 44, color }),
          new TextRun({ text: '  ' + step.em + '  ', size: 36 }),
          new TextRun({ text: step.cn, bold: true, size: 30, color: DARK }),
        ],
      }),
      new Paragraph({
        spacing: { before: 0, after: 80 },
        indent: { left: 800 },
        children: [new TextRun({ text: step.en, italics: true, size: 18, color: GRAY })],
      }),
      new Paragraph({
        spacing: { before: 0, after: 0 },
        indent: { left: 800 },
        children: [new TextRun({ text: step.detail, size: 20, color: DARK })],
      }),
    ],
  });
}

const techRows = [];
for (let row = 0; row < 3; row++) {
  techRows.push(new TableRow({
    height: { value: 1700, rule: 'atLeast' },
    children: [
      techCell(shuffledTech[row * 2],     stepColors[row * 2]),
      techCell(shuffledTech[row * 2 + 1], stepColors[row * 2 + 1]),
    ],
  }));
}

const section2Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('二、科技 时间 旅行 排 序 / Tech Timeline Sequencing', CYBER, 22),
  new Paragraph({
    spacing: { before: 160, after: 80 },
    children: [new TextRun({
      text: '🔢 这 6 个 发明 顺 序 乱 了! 请 在 □ 里 写 1-6, 排 出 「最 早 → 最 晚」 的 顺 序。',
      size: 22, bold: true, color: DARK,
    })],
  }),
  new Paragraph({
    spacing: { before: 40, after: 160 },
    children: [new TextRun({
      text: 'Put these 6 inventions in order from earliest (1) to latest (6).',
      size: 18, italics: true, color: GRAY,
    })],
  }),
  new Paragraph({
    spacing: { before: 0, after: 160 },
    children: [
      new TextRun({ text: '💡 提示  Hint:  ', size: 18, bold: true, color: LIFE_TEAL }),
      new TextRun({ text: '人类 最 早 学 会 用 「火」, 最 晚 才 有 「AI」。',
                    size: 18, color: DARK }),
      new TextRun({ text: '  Fire was first; AI is most recent.',
                    size: 16, italics: true, color: GRAY }),
    ],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [Math.floor(CW / 2), Math.floor(CW / 2)],
    borders: allBorders(noBorder()),
    rows: techRows,
  }),
  new Paragraph({
    spacing: { before: 240, after: 0 },
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({ text: '✅ ', size: 22 }),
      new TextRun({ text: '排 好 了 — 请 老师 检 查! ', bold: true, size: 20, color: CYBER }),
      new TextRun({ text: 'When you finish, ask your teacher to check.',
                    italics: true, size: 16, color: GRAY }),
    ],
  }),
];

// ===== §3 连 一 连 / Match — 发明 / 科技 / 纸 / 过去 / 现在 =====
const matchWords = [
  { char: '发明', en: 'Invent',     em: '💡' },
  { char: '科技', en: 'Technology', em: '🚀' },
  { char: '纸',   en: 'Paper',      em: '📜' },
  { char: '过去', en: 'Past',       em: '⏰' },
  { char: '现在', en: 'Now',        em: '📱' },
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

// ===== §4 描一描, 写一写 — 纸 / 过去 / 现在 (3 chars from 我会写) =====
const section4Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('四、描 一 描, 写 一 写 / Trace and Write  (纸 · 过去 · 现在)', PURPLE, 24),
  new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({
      text: '👉 在 下面 贴 上 田字格 写字 纸, 写 一 写 今天 学 到 的 字: 纸 · 过去 · 现在。',
      size: 22, italics: true, color: GRAY,
    })],
  }),
  new Paragraph({
    spacing: { before: 60, after: 200 },
    children: [new TextRun({
      text: "Insert your grid-paper below and practice today's characters: 纸 · 过去 · 现在.",
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
        // no shading — transparent
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
