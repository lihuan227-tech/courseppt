// Build day2_3dprint_booklet.docx — 玩转 创新 科技 Unit · Day 2: 从 活字 印刷 到 3D 打印
// Based on day2_3dprint.pptx (final 28-slide deck)
// Modeled on build_day1_ai_booklet.js (same 4-section structure)
// Run: node build_day2_3dprint_booklet.js

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
} = require('docx');

const OUT = path.join(__dirname, 'day2_3dprint_booklet.docx');

// ===== Palette (Day 2 — 从 活字 印刷 到 3D 打印) =====
const INK       = '0F1A3A';
const ANCIENT   = 'B85042';   // §1 bar — terracotta (古代 中国 色调)
const ORANGE    = 'FB8C00';   // §2 bar (PRINT_ORANGE)
const MODERN    = '2A47E0';   // §3 bar — modern tech blue
const GREEN     = '2E7D32';   // §4 bar
const PURPLE    = '6A1B9A';
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
    children: [new TextRun({ text: '📜  🖨️', size: 96 })],
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
    children: [new TextRun({ text: 'Day 2 · 从 活字 印刷 到 3D 打印', bold: true, size: 38, color: DARK })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 200 },
    children: [new TextRun({ text: '从 古代 发明 到 未来 制造', bold: true, size: 26, color: ANCIENT })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 200 },
    children: [new TextRun({ text: 'From Movable Type to 3D Printing', italics: true, size: 22, color: GRAY })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 800 },
    children: [new TextRun({ text: '📜  →  🖨️  →  ✨', size: 30, color: ORANGE })],
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

// ===== §1 题库 · 8 个 印刷 + 3D 打印 小问题 =====
const qbCardColors = [ANCIENT, ORANGE, PURPLE, PINK, MODERN, SKY, GREEN, ANCIENT];

const qbQuestions = [
  { em: '📜', cn: '是 谁 发明 了 活字 印刷 术?',
    en: 'Who invented movable type printing?',
    a: '毕昇 (宋朝 人)',
    b: '老子' },
  { em: '🇨🇳', cn: '活字 印刷 是 中国 的 「___ 大 发明」 之 一?',
    en: "It's one of China's ___ Great Inventions?",
    a: '四 大 发明',
    b: '八 大 发明' },
  { em: '📚', cn: '古代 没有 打印机 — 一 本 书 是 怎么 做 出 来 的?',
    en: 'No printers in ancient times — how were books made?',
    a: '把 字 块 排好, 刷 上 墨, 再 压 在 纸 上',
    b: '用 复印机 印 出 来 的' },
  { em: '💡', cn: '活字 印刷 厉害 在 哪 里?',
    en: 'What makes movable type clever?',
    a: '每 个 字 都 可以 重 复 使 用',
    b: '一 块 木 板 只 能 印 一 次' },
  { em: '🖨️', cn: '3D 打印 和 普通 打印, 有 什么 不 一 样?',
    en: 'How is 3D printing different from regular printing?',
    a: '3D 打印 出来 的 是 立体 的, 可以 拿 在 手 上',
    b: '一 样, 都 印 在 纸 上' },
  { em: '🍦', cn: '3D 打印 机 是 怎么 把 东 西 做 出 来 的?',
    en: 'How does a 3D printer make an object?',
    a: '一 层 一 层 堆 起 来 — 像 挤 奶 油',
    b: '「叮」 一 下 就 变 出 来 了' },
  { em: '🦷', cn: '现在 的 牙医, 真的 用 3D 打印 做 假牙 吗?',
    en: 'Do dentists really use 3D printing for crowns?',
    a: '是 的 — 真 的 在 用!',
    b: '从 来 没 用 过' },
  { em: '⚖️', cn: '活字 印刷 和 3D 打印, 都 在 做 同 一 件 事 — 是 什么?',
    en: 'Movable type & 3D printing both do what?',
    a: '把 「想 法」 变 成 真 实 的 东 西',
    b: '都 是 用 来 赚 钱 的' },
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
  shadedBar('一、题库 · 8 个 印刷 + 3D 打印 小 问题  (圈出 正确 答案)', ANCIENT, 24),
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
      new TextRun({ text: '答对 8 题 = 「小 毕昇 + 小 工程师」 徽章! ', bold: true, size: 20, color: ANCIENT }),
      new TextRun({ text: 'All 8 right = Mini Bi Sheng + Mini Engineer badge!', italics: true, size: 16, color: GRAY }),
    ],
  }),
];

// ===== §2 设计 你 想 印 的 东西 — open creative + draw + write =====
const section2Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('二、设计 你 想 印 的 东西 / Design What YOU Want to Print', ORANGE, 22),
  new Paragraph({
    spacing: { before: 160, after: 80 },
    children: [new TextRun({
      text: '🛠️ 想 一 想 — 如果 你 有 一 台 3D 打印机, 你 想 印 什么? 为什么?',
      size: 24, bold: true, color: DARK,
    })],
  }),
  new Paragraph({
    spacing: { before: 40, after: 200 },
    children: [new TextRun({
      text: 'Imagine you have a 3D printer — what would you print? Why?',
      size: 18, italics: true, color: GRAY,
    })],
  }),
  new Paragraph({
    spacing: { before: 100, after: 40 },
    children: [
      new TextRun({ text: '🎨 画一画  Draw it', size: 22, bold: true, color: ORANGE }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 80 },
    children: [
      new TextRun({ text: '画 出 你 的 发明 / 物 品 — 形状? 大小? 颜色?', size: 14, color: GRAY }),
      new TextRun({ text: '  ·  ', size: 14, color: LGRAY }),
      new TextRun({ text: "Draw your invention — shape, size, colors", size: 14, italics: true, color: GRAY }),
    ],
  }),
  new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders(border(ORANGE, 12)),
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
      new TextRun({ text: '✏️ 写一写  Write it', size: 22, bold: true, color: ORANGE }),
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
    children: [new TextRun({ text: '💡 例子  Examples:', size: 18, bold: true, color: ANCIENT })],
  }),
  // Example 1
  new Paragraph({
    spacing: { before: 60, after: 0 },
    indent: { left: 300 },
    children: [
      new TextRun({ text: '① ', bold: true, size: 20, color: ANCIENT }),
      new TextRun({ text: '我 想 打 印 一 只 小 恐 龙, 送 给 弟 弟 玩, 颜 色 是 绿 色 的。',
                    size: 18, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 100 },
    indent: { left: 600 },
    children: [
      new TextRun({ text: 'I want to print a little green dinosaur as a gift for my brother.',
                    size: 13, italics: true, color: GRAY }),
    ],
  }),
  // Example 2
  new Paragraph({
    spacing: { before: 0, after: 0 },
    indent: { left: 300 },
    children: [
      new TextRun({ text: '② ', bold: true, size: 20, color: ANCIENT }),
      new TextRun({ text: '我 想 打 印 一 个 钥 匙 扣, 上 面 写 着 我 的 名 字。',
                    size: 18, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 100 },
    indent: { left: 600 },
    children: [
      new TextRun({ text: "I want to print a keychain with my name on it.",
                    size: 13, italics: true, color: GRAY }),
    ],
  }),
  // Example 3
  new Paragraph({
    spacing: { before: 0, after: 0 },
    indent: { left: 300 },
    children: [
      new TextRun({ text: '③ ', bold: true, size: 20, color: ANCIENT }),
      new TextRun({ text: '我 想 打 印 一 只 假 手, 帮 没 有 手 的 小 朋 友。',
                    size: 18, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 200 },
    indent: { left: 600 },
    children: [
      new TextRun({ text: "I want to print a prosthetic hand to help a friend who needs one.",
                    size: 13, italics: true, color: GRAY }),
    ],
  }),

  // ===== Now YOU write =====
  new Paragraph({
    spacing: { before: 200, after: 80 },
    children: [
      new TextRun({ text: '✨ 现在 你 来 写!  Now Your Turn!', size: 18, bold: true, color: ORANGE }),
    ],
  }),
  // Frame: 我想印___, 给___, 用___
  new Paragraph({
    spacing: { before: 120, after: 60 },
    children: [
      new TextRun({ text: '我 想 印 ', size: 24, bold: true, color: DARK }),
      new TextRun({ text: '_________________', size: 24, color: GRAY }),
      new TextRun({ text: ', 它 是 给 ', size: 24, bold: true, color: DARK }),
      new TextRun({ text: '_____________', size: 24, color: GRAY }),
      new TextRun({ text: ' 的。', size: 24, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 160 },
    children: [
      new TextRun({ text: 'I want to print ', size: 16, italics: true, color: GRAY }),
      new TextRun({ text: '___________ , for ___________ .', size: 16, color: GRAY }),
    ],
  }),
  // Why?
  new Paragraph({
    spacing: { before: 120, after: 60 },
    children: [
      new TextRun({ text: '因为 ', size: 24, bold: true, color: DARK }),
      new TextRun({ text: '_____________________________________', size: 24, color: GRAY }),
      new TextRun({ text: ' 。', size: 24, color: DARK }),
    ],
  }),
  new Paragraph({
    spacing: { before: 0, after: 0 },
    children: [
      new TextRun({ text: 'Because ', size: 16, italics: true, color: GRAY }),
      new TextRun({ text: '___________________________ .', size: 16, color: GRAY }),
    ],
  }),
];

// ===== §3 连一连 / Match — 我会认: 活字印刷 打印 设计 机器 模型 =====
const matchWords = [
  { char: '活字 印刷', py: 'huó zì yìn shuā', en: 'Movable Type', em: '📜' },
  { char: '打印',     py: 'dǎ yìn',          en: 'Print',        em: '🖨️' },
  { char: '设计',     py: 'shè jì',          en: 'Design',       em: '✏️' },
  { char: '机器',     py: 'jī qì',           en: 'Machine',      em: '⚙️' },
  { char: '模型',     py: 'mó xíng',         en: 'Model',        em: '🧊' },
];

// Shuffle right column so pairs don't line up
const matchShuffled = [matchWords[3], matchWords[0], matchWords[4], matchWords[2], matchWords[1]];
const matchRows = matchWords.map((w, i) => {
  const right = matchShuffled[i];
  const colW = Math.floor(CW / 2);
  return new TableRow({
    height: { value: 1100, rule: 'atLeast' },
    children: [
      new TableCell({
        width: { size: colW, type: WidthType.DXA },
        borders: allBorders(border(MODERN, 8)),
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
  shadedBar('三、连一连 / Match  (用 线 连 起来)', MODERN, 24),
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

// ===== §4 描一描, 写一写 — 打印 · 机器 =====
const section4Children = [
  new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] }),
  shadedBar('四、描 一 描, 写 一 写 / Trace and Write  (打印 · 机器)', GREEN, 24),
  new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({
      text: '👉 在 下面 贴 上 田字格 写字 纸, 写 一 写 今天 学 到 的 字: 打印 · 机器。',
      size: 22, italics: true, color: GRAY,
    })],
  }),
  new Paragraph({
    spacing: { before: 60, after: 200 },
    children: [new TextRun({
      text: "Insert your grid-paper below and practice today's characters: 打印 · 机器.",
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
