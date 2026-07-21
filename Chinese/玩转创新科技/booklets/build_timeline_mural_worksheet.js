// Build timeline_mural_worksheet.docx
// 5-page worksheet for Day 5 Session 3 Project 1: 科技 改变 生活 大 卷 轴
// One printable page per era group (古代/知识/传播/工业/数字).
// Run: node build_timeline_mural_worksheet.js

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, PageBreak, HeightRule,
} = require('docx');

const OUT = path.join(__dirname, 'timeline_mural_worksheet.docx');

// ===== Palette =====
const INK     = '0F1A3A';
const DAY     = 'D81B60';   // Day 5 pink
const STAR    = 'F5C242';
const WARM    = 'FFF4D6';
const DARK    = '2C2C2C';
const GRAY    = '888888';
const LGRAY   = 'D8D8D8';
const WHITE   = 'FFFFFF';

// ===== Page geometry (US Letter, 0.75" margins) =====
const PAGE = {
  size: { width: 12240, height: 15840 },
  margin: { top: 720, right: 720, bottom: 720, left: 720 },
};
const CW = 12240 - 720 - 720;

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
        margins: { top: 100, bottom: 100, left: 200, right: 200 },
        children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: WHITE, size })] })],
      })],
    })],
  });
}

function spacer(h = 120) {
  return new Paragraph({ spacing: { before: 0, after: h }, children: [new TextRun({ text: '' })] });
}

// ===== Era data =====
const ERAS = [
  {
    group: 'Group 1',
    eraCn: '古 代 时 代',
    eraEn: 'Ancient World',
    color: 'B85042',    // terracotta
    candidates: [
      { em: '🔥', cn: '火',    en: 'Fire' },
      { em: '🛞', cn: '轮 子', en: 'Wheel' },
      { em: '🏹', cn: '弓 箭', en: 'Bow' },
      { em: '⛵', cn: '帆 船', en: 'Sailboat' },
      { em: '🌾', cn: '农 业', en: 'Agriculture' },
    ],
  },
  {
    group: 'Group 2',
    eraCn: '知 识 时 代',
    eraEn: 'Classical & Medieval',
    color: 'FB8C00',
    candidates: [
      { em: '📜', cn: '纸',       en: 'Paper' },
      { em: '🧭', cn: '指 南 针', en: 'Compass' },
      { em: '🧮', cn: '算 盘',    en: 'Abacus' },
      { em: '🏺', cn: '陶 器',    en: 'Pottery' },
      { em: '🏗️', cn: '水 渠',   en: 'Aqueduct' },
    ],
  },
  {
    group: 'Group 3',
    eraCn: '传 播 时 代',
    eraEn: 'Early Modern',
    color: '2E7D32',
    candidates: [
      { em: '🖨️', cn: '印 刷 术', en: 'Printing' },
      { em: '🔭', cn: '望 远 镜', en: 'Telescope' },
      { em: '⚙️', cn: '蒸 汽 机', en: 'Steam engine' },
      { em: '🧪', cn: '显 微 镜', en: 'Microscope' },
      { em: '🌡️', cn: '温 度 计', en: 'Thermometer' },
    ],
  },
  {
    group: 'Group 4',
    eraCn: '工 业 时 代',
    eraEn: 'Industrial Age',
    color: '2A47E0',
    candidates: [
      { em: '☎️', cn: '电 话', en: 'Phone' },
      { em: '💡', cn: '电 灯', en: 'Light bulb' },
      { em: '🚂', cn: '火 车', en: 'Train' },
      { em: '🚗', cn: '汽 车', en: 'Car' },
      { em: '✈️', cn: '飞 机', en: 'Airplane' },
    ],
  },
  {
    group: 'Group 5',
    eraCn: '数 字 时 代',
    eraEn: 'Digital Age',
    color: '6A1B9A',
    candidates: [
      { em: '💻', cn: '电 脑',        en: 'Computer' },
      { em: '🌐', cn: '互 联 网',      en: 'Internet' },
      { em: '📱', cn: '智 能 手 机',   en: 'Smartphone' },
      { em: '🤖', cn: 'AI',          en: 'AI' },
      { em: '🛰️', cn: 'GPS',         en: 'GPS' },
    ],
  },
];

// ===== Build one worksheet page =====
function buildPage(era, isFirst) {
  const children = [];
  if (!isFirst) {
    children.push(new Paragraph({ children: [new PageBreak()] }));
  }

  // ---- HEADER ----
  // Title bar
  children.push(new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders(noBorder()),
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        shading: { fill: DAY, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 240, right: 240 },
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: '📜  科技 改变 生活 大 卷 轴', bold: true, color: WHITE, size: 32 }),
            ],
          }),
          new Paragraph({
            spacing: { before: 60, after: 0 },
            children: [
              new TextRun({ text: 'Tech Time Travel Timeline Mural · 小 组 任 务 单', color: STAR, size: 20, bold: true }),
            ],
          }),
        ],
      })],
    })],
  }));
  children.push(spacer(120));

  // Group + era + members
  const cwHalf = Math.floor(CW / 2);
  children.push(new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [cwHalf, cwHalf],
    borders: allBorders(border(era.color, 12)),
    rows: [new TableRow({
      children: [
        new TableCell({
          width: { size: cwHalf, type: WidthType.DXA },
          shading: { fill: era.color, type: ShadingType.CLEAR },
          margins: { top: 160, bottom: 160, left: 240, right: 240 },
          verticalAlign: 'center',
          children: [
            new Paragraph({ children: [new TextRun({ text: era.group, bold: true, color: STAR, size: 22 })] }),
            new Paragraph({
              spacing: { before: 40 },
              children: [new TextRun({ text: era.eraCn, bold: true, color: WHITE, size: 36 })],
            }),
            new Paragraph({
              spacing: { before: 30 },
              children: [new TextRun({ text: era.eraEn, color: WARM, size: 18 })],
            }),
          ],
        }),
        new TableCell({
          width: { size: cwHalf, type: WidthType.DXA },
          margins: { top: 140, bottom: 140, left: 240, right: 240 },
          verticalAlign: 'center',
          children: [
            new Paragraph({
              children: [new TextRun({ text: '👥 组 员 姓 名  Group Members:', bold: true, color: DARK, size: 22 })],
            }),
            new Paragraph({
              spacing: { before: 120 },
              children: [new TextRun({ text: '1. _____________________', size: 22, color: GRAY })],
            }),
            new Paragraph({
              spacing: { before: 60 },
              children: [new TextRun({ text: '2. _____________________', size: 22, color: GRAY })],
            }),
            new Paragraph({
              spacing: { before: 60 },
              children: [new TextRun({ text: '3. _____________________', size: 22, color: GRAY })],
            }),
            new Paragraph({
              spacing: { before: 60 },
              children: [new TextRun({ text: '4. _____________________', size: 22, color: GRAY })],
            }),
          ],
        }),
      ],
    })],
  }));
  children.push(spacer(160));

  // ---- §1 VOTE ----
  children.push(shadedBar(`①  全 组 投 票 · Vote — 选 一 个 「最 重 要 的」 发 明`, era.color, 22));
  children.push(spacer(80));

  const cellW = Math.floor(CW / 5);
  children.push(new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [cellW, cellW, cellW, cellW, cellW],
    borders: allBorders(border(era.color, 8)),
    rows: [new TableRow({
      height: { value: 1800, rule: HeightRule.ATLEAST },
      children: era.candidates.map(c => new TableCell({
        width: { size: cellW, type: WidthType.DXA },
        borders: allBorders(border(era.color, 8)),
        margins: { top: 160, bottom: 160, left: 80, right: 80 },
        verticalAlign: 'center',
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: c.em, size: 56 })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 80 },
            children: [new TextRun({ text: c.cn, bold: true, color: era.color, size: 26 })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 30 },
            children: [new TextRun({ text: c.en, color: GRAY, size: 16 })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 120 },
            children: [new TextRun({ text: '☐', size: 36, color: era.color, bold: true })],
          }),
        ],
      })),
    })],
  }));
  children.push(spacer(160));

  // ---- §2 DISCUSSION ----
  children.push(shadedBar('②  小 组 讨 论 · Discussion — 一 起 回 答 这 3 个 问 题', era.color, 22));
  children.push(spacer(80));

  // Three question rows
  const discQs = [
    { num: '1', em: '✅', cn: '我 们 选 哪 个 发 明?', en: 'Which invention did we pick?',
      hint: '写 出 你 们 选 的 名 字  Write the name:' },
    { num: '2', em: '🌟', cn: '为 什 么 它 最 重 要?', en: 'Why is it the most important?',
      hint: '句 型: 「它 帮 了 我 们 ___」 / 「因 为 它, 人 们 可 以 ___」' },
    { num: '3', em: '🤔', cn: '如 果 没 有 它, 怎 么 办?', en: 'What if it didn\'t exist?',
      hint: '句 型: 「没 有 它, 人 们 就 ___」 / 「人 们 会 ___」' },
  ];

  for (const q of discQs) {
    children.push(new Table({
      width: { size: CW, type: WidthType.DXA },
      columnWidths: [CW],
      borders: allBorders(border(era.color, 6)),
      rows: [new TableRow({
        children: [new TableCell({
          width: { size: CW, type: WidthType.DXA },
          margins: { top: 140, bottom: 140, left: 240, right: 240 },
          children: [
            new Paragraph({
              children: [
                new TextRun({ text: `${q.num}.  `, bold: true, size: 28, color: era.color }),
                new TextRun({ text: `${q.em}  `, size: 28 }),
                new TextRun({ text: q.cn, bold: true, size: 24, color: DARK }),
              ],
            }),
            new Paragraph({
              spacing: { before: 40 },
              children: [new TextRun({ text: q.en, color: GRAY, size: 18, italics: true })],
            }),
            new Paragraph({
              spacing: { before: 60 },
              children: [new TextRun({ text: `💡 ${q.hint}`, color: era.color, size: 18, bold: true })],
            }),
            new Paragraph({
              spacing: { before: 100 },
              children: [new TextRun({ text: '_______________________________________________________________________', color: LGRAY, size: 22 })],
            }),
            new Paragraph({
              spacing: { before: 60 },
              children: [new TextRun({ text: '_______________________________________________________________________', color: LGRAY, size: 22 })],
            }),
          ],
        })],
      })],
    }));
    children.push(spacer(80));
  }
  children.push(spacer(80));

  // ---- §3 POSTER PLAN ----
  children.push(shadedBar('③  海 报 设 计 · Poster Plan — 在 大 海 报 上 这 样 画!', era.color, 22));
  children.push(spacer(80));

  // 5 rows: name+emoji, drawing box, why, what if, rating
  // Row 1: name
  children.push(new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders(border(era.color, 6)),
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        margins: { top: 140, bottom: 140, left: 240, right: 240 },
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: '1.  发 明 名 称 + emoji  ', bold: true, size: 24, color: era.color }),
              new TextRun({ text: 'Invention name & emoji', color: GRAY, size: 18, italics: true }),
            ],
          }),
          new Paragraph({
            spacing: { before: 100 },
            children: [new TextRun({ text: '___________________________________________________________  emoji:  ☐  ☐  ☐', color: LGRAY, size: 22 })],
          }),
        ],
      })],
    })],
  }));
  children.push(spacer(80));

  // Row 2: drawing box — bigger
  children.push(new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders(border(era.color, 6)),
    rows: [new TableRow({
      height: { value: 2800, rule: HeightRule.ATLEAST },
      children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        margins: { top: 140, bottom: 140, left: 240, right: 240 },
        verticalAlign: 'top',
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: '2.  画 一 张 大 图  ', bold: true, size: 24, color: era.color }),
              new TextRun({ text: 'Big illustration · 在 海 报 中 间 画', color: GRAY, size: 18, italics: true }),
            ],
          }),
          new Paragraph({
            spacing: { before: 800 },
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: '✏️', size: 28, color: LGRAY })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 60 },
            children: [new TextRun({ text: '( 在 大 海 报 上 画 — 这 里 只 写 提 示 )', color: GRAY, size: 16, italics: true })],
          }),
        ],
      })],
    })],
  }));
  children.push(spacer(80));

  // Row 3: why important
  children.push(new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders(border(era.color, 6)),
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        margins: { top: 140, bottom: 140, left: 240, right: 240 },
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: '3.  为 什 么 重 要?  ', bold: true, size: 24, color: era.color }),
              new TextRun({ text: 'Why is it important?  · 「它 帮 了 ___」', color: GRAY, size: 18, italics: true }),
            ],
          }),
          new Paragraph({
            spacing: { before: 100 },
            children: [new TextRun({ text: '它 帮 了 _____________________________________________________________', color: LGRAY, size: 22 })],
          }),
          new Paragraph({
            spacing: { before: 60 },
            children: [new TextRun({ text: '_______________________________________________________________________', color: LGRAY, size: 22 })],
          }),
        ],
      })],
    })],
  }));
  children.push(spacer(80));

  // Row 4: what if not
  children.push(new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders(border(era.color, 6)),
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        margins: { top: 140, bottom: 140, left: 240, right: 240 },
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: '4.  如 果 没 有 它 ...  ', bold: true, size: 24, color: era.color }),
              new TextRun({ text: 'Without it ...  · 「没 有 它, 人 们 就 ___」', color: GRAY, size: 18, italics: true }),
            ],
          }),
          new Paragraph({
            spacing: { before: 100 },
            children: [new TextRun({ text: '没 有 它, 人 们 就 ___________________________________________________', color: LGRAY, size: 22 })],
          }),
          new Paragraph({
            spacing: { before: 60 },
            children: [new TextRun({ text: '_______________________________________________________________________', color: LGRAY, size: 22 })],
          }),
        ],
      })],
    })],
  }));
  children.push(spacer(80));

  // Row 5: rating
  children.push(new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allBorders(border(era.color, 6)),
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        shading: { fill: WARM, type: ShadingType.CLEAR },
        margins: { top: 140, bottom: 140, left: 240, right: 240 },
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: '5.  影 响 评 分  ', bold: true, size: 24, color: era.color }),
              new TextRun({ text: 'Impact rating  · 圈 出 几 颗 星  Circle stars (1-5):', color: GRAY, size: 18, italics: true }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 100 },
            children: [
              new TextRun({ text: '⭐   ⭐   ⭐   ⭐   ⭐', size: 52, color: era.color }),
            ],
          }),
        ],
      })],
    })],
  }));

  return children;
}

// ===== Build all 5 pages =====
const allChildren = [];
ERAS.forEach((era, i) => {
  allChildren.push(...buildPage(era, i === 0));
});

const doc = new Document({
  sections: [{
    properties: { page: PAGE },
    children: allChildren,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log(`Saved ${OUT}  (${ERAS.length} pages)`);
});
