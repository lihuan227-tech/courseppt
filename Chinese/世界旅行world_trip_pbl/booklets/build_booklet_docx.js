// Build asia_booklet_AwithChina.docx from the same content as the HTML.
// Uses docx-js directly (not LibreOffice) so we get a real Writer document
// with proper page breaks and embedded images.
//
// Run: node build_booklet_docx.js

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, BorderStyle, WidthType, ShadingType, PageBreak,
  HeadingLevel, LevelFormat,
} = require('docx');

const ASSETS = '/tmp/booklet_assets';
const OUT = path.join(__dirname, 'asia_booklet_AwithChina.docx');

// ===== Colors (hex without #) =====
const COLOR_ACCENT = 'FF8F00';  // orange — Asia section
const COLOR_CHINA = 'DE2910';   // red
const COLOR_JAPAN = 'BC002D';   // dark red
const COLOR_INDIA = '138808';   // green

// ===== Page geometry =====
const PAGE = {
  size: { width: 12240, height: 15840 }, // US Letter
  margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }, // 0.75" margins
};
const CONTENT_WIDTH = 12240 - 1080 - 1080; // 10080 DXA

// ===== Helpers =====

// A colored "shaded bar" paragraph — white bold text on a colored background, via a single-cell table.
function shadedBar(text, colorHex) {
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH],
    borders: allBorders('none'),
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: CONTENT_WIDTH, type: WidthType.DXA },
        shading: { fill: colorHex, type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 160, right: 160 },
        children: [new Paragraph({
          children: [new TextRun({ text, bold: true, color: 'FFFFFF', size: 22 })],
        })],
      })],
    })],
  });
}

function allBorders(style = 'single', size = 4, color = '999999') {
  const b = style === 'none'
    ? { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }
    : { style: BorderStyle.SINGLE, size, color };
  return { top: b, bottom: b, left: b, right: b, insideHorizontal: b, insideVertical: b };
}

// A regular paragraph of plain text
function p(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text, size: 22, ...opts })],
  });
}

// Bold paragraph (question label)
function pBold(text) { return p(text, { bold: true }); }

// Heading paragraph (country title, large + colored)
function countryHeading(emojiFlag, zh, en, colorHex) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 200, after: 200 },
    children: [new TextRun({ text: `${emojiFlag}  ${zh} ${en}`, bold: true, size: 44, color: colorHex })],
  });
}

// Image paragraph (centered)
function imgPara(fn, widthPx) {
  const data = fs.readFileSync(fn);
  const ext = path.extname(fn).slice(1).toLowerCase().replace('jpeg', 'jpg');
  // px -> EMU/pt: 96 DPI -> width in pixels is a good docx display size
  // docx width/height are in pixels when using transformation per the library
  const sizePx = widthPx;
  // Estimate height from actual image dimensions
  let heightPx = sizePx; // fallback square
  try {
    const sharp = null; // we don't have sharp; use PIL via a sidecar .json? Skip: rely on width, height auto.
  } catch (_) {}
  // The docx library takes width/height in points (EMU later). Use aspect from raw header parse for PNGs/JPEGs.
  const dims = probeDims(fn);
  if (dims) heightPx = Math.round(sizePx * dims.h / dims.w);
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 80 },
    children: [new ImageRun({
      type: ext === 'jpg' ? 'jpg' : ext,
      data,
      transformation: { width: sizePx, height: heightPx },
      altText: { title: 'booklet image', description: path.basename(fn), name: path.basename(fn) },
    })],
  });
}

// Read PNG/JPEG width & height from file bytes
function probeDims(fn) {
  const buf = fs.readFileSync(fn);
  // PNG
  if (buf.length > 24 && buf[0] === 0x89 && buf[1] === 0x50) {
    return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
  }
  // JPEG
  if (buf[0] === 0xff && buf[1] === 0xd8) {
    let i = 2;
    while (i < buf.length) {
      if (buf[i] !== 0xff) return null;
      const marker = buf[i + 1];
      if (marker >= 0xc0 && marker <= 0xc3) {
        return { h: buf.readUInt16BE(i + 5), w: buf.readUInt16BE(i + 7) };
      }
      const len = buf.readUInt16BE(i + 2);
      i += 2 + len;
    }
  }
  return null;
}

// Trace boxes — a single-row table with N cells, each showing a big faint character inside a dashed border
function traceRow(chars, colorHex) {
  const cellW = Math.floor(CONTENT_WIDTH / chars.length);
  const dashed = { style: BorderStyle.DASHED, size: 6, color: 'CCCCCC' };
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: chars.map(() => cellW),
    borders: allBorders('none'),
    rows: [new TableRow({
      height: { value: 1600, rule: 'exact' },
      children: chars.map(ch => new TableCell({
        width: { size: cellW, type: WidthType.DXA },
        borders: { top: dashed, bottom: dashed, left: dashed, right: dashed },
        margins: { top: 120, bottom: 120, left: 120, right: 120 },
        verticalAlign: 'center',
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: ch, size: 96, color: 'DDDDDD' })],
        })],
      })),
    })],
  });
}

// Flag placeholder box — a table cell with a solid border and a fixed height
function flagBox(innerChildren = []) {
  const solid = { style: BorderStyle.SINGLE, size: 18, color: '333333' };
  return new Table({
    alignment: AlignmentType.CENTER,
    width: { size: 3600, type: WidthType.DXA }, // 2.5" wide
    columnWidths: [3600],
    borders: allBorders('none'),
    rows: [new TableRow({
      height: { value: 2400, rule: 'exact' }, // 1.67" tall (3:2 aspect)
      children: [new TableCell({
        width: { size: 3600, type: WidthType.DXA },
        borders: { top: solid, bottom: solid, left: solid, right: solid },
        verticalAlign: 'center',
        children: innerChildren.length ? innerChildren : [new Paragraph({ children: [new TextRun('')] })],
      })],
    })],
  });
}

// Draw & Write box — large bordered empty area for student drawing
function drawWriteBox(placeholder) {
  const dashed = { style: BorderStyle.DASHED, size: 8, color: 'BBBBBB' };
  const empty = new Paragraph({ children: [new TextRun('')] });
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH],
    borders: allBorders('none'),
    rows: [new TableRow({
      height: { value: 3400, rule: 'exact' }, // ~2.4 inches
      children: [new TableCell({
        width: { size: CONTENT_WIDTH, type: WidthType.DXA },
        borders: { top: dashed, bottom: dashed, left: dashed, right: dashed },
        margins: { top: 120, bottom: 120, left: 160, right: 160 },
        children: [
          new Paragraph({ children: [new TextRun({ text: placeholder, size: 20, color: 'AAAAAA', italics: true })] }),
          empty, empty, empty, empty,
        ],
      })],
    })],
  });
}

function pageBreak() { return new Paragraph({ children: [new PageBreak()] }); }

// ======================================================================
// CONTENT
// ======================================================================

const children = [];

// ------------------ PAGE 1 · COVER ------------------
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 2400, after: 200 },
  children: [new TextRun({ text: '🌏', size: 120 })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 200, after: 200 },
  children: [new TextRun({ text: '探索亚洲', bold: true, size: 80, color: 'C62828' })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 100, after: 600 },
  children: [new TextRun({ text: 'Explore: Asia', bold: true, size: 48 })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 200, after: 1200 },
  children: [new TextRun({ text: '🇨🇳 🇯🇵 🇮🇳', size: 64 })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 800, after: 200 },
  children: [new TextRun({ text: 'Global Explorer Camp', size: 28, color: '666666', italics: true })],
}));
children.push(pageBreak());

// ------------------ PAGE 2 · ABOUT ASIA ------------------
children.push(new Paragraph({
  spacing: { before: 100, after: 200 },
  children: [new TextRun({ text: '🌏 认识亚洲 About Asia', bold: true, size: 44, color: 'C62828' })],
}));
children.push(shadedBar('我会认 I Can Read — 连一连 Match Words to Pictures', COLOR_ACCENT));

// Match table: 词语 | ← 连线 → | 图片
const matchRows = [
  { word: '亚洲', pic: { kind: 'text', val: '🇯🇵' } },
  { word: '中国', pic: { kind: 'img', file: path.join(ASSETS, 'img_0.png'), w: 120 } },
  { word: '日本', pic: { kind: 'img', file: path.join(ASSETS, 'img_1.png'), w: 150 } },
  { word: '印度', pic: { kind: 'img', file: path.join(ASSETS, 'img_2.png'), w: 170 } },
  { word: '首都 (北京)', pic: { kind: 'img', file: path.join(ASSETS, 'img_3.png'), w: 190 } },
];

const matchBorder = { style: BorderStyle.SINGLE, size: 6, color: 'CCCCCC' };
const matchTable = new Table({
  width: { size: CONTENT_WIDTH, type: WidthType.DXA },
  columnWidths: [2400, 3600, 4080],
  borders: {
    top: matchBorder, bottom: matchBorder, left: matchBorder, right: matchBorder,
    insideHorizontal: matchBorder, insideVertical: matchBorder,
  },
  rows: [
    new TableRow({
      tableHeader: true,
      children: ['词语 Words', '← 连线 Draw lines →', '图片 Pictures'].map((t, i) => new TableCell({
        width: { size: [2400, 3600, 4080][i], type: WidthType.DXA },
        shading: { fill: COLOR_ACCENT, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: t, bold: true, color: 'FFFFFF', size: 22 })],
        })],
      })),
    }),
    ...matchRows.map(row => new TableRow({
      children: [
        new TableCell({
          width: { size: 2400, type: WidthType.DXA },
          verticalAlign: 'center',
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: row.word, size: 28 })],
          })],
        }),
        new TableCell({
          width: { size: 3600, type: WidthType.DXA },
          children: [new Paragraph({ children: [new TextRun('')] })],
        }),
        new TableCell({
          width: { size: 4080, type: WidthType.DXA },
          verticalAlign: 'center',
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [row.pic.kind === 'img'
            ? imgPara(row.pic.file, row.pic.w)
            : new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: row.pic.val, size: 52 })] })
          ],
        }),
      ],
    })),
  ],
});
children.push(matchTable);

children.push(p(''));
children.push(shadedBar('⭕ 描一描 Trace', COLOR_ACCENT));
// 2 trace images side by side in a 2-column table
const traceAsia = new Table({
  alignment: AlignmentType.CENTER,
  width: { size: CONTENT_WIDTH, type: WidthType.DXA },
  columnWidths: [CONTENT_WIDTH / 2, CONTENT_WIDTH / 2],
  borders: allBorders('none'),
  rows: [new TableRow({
    children: [
      new TableCell({
        width: { size: CONTENT_WIDTH / 2, type: WidthType.DXA },
        children: [imgPara(path.join(ASSETS, 'img_4.png'), 100)],
      }),
      new TableCell({
        width: { size: CONTENT_WIDTH / 2, type: WidthType.DXA },
        children: [imgPara(path.join(ASSETS, 'img_5.png'), 100)],
      }),
    ],
  })],
});
children.push(traceAsia);
children.push(p(''));
children.push(shadedBar('⭕ 圈一圈 Circle the Correct Answer', COLOR_ACCENT));
children.push(pBold('1. 亚洲有多少个国家？ How many countries are there in Asia?'));
children.push(p('   A. 20     B. 48     C. 100'));
children.push(pBold('2. 亚洲是世界上最______的洲。 Asia is the ______ continent in the world.'));
children.push(p('   A. 小 smallest     B. 冷 coldest     C. 大 largest'));
children.push(pBold('3. 世界上最高的山是？ What is the tallest mountain in the world?'));
children.push(p('   A. 富士山 Mt. Fuji     B. 珠穆朗玛峰 Mt. Everest     C. 黄山 Mt. Huang'));
children.push(p(''));
children.push(shadedBar('🌍 涂色 Color the Continents', COLOR_ACCENT));
children.push(pBold('在下面的世界地图上，给七大洲涂上不同的颜色'));
children.push(p('🔴 亚洲 Asia = 红色 Red     🟡 非洲 Africa = 黄色 Yellow     🔵 欧洲 Europe = 蓝色 Blue'));
children.push(p('🟢 北美洲 N.America = 绿色 Green     🟠 南美洲 S.America = 橙色 Orange     🟣 大洋洲 Oceania = 紫色 Purple'));
children.push(pageBreak());

// ------------------ PAGE 3 · CHINA (1) ------------------
children.push(countryHeading('🇨🇳', '中国', 'China', COLOR_CHINA));
children.push(shadedBar('涂一涂: 中国国旗 Color the Flag', COLOR_CHINA));
children.push(p('🔴 红色 Red + ⭐ 黄色 Yellow（五颗星）'));
children.push(imgPara(path.join(ASSETS, 'img_6.jpeg'), 220));

children.push(shadedBar('描一描 Trace', COLOR_CHINA));
const traceChina = new Table({
  alignment: AlignmentType.CENTER,
  width: { size: CONTENT_WIDTH, type: WidthType.DXA },
  columnWidths: [CONTENT_WIDTH / 2, CONTENT_WIDTH / 2],
  borders: allBorders('none'),
  rows: [new TableRow({
    children: [
      new TableCell({ width: { size: CONTENT_WIDTH / 2, type: WidthType.DXA }, children: [imgPara(path.join(ASSETS, 'img_7.png'), 110)] }),
      new TableCell({ width: { size: CONTENT_WIDTH / 2, type: WidthType.DXA }, children: [imgPara(path.join(ASSETS, 'img_8.png'), 110)] }),
    ],
  })],
});
children.push(traceChina);

children.push(shadedBar('⭕ 圈一圈 Circle the Correct Answer', COLOR_CHINA));
children.push(pBold('1. 中国的首都是？ What is the capital of China?'));
children.push(p('   A. 上海 Shanghai     B. 北京 Beijing     C. 东京 Tokyo'));
children.push(pBold('2. 中国人用什么吃饭？ What do people in China use to eat?'));
children.push(p('   A. 叉子 Fork     B. 手 Hand     C. 筷子 Chopsticks'));
children.push(pBold('3. 中国的国宝是？ What is China’s national treasure?'));
children.push(p('   A. 老虎 Tiger     B. 熊猫 Panda     C. 龙 Dragon'));
children.push(pageBreak());

// ------------------ PAGE 4 · CHINA (2) ------------------
children.push(countryHeading('🇨🇳', '中国', 'China', COLOR_CHINA));
children.push(shadedBar('在地图中找出中国，并涂色标注。 Find China on the map, color it, and label it.', COLOR_CHINA));
children.push(imgPara(path.join(ASSETS, 'img_9.png'), 420));
children.push(shadedBar('画一画,写一写。Draw & Write.', COLOR_CHINA));
children.push(pBold('我最喜欢的中国食物 My Favorite Chinese Food'));
children.push(drawWriteBox('比如：饺子、面条、包子。 For example: dumplings, noodles, buns.'));
children.push(pageBreak());

// ------------------ PAGE 5 · JAPAN (1) ------------------
children.push(countryHeading('🇯🇵', '日本', 'Japan', COLOR_JAPAN));
children.push(shadedBar('涂一涂: 日本国旗 Color the Flag', COLOR_JAPAN));
children.push(p('⬜ 白色 White + 🔴 红色 Red（中间圆形）'));
children.push(flagBox());

children.push(shadedBar('描一描 Trace: 日本', COLOR_JAPAN));
children.push(traceRow(['日', '本', '日', '本'], COLOR_JAPAN));

children.push(shadedBar('⭕ 圈一圈 Circle the Correct Answer', COLOR_JAPAN));
children.push(pBold('1. 在日本见面应该怎么做？ How do you greet in Japan?'));
children.push(p('   A. 拥抱 Hug     B. 握手 Shake hands     C. 鞠躬 Bow'));
children.push(pBold('2. 进别人家要做什么？ What do you do when entering a house?'));
children.push(p('   A. 脱鞋 Take off shoes     B. 洗手 Wash hands     C. 拍手 Clap'));
children.push(pBold('3. 在日本吃拉面发出声音是？ Slurping ramen in Japan is:'));
children.push(p('   A. 不礼貌 Rude     B. 礼貌 Polite     C. 奇怪 Weird'));
children.push(pageBreak());

// ------------------ PAGE 6 · JAPAN (2) ------------------
children.push(countryHeading('🇯🇵', '日本', 'Japan', COLOR_JAPAN));
children.push(shadedBar('在地图中找出日本，并涂色标注。 Find Japan on the map, color it, and label it.', COLOR_JAPAN));
children.push(imgPara(path.join(ASSETS, 'img_10.png'), 420));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: '🏝️ 日本在亚洲东部，隔海与中国、韩国相望。 Japan sits off East Asia, across the sea from China & Korea. 🇯🇵🇨🇳🇰🇷', size: 20, color: '666666' })],
}));
children.push(shadedBar('画一画,写一写。Draw & Write.', COLOR_JAPAN));
children.push(pBold('我最喜欢的日本食物 My Favorite Japanese Food'));
children.push(drawWriteBox('比如：寿司、拉面、天妇罗。 For example: sushi, ramen, tempura.'));
children.push(pageBreak());

// ------------------ PAGE 7 · INDIA (1) ------------------
children.push(countryHeading('🇮🇳', '印度', 'India', COLOR_INDIA));
children.push(shadedBar('涂一涂: 印度国旗 Color the Flag', COLOR_INDIA));
children.push(p('🟧 橙色 Saffron (上) + ⬜ 白色 White (中，阿育王轮 Ashoka Chakra) + 🟩 绿色 Green (下)'));
children.push(flagBox());

children.push(shadedBar('描一描 Trace: 印度', COLOR_INDIA));
children.push(traceRow(['印', '度', '印', '度'], COLOR_INDIA));

children.push(shadedBar('⭕ 圈一圈 Circle the Correct Answer', COLOR_INDIA));
children.push(pBold('1. 印度的首都是？ What is the capital of India?'));
children.push(p('   A. 孟买 Mumbai     B. 新德里 New Delhi     C. 班加罗尔 Bangalore'));
children.push(pBold('2. 印度人吃饭常用什么？ What do people in India often use to eat?'));
children.push(p('   A. 筷子 Chopsticks     B. 手 Hands     C. 刀叉 Knife & Fork'));
children.push(pBold('3. 印度的国宝动物是？ What is India’s national animal?'));
children.push(p('   A. 大象 Elephant     B. 孟加拉虎 Bengal Tiger     C. 狮子 Lion'));
children.push(pageBreak());

// ------------------ PAGE 8 · INDIA (2) ------------------
children.push(countryHeading('🇮🇳', '印度', 'India', COLOR_INDIA));
children.push(shadedBar('在地图中找出印度，并涂色标注。 Find India on the map, color it, and label it.', COLOR_INDIA));
children.push(imgPara(path.join(ASSETS, 'img_11.png'), 420));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: '🐘 印度位于南亚，北靠喜马拉雅山，与巴基斯坦、尼泊尔、孟加拉国相邻。 India sits in South Asia with Pakistan, Nepal & Bangladesh as neighbors.', size: 20, color: '666666' })],
}));
children.push(shadedBar('画一画,写一写。Draw & Write.', COLOR_INDIA));
children.push(pBold('我最喜欢的印度食物 My Favorite Indian Food'));
children.push(drawWriteBox('比如：咖喱、馕、印度飞饼、比尔亚尼饭。 For example: curry, naan, roti, biryani.'));

// ======================================================================
// BUILD DOC
// ======================================================================

const doc = new Document({
  creator: 'Global Explorer Camp',
  title: '探索亚洲 Explore Asia',
  styles: {
    default: { document: { run: { font: 'Noto Sans SC', size: 22 } } },
  },
  sections: [{
    properties: { page: PAGE },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log(`Wrote ${OUT}  (${(buf.length / 1024 / 1024).toFixed(2)} MB)`);
}).catch(err => {
  console.error('FAILED:', err);
  process.exit(1);
});
