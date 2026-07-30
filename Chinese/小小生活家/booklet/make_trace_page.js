// Generate the 描一描,写一写 (Trace & Write) page image for a 小小生活家 booklet.
// Mirrors the Day 1 kitchen booklet page 4: per word —
//   left: stroke-order progression (one line per character)
//   right: pinyin + 田字格 model row (green trace characters)
//   below: full-width empty 田字格 practice row + ◆造句 line
// Stroke data: hanzi-writer-data (makemeahanzi). Rendered SVG -> PNG via sharp.
// Run: node make_trace_page.js 2     (day number; defaults to 2)
const fs = require('fs');
const path = require('path');
const REPO = path.resolve(__dirname, '../../..');
const sharp = require(path.join(REPO, 'node_modules/sharp'));
const HANZI = path.join(REPO, 'node_modules/hanzi-writer-data');

// 我会写 words per day (pinyin per character)
const DAYS = {
  2: [ // 整理和收纳达人
    { chars: '分类', pinyin: ['fēn', 'lèi'] },
    { chars: '书包', pinyin: ['shū', 'bāo'] },
    { chars: '收拾', pinyin: ['shōu', 'shí'] },
  ],
  4: [ // 友善我先行 · 接人待物 (我会写 朋友 + 描一描 友善 · 帮助)
    { chars: '朋友', pinyin: ['péng', 'you'] },
    { chars: '友善', pinyin: ['yǒu', 'shàn'] },
    { chars: '帮助', pinyin: ['bāng', 'zhù'] },
  ],
};

const DAY = process.argv[2] || '2';
const WORDS = DAYS[DAY];
if (!WORDS) { console.error(`No word list for day ${DAY}. Known: ${Object.keys(DAYS).join(', ')}`); process.exit(1); }

// ---- geometry (200 px per inch; page image = 7.0" x 9.5") ----
const W = 1400, H = 1900;
const GREEN = '#6EE86E', GREEN_LIGHT = '#8DEE8D', BLUE = '#1414FF';
const CELL_LINE = '#1A1A1A', DASH = '#A8A8A8';

const BLOCK_TOP = 30;         // first block y
const BLOCK_H = 600;          // per-word block height
const MODEL_CELL = 155;       // model 田字格 cell size
const MODEL_COLS = 4;
const MODEL_X = W - 40 - MODEL_CELL * MODEL_COLS;
const PRAC_COLS = 8;
const PRAC_CELL = Math.floor((W - 80) / PRAC_COLS);
const GLYPH = 78;             // stroke-progression glyph size

function strokesOf(ch) {
  return require(path.join(HANZI, `${ch}.json`)).strokes;
}

// one glyph: paths drawn inside a size x size box at (x, y)
function glyph(paths, x, y, size, fill) {
  const s = size / 1024;
  return `<g transform="translate(${x},${y}) scale(${s}) scale(1,-1) translate(0,-900)">`
    + paths.map(d => `<path d="${d}" fill="${fill}"/>`).join('')
    + `</g>`;
}

function tianzige(x, y, size) {
  const mx = x + size / 2, my = y + size / 2;
  return `<rect x="${x}" y="${y}" width="${size}" height="${size}" fill="none" stroke="${CELL_LINE}" stroke-width="3"/>`
    + `<line x1="${x}" y1="${my}" x2="${x + size}" y2="${my}" stroke="${DASH}" stroke-width="2.5" stroke-dasharray="14 11"/>`
    + `<line x1="${mx}" y1="${y}" x2="${mx}" y2="${y + size}" stroke="${DASH}" stroke-width="2.5" stroke-dasharray="14 11"/>`;
}

let svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">`
  + `<rect width="${W}" height="${H}" fill="#ffffff"/>`;

WORDS.forEach((word, wi) => {
  const top = BLOCK_TOP + wi * BLOCK_H;
  const chars = [...word.chars];

  // --- stroke-order progression, one line per character ---
  // Long characters (e.g. 善, 12 strokes) are scaled down so the row never
  // runs into the 田字格 column on the right.
  const AVAIL = MODEL_X - 40 - 40;   // left margin .. gap before model cells
  chars.forEach((ch, ci) => {
    const st = strokesOf(ch);
    const need = (st.length - 1) * (GLYPH - 6) + GLYPH;
    const g = need > AVAIL ? GLYPH * AVAIL / need : GLYPH;
    const step = g - 6 * (g / GLYPH);
    const y = top + 6 + ci * (GLYPH + 8);
    st.forEach((_, i) => {
      svg += glyph(st.slice(0, i + 1), 40 + i * step, y, g, GREEN_LIGHT);
    });
  });

  // --- pinyin + model row ---
  chars.forEach((ch, ci) => {
    const cx = MODEL_X + ci * MODEL_CELL + MODEL_CELL / 2;
    svg += `<text x="${cx}" y="${top + 48}" font-family="PingFang SC, Helvetica, Arial" font-size="40"`
      + ` fill="${BLUE}" text-anchor="middle">${word.pinyin[ci]}</text>`;
  });
  const modelY = top + 62;
  for (let c = 0; c < MODEL_COLS; c++) svg += tianzige(MODEL_X + c * MODEL_CELL, modelY, MODEL_CELL);
  chars.forEach((ch, ci) => {
    const pad = MODEL_CELL * 0.09;
    svg += glyph(strokesOf(ch), MODEL_X + ci * MODEL_CELL + pad, modelY + pad, MODEL_CELL - 2 * pad, GREEN);
  });

  // --- full-width practice row ---
  const pracY = modelY + MODEL_CELL + 70;
  for (let c = 0; c < PRAC_COLS; c++) svg += tianzige(40 + c * PRAC_CELL, pracY, PRAC_CELL);

  // --- 造句 line ---
  const zjY = pracY + PRAC_CELL + 62;
  svg += `<text x="52" y="${zjY}" font-family="Kaiti SC, STKaiti, Songti SC" font-size="42" fill="#1A1A1A">◆ 造句：</text>`;
  svg += `<line x1="230" y1="${zjY + 10}" x2="${W - 40}" y2="${zjY + 10}" stroke="#BBBBBB" stroke-width="2"/>`;

  // --- separator ---
  if (wi < WORDS.length - 1) {
    const sy = top + BLOCK_H - 34;
    svg += `<line x1="20" y1="${sy}" x2="${W - 20}" y2="${sy}" stroke="#555555" stroke-width="2"/>`;
  }
});

svg += `</svg>`;

const OUT = path.join(__dirname, 'assets', `day${DAY}_trace_page.png`);
sharp(Buffer.from(svg)).png().toFile(OUT).then(i => console.log(`Created ${OUT} (${i.width}x${i.height})`));
