# Zero Waste Day 3 — 海洋小卫士 Deck Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a fully-editable ~54-slide Day 3 (水+塑料) teaching deck following the wilderness `day1_nature` reference's 看→想→判→做 interactive structure, re-themed to 海洋小卫士 / ocean palette.

**Architecture:** New builder `create_day3_ocean.py` on top of the existing `_helpers.py` design system (same system the reference uses). Add 6 new reference-matching helpers to `_helpers.py`. Content is data-driven (a single `ITEMS` array of 6 plastic items) so the loop and trimming are trivial. Render via python-pptx → QA via LibreOffice→images + subagent visual inspection.

**Tech Stack:** Python 3, `python-pptx`, `_helpers.py`, LibreOffice (`scripts/office/soffice.py`), `pdftoppm`, `markitdown`.

**Spec:** `docs/superpowers/specs/2026-06-30-zero-waste-day3-ocean-guardians-full-deck-design.md`

---

## Working directory & conventions
- Build dir: `Chinese/zero_waste零废弃/`
- Slides are `10 × 5.625 in` (16:9). Font KaiTi. Every slide gets `pn(s, n)` page number.
- Verify method for a build step = **render + inspect**, not unit tests (this is a deck).
  Smoke check after each task: `python create_day3_ocean.py && python -m markitdown PPT/day3_water_plastic.pptx | head -N`.

---

## Task 1: Scaffold builder + ocean palette + smoke render

**Files:**
- Create: `Chinese/zero_waste零废弃/create_day3_ocean.py`

- [ ] **Step 1: Create builder skeleton**

```python
# create_day3_ocean.py — Zero Waste Day 3 · 海洋小卫士 (Water + Plastic)
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import _helpers as H
from _helpers import (make_presentation, ns, tb, bg, hb, pn, panel, panel_head, div,
                      cover, learning_goals, photo_slot, teacher_box, sentence_frame_bar,
                      compare_slide, video_slide, answer_panels_slide,
                      vocab_recognize, vocab_write, share_close,
                      CREAM, WHITE, DARK, GRAY, LGRAY, INK, STAR, WARM, IMGBG,
                      OK, ALERT)

# ----- Ocean palette -----
def _rgb(h): return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))
OCEAN = _rgb("0B5C8C")   # primary
TEAL  = _rgb("006970")
AQUA  = _rgb("2A9D8F")
SKY   = _rgb("42A5F5")
CORAL = ALERT            # danger / reveal headers

# 6 per-item accent tints
TINTS = [_rgb("1B6FBA"), _rgb("006970"), _rgb("2A9D8F"),
         _rgb("3A7CA5"), _rgb("E6853A"), _rgb("C8253E")]

def build():
    prs = make_presentation()
    n = [0]
    def page(s):
        n[0] += 1; pn(s, n[0]); return s
    # --- slides appended here in later tasks ---
    prs.save("PPT/day3_water_plastic.pptx")
    print(f"Saved {n[0]} slides")

if __name__ == "__main__":
    build()
```

- [ ] **Step 2: Add cover + Session 1 divider as a smoke test**

Inside `build()` before `prs.save`:

```python
    page(cover(prs, 3, "水 与 塑 料 · 海 洋 小 卫 士",
               "Water & Plastic — Little Ocean Guardians",
               "💧  🌊  🐢  🥤  🛍️  ♻️", OCEAN,
               "我 们 怎 样 节 约 用 水、减 少 塑 料，保 护 海 洋?",
               "How can we save water and cut plastic to protect the ocean?"))
    page(div(prs, "Session 1 · 上 午", "水 很 重 要 + 塑 料 污 染", OCEAN, "🌊"))
```

- [ ] **Step 3: Render smoke test**

Run: `cd Chinese/zero_waste零废弃 && python create_day3_ocean.py`
Expected: `Saved 2 slides` and `PPT/day3_water_plastic.pptx` overwritten.

---

## Task 2: Add 6 new helpers to `_helpers.py`

**Files:**
- Modify: `Chinese/zero_waste零废弃/_helpers.py` (append at end)

Each mirrors a reference slide. Geometry follows existing helpers (10×5.625, `hb` header at top, CREAM bg). Add all six, then smoke-render one of each from the builder.

- [ ] **Step 1: `mission_intro_slide`** (reference slide 5 "You're a Little Explorer!")

```python
def mission_intro_slide(prs, color, header, cards, frame_cn, frame_en):
    """6 preview cards (emoji + label) + intro banner + sentence frame.
    cards: list of (emoji, cn, en) — up to 6."""
    s = ns(prs); bg(s, CREAM); hb(s, header, color)
    band = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(0.90), Inches(9.20), Inches(0.85))
    band.fill.solid(); band.fill.fore_color.rgb = color; band.line.color.rgb = STAR; band.line.width = Pt(2.5)
    tb(s, 0.5, 0.98, 9.0, 0.40, "🌊 今 天 你 是 海 洋 小 卫 士!", sz=17, b=True, c=STAR, a=PP_ALIGN.CENTER)
    tb(s, 0.5, 1.38, 9.0, 0.30, "Today you're a Little Ocean Guardian — save water, cut plastic!", sz=10, c=WHITE, a=PP_ALIGN.CENTER)
    cw, gap = 1.46, 0.10
    total = 6*cw + 5*gap; x0 = (10 - total)/2
    for i, (em, cn, en) in enumerate(cards[:6]):
        x = x0 + i*(cw+gap)
        panel(s, x, 2.05, cw, 1.85, TINTS_FALLBACK(i, color), fill=WHITE, lw=2.5)
        tb(s, x, 2.20, cw, 0.75, em, sz=40, a=PP_ALIGN.CENTER)
        tb(s, x, 3.00, cw, 0.35, cn, sz=13, b=True, c=DARK, a=PP_ALIGN.CENTER)
        tb(s, x, 3.38, cw, 0.28, en, sz=8.5, c=GRAY, a=PP_ALIGN.CENTER)
        tb(s, x, 3.62, cw, 0.24, f"#{i+1}", sz=10, b=True, c=GRAY, a=PP_ALIGN.CENTER)
    sentence_frame_bar(s, 4.55, frame_cn, frame_en, color)
    return s

def TINTS_FALLBACK(i, color):
    return color
```

- [ ] **Step 2: `observe_think_slide`** (reference slide 10: photo + 看/听/感觉 + danger cards + media line)

```python
def observe_think_slide(prs, color, header, photo_cn, photo_en, senses, dangers,
                        media_cn=None, media_url=None):
    """看+想 station slide.
    senses: list of (emoji, cn) up to 3.  dangers: list of (emoji, cn) up to 4."""
    s = ns(prs); bg(s, CREAM); hb(s, header, color)
    photo_slot(s, 0.30, 0.95, 4.35, 3.30, photo_cn, photo_en, color)
    panel(s, 4.80, 0.95, 4.90, 3.30, color, fill=WHITE, lw=2.5)
    panel_head(s, 4.80, 0.95, 4.90, color, "1️⃣ 👀 看 — 你 看 到 / 想 到 什 么?", sz=12)
    for i, (em, cn) in enumerate(senses[:3]):
        y = 1.60 + i*0.80
        tb(s, 4.95, y, 0.55, 0.55, em, sz=26, a=PP_ALIGN.LEFT)
        tb(s, 5.55, y+0.06, 3.95, 0.55, cn, sz=13, b=True, c=DARK, a=PP_ALIGN.LEFT)
    # bottom danger row
    dh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.30), Inches(4.38), Inches(9.40), Inches(0.42))
    dh.fill.solid(); dh.fill.fore_color.rgb = CORAL; dh.line.fill.background()
    tb(s, 0.4, 4.44, 9.2, 0.32, "2️⃣ 🤔 想 — 它 到 了 海 里, 会 伤 害 谁?  My idea: 我 觉 得……",
       sz=12, b=True, c=STAR, a=PP_ALIGN.CENTER)
    dw, dgap = 2.28, 0.10; dx0 = (10 - (4*dw+3*dgap))/2
    for i, (em, cn) in enumerate(dangers[:4]):
        x = dx0 + i*(dw+dgap)
        panel(s, x, 4.86, dw, 0.62, color, fill=WHITE, lw=2)
        tb(s, x+0.10, 4.96, 0.5, 0.42, em, sz=20, a=PP_ALIGN.LEFT)
        tb(s, x+0.62, 4.98, dw-0.7, 0.40, cn, sz=12, b=True, c=DARK, a=PP_ALIGN.LEFT)
    if media_cn and media_url:
        link = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.80), Inches(3.55), Inches(4.90), Inches(0.65))
        link.fill.solid(); link.fill.fore_color.rgb = INK; link.line.color.rgb = STAR; link.line.width = Pt(1.5)
        tb(s, 4.90, 3.60, 4.7, 0.28, f"▶️ {media_cn}", sz=11, b=True, c=STAR, a=PP_ALIGN.CENTER)
        tb(s, 4.90, 3.88, 4.7, 0.26, media_url, sz=8, c=LGRAY, a=PP_ALIGN.CENTER)
        link.click_action.hyperlink.address = media_url
    return s
```

- [ ] **Step 3: `judge_ab_slide` + `reveal_ab_slide`** (reference slides 11 & 12: 3 rows A/B, then green ✓ + 做 action)

```python
def _ab_rows(s, color, rows, reveal=False):
    """3 scenario rows; each row: number badge + scenario + A box (left) + B box (right).
    rows: list of (scenario_cn, a_cn, b_cn, correct)  correct in {'A','B'}."""
    top = 1.75; rh = 0.86
    for i, (sc, a_cn, b_cn, correct) in enumerate(rows[:3]):
        y = top + i*rh
        panel(s, 0.35, y, 9.30, rh-0.12, color, fill=WHITE, lw=2)
        bd = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.45), Inches(y+0.10), Inches(0.42), Inches(0.42))
        bd.fill.solid(); bd.fill.fore_color.rgb = color; bd.line.fill.background()
        tb(s, 0.45, y+0.12, 0.42, 0.38, str(i+1), sz=15, b=True, c=WHITE, a=PP_ALIGN.CENTER)
        tb(s, 1.00, y+0.06, 3.05, 0.62, sc, sz=13, b=True, c=DARK, a=PP_ALIGN.LEFT)
        for label, txt, bx in (("A", a_cn, 4.15), ("B", b_cn, 6.95)):
            good = reveal and (correct == label)
            oc = OK if good else color
            ob = panel(s, bx, y+0.12, 2.60, rh-0.36, oc, fill=(RGBColor(0xEF,0xF7,0xEE) if good else WHITE), lw=(3 if good else 1.5))
            tb(s, bx+0.10, y+0.20, 0.45, 0.35, ("✓ "+label if good else label), sz=12, b=True, c=oc, a=PP_ALIGN.LEFT)
            tb(s, bx+0.62, y+0.20, 1.90, 0.35, txt, sz=12, b=True, c=DARK, a=PP_ALIGN.LEFT)

def judge_ab_slide(prs, color, header, rows, frame_cn, frame_en):
    s = ns(prs); bg(s, CREAM); hb(s, header, color)
    tb(s, 0.4, 0.80, 9.2, 0.32, "⚖️ 判 — 你 觉 得 呢? A 还 是 B?  (先 选, 再 看!)",
       sz=13, b=True, c=DARK, a=PP_ALIGN.CENTER)
    _ab_rows(s, color, rows, reveal=False)
    vb = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(4.40), Inches(9.20), Inches(0.40))
    vb.fill.solid(); vb.fill.fore_color.rgb = color; vb.line.fill.background()
    tb(s, 0.5, 4.45, 9.0, 0.30, "👉 A → 举 左 手 · B → 举 右 手   🤔 3… 2… 1…",
       sz=12, b=True, c=STAR, a=PP_ALIGN.CENTER)
    sentence_frame_bar(s, 4.90, frame_cn, frame_en, color)
    return s

def reveal_ab_slide(prs, color, header, rows, do_action, frame_cn, frame_en):
    s = ns(prs); bg(s, CREAM); hb(s, header, CORAL)
    tb(s, 0.4, 0.80, 9.2, 0.32, "💡 答 案 揭 晓!  绿 色 ✓ = 保 护 海 洋 的 选 择",
       sz=13, b=True, c=DARK, a=PP_ALIGN.CENTER)
    _ab_rows(s, color, rows, reveal=True)
    ab = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(4.40), Inches(9.20), Inches(0.40))
    ab.fill.solid(); ab.fill.fore_color.rgb = OK; ab.line.fill.background()
    tb(s, 0.5, 4.45, 9.0, 0.30, f"🙆 做: {do_action}", sz=12, b=True, c=STAR, a=PP_ALIGN.CENTER)
    sentence_frame_bar(s, 4.90, frame_cn, frame_en, color)
    return s
```

- [ ] **Step 4: `cardgrid_slide`** (reference slide 9: 6-card overview, 3×2)

```python
def cardgrid_slide(prs, color, header, subtitle, items):
    """items: list of (emoji, cn, en) — renders 3 cols × 2 rows."""
    s = ns(prs); bg(s, CREAM); hb(s, header, color)
    tb(s, 0.4, 0.80, 9.2, 0.32, subtitle, sz=13, b=True, c=GRAY, a=PP_ALIGN.CENTER)
    cw, ch, gx, gy = 2.95, 1.75, 0.20, 0.18
    x0 = (10 - (3*cw + 2*gx))/2; y0 = 1.25
    for i, (em, cn, en) in enumerate(items[:6]):
        r, c = divmod(i, 3)
        x = x0 + c*(cw+gx); y = y0 + r*(ch+gy)
        panel(s, x, y, cw, ch, TINTS[i % len(TINTS)], fill=WHITE, lw=2.5)
        tb(s, x, y+0.20, cw, 0.70, em, sz=44, a=PP_ALIGN.CENTER)
        tb(s, x, y+1.00, cw, 0.40, cn, sz=20, b=True, c=TINTS[i % len(TINTS)], a=PP_ALIGN.CENTER)
        tb(s, x, y+1.42, cw, 0.26, en, sz=10, c=GRAY, a=PP_ALIGN.CENTER)
    return s
```

- [ ] **Step 5: `guess_slide`** (reference slide 31: teacher acts a swap, students guess)

```python
def guess_slide(prs, color, header, subtitle, cards, frame_cn, frame_en):
    """cards: list of (emoji, prompt_cn) up to 6 — the mimes/clues."""
    s = ns(prs); bg(s, CREAM); hb(s, header, color)
    tb(s, 0.4, 0.80, 9.2, 0.30, subtitle, sz=12, b=True, c=GRAY, a=PP_ALIGN.CENTER)
    cw, ch, gx, gy = 2.95, 1.35, 0.20, 0.18
    x0 = (10 - (3*cw + 2*gx))/2; y0 = 1.25
    for i, (em, cn) in enumerate(cards[:6]):
        r, c = divmod(i, 3)
        x = x0 + c*(cw+gx); y = y0 + r*(ch+gy)
        panel(s, x, y, cw, ch, TINTS[i % len(TINTS)], fill=WHITE, lw=2)
        tb(s, x, y+0.14, cw, 0.55, em, sz=30, a=PP_ALIGN.CENTER)
        tb(s, x+0.12, y+0.72, cw-0.24, 0.55, cn, sz=12, b=True, c=DARK, a=PP_ALIGN.CENTER)
    sentence_frame_bar(s, 4.62, frame_cn, frame_en, color)
    return s
```

Note: move `TINTS`/`TINTS_FALLBACK` usage cleanly — define `TINTS` in `_helpers.py` too (ocean tints) so grid helpers there can use it:

```python
# add near palette in _helpers.py
OCEAN_TINTS = [RGBColor(0x1B,0x6F,0xBA), RGBColor(0x00,0x69,0x70), RGBColor(0x2A,0x9D,0x8F),
               RGBColor(0x3A,0x7C,0xA5), RGBColor(0xE6,0x85,0x3A), RGBColor(0xC8,0x25,0x3E)]
```
Then in helpers reference `OCEAN_TINTS` instead of `TINTS`; builder keeps its own `TINTS` for local use. (Replace the `TINTS`/`TINTS_FALLBACK` placeholders above with `OCEAN_TINTS` and `color` respectively.)

- [ ] **Step 6: Smoke-render each new helper** — temporarily call each once in `build()`, run, confirm no exception, then remove the temp calls.

Run: `cd Chinese/zero_waste零废弃 && python create_day3_ocean.py`
Expected: `Saved N slides`, no traceback.

---

## Task 3: Content data — the `ITEMS` array + vocab + projects

**Files:**
- Modify: `create_day3_ocean.py` (module-level data)

- [ ] **Step 1: Define the 6-item spine** (module level, before `build`)

```python
# (item_emoji, cn, en, tint_idx, victim_emoji, victim_cn,
#  senses[(em,cn)x3], dangers[(em,cn)x4],
#  rows[(scenario,a,b,correct)x3], do_action, swap_cn)
ITEMS = [
  dict(em="🛍️", cn="塑 料 袋", en="Plastic Bag", tint=0, victim="🐢",
       senses=[("👀","薄薄的, 会飞"),("🌊","在水里像水母"),("🗑️","用一次就扔?")],
       dangers=[("🐢","海龟?"),("🐟","鱼?"),("🐦","海鸟?"),("🌊","漂到海里?")],
       rows=[("去买东西","要塑料袋","带布袋","B"),
             ("袋子破了","再拿一个","重复用/补一下","B"),
             ("买菜","很多小袋","一个大布袋","B")],
       do="假装把布袋挂在手上, 去买东西!", swap="🛍️ 换成: 布 袋 / 可 重 复 袋"),
  dict(em="🥤", cn="吸 管", en="Straw", tint=1, victim="🐢",
       senses=[("👀","细细的, 一次性"),("🌊","会掉进海里"),("🐢","卡在海龟鼻子里?")],
       dangers=[("🐢","海龟?"),("🐬","海豚?"),("🐟","鱼?"),("🗑️","变垃圾?")],
       rows=[("喝果汁","用塑料吸管","不用/直接喝","B"),
             ("店里给吸管","收下","说不用, 谢谢","B"),
             ("想要吸管","塑料的","纸的/钢的","B")],
       do="做喝水动作 — 不用吸管, 直接喝!", swap="🚫 不 用 · 钢 吸 管 / 纸 吸 管"),
  dict(em="🧴", cn="塑 料 瓶", en="Plastic Bottle", tint=2, victim="🐦",
       senses=[("👀","透明, 很多"),("🌊","会碎成小塑料"),("🗑️","喝完就扔?")],
       dangers=[("🐦","海鸟?"),("🐟","鱼?"),("🐳","鲸鱼?"),("🌊","海洋垃圾?")],
       rows=[("口渴了","买瓶装水","带水壶","B"),
             ("水喝完了","再买一瓶","装满水壶","B"),
             ("出门","买饮料瓶","自己的水壶","B")],
       do="举起水壶, 咕嘟咕嘟喝水!", swap="💧 换成: 水 壶 (装满再用)"),
  dict(em="🍴", cn="一 次 性 餐 具", en="Disposable Utensils", tint=3, victim="🐟",
       senses=[("👀","塑料叉/勺"),("🌊","会断成小片"),("🗑️","吃完就扔?")],
       dangers=[("🐟","鱼?"),("🦀","螃蟹?"),("🐢","海龟?"),("🗑️","变垃圾?")],
       rows=[("吃饭","一次性叉子","自带勺子","B"),
             ("外卖给餐具","都收下","说不用","B"),
             ("野餐","塑料餐具","自己的餐具","B")],
       do="假装从书包拿出自己的勺子吃饭!", swap="🍴 换成: 自 带 餐 具"),
  dict(em="📦", cn="塑 料 包 装", en="Plastic Packaging", tint=4, victim="🦭",
       senses=[("👀","一层层包装"),("🌊","会缠住动物"),("🗑️","拆完就扔?")],
       dangers=[("🦭","海豹?"),("🐢","海龟?"),("🐟","鱼?"),("🌊","漂到海里?")],
       rows=[("买零食","很多小包装","大包装/散装","B"),
             ("水果","塑料盒装","散装自己装","B"),
             ("包装绳","套着扔","剪开再扔","B")],
       do="用手比一比 — 剪开包装, 再扔掉!", swap="📦 换成: 少 包 装 / 散 装"),
  dict(em="🥤", cn="一 次 性 杯 子", en="Disposable Cup", tint=5, victim="🌊",
       senses=[("👀","纸/塑料杯+盖"),("🌊","会变海洋垃圾"),("🗑️","喝完就扔?")],
       dangers=[("🌊","海洋垃圾?"),("🐦","海鸟?"),("🐟","鱼?"),("🐢","海龟?")],
       rows=[("买饮料","一次性杯","自己的杯子","B"),
             ("喝水","拿纸杯","用自己的杯","B"),
             ("咖啡店","要一次性杯","带随行杯","B")],
       do="举起自己的杯子, 干杯!", swap="🥤 换成: 自 带 杯 子 / 随 行 杯"),
]

VOCAB_RECOGNIZE = [
  ("💧","水","shuǐ","water","水 是 生 命 之 源, 我 们 每 天 都 要 喝 水。","Water is the source of life.","💧 一杯干净的水"),
  ("🌊","海 洋","hǎi yáng","ocean","海 洋 里 有 很 多 动 物。","The ocean has many animals.","🌊 蓝色的大海"),
  ("🥤","塑 料","sù liào","plastic","塑 料 用 一 次 就 扔, 很 浪 费。","Single-use plastic is wasteful.","🥤 塑料瓶和塑料袋"),
  ("🛢️","污 染","wū rǎn","pollution","塑 料 让 海 洋 污 染。","Plastic pollutes the ocean.","🛢️ 海面上的塑料垃圾"),
  ("🛡️","保 护","bǎo hù","protect","我 们 要 保 护 海 洋。","We protect the ocean.","🛡️ 孩子在捡垃圾"),
  ("🚰","节 约","jié yuē","save","节 约 用 水, 关 好 水 龙 头。","Save water — turn off the tap.","🚰 关水龙头的手"),
]

# vocab_write chars: (char, pinyin, "N strokes", mnemonic)
VOCAB_WRITE = [
  ("水", [("水","shuǐ","4 笔","三点+一竖钩")], "shuǐ", "water"),
  ("保 护", [("保","bǎo","9 笔","人+口+木"), ("护","hù","7 笔","扌+户")], "bǎo hù", "protect"),
  ("塑 料", [("塑","sù","13 笔","朔+土"), ("料","liào","10 笔","米+斗")], "sù liào", "plastic"),
]
```

- [ ] **Step 2: Smoke import** — `python -c "import create_day3_ocean"` → no error.

---

## Task 4: Assemble Session 1 (slides 1–34)

**Files:** `create_day3_ocean.py` (`build()` body)

- [ ] **Step 1: Open (1–3)** — cover (from Task 1), schedule, objectives.

```python
    # 2. Schedule — reuse a simple 3-bar layout via panels
    s = page(ns(prs)); bg(s, CREAM); hb(s, "⏰ 今 日 时 间 安 排  Today's Schedule", OCEAN)
    for i,(lab,tm,desc,cl) in enumerate([
        ("Session 1 · 上午","10:00–11:45","水很重要 + 塑料污染 (故事/视频)", OCEAN),
        ("Session 2 · 下午","2:00–2:45","复习 + 语言 (我会认 6 / 我会写 3)", TEAL),
        ("Session 3 · 下午","3:00–4:30","练习册 + 动手 (海洋拼贴 / 承诺)", AQUA)]):
        y = 1.05 + i*1.42
        p = panel(s, 0.5, y, 9.0, 1.25, cl, fill=cl, lw=0); 
        tb(s, 0.8, y+0.20, 5.5, 0.5, lab, sz=22, b=True, c=WHITE)
        tb(s, 0.8, y+0.75, 5.5, 0.35, tm, sz=13, c=STAR)
        tb(s, 5.6, y+0.42, 3.6, 0.7, desc, sz=13, b=True, c=WHITE)
    # 3. Objectives
    page(learning_goals(prs, OCEAN, [
        ("1","理 解 水 资 源 的 重 要 性","Understand why water matters", OCEAN),
        ("2","认 识 塑 料 对 海 洋 动 物 的 影 响","Plastic harms ocean animals", TEAL),
        ("3","理 解 一 次 性 塑 料 的 问 题","The problem with single-use plastic", AQUA),
        ("4","提 出 减 少 塑 料 的 方 法","Propose ways to reduce plastic", CORAL)]))
```

- [ ] **Step 2: Session 1 opener (4–11)** — divider (Task 1), mission intro, story, discuss, water×2, plastic video, single-use framing.

```python
    page(div(prs, "Session 1 · 上 午", "水 很 重 要 + 塑 料 污 染", OCEAN, "🌊"))
    page(H.mission_intro_slide(prs, OCEAN, "🧭 你 是 海 洋 小 卫 士!  Little Ocean Guardian!",
        [(it["em"], it["cn"], it["en"]) for it in ITEMS],
        "我 是 海 洋 小 卫 士, 我 要 保 护 ______。", "I'm an Ocean Guardian. I'll protect ___."))
    page(video_slide(prs, "先 来 听 个 故 事: 一 滴 水 的 旅 行", "Story: A Water Drop's Journey",
        [("👂","水 从 哪 里 来?","Where is water from?"),
         ("👀","水 去 了 哪 里?","Where does it go?"),
         ("🤔","我 们 用 水 做 什 么?","What do we use water for?")],
        "🌊 看 完 说 一 说: 我 们 为 什 么 需 要 水?",
        "https://www.youtube.com/results?search_query=水循环+儿童+动画", OCEAN))
    # 故事讨论
    s = page(ns(prs)); bg(s, CREAM); hb(s, "💬 故 事 讨 论  Let's Talk", OCEAN)
    for i,(q,en) in enumerate([("水 去 了 哪 里? (云→雨→河→海)","Where did the water go?"),
                               ("我 们 为 什 么 需 要 水?","Why do we need water?")]):
        y=1.0+i*1.0; panel(s,0.5,y,9.0,0.85,OCEAN,fill=WHITE,lw=2.5)
        tb(s,0.7,y+0.12,8.6,0.4,f"{i+1}. {q}",sz=15,b=True,c=DARK)
        tb(s,0.7,y+0.52,8.6,0.3,en,sz=10,c=GRAY)
    sentence_frame_bar(s,3.2,"我 们 需 要 水 来 ______。","We need water to ___.",OCEAN)
    teacher_box(s,0.5,4.05,9.0,"引 导 说 出: 喝、洗、种 植 物、动 物 也 要 水","学 生 抢 答",4,OCEAN)
    # 水很重要
    s = page(ns(prs)); bg(s,CREAM); hb(s,"💧 水 很 重 要  Water Matters", OCEAN)
    tb(s,0.4,0.82,9.2,0.3,"我 们 每 天 都 要 用 水!  We use water every day!",sz=13,b=True,c=GRAY,a=PP_ALIGN.CENTER)
    for i,(em,cn,en) in enumerate([("🥤","喝 水","drink"),("🚿","洗 澡","wash"),
        ("🌱","种 植 物","plants"),("🐟","动 物 也 要 水","animals")]):
        r,c=divmod(i,2); x=0.9+c*4.3; y=1.3+r*1.7
        panel(s,x,y,3.9,1.5,OCEAN,fill=WHITE,lw=2.5)
        tb(s,x,y+0.18,3.9,0.7,em,sz=44,a=PP_ALIGN.CENTER)
        tb(s,x,y+0.95,3.9,0.4,f"{cn}",sz=18,b=True,c=OCEAN,a=PP_ALIGN.CENTER)
    # 珍贵 + 节约
    s = page(ns(prs)); bg(s,CREAM); hb(s,"🌍 地 球 的 水 很 珍 贵 · 节 约 用 水  Save Water", TEAL)
    panel(s,0.4,1.0,4.5,3.4,TEAL,fill=WARM,lw=2.5)
    tb(s,0.5,1.4,4.3,1.2,"🌏",sz=90,a=PP_ALIGN.CENTER)
    tb(s,0.5,3.0,4.3,0.5,"能 喝 的 淡 水 很 少!",sz=17,b=True,c=TEAL,a=PP_ALIGN.CENTER)
    tb(s,0.5,3.6,4.3,0.5,"Only a little water is drinkable.",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    panel(s,5.1,1.0,4.5,3.4,TEAL,fill=WHITE,lw=2.5); panel_head(s,5.1,1.0,4.5,TEAL,"✅ 节 约 用 水 小 妙 招",sz=13)
    for i,(em,cn) in enumerate([("🚰","关 好 水 龙 头"),("🪥","刷 牙 时 关 水"),
        ("♻️","洗 菜 水 浇 花"),("🚿","洗 澡 快 一 点")]):
        y=1.65+i*0.62; tb(s,5.25,y,0.5,0.5,em,sz=22); tb(s,5.85,y+0.06,3.6,0.4,cn,sz=13,b=True,c=DARK)
    page(s)  # already appended via ns; ensure pn — see note
    # 塑料 video
    page(video_slide(prs,"海 洋 告 急! 塑 料 去 哪 了?","Ocean SOS — Where does plastic go?",
        [("🌊","塑 料 去 了 海 里","Plastic reaches the sea"),
         ("🐢","海 龟 会 怎 样?","What happens to turtles?"),
         ("😢","你 有 什 么 感 觉?","How do you feel?")],
        "看 完 说: 塑 料 让 ______ 受 伤。",
        "https://www.youtube.com/results?search_query=ocean+plastic+sea+turtle+kids", CORAL))
    # 一次性塑料
    s=page(ns(prs)); bg(s,CREAM); hb(s,"🥤 什 么 是 一 次 性 塑 料?  Single-Use Plastic", OCEAN)
    tb(s,0.4,0.9,9.2,0.5,"用 一 次 就 扔 掉 的 塑 料 = 一 次 性 塑 料",sz=20,b=True,c=CORAL,a=PP_ALIGN.CENTER)
    tb(s,0.4,1.5,9.2,0.35,"Used once, then thrown away.",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
    for i,(em,cn) in enumerate(ITEMS[i2] for i2,_ in enumerate(ITEMS)):
        pass
    for i,it in enumerate(ITEMS):
        r,c=divmod(i,3); x=0.7+c*3.0; y=2.1+r*1.5
        panel(s,x,y,2.8,1.35,TINTS[it["tint"]],fill=WHITE,lw=2)
        tb(s,x,y+0.15,2.8,0.6,it["em"],sz=34,a=PP_ALIGN.CENTER)
        tb(s,x,y+0.85,2.8,0.4,it["cn"],sz=15,b=True,c=TINTS[it["tint"]],a=PP_ALIGN.CENTER)
```

> Note on `page()` + existing-helper slides: helpers that call `ns(prs)` internally already
> created the slide; `page(s)` just stamps the number on the returned slide. For slides you
> build inline starting with `s = ns(prs)`, call `page(s)` once at the end (or wrap creation).
> Keep a single rule: **every slide object passes through `page()` exactly once.** Fix the two
> inline spots above (`page(s)` right after finishing each inline slide).

- [ ] **Step 3: 6-item overview grid (12)**

```python
    page(H.cardgrid_slide(prs, OCEAN, "🗺️ 今 天 的 6 种 一 次 性 塑 料  Today's 6 Single-Use Plastics",
        "先 认 一 认, 再 一 个 一 个 来!",
        [(it["em"], it["cn"], it["en"]) for it in ITEMS]))
```

- [ ] **Step 4: The 6 stations loop (13–30)**

```python
    for it in ITEMS:
        col = TINTS[it["tint"]]
        hdr = f'{it["em"]} {it["cn"]} · {it["en"]}'
        media = (("看 1 分 钟 视 频", "https://www.youtube.com/results?search_query=ocean+plastic+"+it["en"].replace(" ","+"))
                 if it["cn"].strip() == "吸 管" else (None, None))
        page(H.observe_think_slide(prs, col, f"{hdr}  ①看 +②想", it["cn"]+" 的 真 实 照 片",
            f'Real photo of {it["en"]}', it["senses"], it["dangers"], media[0], media[1]))
        page(H.judge_ab_slide(prs, col, f"{it['em']} {it['cn']} · ⚖️ 你 觉 得 呢?", it["rows"],
            "我 选 A / B。 (G2-G3: 我 选 __, 因 为 __。)", "I choose A/B. (because…)"))
        page(H.reveal_ab_slide(prs, col, f"{it['em']} {it['cn']} · 💡 答 案 揭 晓!", it["rows"],
            it["do"] + "   " + it["swap"],
            "____ 不 好, 我 应 该 ____。", "___ is bad; I should ___.", ))
```

- [ ] **Step 5: compare + action game + guess (31–34)**

```python
    # 31 compare (item → 动物 → 换成)
    s=page(ns(prs)); bg(s,CREAM); hb(s,"📋 6 种 塑 料 对 比  Compare 6 Plastics", OCEAN)
    tb(s,0.4,0.80,9.2,0.3,"每 种 塑 料 伤 害 谁? 换 成 什 么?",sz=12,b=True,c=GRAY,a=PP_ALIGN.CENTER)
    colw=1.52; x0=0.55
    heads=["","🐢 伤 害","♻️ 换 成"]
    # header row
    for i,it in enumerate(ITEMS):
        x=x0+ i*colw + 0.0
    # simple table via cells
    rows_data=[("塑 料", [it["em"] for it in ITEMS]),
               ("动 物", [it["victim"] for it in ITEMS]),
               ("换 成", ["🛍️","🚫","💧","🍴","📦","🥤"])]
    ty=1.3
    for ri,(lab,vals) in enumerate(rows_data):
        y=ty+ri*1.05
        panel(s,0.35,y,1.15,0.95,OCEAN,fill=OCEAN,lw=0); tb(s,0.35,y+0.30,1.15,0.4,lab,sz=13,b=True,c=WHITE,a=PP_ALIGN.CENTER)
        for ci,v in enumerate(vals):
            x=1.60+ci*1.33; panel(s,x,y,1.25,0.95,OCEAN,fill=WHITE,lw=1.5)
            tb(s,x,y+0.22,1.25,0.6,v,sz=26,a=PP_ALIGN.CENTER)
    # 32-33 action game think + reveal
    game=[("🛍️","去买东西不用塑料袋"),("🥤","喝水用自己的水壶"),("🐢","塑料扔进海里"),
          ("💧","刷牙一直开着水"),("🍴","自带勺子吃饭"),("🚰","洗菜水浇花")]
    ans=["✅","✅","❌","❌","✅","✅"]
    s=page(ns(prs)); bg(s,CREAM); hb(s,"🛡️ 好 vs 不 好!  🤔 你 觉 得 呢?", OCEAN)
    tb(s,0.4,0.85,9.2,0.4,"老 师 说 一 件 事 — 学 生 做 动 作!  ✅好=举 手 · ❌不 好=交 叉",sz=13,b=True,c=DARK,a=PP_ALIGN.CENTER)
    for i,(em,cn) in enumerate(game):
        r,c=divmod(i,2); x=0.5+c*4.6; y=1.5+r*1.15
        panel(s,x,y,4.3,1.0,OCEAN,fill=WHITE,lw=2); tb(s,x+0.1,y+0.28,0.6,0.5,em,sz=26)
        tb(s,x+0.75,y+0.30,3.0,0.5,cn,sz=13,b=True,c=DARK); tb(s,x+3.7,y+0.25,0.5,0.5,"?",sz=24,b=True,c=OCEAN)
    s=page(ns(prs)); bg(s,CREAM); hb(s,"🛡️ 好 vs 不 好!  💡 答 案 揭 晓", CORAL)
    for i,(em,cn) in enumerate(game):
        r,c=divmod(i,2); x=0.5+c*4.6; y=1.3+r*1.15; mk=ans[i]; mc=OK if mk=="✅" else ALERT
        panel(s,x,y,4.3,1.0,mc,fill=WHITE,lw=2.5); tb(s,x+0.1,y+0.28,0.6,0.5,em,sz=26)
        tb(s,x+0.75,y+0.30,3.0,0.5,cn,sz=13,b=True,c=DARK); tb(s,x+3.65,y+0.24,0.6,0.55,mk,sz=26,b=True,c=mc)
    sentence_frame_bar(s,5.02,"____ 不 好, 因 为 ____。","___ is bad, because ___.",OCEAN)
    # 34 guess
    page(H.guess_slide(prs, OCEAN, "🔍 我 演 你 猜  Act & Guess!",
        "老 师 演 一 个 动 作 — 学 生 猜 是 哪 个 塑 料 / 替 代 品!",
        [("🛍️","挂布袋去买东西"),("🥤","咕嘟咕嘟喝水壶"),("🍴","从书包拿勺子"),
         ("🚫","摆手说不用吸管"),("📦","剪开包装"),("🥤","举杯干杯")],
        "我 猜 是 ______!", "I guess it's ___!"))
```

---

## Task 5: Assemble Session 2 (35–46) + Session 3 (47–54)

**Files:** `create_day3_ocean.py`

- [ ] **Step 1: Session 2 (35–46)**

```python
    page(div(prs, "Session 2 · 下 午", "复 习 + 语 言 (我 会 认 6 · 我 会 写 3)", TEAL, "📖"))
    # 36 quick review — match item→animal
    s=page(ns(prs)); bg(s,CREAM); hb(s,"🔁 快 速 复 习 — 塑 料 伤 害 谁?", TEAL)
    tb(s,0.4,0.82,9.2,0.3,"把 塑 料 和 它 伤 害 的 动 物 连 起 来 (口 头)",sz=12,b=True,c=GRAY,a=PP_ALIGN.CENTER)
    for i,it in enumerate(ITEMS):
        y=1.25+i*0.62
        panel(s,1.3,y,3.2,0.52,TINTS[it["tint"]],fill=TINTS[it["tint"]],lw=0)
        tb(s,1.45,y+0.10,3.0,0.35,f'{it["em"]} {it["cn"]}',sz=14,b=True,c=WHITE)
        panel(s,5.5,y,3.2,0.52,CORAL,fill=WHITE,lw=2)
        tb(s,5.65,y+0.10,3.0,0.35,f'{it["victim"]} ？',sz=14,b=True,c=DARK)
    # 37 bamboozle
    s=page(ns(prs)); bg(s,CREAM); hb(s,"🎮 复 习 游 戏 · Baamboozle", TEAL)
    panel(s,1.5,1.1,7.0,2.6,TEAL,fill=INK,lw=3); tb(s,1.5,1.9,7.0,1.0,"🎮",sz=80,a=PP_ALIGN.CENTER)
    btn=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(3.0),Inches(3.9),Inches(4.0),Inches(0.6))
    btn.fill.solid(); btn.fill.fore_color.rgb=FIRE_ORANGE if hasattr(H,'FIRE_ORANGE') else OCEAN; btn.line.fill.background()
    tb(s,3.0,4.0,4.0,0.4,"▶️ 点 击 开 始  Play",sz=15,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    url="https://www.baamboozle.com/games"  # replace with real game id
    btn.click_action.hyperlink.address=url
    tb(s,0.4,4.7,9.2,0.3,url,sz=9,c=GRAY,a=PP_ALIGN.CENTER)
    # 38-43 我会认
    for (em,cn,py,en,ex,exen,hint) in VOCAB_RECOGNIZE:
        page(vocab_recognize(prs, OCEAN, em, cn, py, en, ex, exen, hint))
    # 44-46 我会写
    for (phrase, chars, py, en) in VOCAB_WRITE:
        page(vocab_write(prs, OCEAN, phrase, en, chars))
```

- [ ] **Step 2: Session 3 (47–54)**

```python
    page(div(prs, "Session 3 · 下 午", "练 习 册 + 动 手 (海 洋 拼 贴 · 承 诺)", AQUA, "🎒"))
    # 48 booklet preview
    s=page(ns(prs)); bg(s,CREAM); hb(s,"📓 完 成 「海 洋 小 卫 士」练 习 册  Day 3 Booklet", EARTH_BROWN if hasattr(H,'EARTH_BROWN') else AQUA)
    for i,(t,d) in enumerate([("① 圈 污 染","哪 些 是 一 次 性 塑 料?"),("② 连 一 连","塑 料 → 换 成 什 么"),
        ("③ 我 的 承 诺","我 是 海 洋 小 卫 士, 我 可 以 ___"),("④ 描 一 描","水 · 保 护 · 塑 料")]):
        r,c=divmod(i,2); x=0.6+c*4.5; y=1.1+r*1.7
        panel(s,x,y,4.2,1.5,AQUA,fill=WHITE,lw=2.5); tb(s,x+0.2,y+0.2,3.8,0.5,t,sz=17,b=True,c=AQUA)
        tb(s,x+0.2,y+0.75,3.8,0.6,d,sz=12,b=True,c=DARK)
    # 49 hands-on menu
    s=page(ns(prs)); bg(s,CREAM); hb(s,"🎨 动 手 时 间!  Hands-On — 2 个 活 动", AQUA)
    for i,(t,d,cl) in enumerate([("PROJECT 1 · 🖼️ 海 洋 拼 贴 画","用 回 收 材 料 拼 一 片 干 净 的 海 洋", OCEAN),
        ("PROJECT 2 · 🛍️ 装 饰 环 保 袋 + 承 诺","装 饰 布 袋 + 写 下 保 护 海 洋 的 承 诺", AQUA)]):
        x=0.5+i*4.7; panel(s,x,1.1,4.3,3.4,cl,fill=WHITE,lw=3)
        tb(s,x+0.2,1.35,3.9,0.9,t,sz=16,b=True,c=cl,a=PP_ALIGN.CENTER)
        tb(s,x+0.2,3.4,3.9,0.9,d,sz=13,b=True,c=DARK,a=PP_ALIGN.CENTER)
        tb(s,x+0.2,2.2,3.9,0.9,("🖼️" if i==0 else "🛍️"),sz=70,a=PP_ALIGN.CENTER)
    # 50 project 1 detail
    s=page(ns(prs)); bg(s,CREAM); hb(s,"🖼️ Project 1: 海 洋 拼 贴 画  Ocean Collage", OCEAN)
    panel_head(s,0.4,0.95,4.5,OCEAN,"🧺 材 料  Materials"); 
    for i,m in enumerate(["🟦 蓝 纸 (海 水)","🐟 彩 纸 (鱼)","♻️ 干 净 塑 料 片","🍃 树 叶","🖊️ 胶 水"]):
        tb(s,0.55,1.55+i*0.5,4.2,0.4,m,sz=13,b=True,c=DARK)
    panel_head(s,5.1,0.95,4.5,FIRE_ORANGE if hasattr(H,'FIRE_ORANGE') else AQUA,"👉 做 法  Steps")
    for i,st in enumerate(["1️⃣ 想 一 片 干 净 的 海 洋","2️⃣ 用 材 料 拼 出 海 和 动 物","3️⃣ 贴 牢, 晾 一 下","4️⃣ 介 绍 我 的 作 品"]):
        tb(s,5.25,1.55+i*0.55,4.3,0.45,st,sz=13,b=True,c=DARK)
    sentence_frame_bar(s,4.7,"这 是 干 净 的 海 洋, 有 ______。","This is a clean ocean with ___.",OCEAN)
    # 51 project 1 gallery
    s=page(ns(prs)); bg(s,CREAM); hb(s,"🖼️ Project 1 · 参 考 作 品  Examples", OCEAN)
    for i in range(6):
        r,c=divmod(i,3); x=0.4+c*3.15; y=1.1+r*1.75
        photo_slot(s,x,y,2.95,1.6,"海 洋 拼 贴 参 考","collage example",OCEAN)
    # 52 project 2
    s=page(ns(prs)); bg(s,CREAM); hb(s,"🛍️ Project 2: 装 饰 环 保 袋 + 承 诺", AQUA)
    panel_head(s,0.4,0.95,4.5,AQUA,"🧺 材 料 + 做 法")
    for i,st in enumerate(["🛍️ 一 个 布 袋","🖍️ 彩 笔 / 贴 纸","1️⃣ 画 海 洋 动 物","2️⃣ 写 一 句 承 诺","3️⃣ 贴 到 承 诺 墙"]):
        tb(s,0.55,1.55+i*0.55,4.2,0.45,st,sz=13,b=True,c=DARK)
    panel(s,5.1,0.95,4.5,3.3,AQUA,fill=WARM,lw=2.5); tb(s,5.1,1.6,4.5,1.4,"🛍️🌊",sz=70,a=PP_ALIGN.CENTER)
    tb(s,5.2,3.2,4.3,0.8,"我 的 布 袋, 我 天 天 用!",sz=15,b=True,c=AQUA,a=PP_ALIGN.CENTER)
    sentence_frame_bar(s,4.7,"我 是 海 洋 小 卫 士, 我 可 以 ______。","I'm an Ocean Guardian. I can ___.",AQUA)
    # 53 badge
    s=page(ns(prs)); bg(s,CREAM); hb(s,"🏅 Day 3 海 洋 小 卫 士 徽 章  Ocean Guardian Badge", OCEAN)
    c=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.2),Inches(1.0),Inches(3.6),Inches(3.0))
    c.fill.solid(); c.fill.fore_color.rgb=WHITE; c.line.color.rgb=OCEAN; c.line.width=Pt(4)
    tb(s,3.2,1.2,3.6,0.4,"DAY 3",sz=16,b=True,c=CORAL,a=PP_ALIGN.CENTER)
    tb(s,3.2,1.6,3.6,0.9,"🌊🐢",sz=54,a=PP_ALIGN.CENTER)
    tb(s,3.2,2.7,3.6,0.5,"海 洋 小 卫 士",sz=20,b=True,c=OCEAN,a=PP_ALIGN.CENTER)
    tb(s,3.2,3.25,3.6,0.4,"✓ COMPLETED",sz=13,b=True,c=OK,a=PP_ALIGN.CENTER)
    tb(s,0.4,4.25,9.2,0.4,"⭐⭐⭐⭐⭐⭐  6 颗 星 都 拿 到 啦!",sz=20,a=PP_ALIGN.CENTER)
    tb(s,0.4,4.8,9.2,0.3,"学 会 了 6 种 塑 料 · 6 个 替 代 · 我 会 认 6 词 · 我 会 写 3 词",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    # 54 see you tomorrow
    page(share_close(prs, OCEAN,
        ["我 是 海 洋 小 卫 士, 我 可 以 ______。", "____ 不 好, 我 应 该 ____。"],
        "I'm an Ocean Guardian — I can ___ / I should ___.",
        "Day 4 — 家 庭 零 废 弃", "Family Zero Waste — 在 家 怎 样 零 废 弃?", "🏠"))
```

> Uses `EARTH_BROWN`, `FIRE_ORANGE` from `_helpers` — import them at top of builder
> (add to the `from _helpers import (...)` list). Remove `hasattr` guards once imported.

- [ ] **Step 2b: Fix imports** — add `EARTH_BROWN, FIRE_ORANGE, MOSS` to the import list; replace the `hasattr(...)` guards with the plain names.

- [ ] **Step 3: Render full deck**

Run: `cd Chinese/zero_waste零废弃 && python create_day3_ocean.py`
Expected: `Saved 54 slides` (±1), no traceback.

---

## Task 6: Content QA

- [ ] **Step 1: Text/content extraction**

Run: `cd Chinese/zero_waste零废弃 && python -m markitdown PPT/day3_water_plastic.pptx | less`
Check: all 6 items present with victim + rows + swap; vocab 6+3; objectives; story/video; projects; badge; no missing text.

- [ ] **Step 2: Placeholder scan**

Run: `python -m markitdown PPT/day3_water_plastic.pptx | grep -iE "xxxx|lorem|TODO|placeholder|None"`
Expected: no results (real-photo slots say "真实照片", which is intentional — ignore those).

---

## Task 7: Visual QA (subagent, per pptx skill)

- [ ] **Step 1: Render to images**

```bash
cd Chinese/zero_waste零废弃
python /Users/Huan/.claude/plugins/cache/anthropic-agent-skills/claude-api/da20c92503b2/skills/pptx/scripts/office/soffice.py --headless --convert-to pdf PPT/day3_water_plastic.pptx
pdftoppm -jpeg -r 110 day3_water_plastic.pdf /tmp/d3slide
```

- [ ] **Step 2: Dispatch a visual-QA subagent** (Explore/general) with the pptx skill's visual-QA prompt over `/tmp/d3slide-*.jpg`, expected content per slide. Collect issues: overlaps, text overflow, low contrast, misaligned rows (esp. `judge/reveal` A/B rows and the compare table), off-canvas text.

- [ ] **Step 3: Fix issues** in `create_day3_ocean.py` / `_helpers.py`, re-render affected slides:

```bash
pdftoppm -jpeg -r 110 -f N -l N day3_water_plastic.pdf /tmp/d3fix
```

- [ ] **Step 4: Re-verify** until a full pass finds no new issues (at least one fix-and-verify cycle).

---

## Task 8: Finalize

- [ ] **Step 1: Confirm output in place** — `ls -la PPT/day3_water_plastic.pptx` (builder writes there directly).
- [ ] **Step 2: Clean temp** — remove `day3_water_plastic.pdf`, `/tmp/d3slide-*.jpg`.
- [ ] **Step 3: Summarize** to user: slide count, what needs teacher input (verify YouTube/Baamboozle URLs, insert real photos), and offer to `git commit` on a branch.

---

## Self-review notes
- **Spec coverage:** obj1 = story + water slides (6,8,9); obj2 = observe/想 + victim + compare + review; obj3 = single-use framing (11) + 看 step; obj4 = judge/reveal/swap + projects + badge. Vocab matches booklet. ✓
- **Types:** `ITEMS` dict keys (`em,cn,en,tint,victim,senses,dangers,rows,do,swap`) used consistently across Tasks 4–5. Helper signatures fixed in Task 2 and called with matching args. `page()` wraps every slide once (noted rule). ✓
- **Known risk:** the A/B row geometry (`_ab_rows`) and the compare table are the tightest layouts — Task 7 must verify no overflow; adjust `rh`, font sizes if wrapping.
