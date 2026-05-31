# Zero Waste Day 2 — 可再生能源 · 小小能源工程师 (Design)

**Date:** 2026-05-31
**Unit:** `Chinese/zero_waste零废弃/`
**Primary artifact:** `Chinese/zero_waste零废弃/create_day2_energy.py` → `day2_energy.pptx`
**Secondary artifact:** `Chinese/zero_waste零废弃/bamboozle_day2_energy.csv`
**Shared module:** `Chinese/zero_waste零废弃/_helpers.py` (imported, with one new helper added)
**Reference sibling:** `create_day1_trash_v2.py` (same unit, same style)
**Reference style:** wilderness camp `day2_camp+手工+游戏链接✅.pptx` patterns (zone slides, answer panels, sentence-frame bar, A/B/C, video slides, vocab cards)

## Goal

Build Day 2 of the Zero Waste unit as a ~50-slide editable `.pptx` deck (python-pptx, via `_helpers.py`) that teaches **renewable energy** through concept + practice on the narrative spine of **小小能源工程师 (Little Energy Engineers)**. Kids are engineers: Earth is running low on the energy it has been using (fossil), so their mission is to find energy that **never runs out** (太阳 / 风 / 水), then learn to **节约** so it lasts. The deck matches the immersion-classroom layers of the unit's other decks: bilingual support, two recurring target sentence frames, 我会认 / 我会写 vocab, and a hands-on solar-kit project in Session 3.

## Learning Objectives

1. 学生能够理解**地球资源有限** — some energy (fossil) runs out.
2. 学生能够初步理解**可持续发展** — choosing energy that never runs out so the Earth lasts.
3. 学生能够认识**太阳能、风能、水力**等可再生能源.
4. 学生能够思考**如何节约资源**.

**Language objectives (Session 2):**
- 我会认: 地球 · 资源 · 能源 · 太阳能 · 节约 (×5)
- 我会写: 地球 · 节约 (×2)

## Audience & Format

- K–5 mixed age (5–11), 20–30 students
- 80–90% Chinese immersion
- 3 sessions: morning 45 min · afternoon 45 min · afternoon 90 min
- Slide size 10″ × 5.625″, font KaiTi (from `_helpers.make_presentation` / `tb`)

## The Energy Set (palette — all colors already in `_helpers.py`)

| Energy | 中文 | Runs out? | Color | Emoji | Examples |
|---|---|---|---|---|---|
| 🛢️ Fossil | 化石能源 (煤·石油·天然气) | ✗ **用得完** + 污染 | `INK #0F1A3A` / `GRAY` (coal-black / smoke) | 🛢️⛏️ | 煤 · 石油 · 天然气 |
| ☀️ Solar | 太阳能 | ✓ 用不完 | `STAR #F5C242` | ☀️ | 太阳能板 · 太阳能小车 · 太阳能热水 |
| 💨 Wind | 风能 | ✓ 用不完 | `SKY #42A5F5` | 💨 | 风车 · 风力发电机 · 帆船 |
| 💧 Hydro | 水力 | ✓ 用不完 | `DEEP_TEAL #006970` | 💧 | 水坝 · 水车 · 瀑布发电 |

**Day theme color:** `SKY #42A5F5` (clean-energy blue) for `cover` / `div` / generic headers.

## Two target sentence frames (recur via `sentence_frame_bar`)

1. 「**___ 能来自 ___。**」 *(___ energy comes from ___.)* — recognize energy sources (objectives 2–3)
2. 「**我会节约 ___。**」 *(I will save ___.)* — conservation (objective 4)

## New helper to add to `_helpers.py`

`compare_slide(prs, header_text, left, right, frame_cn=None, frame_en=None)` — a two-column "用得完 vs 用不完" contrast:
- LEFT panel = 化石能源 (dark/INK header, ✗ 用得完 + 污染 badge)
- RIGHT panel = 可再生能源 (green/MOSS header, ✓ 用不完 badge), showing ☀️💨💧
- Each side: emoji row + 2–3 short CN+EN bullets
- Optional sentence-frame bar at bottom
This is the only new pattern; every other slide reuses an existing helper. Keep the signature and visual conventions consistent with `zone_slide` / `answer_panels_slide` (rounded panels, `panel_head`, `STAR` accents).

## Style choices (inherited from the unit / day2_camp)

1. **Transparent/high-contrast photo placeholders** via `photo_slot` — teacher inserts real photos at use time.
2. **Sentence-frame bar** recurs on intro + practice slides (`sentence_frame_bar`).
3. **Direct-answer format** — practice slides show the correct answer directly (`answer_panels_slide`), not guess-then-reveal.
4. **Mixed CN + EN** — small EN under bold CN.
5. **Native Chinese phrasing** — not English-translated style.
6. **Teacher notes** on every slide via `notes()` — script, timing, age differentiation (K vs G1–3).

## Slide-by-slide plan (~50)

### Session 1 — 能源故事 → 用得完 vs 用不完 → 三种可再生能源 (24 slides, 45 min)

| # | Slide | Pattern / helper |
|---|---|---|
| 1 | Cover — 能量从哪里来? | `cover` (inquiry: 用不完的能源在哪里?) |
| 2 | Session 1 divider | `div` |
| 3 | Learning goals (4) | `learning_goals` |
| 4 | Hook — 什么需要能量? | warm-up: 灯💡 / 车🚗 / 手机📱 need energy |
| 5 | 能源是什么 | concept — energy helps us do things |
| 6 | Think-Pair-Share — 这些能从哪来? | TPS + timer (teacher_box) |
| 7 | 🛢️ 化石能源 — 煤·石油·天然气 | `zone_slide` (INK) — 挖 / 抽 / 烧 |
| 8 | ⚠️ 用得完！+ 污染 | fossil runs out + 黑烟 (sets up contrast) |
| 9–10 | **用得完 vs 用不完** | **new `compare_slide`** — fossil ✗ vs 太阳/风/水 ✓ |
| 11–12 | ☀️ 太阳能 intro + examples | `zone_slide` (STAR) + frame 「太阳能来自太阳」 |
| 13–14 | 💨 风能 intro + examples | `zone_slide` (SKY) + frame |
| 15–16 | 💧 水力 intro + examples | `zone_slide` (DEEP_TEAL) + frame |
| 17–21 | 这是什么能源? ×5 | `answer_panels_slide` — picture → 「___能来自___」 |
| 22 | 小组游戏 — 能源大分类 | sort 用得完 / 用不完 + 三能源 |
| 23 | 公布答案 + 🥇🥈🥉 | medals + answer key |
| 24 | 出门票 Exit ticket | frame 1 + 3 energy icons |

**5 practice items (17–21):** 太阳能板 → 太阳 · 风车 → 风 · 水坝 → 水 · 煤 → 会用完 (✗) · 太阳能小车 → 太阳

### Session 2 — 语言目标 + 节约 + 游戏 (14 slides, 45 min)

| # | Slide | Pattern / helper |
|---|---|---|
| 25 | S2 divider | `div` |
| 26–30 | 我会认 ×5 — 地球 / 资源 / 能源 / 太阳能 / 节约 | `vocab_recognize` |
| 31 | 词汇配对 | match (names ↔ emoji/items) |
| 32 | 拍词卡 | slap-cards |
| 33 | 句型练习 + Pair Share | **both frames** + example dialogue |
| 34–35 | 我会写 ×2 — 地球 / 节约 | `vocab_write` (田字格) |
| 36 | 如何节约资源 | 关灯 / 关水 / 重复用 + frame 「我会节约 ___」 (objective 4) |
| 37 | 🎮 Bamboozle | `video_slide`-style link + `bamboozle_day2_energy.csv` import note |
| 38 | Closing reflection | 「今天我学会了 ___」 |

### Session 3 — 太阳能套件项目 (12 slides, 90 min)

Reuses existing unit assets (real images + video):
- `互动版_太阳能小车_科学区_大班/太阳能小车_科学区_大班_步骤图/1-彩色版/*.png` + `..._操作视频.mp4`
- `互动版_太阳能发电_科学区_大班/太阳能发电_科学区_大班_步骤图/1-彩色版/*.png` + `..._操作视频.mp4`

| # | Slide | Notes / helper |
|---|---|---|
| 39 | S3 divider | 🔧 我是小小能源工程师! (`div`) |
| 40 | Project menu | 太阳能小车 🚗 / 太阳能发电 💡 — teacher picks by available kit (`ab3_slide`-style 2-card) |
| 41 | 材料 Materials | reuse `1-材料准备.png` (both projects) |
| 42–43 | 步骤 Steps | reuse `步骤图` images (太阳能小车 步骤1、2 / 3、4; 太阳能发电 步骤图1 / 2) |
| 44 | 老师示范 Teacher demo | `video_slide` → reuse 操作视频 (relative path) |
| 45 | ⏱️ 动手时间 Work time | `teacher_box` timer (~35 min) |
| 46 | ☀️ 测试! | take outside to the sun / under a lamp — observe it run |
| 47 | 🖼️ Gallery Walk | 摆桌 → 静走 → 拍桌 |
| 48 | 分享 sentence frame | 「我用太阳能 ____」 |
| 49 | 在家节约 / 活动延伸 | home practice (关灯关水, spot renewable energy outside) |
| 50 | Share + close | `share_close` → Day 3 teaser |

**Total: ~50 slides.**

## Game link integration (Bamboozle)

Slide 37:
- Click-to-open button → baamboozle.com (via `video_slide` play-button pattern or a simple hyperlinked shape)
- CSV callout: `bamboozle_day2_energy.csv` (~10 questions covering 用得完/用不完 + 太阳/风/水 + 节约)
- Teacher setup instructions in notes

The CSV is generated as part of this work (separate artifact, ready for teacher upload), same format as `bamboozle_day1_trash.csv`.

## Materials checklist (Session 3)

**Solar-kit project (teacher picks one):**
- 太阳能小车: solar-car kit (solar panel + motor + wheels/chassis), per the unit's 步骤图 — 1 kit/student or 1/pair
- 太阳能发电: solar-power demo kit (panel + LED/小风扇), per the unit's 步骤图
- Common: scissors, tape, markers; access to sunlight or a bright lamp for testing

## Verification plan

1. Run `python create_day2_energy.py`; confirm it builds `day2_energy.pptx` without error and produces ~50 slides.
2. Render representative slides to images (LibreOffice headless or equivalent) at slide resolution.
3. Fresh-eyes QA pass — overlap, contrast, text overflow, energy-color consistency (☀️STAR / 💨SKY / 💧DEEP_TEAL / 🛢️INK), correct reuse of solar step-images.
4. Confirm Session 3 image/video paths resolve relative to the deck.
5. Fix any issues + re-verify.

## What this does NOT include (deferred / out of scope)

- Real photographs for non-craft slides — `photo_slot` placeholders only; teacher inserts at use time.
- Translation/practice booklet — separate work (as with Day 1).
- Editing the source courseware `.pptx` files in the unit's sub-folders.
- Animations/transitions or audio narration.
- A second non-kit fallback craft — Session 3 is solar-kit only (per decision); classes without kits substitute at teacher discretion.

## Open questions resolved during brainstorm

- ✅ Placement: **Day 2** of `zero_waste零废弃`, file `create_day2_energy.py` → `day2_energy.pptx`.
- ✅ Story spine: **小小能源工程师** (Little Energy Engineers).
- ✅ Energy comparison: **add the full 用得完 vs 用不完 contrast** (fossil → renewable), anchoring objectives 1–2.
- ✅ Sentence frames: **both** — 「___ 能来自 ___。」 + 「我会节约 ___。」
- ✅ Session 3 craft: **solar-kit project only** (太阳能小车 / 太阳能发电), reusing existing step-images + operation videos, teacher-pick menu.
- ✅ Build approach: **python-pptx via `_helpers.py`** (matches sibling `create_day1_trash_v2.py`) + **one new `compare_slide` helper**.
- ✅ Day color: **SKY #42A5F5**.
