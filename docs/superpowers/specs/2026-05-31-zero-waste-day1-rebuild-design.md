# Zero Waste Day 1 — PPT Rebuild Design

**Date:** 2026-05-31
**Unit:** `Chinese/zero_waste零废弃/`
**Replaces:** existing 40-slide `day1_trash.pptx`
**Reference style:** `Downloads/day2_camp+手工+游戏链接✅.pptx`

## Goal

Rebuild Day 1 of the Zero Waste unit as a ~54-slide deck that closely follows the structural and visual patterns of the wilderness camp Day 2 reference, while hitting the four stated learning objectives.

## Learning Objectives

1. Recognize 4 common trash types — **纸、塑料、食物垃圾、金属**
2. Understand trash doesn't disappear — it goes somewhere
3. Initial understanding of 3 disposal methods — 回收 / 填埋 / 焚烧
4. Use sentence frames — 「这是____垃圾。它可以/不可以回收。」

## Audience & Format

- K-5 mixed age (5-11), 20-30 students
- 80-90% Chinese immersion
- 3 sessions: morning 45 min · afternoon 45 min · afternoon 90 min

## Architecture

- New build script: `Chinese/zero_waste零废弃/create_day1_trash_v2.py`
- Output: `Chinese/zero_waste零废弃/day1_trash.pptx` (overwrites existing)
- Reuses existing `_helpers.py` with **new helpers ported from `wilderness_pbl/create_day2_camp.py`**:
  - `answer_panels_slide` — ❓ question header + ✅/❌ panels for "可以回收吗?" scenarios
  - `zone_slide` / `zone_q_slide` — for the 4 category intro slides
  - `ab_slide` — for "猜一猜" 3-option guess
  - `video_slide` — for the YouTube book embed
  - `sentence_frame_bar` — recurring banner showing the target sentence frames
- New artifact: `bamboozle_day1_trash.csv` — ~10 Q&A for in-class Bamboozle game

## Palette

Reuse the Forest & Moss palette already in `_helpers.py`:

- Primary: `EARTH_GREEN` (forest)
- Secondary: `MOSS`
- Per-category accents:
  - 纸 → `EARTH_BROWN`
  - 塑料 → `RECYCLE_BLUE`
  - 食物 → `MOSS`
  - 金属 → `FIRE_ORANGE`

## Slide-by-slide plan

### Cover + S1 setup (3 slides)

| # | Slide | Pattern |
|---|---|---|
| 1 | Cover — 垃圾去哪儿了? | `cover()` — 🗑️🍌📰🥤 + inquiry |
| 2 | Session 1 divider | `div()` — 10:00-10:45 / 11:00-11:45 |
| 3 | Learning goals | `learning_goals()` — 4 goals (one per objective above) |

### Session 1 — Discovery → Sort (22 slides, 45 min)

| # | Slide | Pattern | Time |
|---|---|---|---|
| 4 | Warm-up — 透明垃圾袋 | photo left + 4 items right + 3 questions | 4 min |
| 5 | Think-Pair-Share | 3-step icons + timer card | 3 min |
| 6 | Book intro — 《垃圾哪里去了》 | `video_slide()` w/ YouTube link | 2 min |
| 7-10 | 4 book-pause discussions | photo + ❓ question card per pause | 5 min total |
| 11 | Flow chart 家→桶→车→? | 4 panels + arrows, no answer yet | 2 min |
| 12 | 猜一猜 A/B/C | `ab_slide()` 3-option layout | 2 min |
| 13 | 3 disposal methods | 3 photo cards (回收/填埋/焚烧), 1 sentence each | 3 min |
| 14-17 | 4 category intros | `zone_slide()` ported — emoji + 5 examples + photo + rule | 8 min total |
| 18-22 | 5 "可以回收吗?" scenarios | `answer_panels_slide()` ported — ❓ + ✅/❌ + sentence frame | 8 min total |
| 23 | Sort game | 4 zones (纸/塑料/食物/金属) + 8 cards, table-based | 5 min |
| 24 | Reveal + medals | 🥇🥈🥉 + answer key | 3 min |
| 25 | Exit ticket | 2 sentence frames + examples | (rolls to S2) |

**5 scenarios for slides 18-22:** 报纸 ✅纸 · 塑料瓶 ✅塑料 · 香蕉皮 ✅食物 · 牛奶盒 ✅纸 · 易拉罐 ✅金属

### Session 2 — Vocab + games (14 slides, 45 min)

| # | Slide | Pattern |
|---|---|---|
| 26 | S2 divider | `div()` |
| 27-31 | 我会认 × 5 | `vocab_recognize()` — one per: 纸 / 塑料 / 食物 / 金属 / 回收 |
| 32 | 词汇配对 | matching layout (5 words ↔ shuffled emoji) |
| 33 | 拍词卡 | slap-cards layout |
| 34 | 句型练习 + Pair Share | 2 sentence frames + example dialogue |
| 35-37 | 我会写 × 3 | `vocab_write()` — 纸 / 垃圾 / 回收 with 田字格 |
| 38 | 🎮 Bamboozle game time | clickable hyperlink button + CSV import instructions |
| 39 | Closing reflection | "今天我学会了___" |

### Session 3 — Two craft projects (14 slides, 90 min)

| # | Slide | Notes |
|---|---|---|
| 40 | S3 divider | 🛠️ 我是垃圾分类专家! |
| 41 | Project menu | 2-card choice screen |
| **Project 1 — 垃圾分类转盘** | | |
| 42 | Materials | paper plate · brass fastener · markers · scissors |
| 43 | 4-step process | plate → 4 sections → draw → pointer |
| 44 | Teacher demo | finished spinner visual |
| 45 | Work time + walk-around tips | 30-min timer + 5 prompts |
| **Project 2 — Shrinky Dink 环保钥匙牌** | | |
| 46 | Materials | Shrinky Dink · Sharpies · keyring · oven |
| 47 | 4-step process | draw → cut → bake → keyring |
| 48 | Design ideas | Recycle Hero · 环保小达人 · 4 category icons |
| 49 | Safety + work time | oven safety + 25-min timer |
| **Shared closing for Session 3** | | |
| 50 | 🖼️ Gallery Walk | 摆桌 → 静走 → 拍桌 |
| 51-52 | Sharing sentence frames | spinner frame + keychain frame |
| 53 | Group photo + celebration | "全班合影 — 拿着作品!" |

### Close (1 slide)

| # | Slide | Notes |
|---|---|---|
| 54 | Share + close | `share_close()` — "今天我学会了___ 是 ___ 垃圾" |

**Total: ~54 slides**

## Style choices ported from day2_camp

1. **Transparent photo placeholders** — light gray border, no fill (matches user preference from earlier Day 3 booklet work)
2. **Sentence frame bar** — appears on slides 14-22 to reinforce target frames consistently
3. **❓ above each ✅/❌ panel** — direct-answer format, not "guess then reveal" (matches the user's refactor preference from the camp Day 2 work)
4. **Mixed CN + EN bullets inside panels** — small EN under bold CN
5. **Native Chinese phrasing** — not English-translated style (e.g., 「你扔的垃圾去哪儿了?」 not 「你今天扔过垃圾吗?」)

## Game link integration ("+游戏链接✅")

Slide 38 includes:
- Big "Bamboozle" callout with the trash logo
- Clickable hyperlink button → bamboozle.com (or specific game URL the teacher creates)
- CSV file path callout: `bamboozle_day1_trash.csv` (~10 questions covering all 4 objectives)
- Teacher instructions

The CSV file is generated as part of this work, ready for teacher upload.

## Materials checklist (Session 3)

**Project 1:**
- Paper plates (1 per student)
- Brass fasteners
- Markers / crayons
- Scissors
- Pre-printed pointer template (optional)

**Project 2:**
- Shrinky Dink plastic sheets (3"x3" pieces)
- Sharpies (multiple colors)
- Keyrings
- Hole punch
- Oven (teacher-only, 350°F, 1-3 min)

## Verification plan

1. Run `python3 create_day1_trash_v2.py` — expect "Saved ... (54 slides)"
2. Convert to PDF via `soffice --headless --convert-to pdf`
3. Render every slide to JPG at 100dpi
4. Dispatch fresh-eyes QA subagent against all 54 slides — look for overlap, contrast issues, text overflow
5. Fix any issues + re-verify

## What this does NOT include (deferred / out of scope)

- Real photographs — placeholders only; teacher inserts at use time
- Translation booklet — Day 1 booklet is separate work
- Animation / transitions — static slides only
- Audio narration — none

## Open questions resolved during brainstorm

- ✅ Categories: material types (纸/塑料/食物/金属), not Chinese 4-bin
- ✅ Sessions: 3 (same as previous build)
- ✅ Hand crafts: both options (spinner + Shrinky Dink)
- ✅ Session 1: discovery-first arc (book → flow → guess → 3 methods → 4 categories → 5 scenarios → game)
