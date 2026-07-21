# Zero Waste Day 3 — 家庭 Zero Waste 计划 (Family Plan) Design

**Date:** 2026-05-31
**Unit:** `Chinese/zero_waste零废弃/`
**Primary artifact:** `Chinese/zero_waste零废弃/day3_family_zerowaste.pptx` (built by `day3_family_zerowaste.py`)
**Reference style:** wilderness camp `Chinese/野外生存与探险wilderness_pbl/create_day2_camp.py` (python-pptx helpers, zone/sentence-frame patterns)
**Sibling lessons:** Day 1 = 垃圾分类 / 4-bin sorting (HTML); Day 2 = 可再生能源 / renewable energy

## Goal

Build Day 3 of the Zero Waste unit as a ~45–48 slide python-pptx deck that teaches the **Zero Waste concept + the 5R principle** (deep focus on 减少 Reduce + 重复使用 Reuse), trains kids to **spot waste at home**, and has each child **design a simple family eco-action** using the frame 「我们家可以 ____。」. Distinct from Day 1's sorting focus — this lesson is about everyday family habits.

## Learning Objectives

1. Understand the **Zero Waste (零废弃)** core concept
2. Learn the **5R principle** — 拒绝 / 减少 / 重复使用 / 回收 / 堆肥 (deep on 减少 + 重复使用)
3. Observe and identify **waste behaviors at home**
4. Design a simple **family eco-action plan**

## Language Targets

- **我会认 ×5:** 减少 · 重复使用 · 环保 · 购物袋 · 浪费
- **我会写 ×2:** 减少 · 环保 (with 田字格 grid)
- **Primary sentence frame:** 「我们家可以 ____。」 *(Our family can ____.)*
- **Secondary frame (find-waste):** 「这里有浪费！我们可以 ____。」

## Teaching methods (from brief)

1. Explain 减少 / 重复使用 with **actions + pictures**.
2. **Family-scene images** — find the waste (找浪费).
3. **Writing practice** — 减少, 环保 in 田字格.
4. **Sentence-frame practice** — 我们家可以 ____。

## Audience & Format

- K–5 mixed age (5–11), 20–30 students
- 80–90% Chinese immersion
- 3 sessions: morning 45 min · afternoon 45 min · afternoon 90 min
- 10″ × 5.625″ (16:9), `KaiTi` font, mixed CN (bold) + small EN

## Framing & mascot continuity

Day 1's mascots **熊猫奇奇 & 妙妙** come **home with the kids**: 「奇奇妙妙来我们家做客，帮我们家做 Zero Waste 计划。」 Narrative arc: *What is Zero Waste? → meet the 5R family → hunt for 浪费 at home → make our family's plan.*

## Build approach

`day3_family_zerowaste.py` clones `day2_camp.py`'s helper conventions:
`ns / tb / ap / bg / ib / hb / pn / notes / div / pill / sentence_frame_bar`.

New helpers:
- `r5_card(...)` — one 5R card (emoji + CN + EN + accent color); reused on the 5R lineup and each R's detail slide.
- `find_waste_slide(...)` — family-scene image placeholder + 「圈出浪费」 prompt + answer chips.
- `plan_frame_bar(...)` — sentence-frame bar variant carrying 「我们家可以 ____。」.

## Palette

Unit Forest/Moss base for cohesion: `EARTH_GREEN #2C5F2D` (primary), `MOSS #6BA03D` (secondary), `CREAM`, `WARM`, `ALERT`, `OK`. Each of the 5R has a **fixed accent color** used consistently throughout:

| R | 中文 | EN | emoji | accent | depth |
|---|---|---|---|---|---|
| 1 | 拒绝 | Refuse | 🙅 | red `#C8253E` | light |
| 2 | **减少** | Reduce | ⬇️ | orange `#E07A2C` | **deep** |
| 3 | **重复使用** | Reuse | 🔄 | blue `#3E6EB6` | **deep** |
| 4 | 回收 | Recycle | ♻️ | green `#6BA03D` | light (callback to Day 1 桶) |
| 5 | 堆肥 | Rot | 🍂 | brown `#6B4423` | light |

减少 & 重复使用 (the 我会认/我会写 focus) get intro + examples slides; the other three get one light intro each.

## Style choices (ported from day2_camp + unit prefs)

1. **Transparent / gray photo placeholders** — teacher inserts real photos at use time.
2. **Sentence-frame bar** recurs on plan/practice slides.
3. **Direct-answer format** for practice (show the correct R / action, not guess-then-reveal).
4. **Mixed CN + EN** — small EN under bold CN.
5. **Native Chinese phrasing**, not English-translated style.

## Slide-by-slide plan (~45–48)

### Session 1 — 什么是 Zero Waste? + 5R (≈18 slides, 45 min)

| # | Slide | Notes |
|---|---|---|
| 1 | Cover —「家庭 Zero Waste 计划」 | 🐼🏠 + 奇奇妙妙 teaser |
| 2 | Session 1 divider | 10:00–10:45 |
| 3 | Learning goals | 4 goals above |
| 4 | Hook — 奇奇妙妙来做客 | 一天我们家产生多少垃圾? |
| 5 | 什么是 Zero Waste (零废弃) | concept via action + pictures |
| 6 | 浪费是什么? | 我会认 「浪费」 + examples |
| 7 | Think-Pair-Share | 想 / 说 / 听 + timer |
| 8 | **5R 大家庭 lineup** | 5 colored `r5_card`s |
| 9–10 | **减少 Reduce** intro + examples | action + pictures / 5-example grid + frame bar |
| 11–12 | **重复使用 Reuse** intro + examples | incl. 购物袋 / grid + frame bar |
| 13 | 拒绝 Refuse (light) | one card slide |
| 14 | 回收 Recycle (light) | callback to Day 1 四个桶 |
| 15 | 堆肥 Rot (light) | one card slide |
| 16–17 | 配一配：行为 → 哪个 R ×2 | direct-answer format |
| 18 | S1 exit ticket | 「我们家可以 ____。」 |

**Match items (16–17):** e.g. 自带购物袋 → 减少/重复使用 · 用旧瓶子种花 → 重复使用 · 不要免费传单 → 拒绝 · 剩菜堆肥 → 堆肥 · 旧纸投蓝桶 → 回收.

### Session 2 — 家庭找浪费 + 词汇/书写 (≈11 slides, 45 min)

| # | Slide | Notes |
|---|---|---|
| 19 | S2 divider | — |
| 20–22 | **家庭场景找浪费 ×3** | 厨房 / 客厅 / 浴室 scenes — `find_waste_slide`: 圈出浪费 + 我们可以 ____ |
| 23 | 我会认 ×5 review | 拍词卡 (减少/重复使用/环保/购物袋/浪费) |
| 24 | 词汇配对 | 5 words ↔ shuffled icons |
| 25 | 句型练习「我们家可以 ____」 | Pair-Share example dialogue |
| 26–27 | **我会写 减少 / 环保** | 田字格 grid |
| 28 | 🎮 Bamboozle | click-to-open button + `bamboozle_day3_family.csv` callout |
| 29 | Reflection | 「今天我学会了 ____」 |

**Find-waste answers (20–22):** 厨房 — 开着的水龙头 / 剩饭 / 一次性袋子；客厅 — 没人看的电视/灯 / 一次性水瓶；浴室 — 流着的水 / 用太多纸巾.

### Session 3 — 两个动手活动 (≈14 slides, 90 min)

| # | Slide | Notes |
|---|---|---|
| 30 | S3 divider | 🛠️ 我们家的 Zero Waste 计划 |
| 31 | Project menu | 2-card choice |
| 32–35 | **核心 — 「我们家可以 ____」家庭计划海报** | materials · 想一想(从 5R 选 3 件事) · teacher demo · 30-min work |
| 36–39 | **重复使用 — 装饰环保购物袋** | materials · 4-step · design ideas · 25-min work + 安全 |
| 40 | 🖼️ Gallery Walk | 摆桌 → 静走 → 拍桌 |
| 41–42 | Sharing sentence frames | poster frame + bag frame |
| 43 | Group photo | 全班合影 — 拿着作品! |

### Close (2 slides)

| # | Slide | Notes |
|---|---|---|
| 44 | 在家练习 / 家园共育 | 贴海报 + 全家一起做一件事 |
| 45 | Share + close | 「我们家可以 ____」 + 下次见 |

**Total: ~45 slides** (final count settles during build, like Day 1).

## Activities & materials (Session 3)

- **核心海报:** cardstock/paper (1/student) · markers/crayons · optional 5R sticker icons · pre-printed 「我们家可以 ____」 template lines.
- **环保购物袋:** plain canvas tote *or* paper bag (1/student) · fabric markers/crayons · optional stencils.

## Side artifacts

- `bamboozle_day3_family.csv` — ~10 review questions (5R + 找浪费 + vocab), ready for teacher upload (same format as `bamboozle_day1_trash.csv`).

## Verification plan

1. Run `day3_family_zerowaste.py`; confirm the `.pptx` builds without error and slide count is ~45.
2. Render representative slides to images (LibreOffice/headless) at 1280×720.
3. Fresh-eyes QA pass — overlap, contrast, text overflow, 5R color consistency, frame-bar wording.
4. Fix any issues + re-verify.

## Out of scope (deferred)

- Real photographs — placeholders only; teacher inserts at use time
- Printed translation booklet — separate work
- Embedded videos (none bundled for this concept lesson) and audio narration
- Animation/transitions beyond python-pptx defaults

## Decisions locked during brainstorm

- ✅ Placement: **new Day 3** — `day3_family_zerowaste.py` → `day3_family_zerowaste.pptx`
- ✅ Scope: **3 sessions** (45 + 45 + 90), ~45 slides, matching Day 1's shape
- ✅ 5R: **full 5R introduced, deep focus on 减少 + 重复使用**
- ✅ Activities: **both** — 「我们家可以 ____」family-plan poster (core) + decorate a reusable 购物袋 (longer hands-on)
- ✅ Mascots: reuse Day 1's 熊猫奇奇 & 妙妙 (come-home framing)
- ✅ Build: python-pptx, cloning `day2_camp.py` helper conventions
