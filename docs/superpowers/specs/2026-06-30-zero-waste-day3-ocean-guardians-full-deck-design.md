# Zero Waste · Day 3 — 海洋小卫士 / Little Ocean Guardians (full deck)

**Date:** 2026-06-30
**Unit:** 零废弃 Zero Waste (Chinese, K–5th, bilingual)
**Day:** Day 3 — 水资源 + 塑料污染 (Water resources + Plastic pollution)
**Supersedes:** `2026-05-31-day2-water-plastic-ocean-guardians-design.md` (that was an
afternoon-only 44-slide deck that put craft/games/role-play out of scope). This is the
**full-day deck** modeled on the wilderness `day1_nature` reference.

## Goal

Produce a complete, fully-editable teaching PPTX for Day 3 that **follows the structure
and interaction of the reference deck** `day1_nature+手工+游戏链接✅.pdf` (Wilderness
Adventure Camp, Day 1), re-themed to an **ocean palette** and the **海洋小卫士 (Little
Ocean Guardians)** mission. The deck must use **story + YouTube video** to make the topic
approachable and relatable to young students.

### Learning objectives (drive content)
1. 理解水资源的重要性 — understand why water matters.
2. 认识塑料污染对海洋动物的影响 — plastic pollution harms ocean animals.
3. 理解一次性塑料的问题 — the problem with single-use plastic.
4. 提出减少塑料的方法 — propose ways to reduce plastic.

### How the spine maps to the objectives
The repeated interactive **station = one single-use plastic item** (user's choice). Each
station runs the reference's **看 → 想 → 说 → 判(A/B) → 做** loop:
- **看 Observe** the item in daily life + "用一次就扔" framing → **objective 3**
- **想 Think** what happens when it reaches the ocean / which animal it hurts → **objective 2**
- **判 Judge (A/B)** disposable-vs-reusable daily choices, then **答案揭晓 Reveal** with the
  reusable swap → **objective 4**

**Objective 1 (water matters)** is carried by the Session 1 opener: a water-cycle **story**
+ 2 water slides, before the plastic spine.

## The reference's engine (what we replicate)

Every reference day is wrapped: **Title → Schedule → Objectives → Session 1 (story +
stations + compare + games) → Session 2 (review + 我会认 vocab + 我会写 stroke order) →
Session 3 (booklet + hands-on projects + badge + see-you-tomorrow)**.

Each topic station = **3 slides**: 看+想 (photo + sensory prompts + danger option cards +
media link) · 判 A/B (3 scenario rows, "先选再看") · 答案揭晓 (same rows, green ✓ + 做 TPR
action + K–G1 / G2–G3 sentence frames). Recurring elements: full-width colored header bars,
emoji option cards, real-photo placeholders, INK/STAR sentence-frame bars, K–G1 (simple) vs
G2–G3 ("…，因为…") differentiation.

## Build approach

**New builder `Chinese/zero_waste零废弃/create_day3_ocean.py`** built on the existing
`_helpers.py` toolkit — the same design system the reference deck uses. Output replaces
`Chinese/zero_waste零废弃/PPT/day3_water_plastic.pptx` (current file is preserved in git
history). Canvas 10×5.625in (16:9), KaiTi font, bilingual, emoji-forward, native
python-pptx shapes only (fully editable; image placeholders for teacher photos).

### Existing helpers reused (no change)
`cover`, `learning_goals`, `div`, `photo_slot`, `teacher_box`, `sentence_frame_bar`,
`compare_slide`, `video_slide`, `answer_panels_slide`, `vocab_recognize`, `vocab_write`,
`share_close`, `tianzi_box`, `hb`, `panel`, `panel_head`, `tb`, `bg`, `pn`, `notes`.

### New helpers to add to `create_day3_ocean.py` (modeled on the reference)
1. `mission_intro_slide(...)` — "你是海洋小卫士!" 6-card preview + sentence frame (reference
   slide 5: "You're a Little Explorer!").
2. `observe_think_slide(...)` — per-item 看+想: real-photo LEFT, 看/想/感觉 sensory prompts
   RIGHT, danger/animal option cards bottom, optional media link line (reference slide 10).
3. `judge_ab_slide(...)` — 3 scenario rows, each A-left / B-right, vote bar, frame bar
   (reference slide 11).
4. `reveal_ab_slide(...)` — same 3 rows with green ✓ on the safe choice + a green 做 action
   bar + K–G1/G2–G3 frame bar (reference slide 12).
5. `cardgrid_slide(...)` — 6-item overview grid (reference slide 9 "What's in Nature").
6. `guess_slide(...)` — "我在哪里" teacher-acts/students-guess game (reference slide 31).

(If cleaner during implementation, items 3–4 may instead reuse `build_clean_deck.py`'s
`s_practice` / `s_practice_answers`; decided in the plan. Default: new helpers in
`_helpers.py` to keep one design system.)

## Ocean palette (defined in builder)
- `OCEAN` `#0B5C8C` (primary headers / title) · `TEAL` `#006970` · `AQUA` `#2A9D8F`
- `CREAM` `#F6F8EC` (page bg, from `_helpers`) · `SKY` `#42A5F5`
- `ALERT` `#C8253E` (danger / "答案揭晓" reveal headers, from `_helpers`)
- `OK` `#2E7D32` (safe ✓ choices + 做 action bars, from `_helpers`) · `STAR` `#F5C242`
- 6 per-item accent tints (one per plastic item, ocean-compatible) for visual variety.

## Day 3 content (locked)

**Theme/identity:** 海洋小卫士 / Little Ocean Guardians (matches existing
`booklets/booklet_content.js` Day 3). Subtitle: 节约用水 · 减少塑料.

**The 6 single-use plastic items (the spine):**

| # | Item | Ocean victim (想) | A/B daily choice (判) | Reusable swap (做/换) |
|---|------|-------------------|------------------------|------------------------|
| 1 | 🛍️ 塑料袋 Plastic bag | 海龟 (mistakes for jellyfish) | 买东西→要袋 / 带布袋 | 布袋 cloth bag |
| 2 | 🥤 吸管 Straw | 海龟 / 海豚 | 喝果汁→塑料吸管 / 不用 | 不用 · 钢吸管 |
| 3 | 🧴 塑料瓶 Plastic bottle | 海鸟 / 鱼 | 口渴→买瓶装水 / 带水壶 | 水壶 refill bottle |
| 4 | 🍴 一次性餐具 Disposable utensils | 鱼 / 小动物 | 吃饭→一次性叉 / 自带勺 | 自带餐具 |
| 5 | 📦 塑料包装 Plastic packaging | 海豹 (entangled) | 零食→很多小包装 / 大包装 | 少包装 · 散装 |
| 6 | 🥤 一次性杯子 Disposable cup | 海洋垃圾堆 | 饮料→一次性杯 / 自己的杯 | 自带杯子 |

**Story + video (objective 1 + relatable hook):**
- **Story** (water cycle, objective 1): 《一滴水的旅行》 picture-book style — built with
  `video_slide` (cover + 听/看/想 pre-watch prompts + link). Prompts: 水从哪里来? 水去了哪里?
  我们用水做什么? Teacher may sub any water-cycle read-aloud.
- **Video** (launch plastic): a kid-friendly ocean-plastic clip (e.g. sea-turtle rescue).
  **Links are provided as a YouTube *search* URL + suggested titles, marked
  "VERIFY / replace with a classroom-approved link"** — the builder will NOT hard-code a
  fabricated video ID. Teacher confirms before class.

**Vocabulary (aligned to existing Day 3 booklet):**
- 我会认 ×6: 水 · 海洋 · 塑料 · 污染 · 保护 · 节约 (booklet's 5 + 节约 for the 6-card rhythm)
- 我会写 ×3: 水 · 保护 · 塑料 (exactly the booklet's `writeChars`, so kids practice the same
  characters; K–G1 trace only, G2–G3 full)

**Session 3 hands-on (standalone — unit has no week-long running build):**
- Project 1 · 海洋拼贴画 Nature/Ocean Collage — build a **clean** ocean scene from recycled
  paper / clean plastic scraps (materials + steps + 展示句型 + examples gallery).
- Project 2 · 海洋小卫士承诺 / 装饰环保袋 — decorate a reusable cloth bag + add to a class
  "无塑承诺墙" (frame: 「我是海洋小卫士，我可以 ___。」, from the booklet's promise page).

## Slide outline (~54 slides)

**Open (1–3)**
1. Cover — `cover(...)` Day 3 · 海洋小卫士 · ocean emoji row · inquiry question
2. Today's Schedule — 3 session bars (上午 / 下午 / 下午)
3. Learning Objectives — `learning_goals(...)` 4 objectives + 我会认/我会写

**Session 1 · 水很重要 + 塑料污染 (4–34)**
4. Session 1 divider (`div`, ocean)
5. 你是海洋小卫士! mission intro — 6-card preview + frame `mission_intro_slide`
6. Story Time — 《一滴水的旅行》 `video_slide` (听/看/想 + link) → **obj 1**
7. 故事讨论 — 水去了哪里? 我们为什么需要水? + sentence frames
8. 水很重要 — 喝/洗/种植物/动物 (icon rows) → **obj 1**
9. 地球的水很珍贵 + 节约用水 — "能喝的水很少" simple visual + 关水龙头 tips
10. 塑料污染 video — `video_slide` sea-turtle/ocean clip (VERIFY link) → launches plastic
11. 什么是一次性塑料 — "用一次就扔" framing → **obj 3**
12. 6-item overview grid — `cardgrid_slide` (the 6 plastic items)
13–30. **6 items × 3 slides** — `observe_think_slide` / `judge_ab_slide` / `reveal_ab_slide`
    per the table above → **obj 2, 3, 4**
31. 6-item compare table — `compare_slide` (item → 伤害的动物 → 换成什么)
32–33. 安全 vs 浪费 action game — think slide + reveal (举手 = 好 / 交叉 = 不好)
34. 我在哪里/我演你猜 — `guess_slide` teacher acts a swap, students guess the item

**Session 2 · 复习 + 语言 (35–46)**
35. Session 2 divider (`div`)
36. Quick Review — match item → animal hurt (or item → swap)
37. Baamboozle review — setup + clickable button (CSV: `bamboozle_day3_*.csv`)
38–43. 我会认 ×6 — `vocab_recognize` cards (水·海洋·塑料·污染·保护·节约)
44–46. 我会写 ×3 — `vocab_write` (水 · 保护 · 塑料)

**Session 3 · booklet + 动手 (47–54)**
47. Session 3 divider (`div`)
48. Complete Day 3 Booklet — page previews (圈污染 / 连一连 item→swap / 我的承诺 / 描写)
49. Hands-On Time — 2 projects menu
50. Project 1 · 海洋拼贴画 — materials + steps + 展示句型
51. Project 1 examples gallery (photo grid)
52. Project 2 · 装饰环保袋 / 承诺墙 — materials + steps + 承诺 frame
53. Day 3 海洋小卫士 Badge — completed, 6 stars (recap: 6 items · 6 swaps · vocab)
54. See You Tomorrow — `share_close(...)` → Day 4 家庭零废弃 (Family Zero Waste)

## Differentiation (every station, per reference)
- **K–G1:** 「我选 A。/ 我选 B。」 · 「不可以 ___。安全/不安全。」
- **G2–G3:** 「我选 B，因为 ___。」 · 「___ 不好，因为 ___。我应该 ___。」

## Deliverables
- `Chinese/zero_waste零废弃/create_day3_ocean.py` — the builder
- New helpers added to `Chinese/zero_waste零废弃/_helpers.py` (items 1–6 above)
- `Chinese/zero_waste零废弃/PPT/day3_water_plastic.pptx` — regenerated deck (replaces current)
- `Chinese/zero_waste零废弃/bamboozle_day3_water.csv` — review-game questions
- QA: render to images + subagent visual inspection (per pptx skill), fix-and-verify loop

## Out of scope
- Embedding real photos/video (placeholders + links only; teacher inserts/verifies).
- Rewriting the Day 3 booklet (`booklet_content.js` already has Day 3 content; deck only
  previews it).
- Other unit days; Google Sites embedding.

## Open items to confirm at build time
- Exact ocean-plastic video URL (teacher-approved) — builder ships a search link + titles.
- Whether to keep all **6** items or trim to **4** if class pacing is tight (builder will
  make the item list a single array so trimming is one-line).
