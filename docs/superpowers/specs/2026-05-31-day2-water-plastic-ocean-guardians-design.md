# Day 2 Afternoon · 保护水资源 & 减少塑料污染 — Ocean Guardians deck

**Date:** 2026-05-31
**Unit:** 零废弃 Zero Waste (Chinese, K–lower elementary, bilingual)
**Session:** Day 2 afternoon (reviews/finishes the morning content)

## Goal

Produce a **full, fully-editable teaching PPTX** for an afternoon Chinese lesson on
protecting water and reducing plastic pollution, in the visual style of the existing
`野外生存与探险wilderness_pbl/day2_camp+手工+游戏链接✅.pptx` (built by
`create_day2_camp.py`) — but **recolored to an ocean palette** and reframed around a
**海洋小卫士 (Ocean Guardians)** mission.

### Learning objectives
1. 理解水资源的重要性 — understand why water matters.
2. 认识塑料污染对海洋动物的影响 — plastic pollution harms ocean animals.
3. 理解一次性塑料的问题 — the problem with single-use plastic.
4. 提出减少塑料的方法 — propose ways to reduce plastic.

### Language goals
- 我会认 (recognize): 水、海洋、塑料、污染、保护
- 我会写 (write): 水、保护、塑料

## Approach

Write a **new builder** `create_day2_water.py` that reuses the example builder's
design-system helper conventions (`ns / tb / ap / bg / hb / pill / div / ib`,
mission cards, `💬 我来说` sentence-frame bar, per-slide image placeholders) recolored
to ocean tones and authoring the new content. Native python-pptx shapes only — no flat
screenshots; every element stays editable. Canvas 10×5.625 in (16:9), KaiTi font,
bilingual throughout (中文 large + English caption small), emoji-forward.

This mirrors how the unit's other days were produced (`create_dayN_*.py`), so it stays
consistent and reproducible.

## Design system

- **Layout vocabulary (from the example):** rounded header bar per content slide; white
  content cards with colored borders + numbered oval badges; "对吗？为什么？" decide
  prompts; warm sentence-frame bar at slide bottom; solid-color section dividers;
  `📷` image placeholders for photos the teacher inserts later.
- **Ocean palette:** deep teal `#0B5563` (primary/headers), ocean blue `#1565A0`,
  sky `#4AA3DF`, sand/cream `#FBF7EC` (page bg), seagreen `#2E8B7A`,
  coral `#E0633F` (alert / "hurts the ocean"), sun-yellow `#F5C242` accent,
  dark `#2C2C2C`, gray `#888888`.
- **Story framing:** kids are **海洋小卫士 Ocean Guardians** on a 4-task mission
  (parallels the example's "Explorer Mission" 4 cards).

## Slide outline (~44 slides)

**Open & review**
1. Cover — title + Ocean Guardians mission band
2. Divider — 下午 · 复习 + 新任务
3. 复习上午 — recap + transition
4. 今天的任务 — 海洋小卫士 4 个任务 (4 mission cards)
5. 学习目标 — the 4 objectives, bilingual

**任务 1 · 水很重要** (divider + 3): why we need water · 地球的水很珍贵 · frame「水很重要，因为 ___。」
**任务 2 · 海洋告急** (divider + 4): 塑料去了海里 · 受伤的海洋动物 (decide cards) · 一只海龟的故事 · frame「塑料让 ___ 受伤。」
**任务 3 · 一次性的麻烦** (divider + 3): 什么是一次性塑料 · 用一次 vs 可重复 · frame「___ 只用一次，不好。」
**任务 4 · 我能做到** (divider + 3): 减少塑料 4 个方法 · 海洋小卫士承诺 frame「我会 ___，保护海洋。」

**语言目标 · 我会认 ×5** (overview + 5 cards): 水 · 海洋 · 塑料 · 污染 · 保护 (char + pinyin + en + 组词 + 📷)
**我会写 ×3** (divider + 3): 水 · 保护 · 塑料 (田字格 trace + strokes)

**Game & wrap**
- Bamboozle 复习游戏 (setup + button; CSV `day2_water_baamboozle.csv` covers morning + today)
- 我们学到 (4-task recap cards)
- 海洋小卫士宣言 + close

## Deliverables

- `Chinese/zero_waste零废弃/create_day2_water.py` — the builder
- `Chinese/zero_waste零废弃/day2_water_plastic.pptx` — the deck (NEW file)
- `Chinese/zero_waste零废弃/day2_water_baamboozle.csv` — review-game questions

## Out of scope

Hands-on craft, sort/decide game section, role-play (not selected). Embedding real
photos/video (image placeholders only). Renaming or modifying existing unit files.
