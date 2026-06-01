# Zero Waste Day 1 — Rebuild Design (4-Bin + Courseware Integration)

**Date:** 2026-05-31
**Unit:** `Chinese/zero_waste零废弃/`
**Primary artifact:** `Chinese/zero_waste零废弃/day1_trash.html` (self-contained HTML slide deck)
**Reference courseware:** `day 1 - 基础版_垃圾分类_社会科学_小中大班/社会科学-垃圾分类-课件（基础版）.pptx` (pages 4–29) + bundled videos in `视频资源/`
**Reference style:** wilderness camp `day2_camp` patterns (answer panels, sentence-frame bar, zone slides)

> **Revision note (supersedes earlier draft):** The earlier version organized Day 1 around **material types** (纸/塑料/食物/金属) and targeted a python-pptx build. This version reorganizes the course around the **national 4-bin standard** (可回收物 / 厨余垃圾 / 有害垃圾 / 其他垃圾), integrates the existing 垃圾怪 courseware (p4–29), and targets the **HTML deck** `day1_trash.html` as the deliverable.

## Goal

Rebuild Day 1 of the Zero Waste unit as a ~54-slide self-contained HTML deck that teaches the **four national trash-sorting categories** through concept explanation **and** practice, woven onto the narrative spine of the existing 垃圾怪 (Trash Monster) courseware, while adding the immersion-classroom layers the courseware lacks (English support, target sentence frame, 我会认/我会写 vocab, two hands-on crafts).

## Learning Objectives

1. Recognize the **4 trash categories** — 可回收物 · 厨余垃圾 · 有害垃圾 · 其他垃圾
2. Understand trash doesn't disappear — littering harms the environment, so we must **sort** it
3. Initial understanding of where each bin's trash goes — 回收 / 堆肥 / 特殊处理 / 填埋·焚烧
4. Use the sentence frame — **「这是 ____。它要放进 ____ 桶。」** *(This is ____. It goes in the ____ bin.)*

## Audience & Format

- K–5 mixed age (5–11), 20–30 students
- 80–90% Chinese immersion
- 3 sessions: morning 45 min · afternoon 45 min · afternoon 90 min

## The Four Bins (official standard colors)

| Bin | 中文 | Color | Hex (from `_helpers.py`) | Examples (adopted from courseware) | Where it goes |
|---|---|---|---|---|---|
| ♻️ Recyclable | 可回收物 | 🟦 蓝 | `RECYCLE_BLUE #1B6FBA` | 玻璃罐 · 快递箱 · 旧书本 · 旧衣服 · 塑料瓶 | 回收，做成新东西 |
| 🍎 Food/Kitchen | 厨余垃圾 | 🟩 绿 | `MOSS #6BA03D` | 剩饭 · 剩菜 · 骨头 · 落叶 · 残花 | 堆肥，变成肥料 |
| ⚠️ Hazardous | 有害垃圾 | 🟥 红 | `ALERT #C8253E` | 过期药品 · 体温计 · 废旧电池 · 旧灯泡 | 特殊处理 · **🛑 不要碰，告诉大人** |
| 🗑️ Other/Residual | 其他垃圾 | ⬛ 灰 | `GRAY #88888c` | 尘土 · 砖块 · 脏纸巾 · 塑料包装袋 | 填埋 / 焚烧 |

**有害垃圾 safety treatment:** taught with full equal weight, but every 有害 slide carries a red safety bar — 「🛑 不要碰！要告诉大人。 / Don't touch — tell an adult.」

## Courseware Integration (hybrid)

The deck rebuilds the courseware's content as native HTML slides (so the immersion layers integrate cleanly) **and** embeds the 5 bundled local videos as click-to-play `<video controls>` elements.

**Adopted from courseware (p4–29):**
- **Narrative spine:** 垃圾怪 (Trash Monster) appears from littering → mascots 熊猫奇奇 / 妙妙 + the 4 bins help → learn to sort → defeat the monster → town is clean.
- **Bins-as-characters** framing: 「我喜欢吃 ___」 reused on bin-intro slides.
- **Example sets:** verbatim per the table above.
- **Sort game:** drop items into the right bin to defeat 垃圾怪.
- **活动延伸 (p28):** 家园共育 home-practice line + classroom-bin tip, folded into the close.

**Bundled videos (relative path from deck):** `day 1 - 基础版_垃圾分类_社会科学_小中大班/视频资源/`
- `动画视频-脏脏垃圾怪.mp4` → hook (slide 4)
- `动画视频-四个垃圾桶-片段1~4.mp4` → one per bin intro (可回收 / 厨余 / 有害 / 其他)

Videos play when the deck is opened locally (deck and courseware folder share the unit directory).

## Palette

Forest & Moss base from `_helpers.py`; the 4 bins use official standard colors (above). `EARTH_GREEN #2C5F2D` primary, `MOSS #6BA03D` secondary, `WARM/STAR` for timers/accents, `OK/ALERT` for safety + answers.

## Style choices (ported from day2_camp + earlier prefs)

1. **Transparent photo placeholders** — dashed gray border, no fill (teacher inserts real photos).
2. **Sentence-frame bar** — recurs on bin-intro + practice slides to reinforce 「这是 ___。它要放进 ___ 桶。」
3. **Direct-answer format** — practice slides show the correct bin directly (not guess-then-reveal).
4. **Mixed CN + EN** — small EN under bold CN.
5. **Native Chinese phrasing** — not English-translated style.
6. **New pattern `bin_sort_slide`** — replaces the binary ✅/❌ panel: item shown, then the 1 correct bin highlighted among 4 dimmed bins + reason + frame.

## Slide-by-slide plan (~54)

### Session 1 — 垃圾怪 story → 4 bins → sort (24 slides, 45 min)

| # | Slide | Pattern / integration |
|---|---|---|
| 1 | Cover — 垃圾去哪儿了? | 🗑️🍌📰🔋 + 垃圾怪 teaser |
| 2 | Session 1 divider | 10:00–10:45 / 11:00–11:45 |
| 3 | Learning goals | 4 goals (objectives above) |
| 4 | Hook — 垃圾怪来了! | embed `脏脏垃圾怪.mp4` + 看前问题 |
| 5 | Discussion | ❓ 垃圾怪怎么来的?(乱丢垃圾) 它做了什么?(破坏环境) + 小结 |
| 6 | Think-Pair-Share | 我们能帮忙吗? — 想/说/听 + timer |
| 7 | Mascots + bins help | 熊猫奇奇/妙妙 + 4 bins「别担心，我们来帮忙!」 |
| 8 | 4 bins line-up | color + symbol: 🟦可回收 · 🟩厨余 · 🟥有害 · ⬛其他 |
| 9–10 | 可回收物 intro + examples | embed `片段1` + 「我喜欢吃…」 / 5-example grid + frame bar |
| 11–12 | 厨余垃圾 intro + examples | embed `片段2` / 5-example grid + frame bar |
| 13–14 | 有害垃圾 intro + examples | embed `片段3` / grid + **🛑 safety bar** + frame |
| 15–16 | 其他垃圾 intro + examples | embed `片段4` / grid + frame bar |
| 17–21 | 放进哪个桶? ×5 | `bin_sort_slide` — direct answer + frame |
| 22 | Sort game — 打败垃圾怪! | 8 cards (2/bin) + 4 colored bins |
| 23 | 垃圾怪被打败了! + medals | 🎉 town clean + 🥇🥈🥉 + answer key |
| 24 | Exit ticket | sentence frame + 4 bin icons |

**5 practice items (17–21):** 旧书本 → 可回收 · 骨头 → 厨余 · 废旧电池 → 有害🛑 · 脏纸巾 → 其他 · 玻璃罐 → 可回收
**8 sort-game cards (22):** 纸盒/塑料瓶 → 可回收 · 树叶/剩饭 → 厨余 · 旧电池/旧灯泡 → 有害 · 脏纸巾/塑料袋 → 其他

### Session 2 — Vocab + games (14 slides, 45 min)

| # | Slide | Pattern |
|---|---|---|
| 25 | S2 divider | — |
| 26–30 | 我会认 ×5 | 可回收 / 厨余 / 有害 / 其他 / 分类 |
| 31 | 词汇配对 | 4 bin names ↔ shuffled emoji/items |
| 32 | 拍词卡 | slap-cards |
| 33 | 句型练习 + Pair Share | frame 「这是 ___。它要放进 ___ 桶。」 + example dialogue |
| 34–36 | 我会写 ×3 | 垃圾 / 分类 / 回收 with 田字格 |
| 37 | 🎮 Bamboozle | click-to-open button + `bamboozle_day1_trash.csv` import note |
| 38 | Closing reflection | 「今天我学会了 ___」 |

### Session 3 — Two craft projects (14 slides, 90 min)

| # | Slide | Notes |
|---|---|---|
| 39 | S3 divider | 🛠️ 我是垃圾分类专家! |
| 40 | Project menu | 2-card choice |
| 41–44 | **Project 1 — 垃圾分类转盘** | materials · 4-step · teacher demo · work time (30-min). **4 spinner sections = the 4 bins in standard colors** (this is the courseware's 科学区 sorting aid, hands-on) |
| 45–48 | **Project 2 — Shrinky Dink 环保钥匙牌** | materials · 4-step · design ideas (4 bin icons) · oven safety + work time (25-min) |
| 49 | 🖼️ Gallery Walk | 摆桌 → 静走 → 拍桌 |
| 50–51 | Sharing sentence frames | spinner frame + keychain frame |
| 52 | Group photo | 全班合影 — 拿着作品! |

### Close (2 slides)

| # | Slide | Notes |
|---|---|---|
| 53 | 活动延伸 / 在家练习 | 家园共育 (home sorting + 宝宝巴士《学垃圾分类》) + classroom-bin tip |
| 54 | Share + close | 「今天我学会了 ___ 要放进 ___ 桶」 + 4 bins + 下次见 |

**Total: 54 slides**

## Game link integration (Bamboozle)

Slide 37:
- Big "Bamboozle" callout + click-to-open button → baamboozle.com
- CSV file callout: `bamboozle_day1_trash.csv` (~10 "放哪个桶?" questions covering all 4 bins)
- Teacher setup instructions

The CSV is generated as part of this work (separate artifact, ready for teacher upload).

## Materials checklist (Session 3)

**Project 1 — 垃圾分类转盘:** paper plates (1/student) · brass fasteners · markers/crayons · scissors · optional pointer template
**Project 2 — Shrinky Dink 钥匙牌:** Shrinky Dink sheets (3"×3") · Sharpies (multi-color) · keyrings · hole punch · oven (teacher-only, 350°F, 1–3 min)

## Verification plan

1. Open `day1_trash.html` locally; confirm all 54 slides render and navigation works (arrows / number-jump / `#N` deep-link).
2. Confirm the 5 embedded videos resolve via their relative paths and play.
3. Render representative slides to images (headless Chrome) at 1280×720.
4. Dispatch a fresh-eyes QA pass against the slides — overlap, contrast, text overflow, bin-color consistency.
5. Fix any issues + re-verify.

## What this does NOT include (deferred / out of scope)

- Real photographs — placeholders only; teacher inserts at use time
- Translation booklet — Day 1 booklet is separate work
- Re-authoring the source `.pptx` courseware — the deck embeds its videos and rebuilds its content, but does not edit the original file
- Animation/transitions beyond the deck's built-in entrance reveals; audio narration — none

## Open questions resolved during brainstorm

- ✅ Categories: **national 4-bin standard** (可回收物 / 厨余 / 有害 / 其他) — reverses the earlier material-type decision
- ✅ Bin colors: **official standard** (蓝 / 绿 / 红 / 灰)
- ✅ Sentence frame: **「这是 ___。它要放进 ___ 桶。」** (item → bin)
- ✅ 有害垃圾: full teach + 「不要碰，告诉大人」 safety rule
- ✅ Courseware (p4–29): **hybrid** — adopt narrative/examples/colors as native slides + embed the 5 bundled videos
- ✅ Deliverable: spec now → implementation plan next; HTML rebuilt in a later step
