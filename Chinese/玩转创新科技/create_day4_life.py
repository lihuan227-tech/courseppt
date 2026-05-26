#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
玩转 创新 科技 · Day 4: 科技 改变 生活 / How Tech Changes Life
探究 问题: 哪 个 科技 最 改变 生活?
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from _helpers import *
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

prs = make_presentation()
DAY = LIFE_TEAL
n = 0

# ============================================================
# 1 · COVER
# ============================================================
cover(prs, 4, "科技 改变 生活", "Tech Changes Life",
      "📱 🚗 🏥 🏠 ✨", DAY,
      "哪 个 科技 最 改变 生活?",
      "Which tech changed life the most?")
n += 1; pn(prs.slides[-1], n)
notes(prs.slides[-1], "Day 4 · 重 点: 比较 + 选择 + 表达 观点\n• Session 1: 科技 在 4 个 生活 领域\n• Session 2: 词汇\n• Session 3: 设计 解决 真 实 生活 问题")

# ============================================================
# 2 · SESSION 1 DIVIDER
# ============================================================
s = div(prs, "Session 1", "🌍 上午 45 min · 科技 改变 了 什么?", DAY, "📱"); n += 1; pn(s, n)

# ============================================================
# 3 · LEARNING GOALS
# ============================================================
s = learning_goals(prs, DAY, [
    ("1️⃣", "认识 科技 怎么 改变 交通 / 医疗 / 学习 / 家庭", "See how tech changes 4 life areas", CYBER),
    ("2️⃣", "比较 「以前」 vs 「现在」 的 生活", "Compare 'before vs now'", ORANGE),
    ("3️⃣", "理解 科技 解决 真 实 问题", "Understand: tech solves real problems", GREEN),
    ("4️⃣", "表达 自己 觉得 最 有 帮助 的 一 项 科技", "Say which tech helps you most", PINK),
])
n += 1; pn(s, n)

# ============================================================
# 4 · INQUIRY HOOK
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🤔 探究 问题  Inquiry Question", DAY)
q = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.50), Inches(0.95), Inches(9.00), Inches(1.45))
q.fill.solid(); q.fill.fore_color.rgb = DAY; q.line.color.rgb = STAR; q.line.width = Pt(3)
tb(s, 0.60, 1.05, 8.80, 0.50, "哪 个 科技 最 改变 生活?", sz=28, b=True, c=WHITE, a=PP_ALIGN.CENTER)
tb(s, 0.60, 1.60, 8.80, 0.40, "Which technology changed life the most?", sz=14, c=WARM, a=PP_ALIGN.CENTER)
tb(s, 0.60, 2.00, 8.80, 0.30, "🗳️ 等 一下 我们 来 投票!  We'll vote in a minute!",
   sz=11, b=True, c=STAR, a=PP_ALIGN.CENTER)

# Vote candidates (4 areas)
votes = [
    ("🚗", "交通", "Transportation"),
    ("🏥", "医疗", "Medical"),
    ("📚", "学习", "Learning"),
    ("🏠", "家庭", "Home"),
]
card_w = 2.10; gap = 0.15
total = 4*card_w + 3*gap; start = (10 - total)/2
for i, (em, cn, en) in enumerate(votes):
    x = start + i*(card_w + gap)
    panel(s, x, 2.65, card_w, 2.35, DAY, lw=2)
    tb(s, x, 2.80, card_w, 0.80, em, sz=50, a=PP_ALIGN.CENTER)
    tb(s, x, 3.70, card_w, 0.40, cn, sz=18, b=True, c=DAY, a=PP_ALIGN.CENTER)
    tb(s, x, 4.10, card_w, 0.32, en, sz=10, c=GRAY, a=PP_ALIGN.CENTER)
    # vote line
    tb(s, x+0.15, 4.50, card_w-0.30, 0.35, "票 数: ____", sz=11, b=True, c=DARK, a=PP_ALIGN.CENTER)

tb(s, 0.40, 5.10, 9.20, 0.30, "👋 一会儿 投 你 的 一 票! Vote later for your favorite!",
   sz=11, b=True, c=DAY, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "5 分钟 引入: 不 投票, 先 让 学生 「想」")

# ============================================================
# 5 · 以前 vs 现在 (Then vs Now comparison)
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🕰️ 以前 vs 现在  Then vs Now", CYBER)
tb(s, 0.40, 0.85, 9.20, 0.30, "看 看 — 科技 让 这 些 事 变 得 多 不 一 样!",
   sz=12, b=True, c=DARK, a=PP_ALIGN.CENTER)
tb(s, 0.40, 1.18, 9.20, 0.26, "See how technology changed these things!",
   sz=9, c=GRAY, a=PP_ALIGN.CENTER)

comparisons = [
    ("✉️", "联系 朋友", "Contact friends",
     "写 信  → 走 路 / 火车 送  → 几 周 才 到",
     "📱 一 秒 钟 视频 通话!", CYBER),
    ("📖", "找 答案", "Find answers",
     "去 图书馆  → 翻 厚 书  → 几 小时",
     "💬 问 一 下 — 1 秒 就 有!", ORANGE),
    ("🚶", "去 远 的 地 方", "Go far away",
     "走 路 / 马车 → 几 天 / 几 周",
     "✈️ 飞机 — 半 天 到 任 何 地方!", GREEN),
    ("🩺", "看 病", "See doctor",
     "等 医生 来 → 没 设备 → 命 都 难 保",
     "🏥 视频 + AI 诊断 + 高 科技 仪 器", PINK),
]
card_w = 4.40; gap = 0.30; row_gap = 0.15
total = 2*card_w + gap; start = (10 - total)/2
for i, (em, cn_t, en_t, then, now, cl) in enumerate(comparisons):
    row = i // 2; col = i % 2
    x = start + col*(card_w + gap)
    y = 1.50 + row*(1.90 + row_gap)
    panel(s, x, y, card_w, 1.90, cl, lw=2.5)
    # Title bar
    tb(s, x+0.15, y+0.08, 0.55, 0.40, em, sz=20)
    tb(s, x+0.65, y+0.10, card_w-0.80, 0.35, cn_t, sz=13, b=True, c=cl)
    tb(s, x+0.65, y+0.42, card_w-0.80, 0.28, en_t, sz=8, c=GRAY)
    # Then / now
    tb(s, x+0.20, y+0.80, card_w-0.40, 0.40, "📜 以前: " + then, sz=9, b=True, c=DARK)
    tb(s, x+0.20, y+1.30, card_w-0.40, 0.40, "✨ 现在: " + now, sz=10, b=True, c=cl)

n += 1; pn(s, n)
notes(s, "10 分钟 讨论:\n• 老师 念 「以前」, 让 学生 想象 (难!)\n• 再 念 「现在」, 学生 笑 / 惊\n• 强调 — 「科技 帮 我们 解决 这 些 问题」")

# ============================================================
# 6 · 4 个 生活 领域 (4 life areas detailed)
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🌍 科技 在 生活 里  Tech in 4 Life Areas", DAY)

areas = [
    ("🚗", "交通", "Transportation",
     "汽车 · 飞机 · 高 铁 · 地 铁",
     "Cars, planes, trains", CYBER),
    ("🏥", "医疗", "Medical",
     "X 光 · 手术 机器人 · AI 诊断",
     "X-ray, surgery robots, AI", PINK),
    ("📚", "学习", "Learning",
     "电脑 · 视频 课 · Chatbots · App",
     "Computers, video, apps", ORANGE),
    ("🏠", "家庭", "Home",
     "冰箱 · 洗 衣机 · 智能 音箱",
     "Fridge, washer, smart speaker", GREEN),
]
card_w = 2.20; gap = 0.10
total = 4*card_w + 3*gap; start = (10 - total)/2
for i, (em, cn_t, en_t, cn_d, en_d, cl) in enumerate(areas):
    x = start + i*(card_w + gap)
    panel(s, x, 1.05, card_w, 4.00, cl, lw=2.5)
    head = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.05), Inches(card_w), Inches(0.45))
    head.fill.solid(); head.fill.fore_color.rgb = cl; head.line.fill.background()
    tb(s, x, 1.11, card_w, 0.35, cn_t, sz=14, b=True, c=WHITE, a=PP_ALIGN.CENTER)
    tb(s, x, 1.60, card_w, 1.10, em, sz=72, a=PP_ALIGN.CENTER)
    tb(s, x, 2.85, card_w, 0.30, en_t, sz=11, b=True, c=cl, a=PP_ALIGN.CENTER)
    tb(s, x+0.10, 3.25, card_w-0.20, 0.95, cn_d, sz=10, b=True, c=DARK, a=PP_ALIGN.CENTER)
    tb(s, x+0.10, 4.30, card_w-0.20, 0.60, en_d, sz=8, c=GRAY, a=PP_ALIGN.CENTER)

tb(s, 0.40, 5.20, 9.20, 0.30, "💡 想 一 想 — 你 用 过 哪 些? 哪 个 最 改变 你 的 生活?",
   sz=11, b=True, c=DAY, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)

# ============================================================
# 7 · 投票 + 表达 (Vote + Express)
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🗳️ 你 的 投票!  Your Vote!", PURPLE)

panel(s, 0.40, 0.95, 9.20, 1.20, PURPLE, fill=WARM)
tb(s, 0.55, 1.02, 9.00, 0.35, "🎯 想 一 想: 哪 个 科技 对 「你」 最 重要?",
   sz=14, b=True, c=PURPLE, a=PP_ALIGN.CENTER)
tb(s, 0.55, 1.38, 9.00, 0.30, "Which tech matters most to YOU?", sz=11, c=GRAY, a=PP_ALIGN.CENTER)
tb(s, 0.55, 1.70, 9.00, 0.32, "✋ 举 手 投票 · 然后 说 「为什么」",
   sz=11, b=True, c=DARK, a=PP_ALIGN.CENTER)

# Sentence frames panel
panel(s, 0.40, 2.35, 9.20, 2.55, CYBER)
panel_head(s, 0.40, 2.35, 9.20, CYBER, "✍️ 表达 你 的 想法  Express Your Idea", sz=13)

# Frame 1 — K-2
tb(s, 0.55, 2.95, 9.00, 0.32, "K-2:", sz=11, b=True, c=ORANGE)
tb(s, 0.55, 3.25, 9.00, 0.40, "「我 喜欢 ___」",
   sz=18, b=True, c=DARK, a=PP_ALIGN.CENTER)

# Frame 2 — Grade 3-5
tb(s, 0.55, 3.75, 9.00, 0.32, "Grade 3-5:", sz=11, b=True, c=ORANGE)
tb(s, 0.55, 4.05, 9.00, 0.40, "「我 觉得 ___ 最 重要, 因为 它 帮 我 ___」",
   sz=14, b=True, c=DARK, a=PP_ALIGN.CENTER)
tb(s, 0.55, 4.50, 9.00, 0.28, "I think ___ matters most, because it helps me ___",
   sz=9, c=GRAY, a=PP_ALIGN.CENTER)

tip = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(5.00), Inches(9.20), Inches(0.30))
tip.fill.solid(); tip.fill.fore_color.rgb = DAY; tip.line.fill.background()
tb(s, 0.55, 5.03, 9.00, 0.25, "💬 没有 错 答案 — 每 个 人 的 选择 都 OK!",
   sz=10, b=True, c=WHITE, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)

# ============================================================
# 8 · Session 2 divider
# ============================================================
s = div(prs, "Session 2", "📖 下午 45 min · 词汇 — 我 会 认 + 我 会 写", DAY, "📚"); n += 1; pn(s, n)

# ============================================================
# 9-13 · 我 会 认
# ============================================================
recognize_words = [
    ("🔬", "科技", "kē jì", "Technology", "科技 让 生活 更 方便。", "Tech makes life easier.",
     "🔬 实验室 / 显微镜 / 高 科技", DAY),
    ("🏥", "医院", "yī yuàn", "Hospital", "医院 用 很 多 高 科技。", "Hospitals use lots of tech.",
     "🏥 现代 医院 / 仪器", PINK),
    ("🚗", "汽车", "qì chē", "Car", "现在 有 自动 驾驶 汽车。", "We have self-driving cars now.",
     "🚗 现代 汽车 / 电动 车", CYBER),
    ("📱", "手机", "shǒu jī", "Cell phone", "手机 可以 做 很 多 事。", "Phones can do so many things.",
     "📱 智能 手机 / app 屏幕", ORANGE),
    ("🌟", "生活", "shēng huó", "Life", "科技 改变 我们 的 生活。", "Tech changes our life.",
     "🌟 家庭 生活 / 日常", GREEN),
]
for em, cn, py, en, ex_cn, ex_en, hint, cl in recognize_words:
    s = vocab_recognize(prs, cl, em, cn, py, en, ex_cn, ex_en, hint)
    n += 1; pn(s, n)

# ============================================================
# 14-15 · 我 会 写
# ============================================================
s = vocab_write(prs, DAY, "科技", "Technology",
                [("科", "kē", "9 笔", "「禾」 旁 + 「斗」"),
                 ("技", "jì", "7 笔", "「扌」 旁 + 「支」")])
n += 1; pn(s, n)

s = vocab_write(prs, GREEN, "生活", "Life",
                [("生", "shēng", "5 笔", "像 一 颗 小 草 生 长"),
                 ("活", "huó", "9 笔", "「氵」 旁 + 「舌」")])
n += 1; pn(s, n)

# ============================================================
# 16 · Session 3 divider
# ============================================================
s = div(prs, "Session 3", "🎨 下午 90 min · 科技 改造 生活 · 发明 设计", DAY, "🛠️"); n += 1; pn(s, n)

# ============================================================
# 17 · Project — Problem-solver invention
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🛠️ 用 科技 解决 生活 问题  Tech Solves Problems", DAY)

intro = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(0.95), Inches(9.20), Inches(1.10))
intro.fill.solid(); intro.fill.fore_color.rgb = WARM
intro.line.color.rgb = DAY; intro.line.width = Pt(2.5)
tb(s, 0.55, 1.05, 9.00, 0.40, "🎯 你 们 是 「发明家」 — 设计 一 个 解决 生活 问题 的 新 科技!",
   sz=14, b=True, c=DAY, a=PP_ALIGN.CENTER)
tb(s, 0.55, 1.45, 9.00, 0.30, "You're inventors! Design a new tech to solve a real-life problem",
   sz=10, c=GRAY, a=PP_ALIGN.CENTER)
tb(s, 0.55, 1.75, 9.00, 0.28, "🧑‍🤝‍🧑 4-5 人 一 组 · 60 分钟 制作 · 15 分钟 分享",
   sz=11, b=True, c=DARK, a=PP_ALIGN.CENTER)

# 5-step process
steps = [
    ("1️⃣", "🔍", "找 问题", "Find Problem",
     "什么 让 你 烦 恼?", "What bothers you?", CYBER),
    ("2️⃣", "💡", "想 主意", "Brainstorm",
     "怎么 用 科技 解决?", "How can tech help?", ORANGE),
    ("3️⃣", "✏️", "画 + 做", "Build",
     "草 图 + 用 材料 做", "Sketch + craft", PINK),
    ("4️⃣", "🧪", "测 试", "Test",
     "好 不 好 用?", "Does it work?", GREEN),
    ("5️⃣", "🎤", "讲", "Present",
     "讲 给 全 班 听", "Pitch to the class", PURPLE),
]
card_w = 1.65; gap = 0.10
total = 5*card_w + 4*gap; start = (10 - total)/2
for i, (num, em, cn_t, en_t, cn_d, en_d, cl) in enumerate(steps):
    x = start + i*(card_w + gap)
    panel(s, x, 2.25, card_w, 2.60, cl, lw=2)
    tb(s, x, 2.35, card_w, 0.30, num, sz=12, b=True, c=cl, a=PP_ALIGN.CENTER)
    tb(s, x, 2.65, card_w, 0.65, em, sz=34, a=PP_ALIGN.CENTER)
    tb(s, x, 3.35, card_w, 0.32, cn_t, sz=12, b=True, c=cl, a=PP_ALIGN.CENTER)
    tb(s, x, 3.68, card_w, 0.30, en_t, sz=8, c=GRAY, a=PP_ALIGN.CENTER)
    tb(s, x+0.05, 4.05, card_w-0.10, 0.40, cn_d, sz=9, b=True, c=DARK, a=PP_ALIGN.CENTER)
    tb(s, x+0.05, 4.45, card_w-0.10, 0.30, en_d, sz=7, c=GRAY, a=PP_ALIGN.CENTER)

ex = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(5.05), Inches(9.20), Inches(0.50))
ex.fill.solid(); ex.fill.fore_color.rgb = DAY; ex.line.fill.background()
tb(s, 0.55, 5.12, 9.00, 0.32, "💡 灵感: 不 会 掉 的 雨 伞? 自动 洗 鞋 子 的 鞋柜? 会 提醒 喝 水 的 杯子?",
   sz=11, b=True, c=STAR, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "60 分钟:\n• 5 min 老师 介绍\n• 15 min 小 组 头脑 风暴 + 选 1 个 问题\n• 25 min 制作\n• 15 min 分享 + 投票 最 棒 发明")

# ============================================================
# 18 · 项目 工作 表 + 句型
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "📝 发明 工作 表 · Invention Worksheet", PURPLE)

# Left: problem hunt
panel(s, 0.40, 0.95, 4.55, 4.30, PINK)
panel_head(s, 0.40, 0.95, 4.55, PINK, "🔍 找 问题 灵感  Problem Ideas", sz=12)
problems = [
    "🌧️ 雨 天 没 带 伞",
    "🥱 早 上 起 不 来",
    "💧 总 忘 喝 水",
    "🧦 找 不 到 同 一 双 袜子",
    "📚 书包 太 重",
    "🍱 午饭 凉 了",
]
for i, line in enumerate(problems):
    tb(s, 0.55, 1.55+i*0.50, 4.30, 0.45, line, sz=11, b=True, c=DARK)

# Right: pitch sentence frames
panel(s, 5.05, 0.95, 4.55, 4.30, CYBER)
panel_head(s, 5.05, 0.95, 4.55, CYBER, "✍️ 介绍 你 的 发明  Pitch Frames", sz=12)
tb(s, 5.20, 1.55, 4.30, 0.32, "K-2:", sz=10, b=True, c=ORANGE)
tb(s, 5.20, 1.85, 4.30, 0.45, "「我 的 发明 是 ___」",
   sz=15, b=True, c=DARK)
tb(s, 5.20, 2.30, 4.30, 0.30, "「它 帮 你 ___」", sz=15, b=True, c=DARK)

tb(s, 5.20, 2.85, 4.30, 0.32, "Grade 3-5:", sz=10, b=True, c=ORANGE)
g35 = [
    "「我们 的 问题 是 ___」",
    "「所以 我们 发明 了 ___」",
    "「它 用 ___ 解决 问题」",
]
for i, line in enumerate(g35):
    tb(s, 5.20, 3.20+i*0.45, 4.30, 0.40, line, sz=12, b=True, c=DARK)

tip = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(5.30), Inches(9.20), Inches(0.30))
tip.fill.solid(); tip.fill.fore_color.rgb = DAY; tip.line.fill.background()
tb(s, 0.55, 5.33, 9.00, 0.25, "🏆 最 后: 全 班 投票 「最 棒 发明」 + 「最 有 用」 + 「最 好 玩」!",
   sz=10, b=True, c=WHITE, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)

# ============================================================
# 19 · share + close
# ============================================================
s = share_close(prs, DAY,
    frames_cn=["「我 喜欢 ___ 因为 ___」", "「我 的 发明 解决 了 ___」"],
    frames_en="I like ___ because ___ · My invention solves ___",
    next_day_cn="Day 5 · 畅想 未来 科技 — Final Expo!",
    next_day_en="Day 5 · Imagine future tech — Final Showcase!",
    next_emoji="🔮")
n += 1; pn(s, n)

out = os.path.join(os.path.dirname(__file__), "day4_life.pptx")
prs.save(out)
print(f"Saved {out}  ({len(prs.slides)} slides)")
