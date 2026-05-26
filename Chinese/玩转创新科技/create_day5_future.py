#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
玩转 创新 科技 · Day 5: 畅想 未来 科技 / Imagining Future Tech
探究 问题: 未来 还有 什么 问题 需要 解决?
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from _helpers import *
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

prs = make_presentation()
DAY = FUTURE_PINK
n = 0

# ============================================================
# 1 · COVER
# ============================================================
cover(prs, 5, "畅想 未来 科技", "Imagining Future Tech",
      "🔮 🚀 🤖 🌟 ✨", DAY,
      "未来 还 有 什么 问题 需要 解决?",
      "What problems still need solving?")
n += 1; pn(prs.slides[-1], n)
notes(prs.slides[-1], "Day 5 · 综合 一 周 知识 + 设计 思维 + Final Showcase\n• Session 1: 综合 复习 + 头脑 风暴 未来\n• Session 2: 词汇 + 准备 展示\n• Session 3: 未来 科技 博览会 — Final Presentation")

# ============================================================
# 2 · SESSION 1 DIVIDER
# ============================================================
s = div(prs, "Session 1", "🔮 上午 45 min · 这 一 周 学 了 什么? + 想 未来", DAY, "🌟"); n += 1; pn(s, n)

# ============================================================
# 3 · LEARNING GOALS
# ============================================================
s = learning_goals(prs, DAY, [
    ("1️⃣", "回顾 一 周 学 过 的 科技 知识", "Review the week's tech knowledge", CYBER),
    ("2️⃣", "想 一 想 — 未来 还 有 哪 些 新 科技?", "Brainstorm future technologies", ORANGE),
    ("3️⃣", "用 设计 思维 提 出 「问题」 + 「方案」", "Use design thinking", GREEN),
    ("4️⃣", "为 最 后 展示 做 准备", "Prepare for the final showcase", DAY),
])
n += 1; pn(s, n)

# ============================================================
# 4 · INQUIRY HOOK
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🤔 探究 问题  Inquiry Question", DAY)
q = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.50), Inches(0.95), Inches(9.00), Inches(1.45))
q.fill.solid(); q.fill.fore_color.rgb = DAY; q.line.color.rgb = STAR; q.line.width = Pt(3)
tb(s, 0.60, 1.05, 8.80, 0.50, "未来 还 有 什么 问题 需要 解决?", sz=26, b=True, c=WHITE, a=PP_ALIGN.CENTER)
tb(s, 0.60, 1.60, 8.80, 0.40, "What problems still need solving in the future?",
   sz=14, c=WARM, a=PP_ALIGN.CENTER)
tb(s, 0.60, 2.00, 8.80, 0.30, "🌍 想 一 想 地球 / 学校 / 家 里 — 还 有 什么 不 完美?",
   sz=11, b=True, c=STAR, a=PP_ALIGN.CENTER)

# Big-question categories
cats = [
    ("🌍", "地球", "Earth", "环境 / 气候 / 海 洋"),
    ("🏫", "学习", "School", "更 好 玩? 更 公平?"),
    ("🏥", "健康", "Health", "病 / 老 / 心情"),
    ("🚀", "太空", "Space", "去 火星? 找 外星人?"),
]
card_w = 2.10; gap = 0.15
total = 4*card_w + 3*gap; start = (10 - total)/2
for i, (em, cn, en, q_text) in enumerate(cats):
    x = start + i*(card_w + gap)
    panel(s, x, 2.65, card_w, 2.40, DAY, lw=2)
    tb(s, x, 2.80, card_w, 0.75, em, sz=46, a=PP_ALIGN.CENTER)
    tb(s, x, 3.60, card_w, 0.38, cn, sz=16, b=True, c=DAY, a=PP_ALIGN.CENTER)
    tb(s, x, 3.98, card_w, 0.30, en, sz=9, c=GRAY, a=PP_ALIGN.CENTER)
    tb(s, x+0.10, 4.35, card_w-0.20, 0.65, q_text, sz=10, b=True, c=DARK, a=PP_ALIGN.CENTER)

tb(s, 0.40, 5.15, 9.20, 0.30, "✋ 你 最 想 解决 哪 一 类? Pick the area you most want to fix!",
   sz=11, b=True, c=DAY, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)

# ============================================================
# 5 · 一 周 回顾 (Week recap)
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🔁 一 周 回顾  Week Recap", CYBER)
tb(s, 0.40, 0.85, 9.20, 0.30, "我们 学 过 了 — 这 些 都 是 改变 世界 的 新 科技!",
   sz=12, b=True, c=DARK, a=PP_ALIGN.CENTER)

recap = [
    ("🤖", "Day 1", "AI · 人工 智能", "AI", "「会 学 习 的 电脑 助手」", AI_PURPLE),
    ("🖨️", "Day 2", "3D 打印", "3D Printing", "从 设计 到 真实 物 品", PRINT_ORANGE),
    ("🧠", "Day 3", "Machine Learning", "ML", "电脑 看 例子 学 规律", ML_GREEN),
    ("📱", "Day 4", "科技 改变 生活", "Tech in Life", "交通 · 医疗 · 学习 · 家庭", LIFE_TEAL),
]
card_w = 2.20; gap = 0.10
total = 4*card_w + 3*gap; start = (10 - total)/2
for i, (em, day, cn_t, en_t, summary, cl) in enumerate(recap):
    x = start + i*(card_w + gap)
    panel(s, x, 1.30, card_w, 3.80, cl, lw=2.5)
    head = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.30), Inches(card_w), Inches(0.45))
    head.fill.solid(); head.fill.fore_color.rgb = cl; head.line.fill.background()
    tb(s, x, 1.36, card_w, 0.35, day, sz=13, b=True, c=WHITE, a=PP_ALIGN.CENTER)
    tb(s, x, 1.85, card_w, 0.95, em, sz=58, a=PP_ALIGN.CENTER)
    tb(s, x+0.05, 2.95, card_w-0.10, 0.38, cn_t, sz=13, b=True, c=cl, a=PP_ALIGN.CENTER)
    tb(s, x+0.05, 3.35, card_w-0.10, 0.28, en_t, sz=9, c=GRAY, a=PP_ALIGN.CENTER)
    tb(s, x+0.10, 3.85, card_w-0.20, 1.10, summary, sz=10, b=True, c=DARK, a=PP_ALIGN.CENTER)

tb(s, 0.40, 5.20, 9.20, 0.30, "💡 哪 一 天 你 学 到 最 多?  Which day did you learn the most?",
   sz=11, b=True, c=CYBER, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)

# ============================================================
# 6 · 头脑 风暴 — 未来 科技 主题
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "💡 头脑 风暴 · 未来 科技!  Brainstorm: Future Tech!", DAY)

panel(s, 0.40, 0.95, 9.20, 0.85, DAY, fill=WARM)
tb(s, 0.55, 1.05, 9.00, 0.35, "🚀 没有 「太 疯狂」 的 想法 — 越 大胆 越 好!",
   sz=14, b=True, c=DAY, a=PP_ALIGN.CENTER)
tb(s, 0.55, 1.45, 9.00, 0.28, "No idea is too wild — be bold!",
   sz=10, c=GRAY, a=PP_ALIGN.CENTER)

# 6 future-tech inspiration cards
futures = [
    ("🪽", "会 飞 的 鞋 子", "Flying shoes", CYBER),
    ("🧠", "梦 录 像 机", "Dream recorder", PURPLE),
    ("🌱", "能 种 在 月球 的 蔬菜", "Moon-grown veggies", GREEN),
    ("🐶", "听 懂 动物 说 话", "Animal translator", ORANGE),
    ("💊", "一 颗 药 治 所有 病", "Cure-all pill", PINK),
    ("🌍", "吃 垃圾 的 机器人", "Trash-eating robot", TEAL),
]
card_w = 2.95; gap = 0.12; row_gap = 0.15
total = 3*card_w + 2*gap; start = (10 - total)/2
for i, (em, cn, en, cl) in enumerate(futures):
    row = i // 3; col = i % 3
    x = start + col*(card_w + gap)
    y = 1.95 + row*(1.55 + row_gap)
    panel(s, x, y, card_w, 1.55, cl, lw=2)
    tb(s, x, y+0.10, 0.85, 0.50, em, sz=24, a=PP_ALIGN.CENTER)
    tb(s, x+0.75, y+0.18, card_w-0.85, 0.35, cn, sz=13, b=True, c=cl)
    tb(s, x+0.75, y+0.55, card_w-0.85, 0.28, en, sz=9, c=GRAY)
    tb(s, x+0.15, y+0.95, card_w-0.30, 0.50, "💭 还 能 ___?", sz=10, b=True, c=DARK)

tb(s, 0.40, 5.20, 9.20, 0.30, "✋ 现在 你 来 想 — 你 想 发明 什么?",
   sz=11, b=True, c=DAY, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "10 分钟 头脑 风暴:\n• 老师 念 6 个 灵感, 学生 笑\n• 然后 让 学生 自己 想 一 个\n• 写 在 便 利 贴 上 贴 黑 板 — 大家 看")

# ============================================================
# 7 · 设计 思维 4 步
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🛠️ 设计 思维 4 步  Design Thinking", PURPLE)

panel(s, 0.40, 0.95, 9.20, 0.85, PURPLE, fill=WARM)
tb(s, 0.55, 1.05, 9.00, 0.35, "💡 真 正 的 工程师 + 科学家 用 的 方法",
   sz=14, b=True, c=PURPLE, a=PP_ALIGN.CENTER)
tb(s, 0.55, 1.45, 9.00, 0.28, "The same method real engineers and scientists use",
   sz=10, c=GRAY, a=PP_ALIGN.CENTER)

steps = [
    ("1️⃣", "❓", "找 问题", "Find Problem", "什么 让 你 烦?", CYBER),
    ("2️⃣", "💭", "想 主意", "Imagine", "怎么 解决?", ORANGE),
    ("3️⃣", "✏️", "画 + 做", "Design", "画 出 来! 做 模型!", GREEN),
    ("4️⃣", "🎤", "讲 一 讲", "Share", "讲 给 大家", PINK),
]
card_w = 2.20; gap = 0.12
total = 4*card_w + 3*gap; start = (10 - total)/2
for i, (num, em, cn, en, desc, cl) in enumerate(steps):
    x = start + i*(card_w + gap)
    panel(s, x, 1.95, card_w, 3.05, cl, lw=2.5)
    head = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.95), Inches(card_w), Inches(0.45))
    head.fill.solid(); head.fill.fore_color.rgb = cl; head.line.fill.background()
    tb(s, x, 2.01, card_w, 0.35, num, sz=14, b=True, c=WHITE, a=PP_ALIGN.CENTER)
    tb(s, x, 2.55, card_w, 0.75, em, sz=48, a=PP_ALIGN.CENTER)
    tb(s, x, 3.40, card_w, 0.40, cn, sz=15, b=True, c=cl, a=PP_ALIGN.CENTER)
    tb(s, x, 3.80, card_w, 0.30, en, sz=10, c=GRAY, a=PP_ALIGN.CENTER)
    tb(s, x+0.10, 4.20, card_w-0.20, 0.60, desc, sz=11, b=True, c=DARK, a=PP_ALIGN.CENTER)

tb(s, 0.40, 5.15, 9.20, 0.30, "📓 现在 — 在 工作 表 上 写 下 你 的 想法!",
   sz=11, b=True, c=DAY, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "5 分钟 介绍 + 15 分钟 学生 在 工作 表 上 写\n• K-2: 画 + 1-2 个 词\n• 3-5: 用 4 步 写 完整 想法")

# ============================================================
# 8 · Session 2 divider
# ============================================================
s = div(prs, "Session 2", "📖 下午 45 min · 词汇 + 准备 展示", DAY, "📚"); n += 1; pn(s, n)

# ============================================================
# 9-13 · 我 会 认
# ============================================================
recognize_words = [
    ("🔮", "未来", "wèi lái", "Future", "未来 会 有 飞 行 汽车!", "Future will have flying cars!",
     "🔮 水晶 球 / 未来 城市", DAY),
    ("💡", "发明", "fā míng", "Invent", "我 想 发明 一 个 机器人。", "I want to invent a robot.",
     "💡 灯 泡 / 发明家", ORANGE),
    ("🤖", "机器人", "jī qì rén", "Robot", "机器人 会 帮 我们 做 事。", "Robots help us with tasks.",
     "🤖 各种 机器人", PURPLE),
    ("🌟", "梦想", "mèng xiǎng", "Dream / Aspiration", "我 的 梦想 是 当 科学家。", "My dream is to be a scientist.",
     "🌟 星空 / 想 实现 的 事", PINK),
    ("🔄", "改变", "gǎi biàn", "Change", "科技 改变 世界。", "Technology changes the world.",
     "🔄 转 换 / 蜕变", GREEN),
]
for em, cn, py, en, ex_cn, ex_en, hint, cl in recognize_words:
    s = vocab_recognize(prs, cl, em, cn, py, en, ex_cn, ex_en, hint)
    n += 1; pn(s, n)

# ============================================================
# 14-15 · 我 会 写
# ============================================================
s = vocab_write(prs, DAY, "未来", "Future",
                [("未", "wèi", "5 笔", "「木」 上 加 一 横 — 还没 到"),
                 ("来", "lái", "7 笔", "像 一 棵 树 加 「人」")])
n += 1; pn(s, n)

s = vocab_write(prs, PINK, "梦想", "Dream",
                [("梦", "mèng", "11 笔", "上 「林」 + 下 「夕」 — 林中 晚 上 做 梦"),
                 ("想", "xiǎng", "13 笔", "上 「相」 + 下 「心」 — 心 里 想")])
n += 1; pn(s, n)

# ============================================================
# 16 · Session 3 divider
# ============================================================
s = div(prs, "Session 3", "🎤 下午 90 min · 未来 科技 博览会 · Final Expo!", DAY, "🚀"); n += 1; pn(s, n)

# ============================================================
# 17 · Final Expo Overview
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🚀 未来 科技 博览会 · Future Tech Expo!", DAY)

intro = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(0.95), Inches(9.20), Inches(1.10))
intro.fill.solid(); intro.fill.fore_color.rgb = WARM
intro.line.color.rgb = DAY; intro.line.width = Pt(2.5)
tb(s, 0.55, 1.05, 9.00, 0.40, "🎯 你 是 未来 的 发明家 — 设计 一 个 「2050 年 的 新 科技」!",
   sz=14, b=True, c=DAY, a=PP_ALIGN.CENTER)
tb(s, 0.55, 1.45, 9.00, 0.30, "You're inventors of the future — design a 2050 technology!",
   sz=10, c=GRAY, a=PP_ALIGN.CENTER)
tb(s, 0.55, 1.78, 9.00, 0.28, "🎨 个 人 或 小组 · 60 min 准备 · 30 min 展示",
   sz=11, b=True, c=DARK, a=PP_ALIGN.CENTER)

# 4 expo deliverables
items = [
    ("✏️", "设计 图", "Design Sketch", "画 你 的 发明", CYBER),
    ("🛠️", "模型 / 道具", "Prototype", "用 材料 做 出 来", ORANGE),
    ("📝", "介绍 海 报", "Poster", "名字 + 用 处 + 怎么 工作", GREEN),
    ("🎤", "1 分钟 展示", "Pitch", "上 台 介绍 给 大家", PINK),
]
card_w = 2.20; gap = 0.10
total = 4*card_w + 3*gap; start = (10 - total)/2
for i, (em, cn, en, desc, cl) in enumerate(items):
    x = start + i*(card_w + gap)
    panel(s, x, 2.20, card_w, 2.60, cl, lw=2.5)
    tb(s, x, 2.35, card_w, 0.75, em, sz=48, a=PP_ALIGN.CENTER)
    tb(s, x, 3.15, card_w, 0.40, cn, sz=14, b=True, c=cl, a=PP_ALIGN.CENTER)
    tb(s, x, 3.55, card_w, 0.30, en, sz=10, c=GRAY, a=PP_ALIGN.CENTER)
    tb(s, x+0.10, 3.95, card_w-0.20, 0.65, desc, sz=10, b=True, c=DARK, a=PP_ALIGN.CENTER)

ex = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(5.00), Inches(9.20), Inches(0.55))
ex.fill.solid(); ex.fill.fore_color.rgb = DAY; ex.line.fill.background()
tb(s, 0.55, 5.07, 9.00, 0.32, "🏆 颁 奖: 最 大胆 · 最 有 用 · 最 好 玩 · 最 解决 问题",
   sz=12, b=True, c=STAR, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "60 分钟 准备:\n• 10 min — 选 主题 + 计划\n• 30 min — 制作 模型 + 海 报\n• 20 min — 排 练 1 分钟 介绍\n\n30 分钟 展示:\n• 每 人 / 小 组 1 分钟\n• 老师 / 同学 给 反馈\n• 颁 奖 + 拍 集体 照")

# ============================================================
# 18 · Pitch sentence frames
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🎤 1 分钟 介绍 — 句型  1-Minute Pitch", PURPLE)

panel(s, 0.40, 0.95, 9.20, 0.65, PURPLE, fill=WARM)
tb(s, 0.55, 1.05, 9.00, 0.35, "💬 用 这 些 句型 — 让 你 的 介绍 又 清 又 棒!",
   sz=13, b=True, c=PURPLE, a=PP_ALIGN.CENTER)
tb(s, 0.55, 1.40, 9.00, 0.28, "Use these sentence frames for a clear pitch!",
   sz=9, c=GRAY, a=PP_ALIGN.CENTER)

# K-2 frames
panel(s, 0.40, 1.75, 4.55, 3.30, ORANGE)
panel_head(s, 0.40, 1.75, 4.55, ORANGE, "📘 K-2 · 简单 句型  Simple Frames", sz=12)
k2 = [
    "1️⃣ 「我 发明 了 ___」",
    "2️⃣ 「它 帮 ___」",
    "3️⃣ 「它 在 ___ 时 用」",
    "4️⃣ 「我 觉得 它 很 ___」",
]
for i, line in enumerate(k2):
    tb(s, 0.55, 2.30+i*0.65, 4.30, 0.55, line, sz=13, b=True, c=DARK)

# 3-5 frames
panel(s, 5.05, 1.75, 4.55, 3.30, GREEN)
panel_head(s, 5.05, 1.75, 4.55, GREEN, "📗 3-5 · 完整 介绍  Full Pitch", sz=12)
g35 = [
    "1️⃣ 「问题: ___」",
    "2️⃣ 「我 的 发明 叫 ___」",
    "3️⃣ 「它 怎么 工作: ___」",
    "4️⃣ 「它 解决 了 ___ 因为 ___」",
]
for i, line in enumerate(g35):
    tb(s, 5.20, 2.30+i*0.65, 4.30, 0.55, line, sz=12, b=True, c=DARK)

tip = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(5.20), Inches(9.20), Inches(0.40))
tip.fill.solid(); tip.fill.fore_color.rgb = DAY; tip.line.fill.background()
tb(s, 0.55, 5.25, 9.00, 0.30, "🎯 排 练 2-3 遍 — 上 台 就 不 紧张 了!  Practice 2-3 times!",
   sz=11, b=True, c=WHITE, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)

# ============================================================
# 19 · 毕业 + 颁奖
# ============================================================
s = ns(prs); bg(s, INK, prs)
import random
random.seed(123)
for _ in range(50):
    x = random.uniform(0.2, 9.7); y = random.uniform(0.2, 5.4); sz = random.choice([0.06, 0.10])
    d = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(sz), Inches(sz))
    d.fill.solid(); d.fill.fore_color.rgb = STAR; d.line.fill.background()

tb(s, 0.3, 0.25, 9.4, 0.55, "🎓 毕业 啦!  Congrats, Future Inventors!", sz=20, b=True, c=STAR, a=PP_ALIGN.CENTER)

# Certificate
cert = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.00), Inches(1.10), Inches(8.00), Inches(2.65))
cert.fill.solid(); cert.fill.fore_color.rgb = CREAM
cert.line.color.rgb = STAR; cert.line.width = Pt(4)
tb(s, 1.00, 1.25, 8.00, 0.40, "🏅 「玩 转 创新 科技」 结业 证书",
   sz=18, b=True, c=AI_PURPLE, a=PP_ALIGN.CENTER)
tb(s, 1.00, 1.70, 8.00, 0.32, "Certificate · Innovative Tech Explorer",
   sz=11, c=GRAY, a=PP_ALIGN.CENTER)
tb(s, 1.30, 2.15, 7.40, 0.36, "恭喜  ____________________________",
   sz=14, b=True, c=DARK)
tb(s, 1.30, 2.55, 7.40, 0.32, "完成 5 天 科技 之 旅 — AI · 3D 打印 · ML · 科技 生活 · 未来!",
   sz=11, c=DARK)
tb(s, 1.30, 2.92, 7.40, 0.30, "Completed 5-day tech journey",
   sz=9, c=GRAY)
tb(s, 1.30, 3.35, 7.40, 0.32, "日期 Date: __________   老师 Teacher: __________",
   sz=11, c=DARK)

# Share box
share = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.50), Inches(4.00), Inches(9.00), Inches(1.20))
share.fill.solid(); share.fill.fore_color.rgb = AI_PURPLE
share.line.color.rgb = STAR; share.line.width = Pt(2.5)
tb(s, 0.65, 4.08, 8.70, 0.32, "🎤 一 句 话 总结  One-line Reflection",
   sz=14, b=True, c=STAR, a=PP_ALIGN.CENTER)
tb(s, 0.65, 4.45, 8.70, 0.35, "「我 学 到 了 ___, 我 的 梦想 是 ___」",
   sz=14, b=True, c=WHITE, a=PP_ALIGN.CENTER)
tb(s, 0.65, 4.85, 8.70, 0.30, "I learned ___ · My dream is ___",
   sz=10, c=LGRAY, a=PP_ALIGN.CENTER)

tb(s, 0.3, 5.25, 9.4, 0.30, "👋 继续 玩 · 继续 创造!  Keep playing, keep creating!",
   sz=11, b=True, c=STAR, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "最 后 15 分钟:\n• 颁 「结业 证书」\n• 每 个 学生 一 句话 总结\n• 全 班 合 影\n• 鼓励 把 作品 带 回 家 给 父母 看")

out = os.path.join(os.path.dirname(__file__), "day5_future.pptx")
prs.save(out)
print(f"Saved {out}  ({len(prs.slides)} slides)")
