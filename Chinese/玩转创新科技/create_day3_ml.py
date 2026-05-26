#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
玩转 创新 科技 · Day 3: 机器 学习 (Machine Learning)
3-session classroom deck for K-5 Chinese immersion summer camp.

探究 问题: 电脑 怎么 学 习?

我 会 认: 机器 / 分类 / 数据 / 训练
我 会 写: 分类 / 训练

Structure (matches Day 1 / Day 2):
  Session 1 (11:00–11:45) — 认识 机器 学习  · What is ML?
  Session 2 (2:00–2:45)   — 复习 + 中文 词汇  · 我 会 认 / 我 会 写
  Session 3 (3:00–4:30)   — 训练 真 AI + 我 是 AI 训练师  · Project
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from _helpers import *
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

prs = make_presentation()
DAY = ML_GREEN
n = 0


def arrow(s, x, y, w=0.30, h=0.30, color=DAY):
    a = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(y), Inches(w), Inches(h))
    a.fill.solid(); a.fill.fore_color.rgb = color
    a.line.fill.background()
    return a


# ============================================================
# 1 · COVER
# ============================================================
cover(prs, 3, "Machine Learning", "机器 学习 · 电脑 怎么 学 习?",
      "🧠 📊 🐱 🤖 ✨", DAY,
      "电脑 怎么 学 习?",
      "How do computers learn?")
n += 1; pn(prs.slides[-1], n)
notes(prs.slides[-1], "Day 3 · 核心 概念: 看 例子 → 找 规律 → 猜 答案\n• Session 1: 认识 机器 学习 + 真 实 例子\n• Session 2: 中文 词汇 (我 会 认 / 我 会 写)\n• Session 3: 用 Teachable Machine 训练 真 AI + 项目")


# ============================================================
# 2 · SESSION 1 DIVIDER
# ============================================================
s = div(prs, "Session 1", "🌅 上午 11:00–11:45  ·  机器 学习 实验 室 · ML Lab",
        DAY, "🔬"); n += 1; pn(s, n)


# ============================================================
# 3 · LEARNING GOALS
# ============================================================
s = learning_goals(prs, DAY, [
    ("1️⃣", "通过 真 实 实验 体验 机器 学习",
     "Experience ML through a real experiment", CYBER),
    ("2️⃣", "知道 「数据 / 分类 / 特征 / 训练」 是 什么",
     "Learn: data, classify, features, training", ORANGE),
    ("3️⃣", "发现 AI 为什么 有 时 会 猜 错",
     "Discover why AI sometimes makes mistakes", PINK),
    ("4️⃣", "学 当 一 个 聪明 的 AI 训练 师",
     "Become a thoughtful AI trainer", PURPLE),
])
n += 1; pn(s, n)


# ============================================================
# 4 · LAB OPENING TITLE — 机器 学习 实验 室
# ============================================================
s = ns(prs); bg(s, INK, prs)
# Lab decoration: corner sparkles + beakers
for x, y in [(0.5, 0.5), (9.0, 0.5), (0.5, 4.8), (9.0, 4.8)]:
    d = s.shapes.add_shape(MSO_SHAPE.STAR_5_POINT, Inches(x), Inches(y), Inches(0.40), Inches(0.40))
    d.fill.solid(); d.fill.fore_color.rgb = STAR; d.line.fill.background()
# Top banner
tb(s, 0.3, 0.45, 9.4, 0.40, "🔬 Session 1 · Machine Learning Lab",
   sz=16, b=True, c=NEON, a=PP_ALIGN.CENTER)
# Big title
title_box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.05), Inches(8.8), Inches(2.20))
title_box.fill.solid(); title_box.fill.fore_color.rgb = DAY
title_box.line.color.rgb = STAR; title_box.line.width = Pt(4)
tb(s, 0.8, 1.25, 8.4, 0.85, "机器 学习 实验 室",
   sz=46, b=True, c=WHITE, a=PP_ALIGN.CENTER)
tb(s, 0.8, 2.15, 8.4, 0.50, "我 是 AI 训练 师",
   sz=22, b=True, c=STAR, a=PP_ALIGN.CENTER)
tb(s, 0.8, 2.75, 8.4, 0.40, "Machine Learning Lab · I'm an AI Trainer!",
   sz=13, c=WARM, a=PP_ALIGN.CENTER)
# Lab icons row
tb(s, 0.3, 3.65, 9.4, 1.10, "🔬   🧪   🤖   💻   ✨",
   sz=58, a=PP_ALIGN.CENTER)
# Bottom hype
tb(s, 0.3, 5.00, 9.4, 0.40, "今天 我们 一起 做 真 实 验, 训练 真 AI!",
   sz=14, b=True, c=STAR, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "30 秒 hype slide.\n• 让 学生 大 声 喊 「我 是 AI 训练 师!」\n• 今天 重点: 不 是 听 讲 — 是 做 实验\n• 老师 准备: 笔记本 + 投影 + 摄像头 + Teachable Machine 已 打开")


# ============================================================
# 5 · HOOK STORY — Baby Learning
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🍼 小 baby 怎么 学 习? · How Does a Baby Learn?", DAY)

# Top story panel
story = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(0.95), Inches(9.20), Inches(2.05))
story.fill.solid(); story.fill.fore_color.rgb = WARM
story.line.color.rgb = DAY; story.line.width = Pt(2.5)
# Baby + parent emojis
tb(s, 0.55, 1.05, 1.20, 1.20, "👶", sz=64, a=PP_ALIGN.CENTER)
# Parent teaching dialogue
tb(s, 1.80, 1.05, 7.65, 0.35, "爸爸 妈妈 教 小 baby:",
   sz=13, b=True, c=DAY)
parent_lines = [
    ("🐱", "「这 是 猫。」"),
    ("🐶", "「这 是 狗。」"),
    ("🍎", "「这 是 苹果。」"),
    ("🍌", "「这 是 香蕉。」"),
]
for i, (em, line) in enumerate(parent_lines):
    col = i % 2; row = i // 2
    x = 1.85 + col * 3.85
    y = 1.40 + row * 0.62
    tb(s, x, y, 0.40, 0.40, em, sz=22, a=PP_ALIGN.LEFT)
    tb(s, x+0.45, y+0.05, 3.30, 0.40, line, sz=14, b=True, c=DARK)
tb(s, 1.80, 2.65, 7.65, 0.30, "✨ 看 很 多 次 — baby 学 会 了!",
   sz=12, b=True, c=DAY, a=PP_ALIGN.LEFT)

# Big question + 4 idea bubbles
q = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(3.20), Inches(9.20), Inches(0.55))
q.fill.solid(); q.fill.fore_color.rgb = DAY; q.line.fill.background()
tb(s, 0.55, 3.30, 9.0, 0.40, "🤔 小 baby 是 怎么 学 会 的?  How did baby learn?",
   sz=15, b=True, c=WHITE, a=PP_ALIGN.CENTER)

ideas = [
    ("👀", "看 很 多 次"),
    ("👨‍👩‍👧", "别 人 教"),
    ("🧠", "记 住 了"),
    ("🔍", "找 规律"),
]
idea_w = 2.10; gap = 0.15
total = 4*idea_w + 3*gap; start = (10 - total)/2
for i, (em, txt) in enumerate(ideas):
    x = start + i*(idea_w + gap)
    panel(s, x, 3.95, idea_w, 0.85, ORANGE, fill=WHITE, lw=2.5)
    tb(s, x, 4.02, idea_w, 0.40, em, sz=22, a=PP_ALIGN.CENTER)
    tb(s, x, 4.42, idea_w, 0.32, txt, sz=12, b=True, c=DARK, a=PP_ALIGN.CENTER)

# Bottom takeaway
tk = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(4.95), Inches(9.20), Inches(0.45))
tk.fill.solid(); tk.fill.fore_color.rgb = DAY
tk.line.color.rgb = STAR; tk.line.width = Pt(2)
tb(s, 0.55, 5.00, 9.0, 0.35, "💡 电脑 也 是 这 样 学 的!  Computers learn the same way!",
   sz=13, b=True, c=STAR, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "5 分钟 故事 引入:\n• 老师 用 讲 故事 的 语气 慢 慢 说\n• 问: 小 baby 是 怎么 学 会 的?\n• 让 学生 想 + 答 — 别 急 着 给 答案\n• 引出: 电脑 也 是 看 很 多 例子 学 — 叫 「机器 学习」")


# ============================================================
# 6 · BIG CONCEPT
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "💡 机器 学习 = ... · ML = ...", DAY)

# 3-step big visual
steps = [
    ("📚", "看 很 多 例子", "See lots of examples", CYBER),
    ("🔍", "找 规律", "Find patterns", ORANGE),
    ("✨", "学 会 判断", "Make predictions", PINK),
]
card_w = 2.65; gap = 0.50
total = 3*card_w + 2*gap; start = (10 - total)/2
for i, (em, cn, en, cl) in enumerate(steps):
    x = start + i*(card_w + gap)
    panel(s, x, 1.20, card_w, 3.20, cl, fill=WHITE, lw=3)
    tb(s, x, 1.40, card_w, 1.10, em, sz=68, a=PP_ALIGN.CENTER)
    tb(s, x, 2.65, card_w, 0.55, cn, sz=22, b=True, c=cl, a=PP_ALIGN.CENTER)
    tb(s, x, 3.25, card_w, 0.30, en, sz=11, c=GRAY, a=PP_ALIGN.CENTER)
    # big number
    tb(s, x+0.10, 3.70, card_w-0.20, 0.55, str(i+1), sz=44, b=True, c=cl, a=PP_ALIGN.CENTER)
    if i < 2:
        arrow(s, x + card_w + 0.10, 2.70, w=0.30, h=0.35, color=DAY)

# Big takeaway bar
tb_box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(4.65), Inches(9.20), Inches(0.75))
tb_box.fill.solid(); tb_box.fill.fore_color.rgb = DAY
tb_box.line.color.rgb = STAR; tb_box.line.width = Pt(3)
tb(s, 0.55, 4.75, 9.0, 0.40, "机器 学习 = 看 例子 → 找 规律 → 学 判断",
   sz=17, b=True, c=WHITE, a=PP_ALIGN.CENTER)
tb(s, 0.55, 5.15, 9.0, 0.28, "ML = examples → patterns → predictions",
   sz=10, c=WARM, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "2-3 分钟:\n• 反复 读 3 遍 — 让 学生 一起 说\n• 配 手势: 「看」 (指 眼睛) → 「找」 (指 头脑) → 「判断」 (指 嘴 / 答)\n• 这 是 今天 的 核心 概念!")


# ============================================================
# 7 · WHAT IS DATA?
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🍴 AI 吃 什么 长 大? · What Does AI Eat?", DAY)

# Setup analogy
setup = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(0.95), Inches(4.40), Inches(1.20))
setup.fill.solid(); setup.fill.fore_color.rgb = WARM
setup.line.color.rgb = DAY; setup.line.width = Pt(2.5)
tb(s, 0.55, 1.05, 4.10, 0.45, "👶 小 baby 需要 ...",
   sz=15, b=True, c=DAY, a=PP_ALIGN.LEFT)
tb(s, 0.55, 1.55, 4.10, 0.50, "🍼 食物!  Food!",
   sz=22, b=True, c=DAY, a=PP_ALIGN.LEFT)

# Reveal
reveal = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.20), Inches(0.95), Inches(4.40), Inches(1.20))
reveal.fill.solid(); reveal.fill.fore_color.rgb = DAY
reveal.line.color.rgb = STAR; reveal.line.width = Pt(2.5)
tb(s, 5.35, 1.05, 4.10, 0.45, "🤖 AI 需要 ...",
   sz=15, b=True, c=STAR, a=PP_ALIGN.LEFT)
tb(s, 5.35, 1.55, 4.10, 0.50, "📊 数据!  Data!",
   sz=22, b=True, c=WHITE, a=PP_ALIGN.LEFT)

# Definition
defn = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(2.30), Inches(9.20), Inches(0.65))
defn.fill.solid(); defn.fill.fore_color.rgb = WHITE
defn.line.color.rgb = DAY; defn.line.width = Pt(2.5)
tb(s, 0.55, 2.40, 9.0, 0.45, "📚 数据 = AI 的 学习 材料  Data = AI's learning material",
   sz=15, b=True, c=DAY, a=PP_ALIGN.CENTER)

# Examples grid
tb(s, 0.4, 3.10, 9.2, 0.30, "例子 / Examples:",
   sz=11, b=True, c=DARK, a=PP_ALIGN.CENTER)
data_types = [
    ("🐱", "猫 照片"),
    ("🐶", "狗 照片"),
    ("🍎", "水果 图片"),
    ("😊", "表情 图片"),
    ("♻️", "垃圾 图片"),
]
dw = 1.70; dgap = 0.15
dtotal = 5*dw + 4*dgap; dstart = (10 - dtotal)/2
for i, (em, cn) in enumerate(data_types):
    x = dstart + i*(dw + dgap)
    panel(s, x, 3.50, dw, 1.55, ORANGE, fill=WHITE, lw=2.5)
    tb(s, x, 3.62, dw, 0.75, em, sz=38, a=PP_ALIGN.CENTER)
    tb(s, x, 4.45, dw, 0.45, cn, sz=12, b=True, c=ORANGE, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "3 分钟:\n• 问: 小 baby 需要 什么? → 食物!\n• 那 AI 呢? → 数据!\n• 数据 不 是 食物 — 是 图片 / 文字 / 声音 等 学习 材料\n• 给 学生 看 5 种 数据 例子 — 让 他们 想 还 有 什么?")


# ============================================================
# 8 · CLASSIFICATION
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "📊 AI 最 会 做 什么? · What's AI's Specialty?", DAY)

# Big answer reveal
ans = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(0.95), Inches(8.0), Inches(1.00))
ans.fill.solid(); ans.fill.fore_color.rgb = DAY
ans.line.color.rgb = STAR; ans.line.width = Pt(3)
tb(s, 1.1, 1.05, 7.8, 0.50, "🗂️ 分 类!  Classification!",
   sz=26, b=True, c=WHITE, a=PP_ALIGN.CENTER)
tb(s, 1.1, 1.55, 7.8, 0.30, "AI 把 东西 分 到 不 同 的 组",
   sz=11, c=WARM, a=PP_ALIGN.CENTER)

# 4 examples in 2x2 grid
tb(s, 0.4, 2.10, 9.2, 0.30, "例子 / Examples:",
   sz=12, b=True, c=DARK, a=PP_ALIGN.CENTER)

examples = [
    ("🐾", "动物 分类", "Animal classification", "猫? 狗? 鸟?", CYBER),
    ("🍎", "水果 分类", "Fruit classification", "苹果? 香蕉? 橙子?", ORANGE),
    ("♻️", "垃圾 分类", "Trash sorting", "可 回收? 厨余?", DAY),
    ("😊", "表情 分类", "Emotion classify", "开心? 难过? 生气?", PINK),
]
cw = 4.40; cgap_x = 0.20; cgap_y = 0.15
cstart_x = (10 - 2*cw - cgap_x)/2
for i, (em, cn_t, en_t, ex, cl) in enumerate(examples):
    row = i // 2; col = i % 2
    x = cstart_x + col*(cw + cgap_x)
    y = 2.55 + row*(1.20 + cgap_y)
    panel(s, x, y, cw, 1.20, cl, fill=WHITE, lw=2.5)
    tb(s, x+0.15, y+0.15, 0.95, 0.90, em, sz=44, a=PP_ALIGN.LEFT)
    tb(s, x+1.20, y+0.18, cw-1.40, 0.35, cn_t, sz=15, b=True, c=cl)
    tb(s, x+1.20, y+0.52, cw-1.40, 0.25, en_t, sz=9, c=GRAY)
    tb(s, x+1.20, y+0.78, cw-1.40, 0.35, ex, sz=11, b=True, c=DARK)

# Discussion prompt
disc = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(5.10), Inches(9.20), Inches(0.32))
disc.fill.solid(); disc.fill.fore_color.rgb = ORANGE; disc.line.fill.background()
tb(s, 0.55, 5.13, 9.0, 0.28, "💬 你 还 能 想 到 什么 分类? Can you think of more?",
   sz=11, b=True, c=WHITE, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "3-5 分钟 互动:\n• 问 学生 还 知 道 什么 分类 — 鼓励 创意\n• 例子: 颜色 分类 / 大 小 分类 / 形状 分类 / 字母 vs 数字\n• 高 年级: 介绍 「类别 (categories) / 标签 (labels)」 概念")


# ============================================================
# 9 · FEATURES — 特征
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "👀 AI 看 什么? · What Does AI Look For?", DAY)

# Question
tb(s, 0.4, 0.85, 9.2, 0.32, "我们 怎么 知道 这 是 猫? 这 是 狗?",
   sz=14, b=True, c=DARK, a=PP_ALIGN.CENTER)

# Cat vs Dog comparison
animals = [
    ("🐱", "猫 · Cat", CYBER, [
        ("👂", "尖 耳朵"),
        ("〰️", "胡须"),
        ("🌀", "尾巴 弯"),
        ("⚪", "身体 小"),
    ]),
    ("🐶", "狗 · Dog", ORANGE, [
        ("👃", "鼻子 长"),
        ("👂", "耳朵 软"),
        ("🐕", "身体 大"),
        ("🐾", "脚 印 大"),
    ]),
]
col_w = 4.30; gap = 0.40
total = 2*col_w + gap; start = (10 - total)/2
for i, (em, title, cl, feats) in enumerate(animals):
    x = start + i*(col_w + gap)
    panel(s, x, 1.25, col_w, 2.90, cl, fill=WHITE, lw=3)
    head = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.25), Inches(col_w), Inches(0.50))
    head.fill.solid(); head.fill.fore_color.rgb = cl; head.line.fill.background()
    tb(s, x, 1.33, col_w, 0.40, title, sz=15, b=True, c=WHITE, a=PP_ALIGN.CENTER)
    tb(s, x, 1.85, col_w, 0.80, em, sz=54, a=PP_ALIGN.CENTER)
    for j, (e, txt) in enumerate(feats):
        col_inner = j % 2; row_inner = j // 2
        bx = x + 0.20 + col_inner*((col_w-0.40)/2)
        by = 2.75 + row_inner*0.65
        tb(s, bx, by, 0.45, 0.45, e, sz=20, a=PP_ALIGN.LEFT)
        tb(s, bx+0.50, by+0.05, (col_w-0.40)/2 - 0.50, 0.40, txt, sz=12, b=True, c=DARK)

# Big concept callout
concept = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(4.30), Inches(9.20), Inches(1.10))
concept.fill.solid(); concept.fill.fore_color.rgb = DAY
concept.line.color.rgb = STAR; concept.line.width = Pt(3)
tb(s, 0.55, 4.45, 9.0, 0.40, "💡 这 些 小 线索 叫:",
   sz=14, b=True, c=STAR, a=PP_ALIGN.CENTER)
tb(s, 0.55, 4.85, 9.0, 0.55, "特 征  Features",
   sz=26, b=True, c=WHITE, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "5 分钟:\n• 问 学生: 怎么 一 眼 看 出 是 猫 还是 狗?\n• 让 学生 大 声 喊 出 「尖 耳朵」「胡须」「尾巴」 等\n• 把 这些 写 在 白 板 上\n• 解释: 这 些 帮助 判断 的 小 线索 = 特征\n• AI 也 是 看 特征 来 分类!")


# ============================================================
# 10 · TRANSITION TO EXPERIMENT
# ============================================================
s = ns(prs); bg(s, INK, prs)
for x, y in [(0.4, 0.45), (9.1, 0.5), (0.5, 4.8), (9.0, 4.7), (1.2, 0.6), (8.3, 4.9)]:
    d = s.shapes.add_shape(MSO_SHAPE.STAR_5_POINT, Inches(x), Inches(y), Inches(0.35), Inches(0.35))
    d.fill.solid(); d.fill.fore_color.rgb = STAR; d.line.fill.background()

tb(s, 0.3, 0.55, 9.4, 0.40, "🧪 实 验 时 间!",
   sz=18, b=True, c=NEON, a=PP_ALIGN.CENTER)

# Big question box
qb = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.15), Inches(8.8), Inches(1.65))
qb.fill.solid(); qb.fill.fore_color.rgb = DAY
qb.line.color.rgb = STAR; qb.line.width = Pt(4)
tb(s, 0.8, 1.30, 8.4, 0.70, "AI 能 分 得 清 谁 是 谁 吗?",
   sz=30, b=True, c=WHITE, a=PP_ALIGN.CENTER)
tb(s, 0.8, 2.00, 8.4, 0.40, "Can AI tell who's who?",
   sz=15, c=STAR, a=PP_ALIGN.CENTER)
tb(s, 0.8, 2.40, 8.4, 0.30, "今天 我们 来 训练 一 个 真 正 的 AI!",
   sz=13, c=WARM, a=PP_ALIGN.CENTER)

# Lab equipment row
tb(s, 0.3, 3.20, 9.4, 1.10, "📷   🤖   🧠   🧪   ✨",
   sz=58, a=PP_ALIGN.CENTER)

# TM banner
bn = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(4.50), Inches(8.8), Inches(0.95))
bn.fill.solid(); bn.fill.fore_color.rgb = WHITE
bn.line.color.rgb = STAR; bn.line.width = Pt(3)
tb(s, 0.75, 4.60, 8.5, 0.40, "💻 用 Teachable Machine — 一 起 训练 AI!",
   sz=15, b=True, c=DAY, a=PP_ALIGN.CENTER)
tb(s, 0.75, 5.05, 8.5, 0.30, "Let's train an AI together — live!",
   sz=11, c=GRAY, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "30 秒 hype slide.\n• 老师 切到 投影 — 显示 teachablemachine.withgoogle.com\n• 大 声 宣布: 「今天 我们 不 是 看 视频, 是 做 实验!」\n• 让 学生 鼓掌, 兴奋 起来")


# ============================================================
# 11 · CORE EXPERIMENT SETUP — 大咪 vs 小咪
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🐱 大 咪 vs 小 咪 · Cat A vs Cat B", DAY)

# Top description
tb(s, 0.4, 0.85, 9.2, 0.32, "我们 要 训练 AI 认 两 只 长 得 像 的 猫!",
   sz=14, b=True, c=DARK, a=PP_ALIGN.CENTER)

# Two photo placeholders side by side
photo_slot(s, 0.50, 1.30, 4.30, 2.70, "📷 大 咪 照片",
           "Tabby + white paws", DAY)
photo_slot(s, 5.20, 1.30, 4.30, 2.70, "📷 小 咪 照片",
           "Tabby, NO white paws", ORANGE)

# Cat descriptions below photos
desc1 = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.50), Inches(4.05), Inches(4.30), Inches(0.55))
desc1.fill.solid(); desc1.fill.fore_color.rgb = DAY; desc1.line.fill.background()
tb(s, 0.55, 4.10, 4.20, 0.30, "🐱 大 咪 Da-mi",
   sz=13, b=True, c=STAR, a=PP_ALIGN.CENTER)
tb(s, 0.55, 4.38, 4.20, 0.22, "虎斑 + 白色 脚 印 + 特别 花纹",
   sz=9, c=WARM, a=PP_ALIGN.CENTER)

desc2 = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.20), Inches(4.05), Inches(4.30), Inches(0.55))
desc2.fill.solid(); desc2.fill.fore_color.rgb = ORANGE; desc2.line.fill.background()
tb(s, 5.25, 4.10, 4.20, 0.30, "🐱 小 咪 Xiao-mi",
   sz=13, b=True, c=STAR, a=PP_ALIGN.CENTER)
tb(s, 5.25, 4.38, 4.20, 0.22, "虎斑 + 没 有 白 脚 + 灰白 身 体",
   sz=9, c=WARM, a=PP_ALIGN.CENTER)

# Big prediction question
q = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(4.75), Inches(9.20), Inches(0.65))
q.fill.solid(); q.fill.fore_color.rgb = INK
q.line.color.rgb = STAR; q.line.width = Pt(2)
tb(s, 0.55, 4.82, 9.0, 0.32, "🤔 AI 能 分 得 清 大 咪 和 小 咪 吗? 投 票!",
   sz=13, b=True, c=STAR, a=PP_ALIGN.CENTER)
tb(s, 0.55, 5.13, 9.0, 0.26, "Will AI tell them apart? Vote: 能 / 不 能 / 不 知 道",
   sz=9, c=LGRAY, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "3 分钟 setup:\n• 老师 提前 拍 好 2 只 课堂 猫 (或 用 学生 自带 玩 偶) — 备 好 照片\n• 也 可以 用 老师 家 的 真 猫 照片\n• 重点: 两 只 猫 要 「长 得 像 但 有 差别」 (e.g., 都 是 虎 斑, 一 只 有 白 脚 一 只 没 有)\n• 让 学生 先 看 + 比较 + 找 特征\n• 然后 投票 预测 — 制造 悬念!")


# ============================================================
# 12 · EXPERIMENT ROUND 1 — Few photos
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🔬 实验 第 1 轮 · Round 1: Few Photos", PINK)

# Setup card
setup = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(0.95), Inches(9.20), Inches(0.75))
setup.fill.solid(); setup.fill.fore_color.rgb = WARM
setup.line.color.rgb = PINK; setup.line.width = Pt(2.5)
tb(s, 0.55, 1.02, 9.0, 0.32, "📷 训练 数据: 每 只 猫 只 拍 1-2 张!",
   sz=15, b=True, c=PINK, a=PP_ALIGN.CENTER)
tb(s, 0.55, 1.38, 9.0, 0.28, "Train data: only 1-2 photos per cat",
   sz=10, c=GRAY, a=PP_ALIGN.CENTER)

# Visual: very few photos
panel(s, 0.40, 1.85, 4.40, 2.30, DAY, fill=WHITE, lw=2.5)
panel_head(s, 0.40, 1.85, 4.40, DAY, "🐱 大 咪 — 只 有 1 张", sz=12)
tb(s, 0.50, 2.45, 4.20, 1.60, "📷", sz=110, a=PP_ALIGN.CENTER)

panel(s, 5.20, 1.85, 4.40, 2.30, ORANGE, fill=WHITE, lw=2.5)
panel_head(s, 5.20, 1.85, 4.40, ORANGE, "🐱 小 咪 — 只 有 1 张", sz=12)
tb(s, 5.30, 2.45, 4.20, 1.60, "📷", sz=110, a=PP_ALIGN.CENTER)

# Live test prompt — shorter box so text fits within slide bounds
test = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(4.40), Inches(9.20), Inches(1.00))
test.fill.solid(); test.fill.fore_color.rgb = PINK
test.line.color.rgb = STAR; test.line.width = Pt(3)
tb(s, 0.55, 4.50, 9.0, 0.45, "🧪 现场 测试! AI 会 猜 对 吗?",
   sz=20, b=True, c=WHITE, a=PP_ALIGN.CENTER)
tb(s, 0.55, 5.00, 9.0, 0.30, "Live test! 👉 一起 投 票: 能 / 不 能",
   sz=12, b=True, c=STAR, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "5-7 分钟 现场 演示 (核心 体验!):\n• 老师 在 Teachable Machine 上 只 拍 1-2 张 大 咪 + 1-2 张 小 咪\n• 点 Train → 立 即 测试\n• 给 它 看 大 咪 新 照片 → 看 AI 怎么 说\n• 多 半 会 错! 因为 数据 太 少\n• 让 学生 大 声 喊 出 AI 的 预测\n• 故意 制造 错误 — 这 是 下 一 张 的 讨论 引子")


# ============================================================
# 13 · DISCUSSION — Why AI got it wrong
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🤔 为什么 AI 猜 错 了? · Why Did AI Get It Wrong?", PINK)

tb(s, 0.4, 0.85, 9.2, 0.32, "想 一 想 — AI 为什么 分 不 清?",
   sz=14, b=True, c=DARK, a=PP_ALIGN.CENTER)

# 4 reason cards in 2x2
reasons = [
    ("📷", "图片 太 少", "Too few images", CYBER),
    ("📚", "学 得 不 够", "Didn't learn enough", ORANGE),
    ("🔍", "不 知 道 特征", "Doesn't know features", PURPLE),
    ("🐱", "长 得 太 像", "Cats look too similar", PINK),
]
card_w = 4.30; card_h = 1.65; gap_x = 0.25; gap_y = 0.20
start_x = (10 - 2*card_w - gap_x)/2
for i, (em, cn, en, cl) in enumerate(reasons):
    row = i // 2; col = i % 2
    x = start_x + col*(card_w + gap_x)
    y = 1.30 + row*(card_h + gap_y)
    panel(s, x, y, card_w, card_h, cl, fill=WHITE, lw=3)
    tb(s, x+0.20, y+0.30, 0.90, 0.90, em, sz=42, a=PP_ALIGN.LEFT)
    tb(s, x+1.25, y+0.35, card_w-1.40, 0.50, cn, sz=18, b=True, c=cl)
    tb(s, x+1.25, y+0.90, card_w-1.40, 0.40, en, sz=10, c=GRAY)

# Bottom prompt
tb(s, 0.4, 5.10, 9.2, 0.30, "💬 小 组 讨 论 — 还 有 其他 原因 吗?",
   sz=11, b=True, c=PINK, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "3-5 分钟 讨论:\n• 让 学生 大 声 说 出 想 法\n• 引导 — 把 答案 落 到 「数据 少」 + 「特征 难 抓」\n• 这 是 关键 教学 时刻: AI 不 是 不 聪明, 是 「学 习 材料」 不 够")


# ============================================================
# 14 · EXPERIMENT ROUND 2 — More photos
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🔬 实验 第 2 轮 · Round 2: More Photos!", GREEN)

# Setup card
setup = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(0.95), Inches(9.20), Inches(0.75))
setup.fill.solid(); setup.fill.fore_color.rgb = WARM
setup.line.color.rgb = GREEN; setup.line.width = Pt(2.5)
tb(s, 0.55, 1.02, 9.0, 0.32, "📷 加 多 数据! 每 只 猫 5-10 张 — 不同 角度!",
   sz=15, b=True, c=GREEN, a=PP_ALIGN.CENTER)
tb(s, 0.55, 1.38, 9.0, 0.28, "Add more data! 5-10 photos per cat — different angles!",
   sz=10, c=GRAY, a=PP_ALIGN.CENTER)

# Variety tags
varieties = [
    ("📐", "不同 角度"),
    ("🤸", "不同 姿势"),
    ("💡", "不同 光线"),
    ("📏", "不同 距离"),
]
vw = 2.10; vgap = 0.15
vtotal = 4*vw + 3*vgap; vstart = (10 - vtotal)/2
for i, (em, txt) in enumerate(varieties):
    x = vstart + i*(vw + vgap)
    panel(s, x, 2.00, vw, 1.05, GREEN, fill=WHITE, lw=2.5)
    tb(s, x, 2.10, vw, 0.50, em, sz=26, a=PP_ALIGN.CENTER)
    tb(s, x, 2.62, vw, 0.40, txt, sz=12, b=True, c=GREEN, a=PP_ALIGN.CENTER)

# Many photos visual
tb(s, 0.4, 3.20, 9.2, 0.85, "📷 📷 📷 📷 📷  vs  📷 📷 📷 📷 📷",
   sz=36, a=PP_ALIGN.CENTER)
tb(s, 0.4, 4.05, 9.2, 0.30, "10 张 大 咪 + 10 张 小 咪",
   sz=12, b=True, c=DARK, a=PP_ALIGN.CENTER)

# Live retest prompt
retest = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(4.50), Inches(9.20), Inches(0.90))
retest.fill.solid(); retest.fill.fore_color.rgb = GREEN
retest.line.color.rgb = STAR; retest.line.width = Pt(3)
tb(s, 0.55, 4.62, 9.0, 0.40, "🧪 再 测 试! 现在 AI 会 更 聪 明 吗?",
   sz=18, b=True, c=WHITE, a=PP_ALIGN.CENTER)
tb(s, 0.55, 5.05, 9.0, 0.30, "Retest now! Will AI be smarter?",
   sz=11, c=STAR, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "5-7 分钟 现场 演示:\n• 老师 加 多 照片 — 拍 不同 角度 的 大 咪 + 小 咪\n• 让 学生 帮 忙 决定 拍 什么 (头 / 尾巴 / 全 身)\n• 重新 训练 → 测试\n• 这 次 应 该 答 对 大 部分\n• 让 学生 鼓掌!\n• 引出 下 一 张: 为什么 这 次 好 了?")


# ============================================================
# 15 · DISCUSSION — Why it got better
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "✨ 为什么 现在 AI 更 棒? · Why Did AI Improve?", GREEN)

# 3 reasons
reasons = [
    ("📚", "更 多 数据", "More data", CYBER),
    ("🔍", "学 到 更 多 特征", "Learned more features", ORANGE),
    ("👀", "看 过 更 多 例子", "Saw more examples", PINK),
]
card_w = 2.85; gap = 0.20
total = 3*card_w + 2*gap; start = (10 - total)/2
for i, (em, cn, en, cl) in enumerate(reasons):
    x = start + i*(card_w + gap)
    panel(s, x, 1.05, card_w, 2.85, cl, fill=WHITE, lw=3)
    tb(s, x, 1.20, card_w, 0.95, em, sz=58, a=PP_ALIGN.CENTER)
    tb(s, x, 2.25, card_w, 0.50, cn, sz=18, b=True, c=cl, a=PP_ALIGN.CENTER)
    tb(s, x, 2.78, card_w, 0.32, en, sz=11, c=GRAY, a=PP_ALIGN.CENTER)
    tb(s, x+0.10, 3.25, card_w-0.20, 0.55, str(i+1), sz=40, b=True, c=cl, a=PP_ALIGN.CENTER)

# Big takeaway
tk = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(4.20), Inches(9.20), Inches(1.20))
tk.fill.solid(); tk.fill.fore_color.rgb = GREEN
tk.line.color.rgb = STAR; tk.line.width = Pt(4)
tb(s, 0.55, 4.35, 9.0, 0.55, "💡 更 多 数据 = 更 好 学习!",
   sz=28, b=True, c=WHITE, a=PP_ALIGN.CENTER)
tb(s, 0.55, 4.95, 9.0, 0.35, "More data = better learning!",
   sz=14, b=True, c=STAR, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "3 分钟:\n• 让 学生 一起 喊 「更 多 数据 = 更 好 学习!」\n• 这 是 今天 最 重 要 的 一 句 话\n• 跟 第 5 张 (baby) 呼应: 看 多 了 就 学 会")


# ============================================================
# 16 · CHALLENGE ROUND — Tricky tests
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🎯 大 挑 战! · Challenge Round!", ORANGE)

tb(s, 0.4, 0.85, 9.2, 0.32, "用 「难」 的 图 测试 AI — 它 还 认 得 吗?",
   sz=14, b=True, c=DARK, a=PP_ALIGN.CENTER)
tb(s, 0.4, 1.18, 9.2, 0.26, "Test AI with tricky images — can it still tell?",
   sz=10, c=GRAY, a=PP_ALIGN.CENTER)

# 4 tricky challenges
tricky = [
    ("📐", "角度 奇 怪", "Unusual angle", "倒 着 拍?"),
    ("🌫️", "模糊 的 图", "Blurry photo", "看 不 清?"),
    ("🐈", "长 得 像 的 别 的 猫", "Similar other cat", "不 是 大 咪 / 小 咪!"),
    ("🧸", "玩 具 猫", "Toy cat", "是 真 的 吗?"),
]
card_w = 2.15; gap = 0.20
total = 4*card_w + 3*gap; start = (10 - total)/2
for i, (em, cn, en, ex) in enumerate(tricky):
    x = start + i*(card_w + gap)
    panel(s, x, 1.65, card_w, 2.85, ORANGE, fill=WHITE, lw=3)
    tb(s, x, 1.78, card_w, 0.85, em, sz=44, a=PP_ALIGN.CENTER)
    tb(s, x, 2.70, card_w, 0.40, cn, sz=14, b=True, c=ORANGE, a=PP_ALIGN.CENTER)
    tb(s, x, 3.12, card_w, 0.28, en, sz=9, c=GRAY, a=PP_ALIGN.CENTER)
    tb(s, x+0.08, 3.55, card_w-0.16, 0.80, ex, sz=11, b=True, c=DARK, a=PP_ALIGN.CENTER)

# Big question
qb = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(4.70), Inches(9.20), Inches(0.75))
qb.fill.solid(); qb.fill.fore_color.rgb = DAY
qb.line.color.rgb = STAR; qb.line.width = Pt(3)
tb(s, 0.55, 4.78, 9.0, 0.40, "🤔 AI 还 认 得 吗?",
   sz=20, b=True, c=WHITE, a=PP_ALIGN.CENTER)
tb(s, 0.55, 5.20, 9.0, 0.28, "Can AI still recognize?",
   sz=10, c=WARM, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "5 分钟 演示:\n• 用 这 4 类 「难」 图 测试 已 经 训练 好 的 AI\n• 多 半 又 会 错!\n• 让 学生 看 到: AI 不 是 万 能, 它 只 学 了 「正常」 的 照片\n• 引出 下 一 张: AI 也 会 犯错")


# ============================================================
# 17 · AI MAKES MISTAKES
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "😅 AI 也 会 犯 错! · AI Makes Mistakes!", PINK)

tb(s, 0.4, 0.85, 9.2, 0.32, "AI 可能 会 犯错, 因为:",
   sz=14, b=True, c=DARK, a=PP_ALIGN.CENTER)

# 4 mistake reasons
mistakes = [
    ("📉", "数据 太 少", "Too little data", CYBER),
    ("🔍", "特征 太 像", "Features too similar", ORANGE),
    ("🌫️", "图片 不 清 楚", "Image unclear", PURPLE),
    ("⚠️", "学习 不 完整", "Incomplete learning", PINK),
]
card_w = 2.15; gap = 0.20
total = 4*card_w + 3*gap; start = (10 - total)/2
for i, (em, cn, en, cl) in enumerate(mistakes):
    x = start + i*(card_w + gap)
    panel(s, x, 1.25, card_w, 2.85, cl, fill=WHITE, lw=3)
    tb(s, x, 1.38, card_w, 0.85, em, sz=44, a=PP_ALIGN.CENTER)
    tb(s, x, 2.30, card_w, 0.45, cn, sz=15, b=True, c=cl, a=PP_ALIGN.CENTER)
    tb(s, x, 2.78, card_w, 0.30, en, sz=10, c=GRAY, a=PP_ALIGN.CENTER)
    # Number badge
    badge = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x+card_w/2-0.30), Inches(3.35), Inches(0.60), Inches(0.60))
    badge.fill.solid(); badge.fill.fore_color.rgb = cl; badge.line.fill.background()
    tb(s, x+card_w/2-0.30, 3.40, 0.60, 0.50, str(i+1), sz=22, b=True, c=WHITE, a=PP_ALIGN.CENTER)

# Big message
mb = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(4.35), Inches(9.20), Inches(1.05))
mb.fill.solid(); mb.fill.fore_color.rgb = PINK
mb.line.color.rgb = STAR; mb.line.width = Pt(3)
tb(s, 0.55, 4.50, 9.0, 0.50, "🤖 AI 不 是 完美 的 — 它 在 学习!",
   sz=22, b=True, c=WHITE, a=PP_ALIGN.CENTER)
tb(s, 0.55, 5.05, 9.0, 0.30, "AI is not perfect — it's still learning!",
   sz=12, b=True, c=STAR, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "3 分钟:\n• 让 学生 明白: AI 也 会 错 是 「正常」 的\n• 关键: 我们 怎么 帮 AI 变 更 聪明? → 给 它 更 多 + 更 好 数据!\n• 高 年级: 可以 引入 「测试 数据 vs 训练 数据」 概念")


# ============================================================
# 18 · HIDDEN BIAS — Extension for older students
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🌗 高 年级 加 餐 · Hidden Bias", PURPLE)

tb(s, 0.4, 0.85, 9.2, 0.32, "如果 AI 看 过 的 照片 不 「完 整」 怎么 办?",
   sz=14, b=True, c=DARK, a=PP_ALIGN.CENTER)
tb(s, 0.4, 1.18, 9.2, 0.26, "What if AI only saw certain kinds of photos?",
   sz=10, c=GRAY, a=PP_ALIGN.CENTER)

# 2 scenario panels — each has a "train condition" + "test condition"
scenarios = [
    ("☀️", "只 学 过 白天 的 猫", "Only daytime cats",
     "🌙", "晚 上 怎么 办?", "What about at night?", CYBER),
    ("🪑", "只 见 过 坐 着 的 猫", "Only sitting cats",
     "🏃", "站 着 / 跑 着 怎么 办?", "Standing / running?", ORANGE),
]
card_w = 4.40; gap = 0.40
total = 2*card_w + gap; start = (10 - total)/2
for i, (em1, cn1, en1, em2, cn2, en2, cl) in enumerate(scenarios):
    x = start + i*(card_w + gap)
    panel(s, x, 1.55, card_w, 2.95, cl, fill=WHITE, lw=3)
    # Top half: training
    tb(s, x+0.15, 1.70, 0.70, 0.65, em1, sz=32, a=PP_ALIGN.LEFT)
    tb(s, x+0.95, 1.78, card_w-1.10, 0.30, "训练 时:", sz=10, b=True, c=GRAY)
    tb(s, x+0.95, 2.05, card_w-1.10, 0.40, cn1, sz=14, b=True, c=cl)
    tb(s, x+0.95, 2.40, card_w-1.10, 0.25, en1, sz=9, c=GRAY)
    # Divider
    div_line = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x+0.30), Inches(2.85), Inches(card_w-0.60), Inches(0.03))
    div_line.fill.solid(); div_line.fill.fore_color.rgb = cl; div_line.line.fill.background()
    # Bottom half: test
    tb(s, x+0.15, 3.05, 0.70, 0.65, em2, sz=32, a=PP_ALIGN.LEFT)
    tb(s, x+0.95, 3.13, card_w-1.10, 0.30, "测试 时:", sz=10, b=True, c=GRAY)
    tb(s, x+0.95, 3.40, card_w-1.10, 0.40, cn2, sz=14, b=True, c=cl)
    tb(s, x+0.95, 3.75, card_w-1.10, 0.25, en2, sz=9, c=GRAY)
    # Question mark — inset from edge so it doesn't touch the card border
    tb(s, x+card_w-0.85, 4.05, 0.55, 0.35, "❓", sz=22, a=PP_ALIGN.CENTER)

# Concept callout
concept = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(4.70), Inches(9.20), Inches(0.70))
concept.fill.solid(); concept.fill.fore_color.rgb = PURPLE
concept.line.color.rgb = STAR; concept.line.width = Pt(2.5)
tb(s, 0.55, 4.80, 9.0, 0.35, "💡 学 习 材料 不 完整 = AI 不 公平!",
   sz=14, b=True, c=WHITE, a=PP_ALIGN.CENTER)
tb(s, 0.55, 5.15, 9.0, 0.28, "Incomplete data = unfair AI",
   sz=10, c=STAR, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "5 分钟 (高 年级 适用; K-2 可 跳 过):\n• 这 是 AI ethics 的 入门\n• 引导: 想象 AI 学 认 人 — 但 只 学 过 一 种 人?\n• 真 实 例子: 早期 人脸 识 别 AI 对 深色 皮 肤 不 公平 — 因为 训练 数据 不 多 样\n• 高 年级 可以 讨论: 怎么 让 AI 公平?")


# ============================================================
# 19 · SMALL GROUP REFLECTION
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "💬 小 组 讨 论 · Group Discussion", DAY)

# Big question
q = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(0.95), Inches(9.20), Inches(1.15))
q.fill.solid(); q.fill.fore_color.rgb = DAY
q.line.color.rgb = STAR; q.line.width = Pt(3)
tb(s, 0.55, 1.10, 9.0, 0.50, "🤔 如果 你 要 训练 AI, 你 会 给 它 什么 样 的 图片?",
   sz=18, b=True, c=WHITE, a=PP_ALIGN.CENTER)
tb(s, 0.55, 1.65, 9.0, 0.32, "If YOU train an AI — what photos would you give it?",
   sz=11, c=WARM, a=PP_ALIGN.CENTER)

# 5 idea bubbles in row
ideas = [
    ("📐", "不同 角度", "Different angles"),
    ("📷", "清楚 的 图", "Clear photos"),
    ("📏", "远 + 近", "Near + far"),
    ("🌈", "不同 背景", "Different backgrounds"),
    ("💯", "很 多 例子", "Lots of examples"),
]
iw = 1.75; igap = 0.10
itotal = 5*iw + 4*igap; istart = (10 - itotal)/2
for i, (em, cn, en) in enumerate(ideas):
    x = istart + i*(iw + igap)
    panel(s, x, 2.40, iw, 2.20, ORANGE, fill=WHITE, lw=2.5)
    tb(s, x, 2.55, iw, 0.75, em, sz=38, a=PP_ALIGN.CENTER)
    tb(s, x, 3.40, iw, 0.50, cn, sz=12, b=True, c=ORANGE, a=PP_ALIGN.CENTER)
    tb(s, x+0.05, 3.95, iw-0.10, 0.35, en, sz=9, c=GRAY, a=PP_ALIGN.CENTER)

# Group discussion prompt
disc = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(4.80), Inches(9.20), Inches(0.60))
disc.fill.solid(); disc.fill.fore_color.rgb = INK; disc.line.fill.background()
tb(s, 0.55, 4.88, 9.0, 0.32, "👥 4 人 一 组 — 3 分钟 讨论 后 分 享!",
   sz=13, b=True, c=STAR, a=PP_ALIGN.CENTER)
tb(s, 0.55, 5.18, 9.0, 0.22, "Groups of 4 — discuss 3 min, then share!",
   sz=9, c=LGRAY, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "5 分钟:\n• 小 组 讨论 — 让 学生 自己 想 答案 (不 提示)\n• 然后 分 享 — 老师 收 集 答 案 写 白 板 上\n• 这 是 「学 当 一 个 AI 训练 师」 的 关键 时刻\n• 让 学生 感受 自己 是 设 计 者 不 是 用 户")


# ============================================================
# 20 · WRAP-UP
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🎓 总 结 · Wrap-up", DAY)

# 6 key concepts in 2 rows of 3
concepts = [
    ("📊", "数 据", "Data", CYBER),
    ("🗂️", "分 类", "Classify", ORANGE),
    ("👀", "特 征", "Features", PURPLE),
    ("🏋️", "训 练", "Train", DAY),
    ("🧪", "测 试", "Test", PINK),
    ("😅", "错 误", "Errors", GREEN),
]
card_w = 2.85; card_h = 1.30; gap_x = 0.15; gap_y = 0.15
start_x = (10 - 3*card_w - 2*gap_x)/2
for i, (em, cn, en, cl) in enumerate(concepts):
    row = i // 3; col = i % 3
    x = start_x + col*(card_w + gap_x)
    y = 0.95 + row*(card_h + gap_y)
    panel(s, x, y, card_w, card_h, cl, fill=WHITE, lw=2.5)
    tb(s, x+0.10, y+0.20, 0.85, 0.90, em, sz=42, a=PP_ALIGN.LEFT)
    tb(s, x+1.00, y+0.20, card_w-1.10, 0.50, cn, sz=18, b=True, c=cl)
    tb(s, x+1.00, y+0.75, card_w-1.10, 0.35, en, sz=11, c=GRAY)

# Final takeaway
fin = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(3.85), Inches(9.20), Inches(1.55))
fin.fill.solid(); fin.fill.fore_color.rgb = DAY
fin.line.color.rgb = STAR; fin.line.width = Pt(4)
tb(s, 0.55, 4.00, 9.0, 0.50, "💡 机器 学习 不 是 魔法!",
   sz=24, b=True, c=WHITE, a=PP_ALIGN.CENTER)
tb(s, 0.55, 4.55, 9.0, 0.50, "是 看 很 多 例子, 找 到 规律!",
   sz=22, b=True, c=STAR, a=PP_ALIGN.CENTER)
tb(s, 0.55, 5.10, 9.0, 0.30, "ML is not magic — it's seeing examples + finding patterns!",
   sz=11, c=WARM, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "3 分钟 总 结:\n• 一起 复习 6 个 关键 词: 数据 / 分类 / 特征 / 训练 / 测试 / 错误\n• 全 班 一起 大 声 喊: 「机器 学习 不 是 魔法 — 是 看 很 多 例子 找 规律!」\n• 下午 Session 2 — 学 中文 词汇 让 你 更 像 真 AI 工程师\n• 下午 Session 3 — 你们 自己 来 训练 AI!")


# ============================================================
# 9 · SESSION 2 DIVIDER
# ============================================================
s = div(prs, "Session 2", "📖 下午 2:00–2:45  ·  复习 + 中文 词汇 · 我 会 认 / 我 会 写",
        DAY, "📚"); n += 1; pn(s, n)


# ============================================================
# 10-13 · 我 会 认 (vocabulary recognition)
# ============================================================
recognize_words = [
    ("🤖", "机器", "jī qì", "Machine",
     "机器 人 帮 我们 做 事。", "Robots help us do things.",
     "机器 人 / 机械 手 / 工厂 机器", CYBER),
    ("🗂️", "分类", "fēn lèi", "Classify",
     "我们 给 图片 分类。", "We sort the pictures.",
     "整理 文件夹 / 分 类 物品", ORANGE),
    ("📊", "数据", "shù jù", "Data",
     "电脑 需要 很 多 数据 学 习。", "Computers need lots of data to learn.",
     "图表 / 数字 / 表 格", PINK),
    ("🏋️", "训练", "xùn liàn", "Train",
     "我们 要 训练 一 个 AI。", "We're training an AI.",
     "训练 中 / 运动 训练 / 教 AI", DAY),
]
for em, cn, py, en, ex_cn, ex_en, hint, cl in recognize_words:
    s = vocab_recognize(prs, cl, em, cn, py, en, ex_cn, ex_en, hint)
    n += 1; pn(s, n)


# ============================================================
# 14-15 · 我 会 写 (writing practice)
# ============================================================
s = vocab_write(prs, ORANGE, "分类", "Classify",
                [("分", "fēn", "4 笔", "上 「八」 下 「刀」 — 切 开 分 开"),
                 ("类", "lèi", "9 笔", "上 「米」 下 「大」")])
n += 1; pn(s, n)

s = vocab_write(prs, DAY, "训练", "Train",
                [("训", "xùn", "5 笔", "「讠」 + 「川」 — 用 话 教"),
                 ("练", "liàn", "8 笔", "「纟」 + 「东」 — 反复 练 习")])
n += 1; pn(s, n)


# ============================================================
# 16 · SESSION 3 DIVIDER
# ============================================================
s = div(prs, "Session 3", "🎨 下午 3:00–4:30  ·  训练 真 AI + 我 是 AI 训练师!",
        DAY, "🤖"); n += 1; pn(s, n)


# ============================================================
# TM 5-STEP INSTRUCTIONS + THEMES  (Session 3 opener — hands-on prep)
# Note: Cat-differentiation example was MOVED to Session 1 (Rounds 1 & 2).
# Session 3 now starts directly with the hands-on instructions for student groups.
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "📝 操作 步骤 · How to Use Teachable Machine", DAY)

# LEFT: 5 steps
panel(s, 0.40, 0.95, 4.55, 3.95, DAY, fill=WHITE, lw=3)
panel_head(s, 0.40, 0.95, 4.55, DAY, "5 步 操作  5 Steps", sz=14)
steps = [
    ("1️⃣", "选择 分类 主题", "Pick a theme"),
    ("2️⃣", "收集 图片", "Collect images"),
    ("3️⃣", "点击 训练", "Click Train"),
    ("4️⃣", "测试 AI", "Test the AI"),
    ("5️⃣", "看 AI 会 不 会 错!", "See if AI gets it wrong!"),
]
for i, (num, cn, en) in enumerate(steps):
    y = 1.60 + i * 0.62
    tb(s, 0.55, y, 0.55, 0.50, num, sz=20, b=True, c=DAY, a=PP_ALIGN.LEFT)
    tb(s, 1.15, y+0.02, 3.70, 0.32, cn, sz=13, b=True, c=DARK)
    tb(s, 1.15, y+0.33, 3.70, 0.25, en, sz=9, c=GRAY)

# RIGHT: 5 themes
panel(s, 5.05, 0.95, 4.55, 3.95, ORANGE, fill=WHITE, lw=3)
panel_head(s, 5.05, 0.95, 4.55, ORANGE, "🎨 主题 选择  Themes", sz=14)
themes = [
    ("🐾", "动 物", "Animals"),
    ("🍎", "水 果", "Fruits"),
    ("😊", "表 情", "Emotions"),
    ("✋", "手 势", "Hand gestures"),
    ("🌈", "颜色 物品", "Colored objects"),
]
for i, (em, cn, en) in enumerate(themes):
    y = 1.60 + i * 0.62
    tb(s, 5.20, y, 0.55, 0.50, em, sz=22, a=PP_ALIGN.LEFT)
    tb(s, 5.85, y+0.02, 3.55, 0.32, cn, sz=14, b=True, c=ORANGE)
    tb(s, 5.85, y+0.33, 3.55, 0.25, en, sz=9, c=GRAY)

url = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(5.00), Inches(9.20), Inches(0.42))
url.fill.solid(); url.fill.fore_color.rgb = INK; url.line.fill.background()
tb(s, 0.55, 5.05, 9.0, 0.32, "💻 teachablemachine.withgoogle.com  ·  免费 + 不 用 注册!",
   sz=12, b=True, c=STAR, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "操作 建议:\n• 全 班 一起 看 投影 演示 一 次\n• 然后 小 组 用 iPad 自己 玩\n• K-2: 选 简单 的 (动物 / 水果)\n• 3-5: 可以 尝试 手势 / 表情\n• 重要: 让 学生 自己 拍 照, 不 要 老师 代 劳")


# ============================================================
# 19 · BAD DATA EXAMPLE (错 数据 = 错 结果)
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "😅 错 数据 = 错 结果! · Bad Data = Bad Results", PINK)

# LEFT: training panel (all white shapes labeled "rabbit")
panel(s, 0.40, 0.95, 4.30, 3.80, ORANGE, fill=WHITE, lw=3)
panel_head(s, 0.40, 0.95, 4.30, ORANGE, "📚 训练: 全 部 是 白色!", sz=13)

train = [
    ("白 兔子", "→ 「兔子」 ✅"),
    ("白 狗 狗", "→ 「兔子」 ⁉️"),
    ("白 猫 咪", "→ 「兔子」 ⁉️"),
]
for i, (label, pred) in enumerate(train):
    y = 1.65 + i * 1.00
    oval = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.60), Inches(y), Inches(0.85), Inches(0.85))
    oval.fill.solid(); oval.fill.fore_color.rgb = WHITE
    oval.line.color.rgb = DARK; oval.line.width = Pt(2)
    tb(s, 0.60, y+0.22, 0.85, 0.45, "白", sz=22, b=True, c=DARK, a=PP_ALIGN.CENTER)
    tb(s, 1.55, y+0.10, 1.20, 0.35, label, sz=13, b=True, c=DARK)
    arrow(s, 2.78, y+0.32, w=0.28, h=0.25, color=ORANGE)
    tb(s, 3.10, y+0.20, 1.60, 0.50, pred, sz=11, b=True, c=ORANGE, a=PP_ALIGN.LEFT)

arrow(s, 4.80, 2.60, w=0.45, h=0.45, color=PINK)

# RIGHT: test panel
panel(s, 5.30, 0.95, 4.30, 3.80, PINK, fill=WHITE, lw=3)
panel_head(s, 5.30, 0.95, 4.30, PINK, "🧪 测试: 给 AI 看 新 图", sz=13)
big = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(6.45), Inches(1.65), Inches(2.00), Inches(1.55))
big.fill.solid(); big.fill.fore_color.rgb = WHITE
big.line.color.rgb = DARK; big.line.width = Pt(3)
tb(s, 6.45, 2.05, 2.00, 0.60, "白", sz=36, b=True, c=DARK, a=PP_ALIGN.CENTER)
tb(s, 5.30, 3.30, 4.30, 0.32, "白色 北极熊  Polar Bear",
   sz=14, b=True, c=DARK, a=PP_ALIGN.CENTER)
ai = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.55), Inches(3.78), Inches(3.80), Inches(0.90))
ai.fill.solid(); ai.fill.fore_color.rgb = PINK; ai.line.fill.background()
tb(s, 5.65, 3.85, 3.60, 0.28, "🤖 AI 说:",
   sz=11, b=True, c=STAR, a=PP_ALIGN.LEFT)
tb(s, 5.65, 4.15, 3.60, 0.50, "「兔子!」",
   sz=26, b=True, c=WHITE, a=PP_ALIGN.CENTER)

tk = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(4.85), Inches(9.20), Inches(0.55))
tk.fill.solid(); tk.fill.fore_color.rgb = INK; tk.line.fill.background()
tb(s, 0.55, 4.92, 9.0, 0.40, "💡 错 数据 = 错 结果!  AI 只 学 到 「白色 = 兔子」!",
   sz=13, b=True, c=STAR, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "搞笑 + 重要!\n• 让 学生 一起 喊 「兔子!」 — 加强 记忆\n• 训练 数据 全部 是 白色 → AI 学 到 「白色 = 兔子」\n• 给 它 新 的 白色 东西 (北极熊) → 它 还是 说 「兔子」\n• 真实 例子: 如果 训练 数据 不 多 样, AI 就 不 公平\n• 这 是 AI 公平性 (data bias) 的 入门 话题")


# ============================================================
# 20 · THE MORAL — AI 像 你 in school
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🏫 AI 学 习 像 你 在 学校! · AI Learns Like You at School", DAY)

# Big banner
bn = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.95), Inches(9.2), Inches(0.75))
bn.fill.solid(); bn.fill.fore_color.rgb = WARM
bn.line.color.rgb = DAY; bn.line.width = Pt(2.5)
tb(s, 0.55, 1.02, 9.0, 0.35, "好 好 学 = 好 结果! 你 这 样, AI 也 一 样!",
   sz=15, b=True, c=DAY, a=PP_ALIGN.CENTER)
tb(s, 0.55, 1.38, 9.0, 0.28, "Study well = good results. Same for you AND for AI.",
   sz=10, c=GRAY, a=PP_ALIGN.CENTER)

# 2 columns: Student | AI
cols = [
    ("🎒", "你 (学 生) · You", DAY, [
        ("✅", "认真 听 课 + 做 作业", "Good: pay attention + do homework"),
        ("🏆", "→ 考 试 答 对!", "→ Get answers right!"),
        ("❌", "走 神 + 不 复习", "Bad: zone out + don't review"),
        ("😅", "→ 考 试 答 错!", "→ Get answers wrong!"),
    ]),
    ("🤖", "AI · AI", PINK, [
        ("✅", "好 数据 + 多 例子", "Good: clean data + lots of examples"),
        ("🏆", "→ AI 答 对!", "→ AI predicts correctly!"),
        ("❌", "坏 数据 + 错 标签", "Bad: wrong data + bad labels"),
        ("😅", "→ AI 答 错!", "→ AI gets it wrong!"),
    ]),
]
col_w = 4.40; gap = 0.40
total = 2*col_w + gap; start = (10 - total)/2
for i, (em, title, cl, bullets) in enumerate(cols):
    x = start + i*(col_w + gap)
    panel(s, x, 1.85, col_w, 2.95, cl, fill=WHITE, lw=3)
    head = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.85), Inches(col_w), Inches(0.50))
    head.fill.solid(); head.fill.fore_color.rgb = cl; head.line.fill.background()
    tb(s, x+0.10, 1.92, col_w-0.20, 0.40, f"{em}  {title}",
       sz=14, b=True, c=WHITE, a=PP_ALIGN.CENTER)
    for j, (e, cn, en) in enumerate(bullets):
        y = 2.45 + j*0.58
        tb(s, x+0.18, y, 0.40, 0.45, e, sz=18, a=PP_ALIGN.LEFT)
        tb(s, x+0.65, y+0.04, col_w-0.80, 0.28, cn, sz=12, b=True, c=DARK)
        tb(s, x+0.65, y+0.32, col_w-0.80, 0.24, en, sz=8, c=GRAY)

# Bottom moral
mb = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(4.95), Inches(9.20), Inches(0.45))
mb.fill.solid(); mb.fill.fore_color.rgb = DAY
mb.line.color.rgb = STAR; mb.line.width = Pt(2)
tb(s, 0.55, 5.00, 9.0, 0.35, "💡 给 自己 好 的 学 习, 给 AI 好 的 数据 — 我们 都 是 学 习 者!",
   sz=12, b=True, c=STAR, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "核心 教学 时刻!\n• 让 学生 自己 说: AI 学 习 像 我 在 学校 — 怎么 像?\n• 引导 出: 都 需要 「好 例子」 + 「多 练 习」\n• 数据 公平 = 数据 多 样 (不 只 是 一 种)\n• 这 是 道德 + STEM 结 合 — 让 学生 思考: 怎么 当 「好 AI 训练师」?\n• 联 想: 我 们 平 时 给 自己 看 什么? 玩 游戏? 还是 学 知识?")


# ============================================================
# 21 · PROJECT INTRO (hype slide)
# ============================================================
s = ns(prs); bg(s, INK, prs)
for x, y in [(0.4, 0.5), (9.1, 0.6), (0.5, 4.8), (9.0, 4.7), (1.0, 0.4), (8.5, 4.9)]:
    d = s.shapes.add_shape(MSO_SHAPE.STAR_5_POINT, Inches(x), Inches(y), Inches(0.40), Inches(0.40))
    d.fill.solid(); d.fill.fore_color.rgb = STAR; d.line.fill.background()

tb(s, 0.3, 0.55, 9.4, 0.40, "🏆 Project Time!",
   sz=18, b=True, c=STAR, a=PP_ALIGN.CENTER)

title_box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.15), Inches(8.8), Inches(2.40))
title_box.fill.solid(); title_box.fill.fore_color.rgb = DAY
title_box.line.color.rgb = STAR; title_box.line.width = Pt(4)
tb(s, 0.8, 1.35, 8.4, 0.85, "我 是 AI 训练师!",
   sz=44, b=True, c=WHITE, a=PP_ALIGN.CENTER)
tb(s, 0.8, 2.30, 8.4, 0.50, "I'm an AI Trainer!",
   sz=22, b=True, c=STAR, a=PP_ALIGN.CENTER)
tb(s, 0.8, 2.90, 8.4, 0.40, "小组 合作 — 训练 你 自己 的 AI 模型!",
   sz=14, c=WARM, a=PP_ALIGN.CENTER)

tb(s, 0.3, 3.95, 9.4, 1.00, "🤖   📷   🧠   ✨   🏆",
   sz=52, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "30 秒 hype slide.\n• 让 学生 大 声 喊 「我 是 AI 训练师!」\n• 制造 仪式 感")


# ============================================================
# 22 · PROJECT WORKFLOW
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🛠️ 项目 流程 · Project Workflow", DAY)

# LEFT: theme choices
panel(s, 0.40, 0.95, 3.85, 4.30, PURPLE, fill=WHITE, lw=3)
panel_head(s, 0.40, 0.95, 3.85, PURPLE, "🎨 选 一 个 主题  Pick a Theme", sz=13)
themes = [
    ("🐾", "动 物"),
    ("🍎", "水 果"),
    ("😊", "表 情"),
    ("✋", "手 势"),
    ("🌈", "颜色 物品"),
]
for i, (em, cn) in enumerate(themes):
    y = 1.65 + i * 0.65
    tb(s, 0.60, y, 0.65, 0.55, em, sz=26, a=PP_ALIGN.LEFT)
    tb(s, 1.35, y+0.10, 2.80, 0.40, cn, sz=18, b=True, c=PURPLE)

# RIGHT: 4-step workflow
panel(s, 4.40, 0.95, 5.20, 4.30, DAY, fill=WHITE, lw=3)
panel_head(s, 4.40, 0.95, 5.20, DAY, "📋 4 个 步骤  4 Steps", sz=13)
flow = [
    ("📷", "收集 数据", "Collect data", "多 拍 一 些 — 数据 越 多 越 好!"),
    ("🏋️", "训练 AI", "Train AI", "点 Train 按钮"),
    ("✨", "测试 AI", "Test AI", "给 它 看 新 图 — 它 答 对 吗?"),
    ("🤔", "再 想 一 想", "Reflect", "数据 够 吗? 数据 公平 吗?"),
]
for i, (em, cn, en, hint) in enumerate(flow):
    y = 1.60 + i * 0.85
    tb(s, 4.60, y, 0.65, 0.65, em, sz=28, a=PP_ALIGN.LEFT)
    tb(s, 5.40, y+0.02, 1.85, 0.35, cn, sz=14, b=True, c=DAY)
    tb(s, 5.40, y+0.36, 1.85, 0.28, en, sz=9, c=GRAY)
    tb(s, 7.35, y+0.10, 2.15, 0.55, hint, sz=10, b=True, c=DARK)

tip = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(5.00), Inches(9.20), Inches(0.42))
tip.fill.solid(); tip.fill.fore_color.rgb = STAR; tip.line.fill.background()
tb(s, 0.55, 5.05, 9.0, 0.32, "👥 4 人 一 组 — 多 收 集 数据, 做 一 个 好 AI!",
   sz=12, b=True, c=INK, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "30 分钟 项目:\n• 5 min: 组队 + 选 主题\n• 15 min: 收集 数据 (强调: 越 多 越 好, 多 角 度!)\n• 5 min: 训练 + 测试\n• 5 min: 反思 — 我们 的 数据 公平 吗?\n\n关键: 让 学生 实际 体会 「数据 决定 AI 表现」 — 联 系 Slide 17 + 19 + 20")


# ============================================================
# 23 · REFLECTION
# ============================================================
s = ns(prs); bg(s, CREAM, prs); hb(s, "🤔 想 一 想 · Let's Reflect", DAY)

tb(s, 0.4, 0.85, 9.2, 0.32, "做 完 项目 后, 我们 一起 想 一 想:",
   sz=13, b=True, c=DARK, a=PP_ALIGN.CENTER)

qs = [
    ("📊", "我们 的 数据 够 多 吗?", "Did we have enough data?", CYBER),
    ("⚖️", "我们 的 数据 公平 吗?", "Was our data fair?", PURPLE),
    ("❓", "为什么 AI 猜 错 了?", "Why did AI guess wrong?", PINK),
    ("💡", "怎样 让 AI 更 聪明?", "How can we make AI smarter?", DAY),
]
card_w = 4.40; card_h = 1.75; gap_x = 0.20; gap_y = 0.20
start_x = (10 - 2*card_w - gap_x)/2
for i, (em, cn, en, cl) in enumerate(qs):
    row = i // 2; col = i % 2
    x = start_x + col*(card_w + gap_x)
    y = 1.30 + row*(card_h + gap_y)
    panel(s, x, y, card_w, card_h, cl, fill=WHITE, lw=3)
    tb(s, x+0.20, y+0.30, 0.85, 0.85, em, sz=40, a=PP_ALIGN.LEFT)
    tb(s, x+1.20, y+0.28, card_w-1.35, 0.55, cn, sz=16, b=True, c=cl)
    tb(s, x+1.20, y+0.86, card_w-1.35, 0.40, en, sz=10, c=GRAY)

tb(s, 0.4, 5.05, 9.2, 0.32, "💬 和 小 组 讨论 — 然后 准备 分享!",
   sz=12, b=True, c=DAY, a=PP_ALIGN.CENTER)
n += 1; pn(s, n)
notes(s, "5-10 分钟 讨论:\n• 让 每 组 讨论 4 个 问题\n• 老师 走动 倾听\n• 关键: 联 系 Slide 20 的 「好 学 习 = 好 结果」 — 我们 当 「好 AI 训练师」!\n• 这 是 AI ethics 的 入门")


# ============================================================
# 24 · SHARE + CLOSE
# ============================================================
s = share_close(prs, DAY,
    frames_cn=["「我们 训练 的 AI 会 认 ______」",
               "「数据 越 多, AI 就 ______」"],
    frames_en="Our AI can recognize ___  ·  More data makes AI ___",
    next_day_cn="Day 4 · 科技 改变 生活 — 以前 vs 现在!",
    next_day_en="Day 4 · How tech changes life — then vs now!",
    next_emoji="📱")
n += 1; pn(s, n)
notes(s, "10 分钟:\n• 每 组 1-2 分钟 分享\n• 句型 帮 学生 说 出 「数据 多 少」 + 「好 数据 vs 坏 数据」\n• 收 拾 教 具 准备 Day 4")


out = os.path.join(os.path.dirname(__file__), "day3_ml.pptx")
prs.save(out)
print(f"Saved {out}  ({len(prs.slides)} slides)")
