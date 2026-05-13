#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
仰望星空 · Day 2: 星空、星座与神话传说  Constellations & Myths
Picture book anchor: 牛郎织女  Cowherd and Weaver Girl
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import random, os

prs = Presentation()
prs.slide_width = Inches(10); prs.slide_height = Inches(5.625)
W, H = prs.slide_width, prs.slide_height

# Palette — Mythic Night
NIGHT  = RGBColor(0x0D,0x1B,0x3E)
COSMIC = RGBColor(0x6A,0x1B,0x9A)
STAR   = RGBColor(0xF5,0xC2,0x42)
GOLD   = RGBColor(0xFF,0xB7,0x00)
EARTH  = RGBColor(0x1E,0x88,0xE5)
RED    = RGBColor(0xC8,0x25,0x3E)
PINK   = RGBColor(0xEC,0x40,0x7A)
NEBULA = RGBColor(0x7B,0x1F,0xA2)
SKY    = RGBColor(0x42,0xA5,0xF5)
SILVER = RGBColor(0xB0,0xBE,0xC5)
GREEN  = RGBColor(0x2E,0x7D,0x32)
CREAM  = RGBColor(0xFF,0xF8,0xE7)
WARM   = RGBColor(0xFF,0xF3,0xE0)
WHITE  = RGBColor(0xFF,0xFF,0xFF)
DARK   = RGBColor(0x2C,0x2C,0x2C)
GRAY   = RGBColor(0x88,0x88,0x88)
LGRAY  = RGBColor(0xBB,0xBB,0xBB)
IMGBG  = RGBColor(0xE8,0xE8,0xF0)

def ns(): return prs.slides.add_slide(prs.slide_layouts[6])
def tb(s,l,t,w,h,txt,sz=18,b=False,c=DARK,a=None):
    bx=s.shapes.add_textbox(Inches(l),Inches(t),Inches(w),Inches(h)); tf=bx.text_frame; tf.word_wrap=True
    p=tf.paragraphs[0]
    if a: p.alignment=a
    r=p.add_run(); r.text=txt; r.font.size=Pt(sz); r.font.bold=b; r.font.color.rgb=c; r.font.name='KaiTi'
    return tf
def bg(s,c):
    sh=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,W,H); sh.fill.solid(); sh.fill.fore_color.rgb=c; sh.line.fill.background()
    sp=sh._element; sp.getparent().remove(sp); s.shapes._spTree.insert(2,sp)
def hb(s,txt,c=NIGHT,t=0.15):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(t),Inches(9.4),Inches(0.55))
    sh.fill.solid(); sh.fill.fore_color.rgb=c; sh.line.fill.background()
    tb(s,0.4,t+0.03,9.2,0.5,txt,sz=20,b=True,c=WHITE)
def ib(s,l,t,w,h,lb="📷"):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(l),Inches(t),Inches(w),Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb=IMGBG; sh.line.fill.background()
    tb(s,l+0.1,t+h/2-0.2,w-0.2,0.4,lb,sz=14,c=LGRAY,a=PP_ALIGN.CENTER)
def pn(s,n): tb(s,9.0,5.25,0.8,0.3,str(n),sz=10,c=GRAY,a=PP_ALIGN.RIGHT)
def notes(s,text):
    nf=s.notes_slide.notes_text_frame
    lines=text.split("\n"); nf.text=lines[0]
    for line in lines[1:]:
        p=nf.add_paragraph(); p.text=line
def div(title,sub,color,emoji=""):
    s=ns(); bg(s,color)
    tb(s,1,1.5,8,1.2,f"{emoji} {title}",sz=42,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,1,2.8,8,0.8,sub,sz=22,c=STAR,a=PP_ALIGN.CENTER)
    for x,y in [(0.8,4.7),(1.8,4.5),(7.8,4.5),(8.6,4.7),(2.0,1.0),(8.0,1.0)]:
        d=s.shapes.add_shape(MSO_SHAPE.STAR_5_POINT,Inches(x),Inches(y),Inches(0.35),Inches(0.35))
        d.fill.solid(); d.fill.fore_color.rgb=STAR; d.line.fill.background()
    return s
def tianzi(s,x,y,size,char,color,pinyin=None,char_sz=130):
    box=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(x),Inches(y),Inches(size),Inches(size))
    box.fill.solid(); box.fill.fore_color.rgb=WHITE
    box.line.color.rgb=color; box.line.width=Pt(3)
    mx=x+size/2; my=y+size/2; lw=0.015
    v=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(mx-lw/2),Inches(y),Inches(lw),Inches(size))
    v.fill.solid(); v.fill.fore_color.rgb=LGRAY; v.line.fill.background()
    h=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(x),Inches(my-lw/2),Inches(size),Inches(lw))
    h.fill.solid(); h.fill.fore_color.rgb=LGRAY; h.line.fill.background()
    tb(s,x,y,size,size,char,sz=char_sz,b=True,c=color,a=PP_ALIGN.CENTER)
    if pinyin:
        tb(s,x,y+size+0.05,size,0.30,pinyin,sz=12,b=True,c=GRAY,a=PP_ALIGN.CENTER)

n=0

# 1. Cover
s=ns(); bg(s,NIGHT)
random.seed(11)
for _ in range(50):
    x=random.uniform(0.3,9.7); y=random.uniform(0.3,5.3); sz=random.uniform(0.06,0.16)
    d=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x),Inches(y),Inches(sz),Inches(sz))
    d.fill.solid(); d.fill.fore_color.rgb=STAR; d.line.fill.background()
tb(s,0.5,0.4,9,0.6,"🌌 仰望星空  Looking Up at the Stars",sz=28,b=True,c=STAR,a=PP_ALIGN.CENTER)
tb(s,0.5,1.0,9,0.45,"Day 2 · 星空、星座 与 神话  Constellations & Myths",sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
# Big Dipper (Big-Dipper-style 7 stars)
dipper=[(2.5,2.5),(3.2,2.3),(4.0,2.1),(4.8,2.3),(5.4,2.8),(5.7,3.4),(6.0,4.0)]
prev=None
for (x,y) in dipper:
    sh=s.shapes.add_shape(MSO_SHAPE.STAR_5_POINT,Inches(x),Inches(y),Inches(0.40),Inches(0.40))
    sh.fill.solid(); sh.fill.fore_color.rgb=STAR; sh.line.fill.background()
    if prev:
        # Use connector line
        ln=s.shapes.add_connector(2,Inches(prev[0]+0.2),Inches(prev[1]+0.2),Inches(x+0.2),Inches(y+0.2))
        ln.line.color.rgb=STAR; ln.line.width=Pt(1)
    prev=(x,y)
tb(s,3.0,4.55,4,0.4,"⭐ 北斗七星 Big Dipper",sz=14,b=True,c=STAR,a=PP_ALIGN.CENTER)
tb(s,0.5,4.95,9,0.40,"📖 绘本: 牛郎织女  Cowherd and Weaver Girl",sz=15,b=True,c=STAR,a=PP_ALIGN.CENTER)
tb(s,0.5,5.30,9,0.25,"Picture book — a Chinese love legend in the stars",sz=9,c=LGRAY,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"Day 2 开场:\n• 「昨天 我们 学了 太阳系 — 今天 我们 看 星星 和 星座!」\n• 古人 用 故事 解释 星星 — 我们 来 听 一个 中国 故事")

# 2. Session 1 Divider
s=div("Session 1  上午 11:00–11:45","⭐ 故事课 · 牛郎织女 和 星座  ·  45 min",COSMIC,"💫"); n+=1; pn(s,n)

# 3. Learning Goals
s=ns(); bg(s,CREAM); hb(s,"🎯 今天的学习目标  Today's Learning Goals",NIGHT)
tb(s,0.4,0.85,9.2,0.30,"上完这节课, 你会……  By the end, you'll be able to…",sz=12,b=True,c=DARK,a=PP_ALIGN.CENTER)
goals=[
    ("1","📖","通过 牛郎织女 — 了解 古人 用 故事 解释 星空",PINK),
    ("2","⭐","认识 星座 和 银河 是 什么",STAR),
    ("3","💡","知道 星空 = 科学 + 故事 + 想象",COSMIC),
    ("4","🌙","观察 夜空 — 发挥 想象力!",NEBULA),
]
for i,(num,em,text,cl) in enumerate(goals):
    y=1.30+i*0.95
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(y),Inches(9.2),Inches(0.85))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(2.5)
    nb=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(0.55),Inches(y+0.18),Inches(0.50),Inches(0.50))
    nb.fill.solid(); nb.fill.fore_color.rgb=cl; nb.line.fill.background()
    tb(s,0.55,y+0.22,0.50,0.40,num,sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,1.20,y+0.18,0.60,0.50,em,sz=28,a=PP_ALIGN.CENTER)
    tb(s,1.90,y+0.20,7.6,0.55,text,sz=14,b=True,c=DARK)
n+=1; pn(s,n)
notes(s,"1-2 分钟 — 学习目标")

# 4. Hook — connect dots / What pattern do you see?
s=ns(); bg(s,NIGHT); hb(s,"🔍 连 一 连 — 你 看到 什么?  Connect the Dots!",STAR)
tb(s,0.4,0.85,9.2,0.30,"古人 抬头 — 把 星星 连 起来, 看到 了 什么?",sz=13,b=True,c=STAR,a=PP_ALIGN.CENTER)
tb(s,0.4,1.15,9.2,0.22,"Ancient people connected the stars — what did they see?",sz=9,c=LGRAY,a=PP_ALIGN.CENTER)
# 3 sample constellation cards
cons=[
    ("🐻","大熊座","Big Bear","北斗 是 大熊 的 尾巴!",STAR),
    ("🦁","狮子座","Leo","狮子 王 — 春天 出现",GOLD),
    ("🦂","天蝎座","Scorpio","红色 心 = 心宿 二",RED),
]
for i,(em,cn,en,d,cl) in enumerate(cons):
    x=0.4+i*3.10
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.55),Inches(2.95),Inches(3.10))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    tb(s,x+0.05,1.70,2.85,1.20,em,sz=78,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.95,2.85,0.40,cn,sz=18,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.40,2.85,0.28,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.10,3.78,2.75,0.85,d,sz=11,b=True,c=DARK,a=PP_ALIGN.CENTER)
prompt=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(4.80),Inches(9.2),Inches(0.65))
prompt.fill.solid(); prompt.fill.fore_color.rgb=COSMIC; prompt.line.color.rgb=STAR; prompt.line.width=Pt(2)
tb(s,0.55,4.85,9.0,0.30,"💬 你 觉得 古人 为 什么 这样 看 星星?",sz=12,b=True,c=STAR,a=PP_ALIGN.CENTER)
tb(s,0.55,5.18,9.0,0.20,"Why do you think ancient people saw shapes in the stars?",sz=8,c=WARM,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"4 分钟 — Hook:\n• 「夜空 有 好多 星星 — 你 觉得 它们 像 什么?」\n• 展示 3 个 星座 例子 (大熊 / 狮子 / 天蝎)\n• 引导: 「古人 没 电视, 没 手机 — 他们 看 星星 就 是 在 看 故事!」")

# 5. Picture Book — 牛郎织女
s=ns(); bg(s,CREAM); hb(s,"📖 绘本: 牛郎织女  Cowherd & Weaver",PINK)
tb(s,0.4,0.85,9.2,0.40,"一个 关于 银河 + 七夕 的 中国 古老 故事",sz=15,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.25,9.2,0.26,"An old Chinese story about the Milky Way and Qixi Festival",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
ib(s,0.5,1.65,4.5,3.30,"📚 绘本 / Book illustration")
panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.25),Inches(1.65),Inches(4.35),Inches(3.30))
panel.fill.solid(); panel.fill.fore_color.rgb=WHITE; panel.line.color.rgb=PINK; panel.line.width=Pt(2.5)
head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.25),Inches(1.65),Inches(4.35),Inches(0.50))
head.fill.solid(); head.fill.fore_color.rgb=PINK; head.line.fill.background()
tb(s,5.40,1.72,4.10,0.40,"📚 故事 简介  About the Story",sz=13,b=True,c=WHITE)
parts=[
    "👦 牛郎 — 一个 放牛 的 小伙子",
    "👧 织女 — 天上 王母 的 孙女",
    "💕 他们 相爱, 结婚, 生了 孩子",
    "😢 王母 用 银河 把 他们 分开!",
    "🐦 每年 七月初七 — 喜鹊 搭 桥",
    "✨ 让 他们 一年 见 一次!",
]
for i,line in enumerate(parts):
    tb(s,5.40,2.25+i*0.42,4.10,0.40,line,sz=11,b=True,c=DARK)
n+=1; pn(s,n)
notes(s,"5-6 分钟 — 讲故事:\n• 老师 用 戏剧 语气 讲 牛郎织女\n• 强调: 「夏天 晚上 抬头 — 你 真的 能 看到 牛郎星 和 织女星!」\n• 中间 隔着 银河 (Milky Way)\n• 文化 链接: 七夕 = 中国 情人节")

# ============================================================
# 5b. 绘本 前 · Pre-reading predictions
# ============================================================
s=ns(); bg(s,CREAM); hb(s,"🔮 翻 开 之前 — 想 一 想!  Before We Read",PINK)
tb(s,0.4,0.85,9.2,0.34,"光 看 题目 「牛郎 织女」 — 你 能 猜 到 什么?",sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.18,9.2,0.22,"From the title 'Cowherd & Weaver' — what can you guess?",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
preds=[
    ("👦","「牛郎」 是 做 什么 的?","What does 'Cowherd' do?",PINK),
    ("👧","「织女」 是 做 什么 的?","What does 'Weaver Girl' do?",NEBULA),
    ("🌌","为什么 是 一个 星空 故事?","Why is this a sky story?",STAR),
]
for i,(em,q,en,cl) in enumerate(preds):
    x=0.4+i*3.10
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.55),Inches(2.95),Inches(2.85))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    tb(s,x+0.05,1.70,2.85,1.0,em,sz=66,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.75,2.85,0.50,q,sz=14,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.30,2.85,0.40,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.10,3.80,2.75,0.55,"💡 大胆 猜! 没 标准 答案",sz=10,b=True,c=DARK,a=PP_ALIGN.CENTER)
tps=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(4.55),Inches(9.2),Inches(0.95))
tps.fill.solid(); tps.fill.fore_color.rgb=NIGHT; tps.line.color.rgb=STAR; tps.line.width=Pt(2.5)
tb(s,0.55,4.62,9.0,0.30,"👥 Think-Pair-Share: 跟 同桌 说 你 的 猜 想 (1 分钟)",sz=12,b=True,c=STAR,a=PP_ALIGN.CENTER)
tb(s,0.55,4.95,9.0,0.26,"Turn to a partner — share your guess!",sz=9,c=LGRAY,a=PP_ALIGN.CENTER)
tb(s,0.55,5.22,9.0,0.24,"💬 「我 猜 ___ 因为 ___」",sz=12,b=True,c=STAR,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"🔮 PRE-READING · 3 分钟:\n• 让 学生 先 猜\n• 收集 想法 — 写 1-2 个 在 黑板\n• 引出 故事 — 「我们 来 看 真正 的 故事!」")

# ============================================================
# 5c. 听 故事 时 · During-reading
# ============================================================
s=ns(); bg(s,CREAM); hb(s,"👀 听 故事 时 — 你 的 任务!  While You Listen",STAR)
tb(s,0.4,0.85,9.2,0.34,"3 个 任务 — 一边 听 一边 留心!",sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.18,9.2,0.22,"3 missions — listen + observe!",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
obs=[
    ("1","💕","牛郎 和 织女 怎么 认识 的?","How did they meet?",PINK),
    ("2","😢","为什么 王母 把 他们 分开?","Why did Queen Mother separate them?",NEBULA),
    ("3","🐦","喜鹊 怎么 帮助 他们?","How did the magpies help?",STAR),
]
for i,(num,em,cn,en,cl) in enumerate(obs):
    y=1.55+i*1.05
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(y),Inches(9.2),Inches(0.95))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    nb=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(0.55),Inches(y+0.20),Inches(0.55),Inches(0.55))
    nb.fill.solid(); nb.fill.fore_color.rgb=cl; nb.line.fill.background()
    tb(s,0.55,y+0.27,0.55,0.40,num,sz=20,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,1.30,y+0.18,0.80,0.55,em,sz=32,a=PP_ALIGN.CENTER)
    tb(s,2.30,y+0.14,7.0,0.40,cn,sz=15,b=True,c=cl)
    tb(s,2.30,y+0.54,7.0,0.30,en,sz=10,c=GRAY)
tip=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(4.85),Inches(9.2),Inches(0.60))
tip.fill.solid(); tip.fill.fore_color.rgb=STAR; tip.line.fill.background()
tb(s,0.55,4.93,9.0,0.30,"👂 用 心 听! 念 完 — 我们 一起 讨论",sz=12,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.55,5.23,9.0,0.22,"Listen well — we'll discuss after.",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"👀 DURING-READING · setup 1 分钟 + 念 8 分钟:\n• 念 3 个 任务\n• 老师 讲 故事 (戏剧 化, 加 手势)\n• 念 完后 → 下一页 讨论")

# ============================================================
# 5d. 故事 后 · Post-reading discussion (6 Qs)
# ============================================================
s=ns(); bg(s,CREAM); hb(s,"💭 听 完 故事 — 一起 讨论!  After We Read",COSMIC)
tb(s,0.4,0.85,9.2,0.30,"选 1-2 个 问题 — 全班 / Think-Pair-Share",sz=12,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.15,9.2,0.22,"Pick 1-2 questions — class / Think-Pair-Share",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
qs=[
    ("💕","牛郎 织女 相 爱 — 但 不能 见面, 你 觉得?","Love but apart — how do YOU feel?"),
    ("👑","王母 太 严厉 了 吗?","Was Queen Mother too strict?"),
    ("🐦","喜鹊 真好! 它们 为什么 帮 忙?","Why did magpies help?"),
    ("📅","一年 只 见 一次 — 够 吗?","One meeting a year — enough?"),
    ("🌌","西方 看 同一片 星空 是 大熊座 — 神奇 吗?","Same sky, different story — wow?"),
    ("✨","你 自己 想 给 这 几颗 星 起 什么 名字?","What names would YOU give those stars?"),
]
for i,(em,cn,en) in enumerate(qs):
    col=i%3; row=i//3
    x=0.4+col*3.10; y=1.45+row*1.80
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(2.95),Inches(1.65))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=COSMIC; sh.line.width=Pt(2.5)
    nb=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x+0.10),Inches(y+0.10),Inches(0.42),Inches(0.42))
    nb.fill.solid(); nb.fill.fore_color.rgb=COSMIC; nb.line.fill.background()
    tb(s,x+0.10,y+0.14,0.42,0.34,str(i+1),sz=14,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,x+0.60,y+0.05,0.55,0.50,em,sz=24,a=PP_ALIGN.CENTER)
    tb(s,x+0.15,y+0.60,2.70,0.65,cn,sz=11,b=True,c=DARK)
    tb(s,x+0.15,y+1.25,2.70,0.32,en,sz=8,c=GRAY)
n+=1; pn(s,n)
notes(s,"💭 POST-READING · 5-7 分钟:\n• 选 2-3 题 — 全班 / Think-Pair-Share\n• Q5 + Q6 启发 文化 + 想象\n• 鼓励 文化 比较 — 同一片 天 多个 故事")

# 6. 银河 + 牛郎星 + 织女星 visualization
s=ns(); bg(s,NIGHT); hb(s,"💫 银河 · 牛郎星 · 织女星  Milky Way",COSMIC)
tb(s,0.4,0.85,9.2,0.30,"夏天 晚上, 抬头 — 你 真的 能 看到 这 3 颗星!",sz=13,b=True,c=STAR,a=PP_ALIGN.CENTER)
tb(s,0.4,1.18,9.2,0.22,"On summer nights, you can really see these 3 stars!",sz=9,c=LGRAY,a=PP_ALIGN.CENTER)
# Milky way band (translucent)
band=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(0.5),Inches(2.80),Inches(9.0),Inches(0.40))
band.fill.solid(); band.fill.fore_color.rgb=SILVER; band.line.fill.background()
tb(s,0.5,2.85,9.0,0.30,"~ 银河 Milky Way ~",sz=12,b=True,c=NIGHT,a=PP_ALIGN.CENTER)
# Cowherd star (left of band)
sh=s.shapes.add_shape(MSO_SHAPE.STAR_5_POINT,Inches(2.0),Inches(3.60),Inches(0.65),Inches(0.65))
sh.fill.solid(); sh.fill.fore_color.rgb=STAR; sh.line.fill.background()
tb(s,1.0,4.30,2.5,0.35,"👦 牛郎星",sz=13,b=True,c=STAR,a=PP_ALIGN.CENTER)
tb(s,1.0,4.65,2.5,0.25,"Altair",sz=9,c=LGRAY,a=PP_ALIGN.CENTER)
# Weaver star (right of band)
sh=s.shapes.add_shape(MSO_SHAPE.STAR_5_POINT,Inches(7.3),Inches(1.65),Inches(0.65),Inches(0.65))
sh.fill.solid(); sh.fill.fore_color.rgb=STAR; sh.line.fill.background()
tb(s,6.5,2.30,2.5,0.35,"👧 织女星",sz=13,b=True,c=STAR,a=PP_ALIGN.CENTER)
tb(s,6.5,2.65,2.5,0.25,"Vega",sz=9,c=LGRAY,a=PP_ALIGN.CENTER)
# Bottom takeaway
bottom=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(4.95),Inches(9.2),Inches(0.55))
bottom.fill.solid(); bottom.fill.fore_color.rgb=COSMIC; bottom.line.color.rgb=STAR; bottom.line.width=Pt(2)
tb(s,0.55,5.02,9.0,0.32,"🌌 星空 = 科学 + 故事 + 想象!",sz=14,b=True,c=STAR,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"3-4 分钟 — 银河 + 两颗 星:\n• 中间 一条 白带 = 银河 (Milky Way)\n• 织女星 (Vega) 在 银河 一边\n• 牛郎星 (Altair) 在 另一边\n• 7 月 7 日 — 喜鹊 帮 他们 搭 桥 见面\n• 提示: 「今晚 抬头 — 找 银河!」")

# ============================================================
# 6b. 星座 演 一 演 · Constellation Charades (TPR)
# ============================================================
s=ns(); bg(s,NIGHT); hb(s,"🎭 星座 演 一 演!  Constellation Charades!",STAR)
tb(s,0.4,0.85,9.2,0.34,"全班 起立! 用 身体 摆 出 一个 星座!",sz=14,b=True,c=STAR,a=PP_ALIGN.CENTER)
tb(s,0.4,1.18,9.2,0.22,"Everyone stand! Use your body to make a constellation!",sz=9,c=LGRAY,a=PP_ALIGN.CENTER)
acts=[
    ("🐻","大熊座","Big Dipper","4 人 排成 一个 大 勺子",NEBULA),
    ("🦁","狮子座","Leo","站 起来 像 狮子 — 「吼!」",STAR),
    ("👦","牛郎 + 织女","Cowherd + Weaver","2 人 站 远 — 中间 隔 银河",PINK),
    ("✨","自己 编 一个!","Make your own!","队 4 人 — 创 一个 新 星座!",SKY),
]
for i,(em,cn,en,action,cl) in enumerate(acts):
    col=i%2; row=i//2
    x=0.4+col*4.65; y=1.55+row*1.65
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(4.55),Inches(1.50))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    tb(s,x+0.10,y+0.15,0.95,1.10,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+1.20,y+0.10,3.30,0.40,cn,sz=15,b=True,c=cl)
    tb(s,x+1.20,y+0.50,3.30,0.26,en,sz=9,c=GRAY)
    tb(s,x+1.20,y+0.85,3.30,0.55,action,sz=11,b=True,c=DARK)
tb(s,0.4,4.95,9.2,0.30,"💡 让 同学 猜 — 这 是 什么 星座!",sz=12,b=True,c=STAR,a=PP_ALIGN.CENTER)
tb(s,0.4,5.25,9.2,0.22,"Let others guess — what constellation is it?",sz=9,c=LGRAY,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"🎭 TPR · 6-7 分钟:\n• 4 队 — 每队 4-5 人\n• 1 分钟 — 商量 一个 星座 + 怎么 摆\n• 3 分钟 — 4 队 轮流 上台 — 摆 出 来\n• 同学 猜 (10 秒)\n• 让 K-5 都 动 — 加 想象 + 团队")

# 7. Session 1 wrap — discussion
s=ns(); bg(s,CREAM); hb(s,"🎤 一起 想 一 想  Discuss Together",COSMIC)
qs=[
    ("📖","你 喜欢 牛郎织女 的 故事 吗? 为什么?",PINK),
    ("⭐","你 看 星星 时 — 想到 过 什么 形状?",STAR),
    ("🌌","星空 — 是 科学 还是 故事? (两个 都 是!)",COSMIC),
]
tb(s,0.4,0.85,9.2,0.32,"听完 故事 — 你 觉得……",sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.18,9.2,0.22,"After the story — what do YOU think?",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
for i,(em,q,cl) in enumerate(qs):
    y=1.55+i*1.10
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(y),Inches(9.2),Inches(1.00))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(2.5)
    tb(s,0.55,y+0.22,0.80,0.60,em,sz=32,a=PP_ALIGN.CENTER)
    tb(s,1.45,y+0.30,8.0,0.50,q,sz=15,b=True,c=cl)
n+=1; pn(s,n)
notes(s,"5 分钟 — 讨论:\n• 选 1-2 个 问题 全班 讨论\n• 引导: 「同 一片 星空 — 中国 看到 牛郎织女, 西方 看到 大熊座」\n• 「都 是 想象! 都 是 美 的!」")

# 8. Session 2 Divider
s=div("Session 2  下午 2:00–2:45","📚 词汇课 · 我会认 + 我会写  ·  45 min",EARTH,"📖"); n+=1; pn(s,n)

# 9. Review
s=ns(); bg(s,CREAM); hb(s,"🔁 早上 学了 什么?  Morning Review",EARTH)
items=[
    ("👦👧","牛郎织女","Cowherd & Weaver","古老 中国 故事",PINK),
    ("⭐","星座","Constellations","把 星星 连 起来 看 形状",STAR),
    ("🌌","银河","Milky Way","一条 白色 大 河",COSMIC),
    ("💡","星空","Starry sky","= 科学 + 故事 + 想象!",NEBULA),
]
tb(s,0.4,0.85,9.2,0.32,"想 一 想 — 早上 我们 学了 什么?",sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.18,9.2,0.22,"What did we explore this morning?",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
for i,(em,cn,en,d,cl) in enumerate(items):
    col=i%2; row=i//2
    x=0.4+col*4.65; y=1.55+row*1.85
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(4.55),Inches(1.75))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    tb(s,x+0.10,y+0.20,1.00,1.20,em,sz=42,a=PP_ALIGN.CENTER)
    tb(s,x+1.20,y+0.18,3.25,0.40,cn,sz=15,b=True,c=cl)
    tb(s,x+1.20,y+0.58,3.25,0.28,en,sz=10,c=GRAY)
    tb(s,x+1.20,y+0.92,3.25,0.70,d,sz=11,b=True,c=DARK)
n+=1; pn(s,n)
notes(s,"2-3 分钟 — 复习")

# 10. 我会认 — 5 words
s=ns(); bg(s,CREAM); hb(s,"📖 我会认  I Can Read",STAR)
tb(s,0.4,0.85,9.2,0.32,"5 个 星空 词 — 我们 一起 读!",sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
words=[
    ("⭐","星星","xīng xing","stars",STAR),
    ("✨","星座","xīng zuò","constellation",GOLD),
    ("🌌","银河","yín hé","Milky Way",COSMIC),
    ("👻","神话","shén huà","myth",PINK),
    ("📖","故事","gù shi","story",NEBULA),
]
for i,(em,cn,py,en,cl) in enumerate(words):
    x=0.4+i*1.88
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.55),Inches(1.78),Inches(3.45))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    tb(s,x+0.05,1.70,1.70,0.90,em,sz=52,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.65,1.70,0.55,cn,sz=22,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.25,1.70,0.35,py,sz=12,b=True,c=DARK,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.62,1.70,0.30,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.10,4.05,1.60,0.85,"跟读\n3 遍",sz=11,b=True,c=cl,a=PP_ALIGN.CENTER)
tb(s,0.4,5.10,9.2,0.30,"💬 「我 认识 ___」",sz=12,b=True,c=STAR,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"5-6 分钟 — 我会认 5 个 词")

# 11-12. 我会写 — 星星 + 星座
def write_slide(emoji,word_cn,word_en,chars,color):
    s=ns(); bg(s,CREAM); hb(s,f"✏️ 我会写 · {word_cn}  I Can Write · {word_en}",color)
    tb(s,0.4,0.85,9.2,0.36,f"{emoji} 一起来写「{word_cn}」!",sz=20,b=True,c=color,a=PP_ALIGN.CENTER)
    tb(s,0.4,1.25,9.2,0.26,f"Practice writing {word_cn} ({word_en})",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    if len(chars)==2:
        tianzi(s,0.55,1.65,2.20,chars[0][0],color,pinyin=chars[0][1],char_sz=120)
        tianzi(s,2.95,1.65,2.20,chars[1][0],color,pinyin=chars[1][1],char_sz=120)
    else:
        tianzi(s,1.30,1.65,2.95,chars[0][0],color,pinyin=chars[0][1],char_sz=160)
    panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.30),Inches(1.65),Inches(4.30),Inches(2.85))
    panel.fill.solid(); panel.fill.fore_color.rgb=WHITE; panel.line.color.rgb=color; panel.line.width=Pt(2.5)
    head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.30),Inches(1.65),Inches(4.30),Inches(0.50))
    head.fill.solid(); head.fill.fore_color.rgb=color; head.line.fill.background()
    tb(s,5.45,1.72,4.10,0.40,"✏️ 怎么写  How to Write",sz=13,b=True,c=WHITE)
    for i,(ch,py,hint_cn,hint_en) in enumerate(chars):
        y=2.30+i*0.95
        tb(s,5.45,y,4.10,0.35,f"📐「{ch}」 — {py}",sz=13,b=True,c=DARK)
        tb(s,5.45,y+0.32,4.10,0.30,hint_cn,sz=10,b=True,c=color)
        tb(s,5.45,y+0.58,4.10,0.26,hint_en,sz=9,c=GRAY)
    pf=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(4.65),Inches(9.2),Inches(0.85))
    pf.fill.solid(); pf.fill.fore_color.rgb=WARM; pf.line.color.rgb=color; pf.line.width=Pt(2)
    tb(s,0.55,4.72,9.0,0.32,f"📝 在 田字格 里 写 3 遍",sz=12,b=True,c=color,a=PP_ALIGN.CENTER)
    tb(s,0.55,5.08,9.0,0.32,f"💬 「我 会 写「{word_cn}」!」",sz=13,b=True,c=DARK,a=PP_ALIGN.CENTER)
    return s

s=write_slide("⭐","星星","Stars",[
    ("星","xīng","9 笔 — 上面「日」+ 下面「生」","Sun on top + 生 below"),
    ("星","xīng","写 两遍 一样 — 重复 词","Repeat character"),
],STAR); n+=1; pn(s,n)
notes(s,"5-6 分钟 — 写 星星")

s=write_slide("✨","星座","Constellation",[
    ("星","xīng","9 笔 — 上「日」+ 下「生」","Sun + life"),
    ("座","zuò","10 笔 — 「广」+「人人」+「土」","Roof + two people + earth"),
],GOLD); n+=1; pn(s,n)
notes(s,"5-6 分钟 — 写 星座")

# 13. Session 3 Divider
s=div("Session 3  下午 3:00–4:30","🛠️ 项目课 · 星座 创作  ·  90 min",NEBULA,"⭐"); n+=1; pn(s,n)

# 14. Projects overview
s=ns(); bg(s,CREAM); hb(s,"🛠️ 3 个 项目  3 Projects",NEBULA)
tb(s,0.4,0.85,9.2,0.32,"选 一个 — 自己 来 当 星座 创造者!",sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
projects=[
    ("🔦","星座 手电筒 投影","Constellation Flashlight","纸杯 + 针孔 → 投影 到 墙上!",STAR),
    ("🧵","棉签 星座 拼图","Q-tip Constellations","棉签 + 黑纸 — 拼 你 的 星座",NEBULA),
    ("✍️","创造 你 的 星座 故事","Make YOUR Constellation","起 个 名字 + 写 个 故事",PINK),
]
for i,(em,cn,en,d,cl) in enumerate(projects):
    x=0.4+i*3.10
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.55),Inches(2.95),Inches(3.20))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    tb(s,x+0.05,1.70,2.85,1.20,em,sz=78,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.95,2.85,0.42,cn,sz=17,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.38,2.85,0.30,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.10,3.78,2.75,0.85,d,sz=12,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,4.90,9.2,0.30,"⭐ 完成 后 — 给 你 的 星座 取 名字!",sz=12,b=True,c=NEBULA,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"3 分钟 — 介绍 项目:\n• 项目 1 (手电筒): 需要 提前 准备 纸杯 + 针\n• 项目 2 (棉签): 棉签 + 黑色卡纸\n• 项目 3 (写故事): 适合 喜欢 写字 的 学生")

# 15. Project 1 details — Constellation flashlight
s=ns(); bg(s,CREAM); hb(s,"🔦 项目 1 · 星座 手电筒  Constellation Flashlight",STAR)
ib(s,0.4,0.90,4.5,3.95,"🖼️ 手电筒 投影 示例")
panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.10),Inches(0.90),Inches(4.50),Inches(3.95))
panel.fill.solid(); panel.fill.fore_color.rgb=WHITE; panel.line.color.rgb=STAR; panel.line.width=Pt(2.5)
head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.10),Inches(0.90),Inches(4.50),Inches(0.50))
head.fill.solid(); head.fill.fore_color.rgb=STAR; head.line.fill.background()
tb(s,5.25,0.97,4.30,0.40,"📝 怎么 做  How to Make",sz=14,b=True,c=DARK)
steps=[
    ("1️⃣","拿 一个 纸杯 (paper cup)"),
    ("2️⃣","在 杯底 用 针 戳 小孔 — 摆 成 星座"),
    ("3️⃣","可以 是 大熊座 / 北斗七星 / 自创"),
    ("4️⃣","把 手电筒 放进 杯里 — 照 墙上!"),
    ("5️⃣","关 灯 — 看 你 的 星座 在 墙上 发光!"),
]
for i,(num,txt) in enumerate(steps):
    y=1.55+i*0.62
    tb(s,5.25,y,0.40,0.40,num,sz=16,b=True,c=STAR)
    tb(s,5.75,y+0.04,3.80,0.35,txt,sz=11,c=DARK)
tip=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(4.95),Inches(9.2),Inches(0.55))
tip.fill.solid(); tip.fill.fore_color.rgb=NIGHT; tip.line.fill.background()
tb(s,0.55,5.02,9.0,0.32,"💡 安全 提示: 老师 帮 戳 孔 — 不要 自己 用 针!",sz=12,b=True,c=STAR,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"项目 1 · 25 分钟:\n• 材料: 纸杯 + 针 (老师 操作) + 手电筒 (or 手机 闪光灯)\n• 提前 准备: 大熊座 / 牛郎织女 / 北斗 的 图案 模板\n• 课堂 关 灯 试 — 学生 会 惊叹!")

# 16. Project 2 + 3 (combined)
s=ns(); bg(s,CREAM); hb(s,"🧵✍️ 项目 2 + 3  Two More Options",NEBULA)
left=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.90),Inches(4.55),Inches(4.10))
left.fill.solid(); left.fill.fore_color.rgb=WHITE; left.line.color.rgb=NEBULA; left.line.width=Pt(2.5)
head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.90),Inches(4.55),Inches(0.50))
head.fill.solid(); head.fill.fore_color.rgb=NEBULA; head.line.fill.background()
tb(s,0.55,0.97,4.30,0.40,"🧵 项目 2 · 棉签 星座 拼图",sz=14,b=True,c=WHITE)
items=[
    "🎨 材料: 黑色卡纸 + 棉签 + 白色 颜料 / 贴纸",
    "1️⃣ 想 — 你 要 拼 什么 星座?",
    "2️⃣ 用 白色 圆 贴纸 贴 星星 位置",
    "3️⃣ 用 棉签 蘸 白颜料 — 连 起来",
    "4️⃣ 写 上 星座 的 名字",
    "✨ 给 你 的 队 / 妈妈 / 朋友!",
]
for i,line in enumerate(items):
    tb(s,0.55,1.55+i*0.42,4.30,0.40,line,sz=11,b=True,c=DARK)
right=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.05),Inches(0.90),Inches(4.55),Inches(4.10))
right.fill.solid(); right.fill.fore_color.rgb=WHITE; right.line.color.rgb=PINK; right.line.width=Pt(2.5)
head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.05),Inches(0.90),Inches(4.55),Inches(0.50))
head.fill.solid(); head.fill.fore_color.rgb=PINK; head.line.fill.background()
tb(s,5.20,0.97,4.30,0.40,"✍️ 项目 3 · 创造 你 的 星座 故事",sz=14,b=True,c=WHITE)
parts=[
    "🌟 假装 你 是 古人 — 抬头 看 星星",
    "1️⃣ 把 几 颗 星 连 起来 — 看 像 什么?",
    "2️⃣ 给 它 取 一个 名字 (你的 名字 都 行!)",
    "3️⃣ 编 一个 小 故事:",
    "「很久 以前, 有 一个 ___」",
    "4️⃣ 画 + 写 — 老师 帮 K-1!",
]
for i,line in enumerate(parts):
    tb(s,5.20,1.55+i*0.42,4.30,0.40,line,sz=11,b=True,c=DARK)
n+=1; pn(s,n)
notes(s,"25 分钟:\n• 项目 2 适合 喜欢 手工 的 学生\n• 项目 3 适合 喜欢 创作 的 学生\n• 老师 走动 — 鼓励 每个 想法")

# 17. Share & close
s=ns(); bg(s,CREAM); hb(s,"🎤 分享 + 再见!",COSMIC)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.95),Inches(9.2),Inches(2.10))
sh.fill.solid(); sh.fill.fore_color.rgb=NIGHT; sh.line.color.rgb=STAR; sh.line.width=Pt(3)
tb(s,0.55,1.05,9.0,0.40,"💬 句型 — 分享 你 的 星座!",sz=14,b=True,c=STAR,a=PP_ALIGN.CENTER)
tb(s,0.55,1.55,9.0,0.50,"「我 的 星座 叫 ___」",sz=22,b=True,c=STAR,a=PP_ALIGN.CENTER)
tb(s,0.55,2.10,9.0,0.50,"「它 看 起来 像 ___」",sz=22,b=True,c=STAR,a=PP_ALIGN.CENTER)
preview=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(3.25),Inches(9.2),Inches(2.05))
preview.fill.solid(); preview.fill.fore_color.rgb=WHITE; preview.line.color.rgb=COSMIC; preview.line.width=Pt(2.5)
tb(s,0.55,3.40,9.0,0.40,"🔮 下次 见 (Day 3):",sz=14,b=True,c=COSMIC,a=PP_ALIGN.CENTER)
tb(s,0.55,3.85,9.0,0.40,"🚀 登月 + 火星 计划 + 火箭 实验!",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.55,4.30,9.0,0.30,"Moon landing + Mars + rocket experiments",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,0.55,4.70,9.0,0.30,"👋 今晚 抬头 — 找 找 银河 + 牛郎织女!",sz=11,b=True,c=COSMIC,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"5 分钟 — 分享:\n• 3-4 学生 用 句型 分享 自己 的 星座\n• 「家庭 作业」: 今晚 看 银河")

out=os.path.join(os.path.dirname(__file__),"day2_constellations.pptx")
prs.save(out)
print(f"Saved {out}  ({len(prs.slides)} slides)")
