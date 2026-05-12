#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
我的职业梦想 — Day 4: 社区小帮手 (Community Helpers)
3 sessions × 50 min, Workshop Model (5-phase frame per session):
  🔥 Hook (5) → 📚 Mini-Lesson (10) → 🎯 Active Practice (20-25) → 🌱 Apply (5-10) → 🎤 Share & Close (5)
8 helpers: 医生 · 消防员 · 老师 · 警察 · 邮递员 · 清洁工 · 图书管理员 · 厨师
Session 3 — 2 projects: P1 最重要的工作 (debate) · P2 社区救援大挑战 (rescue match)
我会认: 老师 · 学校 · 医院 · 消防员 · 厨师    我会写: 老师 · 学校
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
prs.slide_width  = Inches(10)
prs.slide_height = Inches(5.625)
W, H = prs.slide_width, prs.slide_height

# === Palette ===
HELP   = RGBColor(0x2E,0x7D,0x32)   # community helper green
HEART  = RGBColor(0xE5,0x3E,0x5E)   # care pink-red
NAVY   = RGBColor(0x1E,0x3A,0x5F)
IDEA   = RGBColor(0xF5,0xC2,0x42)
LAB    = RGBColor(0xE5,0x3E,0x3E)
GREEN  = RGBColor(0x2E,0x7D,0x32)
PURPLE = RGBColor(0x7B,0x1F,0xA2)
RUST   = RGBColor(0xC4,0x52,0x2A)
CREAM  = RGBColor(0xFF,0xF8,0xE7)
WARM   = RGBColor(0xFF,0xF3,0xE0)
BROWN  = RGBColor(0x6B,0x44,0x23)
WHITE  = RGBColor(0xFF,0xFF,0xFF)
DARK   = RGBColor(0x2C,0x2C,0x2C)
GRAY   = RGBColor(0x88,0x88,0x88)
LGRAY  = RGBColor(0xBB,0xBB,0xBB)
IMGBG  = RGBColor(0xE8,0xE8,0xE8)
OK     = RGBColor(0x38,0x8E,0x3C)
SKY    = RGBColor(0x1F,0x77,0xB4)
GOLD   = RGBColor(0xF5,0xA6,0x23)

# Phase colors
PH_HOOK   = LAB
PH_MINI   = HELP
PH_ACTIVE = RUST
PH_APPLY  = GREEN
PH_CLOSE  = IDEA

# Per-helper colors (8 helpers)
DOC    = RGBColor(0xC8,0x25,0x3E)   # 医生 wine red
FIRE   = RGBColor(0xD3,0x18,0x18)   # 消防员 fire red
TEACH  = RGBColor(0x43,0xA0,0x47)   # 老师 green
POLICE = RGBColor(0x15,0x65,0xC0)   # 警察 police blue
MAIL   = RGBColor(0xF5,0x7C,0x00)   # 邮递员 postal orange
CLEAN  = RGBColor(0x00,0x89,0x7B)   # 清洁工 teal
LIB    = RGBColor(0x6A,0x1B,0x9A)   # 图书管理员 purple
CHEF   = RGBColor(0xFF,0x8F,0x00)   # 厨师 amber

# Team colors
T_RED    = RGBColor(0xE5,0x3E,0x3E)
T_BLUE   = RGBColor(0x1F,0x77,0xB4)
T_GREEN  = RGBColor(0x2E,0x7D,0x32)
T_YELLOW = RGBColor(0xF5,0xA6,0x23)

# === Helpers ===
def ns(): return prs.slides.add_slide(prs.slide_layouts[6])
def tb(s,l,t,w,h,txt,sz=18,b=False,c=DARK,a=None):
    bx=s.shapes.add_textbox(Inches(l),Inches(t),Inches(w),Inches(h)); tf=bx.text_frame; tf.word_wrap=True
    p=tf.paragraphs[0]
    if a: p.alignment=a
    r=p.add_run(); r.text=txt; r.font.size=Pt(sz); r.font.bold=b; r.font.color.rgb=c; r.font.name='KaiTi'
    return tf
def ap(tf,txt,sz=18,b=False,c=DARK,a=None):
    p=tf.add_paragraph()
    if a: p.alignment=a
    r=p.add_run(); r.text=txt; r.font.size=Pt(sz); r.font.bold=b; r.font.color.rgb=c; r.font.name='KaiTi'
def bg(s,c):
    sh=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,W,H); sh.fill.solid(); sh.fill.fore_color.rgb=c; sh.line.fill.background()
    sp=sh._element; sp.getparent().remove(sp); s.shapes._spTree.insert(2,sp)
def hb(s,txt,c=HELP,t=0.15):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(t),Inches(9.4),Inches(0.55))
    sh.fill.solid(); sh.fill.fore_color.rgb=c; sh.line.fill.background()
    tb(s,0.4,t+0.03,9.2,0.5,txt,sz=20,b=True,c=WHITE)
def pn(s,n): tb(s,9.0,5.25,0.8,0.3,str(n),sz=10,c=GRAY,a=PP_ALIGN.RIGHT)
def pill(s,l,t,w,h,txt,c,sz=14):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(l),Inches(t),Inches(w),Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb=c; sh.line.fill.background()
    tb(s,l+0.05,t+h/2-0.18,w-0.10,0.4,txt,sz=sz,b=True,c=WHITE,a=PP_ALIGN.CENTER)
def notes(s,text):
    nf=s.notes_slide.notes_text_frame; lines=text.split("\n"); nf.text=lines[0]
    for line in lines[1:]:
        p=nf.add_paragraph(); p.text=line
def div(title,sub,color,emoji=""):
    s=ns(); bg(s,color)
    tb(s,1,1.5,8,1.2,f"{emoji} {title}",sz=42,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,1,2.8,8,0.8,sub,sz=22,c=WHITE,a=PP_ALIGN.CENTER)
    return s
def sentence_frame_bar(s,t,frame_cn,frame_en,accent=IDEA):
    if t > 4.95: t = 4.95
    sf=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(t),Inches(9.4),Inches(0.65))
    sf.fill.solid(); sf.fill.fore_color.rgb=WARM; sf.line.color.rgb=accent; sf.line.width=Pt(2)
    tb(s,0.5,t+0.1,1.7,0.4,"💬 我来说",sz=14,b=True,c=accent)
    tb(s,2.0,t+0.07,7.6,0.3,frame_cn,sz=14,b=True,c=DARK)
    tb(s,2.0,t+0.32,7.6,0.3,frame_en,sz=10,c=GRAY)

def phase_marker(emoji,phase_cn,phase_en,time_min,color,what_cn,what_en):
    s=ns(); bg(s,color)
    tb(s,1,0.85,8,0.7,emoji,sz=80,a=PP_ALIGN.CENTER)
    tb(s,1,1.85,8,0.6,phase_cn,sz=38,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,1,2.50,8,0.4,phase_en,sz=16,c=IDEA,a=PP_ALIGN.CENTER)
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(3.5),Inches(3.10),Inches(3.0),Inches(0.55))
    sh.fill.solid(); sh.fill.fore_color.rgb=IDEA; sh.line.fill.background()
    tb(s,3.5,3.18,3.0,0.4,f"⏱  {time_min} 分钟",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
    tb(s,1,4.00,8,0.5,what_cn,sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,1,4.55,8,0.35,what_en,sz=12,c=IDEA,a=PP_ALIGN.CENTER)
    return s

def score_badge(s):
    teams=[("🔴",T_RED),("🔵",T_BLUE),("🟢",T_GREEN),("🟡",T_YELLOW)]
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(7.40),Inches(0.78),Inches(2.30),Inches(0.32))
    sh.fill.solid(); sh.fill.fore_color.rgb=WARM; sh.line.color.rgb=GRAY; sh.line.width=Pt(0.75)
    tb(s,7.45,0.81,0.45,0.28,"🏆",sz=12,a=PP_ALIGN.CENTER)
    for i,(em,cl) in enumerate(teams):
        tb(s,7.85+i*0.45,0.81,0.40,0.28,f"{em}__",sz=10,b=True,c=cl,a=PP_ALIGN.CENTER)

def group_label(s,t=0.78):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.30),Inches(t),Inches(1.80),Inches(0.32))
    sh.fill.solid(); sh.fill.fore_color.rgb=PH_ACTIVE; sh.line.fill.background()
    tb(s,0.30,t+0.03,1.80,0.28,"👥 分组任务",sz=11,b=True,c=WHITE,a=PP_ALIGN.CENTER)

def tianzi(s,x,y,size,char,color,pinyin=None,char_sz=130):
    box=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(x),Inches(y),Inches(size),Inches(size))
    box.fill.solid(); box.fill.fore_color.rgb=WHITE
    box.line.color.rgb=color; box.line.width=Pt(3)
    mid_x=x+size/2; mid_y=y+size/2; lw=0.015
    v=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(mid_x-lw/2),Inches(y),Inches(lw),Inches(size))
    v.fill.solid(); v.fill.fore_color.rgb=LGRAY; v.line.fill.background()
    h=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(x),Inches(mid_y-lw/2),Inches(size),Inches(lw))
    h.fill.solid(); h.fill.fore_color.rgb=LGRAY; h.line.fill.background()
    tb(s,x,y,size,size,char,sz=char_sz,b=True,c=color,a=PP_ALIGN.CENTER)
    if pinyin:
        tb(s,x,y+size+0.05,size,0.30,pinyin,sz=12,b=True,c=GRAY,a=PP_ALIGN.CENTER)

n=0

# ============================================================
# 1. COVER
# ============================================================
s=ns(); bg(s,CREAM)
tb(s,1,0.40,8,0.55,"我的职业梦想 · My Dream Career",sz=22,b=True,c=HELP,a=PP_ALIGN.CENTER)
tb(s,1,0.95,8,0.4,"Day 4 · 社区小帮手  Community Helpers",sz=16,b=True,c=DARK,a=PP_ALIGN.CENTER)
# 8 helper bubbles arranged in 2 rows of 4
helpers_cover=[("🩺",DOC,"医生"),("🚒",FIRE,"消防员"),("📚",TEACH,"老师"),("👮",POLICE,"警察"),
               ("✉️",MAIL,"邮递员"),("🧹",CLEAN,"清洁工"),("📖",LIB,"管理员"),("👨‍🍳",CHEF,"厨师")]
for i,(em,cl,name) in enumerate(helpers_cover):
    col=i%4; row=i//4
    x=1.0+col*2.0; y=1.55+row*1.55
    sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x),Inches(y),Inches(1.55),Inches(1.30))
    sh.fill.solid(); sh.fill.fore_color.rgb=cl; sh.line.color.rgb=IDEA; sh.line.width=Pt(3)
    tb(s,x,y+0.12,1.55,0.75,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x,y+0.85,1.55,0.32,name,sz=11,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,1,4.75,8,0.45,"❤️ 社区小帮手 = 帮助别人 · 解决问题 · 让生活更安全方便",sz=14,b=True,c=HELP,a=PP_ALIGN.CENTER)
tb(s,1,5.20,8,0.25,"3 sessions × 50 min",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"30 秒 hook:\n• 「今天 — 我们来认识 8 位『社区小帮手』 — 他们就在我们身边!」\n• 全班分成 4 队: 🔴 红 / 🔵 蓝 / 🟢 绿 / 🟡 黄\n• 准备道具: 8 张『职业卡』(每位帮手一张) + 8 张『问题卡』(配对游戏用)")

# ============================================================
# 2. 5-DAY PREVIEW
# ============================================================
s=ns(); bg(s,CREAM); hb(s,"🗺️ 5 天的职业之旅  Our 5-Day Career Journey",NAVY)
tb(s,0.4,0.85,9.2,0.34,"今天是第 4 天 — 认识身边每天在帮助我们的人!",sz=15,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.20,9.2,0.28,"Today is Day 4 — meet the helpers around us every day!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
days_preview=[
    ("Day 1","认识职业世界","Discover Careers","🌍",NAVY,"8 个职业"),
    ("Day 2","小小科学家","Little Scientists","🔬",SKY,"⭐ 爱迪生"),
    ("Day 3","小小企业家","Little Entrepreneurs","💡",GOLD,"⭐ 乔布斯"),
    ("Day 4","社区小帮手","Community Helpers","❤️",HELP,"⭐ 今天!"),
    ("Day 5","AI 与未来","AI & the Future","🤖",PURPLE,"⭐ AI 公司"),
]
for i,(label,cn,en,em,cl,spotlight) in enumerate(days_preview):
    x=0.3+i*1.92
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.55),Inches(1.82),Inches(3.45))
    is_today=(i==3)
    sh.fill.solid(); sh.fill.fore_color.rgb=cl if is_today else WHITE
    sh.line.color.rgb=cl; sh.line.width=Pt(3.5 if is_today else 2.5)
    badge=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x+0.1),Inches(1.65),Inches(0.55),Inches(0.55))
    badge.fill.solid(); badge.fill.fore_color.rgb=WHITE if is_today else cl; badge.line.fill.background()
    tb(s,x+0.1,1.74,0.55,0.4,str(i+1),sz=18,b=True,c=cl if is_today else WHITE,a=PP_ALIGN.CENTER)
    tb(s,x+0.7,1.7,1.1,0.3,label,sz=11,b=True,c=WHITE if is_today else cl)
    tb(s,x+0.05,2.30,1.72,0.7,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.05,1.72,0.4,cn,sz=15,b=True,c=WHITE if is_today else DARK,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.45,1.72,0.3,en,sz=10,c=IDEA if is_today else GRAY,a=PP_ALIGN.CENTER)
    sep=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(x+0.25),Inches(3.85),Inches(1.32),Inches(0.02))
    sep.fill.solid(); sep.fill.fore_color.rgb=WHITE if is_today else cl; sep.line.fill.background()
    tb(s,x+0.05,4.00,1.72,0.85,spotlight,sz=11,b=True,c=WHITE if is_today else cl,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,5.20,"今天我要认识 ___ 位小帮手!","Today I'll meet ___ community helpers!",accent=HELP)
n+=1; pn(s,n)
notes(s,"30 秒 — 复习一下整个 unit:\n• 「Day 1 我们认识了 8 个职业, Day 2 见了科学家工程师, Day 3 当了小老板……」\n• 「今天 Day 4 — 我们看看生活中每天在帮助我们的人!」\n• 不展开 — 接下来 Today's Mission")

# ============================================================
# 3. TODAY'S MISSION
# ============================================================
s=ns(); bg(s,CREAM); hb(s,"🧭 今天的任务  Today's Mission",HELP)
tb(s,0.4,0.85,9.2,0.40,"❤️ 5 个任务 — 认识 8 位小帮手 + 做 2 个项目!",sz=18,b=True,c=HELP,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.28,"5 missions — meet 8 helpers + complete 2 projects",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
missions=[
    ("1","🤝","小帮手是谁?","Who's a helper",HELP),
    ("2","👋","认识 8 位","Meet 8 helpers",HEART),
    ("3","🎯","谁来帮?","Who helps",NAVY),
    ("4","📖","我会认/写","Read & Write",PURPLE),
    ("5","🛡️","救援大挑战","Rescue Challenge",FIRE),
]
for i,(num,em,cn,en,cl) in enumerate(missions):
    x=0.4+i*1.90
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.85),Inches(1.80),Inches(2.30))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(2.5)
    badge=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x+0.1),Inches(1.95),Inches(0.55),Inches(0.55))
    badge.fill.solid(); badge.fill.fore_color.rgb=cl; badge.line.fill.background()
    tb(s,x+0.1,2.04,0.55,0.4,num,sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.65,1.72,0.75,em,sz=46,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.45,1.72,0.36,cn,sz=14,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.80,1.72,0.28,en,sz=9,c=GRAY,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.30,"我想帮助 ___ 。 我可以 ___ 。","I want to help ___. I can ___.",accent=HELP)
n+=1; pn(s,n)
notes(s,"1 分钟 — 预告今天 5 个任务:\n• 快速指 5 张卡 — 让学生知道今天要做什么\n• 关键句型 (今天的核心): 「我想帮助 ___, 我可以 ___」\n• 这两句话会贯穿全天 — 让学生反复练")

# ============================================================
# 4. SESSION 1 DIVIDER
# ============================================================
s=div("Session 1  上午 11:00–11:50","❤️ 故事课  认识 8 位社区小帮手  ·  50 min",HELP,"🌟"); n+=1; pn(s,n)

# ============================================================
# 5. SESSION 1 · LEARNING GOALS
# ============================================================
s=ns(); bg(s,CREAM); hb(s,"🎯 这节课的学习目标  Session 1 Learning Goals",HELP)
tb(s,0.4,0.85,9.2,0.30,"上完这节课, 你会……",sz=13,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.15,9.2,0.22,"By the end of this session, you will be able to…",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
goals=[
    ("1","🤝","懂得「社区小帮手」是什么 — 在社区里帮助别人、解决问题、让生活更安全方便的人。",HELP),
    ("2","👋","认识 8 位常见小帮手: 医生、消防员、老师、警察、邮递员、清洁工、图书管理员、厨师。",NAVY),
    ("3","🎯","说出每位小帮手帮谁、解决什么问题、用什么工具。",PURPLE),
    ("4","🛟","在情境里挑出合适的小帮手 — 知道什么时候找谁帮忙。",FIRE),
    ("5","💬","用句型表达: 「我想帮助 ___」 「我可以 ___」 — 你也是小帮手!",CHEF),
]
for i,(num,em,text,c) in enumerate(goals):
    y=1.42+i*0.80
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(y),Inches(9.2),Inches(0.74))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=c; sh.line.width=Pt(2.5)
    nb=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(0.55),Inches(y+0.13),Inches(0.48),Inches(0.48))
    nb.fill.solid(); nb.fill.fore_color.rgb=c; nb.line.fill.background()
    tb(s,0.55,y+0.18,0.48,0.40,num,sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,1.15,y+0.13,0.55,0.50,em,sz=26,a=PP_ALIGN.CENTER)
    tb(s,1.80,y+0.08,7.75,0.62,text,sz=11,b=True,c=DARK)
n+=1; pn(s,n)
notes(s,"1-2 分钟 — 学习目标预告:\n• 快速过一遍 5 个目标\n• 不细讲 — 让学生知道大方向\n• 关键: 「这节课结束 — 你能说出 8 位小帮手 + 知道什么时候找谁!」")

# ============================================================
# === SESSION 1 · PHASE 1: HOOK (5 min) ===
# ============================================================
s=phase_marker("🔥","HOOK","Wake Up!",5,PH_HOOK,"想象一下: 如果没有他们……","Imagine: what if they were gone?")
n+=1; pn(s,n)

# Hook 1 — Imagine without them
s=ns(); bg(s,CREAM); hb(s,"😱 如果没有他们呢?  What If They Were Gone?",PH_HOOK)
tb(s,0.4,0.85,9.2,0.36,"想象一下 — 如果世界上没有这些人, 会怎么样?",sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.18,9.2,0.24,"Imagine — what if these people didn't exist?",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
scenes=[
    ("🔥","房子着火了…","House on fire…","但 没有 消防员!",FIRE),
    ("🤒","小朋友 生病了…","Kid is sick…","但 没有 医生!",DOC),
    ("📚","小朋友想 学新东西…","Want to learn…","但 没有 老师!",TEACH),
    ("🗑️","街上 都是垃圾…","Streets full of trash…","但 没有 清洁工!",CLEAN),
]
for i,(em,cn,en,bad,cl) in enumerate(scenes):
    col=i%2; row=i//2
    x=0.4+col*4.65; y=1.55+row*1.55
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(4.55),Inches(1.45))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    tb(s,x+0.10,y+0.12,1.0,1.20,em,sz=42,a=PP_ALIGN.CENTER)
    tb(s,x+1.20,y+0.10,3.30,0.36,cn,sz=14,b=True,c=DARK)
    tb(s,x+1.20,y+0.48,3.30,0.26,en,sz=9,c=GRAY)
    tb(s,x+1.20,y+0.82,3.30,0.50,bad,sz=14,b=True,c=cl)
sentence_frame_bar(s,4.70,"如果没有 ___, 就 ___ !","Without ___, then ___!",accent=PH_HOOK)
n+=1; pn(s,n)
notes(s,"🔥 HOOK · 4-5 分钟 — 想象游戏:\n• 老师戏剧化地念 4 个场景:\n  - 「房子着火了 — 但…… 没有消防员!」(惊恐表情)\n  - 「你发烧了 — 但…… 没有医生!」\n• 每个场景后停顿: 「怎么办?!」\n• 让学生小声议论或抢答\n• 引出: 「幸好 — 我们有这些人! 他们就是 — 社区小帮手!」")

# Hook 2 — Reveal: 社区小帮手!
s=ns(); bg(s,CREAM); hb(s,"❤️ 他们都是 — 社区小帮手!  They Are All — Community Helpers!",HELP)
tb(s,0.4,0.85,9.2,0.42,"幸好我们有他们 — 每天都在帮助大家!",sz=18,b=True,c=HELP,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.28,"We're lucky to have them — helping us every day!",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
# Big centered banner
banner=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.5),Inches(1.75),Inches(9.0),Inches(2.50))
banner.fill.solid(); banner.fill.fore_color.rgb=HELP; banner.line.color.rgb=IDEA; banner.line.width=Pt(3)
tb(s,0.5,1.85,9.0,0.55,"社区小帮手",sz=42,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,0.5,2.45,9.0,0.36,"Community Helpers",sz=16,b=True,c=IDEA,a=PP_ALIGN.CENTER)
tb(s,0.5,2.90,9.0,0.40,"= 在社区里 帮助别人、解决问题、让生活更安全方便的人",sz=14,b=True,c=IDEA,a=PP_ALIGN.CENTER)
tb(s,0.5,3.30,9.0,0.30,"People who help others, solve problems, and make life safer + easier.",sz=10,c=WARM,a=PP_ALIGN.CENTER)
tb(s,0.5,3.70,9.0,0.40,"🌟 他们就在我们身边 — 每天 都在 工作!",sz=14,b=True,c=IDEA,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.55,"___ 是社区小帮手 — 因为 ___。","___ is a community helper — because ___.",accent=HELP)
n+=1; pn(s,n)
notes(s,"2-3 分钟 — 揭晓答案:\n• 老师放慢节奏念定义\n• 全班跟读 3 遍: 「社区小帮手 = 帮助别人 + 解决问题 + 让生活更安全方便」\n• 关键: 「他们每天都在工作 — 我们才能安心生活!」\n• 引出下一页: 「这些小帮手 — 他们有 4 个共同点!」")

# ============================================================
# === SESSION 1 · PHASE 2: MINI-LESSON (10 min) ===
# ============================================================
s=phase_marker("📚","MINI-LESSON","Learn Together",10,PH_MINI,"小帮手的 4 个共同点 + 认识 8 位","4 traits + meet 8 helpers")
n+=1; pn(s,n)

# What is a helper? (4 traits)
s=ns(); bg(s,CREAM); hb(s,"💡 小帮手有 4 个共同点  4 Things All Helpers Do",HELP)
tb(s,0.4,0.85,9.2,0.40,"每一位小帮手 — 都做这 4 件事!",sz=16,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.25,9.2,0.28,"Every helper does these 4 things!",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
traits=[
    ("🤝","帮助别人","Help Others","看见有人需要 → 出手!",HELP),
    ("🛠️","解决问题","Solve Problems","用工具 + 方法",NAVY),
    ("🛡️","让生活更安全","Make It Safer","让大家不害怕",FIRE),
    ("✨","让生活更方便","Make It Easier","让大家更轻松",CHEF),
]
for i,(em,cn,en,detail,cl) in enumerate(traits):
    x=0.4+i*2.30
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.65),Inches(2.20),Inches(3.0))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(2.5)
    tb(s,x+0.05,1.80,2.10,0.8,em,sz=46,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.70,2.10,0.42,cn,sz=17,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.12,2.10,0.28,en,sz=9,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.10,3.50,2.00,1.00,detail,sz=11,b=True,c=DARK,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.80,"小帮手 = 帮助 + 解决 + 安全 + 方便。","Helper = help + solve + safer + easier.",accent=HELP)
n+=1; pn(s,n)
notes(s,"3-4 分钟 — 4 个共同点:\n• 老师念 4 个 trait + 配手势:\n  - 帮助别人 (双手向外伸)\n  - 解决问题 (拳头)\n  - 让生活安全 (双手抱护)\n  - 让生活方便 (大拇指)\n• 全班跟读 + 跟做手势 3 遍\n• 关键 message: 「学完今天 — 你也能当小帮手!」")

# 8 helpers gallery
s=ns(); bg(s,CREAM); hb(s,"👋 8 位社区小帮手  Meet the 8 Helpers",HELP)
tb(s,0.4,0.85,9.2,0.34,"他们都帮助谁? 解决什么问题? 一起认识!",sz=13,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.18,9.2,0.24,"Who do they help? What do they solve? Let's meet them!",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
helpers8=[
    ("🩺","医生","Doctor",DOC),
    ("🚒","消防员","Firefighter",FIRE),
    ("📚","老师","Teacher",TEACH),
    ("👮","警察","Police",POLICE),
    ("✉️","邮递员","Mail carrier",MAIL),
    ("🧹","清洁工","Cleaner",CLEAN),
    ("📖","图书管理员","Librarian",LIB),
    ("👨‍🍳","厨师","Chef",CHEF),
]
for i,(em,cn,en,cl) in enumerate(helpers8):
    col=i%4; row=i//4
    x=0.4+col*2.32; y=1.55+row*1.75
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(2.22),Inches(1.60))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    tb(s,x+0.05,y+0.10,2.12,0.70,em,sz=42,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,y+0.85,2.12,0.36,cn,sz=14,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,y+1.22,2.12,0.26,en,sz=9,c=GRAY,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"3-4 分钟 — 一次性把 8 位帮手亮出来:\n• 一个一个指 — 学生齐读名字 3 遍\n• 用提问检测: 「这个是谁?」 (指消防员)\n• 引出: 「接下来 — 我们把他们分成 3 组, 看看他们具体做什么!」")

# Deep-dive Group A: Safety (医生, 消防员, 警察)
s=ns(); bg(s,CREAM); hb(s,"🛡️ 让我们安全的人  Keep Us Safe",FIRE)
tb(s,0.4,0.85,9.2,0.36,"急事来了 — 找他们! 他们让大家不害怕。",sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.18,9.2,0.24,"In an emergency — find them! They make us feel safe.",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
safety=[
    ("🩺","医生","Doctor","🤒 生病的人","🩹 治病","🏥 医院",DOC),
    ("🚒","消防员","Firefighter","🔥 火灾","💧 灭火 + 救人","🚒 消防车",FIRE),
    ("👮","警察","Police","🚓 大家 (安全)","🛡️ 抓坏人 + 保护","🚓 警车",POLICE),
]
for i,(em,cn,en,who,what,where,cl) in enumerate(safety):
    x=0.4+i*3.10
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.55),Inches(2.95),Inches(3.10))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    tb(s,x+0.05,1.65,2.85,0.85,em,sz=58,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.55,2.85,0.40,cn,sz=18,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.95,2.85,0.28,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    sep=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(x+0.30),Inches(3.30),Inches(2.35),Inches(0.02))
    sep.fill.solid(); sep.fill.fore_color.rgb=cl; sep.line.fill.background()
    tb(s,x+0.10,3.40,2.75,0.30,f"👥 帮谁: {who}",sz=10,b=True,c=DARK)
    tb(s,x+0.10,3.72,2.75,0.30,f"🎯 解决: {what}",sz=10,b=True,c=DARK)
    tb(s,x+0.10,4.04,2.75,0.30,f"📍 在哪: {where}",sz=10,b=True,c=DARK)
sentence_frame_bar(s,4.78,"___ 帮助 ___ — 解决 ___ 的问题。","___ helps ___ — solves ___.",accent=FIRE)
n+=1; pn(s,n)
notes(s,"4-5 分钟 — 让我们安全的人:\n• 老师讲: 「当我们生病、着火、危险 — 这 3 位是我们的『紧急』帮手」\n• 念 911 / 119 / 110 — 中美对照:\n  - 🇺🇸 911 = 所有 emergency (police / fire / medical)\n  - 🇨🇳 110 = 警察, 119 = 消防, 120 = 医生\n• 让学生用句型造句 — 抽 2-3 个学生上台说\n• 关键: 「记住电话 — 真的有事要找他们!」")

# Deep-dive Group B: Learning (老师, 图书管理员)
s=ns(); bg(s,CREAM); hb(s,"📚 让我们成长的人  Help Us Learn",TEACH)
tb(s,0.4,0.85,9.2,0.36,"想学新东西 — 找他们! 他们让大家越来越聪明。",sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.18,9.2,0.24,"Want to learn — find them! They help us grow smart.",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
learn=[
    ("📚","老师","Teacher","👧👦 学生","📖 教知识 + 做人","🏫 学校",TEACH),
    ("📖","图书管理员","Librarian","📚 读者","🔎 找书 + 借书","🏛️ 图书馆",LIB),
]
for i,(em,cn,en,who,what,where,cl) in enumerate(learn):
    x=0.4+i*4.65
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.55),Inches(4.55),Inches(3.10))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    tb(s,x+0.05,1.65,4.45,0.85,em,sz=68,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.60,4.45,0.42,cn,sz=20,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.02,4.45,0.28,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    sep=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(x+0.40),Inches(3.40),Inches(3.75),Inches(0.02))
    sep.fill.solid(); sep.fill.fore_color.rgb=cl; sep.line.fill.background()
    tb(s,x+0.20,3.50,4.20,0.30,f"👥 帮谁: {who}",sz=11,b=True,c=DARK)
    tb(s,x+0.20,3.82,4.20,0.30,f"🎯 解决: {what}",sz=11,b=True,c=DARK)
    tb(s,x+0.20,4.14,4.20,0.30,f"📍 在哪: {where}",sz=11,b=True,c=DARK)
sentence_frame_bar(s,4.78,"___ 教我们 ___ — 让我们 ___ 。","___ teaches us ___ — makes us ___.",accent=TEACH)
n+=1; pn(s,n)
notes(s,"3 分钟 — 让我们成长的人:\n• 老师强调: 「你每天都见的两位 — 老师和图书管理员!」\n• 互动: 「你最喜欢老师教什么? 你最喜欢的书是什么?」\n• 抽 2-3 个学生分享 — 把他们的『谢谢老师』 / 『谢谢图书管理员』提前讲一讲")

# Deep-dive Group C: Daily Life (邮递员, 清洁工, 厨师)
s=ns(); bg(s,CREAM); hb(s,"✨ 让生活方便的人  Make Life Easy",CHEF)
tb(s,0.4,0.85,9.2,0.36,"每天默默工作 — 我们的日子才能轻松、干净、好吃!",sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.18,9.2,0.24,"Quiet daily heroes — they make our days clean and delicious!",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
daily=[
    ("✉️","邮递员","Mail carrier","✉️ 收信的人","📬 送信 + 送包裹","🛣️ 邮路",MAIL),
    ("🧹","清洁工","Cleaner","🌍 大家","🗑️ 扫地 + 清垃圾","🏙️ 街道",CLEAN),
    ("👨‍🍳","厨师","Chef","🍽️ 吃饭的人","🍱 做好吃的饭","🍳 厨房",CHEF),
]
for i,(em,cn,en,who,what,where,cl) in enumerate(daily):
    x=0.4+i*3.10
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.55),Inches(2.95),Inches(3.10))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    tb(s,x+0.05,1.65,2.85,0.85,em,sz=58,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.55,2.85,0.40,cn,sz=18,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.95,2.85,0.28,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    sep=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(x+0.30),Inches(3.30),Inches(2.35),Inches(0.02))
    sep.fill.solid(); sep.fill.fore_color.rgb=cl; sep.line.fill.background()
    tb(s,x+0.10,3.40,2.75,0.30,f"👥 帮谁: {who}",sz=10,b=True,c=DARK)
    tb(s,x+0.10,3.72,2.75,0.30,f"🎯 解决: {what}",sz=10,b=True,c=DARK)
    tb(s,x+0.10,4.04,2.75,0.30,f"📍 在哪: {where}",sz=10,b=True,c=DARK)
sentence_frame_bar(s,4.78,"___ 让我们的生活 ___ 。","___ makes our life ___.",accent=CHEF)
n+=1; pn(s,n)
notes(s,"3-4 分钟 — 默默工作的人:\n• 老师指出: 「这 3 位 — 你也许没注意到, 但他们每天都在帮你!」\n• 互动: 「你今天吃饭了吗? — 谢谢谁?」 → 厨师 (爸妈也算!)\n• 「你今天看到的街道干净吗? — 谢谢谁?」 → 清洁工\n• 重要: 「这些『看不见』的帮手 — 也很重要!」")

# ============================================================
# === SESSION 1 · PHASE 3: ACTIVE PRACTICE (20 min) ===
# ============================================================
s=phase_marker("🎯","ACTIVE PRACTICE","Try It Out!",20,PH_ACTIVE,"配对 + 情境判断 — 谁来帮?","Matching + situations — who helps?")
n+=1; pn(s,n)

# Matching game setup
s=ns(); bg(s,CREAM); hb(s,"🃏 配对游戏 · 8 个问题 → 8 位帮手  Match Game",PH_ACTIVE)
score_badge(s)
tb(s,0.4,1.15,9.2,0.36,"老师举一张「问题卡」 — 哪一队 先举手 说出 对应的帮手?",sz=13,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.50,9.2,0.24,"Teacher holds up a problem card — first team to name the helper wins!",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
# 8 problem pills (preview the cards)
problems=[
    ("🔥","房子着火",FIRE),
    ("🤒","小朋友发烧",DOC),
    ("👶","小宝宝走丢",POLICE),
    ("📚","想学算数",TEACH),
    ("📦","包裹要送来",MAIL),
    ("🍱","饿了想吃饭",CHEF),
    ("📚","想找一本书",LIB),
    ("🗑️","街上有垃圾",CLEAN),
]
for i,(em,cn,cl) in enumerate(problems):
    col=i%4; row=i//4
    x=0.4+col*2.32; y=1.95+row*1.40
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(2.22),Inches(1.25))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(2.5)
    tb(s,x+0.05,y+0.08,2.12,0.65,em,sz=36,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,y+0.80,2.12,0.40,cn,sz=12,b=True,c=cl,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.85,"___ 的时候, 我们 找 ___ !","When ___, we find ___!",accent=PH_ACTIVE)
n+=1; pn(s,n)
notes(s,"🎯 ACTIVE 1 · 7-8 分钟 — 配对游戏:\n• 准备 8 张问题卡 (老师课前打印)\n• 玩法:\n  - 1 分钟: 老师讲规则\n  - 5-6 分钟: 一张张出, 各队抢答\n  - 答对 +2 分\n• 老师举卡时戏剧化: 「啊! 房子着火了! — 找谁?」\n• 学生回答: 「消防员!」 全班齐喊\n• 答案 (对照 cl):\n  - 🔥 着火 → 消防员\n  - 🤒 发烧 → 医生\n  - 👶 走丢 → 警察\n  - 📚 学算数 → 老师\n  - 📦 送包裹 → 邮递员\n  - 🍱 饿了 → 厨师 (或家里大人!)\n  - 📚 找书 → 图书管理员\n  - 🗑️ 垃圾 → 清洁工")

# Situation cards practice
s=ns(); bg(s,CREAM); hb(s,"🎬 情境判断 · 你 会 找谁?  Pick the Right Helper",PH_ACTIVE)
tb(s,0.4,0.85,9.2,0.32,"读情境 → 想一想 → 全班一起说出帮手!",sz=13,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.15,9.2,0.22,"Read each scene → think → answer together!",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
situations=[
    ("1","🌙","半夜 — 闻到 烟味 + 火光!","→ ___","消防员 119/911",FIRE),
    ("2","🤕","在公园 — 摔了一跤 + 流血!","→ ___","医生 / 救护车",DOC),
    ("3","👀","在商场 — 看见 可疑的人!","→ ___","警察",POLICE),
    ("4","❓","作业 — 不会做这道题!","→ ___","老师",TEACH),
    ("5","🎁","收到 一个大包裹!","→ ___","邮递员",MAIL),
    ("6","🍝","中午 在学校食堂 吃饭","→ ___","厨师",CHEF),
]
for i,(num,em,scene,arrow,ans,cl) in enumerate(situations):
    col=i%3; row=i//3
    x=0.4+col*3.10; y=1.50+row*1.75
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(2.95),Inches(1.62))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(2.5)
    nb=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x+0.10),Inches(y+0.10),Inches(0.42),Inches(0.42))
    nb.fill.solid(); nb.fill.fore_color.rgb=cl; nb.line.fill.background()
    tb(s,x+0.10,y+0.14,0.42,0.34,num,sz=14,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,x+0.60,y+0.08,0.55,0.45,em,sz=24,a=PP_ALIGN.CENTER)
    tb(s,x+0.15,y+0.62,2.65,0.38,scene,sz=11,b=True,c=DARK)
    tb(s,x+0.15,y+1.00,2.65,0.30,arrow,sz=12,b=True,c=cl)
    tb(s,x+0.15,y+1.28,2.65,0.30,ans,sz=11,b=True,c=cl)
n+=1; pn(s,n)
notes(s,"🎯 ACTIVE 2 · 7-8 分钟 — 情境判断:\n• 6 个情境, 一题一题来\n• 老师先盖住答案 (一手挡住右下角)\n• 全班思考 → 齐答 → 老师揭晓\n• Tricky 题: 第 1 题 (半夜火) — 不只找消防员, 也要喊大人; 第 3 题 (可疑) — 不要自己去, 告诉警察 / 大人\n• 答错不扣分 — 鼓励试错\n• 关键 takeaway: 「真有事 — 不要自己处理, 找大人帮忙打电话!」")

# Tools quiz — match tools to helpers
s=ns(); bg(s,CREAM); hb(s,"🛠️ 工具 → 帮手  Tools → Helpers",PH_ACTIVE)
tb(s,0.4,0.85,9.2,0.32,"看到工具 — 猜猜是哪一位帮手用的?",sz=13,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.15,9.2,0.22,"See the tool — guess whose it is!",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
tools=[
    ("🩺","听诊器","医生",DOC),
    ("🚒","消防车","消防员",FIRE),
    ("📐","粉笔","老师",TEACH),
    ("🚓","警车","警察",POLICE),
    ("📬","邮包","邮递员",MAIL),
    ("🧹","扫把","清洁工",CLEAN),
    ("📚","图书卡","管理员",LIB),
    ("🍳","锅","厨师",CHEF),
]
for i,(em,tool,ans,cl) in enumerate(tools):
    col=i%4; row=i//4
    x=0.4+col*2.32; y=1.55+row*1.75
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(2.22),Inches(1.60))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(2.5)
    tb(s,x+0.05,y+0.08,2.12,0.70,em,sz=42,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,y+0.82,2.12,0.30,tool,sz=12,b=True,c=DARK,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,y+1.12,2.12,0.26,"→",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,y+1.30,2.12,0.30,ans,sz=12,b=True,c=cl,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"🎯 ACTIVE 3 · 4-5 分钟 — 工具配对:\n• 老师指一个工具 — 全班齐答 / 抢答\n• 中文教学: 听诊器、消防车、粉笔、警车、邮包、扫把、图书卡、锅\n• 关键: 「每位帮手都有自己的『超能力工具』 — 帮他们工作!」")

# ============================================================
# === SESSION 1 · PHASE 4: APPLY (10 min) ===
# ============================================================
s=phase_marker("🌱","APPLY","Use What You Learned",10,PH_APPLY,"我想感谢身边的小帮手","Thank a helper near you")
n+=1; pn(s,n)

# Apply — 我想感谢
s=ns(); bg(s,CREAM); hb(s,"❤️ 我想感谢身边的小帮手  Thank a Helper",HELP)
tb(s,0.4,0.85,9.2,0.36,"想一想 — 你身边 有哪一位 小帮手 让你感激?",sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.18,9.2,0.24,"Think — which helper near you do you appreciate?",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
prompts=[
    ("👨‍⚕️","医生 / 护士","Doctor / Nurse","你生病时 照顾你"),
    ("📚","老师","Teacher","教你 新东西"),
    ("👵","爷爷奶奶","Grandparents","做饭、讲故事"),
    ("🧑‍🚒","消防员 / 警察","Firefighter / Police","让你安全"),
    ("👨‍👩‍👧","爸爸妈妈","Parents","每天 照顾你"),
    ("🧑‍🤝‍🧑","好朋友","Friend","跟你 一起玩"),
]
for i,(em,cn,en,ex) in enumerate(prompts):
    col=i%3; row=i//3
    x=0.3+col*3.20; y=1.55+row*1.55
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(3.0),Inches(1.40))
    sh.fill.solid(); sh.fill.fore_color.rgb=WARM; sh.line.color.rgb=HELP; sh.line.width=Pt(2)
    tb(s,x+0.10,y+0.05,0.6,0.55,em,sz=24,a=PP_ALIGN.CENTER)
    tb(s,x+0.75,y+0.10,2.2,0.35,cn,sz=14,b=True,c=HELP)
    tb(s,x+0.75,y+0.42,2.2,0.30,en,sz=10,c=GRAY)
    tb(s,x+0.10,y+0.83,2.85,0.5,f"💭 {ex}",sz=12,c=DARK)
sentence_frame_bar(s,4.78,"我想感谢 ___, 因为 ___ 。","I want to thank ___, because ___.",accent=HELP)
n+=1; pn(s,n)
notes(s,"🌱 APPLY · 8-10 分钟:\n• 2 分钟 — 个人想: 「身边谁帮过我?」\n• 3 分钟 — 同桌交流 (Turn & Talk): 「我想感谢 ___, 因为 ___」\n• 3-5 分钟 — 让 3-4 个学生上台分享\n• 老师跟读 + 鼓掌\n• 关键: 「小帮手不只是穿制服的 — 爸妈、爷爷奶奶、朋友 — 也是!」\n• 不评好坏 — 鼓励每个想法")

# ============================================================
# === SESSION 1 · PHASE 5: SHARE & CLOSE (5 min) ===
# ============================================================
s=ns(); bg(s,CREAM); hb(s,"🎤 总结 + 下一步  Summary + Next!",PH_CLOSE)
score_badge(s)
tb(s,0.4,0.95,9.2,0.4,"🧭 今天早上 学了 什么?",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
left=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.5),Inches(1.45),Inches(4.4),Inches(1.85))
left.fill.solid(); left.fill.fore_color.rgb=HELP; left.line.fill.background()
tb(s,0.5,1.55,4.4,0.5,"❤️ 小帮手 = ?",sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,0.5,2.05,4.4,0.55,"帮 + 解 + 安 + 便!",sz=22,b=True,c=IDEA,a=PP_ALIGN.CENTER)
tb(s,0.5,2.60,4.4,0.30,"Help · Solve · Safe · Easy",sz=11,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,0.5,2.90,4.4,0.30,"4 个共同点 — 缺一不可",sz=11,c=IDEA,a=PP_ALIGN.CENTER)
right=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.10),Inches(1.45),Inches(4.4),Inches(1.85))
right.fill.solid(); right.fill.fore_color.rgb=NAVY; right.line.fill.background()
tb(s,5.10,1.55,4.4,0.5,"👋 我认识的 8 位",sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,5.10,2.05,4.4,0.55,"医 · 消 · 师 · 警",sz=22,b=True,c=IDEA,a=PP_ALIGN.CENTER)
tb(s,5.10,2.55,4.4,0.30,"邮 · 清 · 图 · 厨",sz=18,b=True,c=IDEA,a=PP_ALIGN.CENTER)
tb(s,5.10,2.95,4.4,0.30,"8 helpers — all important!",sz=10,c=WHITE,a=PP_ALIGN.CENTER)
trans=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.5),Inches(3.45),Inches(9.0),Inches(1.45))
trans.fill.solid(); trans.fill.fore_color.rgb=PH_ACTIVE; trans.line.color.rgb=IDEA; trans.line.width=Pt(3)
tb(s,0.5,3.55,9.0,0.45,"📖 下午 → 复习 + 我会认 + 我会写",sz=20,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,0.5,4.00,9.0,0.30,"Afternoon Session 2 → Review + Read + Write",sz=11,c=IDEA,a=PP_ALIGN.CENTER)
tb(s,0.5,4.35,9.0,0.40,"📝 我会认: 老师 · 学校 · 医院 · 消防员 · 厨师",sz=14,b=True,c=IDEA,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"🎤 SHARE & CLOSE · 5 分钟:\n• 1 分钟 — 全班齐喊: 「小帮手! 帮助 + 解决 + 安全 + 方便!」 (3 遍 + 手势)\n• 2 分钟 — 各队代表说一句: 「我们队最喜欢 ___ 帮手, 因为 ___」\n• 1 分钟 — 公布 Session 1 暂时积分\n• 1 分钟 — Tease 下午: 「下午 — 学 5 个新字 + 写 2 个 + Day 4 booklet!」")

# ============================================================
# 22. SESSION 2 DIVIDER
# ============================================================
s=div("Session 2  下午 1:00–1:50","📚 复习 + 我会认 + 我会写  Review · Read · Write",IDEA,"📖"); n+=1; pn(s,n)

# REVIEW — Session 1 recap
s=ns(); bg(s,CREAM); hb(s,"🔄 复习  Review · Session 1",NAVY)
tb(s,0.4,0.85,9.2,0.40,"还记得 早上 学的吗?  Remember what we learned?",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.25,9.2,0.30,"Quick recap before we read & write!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
recap=[
    ("❤️","小帮手 = ?","Helper = ?","帮 + 解 + 安 + 便",HELP),
    ("🛡️","让我们安全","Keep us safe","医生 · 消防员 · 警察",FIRE),
    ("📚","让我们成长","Help us learn","老师 · 图书管理员",TEACH),
    ("✨","让生活方便","Make life easy","邮递员 · 清洁工 · 厨师",CHEF),
]
for i,(em,cn,en,detail,c) in enumerate(recap):
    col=i%2; row=i//2
    x=0.4+col*4.65; y=1.65+row*1.45
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(4.55),Inches(1.30))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=c; sh.line.width=Pt(2)
    tb(s,x+0.10,y+0.12,0.7,0.65,em,sz=32,a=PP_ALIGN.CENTER)
    tb(s,x+0.85,y+0.10,3.6,0.40,cn,sz=15,b=True,c=c)
    tb(s,x+0.85,y+0.45,3.6,0.30,en,sz=10,c=GRAY)
    tb(s,x+0.15,y+0.85,4.30,0.40,detail,sz=11,b=True,c=DARK)
sentence_frame_bar(s,4.65,"今天 早上 我学了 ___ 。","This morning I learned about ___.")
n+=1; pn(s,n)
notes(s,"5 分钟 review:\n• 「早上学了哪 4 个共同点?」抢答\n• 「让我们安全的 3 位是?」/ 「让生活方便的 3 位是?」\n• 让 1-2 个学生说: 「我今天学了 ___」\n• 引出: 「现在 — 我们来学 5 个新字!」")

# ============================================================
# 我会认 — 5 vocabulary words: 老师 学校 医院 消防员 厨师
# ============================================================
read_words=[
    ("📚","老师","lǎo shī","Teacher",TEACH,
        "老师 在 学校 教我们。",
        "📷 老师 / 黑板 / 教室"),
    ("🏫","学校","xué xiào","School",NAVY,
        "我 每天 去 学校 上课。",
        "📷 学校 / 操场 / 教室"),
    ("🏥","医院","yī yuàn","Hospital",DOC,
        "生病了 — 去 医院 看医生。",
        "📷 医院 / 救护车 / 病房"),
    ("🚒","消防员","xiāo fáng yuán","Firefighter",FIRE,
        "消防员 救人 + 灭火 — 真勇敢!",
        "📷 消防员 / 消防车 / 头盔"),
    ("👨‍🍳","厨师","chú shī","Chef",CHEF,
        "厨师 做的 饭菜 真好吃。",
        "📷 厨师 / 高白帽 / 厨房"),
]
for em,cn,py,en,c,sent,img_label in read_words:
    s=ns(); bg(s,CREAM); hb(s,f"👀 我会认 · {cn}  I Can Read",c)
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.0),Inches(4.5),Inches(2.5))
    sh.fill.solid(); sh.fill.fore_color.rgb=WARM; sh.line.fill.background()
    tb(s,0.5,1.10,4.3,1.4,cn,sz=56 if len(cn)==3 else 70,b=True,c=c,a=PP_ALIGN.CENTER)
    tb(s,0.5,2.40,4.3,0.4,f"{py}  {en}",sz=18,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,0.5,2.85,4.3,0.4,"👉 跟我读!  Read after me!",sz=14,c=c,a=PP_ALIGN.CENTER)
    ib_box=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.3),Inches(1.0),Inches(4.4),Inches(2.5))
    ib_box.fill.solid(); ib_box.fill.fore_color.rgb=IMGBG; ib_box.line.fill.background()
    tb(s,5.3,2.05,4.4,0.4,img_label,sz=14,c=LGRAY,a=PP_ALIGN.CENTER)
    sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(3.8),Inches(9.2),Inches(1.2))
    sh2.fill.solid(); sh2.fill.fore_color.rgb=WHITE; sh2.line.color.rgb=c; sh2.line.width=Pt(2)
    tb(s,0.6,3.9,1.5,0.4,"例句",sz=16,b=True,c=c)
    tb(s,0.6,4.3,8.8,0.5,sent,sz=22,b=True,c=DARK)
    n+=1; pn(s,n)
    notes(s,f"3 分钟 — {cn}:\n• 老师指字, 全班齐读 3 遍 (慢→快→唱)\n• 看图: 「这是 ___, 在 哪里?」\n• 读例句, 学生跟读\n• 抽 1-2 个学生用「{cn}」造一个新句子\n• 写到黑板上 — 让学生在空中跟着写一遍")

# ============================================================
# 我会写 · 老师 (stroke order)
# ============================================================
s=ns(); bg(s,CREAM); hb(s,"✏️ 我会写 · 老师  I Can Write · Teacher",TEACH)
tb(s,0.4,0.85,9.2,0.40,"一起来写「老师」!",sz=22,b=True,c=TEACH,a=PP_ALIGN.CENTER)
tb(s,0.4,1.25,9.2,0.30,"Practice writing 老师 (Teacher)",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
tianzi(s,0.55,1.65,2.20,"老",TEACH,pinyin="lǎo (old)",char_sz=120)
tianzi(s,2.95,1.65,2.20,"师",TEACH,pinyin="shī (master)",char_sz=120)
panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.30),Inches(1.65),Inches(4.30),Inches(2.85))
panel.fill.solid(); panel.fill.fore_color.rgb=WHITE
panel.line.color.rgb=TEACH; panel.line.width=Pt(2.5)
head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.30),Inches(1.65),Inches(4.30),Inches(0.50))
head.fill.solid(); head.fill.fore_color.rgb=TEACH; head.line.fill.background()
tb(s,5.45,1.72,4.10,0.4,"✏️ 怎么写  How to write",sz=13,b=True,c=WHITE)
tb(s,5.45,2.30,4.10,0.40,"1️⃣「老」 — 6 笔",sz=14,b=True,c=DARK)
tb(s,5.45,2.65,4.10,0.30,"  上「耂」(老字头) + 下「匕」",sz=10,c=GRAY)
tb(s,5.45,3.05,4.10,0.40,"2️⃣「师」 — 6 笔",sz=14,b=True,c=DARK)
tb(s,5.45,3.40,4.10,0.30,"  左边「丿+一」, 右边像旗子!",sz=10,c=GRAY)
tb(s,5.45,3.85,4.10,0.40,"📝 在 田字格 里 写 3 遍",sz=12,b=True,c=TEACH)
tb(s,5.45,4.20,4.10,0.30,"Practice 3 times in grid paper",sz=9,c=GRAY)
sentence_frame_bar(s,4.65,
    "我会写「老师」! 我的 老师 叫 ___ 。",
    "I can write 老师! My teacher's name is ___.",accent=TEACH)
n+=1; pn(s,n)
notes(s,"6-7 分钟:\n• 演示笔顺, 学生跟写 (空中写)\n• 田字格练 3 遍\n• 记忆法:\n  - 「老」上面像帽子 (老人戴帽子)\n  - 「师」右边像旗子 — 老师举旗带头\n• 让学生说自己老师的名字 — 「我的老师叫 ___ 」\n• 完成 → 颁发「我会写」贴纸; 写得最好的队 +2 分")

# ============================================================
# 我会写 · 学校 (stroke order)
# ============================================================
s=ns(); bg(s,CREAM); hb(s,"✏️ 我会写 · 学校  I Can Write · School",NAVY)
tb(s,0.4,0.85,9.2,0.40,"一起来写「学校」!",sz=22,b=True,c=NAVY,a=PP_ALIGN.CENTER)
tb(s,0.4,1.25,9.2,0.30,"Practice writing 学校 (School)",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
tianzi(s,0.55,1.65,2.20,"学",NAVY,pinyin="xué (study)",char_sz=120)
tianzi(s,2.95,1.65,2.20,"校",NAVY,pinyin="xiào (school)",char_sz=120)
panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.30),Inches(1.65),Inches(4.30),Inches(2.85))
panel.fill.solid(); panel.fill.fore_color.rgb=WHITE
panel.line.color.rgb=NAVY; panel.line.width=Pt(2.5)
head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.30),Inches(1.65),Inches(4.30),Inches(0.50))
head.fill.solid(); head.fill.fore_color.rgb=NAVY; head.line.fill.background()
tb(s,5.45,1.72,4.10,0.4,"✏️ 怎么写  How to write",sz=13,b=True,c=WHITE)
tb(s,5.45,2.30,4.10,0.40,"1️⃣「学」 — 8 笔",sz=14,b=True,c=DARK)
tb(s,5.45,2.65,4.10,0.30,"  上「⺍+冖」, 下「子」 — 房子里教孩子!",sz=10,c=GRAY)
tb(s,5.45,3.05,4.10,0.40,"2️⃣「校」 — 10 笔",sz=14,b=True,c=DARK)
tb(s,5.45,3.40,4.10,0.30,"  左「木」+ 右「交」 — 木头建的地方",sz=10,c=GRAY)
tb(s,5.45,3.85,4.10,0.40,"📝 在 田字格 里 写 3 遍",sz=12,b=True,c=NAVY)
tb(s,5.45,4.20,4.10,0.30,"Practice 3 times in grid paper",sz=9,c=GRAY)
sentence_frame_bar(s,4.65,
    "我会写「学校」! 我的 学校 叫 ___ 。",
    "I can write 学校! My school is called ___.",accent=NAVY)
n+=1; pn(s,n)
notes(s,"6-7 分钟:\n• 演示笔顺, 学生跟写\n• 田字格练 3 遍\n• 记忆法:\n  - 「学」上面是『盖子』(房顶), 下面是『子』(小孩) — 小孩在房子里学习!\n  - 「校」左边是『木』 — 古时候学校用木头建的\n• 让学生说自己学校名 — 「我的学校叫 ___」\n• 完成 → 「我会写」贴纸; 写得最好 +2 分")

# ============================================================
# Day 4 Booklet activity
# ============================================================
s=ns(); bg(s,CREAM); hb(s,"📓 写 Day 4 Booklet  Fill Your Booklet",PURPLE)
tb(s,0.4,0.85,9.2,0.40,"现在 — 打开 你的 Day 4 booklet, 一起完成!",sz=18,b=True,c=PURPLE,a=PP_ALIGN.CENTER)
tb(s,0.4,1.25,9.2,0.28,"Now — open your Day 4 booklet and fill it in together!",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
# 4 booklet sections
sections=[
    ("①","🖼️","看图选择","Multiple Choice","4 题 — 看图猜 帮手",NAVY),
    ("②","❤️","我想帮助 ___","I Want to Help","选你想帮谁 + 画画",HELP),
    ("③","🔗","连一连","Match",  "5 词 ↔ 表情",CHEF),
    ("④","✏️","描一描 写一写","Trace & Write","老师 · 学校",TEACH),
]
for i,(num,em,cn,en,detail,cl) in enumerate(sections):
    x=0.4+i*2.32
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.70),Inches(2.22),Inches(2.95))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    tb(s,x+0.05,1.78,2.12,0.40,num,sz=20,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.20,2.12,0.75,em,sz=40,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.00,2.12,0.36,cn,sz=13,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.36,2.12,0.28,en,sz=9,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.10,3.70,2.02,0.90,detail,sz=10,c=DARK,a=PP_ALIGN.CENTER)
bar=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(4.80),Inches(9.2),Inches(0.65))
bar.fill.solid(); bar.fill.fore_color.rgb=PURPLE; bar.line.color.rgb=IDEA; bar.line.width=Pt(2)
tb(s,0.55,4.85,9.0,0.30,"⏱ 10-15 分钟 — 全做完 → +5 分! 写字最漂亮的队 → +3 分!",sz=13,b=True,c=IDEA,a=PP_ALIGN.CENTER)
tb(s,0.55,5.18,9.0,0.22,"Finish all → +5 pts! Best handwriting → +3 pts!",sz=9,c=WARM,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"📓 BOOKLET 活动 · 10-15 分钟:\n• 老师发 day4_booklet.docx (打印好的)\n• 学生自己写 — 老师巡场指导\n• 4 个 section 大约时间:\n  - ① 看图选择 — 3 分钟\n  - ② 我想帮助 + 画 — 5 分钟\n  - ③ 连一连 — 3 分钟\n  - ④ 描一描 — 5 分钟\n• 写完 → 老师收齐 / 让学生留着带回家\n• 鼓励: 「booklet 是 你今天的 souvenir!」")

# ============================================================
# 33. SESSION 3 DIVIDER — Projects
# ============================================================
s=div("Session 3  下午 2:00–2:50","🛠️ 项目课  2 个大项目!  ·  50 min",PH_ACTIVE,"❤️"); n+=1; pn(s,n)

# ============================================================
# === SESSION 3 · PHASE 1: HOOK (5 min) ===
# ============================================================
s=ns(); bg(s,CREAM); hb(s,"🤔 一个 大问题  A Big Question",LAB)
tb(s,0.4,0.85,9.2,0.40,"老师 来 问 一个 不容易 回答的问题……",sz=16,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.28,"Teacher has a question that's not easy to answer…",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
# Big question box
qbox=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.5),Inches(1.75),Inches(9.0),Inches(2.40))
qbox.fill.solid(); qbox.fill.fore_color.rgb=LAB; qbox.line.color.rgb=IDEA; qbox.line.width=Pt(3.5)
tb(s,0.5,1.95,9.0,0.55,"❓",sz=46,a=PP_ALIGN.CENTER)
tb(s,0.5,2.55,9.0,0.60,"哪一位 小帮手 — 最重要?",sz=30,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,0.5,3.20,9.0,0.40,"Which helper is the MOST important?",sz=14,b=True,c=IDEA,a=PP_ALIGN.CENTER)
tb(s,0.5,3.65,9.0,0.30,"如果只能 留 一位 — 你 留 谁?",sz=14,b=True,c=IDEA,a=PP_ALIGN.CENTER)
# Hint
hint=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(4.30),Inches(9.2),Inches(1.10))
hint.fill.solid(); hint.fill.fore_color.rgb=WARM; hint.line.color.rgb=LAB; hint.line.width=Pt(2)
tb(s,0.55,4.38,9.0,0.30,"🙋 想 1 分钟 — 你 心里 选了 谁?",sz=13,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.55,4.70,9.0,0.30,"💬 一会儿 你要 上台 跟同学 辩论 — 为什么 你 选 ta!",sz=12,b=True,c=LAB,a=PP_ALIGN.CENTER)
tb(s,0.55,5.05,9.0,0.25,"In 1 min — you'll debate on stage why you picked them!",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"🔥 HOOK · 5 分钟:\n• 老师戏剧化地问: 「如果世界上只能留 1 位小帮手 — 你留谁?」\n• 让学生先沉默思考 1 分钟\n• 不收答案 — 留悬念\n• 关键: 「这是个不容易的问题 — 等会儿 你要 真的 辩论!」\n• 引出: 「先看 — 项目怎么进行!」")

# ============================================================
# === SESSION 3 · PHASE 2: MINI-LESSON (10 min) ===
# ============================================================
# Mini 1 — 4-step community helper thinking
s=ns(); bg(s,CREAM); hb(s,"🔁 帮手 思考 4 步  Helper Thinking Loop",HELP)
score_badge(s)
tb(s,0.4,1.10,9.2,0.4,"看见 → 想 → 找 → 帮! 解决一个问题, 4 步走!",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.55,9.2,0.30,"See → Think → Find → Help! 4 steps to solve a problem.",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
steps=[
    ("👀","看","See","看见 什么 问题?",NAVY),
    ("💡","想","Think","需要 什么 帮手?",HELP),
    ("📞","找","Find","怎么 找到 ta?",PH_ACTIVE),
    ("❤️","帮","Help","ta 来 帮你 解决!",GREEN),
]
for i,(em,cn,en,q,cl) in enumerate(steps):
    x=0.4+i*2.30
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(2.05),Inches(2.10),Inches(2.55))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    nb=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x+0.78),Inches(2.15),Inches(0.55),Inches(0.55))
    nb.fill.solid(); nb.fill.fore_color.rgb=cl; nb.line.fill.background()
    tb(s,x+0.78,2.23,0.55,0.4,str(i+1),sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.80,2.0,0.6,em,sz=42,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.50,2.0,0.45,cn,sz=22,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.95,2.0,0.30,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,4.25,2.0,0.30,q,sz=11,c=DARK,a=PP_ALIGN.CENTER)
    if i<3:
        tb(s,x+2.10,3.20,0.30,0.4,"→",sz=22,b=True,c=PURPLE,a=PP_ALIGN.CENTER)
tb(s,0.4,4.78,9.2,0.30,"🌟 真有事 — 记得告诉大人! 大人会 拨电话 找帮手!",sz=14,b=True,c=PURPLE,a=PP_ALIGN.CENTER)
tb(s,0.4,5.08,9.2,0.25,"Real emergency — tell an adult! Adults call the helper for you.",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"📚 MINI 1 · 4 分钟:\n• 4 步走 — 配手势:\n  - 看 (指眼)\n  - 想 (指头)\n  - 找 (拨电话)\n  - 帮 (拥抱)\n• 强调第 3 步: 不要自己处理大事 — 找大人! 大人 dial 911/119/110!\n• 全班齐喊 + 手势 3 遍")

# Mini 2 — 2 projects overview
s=ns(); bg(s,CREAM); hb(s,"🎯 今天 的 2 个项目  2 Projects Today",HELP)
tb(s,0.4,0.85,9.2,0.4,"两个大项目 — 一前一后!",sz=16,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.25,9.2,0.28,"2 mini-projects — one after another!",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
projects=[
    ("P1","⚖️","最重要的工作","Most Important Job","15 min","各队选 1 位 → 辩论 → 揭晓 全都重要!",HELP),
    ("P2","🛡️","社区救援大挑战","Rescue Challenge","20 min","8 个情境 → 抢答 → 帮手 + 工具!",FIRE),
]
for i,(tag,em,cn,en,t,desc,cl) in enumerate(projects):
    x=0.4+i*4.65
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.75),Inches(4.55),Inches(3.40))
    sh.fill.solid(); sh.fill.fore_color.rgb=cl; sh.line.fill.background()
    tb(s,x+0.05,1.85,4.45,0.40,tag,sz=18,b=True,c=IDEA,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.30,4.45,0.95,em,sz=64,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.35,4.45,0.45,cn,sz=20,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.80,4.45,0.30,en,sz=11,c=IDEA,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,4.20,4.45,0.30,f"⏱  {t}",sz=13,b=True,c=IDEA,a=PP_ALIGN.CENTER)
    tb(s,x+0.20,4.55,4.20,0.50,desc,sz=11,c=WHITE,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"📚 MINI 2 · 2 分钟:\n• 介绍 2 个项目 — 流程: P1 → P2\n• 强调总时间 35 分钟 active + apply\n• P1 = 团队辩论 (思辨 — 但答案是 all important!)\n• P2 = 抢答 (能量 + 应用)\n• 完成 = 队 +5 分; 优胜 = 队 +10 分")

# ============================================================
# === SESSION 3 · PHASE 3: ACTIVE PRACTICE (15 min) — Project 1 ===
# ============================================================
# Project 1 setup
s=ns(); bg(s,CREAM); hb(s,"⚖️ 项目 1 · 最重要的工作  Project 1 · Most Important Job",HELP)
group_label(s)
score_badge(s)
tb(s,0.4,1.20,9.2,0.40,"⏱ 15 分钟 — 各队选 1 位 帮手, 辩论 为什么 ta 最重要!",sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
# 4 team boxes — 4 helpers to pick
teams=[
    ("🔴","红队","医生","🩺","生病时 救命",DOC),
    ("🔵","蓝队","消防员","🚒","火灾时 救人",FIRE),
    ("🟢","绿队","老师","📚","教大家 学知识",TEACH),
    ("🟡","黄队","厨师","👨‍🍳","让大家 吃饱",CHEF),
]
for i,(em,team,helper,hem,why,cl) in enumerate(teams):
    x=0.4+i*2.32
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.70),Inches(2.22),Inches(2.85))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    tb(s,x+0.05,1.82,2.12,0.36,f"{em} {team}",sz=13,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.22,2.12,0.80,hem,sz=50,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.08,2.12,0.40,helper,sz=16,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.48,2.12,0.30,"最重要 — 因为:",sz=10,c=DARK,a=PP_ALIGN.CENTER)
    tb(s,x+0.10,3.80,2.02,0.65,why,sz=11,b=True,c=DARK,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.70,"我们 队 选 ___ — 因为 没有 ta, 就 没有 ___ !","Our team picks ___ — because without them, no ___!",accent=HELP)
n+=1; pn(s,n)
notes(s,"⚖️ P1 SETUP · 3 分钟 — 项目 1 开始:\n• 老师分配: 4 队各 1 位帮手 (示例 — 也可以让队自选)\n• 3 分钟队内讨论: 「为什么 ta 最重要?」\n• 给每队 1 张大白纸 — 写 3 个原因 (中英都行)\n• 句型: 「我们队选 ___, 因为 没有 ta, 就没有 ___!」")

# Project 1 — debate prompts (sample arguments)
s=ns(); bg(s,CREAM); hb(s,"🎤 辩论 时间!  Debate Time!",HELP)
tb(s,0.4,0.85,9.2,0.36,"各队 上台 60 秒 — 用 3 个理由 说服 全班!",sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.18,9.2,0.24,"Each team — 60 sec to convince the class with 3 reasons!",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
# 3-step debate guide
guide=[
    ("1","💬","开场","Start","「我们 选 ___ ! 因为……」",HELP),
    ("2","📝","3 个 理由","3 reasons","「第一 ___ 第二 ___ 第三 ___」",NAVY),
    ("3","🌟","结尾","Finish","「所以 — ___ 最重要!」",GOLD),
]
for i,(num,em,cn,en,frame,cl) in enumerate(guide):
    x=0.4+i*3.10
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.55),Inches(2.95),Inches(2.55))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    nb=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x+1.25),Inches(1.65),Inches(0.50),Inches(0.50))
    nb.fill.solid(); nb.fill.fore_color.rgb=cl; nb.line.fill.background()
    tb(s,x+1.25,1.73,0.50,0.4,num,sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.25,2.85,0.65,em,sz=38,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.95,2.85,0.40,cn,sz=15,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.35,2.85,0.26,en,sz=9,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.15,3.65,2.65,0.40,frame,sz=10,b=True,c=DARK,a=PP_ALIGN.CENTER)
# Voting strip
vote=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(4.20),Inches(9.2),Inches(0.85))
vote.fill.solid(); vote.fill.fore_color.rgb=HELP; vote.line.color.rgb=IDEA; vote.line.width=Pt(2.5)
tb(s,0.55,4.28,9.0,0.30,"🗳️ 4 队 讲完 — 全班 投票: 「最厉害的辩手」 (除了自己队)",sz=13,b=True,c=IDEA,a=PP_ALIGN.CENTER)
tb(s,0.55,4.62,9.0,0.30,"After 4 teams pitch — vote for the best speaker (not your own team)!",sz=10,c=WARM,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"⚖️ P1 DEBATE · 6-7 分钟:\n• 每队 60 秒 (4 队 × 60 = 4 分钟)\n• 用 3 步框架 — 开场 + 3 理由 + 结尾\n• 老师举手计时 / 倒计时手势\n• 讲完投票 — 用举手 or 贴纸 (不能投自己队)\n• 「最厉害辩手队」 +3 分\n• 准备下一页 — 揭晓「真正的答案」!")

# Project 1 reveal — they're all important!
s=ns(); bg(s,CREAM); hb(s,"🌟 真正 的 答案  The Real Answer",GOLD)
tb(s,0.4,0.85,9.2,0.36,"听完 4 队 — 你 觉得 谁 最重要?",sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.18,9.2,0.24,"After all 4 teams — who's the most important?",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
# Big reveal banner
banner=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.5),Inches(1.65),Inches(9.0),Inches(2.50))
banner.fill.solid(); banner.fill.fore_color.rgb=GOLD; banner.line.color.rgb=HELP; banner.line.width=Pt(3.5)
tb(s,0.5,1.80,9.0,0.55,"🌟",sz=46,a=PP_ALIGN.CENTER)
tb(s,0.5,2.40,9.0,0.65,"他们 都 一样 重要!",sz=34,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,0.5,3.10,9.0,0.40,"They are ALL equally important!",sz=15,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.5,3.55,9.0,0.40,"社区 需要 每一位 — 缺 一 不可!",sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
# Why?
why=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(4.25),Inches(9.2),Inches(1.15))
why.fill.solid(); why.fill.fore_color.rgb=WARM; why.line.color.rgb=GOLD; why.line.width=Pt(2)
tb(s,0.55,4.32,9.0,0.30,"💡 为什么? 火灾要消防员; 生病要医生; 上学要老师; 吃饭要厨师……",sz=12,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.55,4.65,9.0,0.30,"如果 少 一位 — 社区 就 不完整!",sz=12,b=True,c=GOLD,a=PP_ALIGN.CENTER)
tb(s,0.55,4.98,9.0,0.30,"💬 一起 喊: 「我们 都需要 — 大家 都重要!」",sz=12,b=True,c=HELP,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"⚖️ P1 REVEAL · 4-5 分钟:\n• 老师戏剧化地说: 「等等 — 答案 不是 一位 — 答案是…… 全部都 重要!」\n• 念为什么 — 一个一个场景:\n  - 火灾 → 没有消防员 — 房子 烧了\n  - 生病 → 没有医生 — 没人 救你\n  - 上学 → 没有老师 — 不会 学新东西\n  - 吃饭 → 没有厨师 — 饿肚子\n• 关键 takeaway: 「社区 = 每一位 都不可缺少」\n• 全班齐喊: 「我们 都需要 — 大家 都重要!」\n• 每队都 +5 分 (所有人 都赢)\n• 引出 P2: 「既然 都重要 — 我们来 玩 一个 大救援!」")

# ============================================================
# === SESSION 3 · PHASE 4: APPLY (15 min) — Project 2 Rescue Challenge ===
# ============================================================
s=phase_marker("🌱","APPLY","Rescue Challenge!",15,FIRE,"项目 2 · 社区救援大挑战","P2 · Community Rescue Challenge")
n+=1; pn(s,n)

# Project 2 setup
s=ns(); bg(s,CREAM); hb(s,"🛡️ 项目 2 · 社区救援大挑战  Project 2 · Rescue Challenge",FIRE)
group_label(s)
score_badge(s)
tb(s,0.4,1.20,9.2,0.40,"⏱ 15 分钟 — 8 个 紧急情境, 抢答 谁 来 帮?",sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.55,9.2,0.26,"15 min — 8 emergencies, race to name the helper!",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
# Rules box
rules=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.5),Inches(1.95),Inches(9.0),Inches(2.85))
rules.fill.solid(); rules.fill.fore_color.rgb=WHITE; rules.line.color.rgb=FIRE; rules.line.width=Pt(3)
tb(s,0.5,2.05,9.0,0.40,"📋 游戏规则  Rules",sz=18,b=True,c=FIRE,a=PP_ALIGN.CENTER)
rules_list=[
    ("1","📺","老师 念 一个 情境 — 比如『房子着火』",FIRE),
    ("2","🙋","第一个 举手 的 队 — 派 一个 代表 回答",GOLD),
    ("3","✅","答对 帮手 = +2 分; 还能 说出 工具 = 再 +1 分; 工具 + 电话 = 再 +2 分!",HELP),
    ("4","🔁","8 个 情境 — 全部 答完 → 总分 最高 的队 = 冠军!",PURPLE),
]
for i,(num,em,text,cl) in enumerate(rules_list):
    y=2.55+i*0.50
    tb(s,0.65,y,0.40,0.40,num,sz=18,b=True,c=cl)
    tb(s,1.10,y+0.04,0.50,0.35,em,sz=20,a=PP_ALIGN.CENTER)
    tb(s,1.70,y+0.05,7.70,0.40,text,sz=12,b=True,c=DARK)
sentence_frame_bar(s,4.92,"___ 的时候, 找 ___ ! 用 ___ ! 打 ___ !","When ___, find ___! With ___! Call ___!",accent=FIRE)
n+=1; pn(s,n)
notes(s,"🛡️ P2 SETUP · 2-3 分钟:\n• 老师讲规则 — 强调 4 个赛点:\n  - 找对帮手 +2\n  - 工具 +1\n  - 电话 +2 (911/119/110/120)\n• 准备 8 张 emergency 卡 (PPT 下页有)\n• 各队 派 出 1 名代表 上台 — 轮流!")

# Project 2 scenarios — 8 emergency cards
s=ns(); bg(s,CREAM); hb(s,"🚨 8 个 紧急情境  8 Emergency Scenes",FIRE)
score_badge(s)
tb(s,0.4,1.10,9.2,0.30,"读情境 → 想 → 抢答 帮手 + 工具 + 电话!",sz=12,b=True,c=DARK,a=PP_ALIGN.CENTER)
scenarios=[
    ("1","🔥","房子 着火 + 浓烟!","119/911",FIRE),
    ("2","🤒","小宝宝 突然 高烧!","120/911",DOC),
    ("3","👤","公园里 — 走 丢 一个 小朋友!","110/911",POLICE),
    ("4","📦","门口 — 有 一个 重 包裹!","邮局",MAIL),
    ("5","🐕","小狗 受伤 + 流血!","兽医/120",DOC),
    ("6","🗑️","街上 一堆 垃圾 + 臭!","清洁公司",CLEAN),
    ("7","📚","作业 — 完全 不会!","学校/老师",TEACH),
    ("8","🍱","学校 — 中午 没饭吃!","学校 / 食堂",CHEF),
]
for i,(num,em,scene,phone,cl) in enumerate(scenarios):
    col=i%4; row=i//4
    x=0.4+col*2.32; y=1.50+row*1.85
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(2.22),Inches(1.70))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(2.5)
    nb=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x+0.08),Inches(y+0.08),Inches(0.40),Inches(0.40))
    nb.fill.solid(); nb.fill.fore_color.rgb=cl; nb.line.fill.background()
    tb(s,x+0.08,y+0.12,0.40,0.32,num,sz=14,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,x+0.55,y+0.05,1.55,0.50,em,sz=26,a=PP_ALIGN.CENTER)
    tb(s,x+0.10,y+0.62,2.05,0.50,scene,sz=11,b=True,c=DARK,a=PP_ALIGN.CENTER)
    tb(s,x+0.10,y+1.14,2.05,0.26,"📞 电话:",sz=8,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.10,y+1.38,2.05,0.30,phone,sz=11,b=True,c=cl,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"🛡️ P2 PLAY · 10-12 分钟 — 8 个情境:\n• 一个一个出 — 各队 抢答\n• 答案 (cl 已经对应):\n  - 1 着火 → 消防员 / 消防车 / 119(中) 911(美)\n  - 2 高烧 → 医生 / 救护车 / 120(中) 911(美)\n  - 3 走丢 → 警察 / 110(中) 911(美)\n  - 4 包裹 → 邮递员 / 邮包 / 邮局\n  - 5 狗 流血 → 兽医 (是医生家族) / 救护车 / 120\n  - 6 垃圾 → 清洁工 / 扫把 / 311(美)\n  - 7 作业 → 老师 / 粉笔 / 学校\n  - 8 没饭 → 厨师 / 锅 / 食堂\n• 每题 30-60 秒 — 不超时\n• 计分员 (1 个学生) 在黑板上记\n• 完成所有 8 题 → 揭晓 冠军!")

# Project 2 pledge
s=ns(); bg(s,CREAM); hb(s,"❤️ 我 也 是 小帮手  I'm a Helper Too",HELP)
tb(s,0.4,0.85,9.2,0.40,"想一想 — 你 怎么 帮 你的 社区?",sz=18,b=True,c=HELP,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.30,"Think — how can YOU help your community?",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# 6 example pledges
pledges=[
    ("🌱","帮 老师 浇花","Water teacher's plants"),
    ("🗑️","看到 垃圾 — 捡起来","Pick up trash"),
    ("📚","帮 同学 解题","Help classmates"),
    ("👵","扶 爷爷 过马路","Help elderly cross"),
    ("🐕","照顾 小动物","Care for animals"),
    ("👶","哄 弟弟妹妹","Comfort siblings"),
]
for i,(em,cn,en) in enumerate(pledges):
    col=i%3; row=i//3
    x=0.4+col*3.10; y=1.65+row*1.40
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(2.95),Inches(1.25))
    sh.fill.solid(); sh.fill.fore_color.rgb=WARM; sh.line.color.rgb=HELP; sh.line.width=Pt(2)
    tb(s,x+0.10,y+0.08,0.7,0.6,em,sz=28,a=PP_ALIGN.CENTER)
    tb(s,x+0.85,y+0.10,2.05,0.40,cn,sz=14,b=True,c=HELP)
    tb(s,x+0.85,y+0.50,2.05,0.30,en,sz=10,c=GRAY)
sentence_frame_bar(s,4.55,"我 想 帮助 ___, 我 可以 ___ !","I want to help ___, I can ___!",accent=HELP)
n+=1; pn(s,n)
notes(s,"❤️ PLEDGE · 3-4 分钟 — 最后 升华:\n• 关键句型 — 今天的 核心: 「我想帮助 ___, 我可以 ___」\n• 让 4-5 个学生 上台 用 句型 说\n• 老师 跟读 + 鼓掌\n• 关键 message: 「你不用 等长大 — 现在 就 可以 当 小帮手!」\n• 写到 booklet 上 — 带回家")

# ============================================================
# === SESSION 3 · PHASE 5: SHARE & CLOSE (5 min) ===
# ============================================================
s=ns(); bg(s,CREAM); hb(s,"🏆 颁奖 + Day 4 徽章!  Awards + Day 4 Badge!",PH_CLOSE)
score_badge(s)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.85),Inches(9.2),Inches(1.45))
sh.fill.solid(); sh.fill.fore_color.rgb=PH_CLOSE; sh.line.color.rgb=DARK; sh.line.width=Pt(2.5)
tb(s,0.4,1.00,9.2,0.45,"🏆 今天 的 冠军队  Day Champions",sz=20,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.50,9.2,0.45,"___ 队 — 全天 最高 分!",sz=24,b=True,c=HELP,a=PP_ALIGN.CENTER)
tb(s,0.4,2.00,9.2,0.30,"___ team — highest score across all 3 sessions!",sz=11,c=DARK,a=PP_ALIGN.CENTER)
sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.5),Inches(2.55),Inches(3),Inches(2.6))
sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=HEART; sh.line.width=Pt(6)
tf=tb(s,3.6,2.75,2.8,2.4,"DAY 4",sz=18,b=True,c=HEART,a=PP_ALIGN.CENTER)
ap(tf,"❤️🛡️",sz=42,a=PP_ALIGN.CENTER)
ap(tf,"社区小帮手",sz=14,b=True,c=HELP,a=PP_ALIGN.CENTER)
ap(tf,"✓ COMPLETED",sz=12,b=True,c=OK,a=PP_ALIGN.CENTER)
ap(tf,"🩺🚒📚👮",sz=18,a=PP_ALIGN.CENTER)
# Tease Day 5
tease=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(5.20),Inches(9.2),Inches(0.35))
tease.fill.solid(); tease.fill.fore_color.rgb=PURPLE; tease.line.fill.background()
tb(s,0.55,5.25,9.0,0.28,"🤖 明天 Day 5 — AI 与未来! 机器人 也是 小帮手 吗?",sz=12,b=True,c=IDEA,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"🎤 CLOSE · 5 分钟:\n• 1 分钟 — 公布全天总分, 冠军队上台\n• 1 分钟 — 各队代表说一句: 「今天 我最喜欢 ___」\n• 2 分钟 — 给每个学生发徽章 / 贴纸: 「你也是 小帮手!」\n• 1 分钟 — 全班齐喊: 「我想帮助 ___, 我可以 ___! 我也是 小帮手!」\n• 预告 Day 5: 「明天 — AI 与未来! 机器人 也能 当小帮手 吗?」")

# ============================================================
# BONUS — supplementary video / extension
# ============================================================
s=ns(); bg(s,CREAM); hb(s,"🎬 加餐 视频  Bonus Video",PH_CLOSE)
tb(s,0.4,0.85,9.2,0.40,"还有时间? 看 一 段 关于 社区小帮手 的 视频!",sz=16,b=True,c=HELP,a=PP_ALIGN.CENTER)
tb(s,0.4,1.25,9.2,0.28,"If time permits — one more video about community helpers!",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
vsh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.5),Inches(1.65),Inches(9.0),Inches(2.85))
vsh.fill.solid(); vsh.fill.fore_color.rgb=DARK; vsh.line.color.rgb=PH_CLOSE; vsh.line.width=Pt(3)
tb(s,0.5,2.05,9.0,1.40,"▶",sz=130,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,0.5,3.55,9.0,0.30,"🌟 社区小帮手 — 在 现实中 帮助 大家 的 故事",sz=14,b=True,c=IDEA,a=PP_ALIGN.CENTER)
tb(s,0.5,3.95,9.0,0.30,"老师 在 这里 插入 视频 / Teacher: insert video",sz=10,b=True,c=IDEA,a=PP_ALIGN.CENTER)
hint=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(4.65),Inches(9.2),Inches(0.85))
hint.fill.solid(); hint.fill.fore_color.rgb=WARM; hint.line.color.rgb=PH_CLOSE; hint.line.width=Pt(2)
tb(s,0.55,4.70,9.0,0.30,"🔍 推荐搜索: 'Community helpers for kids' / 'Mr. Rogers community helpers' / 'Sesame Street helpers'",sz=11,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.55,5.02,9.0,0.30,"💭 看完想一想: 你今天 学到 的 谁 — 最让你 感动?",sz=11,b=True,c=HELP,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"🎬 加餐视频 · 3-5 分钟 (可选):\n• 如果 P2 提前结束 — 用 这页\n• 推荐: Sesame Street Community Helpers (有中文版本)\n• 看完让 1-2 个学生分享: 「我今天最 感动 的 是 ___」\n• 一天 收尾 — 让 学生 带着 『我也可以 帮人』的 心情 回家")

# ============================================================
out=os.path.join(os.path.dirname(__file__),"day4_helpers.pptx")
prs.save(out); print(f"Saved {out}  ({n} slides)")
