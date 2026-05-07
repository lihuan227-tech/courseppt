#!/usr/bin/env python3
"""
我的职业梦想 — Day 3: 小小企业家 (Little Entrepreneurs)
3 sessions × 50 min, Workshop Model:
  🔥 Hook (5) → 📚 Mini-Lesson (10) → 🎯 Active Practice (20) → 🌱 Apply (10) → 🎤 Share & Close (5)
4 famous entrepreneurs: 乔布斯 · 沃尔特·迪士尼 · 马云 · 米尔顿·赫尔希
Session 3 project: 📦 产品改造挑战 — 设计新书包!
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
BIZ    = RGBColor(0xD4,0x8E,0x1F)   # primary entrepreneur gold/amber
MONEY  = RGBColor(0x2E,0x7D,0x32)   # money green
IDEA   = RGBColor(0xF5,0xC2,0x42)   # idea spark yellow
NAVY   = RGBColor(0x1E,0x3A,0x5F)
LAB    = RGBColor(0xE5,0x3E,0x3E)   # alarm/problem red
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

# Phase colors (consistent with Day 2)
PH_HOOK   = LAB
PH_MINI   = BIZ
PH_ACTIVE = RUST
PH_APPLY  = GREEN
PH_CLOSE  = IDEA

# Per-entrepreneur colors
JOBS    = RGBColor(0x33,0x33,0x33)  # Apple dark gray
DISNEY  = RGBColor(0x6B,0x46,0xC1)  # Disney magic purple
MAYUN   = RGBColor(0xFF,0x6A,0x00)  # Alibaba orange
HERSHEY = RGBColor(0x6B,0x3A,0x1A)  # chocolate brown

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
def hb(s,txt,c=BIZ,t=0.15):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(t),Inches(9.4),Inches(0.55))
    sh.fill.solid(); sh.fill.fore_color.rgb=c; sh.line.fill.background()
    tb(s,0.4,t+0.03,9.2,0.5,txt,sz=20,b=True,c=WHITE)
def pn(s,n): tb(s,9.0,5.25,0.8,0.3,str(n),sz=10,c=GRAY,a=PP_ALIGN.RIGHT)
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

def tpr_strip(s,t,cue_cn,cue_en):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(t),Inches(9.4),Inches(0.55))
    sh.fill.solid(); sh.fill.fore_color.rgb=PH_HOOK; sh.line.fill.background()
    tb(s,0.5,t+0.05,9.0,0.3,f"🙌 TPR · {cue_cn}",sz=14,b=True,c=WHITE)
    tb(s,0.5,t+0.30,9.0,0.25,cue_en,sz=10,c=IDEA)

# STORY slide — watch video + answer 3 questions
def story_slide(emoji,cn,en,years,country,color,panels,fun_cn,fun_en,video_cn="",video_en=""):
    s=ns(); bg(s,CREAM); hb(s,f"{emoji} 故事 · {cn}  ·  Story of {en}",color)
    score_badge(s)
    idstrip=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.85),Inches(9.2),Inches(0.50))
    idstrip.fill.solid(); idstrip.fill.fore_color.rgb=color; idstrip.line.fill.background()
    tb(s,0.55,0.90,0.7,0.40,emoji,sz=22,a=PP_ALIGN.CENTER)
    tb(s,1.30,0.90,3.5,0.35,cn,sz=17,b=True,c=WHITE)
    tb(s,5.00,0.92,2.4,0.22,f"📅 {years}",sz=10,c=IDEA)
    tb(s,5.00,1.13,2.4,0.20,en,sz=9,c=IDEA)
    tb(s,7.45,0.92,2.0,0.22,f"📍 {country}",sz=10,c=IDEA)
    vr=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.40),Inches(9.2),Inches(0.40))
    vr.fill.solid(); vr.fill.fore_color.rgb=PURPLE; vr.line.color.rgb=IDEA; vr.line.width=Pt(1.5)
    tb(s,0.55,1.43,2.5,0.35,"🎥 先看视频!",sz=13,b=True,c=IDEA)
    tb(s,3.10,1.43,6.50,0.20,video_cn,sz=11,b=True,c=WHITE)
    tb(s,3.10,1.62,6.50,0.18,video_en,sz=8,c=WARM)
    panel_y=1.90
    panel_h=2.45
    for i,(p_em,p_cn,p_en,b_cn,b_en,pause_cn) in enumerate(panels):
        x=0.4+i*3.10
        sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(panel_y),Inches(2.95),Inches(panel_h))
        sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=color; sh.line.width=Pt(2.5)
        nb=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x+0.10),Inches(panel_y+0.07),Inches(0.42),Inches(0.42))
        nb.fill.solid(); nb.fill.fore_color.rgb=color; nb.line.fill.background()
        tb(s,x+0.10,panel_y+0.13,0.42,0.32,str(i+1),sz=14,b=True,c=WHITE,a=PP_ALIGN.CENTER)
        tb(s,x+0.55,panel_y+0.07,2.4,0.50,p_em,sz=28,a=PP_ALIGN.CENTER)
        tb(s,x+0.10,panel_y+0.62,2.85,0.32,p_cn,sz=14,b=True,c=color,a=PP_ALIGN.CENTER)
        tb(s,x+0.10,panel_y+0.92,2.85,0.22,p_en,sz=9,c=GRAY,a=PP_ALIGN.CENTER)
        tb(s,x+0.15,panel_y+1.20,2.75,0.45,b_cn,sz=11,c=DARK)
        tb(s,x+0.15,panel_y+1.62,2.75,0.30,b_en,sz=8,c=GRAY)
        ps_y=panel_y+panel_h-0.43
        ps=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x+0.10),Inches(ps_y),Inches(2.75),Inches(0.38))
        ps.fill.solid(); ps.fill.fore_color.rgb=IDEA; ps.line.fill.background()
        tb(s,x+0.18,ps_y+0.04,2.65,0.30,f"👉 {pause_cn}",sz=10,b=True,c=DARK)
    ff_y=panel_y+panel_h+0.05
    ff=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(ff_y),Inches(9.2),Inches(0.45))
    ff.fill.solid(); ff.fill.fore_color.rgb=WARM; ff.line.color.rgb=color; ff.line.width=Pt(1.5)
    tb(s,0.55,ff_y+0.04,0.6,0.40,"⭐",sz=16,a=PP_ALIGN.CENTER)
    tb(s,1.10,ff_y+0.04,8.4,0.25,fun_cn,sz=11,b=True,c=color)
    tb(s,1.10,ff_y+0.27,8.4,0.18,fun_en,sz=8,c=GRAY)
    return s

# Q&A slide — 2 comprehension Qs + personal connection + TPR + sentence frame
def qa_slide(emoji,cn,en,color,questions,personal_cn,personal_en,tpr_cn,tpr_en,sentence_cn,sentence_en):
    s=ns(); bg(s,CREAM); hb(s,f"{emoji} 互动问答 · {cn}  ·  Q&A",color)
    group_label(s)
    score_badge(s)
    for i,(q_em,q_cn,q_en,hint_cn,hint_en) in enumerate(questions):
        y=1.20+i*0.95
        sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(y),Inches(9.2),Inches(0.85))
        sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=color; sh.line.width=Pt(2)
        tb(s,0.55,y+0.18,0.6,0.5,q_em,sz=24,a=PP_ALIGN.CENTER)
        tb(s,1.30,y+0.10,5.3,0.35,q_cn,sz=14,b=True,c=DARK)
        tb(s,1.30,y+0.45,5.3,0.30,q_en,sz=10,c=GRAY)
        pill=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(6.70),Inches(y+0.15),Inches(2.80),Inches(0.55))
        pill.fill.solid(); pill.fill.fore_color.rgb=IDEA; pill.line.fill.background()
        tb(s,6.75,y+0.18,2.70,0.28,f"💡 {hint_cn}",sz=11,b=True,c=DARK)
        tb(s,6.75,y+0.43,2.70,0.22,hint_en,sz=9,c=DARK)
    pc=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(3.20),Inches(9.4),Inches(0.55))
    pc.fill.solid(); pc.fill.fore_color.rgb=PURPLE; pc.line.fill.background()
    tb(s,0.5,3.25,9.0,0.3,f"💭 想一想 · {personal_cn}",sz=14,b=True,c=WHITE)
    tb(s,0.5,3.50,9.0,0.25,personal_en,sz=10,c=IDEA)
    tpr_strip(s,3.85,tpr_cn,tpr_en)
    sentence_frame_bar(s,4.50,sentence_cn,sentence_en,accent=color)
    return s

n=0

# ============================================================
# 1. COVER
# ============================================================
s=ns(); bg(s,CREAM)
tb(s,1,0.40,8,0.55,"我的职业梦想 · My Dream Career",sz=22,b=True,c=BIZ,a=PP_ALIGN.CENTER)
tb(s,1,0.95,8,0.4,"Day 3 · 小小企业家  Little Entrepreneurs",sz=16,b=True,c=DARK,a=PP_ALIGN.CENTER)
sh1=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(2.0),Inches(1.55),Inches(2.8),Inches(2.8))
sh1.fill.solid(); sh1.fill.fore_color.rgb=BIZ; sh1.line.color.rgb=IDEA; sh1.line.width=Pt(5)
tf1=tb(s,2.0,1.85,2.8,0.4,"金点子",sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
ap(tf1,"💡",sz=70,a=PP_ALIGN.CENTER)
ap(tf1,"发现需要!",sz=14,b=True,c=IDEA,a=PP_ALIGN.CENTER)
sh2=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(5.2),Inches(1.55),Inches(2.8),Inches(2.8))
sh2.fill.solid(); sh2.fill.fore_color.rgb=MONEY; sh2.line.color.rgb=IDEA; sh2.line.width=Pt(5)
tf2=tb(s,5.2,1.85,2.8,0.4,"产品",sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
ap(tf2,"🛍️",sz=70,a=PP_ALIGN.CENTER)
ap(tf2,"帮助别人!",sz=14,b=True,c=IDEA,a=PP_ALIGN.CENTER)
tb(s,1,4.65,8,0.45,"💰 我有一个 金点子!  I have a great idea!",sz=16,b=True,c=BIZ,a=PP_ALIGN.CENTER)
tb(s,1,5.15,8,0.3,"3 sessions × 50 min  ·  Hook → Mini-Lesson → Practice → Apply → Share",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"30 秒 hook:\n• 「今天你们都是 小老板!」\n• 4 队分好: 🔴 / 🔵 / 🟢 / 🟡\n• 准备道具: 一些「钱」 (假钱/贴纸/积木) 给每队当 starting capital")

# ============================================================
# 2. ROADMAP
# ============================================================
s=ns(); bg(s,CREAM); hb(s,"🗺️ 今天 3 节课  Today's 3 Sessions",NAVY)
tb(s,0.4,0.85,9.2,0.35,"每节课都有 5 步: 🔥 → 📚 → 🎯 → 🌱 → 🎤",sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.20,9.2,0.28,"Each session: Hook → Mini-Lesson → Practice → Apply → Share",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
sessions=[
    ("S1","🌟","故事课","Stories","上午","认识 4 位 famous 企业家",BIZ),
    ("S2","📖","词汇课","Words","下午","6 个商业词 + 写「产品·买·卖」",IDEA),
    ("S3","🛠️","项目课","Project","下午","设计 新书包 + 卖给同学!",PH_ACTIVE),
]
for i,(tag,em,cn,en,when,desc,cl) in enumerate(sessions):
    x=0.4+i*3.10
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.65),Inches(2.95),Inches(2.85))
    sh.fill.solid(); sh.fill.fore_color.rgb=cl; sh.line.fill.background()
    tb(s,x+0.05,1.80,2.85,0.45,tag,sz=18,b=True,c=IDEA,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.25,2.85,0.85,em,sz=60,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.20,2.85,0.45,cn,sz=22,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.65,2.85,0.30,en,sz=11,c=IDEA,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,4.00,2.85,0.45,desc,sz=11,c=WHITE,a=PP_ALIGN.CENTER)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(4.65),Inches(9.2),Inches(0.55))
sh.fill.solid(); sh.fill.fore_color.rgb=PURPLE; sh.line.fill.background()
tb(s,0.5,4.70,9.0,0.30,"🏆 4 队全天积分 — 最高分队拿 冠军企业家 徽章!",sz=14,b=True,c=WHITE)
tb(s,0.5,4.98,9.0,0.20,"4 teams · points all day · winners get champion entrepreneur badges",sz=10,c=IDEA)
n+=1; pn(s,n)
notes(s,"1 分钟:\n• 介绍 3 节课的 flow\n• 强调 4 队全天积分\n• 提示: S3 有 mini-fair, 学生设计的产品互相买卖!")

# ============================================================
# 3. SESSION 1 DIVIDER
# ============================================================
s=div("Session 1  上午 11:00–11:50","🌟 故事课  Entrepreneur Stories  ·  50 min",BIZ,"💡"); n+=1; pn(s,n)

# ============================================================
# === SESSION 1 · PHASE 1: HOOK (5 min) ===
# ============================================================
phase_marker("🔥","Hook","A Real Problem!",5,PH_HOOK,
             "大家饿了 — 谁来开 餐厅?","Everyone's hungry — who opens a restaurant?")
n+=1; pn(s:=prs.slides[-1],n)

# 4. Hook — 大家饿了 怎么办?
s=ns(); bg(s,CREAM); hb(s,"🍔 想象一下…  Imagine…",NAVY)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.5),Inches(0.95),Inches(3.0),Inches(2.20))
sh.fill.solid(); sh.fill.fore_color.rgb=BIZ; sh.line.color.rgb=NAVY; sh.line.width=Pt(3)
tb(s,0.5,1.10,3.0,1.3,"🍔",sz=90,a=PP_ALIGN.CENTER)
tb(s,0.5,2.50,3.0,0.4,"大家饿了!",sz=14,b=True,c=IDEA,a=PP_ALIGN.CENTER)
tb(s,0.5,2.85,3.0,0.25,"Everyone's hungry!",sz=10,c=IDEA,a=PP_ALIGN.CENTER)
panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(3.70),Inches(0.95),Inches(6.0),Inches(2.20))
panel.fill.solid(); panel.fill.fore_color.rgb=WHITE; panel.line.color.rgb=NAVY; panel.line.width=Pt(2.5)
tb(s,3.85,1.05,5.7,0.45,"💡 想一想:",sz=15,b=True,c=NAVY)
tb(s,3.85,1.45,5.7,0.40,"大家都饿了, 还没有 餐厅 — 怎么办?",sz=15,b=True,c=DARK)
tb(s,3.85,1.85,5.7,0.30,"Everyone's hungry, no restaurants — what to do?",sz=10,c=GRAY)
ln=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(3.85),Inches(2.20),Inches(5.7),Inches(0.02))
ln.fill.solid(); ln.fill.fore_color.rgb=LGRAY; ln.line.fill.background()
tb(s,3.85,2.28,5.7,0.40,"❓ 你能想出 什么 办法?",sz=15,b=True,c=LAB)
tb(s,3.85,2.65,5.7,0.30,"What ideas can YOU come up with?",sz=10,c=GRAY)
q2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(3.30),Inches(9.3),Inches(0.95))
q2.fill.solid(); q2.fill.fore_color.rgb=LAB; q2.line.fill.background()
tb(s,0.5,3.40,9.1,0.45,"❓ 谁 看到 「需要」, 谁 就 能 帮 — 这就是 企业家!",sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,0.5,3.85,9.1,0.30,"Whoever sees a 'need' can help — that's an entrepreneur!",sz=11,c=IDEA,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.45,"我看到大家 ___, 我可以 ___ !","I see people ___, I can ___!",accent=IDEA)
n+=1; pn(s,n)
notes(s,"🔥 HOOK · 5 分钟:\n• 表演: 老师装很饿 「啊! 我饿了! 哪里有吃的?」 — 戏剧效果\n• 问 3 个学生: 「大家饿了, 你怎么办?」\n  - 引导答案: 开餐厅, 卖零食, 送外卖, 做面包...\n• 收集想法 → 「你们已经在想 entrepreneur 的事了!」\n• 转: 「企业家 = 看到需要 → 想办法 → 做出来 帮人!」")

# ============================================================
# === SESSION 1 · PHASE 2: MINI-LESSON (10 min) ===
# ============================================================
phase_marker("📚","Mini-Lesson","What's an Entrepreneur?",10,PH_MINI,
             "企业家 = 发现需要 → 创造产品 → 帮助别人","See need → make product → help people")
n+=1; pn(s:=prs.slides[-1],n)

# 5. Concept — 3-step entrepreneur
s=ns(); bg(s,CREAM); hb(s,"💡 企业家是谁?  Who's an Entrepreneur?",BIZ)
tb(s,0.4,0.85,9.2,0.35,"3 个步骤 — 你也能做!",sz=15,b=True,c=DARK,a=PP_ALIGN.CENTER)
steps=[
    ("👀","发现 需要","Find a Need","看到大家 还没解决 的 问题",NAVY),
    ("💡","创造 产品","Make a Product","想办法 — 做一个 东西/服务",BIZ),
    ("🤝","帮助 别人","Help People","让别人 喜欢用 → 你成功!",MONEY),
]
for i,(em,cn,en,detail,cl) in enumerate(steps):
    x=0.4+i*3.10
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.30),Inches(2.95),Inches(3.0))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    nb=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x+1.18),Inches(1.42),Inches(0.6),Inches(0.6))
    nb.fill.solid(); nb.fill.fore_color.rgb=cl; nb.line.fill.background()
    tb(s,x+1.18,1.50,0.6,0.45,str(i+1),sz=22,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.10,2.85,0.85,em,sz=58,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.05,2.85,0.50,cn,sz=22,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.55,2.85,0.30,en,sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.10,3.90,2.75,0.35,detail,sz=11,c=DARK,a=PP_ALIGN.CENTER)
hint=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(4.45),Inches(9.2),Inches(0.55))
hint.fill.solid(); hint.fill.fore_color.rgb=WARM; hint.line.color.rgb=BIZ; hint.line.width=Pt(1.5)
tb(s,0.55,4.50,9.0,0.30,"💰 帮的人越多 → 越成功! Help more → succeed more!",sz=14,b=True,c=BIZ,a=PP_ALIGN.CENTER)
tb(s,0.55,4.78,9.0,0.20,"(钱不是目标, 帮人才是 — 帮得多, 钱自然来)",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"📚 MINI 1 · 4 分钟:\n• 念 3 步, 全班跟着做手势 (👀 看, 💡 想, 🤝 握手)\n• 重点: 「钱不是目标 — 帮人 是目标」\n• 例子: 爱迪生 帮人看见黑夜, 莱特兄弟 帮人飞 — 也是 entrepreneur!\n• 全班齐声: 「发现 — 创造 — 帮人!」 (×3)")

# 6. 4 examples — problem → biz idea
s=ns(); bg(s,CREAM); hb(s,"🎯 从小问题 开始!  Start from Small Problems!",BIZ)
tb(s,0.4,0.85,9.2,0.35,"4 个例子 — 都是 你身边 的 问题!",sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
examples=[
    ("🍔","大家饿了 怎么办?","开 餐厅 / 卖 食物","Restaurant / sell food",LAB),
    ("🧸","没有好玩的玩具?","设计 新玩具","Design new toys",PURPLE),
    ("📦","买东西不方便?","做外卖 / 网店 / App","Delivery / online shop / app",MAYUN),
    ("🎒","学习用品不好用?","设计 新书包 / 文具","New backpack / supplies",GREEN),
]
for i,(em,p_cn,sol_cn,sol_en,cl) in enumerate(examples):
    col=i%2; row=i//2
    x=0.4+col*4.65; y=1.30+row*1.55
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(4.55),Inches(1.40))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(2.5)
    tb(s,x+0.10,y+0.10,1.0,1.0,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+1.20,y+0.10,3.30,0.40,p_cn,sz=14,b=True,c=LAB)
    tb(s,x+1.20,y+0.50,0.4,0.30,"→",sz=18,b=True,c=cl)
    tb(s,x+1.20,y+0.78,3.30,0.35,sol_cn,sz=13,b=True,c=cl)
    tb(s,x+1.20,y+1.10,3.30,0.25,sol_en,sz=9,c=GRAY)
sentence_frame_bar(s,4.45,"我看到 ___ 问题, 我可以 ___ !","I see ___ problem, I can ___!",accent=BIZ)
n+=1; pn(s,n)
notes(s,"📚 MINI 2 · 6 分钟:\n• 念每个例子 — 全班跟着说: 「我看到 ___ , 我可以 ___」\n• Bonus: 让 1-2 个学生想新例子 (打哈欠 → 卖咖啡? 吵架 → 卖耳塞?)\n• 强调: 任何小问题都可能是 一个 商业机会")

# ============================================================
# === SESSION 1 · PHASE 3: ACTIVE PRACTICE (20 min) ===
# 4 entrepreneurs × 5 min each
# ============================================================
phase_marker("🎯","Active Practice","Meet 4 Entrepreneurs!",20,PH_ACTIVE,
             "故事 → 问答 → 想想 → 一起做!","Story → Q&A → Think → ACT!")
n+=1; pn(s:=prs.slides[-1],n)

# --- Steve Jobs ---
s=story_slide("🍎","乔布斯","Steve Jobs","1955–2011","🇺🇸 美国",JOBS,
    [("❓","他看到什么需要?","What need did he see?",
       "电脑太大太贵, 普通人不会用 — 他想让每个人都用上.",
       "Computers were huge & hard. He wanted them for everyone.",
       "你家有几个苹果产品?"),
     ("❓","他做了什么?","What did he do?",
       "在车库里创办苹果公司 — 设计了漂亮简单的电脑 + iPhone.",
       "Started Apple in a garage. Beautiful simple computers + iPhone.",
       "你猜他一开始在哪里工作?"),
     ("❓","结果是什么?","What was the result?",
       "几十亿人用 iPhone! 改变了世界怎么 听音乐/拍照/上网.",
       "Billions use iPhones today!",
       "iPhone 之前 — 我们怎么上网?")],
    "🚪 他在朋友家的车库里创办苹果 — 后来公司价值比国家还大!",
    "Started in a friend's garage — Apple worth more than most countries!",
    video_cn="搜: 'Who Was Steve Jobs?' 中文动画  ·  绘本: 'Who Was Steve Jobs?' 系列",
    video_en="Picture book: 'Who Was Steve Jobs?' / Search 'Steve Jobs kids documentary'")
n+=1; pn(s,n)
notes(s,"5 分钟:\n🎥 1-2 分钟视频 — 推荐: BiographyKids YouTube 'Steve Jobs'\n📋 3 分钟答 3 题:\n• Q1: 强调 — 看到「普通人需要简单的电脑」\n• Q2: 「车库」创业故事 — 让学生 picture\n• Q3: 让学生说自家几个苹果产品 (轮流上台)")

s=qa_slide("🍎","乔布斯","Steve Jobs",JOBS,
    [("🏠","乔布斯一开始在哪里工作?","Where did Jobs start?",
      "车库!","A garage!"),
     ("📱","他设计的最有名产品?","Most famous product?",
      "iPhone!","iPhone!")],
    "你家有苹果产品吗? 哪些?","Do you have Apple products? Which ones?",
    "🤲 假装滑动 iPhone → 「click! 滑!」 (×3)","Pretend to swipe an iPhone (×3)",
    "我想像乔布斯一样 — 设计 ___ !","Like Jobs — I want to design ___!")
n+=1; pn(s,n)
notes(s,"2 分钟 Q&A:\n• Q1, Q2 抢答 +2\n• 让 3 个学生说自家苹果产品 (iPhone, iPad, Mac, AirPods, Watch...)\n• TPR 滑 iPhone (3 次)\n• 句型: 让学生说想设计什么 (轮流上台)")

# --- Walt Disney ---
s=story_slide("🏰","沃尔特·迪士尼","Walt Disney","1901–1966","🇺🇸 美国",DISNEY,
    [("❓","他看到什么需要?","What need did he see?",
       "没有给小朋友的电影 / 乐园 — 他想给全家一起的快乐.",
       "No fun movies/parks for kids & families — he wanted joy for all.",
       "你看过哪些迪士尼电影?"),
     ("❓","他做了什么?","What did he do?",
       "画了米老鼠! 拍了动画电影! 建了迪士尼乐园!",
       "Drew Mickey! Made cartoons! Built Disneyland!",
       "你最喜欢哪个迪士尼角色?"),
     ("❓","结果是什么?","What was the result?",
       "全世界几亿人去迪士尼乐园, 看迪士尼电影 — 一直快乐到现在!",
       "Hundreds of millions visit Disney parks today.",
       "你想去哪个迪士尼乐园?")],
    "🐭 米老鼠是他在火车上画出来的 — 一只老鼠改变世界!",
    "He drew Mickey on a train — one mouse changed the world!",
    video_cn="搜: 'Walt Disney for kids' / 米老鼠 Steamboat Willie 短片  ·  绘本: 「Who Was Walt Disney?」",
    video_en="Picture book: 'Who Was Walt Disney?' / Search 'Walt Disney biography kids'")
n+=1; pn(s,n)
notes(s,"5 分钟:\n🎥 1-2 分钟 — 推荐: Steamboat Willie 1928 米老鼠首秀, 或 'Who Was Walt Disney?' 系列\n📋 3 分钟答:\n• Q1: 强调 — 没有给孩子的电影/乐园\n• Q2: 米老鼠 + 动画 + Disneyland — 三个大事\n• Q3: 收集 3-5 个最喜欢的迪士尼角色")

s=qa_slide("🏰","沃尔特·迪士尼","Walt Disney",DISNEY,
    [("🎨","他在哪里画米老鼠?","Where did he draw Mickey?",
      "火车上!","On a train!"),
     ("🎢","他建了什么让全家快乐?","What did he build for families?",
      "迪士尼乐园","Disneyland!")],
    "你最喜欢哪个迪士尼角色 / 电影?","Favorite Disney character or movie?",
    "🐭 戴米老鼠耳朵 (用手) → 喊「Disney!」(×3)","Hands as Mickey ears → 'Disney!' (×3)",
    "我想像迪士尼一样 — 创造 ___ !","Like Disney — I want to create ___!")
n+=1; pn(s,n)
notes(s,"2 分钟:\n• Q1, Q2 抢答\n• 让 3 个学生喊最喜欢的角色 (Elsa, Mickey, Moana...)\n• TPR 米奇耳朵\n• 句型: 让学生说想创造的 (动画? 乐园? 玩具?)")

# --- 马云 Jack Ma ---
s=story_slide("🛒","马云","Jack Ma","1964– 至今","🇨🇳 中国",MAYUN,
    [("❓","他看到什么需要?","What need did he see?",
       "中国人买东西不方便 — 想买好东西要跑很远. 商家也卖不出去.",
       "In China, buying stuff was hard — long trips, sellers struggled too.",
       "你网上买过东西吗?"),
     ("❓","他做了什么?","What did he do?",
       "创办淘宝 / 阿里巴巴 — 让全中国人在网上买 + 卖!",
       "Founded Taobao/Alibaba — online shopping for all of China!",
       "他以前是老师 — 你猜他懂电脑吗?"),
     ("❓","结果是什么?","What was the result?",
       "中国人在网上买几乎所有东西 — 包裹一天几亿个!",
       "Chinese buy almost everything online — millions of packages daily.",
       "你家一周收几个包裹?")],
    "🎓 他以前是英语老师, 完全不懂电脑! 但是他看到了需要.",
    "He was an English teacher who knew NOTHING about computers — but saw the need!",
    video_cn="搜: 'Jack Ma 马云 故事 中文'  ·  CCTV 少儿 「马云 创业」 短片",
    video_en="Search 'Jack Ma for kids' / 'Alibaba founder story'")
n+=1; pn(s,n)
notes(s,"5 分钟:\n🎥 1-2 分钟 — 推荐: 'Jack Ma 创业故事' 央视少儿版本\n📋 3 分钟答:\n• Q1: 中国当年没有方便的网购\n• Q2: 强调 — 他是英语老师, 不懂电脑! 重要的是看到需要 + 找懂的人合作\n• Q3: 让学生说自家一周收几个包裹 (淘宝 / 京东 / Amazon)")

s=qa_slide("🛒","马云","Jack Ma",MAYUN,
    [("📚","马云以前做什么工作?","Jack Ma's old job?",
      "英语老师","English teacher"),
     ("🛍️","他创办的网站叫什么?","His website's name?",
      "淘宝 / 阿里巴巴","Taobao/Alibaba")],
    "你 / 你家用淘宝还是 Amazon? 一周几次?","Taobao or Amazon? How often?",
    "📱 假装滑手机点「下单!」 → 「叮咚! 包裹来了!」","Tap phone → 'Order!' → 'Ding! Package!'",
    "我想卖 ___ 给 ___ !","I want to sell ___ to ___!")
n+=1; pn(s,n)
notes(s,"2 分钟:\n• Q1 答案出乎意料 — 老师! → 让学生 wow\n• Q2 抢答 +2\n• 个人连接: 让 3 个学生说一周收几个包裹\n• TPR 假装下单 → 听叮咚 → 收包裹 (3 次, 表演)")

# --- Milton Hershey ---
s=story_slide("🍫","赫尔希","Milton Hershey","1857–1945","🇺🇸 美国",HERSHEY,
    [("❓","他看到什么需要?","What need did he see?",
       "100 年前, 巧克力很贵 — 只有富人吃得起. 他想让大家都吃到!",
       "100 yrs ago, chocolate was expensive — only rich. He wanted it for all.",
       "你最喜欢什么巧克力?"),
     ("❓","他做了什么?","What did he do?",
       "想出大量做牛奶巧克力的办法 — 又便宜又好吃!",
       "Found a way to mass-produce milk chocolate — cheap & tasty!",
       "你猜他失败了几次?"),
     ("❓","结果是什么?","What was the result?",
       "全世界都吃到了巧克力! 他还建了一个城市叫 Hershey.",
       "Everyone got chocolate! He even built a city called Hershey.",
       "他用钱做什么帮人?")],
    "🏫 他把钱用来给孤儿建学校 — 帮人比赚钱更重要!",
    "He spent his money on schools for orphans — helping > earning!",
    video_cn="搜: 'Milton Hershey for kids' / Hershey 工厂 tour 视频  ·  绘本: 「Who Was Milton Hershey?」",
    video_en="Picture book: 'Who Was Milton Hershey?' / Search 'Hershey chocolate factory kids'")
n+=1; pn(s,n)
notes(s,"5 分钟:\n🎥 1-2 分钟 — 推荐: 'How Hershey's Chocolate is Made' 工厂 tour\n📋 3 分钟答:\n• Q1: 巧克力当年很贵 — 只有富人吃得到\n• Q2: 失败很多次才成功 mass produce — 跟爱迪生一样\n• Q3: 重点话 — 「他把钱用来帮孤儿! 钱不是目标, 帮人是」")

s=qa_slide("🍫","赫尔希","Milton Hershey",HERSHEY,
    [("💰","以前巧克力是给谁吃的?","Who used to eat chocolate?",
      "只有富人","Only the rich"),
     ("🏫","赫尔希用钱做什么好事?","His good deed with money?",
      "建学校给孤儿","Schools for orphans")],
    "你最喜欢什么巧克力 / 糖? 想发明什么新糖?","Favorite chocolate? Wanna invent a new candy?",
    "🍬 假装拿巧克力 → 咬一口 → 「mmmm! 好吃!」","Pretend chocolate bite → 'mmm! yummy!' (×3)",
    "我想发明一个 ___ 糖!","I want to invent a ___ candy!")
n+=1; pn(s,n)
notes(s,"2 分钟:\n• Q1, Q2 抢答\n• TPR 吃巧克力表演 — 全班一起 → 大笑\n• 句型: 让学生发明新糖 (彩虹糖? 会跳的糖? 巨型糖?)\n• 收尾: 「他帮孤儿建学校」 — 钱用来帮人才有意义")

# ============================================================
# === SESSION 1 · PHASE 4: APPLY (10 min) ===
# ============================================================
phase_marker("🌱","Apply","Match + My Idea!",10,PH_APPLY,
             "猜 + 想 — 我也想当企业家!","Match + Brainstorm — I want to be one too!")
n+=1; pn(s:=prs.slides[-1],n)

# Apply 1 — matching game
s=ns(); bg(s,CREAM); hb(s,"🎮 谁做了什么?  Who Made What?",LAB)
group_label(s)
score_badge(s)
tb(s,0.4,1.25,9.2,0.40,"老师说产品 — 第一队答名字 +2 分!",sz=17,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.65,9.2,0.30,"Teacher says product — first team to name = +2!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
matches=[
    ("📱","iPhone","iPhone","🍎","乔布斯","Steve Jobs",JOBS),
    ("🏰","迪士尼乐园","Disneyland","🐭","沃尔特·迪士尼","Walt Disney",DISNEY),
    ("🛒","淘宝","Taobao","👨‍💼","马云","Jack Ma",MAYUN),
    ("🍫","Hershey 巧克力","Hershey choc.","🏭","赫尔希","Milton Hershey",HERSHEY),
]
for i,(em1,p_cn,p_en,em2,a_cn,a_en,cl) in enumerate(matches):
    y=2.10+i*0.62
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(y),Inches(9.2),Inches(0.55))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(2)
    tb(s,0.55,y+0.10,0.5,0.4,em1,sz=22,a=PP_ALIGN.CENTER)
    tb(s,1.10,y+0.07,2.5,0.30,p_cn,sz=14,b=True,c=DARK)
    tb(s,1.10,y+0.32,2.5,0.25,p_en,sz=10,c=GRAY)
    tb(s,3.85,y+0.07,0.5,0.4,"→",sz=20,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,4.45,y+0.10,0.5,0.4,em2,sz=22,a=PP_ALIGN.CENTER)
    tb(s,5.00,y+0.07,4.4,0.30,a_cn,sz=14,b=True,c=cl)
    tb(s,5.00,y+0.32,4.4,0.25,a_en,sz=10,c=GRAY)
sentence_frame_bar(s,4.65,"___ 是 ___ 创办的!","___ was founded by ___!")
n+=1; pn(s,n)
notes(s,"🌱 APPLY 1 · 5 分钟 — 3 轮快游戏:\n• 第 1 轮: 老师说产品 → 第一队抢答名字\n• 第 2 轮: 老师说名字 → 第一队抢答产品\n• 第 3 轮: 用句型 「___ 是 ___ 创办的」 (+3)")

# Apply 2 — 我想做的生意
s=ns(); bg(s,CREAM); hb(s,"💭 我想做的生意!  My Business Idea!",GREEN)
group_label(s)
score_badge(s)
tb(s,0.4,1.25,9.2,0.45,"想一想 — 你想解决什么问题?",sz=20,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.75,9.2,0.30,"Think — what problem do YOU want to solve?",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
prompts=[
    ("🍔","食物","Food","早饭? 零食?"),
    ("🎮","好玩","Fun","新玩具? 游戏?"),
    ("📚","学习","Learning","新文具? App?"),
    ("🐾","动物","Pets","流浪猫的家?"),
]
for i,(em,cn,en,ex) in enumerate(prompts):
    x=0.4+i*2.30
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(2.20),Inches(2.20),Inches(2.10))
    sh.fill.solid(); sh.fill.fore_color.rgb=WARM; sh.line.color.rgb=GREEN; sh.line.width=Pt(2)
    tb(s,x+0.05,2.30,2.10,0.7,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.05,2.10,0.40,cn,sz=16,b=True,c=GREEN,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.45,2.10,0.30,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.80,2.10,0.40,f"💭 {ex}",sz=11,c=DARK,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.45,"我想做 ___ 帮 ___ 解决 ___ !","I want to make ___ to help ___ with ___!")
n+=1; pn(s,n)
notes(s,"🌱 APPLY 2 · 5 分钟:\n• 1 分钟个人想 — 写/画在 booklet\n• 2 分钟组内分享 (每人 30 秒)\n• 2 分钟: 每队选 1 位代表上台说 — +1 分\n• 不评判好坏 — 越奇怪越好!")

# ============================================================
# === SESSION 1 · PHASE 5: SHARE & CLOSE (5 min) ===
# ============================================================
s=ns(); bg(s,CREAM); hb(s,"🎤 总结 + 下一步  Summary + Next!",PH_CLOSE)
score_badge(s)
tb(s,0.4,0.85,9.2,0.4,"🧭 早上学了什么?",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
left=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.40),Inches(9.2),Inches(1.85))
left.fill.solid(); left.fill.fore_color.rgb=BIZ; left.line.fill.background()
tb(s,0.4,1.55,9.2,0.50,"💡 企业家 3 步: 👀 发现 → 💡 创造 → 🤝 帮人!",sz=22,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,0.4,2.10,9.2,0.30,"3 steps: Find a need → Make a product → Help people",sz=11,c=IDEA,a=PP_ALIGN.CENTER)
tb(s,0.4,2.45,9.2,0.30,"🍎 乔布斯 · 🏰 迪士尼 · 🛒 马云 · 🍫 赫尔希",sz=15,b=True,c=IDEA,a=PP_ALIGN.CENTER)
tb(s,0.4,2.80,9.2,0.25,"4 entrepreneurs we met today",sz=10,c=IDEA,a=PP_ALIGN.CENTER)
trans=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(3.40),Inches(9.2),Inches(1.45))
trans.fill.solid(); trans.fill.fore_color.rgb=PH_ACTIVE; trans.line.color.rgb=IDEA; trans.line.width=Pt(3)
tb(s,0.4,3.50,9.2,0.45,"📦 下午你来做 entrepreneur — 设计新书包!",sz=20,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,0.4,3.95,9.2,0.30,"Afternoon: YOU design a better backpack!",sz=11,c=IDEA,a=PP_ALIGN.CENTER)
tb(s,0.4,4.30,9.2,0.40,"🛒 然后卖给同学 — 看谁是销冠 (top seller)!",sz=15,b=True,c=IDEA,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"🎤 SHARE & CLOSE · 5 分钟:\n• 1 分钟: 全班齐喊 — 「发现! 创造! 帮人!」 (×3)\n• 1 分钟: 4 队复习 4 位企业家 (每队 1 个: 🔴 → Jobs, 🔵 → Disney, 🟢 → 马云, 🟡 → Hershey)\n• 1 分钟: Session 1 积分公布\n• 2 分钟: Tease 下午 — 「书包太重怎么办? 你设计! 你卖!」")

# ============================================================
# 22. SESSION 2 DIVIDER (Day 1-style: Review + Read + Write)
# ============================================================
s=div("Session 2  下午 1:00–1:50","📚 复习 + 我会认 + 我会写  Review · Read · Write",IDEA,"📖"); n+=1; pn(s,n)

# ============================================================
# 23. REVIEW
# ============================================================
s=ns(); bg(s,CREAM); hb(s,"🔄 复习  Review · Session 1",NAVY)
tb(s,0.4,0.85,9.2,0.40,"早上学了什么?  What did we learn this morning?",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.25,9.2,0.30,"Quick recap before we read & write!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
recap=[
    ("💡","企业家 3 步","3-step entrepreneur","发现 → 创造 → 帮人",BIZ),
    ("🍎","乔布斯 · 迪士尼","Jobs · Disney","iPhone · 米老鼠 · 乐园",JOBS),
    ("🛒","马云 · 赫尔希","Jack Ma · Hershey","淘宝 · 巧克力",MAYUN),
    ("💭","我也想…","I also want to…","「我想做 ___ 帮 ___」",GREEN),
]
for i,(em,cn,en,detail,c) in enumerate(recap):
    col=i%2; row=i//2
    x=0.4+col*4.65
    y=1.65+row*1.45
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(4.55),Inches(1.30))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE
    sh.line.color.rgb=c; sh.line.width=Pt(2)
    tb(s,x+0.10,y+0.12,0.7,0.65,em,sz=32,a=PP_ALIGN.CENTER)
    tb(s,x+0.85,y+0.10,3.6,0.40,cn,sz=15,b=True,c=c)
    tb(s,x+0.85,y+0.45,3.6,0.30,en,sz=10,c=GRAY)
    tb(s,x+0.15,y+0.85,4.30,0.40,detail,sz=11,b=True,c=DARK)
sentence_frame_bar(s,4.65,
    "今天早上我学了 ___ 。",
    "This morning I learned about ___.")
n+=1; pn(s,n)
notes(s,"5 分钟 review:\n• 「早上学了哪 4 位企业家?」抢答\n• 复习: 3 步 — 发现 → 创造 → 帮人\n• 抽 1-2 个学生说 「我想做 ___ 帮 ___」")

# ============================================================
# 24-29. 我会认 — 6 vocabulary words
# ============================================================
read_words=[
    ("👨‍💼","企业家","qǐ yè jiā","Entrepreneur",BIZ,
        "乔布斯 是一位 企业家。",
        "📷 企业家 / 西装 / 创业"),
    ("🛍️","产品","chǎn pǐn","Product",JOBS,
        "iPhone 是苹果公司 的 产品。",
        "📷 产品 / 包装 / 商店"),
    ("🤝","顾客","gù kè","Customer",MAYUN,
        "顾客 喜欢 买 好 产品。",
        "📷 顾客 / 购物 / 收银"),
    ("🛒","买","mǎi","Buy",MONEY,
        "我 在商店 买 苹果。",
        "📷 买东西 / 购物车 / 购物"),
    ("💰","卖","mài","Sell",HERSHEY,
        "餐厅 卖 好吃 的 菜。",
        "📷 卖东西 / 招牌 / 摊位"),
    ("💵","钱","qián","Money",MONEY,
        "买东西 要 用 钱。",
        "📷 钱 / 钞票 / 硬币"),
]
for em,cn,py,en,c,sent,img_label in read_words:
    s=ns(); bg(s,CREAM); hb(s,f"👀 我会认 · {cn}  I Can Read",c)
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.0),Inches(4.5),Inches(2.5))
    sh.fill.solid(); sh.fill.fore_color.rgb=WARM; sh.line.fill.background()
    tb(s,0.5,1.10,4.3,1.4,cn,sz=64,b=True,c=c,a=PP_ALIGN.CENTER)
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
    notes(s,f"3 分钟 — {cn}:\n• 老师指字, 全班齐读 3 遍 (慢→快→唱)\n• 看图: 「这是 ___, 在做什么?」\n• 读例句, 跟读\n• 抽 1-2 个学生用 {cn} 造一个新句子")

# ============================================================
# 30. 我会写 · 产品
# ============================================================
s=ns(); bg(s,CREAM); hb(s,"✏️ 我会写 · 产品  I Can Write · Product",JOBS)
tb(s,0.4,0.85,9.2,0.40,"练一练 — 写「产品」!",sz=22,b=True,c=JOBS,a=PP_ALIGN.CENTER)
tb(s,0.4,1.25,9.2,0.30,"Practice writing 产品 (Product)",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
char1=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.65),Inches(2.30),Inches(2.85))
char1.fill.solid(); char1.fill.fore_color.rgb=WHITE
char1.line.color.rgb=JOBS; char1.line.width=Pt(3)
tb(s,0.4,1.95,2.30,1.95,"产",sz=130,b=True,c=JOBS,a=PP_ALIGN.CENTER)
tb(s,0.4,4.10,2.30,0.30,"chǎn (produce)",sz=12,b=True,c=GRAY,a=PP_ALIGN.CENTER)
char2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(2.85),Inches(1.65),Inches(2.30),Inches(2.85))
char2.fill.solid(); char2.fill.fore_color.rgb=WHITE
char2.line.color.rgb=JOBS; char2.line.width=Pt(3)
tb(s,2.85,1.95,2.30,1.95,"品",sz=130,b=True,c=JOBS,a=PP_ALIGN.CENTER)
tb(s,2.85,4.10,2.30,0.30,"pǐn (item)",sz=12,b=True,c=GRAY,a=PP_ALIGN.CENTER)
panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.30),Inches(1.65),Inches(4.30),Inches(2.85))
panel.fill.solid(); panel.fill.fore_color.rgb=WHITE
panel.line.color.rgb=JOBS; panel.line.width=Pt(2.5)
head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.30),Inches(1.65),Inches(4.30),Inches(0.50))
head.fill.solid(); head.fill.fore_color.rgb=JOBS; head.line.fill.background()
tb(s,5.45,1.72,4.10,0.4,"✏️ 怎么写  How to write",sz=13,b=True,c=WHITE)
tb(s,5.45,2.30,4.10,0.40,"1️⃣ 「产」 — 6 笔",sz=14,b=True,c=DARK)
tb(s,5.45,2.65,4.10,0.30,"  上面「立」, 下面「丿」",sz=10,c=GRAY)
tb(s,5.45,3.05,4.10,0.40,"2️⃣ 「品」 — 9 笔",sz=14,b=True,c=DARK)
tb(s,5.45,3.40,4.10,0.30,"  3 个「口」 — 像 3 个商品摆在一起!",sz=10,c=GRAY)
tb(s,5.45,3.85,4.10,0.40,"📝 在田字格练 3 遍",sz=12,b=True,c=JOBS)
tb(s,5.45,4.20,4.10,0.30,"Practice 3 times in grid paper",sz=9,c=GRAY)
sentence_frame_bar(s,4.65,
    "我会写「产品」! 我的产品是 ___ 。",
    "I can write 产品! My product is ___.")
n+=1; pn(s,n)
notes(s,"5-6 分钟:\n• 演示笔顺, 学生跟写 (空中写)\n• 田字格练 3 遍\n• 提示: 「品」 = 3 个口 — 像 3 个商品摆在一起 (好记!)\n• 让学生说自己的产品 (下午会做)")

# ============================================================
# 31. 我会写 · 买
# ============================================================
s=ns(); bg(s,CREAM); hb(s,"✏️ 我会写 · 买  I Can Write · Buy",MONEY)
tb(s,0.4,0.85,9.2,0.40,"练一练 — 写「买」!",sz=22,b=True,c=MONEY,a=PP_ALIGN.CENTER)
tb(s,0.4,1.25,9.2,0.30,"Practice writing 买 (Buy)",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.65),Inches(4.0),Inches(2.85))
sh.fill.solid(); sh.fill.fore_color.rgb=WARM; sh.line.color.rgb=MONEY; sh.line.width=Pt(3)
tb(s,0.4,1.95,4.0,2.0,"买",sz=200,b=True,c=MONEY,a=PP_ALIGN.CENTER)
tb(s,0.4,4.10,4.0,0.40,"mǎi · buy",sz=18,c=GRAY,a=PP_ALIGN.CENTER)
panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.55),Inches(1.65),Inches(5.15),Inches(2.85))
panel.fill.solid(); panel.fill.fore_color.rgb=WHITE
panel.line.color.rgb=MONEY; panel.line.width=Pt(2.5)
head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.55),Inches(1.65),Inches(5.15),Inches(0.50))
head.fill.solid(); head.fill.fore_color.rgb=MONEY; head.line.fill.background()
tb(s,4.70,1.72,4.95,0.4,"✏️ 怎么写  How to write",sz=13,b=True,c=WHITE)
tb(s,4.70,2.30,4.95,0.40,"「买」 — 6 笔",sz=14,b=True,c=DARK)
tb(s,4.70,2.70,4.95,0.30,"  上面「⺈」(像帽子), 下面「头」字",sz=10,c=GRAY)
tb(s,4.70,3.10,4.95,0.40,"💡 记忆: 买东西要用头想 — 别乱买!",sz=12,b=True,c=MONEY)
tb(s,4.70,3.45,4.95,0.30,"Memory: 'Use your head when buying!'",sz=9,c=GRAY)
tb(s,4.70,3.85,4.95,0.40,"📝 在田字格练 3 遍",sz=12,b=True,c=MONEY)
tb(s,4.70,4.20,4.95,0.30,"Practice 3 times in grid paper",sz=9,c=GRAY)
sentence_frame_bar(s,4.65,
    "我会写「买」! 我想买 ___ 。",
    "I can write 买! I want to buy ___.")
n+=1; pn(s,n)
notes(s,"5-6 分钟:\n• 演示笔顺\n• 提示: 上面像帽子的是「⺈」, 下面是「头」 — 「买东西要用头 (动脑)」\n• 让学生说想买什么 (轮流上台)")

# ============================================================
# 32. 我会写 · 卖
# ============================================================
s=ns(); bg(s,CREAM); hb(s,"✏️ 我会写 · 卖  I Can Write · Sell",HERSHEY)
tb(s,0.4,0.85,9.2,0.40,"练一练 — 写「卖」!",sz=22,b=True,c=HERSHEY,a=PP_ALIGN.CENTER)
tb(s,0.4,1.25,9.2,0.30,"Practice writing 卖 (Sell)",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.65),Inches(4.0),Inches(2.85))
sh.fill.solid(); sh.fill.fore_color.rgb=WARM; sh.line.color.rgb=HERSHEY; sh.line.width=Pt(3)
tb(s,0.4,1.95,4.0,2.0,"卖",sz=200,b=True,c=HERSHEY,a=PP_ALIGN.CENTER)
tb(s,0.4,4.10,4.0,0.40,"mài · sell",sz=18,c=GRAY,a=PP_ALIGN.CENTER)
panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.55),Inches(1.65),Inches(5.15),Inches(2.85))
panel.fill.solid(); panel.fill.fore_color.rgb=WHITE
panel.line.color.rgb=HERSHEY; panel.line.width=Pt(2.5)
head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.55),Inches(1.65),Inches(5.15),Inches(0.50))
head.fill.solid(); head.fill.fore_color.rgb=HERSHEY; head.line.fill.background()
tb(s,4.70,1.72,4.95,0.4,"✏️ 怎么写  How to write",sz=13,b=True,c=WHITE)
tb(s,4.70,2.30,4.95,0.40,"「卖」 — 8 笔",sz=14,b=True,c=DARK)
tb(s,4.70,2.70,4.95,0.30,"  上面多一个「十」, 下面跟「买」一样!",sz=10,c=GRAY)
tb(s,4.70,3.10,4.95,0.40,"💡 记忆: 「买」+「十」=「卖」!",sz=12,b=True,c=HERSHEY)
tb(s,4.70,3.45,4.95,0.30,"Memory: 买 (buy) + 十 (cross) = 卖 (sell)",sz=9,c=GRAY)
tb(s,4.70,3.85,4.95,0.40,"📝 在田字格练 3 遍",sz=12,b=True,c=HERSHEY)
tb(s,4.70,4.20,4.95,0.30,"Practice 3 times in grid paper",sz=9,c=GRAY)
sentence_frame_bar(s,4.65,
    "我会写「卖」! 我想卖 ___ 。",
    "I can write 卖! I want to sell ___.")
n+=1; pn(s,n)
notes(s,"6-7 分钟:\n• 演示笔顺 — 「卖」比「买」多一个「十」在顶上\n• 重点记忆: 买 + 十 = 卖 (好记!)\n• K 学生只写「买」就行; G3+ 都写\n• 完成 → 「我会写」贴纸; 写得最好 → 队 +2 分")

# ============================================================
# 33. SESSION 3 DIVIDER — Project
# ============================================================
s=div("Session 3  下午 2:00–2:50","🛠️ 项目课  Backpack Redesign  ·  50 min",PH_ACTIVE,"📦"); n+=1; pn(s,n)

# ============================================================
# === SESSION 3 · PHASE 1: HOOK (5 min) ===
# ============================================================
phase_marker("🔥","Hook","Heavy Backpack!",5,PH_HOOK,
             "书包太重 — 怎么办?","Backpack is too heavy — how to fix?")
n+=1; pn(s:=prs.slides[-1],n)

s=ns(); bg(s,CREAM); hb(s,"🎒 想象一下…  Imagine…",LAB)
score_badge(s)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(3.25),Inches(0.95),Inches(3.5),Inches(2.85))
sh.fill.solid(); sh.fill.fore_color.rgb=NAVY; sh.line.color.rgb=DARK; sh.line.width=Pt(4)
tb(s,3.25,1.20,3.5,1.5,"🎒",sz=110,a=PP_ALIGN.CENTER)
tb(s,3.25,2.80,3.5,0.45,"我的书包太重!",sz=18,b=True,c=IDEA,a=PP_ALIGN.CENTER)
tb(s,3.25,3.30,3.5,0.40,"Too heavy!",sz=14,c=IDEA,a=PP_ALIGN.CENTER)
tb(s,0.4,4.00,9.2,0.55,"❓ 书包太重 — 怎么办?",sz=24,b=True,c=LAB,a=PP_ALIGN.CENTER)
tb(s,0.4,4.55,9.2,0.30,"My backpack is too heavy — how do we fix it?",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
tpr_strip(s,4.95,"假装背书包 → 弯腰 → 喊「太重了!」(×3)","Pretend backpack → bend → 'Too heavy!' (×3)")
n+=1; pn(s,n)
notes(s,"🔥 HOOK · 5 分钟:\n• 老师拎一个真的书包 (装重东西) — 表演走路弯腰\n• 让 1-2 个学生试着背 — 喊「太重了!」\n• 转: 「这是真的问题! 你是 entrepreneur — 怎么解决?」\n• 全班收集想法: 加轮子? 减书? 用电子书? 双肩? 飞起来?")

# ============================================================
# === SESSION 3 · PHASE 2: MINI-LESSON (10 min) ===
# ============================================================
phase_marker("📚","Mini-Lesson","Entrepreneur 4 Steps",10,PH_MINI,
             "发现 → 设计 → 做 → 卖!","Find → Design → Make → Sell!")
n+=1; pn(s:=prs.slides[-1],n)

# Design loop
s=ns(); bg(s,CREAM); hb(s,"🔁 企业家 4 步循环  Entrepreneur 4-Step Loop",BIZ)
score_badge(s)
tb(s,0.4,0.85,9.2,0.4,"4 步 — 你也能做!",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.30,"4 steps — you can do it too!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
steps=[
    ("👀","发现","Find","看问题",NAVY),
    ("💡","设计","Design","画想法",BIZ),
    ("🔨","做","Make","做模型",PH_ACTIVE),
    ("🛒","卖","Sell","告诉顾客",MONEY),
]
for i,(em,cn,en,desc,cl) in enumerate(steps):
    x=0.4+i*2.30
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.75),Inches(2.10),Inches(2.55))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    nb=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x+0.78),Inches(1.85),Inches(0.55),Inches(0.55))
    nb.fill.solid(); nb.fill.fore_color.rgb=cl; nb.line.fill.background()
    tb(s,x+0.78,1.93,0.55,0.4,str(i+1),sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.50,2.0,0.6,em,sz=42,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.20,2.0,0.45,cn,sz=22,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.65,2.0,0.30,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.95,2.0,0.30,desc,sz=11,c=DARK,a=PP_ALIGN.CENTER)
    if i<3:
        tb(s,x+2.10,2.85,0.30,0.4,"→",sz=22,b=True,c=PURPLE,a=PP_ALIGN.CENTER)
tb(s,0.4,4.40,9.2,0.30,"🔁 第一次失败? 改一改 → 再来! (像爱迪生试 1000 次)",sz=14,b=True,c=PURPLE,a=PP_ALIGN.CENTER)
tb(s,0.4,4.70,9.2,0.25,"Failed first try? Iterate — like Edison's 1000 tries!",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"📚 MINI 1 · 4 分钟:\n• 4 步走一遍 — 加手势 (👀 看, 💡 头脑风暴, 🔨 拳头敲, 🛒 推车)\n• 全班齐声: 发现! 设计! 做! 卖! (×3)\n• 强调: 失败也没关系 — 改 → 再做")

# Materials + Rules
s=ns(); bg(s,CREAM); hb(s,"🛠️ 材料 + 比赛规则  Materials + Rules",PH_ACTIVE)
score_badge(s)
tb(s,0.4,0.85,9.2,0.4,"每队一套材料 — 30 分钟设计你的新书包!",sz=15,b=True,c=DARK,a=PP_ALIGN.CENTER)
mats=[
    ("📃","白纸","Paper"),
    ("✏️","彩笔","Markers"),
    ("📦","纸盒","Box"),
    ("✂️","剪刀","Scissors"),
    ("🩹","胶带","Tape"),
]
for i,(em,cn,en) in enumerate(mats):
    x=0.4+i*1.85
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.40),Inches(1.70),Inches(2.10))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=PH_ACTIVE; sh.line.width=Pt(2)
    tb(s,x,1.50,1.70,0.85,em,sz=50,a=PP_ALIGN.CENTER)
    tb(s,x,2.55,1.70,0.40,cn,sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
    tb(s,x,2.95,1.70,0.30,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(3.65),Inches(9.2),Inches(1.30))
sh.fill.solid(); sh.fill.fore_color.rgb=WARM; sh.line.color.rgb=PH_ACTIVE; sh.line.width=Pt(2)
tb(s,0.5,3.75,9.0,0.35,"📋 比赛规则  Competition Rules",sz=14,b=True,c=PH_ACTIVE)
tb(s,0.5,4.10,9.0,0.30,"1️⃣ 30 分钟设计 + 做模型 (画 / 用纸 + 胶带拼)",sz=12,c=DARK)
tb(s,0.5,4.40,9.0,0.30,"2️⃣ 最后 5 分钟 — 每队 30 秒推销! 全班投票 — 最想买的 +5 分!",sz=12,c=DARK)
tb(s,0.5,4.70,9.0,0.30,"3️⃣ 创意 + 解决真实问题 = 更高分",sz=12,b=True,c=MONEY)
n+=1; pn(s,n)
notes(s,"📚 MINI 2 · 6 分钟:\n• 1 分钟: 展示材料 — 让 1 学生上台摸\n• 2 分钟: 老师 demo 一些 features (轮子? 灯? 减重?)\n• 1 分钟: 各队选「首席设计师」 (rotation)\n• 2 分钟: 强调规则 + 30 秒推销")

# ============================================================
# === SESSION 3 · PHASE 3: ACTIVE PRACTICE (20 min) ===
# ============================================================
phase_marker("🎯","Active Practice","Brainstorm → Design → Make!",20,PH_ACTIVE,
             "想 → 画 → 做 → 改!","Think → Draw → Build → Improve!")
n+=1; pn(s:=prs.slides[-1],n)

# Step 1: Brainstorm features
s=ns(); bg(s,CREAM); hb(s,"💭 第 1 步 · 想 features  Step 1 · Brainstorm",BIZ)
group_label(s)
score_badge(s)
tb(s,0.4,1.25,9.2,0.4,"⏱️ 5 分钟 — 想 5+ 个 cool 功能!",sz=18,b=True,c=BIZ,a=PP_ALIGN.CENTER)
ideas=[
    ("🎡","加轮子","Wheels"),
    ("🪶","减重量","Lighter"),
    ("💡","加灯","Add light"),
    ("🔋","充电口","USB port"),
    ("🌧️","防雨","Waterproof"),
    ("🎨","好看","Cool design"),
]
for i,(em,cn,en) in enumerate(ideas):
    col=i%3; row=i//3
    x=0.4+col*3.10; y=1.85+row*1.20
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(2.95),Inches(1.05))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=BIZ; sh.line.width=Pt(2)
    tb(s,x+0.10,y+0.15,0.85,0.75,em,sz=38,a=PP_ALIGN.CENTER)
    tb(s,x+1.05,y+0.20,1.85,0.40,cn,sz=15,b=True,c=BIZ)
    tb(s,x+1.05,y+0.55,1.85,0.30,en,sz=10,c=GRAY)
sentence_frame_bar(s,4.40,"我想加 ___ — 因为 ___ !","I want to add ___ — because ___!",accent=BIZ)
n+=1; pn(s,n)
notes(s,"🎯 STEP 1 · 5 分钟:\n• 各队头脑风暴 — 至少 5 个 features (PPT 6 个例子是 starter)\n• 鼓励越奇怪越好: 飞? 会说话? 自己上学?\n• 写在 booklet 一页\n• 老师巡视, 不评判")

# Step 2: Sketch design
s=ns(); bg(s,CREAM); hb(s,"✏️ 第 2 步 · 画设计图  Step 2 · Sketch",BIZ)
group_label(s)
score_badge(s)
tb(s,0.4,1.25,9.2,0.4,"⏱️ 5 分钟 — 画你的新书包!",sz=18,b=True,c=BIZ,a=PP_ALIGN.CENTER)
left=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.85),Inches(5.0),Inches(2.85))
left.fill.solid(); left.fill.fore_color.rgb=WHITE; left.line.color.rgb=BIZ; left.line.width=Pt(2.5)
tb(s,0.4,1.95,5.0,0.45,"📐 画你的新书包",sz=16,b=True,c=BIZ,a=PP_ALIGN.CENTER)
tb(s,0.4,2.40,5.0,0.30,"Draw your new backpack",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,0.6,2.80,4.6,0.30,"• 大小 / shape",sz=12,c=DARK)
tb(s,0.6,3.10,4.6,0.30,"• 标记 features (灯? 轮子? 口袋?)",sz=12,c=DARK)
tb(s,0.6,3.40,4.6,0.30,"• 给它起一个名字!",sz=12,b=True,c=BIZ)
tb(s,0.6,3.70,4.6,0.30,"• 选你的顾客 — 谁会用?",sz=12,c=DARK)
right=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.55),Inches(1.85),Inches(4.05),Inches(2.85))
right.fill.solid(); right.fill.fore_color.rgb=WARM; right.line.color.rgb=BIZ; right.line.width=Pt(2.5)
tb(s,5.55,1.95,4.05,0.45,"💡 提示  Tips",sz=16,b=True,c=BIZ,a=PP_ALIGN.CENTER)
tb(s,5.70,2.55,3.85,0.30,"✓ 至少标 3 个 features",sz=12,c=DARK)
tb(s,5.70,2.85,3.85,0.30,"✓ 加颜色 / 装饰",sz=12,c=DARK)
tb(s,5.70,3.15,3.85,0.30,"✓ 想一个 cool 名字",sz=12,c=DARK)
tb(s,5.70,3.45,3.85,0.30,"✓ 想 — 谁是顾客?",sz=12,c=DARK)
sentence_frame_bar(s,4.85,"我的书包叫 ___ , 给 ___ 用!","My backpack is called ___, for ___!",accent=BIZ)
n+=1; pn(s,n)
notes(s,"🎯 STEP 2 · 5 分钟:\n• 各队 1 张大纸 + 彩笔 — 一起画\n• 强调标 features (画 + 写)\n• 名字很重要 — 像「飞鼠书包」「自走包」「彩虹小包」\n• 老师巡视 — 帮 K 学生写名字")

# Step 3: Make prototype
s=ns(); bg(s,CREAM); hb(s,"🔨 第 3 步 · 做模型  Step 3 · Build!",PH_ACTIVE)
group_label(s)
score_badge(s)
tb(s,0.4,1.25,9.2,0.4,"⏱️ 10 分钟 — 用纸 + 胶带做一个模型!",sz=18,b=True,c=PH_ACTIVE,a=PP_ALIGN.CENTER)
tips=[
    ("📦","用纸盒","Use box","当主体"),
    ("✂️","剪形状","Cut shapes","加 features"),
    ("🩹","胶带粘","Tape it","固定装饰"),
    ("🏷️","贴标签","Add labels","写 features 名字"),
]
for i,(em,cn,en,desc) in enumerate(tips):
    col=i%2; row=i//2
    x=0.4+col*4.65; y=1.85+row*1.30
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(4.55),Inches(1.20))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=PH_ACTIVE; sh.line.width=Pt(2)
    tb(s,x+0.10,y+0.15,0.95,0.95,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+1.10,y+0.15,3.30,0.40,cn,sz=15,b=True,c=PH_ACTIVE)
    tb(s,x+1.10,y+0.55,3.30,0.30,en,sz=10,c=GRAY)
    tb(s,x+1.10,y+0.85,3.30,0.30,desc,sz=11,c=DARK)
sentence_frame_bar(s,4.55,"看! 我做了 ___ — 它能 ___ !","Look! I made ___ — it can ___!",accent=PH_ACTIVE)
n+=1; pn(s,n)
notes(s,"🎯 STEP 3 · 10 分钟 — 动手!\n• 各队一起做 — 一个 paper backpack 模型\n• 不一定漂亮 — 重点是表达 features\n• 鼓励试错 — 「失败也是数据!」\n• 最后 2 分钟 — 各队加标签 / 价格牌")

# ============================================================
# === SESSION 3 · PHASE 4: APPLY (10 min) ===
# ============================================================
phase_marker("🌱","Apply","Pitch Prep!",10,PH_APPLY,
             "起名字 + 定价格 + 准备推销词","Name + Price + Pitch script")
n+=1; pn(s:=prs.slides[-1],n)

s=ns(); bg(s,CREAM); hb(s,"📝 准备推销!  Prepare to Pitch!",GREEN)
group_label(s)
score_badge(s)
tb(s,0.4,1.25,9.2,0.4,"3 个问题 — 准备 30 秒推销!",sz=18,b=True,c=GREEN,a=PP_ALIGN.CENTER)
tb(s,0.4,1.65,9.2,0.30,"3 questions to prep your 30-sec pitch!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
qs=[
    ("🏷️","名字是什么?","What's the name?","___ 包",BIZ),
    ("💰","卖多少钱?","How much?","$ ___",MONEY),
    ("👥","谁是顾客?","Who's the customer?","小学生? 老师?",PURPLE),
]
for i,(em,q_cn,q_en,hint,cl) in enumerate(qs):
    x=0.4+i*3.10
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.95),Inches(2.95),Inches(2.55))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    tb(s,x+0.05,2.10,2.85,0.7,em,sz=48,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.95,2.85,0.45,q_cn,sz=18,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.45,2.85,0.30,q_en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    pill=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x+0.20),Inches(3.85),Inches(2.55),Inches(0.55))
    pill.fill.solid(); pill.fill.fore_color.rgb=WARM; pill.line.color.rgb=cl; pill.line.width=Pt(1.5)
    tb(s,x+0.20,3.95,2.55,0.40,hint,sz=14,b=True,c=cl,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.65,"我的 ___ 包, 卖 $ ___ , 给 ___ 用!","My ___ Bag, $___, for ___!",accent=GREEN)
n+=1; pn(s,n)
notes(s,"🌱 APPLY · 10 分钟 — 准备推销:\n• 3 分钟: 各队讨论 3 个问题, 写在 booklet\n• 5 分钟: 练习 30 秒推销 (轮流说 — 每个队员试一次)\n• 2 分钟: 选「最佳推销员」上台 (rotation)\n• 价格提示: $5–$50 之间\n• 名字越 catchy 越好 (押韵, 玩字, 短)")

# ============================================================
# === SESSION 3 · PHASE 5: SHARE & CLOSE (5 min) ===
# ============================================================
s=ns(); bg(s,CREAM); hb(s,"🛒 Mini-Fair + Day 3 徽章!  Mini-Fair + Badge!",PH_CLOSE)
score_badge(s)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.85),Inches(9.2),Inches(1.65))
sh.fill.solid(); sh.fill.fore_color.rgb=PH_CLOSE; sh.line.color.rgb=DARK; sh.line.width=Pt(2.5)
tb(s,0.4,1.00,9.2,0.45,"🛒 Mini Fair! 4 队 30 秒推销",sz=20,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.50,9.2,0.30,"Each team — 30-sec pitch in front of class!",sz=11,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.85,9.2,0.45,"💰 全班投票 — 最想买的 +5 分!",sz=18,b=True,c=BIZ,a=PP_ALIGN.CENTER)
tb(s,0.4,2.30,9.2,0.20,"Class votes — the one they'd most buy gets +5!",sz=10,c=DARK,a=PP_ALIGN.CENTER)
sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.5),Inches(2.65),Inches(3),Inches(2.55))
sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=IDEA; sh.line.width=Pt(6)
tf=tb(s,3.6,2.85,2.8,2.4,"DAY 3",sz=18,b=True,c=IDEA,a=PP_ALIGN.CENTER)
ap(tf,"💡🛍️",sz=42,a=PP_ALIGN.CENTER)
ap(tf,"小小企业家",sz=15,b=True,c=BIZ,a=PP_ALIGN.CENTER)
ap(tf,"✓ COMPLETED",sz=11,b=True,c=OK,a=PP_ALIGN.CENTER)
ap(tf,"🍎🏰🛒🍫",sz=18,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"🎤 CLOSE · 5 分钟:\n• 2 分钟: 4 队 30 秒推销 — 队代表上台 (rotation)\n• 1 分钟: 全班投票 — 「你最想买哪个?」 (举手数最多 → 队 +5)\n• 1 分钟: 公布全天总分 → 冠军队上台\n• 1 分钟: 发徽章 + tease Day 4 — 「明天 — 帮助者!」\n• 全班齐喊: 「发现! 创造! 帮人! 我也是 entrepreneur!」")

# ============================================================
out=os.path.join(os.path.dirname(__file__),"day3_entrepreneurs.pptx")
prs.save(out); print(f"Saved {out}  ({n} slides)")
