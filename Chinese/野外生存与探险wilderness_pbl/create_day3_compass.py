#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
野外生存与探险 Day 3 — 方向与地图 (罗盘挑战)
Direction & Map (Compass Challenge)
"Explorer Mission" framing — kids learn N/S/E/W, sun-method, build a compass,
draw a camp map, and run a treasure hunt.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(5.625)
W, H = prs.slide_width, prs.slide_height

# --- Palette (continuity + compass accents) ---
PINE   = RGBColor(0x1E,0x4D,0x2B)
SUN    = RGBColor(0xE0,0x7A,0x2C)
CREAM  = RGBColor(0xFD,0xF6,0xE3)
BROWN  = RGBColor(0x6B,0x44,0x23)
SKY    = RGBColor(0x4A,0x90,0xD9)
SUNYEL = RGBColor(0xF5,0xC2,0x42)
ALERT  = RGBColor(0xD0,0x4A,0x3C)
WHITE  = RGBColor(0xFF,0xFF,0xFF)
DARK   = RGBColor(0x2C,0x2C,0x2C)
GRAY   = RGBColor(0x88,0x88,0x88)
LGRAY  = RGBColor(0xBB,0xBB,0xBB)
WARM   = RGBColor(0xFF,0xF3,0xE0)
IMGBG  = RGBColor(0xE8,0xE8,0xE8)
OK     = RGBColor(0x38,0x8E,0x3C)
# Compass-themed
NAVY    = RGBColor(0x1A,0x23,0x7E)   # navigator's blue (primary for D3)
COMPRED = RGBColor(0xC6,0x28,0x28)   # compass north red
GOLD    = RGBColor(0xF9,0xA8,0x25)
N_CL    = COMPRED
S_CL    = NAVY
E_CL    = SUN
W_CL    = OK

# === Helpers ===
def ns(): return prs.slides.add_slide(prs.slide_layouts[6])
def tb(s,l,t,w,h,txt,sz=18,b=False,c=DARK,a=None):
    bx=s.shapes.add_textbox(Inches(l),Inches(t),Inches(w),Inches(h));tf=bx.text_frame;tf.word_wrap=True;p=tf.paragraphs[0]
    if a:p.alignment=a
    r=p.add_run();r.text=txt;r.font.size=Pt(sz);r.font.bold=b;r.font.color.rgb=c;r.font.name='KaiTi';return tf
def ap(tf,txt,sz=18,b=False,c=DARK,a=None):
    p=tf.add_paragraph()
    if a:p.alignment=a
    r=p.add_run();r.text=txt;r.font.size=Pt(sz);r.font.bold=b;r.font.color.rgb=c;r.font.name='KaiTi'
def bg(s,c):
    sh=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,W,H);sh.fill.solid();sh.fill.fore_color.rgb=c;sh.line.fill.background()
    sp=sh._element;sp.getparent().remove(sp);s.shapes._spTree.insert(2,sp)
def ib(s,l,t,w,h,lb="📷"):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(l),Inches(t),Inches(w),Inches(h));sh.fill.solid();sh.fill.fore_color.rgb=IMGBG;sh.line.fill.background()
    tb(s,l+0.1,t+h/2-0.2,w-0.2,0.4,lb,sz=14,c=LGRAY,a=PP_ALIGN.CENTER)
def hb(s,txt,c=NAVY,t=0.15):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(t),Inches(9.4),Inches(0.55));sh.fill.solid();sh.fill.fore_color.rgb=c;sh.line.fill.background()
    tb(s,0.4,t+0.03,9.2,0.5,txt,sz=20,b=True,c=WHITE)
def pn(s,n): tb(s,9.0,5.25,0.8,0.3,str(n),sz=10,c=GRAY,a=PP_ALIGN.RIGHT)
def notes(s,txt):
    s.notes_slide.notes_text_frame.text=txt
def div(title,sub,color,emoji=""):
    s=ns();bg(s,color)
    tb(s,1,1.5,8,1.2,f"{emoji} {title}",sz=42,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,1,2.8,8,0.8,sub,sz=22,c=WHITE,a=PP_ALIGN.CENTER);return s
def pill(s,l,t,w,h,txt,c,sz=14):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(l),Inches(t),Inches(w),Inches(h))
    sh.fill.solid();sh.fill.fore_color.rgb=c;sh.line.fill.background()
    tb(s,l+0.1,t+h/2-0.2,w-0.2,0.4,txt,sz=sz,b=True,c=WHITE,a=PP_ALIGN.CENTER)

def mission_card(s,l,t,w,h,num,task_cn,task_en,emoji,color):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(l),Inches(t),Inches(w),Inches(h))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=color;sh.line.width=Pt(2.5)
    badge=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(l+0.1),Inches(t+0.08),Inches(0.55),Inches(0.55))
    badge.fill.solid();badge.fill.fore_color.rgb=color;badge.line.fill.background()
    tb(s,l+0.1,t+0.18,0.55,0.4,str(num),sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,l+0.05,t+h-0.65,w-0.1,0.35,emoji,sz=24,a=PP_ALIGN.CENTER)
    tb(s,l+0.05,t+h-0.32,w-0.1,0.3,task_cn,sz=12,b=True,c=color,a=PP_ALIGN.CENTER)
    tb(s,l+0.7,t+0.15,w-0.8,0.4,task_en,sz=10,c=GRAY)

def sentence_frame_bar(s,t,frame_cn,frame_en):
    sf=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(t),Inches(9.4),Inches(0.65))
    sf.fill.solid();sf.fill.fore_color.rgb=WARM;sf.line.color.rgb=SUN;sf.line.width=Pt(2)
    tb(s,0.5,t+0.1,1.7,0.4,"💬 我来说",sz=14,b=True,c=SUN)
    tb(s,2.0,t+0.07,7.6,0.3,frame_cn,sz=14,b=True,c=DARK)
    tb(s,2.0,t+0.32,7.6,0.3,frame_en,sz=10,c=GRAY)

def compass_rose(s,cx,cy,radius,size_label=14):
    """Draw a compass-rose (NSEW labels, with N in red)."""
    # Outer circle
    sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(cx-radius),Inches(cy-radius),Inches(radius*2),Inches(radius*2))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=NAVY;sh.line.width=Pt(2.5)
    # Inner circle
    sh2=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(cx-radius*0.3),Inches(cy-radius*0.3),Inches(radius*0.6),Inches(radius*0.6))
    sh2.fill.solid();sh2.fill.fore_color.rgb=GOLD;sh2.line.color.rgb=NAVY;sh2.line.width=Pt(1.5)
    # Cross lines
    ln1=s.shapes.add_connector(1,Inches(cx-radius*0.95),Inches(cy),Inches(cx+radius*0.95),Inches(cy))
    ln1.line.color.rgb=LGRAY;ln1.line.width=Pt(0.8)
    ln2=s.shapes.add_connector(1,Inches(cx),Inches(cy-radius*0.95),Inches(cx),Inches(cy+radius*0.95))
    ln2.line.color.rgb=LGRAY;ln2.line.width=Pt(0.8)
    # NSEW labels
    tb(s,cx-0.3,cy-radius-0.45,0.6,0.4,"北 N",sz=size_label,b=True,c=COMPRED,a=PP_ALIGN.CENTER)
    tb(s,cx-0.3,cy+radius+0.05,0.6,0.4,"南 S",sz=size_label,b=True,c=NAVY,a=PP_ALIGN.CENTER)
    tb(s,cx+radius+0.05,cy-0.2,0.7,0.4,"东 E",sz=size_label,b=True,c=SUN,a=PP_ALIGN.CENTER)
    tb(s,cx-radius-0.75,cy-0.2,0.7,0.4,"西 W",sz=size_label,b=True,c=OK,a=PP_ALIGN.CENTER)

def direction_card(s,l,t,w,h,em,cn,en,abbr,body_pos,color):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(l),Inches(t),Inches(w),Inches(h))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=color;sh.line.width=Pt(3)
    pill(s,l+0.15,t+0.15,0.6,0.4,abbr,color,sz=14)
    tb(s,l+0.1,t+0.6,w-0.2,0.7,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,l+0.1,t+1.5,w-0.2,0.5,cn,sz=24,b=True,c=color,a=PP_ALIGN.CENTER)
    tb(s,l+0.1,t+2.05,w-0.2,0.4,en,sz=14,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,l+0.1,t+h-0.55,w-0.2,0.4,f"≈ 身体的 {body_pos}",sz=12,b=True,c=DARK,a=PP_ALIGN.CENTER)

def video_slide(title_cn,title_en,before_task,after_action,bgc=NAVY):
    s=ns();bg(s,bgc)
    tb(s,1,0.55,8,0.7,"🎬 看视频 Watch Video",sz=32,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,1,1.25,8,0.4,title_cn,sz=20,b=True,c=WARM,a=PP_ALIGN.CENTER)
    tb(s,1,1.6,8,0.3,title_en,sz=12,c=WARM,a=PP_ALIGN.CENTER)
    pre=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(2.05),Inches(4.5),Inches(2.0))
    pre.fill.solid();pre.fill.fore_color.rgb=WHITE;pre.line.fill.background()
    tb(s,0.55,2.15,4.3,0.4,"👂 看之前 Before Watching",sz=14,b=True,c=bgc)
    tb(s,0.55,2.55,4.3,1.4,before_task,sz=12,c=DARK)
    post=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.1),Inches(2.05),Inches(4.5),Inches(2.0))
    post.fill.solid();post.fill.fore_color.rgb=WHITE;post.line.fill.background()
    tb(s,5.25,2.15,4.3,0.4,"🎯 看完后 After Watching",sz=14,b=True,c=SUN)
    tb(s,5.25,2.55,4.3,1.4,after_action,sz=12,c=DARK)
    tb(s,1,4.3,8,0.4,"🔗 视频链接 (老师粘贴) / Teacher pastes video link",sz=12,c=WARM,a=PP_ALIGN.CENTER)
    return s

def ab_slide(title_cn,title_en,question_cn,question_en,a_emoji,a_label,a_caption,b_emoji,b_label,b_caption,answer,reason):
    s=ns();bg(s,CREAM);hb(s,f"🧭 {title_cn}  {title_en}",NAVY)
    tb(s,0.4,0.85,9.2,0.4,question_cn,sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
    tb(s,0.4,1.20,9.2,0.3,question_en,sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    a_box=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.6),Inches(4.5),Inches(2.4))
    a_box.fill.solid();a_box.fill.fore_color.rgb=WHITE;a_box.line.color.rgb=NAVY;a_box.line.width=Pt(2.5)
    pill(s,0.5,1.7,0.7,0.4,"A",NAVY,sz=16)
    tb(s,1.2,1.65,3.6,0.5,a_label,sz=18,b=True,c=NAVY)
    tb(s,0.6,2.2,4.2,0.5,a_emoji,sz=44,a=PP_ALIGN.CENTER)
    tb(s,0.6,3.1,4.2,0.6,a_caption,sz=13,c=DARK,a=PP_ALIGN.CENTER)
    b_box=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.1),Inches(1.6),Inches(4.5),Inches(2.4))
    b_box.fill.solid();b_box.fill.fore_color.rgb=WHITE;b_box.line.color.rgb=ALERT;b_box.line.width=Pt(2.5)
    pill(s,5.2,1.7,0.7,0.4,"B",ALERT,sz=16)
    tb(s,5.9,1.65,3.6,0.5,b_label,sz=18,b=True,c=ALERT)
    tb(s,5.3,2.2,4.2,0.5,b_emoji,sz=44,a=PP_ALIGN.CENTER)
    tb(s,5.3,3.1,4.2,0.6,b_caption,sz=13,c=DARK,a=PP_ALIGN.CENTER)
    sentence_frame_bar(s,4.15,
        "我选 ___, 因为 ____",
        "I choose A/B because…")
    tb(s,0.4,4.92,9.2,0.3,"👉 走到 A 边 / B 边 — 用一句话说为什么。",sz=12,b=True,c=SUN,a=PP_ALIGN.CENTER)
    notes(s,f"老师备课:\n• 答案: {answer}\n• 原因: {reason}")
    return s

def word_card_read(w,py,en,sent,img,color=SUN):
    s=ns();bg(s,CREAM);hb(s,"👀 我会认  I Can Read",color)
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.0),Inches(4.5),Inches(2.5))
    sh.fill.solid();sh.fill.fore_color.rgb=WARM;sh.line.fill.background()
    tb(s,0.5,1.1,4.3,1.4,w,sz=72,b=True,c=NAVY,a=PP_ALIGN.CENTER)
    tb(s,0.5,2.4,4.3,0.4,f"{py}  {en}",sz=20,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,0.5,2.85,4.3,0.4,"👉 跟我读！Read after me!",sz=14,c=color,a=PP_ALIGN.CENTER)
    ib(s,5.3,1.0,4.4,2.5,img)
    sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(3.8),Inches(9.2),Inches(1.2))
    sh2.fill.solid();sh2.fill.fore_color.rgb=WHITE;sh2.line.color.rgb=color;sh2.line.width=Pt(2)
    tb(s,0.6,3.9,1.5,0.4,"例句 Example",sz=14,b=True,c=color)
    tb(s,0.6,4.3,8.8,0.5,sent,sz=20,b=True,c=DARK)
    return s

def word_card_write(w,py,en,strokes_hint,color=NAVY):
    s=ns();bg(s,CREAM);hb(s,"✍️ 我会写  I Can Write",color)
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.0),Inches(4.5),Inches(2.0))
    sh.fill.solid();sh.fill.fore_color.rgb=WARM;sh.line.fill.background()
    tb(s,0.5,1.05,4.3,1.2,w,sz=72,b=True,c=color,a=PP_ALIGN.CENTER)
    tb(s,0.5,2.2,4.3,0.4,f"{py}  {en}",sz=20,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,0.6,3.4,4.6,0.4,"📝 笔顺 Stroke Order",sz=16,b=True,c=color)
    tb(s,0.6,3.8,4.6,1.2,strokes_hint,sz=14,c=DARK)
    tb(s,5.8,3.4,3.8,0.4,"练习步骤 Practice:",sz=14,b=True,c=color)
    tb(s,5.8,3.8,3.8,1.2,"1. 看老师写\n2. 用手指空中写\n3. 在本子上写 3 次",sz=13,c=DARK)
    for i in range(4):
        sq=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(5.8+i*0.95),Inches(1.0),Inches(0.85),Inches(0.85))
        sq.fill.solid();sq.fill.fore_color.rgb=WHITE;sq.line.color.rgb=color;sq.line.width=Pt(1.5)
        ln1=s.shapes.add_connector(1,Inches(5.8+i*0.95),Inches(1.425),Inches(5.8+i*0.95+0.85),Inches(1.425))
        ln1.line.color.rgb=LGRAY;ln1.line.width=Pt(0.5);ln1.line.dash_style=2
        ln2=s.shapes.add_connector(1,Inches(5.8+i*0.95+0.425),Inches(1.0),Inches(5.8+i*0.95+0.425),Inches(1.85))
        ln2.line.color.rgb=LGRAY;ln2.line.width=Pt(0.5);ln2.line.dash_style=2
    tb(s,5.8,1.95,3.8,0.3,"在田字格里写 3 次 ↓",sz=11,c=GRAY)
    return s

# ========================================================================
#                              SLIDES
# ========================================================================
n=0

# 1. COVER
s=ns();bg(s,NAVY)
sh=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,Inches(2.4),W,Inches(2.0))
sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.fill.background()
tb(s,1,0.4,8,0.5,"DAY 3",sz=18,b=True,c=GOLD,a=PP_ALIGN.CENTER)
tb(s,1,0.95,8,0.7,"🧭 方向与地图",sz=42,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,1,1.7,8,0.5,"Direction & Map  ·  罗盘挑战 Compass Challenge",sz=20,c=WARM,a=PP_ALIGN.CENTER)
tb(s,1,2.6,8,0.5,"🧭 探险家任务  Explorer Mission",sz=24,b=True,c=NAVY,a=PP_ALIGN.CENTER)
tb(s,1,3.15,8,0.4,"Find · Use Sun · Make Compass · Map",sz=14,c=BROWN,a=PP_ALIGN.CENTER)
tb(s,1,3.55,8,0.4,"认方向 · 用太阳 · 做罗盘 · 画地图",sz=14,b=True,c=BROWN,a=PP_ALIGN.CENTER)
tb(s,1,4.6,8,0.4,"野外生存与探险 · Wilderness Survival",sz=14,b=True,c=GOLD,a=PP_ALIGN.CENTER)
n+=1;pn(s,n)
notes(s,"开场 (1 分钟):\n• 「探险家们! 今天你们会迷路吗? 不会! 因为我们要学会用方向。」\n• 4 个任务: 认方向 / 用太阳 / 做罗盘 / 画地图。\n• 完成所有任务可以拿到「方向大师」徽章。")

# 2. EXPLORER MISSION — 4 tasks
s=ns();bg(s,CREAM);hb(s,"🧭 今天的任务  Today's Mission",NAVY)
tb(s,0.4,0.9,9.2,0.4,"小探险家们 — 我们要一起做 4 件事!",sz=20,b=True,c=NAVY,a=PP_ALIGN.CENTER)
tb(s,0.4,1.32,9.2,0.3,"Little explorers — 4 jobs today!",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
mission_card(s,0.4,1.85,2.25,2.6,1,"认方向","Find directions","🧭",NAVY)
mission_card(s,2.75,1.85,2.25,2.6,2,"用太阳","Use the sun","☀️",SUN)
mission_card(s,5.10,1.85,2.25,2.6,3,"做罗盘","Make compass","🔧",COMPRED)
mission_card(s,7.45,1.85,2.25,2.6,4,"画地图","Draw a map","🗺️",OK)
sentence_frame_bar(s,4.65,
    "我是小探险家, 我可以…",
    "I am a little explorer, I can…")
n+=1;pn(s,n)
notes(s,"4 个任务一起说一遍, 让学生跟读「认方向 - 用太阳 - 做罗盘 - 画地图」。完成 1 个打 1 个 ✓。")

# 3. EXPLORER PLEDGE
s=ns();bg(s,CREAM);hb(s,"🤝 探险家宣言  Explorer Pledge",NAVY)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(1.0),Inches(1.1),Inches(8.0),Inches(3.2))
sh.fill.solid();sh.fill.fore_color.rgb=WARM;sh.line.color.rgb=NAVY;sh.line.width=Pt(2.5)
tb(s,1.2,1.3,7.6,0.5,"🧭 我会认 4 个方向: 东 南 西 北。",sz=20,b=True,c=NAVY,a=PP_ALIGN.CENTER)
tb(s,1.2,1.85,7.6,0.4,"I know 4 directions: E S W N.",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,1.2,2.35,7.6,0.5,"☀️ 我能用太阳找方向。",sz=20,b=True,c=SUN,a=PP_ALIGN.CENTER)
tb(s,1.2,2.85,7.6,0.4,"I can find directions with the sun.",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,1.2,3.35,7.6,0.5,"🗺️ 我不会迷路。",sz=20,b=True,c=COMPRED,a=PP_ALIGN.CENTER)
tb(s,1.2,3.85,7.6,0.4,"I will not get lost.",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,0.4,4.5,9.2,0.5,"👉 一起举手念! Raise your hand and say it!",sz=14,b=True,c=SUN,a=PP_ALIGN.CENTER)
n+=1;pn(s,n)

# 4. SESSION 1 DIVIDER
s=div("Session 1  上午","🧭 方向 + 罗盘 + 教室方向游戏  Direction · Compass · Classroom",NAVY,"📖");n+=1;pn(s,n)

# 5. HOOK — pre-watching
s=ns();bg(s,CREAM);hb(s,"🤔 想一想  Imagine This!",SUN)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.95),Inches(9.2),Inches(1.6))
sh.fill.solid();sh.fill.fore_color.rgb=WARM;sh.line.color.rgb=SUN;sh.line.width=Pt(2.5)
tb(s,0.6,1.1,8.8,0.6,"假装你迷路了 — 怎么找到回家的路?",sz=24,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.6,1.7,8.8,0.4,"Pretend you're lost — how do you find your way home?",sz=14,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,0.6,2.10,8.8,0.4,"👉 跟同桌说一说 / Tell your partner",sz=12,b=True,c=SUN,a=PP_ALIGN.CENTER)
# 4 hint cards
hints=[("🧭","用罗盘 / Compass"),("☀️","看太阳 / Sun"),("⭐","看星星 / Stars"),("🗺️","看地图 / Map")]
for i,(em,t) in enumerate(hints):
    x=0.4+i*2.35
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(2.85),Inches(2.20),Inches(1.7))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=NAVY;sh.line.width=Pt(2)
    tb(s,x+0.05,2.95,2.1,0.7,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.85,2.1,0.5,t,sz=13,b=True,c=NAVY,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.7,
    "我可以用 ___ 找路。",
    "I can use ___ to find my way.")
n+=1;pn(s,n)
notes(s,"打开课程 (3 分钟):\n• 让学生 think-pair-share。\n• 收集答案 — 板书。\n• 「今天我们要学最重要的: 方向!」")

# 6. VIDEO — North/South/East/West intro
s=video_slide(
    "🧭 方向 4 兄弟  4 Directions for Kids",
    "North · South · East · West",
    "👂 听 / 看:\n1. 一共有几个方向?\n2. 哪个在「上面」?\n3. 太阳从哪边升起?",
    "🎯 看完后:\n• 用手指: 上 北 / 下 南 / 右 东 / 左 西\n• 喊: NSEW!\n• 说: 太阳从 ___ 升起",
    bgc=NAVY);n+=1;pn(s,n)
notes(s,"视频建议 (1-3 分钟):\n• YouTube 搜: 「North South East West for kids song」 (有动画版)\n• 或: 「Compass directions for kids」\n• 关键概念: 上北下南左西右东; 太阳从东边升起")

# 7. 4 DIRECTIONS — Compass Rose intro
s=ns();bg(s,CREAM);hb(s,"🧭 4 个基本方向  4 Basic Directions",NAVY)
tb(s,0.4,0.9,9.2,0.4,"上北 · 下南 · 左西 · 右东",sz=22,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.3,"Up = North · Down = South · Left = West · Right = East",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# Compass rose (left)
compass_rose(s,2.5,3.5,1.4,size_label=15)
# Direction cards (right side, 2x2)
specs=[("🔝","北","North","N","头",N_CL,4.6,1.8),
       ("⬇️","南","South","S","脚",S_CL,7.3,1.8),
       ("➡️","东","East","E","右",E_CL,4.6,3.7),
       ("⬅️","西","West","W","左",W_CL,7.3,3.7)]
for em,cn,en,abbr,bp,cl,x,y in specs:
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(2.3),Inches(1.55))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=cl;sh.line.width=Pt(2.5)
    tb(s,x+0.1,y+0.1,2.1,0.55,f"{em} {abbr}",sz=20,b=True,c=cl)
    tb(s,x+0.1,y+0.65,2.1,0.4,cn,sz=22,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,y+1.05,2.1,0.3,en,sz=12,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,y+1.30,2.1,0.25,f"≈ 身体的 {bp}",sz=10,b=True,c=DARK,a=PP_ALIGN.CENTER)
n+=1;pn(s,n)
notes(s,"重点: \n• 「上北下南」是地图惯例 (但实际方向不一定指向「上」)\n• 红色 N 是国际通用 (compass needle red end)\n• 让学生用手指在空中画 + 字 (东南西北)")

# 8. MNEMONIC — body to map bridge
s=ns();bg(s,CREAM);hb(s,"🧠 记一记  Mnemonic",SUN)
tb(s,0.4,0.85,9.2,0.4,"中文记法 + 英文记法",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
# Chinese mnemonic
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.4),Inches(4.5),Inches(3.0))
sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=COMPRED;sh.line.width=Pt(2.5)
pill(s,0.55,1.55,1.4,0.4,"🇨🇳 中文",COMPRED,sz=13)
tb(s,0.55,2.0,4.2,0.5,"上 北",sz=24,b=True,c=COMPRED)
tb(s,0.55,2.5,4.2,0.5,"下 南",sz=24,b=True,c=NAVY)
tb(s,0.55,3.0,4.2,0.5,"左 西",sz=24,b=True,c=OK)
tb(s,0.55,3.5,4.2,0.5,"右 东",sz=24,b=True,c=SUN)
tb(s,0.55,4.0,4.2,0.4,"💡 一起念 3 次!",sz=12,c=DARK)
# English mnemonic
sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.1),Inches(1.4),Inches(4.5),Inches(3.0))
sh2.fill.solid();sh2.fill.fore_color.rgb=WHITE;sh2.line.color.rgb=NAVY;sh2.line.width=Pt(2.5)
pill(s,5.25,1.55,1.4,0.4,"🇺🇸 English",NAVY,sz=13)
tb(s,5.25,2.0,4.2,0.5,"Never  =  N (North)",sz=18,b=True,c=COMPRED)
tb(s,5.25,2.5,4.2,0.5,"Eat  =  E (East)",sz=18,b=True,c=SUN)
tb(s,5.25,3.0,4.2,0.5,"Soggy  =  S (South)",sz=18,b=True,c=NAVY)
tb(s,5.25,3.5,4.2,0.5,"Waffles  =  W (West)",sz=18,b=True,c=OK)
tb(s,5.25,4.0,4.2,0.4,"💡 Clockwise from N!",sz=12,c=DARK)
sentence_frame_bar(s,4.6,
    "上北 ___ 南 , ___ 西 ___ 东 。",
    "Up north, down south, left west, right east.")
n+=1;pn(s,n)

# 9. BODY DIRECTIONS — front/back/left/right → NSEW
s=ns();bg(s,CREAM);hb(s,"🚶 前后左右 → 东南西北  Body → Map",E_CL)
tb(s,0.4,0.9,9.2,0.4,"如果你面向北边 (North) 站着…",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.3,"If you stand facing NORTH…",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# 4 boxes
mappings=[("👆 前","面前","Front","→","🔝 北","North",N_CL,1.0),
          ("👇 后","身后","Back", "→","⬇️ 南","South",S_CL,2.0),
          ("👉 右","右手","Right","→","➡️ 东","East",E_CL,3.0),
          ("👈 左","左手","Left", "→","⬅️ 西","West",W_CL,4.0)]
for i,(em_l,bp_cn,bp_en,arrow,em_r,nsew_en,cl,y) in enumerate(mappings):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(y+0.7),Inches(9.2),Inches(0.85))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=cl;sh.line.width=Pt(2)
    tb(s,0.6,y+0.85,2.0,0.5,em_l,sz=22,b=True,c=cl)
    tb(s,2.4,y+0.78,1.8,0.4,bp_cn,sz=15,b=True,c=DARK)
    tb(s,2.4,y+1.13,1.8,0.3,bp_en,sz=10,c=GRAY)
    tb(s,4.4,y+0.85,0.8,0.5,arrow,sz=22,b=True,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,5.5,y+0.85,2.0,0.5,em_r,sz=22,b=True,c=cl)
    tb(s,7.6,y+0.85,2.0,0.5,nsew_en,sz=15,b=True,c=DARK)
n+=1;pn(s,n)
notes(s,"老师演示:\n• 「站起来, 面向北 (墙上贴 N 标签)。」\n• 学生:  「我前面是北, 后面是南, 右边是东, 左边是西。」\n• 转一圈 — 如果面向南, 重做一遍! (东西颠倒)")

# 10. SUN METHOD — concept
s=ns();bg(s,CREAM);hb(s,"☀️ 用太阳找方向  Use the Sun",SUN)
tb(s,0.4,0.9,9.2,0.4,"太阳早上从「东」升起, 下午从「西」落下!",sz=20,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.3,"Sun rises in the EAST, sets in the WEST!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# 3 time-of-day cards
times=[("🌅","早上","Morning","太阳在 ☀️ → 东",E_CL),
       ("☀️","中午","Noon", "太阳在头顶 ↑ 南 (北半球)",NAVY),
       ("🌇","下午","Evening","太阳在 ☀️ → 西",W_CL)]
for i,(em,cn,en,desc,cl) in enumerate(times):
    x=0.4+i*3.15
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.85),Inches(3.0),Inches(2.4))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=cl;sh.line.width=Pt(2.5)
    tb(s,x+0.1,1.95,2.8,0.7,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,2.75,2.8,0.5,cn,sz=22,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,3.30,2.8,0.4,en,sz=12,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,3.75,2.8,0.4,desc,sz=12,b=True,c=DARK,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.45,
    "现在是 ___, 太阳在 ___, 那边就是 ___ 。",
    "It's ___, the sun is in the ___, so that's ___.")
n+=1;pn(s,n)
notes(s,"K-5 简化版 (足够准确):\n• 早上 → 太阳指向东\n• 下午 → 太阳指向西\n• 中午太阳在「上面」(北半球偏南)\n• 注意: 实际只在春分/秋分时才正东正西, 但 K-5 不需要细讲。")

# 11. SUN METHOD A/B
s=ab_slide("用太阳找方向  Use the Sun","",
    "现在是早上, 太阳在你的左手边 — 你面向哪里?",
    "Morning. Sun on your LEFT. Which direction are you facing?",
    "🔝","面向北 N","左手是西 → 不对!",
    "⬇️","面向南 S","左手是东 ☀️ → 对!",
    "B 面向南",
    "如果太阳 (东) 在左手, 那么右手是西, 前面是南, 后面是北。"
)
n+=1;pn(s,n)

# 12. CLASSROOM SIMULATION — set up corners
s=ns();bg(s,CREAM);hb(s,"🏫 教室方向  Classroom Directions",NAVY)
tb(s,0.4,0.9,9.2,0.4,"老师在 4 面墙上贴 4 张方向标签!",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.3,"Teacher tapes N/S/E/W on 4 walls.",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# Classroom diagram (rectangle with 4 wall labels)
rm=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(2.5),Inches(2.0),Inches(5.0),Inches(2.5))
rm.fill.solid();rm.fill.fore_color.rgb=WHITE;rm.line.color.rgb=DARK;rm.line.width=Pt(2.5)
tb(s,3.0,2.20,4.0,0.4,"教 室  Classroom",sz=14,b=True,c=GRAY,a=PP_ALIGN.CENTER)
# Wall labels
tb(s,2.5,1.55,5.0,0.4,"🔝 北 N",sz=18,b=True,c=N_CL,a=PP_ALIGN.CENTER)
tb(s,2.5,4.55,5.0,0.4,"⬇️ 南 S",sz=18,b=True,c=S_CL,a=PP_ALIGN.CENTER)
tb(s,7.55,3.0,1.5,0.4,"➡️ 东 E",sz=18,b=True,c=E_CL,a=PP_ALIGN.CENTER)
tb(s,1.0,3.0,1.5,0.4,"⬅️ 西 W",sz=18,b=True,c=W_CL,a=PP_ALIGN.CENTER)
# Rules
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(4.95),Inches(9.2),Inches(0.55))
sh.fill.solid();sh.fill.fore_color.rgb=WARM;sh.line.color.rgb=SUN;sh.line.width=Pt(2)
tb(s,0.5,5.02,9.0,0.4,"🎯 规则: 老师喊「东」 → 学生跑到「东」的墙边! Slow ones sit out.",sz=12,b=True,c=PINE,a=PP_ALIGN.CENTER)
n+=1;pn(s,n)
notes(s,"老师准备 (零 prep):\n• 用 4 张 A4 纸, 写「东 East」「南 South」「西 West」「北 North」, 贴在 4 面墙。\n• 老师喊指令, 学生跑。可加: 「向北 + 向东」让学生站在两墙之间的「东北」角。\n• 玩 5-8 轮即可。")

# 13. SIMON SAYS — direction edition
s=ns();bg(s,CREAM);hb(s,"🎮 方向 Simon Says!",E_CL)
tb(s,0.4,0.9,9.2,0.4,"老师说: 「Simon says — 向北!」 → 学生面向北",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.3,"Teacher: \"Simon says — face NORTH!\" → Students face north",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# Examples
ex=[("👉","Simon says — 向东转!","Face East!"),
    ("👉","Simon says — 走 3 步向南!","3 steps south!"),
    ("👉","向西跳!","Jump west! (no Simon → don't move!)"),
    ("👉","Simon says — 找最北的同学!","Find the most-north classmate!")]
for i,(em,cn,en) in enumerate(ex):
    y=1.85+i*0.7
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(y),Inches(9.2),Inches(0.6))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=NAVY;sh.line.width=Pt(1.5)
    tb(s,0.55,y+0.13,0.5,0.4,em,sz=18)
    tb(s,1.05,y+0.05,5.5,0.3,cn,sz=14,b=True,c=DARK)
    tb(s,1.05,y+0.32,5.5,0.3,en,sz=10,c=GRAY)
    if i==2:
        tb(s,7.0,y+0.13,2.0,0.4,"⚠️ Trap!",sz=12,b=True,c=ALERT,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.85,
    "我向 ___ 转了 ___ 步。",
    "I turned ___ steps to the ___.")
n+=1;pn(s,n)

# 14. COMPASS — what is it
s=ns();bg(s,CREAM);hb(s,"🧭 罗盘 / 指南针 是什么？  What's a Compass?",COMPRED)
ib(s,0.3,1.0,4.4,3.6,"📷 罗盘 / 指南针照片  Compass photo")
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(1.0),Inches(4.8),Inches(3.6))
sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=COMPRED;sh.line.width=Pt(2.5)
head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(1.0),Inches(4.8),Inches(0.5))
head.fill.solid();head.fill.fore_color.rgb=COMPRED;head.line.fill.background()
tb(s,5.05,1.08,4.6,0.4,"🤔 罗盘的秘密  Compass Secrets",sz=15,b=True,c=WHITE)
tf=tb(s,5.1,1.65,4.55,0.4,"·  红色的针 = 指 北 N",sz=13,b=True,c=DARK)
ap(tf,"   Red end of needle = points NORTH",sz=10,c=GRAY)
ap(tf," ",sz=8)
ap(tf,"·  里面有 磁铁 (磁针)",sz=13,b=True,c=DARK)
ap(tf,"   Inside: a magnetized needle",sz=10,c=GRAY)
ap(tf," ",sz=8)
ap(tf,"·  地球本身就像一个大磁铁",sz=13,b=True,c=DARK)
ap(tf,"   Earth is a giant magnet",sz=10,c=GRAY)
ap(tf," ",sz=8)
ap(tf,"·  所以红针总是指向「北」",sz=13,b=True,c=DARK)
ap(tf,"   So the red end always points NORTH",sz=10,c=GRAY)
sentence_frame_bar(s,4.75,
    "罗盘的红针指向 ___ 。",
    "The red needle points to ___.")
n+=1;pn(s,n)
notes(s,"K-5 简化版:\n• 不需要细讲磁场/磁极, 只要知道「红针 = 北」即可。\n• 可以让学生看真实的小罗盘 (老师带几个)。\n• 「红色 = 北 = 危险/重要」是国际通用 (像红绿灯)")

# 15. HOW TO READ a compass
s=ns();bg(s,CREAM);hb(s,"📍 怎么用罗盘？  How to Use a Compass",COMPRED)
tb(s,0.4,0.9,9.2,0.4,"3 步使用法  3 Steps",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
steps=[("1️⃣","平拿着","Hold flat","罗盘要平 — 不要斜!\nKeep flat — not tilted",NAVY),
       ("2️⃣","看红针","Watch the red","等红针停下来\nWait for red to settle",COMPRED),
       ("3️⃣","转身","Turn body","把身体转到「N」对着红针\nTurn so N aligns with red",E_CL)]
for i,(num,cn,en,desc,cl) in enumerate(steps):
    x=0.4+i*3.15
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.4),Inches(3.0),Inches(3.0))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=cl;sh.line.width=Pt(2.5)
    tb(s,x+0.1,1.5,2.8,0.7,num,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,2.3,2.8,0.5,cn,sz=20,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,2.85,2.8,0.4,en,sz=12,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.15,3.4,2.85,0.9,desc,sz=11,c=DARK,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.55,
    "我把罗盘 ___, 看见 ___, 然后 ___ 。",
    "I hold flat, see red, and turn.")
n+=1;pn(s,n)
notes(s,"老师准备:\n• 如果有真罗盘 — 让学生轮流用, 找到「北」(指向窗外/某面墙)。\n• 如果没有 — 用打印的纸罗盘示范。")

# 16. PRACTICE A/B — sun + compass
s=ab_slide("综合练习  Practice","",
    "你迷路了, 没有罗盘, 但是中午太阳在头顶 — 哪边是北?",
    "Lost, no compass. Noon — sun is overhead. Where's North?",
    "👤☀️","身体的影子那边","影子指向 北 (北半球)",
    "👤➡️","太阳的反方向","太阳在头顶 — 没有「反方向」",
    "A 影子那边",
    "在北半球的中午, 太阳偏南, 影子指向北。这是简单实用的方法 — K-5 知道这个就够了!"
)
n+=1;pn(s,n)

# 17. SESSION 2 DIVIDER
s=div("Session 2  下午","🔄 复习 + 语言目标 (我会认 6 / 我会写 3)",SUN,"📖");n+=1;pn(s,n)

# 18. QUICK REVIEW — direction quiz
s=ns();bg(s,CREAM);hb(s,"🔄 快速复习  Quick Review",SUN)
tb(s,0.4,0.85,9.2,0.4,"上午学了什么? Let's review!",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
qs=[("Q1","太阳从哪里升起?", "☀️ → 东 East"),
    ("Q2","上北 — 下面是什么?", "南 South"),
    ("Q3","罗盘红针指哪里?",   "北 N"),
    ("Q4","面向北, 右手是哪边?","东 East")]
for i,(num,q,a) in enumerate(qs):
    col=i%2;row=i//2
    x=0.4+col*4.65;y=1.4+row*1.7
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(4.5),Inches(1.5))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=NAVY;sh.line.width=Pt(2)
    pill(s,x+0.15,y+0.15,0.7,0.4,num,NAVY,sz=12)
    tb(s,x+0.95,y+0.1,3.4,0.45,q,sz=14,b=True,c=DARK)
    tb(s,x+0.15,y+0.65,4.2,0.4,"💬 答: ",sz=13,b=True,c=SUN)
    tb(s,x+0.85,y+0.65,3.5,0.4,a,sz=14,b=True,c=COMPRED)
    tb(s,x+0.15,y+1.05,4.2,0.3,"举手抢答! Hand-raise quick answer!",sz=10,c=GRAY)
n+=1;pn(s,n)

# 19. SENTENCE FRAMES
s=ns();bg(s,CREAM);hb(s,"💬 句型卡  Sentence Frames",SUN)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.0),Inches(4.5),Inches(3.6))
sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=SUN;sh.line.width=Pt(2.5)
pill(s,0.6,1.15,1.6,0.4,"K (TK-K)",SUN,sz=14)
tb(s,0.6,1.7,4.1,0.5,"🔝 这是 ___ 。",sz=22,b=True,c=NAVY)
tb(s,0.6,2.2,4.1,0.4,"This is ___. (north/south/east/west)",sz=10,c=GRAY)
tb(s,0.6,2.7,4.1,0.5,"☀️ 太阳在 ___ 。",sz=22,b=True,c=NAVY)
tb(s,0.6,3.2,4.1,0.4,"The sun is in the ___.",sz=10,c=GRAY)
tb(s,0.6,3.75,4.1,0.4,"💡 例: 这是北。太阳在东。",sz=13,c=DARK)
tb(s,0.6,4.1,4.1,0.4,"Ex: This is North. Sun is in the East.",sz=10,c=GRAY)
sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.1),Inches(1.0),Inches(4.5),Inches(3.6))
sh2.fill.solid();sh2.fill.fore_color.rgb=WHITE;sh2.line.color.rgb=NAVY;sh2.line.width=Pt(2.5)
pill(s,5.3,1.15,1.6,0.4,"G1 - G3",NAVY,sz=14)
tb(s,5.3,1.7,4.1,0.5,"🚶 我向 ___ 走 ___ 步。",sz=20,b=True,c=NAVY)
tb(s,5.3,2.2,4.1,0.4,"I walk ___ steps to the ___.",sz=10,c=GRAY)
tb(s,5.3,2.7,4.1,0.5,"🧭 罗盘指向 ___, 因为 ___ 。",sz=18,b=True,c=NAVY)
tb(s,5.3,3.2,4.1,0.4,"The compass points to ___, because ___.",sz=10,c=GRAY)
tb(s,5.3,3.75,4.1,0.4,"💡 例: 我向北走 5 步。罗盘指向北, 因为里面有磁铁。",sz=11,c=DARK)
n+=1;pn(s,n)

# 20-25. WORD CARDS — 我会认 (6 words)
read_data=[
    ("帐篷","zhàng peng","tent",        "我们在营地搭了一个帐篷。",   "📷 帐篷"),
    ("营地","yíng dì",   "campground",  "我们的营地很安全。",         "📷 营地"),
    ("食物","shí wù",    "food",        "我把食物放在背包里。",       "📷 食物"),
    ("水源","shuǐ yuán", "water source","河边有水源。",              "📷 河边水源"),
    ("安全","ān quán",   "safe",        "在大人旁边最安全。",         "📷 营地 + 大人"),
    ("火",   "huǒ",      "fire",        "火堆要远离帐篷。",           "📷 营火"),
]
for w,py,en,sent,img in read_data:
    s=word_card_read(w,py,en,sent,img,color=SUN);n+=1;pn(s,n)

# 26-28. WORD CARDS — 我会写 (3)
write_data=[
    ("食物","shí wù","food",  "9 笔: 食 (人 + 良) + 物 (牛 + 勿)\n吃的东西都叫食物!"),
    ("营地","yíng dì","camp", "8 笔: 营 (草头 + 吕) + 地 (土 + 也)\n小探险家睡觉的地方"),
    ("安全","ān quán","safe", "12 笔: 安 (宀 + 女) + 全 (人 + 王)\n在家里很安全"),
]
for w,py,en,strokes in write_data:
    s=word_card_write(w,py,en,strokes,color=NAVY);n+=1;pn(s,n)

# 29. SESSION 3 DIVIDER
s=div("Session 3  下午","🛠️ 写 Booklet + 3 个项目  Booklet + Treasure · Map · Compass",BROWN,"🎒");n+=1;pn(s,n)

# 30. DAY 3 BOOKLET
s=ns();bg(s,CREAM);hb(s,"📓 完成 Day 3 练习册  Day 3 Booklet",BROWN)
tb(s,0.4,0.85,9.2,0.4,"老师带着一起做  Teacher leads — do it together",sz=14,c=GRAY,a=PP_ALIGN.CENTER)
ib(s,0.4,1.3,9.2,3.9,"📷 Booklet 截图 / Booklet pages screenshot")
n+=1;pn(s,n)

# 31. PROJECTS OVERVIEW
s=ns();bg(s,CREAM);hb(s,"🛠️ 动手时间！  Hands-On Time — 3 个项目",BROWN)
projects=[
    ("PROJECT 1","🎯 寻宝任务","Treasure Hunt","按指令找位置 (基础)\nBy compass directions",WARM,COMPRED,"基础 / Basic"),
    ("PROJECT 2","🗺️ 我的营地地图","Camp Map","画一个简单地图\nDraw a simple map",RGBColor(0xFF,0xE0,0xB2),NAVY,"中级 / Mid"),
    ("PROJECT 3","🧲 自制指南针","Make a Compass","用针 + 磁铁 + 水\nNeedle + magnet + water",RGBColor(0xDC,0xED,0xC8),OK,"进阶 / Advanced"),
]
for i,(lbl,nm,en,d,bgc,cl,lvl) in enumerate(projects):
    x=0.3+i*3.2
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(0.95),Inches(3.1),Inches(4.15))
    sh.fill.solid();sh.fill.fore_color.rgb=bgc;sh.line.color.rgb=cl;sh.line.width=Pt(2.5)
    tb(s,x+0.1,1.05,2.9,0.35,lbl,sz=12,b=True,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,1.4,2.9,0.6,nm,sz=20,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,2.0,2.9,0.35,en,sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    pill(s,x+0.85,2.45,1.4,0.35,lvl,cl,sz=10)
    ib(s,x+0.2,2.95,2.8,1.2,"📷 示范")
    ls=d.split('\n')
    tf=tb(s,x+0.15,4.20,2.85,0.4,ls[0],sz=12,c=DARK,a=PP_ALIGN.CENTER)
    for ln in ls[1:]:ap(tf,ln,sz=12,c=DARK,a=PP_ALIGN.CENTER)
n+=1;pn(s,n)

# 32. PROJECT 1 — Treasure Hunt
s=ns();bg(s,CREAM);hb(s,"🎯 Project 1: 寻宝任务  Treasure Hunt",COMPRED)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(0.95),Inches(4.4),Inches(0.4))
sh.fill.solid();sh.fill.fore_color.rgb=COMPRED;sh.line.fill.background()
tb(s,0.4,0.98,4.2,0.35,"🧺 材料  Materials",sz=14,b=True,c=WHITE)
tf=tb(s,0.4,1.45,4.4,2.1,"🃏 指令卡片 (老师写好)  Direction cards",sz=13,c=DARK)
ap(tf,"🎁 小奖品 / 小物品 (藏起来)  Hidden small treasures",sz=13,c=DARK)
ap(tf,"🧭 罗盘 (Project 3 做的)  The compass we made",sz=13,c=DARK)
ap(tf,"📏 几张写有「3 步」「5 步」的卡  Step cards",sz=13,c=DARK)
sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(0.95),Inches(4.8),Inches(0.4))
sh2.fill.solid();sh2.fill.fore_color.rgb=SUN;sh2.line.fill.background()
tb(s,5.0,0.98,4.6,0.35,"👉 玩法  Steps",sz=14,b=True,c=WHITE)
tf2=tb(s,5.0,1.45,4.7,2.1,"1️⃣ 老师在教室藏 4-5 个小物品",sz=13,c=DARK)
ap(tf2,"2️⃣ 给学生一张指令卡, 例如:",sz=13,c=DARK)
ap(tf2,"     「向北 5 步, 向东 3 步」",sz=12,c=NAVY)
ap(tf2,"3️⃣ 学生用罗盘走过去找!",sz=13,c=DARK)
ap(tf2,"4️⃣ 找到物品 — 大声说指令",sz=13,c=DARK)
sh3=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(3.85),Inches(9.4),Inches(1.35))
sh3.fill.solid();sh3.fill.fore_color.rgb=WARM;sh3.line.color.rgb=COMPRED;sh3.line.width=Pt(2)
tb(s,0.5,3.95,9,0.35,"🗣️ 展示句型  Say These:",sz=14,b=True,c=COMPRED)
tb(s,0.5,4.35,4.5,0.35,"·  我向 ___ 走 ___ 步。",sz=14,c=DARK)
tb(s,0.5,4.65,4.5,0.35,"·  我找到了 ___ !",sz=14,c=DARK)
tb(s,5.2,4.35,4.5,0.35,"·  这是 ___ 边 (东/南/西/北)。",sz=14,c=DARK)
tb(s,5.2,4.65,4.5,0.35,"·  Treasure is in the ___ corner.",sz=14,c=DARK)
n+=1;pn(s,n)
notes(s,"老师准备 (低 prep):\n• 提前藏 4-5 件小物品 (橡皮、贴纸、糖果都行)\n• 写 4-5 张指令卡片: 「从老师这里, 向北 5 步, 向东 3 步」\n• 学生用 Project 3 做的罗盘走指令路线\n• 找到的人大声读自己的指令")

# 33. PROJECT 2 — Camp Map
s=ns();bg(s,CREAM);hb(s,"🗺️ Project 2: 我的营地地图  Camp Map",NAVY)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(0.95),Inches(4.4),Inches(0.4))
sh.fill.solid();sh.fill.fore_color.rgb=NAVY;sh.line.fill.background()
tb(s,0.4,0.98,4.2,0.35,"🧺 材料  Materials",sz=14,b=True,c=WHITE)
tf=tb(s,0.4,1.45,4.4,2.1,"📄 大白纸 (A4 或更大)  Big paper",sz=13,c=DARK)
ap(tf,"🖍️ 彩笔  Markers",sz=13,c=DARK)
ap(tf,"📏 直尺  Ruler",sz=13,c=DARK)
ap(tf,"🧭 4 方向标签 (东南西北)  N/S/E/W labels",sz=13,c=DARK)
sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(0.95),Inches(4.8),Inches(0.4))
sh2.fill.solid();sh2.fill.fore_color.rgb=SUN;sh2.line.fill.background()
tb(s,5.0,0.98,4.6,0.35,"👉 做法  Steps",sz=14,b=True,c=WHITE)
tf2=tb(s,5.0,1.45,4.7,2.1,"1️⃣ 在纸上写「上北下南左西右东」",sz=13,c=DARK)
ap(tf2,"2️⃣ 画营地的 4 个区域:",sz=13,c=DARK)
ap(tf2,"     ⛺ 帐篷 / 🔥 火 / 🍱 食物 / 🛡️ 安全",sz=12,c=NAVY)
ap(tf2,"3️⃣ 用方向标出位置",sz=13,c=DARK)
ap(tf2,"4️⃣ 给地图起一个名字",sz=13,c=DARK)
sh3=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(3.85),Inches(9.4),Inches(1.35))
sh3.fill.solid();sh3.fill.fore_color.rgb=WARM;sh3.line.color.rgb=NAVY;sh3.line.width=Pt(2)
tb(s,0.5,3.95,9,0.35,"🗣️ 展示句型  Say These:",sz=14,b=True,c=NAVY)
tb(s,0.5,4.35,4.5,0.35,"·  这是我的营地地图。",sz=14,c=DARK)
tb(s,0.5,4.65,4.5,0.35,"·  在 ___ 边, 我画了 ___ 。",sz=14,c=DARK)
tb(s,5.2,4.35,4.5,0.35,"·  火堆在 ___, 远离帐篷。",sz=14,c=DARK)
tb(s,5.2,4.65,4.5,0.35,"·  Tent is in the ___ corner.",sz=14,c=DARK)
n+=1;pn(s,n)
notes(s,"低 prep:\n• 一张大纸 + 彩笔即可\n• 每张纸先画一个十字 + 标 N/S/E/W\n• 学生画自己想象的营地 (借用 Day 2 的 4 区域知识)\n• 老师可以在白板示范一份")

# 34. PROJECT 3 — Make a Compass
s=ns();bg(s,CREAM);hb(s,"🧲 Project 3: 自制指南针  Make a Compass",OK)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(0.95),Inches(4.4),Inches(0.4))
sh.fill.solid();sh.fill.fore_color.rgb=OK;sh.line.fill.background()
tb(s,0.4,0.98,4.2,0.35,"🧺 材料  Materials",sz=14,b=True,c=WHITE)
tf=tb(s,0.4,1.45,4.4,2.1,"🪡 缝衣针  Sewing needle (1 根)",sz=13,c=DARK)
ap(tf,"🧲 磁铁  Magnet (1 块, 老师准备)",sz=13,c=DARK)
ap(tf,"🟫 软木塞 / 泡沫小片  Cork or foam",sz=13,c=DARK)
ap(tf,"🥣 一碗水  Bowl of water",sz=13,c=DARK)
ap(tf,"⚠️ 老师监督 (针很尖)  Teacher supervised!",sz=12,b=True,c=ALERT)
sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(0.95),Inches(4.8),Inches(0.4))
sh2.fill.solid();sh2.fill.fore_color.rgb=SUN;sh2.line.fill.background()
tb(s,5.0,0.98,4.6,0.35,"👉 做法 4 步  4 Steps",sz=14,b=True,c=WHITE)
tf2=tb(s,5.0,1.45,4.7,2.2,"1️⃣ 用磁铁朝同一个方向摩擦针 30 次",sz=13,c=DARK)
ap(tf2,"     Rub needle with magnet, same direction",sz=10,c=GRAY)
ap(tf2,"2️⃣ 把针放在软木 / 泡沫小片上",sz=13,c=DARK)
ap(tf2,"3️⃣ 把软木放在水面上",sz=13,c=DARK)
ap(tf2,"4️⃣ 等针稳定 — 它就指向北! 🧭",sz=13,b=True,c=COMPRED)
sh3=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(3.95),Inches(9.4),Inches(1.25))
sh3.fill.solid();sh3.fill.fore_color.rgb=WARM;sh3.line.color.rgb=OK;sh3.line.width=Pt(2)
tb(s,0.5,4.05,9,0.35,"🎬 视频教程  Video Tutorial:",sz=14,b=True,c=OK)
tb(s,0.5,4.40,9,0.35,"🔗 https://www.youtube.com/shorts/YZiVvUFkibE",sz=12,b=True,c=NAVY)
tb(s,0.5,4.75,9,0.35,"🗣️ 展示句型: 我做的指南针指向 ___ 。/ My compass points to ___.",sz=12,b=True,c=DARK)
n+=1;pn(s,n)
notes(s,"老师准备 (中等 prep):\n• 每组 1 块磁铁, 5-10 根缝衣针, 1 个浅碗水, 软木/泡沫\n• 安全: 针很尖, 老师全程监督, 摩擦后老师收回\n• 摩擦原理: 磁铁让针的「磁极」对齐, 变成小磁铁\n• 视频链接 https://www.youtube.com/shorts/YZiVvUFkibE — 提前看一遍\n• 失败常见原因: 摩擦次数太少 / 摩擦方向不一致 / 软木太大/不平")

# 35. CLOSING — mission complete
s=ns();bg(s,NAVY)
tb(s,1,0.6,8,0.7,"🏆 任务完成!",sz=44,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,1,1.35,8,0.5,"Mission Complete!",sz=20,c=WARM,a=PP_ALIGN.CENTER)
sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.5),Inches(2.0),Inches(3.0),Inches(3.0))
sh.fill.solid();sh.fill.fore_color.rgb=GOLD;sh.line.color.rgb=WHITE;sh.line.width=Pt(4)
tb(s,3.5,2.4,3.0,0.6,"🧭",sz=80,a=PP_ALIGN.CENTER)
tb(s,3.5,3.6,3.0,0.5,"方向大师",sz=22,b=True,c=NAVY,a=PP_ALIGN.CENTER)
tb(s,3.5,4.1,3.0,0.4,"Direction Master",sz=12,c=BROWN,a=PP_ALIGN.CENTER)
tb(s,1,5.0,8,0.4,"明天: Day 4 — 食物与水 / Day 4: Food & Water",sz=14,c=WARM,a=PP_ALIGN.CENTER)
n+=1;pn(s,n)

# === Save ===
out="/Users/Huan/projects/summercourse/Chinese/野外生存与探险wilderness_pbl/day3_compass.pptx"
prs.save(out)
print(f"Saved {out}  ({n} slides)")
