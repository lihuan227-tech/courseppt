#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
野外生存与探险 Day 4 — 紧急情况怎么办？
What to Do in Emergencies (Weather · Danger · Lost)
"Explorer Mission" framing — kids learn weather changes, danger ID, and what
to do when lost. Closing project: collective survival-safety poster.
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

# --- Palette (emergency theme: caution red + warning yellow) ---
PINE   = RGBColor(0x1E,0x4D,0x2B)
SUN    = RGBColor(0xE0,0x7A,0x2C)
CREAM  = RGBColor(0xFD,0xF6,0xE3)
BROWN  = RGBColor(0x6B,0x44,0x23)
WHITE  = RGBColor(0xFF,0xFF,0xFF)
DARK   = RGBColor(0x2C,0x2C,0x2C)
GRAY   = RGBColor(0x88,0x88,0x88)
LGRAY  = RGBColor(0xBB,0xBB,0xBB)
WARM   = RGBColor(0xFF,0xF3,0xE0)
IMGBG  = RGBColor(0xE8,0xE8,0xE8)
OK     = RGBColor(0x38,0x8E,0x3C)
# Emergency-specific
ALERT  = RGBColor(0xC6,0x28,0x28)   # primary: emergency red
WARN   = RGBColor(0xF9,0xA8,0x25)   # warning yellow
WEATHER= RGBColor(0x42,0xA5,0xF5)   # weather blue
LOST   = RGBColor(0x6A,0x1B,0x9A)   # lost purple
GOLD   = RGBColor(0xF9,0xA8,0x25)
HOT    = RGBColor(0xE6,0x52,0x2C)
COLD   = RGBColor(0x78,0xA7,0xB5)

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
def hb(s,txt,c=ALERT,t=0.15):
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

def video_slide(title_cn,title_en,before_task,after_action,bgc=ALERT):
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
    s=ns();bg(s,CREAM);hb(s,f"🚨 {title_cn}  {title_en}",ALERT)
    tb(s,0.4,0.85,9.2,0.4,question_cn,sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
    tb(s,0.4,1.20,9.2,0.3,question_en,sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    a_box=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.6),Inches(4.5),Inches(2.4))
    a_box.fill.solid();a_box.fill.fore_color.rgb=WHITE;a_box.line.color.rgb=WEATHER;a_box.line.width=Pt(2.5)
    pill(s,0.5,1.7,0.7,0.4,"A",WEATHER,sz=16)
    tb(s,1.2,1.65,3.6,0.5,a_label,sz=18,b=True,c=WEATHER)
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
    tb(s,0.5,1.1,4.3,1.4,w,sz=72,b=True,c=ALERT,a=PP_ALIGN.CENTER)
    tb(s,0.5,2.4,4.3,0.4,f"{py}  {en}",sz=20,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,0.5,2.85,4.3,0.4,"👉 跟我读！Read after me!",sz=14,c=color,a=PP_ALIGN.CENTER)
    ib(s,5.3,1.0,4.4,2.5,img)
    sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(3.8),Inches(9.2),Inches(1.2))
    sh2.fill.solid();sh2.fill.fore_color.rgb=WHITE;sh2.line.color.rgb=color;sh2.line.width=Pt(2)
    tb(s,0.6,3.9,1.5,0.4,"例句 Example",sz=14,b=True,c=color)
    tb(s,0.6,4.3,8.8,0.5,sent,sz=20,b=True,c=DARK)
    return s

def word_card_write(w,py,en,strokes_hint,color=ALERT):
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
s=ns();bg(s,ALERT)
sh=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,Inches(2.4),W,Inches(2.0))
sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.fill.background()
tb(s,1,0.4,8,0.5,"DAY 4",sz=18,b=True,c=GOLD,a=PP_ALIGN.CENTER)
tb(s,1,0.95,8,0.7,"🚨 紧急情况怎么办?",sz=42,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,1,1.7,8,0.5,"What to Do in Emergencies",sz=22,c=WARM,a=PP_ALIGN.CENTER)
tb(s,1,2.6,8,0.5,"🧭 探险家任务  Explorer Mission",sz=24,b=True,c=ALERT,a=PP_ALIGN.CENTER)
tb(s,1,3.15,8,0.4,"Weather · Danger · Lost · Survive",sz=14,c=BROWN,a=PP_ALIGN.CENTER)
tb(s,1,3.55,8,0.4,"看天气 · 防危险 · 不迷路 · 会求救",sz=14,b=True,c=BROWN,a=PP_ALIGN.CENTER)
tb(s,1,4.6,8,0.4,"野外生存与探险 · Wilderness Survival",sz=14,b=True,c=GOLD,a=PP_ALIGN.CENTER)
n+=1;pn(s,n)
notes(s,"开场 (1 分钟):\n• 「探险家们! 大自然有时候会给你「惊喜」 — 突然下雨, 突然有动物。」\n• 今天 4 个任务: 看天气 / 防危险 / 不迷路 / 会求救\n• 完成所有任务可以拿到「应急大师」徽章")

# 2. EXPLORER MISSION — 4 tasks
s=ns();bg(s,CREAM);hb(s,"🧭 今天的任务  Today's Mission",ALERT)
tb(s,0.4,0.9,9.2,0.4,"小探险家们 — 4 件事!",sz=20,b=True,c=ALERT,a=PP_ALIGN.CENTER)
tb(s,0.4,1.32,9.2,0.3,"Little explorers — 4 jobs today!",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
mission_card(s,0.4,1.85,2.25,2.6,1,"看天气", "Read weather",  "🌦️",WEATHER)
mission_card(s,2.75,1.85,2.25,2.6,2,"防危险", "Spot danger",   "⚠️",ALERT)
mission_card(s,5.10,1.85,2.25,2.6,3,"不迷路", "Don't get lost","🧭",LOST)
mission_card(s,7.45,1.85,2.25,2.6,4,"会求救", "Call for help", "📢",WARN)
sentence_frame_bar(s,4.65,
    "我是小探险家, 我会 ___ 。",
    "I am a little explorer, I will ___.")
n+=1;pn(s,n)

# 3. PLEDGE
s=ns();bg(s,CREAM);hb(s,"🤝 探险家宣言  Explorer Pledge",ALERT)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(1.0),Inches(1.1),Inches(8.0),Inches(3.2))
sh.fill.solid();sh.fill.fore_color.rgb=WARM;sh.line.color.rgb=ALERT;sh.line.width=Pt(2.5)
tb(s,1.2,1.3,7.6,0.5,"🌦️ 天气变了, 我会找避难。",sz=20,b=True,c=WEATHER,a=PP_ALIGN.CENTER)
tb(s,1.2,1.85,7.6,0.4,"When weather changes, I find shelter.",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,1.2,2.35,7.6,0.5,"⚠️ 危险的东西, 我不碰、不靠近。",sz=20,b=True,c=ALERT,a=PP_ALIGN.CENTER)
tb(s,1.2,2.85,7.6,0.4,"Dangerous things — I don't touch, don't approach.",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,1.2,3.35,7.6,0.5,"🧭 迷路的时候, 我留在原地。",sz=20,b=True,c=LOST,a=PP_ALIGN.CENTER)
tb(s,1.2,3.85,7.6,0.4,"If lost — I stay put.",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,0.4,4.5,9.2,0.5,"👉 一起举手念! Raise your hand and say it!",sz=14,b=True,c=SUN,a=PP_ALIGN.CENTER)
n+=1;pn(s,n)

# 4. SESSION 1 DIVIDER
s=div("Session 1  上午","🌦️ 天气 + ⚠️ 危险 + 🧭 迷路 应对  Weather · Danger · Lost",ALERT,"📖");n+=1;pn(s,n)

# 5. HOOK
s=ns();bg(s,CREAM);hb(s,"🤔 想一想  Imagine This!",WARN)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.95),Inches(9.2),Inches(1.6))
sh.fill.solid();sh.fill.fore_color.rgb=WARM;sh.line.color.rgb=WARN;sh.line.width=Pt(2.5)
tb(s,0.6,1.1,8.8,0.6,"你在森林里 — 突然下大雨怎么办?",sz=22,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.6,1.7,8.8,0.4,"You're in the forest — sudden rain. What now?",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,0.6,2.10,8.8,0.4,"👉 跟同桌说一说 / Tell your partner",sz=12,b=True,c=SUN,a=PP_ALIGN.CENTER)
hints=[("☂️","找避雨"),("🌳","树底下?"),("🏃","快跑回去?"),("📞","告诉大人")]
for i,(em,t) in enumerate(hints):
    x=0.4+i*2.35
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(2.85),Inches(2.20),Inches(1.7))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=WEATHER;sh.line.width=Pt(2)
    tb(s,x+0.05,2.95,2.1,0.7,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.85,2.1,0.5,t,sz=14,b=True,c=WEATHER,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.7,
    "下雨了, 我会 ___ 。",
    "It's raining, I will ___.")
n+=1;pn(s,n)
notes(s,"打开课程: 让学生 think-pair-share。\n• 收集答案 (有些会说错的: 树底下打雷会引电!)\n• 「今天我们学怎么应对各种紧急情况」")

# 6. VIDEO — Weather changes
s=video_slide(
    "🌦️ 天气会变  Weather Changes",
    "Weather Safety for Kids",
    "👂 听 / 看:\n1. 天气有几种?\n2. 哪种最危险?\n3. 看到乌云怎么办?",
    "🎯 看完后:\n• 模仿: 太阳/雨/雪/风 4 个动作\n• 说: 「我看到 ___, 所以 ___」",
    bgc=WEATHER);n+=1;pn(s,n)
notes(s,"视频建议 (1-3 分钟):\n• YouTube 搜「Weather safety for kids」「Weather song for kids」\n• 关键: 看到乌云 / 听到雷 → 找避难所; 不要在树下/水里")

# 7. WEATHER TYPES — 5 types overview
s=ns();bg(s,CREAM);hb(s,"🌤️ 5 种天气  5 Weather Types",WEATHER)
tb(s,0.4,0.9,9.2,0.4,"看图认天气!  Match the weather!",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
weathers=[("☀️","晴天",     "Sunny", "出门天气好",       SUN),
          ("🌧️","下雨",     "Rainy", "要带伞 / 雨衣",    WEATHER),
          ("💨","大风",     "Windy", "树枝会掉",         BROWN),
          ("🥵","很热",     "Hot",   "小心中暑",         HOT),
          ("🥶","很冷",     "Cold",  "穿厚衣服",         COLD)]
for i,(em,cn,en,desc,cl) in enumerate(weathers):
    x=0.3+i*1.92
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.55),Inches(1.78),Inches(2.85))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=cl;sh.line.width=Pt(2.5)
    tb(s,x+0.05,1.65,1.7,0.9,em,sz=48,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.7,1.7,0.5,cn,sz=22,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.25,1.7,0.4,en,sz=12,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.75,1.7,0.55,desc,sz=11,b=True,c=DARK,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.65,
    "今天的天气是 ___, 我要 ___ 。",
    "Today's weather is ___, I should ___.")
n+=1;pn(s,n)

# 8. WEATHER DANGERS
s=ns();bg(s,CREAM);hb(s,"⚠️ 天气变化的危险  Weather Dangers",ALERT)
tb(s,0.4,0.9,9.2,0.4,"不同天气, 不同的小心!",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
dangers=[("🌧️","下雨", "滑倒",       "Slipping",       "地面湿, 鞋子湿"),
         ("🥵","很热", "中暑",       "Heatstroke",      "头晕、太热"),
         ("🥶","很冷", "着凉 / 冻伤", "Cold / frostbite","手脚发紫"),
         ("💨","大风", "树枝掉",     "Falling branches","枯枝会被吹下")]
for i,(em,cn_w,cn_d,en_d,desc) in enumerate(dangers):
    col=i%2;row=i//2
    x=0.4+col*4.65;y=1.5+row*1.6
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(4.5),Inches(1.45))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=ALERT;sh.line.width=Pt(2)
    tb(s,x+0.1,y+0.1,1.0,1.2,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+1.15,y+0.1,3.3,0.4,f"{cn_w} → ⚠️ {cn_d}",sz=15,b=True,c=ALERT)
    tb(s,x+1.15,y+0.55,3.3,0.3,en_d,sz=11,c=GRAY)
    tb(s,x+1.15,y+0.95,3.3,0.4,f"💡 {desc}",sz=11,c=DARK)
sentence_frame_bar(s,4.85,
    "下雨可能 ___, 所以 ___ 。",
    "Rain may cause ___, so ___.")
n+=1;pn(s,n)

# 9. A/B — Dark clouds
s=ab_slide("看到乌云怎么办？","Dark Clouds!",
    "天上突然有黑色乌云 — 你怎么办?",
    "Dark clouds suddenly! What do you do?",
    "🏃","继续走 / 玩","乌云没事, 还能走!",
    "⛺","找避难, 等雨过","回房间 / 进帐篷 / 找大树后",
    "B 找避难",
    "看到乌云通常 5-15 分钟会下大雨 + 闪电。要提前找地方避雨。"
)
n+=1;pn(s,n)

# 10. DANGER OVERVIEW — 3 things to never touch
s=ns();bg(s,CREAM);hb(s,"⚠️ 3 个「不」!  3 Don'ts!",ALERT)
tb(s,0.4,0.9,9.2,0.4,"看到这 3 样 — 不碰, 不吃, 不靠近!",sz=20,b=True,c=ALERT,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.3,"3 things — DON'T touch, DON'T eat, DON'T approach!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
nono=[("🍄","不认识的蘑菇",   "Unknown mushrooms","可能有毒"),
      ("🐍","陌生动物",       "Strange animals",  "蛇、熊、野狗"),
      ("🗑️","不明的东西",     "Unknown items",    "针 / 玻璃 / 旧瓶子")]
for i,(em,cn,en,why) in enumerate(nono):
    x=0.4+i*3.15
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.85),Inches(3.0),Inches(2.7))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=ALERT;sh.line.width=Pt(2.5)
    pill(s,x+0.7,1.95,1.6,0.4,"❌ 不可以",ALERT,sz=12)
    tb(s,x+0.1,2.5,2.8,0.9,em,sz=52,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,3.55,2.8,0.4,cn,sz=15,b=True,c=DARK,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,4.0,2.8,0.3,en,sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,4.30,2.8,0.3,f"💡 {why}",sz=11,b=True,c=ALERT,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.85,
    "我不 ___ 也不 ___, 因为 ___ 。",
    "I don't ___ and don't ___, because ___.")
n+=1;pn(s,n)

# 11. SNAKE — what to do
s=ns();bg(s,CREAM);hb(s,"🐍 看到蛇怎么办？  See a Snake!",ALERT)
ib(s,0.3,1.0,4.4,3.4,"📷 蛇 / Snake — 老师替换照片")
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(1.0),Inches(4.8),Inches(3.4))
sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=ALERT;sh.line.width=Pt(2.5)
head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(1.0),Inches(4.8),Inches(0.5))
head.fill.solid();head.fill.fore_color.rgb=ALERT;head.line.fill.background()
tb(s,5.05,1.08,4.6,0.4,"🚨 4 步规则  4-Step Rule",sz=14,b=True,c=WHITE)
tf=tb(s,5.1,1.65,4.55,0.4,"1️⃣ 停下来  STOP",sz=14,b=True,c=DARK)
ap(tf,"   不要跑, 不要叫",sz=10,c=GRAY)
ap(tf," ",sz=4)
ap(tf,"2️⃣ 慢慢退后  Slowly back away",sz=14,b=True,c=DARK)
ap(tf,"   一步一步, 不要转身跑",sz=10,c=GRAY)
ap(tf," ",sz=4)
ap(tf,"3️⃣ 远离 3 米以上  Stay 3m+ away",sz=14,b=True,c=DARK)
ap(tf,"   蛇够不到的距离",sz=10,c=GRAY)
ap(tf," ",sz=4)
ap(tf,"4️⃣ 告诉大人  Tell an adult",sz=14,b=True,c=DARK)
ap(tf,"   绝对不要去抓!",sz=10,c=ALERT,b=True)
sentence_frame_bar(s,4.55,
    "看到蛇, 我 ___, 不 ___ 。",
    "See a snake — I ___, NOT ___.")
n+=1;pn(s,n)
notes(s,"重点教学:\n• 不要跑 — 跑容易摔倒, 蛇可能追\n• 不要叫 — 叫声会激怒蛇\n• 慢慢退 → 远离 → 找大人\n• 同样的方法适用于熊、野狗等所有大动物")

# 12. SAFE/DANGEROUS — Gesture game
s=ns();bg(s,CREAM);hb(s,"🛡️ 安全 vs 危险!  Safe vs Dangerous!",ALERT)
rb=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(1.0),Inches(9.4),Inches(1.1))
rb.fill.solid();rb.fill.fore_color.rgb=WARM;rb.line.color.rgb=ALERT;rb.line.width=Pt(2)
tb(s,0.5,1.10,9.0,0.4,"老师说一件事 — 学生用动作回答!",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.5,1.50,4.4,0.5,"✅ 安全 = 双手举高 (V)",sz=15,b=True,c=OK,a=PP_ALIGN.CENTER)
tb(s,5.1,1.50,4.4,0.5,"❌ 危险 = 双手交叉 (X)",sz=15,b=True,c=ALERT,a=PP_ALIGN.CENTER)
ex=[("🍄","摸不认识的蘑菇",        "❌"),
    ("⛅","看见乌云就找避雨",      "✅"),
    ("🐍","看见蛇就慢慢退后",      "✅"),
    ("🗑️","捡路上奇怪的瓶子",     "❌"),
    ("🌳","下大雨躲在大树下",      "❌"),
    ("👨‍👩‍👧","告诉大人发生的事",    "✅")]
for i,(em,desc,ans) in enumerate(ex):
    col=i%2;row=i//2
    x=0.3+col*4.7;y=2.3+row*0.85
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(4.5),Inches(0.75))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=ALERT;sh.line.width=Pt(1.5)
    tb(s,x+0.1,y+0.15,0.6,0.5,em,sz=22)
    tb(s,x+0.75,y+0.18,3.2,0.4,desc,sz=12,b=True,c=DARK)
    tb(s,x+3.95,y+0.15,0.5,0.5,ans,sz=22,a=PP_ALIGN.CENTER)
n+=1;pn(s,n)
notes(s,"老师玩法 (3-5 分钟):\n• 念上面 6 个情境, 学生举手或交叉手回答。\n• G1-3 加一句: 「___ 不安全, 因为 ___」\n• 「下大雨躲在大树下」错 — 雷会打高的地方")

# 13. WHAT IS GETTING LOST
s=ns();bg(s,CREAM);hb(s,"🧭 什么是「迷路」？  What is Getting Lost?",LOST)
big=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.95),Inches(9.2),Inches(1.4))
big.fill.solid();big.fill.fore_color.rgb=LOST;big.line.fill.background()
tb(s,0.6,1.1,8.8,0.6,"❓ 不知道自己在哪里, 也不知道怎么回家",sz=22,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,0.6,1.75,8.8,0.4,"You don't know where you are, or how to get back",sz=14,c=WARM,a=PP_ALIGN.CENTER)
tb(s,0.6,2.10,8.8,0.3,"= 迷路  Getting Lost",sz=14,b=True,c=WARM,a=PP_ALIGN.CENTER)
# When can it happen
when=[("🌲","在森林里看不到路"),("🌫️","起雾, 看不远"),("🚶","跟着别人走丢了"),("📵","没有手机 / 信号")]
for i,(em,t) in enumerate(when):
    x=0.4+i*2.35
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(2.65),Inches(2.20),Inches(1.6))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=LOST;sh.line.width=Pt(2)
    tb(s,x+0.05,2.75,2.1,0.7,em,sz=40,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.5,2.1,0.65,t,sz=12,b=True,c=DARK,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.4,
    "迷路 = ___ 。",
    "Lost means ___.")
n+=1;pn(s,n)

# 14. STOP RULE — 4 letters
s=ns();bg(s,CREAM);hb(s,"🛑 迷路了 → STOP!  S.T.O.P.",LOST)
tb(s,0.4,0.9,9.2,0.4,"4 步: 停 → 想 → 看 → 决定",sz=20,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.3,"4 steps: Stop → Think → Observe → Plan",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
stop=[("S","🛑","停 (Stop)",     "立刻停下来, 不乱跑",          ALERT),
      ("T","🤔","想 (Think)",    "上一次大人在哪里?",          WARN),
      ("O","👀","看 (Observe)",  "周围有什么? 大树? 河?",      LOST),
      ("P","📋","决定 (Plan)",   "原地等? 还是去找标志?",      OK)]
for i,(letter,em,cn,desc,cl) in enumerate(stop):
    x=0.3+i*2.4
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.85),Inches(2.25),Inches(2.7))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=cl;sh.line.width=Pt(2.5)
    pill(s,x+0.5,1.95,1.25,0.45,letter,cl,sz=22)
    tb(s,x+0.1,2.55,2.05,0.7,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,3.35,2.05,0.45,cn,sz=15,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,3.85,2.05,0.6,desc,sz=11,c=DARK,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.7,
    "迷路了, 我要 ___ → ___ → ___ → ___ 。",
    "If lost: Stop → Think → Observe → Plan.")
n+=1;pn(s,n)
notes(s,"国际通用的 STOP 规则 (童子军/搜救队都用):\n• Stop: 一发现迷路, 立刻不动\n• Think: 想想最后看到大人的地方\n• Observe: 看周围标志 (路、河、大树)\n• Plan: 决定是「原地等」还是「往明显标志走」\n• K-5 重点: 不乱跑! 大人能找到不动的孩子, 找不到乱跑的。")

# 15. STAY PUT + Find landmark
s=ns();bg(s,CREAM);hb(s,"🌳 不乱跑 + 找标志  Stay + Find Landmark",LOST)
tb(s,0.4,0.9,9.2,0.4,"迷路时, 这 2 件事最重要!",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
# Stay put
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.4),Inches(4.5),Inches(2.95))
sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=ALERT;sh.line.width=Pt(2.5)
pill(s,0.6,1.55,1.6,0.4,"❌ 不!",ALERT,sz=14)
tb(s,0.6,2.0,4.2,0.5,"不要乱跑!",sz=22,b=True,c=ALERT)
tb(s,0.6,2.5,4.2,0.4,"Don't run around!",sz=12,c=GRAY)
tb(s,0.6,3.0,4.2,0.4,"💡 越跑越远, 大人找不到你",sz=12,c=DARK)
tb(s,0.6,3.4,4.2,0.4,"💡 容易摔倒, 容易受伤",sz=12,c=DARK)
tb(s,0.6,3.8,4.2,0.4,"💡 越累越冷越饿",sz=12,c=DARK)
# Find landmark
sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.1),Inches(1.4),Inches(4.5),Inches(2.95))
sh2.fill.solid();sh2.fill.fore_color.rgb=WHITE;sh2.line.color.rgb=OK;sh2.line.width=Pt(2.5)
pill(s,5.3,1.55,1.6,0.4,"✅ 做!",OK,sz=14)
tb(s,5.3,2.0,4.2,0.5,"找明显标志",sz=22,b=True,c=OK)
tb(s,5.3,2.5,4.2,0.4,"Find clear landmarks",sz=12,c=GRAY)
tb(s,5.3,3.0,4.2,0.4,"💡 大树 / 大石头 (站旁边)",sz=12,c=DARK)
tb(s,5.3,3.4,4.2,0.4,"💡 河 / 路 (沿着走)",sz=12,c=DARK)
tb(s,5.3,3.8,4.2,0.4,"💡 高的地方 (能看到远)",sz=12,c=DARK)
sentence_frame_bar(s,4.45,
    "迷路了, 我要在 ___ 旁边等大人。",
    "If lost, I wait near ___ for an adult.")
n+=1;pn(s,n)

# 16. WHISTLE / CALL FOR HELP
s=ns();bg(s,CREAM);hb(s,"📢 怎么求救？  How to Call for Help",WARN)
tb(s,0.4,0.9,9.2,0.4,"3 种求救方法! 3 ways to call for help!",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
help_ways=[("📢","哨子 3 声",     "Whistle x 3",   "国际求救信号"),
           ("👋","挥手大叫",      "Wave + shout",  "「我在这里!」"),
           ("🔥","明亮 / 大颜色", "Bright color",  "穿亮色衣服易被看见")]
for i,(em,cn,en,why) in enumerate(help_ways):
    x=0.4+i*3.15
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.5),Inches(3.0),Inches(2.7))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=WARN;sh.line.width=Pt(2.5)
    tb(s,x+0.1,1.6,2.8,0.9,em,sz=52,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,2.55,2.8,0.5,cn,sz=22,b=True,c=ALERT,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,3.05,2.8,0.4,en,sz=12,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.15,3.55,2.85,0.6,why,sz=11,b=True,c=DARK,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.4,
    "迷路了, 我会 ___ 求救。",
    "If lost, I will ___ for help.")
tb(s,0.4,5.10,9.2,0.3,"💡 哨子 3 声 = 国际求救信号 (whistle 3 short blasts = SOS)",sz=12,b=True,c=ALERT,a=PP_ALIGN.CENTER)
n+=1;pn(s,n)
notes(s,"提醒: \n• 哨子 3 声是国际通用的求救信号 (3 = help)\n• 老师可以让学生练习吹哨子 3 声 (喧闹 — 但只 1 分钟)\n• 穿亮色 = 救援队最容易看见")

# 17. SESSION 2 DIVIDER
s=div("Session 2  下午","🔄 复习 + 语言目标 (我会认 6 / 我会写 3)",SUN,"📖");n+=1;pn(s,n)

# 18. QUICK REVIEW
s=ns();bg(s,CREAM);hb(s,"🔄 快速复习  Quick Review",SUN)
tb(s,0.4,0.85,9.2,0.4,"上午学了什么? Let's review!",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
qs=[("Q1","看到乌云怎么办?",       "⛺ 找避难"),
    ("Q2","看到蛇 — 跑还是退?",    "🚶 慢慢退后"),
    ("Q3","迷路了 — 跑还是停?",    "🛑 停, 不乱跑"),
    ("Q4","求救信号是哨子几声?",   "📢 3 声")]
for i,(num,q,a) in enumerate(qs):
    col=i%2;row=i//2
    x=0.4+col*4.65;y=1.4+row*1.7
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(4.5),Inches(1.5))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=ALERT;sh.line.width=Pt(2)
    pill(s,x+0.15,y+0.15,0.7,0.4,num,ALERT,sz=12)
    tb(s,x+0.95,y+0.1,3.4,0.45,q,sz=14,b=True,c=DARK)
    tb(s,x+0.15,y+0.65,4.2,0.4,"💬 答: ",sz=13,b=True,c=SUN)
    tb(s,x+0.85,y+0.65,3.5,0.4,a,sz=14,b=True,c=ALERT)
    tb(s,x+0.15,y+1.05,4.2,0.3,"举手抢答! Hand-raise!",sz=10,c=GRAY)
n+=1;pn(s,n)

# 19. SENTENCE FRAMES
s=ns();bg(s,CREAM);hb(s,"💬 句型卡  Sentence Frames",SUN)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.0),Inches(4.5),Inches(3.6))
sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=SUN;sh.line.width=Pt(2.5)
pill(s,0.6,1.15,1.6,0.4,"K (TK-K)",SUN,sz=14)
tb(s,0.6,1.7,4.1,0.5,"⚠️ ___ 危险!",sz=22,b=True,c=ALERT)
tb(s,0.6,2.2,4.1,0.4,"___ is dangerous!",sz=10,c=GRAY)
tb(s,0.6,2.7,4.1,0.5,"🛑 我要 ___ 。",sz=22,b=True,c=ALERT)
tb(s,0.6,3.2,4.1,0.4,"I should ___.",sz=10,c=GRAY)
tb(s,0.6,3.75,4.1,0.4,"💡 例: 蛇危险! 我要慢慢退。",sz=13,c=DARK)
tb(s,0.6,4.1,4.1,0.4,"Ex: Snake is dangerous! I'll slowly back away.",sz=10,c=GRAY)
sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.1),Inches(1.0),Inches(4.5),Inches(3.6))
sh2.fill.solid();sh2.fill.fore_color.rgb=WHITE;sh2.line.color.rgb=ALERT;sh2.line.width=Pt(2.5)
pill(s,5.3,1.15,1.6,0.4,"G1 - G3",ALERT,sz=14)
tb(s,5.3,1.7,4.1,0.5,"·  ___ 不安全, 因为 ___ 。",sz=18,b=True,c=ALERT)
tb(s,5.3,2.2,4.1,0.4,"___ is not safe, because ___.",sz=10,c=GRAY)
tb(s,5.3,2.7,4.1,0.5,"·  迷路了, 我应该 ___ 。",sz=18,b=True,c=ALERT)
tb(s,5.3,3.2,4.1,0.4,"If lost, I should ___.",sz=10,c=GRAY)
tb(s,5.3,3.75,4.1,0.4,"💡 例: 大树下不安全, 因为会打雷。",sz=11,c=DARK)
n+=1;pn(s,n)

# 20-25. WORD CARDS — 我会认 (6)
read_data=[
    ("天气",  "tiān qì", "weather",      "今天的天气很好。",          "📷 蓝天白云"),
    ("下雨",  "xià yǔ",  "rain",         "下雨了, 我们要找避难。",    "📷 雨景"),
    ("危险",  "wēi xiǎn","danger",       "陌生动物很危险。",          "📷 ⚠️ 标志"),
    ("蛇",    "shé",     "snake",        "看到蛇要慢慢退后。",        "📷 蛇 (小心)"),
    ("迷路",  "mí lù",   "get lost",     "迷路了不要乱跑。",          "📷 迷路的小孩"),
    ("食物",  "shí wù",  "food",         "我把食物收好了。",          "📷 食物"),
]
for w,py,en,sent,img in read_data:
    s=word_card_read(w,py,en,sent,img,color=SUN);n+=1;pn(s,n)

# 26-28. WORD CARDS — 我会写 (3)
write_data=[
    ("天气","tiān qì","weather","8 笔: 天 (大 + 一) + 气 (气字旁)\n抬头看 → 这就是「天气」!"),
    ("危险","wēi xiǎn","danger","12 笔: 危 (厃 + 㔾) + 险 (阝 + 佥)\n看到 ⚠️ 想到「危险」"),
    ("迷路","mí lù","lost","19 笔: 迷 (走 + 米) + 路 (足 + 各)\n找不到「路」就「迷」了"),
]
for w,py,en,strokes in write_data:
    s=word_card_write(w,py,en,strokes,color=ALERT);n+=1;pn(s,n)

# 29. SESSION 3 DIVIDER
s=div("Session 3  下午","📒 写 Booklet + 全班生存安全 Poster",BROWN,"🛡️");n+=1;pn(s,n)

# 30. DAY 4 BOOKLET
s=ns();bg(s,CREAM);hb(s,"📓 完成 Day 4 练习册  Day 4 Booklet",BROWN)
tb(s,0.4,0.85,9.2,0.4,"老师带着一起做  Teacher leads — do it together",sz=14,c=GRAY,a=PP_ALIGN.CENTER)
ib(s,0.4,1.3,9.2,3.9,"📷 Booklet 截图 / Booklet pages screenshot")
n+=1;pn(s,n)

# 31. SAFETY POSTER — overview
s=ns();bg(s,CREAM);hb(s,"🛡️ 集体项目: 生存安全 Poster  Survival Safety Poster",BROWN)
tb(s,0.4,0.9,9.2,0.4,"🎯 全班一起做一本「小册子」 — 4 页, 每页一个主题!",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.3,"Whole class makes a 4-page mini-book!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# Materials box
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.85),Inches(9.2),Inches(0.85))
sh.fill.solid();sh.fill.fore_color.rgb=WARM;sh.line.color.rgb=BROWN;sh.line.width=Pt(2)
tb(s,0.55,1.95,9.0,0.35,"🧺 材料  Materials:",sz=14,b=True,c=BROWN)
tb(s,0.55,2.30,9.0,0.4,"📒 小册子纸 / 打印模板  ·  🖍️ 彩笔  ·  ✂️ 剪刀  ·  📎 胶水",sz=12,c=DARK)
# 4 pages preview
pages=[("Page 1","🌧️","天气", "遇到下雨怎么办？", WEATHER),
       ("Page 2","🐍","危险", "看到蛇怎么办？",   ALERT),
       ("Page 3","🧭","迷路", "我会怎么做？",      LOST),
       ("Page 4","🍎","食物", "可以吃 vs 不可以吃",OK)]
for i,(lbl,em,cn,q,cl) in enumerate(pages):
    x=0.4+i*2.35
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(2.95),Inches(2.20),Inches(2.2))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=cl;sh.line.width=Pt(2.5)
    pill(s,x+0.45,3.05,1.3,0.35,lbl,cl,sz=11)
    tb(s,x+0.05,3.5,2.1,0.7,em,sz=42,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,4.25,2.1,0.4,cn,sz=18,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,4.65,2.1,0.4,q,sz=10,c=DARK,a=PP_ALIGN.CENTER)
n+=1;pn(s,n)
notes(s,"项目说明 (低 prep):\n• 每个学生 1 张 4 折的纸 (8 页) 或 4 页 stapled\n• 也可以全班合作: 4 个小组, 每组负责 1 页, 最后合并成班级 poster\n• 时间: 30-40 分钟")

# 32. POSTER PAGE 1 — Weather
s=ns();bg(s,CREAM);hb(s,"📒 Page 1: 天气  Weather",WEATHER)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(0.95),Inches(4.4),Inches(0.4))
sh.fill.solid();sh.fill.fore_color.rgb=WEATHER;sh.line.fill.background()
tb(s,0.4,0.98,4.2,0.35,"🎯 主题  Topic",sz=14,b=True,c=WHITE)
tf=tb(s,0.4,1.45,4.4,2.1,"❓ 遇到下雨怎么办?",sz=15,b=True,c=DARK)
ap(tf,"   What to do when it rains?",sz=11,c=GRAY)
ap(tf," ",sz=8)
ap(tf,"💡 想一想 / 画一画:",sz=13,b=True,c=WEATHER)
ap(tf,"·  我看到 ___ (天气标志)",sz=12,c=DARK)
ap(tf,"·  我可以去 ___ (避雨地方)",sz=12,c=DARK)
ap(tf,"·  我不可以 ___ (危险动作)",sz=12,c=DARK)
sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(0.95),Inches(4.8),Inches(0.4))
sh2.fill.solid();sh2.fill.fore_color.rgb=SUN;sh2.line.fill.background()
tb(s,5.0,0.98,4.6,0.35,"✏️ 内容建议  Content Ideas",sz=14,b=True,c=WHITE)
tf2=tb(s,5.0,1.45,4.7,2.1,"画 1: 一个孩子看到乌云",sz=12,c=DARK)
ap(tf2,"画 2: 跑进帐篷 / 屋子 / 大石头底下",sz=12,c=DARK)
ap(tf2,"画 3: ❌ 不要在大树下 (会打雷)",sz=12,c=ALERT)
ap(tf2,"标题: 「下雨时, 我会 ___」",sz=12,b=True,c=WEATHER)
sh3=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(3.85),Inches(9.4),Inches(1.35))
sh3.fill.solid();sh3.fill.fore_color.rgb=WARM;sh3.line.color.rgb=WEATHER;sh3.line.width=Pt(2)
tb(s,0.5,3.95,9,0.35,"🗣️ 展示句型  Say These:",sz=14,b=True,c=WEATHER)
tb(s,0.5,4.35,9,0.35,"·  下雨了, 我可以 ___ , 不可以 ___ 。",sz=14,c=DARK)
tb(s,0.5,4.70,9,0.35,"·  When it rains, I can ___, I cannot ___.",sz=14,c=GRAY)
n+=1;pn(s,n)

# 33. POSTER PAGE 2 — Danger (Snake)
s=ns();bg(s,CREAM);hb(s,"📒 Page 2: 危险  Danger",ALERT)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(0.95),Inches(4.4),Inches(0.4))
sh.fill.solid();sh.fill.fore_color.rgb=ALERT;sh.line.fill.background()
tb(s,0.4,0.98,4.2,0.35,"🎯 主题  Topic",sz=14,b=True,c=WHITE)
tf=tb(s,0.4,1.45,4.4,2.1,"❓ 看到蛇怎么办?",sz=15,b=True,c=DARK)
ap(tf,"   What to do if you see a snake?",sz=11,c=GRAY)
ap(tf," ",sz=8)
ap(tf,"💡 想一想 / 画一画:",sz=13,b=True,c=ALERT)
ap(tf,"·  画一条蛇",sz=12,c=DARK)
ap(tf,"·  画我自己 (慢慢退后)",sz=12,c=DARK)
ap(tf,"·  画大人 (告诉他/她)",sz=12,c=DARK)
sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(0.95),Inches(4.8),Inches(0.4))
sh2.fill.solid();sh2.fill.fore_color.rgb=SUN;sh2.line.fill.background()
tb(s,5.0,0.98,4.6,0.35,"✏️ 内容建议  Content Ideas",sz=14,b=True,c=WHITE)
tf2=tb(s,5.0,1.45,4.7,2.1,"画 1: 🐍 + 我 (远远看见)",sz=12,c=DARK)
ap(tf2,"画 2: ✅ 慢慢退后 (3 米外)",sz=12,c=OK)
ap(tf2,"画 3: ❌ 不要跑, 不要叫, 不要抓",sz=12,c=ALERT)
ap(tf2,"标题: 「看到蛇, 我 ___」",sz=12,b=True,c=ALERT)
sh3=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(3.85),Inches(9.4),Inches(1.35))
sh3.fill.solid();sh3.fill.fore_color.rgb=WARM;sh3.line.color.rgb=ALERT;sh3.line.width=Pt(2)
tb(s,0.5,3.95,9,0.35,"🗣️ 展示句型  Say These:",sz=14,b=True,c=ALERT)
tb(s,0.5,4.35,9,0.35,"·  看到蛇, 我会 ___, 不会 ___ 。",sz=14,c=DARK)
tb(s,0.5,4.70,9,0.35,"·  When I see a snake, I will ___, I will NOT ___.",sz=14,c=GRAY)
n+=1;pn(s,n)

# 34. POSTER PAGE 3 — Lost
s=ns();bg(s,CREAM);hb(s,"📒 Page 3: 迷路  Lost",LOST)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(0.95),Inches(4.4),Inches(0.4))
sh.fill.solid();sh.fill.fore_color.rgb=LOST;sh.line.fill.background()
tb(s,0.4,0.98,4.2,0.35,"🎯 主题  Topic",sz=14,b=True,c=WHITE)
tf=tb(s,0.4,1.45,4.4,2.1,"❓ 迷路了, 我会怎么做?",sz=15,b=True,c=DARK)
ap(tf,"   What will I do if I'm lost?",sz=11,c=GRAY)
ap(tf," ",sz=8)
ap(tf,"💡 想一想 / 画一画:",sz=13,b=True,c=LOST)
ap(tf,"·  画一个标志 (大树 / 河 / 路)",sz=12,c=DARK)
ap(tf,"·  画自己 (站在标志旁)",sz=12,c=DARK)
ap(tf,"·  画哨子 / 大喊",sz=12,c=DARK)
sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(0.95),Inches(4.8),Inches(0.4))
sh2.fill.solid();sh2.fill.fore_color.rgb=SUN;sh2.line.fill.background()
tb(s,5.0,0.98,4.6,0.35,"✏️ 内容建议  Content Ideas",sz=14,b=True,c=WHITE)
tf2=tb(s,5.0,1.45,4.7,2.1,"画 1: STOP 🛑 (我停下)",sz=12,c=DARK)
ap(tf2,"画 2: 找一个大树 / 大石头",sz=12,c=DARK)
ap(tf2,"画 3: 吹哨子 3 声 📢📢📢",sz=12,c=DARK)
ap(tf2,"标题: 「迷路了, 我会 ___」",sz=12,b=True,c=LOST)
sh3=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(3.85),Inches(9.4),Inches(1.35))
sh3.fill.solid();sh3.fill.fore_color.rgb=WARM;sh3.line.color.rgb=LOST;sh3.line.width=Pt(2)
tb(s,0.5,3.95,9,0.35,"🗣️ 展示句型  Say These:",sz=14,b=True,c=LOST)
tb(s,0.5,4.35,9,0.35,"·  迷路了, 我会 ___ → ___ → ___ 。",sz=14,c=DARK)
tb(s,0.5,4.70,9,0.35,"·  If lost, I will ___ → ___ → ___ (Stop · Find · Whistle).",sz=14,c=GRAY)
n+=1;pn(s,n)

# 35. POSTER PAGE 4 — Food
s=ns();bg(s,CREAM);hb(s,"📒 Page 4: 食物  Food",OK)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(0.95),Inches(4.4),Inches(0.4))
sh.fill.solid();sh.fill.fore_color.rgb=OK;sh.line.fill.background()
tb(s,0.4,0.98,4.2,0.35,"🎯 主题  Topic",sz=14,b=True,c=WHITE)
tf=tb(s,0.4,1.45,4.4,2.1,"❓ 可以吃 vs 不可以吃?",sz=15,b=True,c=DARK)
ap(tf,"   Safe to eat vs not?",sz=11,c=GRAY)
ap(tf," ",sz=8)
ap(tf,"💡 想一想 / 画一画:",sz=13,b=True,c=OK)
ap(tf,"·  纸的左边 ✅ 能吃",sz=12,c=OK)
ap(tf,"·  纸的右边 ❌ 不能吃",sz=12,c=ALERT)
ap(tf,"·  画或贴 5+ 个食物",sz=12,c=DARK)
sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(0.95),Inches(4.8),Inches(0.4))
sh2.fill.solid();sh2.fill.fore_color.rgb=SUN;sh2.line.fill.background()
tb(s,5.0,0.98,4.6,0.35,"✏️ 内容建议  Content Ideas",sz=14,b=True,c=WHITE)
tf2=tb(s,5.0,1.45,4.7,2.2,"✅ 能吃: 🍎 苹果 / 🍌 香蕉 / 🥜 坚果 / 🍪 能量棒",sz=12,c=OK)
ap(tf2,"❌ 不能吃: 🍄 不认识的蘑菇 / 🫐 陌生莓果",sz=12,c=ALERT)
ap(tf2,"❌ 不能吃: 🌿 不认识的叶子 / 🐛 路边的虫子",sz=12,c=ALERT)
ap(tf2,"标题: 「不认识的, 不要吃!」",sz=12,b=True,c=OK)
sh3=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(3.85),Inches(9.4),Inches(1.35))
sh3.fill.solid();sh3.fill.fore_color.rgb=WARM;sh3.line.color.rgb=OK;sh3.line.width=Pt(2)
tb(s,0.5,3.95,9,0.35,"🗣️ 展示句型  Say These:",sz=14,b=True,c=OK)
tb(s,0.5,4.35,9,0.35,"·  ___ 可以吃, 因为 ___ 。",sz=14,c=DARK)
tb(s,0.5,4.70,9,0.35,"·  ___ 不可以吃, 因为 ___ 。 / I (don't) eat ___ because ___.",sz=14,c=GRAY)
n+=1;pn(s,n)

# 36. CLOSING
s=ns();bg(s,ALERT)
tb(s,1,0.6,8,0.7,"🏆 任务完成!",sz=44,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,1,1.35,8,0.5,"Mission Complete!",sz=20,c=WARM,a=PP_ALIGN.CENTER)
sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.5),Inches(2.0),Inches(3.0),Inches(3.0))
sh.fill.solid();sh.fill.fore_color.rgb=GOLD;sh.line.color.rgb=WHITE;sh.line.width=Pt(4)
tb(s,3.5,2.4,3.0,0.6,"🛡️",sz=80,a=PP_ALIGN.CENTER)
tb(s,3.5,3.6,3.0,0.5,"应急大师",sz=22,b=True,c=ALERT,a=PP_ALIGN.CENTER)
tb(s,3.5,4.1,3.0,0.4,"Emergency Master",sz=12,c=BROWN,a=PP_ALIGN.CENTER)
tb(s,1,5.0,8,0.4,"明天: Day 5 — 探险家展示日 / Day 5: Final Showcase",sz=14,c=WARM,a=PP_ALIGN.CENTER)
n+=1;pn(s,n)

# === Save ===
out="/Users/Huan/projects/summercourse/Chinese/野外生存与探险wilderness_pbl/day4_emergency.pptx"
prs.save(out)
print(f"Saved {out}  ({n} slides)")
