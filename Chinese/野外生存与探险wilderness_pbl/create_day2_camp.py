#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
野外生存与探险 Day 2 — 搭一个小营地 Build a Mini Camp (Camping & Safety)
"Explorer Mission" framing — kids = small explorers building a safe camp.
Each section = a task: pack, build, choose location, protect.
Reuses Day 1 palette + helper conventions.
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

# --- Palette (continuity with Day 1) ---
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
# Camp-zone accents
TENTCL = RGBColor(0x2D,0x5A,0x3D)
FIRECL = RGBColor(0xE6,0x52,0x2C)
FOODCL = RGBColor(0xC9,0x82,0x46)
SAFCL  = RGBColor(0x3E,0x6E,0xB6)

# === Helpers (same conventions as Day 1) ===
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
def hb(s,txt,c=PINE,t=0.15):
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

# === Specialized Day 2 helpers ===

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
    tb(s,2.0,t+0.32,7.6,0.3,frame_en,sz=10,c=GRAY,a=None)

def zone_slide(emoji,name_cn,name_en,color,rules,frame_cn,frame_en):
    """A camp-zone slide: 4 simple safety rules + sentence frame + image placeholder."""
    s=ns();bg(s,CREAM);hb(s,f"{emoji} 营地 4 区域 · {name_cn} {name_en}",color)
    # left: image
    ib(s,0.3,1.0,4.3,3.3,f"📷 {name_cn} 图片 / 简笔画")
    # right: rule cards
    panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.85),Inches(1.0),Inches(4.85),Inches(3.3))
    panel.fill.solid();panel.fill.fore_color.rgb=WHITE;panel.line.color.rgb=color;panel.line.width=Pt(2.5)
    head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.85),Inches(1.0),Inches(4.85),Inches(0.5))
    head.fill.solid();head.fill.fore_color.rgb=color;head.line.fill.background()
    tb(s,5.0,1.07,4.6,0.4,"🤔 安全的 / 不安全的？",sz=14,b=True,c=WHITE)
    y=1.62
    for icon,rule in rules:
        tb(s,5.05,y,0.4,0.35,icon,sz=18)
        tb(s,5.45,y+0.02,4.0,0.35,rule,sz=13,c=DARK);y+=0.6
    sentence_frame_bar(s,4.45,frame_cn,frame_en)
    return s

def ab_slide(title_cn,title_en,question_cn,question_en,a_emoji,a_label,a_caption,b_emoji,b_label,b_caption,answer,reason):
    """A vs B vote slide — students physically move left/right."""
    s=ns();bg(s,CREAM);hb(s,f"📍 {title_cn}  {title_en}",SAFCL)
    # question
    tb(s,0.4,0.85,9.2,0.4,question_cn,sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
    tb(s,0.4,1.20,9.2,0.3,question_en,sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    # A panel (left)
    a_box=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.6),Inches(4.5),Inches(2.4))
    a_box.fill.solid();a_box.fill.fore_color.rgb=WHITE;a_box.line.color.rgb=SAFCL;a_box.line.width=Pt(2.5)
    pill(s,0.5,1.7,0.7,0.4,"A",SAFCL,sz=16)
    tb(s,1.2,1.65,3.6,0.5,a_label,sz=18,b=True,c=SAFCL)
    tb(s,0.6,2.2,4.2,0.5,a_emoji,sz=44,a=PP_ALIGN.CENTER)
    tb(s,0.6,3.1,4.2,0.6,a_caption,sz=13,c=DARK,a=PP_ALIGN.CENTER)
    # B panel (right)
    b_box=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.1),Inches(1.6),Inches(4.5),Inches(2.4))
    b_box.fill.solid();b_box.fill.fore_color.rgb=WHITE;b_box.line.color.rgb=ALERT;b_box.line.width=Pt(2.5)
    pill(s,5.2,1.7,0.7,0.4,"B",ALERT,sz=16)
    tb(s,5.9,1.65,3.6,0.5,b_label,sz=18,b=True,c=ALERT)
    tb(s,5.3,2.2,4.2,0.5,b_emoji,sz=44,a=PP_ALIGN.CENTER)
    tb(s,5.3,3.1,4.2,0.6,b_caption,sz=13,c=DARK,a=PP_ALIGN.CENTER)
    # vote prompt
    sentence_frame_bar(s,4.15,
        "我选 ___, 因为 ____  /  我应该 ____",
        "I choose A/B because… / I should…")
    # Footer hint
    tb(s,0.4,4.92,9.2,0.3,"👉 走到 A 边 / B 边 — 用一句话说为什么。",sz=12,b=True,c=SUN,a=PP_ALIGN.CENTER)
    notes(s,f"老师备课:\n• 答案 / Answer: {answer}\n• 原因 / Reason: {reason}\n• 玩法: 把教室分成 A/B 两边, 学生走到选择的一边, 然后每边请 1-2 个孩子说出原因。\n• K 级: 「A 安全」「B 不安全」即可。\n• G1-3: 「我选 ___, 因为 ___」整句。\n• 教师可在视频/图片资源里截图替换 emoji。")
    return s

def video_slide(title_cn,title_en,before_task,after_action,bgc=SAFCL):
    """Video slide with before-watching task + after-watching action."""
    s=ns();bg(s,bgc)
    tb(s,1,0.55,8,0.7,"🎬 看视频 Watch Video",sz=32,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,1,1.25,8,0.4,f"{title_cn}",sz=20,b=True,c=WARM,a=PP_ALIGN.CENTER)
    tb(s,1,1.6,8,0.3,title_en,sz=12,c=WARM,a=PP_ALIGN.CENTER)
    # Before
    pre=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(2.05),Inches(4.5),Inches(2.0))
    pre.fill.solid();pre.fill.fore_color.rgb=WHITE;pre.line.fill.background()
    tb(s,0.55,2.15,4.3,0.4,"👂 看之前 Before Watching",sz=14,b=True,c=bgc)
    tb(s,0.55,2.55,4.3,1.4,before_task,sz=12,c=DARK)
    # After
    post=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.1),Inches(2.05),Inches(4.5),Inches(2.0))
    post.fill.solid();post.fill.fore_color.rgb=WHITE;post.line.fill.background()
    tb(s,5.25,2.15,4.3,0.4,"🎯 看完后 After Watching",sz=14,b=True,c=SUN)
    tb(s,5.25,2.55,4.3,1.4,after_action,sz=12,c=DARK)
    # Link placeholder
    tb(s,1,4.3,8,0.4,"🔗 视频链接 (老师粘贴) / Teacher pastes video link",sz=12,c=WARM,a=PP_ALIGN.CENTER)
    return s

def word_card_read(w,py,en,sent,img,color=SUN):
    s=ns();bg(s,CREAM);hb(s,"👀 我会认  I Can Read",color)
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.0),Inches(4.5),Inches(2.5))
    sh.fill.solid();sh.fill.fore_color.rgb=WARM;sh.line.fill.background()
    tb(s,0.5,1.1,4.3,1.4,w,sz=72,b=True,c=PINE,a=PP_ALIGN.CENTER)
    tb(s,0.5,2.4,4.3,0.4,f"{py}  {en}",sz=20,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,0.5,2.85,4.3,0.4,"👉 跟我读！Read after me!",sz=14,c=color,a=PP_ALIGN.CENTER)
    ib(s,5.3,1.0,4.4,2.5,img)
    sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(3.8),Inches(9.2),Inches(1.2))
    sh2.fill.solid();sh2.fill.fore_color.rgb=WHITE;sh2.line.color.rgb=color;sh2.line.width=Pt(2)
    tb(s,0.6,3.9,1.5,0.4,"例句 Example",sz=14,b=True,c=color)
    tb(s,0.6,4.3,8.8,0.5,sent,sz=20,b=True,c=DARK)
    return s

def word_card_write(w,py,en,strokes_hint,color=PINE):
    s=ns();bg(s,CREAM);hb(s,"✍️ 我会写  I Can Write",color)
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.0),Inches(4.5),Inches(2.0))
    sh.fill.solid();sh.fill.fore_color.rgb=WARM;sh.line.fill.background()
    tb(s,0.5,1.05,4.3,1.2,w,sz=72,b=True,c=color,a=PP_ALIGN.CENTER)
    tb(s,0.5,2.2,4.3,0.4,f"{py}  {en}",sz=20,c=GRAY,a=PP_ALIGN.CENTER)
    # Stroke order area
    tb(s,0.6,3.4,4.6,0.4,"📝 笔顺 Stroke Order",sz=16,b=True,c=color)
    tb(s,0.6,3.8,4.6,1.2,strokes_hint,sz=14,c=DARK)
    # Practice steps
    tb(s,5.8,3.4,3.8,0.4,"练习步骤 Practice:",sz=14,b=True,c=color)
    tb(s,5.8,3.8,3.8,1.2,"1. 看老师写\n2. 用手指空中写\n3. 在本子上写 3 次",sz=13,c=DARK)
    # Right: 田字格 boxes (4 squares)
    for i in range(4):
        sq=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(5.8+i*0.95),Inches(1.0),Inches(0.85),Inches(0.85))
        sq.fill.solid();sq.fill.fore_color.rgb=WHITE;sq.line.color.rgb=color;sq.line.width=Pt(1.5)
        # cross hair
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
s=ns();bg(s,PINE)
sh=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,Inches(2.4),W,Inches(2.0))
sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.fill.background()
tb(s,1,0.4,8,0.5,"DAY 2",sz=18,b=True,c=SUN,a=PP_ALIGN.CENTER)
tb(s,1,0.95,8,0.7,"🏕️ 搭一个小营地",sz=42,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,1,1.7,8,0.5,"Build a Mini Camp",sz=22,c=WARM,a=PP_ALIGN.CENTER)
tb(s,1,2.6,8,0.5,"🧭 探险家任务  Explorer Mission",sz=24,b=True,c=PINE,a=PP_ALIGN.CENTER)
tb(s,1,3.15,8,0.4,"Pack · Build · Choose · Protect",sz=14,c=BROWN,a=PP_ALIGN.CENTER)
tb(s,1,3.55,8,0.4,"装包 · 搭建 · 选址 · 保护",sz=14,b=True,c=BROWN,a=PP_ALIGN.CENTER)
tb(s,1,4.6,8,0.4,"野外生存与探险 · Wilderness Survival",sz=14,b=True,c=SUN,a=PP_ALIGN.CENTER)
n+=1;pn(s,n)
notes(s,"开场 (1 分钟):\n• 「探险家们, 准备好了吗？」\n• 今天我们不是学生 — 是「小探险家」, 要建一个安全的营地。\n• 4 个任务: 装包 / 搭建 / 选址 / 保护。\n• 完成所有任务可以拿到「探险家徽章」(教师可打印或盖章)。")

# 2. EXPLORER MISSION — 4 tasks
s=ns();bg(s,CREAM);hb(s,"🧭 今天的任务  Today's Mission",PINE)
tb(s,0.4,0.9,9.2,0.4,"小探险家们 — 我们要一起做 4 件事!",sz=20,b=True,c=PINE,a=PP_ALIGN.CENTER)
tb(s,0.4,1.32,9.2,0.3,"Little explorers — we'll do 4 jobs together!",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
mission_card(s,0.4,1.85,2.25,2.6,1,"装包",  "Pack a backpack",   "🎒",SUN)
mission_card(s,2.75,1.85,2.25,2.6,2,"搭建", "Build the camp",    "🏕️",TENTCL)
mission_card(s,5.10,1.85,2.25,2.6,3,"选址", "Choose location",   "📍",SAFCL)
mission_card(s,7.45,1.85,2.25,2.6,4,"保护", "Stay safe",         "🛡️",ALERT)
sentence_frame_bar(s,4.65,
    "我是小探险家, 我可以…",
    "I am a little explorer, I can…")
n+=1;pn(s,n)
notes(s,"播放节奏:\n• 把 4 个任务说一遍, 让学生跟读「装包-搭建-选址-保护」。\n• 「我们今天有 4 个任务, 完成 1 个打 1 个勾。」\n• 在白板上画 4 格, 每完成一个任务画 ✓。\n• 探险家徽章可以是贴纸 / 印章 / 自制纸徽章。")

# 3. EXPLORER PLEDGE
s=ns();bg(s,CREAM);hb(s,"🤝 探险家宣言  Explorer Pledge",PINE)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(1.0),Inches(1.1),Inches(8.0),Inches(3.2))
sh.fill.solid();sh.fill.fore_color.rgb=WARM;sh.line.color.rgb=PINE;sh.line.width=Pt(2.5)
tb(s,1.2,1.3,7.6,0.5,"🌲 我是小探险家。",sz=22,b=True,c=PINE,a=PP_ALIGN.CENTER)
tb(s,1.2,1.85,7.6,0.4,"I am a little explorer.",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,1.2,2.35,7.6,0.4,"🛡️ 我会安全。",sz=22,b=True,c=ALERT,a=PP_ALIGN.CENTER)
tb(s,1.2,2.85,7.6,0.4,"I will be safe.",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,1.2,3.35,7.6,0.4,"🤝 我会帮助朋友。",sz=22,b=True,c=SUN,a=PP_ALIGN.CENTER)
tb(s,1.2,3.85,7.6,0.4,"I will help my friends.",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,0.4,4.5,9.2,0.5,"👉 一起举手念 — Raise your hand and say it together!",sz=14,b=True,c=SUN,a=PP_ALIGN.CENTER)
n+=1;pn(s,n)
notes(s,"全班一起举右手念这 3 句话。可以让 1-2 个孩子领读。")

# 4. SESSION 1 DIVIDER
s=div("Session 1  上午","🦊 Story · Zones · Pack · Locate · Watch",PINE,"📖");n+=1;pn(s,n)

# 5. STORY pre-watching
s=ns();bg(s,CREAM);hb(s,"🐵 听之前 Before Watching · Curious George Goes Camping",SUN)
tb(s,0.4,0.9,9.2,0.4,"在乔治的故事里, 我们要找 3 件事 —",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.3,"In the George story, listen for 3 things —",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# 3 task cards
for i,(em,cn,en) in enumerate([("👂","乔治带了什么？","What does George pack?"),
                                ("👀","谁在营地里？","Who is at the camp?"),
                                ("🤔","乔治做对了吗？","Does George do it right?")]):
    x=0.4+i*3.15
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.85),Inches(3.0),Inches(2.5))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=SUN;sh.line.width=Pt(2)
    tb(s,x+0.1,1.95,2.8,0.7,em,sz=42,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,2.7,2.8,0.5,cn,sz=16,b=True,c=PINE,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,3.2,2.8,0.5,en,sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,3.7,2.8,0.6,"举手 ✋\nRaise hand!",sz=12,b=True,c=SUN,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.55,
    "乔治带了 ___ 和 ___。",
    "George packs ___ and ___.")
n+=1;pn(s,n)
notes(s,"看视频前 1 分钟设定任务:\n• 把 3 张卡片读一遍, 让学生记住自己要找什么。\n• 「看到答案的时候, 举手!」(老师可暂停视频)\n• 这样学生从「被动看」变「主动找」。")

# 6. VIDEO — Curious George
s=video_slide(
    "🐵 Curious George Goes Camping",
    "好奇的乔治去露营",
    "👂 听 / 看:\n1. 乔治带了什么？\n2. 营地有谁？\n3. 乔治哪里做错了？\n\n🛑 在看到「错」的时候, 老师按暂停!",
    "🎯 看完后 (二选一):\n• 演一演: 谁演乔治？谁演动物？\n• 说一说: 「乔治应该 ____」\n• 决定: 乔治可以再来吗?",
    bgc=SUN);n+=1;pn(s,n)
notes(s,"视频建议: YouTube 搜索「Curious George Goes Camping」(动画版, 约 11 分钟; 也有 1-2 分钟剪辑版本)。\n• 关键暂停点: 乔治把食物留在桌上 / 乔治玩火 / 乔治追小动物。\n• 暂停时问: 「这样可以吗？为什么不可以？」")

# 7. POST-WATCH DECISION
s=ns();bg(s,CREAM);hb(s,"🎯 看完后 — 你来决定 You Decide",SUN)
tb(s,0.4,0.9,9.2,0.4,"乔治做对了吗？  Did George do it right?",sz=20,b=True,c=DARK,a=PP_ALIGN.CENTER)
# Three columns: SPEAK / ACT / DECIDE
cols=[("🗣️","说一说","Speak","乔治应该 ___ 。\nGeorge should ___."),
      ("🎭","演一演","Act","演一个「错的」乔治, 一个「对的」乔治。\nAct out wrong vs right George."),
      ("✋","举手投票","Vote","可以 / 不可以 让乔治再来？\nCan George come back?")]
for i,(em,cn,en,desc) in enumerate(cols):
    x=0.4+i*3.15
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.5),Inches(3.0),Inches(2.7))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=PINE;sh.line.width=Pt(2)
    pill(s,x+0.1,1.6,1.0,0.35,en,PINE,sz=11)
    tb(s,x+0.1,2.0,2.8,0.7,em,sz=36,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,2.75,2.8,0.4,cn,sz=18,b=True,c=PINE,a=PP_ALIGN.CENTER)
    tb(s,x+0.15,3.2,2.7,0.95,desc,sz=11,c=DARK)
sentence_frame_bar(s,4.4,
    "乔治不安全, 因为 ____ 。我应该 ____ 。",
    "George wasn't safe because ___. I should ___.")
n+=1;pn(s,n)
notes(s,"3 选 1 — 时间 5 分钟:\n• 让学生选一种方式响应 (说 / 演 / 投票)。\n• 演 — 邀请 2 个志愿者演 30 秒小情景。\n• 投票 — 喊「可以」「不可以」, 数举手。")

# 8. CAMP ZONES INTRO — 4 zones
s=ns();bg(s,CREAM);hb(s,"🏕️ 营地有 4 个区域  Camp Has 4 Zones",PINE)
tb(s,0.4,0.9,9.2,0.4,"一个安全的营地 = 4 个区域要分开!",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.28,9.2,0.3,"A safe camp = 4 zones, separated!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
zones=[("🏕️","帐篷","Tent",TENTCL),
       ("🔥","火",  "Fire",FIRECL),
       ("🍱","食物","Food",FOODCL),
       ("🛡️","安全","Safety",SAFCL)]
for i,(em,cn,en,c) in enumerate(zones):
    x=0.4+i*2.35
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.75),Inches(2.20),Inches(2.5))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=c;sh.line.width=Pt(2.5)
    tb(s,x+0.05,1.85,2.1,0.9,em,sz=52,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.85,2.1,0.5,cn,sz=22,b=True,c=c,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.4,2.1,0.4,en,sz=12,c=GRAY,a=PP_ALIGN.CENTER)
    pill(s,x+0.55,3.85,1.1,0.3,f"区域 {i+1}",c,sz=10)
sentence_frame_bar(s,4.4,
    "营地有 ___ 个区域。每个区域 ___ 。",
    "The camp has ___ zones. Each zone ___.")
n+=1;pn(s,n)
notes(s,"重点: 这 4 个区域要分开!\n• 帐篷 — 睡觉的地方\n• 火 — 取暖、烧饭\n• 食物 — 吃东西、收食物\n• 安全 — 厕所、急救、瞭望\n• 让学生用手指数 1-2-3-4。")

# 9-12. ZONE SLIDES
n+=1;s=zone_slide("🏕️","帐篷区","Tent Zone",TENTCL,
    [("✅","平的地方  Flat ground"),
     ("✅","干的地方  Dry ground"),
     ("❌","不要在大树下  Not under big trees"),
     ("❌","不要离火太近  Not near fire")],
    "帐篷可以 ___ , 不可以 ___ 。",
    "Tent can ___, not ___.")
pn(s,n)
notes(s,"为什么不在大树下？枯枝可能掉下来 (widowmaker)。为什么干？避免地里湿气和洪水。")

n+=1;s=zone_slide("🔥","火区","Fire Zone",FIRECL,
    [("✅","用石头围起来  Surround with stones"),
     ("✅","旁边有水  Water nearby"),
     ("❌","不要在树下  Not under trees"),
     ("❌","不要一个人离开  Don't leave alone")],
    "火 ___ 安全, ___ 不安全。",
    "Fire is safe when ___, not safe when ___.")
pn(s,n)
notes(s,"大人在场才能用火。火堆离帐篷至少 4 米。睡前/离开前一定要灭火。")

n+=1;s=zone_slide("🍱","食物区","Food Zone",FOODCL,
    [("✅","食物挂在树上  Hang food in a tree"),
     ("✅","食物放盒子里  Put food in a box"),
     ("❌","不要放在帐篷里  Not in the tent!"),
     ("❌","不要丢在地上  Don't leave on ground")],
    "食物要 ___ , 不能 ___ 。",
    "Food should ___, not ___.")
pn(s,n)
notes(s,"为什么不放在帐篷？熊和动物会找到食物的味道, 进帐篷很危险。这叫「Bear Bag」, 把食物吊在远离帐篷的树枝上。")

n+=1;s=zone_slide("🛡️","安全区","Safety Zone",SAFCL,
    [("✅","厕所离帐篷远一点  Bathroom far from tent"),
     ("✅","急救包放好  First aid ready"),
     ("✅","知道方向  Know directions"),
     ("✅","跟大人在一起  Stay with adults")],
    "我应该 ___ , 因为 ___ 。",
    "I should ___, because ___.")
pn(s,n)
notes(s,"急救包: 创可贴、消毒湿巾、水、哨子。哨子是迷路时的工具!")

# 13. WRONG CAMP ACTIVITY INTRO — using real image
s=ns();bg(s,CREAM);hb(s,"🔍 这里能搭营吗？  Can we camp HERE?",ALERT)
tb(s,0.4,0.9,9.2,0.4,"看一看 — 这个地方搭营好不好？为什么？",sz=18,b=True,c=ALERT,a=PP_ALIGN.CENTER)
tb(s,0.4,1.27,9.2,0.3,"Look — is this a good spot? Why or why not?",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# Real image: D2_露营地点选择.png (1402x1122 ≈ 5:4). Fit into 4.6"w × 3.0"h on left.
import os
_img_path="/Users/Huan/projects/summercourse/Chinese/野外生存与探险wilderness_pbl/D2_露营地点选择.png"
if os.path.exists(_img_path):
    s.shapes.add_picture(_img_path,Inches(0.3),Inches(1.6),Inches(5.6),Inches(3.05))
else:
    ib(s,0.3,1.6,5.6,3.05,"📷 D2_露营地点选择.png")
# Right: question prompts
panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(6.05),Inches(1.6),Inches(3.65),Inches(3.05))
panel.fill.solid();panel.fill.fore_color.rgb=WHITE;panel.line.color.rgb=ALERT;panel.line.width=Pt(2.5)
head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(6.05),Inches(1.6),Inches(3.65),Inches(0.45))
head.fill.solid();head.fill.fore_color.rgb=ALERT;head.line.fill.background()
tb(s,6.2,1.66,3.45,0.35,"🤔 想一想 — 为什么不好？",sz=13,b=True,c=WHITE)
tf=tb(s,6.2,2.15,3.4,0.35,"❓ 帐篷搭在哪里？",sz=12,b=True,c=DARK)
ap(tf," ",sz=6)
ap(tf,"❓ 离水太近吗？",sz=12,b=True,c=DARK)
ap(tf," ",sz=6)
ap(tf,"❓ 上面有什么 (大树、岩石)?",sz=12,b=True,c=DARK)
ap(tf," ",sz=6)
ap(tf,"❓ 地平不平？(斜坡?)",sz=12,b=True,c=DARK)
ap(tf," ",sz=6)
ap(tf,"❓ 下大雨会怎样？",sz=12,b=True,c=DARK)
# Sentence frame at bottom
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(4.78),Inches(9.4),Inches(0.5))
sh.fill.solid();sh.fill.fore_color.rgb=WARM;sh.line.color.rgb=SUN;sh.line.width=Pt(2)
tb(s,0.5,4.85,9.0,0.35,"💬 这里 ___ 安全, 因为 ___ 。/ This place is (not) safe, because ___.",sz=12,b=True,c=DARK)
n+=1;pn(s,n)
notes(s,"老师玩法 (5 分钟):\n• 让学生先 30 秒自己看图, 找问题\n• 然后举手说: 「我觉得 ___ 不好, 因为 ___」\n• 把答案写在白板上, 数一数找到几个问题。\n• 关键问题示例:\n  - 帐篷太靠近水/河边 (会被洪水)\n  - 帐篷在斜坡上 (睡觉会滚)\n  - 帐篷在大树下 (枯枝)\n  - 火堆太靠近帐篷\n  - 食物放在外面没收\n• 找到 3+ 个 = 探险家 1 级 ⭐")

# 14. WRONG CAMP — answer reveal
s=ns();bg(s,CREAM);hb(s,"💡 答案揭晓  The 5 Mistakes",ALERT)
mistakes=[
    ("🏕️🌳","帐篷在大树下 — 枯枝会掉",   "Tent under a big tree — branches can fall"),
    ("🔥🏕️","火堆太靠近帐篷",         "Fire too close to tent"),
    ("🍱🏕️","食物放在帐篷里 — 熊会来", "Food in tent — bears!"),
    ("👤🔥","没有人看着火",            "Nobody is watching the fire"),
    ("🏕️📐","帐篷在斜坡上",           "Tent on a slope"),
]
for i,(em,cn,en) in enumerate(mistakes):
    y=1.0+i*0.78
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(y),Inches(9.2),Inches(0.7))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=ALERT;sh.line.width=Pt(2)
    pill(s,0.55,y+0.15,0.5,0.4,str(i+1),ALERT,sz=14)
    tb(s,1.15,y+0.05,1.5,0.6,em,sz=24)
    tb(s,2.7,y+0.08,5.4,0.3,cn,sz=14,b=True,c=DARK)
    tb(s,2.7,y+0.38,5.4,0.3,en,sz=10,c=GRAY)
n+=1;pn(s,n)
notes(s,"对一对答案。每个错误让 1 个孩子用「___ 不安全, 因为 ___」造句。")

# 15. MOVEMENT GAME — I am a camp object
s=ns();bg(s,CREAM);hb(s,"🚶‍♀️ 我是营地零件  I Am a Camp Object",TENTCL)
tb(s,0.4,0.85,9.2,0.4,"老师叫到「火」, 你就跑到「火区」!",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.23,9.2,0.3,"Teacher says \"fire,\" you run to the FIRE zone!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# 4 corners of room
zones2=[("🏕️","帐篷区","Tent",TENTCL,1.0),
        ("🔥","火区","Fire",FIRECL,3.4),
        ("🍱","食物区","Food",FOODCL,5.8),
        ("🛡️","安全区","Safety",SAFCL,8.2)]
for em,cn,en,c,x in zones2:
    sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x-0.6),Inches(1.85),Inches(1.5),Inches(1.5))
    sh.fill.solid();sh.fill.fore_color.rgb=c;sh.line.color.rgb=WHITE;sh.line.width=Pt(2)
    tb(s,x-0.6,2.2,1.5,0.6,em,sz=36,a=PP_ALIGN.CENTER)
    tb(s,x-0.6,3.4,1.5,0.3,cn,sz=12,b=True,c=c,a=PP_ALIGN.CENTER)
    tb(s,x-0.6,3.7,1.5,0.3,en,sz=9,c=GRAY,a=PP_ALIGN.CENTER)
# rules box
rule=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(4.2),Inches(9.2),Inches(0.95))
rule.fill.solid();rule.fill.fore_color.rgb=WARM;rule.line.color.rgb=PINE;rule.line.width=Pt(2)
tb(s,0.6,4.3,8.8,0.3,"🎯 规则 Rules:  老师喊词 → 学生跑到对应的角落 → 慢的人出局",sz=12,b=True,c=PINE)
tb(s,0.6,4.6,8.8,0.3,"Teacher calls a word → run to the corner → slow ones sit out",sz=11,c=GRAY)
tb(s,0.6,4.85,8.8,0.3,"💡 进阶: 喊错的物品 (e.g. 帐篷 + 火) 站在 2 个角落中间表示「太近了, 不安全」",sz=10,c=ALERT)
n+=1;pn(s,n)
notes(s,"老师准备 (低投入):\n• 在教室 4 个角落贴 4 张纸 (帐篷/火/食物/安全)。\n• 喊词游戏 — 玩 5-8 轮。\n• 进阶: 喊「帐篷 + 火」让 2 组学生面对面, 然后让他们说「太近了!」\n• 也可让学生戴贴纸 (我是火/我是帐篷) 自己找位置。")

# 17. BACKPACK — basics
s=ns();bg(s,CREAM);hb(s,"🎒 装什么进背包？  What goes in the backpack?",SUN)
items=[("💧","水","Water"),("🍎","食物","Food"),("🔦","手电筒","Flashlight"),
       ("🧥","外套","Jacket"),("🆘","急救包","First aid"),("🗺️","地图","Map"),
       ("📢","哨子","Whistle"),("🧴","防晒","Sunscreen"),("🧦","袜子","Socks")]
for i,(em,cn,en) in enumerate(items):
    col=i%5;row=i//5
    x=0.4+col*1.92;y=1.0+row*1.7
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(1.78),Inches(1.55))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=SUN;sh.line.width=Pt(2)
    tb(s,x+0.05,y+0.05,1.7,0.65,em,sz=36,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,y+0.78,1.7,0.4,cn,sz=15,b=True,c=PINE,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,y+1.18,1.7,0.3,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.5,
    "我带 ___ 和 ___ 。我不带 ___ 。",
    "I bring ___ and ___. I don't bring ___.")
n+=1;pn(s,n)
notes(s,"先看 9 个东西, 让学生熟悉。下一页是 30 秒挑战!")

# 18. BACKPACK — 30s timer challenge
s=ns();bg(s,SUN)
tb(s,1,0.5,8,0.7,"⏱️ 30 秒挑战!",sz=44,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,1,1.25,8,0.4,"30-Second Challenge",sz=18,c=WARM,a=PP_ALIGN.CENTER)
# Big timer
sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.5),Inches(2.0),Inches(3.0),Inches(3.0))
sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=ALERT;sh.line.width=Pt(6)
tb(s,3.5,2.7,3.0,1.0,"30",sz=110,b=True,c=ALERT,a=PP_ALIGN.CENTER)
tb(s,3.5,3.85,3.0,0.4,"秒  seconds",sz=18,b=True,c=PINE,a=PP_ALIGN.CENTER)
# Instructions on sides
tb(s,0.4,2.0,2.9,0.4,"🎯 任务 Task:",sz=16,b=True,c=WHITE)
tb(s,0.4,2.4,2.9,1.5,"选出你最重要的 5 件东西!\nPick your TOP 5 items.\n\n👉 圈一圈 / 在卡片上贴星",sz=12,c=WHITE)
tb(s,6.7,2.0,2.9,0.4,"⚠️ 规则 Rules:",sz=16,b=True,c=WHITE)
tb(s,6.7,2.4,2.9,1.5,"30 秒到了 — 停!\nWhen time's up — stop!\n\n👉 跟同桌比一比",sz=12,c=WHITE)
n+=1;pn(s,n)
notes(s,"老师准备:\n• 印 9 张物品小卡片 / 让学生用上一页, 圈 5 个。\n• 用手机倒数 30 秒。\n• 时间到 — 让 2-3 个学生说「我选 ___ 因为 ___」。\n• 没有正确答案, 但要鼓励「水 + 食物 + 急救」是必带。")

# 19-21. SITUATION CHANGE — desert / forest rain / cold mountain
def situation_slide(em,situation_cn,situation_en,color,must_have,why):
    s=ns();bg(s,CREAM);hb(s,f"{em} 情况变了!  Situation Change!",color)
    big=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.95),Inches(9.2),Inches(1.4))
    big.fill.solid();big.fill.fore_color.rgb=color;big.line.fill.background()
    tb(s,0.6,1.05,8.8,0.6,situation_cn,sz=26,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,0.6,1.65,8.8,0.4,situation_en,sz=14,c=WARM,a=PP_ALIGN.CENTER)
    tb(s,0.6,2.0,8.8,0.3,"👉 你要带什么？  What do you pack now?",sz=14,b=True,c=WARM,a=PP_ALIGN.CENTER)
    # 3 items
    for i,(emi,cn,en) in enumerate(must_have):
        x=0.6+i*3.05
        sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(2.55),Inches(2.85),Inches(1.55))
        sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=color;sh.line.width=Pt(2)
        tb(s,x+0.05,2.6,2.75,0.7,emi,sz=44,a=PP_ALIGN.CENTER)
        tb(s,x+0.05,3.4,2.75,0.4,cn,sz=18,b=True,c=color,a=PP_ALIGN.CENTER)
        tb(s,x+0.05,3.85,2.75,0.3,en,sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    sentence_frame_bar(s,4.25,
        "在 ___ 我带 ___ 。",
        "In ___, I bring ___.")
    return s

s=situation_slide("🏜️","沙漠 — 太阳很大, 没有水","Desert — strong sun, no water",DESERT_DUMMY := RGBColor(0xD4,0xA5,0x74),
    [("💧","多带水","Extra water"),("🧴","防晒","Sunscreen"),("🧢","帽子","Hat")],
    "")
n+=1;pn(s,n)
notes(s,"沙漠 — 重点是水! 1 天每人需要 4 升水。皮肤防晒。中午躲太阳。")

s=situation_slide("🌧️","森林 + 下雨","Forest + raining",SAFCL,
    [("🧥","雨衣","Raincoat"),("👢","胶鞋","Boots"),("🔥","防水火柴","Waterproof matches")],
    "")
n+=1;pn(s,n)
notes(s,"下雨 — 鞋子和外衣都要防水。湿衣服 = 失温危险。")

s=situation_slide("🏔️","高山 — 很冷","High mountain — very cold",RGBColor(0x54,0x6E,0x7A),
    [("🧤","手套","Gloves"),("🧣","围巾","Scarf"),("🍫","高能量食物","High-energy food")],
    "")
n+=1;pn(s,n)
notes(s,"冷 — 多层衣服比一件厚衣服好 (layering)。手套和围巾保护手指和脖子。")

# 22. RE-DECIDE reflection
s=ns();bg(s,CREAM);hb(s,"🤔 再想一想 Re-decide",SUN)
tb(s,0.4,0.9,9.2,0.4,"地方不一样, 装的东西也不一样!",sz=20,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.3,"Different place → different pack!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# table 3 cols
heads=[("🏜️","沙漠"),("🌧️","森林雨"),("🏔️","山地冷")]
for i,(em,name) in enumerate(heads):
    x=0.4+i*3.15;
    pill(s,x,1.85,3.0,0.45,f"{em}  {name}",SAFCL,sz=14)
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(2.4),Inches(3.0),Inches(2.0))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=SAFCL;sh.line.width=Pt(2)
    tb(s,x+0.15,2.5,2.7,0.4,"🥇 必带:",sz=12,b=True,c=PINE)
    tb(s,x+0.15,2.85,2.7,0.4,["水 / 防晒 / 帽子","雨衣 / 胶鞋 / 火柴","手套 / 围巾 / 食物"][i],sz=11,c=DARK)
    tb(s,x+0.15,3.4,2.7,0.4,"❌ 不必带:",sz=12,b=True,c=ALERT)
    tb(s,x+0.15,3.75,2.7,0.4,["雨衣","防晒","防晒"][i],sz=11,c=DARK)
sentence_frame_bar(s,4.55,
    "在 ___ 我带 ___ , 不带 ___ 。",
    "In ___, I bring ___ but NOT ___.")
n+=1;pn(s,n)
notes(s,"对比 3 种环境, 让学生发现「装包 = 看情况」。")

# 23-26. TENT LOCATION A/B (4 rounds)
s=ab_slide("选址 1: 在哪里搭帐篷？","Location 1: where to pitch the tent?",
    "你想在 A 还是 B 搭帐篷？","Where would you pitch the tent — A or B?",
    "🟢","平地","平的, 没有大石头",
    "⛰️","斜坡","斜的, 一边高一边低",
    "A 平地", "睡在斜坡上会滚下去, 帐篷也站不稳。")
n+=1;pn(s,n)

s=ab_slide("选址 2: 离河多远？","Location 2: how far from the river?",
    "在 A 还是 B 搭帐篷？","Pitch the tent at A or B?",
    "💧","河边 (3 米)","就在河旁边",
    "🌳","远离河 (30 米)","远一点的高地",
    "B 远离河", "下大雨河水会上来 (洪水), 河边 3 米的帐篷会被淹。")
n+=1;pn(s,n)

s=ab_slide("选址 3: 在大树下还是空地?","Location 3: under tree or open?",
    "选 A 还是 B?","Choose A or B?",
    "🌲","大树下","有阴凉但有枯枝",
    "🌿","空地","没遮挡但安全",
    "B 空地", "大树底下的枯枝叫「widowmaker」, 风一吹会掉下来。下雨打雷时, 大树会引雷。")
n+=1;pn(s,n)

s=ab_slide("选址 4: 离火堆多远?","Location 4: how far from fire?",
    "选 A 还是 B?","Choose A or B?",
    "🏕️🔥","靠近火 (1 米)","暖和, 但火星会飞过来",
    "🏕️➡️🔥","离火 4 米","暖和靠走过去, 不会着火",
    "B 离火 4 米", "火星可能飞 1-2 米, 把帐篷烧着。安全距离 = 4 米以上。")
n+=1;pn(s,n)

# 27. REASONING SUMMARY
s=ns();bg(s,CREAM);hb(s,"🧠 你说得对!  You Got It!",PINE)
tb(s,0.4,0.9,9.2,0.4,"4 个选址规则:  4 location rules:",sz=20,b=True,c=DARK,a=PP_ALIGN.CENTER)
rules=[("🟢 1.","平地 — 不要斜坡","Flat — no slope"),
       ("💧 2.","离河 30 米 — 不要被洪水冲","30m from river — avoid floods"),
       ("🌿 3.","空地 — 不在大树下","Open — not under big trees"),
       ("🔥 4.","离火堆 4 米以上","4m+ from the fire")]
for i,(em,cn,en) in enumerate(rules):
    y=1.5+i*0.78
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(y),Inches(9.2),Inches(0.7))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=PINE;sh.line.width=Pt(2)
    tb(s,0.6,y+0.1,1.4,0.5,em,sz=20,b=True,c=PINE)
    tb(s,2.0,y+0.05,5.4,0.35,cn,sz=15,b=True,c=DARK)
    tb(s,2.0,y+0.4,5.4,0.3,en,sz=11,c=GRAY)
n+=1;pn(s,n)

# 28. SESSION 2.5 — Video interactions divider
s=div("🎬 看视频 学探险","Watch & Learn  ·  Camping setup · Safe campsite · Backpack",SAFCL,"📺");n+=1;pn(s,n)

# 29. VIDEO 1 — Camping setup
s=video_slide("How to Set Up a Tent (Kids)","怎么搭帐篷",
    "👂 听 / 看:\n1. 搭帐篷的第一步是什么？\n2. 一共有几步？\n3. 需要几个人?",
    "🎯 看完后:\n• 演一演: 跟同学一起「假装搭帐篷」\n• 数一数: 共 ___ 步\n• 说一说: 第一步是 ___",
    bgc=PINE);n+=1;pn(s,n)
notes(s,"视频建议: YouTube 搜「How to set up a tent for kids」(选 1-3 分钟版本)。\n关键步骤: 1) 选地方 2) 铺地布 3) 立支架 4) 拉绳子。")

# 30. VIDEO 2 — Safe vs unsafe campsite
s=video_slide("Safe vs Unsafe Campsite","安全 vs 不安全的营地",
    "👂 听 / 看:\n1. 这个营地哪里安全？\n2. 哪里不安全？\n3. 你能找到 3 个错？",
    "🎯 看完后:\n• 圈一圈: 在 worksheet 上画 ✅ 和 ❌\n• 说一说: 「___ 不安全, 因为 ___」\n• 决定: 让乔治再来吗？",
    bgc=ALERT);n+=1;pn(s,n)
notes(s,"视频建议: YouTube 搜「camping safety for kids」或 Smokey Bear 的视频。或老师自己拍 30 秒手机视频。")

# 31. VIDEO 3 — Packing backpack
s=video_slide("Packing a Backpack","怎么装背包",
    "👂 听 / 看:\n1. 哪些东西先放进去？\n2. 哪些放在最上面？\n3. 重的放哪里？",
    "🎯 看完后:\n• 排顺序: 1)___ 2)___ 3)___\n• 说一说: 重的东西放 ___\n• 比一比: 跟你装的一样吗?",
    bgc=SUN);n+=1;pn(s,n)
notes(s,"原则: 重的放中间靠背 (重心稳)。睡袋放下面。常用的 (水/急救) 放最上。")

# 32. SESSION 2 DIVIDER
s=div("Session 2  下午","🔄 复习 + 语言目标 (我会认/我会写)",PINE,"📖");n+=1;pn(s,n)

# 33. SENTENCE FRAMES — 2 levels
s=ns();bg(s,CREAM);hb(s,"💬 句型卡  Sentence Frames",SUN)
# K column
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.0),Inches(4.5),Inches(3.6))
sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=SUN;sh.line.width=Pt(2.5)
pill(s,0.6,1.15,1.6,0.4,"K (TK-K)",SUN,sz=14)
tb(s,0.6,1.7,4.1,0.5,"✅ 可以 / ❌ 不可以",sz=22,b=True,c=PINE)
tb(s,0.6,2.2,4.1,0.4,"can / cannot",sz=12,c=GRAY)
tb(s,0.6,2.7,4.1,0.5,"🛡️ 安全 / ⚠️ 不安全",sz=22,b=True,c=PINE)
tb(s,0.6,3.2,4.1,0.4,"safe / not safe",sz=12,c=GRAY)
tb(s,0.6,3.75,4.1,0.4,"💡 例: 火可以, 在帐篷里不可以",sz=13,c=DARK)
tb(s,0.6,4.1,4.1,0.4,"Ex: Fire is OK, in tent is not",sz=11,c=GRAY)
# G1-3 column
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.1),Inches(1.0),Inches(4.5),Inches(3.6))
sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=PINE;sh.line.width=Pt(2.5)
pill(s,5.3,1.15,1.6,0.4,"G1 - G3",PINE,sz=14)
tb(s,5.3,1.7,4.1,0.5,"___ 不安全, 因为 ___ 。",sz=20,b=True,c=PINE)
tb(s,5.3,2.2,4.1,0.4,"___ is not safe, because ___.",sz=12,c=GRAY)
tb(s,5.3,2.7,4.1,0.5,"我应该 ___ 。",sz=20,b=True,c=PINE)
tb(s,5.3,3.2,4.1,0.4,"I should ___.",sz=12,c=GRAY)
tb(s,5.3,3.75,4.1,0.4,"💡 例: 帐篷在树下不安全, 因为枯枝会掉。",sz=13,c=DARK)
tb(s,5.3,4.1,4.1,0.4,"Ex: Tent under tree is not safe, branch will fall.",sz=11,c=GRAY)
n+=1;pn(s,n)
notes(s,"K 用 2 字短语 (可以/不可以, 安全/不安全)。G1-3 整句, 带「因为」、「应该」。")

# 34. VOCAB OVERVIEW
s=ns();bg(s,CREAM);hb(s,"📚 今天的字  Today's Words",PINE)
tb(s,0.4,0.9,9.2,0.4,"我会认 5 个词  ·  我会写 3 个字",sz=20,b=True,c=DARK,a=PP_ALIGN.CENTER)
read_words=[("营","yíng","camp"),("帐篷","zhàng peng","tent"),("火","huǒ","fire"),
            ("包","bāo","bag"),("安全","ān quán","safe")]
write_chars=[("火","huǒ","fire"),("包","bāo","bag"),("水","shuǐ","water")]
# Read row
tb(s,0.4,1.55,9.2,0.4,"👀 我会认  I Can Read",sz=16,b=True,c=SUN)
for i,(w,py,en) in enumerate(read_words):
    x=0.4+i*1.92
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(2.0),Inches(1.78),Inches(1.2))
    sh.fill.solid();sh.fill.fore_color.rgb=WARM;sh.line.color.rgb=SUN;sh.line.width=Pt(2)
    tb(s,x+0.05,2.05,1.7,0.55,w,sz=28,b=True,c=PINE,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.65,1.7,0.3,py,sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.92,1.7,0.3,en,sz=10,c=DARK,a=PP_ALIGN.CENTER)
# Write row
tb(s,0.4,3.4,9.2,0.4,"✍️ 我会写  I Can Write",sz=16,b=True,c=PINE)
for i,(w,py,en) in enumerate(write_chars):
    x=2.0+i*2.05
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(3.85),Inches(1.9),Inches(1.2))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=PINE;sh.line.width=Pt(2)
    tb(s,x+0.05,3.92,1.8,0.55,w,sz=32,b=True,c=PINE,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,4.55,1.8,0.3,py,sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,4.82,1.8,0.3,en,sz=10,c=DARK,a=PP_ALIGN.CENTER)
n+=1;pn(s,n)

# 35-39. WORD CARDS — 我会认
read_data=[
    ("营","yíng","camp",          "我们去搭一个营地。",       "📷 营地 / Tent + fire camp"),
    ("帐篷","zhàng peng","tent",  "我的帐篷是绿色的。",       "📷 绿色帐篷"),
    ("火","huǒ","fire",            "火很热, 不要靠近。",       "📷 营火 (用石头围)"),
    ("包","bāo","bag",             "我的背包里有水和食物。",   "📷 背包 + 物品"),
    ("安全","ān quán","safe",      "在大人旁边最安全。",       "📷 大人 + 孩子在营地"),
]
for w,py,en,sent,img in read_data:
    s=word_card_read(w,py,en,sent,img,color=SUN);n+=1;pn(s,n)

# 40-42. WORD CARDS — 我会写
write_data=[
    ("火","huǒ","fire",  "4 笔: 点 - 撇 - 撇 - 捺\n像火苗的样子! 🔥"),
    ("包","bāo","bag",   "5 笔: 撇 - 横折钩 - 横折 - 横 - 竖弯钩\n外面像「勹」, 里面是「巳」"),
    ("水","shuǐ","water","4 笔: 竖钩 - 横撇 - 撇 - 捺\n中间一竖, 两边像水滴 💧"),
]
for w,py,en,strokes in write_data:
    s=word_card_write(w,py,en,strokes,color=PINE);n+=1;pn(s,n)

# 43. SESSION 3 DIVIDER
s=div("Session 3  下午","🎒 写 Booklet + 2 个项目  Booklet + Camp Model + Pack Challenge",BROWN,"🛠️");n+=1;pn(s,n)

# 44. Day 2 Booklet
s=ns();bg(s,CREAM);hb(s,"📓 完成 Day 2 练习册  Day 2 Booklet",BROWN)
tb(s,0.4,0.85,9.2,0.35,"老师带着一起做  Teacher leads — do it together",sz=14,b=True,c=BROWN,a=PP_ALIGN.CENTER)
ib(s,0.4,1.25,9.2,4.0,"📷 Booklet 截图")
n+=1;pn(s,n)
notes(s,"老师带着学生一页一页完成 Day 2 练习册。可以用「先看 → 一起读 → 自己做 → 同桌检查」。")

# 45. Projects overview — 2 cards side by side
s=ns();bg(s,CREAM);hb(s,"🛠️ 动手时间！  Hands-On Time — 选一个项目",BROWN)
projects=[
    ("PROJECT 1","🏕️ 最安全的露营地","Camp Model — Design the safest camp",
     "适合高 level · 手工有难度",TENTCL),
    ("PROJECT 2","🎒 露营背包挑战","Pack Your Backpack — Pack the right items",
     "适合低 level · 简单上手",SUN),
]
for i,(lbl,nm,en,lvl,cl) in enumerate(projects):
    x=0.4+i*4.6
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(0.95),Inches(4.4),Inches(4.15))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=cl;sh.line.width=Pt(2.5)
    tb(s,x+0.1,1.05,4.2,0.35,lbl,sz=12,b=True,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,1.4,4.2,0.55,nm,sz=20,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,1.95,4.2,0.35,en,sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    pill(s,x+1.2,2.35,2.0,0.35,lvl.split(" · ")[0],cl,sz=11)
    ib(s,x+0.2,2.85,4.0,1.5,"📷 示范")
    tf=tb(s,x+0.15,4.45,4.1,0.35,lvl.split(" · ")[0],sz=12,c=DARK,a=PP_ALIGN.CENTER)
    ap(tf,lvl.split(" · ")[1] if " · " in lvl else "",sz=12,c=DARK,a=PP_ALIGN.CENTER)
n+=1;pn(s,n)
notes(s,"两个项目根据班级 level 二选一。Project 1 (Camp Model) 需要更多手工材料和时间, 适合 G1-3。Project 2 (Pack Challenge) 用图卡, 简单, 适合 K-G1。")

# 46. Project 1 — Camp Model (high-level)
s=ns();bg(s,CREAM);hb(s,"🏕️ Project 1: 最安全的露营地  Camp Model",TENTCL)
# Left: materials
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(0.95),Inches(4.4),Inches(0.4))
sh.fill.solid();sh.fill.fore_color.rgb=TENTCL;sh.line.fill.background()
tb(s,0.4,0.98,4.2,0.35,"🧺 材料  Materials",sz=14,b=True,c=WHITE)
tf=tb(s,0.4,1.45,4.4,2.3,"📋 纸板 (底板)  Cardboard base",sz=13,c=DARK)
ap(tf,"📦 小盒子 (帐篷)  Small box (tent)",sz=13,c=DARK)
ap(tf,"📄 纸 + 彩笔  Paper + markers",sz=13,c=DARK)
ap(tf,"📎 胶带  Tape",sz=13,c=DARK)
# Right: steps
sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(0.95),Inches(4.8),Inches(0.4))
sh2.fill.solid();sh2.fill.fore_color.rgb=SUN;sh2.line.fill.background()
tb(s,5.0,0.98,4.6,0.35,"👉 做法  Steps",sz=14,b=True,c=WHITE)
tf2=tb(s,5.0,1.45,4.7,2.3,"1️⃣ 选环境: 🌊 海边 or 🌳 森林?  Choose env",sz=13,c=DARK)
ap(tf2,"2️⃣ 摆出 4 个区域: ⛺ 帐篷 / 🔥 生火区 (远离帐篷!) / 🍱 食物区 / 🟢 安全区",sz=12,c=DARK)
ap(tf2,"3️⃣ 用纸/盒子做出来  Build it",sz=13,c=DARK)
ap(tf2,"4️⃣ 跟同学讲一讲  Explain to a friend",sz=13,c=DARK)
# Bottom: sentence frames
sh3=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(3.9),Inches(9.4),Inches(1.25))
sh3.fill.solid();sh3.fill.fore_color.rgb=WARM;sh3.line.color.rgb=TENTCL;sh3.line.width=Pt(2)
tb(s,0.5,4.0,9,0.35,"🗣️ 展示句型:",sz=14,b=True,c=TENTCL)
tb(s,0.5,4.4,9,0.35,"· 这是我的营地。",sz=14,c=DARK)
tb(s,0.5,4.7,9,0.35,"· 我选了 ___。 (海边 / 森林)",sz=14,c=DARK)
tb(s,5.0,4.4,4.7,0.35,"· 我把 ___ 放在 ___, 因为 ___。",sz=14,c=DARK)
n+=1;pn(s,n)
notes(s,"高 level 项目: 用纸盒、纸板、彩笔做一个 3D 营地模型。重点不在精美, 而在 4 区分开 + 能讲出原因。")

# 47. Project 2 — Backpack Challenge (low-level)
s=ns();bg(s,CREAM);hb(s,"🎒 Project 2: 露营背包挑战  Pack Your Backpack",SUN)
# Left: materials
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(0.95),Inches(3.0),Inches(0.4))
sh.fill.solid();sh.fill.fore_color.rgb=SUN;sh.line.fill.background()
tb(s,0.4,0.98,2.8,0.35,"🧺 材料  Materials",sz=14,b=True,c=WHITE)
tf=tb(s,0.4,1.45,3.0,2.3,"🃏 物品图卡  Item picture cards (or worksheet)",sz=12,c=DARK)
ap(tf,"✏️ 铅笔 / Pencil",sz=12,c=DARK)
ap(tf,"✅ ❌ 贴纸 (or stickers)",sz=12,c=DARK)
# Middle: items
sh_ex=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(3.45),Inches(0.95),Inches(3.0),Inches(0.4))
sh_ex.fill.solid();sh_ex.fill.fore_color.rgb=BROWN;sh_ex.line.fill.background()
tb(s,3.55,0.98,2.8,0.35,"💡 物品  Items",sz=14,b=True,c=WHITE)
tf_ex=tb(s,3.55,1.45,3.0,2.3,"✔️ 有用: 帐篷 / 水壶 / 手电筒 / 急救包 / 食物",sz=11,c=DARK)
ap(tf_ex,"❌ 没用: 玩具 / 电视 / 大蛋糕 / 游戏机 😂",sz=11,c=DARK)
# Right: task
sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(6.6),Inches(0.95),Inches(3.1),Inches(0.4))
sh2.fill.solid();sh2.fill.fore_color.rgb=PINE;sh2.line.fill.background()
tb(s,6.7,0.98,2.9,0.35,"👉 任务  Task",sz=14,b=True,c=WHITE)
tf2=tb(s,6.7,1.45,3.0,2.3,"1️⃣ 看 9 个物品",sz=12,c=DARK)
ap(tf2,"2️⃣ 选出有用的 (✔️) 和没用的 (❌)",sz=12,c=DARK)
ap(tf2,"3️⃣ 分类:",sz=12,c=DARK)
ap(tf2,"   😴 睡觉: ___",sz=11,c=DARK)
ap(tf2,"   🍱 吃东西: ___",sz=11,c=DARK)
ap(tf2,"   🛡️ 安全: ___",sz=11,c=DARK)
# Bottom: sentence frames
sh3=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(3.9),Inches(9.4),Inches(1.25))
sh3.fill.solid();sh3.fill.fore_color.rgb=WARM;sh3.line.color.rgb=OK;sh3.line.width=Pt(2)
tb(s,0.5,4.0,9,0.35,"🗣️ 展示句型:",sz=14,b=True,c=OK)
tb(s,0.5,4.4,9,0.35,"· 这个有用 / 没用。",sz=14,c=DARK)
tb(s,0.5,4.7,9,0.35,"· 我选 ___, 因为 ___。",sz=14,c=DARK)
tb(s,5.0,4.4,4.7,0.35,"· 睡觉需要 ___。 / 吃东西需要 ___。 / 安全需要 ___。",sz=12,c=DARK)
n+=1;pn(s,n)
notes(s,"低 level 项目: 用图卡和打勾, 不需要手工。重点是分类 + 用句型说原因。可以让学生先剪图卡再分类。")

# 48. CLOSING — mission complete
s=ns();bg(s,PINE)
tb(s,1,0.6,8,0.7,"🏆 任务完成!",sz=44,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,1,1.35,8,0.5,"Mission Complete!",sz=20,c=WARM,a=PP_ALIGN.CENTER)
# Badge
sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.5),Inches(2.0),Inches(3.0),Inches(3.0))
sh.fill.solid();sh.fill.fore_color.rgb=SUNYEL;sh.line.color.rgb=WHITE;sh.line.width=Pt(4)
tb(s,3.5,2.4,3.0,0.6,"⭐",sz=80,a=PP_ALIGN.CENTER)
tb(s,3.5,3.6,3.0,0.5,"小探险家",sz=22,b=True,c=PINE,a=PP_ALIGN.CENTER)
tb(s,3.5,4.1,3.0,0.4,"Little Explorer",sz=12,c=BROWN,a=PP_ALIGN.CENTER)
tb(s,1,5.0,8,0.4,"明天: Day 3 — 找路 + 信号 / Day 3: Navigation + Signals",sz=14,c=WARM,a=PP_ALIGN.CENTER)
n+=1;pn(s,n)
notes(s,"颁奖时间 (1 分钟): 给每个学生贴 1 张「小探险家」贴纸 / 印章。\n说: 「你今天完成了 4 个任务 — 你是真正的小探险家!」\n预告 Day 3 内容。")

# === Save ===
out="/Users/Huan/projects/summercourse/Chinese/野外生存与探险wilderness_pbl/day2_camp.pptx"
prs.save(out)
print(f"Saved {out}  ({n} slides)")
