#!/usr/bin/env python3
"""
我的职业梦想 — Day 1: 认识职业世界 (Discover the World of Careers)
Same structure as 野外生存/小小艺术家 PBL, with a career-themed palette.
Show-don't-tell: students guess, act, choose, match, solve problems.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
prs.slide_width  = Inches(10)
prs.slide_height = Inches(5.625)
W, H = prs.slide_width, prs.slide_height

# --- Palette: Career Adventure (navy + warm gold + helping green) ---
NAVY   = RGBColor(0x1E,0x3A,0x5F)   # primary — professional, trustworthy
GOLD   = RGBColor(0xF5,0xA6,0x23)   # secondary — energy, achievement
HELP   = RGBColor(0x2E,0x7D,0x32)   # accent — green for "helping"
CREAM  = RGBColor(0xFF,0xF8,0xE7)
WARM   = RGBColor(0xFF,0xF3,0xE0)
BROWN  = RGBColor(0x6B,0x44,0x23)
WHITE  = RGBColor(0xFF,0xFF,0xFF)
DARK   = RGBColor(0x2C,0x2C,0x2C)
GRAY   = RGBColor(0x88,0x88,0x88)
LGRAY  = RGBColor(0xBB,0xBB,0xBB)
IMGBG  = RGBColor(0xE8,0xE8,0xE8)
ALERT  = RGBColor(0xD0,0x4A,0x3C)
OK     = RGBColor(0x38,0x8E,0x3C)

# Job-specific colors
DOC    = RGBColor(0xE5,0x3E,0x3E)   # 医生 red
TEACH  = RGBColor(0x43,0xA0,0x47)   # 老师 green
POLICE = RGBColor(0x19,0x76,0xD2)   # 警察 blue
CHEF   = RGBColor(0xFB,0x8C,0x00)   # 厨师 orange
ENG    = RGBColor(0xFB,0xC0,0x2D)   # 工程师 yellow
ENV    = RGBColor(0x00,0x69,0x6B)   # 环境工程师 deep teal
WILD   = RGBColor(0x6D,0x4C,0x41)   # 野生动物保护员 brown
CITY   = RGBColor(0x7B,0x1F,0xA2)   # 城市规划师 purple

# === Helpers (same conventions as Day 2 camp / Day 1 art) ===
def ns(): return prs.slides.add_slide(prs.slide_layouts[6])

def tb(s,l,t,w,h,txt,sz=18,b=False,c=DARK,a=None):
    bx=s.shapes.add_textbox(Inches(l),Inches(t),Inches(w),Inches(h))
    tf=bx.text_frame; tf.word_wrap=True
    p=tf.paragraphs[0]
    if a: p.alignment=a
    r=p.add_run(); r.text=txt
    r.font.size=Pt(sz); r.font.bold=b; r.font.color.rgb=c; r.font.name='KaiTi'
    return tf

def ap(tf,txt,sz=18,b=False,c=DARK,a=None):
    p=tf.add_paragraph()
    if a: p.alignment=a
    r=p.add_run(); r.text=txt
    r.font.size=Pt(sz); r.font.bold=b; r.font.color.rgb=c; r.font.name='KaiTi'

def bg(s,c):
    sh=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,W,H)
    sh.fill.solid(); sh.fill.fore_color.rgb=c; sh.line.fill.background()
    sp=sh._element; sp.getparent().remove(sp); s.shapes._spTree.insert(2,sp)

def ib(s,l,t,w,h,lb="📷"):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(l),Inches(t),Inches(w),Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb=IMGBG; sh.line.fill.background()
    tb(s,l+0.1,t+h/2-0.2,w-0.2,0.4,lb,sz=14,c=LGRAY,a=PP_ALIGN.CENTER)

def hb(s,txt,c=NAVY,t=0.15):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(t),Inches(9.4),Inches(0.55))
    sh.fill.solid(); sh.fill.fore_color.rgb=c; sh.line.fill.background()
    tb(s,0.4,t+0.03,9.2,0.5,txt,sz=20,b=True,c=WHITE)

def pn(s,n): tb(s,9.0,5.25,0.8,0.3,str(n),sz=10,c=GRAY,a=PP_ALIGN.RIGHT)

def notes(s,text):
    nf=s.notes_slide.notes_text_frame
    lines=text.split("\n")
    nf.text=lines[0]
    for line in lines[1:]:
        p=nf.add_paragraph(); p.text=line

def pill(s,l,t,w,h,txt,c,sz=14):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(l),Inches(t),Inches(w),Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb=c; sh.line.fill.background()
    tb(s,l+0.1,t+h/2-0.2,w-0.2,0.4,txt,sz=sz,b=True,c=WHITE,a=PP_ALIGN.CENTER)

def div(title,sub,color,emoji=""):
    s=ns(); bg(s,color)
    tb(s,1,1.5,8,1.2,f"{emoji} {title}",sz=42,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,1,2.8,8,0.8,sub,sz=22,c=WHITE,a=PP_ALIGN.CENTER)
    return s

def sentence_frame_bar(s,t,frame_cn,frame_en):
    sf=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(t),Inches(9.4),Inches(0.65))
    sf.fill.solid(); sf.fill.fore_color.rgb=WARM
    sf.line.color.rgb=GOLD; sf.line.width=Pt(2)
    tb(s,0.5,t+0.1,1.7,0.4,"💬 我来说",sz=14,b=True,c=GOLD)
    tb(s,2.0,t+0.07,7.6,0.3,frame_cn,sz=14,b=True,c=DARK)
    tb(s,2.0,t+0.32,7.6,0.3,frame_en,sz=10,c=GRAY,a=None)

# === Specialized helpers ===

def mission_card(s,l,t,w,h,num,task_cn,task_en,emoji,color):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(l),Inches(t),Inches(w),Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE
    sh.line.color.rgb=color; sh.line.width=Pt(2.5)
    badge=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(l+0.1),Inches(t+0.08),Inches(0.55),Inches(0.55))
    badge.fill.solid(); badge.fill.fore_color.rgb=color; badge.line.fill.background()
    tb(s,l+0.1,t+0.18,0.55,0.4,str(num),sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,l+0.7,t+0.15,w-0.8,0.4,task_en,sz=10,c=GRAY)
    tb(s,l+0.05,t+0.85,w-0.1,0.7,emoji,sz=44,a=PP_ALIGN.CENTER)
    tb(s,l+0.05,t+1.55,w-0.1,0.4,task_cn,sz=18,b=True,c=color,a=PP_ALIGN.CENTER)

def guess_clue_slide(emoji,job_label,color,clues,frame_cn,frame_en):
    """Clue slide — students hear/see clues, guess the job (no answer shown)."""
    s=ns(); bg(s,CREAM)
    hb(s,f"{emoji} 猜猜我是谁？  Guess Who I Am!",color)
    tb(s,0.4,0.85,9.2,0.32,"听听我的提示 — 我是谁？",sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
    tb(s,0.4,1.18,9.2,0.26,"Listen to my clues — who am I?",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    # LEFT — picture placeholder (mystery)
    img_box=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.55),Inches(4.30),Inches(2.85))
    img_box.fill.solid(); img_box.fill.fore_color.rgb=IMGBG
    img_box.line.color.rgb=color; img_box.line.width=Pt(2)
    tb(s,0.4,2.55,4.30,0.7,"❓",sz=70,c=color,a=PP_ALIGN.CENTER)
    tb(s,0.4,3.40,4.30,0.40,"我是谁？",sz=18,b=True,c=color,a=PP_ALIGN.CENTER)
    tb(s,0.4,3.80,4.30,0.30,"Who am I?",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    # RIGHT — clues
    panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.85),Inches(1.55),Inches(4.85),Inches(2.85))
    panel.fill.solid(); panel.fill.fore_color.rgb=WHITE
    panel.line.color.rgb=color; panel.line.width=Pt(2.5)
    head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.85),Inches(1.55),Inches(4.85),Inches(0.50))
    head.fill.solid(); head.fill.fore_color.rgb=color; head.line.fill.background()
    tb(s,5.0,1.62,4.6,0.4,"🔎 提示 Clues",sz=14,b=True,c=WHITE)
    y=2.20
    for i,(c_cn,c_en) in enumerate(clues):
        badge=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(5.0),Inches(y+0.05),Inches(0.4),Inches(0.4))
        badge.fill.solid(); badge.fill.fore_color.rgb=color; badge.line.fill.background()
        tb(s,5.0,y+0.10,0.4,0.3,str(i+1),sz=12,b=True,c=WHITE,a=PP_ALIGN.CENTER)
        tb(s,5.5,y,4.1,0.4,c_cn,sz=14,b=True,c=DARK)
        tb(s,5.5,y+0.32,4.1,0.30,c_en,sz=9,c=GRAY)
        y+=0.65
    sentence_frame_bar(s,4.55,frame_cn,frame_en)
    return s

def reveal_job_slide(emoji,job_cn,job_en,color,what_cn,what_en,where_cn,where_en):
    """Reveal slide — show the job icon, name, what they do, where they work."""
    s=ns(); bg(s,CREAM)
    hb(s,f"{emoji} 答案揭晓!  Answer Revealed!",color)
    # Big emoji + job name
    bigbox=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.95),Inches(4.30),Inches(3.45))
    bigbox.fill.solid(); bigbox.fill.fore_color.rgb=WHITE
    bigbox.line.color.rgb=color; bigbox.line.width=Pt(3)
    tb(s,0.4,1.30,4.30,1.50,emoji,sz=130,a=PP_ALIGN.CENTER)
    tb(s,0.4,2.95,4.30,0.55,job_cn,sz=30,b=True,c=color,a=PP_ALIGN.CENTER)
    tb(s,0.4,3.55,4.30,0.40,job_en,sz=14,c=GRAY,a=PP_ALIGN.CENTER)
    pill(s,1.4,4.05,2.3,0.30,"我就是 ___ !",GOLD,sz=11)
    # RIGHT — facts
    panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.85),Inches(0.95),Inches(4.85),Inches(3.45))
    panel.fill.solid(); panel.fill.fore_color.rgb=WHITE
    panel.line.color.rgb=color; panel.line.width=Pt(2.5)
    tb(s,5.0,1.10,4.6,0.40,"🛠️ 做什么? What I do",sz=15,b=True,c=color)
    tb(s,5.0,1.55,4.6,0.40,what_cn,sz=14,b=True,c=DARK)
    tb(s,5.0,1.95,4.6,0.30,what_en,sz=10,c=GRAY)
    tb(s,5.0,2.50,4.6,0.40,"📍 在哪里? Where",sz=15,b=True,c=color)
    tb(s,5.0,2.95,4.6,0.40,where_cn,sz=14,b=True,c=DARK)
    tb(s,5.0,3.35,4.6,0.30,where_en,sz=10,c=GRAY)
    sentence_frame_bar(s,4.55,
        f"我是 {job_cn} 。我帮助 ___ 。",
        f"I am a {job_en}. I help ___.")
    return s

def mystery_job_q_slide(emoji,job_label,color,picture_label,q1_cn,q1_en,q2_cn,q2_en):
    """Mystery job — picture + 2 simple A-or-B questions, NO label revealed yet."""
    s=ns(); bg(s,CREAM)
    hb(s,f"{emoji} 神秘职业  Mystery Job!",color)
    tb(s,0.4,0.85,9.2,0.32,"看看图 — 他们在做什么？",sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
    tb(s,0.4,1.18,9.2,0.26,"Look — what are they doing?",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    # LEFT — picture placeholder
    img_box=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.55),Inches(4.30),Inches(2.95))
    img_box.fill.solid(); img_box.fill.fore_color.rgb=IMGBG
    img_box.line.color.rgb=color; img_box.line.width=Pt(2.5)
    tb(s,0.4,2.55,4.30,0.65,emoji,sz=60,a=PP_ALIGN.CENTER)
    tb(s,0.4,3.30,4.30,0.40,picture_label,sz=12,b=True,c=color,a=PP_ALIGN.CENTER)
    tb(s,0.4,3.70,4.30,0.30,"📷 (paste real photo here)",sz=9,c=LGRAY,a=PP_ALIGN.CENTER)
    # RIGHT — 2 questions
    panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.85),Inches(1.55),Inches(4.85),Inches(2.95))
    panel.fill.solid(); panel.fill.fore_color.rgb=WHITE
    panel.line.color.rgb=color; panel.line.width=Pt(2.5)
    head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.85),Inches(1.55),Inches(4.85),Inches(0.50))
    head.fill.solid(); head.fill.fore_color.rgb=color; head.line.fill.background()
    tb(s,5.0,1.62,4.6,0.4,"🤔 想一想 Think First",sz=14,b=True,c=WHITE)
    badge1=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(5.0),Inches(2.30),Inches(0.45),Inches(0.45))
    badge1.fill.solid(); badge1.fill.fore_color.rgb=color; badge1.line.fill.background()
    tb(s,5.0,2.36,0.45,0.3,"1",sz=14,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,5.55,2.25,4.05,0.40,q1_cn,sz=14,b=True,c=DARK)
    tb(s,5.55,2.62,4.05,0.30,q1_en,sz=9,c=GRAY)
    badge2=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(5.0),Inches(3.20),Inches(0.45),Inches(0.45))
    badge2.fill.solid(); badge2.fill.fore_color.rgb=color; badge2.line.fill.background()
    tb(s,5.0,3.26,0.45,0.3,"2",sz=14,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,5.55,3.15,4.05,0.40,q2_cn,sz=14,b=True,c=DARK)
    tb(s,5.55,3.52,4.05,0.30,q2_en,sz=9,c=GRAY)
    sentence_frame_bar(s,4.65,
        "我觉得他们 ___ 。",
        "I think they are ___.")
    return s

def mystery_job_label_slide(emoji,full_job_cn,full_job_en,kid_label_cn,kid_label_en,color,what_does):
    """Reveal the kid-friendly label for a mystery job."""
    s=ns(); bg(s,CREAM)
    hb(s,f"{emoji} 答案揭晓!  Answer!",color)
    # Big card
    big=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.6),Inches(0.95),Inches(8.8),Inches(2.7))
    big.fill.solid(); big.fill.fore_color.rgb=color; big.line.fill.background()
    tb(s,0.6,1.10,8.8,0.85,emoji,sz=70,a=PP_ALIGN.CENTER)
    tb(s,0.6,2.05,8.8,0.55,full_job_cn,sz=32,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,0.6,2.65,8.8,0.40,full_job_en,sz=14,c=WARM,a=PP_ALIGN.CENTER)
    # Kid-friendly label
    pill(s,2.5,3.10,5.0,0.50,kid_label_cn,GOLD,sz=18)
    tb(s,0.6,3.65,8.8,0.30,kid_label_en,sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    # What they do
    tb(s,0.6,4.05,8.8,0.40,what_does,sz=13,b=True,c=DARK,a=PP_ALIGN.CENTER)
    sentence_frame_bar(s,4.65,
        f"{full_job_cn} 是 ___ 的人。",
        f"A {full_job_en} is a person who ___.")
    return s

def scenario_q_slide(em,scene_cn,scene_en,color,help_options):
    """Problem scenario — students decide who can help (3 options shown, no answer)."""
    s=ns(); bg(s,CREAM)
    hb(s,"🆘 谁来帮忙？  Who Can Help?",color)
    # scenario banner
    big=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.95),Inches(9.2),Inches(1.40))
    big.fill.solid(); big.fill.fore_color.rgb=color; big.line.fill.background()
    tb(s,0.4,1.05,9.2,0.50,em,sz=28,a=PP_ALIGN.CENTER)
    tb(s,0.4,1.55,9.2,0.45,scene_cn,sz=20,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,0.4,2.00,9.2,0.30,scene_en,sz=11,c=WARM,a=PP_ALIGN.CENTER)
    # 3 choice cards (same 3 mystery jobs)
    for i,(j_em,j_cn,j_en,j_col) in enumerate(help_options):
        x=0.4+i*3.15
        sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(2.55),Inches(3.0),Inches(1.85))
        sh.fill.solid(); sh.fill.fore_color.rgb=WHITE
        sh.line.color.rgb=j_col; sh.line.width=Pt(2.5)
        tb(s,x+0.05,2.65,2.9,0.65,j_em,sz=38,a=PP_ALIGN.CENTER)
        tb(s,x+0.05,3.38,2.9,0.40,j_cn,sz=15,b=True,c=j_col,a=PP_ALIGN.CENTER)
        tb(s,x+0.05,3.80,2.9,0.30,j_en,sz=9,c=GRAY,a=PP_ALIGN.CENTER)
        pill(s,x+0.85,4.10,1.30,0.25,f"选 {chr(65+i)}",j_col,sz=10)
    sentence_frame_bar(s,4.55,
        "我选 ___ 来帮忙。",
        "I choose ___ to help.")
    return s

def scenario_a_slide(em,scene_cn,scene_en,answer_emoji,answer_cn,answer_en,color,why_cn,why_en):
    """Scenario reveal — show which job helps + 1-line why."""
    s=ns(); bg(s,CREAM)
    hb(s,"💡 答案揭晓!  Who Helps!",color)
    # scenario reminder
    big=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.95),Inches(9.2),Inches(0.85))
    big.fill.solid(); big.fill.fore_color.rgb=color; big.line.fill.background()
    tb(s,0.4,1.00,9.2,0.40,em+" "+scene_cn,sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,0.4,1.42,9.2,0.30,scene_en,sz=10,c=WARM,a=PP_ALIGN.CENTER)
    # answer card (big)
    ans=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(2.0),Inches(2.05),Inches(6.0),Inches(2.30))
    ans.fill.solid(); ans.fill.fore_color.rgb=WHITE
    ans.line.color.rgb=color; ans.line.width=Pt(3)
    tb(s,2.0,2.15,6.0,0.85,answer_emoji,sz=64,a=PP_ALIGN.CENTER)
    tb(s,2.0,3.10,6.0,0.50,answer_cn,sz=24,b=True,c=color,a=PP_ALIGN.CENTER)
    tb(s,2.0,3.65,6.0,0.35,answer_en,sz=12,c=GRAY,a=PP_ALIGN.CENTER)
    pill(s,3.5,4.00,3.0,0.30,"✅ 来帮忙!",OK,sz=11)
    # Why
    tb(s,0.4,4.45,9.2,0.30,f"💡 因为 {why_cn}",sz=13,b=True,c=DARK,a=PP_ALIGN.CENTER)
    tb(s,0.4,4.78,9.2,0.30,f"Because {why_en}",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    return s

# ============================================================
# === BUILD SLIDES ==========================================
# ============================================================
n=0

# 1. COVER
s=ns(); bg(s,NAVY)
sh=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,Inches(2.4),W,Inches(2.0))
sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.fill.background()
tb(s,1,0.4,8,0.5,"DAY 1",sz=18,b=True,c=GOLD,a=PP_ALIGN.CENTER)
tb(s,1,0.95,8,0.7,"💼 我的职业梦想",sz=42,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,1,1.7,8,0.5,"My Career Dream",sz=22,c=WARM,a=PP_ALIGN.CENTER)
tb(s,1,2.6,8,0.5,"🌍 认识职业世界  Discover the World of Careers",sz=22,b=True,c=NAVY,a=PP_ALIGN.CENTER)
tb(s,1,3.15,8,0.4,"Guess · Act · Choose · Solve · Dream",sz=14,c=BROWN,a=PP_ALIGN.CENTER)
tb(s,1,3.55,8,0.4,"猜猜 · 演演 · 选选 · 解决 · 梦想",sz=14,b=True,c=BROWN,a=PP_ALIGN.CENTER)
tb(s,1,4.6,8,0.4,"小小职业人 · Future Career Heroes",sz=14,b=True,c=GOLD,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"开场 (1 分钟):\n• 「同学们! 长大后, 你想做什么?」\n• 今天我们去「职业世界」看看 — 不是听老师讲, 是自己猜、演、选!\n• 5 个步骤完成 → 拿「小小职业人」徽章!")

# 2. TODAY'S MISSION — 5 steps
s=ns(); bg(s,CREAM); hb(s,"🧭 今天的任务  Today's Mission",NAVY)
tb(s,0.4,0.85,9.2,0.45,"🌍 我们要认识 8 个职业!",sz=24,b=True,c=NAVY,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.28,"We're going to discover 8 jobs!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,0.4,1.55,9.2,0.32,"👉 5 个游戏 — 玩中学, 边玩边学职业 ↓",sz=14,b=True,c=BROWN,a=PP_ALIGN.CENTER)
mission_card(s,0.4,1.95,1.80,2.20,1,"猜猜","Guess",     "🔎",DOC)
mission_card(s,2.30,1.95,1.80,2.20,2,"演演","Act",       "🎭",CHEF)
mission_card(s,4.20,1.95,1.80,2.20,3,"神秘职业","Mystery","✨",ENV)
mission_card(s,6.10,1.95,1.80,2.20,4,"谁来帮忙","Who Helps?","🆘",HELP)
mission_card(s,8.00,1.95,1.80,2.20,5,"我的梦想","My Dream","💭",GOLD)
sentence_frame_bar(s,4.40,
    "我想认识 ___ 。",
    "I want to learn about ___.")
n+=1; pn(s,n)
notes(s,"1-2 分钟:\n• 介绍 5 个游戏 — 让学生跟读: 猜猜 → 演演 → 神秘 → 帮忙 → 梦想\n• 「完成 5 个游戏, 拿到「小小职业人」徽章!」\n• 在白板画 5 格, 每完成一个画 ✓")

# 3. SESSION 1 DIVIDER
s=div("Session 1  上午 11:00–11:45","🎯 Career Adventure · 职业大冒险",NAVY,"🌟"); n+=1; pn(s,n)

# 4. STEP 1 — Warm-up: 你想当 ___ 吗?
s=ns(); bg(s,CREAM); hb(s,"👋 你想当...?  Do You Want to Be...?",GOLD)
tb(s,0.4,0.85,9.2,0.45,"举手 / 站起来 — 你想当吗？",sz=22,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.30,"Raise your hand / stand up — do you want to be one?",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
warm=[("🩺","医生","Doctor",DOC),
      ("📚","老师","Teacher",TEACH),
      ("👮","警察","Police",POLICE),
      ("👨‍🍳","厨师","Chef",CHEF),
      ("👷","工程师","Engineer",ENG)]
for i,(em,cn,en,c) in enumerate(warm):
    x=0.4+i*1.88
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.85),Inches(1.78),Inches(2.40))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE
    sh.line.color.rgb=c; sh.line.width=Pt(2.5)
    tb(s,x+0.05,1.95,1.7,0.85,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.85,1.7,0.45,cn,sz=18,b=True,c=c,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.30,1.7,0.30,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    pill(s,x+0.34,3.70,1.10,0.30,"举手 ✋",c,sz=10)
sentence_frame_bar(s,4.45,
    "我想当 ___ ! / 我不想当 ___ 。",
    "I want to be a ___! / I don't want to be ___.")
n+=1; pn(s,n)
notes(s,"5 分钟:\n• 老师指着每个职业: 「你想当医生吗?」让想当的孩子站起来或举手。\n• 不解释 — 直接玩!\n• 简单介绍: 职业 = 长大以后做的工作 (1 句话, 不要多)。\n• 让 1-2 个孩子说: 「我想当 ___ !」")

# 5. INTRO — 职业是什么?
s=ns(); bg(s,CREAM); hb(s,"💡 职业是什么？  What is a Career?",NAVY)
# 3 visual definitions
defs=[("🤝","帮助别人","Help others",HELP),
      ("🛠️","解决问题","Solve problems",GOLD),
      ("✨","创造价值","Create value",CITY)]
for i,(em,cn,en,c) in enumerate(defs):
    x=0.4+i*3.15
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.10),Inches(3.0),Inches(2.55))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE
    sh.line.color.rgb=c; sh.line.width=Pt(2.5)
    tb(s,x+0.05,1.20,2.9,0.85,em,sz=56,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.10,2.9,0.45,cn,sz=20,b=True,c=c,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.55,2.9,0.30,en,sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    pill(s,x+0.75,2.95,1.50,0.30,"= 职业!",c,sz=10)
tb(s,0.4,3.85,9.2,0.40,"职业 = 长大以后做的工作",sz=20,b=True,c=NAVY,a=PP_ALIGN.CENTER)
tb(s,0.4,4.25,9.2,0.30,"A career = the work people do as adults",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.65,
    "职业可以 ___ 。",
    "A career can ___.")
n+=1; pn(s,n)
notes(s,"1 分钟 — 不要长讲!\n• 「职业是什么呢?」\n• 让学生说: 帮助别人 / 解决问题 / 做东西\n• 总结一句: 「职业 = 长大以后做的工作。」\n• 然后直接进游戏 — 不解释太多!")

# 6. STEP 2 INTRO — Guess Who I Am
s=ns(); bg(s,CREAM); hb(s,"🔎 猜猜我是谁?  Guess Who I Am!",DOC)
tb(s,0.4,0.95,9.2,0.45,"听 2 个提示, 你能猜出来吗？",sz=22,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.40,9.2,0.30,"Listen to 2 clues — can you guess?",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,0.4,1.95,9.2,0.40,"📖 5 个职业 — 我们一个一个猜!",sz=18,b=True,c=NAVY,a=PP_ALIGN.CENTER)
icons=[("🩺",DOC),("📚",TEACH),("👮",POLICE),("👨‍🍳",CHEF),("👷",ENG)]
for i,(em,c) in enumerate(icons):
    x=0.4+i*1.88
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(2.55),Inches(1.78),Inches(1.65))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE
    sh.line.color.rgb=c; sh.line.width=Pt(2.5)
    tb(s,x+0.05,2.65,1.7,0.80,"❓",sz=42,c=c,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.50,1.7,0.40,f"职业 {i+1}",sz=14,b=True,c=c,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.90,1.7,0.25,"Job "+str(i+1),sz=9,c=GRAY,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.55,
    "我猜是 ___ !",
    "I guess it's a ___!")
n+=1; pn(s,n)
notes(s,"1 分钟 — 引导:\n• 「老师有 5 个秘密职业, 每个给 2 个提示。」\n• 「猜对了, 我们就揭晓答案!」\n• 让学生跟读: 「我猜是 ___ !」\n• 翻到下一页开始第一个谜题。")

# 7-16. Five GUESS clue + reveal pairs
guesses=[
    ("🩺","Doctor",DOC,
     [("我帮助生病的人。","I help sick people."),
      ("我在医院工作。","I work in a hospital.")],
     "医生","Doctor","看病、给药、做手术","check, medicine, surgery","医院 / 诊所","hospital / clinic"),
    ("📚","Teacher",TEACH,
     [("我教小朋友。","I teach children."),
      ("我在学校工作。","I work at school.")],
     "老师","Teacher","教课、改作业、讲故事","teach, grade, tell stories","学校 / 教室","school / classroom"),
    ("👮","Police",POLICE,
     [("我帮助大家安全。","I keep people safe."),
      ("我开警车, 我穿制服。","I drive a police car, I wear a uniform.")],
     "警察","Police Officer","抓坏人、保护大家","catch bad guys, protect people","街上 / 警察局","street / police station"),
    ("👨‍🍳","Chef",CHEF,
     [("我做饭给大家吃。","I cook food for people."),
      ("我戴白色高高的帽子。","I wear a tall white hat.")],
     "厨师","Chef","做菜、做点心","cook meals, make desserts","餐厅 / 厨房","restaurant / kitchen"),
    ("👷","Engineer",ENG,
     [("我会用工具做东西。","I use tools to build things."),
      ("我建桥、机器、楼房。","I build bridges, machines, buildings.")],
     "工程师","Engineer","设计、建造、修理","design, build, repair","公司 / 工地","company / construction site"),
]
for em,j_en,col,clues,job_cn,job_en,what_cn,what_en,where_cn,where_en in guesses:
    s=guess_clue_slide(em,j_en,col,clues,
        f"我猜是 ___ !",
        f"I guess it's a ___!")
    n+=1; pn(s,n)
    notes(s,f"猜谜 (1-2 分钟):\n• 老师慢慢读 2 个提示, 让学生先想。\n• 让 2-3 个学生举手猜: 「我猜是 ___ !」\n• 不要直接给答案 — 让多个孩子先猜。\n• 然后翻页揭晓!")

    s=reveal_job_slide(em,job_cn,j_en,col,what_cn,what_en,where_cn,where_en)
    n+=1; pn(s,n)
    notes(s,f"揭晓 {job_cn}!\n• 让全班一起喊: 「{job_cn}!」\n• 复习: 做什么? 在哪里?\n• 用句型: 「我是 {job_cn}, 我帮助 ___ 。」")

# 17. STEP 3 — Career Charades
s=ns(); bg(s,CREAM); hb(s,"🎭 演一演!  Career Charades",CHEF)
tb(s,0.4,0.95,9.2,0.45,"上台演职业 — 不能说话!",sz=22,b=True,c=CHEF,a=PP_ALIGN.CENTER)
tb(s,0.4,1.40,9.2,0.30,"Act out a job — no talking!",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
# 3 instructional cards
steps=[("1️⃣","抽一张卡片","Pick a card",CHEF),
       ("2️⃣","只能演, 不能说","Act only — no words!",DOC),
       ("3️⃣","大家一起猜","Class guesses!",HELP)]
for i,(em,cn,en,c) in enumerate(steps):
    x=0.4+i*3.15
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.95),Inches(3.0),Inches(2.30))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE
    sh.line.color.rgb=c; sh.line.width=Pt(2.5)
    tb(s,x+0.05,2.05,2.9,0.7,em,sz=42,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.85,2.9,0.40,cn,sz=16,b=True,c=c,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.25,2.9,0.30,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.65,2.9,0.40,"⏱️ 30 秒",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.40,
    "你是 ___ 吗？  我是 ___ 。",
    "Are you a ___? I am a ___.")
n+=1; pn(s,n)
notes(s,"8 分钟 — 道具准备 (低投入):\n• 把 5 个职业名写在小纸条上, 折起来。\n• 选 4-5 个志愿者上台演 (30 秒), 不能说话, 不能写。\n• 全班用句型猜: 「你是 医生 吗?」\n• 演完, 演员说: 「我是 ___ 。」\n• K 学生: 简单动作 + 老师帮忙\n• G1-3: 加道具 (假装) 让动作更明显")

# 18. STEP 3.5 INTRO — 神秘职业!
s=ns(); bg(s,CREAM); hb(s,"✨ 神秘职业  Mystery Jobs!",ENV)
tb(s,0.4,0.85,9.2,0.50,"还有 3 个特别的职业 — 你猜过吗？",sz=22,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.40,9.2,0.30,"3 special jobs — have you heard of them?",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
mysteries=[("🌍","环境工程师","Environmental Engineer",ENV),
           ("🦁","野生动物保护员","Wildlife Protector",WILD),
           ("🏙️","城市规划师","City Planner",CITY)]
for i,(em,cn,en,c) in enumerate(mysteries):
    x=0.4+i*3.15
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(2.00),Inches(3.0),Inches(2.30))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE
    sh.line.color.rgb=c; sh.line.width=Pt(2.5)
    tb(s,x+0.05,2.10,2.9,0.85,"❓",sz=46,c=c,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.00,2.9,0.40,em,sz=30,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.55,2.9,0.40,"神秘 "+str(i+1),sz=14,b=True,c=c,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.95,2.9,0.25,"Mystery "+str(i+1),sz=9,c=GRAY,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.55,
    "他们是 ___ 的人？",
    "They are people who ___?")
n+=1; pn(s,n)
notes(s,"30 秒过渡:\n• 「现在 — 3 个新职业! 大人也不一定全知道哦!」\n• 跟读: 环境工程师 / 野生动物保护员 / 城市规划师\n• 我们看图来猜!")

# 19-24. Three MYSTERY JOBS — Q + Reveal pairs

# Mystery 1: 环境工程师
s=mystery_job_q_slide("🌍","Environmental Engineer",ENV,
    "💧 检测水 / 清洁 photo here",
    "他们在做什么？","What are they doing?",
    "是在玩水, 还是让水变干净？","Playing with water — or cleaning it?")
n+=1; pn(s,n)
notes(s,"3-4 分钟:\n• 投影一张「人们检测水/清洁水」的真实照片 (建议老师贴一张)。\n• 老师追问: 「他们手里拿什么? 那个瓶子是干什么的?」\n• 让学生先猜, 不要直接给答案!\n• 让 2-3 个孩子说: 「我觉得他们在 ___ 。」")

s=mystery_job_label_slide("🌍","环境工程师","Environmental Engineer",
    "保护水和空气的人","helps keep water and air clean",ENV,
    "他们让水变干净, 让空气变好!")
n+=1; pn(s,n)
notes(s,"揭晓 (1-2 分钟):\n• 「答案是 — 环境工程师!」\n• 让全班跟读 3 遍: 「环境工程师 — 保护水和空气的人」\n• 简单解释: 「水脏了 → 他们让水变干净。空气脏了 → 他们检查空气。」")

# Mystery 2: 野生动物保护员
s=mystery_job_q_slide("🦁","Wildlife Protector",WILD,
    "🦒 救动物 / 观察 photo here",
    "他们在做什么？","What are they doing?",
    "是在抓动物, 还是保护动物？","Catching animals — or protecting them?")
n+=1; pn(s,n)
notes(s,"3-4 分钟:\n• 投影一张「人们在帮助 / 观察动物」的真实照片。\n• 老师追问: 「动物开心吗? 这些人对动物好不好?」\n• 让学生猜, 不直接给答案!\n• 让 2-3 个孩子说: 「我觉得他们在 ___ 。」")

s=mystery_job_label_slide("🦁","野生动物保护员","Wildlife Protector",
    "保护动物的人","protects animals",WILD,
    "他们救受伤的动物, 看着不让坏人伤害动物!")
n+=1; pn(s,n)
notes(s,"揭晓:\n• 跟读: 「野生动物保护员 — 保护动物的人」\n• 解释: 「动物受伤 → 他们救; 有人想伤害动物 → 他们保护。」")

# Mystery 3: 城市规划师
s=mystery_job_q_slide("🏙️","City Planner",CITY,
    "🗺️ 城市地图 / 模型 photo here",
    "他们在做什么？","What are they doing?",
    "是在画画, 还是设计城市？","Drawing pictures — or designing a city?")
n+=1; pn(s,n)
notes(s,"3-4 分钟:\n• 投影一张「城市地图 / 模型」照片。\n• 老师追问: 「这是地图吗? 他们要建什么?」\n• 让学生猜!\n• 让 2-3 个孩子说: 「我觉得他们在 ___ 。」")

s=mystery_job_label_slide("🏙️","城市规划师","City Planner",
    "设计城市的人","designs better cities",CITY,
    "他们决定哪里建公园, 哪里建路, 让城市更好住!")
n+=1; pn(s,n)
notes(s,"揭晓:\n• 跟读: 「城市规划师 — 设计城市的人」\n• 解释: 「公园在哪里? 路怎么走? 学校建哪里? 都是他们设计的!」")

# 25. MYSTERY QUICK CHECK — Match
s=ns(); bg(s,CREAM); hb(s,"🎯 快速检查!  Quick Check!",GOLD)
tb(s,0.4,0.85,9.2,0.40,"你来回答 — 谁是 ___ 的人？",sz=20,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.25,9.2,0.28,"You answer — who is the person who ___?",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
checks=[("💧","谁让水变干净？","Who keeps water clean?","🌍","环境工程师",ENV),
        ("🦒","谁保护动物？","Who protects animals?","🦁","野生动物保护员",WILD),
        ("🗺️","谁设计城市？","Who designs cities?","🏙️","城市规划师",CITY)]
for i,(qem,q_cn,q_en,a_em,a_cn,c) in enumerate(checks):
    y=1.65+i*1.05
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(y),Inches(9.2),Inches(0.95))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE
    sh.line.color.rgb=c; sh.line.width=Pt(2)
    tb(s,0.55,y+0.20,0.6,0.55,qem,sz=30,a=PP_ALIGN.CENTER)
    tb(s,1.20,y+0.10,4.5,0.40,q_cn,sz=15,b=True,c=DARK)
    tb(s,1.20,y+0.50,4.5,0.30,q_en,sz=10,c=GRAY)
    # arrow
    tb(s,5.85,y+0.20,0.5,0.5,"→",sz=24,b=True,c=c,a=PP_ALIGN.CENTER)
    tb(s,6.40,y+0.20,0.6,0.55,a_em,sz=28,a=PP_ALIGN.CENTER)
    tb(s,7.05,y+0.25,2.5,0.4,a_cn,sz=14,b=True,c=c)
sentence_frame_bar(s,4.85,
    "___ 是 ___ 的人。",
    "A ___ is the person who ___.")
n+=1; pn(s,n)
notes(s,"3 分钟 — 老师指着问题, 全班一起喊答案!\n• 也可以让学生举手 / 用手势指着对应的图。\n• 让 1-2 个学生用句型说: 「环境工程师是让水变干净的人。」")

# 26. STEP 4 INTRO — Who Can Help?
s=ns(); bg(s,CREAM); hb(s,"🆘 谁来帮忙？  Who Can Help?",HELP)
tb(s,0.4,0.95,9.2,0.45,"出大事了! 谁能帮忙？",sz=24,b=True,c=ALERT,a=PP_ALIGN.CENTER)
tb(s,0.4,1.45,9.2,0.30,"There's a problem! Who can help?",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,0.4,2.00,9.2,0.40,"看 3 个问题, 你来选职业 — A、B 还是 C？",sz=16,b=True,c=NAVY,a=PP_ALIGN.CENTER)
tb(s,0.4,2.45,9.2,0.30,"3 problems — pick a job: A, B, or C?",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
# 3 helpers as choices
helpers=[("A","🌍","环境工程师",ENV),
         ("B","🦁","野生动物保护员",WILD),
         ("C","🏙️","城市规划师",CITY)]
for i,(letter,em,cn,c) in enumerate(helpers):
    x=0.4+i*3.15
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(2.95),Inches(3.0),Inches(1.50))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE
    sh.line.color.rgb=c; sh.line.width=Pt(2.5)
    pill(s,x+0.10,3.05,0.7,0.30,letter,c,sz=14)
    tb(s,x+0.85,3.00,2.05,0.45,em,sz=30,a=PP_ALIGN.LEFT)
    tb(s,x+0.05,3.65,2.9,0.40,cn,sz=16,b=True,c=c,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,4.05,2.9,0.30,"Helper "+letter,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.65,
    "我选 ___ 来帮忙。",
    "I choose ___ to help.")
n+=1; pn(s,n)
notes(s,"30 秒 — 介绍 3 个 helper:\n• 「这 3 个就是我们的「super helpers」!」\n• 「下面看 3 个问题, 你来选 A、B 还是 C。」")

# 27-32. Three SCENARIOS — Q + A pairs
help_options=[("🌍","环境工程师","Env Engineer",ENV),
              ("🦁","野生动物保护员","Wildlife",WILD),
              ("🏙️","城市规划师","City Planner",CITY)]

# Scenario 1: River dirty
s=scenario_q_slide("🐟","河水变脏了, 鱼生病了。","The river is dirty — fish are sick.",ALERT,help_options)
n+=1; pn(s,n)
notes(s,"问题 1 (2-3 分钟):\n• 老师读问题, 让学生想 30 秒。\n• 「A、B、C — 你选谁?」让学生举 A/B/C 手指。\n• 不直接给答案!\n• 翻页揭晓 ✓")

s=scenario_a_slide("🐟","河水变脏了, 鱼生病了。","The river is dirty — fish are sick.",
    "🌍","环境工程师","Environmental Engineer",ENV,
    "他们让水变干净, 鱼就好了!","they clean the water and fish get better!")
n+=1; pn(s,n)
notes(s,"揭晓: 环境工程师!\n• 让全班说: 「环境工程师帮助鱼!」\n• 句型: 「环境工程师帮助 ___ 。」")

# Scenario 2: Sea turtle stuck
s=scenario_q_slide("🐢","海龟被垃圾卡住了。","A sea turtle is stuck in trash.",ALERT,help_options)
n+=1; pn(s,n)
notes(s,"问题 2:\n• 「海龟受伤了 — 谁来救?」\n• 让学生举 A/B/C。\n• 翻页揭晓!")

s=scenario_a_slide("🐢","海龟被垃圾卡住了。","A sea turtle is stuck in trash.",
    "🦁","野生动物保护员","Wildlife Protector",WILD,
    "他们救受伤的动物!","they rescue hurt animals!")
n+=1; pn(s,n)
notes(s,"揭晓: 野生动物保护员!\n• 句型: 「野生动物保护员帮助 ___ 。」")

# Scenario 3: City messy
s=scenario_q_slide("🏘️","城市没有公园, 很乱。","The city has no park — it's messy.",ALERT,help_options)
n+=1; pn(s,n)
notes(s,"问题 3:\n• 「城市没有公园 — 小朋友没地方玩!」\n• 让学生举 A/B/C。\n• 翻页揭晓!")

s=scenario_a_slide("🏘️","城市没有公园, 很乱。","The city has no park — it's messy.",
    "🏙️","城市规划师","City Planner",CITY,
    "他们设计公园, 让城市更好!","they design parks and make the city better!")
n+=1; pn(s,n)
notes(s,"揭晓: 城市规划师!\n• 句型: 「城市规划师帮助 ___ 。」")

# 33. STEP 5 — What if we didn't have this job?
s=ns(); bg(s,CREAM); hb(s,"💭 如果没有...?  What If We Didn't Have...?",NAVY)
tb(s,0.4,0.85,9.2,0.45,"想一想 — 如果没有这个职业, 会怎样？",sz=20,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.30,"Think — what if we didn't have this job?",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
ifs=[("🌍","环境工程师","水会变 ___ ?","water gets ___?",ENV,"脏 / dirty"),
     ("🦁","野生动物保护员","动物会 ___ ?","animals will ___?",WILD,"受伤 / get hurt"),
     ("🏙️","城市规划师","城市会 ___ ?","city becomes ___?",CITY,"很乱 / messy")]
for i,(em,job_cn,q_cn,q_en,c,hint) in enumerate(ifs):
    y=1.70+i*0.95
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(y),Inches(9.2),Inches(0.85))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE
    sh.line.color.rgb=c; sh.line.width=Pt(2)
    tb(s,0.55,y+0.18,0.6,0.5,em,sz=28,a=PP_ALIGN.CENTER)
    tb(s,1.25,y+0.10,3.0,0.40,job_cn,sz=14,b=True,c=c)
    tb(s,1.25,y+0.45,3.0,0.30,"如果没有...",sz=10,c=GRAY)
    tb(s,4.50,y+0.10,3.0,0.40,q_cn,sz=14,b=True,c=DARK)
    tb(s,4.50,y+0.45,3.0,0.30,q_en,sz=10,c=GRAY)
    pill(s,7.70,y+0.25,1.70,0.35,hint,GOLD,sz=11)
sentence_frame_bar(s,4.65,
    "如果没有 ___ , ___ 会 ___ 。",
    "Without ___, ___ would ___.")
n+=1; pn(s,n)
notes(s,"5 分钟 — 加深意义:\n• 一个一个问: 「如果没有环境工程师呢?」\n• 让学生想 30 秒, 再说。\n• 老师总结: 「所以这个职业很重要!」\n• 让 1-2 个学生用句型说: 「如果没有 ___ , ___ 会 ___ 。」")

# 34. SESSION 2 DIVIDER
s=div("Session 2  下午 2:00–2:45","📚 复习 + 我会认 + 我会写  Review · Read · Write",GOLD,"📖"); n+=1; pn(s,n)

# 35. REVIEW — what we learned in Session 1
s=ns(); bg(s,CREAM); hb(s,"🔄 复习  Review · Session 1",NAVY)
tb(s,0.4,0.85,9.2,0.40,"早上学了什么？  What did we learn this morning?",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.25,9.2,0.30,"Quick recap before we read & write!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# 4 review cards (5 things to recall: 5 jobs / 3 mystery / 3 helpers / sentence)
recap=[("🩺","5 个常见职业","Common jobs", "医生·老师·警察·厨师·工程师", DOC),
       ("✨","3 个神秘职业","Mystery jobs","环境工程师·野生动物保护员·城市规划师", ENV),
       ("🆘","谁来帮忙","Who can help","河 / 海龟 / 城市 → 谁来?", HELP),
       ("💭","我的梦想","My dream","「我想当 ___ , 因为 ___ 。」", GOLD)]
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
notes(s,"5-8 分钟 — 让学生先说 (don't tell):\n• 「早上学了什么职业?」让学生抢答。\n• 全班跟读: 5 个常见 + 3 个神秘。\n• 让 1-2 个学生用句型: 「我想当 ___ 。」\n• 高年级: 加 「因为 ___ 」")

# 36. 我会认 — 5 jobs (familiar ones)
s=ns(); bg(s,CREAM); hb(s,"📖 我会认  I Can Read · 5 个职业",NAVY)
tb(s,0.4,0.78,9.2,0.32,"看字, 不看图 — 你认得吗？",sz=15,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.10,9.2,0.26,"Read the characters — can you say them?",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
vocab5=[("🩺","医生","Doctor",DOC),
        ("📚","老师","Teacher",TEACH),
        ("👮","警察","Police",POLICE),
        ("👨‍🍳","厨师","Chef",CHEF),
        ("👷","工程师","Engineer",ENG)]
for i,(em,cn,en,c) in enumerate(vocab5):
    x=0.4+i*1.88
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.55),Inches(1.78),Inches(2.95))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE
    sh.line.color.rgb=c; sh.line.width=Pt(2.5)
    # Big Chinese characters (focus on reading)
    tb(s,x+0.05,1.70,1.7,0.85,cn,sz=36,b=True,c=c,a=PP_ALIGN.CENTER)
    # Then emoji as visual support
    tb(s,x+0.05,2.65,1.7,0.65,em,sz=32,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.40,1.7,0.30,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    pill(s,x+0.34,3.85,1.10,0.30,"读 read",c,sz=10)
sentence_frame_bar(s,4.70,
    "这是 ___ 。 / 我看到 ___ 。",
    "This is a ___. / I see a ___.")
n+=1; pn(s,n)
notes(s,"10 分钟 — 认字:\n• 第一遍: 老师指字, 学生读 (不看图)。\n• 第二遍: 老师藏字, 学生看图说字。\n• 第三遍: 学生闭眼, 老师只说一个字, 学生开眼指。\n• 玩配对: 把字卡和图卡分开, 让学生配对。\n• K: 看图记词。\n• G1-3: 看字读词, 不看图。")

# 37. 我会写 — 医生 (stroke order practice)
s=ns(); bg(s,CREAM); hb(s,"✏️ 我会写 · 医生  I Can Write · Doctor",DOC)
tb(s,0.4,0.85,9.2,0.40,"练一练 — 写「医生」!",sz=22,b=True,c=DOC,a=PP_ALIGN.CENTER)
tb(s,0.4,1.25,9.2,0.30,"Practice writing 医生 (Doctor)",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# Big character display + stroke info side
char1=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.65),Inches(2.30),Inches(2.85))
char1.fill.solid(); char1.fill.fore_color.rgb=WHITE
char1.line.color.rgb=DOC; char1.line.width=Pt(3)
tb(s,0.4,1.95,2.30,1.95,"医",sz=130,b=True,c=DOC,a=PP_ALIGN.CENTER)
tb(s,0.4,4.10,2.30,0.30,"yī (medical)",sz=12,b=True,c=GRAY,a=PP_ALIGN.CENTER)
char2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(2.85),Inches(1.65),Inches(2.30),Inches(2.85))
char2.fill.solid(); char2.fill.fore_color.rgb=WHITE
char2.line.color.rgb=DOC; char2.line.width=Pt(3)
tb(s,2.85,1.95,2.30,1.95,"生",sz=130,b=True,c=DOC,a=PP_ALIGN.CENTER)
tb(s,2.85,4.10,2.30,0.30,"shēng (life/born)",sz=12,b=True,c=GRAY,a=PP_ALIGN.CENTER)
# Right: hints panel
panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.30),Inches(1.65),Inches(4.30),Inches(2.85))
panel.fill.solid(); panel.fill.fore_color.rgb=WHITE
panel.line.color.rgb=DOC; panel.line.width=Pt(2.5)
head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.30),Inches(1.65),Inches(4.30),Inches(0.50))
head.fill.solid(); head.fill.fore_color.rgb=DOC; head.line.fill.background()
tb(s,5.45,1.72,4.10,0.4,"✏️ 怎么写 How to write",sz=13,b=True,c=WHITE)
tb(s,5.45,2.30,4.10,0.40,"1️⃣ 「医」 — 7 笔",sz=14,b=True,c=DARK)
tb(s,5.45,2.65,4.10,0.30,"  先 ㄣ, 然后里面",sz=10,c=GRAY)
tb(s,5.45,3.05,4.10,0.40,"2️⃣ 「生」 — 5 笔",sz=14,b=True,c=DARK)
tb(s,5.45,3.40,4.10,0.30,"  从上往下, 从左往右",sz=10,c=GRAY)
tb(s,5.45,3.85,4.10,0.40,"📝 在田字格练 3 遍",sz=12,b=True,c=DOC)
tb(s,5.45,4.20,4.10,0.30,"Practice 3 times in grid paper",sz=9,c=GRAY)
sentence_frame_bar(s,4.65,
    "我会写「医生」! 我想当医生。",
    "I can write 医生! I want to be a doctor.")
n+=1; pn(s,n)
notes(s,"7-8 分钟 — 写字练习:\n• 老师在白板演示笔顺。\n• 学生跟着空写 (右手食指在空中)。\n• 然后在田字格练习本上写 3-5 遍。\n• 老师巡视, 帮助纠正笔顺。\n• 提示: 「医」里面像个人, 「生」像树苗发芽。\n• 完成后画 ✓。")

# 38. 我会写 — 老师 (stroke order practice)
s=ns(); bg(s,CREAM); hb(s,"✏️ 我会写 · 老师  I Can Write · Teacher",TEACH)
tb(s,0.4,0.85,9.2,0.40,"练一练 — 写「老师」!",sz=22,b=True,c=TEACH,a=PP_ALIGN.CENTER)
tb(s,0.4,1.25,9.2,0.30,"Practice writing 老师 (Teacher)",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
char3=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.65),Inches(2.30),Inches(2.85))
char3.fill.solid(); char3.fill.fore_color.rgb=WHITE
char3.line.color.rgb=TEACH; char3.line.width=Pt(3)
tb(s,0.4,1.95,2.30,1.95,"老",sz=130,b=True,c=TEACH,a=PP_ALIGN.CENTER)
tb(s,0.4,4.10,2.30,0.30,"lǎo (old/respected)",sz=12,b=True,c=GRAY,a=PP_ALIGN.CENTER)
char4=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(2.85),Inches(1.65),Inches(2.30),Inches(2.85))
char4.fill.solid(); char4.fill.fore_color.rgb=WHITE
char4.line.color.rgb=TEACH; char4.line.width=Pt(3)
tb(s,2.85,1.95,2.30,1.95,"师",sz=130,b=True,c=TEACH,a=PP_ALIGN.CENTER)
tb(s,2.85,4.10,2.30,0.30,"shī (master/teacher)",sz=12,b=True,c=GRAY,a=PP_ALIGN.CENTER)
panel2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.30),Inches(1.65),Inches(4.30),Inches(2.85))
panel2.fill.solid(); panel2.fill.fore_color.rgb=WHITE
panel2.line.color.rgb=TEACH; panel2.line.width=Pt(2.5)
head2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.30),Inches(1.65),Inches(4.30),Inches(0.50))
head2.fill.solid(); head2.fill.fore_color.rgb=TEACH; head2.line.fill.background()
tb(s,5.45,1.72,4.10,0.4,"✏️ 怎么写 How to write",sz=13,b=True,c=WHITE)
tb(s,5.45,2.30,4.10,0.40,"1️⃣ 「老」 — 6 笔",sz=14,b=True,c=DARK)
tb(s,5.45,2.65,4.10,0.30,"  从上往下, 像「土」+「ㄜ」",sz=10,c=GRAY)
tb(s,5.45,3.05,4.10,0.40,"2️⃣ 「师」 — 6 笔",sz=14,b=True,c=DARK)
tb(s,5.45,3.40,4.10,0.30,"  左边 ㄐ, 右边「巾」",sz=10,c=GRAY)
tb(s,5.45,3.85,4.10,0.40,"📝 在田字格练 3 遍",sz=12,b=True,c=TEACH)
tb(s,5.45,4.20,4.10,0.30,"Practice 3 times in grid paper",sz=9,c=GRAY)
sentence_frame_bar(s,4.65,
    "我会写「老师」! 老师教我们。",
    "I can write 老师! The teacher teaches us.")
n+=1; pn(s,n)
notes(s,"7-8 分钟:\n• 演示笔顺, 学生跟写。\n• 田字格练 3 遍。\n• 提示: 「老」上半像帽子, 「师」右边像旗子。\n• 完成 → 颁发「我会写」贴纸。\n• 时间够: 让学生用「老师」造句, e.g. 「我的老师叫 ___ 」")

# 39. SESSION 3 DIVIDER
s=div("Session 3  下午 3:00–4:30","🎨 Output: 小本子 + 职业帽子  Booklet + Career Hat",CITY,"💼"); n+=1; pn(s,n)

# 40. STEP 1 — Complete the Booklet
s=ns(); bg(s,CREAM); hb(s,"📖 完成小本子  Finish My Booklet",CITY)
tb(s,0.4,0.85,9.2,0.40,"画 + 写 — 把梦想记下来!",sz=20,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.25,9.2,0.30,"Draw and write — record your dream!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# 3 pages preview
pages=[("1","封面","Cover","「我的职业梦想」+ 你的名字","Title + your name",CITY),
       ("2","我想当","My Job","🔘 我想当 ___ 。 (圈一个或写)","Circle / write your job",GOLD),
       ("3","图画","My Picture","✏️ 画你做这个职业的样子","Draw yourself doing the job",HELP)]
for i,(num,cn,en,task_cn,task_en,c) in enumerate(pages):
    x=0.4+i*3.15
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.65),Inches(3.0),Inches(2.85))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE
    sh.line.color.rgb=c; sh.line.width=Pt(2.5)
    pill(s,x+0.10,1.75,1.0,0.35,f"第 {num} 页",c,sz=12)
    tb(s,x+0.05,2.20,2.9,0.45,cn,sz=18,b=True,c=c,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.62,2.9,0.30,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.10,3.05,2.8,0.55,task_cn,sz=12,b=True,c=DARK)
    tb(s,x+0.10,3.55,2.8,0.40,task_en,sz=9,c=GRAY)
    tb(s,x+0.10,4.00,2.8,0.35,"⏱️ 约 7 分钟",sz=9,c=GOLD,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.65,
    "高年级: 我想当 ___, 因为 ___ 。",
    "Higher level: I want to be ___, because ___.")
n+=1; pn(s,n)
notes(s,"20-25 分钟 — 学生独立做:\n• K 学生: 圈选 + 画画\n• G1-3: 写句子「我想当 ___, 因为 ___ 。」\n• 老师巡视, 帮助拼字。\n• 完成后加贴纸鼓励。")

# 41. STEP 2 INTRO — Career Hat (主推)
s=ns(); bg(s,CREAM); hb(s,"🎩 职业帽子  Career Hat!",GOLD)
tb(s,0.4,0.85,9.2,0.45,"做一顶你的职业帽子 — 戴上演一演!",sz=20,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.30,"Make a hat for your job — wear it and act!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# 4 example hats
hats=[("👨‍⚕️","医生帽","Doctor's hat","白色 + 红十字",DOC),
      ("👮","警察帽","Police hat","蓝色 + 警徽",POLICE),
      ("👨‍🍳","厨师帽","Chef's hat","白色, 高高的",CHEF),
      ("👷","工程师帽","Engineer hat","黄色 + 安全帽",ENG)]
for i,(em,cn,en,hint,c) in enumerate(hats):
    x=0.4+i*2.35
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.70),Inches(2.20),Inches(2.85))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE
    sh.line.color.rgb=c; sh.line.width=Pt(2.5)
    tb(s,x+0.05,1.80,2.10,1.00,em,sz=60,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.85,2.10,0.45,cn,sz=16,b=True,c=c,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.30,2.10,0.30,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.70,2.10,0.30,hint,sz=10,b=True,c=DARK,a=PP_ALIGN.CENTER)
    pill(s,x+0.40,4.10,1.40,0.30,"小小 "+cn[0:2],c,sz=10)
sentence_frame_bar(s,4.65,
    "我是小小 ___ ! 我戴 ___ 帽。",
    "I'm a junior ___! I wear a ___ hat.")
n+=1; pn(s,n)
notes(s,"主推项目 — 30 分钟:\n• 简单介绍每个职业的帽子样式 (上图)。\n• 学生选一种, 然后用纸条做帽圈, 加图案 / 颜色 / 装饰。\n• 戴上帽子, 演一演 (1 分钟): 「我是小小医生 — 听一听!」\n• 帽子贴在头围纸条上, 一戴就成形。\n• 鼓励创意! 学生可以做今天学的 8 个职业的任何一个。")

# 42. HAT MATERIALS & HOW-TO
s=ns(); bg(s,CREAM); hb(s,"🛠️ 帽子材料 & 做法  Hat Materials & How-To",GOLD)
tb(s,0.4,0.85,9.2,0.40,"简单 4 步 — 30 分钟做一顶帽子!",sz=20,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.25,9.2,0.30,"4 simple steps — make a hat in 30 minutes!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# Two columns: materials left, how-to right
mat=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.65),Inches(4.50),Inches(2.95))
mat.fill.solid(); mat.fill.fore_color.rgb=WHITE
mat.line.color.rgb=CHEF; mat.line.width=Pt(2.5)
matH=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.65),Inches(4.50),Inches(0.50))
matH.fill.solid(); matH.fill.fore_color.rgb=CHEF; matH.line.fill.background()
tb(s,0.55,1.72,4.30,0.4,"📦 材料 Materials",sz=13,b=True,c=WHITE)
materials=[("📄","彩色纸 / 卡纸","colored paper / cardstock"),
           ("✂️","剪刀, 胶水, 胶带","scissors, glue, tape"),
           ("🖍️","彩笔, 贴纸","markers, stickers"),
           ("🪶","羽毛, 棉花 (装饰)","feathers, cotton (deco)"),
           ("🪡","线 / 松紧带 (头围)","string / elastic band")]
for i,(em,cn,en) in enumerate(materials):
    y=2.30+i*0.42
    tb(s,0.55,y,0.40,0.30,em,sz=14)
    tb(s,0.95,y,3.85,0.30,cn,sz=11,b=True,c=DARK)
    tb(s,0.95,y+0.22,3.85,0.20,en,sz=8,c=GRAY)
# How-to right
how=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.10),Inches(1.65),Inches(4.50),Inches(2.95))
how.fill.solid(); how.fill.fore_color.rgb=WHITE
how.line.color.rgb=HELP; how.line.width=Pt(2.5)
howH=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.10),Inches(1.65),Inches(4.50),Inches(0.50))
howH.fill.solid(); howH.fill.fore_color.rgb=HELP; howH.line.fill.background()
tb(s,5.25,1.72,4.30,0.4,"🛠️ 4 步做法 4 Steps",sz=13,b=True,c=WHITE)
howto=[("1️⃣","剪一条头围纸条","Cut a paper headband"),
       ("2️⃣","加图案 (红十字 / 警徽 / 帽尖)","Add design (cross / badge / peak)"),
       ("3️⃣","装饰 (颜色 / 贴纸)","Decorate (colors / stickers)"),
       ("4️⃣","卷起来粘成圈 — 戴上!","Tape into a ring — wear it!")]
for i,(em,cn,en) in enumerate(howto):
    y=2.30+i*0.55
    tb(s,5.25,y,0.45,0.4,em,sz=14,b=True,c=HELP)
    tb(s,5.75,y,3.75,0.35,cn,sz=12,b=True,c=DARK)
    tb(s,5.75,y+0.27,3.75,0.25,en,sz=9,c=GRAY)
sentence_frame_bar(s,4.75,
    "我做 ___ 的帽子。我戴上, 我是小小 ___ !",
    "I made a ___ hat. I wear it — I'm a junior ___!")
n+=1; pn(s,n)
notes(s,"老师准备 (低投入):\n• 提前剪好头围纸条 (省时间)。\n• 准备彩纸 + 装饰品 1 桌一份。\n• 老师演示 1 顶 (e.g. 厨师帽: 白色, 加纸团做高顶)。\n• 学生 25 分钟做完, 5 分钟戴上展示。\n• 鼓励合作 — 同桌互相帮忙。")

# 43. PROJECT ALTERNATIVES — 2-3 other ideas
s=ns(); bg(s,CREAM); hb(s,"💡 其他项目选择  Other Project Ideas",NAVY)
tb(s,0.4,0.85,9.2,0.40,"不想做帽子？还可以做这些!",sz=20,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.25,9.2,0.30,"Don't want a hat? Try one of these!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# 3 alternative projects
alts=[("🎖️","职业徽章","Career Badge",
       "圆纸片 + 别针 + 图案", "Paper circle + pin + icon",
       "5 分钟 / 简单", DOC),
      ("🪧","职业小海报","Mini Poster",
       "1 张纸: 「我是 / 我用 / 我帮助」", "1 page: name + tools + helps",
       "15 分钟 / 中等", HELP),
      ("🎴","职业卡片","Career Card",
       "像游戏卡: 名字 + 工具 + 技能", "Like a trading card",
       "10 分钟 / 创意", CITY)]
for i,(em,cn,en,what_cn,what_en,time_cn,c) in enumerate(alts):
    x=0.4+i*3.15
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.65),Inches(3.0),Inches(3.05))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE
    sh.line.color.rgb=c; sh.line.width=Pt(2.5)
    tb(s,x+0.05,1.80,2.9,0.85,em,sz=52,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.75,2.9,0.45,cn,sz=18,b=True,c=c,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.20,2.9,0.30,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.10,3.60,2.8,0.40,what_cn,sz=11,b=True,c=DARK,a=PP_ALIGN.CENTER)
    tb(s,x+0.10,4.00,2.8,0.30,what_en,sz=9,c=GRAY,a=PP_ALIGN.CENTER)
    pill(s,x+0.50,4.35,2.0,0.30,"⏱️ "+time_cn,c,sz=10)
sentence_frame_bar(s,4.85,
    "我做 ___ 。我是小小 ___ !",
    "I made a ___. I'm a junior ___!")
n+=1; pn(s,n)
notes(s,"学生可以选项目 (老师推荐 1-2 个让全班一起做):\n• 帽子 (主推) — 视觉强, 戴上立刻有「I am」感觉\n• 徽章 — 最快, 适合时间紧\n• 海报 — 文字最多, 适合 G2-3, 句型练习\n• 卡片 — 最有趣, 像游戏卡, 可以收集交换\n• 全部 都让学生说: 「我是小小 ___ , 我帮助 ___ 。」\n• 老师可以选 2 种让学生选一种, 不用提供全部 4 种 (太复杂)。")

# 44. STEP 3 — Gallery Walk
s=ns(); bg(s,CREAM); hb(s,"🚶 画展!  Gallery Walk!",HELP)
tb(s,0.4,0.85,9.2,0.45,"展示你的工具箱 — 给同学看!",sz=22,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.30,"Show your toolbox to your friends!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# K vs G1-3 differentiation
k_box=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.75),Inches(4.50),Inches(2.85))
k_box.fill.solid(); k_box.fill.fore_color.rgb=WHITE
k_box.line.color.rgb=HELP; k_box.line.width=Pt(2.5)
khead=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.75),Inches(4.50),Inches(0.55))
khead.fill.solid(); khead.fill.fore_color.rgb=HELP; khead.line.fill.background()
tb(s,0.55,1.83,4.30,0.4,"🌱 K — 简单版",sz=14,b=True,c=WHITE)
tb(s,0.55,2.50,4.30,0.55,"我是小小 ___ 。",sz=22,b=True,c=DARK)
tb(s,0.55,3.10,4.30,0.40,"I'm a junior ___.",sz=12,c=GRAY)
tb(s,0.55,3.70,4.30,0.40,"💡 1 句 + 指着工具",sz=12,b=True,c=HELP)
tb(s,0.55,4.10,4.30,0.30,"1 sentence + point at tools",sz=9,c=GRAY)
g_box=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.10),Inches(1.75),Inches(4.50),Inches(2.85))
g_box.fill.solid(); g_box.fill.fore_color.rgb=WHITE
g_box.line.color.rgb=GOLD; g_box.line.width=Pt(2.5)
ghead=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.10),Inches(1.75),Inches(4.50),Inches(0.55))
ghead.fill.solid(); ghead.fill.fore_color.rgb=GOLD; ghead.line.fill.background()
tb(s,5.25,1.83,4.30,0.4,"🌟 G1-3 — 完整版",sz=14,b=True,c=WHITE)
tb(s,5.25,2.45,4.30,0.45,"我是小小 ___ 。",sz=18,b=True,c=DARK)
tb(s,5.25,2.95,4.30,0.45,"我用 ___ 。",sz=18,b=True,c=DARK)
tb(s,5.25,3.45,4.30,0.45,"我帮助 ___ 。",sz=18,b=True,c=DARK)
tb(s,5.25,4.00,4.30,0.40,"💡 3 句 + 演示 1 个工具",sz=12,b=True,c=GOLD)
tb(s,5.25,4.40,4.30,0.30,"3 sentences + demo 1 tool",sz=9,c=GRAY)
n+=1; pn(s,n)
notes(s,"15 分钟 — 玩法:\n• 桌面摆好工具箱, 学生轮流当「展览员」和「客人」。\n• 5 分钟换一组。\n• K: 1 句 + 指。\n• G1-3: 3 句 + 演示 (例如「我用听诊器 — 听一听!」)\n• 老师在旁拍照记录。")

# 45. MISSION COMPLETE!
s=ns(); bg(s,GOLD)
tb(s,1,0.7,8,0.7,"🏆 任务完成!",sz=46,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,1,1.55,8,0.5,"Mission Complete!",sz=22,c=WARM,a=PP_ALIGN.CENTER)
# 5 checkmarks
checks=[("✅","猜猜","Guessed"),
        ("✅","演演","Acted"),
        ("✅","神秘","Mystery"),
        ("✅","帮忙","Helped"),
        ("✅","梦想","Dreamed")]
for i,(em,cn,en) in enumerate(checks):
    x=0.4+i*1.88
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(2.50),Inches(1.78),Inches(1.50))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE
    sh.line.color.rgb=NAVY; sh.line.width=Pt(2.5)
    tb(s,x+0.05,2.60,1.7,0.50,em,sz=30,c=OK,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.20,1.7,0.40,cn,sz=14,b=True,c=NAVY,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.60,1.7,0.30,en,sz=9,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,0.4,4.30,9.2,0.45,"🌟 你是小小职业人! Future Career Hero!",sz=22,b=True,c=NAVY,a=PP_ALIGN.CENTER)
tb(s,0.4,4.80,9.2,0.30,"明天继续 — 探索更多职业!",sz=12,c=DARK,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"庆祝 (1 分钟):\n• 全班一起喊: 「我是小小职业人!」\n• 颁发徽章 / 贴纸 (在工具箱上贴)\n• 「下次见面 — 介绍更多酷职业!」")

# Save
import os
out=os.path.join(os.path.dirname(__file__),"day1_career.pptx")
prs.save(out)
print(f"Saved {out}  ({n} slides)")
