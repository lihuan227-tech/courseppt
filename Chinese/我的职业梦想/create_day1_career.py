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
    # Clamp so bar (height 0.65) never overflows the 5.625" slide bottom
    if t > 4.95: t = 4.95
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

def reveal_job_slide(emoji,job_cn,job_en,color,what_cn,what_en,where_cn,where_en,video_url=""):
    """Reveal slide — show the job icon, name, what they do, where they work + video."""
    s=ns(); bg(s,CREAM)
    hb(s,f"{emoji} 答案揭晓!  Answer Revealed!",color)
    # Big emoji + job name (slightly shorter to make room for video bar)
    bigbox=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.95),Inches(4.30),Inches(3.05))
    bigbox.fill.solid(); bigbox.fill.fore_color.rgb=WHITE
    bigbox.line.color.rgb=color; bigbox.line.width=Pt(3)
    tb(s,0.4,1.20,4.30,1.40,emoji,sz=120,a=PP_ALIGN.CENTER)
    tb(s,0.4,2.65,4.30,0.55,job_cn,sz=28,b=True,c=color,a=PP_ALIGN.CENTER)
    tb(s,0.4,3.20,4.30,0.40,job_en,sz=13,c=GRAY,a=PP_ALIGN.CENTER)
    pill(s,1.4,3.65,2.3,0.30,"我就是 ___ !",GOLD,sz=11)
    # RIGHT — facts (also shorter)
    panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.85),Inches(0.95),Inches(4.85),Inches(3.05))
    panel.fill.solid(); panel.fill.fore_color.rgb=WHITE
    panel.line.color.rgb=color; panel.line.width=Pt(2.5)
    tb(s,5.0,1.10,4.6,0.40,"🛠️ 做什么? What I do",sz=15,b=True,c=color)
    tb(s,5.0,1.55,4.6,0.40,what_cn,sz=14,b=True,c=DARK)
    tb(s,5.0,1.95,4.6,0.30,what_en,sz=10,c=GRAY)
    tb(s,5.0,2.40,4.6,0.40,"📍 在哪里? Where",sz=15,b=True,c=color)
    tb(s,5.0,2.80,4.6,0.40,where_cn,sz=14,b=True,c=DARK)
    tb(s,5.0,3.20,4.6,0.30,where_en,sz=10,c=GRAY)
    # NEW: video bar full-width
    vid=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(4.10),Inches(9.30),Inches(0.55))
    vid.fill.solid(); vid.fill.fore_color.rgb=WHITE; vid.line.color.rgb=color; vid.line.width=Pt(1.5)
    tb(s,0.5,4.18,1.0,0.4,"📺 视频",sz=13,b=True,c=color)
    tb(s,1.5,4.18,8.1,0.30,video_url if video_url else "(老师在这里贴 1 分钟职业小视频)",sz=10,c=DARK)
    tb(s,1.5,4.45,8.1,0.20,"Watch ~1 min so kids see the job for real",sz=8,c=GRAY)
    sentence_frame_bar(s,4.85,
        f"我是 {job_cn} 。我帮助 ___ 。",
        f"I am a {job_en}. I help ___.")
    return s

def try_it_slide(emoji,job_cn,job_en,color,activity_cn,activity_en,
                 step1,step2,step3,say_cn,say_en):
    """1-2 minute hands-on activity right after a job reveal.
    Students literally do something the job does."""
    s=ns(); bg(s,CREAM)
    hb(s,f"🎬 试一试 · 当 {job_cn}!  Try It Like a {job_en}!",color)
    # Activity title banner
    title=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.95),Inches(9.20),Inches(0.7))
    title.fill.solid(); title.fill.fore_color.rgb=color; title.line.fill.background()
    tb(s,0.5,1.00,1.0,0.6,emoji,sz=32,a=PP_ALIGN.CENTER)
    tb(s,1.6,1.05,8.0,0.40,activity_cn,sz=20,b=True,c=WHITE)
    tb(s,1.6,1.42,8.0,0.28,activity_en,sz=11,c=WHITE)
    # 3 steps
    for i,step in enumerate([step1,step2,step3]):
        x=0.4+i*3.10
        sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.85),Inches(2.95),Inches(2.55))
        sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=color; sh.line.width=Pt(2.5)
        badge=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x+1.225),Inches(1.95),Inches(0.50),Inches(0.50))
        badge.fill.solid(); badge.fill.fore_color.rgb=color; badge.line.fill.background()
        tb(s,x+1.225,2.04,0.50,0.35,str(i+1),sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
        tb(s,x+0.10,2.55,2.75,0.40,step[0],sz=44,a=PP_ALIGN.CENTER)
        tb(s,x+0.10,3.10,2.75,0.45,step[1],sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
        tb(s,x+0.10,3.55,2.75,0.70,step[2],sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    # Time pill
    pill(s,4.0,4.55,2.0,0.30,"⏱️ 1-2 分钟  1-2 min",GOLD,sz=11)
    # Sentence frame
    sentence_frame_bar(s,4.95,say_cn,say_en)
    return s

def mystery_job_q_slide(emoji,job_label,color,picture_label,q1_cn,q1_en,q2_cn,q2_en,video_url=""):
    """Mystery job — picture + 2 simple A-or-B questions, NO label revealed yet."""
    s=ns(); bg(s,CREAM)
    hb(s,f"{emoji} 神秘职业  Mystery Job!",color)
    tb(s,0.4,0.85,9.2,0.32,"看看图 — 他们在做什么？",sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
    tb(s,0.4,1.18,9.2,0.26,"Look — what are they doing?",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    # LEFT — bigger picture placeholder + video link below
    img_box=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.55),Inches(4.30),Inches(2.40))
    img_box.fill.solid(); img_box.fill.fore_color.rgb=IMGBG
    img_box.line.color.rgb=color; img_box.line.width=Pt(2.5)
    tb(s,0.4,2.05,4.30,0.55,emoji,sz=44,a=PP_ALIGN.CENTER)
    tb(s,0.4,2.65,4.30,0.35,picture_label,sz=12,b=True,c=color,a=PP_ALIGN.CENTER)
    tb(s,0.4,3.05,4.30,0.30,"📷 在这里贴真实照片",sz=9,c=LGRAY,a=PP_ALIGN.CENTER)
    tb(s,0.4,3.40,4.30,0.30,"Paste a real photo here",sz=9,c=LGRAY,a=PP_ALIGN.CENTER)
    # Video link box under picture
    vid_box=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(4.05),Inches(4.30),Inches(0.65))
    vid_box.fill.solid(); vid_box.fill.fore_color.rgb=WHITE
    vid_box.line.color.rgb=color; vid_box.line.width=Pt(1.5)
    tb(s,0.5,4.10,1.0,0.4,"📺 视频",sz=12,b=True,c=color)
    tb(s,1.5,4.10,3.15,0.30,video_url if video_url else "(老师在这里贴视频链接)",sz=9,c=DARK)
    tb(s,1.5,4.40,3.15,0.25,"Watch ~1 min before discussion",sz=8,c=GRAY)
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

# 1.5  UNIT PREVIEW — 5-Day Career Journey
s=ns(); bg(s,CREAM); hb(s,"🗺️ 5 天的职业之旅  Our 5-Day Career Journey",NAVY)
tb(s,0.4,0.85,9.2,0.34,"5 天 · 5 个主题 · 每天认识一位/一群改变世界的人",sz=15,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.20,9.2,0.28,"5 days · 5 themes · meet world-changers along the way",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
days_preview=[
    ("Day 1","认识职业世界","Discover Careers","🌍",NAVY,"今天 · 8 个职业"),
    ("Day 2","小小科学家","Little Scientists","🔬",ENV,"⭐ 爱因斯坦"),
    ("Day 3","小小企业家","Little Entrepreneurs","💡",GOLD,"⭐ 乔布斯"),
    ("Day 4","帮助别人","Helpers","❤️",DOC,"⭐ 医生 + 老师"),
    ("Day 5","AI 与未来","AI & the Future","🤖",CITY,"⭐ AI 公司 / 创始人"),
]
for i,(label,cn,en,em,cl,spotlight) in enumerate(days_preview):
    x=0.3+i*1.92
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.55),Inches(1.82),Inches(3.45))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE
    sh.line.color.rgb=cl; sh.line.width=Pt(2.5)
    badge=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x+0.1),Inches(1.65),Inches(0.55),Inches(0.55))
    badge.fill.solid(); badge.fill.fore_color.rgb=cl; badge.line.fill.background()
    tb(s,x+0.1,1.74,0.55,0.4,str(i+1),sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,x+0.7,1.7,1.1,0.3,label,sz=11,b=True,c=cl)
    tb(s,x+0.05,2.30,1.72,0.7,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.05,1.72,0.4,cn,sz=15,b=True,c=DARK,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.45,1.72,0.3,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    sep=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(x+0.25),Inches(3.85),Inches(1.32),Inches(0.02))
    sep.fill.solid(); sep.fill.fore_color.rgb=cl; sep.line.fill.background()
    tb(s,x+0.05,4.00,1.72,0.85,spotlight,sz=11,b=True,c=cl,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,5.20,
    "我最想认识 ___ 。",
    "I most want to meet ___.")
n+=1; pn(s,n)
notes(s,"30 秒 — 只是预告:\n• 「这 5 天我们会认识各种各样的人 — 看看他们怎么改变世界!」\n• 不展开 — Day 1 内容下面继续。")

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

# 4.5  📖 PICTURE BOOK INTRO — "What Do People Do All Day?" by Richard Scarry
# Helper for picture book scene slides (3-level questions: see → job → problem)
def book_scene_slide(emoji,scene_cn,scene_en,color,
                     l1_q_cn,l1_q_en,l1_hint,
                     l2_q_cn,l2_q_en,l2_hint,
                     l3_q_cn,l3_q_en,l3_hint):
    s=ns(); bg(s,CREAM)
    hb(s,f"{emoji} 看故事 · {scene_cn}",color)
    # Big picture placeholder on left half
    img=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.95),Inches(4.30),Inches(4.10))
    img.fill.solid(); img.fill.fore_color.rgb=IMGBG; img.line.color.rgb=color; img.line.width=Pt(3)
    tb(s,0.4,2.10,4.30,1.2,emoji,sz=120,a=PP_ALIGN.CENTER)
    tb(s,0.4,3.50,4.30,0.40,scene_cn,sz=18,b=True,c=color,a=PP_ALIGN.CENTER)
    tb(s,0.4,3.92,4.30,0.30,scene_en,sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,0.4,4.45,4.30,0.30,"📷 在这里贴书中场景截图",sz=9,c=LGRAY,a=PP_ALIGN.CENTER)
    tb(s,0.4,4.72,4.30,0.25,"Paste book scene screenshot",sz=8,c=LGRAY,a=PP_ALIGN.CENTER)
    # Right column — 3 layered question cards (level 1, 2, 3)
    levels=[
        ("🟢","L1 你看到？  See",OK,l1_q_cn,l1_q_en,l1_hint),
        ("🟡","L2 是什么工作？  Job",GOLD,l2_q_cn,l2_q_en,l2_hint),
        ("🟠","L3 解决什么问题？  Solves",CHEF,l3_q_cn,l3_q_en,l3_hint),
    ]
    for i,(em,label,cl,q_cn,q_en,hint) in enumerate(levels):
        y=0.95+i*1.40
        sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.85),Inches(y),Inches(4.85),Inches(1.30))
        sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(2.5)
        head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.85),Inches(y),Inches(4.85),Inches(0.40))
        head.fill.solid(); head.fill.fore_color.rgb=cl; head.line.fill.background()
        tb(s,4.95,y+0.04,4.65,0.32,f"{em}  {label}",sz=12,b=True,c=WHITE)
        tb(s,4.95,y+0.45,4.65,0.35,q_cn,sz=14,b=True,c=DARK)
        tb(s,4.95,y+0.78,4.65,0.25,q_en,sz=9,c=GRAY)
        tb(s,4.95,y+1.02,4.65,0.25,f"💡 老师提示: {hint}",sz=9,c=cl)
    return s

# 4.5b — BOOK INTRO (new video: interests can become your future job)
s=ns(); bg(s,CREAM); hb(s,"📖 听一个故事  Story Time",GOLD)
tb(s,0.4,0.85,9.2,0.50,"💡 你的兴趣 — 可能就是你未来的工作!",sz=22,b=True,c=GOLD,a=PP_ALIGN.CENTER)
tb(s,0.4,1.40,9.2,0.30,"Your hobby today — could become your job tomorrow!",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
# Cover panel
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.85),Inches(4.30),Inches(3.20))
sh.fill.solid(); sh.fill.fore_color.rgb=WARM; sh.line.color.rgb=GOLD; sh.line.width=Pt(3)
tb(s,0.4,2.10,4.30,1.0,"📖",sz=110,a=PP_ALIGN.CENTER)
tb(s,0.4,3.20,4.30,0.45,"我喜欢做的事",sz=22,b=True,c=GOLD,a=PP_ALIGN.CENTER)
tb(s,0.4,3.70,4.30,0.30,"What I Love → What I'll Do",sz=12,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,4.10,4.30,0.30,"📷 在这里贴书的封面",sz=9,c=LGRAY,a=PP_ALIGN.CENTER)
tb(s,0.4,4.40,4.30,0.30,"Paste book cover here",sz=9,c=LGRAY,a=PP_ALIGN.CENTER)
# Right: video link + listening prompts
panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.85),Inches(1.85),Inches(4.85),Inches(3.20))
panel.fill.solid(); panel.fill.fore_color.rgb=WHITE; panel.line.color.rgb=GOLD; panel.line.width=Pt(2.5)
head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.85),Inches(1.85),Inches(4.85),Inches(0.50))
head.fill.solid(); head.fill.fore_color.rgb=GOLD; head.line.fill.background()
tb(s,5.00,1.93,4.65,0.4,"🎬 听故事  Listen",sz=14,b=True,c=WHITE)
tb(s,5.00,2.45,4.65,0.30,"📺 视频链接 / Video:",sz=11,b=True,c=GOLD)
tb(s,5.00,2.75,4.65,0.30,"youtube.com/watch?v=y-jgklhlX3A",sz=10,b=True,c=NAVY)
tb(s,5.00,3.05,4.65,0.30,"(老师: ~3-5 分钟, 看完一起聊)",sz=10,c=GRAY)
tb(s,5.00,3.50,4.65,0.30,"👂 听的时候, 我们一起找:",sz=11,b=True,c=GOLD)
tb(s,5.00,3.80,4.65,0.30,"1. 主角喜欢做什么?",sz=10,c=DARK)
tb(s,5.00,4.05,4.65,0.30,"2. 长大后变成什么工作?",sz=10,c=DARK)
tb(s,5.00,4.30,4.65,0.30,"3. 兴趣 和 工作 怎么连起来的?",sz=10,c=DARK)
tb(s,5.00,4.65,4.65,0.30,"📌 听完一起回答 4 个问题!",sz=10,b=True,c=GOLD)
sentence_frame_bar(s,5.20,
    "故事里 ___ 喜欢 ___ , 长大变成 ___ 。",
    "In the story, ___ loved ___ and became a ___.")
n+=1; pn(s,n)
notes(s,"3-5 分钟:\n• 主资源: https://www.youtube.com/watch?v=y-jgklhlX3A\n• 听之前先说: 「这个故事告诉我们 — 你今天喜欢的, 可能就是你长大后做的!」\n• 听完不要急着讲解 — 翻页让学生自己回答 4 个问题")

# 4.5c — POST-STORY DISCUSSION (4 lead-in questions)
s=ns(); bg(s,CREAM); hb(s,"🤔 听完故事 · 一起讨论  Let's Talk",GOLD)
tb(s,0.4,0.85,9.2,0.40,"4 个问题 — 一个比一个深!",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.28,"4 questions — each goes a bit deeper.",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
discuss=[
    ("🟢","L1","主角喜欢做什么？","What did the main character love?",DOC),
    ("🟡","L2","他长大后变成什么工作？","What job did they grow up to become?",HELP),
    ("🟠","L3","兴趣怎么变成工作的？","How did the hobby BECOME the job?",CHEF),
    ("🔴","L4","你的兴趣 — 可能是什么工作？","Your hobby — could be what job?",NAVY),
]
for i,(em,lvl,q_cn,q_en,cl) in enumerate(discuss):
    col=i%2; row=i//2
    x=0.4+col*4.7; y=1.65+row*1.55
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(4.50),Inches(1.40))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(2.5)
    tb(s,x+0.10,y+0.10,0.7,0.55,em,sz=28,a=PP_ALIGN.CENTER)
    tb(s,x+0.85,y+0.10,3.55,0.35,lvl,sz=12,b=True,c=cl)
    tb(s,x+0.85,y+0.45,3.55,0.45,q_cn,sz=13,b=True,c=DARK)
    tb(s,x+0.85,y+0.95,3.55,0.30,q_en,sz=10,c=GRAY)
sentence_frame_bar(s,4.85,
    "我觉得 ___ , 因为 ___ 。",
    "I think ___, because ___.")
n+=1; pn(s,n)
notes(s,"5-6 分钟:\n• L1 简单, 大家一起回答\n• L2 让 1-2 个孩子说\n• L3 是关键 — 引出「兴趣 → 工作」的桥梁\n• L4 是个人的 — 让 3-4 个孩子说自己的想法\n• 不给「正确答案」 — 接受所有合理回答")

# 4.5d — INTEREST → CAREER MAPPING + ANSWER KEY (paired rows)
s=ns(); bg(s,CREAM); hb(s,"🔗 兴趣 → 工作  Interest → Career  (答案 Answer Key)",GOLD)
tb(s,0.4,0.85,9.2,0.40,"看! 一个兴趣 → 可以变成一个工作:",sz=16,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.25,9.2,0.28,"See — each hobby maps to a real job!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# Aligned pairs (answer key)
pairs=[
    (("🎨","画画"),       ("🎨","画家 / 设计师")),
    (("🧱","搭积木"),     ("👷","工程师 / 建筑师")),
    (("🐶","爱小动物"),   ("🦒","兽医")),
    (("🍳","做饭"),       ("👨‍🍳","厨师")),
    (("⚽","运动"),       ("⚽","运动员 / 教练")),
    (("🎮","玩游戏"),     ("🎮","游戏设计师")),
]
# Column headers
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.60),Inches(4.20),Inches(0.40))
sh.fill.solid(); sh.fill.fore_color.rgb=GOLD; sh.line.fill.background()
tb(s,0.4,1.65,4.20,0.30,"🎨 兴趣 Interests",sz=13,b=True,c=WHITE,a=PP_ALIGN.CENTER)
# Arrow column header
tb(s,4.65,1.65,0.7,0.30,"→",sz=22,b=True,c=GOLD,a=PP_ALIGN.CENTER)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.40),Inches(1.60),Inches(4.20),Inches(0.40))
sh.fill.solid(); sh.fill.fore_color.rgb=NAVY; sh.line.fill.background()
tb(s,5.40,1.65,4.20,0.30,"💼 工作 Jobs",sz=13,b=True,c=WHITE,a=PP_ALIGN.CENTER)
# Aligned rows with arrow between
for i,(left,right) in enumerate(pairs):
    y=2.10+i*0.48
    em_l,cn_l=left; em_r,cn_r=right
    # Left card
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(y),Inches(4.20),Inches(0.40))
    sh.fill.solid(); sh.fill.fore_color.rgb=WARM; sh.line.color.rgb=GOLD; sh.line.width=Pt(1.5)
    tb(s,0.55,y+0.05,0.5,0.30,em_l,sz=16,a=PP_ALIGN.CENTER)
    tb(s,1.10,y+0.06,3.4,0.30,cn_l,sz=13,b=True,c=DARK)
    # Arrow
    tb(s,4.65,y,0.7,0.40,"→",sz=22,b=True,c=GOLD,a=PP_ALIGN.CENTER)
    # Right card
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.40),Inches(y),Inches(4.20),Inches(0.40))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=NAVY; sh.line.width=Pt(1.5)
    tb(s,5.55,y+0.05,0.5,0.30,em_r,sz=16,a=PP_ALIGN.CENTER)
    tb(s,6.10,y+0.06,3.4,0.30,cn_r,sz=13,b=True,c=NAVY)
# Answer key bottom note (small, doesn't overflow)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(5.05),Inches(9.20),Inches(0.45))
sh.fill.solid(); sh.fill.fore_color.rgb=GOLD; sh.line.fill.background()
tb(s,0.5,5.10,9.0,0.35,"✅ 答案 Answer Key — 一个兴趣可以变很多工作, 还有更多!",sz=12,b=True,c=WHITE,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"4-5 分钟:\n• 一行一行读: 「喜欢 画画 → 可以当 画家」\n• 强调: 不止 1 个工作 — 一个兴趣可以变成 很多 工作!\n• 让学生加: 「喜欢 画画 还可以当 ___ ?」 (动画师 / 服装设计师 / 漫画家...)")

# 4.5e — Key principle (兴趣 + 努力 + 帮别人 = 工作)
s=ns(); bg(s,GOLD)
tb(s,1,1.0,8,0.8,"🌟 一个大发现!  A Big Discovery!",sz=26,b=True,c=WHITE,a=PP_ALIGN.CENTER)
big=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.6),Inches(2.05),Inches(8.8),Inches(1.85))
big.fill.solid(); big.fill.fore_color.rgb=WHITE; big.line.color.rgb=NAVY; big.line.width=Pt(4)
tb(s,0.6,2.30,8.8,0.7,"我喜欢的 + 我能帮别人的 = 我的工作",sz=26,b=True,c=NAVY,a=PP_ALIGN.CENTER)
tb(s,0.6,3.10,8.8,0.40,"What I love  +  what helps others  =  my future job",sz=14,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,0.6,3.55,8.8,0.30,"💡 兴趣 + 努力 + 时间 → 工作",sz=14,b=True,c=GOLD,a=PP_ALIGN.CENTER)
tb(s,1,4.20,8,0.40,"💭 想一想: 我的「喜欢」 + 「擅长」, 可能是什么工作?",sz=14,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,1,4.65,8,0.30,"Think: My 'love' + 'good at' — what job could that be?",sz=11,c=WARM,a=PP_ALIGN.CENTER)
tb(s,1,5.05,8,0.30,"长大不可怕 — 你的兴趣会带你去!",sz=12,b=True,c=WHITE,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)
notes(s,"1 分钟 — 重要时刻:\n• 让全班齐读 2 遍: 「我喜欢的 + 我能帮别人的 = 我的工作」\n• 这是这本书的核心信息 — 长大不可怕, 你已经在走向你的工作了\n• 后面的所有职业都会回到这一点")

# 7-16. Five GUESS clue + reveal pairs
guesses=[
    ("🩺","Doctor",DOC,
     [("我能听见你的心跳。","I can hear your heartbeat."),
      ("我用手摸一摸, 就知道你哪里不舒服。","I touch you and can tell where it hurts.")],
     "医生","Doctor","看病、给药、做手术","check, medicine, surgery","医院 / 诊所","hospital / clinic",
     "搜: 'doctor for kids' (Nat Geo Kids / SciShow Kids)"),
    ("📚","Teacher",TEACH,
     [("我每天要说很多话, 嗓子常常很累。","I talk a LOT every day — my throat often gets tired."),
      ("我最喜欢看你眼睛里有「啊! 我懂了」的样子。","My favorite thing is the 'aha — I get it!' look in your eyes.")],
     "老师","Teacher","教课、改作业、讲故事","teach, grade, tell stories","学校 / 教室","school / classroom",
     "搜: 'a day in the life of a teacher for kids'"),
    ("👮","Police",POLICE,
     [("我有一个手电筒和一本小本子。","I have a flashlight and a small notebook."),
      ("我能让一辆车停下来, 也能让它开走。","I can make a car stop — and make it go again.")],
     "警察","Police Officer","抓坏人、保护大家","catch bad guys, protect people","街上 / 警察局","street / police station",
     "搜: 'police officer for kids' (Cocomelon / Sesame Street)"),
    ("👨‍🍳","Chef",CHEF,
     [("我每天闻到的味道最多。","I smell the most flavors every day."),
      ("我用刀的时候不能开小差。","I cannot get distracted when I use a knife.")],
     "厨师","Chef","做菜、做点心","cook meals, make desserts","餐厅 / 厨房","restaurant / kitchen",
     "搜: 'kid chef cooking' or 'MasterChef Junior' (1-min clip)"),
    ("👷","Engineer",ENG,
     [("我喜欢画图, 但不是画画。","I love to draw — but not for art."),
      ("我画的图后来会变成大楼或者大桥。","My drawings later become buildings and bridges.")],
     "工程师","Engineer","设计、建造、修理","design, build, repair","公司 / 工地","company / construction site",
     "搜: 'what does an engineer do for kids' (Crash Course Kids)"),
]
try_it_data={
    "医生": ("👂 听心跳游戏","Listen for a Heartbeat",
            ("🤝","两人一组","Pair up"),
            ("👂","耳朵贴对方手腕","Ear on partner's wrist"),
            ("💓","数 10 秒里的心跳","Count beats for 10 sec"),
            "我数到 ___ 下心跳!","I counted ___ heartbeats!"),
    "老师": ("🎓 30 秒小老师","30-Second Teacher",
            ("🤔","想一件你会的事","Pick something you can do"),
            ("✋","上台 30 秒教大家","Teach for 30 sec"),
            ("👏","全班鼓掌评分","Class claps a score"),
            "今天我教大家 ___ 。","Today I teach you ___."),
    "警察": ("🚦 我来指挥交通","Direct the Traffic",
            ("✋","老师当车, 学生当警察","Teacher = car, student = police"),
            ("🛑","学生举手 → 老师停, 挥手 → 老师走","Hand up = stop, wave = go"),
            ("🔄","换人玩, 各试一次","Switch roles"),
            "停! 走! 谢谢!","Stop! Go! Thank you!"),
    "厨师": ("🍳 神秘菜单挑战","Mystery Menu Challenge",
            ("🎲","抽 3 种食材 (老师准备)","Draw 3 ingredients"),
            ("🧠","30 秒想一道新菜","30 sec to invent a dish"),
            ("📣","给菜起个名 + 介绍","Name + present your dish"),
            "我做的是 ___ , 用了 ___ 、___ 、___ 。","My dish is ___, with ___, ___, ___."),
    "工程师": ("🌉 1 张纸的桥","The 1-Paper Bridge",
            ("📄","拿 1 张 A4 纸","Take ONE sheet of paper"),
            ("📐","折 / 卷 / 设计成桥","Fold / roll into a bridge"),
            ("🔨","看能不能撑住一支笔","Test: can it hold a pencil?"),
            "我的桥能撑住 ___ 。","My bridge held ___."),
}
for em,j_en,col,clues,job_cn,job_en,what_cn,what_en,where_cn,where_en,video in guesses:
    s=guess_clue_slide(em,j_en,col,clues,
        f"我猜是 ___ !",
        f"I guess it's a ___!")
    n+=1; pn(s,n)
    notes(s,f"猜谜 (1-2 分钟):\n• 老师慢慢读 2 个提示, 让学生先想。\n• 让 2-3 个学生举手猜: 「我猜是 ___ !」\n• 不要直接给答案 — 让多个孩子先猜。\n• 然后翻页揭晓!")

    s=reveal_job_slide(em,job_cn,j_en,col,what_cn,what_en,where_cn,where_en,video_url=video)
    n+=1; pn(s,n)
    notes(s,f"揭晓 {job_cn}!\n• 让全班一起喊: 「{job_cn}!」\n• 看 1 分钟职业小视频 (老师提前在 YouTube 找好)\n• 复习: 做什么? 在哪里?\n• 用句型: 「我是 {job_cn}, 我帮助 ___ 。」")

    # NEW: Try-it micro-activity right after reveal
    if job_cn in try_it_data:
        a_cn,a_en,st1,st2,st3,say_cn,say_en=try_it_data[job_cn]
        s=try_it_slide(em,job_cn,j_en,col,a_cn,a_en,st1,st2,st3,say_cn,say_en)
        n+=1; pn(s,n)
        notes(s,f"试一试 (~2 分钟):\n• 让全班真的做一遍 — 不只看, 要做!\n• 给积极参与的孩子一个 ✓ 在白板上。\n• 简单收尾: 「现在你也是小小 {job_cn}!」")

# 16.5. COOL JOBS SHOWCASE — wide variety of "wow" jobs
s=ns(); bg(s,CREAM); hb(s,"🌟 还有这些超酷的职业!  More Cool Careers!",GOLD)
tb(s,0.4,0.85,9.2,0.34,"世界上有几千种职业 — 看看这 9 个超酷的!",sz=15,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.20,9.2,0.26,"There are thousands of jobs — here are 9 cool ones!",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
cool_jobs=[
    ("🚀","宇航员","Astronaut","在太空漂!",NAVY),
    ("✈️","飞行员","Pilot","开飞机看云!",POLICE),
    ("🦒","兽医","Vet","给小动物看病!",WILD),
    ("🎮","游戏设计师","Game Designer","你玩的游戏 = 他做的!",CITY),
    ("🌊","海洋生物学家","Marine Biologist","和海豚做朋友!",ENV),
    ("🚒","消防员","Firefighter","坐红色大车救人!",DOC),
    ("🎬","动画师","Animator","你看的卡通 = 他画的!",CHEF),
    ("🍫","巧克力师","Chocolatier","每天发明新糖果!",GOLD),
    ("🎩","魔术师","Magician","让东西「消失」!",HELP),
]
for i,(em,cn,en,wow,cl) in enumerate(cool_jobs):
    col=i%3; row=i//3
    # 3×3 grid — tighter cards to fit 9
    x=0.30+col*3.20; y=1.45+row*1.20
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(3.0),Inches(1.05))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(2.5)
    tb(s,x+0.10,y+0.07,0.7,0.55,em,sz=28,a=PP_ALIGN.CENTER)
    tb(s,x+0.85,y+0.08,2.10,0.32,cn,sz=14,b=True,c=cl)
    tb(s,x+0.85,y+0.36,2.10,0.24,en,sz=9,c=GRAY)
    tb(s,x+0.10,y+0.65,2.85,0.35,f"💫 {wow}",sz=10,c=DARK)
sentence_frame_bar(s,5.10,
    "我最想知道 ___ 的工作!",
    "I most want to know about ___'s job!")
n+=1; pn(s,n)
notes(s,"3-4 分钟:\n• 快速过 9 个 — 每个 ~20 秒\n• 让举手: 「最想当哪个? 为什么?」\n• 不深入 — 只是激发兴趣\n• 提醒: 「这只是 9 个! 长大有几千种工作可以选!」")

# 16.6  🔍 SILHOUETTE GUESS GAME — pick 3 jobs from the Cool Jobs Showcase (slide 31)
# Same format on each slide: silhouette/clues on top, ANSWER strip at bottom
# (teacher covers the answer strip with a sticky-note OR just doesn't scroll until students guess)
def silhouette_guess_slide(emoji,job_cn,job_en,color,clue1_cn,clue1_en,clue2_cn,clue2_en,clue3_cn,clue3_en,fact):
    s=ns(); bg(s,CREAM)
    hb(s,"🔍 看图猜职业!  Guess the Job from the Silhouette!",color)
    # LEFT — silhouette area (dark colored box with shadow emoji)
    img=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.95),Inches(4.30),Inches(3.10))
    img.fill.solid(); img.fill.fore_color.rgb=DARK; img.line.color.rgb=color; img.line.width=Pt(3)
    # Silhouette emoji in semi-faded color
    tb(s,0.4,1.55,4.30,1.5,emoji,sz=140,c=color,a=PP_ALIGN.CENTER)
    tb(s,0.4,3.30,4.30,0.40,"⬛ 谁的剪影？",sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,0.4,3.70,4.30,0.30,"Whose silhouette?",sz=11,c=LGRAY,a=PP_ALIGN.CENTER)
    # RIGHT — 3 clues stacked (gradually easier)
    panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.85),Inches(0.95),Inches(4.85),Inches(3.10))
    panel.fill.solid(); panel.fill.fore_color.rgb=WHITE; panel.line.color.rgb=color; panel.line.width=Pt(2.5)
    head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.85),Inches(0.95),Inches(4.85),Inches(0.50))
    head.fill.solid(); head.fill.fore_color.rgb=color; head.line.fill.background()
    tb(s,5.0,1.03,4.65,0.4,"🔎 3 个提示  3 Clues",sz=14,b=True,c=WHITE)
    clues=[(clue1_cn,clue1_en),(clue2_cn,clue2_en),(clue3_cn,clue3_en)]
    for i,(cn,en) in enumerate(clues):
        y=1.65+i*0.78
        badge=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(5.0),Inches(y),Inches(0.40),Inches(0.40))
        badge.fill.solid(); badge.fill.fore_color.rgb=color; badge.line.fill.background()
        tb(s,5.0,y+0.04,0.40,0.32,str(i+1),sz=14,b=True,c=WHITE,a=PP_ALIGN.CENTER)
        tb(s,5.50,y-0.02,4.10,0.45,cn,sz=14,b=True,c=DARK)
        tb(s,5.50,y+0.40,4.10,0.30,en,sz=9,c=GRAY)
    # BOTTOM — guess sentence frame (NO answer shown — teacher reveals verbally)
    rev=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(4.20),Inches(9.30),Inches(0.95))
    rev.fill.solid(); rev.fill.fore_color.rgb=WARM; rev.line.color.rgb=color; rev.line.width=Pt(2.5)
    tb(s,0.5,4.30,1.6,0.35,"💬 我来猜!",sz=14,b=True,c=color)
    tb(s,2.0,4.27,7.4,0.45,"我猜是 ___ !",sz=22,b=True,c=DARK)
    tb(s,2.0,4.75,7.4,0.30,"I guess it's a ___!",sz=11,c=GRAY)
    return s

# Round 1: 🚀 Astronaut
s=silhouette_guess_slide("🚀","宇航员","Astronaut",NAVY,
    "我穿的衣服自带氧气罐, 不带不行。","My suit comes with its own oxygen tank — I can't go without it.",
    "我「上班的地方」 — 没有空气, 也没有重力!","My 'workplace' has no air — and no gravity!",
    "我能看到整个地球, 像一个蓝色的球。","I can see the whole Earth — like a blue marble.",
    "上过太空 = 全世界只有几百人!")
n+=1; pn(s,n)
notes(s,"猜谜 (1-2 分钟):\n• 慢慢读 3 个提示 — 让学生在第 1 个就开始想\n• 第 2 提示加难度, 第 3 提示几乎送答案\n• 让学生喊: 「我猜是 ___ !」\n• 答案: 宇航员 — 老师等学生猜过后口头揭晓 (PPT 不显示)")

# Round 2: 🦒 Vet
s=silhouette_guess_slide("🦒","兽医","Vet (Animal Doctor)",WILD,
    "我的「病人」从来不会用嘴说哪里疼。","My 'patients' can never tell me with words where it hurts.",
    "我会被汪汪、喵喵、咩咩声「叫」去工作。","I get called to work by barking, meowing, even baa-ing.",
    "我给小狗、小猫、有时还有小老虎打针。","I give shots to dogs, cats — sometimes baby tigers.",
    "你家的宠物生病了 → 找我!")
n+=1; pn(s,n)
notes(s,"猜谜 (1-2 分钟):\n• 「不会说话的病人」是关键 — 让学生想是谁\n• 提到「小老虎」会让全班兴奋\n• 联想: 「你家有宠物吗? 它生病时谁帮忙?」")

# Round 3: 🎮 Game Designer
s=silhouette_guess_slide("🎮","游戏设计师","Game Designer",CITY,
    "我做的东西 — 你每天都想玩, 玩到爸爸妈妈喊「停!」","I make things you want to play — until your parents shout 'stop!'",
    "我画一个小角色, 让他跳过 100 个关卡。","I draw a tiny character, then send him through 100 levels.",
    "Minecraft、Roblox、马里奥 — 都是这种工作做的。","Minecraft, Roblox, Mario — all made by people in this job.",
    "你最喜欢的游戏 = 一群「他们」做的!")
n+=1; pn(s,n)
notes(s,"猜谜 (1-2 分钟):\n• 提示 1 学生立刻笑 — 因为戳到他们日常\n• 提到具体游戏 (Minecraft / Roblox / 马里奥) 学生绝对炸\n• 收尾: 「这个工作, 30 年前根本没有 — 现在是最热门的工作之一!」")

# 16.7  SPOTLIGHT — 机器人工程师 (Robotics Engineer)
ROBOT=RGBColor(0x55,0x6B,0x83)  # cool steel
s=guess_clue_slide("🤖","Robotics Engineer",ROBOT,
    [("我的「员工」不需要睡觉, 也不需要吃饭。","My 'workers' don't need sleep — or food."),
     ("我教机器一步一步做事, 它们才会动。","I teach machines step by step — only then do they move.")],
    "我猜是 ___ !",
    "I guess it's a ___!")
n+=1; pn(s,n)
notes(s,"猜谜 (1-2 分钟):\n• 「不睡觉不吃饭的员工」 — 让学生想这是什么\n• 引导: 「在工厂里, 在医院里有这样的员工吗?」")

# (Robotics reveal + Try It + Fun Facts slides deleted per user request)

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

# 18. STEP 3.5 INTRO — 神秘职业! (improved cards with teaser hints)
s=ns(); bg(s,CREAM); hb(s,"✨ 神秘职业  Mystery Jobs!",ENV)
tb(s,0.4,0.85,9.2,0.50,"我再跟你介绍三个职业。",sz=22,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.40,9.2,0.30,"3 more jobs — have you heard of them?",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
mysteries=[
    ("🌍","Mystery 1","和地球有关的工作","About our Earth",      "💧 水 · 空气 · 树",ENV),
    ("🦁","Mystery 2","和动物有关的工作","About wild animals",   "🦒 救动物 · 看动物",WILD),
    ("🏙️","Mystery 3","和我们的「家」有关","About where we live","🗺️ 街道 · 公园 · 路",CITY),
]
for i,(em,tag,hint_cn,hint_en,domain,c) in enumerate(mysteries):
    x=0.4+i*3.15
    # Card box
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.85),Inches(3.0),Inches(2.95))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE
    sh.line.color.rgb=c; sh.line.width=Pt(3)
    # Tag badge at top
    badge=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x+0.65),Inches(1.95),Inches(1.70),Inches(0.36))
    badge.fill.solid(); badge.fill.fore_color.rgb=c; badge.line.fill.background()
    tb(s,x+0.65,2.00,1.70,0.30,tag,sz=12,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    # Big emoji
    tb(s,x+0.05,2.40,2.9,0.85,em,sz=58,a=PP_ALIGN.CENTER)
    # Teaser hint (CN bold, EN italics gray)
    tb(s,x+0.05,3.35,2.9,0.40,hint_cn,sz=15,b=True,c=c,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.78,2.9,0.30,hint_en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    # Domain keywords (separator + small text)
    sep=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(x+0.50),Inches(4.18),Inches(2.0),Inches(0.02))
    sep.fill.solid(); sep.fill.fore_color.rgb=c; sep.line.fill.background()
    tb(s,x+0.05,4.30,2.9,0.40,domain,sz=11,c=DARK,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.95,
    "我猜是 ___ 的工作?",
    "I guess this is a ___ job?")
n+=1; pn(s,n)
notes(s,"30 秒过渡:\n• 「现在 — 3 个新职业! 看看你能不能猜出来。」\n• 不要直接说出职业名 — 让学生从「领域提示」里想\n• 我们一个一个看图猜!")

# 19-24. Three MYSTERY JOBS — Q + Reveal pairs

# Mystery 1: 环境工程师
s=mystery_job_q_slide("🌍","Environmental Engineer",ENV,
    "💧 检测水 / 清洁",
    "他们在做什么？","What are they doing?",
    "是在玩水, 还是让水变干净？","Playing with water — or cleaning it?",
    video_url="搜: 'environmental engineer for kids'  (YouTube)")
n+=1; pn(s,n)
notes(s,"3-4 分钟:\n• 投影一张「人们检测水/清洁水」的真实照片 (建议老师贴一张)。\n• 老师追问: 「他们手里拿什么? 那个瓶子是干什么的?」\n• 让学生先猜, 不要直接给答案!\n• 让 2-3 个孩子说: 「我觉得他们在 ___ 。」")

s=mystery_job_label_slide("🌍","环境工程师","Environmental Engineer",
    "保护水和空气的人","helps keep water and air clean",ENV,
    "他们让水变干净, 让空气变好!")
n+=1; pn(s,n)
notes(s,"揭晓 (1-2 分钟):\n• 「答案是 — 环境工程师!」\n• 让全班跟读 3 遍: 「环境工程师 — 保护水和空气的人」\n• 简单解释: 「水脏了 → 他们让水变干净。空气脏了 → 他们检查空气。」")

# Mystery 2: 野生动物保护员
s=mystery_job_q_slide("🦁","Wildlife Protector",WILD,
    "🦒 救动物 / 观察",
    "他们在做什么？","What are they doing?",
    "是在抓动物, 还是保护动物？","Catching animals — or protecting them?",
    video_url="搜: 'wildlife ranger video for kids'  (YouTube / Nat Geo Kids)")
n+=1; pn(s,n)
notes(s,"3-4 分钟:\n• 投影一张「人们在帮助 / 观察动物」的真实照片。\n• 老师追问: 「动物开心吗? 这些人对动物好不好?」\n• 让学生猜, 不直接给答案!\n• 让 2-3 个孩子说: 「我觉得他们在 ___ 。」")

s=mystery_job_label_slide("🦁","野生动物保护员","Wildlife Protector",
    "保护动物的人","protects animals",WILD,
    "他们救受伤的动物, 看着不让坏人伤害动物!")
n+=1; pn(s,n)
notes(s,"揭晓:\n• 跟读: 「野生动物保护员 — 保护动物的人」\n• 解释: 「动物受伤 → 他们救; 有人想伤害动物 → 他们保护。」")

# Mystery 3: 城市规划师
s=mystery_job_q_slide("🏙️","City Planner",CITY,
    "🗺️ 城市地图 / 模型",
    "他们在做什么？","What are they doing?",
    "是在画画, 还是设计城市？","Drawing pictures — or designing a city?",
    video_url="搜: 'what is urban planning for kids'  (YouTube)")
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

# 33.5  JOBS KEEP CHANGING — Session 1 closing slide (3 columns: before / now / future)
s=ns(); bg(s,CREAM); hb(s,"⏳ 工作一直在变!  Jobs Keep Changing!",HELP)
tb(s,0.4,0.85,9.2,0.40,"以前 → 现在 → 未来 — 工作一直在变!",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.28,"Past → Now → Future — jobs change all the time!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# 3 columns
cols=[
    ("📜","以前",BROWN,"Before",[("✉️","送信"),("📚","抄书"),("🐎","养马"),("🕯️","点路灯")]),
    ("🌍","现在",NAVY,"Now",[("🎮","游戏设计师"),("🤖","AI 工程师"),("📺","YouTuber"),("🚁","无人机飞手")]),
    ("🚀","未来?",GOLD,"Future?",[("🛸","火星导游"),("🤖","机器人医生"),("🌌","太空建筑师"),("🧬","基因设计师")]),
]
for i,(em,head_cn,cl,head_en,items) in enumerate(cols):
    x=0.30+i*3.20
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.65),Inches(3.0),Inches(2.85))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(2.5)
    tb(s,x+0.10,1.75,2.80,0.45,f"{em} {head_cn}",sz=18,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.10,2.20,2.80,0.30,head_en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    for j,(je,jcn) in enumerate(items):
        y=2.55+j*0.42
        tb(s,x+0.20,y,0.5,0.35,je,sz=18,a=PP_ALIGN.CENTER)
        tb(s,x+0.75,y+0.03,2.10,0.32,jcn,sz=12,c=DARK)
# Predict-disappearing strip at bottom
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(4.65),Inches(9.40),Inches(0.95))
sh.fill.solid(); sh.fill.fore_color.rgb=GOLD; sh.line.fill.background()
tb(s,0.45,4.75,9.10,0.35,"🔮 你来预测: 今天的哪些工作, 以后可能没有了？",sz=15,b=True,c=WHITE)
tb(s,0.45,5.10,9.10,0.30,"Predict: which jobs today might disappear in the future?",sz=10,c=WARM)
tb(s,0.45,5.40,9.10,0.20,"💭 提示: 司机 (自动驾驶) · 收银员 (自助结账) · 邮递员 (无人机) · ...",sz=9,c=WHITE)
n+=1; pn(s,n)
notes(s,"3-4 分钟 — 收尾互动:\n• 关键句: 「以前没有这些工作, 现在有了 — 等你长大, 也会有现在没有的工作!」\n• 让 2-3 个孩子说: 「我猜未来会有 ___ 工作」\n• 反过来问: 「今天哪些工作以后可能消失?」 — 司机、收银员、邮递员、电梯操作员、流水线工人...\n• 不评判 — 想象力是这一节的目标\n• 收尾: 「重要的不是猜对 — 是 学会变, 学会喜欢, 学会帮人! 」")

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

# 36. 我会认 — 5 jobs, ONE SLIDE PER WORD (modeled on 世界旅行 word card pattern)
read_words=[
    ("🩺","医生","yī shēng","Doctor",DOC,
        "医生帮助生病的人。",
        "📷 医生 / 听诊器 / 白大褂"),
    ("📚","老师","lǎo shī","Teacher",TEACH,
        "老师在学校教我们读书。",
        "📷 老师 / 黑板 / 教室"),
    ("👮","警察","jǐng chá","Police",POLICE,
        "警察让大家很安全。",
        "📷 警察 / 警车 / 制服"),
    ("👨‍🍳","厨师","chú shī","Chef",CHEF,
        "厨师做的菜很好吃。",
        "📷 厨师 / 厨房 / 高白帽"),
    ("👷","工程师","gōng chéng shī","Engineer",ENG,
        "工程师设计了这座大桥。",
        "📷 工程师 / 安全帽 / 图纸"),
]
for em,cn,py,en,c,sent,img_label in read_words:
    s=ns(); bg(s,CREAM); hb(s,f"👀 我会认 · {cn}  I Can Read",c)
    # Left: big character card
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.0),Inches(4.5),Inches(2.5))
    sh.fill.solid(); sh.fill.fore_color.rgb=WARM; sh.line.fill.background()
    tb(s,0.5,1.10,4.3,1.4,cn,sz=72,b=True,c=c,a=PP_ALIGN.CENTER)
    tb(s,0.5,2.40,4.3,0.4,f"{py}  {en}",sz=20,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,0.5,2.85,4.3,0.4,"👉 跟我读！Read after me!",sz=14,c=c,a=PP_ALIGN.CENTER)
    # Right: image placeholder
    ib(s,5.3,1.0,4.4,2.5,img_label)
    # Bottom: example sentence
    sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(3.8),Inches(9.2),Inches(1.2))
    sh2.fill.solid(); sh2.fill.fore_color.rgb=WHITE; sh2.line.color.rgb=c; sh2.line.width=Pt(2)
    tb(s,0.6,3.9,1.5,0.4,"例句",sz=16,b=True,c=c)
    tb(s,0.6,4.3,8.8,0.5,sent,sz=22,b=True,c=DARK)
    n+=1; pn(s,n)
    notes(s,f"2 分钟 — {cn}:\n• 老师指字, 全班齐读 3 遍\n• 看图: 「这是 ___ , 在做什么?」\n• 读例句, 跟读\n• 抽 1-2 个学生用 {cn} 造一个新句子")

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
