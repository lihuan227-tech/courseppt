#!/usr/bin/env python3
"""
小小生活家 Little Homemaker — Day 2: 整理和收纳达人 Organizing & Storage Master
Same structure/design system as Day 1 (厨房小帮手), adapted for organizing & storage.
Palette: Fresh Organizing (teal + coral) — distinct from Day 1 breakfast green/orange.
The Day 1 "6 danger cards" become the "7 storage-tool cards" (guess -> reveal).
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

# --- Palette: Fresh Organizing (清爽整理) ---
TEAL = RGBColor(0x16,0x84,0x8A)   # primary: teal
CORAL = RGBColor(0xEF,0x6B,0x53)  # accent: coral
NAVY = RGBColor(0x1E,0x4B,0x54)   # deep teal-navy
SUNNY = RGBColor(0xF4,0xC1,0x3D)  # yellow highlight
CREAM = RGBColor(0xF8,0xF4,0xEC)  # background cream
WARM = RGBColor(0xFD,0xEF,0xE6)   # warm light box (coral-tinted)
WHITE = RGBColor(0xFF,0xFF,0xFF)
DARK = RGBColor(0x2C,0x2C,0x2C)
GRAY = RGBColor(0x88,0x88,0x88)
LGRAY = RGBColor(0xBB,0xBB,0xBB)
IMGBG = RGBColor(0xEC,0xEC,0xE6)
SKY = RGBColor(0x19,0x76,0xD2)
GREEN_OK = RGBColor(0x2E,0x9E,0x7A)  # correct / 合理
RED = RGBColor(0xD8,0x45,0x3A)       # wrong / 不合理
GOLD = RGBColor(0xD1,0x8F,0x0A)      # readable amber for text on white

# Per-tool colors (7)
C_VAC = TEAL
C_LAZY = CORAL
C_PEG = RGBColor(0x3E,0x8E,0xC4)   # blue
C_DOOR = RGBColor(0x8A,0x6F,0xB0)  # purple
C_DIV = RGBColor(0x2E,0x9E,0x7A)   # green-teal
C_MAT = RGBColor(0xE8,0x9A,0x33)   # amber
C_SINK = RGBColor(0x4C,0x86,0x8C)  # slate-teal

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
def hb(s,txt,c=TEAL,t=0.15):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(t),Inches(9.4),Inches(0.55));sh.fill.solid();sh.fill.fore_color.rgb=c;sh.line.fill.background()
    tb(s,0.4,t+0.03,9.2,0.5,txt,sz=20,b=True,c=WHITE)
def pn(s,n):
    chip=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(9.18),Inches(5.28),Inches(0.5),Inches(0.30))
    chip.fill.solid();chip.fill.fore_color.rgb=WHITE;chip.line.color.rgb=LGRAY;chip.line.width=Pt(0.75)
    bx=s.shapes.add_textbox(Inches(9.18),Inches(5.30),Inches(0.5),Inches(0.26));tf=bx.text_frame;tf.word_wrap=False
    tf.margin_left=0;tf.margin_right=0;tf.margin_top=0;tf.margin_bottom=0
    p=tf.paragraphs[0];p.alignment=PP_ALIGN.CENTER
    r=p.add_run();r.text=str(n);r.font.size=Pt(10);r.font.color.rgb=GRAY;r.font.name='KaiTi'
def notes(s,txt): s.notes_slide.notes_text_frame.text=txt
def div(title,sub,color,emoji=""):
    s=ns();bg(s,color)
    tb(s,0.5,1.5,9,1.2,f"{emoji} {title}",sz=42,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    lines=sub.split("\n")
    tf=tb(s,0.4,2.75,9.2,1.6,lines[0],sz=22,c=WHITE,a=PP_ALIGN.CENTER)
    for ln in lines[1:]:ap(tf,ln,sz=20,c=WHITE,a=PP_ALIGN.CENTER)
    return s

n=0

# ============================================================
# 1 COVER — Organizer badge
# ============================================================
s=ns();n+=1;bg(s,CREAM)
tb(s,1,0.25,8,0.7,"Organizing & Storage Master",sz=30,b=True,c=TEAL,a=PP_ALIGN.CENTER)
tb(s,1,0.82,8,0.45,"整理和收纳达人",sz=20,c=TEAL,a=PP_ALIGN.CENTER)
sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.25),Inches(1.5),Inches(3.5),Inches(3.5))
sh.fill.solid();sh.fill.fore_color.rgb=TEAL;sh.line.color.rgb=CORAL;sh.line.width=Pt(6)
sh2=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.55),Inches(1.8),Inches(2.9),Inches(2.9))
sh2.fill.solid();sh2.fill.fore_color.rgb=WHITE;sh2.line.color.rgb=CORAL;sh2.line.width=Pt(2)
tf=tb(s,3.6,2.05,2.8,0.4,"DAY 2",sz=16,b=True,c=CORAL,a=PP_ALIGN.CENTER)
ap(tf,"📦",sz=48,a=PP_ALIGN.CENTER)
ap(tf,"整理和收纳",sz=18,b=True,c=TEAL,a=PP_ALIGN.CENTER)
ap(tf,"ORGANIZING & STORAGE",sz=9,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,1,5.05,8,0.4,"🧹 一起把东西送回家！Let's give everything a home!",sz=14,b=True,c=CORAL,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 2 SCHEDULE
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"⏰ 今日时间安排  Today's Schedule")
for i,(nm,tm,dc,cl) in enumerate([
    ("Session 1  上午","11:00-11:45","找橡皮 + 四步整理法 + 合理判断 + 收纳工具",TEAL),
    ("Session 2  下午","2:00-2:45","复习 + 语言目标 (认字 + 写字)",CORAL),
    ("Session 3  下午","3:00-4:30","小组收纳挑战 + 方案测试 + 给建议",NAVY)]):
    y=0.9+i*1.5
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.5),Inches(y),Inches(9),Inches(1.2))
    sh.fill.solid();sh.fill.fore_color.rgb=cl;sh.line.fill.background()
    tb(s,0.7,y+0.15,4,0.4,nm,sz=20,b=True,c=WHITE)
    tb(s,0.7,y+0.6,3,0.4,tm,sz=15,c=WARM)
    tb(s,4.6,y+0.35,5.0,0.6,dc,sz=14,c=WHITE)
pn(s,n)

# ============================================================
# 3 OBJECTIVES
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🎯 教学目标  Learning Objectives")
tb(s,0.5,0.85,9,0.5,"📦 内容目标  Content:",sz=19,b=True,c=TEAL)
tf=tb(s,0.7,1.32,9,1.4,"1. 明白整理的意义：更快找到东西、减少丢失、节省时间",sz=14,c=DARK)
ap(tf,"2. 学会「四步整理法」：拿出来 → 分类 → 做决定 → 放回家",sz=14,c=DARK)
ap(tf,"3. 判断收纳是否合理：分类、固定位置、方便使用",sz=14,c=DARK)
ap(tf,"4. 认识收纳工具，并为真实问题选择合适的工具",sz=14,c=DARK)
tb(s,0.5,3.05,9,0.5,"🗣️ 语言目标  Language:",sz=19,b=True,c=CORAL)
tb(s,0.7,3.5,5.0,0.9,"👀 我会认：混乱 分类 放回\n　　　　　收拾 书包 (复习 整理)",sz=14,b=True,c=DARK)
tb(s,5.7,3.5,4.0,0.9,"✍️ 我会写：分类 书包 收拾",sz=14,b=True,c=DARK)
tb(s,0.5,4.6,9,0.5,"🎨 实践目标：小组收纳设计挑战 + 方案测试 + 给建议",sz=14,c=NAVY)
pn(s,n)

# ============================================================
# 4 SESSION 1 DIVIDER
# ============================================================
div("Session 1  上午","为什么整理 + 四步整理法 + 合理判断 + 收纳好帮手\n🔎 找一找  🧺 四步法  ⚖️ 合理吗  🧰 工具",TEAL,"🧹")
n+=1

# ============================================================
# 5 MISSION intro
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🧑‍🔧 你是整理达人!  You're an Organizing Master!",TEAL)
hbx=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.0),Inches(9.2),Inches(0.95))
hbx.fill.solid();hbx.fill.fore_color.rgb=WARM;hbx.line.color.rgb=CORAL;hbx.line.width=Pt(2.5)
tb(s,0.6,1.10,8.8,0.45,"🗂️ 今天你要学会整理和收纳，让东西又快又好地找到!",sz=22,b=True,c=TEAL,a=PP_ALIGN.CENTER)
tb(s,0.6,1.55,8.8,0.35,"Today you learn to organize so everything is easy to find!",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
missions=[("🔎","为什么整理","Why organize",TEAL),
          ("🧺","四步整理法","4-step method",CORAL),
          ("⚖️","合理吗?","Is it smart?",C_PEG),
          ("🧰","选对工具","Pick a tool",NAVY)]
for i,(em,cn,en,cl) in enumerate(missions):
    x=0.55+i*2.30;y=2.25
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(2.10),Inches(1.7))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=cl;sh.line.width=Pt(2.5)
    tb(s,x+0.1,y+0.2,1.9,0.7,em,sz=34,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,y+0.95,1.9,0.4,cn,sz=15,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,y+1.32,1.9,0.3,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
sf=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(4.25),Inches(9.4),Inches(0.95))
sf.fill.solid();sf.fill.fore_color.rgb=WARM;sf.line.color.rgb=CORAL;sf.line.width=Pt(2)
tb(s,0.5,4.33,1.7,0.4,"💬 我来说",sz=14,b=True,c=CORAL)
tb(s,2.0,4.30,7.6,0.3,"我是整理达人，我要把 ____ 送回家。",sz=15,b=True,c=DARK)
tb(s,2.0,4.62,7.6,0.3,"I'm an organizing master. I'll give ___ a home.",sz=11,c=GRAY)
tb(s,2.0,4.88,7.6,0.3,"举起手，我们开始! Raise your hand, let's start! 🙌",sz=12,b=True,c=TEAL)
pn(s,n)

# ============================================================
# 6 找橡皮挑战 — hook (setup + rules)
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🔎 找橡皮挑战  The Eraser Hunt",CORAL)
tb(s,0.4,0.85,9.2,0.32,"两名同学同时找橡皮，其他人帮忙计时 — 哪个箱子更快?",sz=13,c=GRAY,a=PP_ALIGN.CENTER)
boxes=[("📦","混乱箱  Messy Box",RED,["书、纸、玩具、铅笔、","衣服全混在一起","橡皮藏在里面"]),
       ("🗂️","整理箱  Tidy Box",GREEN_OK,["物品已经分类","橡皮放在「文具区」","或笔袋里"])]
for i,(em,title,cl,lines) in enumerate(boxes):
    x=0.3+i*4.85
    card=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.3),Inches(4.6),Inches(2.5))
    card.fill.solid();card.fill.fore_color.rgb=WHITE;card.line.color.rgb=cl;card.line.width=Pt(3)
    hd=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.3),Inches(4.6),Inches(0.55))
    hd.fill.solid();hd.fill.fore_color.rgb=cl;hd.line.fill.background()
    tb(s,x+0.15,1.37,4.3,0.42,f"{em} {title}",sz=17,b=True,c=WHITE)
    tf=tb(s,x+0.3,2.0,4.1,0.4,f"· {lines[0]}",sz=15,c=DARK)
    for l in lines[1:]:ap(tf,f"· {l}",sz=15,c=DARK)
rb=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(4.0),Inches(9.4),Inches(1.2))
rb.fill.solid();rb.fill.fore_color.rgb=WARM;rb.line.color.rgb=CORAL;rb.line.width=Pt(2)
tb(s,0.5,4.08,9.0,0.35,"🎮 玩法  How to Play:",sz=14,b=True,c=CORAL)
tb(s,0.5,4.45,9.2,0.35,"第一轮：在混乱箱里找橡皮  ·  第二轮：在整理箱里找橡皮  ·  全班计时!",sz=13,b=True,c=DARK)
tb(s,0.5,4.80,9.2,0.35,"⏱️ 把两次时间写在白板上，让结果更直观。",sz=12,c=GRAY)
notes(s,"准备两个箱子/书包。两名学生分别在混乱箱、整理箱里找橡皮，其他人计时。把两次时间写白板上。")
pn(s,n)

# ============================================================
# 7 找橡皮 — discussion + 总结
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"💡 为什么要整理?  Why Organize?",CORAL)
qs=[("1️⃣","哪一次找得更快?为什么第二次更容易?","Which was faster? Why?",TEAL),
    ("2️⃣","东西太乱会带来什么问题?整理只是为了漂亮吗?","What problems does mess cause?",RED)]
for i,(num,q_cn,q_en,cl) in enumerate(qs):
    y=1.0+i*1.1
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(y),Inches(9.2),Inches(0.95))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=cl;sh.line.width=Pt(2.5)
    nb=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(0.6),Inches(y+0.16),Inches(0.62),Inches(0.62))
    nb.fill.solid();nb.fill.fore_color.rgb=cl;nb.line.fill.background()
    tb(s,0.6,y+0.20,0.62,0.52,num,sz=20,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,1.45,y+0.12,8.0,0.5,q_cn,sz=17,b=True,c=DARK)
    tb(s,1.45,y+0.58,8.0,0.35,q_en,sz=11,c=GRAY)
sf=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(3.35),Inches(9.4),Inches(1.85))
sf.fill.solid();sf.fill.fore_color.rgb=WARM;sf.line.color.rgb=TEAL;sf.line.width=Pt(2.5)
head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(3.35),Inches(9.4),Inches(0.5))
head.fill.solid();head.fill.fore_color.rgb=TEAL;head.line.fill.background()
tb(s,0.5,3.42,9.0,0.4,"✨ 小结  整理的好处  Why It Helps",sz=15,b=True,c=WHITE)
benefits=["🔎 更快找到东西","🎒 减少丢失","⏰ 节省时间","😊 心情更好"]
for i,bft in enumerate(benefits):
    col=i%2;row=i//2
    x=0.65+col*4.7;y=3.95+row*0.47
    tb(s,x,y,4.4,0.4,bft,sz=15,b=True,c=DARK)
tb(s,0.5,4.90,9.2,0.28,"📌 整理不只是让地方漂亮 — 更是帮我们更快、更省心地生活。",sz=12,b=True,c=TEAL,a=PP_ALIGN.CENTER)
notes(s,"总结：整理不只是好看，还能更快找到东西、减少丢失、节省时间。可对比白板上两次寻找时间。")
pn(s,n)

# ============================================================
# 8 老师的混乱书包 — observe & diagnose
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🎒 老师的混乱书包  The Messy Backpack",CORAL)
tb(s,0.4,0.85,9.2,0.32,"老师边找边表演 — 学生观察，帮老师找出问题!",sz=13,c=GRAY,a=PP_ALIGN.CENTER)
ib(s,0.3,1.3,4.4,3.4,"📷 混乱的书包 / A messy backpack")
panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(1.3),Inches(4.8),Inches(3.4))
panel.fill.solid();panel.fill.fore_color.rgb=WHITE;panel.line.color.rgb=CORAL;panel.line.width=Pt(2.5)
head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(1.3),Inches(4.8),Inches(0.5))
head.fill.solid();head.fill.fore_color.rgb=CORAL;head.line.fill.background()
tb(s,5.05,1.37,4.6,0.4,"🕵️ 帮老师想一想  Help the Teacher",sz=14,b=True,c=WHITE)
tf=tb(s,5.05,1.95,4.6,0.4,"😩 「我的作业在哪里?」",sz=13,c=DARK)
for q in ["💧 「水杯怎么漏水了?」","✏️ 「我的铅笔是不是丢了?」","🧥 「外套为什么压在书下面?」","","🤔 书包有什么问题?先做什么?"]:
    ap(tf,"",sz=4);ap(tf,q,sz=13,c=DARK if not q.startswith("🤔") else TEAL,b=q.startswith("🤔"))
sf=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(4.85),Inches(9.4),Inches(0.5))
sf.fill.solid();sf.fill.fore_color.rgb=WARM;sf.line.color.rgb=TEAL;sf.line.width=Pt(2)
tb(s,0.5,4.93,9.0,0.4,"💬 哪些东西不应该放在一起?怎样下次很快找到作业?",sz=13,b=True,c=DARK,a=PP_ALIGN.CENTER)
notes(s,"老师拿混乱书包边找边演。先不给答案，让学生观察诊断问题，把建议写白板 → 引出四步整理法。")
pn(s,n)

# ============================================================
# 9 四步整理法 — 4 steps + gestures
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🧺 四步整理法  The 4-Step Method",TEAL)
tb(s,0.4,0.82,9.2,0.32,"记住这四步，配上动作，谁都能整理好!",sz=13,c=GRAY,a=PP_ALIGN.CENTER)
steps=[("1️⃣","🙌","全部拿出来","Take everything out","动作：双手向外张开",TEAL),
       ("2️⃣","👐","分类","Sort similar things","动作：双手分到两边",CORAL),
       ("3️⃣","🤔","做决定","Keep / put back / recycle / donate","动作：手指放下巴思考",C_PEG),
       ("4️⃣","🏠","放回固定位置","Give everything a home","动作：双手做屋顶形状",GREEN_OK)]
for i,(num,em,cn,en,act,cl) in enumerate(steps):
    x=0.3+i*2.4
    card=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.25),Inches(2.25),Inches(3.35))
    card.fill.solid();card.fill.fore_color.rgb=WHITE;card.line.color.rgb=cl;card.line.width=Pt(2.5)
    nb=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x+0.85),Inches(1.4),Inches(0.55),Inches(0.55))
    nb.fill.solid();nb.fill.fore_color.rgb=cl;nb.line.fill.background()
    tb(s,x+0.85,1.44,0.55,0.48,num,sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,2.05,2.05,0.8,em,sz=40,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.95,2.15,0.45,cn,sz=16,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.08,3.4,2.1,0.5,en,sz=9,c=GRAY,a=PP_ALIGN.CENTER)
    ab=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x+0.12),Inches(3.95),Inches(2.0),Inches(0.5))
    ab.fill.solid();ab.fill.fore_color.rgb=WARM;ab.line.fill.background()
    tb(s,x+0.18,4.02,1.9,0.4,act,sz=10,b=True,c=cl,a=PP_ALIGN.CENTER)
tb(s,0.4,4.75,9.2,0.35,"👉 做决定 = 留下 / 放回 / 回收 / 捐出  Keep · Put back · Recycle · Donate",sz=12,b=True,c=TEAL,a=PP_ALIGN.CENTER)
notes(s,"四步整理法配动作，帮低龄学生记忆：拿出来(张开)→分类(分两边)→做决定(摸下巴)→放回家(屋顶)。")
pn(s,n)

# ============================================================
# 10 三个核心概念
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🌟 整理的三个核心概念  3 Big Ideas",TEAL)
tb(s,0.4,0.9,9.2,0.35,"这节课最重要的三件事 — 记住它们，你就是整理达人!",sz=13,c=GRAY,a=PP_ALIGN.CENTER)
ideas=[("🗂️","分类","Sort","相同用途的东西\n放在一起",TEAL),
       ("🏠","固定位置","A home","每件东西都有\n自己的「家」",CORAL),
       ("✅","方便使用","Easy to use","不只整齐，还要\n好找、好拿、好放回",GREEN_OK)]
for i,(em,cn,en,d,cl) in enumerate(ideas):
    x=0.3+i*3.2
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.4),Inches(3.0),Inches(3.3))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=cl;sh.line.width=Pt(3)
    tb(s,x+0.1,1.6,2.8,0.8,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,2.55,2.8,0.45,cn,sz=20,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,3.0,2.8,0.3,en,sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    ls=d.split('\n')
    tf=tb(s,x+0.15,3.5,2.7,0.5,ls[0],sz=13,c=DARK,a=PP_ALIGN.CENTER)
    for l in ls[1:]:ap(tf,l,sz=13,c=DARK,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 11-12 合理还是不合理 — two-phase judgment game
# ============================================================
JUDGE=[
    ("所有东西塞进一个大抽屉，桌面看起来很干净","只是把东西藏起来 — 没分类、没固定位置","❌"),
    ("书按大小排得很整齐，但每天用的练习本在最下面","好看却不方便 — 常用的要放在好拿的地方","❌"),
    ("水杯和作业本放在同一个袋子里","水杯可能弄湿作业 — 干和湿要分开放","❌"),
    ("盒子很多，但每个盒子都没有标签","找起来还是慢 — 贴上标签才一目了然","❌"),
]
def judge_slide(reveal):
    global n
    s=ns();bg(s,CREAM)
    if reveal:
        hb(s,"⚖️ 合理还是不合理?  💡 答案揭晓 Reveal!",CORAL)
    else:
        hb(s,"⚖️ 合理还是不合理?  🤔 你觉得呢? Think!",TEAL)
    rb=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(0.9),Inches(9.4),Inches(0.5))
    rb.fill.solid();rb.fill.fore_color.rgb=WARM;rb.line.color.rgb=CORAL;rb.line.width=Pt(1.5)
    if reveal:
        tb(s,0.45,0.96,9.0,0.38,"💡 这四种都「不合理」 — 看看应该怎么改!",sz=13,b=True,c=RED,a=PP_ALIGN.CENTER)
    else:
        tb(s,0.45,0.96,9.0,0.38,"👍 合理 = 举手V   👎 不合理 = 双手交叉X   想一想:怎么改更好?",sz=13,b=True,c=DARK,a=PP_ALIGN.CENTER)
    for i,(sit,fix,mark) in enumerate(JUDGE):
        y=1.55+i*0.90
        card=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(y),Inches(9.4),Inches(0.82))
        card.fill.solid();card.fill.fore_color.rgb=WHITE
        card.line.color.rgb=RED if reveal else LGRAY;card.line.width=Pt(1.5)
        nb=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(0.42),Inches(y+0.10),Inches(0.42),Inches(0.42))
        nb.fill.solid();nb.fill.fore_color.rgb=CORAL if reveal else TEAL;nb.line.fill.background()
        tb(s,0.42,y+0.13,0.42,0.36,f"{i+1}",sz=13,b=True,c=WHITE,a=PP_ALIGN.CENTER)
        tb(s,0.98,y+0.09,7.7,0.4,sit,sz=12,b=True,c=DARK)
        if reveal:
            tb(s,0.98,y+0.45,7.7,0.32,f"➡️ {fix}",sz=11,c=GREEN_OK)
            tb(s,8.9,y+0.18,0.6,0.5,"❌",sz=22,a=PP_ALIGN.CENTER)
        else:
            tb(s,8.9,y+0.18,0.6,0.5,"?",sz=22,b=True,c=GRAY,a=PP_ALIGN.CENTER)
    bs2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(5.18),Inches(9.4),Inches(0.34))
    if reveal:
        bs2.fill.solid();bs2.fill.fore_color.rgb=GREEN_OK;bs2.line.fill.background()
        tb(s,0.45,5.20,9.0,0.28,"✨ 合理的整理 = 分类 + 固定位置 + 方便使用",sz=11,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    else:
        bs2.fill.solid();bs2.fill.fore_color.rgb=TEAL;bs2.line.fill.background()
        tb(s,0.45,5.20,9.0,0.28,"👉 先判断，再想「怎么改更好」 — 3...2...1!",sz=11,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    notes(s,"两阶段:Question页学生判断合理/不合理并想怎么改;Reveal页揭晓——四种都不合理，逐一说改法。")
    n+=1;pn(s,n);return s
judge_slide(False)
judge_slide(True)

# ============================================================
# 13 收纳好帮手 overview (7 tools grid)
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🧰 收纳好帮手  Storage Helpers",TEAL)
tb(s,0.4,0.88,9,0.35,"这些聪明的工具能帮我们省空间、好整理 — 我们来认识 7 样!",sz=13,c=GRAY,a=PP_ALIGN.CENTER)
tools_ov=[
    ("🛍️","真空压缩袋",C_VAC),("🔄","旋转收纳盘",C_LAZY),("🧷","洞洞板",C_PEG),
    ("🚪","门后挂袋",C_DOOR),("📏","抽屉分隔板",C_DIV),("🧺","玩具收纳垫",C_MAT),
    ("🚰","水槽置物架",C_SINK),
]
for i,(em,cn,cl) in enumerate(tools_ov):
    col=i%4;row=i//4
    x=0.3+col*2.4;y=1.4+row*1.85
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(2.25),Inches(1.65))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=cl;sh.line.width=Pt(2.5)
    tb(s,x+0.1,y+0.12,2.05,0.65,em,sz=34,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,y+0.9,2.15,0.6,cn,sz=15,b=True,c=cl,a=PP_ALIGN.CENTER)
note=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(7.5),Inches(3.25),Inches(2.25),Inches(1.65))
note.fill.solid();note.fill.fore_color.rgb=WARM;note.line.color.rgb=CORAL;note.line.width=Pt(2.5)
tf=tb(s,7.6,3.42,2.05,0.4,"⭐ 课堂演示",sz=14,b=True,c=CORAL,a=PP_ALIGN.CENTER)
ap(tf,"最推荐现场展示:",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
ap(tf,"真空袋 · 旋转盘",sz=12,b=True,c=DARK,a=PP_ALIGN.CENTER)
ap(tf,"玩具垫 · 洞洞板",sz=12,b=True,c=DARK,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 14-20 tool cards (guess -> reveal), one per tool
# ============================================================
def tool_card(em,cn,en,color,show,guesses,reveal,concept,demo=False,fit=None):
    global n
    s=ns();bg(s,CREAM)
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(0.15),Inches(9.4),Inches(0.7))
    sh.fill.solid();sh.fill.fore_color.rgb=color;sh.line.fill.background()
    tb(s,0.5,0.22,1.0,0.55,em,sz=28,c=WHITE)
    tb(s,1.5,0.20,4.0,0.5,cn,sz=25,b=True,c=WHITE)
    tb(s,1.5,0.62,4.0,0.25,en,sz=11,c=WARM)
    if demo:
        pill=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(7.4),Inches(0.27),Inches(2.2),Inches(0.45))
        pill.fill.solid();pill.fill.fore_color.rgb=SUNNY;pill.line.color.rgb=WHITE;pill.line.width=Pt(1.5)
        tb(s,7.5,0.32,2.0,0.4,"⭐ 课堂演示 Demo",sz=12,b=True,c=DARK,a=PP_ALIGN.CENTER)
    # Left image
    ib(s,0.3,1.05,4.3,2.5,f"📷 {cn}")
    # Right panel — show + guess
    panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.85),Inches(1.05),Inches(4.85),Inches(2.5))
    panel.fill.solid();panel.fill.fore_color.rgb=WHITE;panel.line.color.rgb=color;panel.line.width=Pt(2.5)
    tb(s,5.05,1.15,4.5,0.3,"🔎 怎么演示  How to show:",sz=13,b=True,c=color)
    tb(s,5.05,1.46,4.55,0.55,show,sz=11,c=DARK)
    tb(s,5.05,2.12,4.5,0.3,"❓ 让学生猜  Ask kids:",sz=13,b=True,c=color)
    tf=tb(s,5.05,2.42,4.55,0.3,f"· {guesses[0]}",sz=11,c=DARK)
    for g in guesses[1:]:ap(tf,f"· {g}",sz=11,c=DARK)
    # Reveal strip
    rv=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(3.68),Inches(9.4),Inches(0.82))
    rv.fill.solid();rv.fill.fore_color.rgb=WARM;rv.line.color.rgb=color;rv.line.width=Pt(2)
    tb(s,0.5,3.74,1.2,0.35,"💡 揭晓",sz=13,b=True,c=color)
    tb(s,1.6,3.74,8.0,0.72,reveal,sz=12,b=True,c=DARK)
    # Key concept strip
    kc=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(4.6),Inches(9.4),Inches(0.55))
    kc.fill.solid();kc.fill.fore_color.rgb=color;kc.line.fill.background()
    txt=f"🔑 关键概念: {concept}"
    if fit:txt+=f"   ·   适合: {fit}"
    tb(s,0.5,4.67,9.2,0.4,txt,sz=12,b=True,c=WHITE)
    n+=1;pn(s,n);return s

tool_card("🛍️","真空压缩袋","Vacuum Storage Bag",C_VAC,
    "透明大袋子上有一个圆形接口，先放进一件蓬松的外套或小被子，再用吸尘器/气泵抽气，袋子迅速变扁。",
    ["圆形接口是做什么的?","为什么袋子要密封?","衣服会发生什么变化?"],
    "抽走空气，让厚衣服和被子占用更少的空间。",
    "东西没有减少，但占用的空间变小了。",demo=True)
tool_card("🔄","旋转收纳盘","Lazy Susan Turntable",C_LAZY,
    "圆形托盘，先不要转。放上调味料、小瓶子或画笔，再轻轻旋转。",
    ["它为什么是圆形的?","后面的东西拿不到怎么办?","它可以放在哪里?"],
    "一转，后面的东西转到前面，不用把其他东西全搬开。",
    "让难拿到的东西变得容易拿到。",demo=True,fit="厨房柜 · 冰箱 · 浴室 · 美术区")
tool_card("🧷","洞洞板","Pegboard",C_PEG,
    "一块上面有很多小洞的板，加上几个不同形状的挂钩，挂上小物品。",
    ["为什么板上有这么多洞?","挂钩可以移动吗?","什么东西适合挂在墙上?"],
    "挂钩可以插在不同位置，根据物品大小自由改变。",
    "收纳的位置可以根据需要改变。",demo=True,fit="剪刀 · 胶带 · 小篮子 · 工具 · 耳机")
tool_card("🚪","门后透明挂袋","Over-the-Door Organizer",C_DOOR,
    "先把挂袋折起来，让学生看到很多透明小口袋，再挂到门后。",
    ["为什么有这么多口袋?","为什么口袋是透明的?","铺桌上、放地上、还是挂起来?"],
    "挂在门后，利用垂直空间来收纳。",
    "地面没空间时，可以向上收纳。",fit="小玩具 · 袜子 · 发饰 · 美术材料 · 水瓶")
tool_card("📏","伸缩抽屉分隔板","Adjustable Drawer Dividers",C_DIV,
    "两根像短木板或塑料条的分隔板，放进一个大盒子里，把盒子分成几个区域。",
    ["它为什么可以变长、变短?","它是单独装东西的吗?","它可以怎样改变大抽屉?"],
    "把一个大空间分成几个小空间，物品各归各位。",
    "大空间不分类，也很容易变乱。",fit="袜子 · 内衣 · 文具 · 厨房工具")
tool_card("🧺","玩具收纳垫","Toy Storage Mat",C_MAT,
    "平铺在地上像圆形游戏垫，周围有一根绳子。放上积木，再拉起绳子，整张垫子变成收纳袋。",
    ["为什么垫子周围有绳子?","玩完玩具怎样快速整理?","拉动绳子会发生什么?"],
    "在垫子上玩，玩完拉起绳子，玩具就装进袋子里。",
    "好的收纳工具，能让整理变得更快更容易。",demo=True,fit="积木 · 拼图 · 小汽车 · 乐高")
tool_card("🚰","水槽伸缩置物架","Expandable Sink Organizer",C_SINK,
    "一个可以拉长、缩短的架子，架在水槽两边，底部有很多小孔。",
    ["它为什么可以改变长度?","它应该架在哪里?","为什么底部有很多小孔?"],
    "架在水槽两边放海绵、抹布、刷子；小孔让水流出去，不积水。",
    "收纳工具还要考虑物品是否潮湿。",fit="海绵 · 抹布 · 洗碗刷")

# ============================================================
# 21 哪个工具最适合 — matching game
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🧩 哪个工具最适合?  Which Tool Fits?",CORAL)
tb(s,0.4,0.85,9.2,0.32,"老师说一个家里的问题，学生举起对应工具的图片卡!",sz=13,c=GRAY,a=PP_ALIGN.CENTER)
matches=[("🧊 冬天的厚被子占太多空间","🛍️ 真空压缩袋"),
         ("🍶 冰箱最里面的小瓶子拿不到","🔄 旋转收纳盘"),
         ("🧱 积木玩完要一块块捡起来","🧺 玩具收纳垫"),
         ("🔌 书桌旁边有很多充电线","📦 电线收纳盒"),
         ("🧺 洗衣机旁边有很窄的缝隙","🛒 缝隙推车"),
         ("📺 遥控器总在沙发上找不到","🛋️ 沙发扶手收纳袋")]
for i,(prob,tool) in enumerate(matches):
    y=1.3+i*0.66
    pb=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(y),Inches(5.1),Inches(0.56))
    pb.fill.solid();pb.fill.fore_color.rgb=WHITE;pb.line.color.rgb=CORAL;pb.line.width=Pt(1.5)
    tb(s,0.55,y+0.10,4.9,0.4,prob,sz=13,b=True,c=DARK)
    ar=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.6),Inches(y),Inches(4.1),Inches(0.56))
    ar.fill.solid();ar.fill.fore_color.rgb=TEAL;ar.line.fill.background()
    tb(s,5.75,y+0.10,3.9,0.4,tool,sz=13,b=True,c=WHITE)
tb(s,0.4,5.28,9.2,0.28,"💡 不是所有工具都适合所有东西 — 要按问题选工具!",sz=11,b=True,c=CORAL,a=PP_ALIGN.CENTER)
notes(s,"升级玩法：老师念家庭问题，学生举对应工具图片卡。左=问题，右=最适合的工具(答案)。")
pn(s,n)

# ============================================================
# 22 教师总结 — choosing a tool
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🧠 怎样选收纳工具?  Choosing a Tool",TEAL)
tb(s,0.4,0.88,9.2,0.32,"选工具前，先问自己这几个问题:",sz=13,c=GRAY,a=PP_ALIGN.CENTER)
crit=["📐 物品有多大?","🔁 经常使用吗?","📍 应该放在哪里?",
      "💧 需要防水/通风/标签吗?","🙋 孩子能自己拿到并放回吗?","🪄 是否真的节省空间?"]
for i,ct in enumerate(crit):
    col=i%2;row=i//2
    x=0.4+col*4.75;y=1.35+row*0.72
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(4.55),Inches(0.6))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=TEAL;sh.line.width=Pt(1.5)
    tb(s,x+0.15,y+0.12,4.3,0.4,ct,sz=14,b=True,c=DARK)
gold=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(3.9),Inches(9.4),Inches(1.3))
gold.fill.solid();gold.fill.fore_color.rgb=WARM;gold.line.color.rgb=CORAL;gold.line.width=Pt(2.5)
tb(s,0.5,4.0,9.0,0.4,"✨ 记住这句话  Remember:",sz=14,b=True,c=CORAL)
tb(s,0.5,4.42,9.2,0.7,"最酷的收纳工具，不一定最贵 — 而是能巧妙利用空间、解决真实问题，\n还让人愿意把东西放回去。",sz=14,b=True,c=DARK)
pn(s,n)

# ============================================================
# 23 小组收纳设计挑战 — Project (materials + task)
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🛠️ 小组收纳设计挑战  Group Design Challenge",NAVY)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(0.92),Inches(4.4),Inches(0.4))
sh.fill.solid();sh.fill.fore_color.rgb=TEAL;sh.line.fill.background()
tb(s,0.4,0.95,4.2,0.35,"🧺 每组材料  Materials",sz=14,b=True,c=WHITE)
tf=tb(s,0.4,1.4,4.4,2.5,"📚 3–5 本书",sz=12,c=DARK)
for m in ["✏️ 铅笔、彩笔、橡皮、剪刀","📄 练习纸和废纸","🧸 小玩具","🧥 小外套或布袋 · 🥤 水杯",
          "📦 2–4 个空盒子/篮子/文件夹","🏷️ 标签纸和马克笔"]:
    ap(tf,m,sz=12,c=DARK)
sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(0.92),Inches(4.8),Inches(0.4))
sh2.fill.solid();sh2.fill.fore_color.rgb=CORAL;sh2.line.fill.background()
tb(s,5.0,0.95,4.6,0.35,"🎯 任务  The Task",sz=14,b=True,c=WHITE)
task=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(1.4),Inches(4.8),Inches(0.9))
task.fill.solid();task.fill.fore_color.rgb=WARM;task.line.color.rgb=CORAL;task.line.width=Pt(2)
tb(s,5.05,1.48,4.55,0.8,"把桌上的物品整理好，让别人能在 ⏱️ 10 秒内找到指定物品!",sz=13,b=True,c=DARK)
tf2=tb(s,5.0,2.45,4.7,1.5,"🤔 讨论:",sz=13,b=True,c=NAVY)
for q in ["· 哪些可以放一起?哪些必须分开?","· 每一类放在哪里?需要标签吗?","· 常用的放哪?不常用的放哪?"]:
    ap(tf2,q,sz=12,c=DARK)
sf=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(4.15),Inches(9.4),Inches(1.0))
sf.fill.solid();sf.fill.fore_color.rgb=WARM;sf.line.color.rgb=TEAL;sf.line.width=Pt(2.5)
tb(s,0.5,4.24,9.0,0.4,"⏱️ 成功标准  Success:",sz=14,b=True,c=TEAL)
tb(s,0.5,4.64,9.2,0.5,"整理好以后，请「别的组」来找指定物品 — 10 秒内找到就算成功! 🎉",sz=15,b=True,c=DARK)
notes(s,"给每组一个「混乱桌面」+ 收纳容器和标签。任务：整理到别人10秒内能找到指定物品。")
pn(s,n)

# ============================================================
# 24 小组挑战 — sentence frames
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"💬 收纳设计句型  Say These",TEAL)
tb(s,0.4,0.9,9.2,0.35,"一边整理，一边用这些句型说出你的想法:",sz=13,c=GRAY,a=PP_ALIGN.CENTER)
frames=["我们把 ______ 和 ______ 放在一起，因为 ______ 。",
        "______ 应该单独放，因为 ______ 。",
        "我们给这个盒子贴上「______」标签。",
        "经常使用的 ______ 应该放在 ______ 。",
        "这样整理以后，我们可以更快地 ______ 。"]
for i,fr in enumerate(frames):
    y=1.4+i*0.72
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(y),Inches(9.2),Inches(0.6))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=TEAL if i%2==0 else CORAL;sh.line.width=Pt(2)
    tb(s,0.6,y+0.12,0.5,0.4,"💬",sz=18)
    tb(s,1.15,y+0.13,8.4,0.4,fr,sz=15,b=True,c=DARK)
tb(s,0.4,5.28,9.2,0.28,"💡 把这张打印出来，贴在每组桌上当参考。",sz=11,b=True,c=NAVY,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 25 收纳方案测试
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🔍 收纳方案测试  Test the Solution",CORAL)
tb(s,0.4,0.85,9.2,0.32,"每组完成后 — 让「别的组」来找东西，才算真的整理好!",sz=13,c=GRAY,a=PP_ALIGN.CENTER)
lp=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(1.3),Inches(4.6),Inches(3.5))
lp.fill.solid();lp.fill.fore_color.rgb=WHITE;lp.line.color.rgb=TEAL;lp.line.width=Pt(2.5)
lh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(1.3),Inches(4.6),Inches(0.5))
lh.fill.solid();lh.fill.fore_color.rgb=TEAL;lh.line.fill.background()
tb(s,0.45,1.37,4.4,0.4,"⏱️ 老师随机任务  Find it in 10s!",sz=14,b=True,c=WHITE)
tf=tb(s,0.5,1.95,4.3,2.7,"🔵 找到一本蓝色的书",sz=14,c=DARK)
for t in ["🩹 找到一块橡皮","✂️ 找到剪刀","📄 找到一张空白纸","🥤 找到水杯"]:
    ap(tf,"",sz=6);ap(tf,t,sz=14,c=DARK)
rp=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.1),Inches(1.3),Inches(4.6),Inches(3.5))
rp.fill.solid();rp.fill.fore_color.rgb=WHITE;rp.line.color.rgb=CORAL;rp.line.width=Pt(2.5)
rh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.1),Inches(1.3),Inches(4.6),Inches(0.5))
rh.fill.solid();rh.fill.fore_color.rgb=CORAL;rh.line.fill.background()
tb(s,5.25,1.37,4.4,0.4,"🗣️ 参观组反馈  Feedback",sz=14,b=True,c=WHITE)
tf2=tb(s,5.3,1.95,4.3,2.7,"· 什么地方很清楚?",sz=14,c=DARK)
for t in ["· 什么东西不容易找到?","· 标签有没有帮助?","· 哪些物品还可以换位置?"]:
    ap(tf2,"",sz=6);ap(tf2,t,sz=14,c=DARK)
tb(s,0.4,4.9,9.2,0.4,"📌 整理不是自己觉得整齐就行 — 别人也要看懂、能找到!",sz=13,b=True,c=CORAL,a=PP_ALIGN.CENTER)
notes(s,"不让本组自测，让别组来找东西并计时。参观组给反馈。强调:整理要让别人也能找到。")
pn(s,n)

# ============================================================
# 26 整理规则 + 全班口令
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"📜 整理不是一次完成  Make It a Habit",NAVY)
tb(s,0.4,0.88,9.2,0.32,"今天整理好，三天后会不会又变乱?我们一起定规则!",sz=13,c=GRAY,a=PP_ALIGN.CENTER)
rules=["✅ 用完马上放回原位","✅ 每个盒子只放一种物品","✅ 找不到「家」的先放待整理箱",
       "✅ 每天下课前用 2 分钟检查","✅ 标签朝外，大家都看得见","✅ 新东西进来，先决定放哪"]
for i,r in enumerate(rules):
    col=i%2;row=i//2
    x=0.4+col*4.75;y=1.32+row*0.62
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(4.55),Inches(0.52))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=TEAL;sh.line.width=Pt(1.5)
    tb(s,x+0.15,y+0.09,4.3,0.35,r,sz=13,b=True,c=DARK)
chant=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(3.5),Inches(9.4),Inches(1.05))
chant.fill.solid();chant.fill.fore_color.rgb=CORAL;chant.line.fill.background()
tb(s,0.5,3.58,9.0,0.4,"📣 全班口令  Class Chant",sz=15,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,0.5,4.02,9.0,0.4,"老师:「用完东西——」   全班:「送它回家!」",sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,0.4,4.68,9.2,0.4,"🔑 三个核心:  分类  ·  固定位置  ·  方便使用",sz=14,b=True,c=NAVY,a=PP_ALIGN.CENTER)
notes(s,"让学生一起定简单规则，设一个全班口令。强调三个核心概念:分类、固定位置、方便使用。")
pn(s,n)

# ============================================================
# 27 SESSION 2 DIVIDER
# ============================================================
div("Session 2  下午","复习 + 语言目标 (认字 + 写字)\n我会认 5 词 · 我会写 3 词",CORAL,"📖")
n+=1

# ============================================================
# 28 QUICK REVIEW — 四步整理法 order
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🔄 快速复习  Quick Review",CORAL)
tb(s,0.4,0.85,9,0.3,"四步整理法，你还记得顺序吗?(口头排一排)  Put the 4 steps in order!",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
review=[("🙌","全部拿出来",TEAL),("👐","分类",CORAL),("🤔","做决定",C_PEG),("🏠","放回家",GREEN_OK)]
for i,(em,cn,cl) in enumerate(review):
    x=0.3+i*2.4
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.5),Inches(2.15),Inches(2.4))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=cl;sh.line.width=Pt(2.5)
    tb(s,x+0.1,1.7,1.95,0.8,em,sz=40,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.7,2.05,0.5,cn,sz=15,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.25,2.05,0.4,"?",sz=26,b=True,c=GRAY,a=PP_ALIGN.CENTER)
    if i<3:tb(s,x+2.02,1.95,0.5,0.5,"➡️",sz=22,c=CORAL)
tb(s,0.4,4.3,9.2,0.4,"💬 先 ____ ，再 ____ ，然后 ____ ，最后 ____ 。",sz=15,b=True,c=DARK,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 29-33  我会认 word cards
# ============================================================
def word_card_read(w,py,en,sent,img):
    global n
    s=ns();bg(s,CREAM);hb(s,"👀 我会认  I Can Read",CORAL)
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.0),Inches(4.5),Inches(2.5))
    sh.fill.solid();sh.fill.fore_color.rgb=WARM;sh.line.fill.background()
    tb(s,0.5,1.1,4.3,1.4,w,sz=72,b=True,c=TEAL,a=PP_ALIGN.CENTER)
    tb(s,0.5,2.4,4.3,0.4,f"{py}  {en}",sz=20,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,0.5,2.85,4.3,0.4,"👉 跟我读！Read after me!",sz=14,c=CORAL,a=PP_ALIGN.CENTER)
    ib(s,5.3,1.0,4.4,2.5,img)
    sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(3.8),Inches(9.2),Inches(1.2))
    sh2.fill.solid();sh2.fill.fore_color.rgb=WHITE;sh2.line.color.rgb=CORAL;sh2.line.width=Pt(2)
    tb(s,0.6,3.9,1.5,0.4,"例句",sz=16,b=True,c=CORAL)
    tb(s,0.6,4.3,8.8,0.5,sent,sz=22,b=True,c=DARK)
    n+=1;pn(s,n);return s
read_words=[
    ("混乱","hùn luàn","messy","书包太混乱，找不到作业。","📷 混乱"),
    ("分类","fēn lèi","sort","我把东西分类放好。","📷 分类"),
    ("放回","fàng huí","put back","用完要放回原处。","📷 放回"),
    ("收拾","shōu shí","tidy up","玩完玩具要收拾。","📷 收拾"),
    ("书包","shū bāo","backpack","我的书包很整齐。","📷 书包"),
]
for w,py,en,sent,img in read_words:
    word_card_read(w,py,en,sent,img)

# ============================================================
# 34 WORD GAMES
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🎮 练一练  Word Games (选一个玩！)",CORAL)
games=[
    ("1️⃣","拍苍蝇\nFly Swatter","把字卡贴在\n白板上，老师\n说词语，学生拍！",WARM),
    ("2️⃣","举牌游戏\nShow Me","每人 5 张字卡\n老师说词语\n举正确的卡",RGBColor(0xFD,0xEF,0xE6)),
    ("3️⃣","抢椅子\nMusical Chairs","椅子上放字卡\n音乐停，读出词",RGBColor(0xE2,0xF2,0xF1)),
    ("4️⃣","传话筒\nPass the Mic","传球，停下的人\n读字卡并造句",RGBColor(0xE3,0xF2,0xFD)),
]
for i,(num,nm,desc,bgc) in enumerate(games):
    x=0.3+i*2.4
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(0.9),Inches(2.2),Inches(4.2))
    sh.fill.solid();sh.fill.fore_color.rgb=bgc;sh.line.fill.background()
    tb(s,x+0.1,1.0,2.0,0.4,num,sz=24,a=PP_ALIGN.CENTER)
    ls=nm.split('\n')
    tf=tb(s,x+0.1,1.45,2.0,0.85,ls[0],sz=16,b=True,c=DARK,a=PP_ALIGN.CENTER)
    for l in ls[1:]:ap(tf,l,sz=16,b=True,c=DARK,a=PP_ALIGN.CENTER)
    ls2=desc.split('\n')
    tf2=tb(s,x+0.15,2.5,1.9,1.5,ls2[0],sz=12,c=DARK,a=PP_ALIGN.CENTER)
    for l in ls2[1:]:ap(tf2,l,sz=12,c=DARK,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,4.75,2.0,0.3,"低 prep ✅",sz=11,b=True,c=GREEN_OK,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 35-37  我会写 cards
# ============================================================
def word_card_write(w,py,en,img):
    global n
    s=ns();bg(s,CREAM);hb(s,"✍️ 我会写  I Can Write",TEAL)
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.0),Inches(4.5),Inches(2.0))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=TEAL;sh.line.width=Pt(3)
    tb(s,0.5,1.05,4.3,1.2,w,sz=72,b=True,c=TEAL,a=PP_ALIGN.CENTER)
    tb(s,0.5,2.2,4.3,0.4,f"{py}  {en}",sz=20,c=GRAY,a=PP_ALIGN.CENTER)
    ib(s,5.3,1.0,4.4,2.0,img)
    sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(3.3),Inches(5.0),Inches(1.8))
    sh2.fill.solid();sh2.fill.fore_color.rgb=WARM;sh2.line.fill.background()
    tb(s,0.6,3.4,4.6,0.4,"📝 笔顺 Stroke Order",sz=16,b=True,c=TEAL)
    ib(s,0.6,3.9,4.6,1.0,"📷 插入笔顺图片")
    tf=tb(s,5.8,3.4,3.8,0.4,"练习步骤 Practice:",sz=14,b=True,c=TEAL)
    ap(tf,"1. 空中写 Air Write",sz=13,c=DARK)
    ap(tf,"2. 手心写 Palm Write",sz=13,c=DARK)
    ap(tf,"3. 纸上写 3 times",sz=13,c=DARK)
    n+=1;pn(s,n);return s
write_words=[
    ("分类","fēn lèi","sort","📷 分类"),
    ("书包","shū bāo","backpack","📷 书包"),
    ("收拾","shōu shí","tidy up","📷 收拾"),
]
for w,py,en,img in write_words:
    word_card_write(w,py,en,img)

# ============================================================
# 38 SENTENCE FRAMES — 我会说
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"💬 我会说  Sentence Frames (K · G1–3)",CORAL)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(0.95),Inches(4.55),Inches(4.0))
sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=CORAL;sh.line.width=Pt(2.5)
pb=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.45),Inches(1.05),Inches(1.6),Inches(0.4))
pb.fill.solid();pb.fill.fore_color.rgb=CORAL;pb.line.fill.background()
tb(s,0.55,1.10,1.5,0.35,"K (TK-K)",sz=12,b=True,c=WHITE,a=PP_ALIGN.CENTER)
k_frames=[("我把 ____ 放回 ____ 。","I put ___ back in ___."),
          ("这是 ____ 的家。","This is ___'s home."),
          ("我会分类。","I can sort."),
          ("我会收拾。","I can tidy up."),
          ("找到了！","Found it!")]
for i,(cn,en) in enumerate(k_frames):
    y=1.55+i*0.62
    tb(s,0.5,y,4.3,0.4,f"·  {cn}",sz=17,b=True,c=TEAL)
    tb(s,0.5,y+0.32,4.3,0.25,en,sz=9,c=GRAY)
sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.05),Inches(0.95),Inches(4.65),Inches(4.0))
sh2.fill.solid();sh2.fill.fore_color.rgb=WHITE;sh2.line.color.rgb=TEAL;sh2.line.width=Pt(2.5)
pb2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.20),Inches(1.05),Inches(1.7),Inches(0.4))
pb2.fill.solid();pb2.fill.fore_color.rgb=TEAL;pb2.line.fill.background()
tb(s,5.30,1.10,1.6,0.35,"G1 - G3",sz=12,b=True,c=WHITE,a=PP_ALIGN.CENTER)
g_frames=[("我把 ___ 和 ___ 放一起，因为 ___ 。","I put ___ and ___ together because ___."),
          ("___ 应该放在 ___ ，因为 ___ 。","___ goes in ___ because ___."),
          ("我们给它贴上「___」标签。","We label it ___."),
          ("这样整理后，可以更快地 ___ 。","Now we can ___ faster.")]
for i,(cn,en) in enumerate(g_frames):
    y=1.65+i*0.78
    tb(s,5.25,y,4.4,0.45,f"·  {cn}",sz=14,b=True,c=TEAL)
    tb(s,5.25,y+0.42,4.4,0.25,en,sz=9,c=GRAY)
tb(s,0.4,5.1,9.2,0.3,"💡 把这张 PPT 截屏打印，贴在每张桌子上 — 学生整堂课参考。",sz=11,b=True,c=NAVY,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 39 SESSION 3 DIVIDER
# ============================================================
div("Session 3  下午","动手 + 给建议\n🛠️ 小组收纳挑战  🔍 方案测试  💡 给建议",NAVY,"🧑‍🔧")
n+=1

# ============================================================
# 40 给建议 — tiered advice activity
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"💡 整理小顾问  Give Advice!",NAVY)
tb(s,0.4,0.88,9.2,0.32,"看情景，给整理建议 — 按年龄有两种任务!",sz=13,c=GRAY,a=PP_ALIGN.CENTER)
# Low level
lp=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(1.3),Inches(4.6),Inches(3.5))
lp.fill.solid();lp.fill.fore_color.rgb=WHITE;lp.line.color.rgb=GOLD;lp.line.width=Pt(2.5)
lh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(1.3),Inches(4.6),Inches(0.5))
lh.fill.solid();lh.fill.fore_color.rgb=GOLD;lh.line.fill.background()
tb(s,0.45,1.37,4.4,0.4,"🐣 低龄  4–6 岁",sz=14,b=True,c=WHITE)
ib(s,0.5,1.95,4.2,1.35,"📷 玩具到处都是 / Toys everywhere")
tf=tb(s,0.5,3.4,4.3,0.4,"🗣️ 简单建议:",sz=13,b=True,c=GOLD)
ap(tf,"把 ____ 放回 ____ 。",sz=15,b=True,c=DARK)
ap(tf,"用玩具垫，一拉就收好!",sz=12,c=DARK)
# High level
rp=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.1),Inches(1.3),Inches(4.6),Inches(3.5))
rp.fill.solid();rp.fill.fore_color.rgb=WHITE;rp.line.color.rgb=TEAL;rp.line.width=Pt(2.5)
rh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.1),Inches(1.3),Inches(4.6),Inches(0.5))
rh.fill.solid();rh.fill.fore_color.rgb=TEAL;rh.line.fill.background()
tb(s,5.25,1.37,4.4,0.4,"🐔 高龄  7 岁以上",sz=14,b=True,c=WHITE)
ib(s,5.3,1.95,4.2,1.35,"📷 书桌又乱又满 / A messy desk")
tf2=tb(s,5.3,3.4,4.3,0.4,"🗣️ 说清楚 + 说理由:",sz=13,b=True,c=TEAL)
ap(tf2,"先把 ___ 分类，再 ___ ，",sz=14,b=True,c=DARK)
ap(tf2,"因为 ___ 。常用的放手边，贴标签。",sz=12,c=DARK)
tb(s,0.4,4.95,9.2,0.35,"💬 用四步整理法 + 三个核心概念，给出你的建议!",sz=13,b=True,c=NAVY,a=PP_ALIGN.CENTER)
notes(s,"给建议活动，2个层次:低龄(4-6)给简单建议+简单句;高龄(7+)分类+说理由+标签。可换不同情景卡。")
pn(s,n)

# ============================================================
# 41 Day 2 Badge
# ============================================================
s=ns();n+=1;bg(s,CREAM)
tb(s,0.5,0.3,9,0.7,"🎖️ Day 2 整理达人徽章  Master Badge",sz=24,b=True,c=TEAL,a=PP_ALIGN.CENTER)
sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.5),Inches(1.05),Inches(3),Inches(3))
sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=TEAL;sh.line.width=Pt(5)
tf=tb(s,3.6,1.28,2.8,0.4,"DAY 2",sz=18,b=True,c=CORAL,a=PP_ALIGN.CENTER)
ap(tf,"📦",sz=40,a=PP_ALIGN.CENTER)
ap(tf,"整理和收纳",sz=19,b=True,c=TEAL,a=PP_ALIGN.CENTER)
ap(tf,"✓ COMPLETED",sz=13,b=True,c=GREEN_OK,a=PP_ALIGN.CENTER)
ap(tf,"🗂️🏠✅🧰",sz=16,a=PP_ALIGN.CENTER)
sb=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(1.3),Inches(4.20),Inches(7.4),Inches(0.65))
sb.fill.solid();sb.fill.fore_color.rgb=SUNNY;sb.line.color.rgb=CORAL;sb.line.width=Pt(2.5)
tb(s,1.3,4.25,7.4,0.55,"⭐  ⭐  ⭐  ⭐  ⭐  ⭐",sz=32,b=True,c=CORAL,a=PP_ALIGN.CENTER)
tb(s,1,4.90,8,0.4,"今天学会了整理和收纳! 🎉",sz=16,b=True,c=TEAL,a=PP_ALIGN.CENTER)
tb(s,1,5.24,8,0.3,"四步整理法 · 三个核心 · 7 种收纳工具 · 5 个词",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 42 Tomorrow preview
# ============================================================
s=ns();n+=1;bg(s,TEAL)
tb(s,0.5,0.9,9,0.8,"🌟 明天见！  See You Tomorrow!",sz=32,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tf=tb(s,1.5,2.2,7,2.5,"Day 3 — 待定 (To Be Continued)",sz=26,b=True,c=SUNNY,a=PP_ALIGN.CENTER)
ap(tf,"",sz=10)
ap(tf,"🏠 继续做小小生活家！",sz=20,b=True,c=WHITE,a=PP_ALIGN.CENTER)
ap(tf,"Keep being a little homemaker!",sz=14,c=WARM,a=PP_ALIGN.CENTER)
ap(tf,"",sz=10)
ap(tf,"明天见，整理达人！",sz=15,c=WARM,a=PP_ALIGN.CENTER)
pn(s,n)

OUT='/Users/huanli/projects/courseppt/Chinese/小小生活家/day2_organizing.pptx'
prs.save(OUT);print(f"Created {n} slides → {OUT}")
