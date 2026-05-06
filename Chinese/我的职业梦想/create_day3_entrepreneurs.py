#!/usr/bin/env python3
"""
我的职业梦想 — Day 3: 小小企业家 (Little Entrepreneurs)
Focus: famous entrepreneurs (Jobs etc.) + business idea pitch project
"我有一个金点子! — 我想做一个 ___ 帮助 ___ "
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

# --- Palette: Startup energy (warm gold + black + accent red) ---
GOLD   = RGBColor(0xF5,0xA6,0x23)   # primary — startup energy
JET    = RGBColor(0x14,0x14,0x14)
APPLE  = RGBColor(0xC7,0xC7,0xC7)   # silver
TESLA  = RGBColor(0xC8,0x1B,0x1B)   # red
ALI    = RGBColor(0xFF,0x66,0x00)   # alibaba orange
MS     = RGBColor(0x00,0xA4,0xEF)   # microsoft cyan
NAVY   = RGBColor(0x1E,0x3A,0x5F)
GREEN  = RGBColor(0x2E,0x7D,0x32)
PURPLE = RGBColor(0x7B,0x1F,0xA2)
CREAM  = RGBColor(0xFF,0xF8,0xE7)
WARM   = RGBColor(0xFF,0xF3,0xE0)
BROWN  = RGBColor(0x6B,0x44,0x23)
WHITE  = RGBColor(0xFF,0xFF,0xFF)
DARK   = RGBColor(0x2C,0x2C,0x2C)
GRAY   = RGBColor(0x88,0x88,0x88)
LGRAY  = RGBColor(0xBB,0xBB,0xBB)
IMGBG  = RGBColor(0xE8,0xE8,0xE8)
OK     = RGBColor(0x38,0x8E,0x3C)

# === Helpers (same conventions) ===
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
def hb(s,txt,c=GOLD,t=0.15):
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
def sentence_frame_bar(s,t,frame_cn,frame_en,accent=GOLD):
    if t > 4.95: t = 4.95
    sf=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(t),Inches(9.4),Inches(0.65))
    sf.fill.solid(); sf.fill.fore_color.rgb=WARM; sf.line.color.rgb=accent; sf.line.width=Pt(2)
    tb(s,0.5,t+0.1,1.7,0.4,"💬 我来说",sz=14,b=True,c=accent)
    tb(s,2.0,t+0.07,7.6,0.3,frame_cn,sz=14,b=True,c=DARK)
    tb(s,2.0,t+0.32,7.6,0.3,frame_en,sz=10,c=GRAY)


def mission_card(s,l,t,w,h,num,task_cn,task_en,emoji,color):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(l),Inches(t),Inches(w),Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=color; sh.line.width=Pt(2.5)
    badge=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(l+0.1),Inches(t+0.08),Inches(0.55),Inches(0.55))
    badge.fill.solid(); badge.fill.fore_color.rgb=color; badge.line.fill.background()
    tb(s,l+0.1,t+0.18,0.55,0.4,str(num),sz=18,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,l+0.7,t+0.15,w-0.8,0.4,task_en,sz=10,c=GRAY)
    tb(s,l+0.05,t+0.85,w-0.1,0.7,emoji,sz=44,a=PP_ALIGN.CENTER)
    tb(s,l+0.05,t+1.55,w-0.1,0.4,task_cn,sz=18,b=True,c=color,a=PP_ALIGN.CENTER)

def person_card(emoji,cn,en,years,company,one_line_cn,one_line_en,fact_cn,fact_en,color):
    s=ns(); bg(s,CREAM); hb(s,f"{emoji} 著名企业家 · {cn}",color)
    img=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.95),Inches(3.7),Inches(3.7))
    img.fill.solid(); img.fill.fore_color.rgb=IMGBG; img.line.color.rgb=color; img.line.width=Pt(3)
    tb(s,0.4,2.15,3.7,1.0,emoji,sz=110,a=PP_ALIGN.CENTER)
    tb(s,0.4,3.30,3.7,0.45,cn,sz=22,b=True,c=color,a=PP_ALIGN.CENTER)
    tb(s,0.4,3.75,3.7,0.30,en,sz=12,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,0.4,4.10,3.7,0.30,f"{company} · {years}",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.30),Inches(0.95),Inches(5.40),Inches(3.7))
    panel.fill.solid(); panel.fill.fore_color.rgb=WHITE; panel.line.color.rgb=color; panel.line.width=Pt(2.5)
    head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.30),Inches(0.95),Inches(5.40),Inches(0.50))
    head.fill.solid(); head.fill.fore_color.rgb=color; head.line.fill.background()
    tb(s,4.45,1.03,5.1,0.4,"💡 一句话介绍  In One Line",sz=14,b=True,c=WHITE)
    tb(s,4.45,1.55,5.1,0.45,one_line_cn,sz=17,b=True,c=DARK)
    tb(s,4.45,2.00,5.1,0.32,one_line_en,sz=11,c=GRAY)
    tb(s,4.45,2.50,5.1,0.35,"🌟 他改变了什么？",sz=14,b=True,c=color)
    tb(s,4.45,2.85,5.1,0.40,fact_cn,sz=14,c=DARK)
    tb(s,4.45,3.30,5.1,0.32,fact_en,sz=10,c=GRAY)
    sentence_frame_bar(s,4.80,
        f"{cn} 做了 ___ 。 我也想 ___ 。",
        f"{en} did ___. I also want to ___.")

n=0

# 1. COVER
s=ns(); bg(s,CREAM)
tb(s,1,0.45,8,0.6,"我的职业梦想 · My Dream Career",sz=26,b=True,c=GOLD,a=PP_ALIGN.CENTER)
tb(s,1,1.05,8,0.4,"Day 3",sz=18,b=True,c=JET,a=PP_ALIGN.CENTER)
sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.25),Inches(1.55),Inches(3.5),Inches(3.5))
sh.fill.solid(); sh.fill.fore_color.rgb=GOLD; sh.line.color.rgb=JET; sh.line.width=Pt(6)
sh2=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.55),Inches(1.85),Inches(2.9),Inches(2.9))
sh2.fill.solid(); sh2.fill.fore_color.rgb=WHITE; sh2.line.color.rgb=JET; sh2.line.width=Pt(2)
tf=tb(s,3.55,2.05,2.9,0.4,"DAY 3",sz=14,b=True,c=GOLD,a=PP_ALIGN.CENTER)
ap(tf,"💡",sz=68,a=PP_ALIGN.CENTER)
ap(tf,"小小企业家",sz=20,b=True,c=JET,a=PP_ALIGN.CENTER)
ap(tf,"LITTLE ENTREPRENEURS",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,1,5.15,8,0.4,"💡 你有一个金点子吗? Do you have a brilliant idea?",sz=14,b=True,c=TESLA,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)

# 2. TODAY'S MISSION
s=ns(); bg(s,CREAM); hb(s,"🧭 今天的任务  Today's Mission",GOLD)
tb(s,0.4,0.85,9.2,0.45,"💡 认识 4 位改变世界的企业家!",sz=24,b=True,c=GOLD,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.28,"Meet 4 entrepreneurs who changed the world.",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,0.4,1.55,9.2,0.32,"👉 5 个任务 ↓",sz=14,b=True,c=BROWN,a=PP_ALIGN.CENTER)
mission_card(s,0.4,1.95,1.80,2.20,1,"企业家做什么","What",   "💼",GOLD)
mission_card(s,2.30,1.95,1.80,2.20,2,"认识 4 位","Meet 4",     "👋",JET)
mission_card(s,4.20,1.95,1.80,2.20,3,"猜公司","Guess Brands",  "🏢",TESLA)
mission_card(s,6.10,1.95,1.80,2.20,4,"我会认/写","Read & Write","📖",PURPLE)
mission_card(s,8.00,1.95,1.80,2.20,5,"金点子!","Pitch!",       "💡",GREEN)
sentence_frame_bar(s,4.40,"我想发明 ___ 。","I want to invent ___.")
n+=1; pn(s,n)

# 3. SESSION 1 DIVIDER
s=div("Session 1  上午","💼 企业家是什么 + 4 位巨人",GOLD,"🌟"); n+=1; pn(s,n)

# 4. WHAT IS AN ENTREPRENEUR?
s=ns(); bg(s,CREAM); hb(s,"💡 企业家是什么？  What is an Entrepreneur?",GOLD)
tb(s,0.4,0.85,9.2,0.4,"企业家做 4 件事:",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.28,"Entrepreneurs do 4 things:",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
defs=[
    ("👀","看见机会","See chance","看到别人没看到的需要"),
    ("💡","想金点子","Have ideas","用新办法解决问题"),
    ("🤝","组队伍","Build a team","找人一起做"),
    ("🚀","坚持做","Keep going","失败 1000 次, 再试 1001 次"),
]
for i,(em,cn,en,desc) in enumerate(defs):
    x=0.4+i*2.30
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.65),Inches(2.2),Inches(3.0))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=GOLD; sh.line.width=Pt(2.5)
    tb(s,x+0.05,1.80,2.10,0.8,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.65,2.10,0.45,cn,sz=18,b=True,c=GOLD,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.10,2.10,0.30,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.10,3.50,2.00,1.10,desc,sz=12,c=DARK,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.85,"企业家就是 ___ 。","An entrepreneur ___ .")
n+=1; pn(s,n)

# 5. INTRO TO 4 FAMOUS ENTREPRENEURS
s=ns(); bg(s,CREAM); hb(s,"👋 4 位改变世界的企业家  4 World-Changers",JET)
tb(s,0.4,0.85,9.2,0.34,"猜猜看 — 你认识他们的公司吗？",sz=15,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.20,9.2,0.28,"Guess — do you know their companies?",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
ent=[
    ("🍎","乔布斯","Steve Jobs","iPhone",JET),
    ("💻","比尔·盖茨","Bill Gates","Windows",MS),
    ("🛒","马云","Jack Ma","淘宝/支付宝",ALI),
    ("🚀","马斯克","Elon Musk","Tesla/SpaceX",TESLA),
]
for i,(em,cn,en,sym,cl) in enumerate(ent):
    x=0.4+i*2.30
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.55),Inches(2.20),Inches(3.30))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    tb(s,x+0.05,1.70,2.10,0.9,em,sz=68,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.65,2.10,0.45,cn,sz=20,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.10,2.10,0.30,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    sep=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(x+0.30),Inches(3.55),Inches(1.60),Inches(0.02))
    sep.fill.solid(); sep.fill.fore_color.rgb=cl; sep.line.fill.background()
    tb(s,x+0.05,3.70,2.10,0.7,sym,sz=14,b=True,c=cl,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,5.05,"我用过 ___ 的产品。","I have used ___'s product.")
n+=1; pn(s,n)

# 6. JOBS spotlight
person_card("🍎","乔布斯","Steve Jobs","1955–2011","🇺🇸 美国 · Apple",
    "他让电脑变成人人能用的「漂亮」东西。",
    "He made computers beautiful and easy for everyone.",
    "iPhone · iPad · Mac · iPod — 改变了我们听音乐、打电话、用电脑的方式。",
    "iPhone, iPad, Mac, iPod — changed music, phones, and computers forever.",
    JET)
n+=1; pn(s,n)
notes(s,"3-4 分钟:\n• 名言: 'Stay hungry, stay foolish.' 保持饥饿, 保持愚笨\n• 他被自己创办的公司开除过, 后来又回来\n• 重点: 「设计 = 不只是好看, 还要好用」")

# 7. GATES spotlight
person_card("💻","比尔·盖茨","Bill Gates","1955– 至今","🇺🇸 美国 · Microsoft",
    "他让每台电脑都能跑 Windows — 还把财富捐出来救人。",
    "He put Windows on every PC — then gave away his fortune to save lives.",
    "Microsoft · Windows · Office; 比尔和梅琳达基金会帮助消灭疟疾、小儿麻痹症。",
    "Microsoft Windows/Office; Gates Foundation fights malaria & polio worldwide.",
    MS)
n+=1; pn(s,n)

# 8. JACK MA spotlight
person_card("🛒","马云","Jack Ma","1964– 至今","🇨🇳 中国 · 阿里巴巴",
    "他帮中国小老板把东西卖到全世界。",
    "He helped tiny Chinese shops sell to the whole world online.",
    "淘宝 / 支付宝 · 改变了我们买东西、付钱的方式 (英语老师创业!)。",
    "Taobao & Alipay — changed how we shop and pay (started as an English teacher!).",
    ALI)
n+=1; pn(s,n)
notes(s,"重点: 失败 ≠ 输!\n• 马云高考考了 3 次, 申请哈佛 10 次都被拒\n• 「今天很残酷, 明天更残酷, 后天很美好 — 但大多数人死在明天晚上」")

# 9. MUSK spotlight
person_card("🚀","马斯克","Elon Musk","1971– 至今","🇿🇦 南非 · 美国",
    "他想造电车, 还想把人送上火星。",
    "Building electric cars — and trying to send people to Mars.",
    "Tesla 电车 · SpaceX 火箭 · Starlink 卫星互联网 — 让科技为地球和宇宙服务。",
    "Tesla EVs · SpaceX rockets · Starlink — tech for Earth and beyond.",
    TESLA)
n+=1; pn(s,n)

# 10. WHAT IS A "金点子"
s=ns(); bg(s,CREAM); hb(s,"💎 什么是金点子？  What's a Brilliant Idea?",GOLD)
tb(s,0.4,0.85,9.2,0.34,"金点子 = 解决一个真问题 + 让生活更好",sz=15,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.20,9.2,0.28,"A great idea = solves a real problem + makes life better",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
ideas=[
    ("📱","iPhone 之前","Before iPhone","手机只能打电话"),
    ("📱","iPhone 之后","After iPhone","手机能拍照、上网、玩游戏"),
    ("🛒","淘宝之前","Before Taobao","买东西要去商店"),
    ("🛒","淘宝之后","After Taobao","在家点一下, 第二天到家"),
]
for i,(em,cn,en,desc) in enumerate(ideas):
    col=i%2; row=i//2
    x=0.4+col*4.7; y=1.55+row*1.65
    bg_color=WARM if i%2==0 else WHITE
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(4.5),Inches(1.45))
    sh.fill.solid(); sh.fill.fore_color.rgb=bg_color; sh.line.color.rgb=GOLD; sh.line.width=Pt(2)
    tb(s,x+0.10,y+0.05,0.6,0.55,em,sz=30,a=PP_ALIGN.CENTER)
    tb(s,x+0.75,y+0.10,3.7,0.35,cn,sz=15,b=True,c=GOLD)
    tb(s,x+0.75,y+0.42,3.7,0.30,en,sz=10,c=GRAY)
    tb(s,x+0.10,y+0.85,4.35,0.55,f"💭 {desc}",sz=12,c=DARK)
sentence_frame_bar(s,4.95,"以前 ___, 现在 ___ — 真好!","Before ___, now ___ — much better!")
n+=1; pn(s,n)

# 11. SESSION 2 DIVIDER
s=div("Session 2  下午","📖 复习 + 我会认 + 我会写",JET,"📖"); n+=1; pn(s,n)

# 12. REVIEW
s=ns(); bg(s,CREAM); hb(s,"🔄 复习  Review · Session 1",GOLD)
tb(s,0.4,0.85,9.2,0.4,"还记得吗？  Do you remember?",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
qs=[
    ("💼","企业家做什么？","看机会、想点子、组队伍、坚持"),
    ("🍎","谁创办了苹果?","乔布斯 (Steve Jobs)"),
    ("💻","谁让每台电脑都有 Windows?","比尔·盖茨"),
    ("🛒","谁让中国人在网上买东西?","马云"),
    ("🚀","谁造电车 + 火箭?","马斯克"),
]
for i,(em,q,a) in enumerate(qs):
    y=1.45+i*0.70
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(y),Inches(9.2),Inches(0.6))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=GOLD; sh.line.width=Pt(1.5)
    tb(s,0.55,y+0.10,0.5,0.4,em,sz=22,a=PP_ALIGN.CENTER)
    tb(s,1.15,y+0.10,4.0,0.4,q,sz=14,b=True,c=DARK)
    tb(s,5.30,y+0.10,4.2,0.4,f"→ {a}",sz=14,b=True,c=GOLD)
n+=1; pn(s,n)

# 13. 我会认 — 4 business words
s=ns(); bg(s,CREAM); hb(s,"📖 我会认  I Can Read · 4 个商业词",GOLD)
words=[
    ("生意","shēng yì","business","💼"),
    ("点子","diǎn zi","idea","💡"),
    ("公司","gōng sī","company","🏢"),
    ("梦想","mèng xiǎng","dream","✨"),
]
for i,(cn,py,en,em) in enumerate(words):
    col=i%2; row=i//2
    x=0.4+col*4.7; y=0.95+row*2.0
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(4.5),Inches(1.85))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=GOLD; sh.line.width=Pt(2.5)
    tb(s,x+0.15,y+0.10,1.0,1.5,em,sz=70,a=PP_ALIGN.CENTER)
    tb(s,x+1.30,y+0.20,3.0,0.7,cn,sz=40,b=True,c=GOLD)
    tb(s,x+1.30,y+0.95,3.0,0.35,py,sz=14,c=GRAY)
    tb(s,x+1.30,y+1.30,3.0,0.40,en,sz=14,b=True,c=DARK)
n+=1; pn(s,n)

# 14. 我会写 — 想 (think / want)
s=ns(); bg(s,CREAM); hb(s,"✏️ 我会写 · 想  I Can Write · 'think / want'",GOLD)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.95),Inches(4.5),Inches(3.7))
sh.fill.solid(); sh.fill.fore_color.rgb=WARM; sh.line.color.rgb=GOLD; sh.line.width=Pt(3)
tb(s,0.4,1.30,4.5,2.5,"想",sz=300,b=True,c=GOLD,a=PP_ALIGN.CENTER)
tb(s,0.4,4.10,4.5,0.40,"xiǎng · think / want",sz=18,c=GRAY,a=PP_ALIGN.CENTER)
panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.05),Inches(0.95),Inches(4.65),Inches(3.7))
panel.fill.solid(); panel.fill.fore_color.rgb=WHITE; panel.line.color.rgb=GOLD; panel.line.width=Pt(2)
tb(s,5.20,1.10,4.4,0.5,"📝 笔顺  13 笔",sz=18,b=True,c=GOLD)
tb(s,5.20,1.65,4.4,0.4,"上面: 「相」(look at)",sz=14,c=DARK)
tb(s,5.20,2.05,4.4,0.4,"下面: 「心」(heart)",sz=14,c=DARK)
tb(s,5.20,2.45,4.4,0.4,"用「心」「看」 = 想!",sz=14,b=True,c=GOLD)
tb(s,5.20,2.95,4.4,0.35,"📝 练习 Practice:",sz=14,b=True,c=GOLD)
tb(s,5.20,3.35,4.4,0.35,"1️⃣ 空中写 · 2️⃣ 手心写 · 3️⃣ 纸上 3 次",sz=13,c=DARK)
sentence_frame_bar(s,4.85,"我 想 做 ___ 。 我 想 当 企业家。","I want to do ___. I want to be an entrepreneur.")
n+=1; pn(s,n)

# 15. SESSION 3 DIVIDER — Pitch project
s=div("Session 3  下午","💎 我有一个金点子! Pitch Day",TESLA,"🎤"); n+=1; pn(s,n)

# 16. PROJECT INTRO — Pitch
s=ns(); bg(s,CREAM); hb(s,"🎤 我有一个金点子!  My Brilliant Idea!",TESLA)
tb(s,0.4,0.85,9.2,0.5,"💡 想一个能帮别人的点子, 60 秒讲给大家听!",sz=20,b=True,c=TESLA,a=PP_ALIGN.CENTER)
tb(s,0.4,1.45,9.2,0.34,"Think of an idea that helps people — pitch in 60 seconds!",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(2.5),Inches(1.90),Inches(5.0),Inches(2.85))
sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=TESLA; sh.line.width=Pt(3)
tb(s,2.5,2.00,5.0,0.45,"🎤 我的点子卡片",sz=18,b=True,c=TESLA,a=PP_ALIGN.CENTER)
tb(s,2.5,2.45,5.0,0.30,"My Pitch Card",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,2.7,2.90,4.6,0.40,"1. 我的点子叫: ___",sz=13,b=True,c=DARK)
tb(s,2.7,3.30,4.6,0.40,"2. 它解决: ___ 的问题",sz=13,b=True,c=DARK)
tb(s,2.7,3.70,4.6,0.40,"3. 它怎么用: ___",sz=13,b=True,c=DARK)
tb(s,2.7,4.10,4.6,0.40,"4. 我会找 ___ 一起做",sz=13,b=True,c=DARK)
tb(s,0.4,4.85,9.2,0.32,"⏱️ 15 分钟想 + 写 + 画 + 60 秒讲",sz=13,b=True,c=BROWN,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)

# 17. STEP 1 — IDEA brainstorm
s=ns(); bg(s,CREAM); hb(s,"1️⃣ 想点子  Step 1 · Brainstorm",TESLA)
tb(s,0.4,0.85,9.2,0.4,"看看你身边 — 谁需要帮助？",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.30,"Who needs help around you?",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
prompts=[
    ("👶","小朋友","Kids","作业辅导 App?"),
    ("👴","老人","Elderly","送药机器人?"),
    ("🐶","宠物","Pets","自动喂食器?"),
    ("🌳","环境","Earth","可吃的吸管?"),
    ("🎒","学校","School","背包整理盒?"),
    ("🍔","吃饭","Food","健康零食?"),
]
for i,(em,cn,en,ex) in enumerate(prompts):
    col=i%3; row=i//3
    x=0.3+col*3.2; y=1.65+row*1.55
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(3.0),Inches(1.40))
    sh.fill.solid(); sh.fill.fore_color.rgb=WARM; sh.line.color.rgb=TESLA; sh.line.width=Pt(2)
    tb(s,x+0.10,y+0.05,0.6,0.55,em,sz=28,a=PP_ALIGN.CENTER)
    tb(s,x+0.75,y+0.10,2.2,0.35,cn,sz=15,b=True,c=TESLA)
    tb(s,x+0.75,y+0.42,2.2,0.30,en,sz=10,c=GRAY)
    tb(s,x+0.10,y+0.83,2.85,0.5,f"💭 {ex}",sz=12,c=DARK)
sentence_frame_bar(s,4.95,"我想帮 ___ 。","I want to help ___.")
n+=1; pn(s,n)

# 18. STEP 2 — DESIGN (4 questions)
s=ns(); bg(s,CREAM); hb(s,"2️⃣ 设计点子  Step 2 · Design",TESLA)
tb(s,0.4,0.85,9.2,0.4,"问自己 4 个问题 — 让点子更好",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
qs=[
    ("❓","解决什么？","Which problem?"),
    ("🧩","怎么用？","How to use?"),
    ("👥","谁会买？","Who will buy?"),
    ("✨","有什么不同？","What's special?"),
]
for i,(em,cn,en) in enumerate(qs):
    x=0.4+i*2.30
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.55),Inches(2.20),Inches(2.85))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=TESLA; sh.line.width=Pt(2.5)
    tb(s,x+0.05,1.70,2.10,0.7,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.50,2.10,0.45,cn,sz=18,b=True,c=TESLA,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.00,2.10,0.30,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.10,3.50,2.00,0.85,"____________",sz=14,c=GRAY,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.60,"我的点子是 ___ 。 它能 ___ 。","My idea is ___. It can ___.")
n+=1; pn(s,n)

# 19. STEP 3 — PITCH
s=ns(); bg(s,CREAM); hb(s,"3️⃣ 60 秒讲  Step 3 · 60-Second Pitch",TESLA)
tb(s,0.4,0.85,9.2,0.4,"轮到你了 — 60 秒, 全班决定要不要投资!",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.30,"60 seconds — class decides whether to 'invest'!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
steps=[
    ("🎙️","Open","「大家好, 我的点子叫 ___」"),
    ("⚠️","Problem","「___ 的人有 ___ 问题」"),
    ("💡","Solution","「我做了一个 ___ 帮他们 ___」"),
    ("🤝","Ask","「谁愿意当我的合伙人?」"),
]
for i,(em,cn,desc) in enumerate(steps):
    x=0.4+i*2.30
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.65),Inches(2.20),Inches(2.85))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=TESLA; sh.line.width=Pt(2.5)
    tb(s,x+0.05,1.80,2.10,0.7,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.65,2.10,0.45,cn,sz=20,b=True,c=TESLA,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.30,2.10,1.10,desc,sz=12,c=DARK,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.65,"我投 ___ 一票, 因为 ___ 。","I invest in ___ because ___.")
n+=1; pn(s,n)

# 20. CLOSING BADGE
s=ns(); bg(s,CREAM)
tb(s,0.5,0.4,9,0.8,"🎖️ Day 3 徽章  Little Entrepreneur Badge",sz=26,b=True,c=GOLD,a=PP_ALIGN.CENTER)
sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.5),Inches(1.4),Inches(3),Inches(3))
sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=GOLD; sh.line.width=Pt(6)
tf=tb(s,3.6,1.65,2.8,2.7,"DAY 3",sz=18,b=True,c=GOLD,a=PP_ALIGN.CENTER)
ap(tf,"💡",sz=42,a=PP_ALIGN.CENTER)
ap(tf,"小小企业家",sz=18,b=True,c=JET,a=PP_ALIGN.CENTER)
ap(tf,"✓ COMPLETED",sz=12,b=True,c=OK,a=PP_ALIGN.CENTER)
ap(tf,"🍎💻🛒🚀",sz=16,a=PP_ALIGN.CENTER)
tb(s,1,4.55,8,0.4,"🎉 你也是一位小小企业家! You are an Entrepreneur now!",sz=15,b=True,c=GOLD,a=PP_ALIGN.CENTER)
tb(s,1,5.0,8,0.4,"明天 Day 4 — 帮助别人 Helpers · 医生 + 老师",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)

out=os.path.join(os.path.dirname(__file__),"day3_entrepreneurs.pptx")
prs.save(out); print(f"Saved {out}  ({n} slides)")
