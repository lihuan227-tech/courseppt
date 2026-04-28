#!/usr/bin/env python3
"""
小小艺术家 Little Artist Unit — Day 1: 艺术是表达 Art is Expression
Structure modeled on world-trip Day 1, distinct palette (magenta + sunshine + sky).
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

# --- Palette: Playful Art Studio ---
MAGENTA = RGBColor(0xD8,0x1B,0x60)  # primary berry-magenta
YELLOW  = RGBColor(0xFF,0xC1,0x07)  # sunshine
SKY     = RGBColor(0x42,0xA5,0xF5)  # sky blue
CORAL   = RGBColor(0xFF,0x70,0x43)  # coral orange
PURPLE  = RGBColor(0x9C,0x27,0xB0)  # plum
GREEN_L = RGBColor(0x66,0xBB,0x6A)  # leaf green
CREAM   = RGBColor(0xFF,0xF8,0xE7)  # warm cream background
WHITE   = RGBColor(0xFF,0xFF,0xFF)
DARK    = RGBColor(0x2C,0x2C,0x2C)
GRAY    = RGBColor(0x88,0x88,0x88)
LGRAY   = RGBColor(0xBB,0xBB,0xBB)
WARM    = RGBColor(0xFF,0xF3,0xE0)
IMGBG   = RGBColor(0xF0,0xE8,0xF0)   # soft rose-tinted gray
GREEN_OK= RGBColor(0x38,0x8E,0x3C)

# Emotion colors
E_HAPPY = RGBColor(0xFF,0xB3,0x00)  # 开心 bright gold
E_SAD   = RGBColor(0x5C,0x6B,0xC0)  # 难过 indigo
E_ANGRY = RGBColor(0xE5,0x3E,0x3E)  # 生气 red
E_LOVE  = RGBColor(0xEC,0x40,0x7A)  # 喜欢 pink
E_MOOD  = RGBColor(0xAB,0x47,0xBC)  # 心情 purple

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
def hb(s,txt,c=MAGENTA,t=0.15):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(t),Inches(9.4),Inches(0.55));sh.fill.solid();sh.fill.fore_color.rgb=c;sh.line.fill.background()
    tb(s,0.4,t+0.03,9.2,0.5,txt,sz=20,b=True,c=WHITE)
def pn(s,n): tb(s,9.0,5.25,0.8,0.3,str(n),sz=10,c=GRAY,a=PP_ALIGN.RIGHT)
def div(title,sub,color,emoji=""):
    s=ns();bg(s,color)
    tb(s,1,1.5,8,1.2,f"{emoji} {title}",sz=42,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,1,2.8,8,0.8,sub,sz=22,c=WHITE,a=PP_ALIGN.CENTER)
    # playful dots motif — pick 4 colors that contrast with the bg
    palette=[YELLOW,CORAL,SKY,GREEN_L,PURPLE,MAGENTA]
    dot_colors=[c for c in palette if c!=color][:4]
    positions=[(0.8,4.7),(1.6,4.5),(7.8,4.5),(8.6,4.7)]
    for (x,y),cl in zip(positions,dot_colors):
        d=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x),Inches(y),Inches(0.4),Inches(0.4))
        d.fill.solid();d.fill.fore_color.rgb=cl;d.line.fill.background()
    return s

def dot(s,x,y,r,color):
    d=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x),Inches(y),Inches(r),Inches(r))
    d.fill.solid();d.fill.fore_color.rgb=color;d.line.fill.background()

def example_slide(cat_emoji, cat_cn, work_cn, work_en, artist, fact, img_lb, color):
    """Museum-card style: colored header + big image + caption card."""
    s=ns();bg(s,CREAM)
    hb(s,f"{cat_emoji} {cat_cn}  ·  《{work_cn}》",color)
    ib(s,0.5,0.95,9,3.5,img_lb)
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.5),Inches(4.55),Inches(9),Inches(0.85))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=color;sh.line.width=Pt(2)
    tb(s,0.7,4.6,4.5,0.4,work_en,sz=15,b=True,c=color)
    tb(s,0.7,4.95,4.5,0.3,artist,sz=11,c=GRAY)
    tb(s,5.3,4.62,4.2,0.75,fact,sz=12,c=DARK)
    return s

def example_slide_generic(cat_emoji, cat_cn, name_cn, name_en, sub_label, fact, img_lb, color):
    """Like example_slide but without 《》 brackets — for instruments, dance types, etc."""
    s=ns();bg(s,CREAM)
    hb(s,f"{cat_emoji} {cat_cn}  ·  {name_cn}",color)
    ib(s,0.5,0.95,9,3.5,img_lb)
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.5),Inches(4.55),Inches(9),Inches(0.85))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=color;sh.line.width=Pt(2)
    tb(s,0.7,4.6,4.5,0.4,name_en,sz=15,b=True,c=color)
    tb(s,0.7,4.95,4.5,0.3,sub_label,sz=11,c=GRAY)
    tb(s,5.3,4.62,4.2,0.75,fact,sz=12,c=DARK)
    return s

def type_overview_slide(emoji, name_cn, name_en, color, hint, subtypes):
    """Category overview — grid of sub-types to introduce before examples."""
    s=ns();bg(s,CREAM);hb(s,f"{emoji} {name_cn}  {name_en}",color)
    tb(s,0.4,0.85,9,0.35,hint,sz=13,c=GRAY,a=PP_ALIGN.CENTER)
    cols = 3 if len(subtypes)<=6 else 4
    rows = (len(subtypes)+cols-1)//cols
    card_w = (9.4 - 0.15*(cols-1))/cols
    card_h = 1.7 if rows<=2 else 1.45
    y_start = 1.3 if rows<=2 else 1.25
    y_step = card_h + 0.25 if rows<=2 else card_h + 0.2
    for i,(sem,scn,sen) in enumerate(subtypes):
        col=i%cols;row=i//cols
        x=0.3+col*(card_w+0.15);y=y_start+row*y_step
        sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(card_w),Inches(card_h))
        sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=color;sh.line.width=Pt(2.5)
        tb(s,x+0.1,y+0.1,card_w-0.2,0.55,sem,sz=30,c=color,a=PP_ALIGN.CENTER)
        tb(s,x+0.1,y+0.75,card_w-0.2,0.4,scn,sz=17,b=True,c=DARK,a=PP_ALIGN.CENTER)
        tb(s,x+0.1,y+1.15,card_w-0.2,0.3,sen,sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    return s

def word_card_read(w,py,en,sent,img,color=MAGENTA):
    s=ns();bg(s,CREAM);hb(s,"👀 我会认  I Can Read",color)
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.0),Inches(4.5),Inches(2.5))
    sh.fill.solid();sh.fill.fore_color.rgb=WARM;sh.line.color.rgb=color;sh.line.width=Pt(2)
    tb(s,0.5,1.1,4.3,1.4,w,sz=72,b=True,c=color,a=PP_ALIGN.CENTER)
    tb(s,0.5,2.4,4.3,0.4,f"{py}  {en}",sz=20,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,0.5,2.85,4.3,0.4,"👉 跟我读！Read after me!",sz=14,c=color,a=PP_ALIGN.CENTER)
    ib(s,5.3,1.0,4.4,2.5,img)
    sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(3.8),Inches(9.2),Inches(1.2))
    sh2.fill.solid();sh2.fill.fore_color.rgb=WHITE;sh2.line.color.rgb=color;sh2.line.width=Pt(2)
    tb(s,0.6,3.9,1.5,0.4,"例句",sz=16,b=True,c=color)
    tb(s,0.6,4.3,8.8,0.5,sent,sz=22,b=True,c=DARK)
    return s

def word_card_write(w,py,en,img,color=MAGENTA):
    s=ns();bg(s,CREAM);hb(s,"✍️ 我会写  I Can Write",color)
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.0),Inches(4.5),Inches(2.0))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=color;sh.line.width=Pt(3)
    tb(s,0.5,1.05,4.3,1.2,w,sz=72,b=True,c=color,a=PP_ALIGN.CENTER)
    tb(s,0.5,2.2,4.3,0.4,f"{py}  {en}",sz=20,c=GRAY,a=PP_ALIGN.CENTER)
    ib(s,5.3,1.0,4.4,2.0,img)
    sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(3.3),Inches(5.0),Inches(1.8))
    sh2.fill.solid();sh2.fill.fore_color.rgb=WARM;sh2.line.fill.background()
    tb(s,0.6,3.4,4.6,0.4,"📝 笔顺 Stroke Order",sz=16,b=True,c=color)
    ib(s,0.6,3.9,4.6,1.0,"📷 插入笔顺图片")
    tf=tb(s,5.8,3.4,3.8,0.4,"练习步骤 Practice:",sz=14,b=True,c=color)
    ap(tf,"1. 空中写 Air Write",sz=13,c=DARK)
    ap(tf,"2. 手心写 Palm Write",sz=13,c=DARK)
    ap(tf,"3. 纸上写 3 times",sz=13,c=DARK)
    return s

n=0

# ============================================================
# 1 COVER — Artist Palette badge
# ============================================================
s=ns();n+=1;bg(s,CREAM)
tb(s,1,0.3,8,0.7,"Little Artist Studio",sz=32,b=True,c=MAGENTA,a=PP_ALIGN.CENTER)
tb(s,1,0.9,8,0.45,"小小艺术家",sz=22,c=MAGENTA,a=PP_ALIGN.CENTER)
# big circle badge w/ paint blobs around
sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.25),Inches(1.55),Inches(3.5),Inches(3.5))
sh.fill.solid();sh.fill.fore_color.rgb=MAGENTA;sh.line.color.rgb=YELLOW;sh.line.width=Pt(6)
sh2=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.55),Inches(1.85),Inches(2.9),Inches(2.9))
sh2.fill.solid();sh2.fill.fore_color.rgb=WHITE;sh2.line.color.rgb=YELLOW;sh2.line.width=Pt(2)
tf=tb(s,3.6,2.0,2.8,0.4,"DAY 1",sz=16,b=True,c=CORAL,a=PP_ALIGN.CENTER)
ap(tf,"🎨",sz=56,a=PP_ALIGN.CENTER)
ap(tf,"艺术是表达",sz=20,b=True,c=MAGENTA,a=PP_ALIGN.CENTER)
ap(tf,"ART IS EXPRESSION",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# paint-splash dots around the badge
for x,y,r,c in [(1.2,1.9,0.6,YELLOW),(1.5,3.8,0.5,SKY),(7.8,1.6,0.55,CORAL),(8.2,3.6,0.5,GREEN_L),(1.0,4.7,0.35,PURPLE),(8.6,4.6,0.35,E_LOVE)]:
    dot(s,x,y,r,c)
tb(s,1,5.05,8,0.4,"🎨 拿起你的画笔，我们出发！Pick up your brush!",sz=14,b=True,c=MAGENTA,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 2 SCHEDULE
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"⏰ 今日时间安排  Today's Schedule")
for i,(nm,tm,dc,cl) in enumerate([
    ("Session 1  上午","11:00-11:45","理解艺术是表达 + 认识艺术形式",MAGENTA),
    ("Session 2  下午","2:00-2:45","复习 + 语言目标 (心情词语)",SKY),
    ("Session 3  下午","3:00-4:30","Project：我的心情画 / 我是谁 自画像",CORAL)]):
    y=0.9+i*1.5
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.5),Inches(y),Inches(9),Inches(1.2))
    sh.fill.solid();sh.fill.fore_color.rgb=cl;sh.line.fill.background()
    tb(s,0.7,y+0.15,4,0.4,nm,sz=20,b=True,c=WHITE)
    tb(s,0.7,y+0.6,3,0.4,tm,sz=15,c=WARM)
    tb(s,4.6,y+0.35,5.0,0.6,dc,sz=15,c=WHITE)
pn(s,n)

# ============================================================
# 3 OBJECTIVES
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🎯 教学目标  Learning Objectives")
tb(s,0.5,0.9,9,0.5,"🎨 内容目标  Content:",sz=20,b=True,c=MAGENTA)
tf=tb(s,0.7,1.4,9,1.3,"1. 理解艺术是表达  Art is expression",sz=15,c=DARK)
ap(tf,"2. 认识艺术形式：绘画、音乐、舞蹈……",sz=15,c=DARK)
ap(tf,"3. 发现生活中的艺术，判断什么是艺术",sz=15,c=DARK)
ap(tf,"4. 理解艺术可以表达心情和想法",sz=15,c=DARK)
tb(s,0.5,3.0,9,0.5,"🗣️ 语言目标  Language:",sz=20,b=True,c=SKY)
tb(s,0.7,3.5,9,0.4,"👀 我会认：开心  难过  生气  喜欢  心情",sz=15,b=True,c=DARK)
tb(s,0.7,4.0,9,0.4,"✍️ 我会写：开心  喜欢  生气",sz=15,b=True,c=DARK)
tb(s,0.5,4.65,9,0.5,"🖌️ 实践目标：完成 1 幅作品 (心情画 / 自画像)",sz=15,c=CORAL)
pn(s,n)

# ============================================================
# 4 SESSION 1 DIVIDER
# ============================================================
div("Session 1  上午","艺术是什么？  What is Art?\n🎨 绘画  🎵 音乐  💃 舞蹈  🎭 戏剧  🎬 电影",MAGENTA,"🖼️")
n+=1

# ============================================================
# 5 WHAT IS ART — big idea slide
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"💡 艺术是什么？  What is Art?",MAGENTA)
# Big centered callout
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(1.2),Inches(1.1),Inches(7.6),Inches(1.6))
sh.fill.solid();sh.fill.fore_color.rgb=MAGENTA;sh.line.fill.background()
tb(s,1.4,1.2,7.2,0.7,"艺术 = 表达",sz=44,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tb(s,1.4,1.95,7.2,0.7,"Art is Expression",sz=22,c=YELLOW,a=PP_ALIGN.CENTER)
# Three playful bubbles
bubbles=[
    ("🎨","表达看到的","Express what you SEE",YELLOW),
    ("💭","表达想到的","Express what you THINK",SKY),
    ("❤️","表达感受的","Express how you FEEL",CORAL),
]
for i,(em,cn,en,cl) in enumerate(bubbles):
    x=0.5+i*3.15
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(3.0),Inches(3.0),Inches(2.1))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=cl;sh.line.width=Pt(3)
    tb(s,x+0.1,3.1,2.8,0.7,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,3.9,2.8,0.45,cn,sz=18,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,4.4,2.8,0.35,en,sz=12,c=GRAY,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 6 ART FORMS — 6-grid
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🎨 艺术形式  Art Forms")
tb(s,0.4,0.9,9,0.4,"艺术有很多种，你认识哪些？",sz=14,c=GRAY,a=PP_ALIGN.CENTER)
forms=[
    ("🎨","绘画","Painting",MAGENTA),
    ("🎵","音乐","Music",SKY),
    ("💃","舞蹈","Dance",CORAL),
    ("🗿","雕塑","Sculpture",GREEN_L),
    ("🎭","戏剧","Drama",PURPLE),
    ("🎬","电影","Film",YELLOW),
]
for i,(em,cn,en,cl) in enumerate(forms):
    col=i%3;row=i//3
    x=0.3+col*3.2;y=1.4+row*1.95
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(3.0),Inches(1.75))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=cl;sh.line.width=Pt(3)
    tb(s,x+0.1,y+0.1,2.8,0.7,em,sz=40,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,y+0.95,2.8,0.45,cn,sz=22,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,y+1.4,2.8,0.3,en,sz=12,c=GRAY,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# PAINTING 绘画 — overview + 5 example slides
# ============================================================
s=type_overview_slide("🎨","绘画","Painting",MAGENTA,
    "绘画有很多种！看看你认识哪些？",
    [("🖼️","油画","Oil Painting"),("🖌️","水墨画","Chinese Ink"),
     ("🎨","水彩画","Watercolor"),("🖍️","蜡笔画","Crayon"),
     ("✏️","素描","Pencil Sketch"),("👶","儿童画","Kids' Art")]);n+=1;pn(s,n)

s=example_slide("🎨","油画","蒙娜丽莎","Mona Lisa",
    "达·芬奇 Leonardo da Vinci · 约 1503 年 · 法国卢浮宫",
    "她的微笑是世界上最有名的微笑！\nThe most famous smile in the world!",
    "📷 蒙娜丽莎 Mona Lisa",MAGENTA);n+=1;pn(s,n)

s=example_slide("🎨","油画","向日葵","Sunflowers",
    "梵高 Vincent van Gogh · 1888 年 · 荷兰",
    "亮黄色 + 粗笔触 = 像阳光一样温暖！\nBright yellow strokes = warm like sunshine!",
    "📷 梵高向日葵 Van Gogh Sunflowers",MAGENTA);n+=1;pn(s,n)

s=example_slide("🖌️","水墨画","虾","Shrimp",
    "齐白石 Qi Baishi · 中国 · 近代",
    "只用黑色墨水，虾就像活的一样游！\nJust black ink — shrimp look alive!",
    "📷 齐白石 虾 Qi Baishi Shrimp",MAGENTA);n+=1;pn(s,n)

s=example_slide_generic("🎨","水彩画","花园一角","A Corner of the Garden",
    "水彩画示例 Watercolor example",
    "水彩颜色透明，感觉很温柔。\nWatercolor is transparent & gentle.",
    "📷 水彩画 Watercolor",MAGENTA);n+=1;pn(s,n)

s=example_slide_generic("🖍️","蜡笔画 / 儿童画","我的家","My Family",
    "小朋友画的 A kid's drawing",
    "简单的线条和颜色，也是很棒的艺术！\nSimple lines + colors = great art too!",
    "📷 儿童蜡笔画 Kids' crayon art",MAGENTA);n+=1;pn(s,n)

# ============================================================
# MUSIC 音乐 — overview + 4 example slides
# ============================================================
s=type_overview_slide("🎵","音乐","Music",SKY,
    "音乐可以用不同的乐器 + 声音！",
    [("🎹","钢琴","Piano"),("🎻","小提琴","Violin"),
     ("🪕","古筝","Guzheng"),("🥁","鼓","Drums"),
     ("🎤","唱歌","Singing"),("🎺","小号","Trumpet")]);n+=1;pn(s,n)

s=example_slide_generic("🎹","钢琴 Piano","《小星星》","Twinkle Twinkle Little Star",
    "莫扎特改编 Mozart variations",
    "钢琴有 88 个键。白键 + 黑键 = 任何歌都能弹！\n88 keys = play any song!",
    "📷 钢琴 Piano",SKY);n+=1;pn(s,n)

s=example_slide_generic("🎻","小提琴 Violin","《梁祝》","Butterfly Lovers",
    "中国著名小提琴协奏曲",
    "小提琴只有 4 根弦，却能唱出故事！\n4 strings tell a whole love story!",
    "📷 小提琴 Violin",SKY);n+=1;pn(s,n)

s=example_slide_generic("🪕","古筝 Guzheng","中国古典音乐","Chinese Classical",
    "中国传统乐器 Traditional Chinese",
    "古筝有 21 根弦，声音像流水一样美。\n21 strings — sounds like flowing water.",
    "📷 古筝 Guzheng",SKY);n+=1;pn(s,n)

s=example_slide_generic("🥁","鼓 Drums","咚咚咚！","Boom Boom Boom!",
    "打击乐 Percussion",
    "鼓声又响又有力，像心跳一样！\nDrums = loud & strong, like a heartbeat!",
    "📷 鼓 Drums",SKY);n+=1;pn(s,n)

# ============================================================
# DANCE 舞蹈 — overview + 3 example slides
# ============================================================
s=type_overview_slide("💃","舞蹈","Dance",CORAL,
    "舞蹈用身体讲故事！",
    [("🩰","芭蕾","Ballet"),("🪭","中国舞","Chinese Dance"),
     ("🕺","街舞","Hip-hop"),("💫","现代舞","Modern"),
     ("🎎","民族舞","Folk"),("🎭","踢踏舞","Tap")]);n+=1;pn(s,n)

s=example_slide("🩰","芭蕾","天鹅湖","Swan Lake",
    "柴可夫斯基 Tchaikovsky · 1877 年",
    "跳舞的人像天鹅一样优雅！\nDancers move like graceful swans!",
    "📷 天鹅湖 Swan Lake ballet",CORAL);n+=1;pn(s,n)

s=example_slide_generic("🪭","中国古典舞","扇子舞","Fan Dance",
    "中国民族舞 Chinese folk",
    "扇子一开一合，像花一样漂亮！\nFans open & close like blooming flowers!",
    "📷 扇子舞 Fan Dance",CORAL);n+=1;pn(s,n)

s=example_slide_generic("🕺","街舞","Hip-hop","Street Dance",
    "现代都市舞蹈 Modern urban",
    "快节奏 + 酷动作 = 自由表达！\nFast beats + cool moves = freedom!",
    "📷 街舞 Hip-hop",CORAL);n+=1;pn(s,n)

# ============================================================
# SCULPTURE 雕塑 — overview + 2 example slides
# ============================================================
s=type_overview_slide("🗿","雕塑","Sculpture",GREEN_L,
    "雕塑是立体的艺术，用石头/木头/黏土做！",
    [("🗿","石雕","Stone"),("🪵","木雕","Wood"),
     ("🏺","陶塑","Clay"),("🧊","冰雕","Ice"),
     ("🎨","泥塑","Ceramic"),("🔩","金属","Metal")]);n+=1;pn(s,n)

s=example_slide("🗿","雕塑","大卫","David",
    "米开朗基罗 Michelangelo · 1504 年 · 意大利",
    "用大理石做的，5米多高！\nMarble, over 17 feet tall!",
    "📷 大卫雕像 Statue of David",GREEN_L);n+=1;pn(s,n)

s=example_slide_generic("🏺","雕塑","兵马俑","Terracotta Warriors",
    "中国秦朝 · 2000 多年前",
    "8000 多个泥塑士兵，每一个脸都不一样！\n8,000+ soldiers, each face is different!",
    "📷 兵马俑 Terracotta Warriors",GREEN_L);n+=1;pn(s,n)

# ============================================================
# DRAMA 戏剧 — overview + 2 example slides
# ============================================================
s=type_overview_slide("🎭","戏剧","Drama",PURPLE,
    "戏剧 = 演员 + 故事 + 舞台！",
    [("🎭","京剧","Peking Opera"),("🎼","音乐剧","Musical"),
     ("🎪","话剧","Spoken Drama"),("🎠","木偶戏","Puppet"),
     ("🎬","歌剧","Opera"),("🎠","皮影","Shadow Play")]);n+=1;pn(s,n)

s=example_slide_generic("🎭","京剧","《霸王别姬》","Farewell My Concubine",
    "中国传统戏曲 · 200 多年历史",
    "演员画脸谱，不同颜色 = 不同人物！\nPainted faces — each color = a character!",
    "📷 京剧脸谱 Peking Opera",PURPLE);n+=1;pn(s,n)

s=example_slide("🎼","音乐剧","狮子王","The Lion King",
    "百老汇音乐剧 · 1997 年首演",
    "唱歌 + 跳舞 + 演戏 = 音乐剧！\nSing + dance + act = musical!",
    "📷 狮子王音乐剧 Lion King",PURPLE);n+=1;pn(s,n)

# ============================================================
# FILM 电影 — overview + 2 example slides
# ============================================================
s=type_overview_slide("🎬","电影","Film",YELLOW,
    "电影用镜头讲故事！",
    [("🎞️","动画","Animation"),("🎬","真人","Live Action"),
     ("🦸","超级英雄","Superhero"),("🎠","儿童","Kids"),
     ("🎪","纪录片","Documentary"),("🎭","短片","Short Film")]);n+=1;pn(s,n)

s=example_slide("🎞️","动画电影","哪吒之魔童降世","Ne Zha",
    "中国动画 · 2019 年",
    '"我命由我不由天！" 中国动画很精彩！\nAmazing Chinese animation!',
    "📷 哪吒 Ne Zha",YELLOW);n+=1;pn(s,n)

s=example_slide("🎞️","动画电影","功夫熊猫","Kung Fu Panda",
    "梦工厂 DreamWorks · 2008 年",
    "熊猫也可以当功夫大师！\nEven a panda can be a kung fu master!",
    "📷 功夫熊猫 Kung Fu Panda",YELLOW);n+=1;pn(s,n)

# ============================================================
# 11 ART IN DAILY LIFE — Is this art?
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🔍 生活中的艺术  Art in Daily Life",YELLOW)
tb(s,0.4,0.9,9,0.4,"哪些是艺术？Point to the art you see every day!",sz=14,c=GRAY,a=PP_ALIGN.CENTER)
items=[
    ("👕","衣服图案","Clothes patterns",MAGENTA),
    ("🏛️","漂亮建筑","Buildings",SKY),
    ("🍱","摆盘艺术","Food plating",CORAL),
    ("📱","手机壁纸","Phone wallpaper",PURPLE),
    ("📚","绘本插图","Book drawings",GREEN_L),
    ("🎪","公园雕塑","Park sculpture",YELLOW),
]
for i,(em,cn,en,cl) in enumerate(items):
    col=i%3;row=i//3
    x=0.3+col*3.2;y=1.35+row*2.0
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(3.0),Inches(1.8))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=cl;sh.line.width=Pt(2.5)
    tb(s,x+0.1,y+0.1,2.8,0.7,em,sz=36,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,y+0.95,2.8,0.45,cn,sz=17,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,y+1.4,2.8,0.3,en,sz=11,c=GRAY,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 12 ART EXPRESSES EMOTIONS
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"❤️ 艺术表达心情  Art Expresses Emotions",E_MOOD)
tb(s,0.4,0.9,9,0.4,"不同颜色 / 不同音乐 = 不同心情",sz=15,c=GRAY,a=PP_ALIGN.CENTER)
mo=[
    ("😄","开心","HAPPY",E_HAPPY,"亮色 Bright colors\n快音乐 Fast music"),
    ("😢","难过","SAD",E_SAD,"蓝色 Blue tones\n慢音乐 Slow music"),
    ("😡","生气","ANGRY",E_ANGRY,"红色 Red\n大声音 Loud sound"),
    ("😍","喜欢","LOVE",E_LOVE,"粉色 Pink\n温柔音乐 Gentle music"),
]
for i,(em,cn,en,cl,d) in enumerate(mo):
    x=0.3+i*2.4
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.4),Inches(2.2),Inches(3.7))
    sh.fill.solid();sh.fill.fore_color.rgb=cl;sh.line.fill.background()
    # Dark text on the bright yellow (happy) card for contrast; white on the others
    txt_c = DARK if cl == E_HAPPY else WHITE
    tb(s,x+0.1,1.5,2.0,0.7,em,sz=46,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,2.3,2.0,0.4,cn,sz=24,b=True,c=txt_c,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,2.75,2.0,0.3,en,sz=12,c=txt_c,a=PP_ALIGN.CENTER)
    ls=d.split('\n')
    tf=tb(s,x+0.1,3.25,2.0,0.4,ls[0],sz=12,c=txt_c,a=PP_ALIGN.CENTER)
    for l in ls[1:]:ap(tf,l,sz=12,c=txt_c,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 13 IS THIS ART? — judgment game
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🤔 这是艺术吗？  Is This Art?",YELLOW)
tb(s,0.4,0.85,9,0.35,"看图，举手说「是」或「不是」！Show me yes 👍 or no 👎",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
# 4 cards with big placeholders
examples=[
    ("🖼️","蒙娜丽莎","Mona Lisa","✅ 是",GREEN_OK),
    ("🌳","一棵树","A tree","🤔 看情况",YELLOW),
    ("🎵","一首歌","A song","✅ 是",GREEN_OK),
    ("🗑️","一堆垃圾","Trash","❓ 讨论一下",CORAL),
]
for i,(em,cn,en,ans,cl) in enumerate(examples):
    col=i%2;row=i//2
    x=0.3+col*4.75;y=1.25+row*1.95
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(4.6),Inches(1.8))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=cl;sh.line.width=Pt(2.5)
    tb(s,x+0.15,y+0.15,1.0,0.9,em,sz=38,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+1.3,y+0.15,3.2,0.45,cn,sz=18,b=True,c=DARK)
    tb(s,x+1.3,y+0.6,3.2,0.35,en,sz=12,c=GRAY)
    sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x+1.3),Inches(y+1.05),Inches(3.1),Inches(0.55))
    sh2.fill.solid();sh2.fill.fore_color.rgb=cl;sh2.line.fill.background()
    tb(s,x+1.4,y+1.1,3.0,0.5,ans,sz=16,b=True,c=WHITE)
pn(s,n)

# ============================================================
# 14 SESSION 2 DIVIDER
# ============================================================
div("Session 2  下午","复习 + 语言目标\n🗣️ 我会认 5 个词  ·  ✍️ 我会写 3 个词",SKY,"📖")
n+=1

# ============================================================
# 15 Review mood match
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🔄 快速复习  Quick Review — 连心情",SKY)
tb(s,0.4,0.85,9,0.35,"把表情和心情连起来 (口头)",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
faces=[("😄",E_HAPPY),("😢",E_SAD),("😡",E_ANGRY),("😍",E_LOVE),("🎭",E_MOOD)]
words=["难过","喜欢","心情","开心","生气"]
for i,(em,cl) in enumerate(faces):
    y=1.25+i*0.75
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.8),Inches(y),Inches(3.4),Inches(0.65))
    sh.fill.solid();sh.fill.fore_color.rgb=cl;sh.line.fill.background()
    tb(s,0.95,y+0.1,3.2,0.5,em,sz=26,c=WHITE,a=PP_ALIGN.CENTER)
for i,w in enumerate(words):
    y=1.25+i*0.75
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.8),Inches(y),Inches(3.4),Inches(0.65))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=SKY;sh.line.width=Pt(2)
    tb(s,5.95,y+0.12,3.2,0.5,w,sz=22,b=True,c=SKY,a=PP_ALIGN.CENTER)
tb(s,4.25,3.1,1.5,0.5,"?",sz=44,b=True,c=CORAL,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 16-20 我会认 word cards
# ============================================================
read_words=[
    ("开心","kāi xīn","happy","今天我很开心！","📷 开心的脸",E_HAPPY),
    ("难过","nán guò","sad","他难过的时候会画蓝色。","📷 难过的脸",E_SAD),
    ("生气","shēng qì","angry","小朋友生气了，他画了红色。","📷 生气的脸",E_ANGRY),
    ("喜欢","xǐ huān","like","我喜欢画画！","📷 喜欢的脸",E_LOVE),
    ("心情","xīn qíng","mood","画画可以表达我的心情。","📷 各种心情",E_MOOD),
]
for w,py,en,sent,img,cl in read_words:
    s=word_card_read(w,py,en,sent,img,cl);n+=1;pn(s,n)

# ============================================================
# 21 WORD GAMES
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🎮 练一练  Word Games (选一个玩！)",SKY)
games=[
    ("1️⃣","拍苍蝇\nFly Swatter","老师说词\n学生拍正确的字卡",RGBColor(0xFF,0xF3,0xE0)),
    ("2️⃣","表情配词\nFace Match","看表情\n读对应的词",RGBColor(0xE3,0xF2,0xFD)),
    ("3️⃣","心情接龙\nMood Chain","一个接一个\n说心情词",RGBColor(0xE8,0xF5,0xE9)),
    ("4️⃣","我演你猜\nActing","做表情\n大家猜心情",RGBColor(0xFC,0xE4,0xEC)),
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
# 22-24 我会写 word cards
# ============================================================
write_words=[
    ("开心","kāi xīn","happy","📷 开心的脸",E_HAPPY),
    ("喜欢","xǐ huān","like","📷 爱心",E_LOVE),
    ("生气","shēng qì","angry","📷 生气的脸",E_ANGRY),
]
for w,py,en,img,cl in write_words:
    s=word_card_write(w,py,en,img,cl);n+=1;pn(s,n)

# ============================================================
# 25 SESSION 3 DIVIDER
# ============================================================
div("Session 3  下午","动手做艺术！  Make Art!\n🎨 我的心情画  ·  👤 我是谁 自画像",CORAL,"🖌️")
n+=1

# ============================================================
# 26 Booklet
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,'📓 完成"艺术是表达"练习册  Day 1 Booklet',CORAL)
ib(s,0.4,0.9,9.2,4.3,"📷 练习册截图 / Booklet pages")
pn(s,n)

# ============================================================
# 27 Projects overview — 2 levels
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🎨 今天做什么？  Two Project Levels",CORAL)
tb(s,0.4,0.9,9,0.4,"老师会根据你的年级选一个！Teacher will pick for your level.",sz=13,c=GRAY,a=PP_ALIGN.CENTER)
projects=[
    ("低 Level\n(K-2)","🎨 我的心情画","My Mood Painting","用颜色画出你今天的心情\nUse colors to show your mood",E_HAPPY,MAGENTA),
    ("高 Level\n(3-5)","👤 我是谁","Self-Portrait: Who Am I","画你自己 + 加喜欢的东西\nDraw yourself + things you love",E_MOOD,PURPLE),
]
for i,(lvl,nm,en,d,ltcl,cl) in enumerate(projects):
    x=0.5+i*4.6
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.4),Inches(4.3),Inches(3.75))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=cl;sh.line.width=Pt(4)
    # level label
    sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x+0.25),Inches(1.55),Inches(1.4),Inches(0.7))
    sh2.fill.solid();sh2.fill.fore_color.rgb=ltcl;sh2.line.fill.background()
    tf=tb(s,x+0.3,1.62,1.3,0.5,lvl.split('\n')[0],sz=13,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    ap(tf,lvl.split('\n')[1],sz=10,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,x+0.2,2.35,4.0,0.6,nm,sz=26,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.2,3.0,4.0,0.3,en,sz=13,c=GRAY,a=PP_ALIGN.CENTER)
    ib(s,x+0.5,3.4,3.4,1.2,"📷 作品示范")
    ls=d.split('\n')
    tf=tb(s,x+0.2,4.65,4.0,0.25,ls[0],sz=12,c=DARK,a=PP_ALIGN.CENTER)
    for l in ls[1:]:ap(tf,l,sz=11,c=GRAY,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 28 Project 低 Level — Mood Painting
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🎨 低 Level: 我的心情画  Mood Painting",MAGENTA)
# Left: materials
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(0.95),Inches(4.4),Inches(0.4))
sh.fill.solid();sh.fill.fore_color.rgb=MAGENTA;sh.line.fill.background()
tb(s,0.4,0.98,4.2,0.35,"🧺 材料  Materials",sz=14,b=True,c=WHITE)
tf=tb(s,0.4,1.45,4.4,1.8,"📄 白纸  White paper",sz=13,c=DARK)
ap(tf,"🖍️ 蜡笔 / 彩笔  Crayons / markers",sz=13,c=DARK)
ap(tf,"🎨 水彩 (可选)  Watercolor",sz=13,c=DARK)
ap(tf,"🖼️ 画板  Easel / board",sz=13,c=DARK)
# Middle: color → mood key
sh_k=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(3.35),Inches(4.4),Inches(0.4))
sh_k.fill.solid();sh_k.fill.fore_color.rgb=YELLOW;sh_k.line.fill.background()
tb(s,0.4,3.38,4.2,0.35,"🎨 颜色小贴士  Color Tips",sz=14,b=True,c=DARK)
tf_k=tb(s,0.4,3.85,4.4,1.4,"🟡 黄色 = 开心",sz=13,c=DARK)
ap(tf_k,"🔵 蓝色 = 难过",sz=13,c=DARK)
ap(tf_k,"🔴 红色 = 生气",sz=13,c=DARK)
ap(tf_k,"💗 粉色 = 喜欢",sz=13,c=DARK)
# Right: steps
sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(0.95),Inches(4.8),Inches(0.4))
sh2.fill.solid();sh2.fill.fore_color.rgb=SKY;sh2.line.fill.background()
tb(s,5.0,0.98,4.6,0.35,"👉 做法  Steps",sz=14,b=True,c=WHITE)
tf2=tb(s,5.0,1.45,4.7,1.8,"1️⃣ 想一想：今天我的心情是什么？",sz=13,c=DARK)
ap(tf2,"2️⃣ 选颜色：开心 / 难过 / 生气 / 喜欢",sz=13,c=DARK)
ap(tf2,"3️⃣ 用线条或形状画出心情",sz=13,c=DARK)
ap(tf2,"4️⃣ 可以是抽象的，没有对错！",sz=13,c=DARK)
# Sentence frames
sh3=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(3.35),Inches(4.8),Inches(0.4))
sh3.fill.solid();sh3.fill.fore_color.rgb=GREEN_OK;sh3.line.fill.background()
tb(s,5.0,3.38,4.6,0.35,"🗣️ 展示句型  Say These",sz=14,b=True,c=WHITE)
tf3=tb(s,5.0,3.85,4.7,1.4,"· 这是我的心情画。",sz=13,c=DARK)
ap(tf3,"· 我今天很开心 / 难过。",sz=13,c=DARK)
ap(tf3,"· 我喜欢黄色，因为……",sz=13,c=DARK)
pn(s,n)

# ============================================================
# 29 Project 高 Level — Self-portrait
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"👤 高 Level: 我是谁 自画像  Self-Portrait",PURPLE)
# Left: materials
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(0.95),Inches(4.4),Inches(0.4))
sh.fill.solid();sh.fill.fore_color.rgb=PURPLE;sh.line.fill.background()
tb(s,0.4,0.98,4.2,0.35,"🧺 材料  Materials",sz=14,b=True,c=WHITE)
tf=tb(s,0.4,1.45,4.4,1.8,"📄 白纸 / 画板",sz=13,c=DARK)
ap(tf,"✏️ 铅笔 + 彩笔",sz=13,c=DARK)
ap(tf,"🪞 小镜子 (看自己)",sz=13,c=DARK)
ap(tf,"💭 想一想：我喜欢什么？",sz=13,c=DARK)
# Middle: what to include
sh_k=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(3.35),Inches(4.4),Inches(0.4))
sh_k.fill.solid();sh_k.fill.fore_color.rgb=CORAL;sh_k.line.fill.background()
tb(s,0.4,3.38,4.2,0.35,"💡 可以画什么  Ideas",sz=14,b=True,c=WHITE)
tf_k=tb(s,0.4,3.85,4.4,1.4,"🎨 你喜欢的东西 (宠物、食物、玩具)",sz=13,c=DARK)
ap(tf_k,"🌈 你的心情颜色",sz=13,c=DARK)
ap(tf_k,"⭐ 你擅长什么 (跳舞、画画、运动)",sz=13,c=DARK)
# Right: steps
sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(0.95),Inches(4.8),Inches(0.4))
sh2.fill.solid();sh2.fill.fore_color.rgb=MAGENTA;sh2.line.fill.background()
tb(s,5.0,0.98,4.6,0.35,"👉 做法  Steps",sz=14,b=True,c=WHITE)
tf2=tb(s,5.0,1.45,4.7,1.8,"1️⃣ 先用铅笔画出自己的脸",sz=13,c=DARK)
ap(tf2,"2️⃣ 加头发、衣服、眼睛",sz=13,c=DARK)
ap(tf2,"3️⃣ 旁边画你喜欢的东西",sz=13,c=DARK)
ap(tf2,"4️⃣ 用颜色表达你的心情",sz=13,c=DARK)
# Sentence frames
sh3=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(3.35),Inches(4.8),Inches(0.4))
sh3.fill.solid();sh3.fill.fore_color.rgb=GREEN_OK;sh3.line.fill.background()
tb(s,5.0,3.38,4.6,0.35,"🗣️ 展示句型  Say These",sz=14,b=True,c=WHITE)
tf3=tb(s,5.0,3.85,4.7,1.4,"· 这是我！",sz=13,c=DARK)
ap(tf3,"· 我喜欢……",sz=13,c=DARK)
ap(tf3,"· 我画了……因为……",sz=13,c=DARK)
pn(s,n)

# ============================================================
# 30 Share & Gallery
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🖼️ 小小画展  Gallery Walk",YELLOW)
tb(s,0.4,0.9,9,0.5,"把作品贴在墙上，大家一起看！",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.4,9,0.4,"Hang your art and walk around like a museum!",sz=13,c=GRAY,a=PP_ALIGN.CENTER)
steps=[
    ("1️⃣","🖼️","贴作品\nHang it up"),
    ("2️⃣","👀","看同学的\nLook around"),
    ("3️⃣","🗣️","说一说\nTalk about it"),
    ("4️⃣","⭐","贴星星\nGive stars"),
]
for i,(num,em,d) in enumerate(steps):
    x=0.3+i*2.4
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(2.0),Inches(2.2),Inches(2.9))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=[MAGENTA,SKY,CORAL,YELLOW][i];sh.line.width=Pt(3)
    tb(s,x+0.1,2.1,2.0,0.5,num,sz=22,b=True,c=[MAGENTA,SKY,CORAL,YELLOW][i],a=PP_ALIGN.CENTER)
    tb(s,x+0.1,2.6,2.0,0.8,em,sz=40,a=PP_ALIGN.CENTER)
    ls=d.split('\n')
    tf=tb(s,x+0.1,3.6,2.0,0.4,ls[0],sz=14,b=True,c=DARK,a=PP_ALIGN.CENTER)
    for l in ls[1:]:ap(tf,l,sz=12,c=GRAY,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 31 Day 1 Artist Badge
# ============================================================
s=ns();n+=1;bg(s,CREAM)
tb(s,0.5,0.4,9,0.8,"🎖️ Day 1 小艺术家徽章  Artist Badge",sz=26,b=True,c=MAGENTA,a=PP_ALIGN.CENTER)
# Big palette-shaped badge: use circle with 6 small color dots
sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.5),Inches(1.4),Inches(3),Inches(3))
sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=MAGENTA;sh.line.width=Pt(5)
tf=tb(s,3.6,1.65,2.8,2.7,"DAY 1",sz=18,b=True,c=CORAL,a=PP_ALIGN.CENTER)
ap(tf,"🎨",sz=40,a=PP_ALIGN.CENTER)
ap(tf,"艺术是表达",sz=20,b=True,c=MAGENTA,a=PP_ALIGN.CENTER)
ap(tf,"✓ COMPLETED",sz=13,b=True,c=GREEN_OK,a=PP_ALIGN.CENTER)
# paint dots around
for x,y,cl in [(2.5,2.2,YELLOW),(7.0,2.2,SKY),(2.5,3.7,CORAL),(7.0,3.7,PURPLE),(2.0,3.0,E_LOVE),(7.5,3.0,GREEN_L)]:
    dot(s,x,y,0.35,cl)
tb(s,1,4.55,8,0.4,"恭喜你完成 Day 1！You are an artist! 🎉",sz=16,b=True,c=MAGENTA,a=PP_ALIGN.CENTER)
tb(s,1,5.0,8,0.4,"认识了 6 种艺术形式 · 5 个心情词 · 1 幅作品",sz=13,c=GRAY,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 32 Tomorrow preview
# ============================================================
s=ns();n+=1;bg(s,MAGENTA)
tb(s,0.5,0.9,9,0.8,"🎨 明天见！  See You Tomorrow!",sz=32,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tf=tb(s,1.5,2.2,7,2.5,"Day 2 — 颜色的秘密",sz=28,b=True,c=YELLOW,a=PP_ALIGN.CENTER)
ap(tf,"Secret of Colors",sz=16,c=WHITE,a=PP_ALIGN.CENTER)
ap(tf,"",sz=10)
ap(tf,"🌈 红 + 黄 = 什么色？",sz=20,b=True,c=WHITE,a=PP_ALIGN.CENTER)
ap(tf,"What happens when we mix colors?",sz=14,c=WHITE,a=PP_ALIGN.CENTER)
ap(tf,"",sz=10)
ap(tf,"明天见，小艺术家！",sz=15,b=True,c=YELLOW,a=PP_ALIGN.CENTER)
pn(s,n)

OUT='/Users/Huan/projects/summercourse/Chinese/小小艺术家little_artist_pbl/day1_art_is_expression.pptx'
prs.save(OUT);print(f"Created {n} slides → {OUT}")
