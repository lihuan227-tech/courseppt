#!/usr/bin/env python3
"""
野外生存与探险 Wilderness Survival Unit — Day 1: 认识自然与安全规则
Structure modeled on 世界旅行 Day 1 Asia v2, adapted for 6 environments.
Palette: Adventure (pine green + sunset orange) — distinct from world-trip teal.
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

# --- Palette: Adventure / Wilderness ---
PINE = RGBColor(0x1E,0x4D,0x2B)   # primary: deep pine green
SUN = RGBColor(0xE0,0x7A,0x2C)    # accent: sunset orange
CREAM = RGBColor(0xFD,0xF6,0xE3)  # background cream
BROWN = RGBColor(0x6B,0x44,0x23)  # soil brown
SKY = RGBColor(0x4A,0x90,0xD9)    # sky blue
SUNYEL = RGBColor(0xF5,0xC2,0x42) # sunshine yellow
ALERT = RGBColor(0xD0,0x4A,0x3C)  # warning red
WHITE = RGBColor(0xFF,0xFF,0xFF)
DARK = RGBColor(0x2C,0x2C,0x2C)
GRAY = RGBColor(0x88,0x88,0x88)
LGRAY = RGBColor(0xBB,0xBB,0xBB)
WARM = RGBColor(0xFF,0xF3,0xE0)
IMGBG = RGBColor(0xE8,0xE8,0xE8)
GREEN_OK = RGBColor(0x38,0x8E,0x3C)

# Per-environment colors
FOREST = RGBColor(0x2D,0x5A,0x3D)
MOUNTAIN = RGBColor(0x54,0x6E,0x7A)
GRASS = RGBColor(0x7C,0xB3,0x42)
RIVER = RGBColor(0x19,0x76,0xD2)
DESERT = RGBColor(0xD4,0xA5,0x74)
SNOW = RGBColor(0x78,0xA7,0xB5)

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
def div(title,sub,color,emoji=""):
    s=ns();bg(s,color);tb(s,1,1.5,8,1.2,f"{emoji} {title}",sz=42,b=True,c=WHITE,a=PP_ALIGN.CENTER);tb(s,1,2.8,8,0.8,sub,sz=22,c=WHITE,a=PP_ALIGN.CENTER);return s
def vs(title,bgc):
    s=ns();bg(s,bgc);tb(s,1,0.8,8,0.8,"🎬 看视频  Watch Video",sz=36,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    tb(s,1,1.8,8,0.5,title,sz=22,c=WARM,a=PP_ALIGN.CENTER)
    ib(s,1.5,2.5,7,2.5,"📷 插入视频截图或粘贴视频链接");tb(s,1,5.1,8,0.3,"🔗 视频链接: ____________________",sz=14,c=LGRAY,a=PP_ALIGN.CENTER)
    return s

def aspect_slide(emoji,cn,en,env_color,aspect_label,aspect_color,questions,frame):
    """One slide for ONE aspect of ONE environment, inquiry-based (questions, no answers)."""
    s=ns();bg(s,CREAM)
    # Header bar (env color) — environment identity on left, aspect label on right
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(0.15),Inches(9.4),Inches(0.7))
    sh.fill.solid();sh.fill.fore_color.rgb=env_color;sh.line.fill.background()
    tb(s,0.5,0.22,1.0,0.55,emoji,sz=28,c=WHITE)
    tb(s,1.5,0.23,3.0,0.55,cn,sz=26,b=True,c=WHITE)
    tb(s,1.5,0.62,3.0,0.25,en,sz=11,c=WARM)
    # Aspect pill on the right of header
    pill=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.8),Inches(0.27),Inches(4.8),Inches(0.45))
    pill.fill.solid();pill.fill.fore_color.rgb=aspect_color;pill.line.color.rgb=WHITE;pill.line.width=Pt(1.5)
    tb(s,4.9,0.32,4.6,0.4,aspect_label,sz=15,b=True,c=WHITE,a=PP_ALIGN.CENTER)
    # Image placeholder (left)
    ib(s,0.3,1.05,4.3,3.3,f"📷 {cn} 图片 / 视频")
    # Inquiry panel (right) — questions, no answers
    panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.85),Inches(1.05),Inches(4.85),Inches(3.3))
    panel.fill.solid();panel.fill.fore_color.rgb=WHITE
    panel.line.color.rgb=aspect_color;panel.line.width=Pt(2.5)
    # Mini header inside the panel
    head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.85),Inches(1.05),Inches(4.85),Inches(0.5))
    head.fill.solid();head.fill.fore_color.rgb=aspect_color;head.line.fill.background()
    tb(s,5.0,1.13,4.6,0.4,"🤔 一起想一想  Let's Think Together",sz=14,b=True,c=WHITE)
    # Questions stacked with breathing room
    tf=tb(s,5.05,1.7,4.55,0.5,f"❓ {questions[0]}",sz=14,c=DARK)
    for q in questions[1:]:
        ap(tf,"",sz=8)
        ap(tf,f"❓ {q}",sz=14,c=DARK)
    # Sentence frame at the bottom (still inquiry — student fills the blanks)
    sf=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(4.5),Inches(9.4),Inches(0.65))
    sf.fill.solid();sf.fill.fore_color.rgb=WARM;sf.line.color.rgb=SUN;sf.line.width=Pt(2)
    tb(s,0.5,4.6,1.7,0.4,"💬 我来说",sz=14,b=True,c=SUN)
    tb(s,2.0,4.6,7.6,0.4,frame,sz=14,c=DARK)
    return s

def word_card_read(w,py,en,sent,img):
    """我会认 card — large character + pinyin + sentence."""
    s=ns();bg(s,CREAM);hb(s,"👀 我会认  I Can Read",SUN)
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.0),Inches(4.5),Inches(2.5))
    sh.fill.solid();sh.fill.fore_color.rgb=WARM;sh.line.fill.background()
    tb(s,0.5,1.1,4.3,1.4,w,sz=72,b=True,c=PINE,a=PP_ALIGN.CENTER)
    tb(s,0.5,2.4,4.3,0.4,f"{py}  {en}",sz=20,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,0.5,2.85,4.3,0.4,"👉 跟我读！Read after me!",sz=14,c=SUN,a=PP_ALIGN.CENTER)
    ib(s,5.3,1.0,4.4,2.5,img)
    sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(3.8),Inches(9.2),Inches(1.2))
    sh2.fill.solid();sh2.fill.fore_color.rgb=WHITE;sh2.line.color.rgb=SUN;sh2.line.width=Pt(2)
    tb(s,0.6,3.9,1.5,0.4,"例句",sz=16,b=True,c=SUN)
    tb(s,0.6,4.3,8.8,0.5,sent,sz=22,b=True,c=DARK)
    return s

def word_card_write(w,py,en,img):
    s=ns();bg(s,CREAM);hb(s,"✍️ 我会写  I Can Write",PINE)
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.0),Inches(4.5),Inches(2.0))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=PINE;sh.line.width=Pt(3)
    tb(s,0.5,1.05,4.3,1.2,w,sz=72,b=True,c=PINE,a=PP_ALIGN.CENTER)
    tb(s,0.5,2.2,4.3,0.4,f"{py}  {en}",sz=20,c=GRAY,a=PP_ALIGN.CENTER)
    ib(s,5.3,1.0,4.4,2.0,img)
    sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(3.3),Inches(5.0),Inches(1.8))
    sh2.fill.solid();sh2.fill.fore_color.rgb=WARM;sh2.line.fill.background()
    tb(s,0.6,3.4,4.6,0.4,"📝 笔顺 Stroke Order",sz=16,b=True,c=PINE)
    ib(s,0.6,3.9,4.6,1.0,"📷 插入笔顺图片")
    tf=tb(s,5.8,3.4,3.8,0.4,"练习步骤 Practice:",sz=14,b=True,c=PINE)
    ap(tf,"1. 空中写 Air Write",sz=13,c=DARK)
    ap(tf,"2. 手心写 Palm Write",sz=13,c=DARK)
    ap(tf,"3. 纸上写 3 times",sz=13,c=DARK)
    return s

n=0

# ============================================================
# 1 COVER — Explorer Badge
# ============================================================
s=ns();n+=1;bg(s,CREAM)
tb(s,1,0.25,8,0.7,"Wilderness Adventure Camp",sz=32,b=True,c=PINE,a=PP_ALIGN.CENTER)
tb(s,1,0.85,8,0.45,"野外生存与探险夏令营",sz=20,c=PINE,a=PP_ALIGN.CENTER)
# Big round badge
sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.25),Inches(1.5),Inches(3.5),Inches(3.5))
sh.fill.solid();sh.fill.fore_color.rgb=PINE;sh.line.color.rgb=SUN;sh.line.width=Pt(6)
# Inner badge ring
sh2=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.55),Inches(1.8),Inches(2.9),Inches(2.9))
sh2.fill.solid();sh2.fill.fore_color.rgb=WHITE;sh2.line.color.rgb=SUN;sh2.line.width=Pt(2)
tf=tb(s,3.6,2.0,2.8,0.4,"DAY 1",sz=16,b=True,c=SUN,a=PP_ALIGN.CENTER)
ap(tf,"🏕️",sz=60,a=PP_ALIGN.CENTER)
ap(tf,"认识自然",sz=20,b=True,c=PINE,a=PP_ALIGN.CENTER)
ap(tf,"NATURE & SAFETY",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,1,5.05,8,0.4,"🎒 准备好背包，我们出发！Let's go, explorers!",sz=14,b=True,c=SUN,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 2 SCHEDULE
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"⏰ 今日时间安排  Today's Schedule")
for i,(nm,tm,dc,cl) in enumerate([
    ("Session 1  上午","11:00-11:45","认识6种自然环境 + 户外安全规则",PINE),
    ("Session 2  下午","2:00-2:45","复习总结 + 语言目标 (认字写字)",SUN),
    ("Session 3  下午","3:00-4:30","写Booklet + Project 项目活动",BROWN)]):
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
tb(s,0.5,0.9,9,0.5,"🌲 内容目标  Content:",sz=20,b=True,c=PINE)
tf=tb(s,0.7,1.4,9,1.3,"1. 了解 6 种常见自然环境：森林、山地、草地、河边、沙漠、雪地",sz=15,c=DARK)
ap(tf,"2. 了解不同自然环境的基本特点",sz=15,c=DARK)
ap(tf,"3. 了解户外可能遇到的危险与基本安全规则",sz=15,c=DARK)
ap(tf,"4. 建立基本的户外安全意识",sz=15,c=DARK)
tb(s,0.5,3.0,9,0.5,"🗣️ 语言目标  Language:",sz=20,b=True,c=SUN)
tb(s,0.7,3.5,4.4,0.9,"👀 我会认：森林 山地 草地\n　　　　　河边 沙漠 雪地",sz=15,b=True,c=DARK)
tb(s,5.3,3.5,4.3,0.9,"✍️ 我会写：森林 山地 河边",sz=15,b=True,c=DARK)
tb(s,0.5,4.55,9,0.5,"🎨 实践目标：完成 Booklet + 2 个项目 + 1 个游戏",sz=15,c=BROWN)
pn(s,n)

# ============================================================
# 4 SESSION 1 DIVIDER
# ============================================================
div("Session 1  上午","认识 6 种自然环境 + 安全规则\n🌲 森林  🏔️ 山地  🌾 草地  🏞️ 河边  🏜️ 沙漠  ❄️ 雪地",PINE,"🧭")
n+=1

# ============================================================
# 5 OVERVIEW — 6 environments preview
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🌍 大自然里有什么？  What's in Nature?")
tb(s,0.4,0.9,9,0.4,"今天我们一起认识 6 种常见自然环境！",sz=14,c=GRAY,a=PP_ALIGN.CENTER)
envs=[
    ("🌲","森林","Forest",FOREST),
    ("🏔️","山地","Mountain",MOUNTAIN),
    ("🌾","草地","Grassland",GRASS),
    ("🏞️","河边","Riverside",RIVER),
    ("🏜️","沙漠","Desert",DESERT),
    ("❄️","雪地","Snow",SNOW),
]
for i,(em,cn,en,cl) in enumerate(envs):
    col=i%3;row=i//3
    x=0.3+col*3.2;y=1.45+row*1.95
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(3.0),Inches(1.75))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=cl;sh.line.width=Pt(3)
    tb(s,x+0.1,y+0.1,2.8,0.7,em,sz=40,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,y+0.95,2.8,0.45,cn,sz=22,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,y+1.4,2.8,0.3,en,sz=12,c=GRAY,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 6-29  Inquiry slides — 6 environments × 4 aspects = 24 slides
#       Each aspect gets its own slide. Students discover answers
#       through questions (inquiry-based), not through statements.
# ============================================================
ASPECTS=[
    ("👀 长什么样  Looks Like", FOREST,   "我看到 ____ 和 ____。这里 ____。"),
    ("✨ 特点  Features",       SUN,      "____ 里有 ____。这里的 ____ 让我 ____。"),
    ("⚠️ 小心危险  Dangers",    ALERT,    "在这里要小心 ____ 和 ____。"),
    ("✅ 安全规则  Safety Rules",GREEN_OK, "我应该 ____。我不能 ____。"),
]

inquiry_data=[
    ("🌲","森林","Forest",FOREST,[
        # Looks Like
        ["你看到了什么？树多还是少？",
         "阳光照到地上多吗？为什么？",
         "地上有什么？你能听到什么声音？"],
        # Features
        ["森林里可能住着哪些动物？",
         "这里的空气闻起来怎么样？",
         "你能在森林里找到吃的吗？什么样的？"],
        # Dangers
        ["一个人在森林里，可能会发生什么？",
         "看到不认识的蘑菇，能不能吃？",
         "听到树丛「沙沙」响，可能是什么？"],
        # Safety Rules
        ["怎么做才不会迷路？",
         "靠近野生动物，要安静还是要发出声音？",
         "不认识的果子和蘑菇，能尝一口吗？"],
    ]),
    ("🏔️","山地","Mountain",MOUNTAIN,[
        ["山是什么样子的？高还是矮？",
         "山路平不平？容易走吗？",
         "山顶上能看到什么？"],
        ["站在山顶，你能看到多远？",
         "山上和山下，哪里更冷？为什么？",
         "山上的风大吗？你怎么知道？"],
        ["在山上最容易怎样受伤？",
         "天气突然变冷下雨，会怎样？",
         "上面的石头掉下来会怎样？"],
        ["爬山要穿什么样的鞋？为什么？",
         "出发前要看什么？(提示：天上)",
         "下山时手要扶着什么？"],
    ]),
    ("🌾","草地","Grassland",GRASS,[
        ["草地是什么颜色的？",
         "草地是平的还是有起伏？",
         "除了草，还能看到什么？"],
        ["太阳能晒到草地吗？亮不亮？",
         "你想在草地上玩什么？",
         "草丛里可能有哪些小动物？"],
        ["一直晒太阳，皮肤会怎样？",
         "草丛里可能藏着什么让你痒？",
         "看到一只蜜蜂飞过来，能挥手赶它吗？"],
        ["头上要戴什么？",
         "皮肤上要涂什么？",
         "看到蜜蜂应该怎么办？"],
    ]),
    ("🏞️","河边","Riverside",RIVER,[
        ["河里的水在动还是不动？",
         "水深还是浅？你怎么判断？",
         "岸边是泥、沙还是石头？"],
        ["用手摸一摸，河水冷还是热？",
         "你觉得水里可能有哪些小动物？",
         "河边为什么常常长柳树和芦苇？"],
        ["河边最危险的是什么？",
         "水看起来很浅，真的就浅吗？",
         "脚下湿湿的石头，走上去会怎样？"],
        ["谁必须一直在你身边？",
         "没有大人允许，可以下水吗？",
         "穿什么样的鞋才不会滑？"],
    ]),
    ("🏜️","沙漠","Desert",DESERT,[
        ["沙漠里到处是什么？",
         "树多吗？为什么？",
         "白天和晚上温度一样吗？"],
        ["什么样的植物能在沙漠里活下来？",
         "你能想到哪些住在沙漠里的动物？",
         "沙丘是怎么形成的？"],
        ["在沙漠里最怕发生什么？",
         "为什么沙漠里很容易迷路？",
         "沙尘暴来了会怎样？"],
        ["一天大约要喝多少水？",
         "穿什么样的衣服比较好？长袖还是短袖？",
         "能不能一个人离开队伍？"],
    ]),
    ("❄️","雪地","Snow",SNOW,[
        ["雪地是什么颜色？",
         "树枝上有什么？",
         "走在雪上会发出什么声音？"],
        ["在雪地里可以玩什么游戏？",
         "这里的空气暖和还是寒冷？",
         "摸一摸雪，是软的还是硬的？"],
        ["太冷了，手和脚会有什么感觉？",
         "雪反光太亮，眼睛会怎样？",
         "看起来是雪，下面会不会是水？"],
        ["应该穿什么样的衣服和鞋？",
         "眼睛要戴什么来挡光？",
         "湖面、河面上能走吗？为什么？"],
    ]),
]

for em,cn,en_name,env_color,aspect_questions in inquiry_data:
    for (aspect_label,aspect_color,frame),questions in zip(ASPECTS,aspect_questions):
        s=aspect_slide(em,cn,en_name,env_color,aspect_label,aspect_color,questions,frame)
        n+=1;pn(s,n)

# ============================================================
# 12 COMPARISON TABLE
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🗂️ 6 种环境对比  Compare 6 Environments",SUN)
tb(s,0.4,0.85,9,0.3,"每种环境都不一样，你能记住它们的危险吗？",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
ts=s.shapes.add_table(4,7,Inches(0.3),Inches(1.2),Inches(9.4),Inches(3.9));t=ts.table
t.columns[0].width=Inches(1.4)
for i in range(1,7):t.columns[i].width=Inches(1.333)
rows=[
    ["","🌲 森林","🏔️ 山地","🌾 草地","🏞️ 河边","🏜️ 沙漠","❄️ 雪地"],
    ["✨ 特点","树多\n阴凉","陡高\n风大","平坦\n有花","有水\n有鱼","热/沙\n干燥","冷白\n有雪"],
    ["⚠️ 危险","迷路\n动物","滑倒\n天气","晒\n虫子","溺水","中暑\n迷路","冻伤\n薄冰"],
    ["✅ 规则","跟紧老师","穿好鞋","戴帽子","远离水","多喝水","穿厚衣"],
]
for r,rd in enumerate(rows):
    for c,ct in enumerate(rd):
        cl=t.cell(r,c);cl.text="";tf=cl.text_frame;tf.word_wrap=True
        p=tf.paragraphs[0];p.alignment=PP_ALIGN.CENTER
        rn=p.add_run();rn.text=ct.split('\n')[0];rn.font.name='KaiTi'
        rn.font.size=Pt(13 if r==0 else 11);rn.font.bold=(r==0 or c==0)
        for line in ct.split('\n')[1:]:
            p2=tf.add_paragraph();p2.alignment=PP_ALIGN.CENTER
            rn2=p2.add_run();rn2.text=line;rn2.font.name='KaiTi';rn2.font.size=Pt(11);rn2.font.color.rgb=DARK
        if r==0:
            rn.font.color.rgb=WHITE;cl.fill.solid();cl.fill.fore_color.rgb=PINE
        elif c==0:
            rn.font.color.rgb=DARK;cl.fill.solid();cl.fill.fore_color.rgb=WARM
        else:
            rn.font.color.rgb=DARK
            if r%2==0:cl.fill.solid();cl.fill.fore_color.rgb=RGBColor(0xF5,0xF5,0xF5)
pn(s,n)

# ============================================================
# 13 SESSION 2 DIVIDER
# ============================================================
div("Session 2  下午","复习 + 语言目标 (认字 + 写字)\n我会认 6 个词  ·  我会写 3 个词",SUN,"📖")
n+=1

# ============================================================
# 14 Quick review — danger match
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🔄 快速复习  Quick Review — 找危险",SUN)
tb(s,0.4,0.85,9,0.3,"把环境和它最大的危险连起来 (口头)  Match the environment with its biggest danger!",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
envs_q=[("🌲","森林",FOREST),("🏔️","山地",MOUNTAIN),("🌾","草地",GRASS),("🏞️","河边",RIVER),("🏜️","沙漠",DESERT),("❄️","雪地",SNOW)]
dangers_q=["🥵 中暑","🌊 溺水","🦟 虫咬","🪨 滑倒","🥶 冻伤","🐻 动物"]
# Left col: environments
for i,(em,cn,cl) in enumerate(envs_q):
    y=1.3+i*0.65
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.8),Inches(y),Inches(3.4),Inches(0.55))
    sh.fill.solid();sh.fill.fore_color.rgb=cl;sh.line.fill.background()
    tb(s,0.95,y+0.08,3.2,0.4,f"{em} {cn}",sz=16,b=True,c=WHITE)
# Right col: dangers (shuffled order)
for i,d in enumerate(dangers_q):
    y=1.3+i*0.65
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.8),Inches(y),Inches(3.4),Inches(0.55))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=ALERT;sh.line.width=Pt(2)
    tb(s,5.95,y+0.08,3.2,0.4,d,sz=15,b=True,c=ALERT)
tb(s,4.25,2.9,1.5,0.4,"?",sz=40,b=True,c=SUN,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 15-20 我会认 word cards
# ============================================================
read_words=[
    ("森林","sēn lín","forest","森林里有很多大树。","📷 森林"),
    ("山地","shān dì","mountain","爬山要小心滑倒。","📷 山地"),
    ("草地","cǎo dì","grassland","草地上有很多花和虫。","📷 草地"),
    ("河边","hé biān","riverside","河边很危险，不能独自去。","📷 河边"),
    ("沙漠","shā mò","desert","沙漠里很热，要多喝水。","📷 沙漠"),
    ("雪地","xuě dì","snow","雪地里很冷，要穿厚衣服。","📷 雪地"),
]
for w,py,en,sent,img in read_words:
    s=word_card_read(w,py,en,sent,img);n+=1;pn(s,n)

# ============================================================
# 21 WORD GAMES
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🎮 练一练  Word Games (选一个玩！)",SUN)
games=[
    ("1️⃣","拍苍蝇\nFly Swatter","把字卡贴在\n白板上，老师\n说词语，学生拍！",WARM),
    ("2️⃣","举牌游戏\nShow Me","每人 6 张字卡\n老师说词语\n举正确的卡",RGBColor(0xFF,0xF3,0xE0)),
    ("3️⃣","抢椅子\nMusical Chairs","椅子上放字卡\n音乐停，读出词",RGBColor(0xE8,0xF5,0xE9)),
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
# 22-24 我会写 cards
# ============================================================
write_words=[
    ("森林","sēn lín","forest","📷 森林"),
    ("山地","shān dì","mountain","📷 山地"),
    ("河边","hé biān","riverside","📷 河边"),
]
for w,py,en,img in write_words:
    s=word_card_write(w,py,en,img);n+=1;pn(s,n)

# ============================================================
# 25 SESSION 3 DIVIDER
# ============================================================
div("Session 3  下午","写 Booklet + 动手做项目\n🖼️ 自然拼贴画  ·  🚧 安全标志  ·  🎭 我演你猜",BROWN,"🎒")
n+=1

# ============================================================
# 26 Booklet
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,'📓 完成"认识自然"练习册  Day 1 Booklet',BROWN)
ib(s,0.4,0.9,9.2,4.3,"📷 练习册截图 / Booklet pages")
pn(s,n)

# ============================================================
# 27 Projects overview
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🎨 动手时间！  Hands-On Time — 3 个活动",BROWN)
projects=[
    ("PROJECT 1","🖼️ 自然拼贴画","Nature Texture Collage","用树叶、石子、纸巾\n拼一幅自然环境",WARM,PINE),
    ("PROJECT 2","🚧 安全标志设计","Safety Sign Design","用彩笔和形状纸\n设计一个安全标志",RGBColor(0xFF,0xE0,0xB2),SUN),
    ("ACTIVITY","🎭 我演你猜","Charades","抽环境卡\n只用动作表演！",RGBColor(0xDC,0xED,0xC8),GREEN_OK),
]
for i,(lbl,nm,en,d,bgc,cl) in enumerate(projects):
    x=0.3+i*3.2
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(0.95),Inches(3.1),Inches(4.15))
    sh.fill.solid();sh.fill.fore_color.rgb=bgc;sh.line.color.rgb=cl;sh.line.width=Pt(2)
    tb(s,x+0.1,1.05,2.9,0.35,lbl,sz=12,b=True,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,1.4,2.9,0.6,nm,sz=20,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,2.0,2.9,0.35,en,sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    ib(s,x+0.2,2.45,2.8,1.6,"📷 示范")
    ls=d.split('\n')
    tf=tb(s,x+0.15,4.15,2.85,0.5,ls[0],sz=12,c=DARK,a=PP_ALIGN.CENTER)
    for ln in ls[1:]:ap(tf,ln,sz=12,c=DARK,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 28 Project 1 — Nature Texture Collage
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🖼️ Project 1: 自然拼贴画  Nature Collage",PINE)
# Left: materials
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(0.95),Inches(4.4),Inches(0.4))
sh.fill.solid();sh.fill.fore_color.rgb=PINE;sh.line.fill.background()
tb(s,0.4,0.98,4.2,0.35,"🧺 材料  Materials",sz=14,b=True,c=WHITE)
tf=tb(s,0.4,1.45,4.4,2.3,"🍂 干树叶  Dry leaves",sz=13,c=DARK)
ap(tf,"🪵 小树枝  Small twigs",sz=13,c=DARK)
ap(tf,"🪨 石子  Pebbles",sz=13,c=DARK)
ap(tf,"🧻 纸巾 (揉成雪)  Tissue (snow)",sz=13,c=DARK)
ap(tf,"🌊 铝箔纸 (代表河流)  Foil (river)",sz=13,c=DARK)
ap(tf,"💧 胶水  Glue",sz=13,c=DARK)
# Right: steps
sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(0.95),Inches(4.8),Inches(0.4))
sh2.fill.solid();sh2.fill.fore_color.rgb=SUN;sh2.line.fill.background()
tb(s,5.0,0.98,4.6,0.35,"👉 做法  Steps",sz=14,b=True,c=WHITE)
tf2=tb(s,5.0,1.45,4.7,2.3,"1️⃣ 选一个自然环境 (森林/河边/沙漠/雪地)",sz=13,c=DARK)
ap(tf2,"2️⃣ 用不同材料拼贴出这个环境",sz=13,c=DARK)
ap(tf2,"3️⃣ 贴牢、晾一下",sz=13,c=DARK)
ap(tf2,"4️⃣ 向同学介绍自己的作品",sz=13,c=DARK)
# Bottom: sentence frames
sh3=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(3.9),Inches(9.4),Inches(1.25))
sh3.fill.solid();sh3.fill.fore_color.rgb=WARM;sh3.line.color.rgb=PINE;sh3.line.width=Pt(2)
tb(s,0.5,4.0,9,0.35,"🗣️ 展示句型  Say These:",sz=14,b=True,c=PINE)
tb(s,0.5,4.4,4.5,0.35,"· 这是森林。",sz=14,c=DARK)
tb(s,0.5,4.7,4.5,0.35,"· 这是河边。",sz=14,c=DARK)
tb(s,5.2,4.4,4.5,0.35,"· 这里有树叶、石头和小河。",sz=14,c=DARK)
pn(s,n)

# ============================================================
# 29 Project 2 — Safety Sign
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🚧 Project 2: 安全标志设计  Safety Sign",SUN)
# Left: materials
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(0.95),Inches(4.4),Inches(0.4))
sh.fill.solid();sh.fill.fore_color.rgb=PINE;sh.line.fill.background()
tb(s,0.4,0.98,4.2,0.35,"🧺 材料  Materials",sz=14,b=True,c=WHITE)
tf=tb(s,0.4,1.45,4.4,1.5,"🖍️ 彩笔  Markers",sz=13,c=DARK)
ap(tf,"⭕ 圆形纸  Circle paper",sz=13,c=DARK)
ap(tf,"🔺 三角形纸  Triangle paper",sz=13,c=DARK)
ap(tf,"📎 胶水/胶带  Glue or tape",sz=13,c=DARK)
# Middle: example signs
sh_ex=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(3.1),Inches(4.4),Inches(0.4))
sh_ex.fill.solid();sh_ex.fill.fore_color.rgb=BROWN;sh_ex.line.fill.background()
tb(s,0.4,3.13,4.2,0.35,"💡 示例主题  Example Themes",sz=14,b=True,c=WHITE)
tf_ex=tb(s,0.4,3.6,4.4,1.5,"⚠️ 河边危险  Danger by river",sz=13,c=DARK)
ap(tf_ex,"🏃 不要乱跑  Don't run around",sz=13,c=DARK)
ap(tf_ex,"💧 小心滑倒  Slippery!",sz=13,c=DARK)
ap(tf_ex,"🛡️ 注意安全  Stay safe",sz=13,c=DARK)
# Right: steps + sentence frames
sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(0.95),Inches(4.8),Inches(0.4))
sh2.fill.solid();sh2.fill.fore_color.rgb=PINE;sh2.line.fill.background()
tb(s,5.0,0.98,4.6,0.35,"👉 做法  Steps",sz=14,b=True,c=WHITE)
tf2=tb(s,5.0,1.45,4.7,1.5,"1️⃣ 老师先示范几个安全标志",sz=13,c=DARK)
ap(tf2,"2️⃣ 选一个安全主题",sz=13,c=DARK)
ap(tf2,"3️⃣ 用圆形/三角形 + 彩笔设计",sz=13,c=DARK)
ap(tf2,"4️⃣ 说一说你的标志表示什么",sz=13,c=DARK)
sh3=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.9),Inches(3.1),Inches(4.8),Inches(0.4))
sh3.fill.solid();sh3.fill.fore_color.rgb=GREEN_OK;sh3.line.fill.background()
tb(s,5.0,3.13,4.6,0.35,"🗣️ 展示句型  Say These",sz=14,b=True,c=WHITE)
tf3=tb(s,5.0,3.6,4.7,1.5,"· 这是我的安全标志。",sz=13,c=DARK)
ap(tf3,"· 它表示河边危险。",sz=13,c=DARK)
ap(tf3,"· 它告诉我们不要乱跑。",sz=13,c=DARK)
pn(s,n)

# ============================================================
# 30 Activity — Charades rules
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🎭 我演你猜  Charades — 规则 Rules",GREEN_OK)
rules_data=[
    ("1️⃣","上台抽卡","一名学生上台，抽一张环境卡"),
    ("2️⃣","只能表演","只能用动作表演，不能说话"),
    ("3️⃣","大家猜","其他同学根据动作猜环境"),
    ("4️⃣","一起说","猜对后，全班一起大声说！"),
]
for i,(num,t,d) in enumerate(rules_data):
    x=0.3+i*2.4
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.0),Inches(2.2),Inches(2.5))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=GREEN_OK;sh.line.width=Pt(3)
    tb(s,x+0.1,1.1,2.0,0.6,num,sz=30,b=True,c=GREEN_OK,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,1.8,2.0,0.5,t,sz=16,b=True,c=DARK,a=PP_ALIGN.CENTER)
    tb(s,x+0.1,2.35,2.0,1.1,d,sz=12,c=DARK,a=PP_ALIGN.CENTER)
# Ask sentences
sh3=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(3.8),Inches(9.4),Inches(1.4))
sh3.fill.solid();sh3.fill.fore_color.rgb=WARM;sh3.line.color.rgb=SUN;sh3.line.width=Pt(2)
tb(s,0.5,3.9,9,0.4,"🗣️ 可用提问句型  Ask:",sz=15,b=True,c=SUN)
tb(s,0.5,4.3,4.6,0.35,"· 这是哪里？",sz=14,c=DARK)
tb(s,0.5,4.65,4.6,0.35,"· 是森林吗？",sz=14,c=DARK)
tb(s,5.2,4.3,4.5,0.35,"· 我觉得是沙漠。",sz=14,c=DARK)
tb(s,5.2,4.65,4.5,0.35,"· 对不对？",sz=14,c=DARK)
pn(s,n)

# ============================================================
# 31 Charades hints
# ============================================================
s=ns();n+=1;bg(s,CREAM);hb(s,"🎭 我演你猜  表演提示 Acting Hints",GREEN_OK)
hints=[
    ("🌲","森林","弯腰走路、拨开树枝",FOREST),
    ("🏔️","山地","小心走路、假装爬坡",MOUNTAIN),
    ("🌾","草地","轻松走路、看远方",GRASS),
    ("🏞️","河边","停下来、低头看水",RIVER),
    ("🏜️","沙漠","擦汗、感觉很热",DESERT),
    ("❄️","雪地","发抖、抱紧身体",SNOW),
]
for i,(em,cn,act,cl) in enumerate(hints):
    col=i%3;row=i//3
    x=0.3+col*3.2;y=0.95+row*2.05
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(3.0),Inches(1.85))
    sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=cl;sh.line.width=Pt(2.5)
    tb(s,x+0.1,y+0.1,2.8,0.55,f"{em} {cn}",sz=19,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.15,y+0.7,2.7,1.1,act,sz=14,c=DARK,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 32 Day 1 Badge / Visa stamp
# ============================================================
s=ns();n+=1;bg(s,CREAM)
tb(s,0.5,0.4,9,0.8,"🎖️ Day 1 探险家徽章  Explorer Badge",sz=26,b=True,c=PINE,a=PP_ALIGN.CENTER)
sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.5),Inches(1.4),Inches(3),Inches(3))
sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=PINE;sh.line.width=Pt(5)
tf=tb(s,3.6,1.65,2.8,2.7,"DAY 1",sz=18,b=True,c=SUN,a=PP_ALIGN.CENTER)
ap(tf,"🏕️",sz=40,a=PP_ALIGN.CENTER)
ap(tf,"认识自然",sz=20,b=True,c=PINE,a=PP_ALIGN.CENTER)
ap(tf,"✓ COMPLETED",sz=13,b=True,c=GREEN_OK,a=PP_ALIGN.CENTER)
ap(tf,"🌲🏔️🌾🏞️🏜️❄️",sz=14,a=PP_ALIGN.CENTER)
tb(s,1,4.55,8,0.4,"恭喜你完成 Day 1！Congratulations, young explorer! 🎉",sz=16,b=True,c=PINE,a=PP_ALIGN.CENTER)
tb(s,1,5.0,8,0.4,"学会了 6 种环境 · 6 条安全规则 · 3 个词",sz=13,c=GRAY,a=PP_ALIGN.CENTER)
pn(s,n)

# ============================================================
# 33 Tomorrow preview
# ============================================================
s=ns();n+=1;bg(s,PINE)
tb(s,0.5,0.9,9,0.8,"🔭 明天见！  See You Tomorrow!",sz=32,b=True,c=WHITE,a=PP_ALIGN.CENTER)
tf=tb(s,1.5,2.2,7,2.5,"Day 2 — 野外工具与装备",sz=28,b=True,c=SUN,a=PP_ALIGN.CENTER)
ap(tf,"Wilderness Tools & Gear",sz=16,c=WARM,a=PP_ALIGN.CENTER)
ap(tf,"",sz=10)
ap(tf,"🎒 背包里要带什么？",sz=20,b=True,c=WHITE,a=PP_ALIGN.CENTER)
ap(tf,"What do explorers pack?",sz=14,c=WARM,a=PP_ALIGN.CENTER)
ap(tf,"",sz=10)
ap(tf,"明天见，小探险家！",sz=15,c=WARM,a=PP_ALIGN.CENTER)
pn(s,n)

OUT='/Users/Huan/projects/summercourse/Chinese/野外生存与探险wilderness_pbl/day1_nature.pptx'
prs.save(OUT);print(f"Created {n} slides → {OUT}")
