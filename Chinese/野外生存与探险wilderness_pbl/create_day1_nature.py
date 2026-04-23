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

def env_slide(emoji,cn,en,color,look,features,dangers,rules):
    """One environment: colored header + 4-box grid (look / features / dangers / rules)."""
    s=ns();bg(s,CREAM)
    # Big colored header bar
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.3),Inches(0.15),Inches(9.4),Inches(0.7))
    sh.fill.solid();sh.fill.fore_color.rgb=color;sh.line.fill.background()
    tb(s,0.5,0.22,1.0,0.55,emoji,sz=28,c=WHITE)
    tb(s,1.6,0.23,4.0,0.55,cn,sz=28,b=True,c=WHITE)
    tb(s,5.7,0.38,3.9,0.4,en,sz=16,c=WARM)
    # 4-box 2x2 grid
    boxes=[
        ("👀 长什么样  Looks Like",look,color,False),
        ("✨ 特点  Features",features,SUN,False),
        ("⚠️ 小心危险  Dangers",dangers,ALERT,True),
        ("✅ 安全规则  Safety Rules",rules,GREEN_OK,True),
    ]
    for i,(title,items,cl,is_rule) in enumerate(boxes):
        col=i%2;row=i//2
        x=0.3+col*4.75;y=1.05+row*2.1
        sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(4.6),Inches(1.95))
        sh.fill.solid();sh.fill.fore_color.rgb=WHITE;sh.line.color.rgb=cl;sh.line.width=Pt(2.5)
        # title bar
        sh2=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(4.6),Inches(0.42))
        sh2.fill.solid();sh2.fill.fore_color.rgb=cl;sh2.line.fill.background()
        tb(s,x+0.15,y+0.05,4.3,0.35,title,sz=14,b=True,c=WHITE)
        # items
        tf=tb(s,x+0.2,y+0.5,4.3,1.4,items[0],sz=13,c=DARK)
        for it in items[1:]:
            ap(tf,it,sz=13,c=DARK)
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
# 6-11  Six environment slides
# ============================================================
environments_data=[
    ("🌲","森林","Forest",FOREST,
        ["• 树木茂密，阳光不多","• 地上有落叶和小草","• 能听见鸟叫虫鸣"],
        ["• 空气很新鲜","• 有很多动物：鸟、松鼠、鹿","• 有蘑菇和浆果"],
        ["• 容易迷路 🗺️","• 遇到野生动物 🐺","• 毒蘑菇、毒蛇"],
        ["• 跟紧老师，不乱跑","• 大声说话，不安静靠近动物","• 不吃不认识的蘑菇/果子"]),
    ("🏔️","山地","Mountain",MOUNTAIN,
        ["• 又高又陡","• 有大石头和山路","• 山顶空气稀薄"],
        ["• 风景美丽","• 可以爬山看远方","• 气温比山下低"],
        ["• 容易滑倒摔伤","• 高处掉石头 🪨","• 天气变化快、下雨冷"],
        ["• 穿结实的鞋","• 不乱扔石头","• 带外套，注意天气预报"]),
    ("🌾","草地","Grassland",GRASS,
        ["• 一大片绿绿的草","• 地势平坦","• 有野花和小虫"],
        ["• 阳光很充足","• 适合跑步、野餐","• 有蝴蝶和小动物"],
        ["• 烈日晒伤 ☀️","• 虫子叮咬 🦟","• 突然有蜜蜂"],
        ["• 戴帽子 + 防晒霜","• 喷防虫液","• 看到蜜蜂不要挥手"]),
    ("🏞️","河边","Riverside",RIVER,
        ["• 有流动的水","• 岸边有石头和泥","• 可能有小鱼小虾"],
        ["• 水很凉很清","• 可以摸小石头","• 常有柳树和芦苇"],
        ["• 溺水 ⚠️（最危险！）","• 滑倒摔进水里","• 水底看不清，有尖石"],
        ["• 永远和大人在一起","• 不独自下水","• 穿防滑鞋"]),
    ("🏜️","沙漠","Desert",DESERT,
        ["• 到处是沙子","• 很少看到树","• 白天非常热晚上冷"],
        ["• 有仙人掌和耐旱植物","• 沙丘很漂亮","• 有骆驼、蜥蜴"],
        ["• 中暑 🥵 脱水","• 迷路（到处一样）","• 沙尘暴"],
        ["• 多喝水 (1 天 2-3 升)","• 戴帽子、穿长袖","• 不能独自离开队伍"]),
    ("❄️","雪地","Snow",SNOW,
        ["• 到处白茫茫","• 地上都是雪","• 树枝也有雪"],
        ["• 可以打雪仗、堆雪人","• 空气非常冷","• 走路会发出吱吱声"],
        ["• 冻伤手脚 🥶","• 雪盲（太亮看不见）","• 踩到薄冰掉进水里"],
        ["• 穿厚厚的衣服、手套","• 戴墨镜或雪镜","• 不在湖面/河面上走"]),
]
for em,cn,en,color,look,feat,dng,rul in environments_data:
    s=env_slide(em,cn,en,color,look,feat,dng,rul)
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
