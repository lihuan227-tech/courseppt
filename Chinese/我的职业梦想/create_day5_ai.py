#!/usr/bin/env python3
"""
我的职业梦想 — Day 5: AI 与未来 (AI & the Future)
Focus: famous AI companies + AI founders + how AI will change the world.
Project: 我的 AI 帮手 — design an AI tool that helps someone.
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

# --- Palette: Future tech (deep purple + electric cyan + neon green) ---
PURPLE = RGBColor(0x6A,0x1B,0x9A)   # primary deep purple
CYAN   = RGBColor(0x00,0xBC,0xD4)   # electric cyan
NEON   = RGBColor(0x00,0xE6,0x76)   # neon green accent
NAVY   = RGBColor(0x1E,0x3A,0x5F)
GOLD   = RGBColor(0xF5,0xA6,0x23)
JET    = RGBColor(0x14,0x14,0x14)
# AI company colors
OPENAI = RGBColor(0x10,0xA3,0x7F)   # OpenAI green
ANTHRO = RGBColor(0xCC,0x78,0x5C)   # Anthropic clay
DEEP   = RGBColor(0x42,0x85,0xF4)   # DeepMind blue
META   = RGBColor(0x00,0x66,0xFF)   # Meta blue (Yann LeCun)
CREAM  = RGBColor(0xFF,0xF8,0xE7)
WARM   = RGBColor(0xFF,0xF3,0xE0)
BROWN  = RGBColor(0x6B,0x44,0x23)
WHITE  = RGBColor(0xFF,0xFF,0xFF)
DARK   = RGBColor(0x2C,0x2C,0x2C)
GRAY   = RGBColor(0x88,0x88,0x88)
LGRAY  = RGBColor(0xBB,0xBB,0xBB)
IMGBG  = RGBColor(0xE8,0xE8,0xE8)
OK     = RGBColor(0x38,0x8E,0x3C)
ALERT  = RGBColor(0xD0,0x4A,0x3C)

# === Helpers ===
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
def hb(s,txt,c=PURPLE,t=0.15):
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

def company_card(emoji,company_cn,company_en,founder_cn,founder_en,year,country,
                 product_cn,product_en,changes_cn,changes_en,color):
    s=ns(); bg(s,CREAM); hb(s,f"{emoji} 著名 AI 公司 · {company_cn}",color)
    img=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.95),Inches(3.7),Inches(3.7))
    img.fill.solid(); img.fill.fore_color.rgb=IMGBG; img.line.color.rgb=color; img.line.width=Pt(3)
    tb(s,0.4,2.15,3.7,1.0,emoji,sz=110,a=PP_ALIGN.CENTER)
    tb(s,0.4,3.30,3.7,0.45,company_cn,sz=22,b=True,c=color,a=PP_ALIGN.CENTER)
    tb(s,0.4,3.75,3.7,0.30,company_en,sz=12,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,0.4,4.10,3.7,0.30,f"{country} · {year}",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.30),Inches(0.95),Inches(5.40),Inches(3.7))
    panel.fill.solid(); panel.fill.fore_color.rgb=WHITE; panel.line.color.rgb=color; panel.line.width=Pt(2.5)
    head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.30),Inches(0.95),Inches(5.40),Inches(0.50))
    head.fill.solid(); head.fill.fore_color.rgb=color; head.line.fill.background()
    tb(s,4.45,1.03,5.1,0.4,f"👤 创始人  Founder: {founder_cn} ({founder_en})",sz=13,b=True,c=WHITE)
    tb(s,4.45,1.55,5.1,0.35,"🤖 主要产品  Main Product",sz=13,b=True,c=color)
    tb(s,4.45,1.90,5.1,0.40,product_cn,sz=14,b=True,c=DARK)
    tb(s,4.45,2.25,5.1,0.30,product_en,sz=10,c=GRAY)
    tb(s,4.45,2.65,5.1,0.35,"🌟 怎么改变世界？",sz=13,b=True,c=color)
    tb(s,4.45,3.00,5.1,0.40,changes_cn,sz=13,c=DARK)
    tb(s,4.45,3.40,5.1,0.30,changes_en,sz=10,c=GRAY)
    sentence_frame_bar(s,4.80,
        f"{company_cn} 帮我们 ___ 。 我也想用 AI ___ 。",
        f"{company_en} helps us ___. I want to use AI to ___.")

n=0

# 1. COVER
s=ns(); bg(s,CREAM)
tb(s,1,0.45,8,0.6,"我的职业梦想 · My Dream Career",sz=26,b=True,c=PURPLE,a=PP_ALIGN.CENTER)
tb(s,1,1.05,8,0.4,"Day 5",sz=18,b=True,c=CYAN,a=PP_ALIGN.CENTER)
sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.25),Inches(1.55),Inches(3.5),Inches(3.5))
sh.fill.solid(); sh.fill.fore_color.rgb=PURPLE; sh.line.color.rgb=CYAN; sh.line.width=Pt(6)
sh2=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.55),Inches(1.85),Inches(2.9),Inches(2.9))
sh2.fill.solid(); sh2.fill.fore_color.rgb=WHITE; sh2.line.color.rgb=NEON; sh2.line.width=Pt(2)
tf=tb(s,3.55,2.05,2.9,0.4,"DAY 5",sz=14,b=True,c=CYAN,a=PP_ALIGN.CENTER)
ap(tf,"🤖",sz=68,a=PP_ALIGN.CENTER)
ap(tf,"AI 与未来",sz=20,b=True,c=PURPLE,a=PP_ALIGN.CENTER)
ap(tf,"AI & THE FUTURE",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,1,5.15,8,0.4,"🚀 你和 AI 一起, 想做什么? You + AI — what will you build?",sz=14,b=True,c=NEON,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)

# 2. TODAY'S MISSION
s=ns(); bg(s,CREAM); hb(s,"🧭 今天的任务  Today's Mission",PURPLE)
tb(s,0.4,0.85,9.2,0.45,"🤖 认识 4 家正在改变世界的 AI 公司!",sz=24,b=True,c=PURPLE,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.28,"Meet 4 AI companies and their founders.",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,0.4,1.55,9.2,0.32,"👉 5 个任务 ↓",sz=14,b=True,c=BROWN,a=PP_ALIGN.CENTER)
mission_card(s,0.4,1.95,1.80,2.20,1,"AI 是什么","What",       "🤖",PURPLE)
mission_card(s,2.30,1.95,1.80,2.20,2,"4 家公司","Meet 4",       "🏢",CYAN)
mission_card(s,4.20,1.95,1.80,2.20,3,"改变世界","World Change", "🌍",NEON)
mission_card(s,6.10,1.95,1.80,2.20,4,"我会认/写","Read & Write","📖",GOLD)
mission_card(s,8.00,1.95,1.80,2.20,5,"我的 AI 帮手","My AI",   "💡",META)
sentence_frame_bar(s,4.40,"AI 让我能 ___ 。","AI lets me ___ .")
n+=1; pn(s,n)

# 3. SESSION 1 DIVIDER
s=div("Session 1  上午","🤖 AI 是什么 + 4 家 AI 公司",PURPLE,"🌟"); n+=1; pn(s,n)

# 4. WHAT IS AI?
s=ns(); bg(s,CREAM); hb(s,"💡 AI 是什么？  What is AI?",PURPLE)
tb(s,0.4,0.85,9.2,0.4,"AI = 人工智能 — 让电脑也能 4 件事:",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.28,"AI = computers that can do 4 things like humans:",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
defs=[
    ("👀","看","See","看图说话"),
    ("👂","听","Hear","听人说话"),
    ("💬","说","Speak","回答问题、写字"),
    ("🧠","学","Learn","越用越聪明"),
]
for i,(em,cn,en,desc) in enumerate(defs):
    x=0.4+i*2.30
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.65),Inches(2.2),Inches(3.0))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=PURPLE; sh.line.width=Pt(2.5)
    tb(s,x+0.05,1.80,2.10,0.8,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.65,2.10,0.45,cn,sz=18,b=True,c=PURPLE,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.10,2.10,0.30,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.10,3.50,2.00,1.10,desc,sz=12,c=DARK,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.85,"AI 像 ___ 一样, 能 ___ 。","AI is like ___ and can ___.")
n+=1; pn(s,n)
notes(s,"3 分钟:\n• AI 不是机器人 — AI 是「软件大脑」\n• 你已经在用 AI: Siri, Alexa, 抖音推荐, Google 翻译")

# 5. INTRO 4 FAMOUS AI COMPANIES
s=ns(); bg(s,CREAM); hb(s,"👋 4 家 AI 公司  4 AI Companies",CYAN)
tb(s,0.4,0.85,9.2,0.34,"猜猜看 — 你听说过哪几家？",sz=15,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.20,9.2,0.28,"Guess — which ones have you heard of?",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
companies=[
    ("🟢","OpenAI","Sam Altman","ChatGPT",OPENAI),
    ("🟠","Anthropic","Dario Amodei","Claude",ANTHRO),
    ("🔵","DeepMind","Demis Hassabis","AlphaGo",DEEP),
    ("🟣","Meta AI","Yann LeCun","开源 LLM",META),
]
for i,(em,cn,founder,product,cl) in enumerate(companies):
    x=0.4+i*2.30
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.55),Inches(2.20),Inches(3.30))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    tb(s,x+0.05,1.70,2.10,0.9,em,sz=68,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.65,2.10,0.45,cn,sz=20,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.10,2.10,0.30,f"{founder}",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    sep=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(x+0.30),Inches(3.55),Inches(1.60),Inches(0.02))
    sep.fill.solid(); sep.fill.fore_color.rgb=cl; sep.line.fill.background()
    tb(s,x+0.05,3.70,2.10,0.7,product,sz=14,b=True,c=cl,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,5.05,"我用过 ___ 。 我没用过 ___ 。","I've used ___. I haven't used ___.")
n+=1; pn(s,n)

# 6. OPENAI
company_card("🟢","OpenAI","OpenAI","萨姆·奥特曼","Sam Altman","2015","🇺🇸 美国",
    "ChatGPT — 你跟它说话, 它写文章、回答问题、帮你做作业",
    "ChatGPT — talk to it, it writes, answers, helps homework",
    "让全世界的人都可以「跟 AI 聊天」, 第一个让 AI 进入每个家庭。",
    "Made AI conversational for everyone — first to bring AI into every home.",
    OPENAI)
n+=1; pn(s,n)

# 7. ANTHROPIC
company_card("🟠","Anthropic","Anthropic","达里奥·阿莫迪","Dario Amodei","2021","🇺🇸 美国",
    "Claude — 一个安全、诚实、会思考的 AI 助手",
    "Claude — a safe, honest, thoughtful AI assistant",
    "专注「AI 安全」 — 让 AI 不只是聪明, 还要诚实和有用。",
    "Focuses on AI safety — making AI honest and helpful, not just smart.",
    ANTHRO)
n+=1; pn(s,n)
notes(s,"重点:\n• AI 不只要聪明, 还要安全\n• 「Claude」就是这家公司的 AI 助手\n• 创始人是 4 兄妹科学家家庭")

# 8. DEEPMIND
company_card("🔵","DeepMind","Google DeepMind","戴米斯·哈萨比斯","Demis Hassabis","2010","🇬🇧 英国 · Google",
    "AlphaGo 下围棋打败世界冠军 · AlphaFold 预测蛋白质结构",
    "AlphaGo beat the world's best Go player · AlphaFold cracked protein folding",
    "用 AI 解决科学难题 — 帮医生发现新药, 救命。",
    "Uses AI for hard science — helps doctors find new medicines, saves lives.",
    DEEP)
n+=1; pn(s,n)
notes(s,"重点:\n• 2016 年 AlphaGo 打败李世石 — 改变了人对 AI 的看法\n• AlphaFold 帮助治疗癌症、阿尔兹海默症\n• 创始人 Demis 13 岁就是国际象棋大师")

# 9. META AI
company_card("🟣","Meta AI","Meta AI","杨立昆","Yann LeCun","2013","🇫🇷 法国 · Meta",
    "Llama — 把强大的 AI 「开源」, 让全世界免费用",
    "Llama — open-source AI that anyone can download and use for free",
    "「图灵奖」得主 — 让小公司也能用上最好的 AI, 不只是大公司。",
    "Turing Award winner — democratized AI so small teams (not just big tech) can use it.",
    META)
n+=1; pn(s,n)

# 10. HOW AI CHANGES THE WORLD
s=ns(); bg(s,CREAM); hb(s,"🌍 AI 怎么改变世界？  How AI Changes the World",NEON)
tb(s,0.4,0.85,9.2,0.34,"AI 已经在帮我们做这些事:",sz=15,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.20,9.2,0.28,"AI is already helping us with:",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
ways=[
    ("🩺","医疗","Health","看 X 光、找药"),
    ("📚","学习","Learn","一对一辅导"),
    ("🎨","创作","Create","画画、写歌"),
    ("🚗","开车","Drive","自动驾驶"),
    ("🌍","翻译","Translate","100 种语言"),
    ("♿","帮残疾","Accessible","盲人「看见」"),
]
for i,(em,cn,en,how) in enumerate(ways):
    col=i%3; row=i//3
    x=0.3+col*3.2; y=1.55+row*1.65
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(3.0),Inches(1.45))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=NEON; sh.line.width=Pt(2)
    tb(s,x+0.10,y+0.05,0.6,0.55,em,sz=30,a=PP_ALIGN.CENTER)
    tb(s,x+0.75,y+0.10,2.2,0.35,cn,sz=15,b=True,c=PURPLE)
    tb(s,x+0.75,y+0.42,2.2,0.30,en,sz=10,c=GRAY)
    tb(s,x+0.10,y+0.85,2.85,0.55,f"→ {how}",sz=12,c=DARK)
sentence_frame_bar(s,4.95,"AI 帮我 ___ — 我觉得 ___ 。","AI helps me ___. I think ___.")
n+=1; pn(s,n)

# 11. SESSION 2 DIVIDER
s=div("Session 2  下午","📖 复习 + 我会认 + 我会写",CYAN,"📖"); n+=1; pn(s,n)

# 12. REVIEW
s=ns(); bg(s,CREAM); hb(s,"🔄 复习  Review · Session 1",PURPLE)
tb(s,0.4,0.85,9.2,0.4,"还记得吗？  Do you remember?",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
qs=[
    ("🤖","AI 能做哪 4 件事？","看、听、说、学"),
    ("🟢","ChatGPT 是哪家公司的？","OpenAI (Sam Altman)"),
    ("🟠","Claude 是哪家公司的？","Anthropic (Dario Amodei)"),
    ("🔵","下围棋打败人的 AI 叫什么？","AlphaGo (DeepMind)"),
    ("🟣","「图灵奖」得主, 把 AI 开源的是谁？","杨立昆 Yann LeCun"),
]
for i,(em,q,a) in enumerate(qs):
    y=1.45+i*0.70
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(y),Inches(9.2),Inches(0.6))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=PURPLE; sh.line.width=Pt(1.5)
    tb(s,0.55,y+0.10,0.5,0.4,em,sz=22,a=PP_ALIGN.CENTER)
    tb(s,1.15,y+0.10,4.0,0.4,q,sz=14,b=True,c=DARK)
    tb(s,5.30,y+0.10,4.2,0.4,f"→ {a}",sz=14,b=True,c=PURPLE)
n+=1; pn(s,n)

# 13. 我会认 — 4 AI words
s=ns(); bg(s,CREAM); hb(s,"📖 我会认  I Can Read · 4 个 AI 词",PURPLE)
words=[
    ("机器人","jī qì rén","robot","🤖"),
    ("聪明","cōng míng","smart","🧠"),
    ("未来","wèi lái","future","🚀"),
    ("帮手","bāng shǒu","helper / sidekick","🤝"),
]
for i,(cn,py,en,em) in enumerate(words):
    col=i%2; row=i//2
    x=0.4+col*4.7; y=0.95+row*2.0
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(4.5),Inches(1.85))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=PURPLE; sh.line.width=Pt(2.5)
    tb(s,x+0.15,y+0.10,1.0,1.5,em,sz=70,a=PP_ALIGN.CENTER)
    tb(s,x+1.30,y+0.20,3.0,0.7,cn,sz=36,b=True,c=PURPLE)
    tb(s,x+1.30,y+0.95,3.0,0.35,py,sz=14,c=GRAY)
    tb(s,x+1.30,y+1.30,3.0,0.40,en,sz=14,b=True,c=DARK)
n+=1; pn(s,n)

# 14. 我会写 — 未 (future)
s=ns(); bg(s,CREAM); hb(s,"✏️ 我会写 · 未  I Can Write · 'not yet / future'",PURPLE)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.95),Inches(4.5),Inches(3.7))
sh.fill.solid(); sh.fill.fore_color.rgb=WARM; sh.line.color.rgb=PURPLE; sh.line.width=Pt(3)
tb(s,0.4,1.30,4.5,2.5,"未",sz=300,b=True,c=PURPLE,a=PP_ALIGN.CENTER)
tb(s,0.4,4.10,4.5,0.40,"wèi · not yet / future",sz=18,c=GRAY,a=PP_ALIGN.CENTER)
panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.05),Inches(0.95),Inches(4.65),Inches(3.7))
panel.fill.solid(); panel.fill.fore_color.rgb=WHITE; panel.line.color.rgb=PURPLE; panel.line.width=Pt(2)
tb(s,5.20,1.10,4.4,0.5,"📝 笔顺  5 笔",sz=18,b=True,c=PURPLE)
tb(s,5.20,1.65,4.4,0.4,"很像「木」 + 上面一短横",sz=14,c=DARK)
tb(s,5.20,2.05,4.4,0.4,"未来 = 还没到的时间",sz=14,b=True,c=PURPLE)
tb(s,5.20,2.45,4.4,0.4,"未 + 来 = 没来 → 未来",sz=13,c=DARK)
tb(s,5.20,2.95,4.4,0.35,"📝 练习 Practice:",sz=14,b=True,c=PURPLE)
tb(s,5.20,3.35,4.4,0.35,"1️⃣ 空中写 · 2️⃣ 手心写 · 3️⃣ 纸上 3 次",sz=13,c=DARK)
sentence_frame_bar(s,4.85,"未 来, 我想 ___ 。","In the future, I want to ___.")
n+=1; pn(s,n)

# 15. SESSION 3 DIVIDER — My AI Helper
s=div("Session 3  下午","💡 我的 AI 帮手 My AI Helper",NEON,"💡"); n+=1; pn(s,n)

# 16. PROJECT INTRO — Design my AI helper
s=ns(); bg(s,CREAM); hb(s,"💡 我的 AI 帮手  My AI Helper",NEON)
tb(s,0.4,0.85,9.2,0.5,"🤖 设计一个 AI — 帮谁? 解决什么问题?",sz=20,b=True,c=NEON,a=PP_ALIGN.CENTER)
tb(s,0.4,1.45,9.2,0.34,"Design an AI helper — who does it help and what does it solve?",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(2.5),Inches(1.90),Inches(5.0),Inches(2.85))
sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=NEON; sh.line.width=Pt(3)
tb(s,2.5,2.00,5.0,0.45,"💡 我的 AI 帮手卡",sz=18,b=True,c=PURPLE,a=PP_ALIGN.CENTER)
tb(s,2.5,2.45,5.0,0.30,"My AI Helper Card",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,2.7,2.90,4.6,0.40,"1. 我的 AI 叫: ___",sz=13,b=True,c=DARK)
tb(s,2.7,3.30,4.6,0.40,"2. 它帮: ___ (谁?)",sz=13,b=True,c=DARK)
tb(s,2.7,3.70,4.6,0.40,"3. 它能 ___ (做什么?)",sz=13,b=True,c=DARK)
tb(s,2.7,4.10,4.6,0.40,"4. 一句标语: ___",sz=13,b=True,c=DARK)
tb(s,0.4,4.85,9.2,0.32,"⏱️ 20 分钟设计 + 画 + 5 分钟讲",sz=13,b=True,c=BROWN,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)

# 17. STEP 1 — pick who to help
s=ns(); bg(s,CREAM); hb(s,"1️⃣ 帮谁?  Step 1 · Who Do You Help?",NEON)
tb(s,0.4,0.85,9.2,0.4,"AI 可以帮各种各样的人 — 选 1 个!",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.30,"AI can help many kinds of people — pick ONE!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
who=[
    ("👨‍🦯","盲人","Blind","用语音「看」周围"),
    ("👶","小朋友","Kids","讲故事 / 学新东西"),
    ("👴","老人","Elderly","提醒吃药 / 找路"),
    ("🌍","地球","The Earth","预测天气, 减少浪费"),
    ("🐾","动物","Animals","识别濒危物种"),
    ("🩺","医生","Doctors","看片子, 早发现病"),
]
for i,(em,cn,en,ex) in enumerate(who):
    col=i%3; row=i//3
    x=0.3+col*3.2; y=1.65+row*1.55
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(3.0),Inches(1.40))
    sh.fill.solid(); sh.fill.fore_color.rgb=WARM; sh.line.color.rgb=NEON; sh.line.width=Pt(2)
    tb(s,x+0.10,y+0.05,0.6,0.55,em,sz=28,a=PP_ALIGN.CENTER)
    tb(s,x+0.75,y+0.10,2.2,0.35,cn,sz=15,b=True,c=PURPLE)
    tb(s,x+0.75,y+0.42,2.2,0.30,en,sz=10,c=GRAY)
    tb(s,x+0.10,y+0.83,2.85,0.5,f"💭 {ex}",sz=12,c=DARK)
sentence_frame_bar(s,4.95,"我的 AI 帮 ___ 。","My AI helps ___.")
n+=1; pn(s,n)

# 18. STEP 2 — Design (4 questions)
s=ns(); bg(s,CREAM); hb(s,"2️⃣ 设计 AI  Step 2 · Design Your AI",NEON)
tb(s,0.4,0.85,9.2,0.4,"问自己 4 个问题 — 让 AI 更有用",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
qs=[
    ("👀","它能看吗？","Can it see?"),
    ("👂","它能听吗？","Can it hear?"),
    ("💬","它能说吗？","Can it speak?"),
    ("🦾","它能做吗？","Can it do things?"),
]
for i,(em,cn,en) in enumerate(qs):
    x=0.4+i*2.30
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.55),Inches(2.20),Inches(2.85))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=NEON; sh.line.width=Pt(2.5)
    tb(s,x+0.05,1.70,2.10,0.7,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.50,2.10,0.45,cn,sz=18,b=True,c=PURPLE,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.00,2.10,0.30,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.10,3.50,2.00,0.85,"☐ 是  ☐ 不是",sz=14,c=DARK,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.60,"我的 AI 能 ___ 和 ___ 。","My AI can ___ and ___.")
n+=1; pn(s,n)

# 19. STEP 3 — Demo + Vote (also: AI safety!)
s=ns(); bg(s,CREAM); hb(s,"3️⃣ 演示 + 投票  Step 3 · Demo & Vote",NEON)
tb(s,0.4,0.85,9.2,0.4,"展示你的 AI — 但要先回答 1 个问题:",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.3,"Pitch your AI — but first answer this:",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# Safety question card
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(1.5),Inches(1.65),Inches(7.0),Inches(1.4))
sh.fill.solid(); sh.fill.fore_color.rgb=WARM; sh.line.color.rgb=ALERT; sh.line.width=Pt(3)
tb(s,1.5,1.75,7.0,0.45,"⚠️ 你的 AI 安全吗？",sz=18,b=True,c=ALERT,a=PP_ALIGN.CENTER)
tb(s,1.5,2.20,7.0,0.35,"会不会伤害人？会不会说假话？",sz=14,c=DARK,a=PP_ALIGN.CENTER)
tb(s,1.5,2.60,7.0,0.30,"Is your AI safe? Could it harm anyone? Could it lie?",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# Demo steps
demo=[
    ("🎤","Pitch","60 秒讲完"),
    ("🛡️","Safety","回答安全问题"),
    ("🗳️","Vote","全班投票"),
]
for i,(em,cn,desc) in enumerate(demo):
    x=1.0+i*2.85
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(3.30),Inches(2.5),Inches(1.55))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=NEON; sh.line.width=Pt(2.5)
    tb(s,x+0.05,3.40,2.40,0.55,em,sz=30,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.95,2.40,0.4,cn,sz=16,b=True,c=PURPLE,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,4.40,2.40,0.4,desc,sz=11,c=DARK,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,5.00,"我的 AI 是安全的, 因为 ___ 。","My AI is safe because ___.")
n+=1; pn(s,n)
notes(s,"5 分钟:\n• 安全很重要 — 这是 Anthropic 的核心理念\n• AI 不是越聪明越好, 还要安全 + 诚实 + 有用")

# 20. CLOSING + UNIT BADGE
s=ns(); bg(s,CREAM)
tb(s,0.5,0.4,9,0.8,"🎖️ Day 5 + 整个单元徽章  Unit Complete!",sz=24,b=True,c=PURPLE,a=PP_ALIGN.CENTER)
sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.5),Inches(1.4),Inches(3),Inches(3))
sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=NEON; sh.line.width=Pt(6)
tf=tb(s,3.6,1.65,2.8,2.7,"DAY 5 + 全单元",sz=14,b=True,c=NEON,a=PP_ALIGN.CENTER)
ap(tf,"🤖",sz=42,a=PP_ALIGN.CENTER)
ap(tf,"AI 与未来",sz=18,b=True,c=PURPLE,a=PP_ALIGN.CENTER)
ap(tf,"✓ COMPLETED",sz=12,b=True,c=OK,a=PP_ALIGN.CENTER)
ap(tf,"🌍🔬💡❤️🤖",sz=16,a=PP_ALIGN.CENTER)
tb(s,1,4.55,8,0.4,"🎉 5 天的旅程结束! 你认识了科学家、企业家、医生、老师、AI 创始人!",sz=13,b=True,c=PURPLE,a=PP_ALIGN.CENTER)
tb(s,1,5.0,8,0.4,"长大后, 你会成为 ___ — 改变这个世界吗?",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)

out=os.path.join(os.path.dirname(__file__),"day5_ai.pptx")
prs.save(out); print(f"Saved {out}  ({n} slides)")
