#!/usr/bin/env python3
"""
我的职业梦想 — Day 4: 帮助别人 (Helpers — Doctors & Teachers)
Focus: famous doctors and teachers + 我想感谢的人 thank-you card project
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

# --- Palette: Care + Wisdom (helping green + heart pink + chalk navy) ---
HELP   = RGBColor(0x2E,0x7D,0x32)   # primary helping green
HEART  = RGBColor(0xE5,0x3E,0x5E)   # care pink-red
NAVY   = RGBColor(0x1E,0x3A,0x5F)   # chalk navy
GOLD   = RGBColor(0xF5,0xA6,0x23)
DOC    = RGBColor(0xE5,0x3E,0x3E)   # 医生 red
TEACH  = RGBColor(0x43,0xA0,0x47)   # 老师 green
WISE   = RGBColor(0x6B,0x44,0x23)   # 孔子 brown
NIGHT  = RGBColor(0x52,0x6E,0x8E)   # 南丁格尔 navy
KELLER = RGBColor(0xA0,0x6B,0x9F)   # 安·沙利文 purple
DOCTOR = RGBColor(0xC8,0x3E,0x46)   # 希波克拉底 wine
CREAM  = RGBColor(0xFF,0xF8,0xE7)
WARM   = RGBColor(0xFF,0xF3,0xE0)
BROWN  = RGBColor(0x6B,0x44,0x23)
WHITE  = RGBColor(0xFF,0xFF,0xFF)
DARK   = RGBColor(0x2C,0x2C,0x2C)
GRAY   = RGBColor(0x88,0x88,0x88)
LGRAY  = RGBColor(0xBB,0xBB,0xBB)
IMGBG  = RGBColor(0xE8,0xE8,0xE8)
OK     = RGBColor(0x38,0x8E,0x3C)
PURPLE = RGBColor(0x7B,0x1F,0xA2)

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
def hb(s,txt,c=HELP,t=0.15):
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

def person_card(emoji,cn,en,years,country,one_line_cn,one_line_en,fact_cn,fact_en,color,role_label="著名"):
    s=ns(); bg(s,CREAM); hb(s,f"{emoji} {role_label} · {cn}",color)
    img=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.95),Inches(3.7),Inches(3.7))
    img.fill.solid(); img.fill.fore_color.rgb=IMGBG; img.line.color.rgb=color; img.line.width=Pt(3)
    tb(s,0.4,2.15,3.7,1.0,emoji,sz=110,a=PP_ALIGN.CENTER)
    tb(s,0.4,3.30,3.7,0.45,cn,sz=22,b=True,c=color,a=PP_ALIGN.CENTER)
    tb(s,0.4,3.75,3.7,0.30,en,sz=12,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,0.4,4.10,3.7,0.30,f"{country} · {years}",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
    panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.30),Inches(0.95),Inches(5.40),Inches(3.7))
    panel.fill.solid(); panel.fill.fore_color.rgb=WHITE; panel.line.color.rgb=color; panel.line.width=Pt(2.5)
    head=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(4.30),Inches(0.95),Inches(5.40),Inches(0.50))
    head.fill.solid(); head.fill.fore_color.rgb=color; head.line.fill.background()
    tb(s,4.45,1.03,5.1,0.4,"💡 一句话介绍  In One Line",sz=14,b=True,c=WHITE)
    tb(s,4.45,1.55,5.1,0.45,one_line_cn,sz=17,b=True,c=DARK)
    tb(s,4.45,2.00,5.1,0.32,one_line_en,sz=11,c=GRAY)
    tb(s,4.45,2.50,5.1,0.35,"🌟 他/她做了什么？",sz=14,b=True,c=color)
    tb(s,4.45,2.85,5.1,0.40,fact_cn,sz=14,c=DARK)
    tb(s,4.45,3.30,5.1,0.32,fact_en,sz=10,c=GRAY)
    sentence_frame_bar(s,4.80,
        f"{cn} 帮助了 ___ 。 我也想 ___ 。",
        f"{en} helped ___. I also want to ___.")

n=0

# 1. COVER
s=ns(); bg(s,CREAM)
tb(s,1,0.45,8,0.6,"我的职业梦想 · My Dream Career",sz=26,b=True,c=HELP,a=PP_ALIGN.CENTER)
tb(s,1,1.05,8,0.4,"Day 4",sz=18,b=True,c=HEART,a=PP_ALIGN.CENTER)
sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.25),Inches(1.55),Inches(3.5),Inches(3.5))
sh.fill.solid(); sh.fill.fore_color.rgb=HELP; sh.line.color.rgb=HEART; sh.line.width=Pt(6)
sh2=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.55),Inches(1.85),Inches(2.9),Inches(2.9))
sh2.fill.solid(); sh2.fill.fore_color.rgb=WHITE; sh2.line.color.rgb=HEART; sh2.line.width=Pt(2)
tf=tb(s,3.55,2.05,2.9,0.4,"DAY 4",sz=14,b=True,c=HEART,a=PP_ALIGN.CENTER)
ap(tf,"❤️",sz=68,a=PP_ALIGN.CENTER)
ap(tf,"帮助别人",sz=20,b=True,c=HELP,a=PP_ALIGN.CENTER)
ap(tf,"HELPERS",sz=10,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,1,5.15,8,0.4,"❤️ 谁在帮助你? Who has helped YOU?",sz=14,b=True,c=HEART,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)

# 2. TODAY'S MISSION
s=ns(); bg(s,CREAM); hb(s,"🧭 今天的任务  Today's Mission",HELP)
tb(s,0.4,0.85,9.2,0.45,"❤️ 认识 4 位帮助别人的「英雄」!",sz=24,b=True,c=HELP,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.28,"Meet 4 helpers — doctors, nurses, and teachers.",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,0.4,1.55,9.2,0.32,"👉 5 个任务 ↓",sz=14,b=True,c=BROWN,a=PP_ALIGN.CENTER)
mission_card(s,0.4,1.95,1.80,2.20,1,"帮助是什么","What",     "🤝",HELP)
mission_card(s,2.30,1.95,1.80,2.20,2,"认识 4 位","Meet 4",     "👋",HEART)
mission_card(s,4.20,1.95,1.80,2.20,3,"医生 vs 老师","Compare", "⚖️",NAVY)
mission_card(s,6.10,1.95,1.80,2.20,4,"我会认/写","Read & Write","📖",PURPLE)
mission_card(s,8.00,1.95,1.80,2.20,5,"感谢卡","Thank-you",     "💌",GOLD)
sentence_frame_bar(s,4.40,"我想感谢 ___ 。","I want to thank ___.")
n+=1; pn(s,n)

# 3. SESSION 1 DIVIDER
s=div("Session 1  上午","❤️ 帮助是什么 + 4 位英雄",HELP,"🌟"); n+=1; pn(s,n)

# 4. WHAT IS A HELPER?
s=ns(); bg(s,CREAM); hb(s,"💡 帮助别人是什么？  What is Helping?",HELP)
tb(s,0.4,0.85,9.2,0.4,"帮助别人有 4 种力量:",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.28,"Helpers use 4 superpowers:",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
defs=[
    ("👂","会听","Listen","真的听别人说话"),
    ("🧠","会想","Think","想办法解决问题"),
    ("💪","会做","Do","真的去做, 不只说"),
    ("❤️","有爱心","Care","把别人当作家人"),
]
for i,(em,cn,en,desc) in enumerate(defs):
    x=0.4+i*2.30
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.65),Inches(2.2),Inches(3.0))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=HELP; sh.line.width=Pt(2.5)
    tb(s,x+0.05,1.80,2.10,0.8,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.65,2.10,0.45,cn,sz=18,b=True,c=HELP,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.10,2.10,0.30,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.10,3.50,2.00,1.10,desc,sz=12,c=DARK,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.85,"帮助别人就是 ___ 。","Helping is ___ .")
n+=1; pn(s,n)

# 5. INTRO TO 4 FAMOUS HELPERS
s=ns(); bg(s,CREAM); hb(s,"👋 4 位帮助世界的英雄  4 Helping Heroes",HEART)
tb(s,0.4,0.85,9.2,0.34,"猜猜看 — 谁是医生？谁是老师？",sz=15,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.20,9.2,0.28,"Guess — who's a doctor? who's a teacher?",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
helpers=[
    ("🩺","希波克拉底","Hippocrates","医生之父",DOCTOR),
    ("🕯️","南丁格尔","F. Nightingale","护士之母",NIGHT),
    ("📚","孔子","Confucius","老师之祖",WISE),
    ("✋","安·沙利文","Anne Sullivan","海伦的老师",KELLER),
]
for i,(em,cn,en,sym,cl) in enumerate(helpers):
    x=0.4+i*2.30
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.55),Inches(2.20),Inches(3.30))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=cl; sh.line.width=Pt(3)
    tb(s,x+0.05,1.70,2.10,0.9,em,sz=68,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.65,2.10,0.45,cn,sz=18,b=True,c=cl,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.10,2.10,0.30,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    sep=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(x+0.30),Inches(3.55),Inches(1.60),Inches(0.02))
    sep.fill.solid(); sep.fill.fore_color.rgb=cl; sep.line.fill.background()
    tb(s,x+0.05,3.70,2.10,0.7,sym,sz=14,b=True,c=cl,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,5.05,"我猜 ___ 是医生, ___ 是老师。","I guess ___ is a doctor, ___ is a teacher.")
n+=1; pn(s,n)

# 6. HIPPOCRATES spotlight
person_card("🩺","希波克拉底","Hippocrates","公元前 460–370","🇬🇷 希腊 · 「医学之父」",
    "他第一个说: 「医生要照顾病人, 不能伤害他们。」",
    "First to say doctors must care for patients — never harm them.",
    "今天医生当上医生时, 还要发「希波克拉底誓言」 — 答应做好医生!",
    "Today doctors still take the Hippocratic Oath when they begin practice.",
    DOCTOR, role_label="著名医生")
n+=1; pn(s,n)

# 7. NIGHTINGALE spotlight
person_card("🕯️","南丁格尔","Florence Nightingale","1820–1910","🇬🇧 英国 · 「现代护士之母」",
    "战争中, 她半夜提着灯, 一床一床照顾受伤的士兵。",
    "In war, she carried a lamp at night to care for wounded soldiers.",
    "她让医院变干净, 救了千千万万人 — 5 月 12 日是「国际护士节」纪念她。",
    "She made hospitals clean and saved countless lives. May 12 = International Nurses' Day.",
    NIGHT, role_label="著名护士")
n+=1; pn(s,n)
notes(s,"3-4 分钟:\n• 「The Lady with the Lamp」\n• 现代护士的标志\n• 她还是数学家 — 用图表说服政府改革医院")

# 8. CONFUCIUS spotlight
person_card("📚","孔子","Confucius","公元前 551–479","🇨🇳 中国 · 「万世师表」",
    "他说: 「学而时习之, 不亦说乎?」 学习是快乐的事!",
    "He said: 'Isn't it joyful to learn and practice?'",
    "中国第一位大老师 — 教过 3000 多个学生; 他的话影响了 2500 年的中国教育。",
    "China's first great teacher — taught 3,000 students; his ideas shaped 2,500 years of education.",
    WISE, role_label="著名老师")
n+=1; pn(s,n)

# 9. ANNE SULLIVAN spotlight
person_card("✋","安·沙利文","Anne Sullivan","1866–1936","🇺🇸 美国 · 海伦·凯勒的老师",
    "她教又盲又聋的海伦说话、写字、上大学。",
    "She taught Helen Keller — who was blind & deaf — to speak, write, and graduate college.",
    "她在海伦手心写字, 让她「摸到」字母。耐心 + 爱心改变了一个人的一生。",
    "She wrote letters in Helen's palm. Patience + love changed one life forever.",
    KELLER, role_label="著名老师")
n+=1; pn(s,n)
notes(s,"重点: 「老师 ≠ 只在教室」\n• 一对一的耐心也是教\n• 影片推荐: The Miracle Worker")

# 10. DOCTOR vs TEACHER comparison
s=ns(); bg(s,CREAM); hb(s,"⚖️ 医生 vs 老师  Doctor vs Teacher",NAVY)
tb(s,0.4,0.85,9.2,0.34,"两个职业都帮人, 但帮的方式不一样",sz=15,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.20,9.2,0.28,"Both help — but in different ways",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
# Two columns
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(1.55),Inches(4.55),Inches(3.4))
sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=DOC; sh.line.width=Pt(2.5)
tb(s,0.4,1.65,4.55,0.5,"🩺 医生  Doctor",sz=20,b=True,c=DOC,a=PP_ALIGN.CENTER)
doc_rows=[("帮人","治病、救命"),("工具","听诊器、药、手术刀"),("地方","医院、诊所"),("感觉","紧张, 但救人")]
for i,(k,v) in enumerate(doc_rows):
    y=2.20+i*0.65
    tb(s,0.6,y,1.5,0.4,k,sz=13,b=True,c=DOC)
    tb(s,2.1,y,2.7,0.4,v,sz=13,c=DARK)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.05),Inches(1.55),Inches(4.55),Inches(3.4))
sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=TEACH; sh.line.width=Pt(2.5)
tb(s,5.05,1.65,4.55,0.5,"📚 老师  Teacher",sz=20,b=True,c=TEACH,a=PP_ALIGN.CENTER)
teach_rows=[("帮人","教知识、教做人"),("工具","书、粉笔、电脑"),("地方","学校、教室"),("感觉","开心 — 看学生进步")]
for i,(k,v) in enumerate(teach_rows):
    y=2.20+i*0.65
    tb(s,5.25,y,1.5,0.4,k,sz=13,b=True,c=TEACH)
    tb(s,6.75,y,2.7,0.4,v,sz=13,c=DARK)
sentence_frame_bar(s,5.10,"我想当 ___ , 因为 ___ 。","I want to be a ___ because ___.")
n+=1; pn(s,n)

# 11. SESSION 2 DIVIDER
s=div("Session 2  下午","📖 复习 + 我会认 + 我会写",HEART,"📖"); n+=1; pn(s,n)

# 12. REVIEW
s=ns(); bg(s,CREAM); hb(s,"🔄 复习  Review · Session 1",HELP)
tb(s,0.4,0.85,9.2,0.4,"还记得吗？  Do you remember?",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
qs=[
    ("🤝","帮助别人有哪 4 种力量？","会听、会想、会做、有爱心"),
    ("🩺","谁是「医学之父」？","希波克拉底"),
    ("🕯️","谁是「现代护士之母」？","南丁格尔"),
    ("📚","谁是中国「万世师表」？","孔子"),
    ("✋","谁教海伦·凯勒说话写字？","安·沙利文"),
]
for i,(em,q,a) in enumerate(qs):
    y=1.45+i*0.70
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(y),Inches(9.2),Inches(0.6))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=HELP; sh.line.width=Pt(1.5)
    tb(s,0.55,y+0.10,0.5,0.4,em,sz=22,a=PP_ALIGN.CENTER)
    tb(s,1.15,y+0.10,4.0,0.4,q,sz=14,b=True,c=DARK)
    tb(s,5.30,y+0.10,4.2,0.4,f"→ {a}",sz=14,b=True,c=HELP)
n+=1; pn(s,n)

# 13. 我会认 — 4 helping words
s=ns(); bg(s,CREAM); hb(s,"📖 我会认  I Can Read · 4 个「帮助」词",HELP)
words=[
    ("帮助","bāng zhù","help","🤝"),
    ("医生","yī shēng","doctor","🩺"),
    ("老师","lǎo shī","teacher","📚"),
    ("谢谢","xiè xiè","thank you","💌"),
]
for i,(cn,py,en,em) in enumerate(words):
    col=i%2; row=i//2
    x=0.4+col*4.7; y=0.95+row*2.0
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(4.5),Inches(1.85))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=HELP; sh.line.width=Pt(2.5)
    tb(s,x+0.15,y+0.10,1.0,1.5,em,sz=70,a=PP_ALIGN.CENTER)
    tb(s,x+1.30,y+0.20,3.0,0.7,cn,sz=40,b=True,c=HELP)
    tb(s,x+1.30,y+0.95,3.0,0.35,py,sz=14,c=GRAY)
    tb(s,x+1.30,y+1.30,3.0,0.40,en,sz=14,b=True,c=DARK)
n+=1; pn(s,n)

# 14. 我会写 — 谢 (thank)
s=ns(); bg(s,CREAM); hb(s,"✏️ 我会写 · 谢  I Can Write · 'thanks'",HELP)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.4),Inches(0.95),Inches(4.5),Inches(3.7))
sh.fill.solid(); sh.fill.fore_color.rgb=WARM; sh.line.color.rgb=HELP; sh.line.width=Pt(3)
tb(s,0.4,1.30,4.5,2.5,"谢",sz=300,b=True,c=HELP,a=PP_ALIGN.CENTER)
tb(s,0.4,4.10,4.5,0.40,"xiè · thank",sz=18,c=GRAY,a=PP_ALIGN.CENTER)
panel=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(5.05),Inches(0.95),Inches(4.65),Inches(3.7))
panel.fill.solid(); panel.fill.fore_color.rgb=WHITE; panel.line.color.rgb=HELP; panel.line.width=Pt(2)
tb(s,5.20,1.10,4.4,0.5,"📝 笔顺  12 笔",sz=18,b=True,c=HELP)
tb(s,5.20,1.65,4.4,0.4,"左边: 「言」(speak)",sz=14,c=DARK)
tb(s,5.20,2.05,4.4,0.4,"右边: 「射」(shoot)",sz=14,c=DARK)
tb(s,5.20,2.45,4.4,0.4,"用「话」表达感激!",sz=14,b=True,c=HELP)
tb(s,5.20,2.95,4.4,0.35,"📝 练习 Practice:",sz=14,b=True,c=HELP)
tb(s,5.20,3.35,4.4,0.35,"1️⃣ 空中写 · 2️⃣ 手心写 · 3️⃣ 纸上 3 次",sz=13,c=DARK)
sentence_frame_bar(s,4.85,"谢 谢 你 ___ 。","Thank you ___ .")
n+=1; pn(s,n)

# 15. SESSION 3 DIVIDER — Thank-you card
s=div("Session 3  下午","💌 我想感谢的人 Thank-You Card",HEART,"💌"); n+=1; pn(s,n)

# 16. PROJECT INTRO — Thank-you card
s=ns(); bg(s,CREAM); hb(s,"💌 我想感谢的人  My Thank-You Card",HEART)
tb(s,0.4,0.85,9.2,0.5,"❤️ 想一想 — 是谁帮过你？做一张感谢卡!",sz=20,b=True,c=HEART,a=PP_ALIGN.CENTER)
tb(s,0.4,1.45,9.2,0.34,"Think — who has helped YOU? Make them a thank-you card.",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(2.5),Inches(1.90),Inches(5.0),Inches(2.85))
sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=HEART; sh.line.width=Pt(3)
tb(s,2.5,2.00,5.0,0.45,"💌 我的感谢卡",sz=18,b=True,c=HEART,a=PP_ALIGN.CENTER)
tb(s,2.5,2.45,5.0,0.30,"My Thank-You Card",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
tb(s,2.7,2.90,4.6,0.40,"1. 我想感谢: ___",sz=13,b=True,c=DARK)
tb(s,2.7,3.30,4.6,0.40,"2. 你 (他/她) 帮过我: ___",sz=13,b=True,c=DARK)
tb(s,2.7,3.70,4.6,0.40,"3. 让我感觉: ___",sz=13,b=True,c=DARK)
tb(s,2.7,4.10,4.6,0.40,"4. 我想说: 谢谢你 ___ !",sz=13,b=True,c=DARK)
tb(s,0.4,4.85,9.2,0.32,"⏱️ 25 分钟做卡 + 5 分钟分享 + 真送出去!",sz=13,b=True,c=BROWN,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)

# 17. STEP 1 — 想 (who helped me)
s=ns(); bg(s,CREAM); hb(s,"1️⃣ 想一想  Step 1 · Who Helped Me?",HEART)
tb(s,0.4,0.85,9.2,0.4,"想 1 个真正帮过你的人 — 不一定是名人",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.30,"Think of ONE real helper in your life — not famous!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
prompts=[
    ("👨‍⚕️","医生 / 护士","Doctor / Nurse","你生病时照顾你"),
    ("👨‍🏫","老师","Teacher","教你新东西"),
    ("👵","爷爷奶奶","Grandparents","做饭、讲故事"),
    ("🧑‍🚒","消防员 / 警察","Firefighter / Police","让你安全"),
    ("👨‍👩‍👧","爸爸妈妈","Parents","每天照顾你"),
    ("🧑‍🤝‍🧑","好朋友","Friend","和你一起玩"),
]
for i,(em,cn,en,ex) in enumerate(prompts):
    col=i%3; row=i//3
    x=0.3+col*3.2; y=1.65+row*1.55
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(3.0),Inches(1.40))
    sh.fill.solid(); sh.fill.fore_color.rgb=WARM; sh.line.color.rgb=HEART; sh.line.width=Pt(2)
    tb(s,x+0.10,y+0.05,0.6,0.55,em,sz=24,a=PP_ALIGN.CENTER)
    tb(s,x+0.75,y+0.10,2.2,0.35,cn,sz=14,b=True,c=HEART)
    tb(s,x+0.75,y+0.42,2.2,0.30,en,sz=10,c=GRAY)
    tb(s,x+0.10,y+0.83,2.85,0.5,f"💭 {ex}",sz=12,c=DARK)
sentence_frame_bar(s,4.95,"我想感谢 ___ 。","I want to thank ___.")
n+=1; pn(s,n)

# 18. STEP 2 — Make the card
s=ns(); bg(s,CREAM); hb(s,"2️⃣ 做卡  Step 2 · Make the Card",HEART)
tb(s,0.4,0.85,9.2,0.4,"4 步做一张漂亮的卡",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.30,"4 steps to a beautiful card",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
steps=[
    ("📐","折","Fold","对折一张彩纸"),
    ("✏️","写","Write","写「谢谢你 ___」"),
    ("🎨","画","Draw","画 ta + 你"),
    ("✨","装饰","Decorate","贴贴纸 / 闪粉"),
]
for i,(em,cn,en,desc) in enumerate(steps):
    x=0.4+i*2.30
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.65),Inches(2.20),Inches(2.85))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=HEART; sh.line.width=Pt(2.5)
    tb(s,x+0.05,1.80,2.10,0.7,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.65,2.10,0.45,cn,sz=20,b=True,c=HEART,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.10,2.10,0.30,en,sz=10,c=GRAY,a=PP_ALIGN.CENTER)
    tb(s,x+0.10,3.50,2.00,0.85,desc,sz=12,c=DARK,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.65,"谢谢你 ___ , 因为 ___ 。","Thank you ___ because ___.")
n+=1; pn(s,n)

# 19. STEP 3 — Share + deliver
s=ns(); bg(s,CREAM); hb(s,"3️⃣ 分享 + 送出去  Step 3 · Share & Deliver",HEART)
tb(s,0.4,0.85,9.2,0.4,"读给同学听 — 然后真的送出去!",sz=18,b=True,c=DARK,a=PP_ALIGN.CENTER)
tb(s,0.4,1.30,9.2,0.30,"Read aloud — then really deliver it!",sz=11,c=GRAY,a=PP_ALIGN.CENTER)
steps=[
    ("📣","Read","站起来读 30 秒"),
    ("👏","Cheer","全班拍手"),
    ("📸","Photo","老师拍合影"),
    ("🚶","Deliver","回家送给 ta!"),
]
for i,(em,cn,desc) in enumerate(steps):
    x=0.4+i*2.30
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.65),Inches(2.20),Inches(2.85))
    sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=HEART; sh.line.width=Pt(2.5)
    tb(s,x+0.05,1.80,2.10,0.7,em,sz=44,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,2.65,2.10,0.45,cn,sz=20,b=True,c=HEART,a=PP_ALIGN.CENTER)
    tb(s,x+0.05,3.30,2.10,1.10,desc,sz=12,c=DARK,a=PP_ALIGN.CENTER)
sentence_frame_bar(s,4.65,"我送给 ___ 一张感谢卡。","I gave ___ a thank-you card.")
n+=1; pn(s,n)

# 20. CLOSING BADGE
s=ns(); bg(s,CREAM)
tb(s,0.5,0.4,9,0.8,"🎖️ Day 4 徽章  Helper Badge",sz=26,b=True,c=HELP,a=PP_ALIGN.CENTER)
sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(3.5),Inches(1.4),Inches(3),Inches(3))
sh.fill.solid(); sh.fill.fore_color.rgb=WHITE; sh.line.color.rgb=HEART; sh.line.width=Pt(6)
tf=tb(s,3.6,1.65,2.8,2.7,"DAY 4",sz=18,b=True,c=HEART,a=PP_ALIGN.CENTER)
ap(tf,"❤️",sz=42,a=PP_ALIGN.CENTER)
ap(tf,"帮助别人",sz=18,b=True,c=HELP,a=PP_ALIGN.CENTER)
ap(tf,"✓ COMPLETED",sz=12,b=True,c=OK,a=PP_ALIGN.CENTER)
ap(tf,"🩺🕯️📚✋",sz=16,a=PP_ALIGN.CENTER)
tb(s,1,4.55,8,0.4,"🎉 你也是一位小帮手! You are a Helper too!",sz=15,b=True,c=HELP,a=PP_ALIGN.CENTER)
tb(s,1,5.0,8,0.4,"明天 Day 5 — AI 与未来 · OpenAI · DeepMind · Anthropic",sz=12,c=GRAY,a=PP_ALIGN.CENTER)
n+=1; pn(s,n)

out=os.path.join(os.path.dirname(__file__),"day4_helpers.pptx")
prs.save(out); print(f"Saved {out}  ({n} slides)")
