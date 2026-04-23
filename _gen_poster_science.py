#!/usr/bin/env python3
"""
第 66 屆第四區分區高中科展 — 榮譽海報
Verdant Triumph design philosophy — full-height layout
"""
import math
from PIL import Image, ImageDraw, ImageFont

# ── Canvas ─────────────────────────────────────────────────────────────
W, H = 2480, 3508          # A4 @ 300 dpi
img  = Image.new("RGBA", (W, H), (0, 0, 0, 255))
draw = ImageDraw.Draw(img)

# ── Palette ────────────────────────────────────────────────────────────
BG_DARK    = (11,  38,  24)
BG_MID     = (16,  52,  33)
GOLD       = (212, 175,  55)
GOLD_LT    = (255, 223,  96)
GOLD_DIM   = (140, 108,  22)
CREAM      = (255, 250, 215)
GRN_LEAF   = (22,  78,  44)
GRN_VEIN   = (44, 120,  68)

# ── Font paths ──────────────────────────────────────────────────────────
CDIR  = r"C:\Users\user\.claude\plugins\cache\anthropic-agent-skills\document-skills\b0cbd3df1533\skills\canvas-design\canvas-fonts"
WDIR  = r"C:\Windows\Fonts"

def ttf(path, size):
    try:    return ImageFont.truetype(path, size)
    except: return ImageFont.load_default()

f_xl   = ttf(rf"{WDIR}\msjhbd.ttf",  280)   # 狂賀！
f_lg   = ttf(rf"{WDIR}\msjhbd.ttf",  104)   # 副標 / 得獎欄
f_md   = ttf(rf"{WDIR}\msjhbd.ttf",   82)   # 老師姓名 / 指導標籤
f_sm   = ttf(rf"{WDIR}\msjh.ttf",     66)   # 學生姓名
f_xs   = ttf(rf"{WDIR}\msjh.ttf",     50)   # 標籤 / 落款
f_num  = ttf(rf"{CDIR}\CrimsonPro-Bold.ttf", 520)   # 大數字 66
f_lat  = ttf(rf"{CDIR}\InstrumentSans-Regular.ttf",  52)
f_latb = ttf(rf"{CDIR}\InstrumentSans-Bold.ttf",     58)

# ── Helpers ────────────────────────────────────────────────────────────
def center_x(text, font, y, fill, draw_obj=None):
    d = draw_obj or draw
    bb = d.textbbox((0,0), text, font=font)
    x  = (W - (bb[2]-bb[0])) // 2
    d.text((x, y), text, font=font, fill=fill)
    return bb[3]-bb[1]   # height

def hline(y, x0, x1, c, w=2):
    draw.line([(x0,y),(x1,y)], fill=c, width=w)

def gold_rule(y, w1=2, w2=1):
    hline(y,        220, W-220, GOLD,    w1)
    hline(y+9,      260, W-260, GOLD_DIM, w2)

def diamond(cx, cy, r=28):
    draw.polygon([(cx,cy-r),(cx+r*3//5,cy),(cx,cy+r),(cx-r*3//5,cy)], fill=GOLD)
    draw.polygon([(cx,cy-r//2),(cx+r//4,cy),(cx,cy+r//2),(cx-r//4,cy)], fill=BG_DARK)

# ── Background ─────────────────────────────────────────────────────────
draw.rectangle([0,0,W,H], fill=BG_DARK)
# Subtle inner glow
for i in range(50):
    v = int(BG_MID[1] * (1 - i/50))
    draw.rectangle([130+i, 130+i, W-130-i, H-130-i],
                   outline=(BG_MID[0], v, BG_MID[2]), width=1)

# ── Botanical leaf helper ───────────────────────────────────────────────
def leaf(cx_, cy_, length, ang_deg, alpha=30):
    ang   = math.radians(ang_deg)
    a, b  = length//2, length//7
    over  = Image.new("RGBA", (W,H), (0,0,0,0))
    od    = ImageDraw.Draw(over)
    pts   = []
    for t in range(360):
        r  = math.radians(t)
        lx = a*math.cos(r); ly = b*math.sin(r)
        rx = lx*math.cos(ang)-ly*math.sin(ang)
        ry = lx*math.sin(ang)+ly*math.cos(ang)
        pts.append((cx_+rx, cy_+ry))
    al = int(255*alpha/100)
    od.polygon(pts, fill=(*GRN_LEAF, al))
    tip  = (cx_+a*math.cos(ang), cy_+a*math.sin(ang))
    base = (cx_-a*math.cos(ang), cy_-a*math.sin(ang))
    od.line([base, tip], fill=(*GRN_VEIN, min(255,al+50)), width=max(2,length//65))
    for k in range(1,6):
        t2 = k/6
        mx = base[0]+t2*(tip[0]-base[0]); my = base[1]+t2*(tip[1]-base[1])
        vl = b*.85*math.sin(math.pi*t2)
        p  = ang+math.pi/2
        for s in (1,-1):
            od.line([(mx,my),(mx+s*vl*math.cos(p),my+s*vl*math.sin(p))],
                    fill=(*GRN_VEIN, int(al*.7)), width=max(1,length//110))
    base_img = img.convert("RGBA")
    img.paste(Image.alpha_composite(base_img, over).convert("RGBA"))

# Draw background leaves (4 corner clusters)
specs = [
    (290,290,460,-20),(190,380,340, 18),(410,180,300,-52),(200,200,260, 48),(460,360,220,-8),
    (W-290,290,460,200),(W-190,380,340,162),(W-410,180,300,232),(W-200,200,260,132),(W-460,360,220,188),
    (290,H-290,460, 22),(190,H-380,340,-18),(410,H-180,300, 52),(200,H-200,260,-48),
    (W-290,H-290,460,158),(W-190,H-380,340,198),(W-410,H-180,300,128),(W-200,H-200,260,228),
]
for s in specs:
    leaf(*s, alpha=22)

# Refresh draw after paste operations
draw = ImageDraw.Draw(img)

# ── Triple border ───────────────────────────────────────────────────────
for m,w,c in [(58,6,GOLD),(78,2,GOLD_DIM),(104,4,GOLD)]:
    draw.rectangle([m,m,W-m,H-m], outline=c, width=w)

# Corner ornaments
for cx_,cy_ in [(108,108),(W-108,108),(108,H-108),(W-108,H-108)]:
    diamond(cx_,cy_,28)
    diamond(cx_,cy_,10)

# ═══════════════════════════════════════════════════════════════════════
# CONTENT  —  full-height layout
# ═══════════════════════════════════════════════════════════════════════

# 1. BADGE ──────────────────────────────────────────────── y~180
BY, BW, BHt = 188, 440, 90
BX = (W-BW)//2
draw.rectangle([BX,BY,BX+BW,BY+BHt], fill=GOLD, outline=GOLD_DIM, width=3)
center_x("— 榮  譽  榜 —", f_xs, BY+18, BG_DARK)

# Thin rules below badge
hline(BY+BHt+22, 200, W-200, GOLD,    2)
hline(BY+BHt+32, 240, W-240, GOLD_DIM,1)

# 2. MAIN TITLE ─────────────────────────────────────────── y~380
TY = 380
bb = draw.textbbox((0,0),"狂賀！",font=f_xl)
tw = bb[2]-bb[0]; x0=(W-tw)//2
draw.text((x0+7,TY+7), "狂賀！", font=f_xl, fill=GOLD_DIM)
draw.text((x0,  TY),   "狂賀！", font=f_xl, fill=GOLD_LT)
th = bb[3]-bb[1]

# flanking dots
my = TY + th//2
for sign,gap in [(-1,430),(1,430)]:
    dx = W//2+sign*gap
    for r,fc in [(22,GOLD),(12,BG_DARK),(6,GOLD)]:
        draw.ellipse([dx-r,my-r,dx+r,my+r], fill=fc)

# 3. EXHIBITION LINE ────────────────────────────────────── y~750
EY = TY + th + 60
center_x("第  66  屆  第四區分區高中科展", f_lg, EY, CREAM)
bb2 = draw.textbbox((0,0),"第  66  屆  第四區分區高中科展", font=f_lg)
eh  = bb2[3]-bb2[1]

# 4. AWARD BAND ─────────────────────────────────────────── y~920
AY = EY + eh + 50
AH = 116
draw.rectangle([160,AY,W-160,AY+AH], fill=GOLD, outline=GOLD_DIM, width=3)
at = "榮  獲  「 環  境  學  科 」  優  等"
bba = draw.textbbox((0,0),at,font=f_lg)
center_x(at, f_lg, AY+(AH-(bba[3]-bba[1]))//2, BG_DARK)

# 5. DIVIDER ────────────────────────────────────────────── y~1090
DY1 = AY+AH+70
gold_rule(DY1)
diamond(W//2, DY1+5, 32)

# 6. ADVISOR ────────────────────────────────────────────── y~1180
ADY = DY1+80
center_x("指  導  老  師", f_md, ADY, GOLD)
adh = draw.textbbox((0,0),"指  導  老  師",font=f_md)[3]
ADY2 = ADY+adh-draw.textbbox((0,0),"指  導  老  師",font=f_md)[1]+50

# Advisor name
anm = "張  祐  誠  老  師"
bbn = draw.textbbox((0,0),anm,font=f_lg)
nw  = bbn[2]-bbn[0]; nh=bbn[3]-bbn[1]
nx  = (W-nw)//2
draw.text((nx,ADY2), anm, font=f_lg, fill=CREAM)
uy  = ADY2+nh+14
hline(uy,   nx,       nx+nw,       GOLD_DIM,2)
hline(uy+7, nx+nw//4, nx+nw*3//4,  GOLD,    1)

# 7. DIVIDER ────────────────────────────────────────────── y~1500
DY2 = uy+70
hline(DY2,   240, W-240, GOLD_DIM,1)
hline(DY2+6, 240, W-240, GOLD_DIM,1)

# 8. STUDENTS ───────────────────────────────────────────── y~1580
STY = DY2+60
center_x("得  獎  學  生", f_md, STY, GOLD)
bbs = draw.textbbox((0,0),"得  獎  學  生",font=f_md)
STY2= STY+(bbs[3]-bbs[1])+70

students = [
    ("高一心班","宋  彥  霖"),
    ("高一意班","鐘  宥  昕"),
    ("高一正班","廖  柃  柃"),
]
row_gap = 138
for cls, name in students:
    bc = draw.textbbox((0,0),cls,  font=f_xs)
    bn = draw.textbbox((0,0),name, font=f_sm)
    cw = bc[2]-bc[0]; ch=bc[3]-bc[1]
    nw = bn[2]-bn[0]; nh=bn[3]-bn[1]
    gap= 44
    total = cw+gap+nw
    xs  = (W-total)//2
    # tag box
    draw.rectangle([xs-12, STY2-6, xs+cw+12, STY2+ch+10], outline=GOLD_DIM, width=2)
    draw.text((xs, STY2+(nh-ch)//2+4), cls,  font=f_xs, fill=GOLD)
    draw.text((xs+cw+gap, STY2),        name, font=f_sm, fill=CREAM)
    STY2 += row_gap

# 9. DIVIDER ────────────────────────────────────────────── bottom
DY3 = STY2 + 50
gold_rule(DY3)
diamond(W//2, DY3+5, 32)

# 10. FOOTER ────────────────────────────────────────────── y~2100
FY = DY3+80
scn = "＊  學  校  名  稱  ＊"
bsc = draw.textbbox((0,0),scn,font=f_md)
sw  = bsc[2]-bsc[0]; sh=bsc[3]-bsc[1]
sx  = (W-sw)//2
draw.rectangle([sx-20,FY-6,sx+sw+20,FY+sh+10], fill=GOLD, outline=GOLD_DIM, width=2)
draw.text((sx,FY), scn, font=f_md, fill=BG_DARK)
FY2 = FY+sh+32

center_x("全  校  師  生  同  賀", f_sm, FY2, CREAM)
bfc = draw.textbbox((0,0),"全  校  師  生  同  賀",font=f_sm)
FY3 = FY2+(bfc[3]-bfc[1])+24
center_x("A.D. 2026", f_lat, FY3, GOLD_DIM)

# 11. LARGE "66" WATERMARK ──────────────────────────────── y~2380
WMY = FY3 + 80
bb66= draw.textbbox((0,0),"66",font=f_num)
w66 = bb66[2]-bb66[0]; h66=bb66[3]-bb66[1]
x66 = (W-w66)//2

# Ghost circle
rc = 300
draw.ellipse([W//2-rc, WMY-30, W//2+rc, WMY+h66+30],
             outline=(*GOLD_DIM, 255), width=3)
draw.ellipse([W//2-rc-16, WMY-46, W//2+rc+16, WMY+h66+46],
             outline=(*GOLD, 255), width=1)

# Semi-transparent "66" via overlay
wm_ov = Image.new("RGBA",(W,H),(0,0,0,0))
wm_d  = ImageDraw.Draw(wm_ov)
wm_d.text((x66, WMY),"66",font=f_num,fill=(*GOLD,70))
base = img.convert("RGBA")
img.paste(Image.alpha_composite(base, wm_ov).convert("RGBA"))
draw = ImageDraw.Draw(img)

# Subtitle below watermark
ENVY = WMY + h66 + 42
center_x("ENVIRONMENTAL  SCIENCE  EXCELLENCE", f_lat, ENVY, GOLD_DIM)

hline(ENVY+60, 420, W-420, GOLD_DIM, 1)

# Bottom motto — very subtle
MOTY = ENVY + 100
center_x("以  知  識  護  大  地  ·  以  榮  耀  回  饋  師  恩", f_xs, MOTY, GOLD_DIM)

# 12. FINAL FOREGROUND LEAF LAYER ──── very subtle, drawn last
draw = ImageDraw.Draw(img)   # refresh
for s in [
    (310,320,380,-28,16),(W-310,320,380,208,16),
    (310,H-320,350, 28,16),(W-310,H-320,350,152,16),
]:
    leaf(*s)
draw = ImageDraw.Draw(img)

# ── Save ───────────────────────────────────────────────────────────────
out_img = img.convert("RGB")
out_png = r"C:\Users\user\allen\第66屆科展榮譽海報.png"
out_pdf = r"C:\Users\user\allen\第66屆科展榮譽海報.pdf"
out_img.save(out_png, "PNG",  dpi=(300,300))
out_img.save(out_pdf, "PDF",  resolution=300)
print("Done.")
print(f"  PNG → {out_png}")
print(f"  PDF → {out_pdf}")
