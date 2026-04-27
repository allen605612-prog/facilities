#!/usr/bin/env python3
"""產生得獎海報，仿照參考圖風格（深藍絲絨背景＋金色星星＋彩帶）"""
import math, random
from PIL import Image, ImageDraw, ImageFont, ImageFilter
import os

W, H = 800, 1200
OUT = r"C:\Users\user\allen\poster_award.png"

# ── 字型 ───────────────────────────────────────────────────
FONT_PATHS = [
    r"C:\Windows\Fonts\kaiu.ttf",      # 標楷體
    r"C:\Windows\Fonts\msjh.ttc",      # 微軟正黑體
    r"C:\Windows\Fonts\mingliu.ttc",   # 細明體
]
def load_font(size, index=0):
    for fp in FONT_PATHS:
        try:
            return ImageFont.truetype(fp, size)
        except Exception:
            continue
    return ImageFont.load_default()

# ── 顏色 ───────────────────────────────────────────────────
NAVY       = (10, 18, 40)
NAVY_MID   = (18, 30, 65)
GOLD       = (212, 175, 55)
GOLD_LIGHT = (255, 223, 100)
GOLD_DARK  = (160, 120, 20)
WHITE      = (255, 255, 255)
OFF_WHITE  = (240, 230, 200)

rng = random.Random(42)

img = Image.new("RGB", (W, H), NAVY)
draw = ImageDraw.Draw(img)

# ── 1. 布幕漸層背景 ───────────────────────────────────────
for y in range(H):
    t = y / H
    r = int(NAVY[0] + (NAVY_MID[0]-NAVY[0]) * math.sin(t * math.pi))
    g = int(NAVY[1] + (NAVY_MID[1]-NAVY[1]) * math.sin(t * math.pi))
    b = int(NAVY[2] + (NAVY_MID[2]-NAVY[2]) * math.sin(t * math.pi))
    draw.line([(0, y), (W, y)], fill=(r, g, b))

# ── 2. 布幕皺褶條紋 ───────────────────────────────────────
for i in range(16):
    cx = int(W * i / 15)
    bright = 30 if i % 2 == 0 else -15
    for x in range(max(0, cx-30), min(W, cx+30)):
        alpha = 1 - abs(x - cx) / 30
        fade = int(bright * alpha)
        for y in range(H):
            px = img.getpixel((x, y))
            img.putpixel((x, y), (
                max(0, min(255, px[0]+fade)),
                max(0, min(255, px[1]+fade)),
                max(0, min(255, px[2]+fade)),
            ))

# ── 3. 彩帶與亮粉 ─────────────────────────────────────────
def draw_confetti(draw, n=120):
    colors = [GOLD, GOLD_LIGHT, (255,200,50), (255,240,150), (200,150,30)]
    for _ in range(n):
        x = rng.randint(0, W)
        y = rng.randint(int(H*0.1), int(H*0.85))
        c = rng.choice(colors)
        kind = rng.randint(0, 2)
        if kind == 0:   # 小矩形彩帶
            ang = rng.uniform(-60, 60)
            l, w2 = rng.randint(6,18), rng.randint(2,5)
            pts = [(-l/2,-w2/2),(l/2,-w2/2),(l/2,w2/2),(-l/2,w2/2)]
            rad = math.radians(ang)
            cos_a, sin_a = math.cos(rad), math.sin(rad)
            rot = [(x + p[0]*cos_a - p[1]*sin_a,
                    y + p[0]*sin_a + p[1]*cos_a) for p in pts]
            draw.polygon(rot, fill=c)
        elif kind == 1: # 圓點
            r2 = rng.randint(2,5)
            draw.ellipse([x-r2,y-r2,x+r2,y+r2], fill=c)
        else:           # 細線
            x2 = x + rng.randint(-15,15)
            y2 = y + rng.randint(5,20)
            draw.line([x,y,x2,y2], fill=c, width=2)

draw_confetti(draw)

# ── 4. 金色星星（大中小） ─────────────────────────────────
def star_polygon(cx, cy, r_out, r_in, n=5, angle_offset=-90):
    pts = []
    for i in range(n*2):
        angle = math.radians(angle_offset + i * 180/n)
        r = r_out if i % 2 == 0 else r_in
        pts.append((cx + r*math.cos(angle), cy + r*math.sin(angle)))
    return pts

def draw_star_3d(draw, cx, cy, r_out, r_in):
    outer = star_polygon(cx, cy, r_out, r_in)
    # 底色深金
    draw.polygon(outer, fill=GOLD_DARK, outline=GOLD_DARK)
    # 漸層亮面（左上半顆星）
    inner = star_polygon(cx-r_out*0.06, cy-r_out*0.06, r_out*0.85, r_in*0.85)
    draw.polygon(inner, fill=GOLD)
    # 高光
    hi = star_polygon(cx-r_out*0.12, cy-r_out*0.12, r_out*0.55, r_in*0.55)
    draw.polygon(hi, fill=GOLD_LIGHT)

# 三顆星：左小、右中、中大（整體下移 150px）
draw_star_3d(draw, 290, 305, 55, 22)   # 左
draw_star_3d(draw, 510, 310, 60, 24)   # 右
draw_star_3d(draw, 400, 245, 80, 32)   # 中（最大，最高）

# ── 5. 文字 ──────────────────────────────────────────────
def centered_text(draw, text, y, font, fill):
    bb = draw.textbbox((0,0), text, font=font)
    tw = bb[2]-bb[0]
    draw.text(((W-tw)//2, y), text, font=font, fill=fill)

def shadow_text(draw, text, y, font, fill, shadow=(0,0,0)):
    bb = draw.textbbox((0,0), text, font=font)
    tw = bb[2]-bb[0]
    x = (W-tw)//2
    draw.text((x+2, y+2), text, font=font, fill=shadow)
    draw.text((x, y), text, font=font, fill=fill)

f_crazy  = load_font(54)
f_names  = load_font(36)
f_achiev = load_font(56)
f_thanks = load_font(36)

# 「狂賀」
shadow_text(draw, "狂 賀", 390, f_crazy, GOLD_LIGHT)

# 學生姓名（兩行）
centered_text(draw, "高一心班宋彥霖、高一意班鐘宥昕", 468, f_names, OFF_WHITE)
centered_text(draw, "高一正班廖柃柃  同學", 522, f_names, OFF_WHITE)

# 分隔細線
draw.line([(80, 580), (W-80, 580)], fill=GOLD_DARK, width=1)

# 主要成就（金色大字）
shadow_text(draw, "榮獲第 66 屆第四區", 598, f_achiev, GOLD_LIGHT, shadow=(80,60,0))
shadow_text(draw, "分區高中科展", 668, f_achiev, GOLD_LIGHT, shadow=(80,60,0))
shadow_text(draw, "「環境學科」優等！", 738, f_achiev, GOLD_LIGHT, shadow=(80,60,0))

# ── 6. 金色橫條 ───────────────────────────────────────────
bar_y = 840
draw.rectangle([0, bar_y, W, bar_y+56], fill=GOLD_DARK)
draw.rectangle([0, bar_y+4, W, bar_y+52], fill=GOLD)
# 橫條亮邊
draw.line([(0,bar_y+4),(W,bar_y+4)], fill=GOLD_LIGHT, width=3)
draw.line([(0,bar_y+52),(W,bar_y+52)], fill=GOLD_DARK, width=2)

# ── 7. 底部感謝文字 ───────────────────────────────────────
centered_text(draw, "感謝 張祐誠 老師指導", 934, f_thanks, WHITE)

# ── 8. 輕微暗角 ───────────────────────────────────────────
vignette = Image.new("L", (W, H), 255)
vd = ImageDraw.Draw(vignette)
for i in range(120):
    t = i / 120
    alpha = int(255 * t)
    vd.rectangle([i, i, W-i, H-i], outline=alpha)
img_v = Image.new("RGB", (W, H), (0,0,0))
img = Image.composite(img, img_v, vignette)

img.save(OUT, "PNG")
print("OK:", OUT)
