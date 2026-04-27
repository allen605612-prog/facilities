#!/usr/bin/env python3
"""第66屆雲林縣國民中小學科展 - 批次產生7張得獎海報"""
import math, random, os
from PIL import Image, ImageDraw, ImageFont

W, H = 800, 1200
OUT_DIR = r"C:\Users\user\allen"

FONT_PATHS = [
    r"C:\Windows\Fonts\kaiu.ttf",
    r"C:\Windows\Fonts\msjh.ttc",
    r"C:\Windows\Fonts\mingliu.ttc",
]

def load_font(size):
    for fp in FONT_PATHS:
        try:
            return ImageFont.truetype(fp, size)
        except Exception:
            continue
    return ImageFont.load_default()

NAVY      = (10, 18, 40)
NAVY_MID  = (18, 30, 65)
GOLD      = (212, 175, 55)
GOLD_LIGHT= (255, 223, 100)
GOLD_DARK = (160, 120, 20)
WHITE     = (255, 255, 255)
OFF_WHITE = (240, 230, 200)

AWARDS = [
    {"subject": "數學科",
     "rank": "第一名",
     "names": [("美二甲", ["廖采玲", "陳彙喬", "黃柏之"])],
     "teacher": "羅婉萍", "out": "poster_66_01_數學科第一名.png"},
    {"subject": "數學科",
     "rank": "第二名",
     "names": [("美一乙", ["廖泱泱"])],
     "teacher": "詩富強", "out": "poster_66_02_數學科第二名.png"},
    {"subject": "生物科",
     "rank": "第二名",
     "names": [("音二甲", ["洪靖宇", "蔡承恩"])],
     "teacher": "許議鶴", "out": "poster_66_03_生物科第二名.png"},
    {"subject": "生物科",
     "rank": "第三名",
     "names": [("美二乙", ["黃家恩"]), ("美二甲", ["王嘉裕"]), ("音二甲", ["陳昱勳"])],
     "teacher": "郭靜如", "out": "poster_66_04_生物科第三名.png"},
    {"subject": "生活與應用科學科(一)",
     "rank": "第四名",
     "names": [("初一誠", ["陳沛青"]), ("美一甲", ["鍾子鵬"]), ("美一乙", ["陳意凡"])],
     "teacher": "林俊杰", "out": "poster_66_05_生活應用科學一第四名.png"},
    {"subject": "生活與應用科學科(二)",
     "rank": "第三名",
     "names": [("音二甲", ["籃苡熏", "萬奕岑", "陳姸汝"])],
     "teacher": "許議鶴", "out": "poster_66_06_生活應用科學二第三名.png"},
    {"subject": "生活與應用科學科(三)",
     "rank": "第四名",
     "names": [("初二正", ["許呈睿", "黃品語", "洪啟鈞"])],
     "teacher": "許議鶴", "out": "poster_66_07_生活應用科學三第四名.png"},
]

def format_name_lines(names_data):
    """回傳 1 或 2 行顯示用字串，末行附「同學」"""
    parts = []
    for cls, names in names_data:
        parts.append(cls + " " + "、".join(names))
    full = "、".join(parts)
    if len(full) <= 18:
        return [full + " 同學"]
    mid = len(parts) // 2 + (1 if len(parts) % 2 else 0)
    return ["、".join(parts[:mid]), "、".join(parts[mid:]) + " 同學"]

def tw(draw, text, font):
    bb = draw.textbbox((0, 0), text, font=font)
    return bb[2] - bb[0]

def centered_text(draw, text, y, font, fill):
    x = (W - tw(draw, text, font)) // 2
    draw.text((x, y), text, font=font, fill=fill)

def shadow_text(draw, text, y, font, fill, shd=(0, 0, 0)):
    x = (W - tw(draw, text, font)) // 2
    draw.text((x+2, y+2), text, font=font, fill=shd)
    draw.text((x, y),   text, font=font, fill=fill)

def draw_bg(img, draw):
    for y in range(H):
        t = y / H
        c = tuple(int(NAVY[i] + (NAVY_MID[i]-NAVY[i]) * math.sin(t*math.pi)) for i in range(3))
        draw.line([(0, y), (W, y)], fill=c)
    for i in range(16):
        cx = int(W * i / 15)
        bright = 30 if i % 2 == 0 else -15
        for x in range(max(0, cx-30), min(W, cx+30)):
            fade = int(bright * (1 - abs(x-cx)/30))
            for y2 in range(H):
                px = img.getpixel((x, y2))
                img.putpixel((x, y2), tuple(max(0, min(255, px[j]+fade)) for j in range(3)))

def draw_confetti(draw, seed):
    rng = random.Random(seed)
    colors = [GOLD, GOLD_LIGHT, (255,200,50), (255,240,150), (200,150,30)]
    for _ in range(120):
        x = rng.randint(0, W)
        y = rng.randint(int(H*.1), int(H*.85))
        c = rng.choice(colors)
        kind = rng.randint(0, 2)
        if kind == 0:
            ang = rng.uniform(-60, 60)
            l, w2 = rng.randint(6,18), rng.randint(2,5)
            pts = [(-l/2,-w2/2),(l/2,-w2/2),(l/2,w2/2),(-l/2,w2/2)]
            ca, sa = math.cos(math.radians(ang)), math.sin(math.radians(ang))
            rot = [(x+p[0]*ca-p[1]*sa, y+p[0]*sa+p[1]*ca) for p in pts]
            draw.polygon(rot, fill=c)
        elif kind == 1:
            r2 = rng.randint(2, 5)
            draw.ellipse([x-r2,y-r2,x+r2,y+r2], fill=c)
        else:
            draw.line([x, y, x+rng.randint(-15,15), y+rng.randint(5,20)], fill=c, width=2)

def star_pts(cx, cy, ro, ri, n=5, off=-90):
    return [(cx+(ro if i%2==0 else ri)*math.cos(math.radians(off+i*180/n)),
             cy+(ro if i%2==0 else ri)*math.sin(math.radians(off+i*180/n)))
            for i in range(n*2)]

def draw_star_3d(draw, cx, cy, ro, ri):
    draw.polygon(star_pts(cx, cy, ro, ri), fill=GOLD_DARK)
    draw.polygon(star_pts(cx-ro*.06, cy-ro*.06, ro*.85, ri*.85), fill=GOLD)
    draw.polygon(star_pts(cx-ro*.12, cy-ro*.12, ro*.55, ri*.55), fill=GOLD_LIGHT)

def make_poster(award, seed):
    img = Image.new("RGB", (W, H), NAVY)
    draw = ImageDraw.Draw(img)
    draw_bg(img, draw)
    draw_confetti(draw, seed)

    # 星星
    draw_star_3d(draw, 290, 305, 55, 22)
    draw_star_3d(draw, 510, 310, 60, 24)
    draw_star_3d(draw, 400, 245, 80, 32)

    f_crazy  = load_font(54)
    f_names  = load_font(36)
    f_achiev = load_font(56)
    f_thanks = load_font(36)

    shadow_text(draw, "狂 賀", 390, f_crazy, GOLD_LIGHT)

    name_lines = format_name_lines(award["names"])
    y_n = 468
    for line in name_lines:
        centered_text(draw, line, y_n, f_names, OFF_WHITE)
        y_n += 54

    y_div = y_n + 4
    draw.line([(80, y_div), (W-80, y_div)], fill=GOLD_DARK, width=1)

    # 成就文字
    subj = award["subject"]
    rank = award["rank"]
    MAX_W = W - 80
    line_a1 = "榮獲第 66 屆雲林縣"
    line_a2 = "國民中小學科展"
    line_a3 = f"「{subj}」{rank}！"

    if tw(draw, line_a3, f_achiev) <= MAX_W:
        # 3行排版
        y_a = y_div + 148
        lh = 70
        shadow_text(draw, line_a1,  y_a,     f_achiev, GOLD_LIGHT, (80,60,0))
        shadow_text(draw, line_a2,  y_a+lh,  f_achiev, GOLD_LIGHT, (80,60,0))
        shadow_text(draw, line_a3,  y_a+lh*2,f_achiev, GOLD_LIGHT, (80,60,0))
        y_last = y_a + lh*2
    else:
        # 4行排版（科目名稱太長，拆成兩行）
        y_a = y_div + 100
        lh = 63
        subj_line = f"「{subj}」"
        f_subj = f_achiev
        for sz in range(56, 27, -1):
            f_t = load_font(sz)
            if tw(draw, subj_line, f_t) <= MAX_W:
                f_subj = f_t
                break
        shadow_text(draw, line_a1,    y_a,      f_achiev, GOLD_LIGHT, (80,60,0))
        shadow_text(draw, line_a2,    y_a+lh,   f_achiev, GOLD_LIGHT, (80,60,0))
        shadow_text(draw, subj_line,  y_a+lh*2, f_subj,   GOLD_LIGHT, (80,60,0))
        shadow_text(draw, f"{rank}！", y_a+lh*3, f_achiev, GOLD_LIGHT, (80,60,0))
        y_last = y_a + lh*3

    # 金色橫條
    bar_y = min(y_last + 95, H - 198)
    draw.rectangle([0, bar_y,   W, bar_y+56], fill=GOLD_DARK)
    draw.rectangle([0, bar_y+4, W, bar_y+52], fill=GOLD)
    draw.line([(0, bar_y+4),  (W, bar_y+4)],  fill=GOLD_LIGHT, width=3)
    draw.line([(0, bar_y+52), (W, bar_y+52)], fill=GOLD_DARK,  width=2)

    centered_text(draw, f"感謝 {award['teacher']} 老師指導", bar_y+88, f_thanks, WHITE)

    # 暗角
    vig = Image.new("L", (W, H), 255)
    vd = ImageDraw.Draw(vig)
    for i in range(120):
        vd.rectangle([i, i, W-i, H-i], outline=int(255*i/120))
    img = Image.composite(img, Image.new("RGB", (W, H), 0), vig)

    out_path = os.path.join(OUT_DIR, award["out"])
    img.save(out_path, "PNG")
    print("OK:", award["out"])

for i, award in enumerate(AWARDS):
    make_poster(award, seed=42+i)
