from PIL import Image, ImageDraw, ImageFont, ImageFilter
import math, random, colorsys

# ── Canvas ───────────────────────────────────────────────────────────────────
W, H = 800, 1130
img = Image.new("RGB", (W, H), (10, 12, 22))
draw = ImageDraw.Draw(img)

# ── Background: dark navy with subtle radial glow ────────────────────────────
bg = Image.new("RGB", (W, H), (8, 10, 20))
bg_draw = ImageDraw.Draw(bg)

# Layered dark gradient from center top
for i in range(H):
    ratio = i / H
    # Very subtle blue tint at top, pure black at bottom
    r = int(8 + 12 * (1 - ratio))
    g = int(10 + 8 * (1 - ratio))
    b = int(20 + 30 * (1 - ratio))
    bg_draw.line([(0, i), (W, i)], fill=(r, g, b))

# Soft radial glow from center
glow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
gd = ImageDraw.Draw(glow)
cx, cy = W // 2, H // 3
for r_ring in range(380, 0, -2):
    alpha = int(18 * (1 - r_ring / 380) ** 2)
    gd.ellipse([cx - r_ring, cy - r_ring, cx + r_ring, cy + r_ring],
               fill=(30, 50, 90, alpha))
bg = Image.alpha_composite(bg.convert("RGBA"), glow).convert("RGB")

# Fabric / drape texture lines (vertical subtle highlight lines)
fab = ImageDraw.Draw(bg)
random.seed(42)
for i in range(0, W, 14):
    x = i + random.randint(-5, 5)
    alpha_val = random.randint(3, 12)
    for y in range(H):
        ratio = abs(math.sin(y * 0.003 + x * 0.05)) * 0.3
        r, g, b = bg.getpixel((min(x, W-1), y))
        r = min(255, r + int(alpha_val * ratio))
        g = min(255, g + int(alpha_val * ratio))
        b = min(255, b + int(alpha_val * ratio))
        if 0 <= x < W:
            bg.putpixel((x, y), (r, g, b))

img = bg.copy()
draw = ImageDraw.Draw(img)

# ── Fonts ─────────────────────────────────────────────────────────────────────
font_path_bold   = "C:/Windows/Fonts/msjhbd.ttc"
font_path_reg    = "C:/Windows/Fonts/msjh.ttc"
font_path_kaiu   = "C:/Windows/Fonts/kaiu.ttf"

font_title  = ImageFont.truetype(font_path_kaiu, 130)   # 恭賀
font_sub    = ImageFont.truetype(font_path_bold, 34)     # main body
font_name   = ImageFont.truetype(font_path_kaiu, 52)     # name highlight
font_footer = ImageFont.truetype(font_path_reg,  26)     # footer
font_label  = ImageFont.truetype(font_path_reg,  18)     # small label

# ── Gold palette ──────────────────────────────────────────────────────────────
GOLD_DEEP    = (180, 130,  30)
GOLD_MID     = (212, 170,  60)
GOLD_BRIGHT  = (255, 215,  80)
GOLD_LIGHT   = (255, 235, 140)
GOLD_SHIMMER = (255, 248, 200)
CREAM        = (255, 245, 180)

def gold_gradient_text(draw, text, font, x, y, align="center", canvas_w=W):
    """Draw text with a vertical gold shimmer gradient."""
    # Create a small layer just for the text
    tmp = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    td = ImageDraw.Draw(tmp)
    bb = td.textbbox((0, 0), text, font=font)
    tw, th = bb[2] - bb[0], bb[3] - bb[1]
    if align == "center":
        tx = (canvas_w - tw) // 2
    elif align == "right":
        tx = canvas_w - tw - x
    else:
        tx = x
    # Draw base text
    td.text((tx, y), text, font=font, fill=GOLD_BRIGHT)

    # Build gradient mask
    grad = Image.new("L", (tw, th))
    gd2 = ImageDraw.Draw(grad)
    for gy in range(th):
        ratio = gy / max(th - 1, 1)
        # shimmer: pale gold top → deep gold mid → shimmer at bottom
        if ratio < 0.3:
            v = int(200 + 55 * (ratio / 0.3))
        elif ratio < 0.6:
            v = 255
        else:
            v = int(255 - 60 * ((ratio - 0.6) / 0.4))
        gd2.line([(0, gy), (tw, gy)], fill=v)

    # Colorize with gold
    gold_layer = Image.new("RGBA", (tw, th))
    glc = ImageDraw.Draw(gold_layer)
    for gy in range(th):
        ratio = gy / max(th - 1, 1)
        r = int(GOLD_DEEP[0] + (GOLD_SHIMMER[0] - GOLD_DEEP[0]) * (math.sin(ratio * math.pi)))
        g2= int(GOLD_DEEP[1] + (GOLD_SHIMMER[1] - GOLD_DEEP[1]) * (math.sin(ratio * math.pi)))
        b = int(GOLD_DEEP[2] + (GOLD_SHIMMER[2] - GOLD_DEEP[2]) * (math.sin(ratio * math.pi)))
        glc.line([(0, gy), (tw, gy)], fill=(r, g2, b, 255))

    # Mask the gold with the text shape
    text_mask = Image.new("L", (tw, th), 0)
    tmp_crop = tmp.crop((tx, y, tx + tw, y + th))
    text_mask = tmp_crop.split()[3] if tmp_crop.mode == "RGBA" else text_mask

    gold_layer.putalpha(text_mask)
    return gold_layer, tx, y

def draw_gold_text(base_img, text, font, y, align="center", shadow=True):
    gl, tx, ty = gold_gradient_text(ImageDraw.Draw(base_img), text, font, 20, y, align)
    if shadow:
        shadow_layer = Image.new("RGBA", base_img.size, (0, 0, 0, 0))
        shadow_layer.paste((0, 0, 0, 120), (tx + 3, ty + 4), gl.split()[3])
        base_img = Image.alpha_composite(base_img.convert("RGBA"), shadow_layer).convert("RGB")
    rgba = base_img.convert("RGBA")
    rgba.alpha_composite(gl, (tx, ty))
    return rgba.convert("RGB")

# ── Confetti helper ────────────────────────────────────────────────────────────
def draw_confetti(draw_obj, rng_seed=7, count=120):
    random.seed(rng_seed)
    colors = [GOLD_DEEP, GOLD_MID, GOLD_BRIGHT, GOLD_LIGHT, GOLD_SHIMMER, (255, 255, 255)]
    for _ in range(count):
        x = random.randint(0, W)
        y = random.randint(0, H * 2 // 3)
        sz = random.randint(2, 9)
        col = random.choice(colors)
        alpha = random.randint(120, 240)
        shape = random.choice(["rect", "circle", "line"])
        angle = random.uniform(0, 360)
        if shape == "circle":
            draw_obj.ellipse([x-sz//2, y-sz//2, x+sz//2, y+sz//2], fill=col + (alpha,) if len(col)==3 else col)
        elif shape == "rect":
            # rotated rect approximation
            draw_obj.rectangle([x-sz, y-sz//2, x+sz, y+sz//2], fill=col)
        else:
            x2 = x + int(sz * 2 * math.cos(math.radians(angle)))
            y2 = y + int(sz * 2 * math.sin(math.radians(angle)))
            draw_obj.line([(x, y), (x2, y2)], fill=col, width=2)

# ── STAR helper ───────────────────────────────────────────────────────────────
def draw_star(draw_obj, cx, cy, r_outer, r_inner, points=5, fill=GOLD_BRIGHT, outline=GOLD_SHIMMER, width=2):
    coords = []
    for i in range(points * 2):
        angle = math.radians(i * 180 / points - 90)
        r = r_outer if i % 2 == 0 else r_inner
        coords.append((cx + r * math.cos(angle), cy + r * math.sin(angle)))
    draw_obj.polygon(coords, fill=fill, outline=outline)

def draw_star_3d(base_img, cx, cy, r_outer, r_inner, points=5):
    """Multi-layer gold star with glow effect."""
    # Glow
    glow_layer = Image.new("RGBA", base_img.size, (0, 0, 0, 0))
    gd = ImageDraw.Draw(glow_layer)
    for spread in range(12, 0, -2):
        a = int(15 * spread)
        draw_star(gd, cx, cy, r_outer + spread, r_inner + spread//2, fill=(255, 200, 50, a))
    base_img = Image.alpha_composite(base_img.convert("RGBA"), glow_layer).convert("RGB")
    d = ImageDraw.Draw(base_img)
    # Shadow
    draw_star(d, cx+3, cy+4, r_outer, r_inner, fill=(30, 20, 0))
    # Base dark gold
    draw_star(d, cx, cy, r_outer, r_inner, fill=GOLD_DEEP)
    # Mid gold
    draw_star(d, cx, cy, int(r_outer*0.88), int(r_inner*0.88), fill=GOLD_MID)
    # Bright highlight
    draw_star(d, cx, cy, int(r_outer*0.72), int(r_inner*0.72), fill=GOLD_BRIGHT)
    # Shimmer center
    draw_star(d, cx, cy, int(r_outer*0.40), int(r_inner*0.40), fill=GOLD_SHIMMER)
    return base_img

# ── DECORATIVE LINES ──────────────────────────────────────────────────────────
def draw_ornamental_line(draw_obj, y, margin=60, color=GOLD_MID):
    draw_obj.line([(margin, y), (W - margin, y)], fill=color, width=1)
    # Diamond accents
    for dx in [W//2, margin + 30, W - margin - 30]:
        sz = 5
        draw_obj.polygon([
            (dx, y - sz), (dx + sz, y), (dx, y + sz), (dx - sz, y)
        ], fill=GOLD_BRIGHT)

# ══════════════════════════════════════════════════════════════════════════════
# BUILD THE POSTER
# ══════════════════════════════════════════════════════════════════════════════

# 1) Confetti (bottom layer)
confetti_layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
draw_confetti(ImageDraw.Draw(confetti_layer), rng_seed=17, count=150)
img = Image.alpha_composite(img.convert("RGBA"), confetti_layer).convert("RGB")

# 2) Stars arrangement (top center cluster, like reference image)
img = draw_star_3d(img, W//2,      160, 72, 30)   # center large
img = draw_star_3d(img, W//2 - 80, 215, 52, 22)   # left mid
img = draw_star_3d(img, W//2 + 85, 230, 42, 18)   # right small
# Tiny accent stars
d = ImageDraw.Draw(img)
draw_star(d, W//2 - 140, 270, 18, 7, fill=GOLD_MID, outline=GOLD_BRIGHT)
draw_star(d, W//2 + 155, 265, 14, 6, fill=GOLD_MID, outline=GOLD_BRIGHT)
draw_star(d, W//2 + 30,  95, 10, 4, fill=GOLD_BRIGHT, outline=GOLD_SHIMMER)
draw_star(d, W//2 - 45,  90, 8,  3, fill=GOLD_LIGHT)
draw_star(d, W//2 + 120, 150, 7, 3, fill=GOLD_SHIMMER)

# 3) Top decorative horizontal line
draw = ImageDraw.Draw(img)
draw_ornamental_line(draw, 310, margin=55)

# 4) "恭賀" title  ──────────────────────────────────────────────────────────
img = draw_gold_text(img, "恭賀", font_title, y=325, shadow=True)

# 5) Subtitle line below title
draw = ImageDraw.Draw(img)
draw_ornamental_line(draw, 475, margin=55)

# 6) Main body text — split across lines nicely
# Line 1: 本校
line1 = "本校"
# Line 2: highlight name
name_text = "朱宥恩同學"
# Line 3: event
line3a = "參加第十屆"
line3b = "世界青少年武術錦標賽"
# Line 4: award
line4 = "榮獲第二名"

# Draw body lines with gold gradient
draw = ImageDraw.Draw(img)

y_cursor = 495

# "本校" — small label
img = draw_gold_text(img, line1, font_sub, y=y_cursor, shadow=True)
y_cursor += 48

# "朱宥恩同學" — large highlighted name
img = draw_gold_text(img, name_text, font_name, y=y_cursor, shadow=True)
y_cursor += 68

# separator dots
draw = ImageDraw.Draw(img)
for xi in range(W//2 - 40, W//2 + 50, 20):
    draw.ellipse([xi-3, y_cursor-3, xi+3, y_cursor+3], fill=GOLD_MID)
y_cursor += 20

# "參加第十屆"
img = draw_gold_text(img, line3a, font_sub, y=y_cursor, shadow=True)
y_cursor += 46

# "世界青少年武術錦標賽"
img = draw_gold_text(img, line3b, font_sub, y=y_cursor, shadow=True)
y_cursor += 56

# "榮獲第二名" — big, bright
font_award = ImageFont.truetype(font_path_kaiu, 64)
img = draw_gold_text(img, line4, font_award, y=y_cursor, shadow=True)
y_cursor += 82

# 7) Mid ornamental line
draw = ImageDraw.Draw(img)
draw_ornamental_line(draw, y_cursor + 8, margin=55)
y_cursor += 30

# 8) Gold band at bottom ─────────────────────────────────────────────────────
band_y = H - 130
band_h = 88
for gy in range(band_h):
    ratio = gy / band_h
    r = int(GOLD_DEEP[0] + (GOLD_MID[0] - GOLD_DEEP[0]) * math.sin(ratio * math.pi))
    g2= int(GOLD_DEEP[1] + (GOLD_MID[1] - GOLD_DEEP[1]) * math.sin(ratio * math.pi))
    b = int(GOLD_DEEP[2] + (GOLD_MID[2] - GOLD_DEEP[2]) * math.sin(ratio * math.pi))
    draw.line([(0, band_y + gy), (W, band_y + gy)], fill=(r, g2, b))

# 9) Footer text on gold band ─────────────────────────────────────────────────
# "正心中學全體師生"
fw_font = ImageFont.truetype(font_path_bold, 30)
footer_text = "正心中學全體師生"
bb = draw.textbbox((0, 0), footer_text, font=fw_font)
ftw = bb[2] - bb[0]
ftx = (W - ftw) // 2
fty = band_y + (band_h - (bb[3] - bb[1])) // 2

# Dark text on gold band (shadow)
draw.text((ftx + 2, fty + 2), footer_text, font=fw_font, fill=(80, 50, 0))
draw.text((ftx, fty), footer_text, font=fw_font, fill=(10, 8, 2))

# 10) Very thin top rule line
draw.line([(0, band_y), (W, band_y)], fill=GOLD_SHIMMER, width=2)
draw.line([(0, band_y + band_h), (W, band_y + band_h)], fill=GOLD_SHIMMER, width=1)

# 11) Add more confetti / sparkles above the stars
sparkle_layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
sd = ImageDraw.Draw(sparkle_layer)
random.seed(99)
for _ in range(60):
    x = random.randint(20, W - 20)
    y = random.randint(10, 310)
    r = random.randint(1, 4)
    a = random.randint(80, 200)
    col = random.choice([GOLD_BRIGHT, GOLD_SHIMMER, GOLD_LIGHT, (255, 255, 255)])
    sd.ellipse([x-r, y-r, x+r, y+r], fill=col + (a,))
img = Image.alpha_composite(img.convert("RGBA"), sparkle_layer).convert("RGB")

# 12) Edge vignette (dark frame)
vignette = Image.new("RGBA", (W, H), (0, 0, 0, 0))
vd = ImageDraw.Draw(vignette)
for i in range(80):
    a = int(160 * (1 - i / 80) ** 2)
    vd.rectangle([i, i, W - i, H - i], outline=(0, 0, 0, a))
img = Image.alpha_composite(img.convert("RGBA"), vignette).convert("RGB")

# 13) Final subtle sharpening
from PIL import ImageEnhance
img = ImageEnhance.Sharpness(img).enhance(1.3)
img = ImageEnhance.Contrast(img).enhance(1.08)

# ── Save ───────────────────────────────────────────────────────────────────────
out_path = "C:/Users/user/allen/朱宥恩武術錦標賽海報.png"
img.save(out_path, "PNG", dpi=(300, 300))
print(f"Saved: {out_path}  ({W}×{H}px)")
