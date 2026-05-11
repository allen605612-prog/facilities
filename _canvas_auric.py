#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Auric Meridian — A Systematic Cartography of Luminous Achievement
4000 × 5600 px  |  300 DPI
"""
import numpy as np
from PIL import Image, ImageDraw, ImageFont
import math, random
from pathlib import Path

W, H = 4000, 5600
DPI  = 300
FONT = Path(r"C:\Users\user\.claude\plugins\cache\anthropic-agent-skills\document-skills\b0cbd3df1533\skills\canvas-design\canvas-fonts")

def fnt(name, size): return ImageFont.truetype(str(FONT / name), size)

# ── Palette ──────────────────────────────────────────────────
VOID   = np.array([ 3,  5, 16], np.float32)
DEEP   = np.array([ 7, 12, 32], np.float32)
MID    = np.array([12, 20, 50], np.float32)
GOLD   = (212, 175,  55)
GOLD_L = (255, 220,  88)
GOLD_P = (242, 208, 140)
GOLD_D = (130,  95,  22)
CREAM  = (248, 238, 210)
SILVER = (168, 176, 198)
WHITE  = (255, 252, 248)

cx, cy  = W // 2, int(H * 0.415)
R_MAIN  = 1480
R_OUTER = 1548

# ═══════════════════════════════════════════════════════════════
# 1. BACKGROUND  (vectorised radial + vertical gradient)
# ═══════════════════════════════════════════════════════════════
xs = np.arange(W, dtype=np.float32)[np.newaxis, :]
ys = np.arange(H, dtype=np.float32)[:, np.newaxis]
dist  = np.sqrt((xs - cx)**2 + (ys - cy)**2)
t_rad = np.clip(dist / (W * 0.64), 0, 1)[:, :, np.newaxis]
t_ver = np.clip(ys / H * 0.38,     0, 0.34)[:, :, np.newaxis]

arr = MID + (DEEP - MID) * t_rad + (VOID - MID) * t_ver
arr = np.clip(arr, 0, 255).astype(np.uint8)
img  = Image.fromarray(arr, 'RGB')
draw = ImageDraw.Draw(img, 'RGBA')

# ═══════════════════════════════════════════════════════════════
# 2. AMBIENT STARFIELD
# ═══════════════════════════════════════════════════════════════
rng = random.Random(2614)
for _ in range(3200):
    x = rng.randint(120, W - 120)
    y = rng.randint(120, H - 120)
    d = math.hypot(x - cx, y - cy)
    if d < R_MAIN * 0.92:
        continue
    t_f = min(1.0, (d - R_MAIN) / (W * 0.32))
    al  = int(rng.randint(12, 55) * (0.25 + 0.75 * t_f))
    sz  = rng.choice([0, 0, 0, 1, 1, 2])
    col = rng.choice([GOLD_P, SILVER, WHITE, CREAM])
    draw.ellipse([x-sz, y-sz, x+sz, y+sz], fill=(*col, al))

# ═══════════════════════════════════════════════════════════════
# 2b. INNER SPHERE AMBIENT LIFT  (very faint warm glow inside circle)
# ═══════════════════════════════════════════════════════════════
for r_lift in range(R_MAIN, 0, -12):
    t = r_lift / R_MAIN
    al = int(8 * (1 - t) ** 1.6)
    draw.ellipse([cx-r_lift, cy-r_lift, cx+r_lift, cy+r_lift],
                 fill=(18, 25, 60, al))

# ═══════════════════════════════════════════════════════════════
# 3. LATITUDE PARALLELS  (ellipses inside sphere)
# ═══════════════════════════════════════════════════════════════
for lat_deg in range(-75, 76, 15):
    lat = math.radians(lat_deg)
    r_x = R_MAIN * math.cos(lat)
    c_y = cy + int(R_MAIN * math.sin(lat))
    r_y = int(r_x * 0.13)
    if r_x < 30:
        continue
    al  = 90 if lat_deg == 0 else 58
    lw  =  2 if lat_deg == 0 else  1
    draw.ellipse([cx - int(r_x), c_y - r_y, cx + int(r_x), c_y + r_y],
                 outline=(*GOLD_D, al), width=lw)

# ═══════════════════════════════════════════════════════════════
# 4. MERIDIAN LINES  (orthographic sphere projection)
# ═══════════════════════════════════════════════════════════════
for lon_deg in range(0, 180, 15):
    lon = math.radians(lon_deg)
    pts = []
    for i in range(240):
        t  = i / 239
        la = math.radians(-89 + 178 * t)
        sx = math.cos(la) * math.sin(lon)
        sy = math.sin(la)
        if sx * sx + sy * sy > 1.002:
            continue
        pts.append((cx + sx * R_MAIN, cy + sy * R_MAIN))
    for j in range(len(pts) - 1):
        draw.line([pts[j], pts[j+1]], fill=(*GOLD_D, 62), width=1)

# ═══════════════════════════════════════════════════════════════
# 5. RADIAL SPOKES  (ghost lines from centre to sphere edge)
# ═══════════════════════════════════════════════════════════════
for deg in range(0, 360, 15):
    rad = math.radians(deg)
    x0  = cx + R_MAIN * 0.06 * math.cos(rad)
    y0  = cy + R_MAIN * 0.06 * math.sin(rad)
    xe  = cx + R_MAIN        * math.cos(rad)
    ye  = cy + R_MAIN        * math.sin(rad)
    draw.line([x0, y0, xe, ye], fill=(*GOLD_D, 28), width=1)

# ═══════════════════════════════════════════════════════════════
# 6. CONCENTRIC RINGS
# ═══════════════════════════════════════════════════════════════
for frac, al in [(0.07, 70), (0.17, 62), (0.29, 55),
                 (0.44, 48), (0.60, 40), (0.77, 32)]:
    r = int(R_MAIN * frac)
    draw.ellipse([cx-r, cy-r, cx+r, cy+r], outline=(*GOLD, al), width=1)

# ═══════════════════════════════════════════════════════════════
# 7. GOLD DOT FIELD  (dense at centre, diffuse at edges)
# ═══════════════════════════════════════════════════════════════
rng2 = random.Random(66)
for _ in range(9000):
    r_norm = rng2.betavariate(0.75, 3.2)
    r      = r_norm * R_MAIN * 0.96
    angle  = rng2.uniform(0, math.tau)
    x = cx + r * math.cos(angle)
    y = cy + r * math.sin(angle)
    sf = (1 - r_norm) ** 0.75
    sz = rng2.uniform(0.4, 3.0) * sf + 0.3
    al = min(240, int(rng2.uniform(70, 230) * sf + 35))
    p  = rng2.random()
    c  = GOLD_L if p < 0.40 else GOLD if p < 0.72 else GOLD_P if p < 0.90 else CREAM
    ri = max(1, int(sz))
    draw.ellipse([x-ri, y-ri, x+ri, y+ri], fill=(*c, al))

# ═══════════════════════════════════════════════════════════════
# 8. MAIN SPHERE BORDER  +  OUTER INSTRUMENT RING
# ═══════════════════════════════════════════════════════════════
draw.ellipse([cx-R_MAIN-1, cy-R_MAIN-1, cx+R_MAIN+1, cy+R_MAIN+1],
             outline=(*GOLD, 155), width=3)
draw.ellipse([cx-R_OUTER,    cy-R_OUTER,    cx+R_OUTER,    cy+R_OUTER],
             outline=(*GOLD, 145), width=5)
draw.ellipse([cx-R_OUTER-14, cy-R_OUTER-14, cx+R_OUTER+14, cy+R_OUTER+14],
             outline=(*GOLD_D, 60), width=1)

# Tick marks around outer ring
for deg in range(0, 360, 2):
    rad = math.radians(deg)
    if   deg % 90 == 0: t_l, t_w, t_al, t_col = 40, 4, 240, GOLD_L
    elif deg % 15 == 0: t_l, t_w, t_al, t_col = 24, 2, 190, GOLD
    elif deg %  5 == 0: t_l, t_w, t_al, t_col = 14, 1, 125, GOLD_D
    else:               t_l, t_w, t_al, t_col =  7, 1,  60, GOLD_D
    r0 = R_OUTER + 14
    x1 = cx + r0         * math.cos(rad)
    y1 = cy + r0         * math.sin(rad)
    x2 = cx + (r0-t_l)   * math.cos(rad)
    y2 = cy + (r0-t_l)   * math.sin(rad)
    draw.line([x1, y1, x2, y2], fill=(*t_col, t_al), width=t_w)

# Cardinal diamond markers (outside ring)
for deg in [0, 90, 180, 270]:
    rad = math.radians(deg)
    r_m = R_OUTER + 60
    xm  = cx + r_m * math.cos(rad)
    ym  = cy + r_m * math.sin(rad)
    s   = 18
    draw.polygon([(xm, ym-s), (xm+s, ym), (xm, ym+s), (xm-s, ym)],
                 fill=(*GOLD, 195))
    draw.polygon([(xm, ym-s*0.5), (xm+s*0.5, ym), (xm, ym+s*0.5), (xm-s*0.5, ym)],
                 fill=(*GOLD_L, 230))

# ═══════════════════════════════════════════════════════════════
# 9. CENTRAL STELLAR CORE
# ═══════════════════════════════════════════════════════════════
# Soft radial glow
for r_g in range(220, 0, -3):
    t  = r_g / 220
    al = int(185 * (1 - t) ** 2.4)
    lr = int(GOLD_L[0] + (255 - GOLD_L[0]) * (1-t) * 0.55)
    lg = int(GOLD_L[1] * (0.55 + 0.45 * (1-t)))
    lb = int(GOLD_L[2] * (1 - t * 0.65))
    draw.ellipse([cx-r_g, cy-r_g, cx+r_g, cy+r_g], fill=(lr, lg, lb, al))

# 8-point star
def star8(cx, cy, r_o, r_i, col, al=255):
    pts = [(cx + (r_o if i%2==0 else r_i)*math.cos(math.radians(i*22.5-90)),
            cy + (r_o if i%2==0 else r_i)*math.sin(math.radians(i*22.5-90)))
           for i in range(16)]
    draw.polygon(pts, fill=(*col, al))

star8(cx, cy,  60,  24, GOLD_D)
star8(cx, cy,  50,  20, GOLD)
star8(cx, cy,  37,  15, GOLD_L)
star8(cx, cy,  20,   8, CREAM)
draw.ellipse([cx-7, cy-7, cx+7, cy+7], fill=(*WHITE, 255))

# ═══════════════════════════════════════════════════════════════
# 10. TYPOGRAPHY
# ═══════════════════════════════════════════════════════════════
f_title  = fnt("Italiana-Regular.ttf",     115)
f_sub    = fnt("Jura-Light.ttf",            42)
f_mono   = fnt("IBMPlexMono-Regular.ttf",   28)
f_corner = fnt("DMMono-Regular.ttf",        26)

# Thin horizontal rule above title
rule_y = cy + R_OUTER + 100
rw = 640
draw.line([cx - rw//2, rule_y, cx + rw//2, rule_y], fill=(*GOLD_D, 100), width=1)
for dx in [-rw//2, 0, rw//2]:
    s = 7
    draw.polygon([(cx+dx, rule_y-s), (cx+dx+s, rule_y),
                  (cx+dx, rule_y+s), (cx+dx-s, rule_y)],
                 fill=(*GOLD, 150))

# Title
title = "AURIC  MERIDIAN"
bb = draw.textbbox((0, 0), title, font=f_title)
tw, th = bb[2]-bb[0], bb[3]-bb[1]
tx = cx - tw // 2
ty = rule_y + 36
draw.text((tx + 2, ty + 2), title, font=f_title, fill=(0, 0, 0, 60))   # shadow
draw.text((tx,     ty    ), title, font=f_title, fill=(*GOLD_L, 218))

# Subtitle
sub = "A SYSTEMATIC CARTOGRAPHY OF LUMINOUS ACHIEVEMENT"
bb2 = draw.textbbox((0, 0), sub, font=f_sub)
sw, sh = bb2[2]-bb2[0], bb2[3]-bb2[1]
sx2 = cx - sw // 2
sy2 = ty + th + 20
draw.text((sx2, sy2), sub, font=f_sub, fill=(*GOLD_D, 175))

# Thin rule below subtitle
rule2_y = sy2 + sh + 30
draw.line([cx - rw//2, rule2_y, cx + rw//2, rule2_y], fill=(*GOLD_D, 78), width=1)

# Bottom metadata (monospaced, very faint)
meta = "FIELD STUDY  No. LXVI   ·   OBSERVATIONAL SERIES   ·   2026"
bbm  = draw.textbbox((0, 0), meta, font=f_mono)
mx   = cx - (bbm[2]-bbm[0]) // 2
my   = H - 220
draw.text((mx, my), meta, font=f_mono, fill=(*SILVER, 88))

# Corner coordinate annotations
corners = [
    (210, 210, "α  23h 14m 06s"),
    (W-210, 210, "δ  +66° 06′ 22″"),
    (210, H-210, "r  1.000 AU"),
    (W-210, H-210, "T  2026.329"),
]
for (x, y, label) in corners:
    bbc = draw.textbbox((0, 0), label, font=f_corner)
    lw2, lh2 = bbc[2]-bbc[0], bbc[3]-bbc[1]
    draw.text((x - lw2//2, y - lh2//2), label, font=f_corner, fill=(*GOLD_D, 72))

# ═══════════════════════════════════════════════════════════════
# 11. EDGE VIGNETTE
# ═══════════════════════════════════════════════════════════════
vig = Image.new("RGBA", (W, H), (0, 0, 0, 0))
vd  = ImageDraw.Draw(vig)
for i in range(200):
    t  = 1 - i / 200
    al = int(230 * t ** 3.0)
    vd.rectangle([i*2, i*2, W-i*2, H-i*2], outline=(0, 0, 0, al))
img = img.convert("RGBA")
img.alpha_composite(vig)
img = img.convert("RGB")

# ═══════════════════════════════════════════════════════════════
# SAVE
# ═══════════════════════════════════════════════════════════════
out = r"C:\Users\user\allen\auric_meridian.png"
img.save(out, "PNG", dpi=(DPI, DPI))
print(f"Saved → {out}  ({W}×{H} @ {DPI} DPI)")
