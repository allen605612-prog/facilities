import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, Polygon
import numpy as np
plt.rcParams['font.family'] = 'Microsoft JhengHei'; plt.rcParams['axes.unicode_minus'] = False
OUT = r'C:\Users\user\allen\_exam_fig\\'

# 圖一：直尺讀數（鉛筆右端在 12.36 cm）
fig, ax = plt.subplots(figsize=(5.6, 1.9), dpi=200)
x0, x1 = 10.6, 13.6
ax.add_patch(Rectangle((x0, 0), x1 - x0, 1, fc='#fdf3d0', ec='k', lw=1))
for i in range(int(x0 * 10) + 1, int(x1 * 10)):
    x = i / 10
    h = .45 if i % 10 == 0 else (.3 if i % 5 == 0 else .18)
    ax.plot([x, x], [1, 1 - h], color='k', lw=.8)
    if i % 10 == 0: ax.text(x, .35, str(i // 10), ha='center', va='center', fontsize=12)
ax.text(13.45, .15, 'cm', fontsize=10)
ax.add_patch(Polygon([(x0 - .3, 1.12), (12.2, 1.12), (12.36, 1.35), (12.2, 1.58), (x0 - .3, 1.58)], closed=True, fc='#f4c542', ec='k', lw=1))
ax.plot([12.36, 12.36], [1.0, 1.35], color='#c0392b', lw=1, ls=':')
ax.text(10.8, 1.35, '鉛筆', va='center', fontsize=11)
ax.set_xlim(x0 - .2, x1 + .1); ax.set_ylim(-.1, 1.7); ax.axis('off')
fig.tight_layout(); fig.savefig(OUT + 'c1_ruler.png'); plt.close(fig)

# 圖二：排水法（32.0 mL → 40.0 mL）
fig, axs = plt.subplots(1, 2, figsize=(4.0, 3.2), dpi=200)
for ax, lv, title, stone in [(axs[0], 32.0, '放入石頭前', False), (axs[1], 40.0, '放入石頭後', True)]:
    lo, hi = 26, 44
    ax.add_patch(Rectangle((0, lo), 1, hi - lo, fill=False, lw=1.3))
    ax.add_patch(Rectangle((0, lo), 1, lv - lo, fc='#bfe3f5', ec='none'))
    for v in range(lo + 1, hi):
        w = .45 if v % 5 == 0 else .25
        ax.plot([0, w], [v, v], color='k', lw=.8)
        if v % 5 == 0: ax.text(-.12, v, str(v), ha='right', va='center', fontsize=9)
    ax.plot([0, 1], [lv, lv], color='#1f5fbf', lw=1.5)
    if stone: ax.add_patch(Polygon([(.45, lo + .3), (.8, lo + .5), (.75, lo + 2.2), (.5, lo + 2.6), (.35, lo + 1.5)], closed=True, fc='#8d8d8d', ec='k'))
    ax.text(.5, lo - 1.8, title, ha='center', fontsize=10)
    ax.text(1.05, hi - .6, 'mL', fontsize=8)
    ax.set_xlim(-.6, 1.4); ax.set_ylim(lo - 2.5, hi + .5); ax.axis('off')
fig.tight_layout(); fig.savefig(OUT + 'c1_cylinder.png'); plt.close(fig)

# 圖三：質量－體積關係圖
fig, ax = plt.subplots(figsize=(4.2, 3.2), dpi=200)
V = np.array([0, 10])
for d, name, c, ls in [(7.9, '甲', '#c0392b', '-'), (2.7, '乙', '#1f5fbf', '--'), (1.0, '丙', '#2e8b57', '-.')]:
    vmax = min(25, 80 / d)
    ax.plot([0, vmax], [0, d * vmax], color=c, lw=2, ls=ls)
    ax.text(vmax + .3, d * vmax, name, color=c, fontsize=12, va='center')
ax.plot([10], [79], 'o', color='#c0392b', ms=4); ax.plot([20], [54], 's', color='#1f5fbf', ms=4); ax.plot([20], [20], '^', color='#2e8b57', ms=4)
ax.set_xlim(0, 27); ax.set_ylim(0, 85)
ax.set_xticks(range(0, 26, 5)); ax.set_yticks(range(0, 81, 10))
ax.set_xlabel('體積（cm³）'); ax.set_ylabel('質量（g）'); ax.grid(ls=':', alpha=.6)
fig.tight_layout(); fig.savefig(OUT + 'c1_mv.png'); plt.close(fig)

# 圖四：繩波（波長 20 cm、振幅 5 cm）
fig, ax = plt.subplots(figsize=(5.6, 2.3), dpi=200)
x = np.linspace(0, 60, 600)
ax.plot(x, 5 * np.sin(2 * np.pi * x / 20), color='#1f5fbf', lw=2)
ax.axhline(0, color='k', lw=.8, ls='--')
ax.set_xlim(0, 62); ax.set_ylim(-8, 8)
ax.set_xticks(range(0, 61, 10)); ax.set_yticks([-5, 0, 5])
ax.set_xlabel('位置（cm）'); ax.set_ylabel('位移（cm）'); ax.grid(ls=':', alpha=.5)
ax.annotate('', xy=(61, 6.8), xytext=(50, 6.8), arrowprops=dict(arrowstyle='->', lw=1.5))
ax.text(50, 7.1, '波前進方向', fontsize=9, va='bottom')
fig.tight_layout(); fig.savefig(OUT + 'c3_wave.png'); plt.close(fig)
print('ok')
