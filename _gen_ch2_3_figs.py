import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, Rectangle, FancyArrowPatch
import numpy as np
plt.rcParams['font.family'] = 'Microsoft JhengHei'; plt.rcParams['axes.unicode_minus'] = False
OUT = r'C:\Users\user\allen\_exam_fig\\'

# 圖一：甲、乙兩台車的加速度－外力關係圖
fig, ax = plt.subplots(figsize=(4.2, 3.2), dpi=200)
F = np.array([0, 6])
ax.plot(F, F / 1.0, color='#c0392b', lw=2, label='甲')
ax.plot(F, F / 2.0, color='#1f5fbf', lw=2, ls='--', label='乙')
ax.text(6.1, 6, '甲', color='#c0392b', fontsize=12, va='center'); ax.text(6.1, 3, '乙', color='#1f5fbf', fontsize=12, va='center')
ax.set_xlim(0, 7); ax.set_ylim(0, 7); ax.set_xticks(range(0, 7)); ax.set_yticks(range(0, 7))
ax.set_xlabel('外力（N）'); ax.set_ylabel('加速度（m/s²）'); ax.grid(ls=':', alpha=.6)
fig.tight_layout(); fig.savefig(OUT + 'c2_aF.png'); plt.close(fig)

# 圖二：俯視的等速圓周運動，小球在 P 點，逆時針方向轉動
fig, ax = plt.subplots(figsize=(3.6, 3.4), dpi=200)
th = np.linspace(0, 2 * np.pi, 300)
ax.plot(np.cos(th), np.sin(th), color='k', lw=1.2, ls='--')
ax.plot(0, 0, 'k.', ms=5); ax.text(-.08, -.15, 'O', fontsize=11)
ax.add_patch(Circle((1, 0), .07, color='#f39c12', zorder=5)); ax.text(1.02, -.2, 'P', fontsize=11)
kw = dict(arrowstyle='-|>', mutation_scale=14, lw=1.6, color='#1f5fbf')
for (dx, dy, lab, tx, ty) in [(0, .75, '甲', 1.05, .8), (-.7, 0, '乙', .3, .07), (.7, 0, '丙', 1.75, .07), (0, -.75, '丁', 1.05, -.85)]:
    ax.add_patch(FancyArrowPatch((1, 0), (1 + dx, dy), **kw)); ax.text(tx, ty, lab, fontsize=12)
ax.add_patch(FancyArrowPatch((.707, -.707), (.93, -.37), connectionstyle='arc3,rad=.15', arrowstyle='-|>', mutation_scale=12, color='#7f8c8d'))
ax.text(.35, -.95, '轉動方向', fontsize=9, color='#7f8c8d')
ax.set_xlim(-1.2, 2.0); ax.set_ylim(-1.2, 1.2); ax.set_aspect('equal'); ax.axis('off')
fig.tight_layout(); fig.savefig(OUT + 'c2_circle.png'); plt.close(fig)

# 圖三：單擺，由甲點靜止釋放，擺到丙點
fig, ax = plt.subplots(figsize=(4.0, 3.0), dpi=200)
L = 2.0; piv = (0, 2.2)
ax.add_patch(Rectangle((-1.3, 2.2), 2.6, .12, color='#8d6e63'))
for ang, lab in [(-40, '甲'), (0, '乙'), (40, '丙')]:
    a = np.radians(ang); x, y = piv[0] + L * np.sin(a), piv[1] - L * np.cos(a)
    ax.plot([piv[0], x], [piv[1], y], color='k', lw=1, ls='-' if ang == -40 else ':')
    ax.add_patch(Circle((x, y), .1, color='#f39c12', zorder=5, alpha=1 if ang == -40 else .55))
    ax.text(x, y - .3, lab, ha='center', fontsize=12)
ax.plot([-1.6, 1.6], [.2 - .02, .2 - .02], color='#7f8c8d', lw=.8, ls='--'); ax.text(1.62, .15, '最低點', fontsize=8, color='#7f8c8d')
arc = np.radians(np.linspace(-38, 38, 60))
ax.plot(piv[0] + (L + .05) * np.sin(arc), piv[1] - (L + .05) * np.cos(arc), color='#c0392b', lw=.8, ls='--')
ax.set_xlim(-1.8, 2.1); ax.set_ylim(-.3, 2.5); ax.set_aspect('equal'); ax.axis('off')
fig.tight_layout(); fig.savefig(OUT + 'c3_pendulum.png'); plt.close(fig)

# 圖四：書本靜置桌面，三個力
fig, ax = plt.subplots(figsize=(3.4, 3.0), dpi=200)
ax.add_patch(Rectangle((-1.6, -.25), 3.2, .25, color='#a1887f'))
ax.add_patch(Rectangle((-1.4, -1.6), .2, 1.35, color='#a1887f')); ax.add_patch(Rectangle((1.2, -1.6), .2, 1.35, color='#a1887f'))
ax.add_patch(Rectangle((-.6, 0), 1.2, .45, fc='#f4d03f', ec='k'))
kw = dict(arrowstyle='-|>', mutation_scale=14, lw=1.8)
ax.add_patch(FancyArrowPatch((-.15, .45), (-.15, 1.35), color='#1f5fbf', **kw)); ax.text(-.5, 1.2, r'$F_1$', fontsize=12, color='#1f5fbf')
ax.add_patch(FancyArrowPatch((.15, .22), (.15, -.65), color='#c0392b', **kw)); ax.text(-.45, -.62, r'$F_2$', fontsize=12, color='#c0392b')
ax.add_patch(FancyArrowPatch((.45, .22), (.45, -1.1), color='#2e8b57', **kw)); ax.text(.58, -1.15, r'$F_3$', fontsize=12, color='#2e8b57')
ax.set_xlim(-1.8, 1.8); ax.set_ylim(-1.7, 1.6); ax.set_aspect('equal'); ax.axis('off')
fig.tight_layout(); fig.savefig(OUT + 'c2_book.png'); plt.close(fig)
print('ok')
