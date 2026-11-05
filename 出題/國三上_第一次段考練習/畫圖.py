import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, Rectangle
import numpy as np
plt.rcParams['font.family'] = 'Microsoft JhengHei'; plt.rcParams['axes.unicode_minus'] = False
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fig') + os.sep

# 圖一：遙控車的速度－時間圖（0~4 s 加速、4~8 s 等速、8~10 s 減速）
fig, ax = plt.subplots(figsize=(4.4, 2.8), dpi=200)
ax.plot([0, 4, 8, 10], [0, 8, 8, 0], color='#1f5fbf', lw=2)
for x, lab in [(2, '甲段'), (6, '乙段'), (9, '丙段')]:
    ax.text(x, 9 if x == 6 else 5.2, lab, ha='center', fontsize=10, color='#c0392b')
ax.set_xlim(0, 11); ax.set_ylim(0, 10.5); ax.set_xticks(range(0, 11, 2)); ax.set_yticks(range(0, 11, 2))
ax.set_xlabel('時間（s）'); ax.set_ylabel('速度（m/s）'); ax.grid(ls=':', alpha=.6)
fig.tight_layout(); fig.savefig(OUT + 'q_vt.png'); plt.close(fig)

# 圖二：施力－位移圖（0~4 m 為 10 N，4~8 m 為 5 N）
fig, ax = plt.subplots(figsize=(4.2, 2.6), dpi=200)
ax.plot([0, 4, 4, 8], [10, 10, 5, 5], color='#c0392b', lw=2)
ax.fill_between([0, 4], 0, 10, color='#c0392b', alpha=.08); ax.fill_between([4, 8], 0, 5, color='#c0392b', alpha=.08)
ax.set_xlim(0, 9); ax.set_ylim(0, 12); ax.set_xticks(range(0, 10)); ax.set_yticks(range(0, 13, 2))
ax.set_xlabel('沿施力方向的位移（m）'); ax.set_ylabel('水平施力（N）'); ax.grid(ls=':', alpha=.6)
fig.tight_layout(); fig.savefig(OUT + 'q_fs.png'); plt.close(fig)

# 圖三：單擺，由甲點靜止釋放，經最低點乙擺到丙點
fig, ax = plt.subplots(figsize=(3.8, 2.8), dpi=200)
L = 2.0; piv = (0, 2.2)
ax.add_patch(Rectangle((-1.3, 2.2), 2.6, .12, color='#8d6e63'))
for ang, lab in [(-45, '甲'), (0, '乙'), (45, '丙')]:
    a = np.radians(ang); x, y = piv[0] + L * np.sin(a), piv[1] - L * np.cos(a)
    ax.plot([piv[0], x], [piv[1], y], color='k', lw=1, ls='-' if ang == -45 else ':')
    ax.add_patch(Circle((x, y), .1, color='#f39c12', zorder=5, alpha=1 if ang == -45 else .55))
    ax.text(x, y - .32, lab, ha='center', fontsize=12)
arc = np.radians(np.linspace(-43, 43, 60))
ax.plot(piv[0] + (L + .05) * np.sin(arc), piv[1] - (L + .05) * np.cos(arc), color='#c0392b', lw=.8, ls='--')
ax.set_xlim(-1.8, 1.8); ax.set_ylim(-.3, 2.5); ax.set_aspect('equal'); ax.axis('off')
fig.tight_layout(); fig.savefig(OUT + 'q_pendulum.png'); plt.close(fig)
print('ok')
