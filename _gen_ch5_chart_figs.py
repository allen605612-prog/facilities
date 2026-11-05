import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
plt.rcParams['font.family'] = 'Microsoft JhengHei'; plt.rcParams['axes.unicode_minus'] = False
out = r'C:\Users\user\allen\_exam_fig'

# 圖一：地下水位與累計地層下陷
yr = [1951, 1961, 1971, 1975, 1981, 1991, 2007]
wl = [-5, -20, -45, -42, -25, -10, -5]
sub = [0, -40, -150, -200, -210, -215, -215]
fig, ax = plt.subplots(figsize=(6.4, 3.4), dpi=200)
ax.plot(yr, wl, 'o-', color='#1f5fbf', lw=2, label='地下水位')
ax.set_ylim(-50, 0); ax.set_ylabel('地下水位（公尺）'); ax.set_xlabel('年（西元）')
ax.set_xticks([1951, 1961, 1971, 1981, 1991, 2001]); ax.grid(ls=':', alpha=.6)
ax2 = ax.twinx()
ax2.plot(yr, sub, 's--', color='#c0392b', lw=2, label='累計地層下陷量')
ax2.set_ylim(-250, 0); ax2.set_ylabel('累計地層下陷量（公分）')
h1, l1 = ax.get_legend_handles_labels(); h2, l2 = ax2.get_legend_handles_labels()
ax.legend(h1 + h2, l1 + l2, loc='lower left', fontsize=9)
fig.tight_layout(); fig.savefig(out + r'\fig1.png'); plt.close(fig)

# 圖二：河流沉積物剖面
layers = [('甲', '沙', '....'), ('乙', '礫石', 'OO'), ('丙', '泥', '---'), ('丁', '沙', '....')]
fig, ax = plt.subplots(figsize=(4.2, 3.2), dpi=200)
for i, (n, kind, h) in enumerate(layers):
    y = 3 - i
    ax.add_patch(Rectangle((0, y), 2, 1, fill=False, hatch=h, lw=1.2))
    ax.text(-0.25, y + .5, n, ha='right', va='center', fontsize=13)
ax.text(1, 4.25, '地表', ha='center', fontsize=11)
for j, (kind, h) in enumerate([('礫石', 'OO'), ('沙', '....'), ('泥', '---')]):
    ax.add_patch(Rectangle((2.6, 3.1 - j * 1.1), .6, .6, fill=False, hatch=h))
    ax.text(3.35, 3.4 - j * 1.1, kind, va='center', fontsize=11)
ax.set_xlim(-0.8, 4.2); ax.set_ylim(-0.1, 4.6); ax.axis('off')
fig.tight_layout(); fig.savefig(out + r'\fig2.png'); plt.close(fig)
print('ok')
