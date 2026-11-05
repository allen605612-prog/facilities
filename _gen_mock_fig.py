import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon, Rectangle
plt.rcParams['font.family'] = 'Microsoft JhengHei'
W, H, D = 10, 6, 0            # 寬、高（已取消斷層，右側不下移）
fig, ax = plt.subplots(figsize=(5.2, 3.2), dpi=200)
blocks = [(Polygon([(0, 0), (W, 0), (W, H), (0, H)], closed=True, transform=ax.transData), 0)]
for clip, off in blocks:
    items = [Rectangle((0, -3), W, 6 + off, fc='#d9d9d9', ec='k', hatch='---', lw=.8),       # 乙：泥質（頁岩）
             Rectangle((0, 3 + off), W, 6, fc='#f3e3b5', ec='k', hatch='...', lw=.8),       # 甲：砂岩
             Polygon([(2.0, -3 + off), (2.9, -3 + off), (9.4, 9 + off), (8.5, 9 + off)], closed=True, fc='#333333', ec='k')]  # 丙：岩脈
    for it in items:
        ax.add_patch(it); it.set_clip_path(clip)
ax.add_patch(Rectangle((0, 0), W, H, fill=False, lw=1.2))
kw = dict(fontsize=13, ha='center', va='center', bbox=dict(fc='white', ec='none', pad=1))
ax.text(1.2, 4.6, '甲', **kw); ax.text(1.2, 1.5, '乙', **kw)
ax.text(4.6, 1.0, '丙', **kw)
ax.set_xlim(-.2, W + .2); ax.set_ylim(-.2, H + .2); ax.set_aspect('equal'); ax.axis('off')
fig.tight_layout(); fig.savefig(r'C:\Users\user\allen\_exam_fig\mock_strata.png'); plt.close(fig)
print('ok')
