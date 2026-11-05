import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
plt.rcParams['font.family'] = 'Microsoft JhengHei'; plt.rcParams['axes.unicode_minus'] = False
yr = [1960, 1965, 1970, 1975, 1978, 1985, 1990, 1995, 2000, 2005, 2010, 2015, 2020]
pos = [0, 30, 60, 85, 100, 75, 45, 15, -20, -45, -70, -85, -100]
fig, ax = plt.subplots(figsize=(6.4, 3.4), dpi=200)
ax.plot(yr, pos, 'o-', color='#1f5fbf', lw=2)
ax.axhline(0, color='k', lw=.8)
ax.axvline(1978, color='#c0392b', ls='--', lw=1.2)
ax.text(1979, -100, '上游水庫完工', color='#c0392b', fontsize=10)
ax.set_xlabel('年（西元）'); ax.set_ylabel('海岸線位置（公尺）')
ax.set_ylim(-120, 120); ax.grid(ls=':', alpha=.6)
ax.text(1961, 105, '↑ 向海方向', fontsize=9); ax.text(1961, -112, '↓ 向陸地方向', fontsize=9)
fig.tight_layout(); fig.savefig(r'C:\Users\user\allen\_exam_fig\h5_coast.png'); plt.close(fig)
print('ok')
