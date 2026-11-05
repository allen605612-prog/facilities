# 模擬測驗題庫後處理：修正第28題(D)、重排選項使答案 ABCD 平均（接在 _parse_mock.py 之後執行）
import json, random, re
f = '_ch5_mock_data.json'
P = json.load(open(f, encoding='utf-8'))
Q = [it for p in P for it in p['items']]
L = 'ABCD'

q28 = Q[27]
assert q28['o'][3].startswith('沉積岩只有埋藏'), q28['o'][3]
q28['o'][3] = '沉積岩只能變成變質岩，無法再形成新的沉積岩'

FIXED = {6: 'B', 22: 'A'}    # 第6題選項本身就是A～D礦物代號、第22題上中下游有順序，不動
assert all(Q[k - 1]['a'] == v for k, v in FIXED.items())
rng = random.Random(115)
while True:                  # 產生 A13 B13 C12 D12、同一答案最多連續2題的目標序列
    pool = list('A' * 13 + 'B' * 13 + 'C' * 12 + 'D' * 12)
    for v in FIXED.values(): pool.remove(v)
    rng.shuffle(pool)
    seq = [FIXED.get(i + 1) or pool.pop() for i in range(50)]
    if not re.search(r'(.)\1\1', ''.join(seq)): break

for it, tgt in zip(Q, seq):
    a, b = L.index(it['a']), L.index(tgt)
    if a == b: continue
    it['o'][a], it['o'][b] = it['o'][b], it['o'][a]          # 正解與目標位置互換
    sw = {L[a]: L[b], L[b]: L[a]}
    it['e'] = re.sub(r'\(([A-D])\)', lambda m: f'({sw.get(m.group(1), m.group(1))})', it['e'])
    it['a'] = tgt
json.dump(P, open(f, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print(''.join(seq), {c: seq.count(c) for c in L})
