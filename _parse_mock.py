# 解析「模擬測驗」docx → _ch5_mock_data.json（題目、選項、表格、題組、答案、解析）
import docx, json, re
from docx.oxml.ns import qn
SRC = r'C:\Users\user\Downloads\國三自然_第5章水與陸地_模擬測驗.md.docx'
d = docx.Document(SRC)
paras = {p._p: p for p in d.paragraphs}
tables = {t._tbl: t for t in d.tables}
seq = []
for el in d.element.body:
    if el.tag == qn('w:p'): seq.append(('p', paras[el]))
    elif el.tag == qn('w:tbl'): seq.append(('t', tables[el]))

def widths_for(rows):
    n = len(rows[0])
    if n == 2: return [2400, 7000]
    first = 2000 if max(len(r[0]) for r in rows) > 3 else 1400
    rest = (9400 - first) // (n - 1)
    return [first] + [rest] * (n - 1)

EMPH = ['最無助於', '最不合理', '最不準確', '不是', '錯誤']
parts, part, pend, answers_expl = [], None, None, []
in_ans = False
for kind, x in seq:
    if kind == 't':
        rows = [[c.text.strip() for c in r.cells] for r in x.rows]
        if in_ans: continue
        pend.setdefault('pre', []).append({'table': rows, 'widths': widths_for(rows)})
        continue
    s = x.text.strip()
    if not s: continue
    st = x.style.name
    if part is None and not st.startswith('Heading'): continue   # 卷頭文字
    if st.startswith('Heading'):
        if '答案與解析' in s: in_ans = True; continue
        if in_ans: continue
        part = {'name': re.sub(r'（第.*題）', '', s).strip(), 'items': []}; parts.append(part); pend = {}
        continue
    if in_ans:
        m = re.match(r'\(([A-D])\)\s*(.*)', s); answers_expl.append((m.group(1), m.group(2))); continue
    m = re.match(r'（\s*）\s*(\d+)\.\s*(.*)', s)
    if m:
        item = {'q': m.group(2), 'o': []}
        item.update({k: v for k, v in pend.items()}); pend = {}
        part['items'].append(item); continue
    if s.startswith('(A)') or re.match(r'^\([A-D]\)', s):
        part['items'][-1]['o'] += [o.strip() for o in re.split(r'\([A-D]\)\s*', s) if o.strip()]
        continue
    if s.startswith('（此處假設有一圖'):
        pend.setdefault('pre', []).append({'fig': ['mock_strata.png', 380, 234]}); continue
    if re.match(r'^【5-\d', s): pend['sec'] = s; continue
    if s.startswith('【題組'): pend['group'] = s; continue
    pend.setdefault('pre', []).append({'text': s})
n = 0
for p in parts:
    for it in p['items']:
        a, e = answers_expl[n]; n += 1
        it['a'], it['e'] = a, e
        assert len(it['o']) == 4, (n, it['o'])
        for w in EMPH:
            if w in it['q']: it['emph'] = w; break
print('題數', n, '解析', len(answers_expl))
json.dump(parts, open('_ch5_mock_data.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
