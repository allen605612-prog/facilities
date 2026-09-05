import sys, zipfile
from lxml import etree
sys.stdout.reconfigure(encoding='utf-8')

COPY = r"D:\D\114設備組\修正對照表-內控-設備組 - 複製.docx"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def w(n): return f"{{{W}}}{n}"

with zipfile.ZipFile(COPY) as z:
    raw = z.read("word/document.xml")

tree = etree.fromstring(raw)
body = tree.find(w("body"))
tbl  = body.find(w("tbl"))
rows = tbl.findall(w("tr"))
print(f"共 {len(rows)} 列")
for i, tr in enumerate(rows):
    cells = tr.findall(w("tc"))
    cols = []
    for tc in cells:
        text = "".join(x.text or "" for x in tc.iter(w("t")))
        cols.append(text[:50])
    print(f"列{i}: {' | '.join(cols)}")
