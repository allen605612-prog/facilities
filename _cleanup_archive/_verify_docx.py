"""簡單驗證輸出的 docx 表格內容"""
import sys, zipfile, xml.etree.ElementTree as ET
sys.stdout.reconfigure(encoding='utf-8')

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def tag(n): return f"{{{W}}}{n}"

path = r"D:\D\114設備組\修正對照表-內控-設備組.docx"
with zipfile.ZipFile(path) as z:
    with z.open("word/document.xml") as f:
        tree = ET.parse(f)

root = tree.getroot()
body = root.find(tag("body"))
tbl = body.find(tag("tbl"))

rows = tbl.findall(tag("tr"))
print(f"共 {len(rows)} 列（含標題）\n")
for i, tr in enumerate(rows):
    cells = tr.findall(tag("tc"))
    texts = []
    for tc in cells:
        cell_text = ""
        for p in tc.findall(tag("p")):
            for r in p.findall(tag("r")):
                t = r.find(tag("t"))
                if t is not None and t.text:
                    cell_text += t.text
        texts.append(cell_text[:40])
    print(f"列{i}: {' | '.join(texts)}")
