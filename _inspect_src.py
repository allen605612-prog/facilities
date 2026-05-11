"""讀取設備組內控修訂.docx，列出所有段落文字（含刪除線、紅色標記）"""
import sys, zipfile
from lxml import etree
sys.stdout.reconfigure(encoding='utf-8')

SRC = r"D:\D\114設備組\設備組內控修訂.docx"
W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def w(n): return f"{{{W}}}{n}"

with zipfile.ZipFile(SRC) as z:
    raw = z.read("word/document.xml")

tree = etree.fromstring(raw)
body = tree.find(w("body"))

for para in body.iter(w("p")):
    # 收集段落內所有 run 的文字與格式
    runs = []
    for r in para.findall(w("r")):
        rpr = r.find(w("rPr"))
        flags = []
        if rpr is not None:
            strike = rpr.find(w("strike"))
            if strike is not None and strike.get(w("val"), "true") != "false":
                flags.append("strike")
            color_el = rpr.find(w("color"))
            if color_el is not None:
                c = color_el.get(w("val"), "")
                if c.upper() in ("FF0000", "FF000000"):
                    flags.append("red")
        t_el = r.find(w("t"))
        if t_el is not None and t_el.text:
            runs.append((t_el.text, flags))

    if not runs:
        continue
    line = ""
    for text, flags in runs:
        tag = ""
        if "strike" in flags:
            tag = "[刪]"
        elif "red" in flags:
            tag = "[紅]"
        line += tag + text
    if line.strip():
        print(line[:120])
