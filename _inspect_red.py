"""只列出含有紅色字或混合刪除線+正常文字的段落（修改項目）"""
import sys, zipfile
from lxml import etree
sys.stdout.reconfigure(encoding='utf-8')

SRC = r"D:\D\114設備組\設備組內控修訂.docx"
W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def w(n): return f"{{{W}}}{n}"

RED_COLORS = {"FF0000", "FF000000", "C00000", "C0000000"}

with zipfile.ZipFile(SRC) as z:
    raw = z.read("word/document.xml")

tree = etree.fromstring(raw)
body = tree.find(w("body"))

for para in body.iter(w("p")):
    runs = []
    has_red = False
    has_strike = False
    has_normal = False

    for r in para.findall(w("r")):
        rpr = r.find(w("rPr"))
        is_strike = False
        is_red = False
        if rpr is not None:
            s = rpr.find(w("strike"))
            if s is not None and s.get(w("val"), "true") != "false":
                is_strike = True
            c = rpr.find(w("color"))
            if c is not None and c.get(w("val"), "").upper().lstrip("0") in {x.lstrip("0") for x in RED_COLORS}:
                is_red = True
        t_el = r.find(w("t"))
        if t_el is not None and t_el.text and t_el.text.strip():
            if is_red:
                has_red = True
                runs.append(f"[紅]{t_el.text}")
            elif is_strike:
                has_strike = True
                runs.append(f"[刪]{t_el.text}")
            else:
                has_normal = True
                runs.append(t_el.text)

    # 只顯示含紅字、或含有刪除線且有正常文字的段落
    if has_red or (has_strike and has_normal):
        print("".join(runs)[:150])
