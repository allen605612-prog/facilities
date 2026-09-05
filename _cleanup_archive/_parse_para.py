"""
輸出每段落的完整格式+文字，方便理解文件結構
"""
import xml.etree.ElementTree as ET
import sys

sys.stdout.reconfigure(encoding='utf-8')

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def tag(name):
    return f"{{{W}}}{name}"

def get_rpr_info(elem):
    color_elem = elem.find(tag("color"))
    color = color_elem.get(tag("val"), "000000").upper() if color_elem is not None else "000000"
    strike_elem = elem.find(tag("strike"))
    strike = False
    if strike_elem is not None:
        val = strike_elem.get(tag("val"), "true")
        strike = val not in ("false", "0")
    return color, strike

tree = ET.parse(r"C:\Users\user\allen\_unpack_src\word\document.xml")
root = tree.getroot()
body = root.find(tag("body"))

for i, p in enumerate(body.findall(tag("p"))):
    # Get paragraph-level rpr
    ppr = p.find(tag("pPr"))
    para_color = "000000"
    para_strike = False
    if ppr is not None:
        ppr_rpr = ppr.find(tag("rPr"))
        if ppr_rpr is not None:
            para_color, para_strike = get_rpr_info(ppr_rpr)

    # Get all runs
    runs = []
    for r in p.findall(tag("r")):
        t = r.find(tag("t"))
        text = t.text if t is not None and t.text else ""
        if not text:
            continue
        rpr = r.find(tag("rPr"))
        if rpr is not None:
            color, strike = get_rpr_info(rpr)
        else:
            color, strike = para_color, para_strike
        runs.append((text, color, strike))

    if not runs:
        continue

    # Categorize paragraph
    cats = set()
    for text, color, strike in runs:
        is_red = color == "FF0000"
        if is_red and strike:
            cats.add("紅刪")
        elif is_red and not strike:
            cats.add("紅字")
        elif not is_red and strike:
            cats.add("黑刪")
        else:
            cats.add("黑字")

    full_text = "".join(t for t, c, s in runs)
    cat_str = "+".join(sorted(cats))
    print(f"[P{i:03d}] [{cat_str}] {full_text[:80]}")
