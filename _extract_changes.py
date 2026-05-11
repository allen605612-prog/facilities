"""
詳細提取每個有格式變化的段落，輸出：現行規定 vs 修正規定
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

# 目標段落：有黑刪、紅字、紅刪的段落
target_ids = {247, 248, 249, 254, 279, 328, 334, 339, 345, 375, 377}

for i, p in enumerate(body.findall(tag("p"))):
    if i not in target_ids:
        continue

    ppr = p.find(tag("pPr"))
    para_color = "000000"
    para_strike = False
    if ppr is not None:
        ppr_rpr = ppr.find(tag("rPr"))
        if ppr_rpr is not None:
            para_color, para_strike = get_rpr_info(ppr_rpr)

    print(f"\n{'='*60}")
    print(f"P{i:03d}")

    normal_text = []   # 黑字，無刪除線
    black_strike = []  # 黑字，有刪除線（現行規定）
    red_text = []      # 紅字，無刪除線（修正規定）
    red_strike = []    # 紅字，有刪除線（被刪除的舊規定）

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

        is_red = color == "FF0000"
        if is_red and strike:
            red_strike.append(text)
        elif is_red and not strike:
            red_text.append(text)
        elif not is_red and strike:
            black_strike.append(text)
        else:
            normal_text.append(text)

    print(f"黑字（正常）: {''.join(normal_text)}")
    print(f"黑刪（現行）: {''.join(black_strike)}")
    print(f"紅字（修正）: {''.join(red_text)}")
    print(f"紅刪（整段刪）: {''.join(red_strike)}")
