"""
解析 設備組內控修訂.docx，分類文字：
- 黑色刪除線 (strike, color!=FF0000 or no color) → 現行規定
- 紅色無刪除線 (color=FF0000, no strike) → 修正規定
- 紅色有刪除線 (color=FF0000, strike) → 刪除的舊規定
"""
import xml.etree.ElementTree as ET
import re

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def tag(name):
    return f"{{{W}}}{name}"

def get_rpr(r_elem):
    rpr = r_elem.find(tag("rPr"))
    if rpr is None:
        return {}
    result = {}
    color_elem = rpr.find(tag("color"))
    if color_elem is not None:
        result["color"] = color_elem.get(tag("val"), "").upper()
    strike_elem = rpr.find(tag("strike"))
    if strike_elem is not None:
        val = strike_elem.get(tag("val"), "true")
        result["strike"] = val not in ("false", "0")
    else:
        result["strike"] = False
    return result

def get_para_rpr(p_elem):
    ppr = p_elem.find(tag("pPr"))
    if ppr is None:
        return {}
    rpr = ppr.find(tag("rPr"))
    if rpr is None:
        return {}
    result = {}
    color_elem = rpr.find(tag("color"))
    if color_elem is not None:
        result["color"] = color_elem.get(tag("val"), "").upper()
    strike_elem = rpr.find(tag("strike"))
    if strike_elem is not None:
        val = strike_elem.get(tag("val"), "true")
        result["strike"] = val not in ("false", "0")
    else:
        result["strike"] = False
    return result

def get_text(r_elem):
    t = r_elem.find(tag("t"))
    if t is not None and t.text:
        return t.text
    return ""

tree = ET.parse(r"C:\Users\user\allen\_unpack_src\word\document.xml")
root = tree.getroot()
body = root.find(tag("body"))

segments = []

for p in body.findall(tag("p")):
    para_rpr = get_para_rpr(p)
    para_texts = []
    for r in p.findall(tag("r")):
        text = get_text(r)
        if not text:
            continue
        rpr = get_rpr(r)
        # inherit from paragraph if run has no explicit setting
        color = rpr.get("color", para_rpr.get("color", "000000"))
        strike = rpr.get("strike", para_rpr.get("strike", False))

        is_red = color == "FF0000"

        para_texts.append({
            "text": text,
            "color": color,
            "strike": strike,
            "is_red": is_red,
        })

    if para_texts:
        segments.append(para_texts)

# 統計不同格式組合
print("=== 格式分類統計 ===")
from collections import Counter
combos = Counter()
for para in segments:
    for run in para:
        key = f"red={run['is_red']}, strike={run['strike']}"
        combos[key] += 1

for k, v in combos.items():
    print(f"  {k}: {v} 個 run")

print()
print("=== 非紅色文字 (可能是黑色) ===")
for para in segments:
    for run in para:
        if not run["is_red"] and run["text"].strip():
            print(f"  color={run['color']}, strike={run['strike']}: [{run['text']}]")

print()
print("=== 紅色無刪除線 ===")
for para in segments:
    for run in para:
        if run["is_red"] and not run["strike"] and run["text"].strip():
            print(f"  [{run['text']}]")
