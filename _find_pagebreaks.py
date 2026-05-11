"""找出文件中的分頁符號位置，估算各節起始頁碼"""
import sys, xml.etree.ElementTree as ET
sys.stdout.reconfigure(encoding='utf-8')

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def tag(n): return f"{{{W}}}{n}"

tree = ET.parse(r"C:\Users\user\allen\_unpack_src\word\document.xml")
root = tree.getroot()
body = root.find(tag("body"))

page_breaks = []
for i, p in enumerate(body.findall(tag("p"))):
    # 找明確分頁
    for r in p.findall(tag("r")):
        for br in r.findall(tag("br")):
            if br.get(tag("type")) == "page":
                page_breaks.append(i)
    # 找段落屬性中的 pageBreakBefore
    ppr = p.find(tag("pPr"))
    if ppr is not None:
        pb = ppr.find(tag("pageBreakBefore"))
        if pb is not None and pb.get(tag("val"), "true") not in ("false", "0"):
            page_breaks.append(i)

print("分頁位置（段落索引）:", page_breaks)
print(f"共 {len(page_breaks)} 個分頁符，文件起始頁 166")
print()

# 輸出分頁附近的段落文字
for idx in page_breaks:
    paras = body.findall(tag("p"))
    texts = []
    for off in range(0, min(3, len(paras)-idx)):
        p = paras[idx+off]
        t = "".join(
            r.find(tag("t")).text
            for r in p.findall(tag("r"))
            if r.find(tag("t")) is not None and r.find(tag("t")).text
        )
        if t.strip():
            texts.append(t[:60])
    print(f"P{idx:03d}: {''.join(texts[:1])}")
