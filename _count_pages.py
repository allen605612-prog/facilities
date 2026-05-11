"""
推算各節的頁碼
策略：從已知頁碼反推。
已知：
  - pgNumType start=166（文件從第166頁開始）
  - P000 = page 166（系統開發及程式修改作業）
  - P093 = pageBreak → 程式及資料之存取作業 = 第170頁
  - P157 = pageBreak → 資料輸出入及處理作業 = 第173頁
  - P255 = pageBreak → 硬體及系統軟體之使用及維護作業
  - P313 = pageBreak → 系統復原計畫及測試作業
  - P384 = pageBreak → ???

  另外 P093~P156 = 3頁(170-172), P157~P201 = 3頁(173-175)

計算從 P201（最後一頁刪除 = 175）到 P255（下一個 pageBreak）有幾個「自然分頁」。
"""
import sys, xml.etree.ElementTree as ET
sys.stdout.reconfigure(encoding='utf-8')

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def tag(n): return f"{{{W}}}{n}"

tree = ET.parse(r"C:\Users\user\allen\_unpack_src\word\document.xml")
root = tree.getroot()
body = root.find(tag("body"))
paras = body.findall(tag("p"))

# 統計各區段的段落數
sections = [
    ("系統開發及程式修改作業+文書", 0, 92, 166, 169),      # 4 pages
    ("程式及資料之存取作業",        93, 156, 170, 172),     # 3 pages
    ("資料輸出入及處理作業",        157, 201, 173, 175),    # 3 pages
]

print("已知區段資訊：")
for name, start, end, pg_start, pg_end in sections:
    n = end - start + 1
    pages = pg_end - pg_start + 1
    print(f"  {name}: P{start:03d}~P{end:03d} = {n} paras, {pages} pages ({pg_start}~{pg_end})")
    print(f"    → 每頁約 {n/pages:.1f} 段落")

# 從P202到P254 = 53個段落（檔案及設備之安全作業）
# 分頁 P255, P313, P384

# 使用 15~20 para/page 估算
avg_per_page = 16  # 取中間值

print("\n推算後段頁碼（文件從175結束，P202起約第176頁）：")
# P202：從175頁結束後開始，可能175末頁就是下一節開頭，或換頁到176
# 由於刪除章節 P157~P201 共173~175（3頁），且P201後沒有顯式分頁，
# ◎檔案及設備之安全作業 可能緊接在175頁後半，但更可能在176頁開頭（樣式包含分頁）

# Let's just count the content roughly
paras_202_254 = 254 - 202 + 1  # 53 paras
paras_256_312 = 312 - 256 + 1  # 57 paras
paras_314_383 = 383 - 314 + 1  # 70 paras
paras_385_end = len(paras) - 385  # remaining

print(f"\n  P202-P254 (◎檔案及設備之安全作業):   {paras_202_254} 段落")
print(f"  P256-P312 (◎硬體及系統軟體使用維護): {paras_256_312} 段落")
print(f"  P314-P383 (◎系統復原計畫及測試):     {paras_314_383} 段落")
print(f"  P385-end  (◎資訊安全之檢查作業):     {paras_385_end} 段落")
print(f"  (文件共 {len(paras)} 個段落)")

# 根據每頁16段估算
p176 = 176  # 檔案及設備之安全作業 起始頁
n1 = paras_202_254 / avg_per_page
p2 = int(p176 + n1 + 0.5)  # 硬體 start
n2 = paras_256_312 / avg_per_page
p3 = int(p2 + n2 + 0.5)
n3 = paras_314_383 / avg_per_page
p4 = int(p3 + n3 + 0.5)

print(f"\n估算結果（每頁 {avg_per_page} 段）：")
print(f"  ◎檔案及設備之安全作業:      第 {p176} 頁 ({paras_202_254}段，約{n1:.1f}頁)")
print(f"  ◎硬體及系統軟體之使用及維護: 第 {p2} 頁 ({paras_256_312}段，約{n2:.1f}頁)")
print(f"  ◎系統復原計畫及測試:         第 {p3} 頁 ({paras_314_383}段，約{n3:.1f}頁)")
print(f"  ◎資訊安全之檢查:             第 {p4} 頁")
