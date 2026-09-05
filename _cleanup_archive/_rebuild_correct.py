"""
從乾淨模板重建修正對照表，完全仿照預填格式：
- 標楷體 sz=28
- 頁碼用 <w:br/> 內嵌在同一段落（格式: P177）
- 修正規定「刪除」加底線
"""
import sys, os, zipfile, shutil, xml.etree.ElementTree as ET
sys.stdout.reconfigure(encoding='utf-8')

SRC_FOLDER = r"C:\Users\user\allen\_unpack_cmp"
DST_DOCX   = r"D:\D\114設備組\修正對照表-內控-設備組.docx"
TMP_DOCX   = r"C:\Users\user\allen\_rebuilt.docx"
DOC_XML    = os.path.join(SRC_FOLDER, "word", "document.xml")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace("", W)
def tag(n): return f"{{{W}}}{n}"

# ── 11 筆修正資料（修正規定, 現行規定, 說明, 頁碼）──────────────────────────
changes = [
    ("3.4.異地備援：",
     "3.4.異地備援：(因經費與管理的緣故尚未施行)",
     "異地備援措施已逐步實施，刪除「尚未施行」括號說明文字",
     "P177"),
    ("3.4.1.檔案備份資料存放於異地或雲端，是否具備保全性及完整記載資料存放紀錄。",
     "3.4.1.檔案備份資料存放於異地，是否具備保全性及完整記載資料存放紀錄。",
     "因應雲端備份趨勢，增列「雲端」作為備援存放選項",
     "P177"),
    ("3.4.2.每學期進行一次異地備份還原演練，確保回復程序能在指定時間內完成復原作業程序。",
     "3.4.2.於異地設置實體主機，是否定期檢查測試回復程序，確保回復程序能在指定時間內完成復原作業程序。",
     "明確規定異地備援測試頻率，以每學期辦理一次演練取代模糊之「定期」規定",
     "P177"),
    ("5.1.正心中學電腦主機房管理辦法。",
     "5.1.正心中學電子計算機中心電腦主機房管理辦法。",
     "配合組織名稱調整，刪除「電子計算機中心」，精簡文件名稱",
     "P177"),
    ("2.3.3.設備組不定期實體稽查及程式軟體稽查是否有使用未經授權之軟體及其他不當之軟體。",
     "2.3.3.電子計算機中心不定期實體稽查及程式軟體稽查是否有使用未經授權之軟體及其他不當之軟體。",
     "配合組織名稱調整，將「電子計算機中心」更正為「設備組」",
     "P179"),
    ("刪除",
     "．由不同單位人員參加成立緊急應變小組，並加強訓練其緊急事故應變能力。（系統復原計畫及測試作業 2.2）",
     "目前無此業務，刪除緊急應變小組相關作業程序條文",
     "P182"),
    ("．重大事故硬體或軟體復原，應由庶務組與電腦廠商簽訂重大意外事故系統復原合約。合約內容應包含修護完成交期、保固期間、違約損失賠償罰則及應變方式等條文。",
     "．重大事故硬體或軟體復原，應由設備組與電腦廠商簽訂重大意外事故系統復原合約。合約內容應包含修護完成交期、保固期間、違約損失賠償罰則及應變方式等條文。",
     "依採購及合約作業職責分工，簽訂廠商合約由庶務組辦理，非設備組職掌",
     "P182"),
    ("．設備組人員或維修外包廠商應將測試結果對使用方詳述說明。",
     "．設備組人員應將測試結果詳述說明。併同測試資料及程式規範送交教務主任核示後建檔。",
     "增列「維修外包廠商」為說明義務人，明確對象為「使用方」；刪除送交教務主任核示建檔程序",
     "P182"),
    ("刪除",
     "．是否規劃由不同單位人員參加成立緊急應變小組，並加強訓練其緊急應變能力。（系統復原計畫及測試作業 3.2）",
     "配合作業程序修訂，同步刪除緊急應變小組之對應控制重點",
     "P183"),
    ("刪除",
     "3.3.本校郵件伺服器是否設置防火牆及防毒軟體，以防止駭客或電腦病毒之侵害。(現在都是Google雲端檢查)",
     "本校郵件服務已改用Google雲端，由Google提供安全防護，無需自設防火牆，刪除此控制重點",
     "P184"),
    ("3.5.網管人員是否定期檢視郵件伺服器上郵件收發情形，若有異常狀況是否陳報權責主管處理。",
     "3.5.設備組人員是否定期檢視郵件伺服器上郵件收發情形，若有異常狀況是否陳報權責主管處理。",
     "明確指定由「網管人員」負責郵件伺服器監控，使職責劃分更清晰",
     "P185"),
]

def safe(t):
    return t.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

# ── 建立 <w:tc> 元素 ────────────────────────────────────────────────────────

def make_tc_col1(text):
    """修正規定欄（5983 dxa）：刪除加底線；其他正常"""
    is_del = (text == "刪除")
    u_ppr = "<w:u w:val=\"single\"/>" if is_del else ""
    u_rpr = "<w:u w:val=\"single\"/>" if is_del else ""
    xml = (
        f'<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:tcPr><w:tcW w:w="5983" w:type="dxa"/></w:tcPr>'
        f'<w:p>'
        f'<w:pPr>'
        f'<w:snapToGrid w:val="0"/>'
        f'<w:spacing w:line="240" w:lineRule="atLeast"/>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="標楷體" w:eastAsia="標楷體" w:hAnsi="標楷體" w:hint="eastAsia"/>'
        f'<w:sz w:val="28"/><w:szCs w:val="28"/>{u_ppr}'
        f'</w:rPr>'
        f'</w:pPr>'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="標楷體" w:eastAsia="標楷體" w:hAnsi="標楷體" w:hint="eastAsia"/>'
        f'<w:sz w:val="28"/><w:szCs w:val="28"/>{u_rpr}'
        f'</w:rPr>'
        f'<w:t>{safe(text)}</w:t>'
        f'</w:r>'
        f'</w:p>'
        f'</w:tc>'
    )
    return ET.fromstring(xml)

def make_tc_col2(content, page):
    """現行規定欄（5670 dxa）：內文 + <w:br/> + 頁碼"""
    xml = (
        f'<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:tcPr><w:tcW w:w="5670" w:type="dxa"/></w:tcPr>'
        f'<w:p>'
        f'<w:pPr>'
        f'<w:pStyle w:val="2"/>'
        f'<w:ind w:leftChars="200" w:left="440" w:firstLineChars="4" w:firstLine="11"/>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="標楷體" w:eastAsia="標楷體" w:hAnsi="標楷體" w:hint="eastAsia"/>'
        f'<w:color w:val="000000" w:themeColor="text1"/>'
        f'<w:kern w:val="0"/><w:szCs w:val="28"/>'
        f'</w:rPr>'
        f'</w:pPr>'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="標楷體" w:eastAsia="標楷體" w:hAnsi="標楷體" w:hint="eastAsia"/>'
        f'<w:color w:val="000000" w:themeColor="text1"/>'
        f'<w:kern w:val="0"/><w:szCs w:val="28"/>'
        f'</w:rPr>'
        f'<w:t xml:space="preserve">{safe(content)}</w:t>'
        f'</w:r>'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="標楷體" w:eastAsia="標楷體" w:hAnsi="標楷體"/>'
        f'<w:color w:val="000000" w:themeColor="text1"/>'
        f'<w:kern w:val="0"/><w:szCs w:val="28"/>'
        f'</w:rPr>'
        f'<w:br/>'
        f'<w:t>{safe(page)}</w:t>'
        f'</w:r>'
        f'</w:p>'
        f'</w:tc>'
    )
    return ET.fromstring(xml)

def make_tc_col3(text):
    """說明欄（2551 dxa）"""
    xml = (
        f'<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:tcPr><w:tcW w:w="2551" w:type="dxa"/></w:tcPr>'
        f'<w:p>'
        f'<w:pPr>'
        f'<w:pStyle w:val="2"/>'
        f'<w:ind w:left="-108"/>'
        f'<w:jc w:val="left"/>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="標楷體" w:eastAsia="標楷體" w:hAnsi="標楷體" w:hint="eastAsia"/>'
        f'<w:szCs w:val="28"/>'
        f'</w:rPr>'
        f'</w:pPr>'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="標楷體" w:eastAsia="標楷體" w:hAnsi="標楷體" w:hint="eastAsia"/>'
        f'<w:szCs w:val="28"/>'
        f'</w:rPr>'
        f'<w:t>{safe(text)}</w:t>'
        f'</w:r>'
        f'</w:p>'
        f'</w:tc>'
    )
    return ET.fromstring(xml)

def make_row(xiu, xian, shuo, page):
    tr_xml = (
        f'<w:tr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:trPr><w:trHeight w:val="1011"/></w:trPr>'
        f'</w:tr>'
    )
    tr = ET.fromstring(tr_xml)
    tr.append(make_tc_col1(xiu))
    tr.append(make_tc_col2(xian, page))
    tr.append(make_tc_col3(shuo))
    return tr

# ── 修改 document.xml ───────────────────────────────────────────────────────
tree = ET.parse(DOC_XML)
root = tree.getroot()
body = root.find(tag("body"))
tbl  = body.find(tag("tbl"))
rows = tbl.findall(tag("tr"))
print(f"原始共 {len(rows)} 列")

# 移除 row 5（空的佔位列）
if len(rows) >= 6:
    empty = rows[5]
    cells = empty.findall(tag("tc"))
    texts = "".join(
        t.text for tc in cells for r in tc.iter(tag("r"))
        for t in [r.find(tag("t"))] if t is not None and t.text
    )
    if not texts.strip():
        tbl.remove(empty)
        print("移除空佔位列")

for c in changes:
    tbl.append(make_row(*c))

final_rows = tbl.findall(tag("tr"))
print(f"加入後共 {len(final_rows)} 列")

# 寫回 XML
tree.write(DOC_XML, encoding="unicode", xml_declaration=False)
with open(DOC_XML, "r", encoding="utf-8") as f:
    content = f.read()
with open(DOC_XML, "w", encoding="utf-8") as f:
    f.write('<?xml version="1.0" encoding="UTF-8"?>' + content)

# ── 用 Python zipfile 重新打包 ───────────────────────────────────────────────
import zipfile as zf

# 讀取所有需要打包的檔案
files_to_pack = []
for dirpath, dirnames, filenames in os.walk(SRC_FOLDER):
    for fname in filenames:
        full = os.path.join(dirpath, fname)
        arcname = os.path.relpath(full, SRC_FOLDER).replace("\\", "/")
        files_to_pack.append((full, arcname))

# 寫入臨時 docx
with zf.ZipFile(TMP_DOCX, "w", zf.ZIP_DEFLATED) as z:
    for full, arcname in files_to_pack:
        z.write(full, arcname)

# 複製到目的地（處理中文路徑）
shutil.copy2(TMP_DOCX, DST_DOCX)
print(f"✓ 已輸出：{DST_DOCX}")
print(f"✓ 臨時檔：{TMP_DOCX}")
