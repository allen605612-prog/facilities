"""
修正對照表：保留前5列（原始列），清除重複列，加入含頁碼的新列
"""
import sys, xml.etree.ElementTree as ET
sys.stdout.reconfigure(encoding='utf-8')

ET.register_namespace("", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

# 改用 _unpack_dst2（第二次 unpack），移除重複列並重寫
DST_XML = r"C:\Users\user\allen\_unpack_dst2\word\document.xml"

tree = ET.parse(DST_XML)
root = tree.getroot()

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def tag(n): return f"{{{W}}}{n}"

body = root.find(tag("body"))
tbl = body.find(tag("tbl"))
rows = tbl.findall(tag("tr"))

print(f"目前共 {len(rows)} 列")

# 只保留前 6 列（標題 + 4 原始資料列 + 1 空列）
KEEP = 6
for r in rows[KEEP:]:
    tbl.remove(r)

print(f"清除重複列後剩 {len(tbl.findall(tag('tr')))} 列")

# 新列資料
changes = [
    ("3.4.異地備援：",
     "3.4.異地備援：(因經費與管理的緣故尚未施行)",
     "異地備援措施已逐步實施，刪除「尚未施行」括號說明文字",
     "第177頁"),
    ("3.4.1.檔案備份資料存放於異地或雲端，是否具備保全性及完整記載資料存放紀錄。",
     "3.4.1.檔案備份資料存放於異地，是否具備保全性及完整記載資料存放紀錄。",
     "因應雲端備份趨勢，增列「雲端」作為備援存放選項",
     "第177頁"),
    ("3.4.2.每學期進行一次異地備份還原演練，確保回復程序能在指定時間內完成復原作業程序。",
     "3.4.2.於異地設置實體主機，是否定期檢查測試回復程序，確保回復程序能在指定時間內完成復原作業程序。",
     "明確規定異地備援測試頻率，以每學期辦理一次演練取代模糊之「定期」規定",
     "第177頁"),
    ("5.1.正心中學電腦主機房管理辦法。",
     "5.1.正心中學電子計算機中心電腦主機房管理辦法。",
     "配合組織名稱調整，刪除「電子計算機中心」，精簡文件名稱",
     "第177頁"),
    ("2.3.3.設備組不定期實體稽查及程式軟體稽查是否有使用未經授權之軟體及其他不當之軟體。",
     "2.3.3.電子計算機中心不定期實體稽查及程式軟體稽查是否有使用未經授權之軟體及其他不當之軟體。",
     "配合組織名稱調整，將「電子計算機中心」更正為「設備組」",
     "第179頁"),
    ("刪除",
     "．由不同單位人員參加成立緊急應變小組，並加強訓練其緊急事故應變能力。（系統復原計畫及測試作業 2.2）",
     "目前無此業務，刪除緊急應變小組相關作業程序條文",
     "第182頁"),
    ("．重大事故硬體或軟體復原，應由庶務組與電腦廠商簽訂重大意外事故系統復原合約。合約內容應包含修護完成交期、保固期間、違約損失賠償罰則及應變方式等條文。",
     "．重大事故硬體或軟體復原，應由設備組與電腦廠商簽訂重大意外事故系統復原合約。合約內容應包含修護完成交期、保固期間、違約損失賠償罰則及應變方式等條文。",
     "依採購及合約作業職責分工，簽訂廠商合約由庶務組辦理，非設備組職掌",
     "第182頁"),
    ("．設備組人員或維修外包廠商應將測試結果對使用方詳述說明。",
     "．設備組人員應將測試結果詳述說明。併同測試資料及程式規範送交教務主任核示後建檔。",
     "增列「維修外包廠商」為說明義務人，明確對象為「使用方」；刪除送交教務主任核示建檔程序",
     "第182頁"),
    ("刪除",
     "．是否規劃由不同單位人員參加成立緊急應變小組，並加強訓練其緊急應變能力。（系統復原計畫及測試作業 3.2）",
     "配合作業程序修訂，同步刪除緊急應變小組之對應控制重點",
     "第183頁"),
    ("刪除",
     "3.3.本校郵件伺服器是否設置防火牆及防毒軟體，以防止駭客或電腦病毒之侵害。(現在都是Google雲端檢查)",
     "本校郵件服務已改用Google雲端，由Google提供安全防護，無需自設防火牆，刪除此控制重點",
     "第184頁"),
    ("3.5.網管人員是否定期檢視郵件伺服器上郵件收發情形，若有異常狀況是否陳報權責主管處理。",
     "3.5.設備組人員是否定期檢視郵件伺服器上郵件收發情形，若有異常狀況是否陳報權責主管處理。",
     "明確指定由「網管人員」負責郵件伺服器監控，使職責劃分更清晰",
     "第185頁"),
]

def safe(t): return t.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def mk_p(text, center=False, small=False, underline=False):
    sz = "20" if small else "24"
    jc = "center" if center else "left"
    u = "<w:u w:val=\"single\"/>" if underline else ""
    return (f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:pPr><w:jc w:val="{jc}"/><w:rPr>'
            f'<w:rFonts w:ascii="標楷體" w:eastAsia="標楷體" w:hAnsi="標楷體"/>'
            f'<w:sz w:val="{sz}"/>{u}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>'
            f'<w:rFonts w:ascii="標楷體" w:eastAsia="標楷體" w:hAnsi="標楷體"/>'
            f'<w:sz w:val="{sz}"/>{u}</w:rPr>'
            f'<w:t xml:space="preserve">{safe(text)}</w:t></w:r></w:p>')

def mk_tc(w_val, *para_xmls):
    paras = "\n".join(para_xmls)
    return (f'<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:tcPr><w:tcW w:w="{w_val}" w:type="dxa"/></w:tcPr>'
            f'{paras}</w:tc>')

def mk_row(xiu, xian, shuo, page):
    if xiu == "刪除":
        c1 = mk_tc(5983, mk_p("刪除", underline=True))
    else:
        c1 = mk_tc(5983, mk_p(xiu))
    c2 = mk_tc(5670, mk_p(xian), mk_p(page, center=True, small=True))
    c3 = mk_tc(2551, mk_p(shuo))
    return ET.fromstring(
        f'<w:tr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:trPr><w:trHeight w:val="1440" w:hRule="atLeast"/></w:trPr>'
        f'{c1}{c2}{c3}</w:tr>'
    )

for c in changes:
    tbl.append(mk_row(*c))

print(f"加入新列後共 {len(tbl.findall(tag('tr')))} 列")

# 直接寫回（保持 UTF-8）
tree.write(DST_XML, encoding="unicode", xml_declaration=False)

# 加回 XML 宣告
with open(DST_XML, "r", encoding="utf-8") as f:
    content = f.read()
with open(DST_XML, "w", encoding="utf-8") as f:
    f.write('<?xml version="1.0" encoding="UTF-8"?>' + content)

print("✓ XML 已更新")
