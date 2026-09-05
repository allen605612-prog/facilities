"""
以複製檔為基底，移除空佔位列，注入 11 筆新列，輸出至目的地。
完全保留原始 XML 命名空間與格式。
"""
import sys, zipfile, shutil
from copy import deepcopy
from lxml import etree
sys.stdout.reconfigure(encoding='utf-8')

COPY = r"D:\D\114設備組\修正對照表-內控-設備組 - 複製.docx"
DST  = r"D:\D\114設備組\修正對照表-內控-設備組.docx"
TMP  = r"C:\Users\user\allen\_from_copy_tmp.docx"

W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
def w(n):   return f"{{{W}}}{n}"
def w14(n): return f"{{{W14}}}{n}"

# ── 11 筆修正資料 ─────────────────────────────────────────────────────────
changes = [
    ("3.4.異地備援：",
     "3.4.異地備援：(因經費與管理的緣故尚未施行)", "P177",
     "異地備援措施已逐步實施，刪除「尚未施行」括號說明文字"),
    ("3.4.1.檔案備份資料存放於異地或雲端，是否具備保全性及完整記載資料存放紀錄。",
     "3.4.1.檔案備份資料存放於異地，是否具備保全性及完整記載資料存放紀錄。", "P177",
     "因應雲端備份趨勢，增列「雲端」作為備援存放選項"),
    ("3.4.2.每學期進行一次異地備份還原演練，確保回復程序能在指定時間內完成復原作業程序。",
     "3.4.2.於異地設置實體主機，是否定期檢查測試回復程序，確保回復程序能在指定時間內完成復原作業程序。", "P177",
     "明確規定異地備援測試頻率，以每學期辦理一次演練取代模糊之「定期」規定"),
    ("5.1.正心中學電腦主機房管理辦法。",
     "5.1.正心中學電子計算機中心電腦主機房管理辦法。", "P177",
     "配合組織名稱調整，刪除「電子計算機中心」，精簡文件名稱"),
    ("2.3.3.設備組不定期實體稽查及程式軟體稽查是否有使用未經授權之軟體及其他不當之軟體。",
     "2.3.3.電子計算機中心不定期實體稽查及程式軟體稽查是否有使用未經授權之軟體及其他不當之軟體。", "P179",
     "配合組織名稱調整，將「電子計算機中心」更正為「設備組」"),
    ("刪除",
     "．由不同單位人員參加成立緊急應變小組，並加強訓練其緊急事故應變能力。（系統復原計畫及測試作業 2.2）", "P182",
     "目前無此業務，刪除緊急應變小組相關作業程序條文"),
    ("．重大事故硬體或軟體復原，應由庶務組與電腦廠商簽訂重大意外事故系統復原合約。合約內容應包含修護完成交期、保固期間、違約損失賠償罰則及應變方式等條文。",
     "．重大事故硬體或軟體復原，應由設備組與電腦廠商簽訂重大意外事故系統復原合約。合約內容應包含修護完成交期、保固期間、違約損失賠償罰則及應變方式等條文。", "P182",
     "依採購及合約作業職責分工，簽訂廠商合約由庶務組辦理，非設備組職掌"),
    ("．設備組人員或維修外包廠商應將測試結果對使用方詳述說明。",
     "．設備組人員應將測試結果詳述說明。併同測試資料及程式規範送交教務主任核示後建檔。", "P182",
     "增列「維修外包廠商」為說明義務人，明確對象為「使用方」；刪除送交教務主任核示建檔程序"),
    ("刪除",
     "．是否規劃由不同單位人員參加成立緊急應變小組，並加強訓練其緊急應變能力。（系統復原計畫及測試作業 3.2）", "P183",
     "配合作業程序修訂，同步刪除緊急應變小組之對應控制重點"),
    ("刪除",
     "3.3.本校郵件伺服器是否設置防火牆及防毒軟體，以防止駭客或電腦病毒之侵害。(現在都是Google雲端檢查)", "P184",
     "本校郵件服務已改用Google雲端，由Google提供安全防護，無需自設防火牆，刪除此控制重點"),
    ("3.5.網管人員是否定期檢視郵件伺服器上郵件收發情形，若有異常狀況是否陳報權責主管處理。",
     "3.5.設備組人員是否定期檢視郵件伺服器上郵件收發情形，若有異常狀況是否陳報權責主管處理。", "P185",
     "明確指定由「網管人員」負責郵件伺服器監控，使職責劃分更清晰"),
]

# ── 以複製檔的「列1」為格式參考，複製出新列 ──────────────────────────────
with zipfile.ZipFile(COPY) as z:
    raw_bytes = {name: z.read(name) for name in z.namelist()}

doc_xml_bytes = raw_bytes["word/document.xml"]
tree = etree.fromstring(doc_xml_bytes)
body = tree.find(w('body'))
tbl  = body.find(w('tbl'))
rows = tbl.findall(w('tr'))
print(f"原始共 {len(rows)} 列")

# 取樣本列（列1 = 刪除/系統開發那列）當格式模板
sample = rows[1]

def make_row(xiu, xian, page, shuo):
    """仿照 sample 列格式，建立新資料列"""
    is_del = (xiu == "刪除")
    tr = etree.Element(w('tr'))
    trPr = etree.SubElement(tr, w('trPr'))
    etree.SubElement(trPr, w('trHeight')).set(w('val'), '1440')

    # ── 欄 1：修正規定 ──────────────────────────────────────────────────
    sample_tc1 = sample.findall(w('tc'))[0]
    tc1 = etree.SubElement(tr, w('tc'))
    tc1Pr = etree.SubElement(tc1, w('tcPr'))
    tcW1 = etree.SubElement(tc1Pr, w('tcW'))
    tcW1.set(w('w'), '5983'); tcW1.set(w('type'), 'dxa')

    p1 = etree.SubElement(tc1, w('p'))
    pPr1 = etree.SubElement(p1, w('pPr'))
    sp1 = etree.SubElement(pPr1, w('spacing'))
    sp1.set(w('before'),'100'); sp1.set(w('beforeAutospacing'),'1')
    sp1.set(w('after'),'100');  sp1.set(w('afterAutospacing'),'1')
    sp1.set(w('line'),'240');   sp1.set(w('lineRule'),'auto')
    etree.SubElement(pPr1, w('jc')).set(w('val'), 'left')
    rPr1p = etree.SubElement(pPr1, w('rPr'))
    f1p = etree.SubElement(rPr1p, w('rFonts'))
    f1p.set(w('ascii'),'標楷體'); f1p.set(w('eastAsia'),'標楷體'); f1p.set(w('hAnsi'),'標楷體')
    etree.SubElement(rPr1p, w('sz')).set(w('val'),'28')
    etree.SubElement(rPr1p, w('szCs')).set(w('val'),'28')

    r1 = etree.SubElement(p1, w('r'))
    rPr1 = etree.SubElement(r1, w('rPr'))
    f1 = etree.SubElement(rPr1, w('rFonts'))
    f1.set(w('ascii'),'標楷體'); f1.set(w('eastAsia'),'標楷體')
    f1.set(w('hAnsi'),'標楷體'); f1.set(w('hint'),'eastAsia')
    etree.SubElement(rPr1, w('sz')).set(w('val'),'28')
    etree.SubElement(rPr1, w('szCs')).set(w('val'),'28')
    if is_del:
        etree.SubElement(rPr1, w('u')).set(w('val'),'single')
    t1 = etree.SubElement(r1, w('t'))
    t1.text = xiu

    # ── 欄 2：現行規定 ──────────────────────────────────────────────────
    tc2 = etree.SubElement(tr, w('tc'))
    tc2Pr = etree.SubElement(tc2, w('tcPr'))
    tcW2 = etree.SubElement(tc2Pr, w('tcW'))
    tcW2.set(w('w'), '5670'); tcW2.set(w('type'), 'dxa')

    p2 = etree.SubElement(tc2, w('p'))
    pPr2 = etree.SubElement(p2, w('pPr'))
    etree.SubElement(pPr2, w('pStyle')).set(w('val'), '2')
    ind2 = etree.SubElement(pPr2, w('ind'))
    ind2.set(w('leftChars'),'200'); ind2.set(w('left'),'440')
    ind2.set(w('firstLineChars'),'4'); ind2.set(w('firstLine'),'11')
    rPr2p = etree.SubElement(pPr2, w('rPr'))
    f2p = etree.SubElement(rPr2p, w('rFonts'))
    f2p.set(w('ascii'),'標楷體'); f2p.set(w('eastAsia'),'標楷體')
    f2p.set(w('hAnsi'),'標楷體'); f2p.set(w('hint'),'eastAsia')
    etree.SubElement(rPr2p, w('szCs')).set(w('val'),'28')

    r2a = etree.SubElement(p2, w('r'))
    rPr2a = etree.SubElement(r2a, w('rPr'))
    f2a = etree.SubElement(rPr2a, w('rFonts'))
    f2a.set(w('ascii'),'標楷體'); f2a.set(w('eastAsia'),'標楷體')
    f2a.set(w('hAnsi'),'標楷體'); f2a.set(w('hint'),'eastAsia')
    etree.SubElement(rPr2a, w('szCs')).set(w('val'),'28')
    t2a = etree.SubElement(r2a, w('t'))
    t2a.text = xian
    t2a.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

    r2b = etree.SubElement(p2, w('r'))
    rPr2b = etree.SubElement(r2b, w('rPr'))
    f2b = etree.SubElement(rPr2b, w('rFonts'))
    f2b.set(w('ascii'),'標楷體'); f2b.set(w('eastAsia'),'標楷體'); f2b.set(w('hAnsi'),'標楷體')
    etree.SubElement(rPr2b, w('szCs')).set(w('val'),'28')
    etree.SubElement(r2b, w('br'))
    t2b = etree.SubElement(r2b, w('t'))
    t2b.text = page

    # ── 欄 3：說明 ──────────────────────────────────────────────────────
    tc3 = etree.SubElement(tr, w('tc'))
    tc3Pr = etree.SubElement(tc3, w('tcPr'))
    tcW3 = etree.SubElement(tc3Pr, w('tcW'))
    tcW3.set(w('w'), '2551'); tcW3.set(w('type'), 'dxa')

    p3 = etree.SubElement(tc3, w('p'))
    pPr3 = etree.SubElement(p3, w('pPr'))
    sp3 = etree.SubElement(pPr3, w('spacing'))
    sp3.set(w('before'),'100'); sp3.set(w('beforeAutospacing'),'1')
    sp3.set(w('after'),'100');  sp3.set(w('afterAutospacing'),'1')
    sp3.set(w('line'),'240');   sp3.set(w('lineRule'),'auto')
    etree.SubElement(pPr3, w('jc')).set(w('val'), 'center')
    rPr3p = etree.SubElement(pPr3, w('rPr'))
    f3p = etree.SubElement(rPr3p, w('rFonts'))
    f3p.set(w('ascii'),'標楷體'); f3p.set(w('eastAsia'),'標楷體'); f3p.set(w('hAnsi'),'標楷體')
    etree.SubElement(rPr3p, w('sz')).set(w('val'),'28')
    etree.SubElement(rPr3p, w('szCs')).set(w('val'),'28')

    r3 = etree.SubElement(p3, w('r'))
    rPr3 = etree.SubElement(r3, w('rPr'))
    f3 = etree.SubElement(rPr3, w('rFonts'))
    f3.set(w('ascii'),'標楷體'); f3.set(w('eastAsia'),'標楷體')
    f3.set(w('hAnsi'),'標楷體'); f3.set(w('hint'),'eastAsia')
    etree.SubElement(rPr3, w('sz')).set(w('val'),'28')
    etree.SubElement(rPr3, w('szCs')).set(w('val'),'28')
    t3 = etree.SubElement(r3, w('t'))
    t3.text = shuo

    return tr

# ── 移除空佔位列（最後一列，無文字） ─────────────────────────────────────
last = rows[-1]
last_text = "".join(x.text for x in last.iter(w('t')) if x.text).strip()
if not last_text:
    tbl.remove(last)
    print("移除空佔位列")

# ── 注入新列 ──────────────────────────────────────────────────────────────
for c in changes:
    tbl.append(make_row(*c))
print(f"注入後共 {len(tbl.findall(w('tr')))} 列")

# ── 改為 A4 直向，按比例縮放欄寬 ─────────────────────────────────────────
# A4 直向：11906 x 16838，邊距 1134 each → 可用寬 9638 dxa
MARGIN = 1134
PG_W, PG_H = 11906, 16838
CONTENT_W = PG_W - MARGIN * 2   # 9638

ORIG_COLS = [5983, 5670, 2551]   # 原始欄寬（總計 14204）
ORIG_TOTAL = sum(ORIG_COLS)
NEW_COLS = [round(c / ORIG_TOTAL * CONTENT_W) for c in ORIG_COLS]
# 修正捨入誤差
NEW_COLS[-1] = CONTENT_W - sum(NEW_COLS[:-1])

# 更新 tblPr / tblW
tblPr = tbl.find(w('tblPr'))
tblW_el = tblPr.find(w('tblW'))
tblW_el.set(w('w'), str(CONTENT_W))
tblInd_el = tblPr.find(w('tblInd'))
if tblInd_el is not None:
    tblInd_el.set(w('w'), '0')

# 更新 tblGrid
tblGrid = tbl.find(w('tblGrid'))
if tblGrid is not None:
    for col_el, new_w in zip(tblGrid.findall(w('gridCol')), NEW_COLS):
        col_el.set(w('w'), str(new_w))

# 更新每列每格的 tcW
for tr in tbl.findall(w('tr')):
    cells = tr.findall(w('tc'))
    for tc, new_w in zip(cells, NEW_COLS):
        tcPr = tc.find(w('tcPr'))
        if tcPr is not None:
            tcW_el = tcPr.find(w('tcW'))
            if tcW_el is not None:
                tcW_el.set(w('w'), str(new_w))

print(f"欄寬：{NEW_COLS}（合計 {sum(NEW_COLS)}）")

# 更新 sectPr → A4 直向
sectPr = body.find(w('sectPr'))
pgSz = sectPr.find(w('pgSz'))
pgSz.set(w('w'), str(PG_W))
pgSz.set(w('h'), str(PG_H))
for attr in (w('orient'), w('code')):
    if attr in pgSz.attrib:
        del pgSz.attrib[attr]
pgMar = sectPr.find(w('pgMar'))
for side in ('top','right','bottom','left'):
    pgMar.set(w(side), str(MARGIN))

# ── 序列化並打包 ──────────────────────────────────────────────────────────
new_doc_bytes = etree.tostring(tree, xml_declaration=True,
                               encoding='UTF-8', standalone=True)

with zipfile.ZipFile(TMP, 'w', zipfile.ZIP_DEFLATED) as zout:
    for name, data in raw_bytes.items():
        if name == "word/document.xml":
            zout.writestr(name, new_doc_bytes)
        else:
            zout.writestr(name, data)

shutil.copy2(TMP, DST)
print(f"✓ 已輸出：{DST}")
