"""
用 lxml 把 document.xml 的命名空間前綴修正為正確的 w:, mc:, w14: ...
並重新打包成可以被 Word 開啟的 docx。
"""
import sys, zipfile, shutil, io
from copy import deepcopy
from lxml import etree
sys.stdout.reconfigure(encoding='utf-8')

SRC_DOCX = r"C:\Users\user\allen\_output.docx"
TMP_DOCX = r"C:\Users\user\allen\_fixed.docx"
DST_DOCX = r"D:\D\114設備組\修正對照表-內控-設備組.docx"

# ── 命名空間定義 ──────────────────────────────────────────────────────────
W    = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC   = "http://schemas.openxmlformats.org/markup-compatibility/2006"
W14  = "http://schemas.microsoft.com/office/word/2010/wordml"
W15  = "http://schemas.microsoft.com/office/word/2012/wordml"
W16SE= "http://schemas.microsoft.com/office/word/2015/wordml/symex"
W16CID="http://schemas.microsoft.com/office/word/2016/wordml/cid"
W16  = "http://schemas.microsoft.com/office/word/2018/wordml"
W16CEX="http://schemas.microsoft.com/office/word/2018/wordml/cex"
W16SDTDH="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
WP   = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
WPS  = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
A    = "http://schemas.openxmlformats.org/drawingml/2006/main"
V    = "urn:schemas-microsoft-com:vml"
O    = "urn:schemas-microsoft-com:office:office"
R    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

NSMAP = {
    'w': W, 'mc': MC, 'w14': W14, 'w15': W15,
    'w16se': W16SE, 'w16cid': W16CID, 'w16': W16,
    'w16cex': W16CEX, 'w16sdtdh': W16SDTDH,
    'wp14': WP14, 'wp': WP, 'wps': WPS,
    'a': A, 'v': V, 'o': O, 'r': R,
}

def qn(ns, name): return f"{{{ns}}}{name}"

# ── 修正資料 ──────────────────────────────────────────────────────────────
changes = [
    ("3.4.異地備援：",
     "3.4.異地備援：(因經費與管理的緣故尚未施行)",
     "異地備援措施已逐步實施，刪除「尚未施行」括號說明文字", "P177"),
    ("3.4.1.檔案備份資料存放於異地或雲端，是否具備保全性及完整記載資料存放紀錄。",
     "3.4.1.檔案備份資料存放於異地，是否具備保全性及完整記載資料存放紀錄。",
     "因應雲端備份趨勢，增列「雲端」作為備援存放選項", "P177"),
    ("3.4.2.每學期進行一次異地備份還原演練，確保回復程序能在指定時間內完成復原作業程序。",
     "3.4.2.於異地設置實體主機，是否定期檢查測試回復程序，確保回復程序能在指定時間內完成復原作業程序。",
     "明確規定異地備援測試頻率，以每學期辦理一次演練取代模糊之「定期」規定", "P177"),
    ("5.1.正心中學電腦主機房管理辦法。",
     "5.1.正心中學電子計算機中心電腦主機房管理辦法。",
     "配合組織名稱調整，刪除「電子計算機中心」，精簡文件名稱", "P177"),
    ("2.3.3.設備組不定期實體稽查及程式軟體稽查是否有使用未經授權之軟體及其他不當之軟體。",
     "2.3.3.電子計算機中心不定期實體稽查及程式軟體稽查是否有使用未經授權之軟體及其他不當之軟體。",
     "配合組織名稱調整，將「電子計算機中心」更正為「設備組」", "P179"),
    ("刪除",
     "．由不同單位人員參加成立緊急應變小組，並加強訓練其緊急事故應變能力。（系統復原計畫及測試作業 2.2）",
     "目前無此業務，刪除緊急應變小組相關作業程序條文", "P182"),
    ("．重大事故硬體或軟體復原，應由庶務組與電腦廠商簽訂重大意外事故系統復原合約。合約內容應包含修護完成交期、保固期間、違約損失賠償罰則及應變方式等條文。",
     "．重大事故硬體或軟體復原，應由設備組與電腦廠商簽訂重大意外事故系統復原合約。合約內容應包含修護完成交期、保固期間、違約損失賠償罰則及應變方式等條文。",
     "依採購及合約作業職責分工，簽訂廠商合約由庶務組辦理，非設備組職掌", "P182"),
    ("．設備組人員或維修外包廠商應將測試結果對使用方詳述說明。",
     "．設備組人員應將測試結果詳述說明。併同測試資料及程式規範送交教務主任核示後建檔。",
     "增列「維修外包廠商」為說明義務人，明確對象為「使用方」；刪除送交教務主任核示建檔程序", "P182"),
    ("刪除",
     "．是否規劃由不同單位人員參加成立緊急應變小組，並加強訓練其緊急應變能力。（系統復原計畫及測試作業 3.2）",
     "配合作業程序修訂，同步刪除緊急應變小組之對應控制重點", "P183"),
    ("刪除",
     "3.3.本校郵件伺服器是否設置防火牆及防毒軟體，以防止駭客或電腦病毒之侵害。(現在都是Google雲端檢查)",
     "本校郵件服務已改用Google雲端，由Google提供安全防護，無需自設防火牆，刪除此控制重點", "P184"),
    ("3.5.網管人員是否定期檢視郵件伺服器上郵件收發情形，若有異常狀況是否陳報權責主管處理。",
     "3.5.設備組人員是否定期檢視郵件伺服器上郵件收發情形，若有異常狀況是否陳報權責主管處理。",
     "明確指定由「網管人員」負責郵件伺服器監控，使職責劃分更清晰", "P185"),
]

# ── lxml 輔助函式 ──────────────────────────────────────────────────────────
def w(name): return qn(W, name)

def rpr_base(underline=False):
    rpr = etree.Element(w('rPr'))
    f = etree.SubElement(rpr, w('rFonts'))
    f.set(w('ascii'), '標楷體'); f.set(w('eastAsia'), '標楷體')
    f.set(w('hAnsi'), '標楷體'); f.set(w('hint'), 'eastAsia')
    etree.SubElement(rpr, w('sz')).set(w('val'), '28')
    etree.SubElement(rpr, w('szCs')).set(w('val'), '28')
    if underline:
        etree.SubElement(rpr, w('u')).set(w('val'), 'single')
    return rpr

def make_para_col1(text, underline=False):
    p = etree.Element(w('p'))
    ppr = etree.SubElement(p, w('pPr'))
    etree.SubElement(ppr, w('snapToGrid')).set(w('val'), '0')
    sp = etree.SubElement(ppr, w('spacing'))
    sp.set(w('line'), '240'); sp.set(w('lineRule'), 'atLeast')
    ppr.append(rpr_base(underline))
    r = etree.SubElement(p, w('r'))
    r.append(rpr_base(underline))
    t = etree.SubElement(r, w('t'))
    t.text = text
    return p

def make_para_col2_content(text):
    p = etree.Element(w('p'))
    ppr = etree.SubElement(p, w('pPr'))
    etree.SubElement(ppr, w('pStyle')).set(w('val'), '2')
    ind = etree.SubElement(ppr, w('ind'))
    ind.set(w('leftChars'), '200'); ind.set(w('left'), '440')
    ind.set(w('firstLineChars'), '4'); ind.set(w('firstLine'), '11')
    rpr = etree.SubElement(ppr, w('rPr'))
    f = etree.SubElement(rpr, w('rFonts'))
    f.set(w('ascii'), '標楷體'); f.set(w('eastAsia'), '標楷體')
    f.set(w('hAnsi'), '標楷體'); f.set(w('hint'), 'eastAsia')
    etree.SubElement(rpr, w('color')).set(w('val'), '000000')
    etree.SubElement(rpr, w('kern')).set(w('val'), '0')
    etree.SubElement(rpr, w('szCs')).set(w('val'), '28')
    # run 1: content
    r1 = etree.SubElement(p, w('r'))
    rpr1 = etree.SubElement(r1, w('rPr'))
    f1 = etree.SubElement(rpr1, w('rFonts'))
    f1.set(w('ascii'), '標楷體'); f1.set(w('eastAsia'), '標楷體')
    f1.set(w('hAnsi'), '標楷體'); f1.set(w('hint'), 'eastAsia')
    etree.SubElement(rpr1, w('color')).set(w('val'), '000000')
    etree.SubElement(rpr1, w('kern')).set(w('val'), '0')
    etree.SubElement(rpr1, w('szCs')).set(w('val'), '28')
    t1 = etree.SubElement(r1, w('t'))
    t1.text = text
    t1.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return p

def append_page_run(p, page):
    r2 = etree.SubElement(p, w('r'))
    rpr2 = etree.SubElement(r2, w('rPr'))
    f2 = etree.SubElement(rpr2, w('rFonts'))
    f2.set(w('ascii'), '標楷體'); f2.set(w('eastAsia'), '標楷體')
    f2.set(w('hAnsi'), '標楷體')
    etree.SubElement(rpr2, w('color')).set(w('val'), '000000')
    etree.SubElement(rpr2, w('kern')).set(w('val'), '0')
    etree.SubElement(rpr2, w('szCs')).set(w('val'), '28')
    etree.SubElement(r2, w('br'))
    t2 = etree.SubElement(r2, w('t'))
    t2.text = page

def make_para_col3(text):
    p = etree.Element(w('p'))
    ppr = etree.SubElement(p, w('pPr'))
    etree.SubElement(ppr, w('pStyle')).set(w('val'), '2')
    etree.SubElement(ppr, w('ind')).set(w('left'), '-108')
    etree.SubElement(ppr, w('jc')).set(w('val'), 'left')
    rpr = etree.SubElement(ppr, w('rPr'))
    f = etree.SubElement(rpr, w('rFonts'))
    f.set(w('ascii'), '標楷體'); f.set(w('eastAsia'), '標楷體')
    f.set(w('hAnsi'), '標楷體'); f.set(w('hint'), 'eastAsia')
    etree.SubElement(rpr, w('szCs')).set(w('val'), '28')
    r = etree.SubElement(p, w('r'))
    rpr2 = etree.SubElement(r, w('rPr'))
    f2 = etree.SubElement(rpr2, w('rFonts'))
    f2.set(w('ascii'), '標楷體'); f2.set(w('eastAsia'), '標楷體')
    f2.set(w('hAnsi'), '標楷體'); f2.set(w('hint'), 'eastAsia')
    etree.SubElement(rpr2, w('szCs')).set(w('val'), '28')
    t = etree.SubElement(r, w('t'))
    t.text = text
    return p

def make_tc(width, *paras):
    tc = etree.Element(w('tc'))
    tcPr = etree.SubElement(tc, w('tcPr'))
    tcW = etree.SubElement(tcPr, w('tcW'))
    tcW.set(w('w'), str(width)); tcW.set(w('type'), 'dxa')
    for p in paras:
        tc.append(p)
    return tc

def make_data_row(xiu, xian, shuo, page):
    tr = etree.Element(w('tr'))
    trPr = etree.SubElement(tr, w('trPr'))
    trH = etree.SubElement(trPr, w('trHeight'))
    trH.set(w('val'), '1011')

    # col1
    is_del = (xiu == '刪除')
    tc1 = make_tc(5983, make_para_col1(xiu, underline=is_del))

    # col2
    p2 = make_para_col2_content(xian)
    append_page_run(p2, page)
    tc2 = make_tc(5670, p2)

    # col3
    tc3 = make_tc(2551, make_para_col3(shuo))

    tr.append(tc1); tr.append(tc2); tr.append(tc3)
    return tr

# ── 讀取原始 docx，取得 tblPr 和 sectPr ─────────────────────────────────
with zipfile.ZipFile(SRC_DOCX) as z:
    with z.open("word/document.xml") as f:
        old_tree = etree.parse(f)

old_root = old_tree.getroot()
old_body = old_root.find(f'{{{W}}}body')
old_tbl  = old_body.find(f'{{{W}}}tbl')
old_tblPr = old_tbl.find(f'{{{W}}}tblPr') if old_tbl is not None else None
old_sectPr = old_body.find(f'{{{W}}}sectPr')

# 取得原始列 00-04（標題 + 4 預填列）
old_rows = old_tbl.findall(f'{{{W}}}tr') if old_tbl is not None else []
print(f"從原始檔讀取 {len(old_rows)} 列")

# ── 建構新的 document.xml ─────────────────────────────────────────────────
doc = etree.Element(w('document'), nsmap=NSMAP)
doc.set(f'{{{MC}}}Ignorable',
        'w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14')

body = etree.SubElement(doc, w('body'))

# 建表格
tbl = etree.SubElement(body, w('tbl'))
if old_tblPr is not None:
    tbl.append(deepcopy(old_tblPr))

# 加入原始列 00-04
for row in old_rows[:5]:   # 最多 5 列（00-04）
    tbl.append(deepcopy(row))

# 加入 11 筆新列
for c in changes:
    tbl.append(make_data_row(*c))

# 加入 sectPr
if old_sectPr is not None:
    body.append(deepcopy(old_sectPr))
else:
    # 預設 A4
    sp = etree.SubElement(body, w('sectPr'))
    pgSz = etree.SubElement(sp, w('pgSz'))
    pgSz.set(w('w'), '11906'); pgSz.set(w('h'), '16838')

print(f"新文件共 {len(tbl.findall(w('tr')))} 列")

# ── 序列化 ───────────────────────────────────────────────────────────────
new_xml = etree.tostring(doc, xml_declaration=True, encoding='UTF-8',
                         standalone=True, pretty_print=False)

# ── 重新打包 ─────────────────────────────────────────────────────────────
with zipfile.ZipFile(SRC_DOCX) as zin, zipfile.ZipFile(TMP_DOCX, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        if item.filename == "word/document.xml":
            zout.writestr(item, new_xml)
        else:
            zout.writestr(item, zin.read(item.filename))

shutil.copy2(TMP_DOCX, DST_DOCX)
print(f"✓ 已輸出：{DST_DOCX}")
