"""
從零建構乾淨的 DOCX（不依賴任何損壞的舊檔）
"""
import sys, zipfile, shutil
from lxml import etree
sys.stdout.reconfigure(encoding='utf-8')

TMP  = r"C:\Users\user\allen\_clean.docx"
DST  = r"D:\D\114設備組\修正對照表-內控-設備組.docx"

# ── 命名空間 ──
W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
W14= "http://schemas.microsoft.com/office/word/2010/wordml"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
PKG_CT  = "http://schemas.openxmlformats.org/package/2006/content-types"
WORD_DOC= "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
WORD_STY= "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
WORD_SET= "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"

def w(n): return f"{{{W}}}{n}"
def r(n): return f"{{{R}}}{n}"

# ── 表格資料 ─────────────────────────────────────────────────────────────
ORIG_ROWS = [
    # (修正規定,  現行規定,  頁碼,  說明)
    ("刪除", "系統開發及程式修改作業",  "P166~P169", "本校無該項業務"),
    ("刪除", "系統文書編製作業",         "",          "本校無該項業務"),
    ("刪除", "程式及資料之存取作業",     "P170~P172", "本校無該項業務"),
    ("刪除", "資料輸出入及處理作業",     "P173~P175", "本校無該項業務"),
]
NEW_ROWS = [
    ("3.4.異地備援：",
     "3.4.異地備援：(因經費與管理的緣故尚未施行)",
     "P177", "異地備援措施已逐步實施，刪除「尚未施行」括號說明文字"),
    ("3.4.1.檔案備份資料存放於異地或雲端，是否具備保全性及完整記載資料存放紀錄。",
     "3.4.1.檔案備份資料存放於異地，是否具備保全性及完整記載資料存放紀錄。",
     "P177", "因應雲端備份趨勢，增列「雲端」作為備援存放選項"),
    ("3.4.2.每學期進行一次異地備份還原演練，確保回復程序能在指定時間內完成復原作業程序。",
     "3.4.2.於異地設置實體主機，是否定期檢查測試回復程序，確保回復程序能在指定時間內完成復原作業程序。",
     "P177", "明確規定異地備援測試頻率，以每學期辦理一次演練取代模糊之「定期」規定"),
    ("5.1.正心中學電腦主機房管理辦法。",
     "5.1.正心中學電子計算機中心電腦主機房管理辦法。",
     "P177", "配合組織名稱調整，刪除「電子計算機中心」，精簡文件名稱"),
    ("2.3.3.設備組不定期實體稽查及程式軟體稽查是否有使用未經授權之軟體及其他不當之軟體。",
     "2.3.3.電子計算機中心不定期實體稽查及程式軟體稽查是否有使用未經授權之軟體及其他不當之軟體。",
     "P179", "配合組織名稱調整，將「電子計算機中心」更正為「設備組」"),
    ("刪除",
     "．由不同單位人員參加成立緊急應變小組，並加強訓練其緊急事故應變能力。（系統復原計畫及測試作業 2.2）",
     "P182", "目前無此業務，刪除緊急應變小組相關作業程序條文"),
    ("．重大事故硬體或軟體復原，應由庶務組與電腦廠商簽訂重大意外事故系統復原合約。合約內容應包含修護完成交期、保固期間、違約損失賠償罰則及應變方式等條文。",
     "．重大事故硬體或軟體復原，應由設備組與電腦廠商簽訂重大意外事故系統復原合約。合約內容應包含修護完成交期、保固期間、違約損失賠償罰則及應變方式等條文。",
     "P182", "依採購及合約作業職責分工，簽訂廠商合約由庶務組辦理，非設備組職掌"),
    ("．設備組人員或維修外包廠商應將測試結果對使用方詳述說明。",
     "．設備組人員應將測試結果詳述說明。併同測試資料及程式規範送交教務主任核示後建檔。",
     "P182", "增列「維修外包廠商」為說明義務人，明確對象為「使用方」；刪除送交教務主任核示建檔程序"),
    ("刪除",
     "．是否規劃由不同單位人員參加成立緊急應變小組，並加強訓練其緊急應變能力。（系統復原計畫及測試作業 3.2）",
     "P183", "配合作業程序修訂，同步刪除緊急應變小組之對應控制重點"),
    ("刪除",
     "3.3.本校郵件伺服器是否設置防火牆及防毒軟體，以防止駭客或電腦病毒之侵害。(現在都是Google雲端檢查)",
     "P184", "本校郵件服務已改用Google雲端，由Google提供安全防護，無需自設防火牆，刪除此控制重點"),
    ("3.5.網管人員是否定期檢視郵件伺服器上郵件收發情形，若有異常狀況是否陳報權責主管處理。",
     "3.5.設備組人員是否定期檢視郵件伺服器上郵件收發情形，若有異常狀況是否陳報權責主管處理。",
     "P185", "明確指定由「網管人員」負責郵件伺服器監控，使職責劃分更清晰"),
]

# ── 建 XML 元素輔助 ───────────────────────────────────────────────────────
KAI = "標楷體"

def font_rpr(sz="28", underline=False, color=None):
    rpr = etree.Element(w('rPr'))
    f = etree.SubElement(rpr, w('rFonts'))
    f.set(w('ascii'), KAI); f.set(w('eastAsia'), KAI); f.set(w('hAnsi'), KAI)
    etree.SubElement(rpr, w('sz')).set(w('val'), sz)
    etree.SubElement(rpr, w('szCs')).set(w('val'), sz)
    if underline:
        etree.SubElement(rpr, w('u')).set(w('val'), 'single')
    if color:
        etree.SubElement(rpr, w('color')).set(w('val'), color)
    return rpr

def make_run(text, sz="28", underline=False, br_before=False):
    r_ = etree.Element(w('r'))
    r_.append(font_rpr(sz, underline))
    if br_before:
        etree.SubElement(r_, w('br'))
    t = etree.SubElement(r_, w('t'))
    t.text = text
    if text and (text[0] == ' ' or text[-1] == ' '):
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return r_

def simple_para(text, jc=None, sz="28", underline=False, indent=None):
    p = etree.Element(w('p'))
    ppr = etree.SubElement(p, w('pPr'))
    if jc:
        etree.SubElement(ppr, w('jc')).set(w('val'), jc)
    if indent:
        ind = etree.SubElement(ppr, w('ind'))
        for k, v in indent.items():
            ind.set(w(k), v)
    sp = etree.SubElement(ppr, w('spacing'))
    sp.set(w('line'), '240'); sp.set(w('lineRule'), 'atLeast')
    ppr.append(font_rpr(sz, underline))
    p.append(make_run(text, sz, underline))
    return p

def make_tc(width, *paras):
    tc = etree.Element(w('tc'))
    tcPr = etree.SubElement(tc, w('tcPr'))
    tcW = etree.SubElement(tcPr, w('tcW'))
    tcW.set(w('w'), str(width)); tcW.set(w('type'), 'dxa')
    for p in paras:
        tc.append(p)
    return tc

def make_tr(height=None, header=False):
    tr = etree.Element(w('tr'))
    trPr = etree.SubElement(tr, w('trPr'))
    if height:
        etree.SubElement(trPr, w('trHeight')).set(w('val'), str(height))
    if header:
        etree.SubElement(trPr, w('tblHeader'))
    return tr

# ── 標題列 ───────────────────────────────────────────────────────────────
def make_header_row():
    tr = make_tr(header=True)
    titles = ["修　正　規　定", "現　行　規　定", "說　　　　明"]
    widths = [5983, 5670, 2551]
    for title, width in zip(titles, widths):
        tc = etree.Element(w('tc'))
        tcPr = etree.SubElement(tc, w('tcPr'))
        tcW = etree.SubElement(tcPr, w('tcW'))
        tcW.set(w('w'), str(width)); tcW.set(w('type'), 'dxa')
        etree.SubElement(tcPr, w('vAlign')).set(w('val'), 'center')
        p = etree.Element(w('p'))
        ppr = etree.SubElement(p, w('pPr'))
        etree.SubElement(ppr, w('jc')).set(w('val'), 'center')
        sp = etree.SubElement(ppr, w('spacing'))
        sp.set(w('line'), '240'); sp.set(w('lineRule'), 'atLeast')
        ppr.append(font_rpr("28"))
        p.append(make_run(title, "28"))
        tc.append(p)
        tr.append(tc)
    return tr

# ── 資料列（原始 + 新增）─────────────────────────────────────────────────
def make_data_row(xiu, xian, page, shuo):
    tr = make_tr(height=1011)
    is_del = (xiu == "刪除")

    # col1 修正規定
    p1 = etree.Element(w('p'))
    ppr1 = etree.SubElement(p1, w('pPr'))
    sp1 = etree.SubElement(ppr1, w('spacing'))
    sp1.set(w('line'), '240'); sp1.set(w('lineRule'), 'atLeast')
    ppr1.append(font_rpr("28", underline=is_del))
    p1.append(make_run(xiu, "28", underline=is_del))
    tr.append(make_tc(5983, p1))

    # col2 現行規定（+ 換行 + 頁碼）
    p2 = etree.Element(w('p'))
    ppr2 = etree.SubElement(p2, w('pPr'))
    ind2 = etree.SubElement(ppr2, w('ind'))
    ind2.set(w('leftChars'), '200'); ind2.set(w('left'), '440')
    ind2.set(w('firstLineChars'), '4'); ind2.set(w('firstLine'), '11')
    sp2 = etree.SubElement(ppr2, w('spacing'))
    sp2.set(w('line'), '240'); sp2.set(w('lineRule'), 'atLeast')
    ppr2.append(font_rpr("28"))
    p2.append(make_run(xian, "28"))
    if page:
        p2.append(make_run(page, "28", br_before=True))
    tr.append(make_tc(5670, p2))

    # col3 說明
    p3 = etree.Element(w('p'))
    ppr3 = etree.SubElement(p3, w('pPr'))
    sp3 = etree.SubElement(ppr3, w('spacing'))
    sp3.set(w('line'), '240'); sp3.set(w('lineRule'), 'atLeast')
    ppr3.append(font_rpr("28"))
    p3.append(make_run(shuo, "28"))
    tr.append(make_tc(2551, p3))

    return tr

# ── 建構 document.xml ─────────────────────────────────────────────────────
NSMAP_DOC = {'w': W, 'mc': MC, 'w14': W14,
             'r': R}

doc = etree.Element(w('document'), nsmap=NSMAP_DOC)
doc.set(f'{{{MC}}}Ignorable', 'w14')
body = etree.SubElement(doc, w('body'))

# 表格
tbl = etree.SubElement(body, w('tbl'))

# 表格屬性
tblPr = etree.SubElement(tbl, w('tblPr'))
tblW = etree.SubElement(tblPr, w('tblW'))
tblW.set(w('w'), '14204'); tblW.set(w('type'), 'dxa')
tblInd = etree.SubElement(tblPr, w('tblInd'))
tblInd.set(w('w'), '108'); tblInd.set(w('type'), 'dxa')
tblBorders = etree.SubElement(tblPr, w('tblBorders'))
for side in ('top','left','bottom','right','insideH','insideV'):
    b = etree.SubElement(tblBorders, w(side))
    b.set(w('val'),'single'); b.set(w('sz'),'4')
    b.set(w('space'),'0'); b.set(w('color'),'auto')

# 欄寬定義
tblGrid = etree.SubElement(tbl, w('tblGrid'))
for cw in (5983, 5670, 2551):
    etree.SubElement(tblGrid, w('gridCol')).set(w('w'), str(cw))

# 標題列
tbl.append(make_header_row())

# 原始 4 列
for row in ORIG_ROWS:
    tbl.append(make_data_row(*row))

# 新增 11 列
for row in NEW_ROWS:
    tbl.append(make_data_row(*row))

# 段落結尾 + sectPr（A4 直向）
etree.SubElement(body, w('p'))
sectPr = etree.SubElement(body, w('sectPr'))
pgSz = etree.SubElement(sectPr, w('pgSz'))
pgSz.set(w('w'), '16838'); pgSz.set(w('h'), '11906')
pgSz.set(w('orient'), 'landscape')
pgMar = etree.SubElement(sectPr, w('pgMar'))
pgMar.set(w('top'),'1134'); pgMar.set(w('right'),'1134')
pgMar.set(w('bottom'),'1134'); pgMar.set(w('left'),'1134')
pgMar.set(w('header'),'851'); pgMar.set(w('footer'),'992')
pgMar.set(w('gutter'),'0')

doc_xml = etree.tostring(doc, xml_declaration=True, encoding='UTF-8',
                         standalone=True, pretty_print=False)

# ── 各支援 XML ────────────────────────────────────────────────────────────
content_types = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

rels = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""

doc_rels = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

# ── 打包 ──────────────────────────────────────────────────────────────────
with zipfile.ZipFile(TMP, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels",         rels)
    z.writestr("word/document.xml",   doc_xml)
    z.writestr("word/_rels/document.xml.rels", doc_rels)

shutil.copy2(TMP, DST)
print(f"✓ 共 {len(tbl.findall(w('tr')))} 列")
print(f"✓ 已輸出：{DST}")
