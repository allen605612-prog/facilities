"""
在合冊末尾新增「§自由落體」章節，格式與加速度章一致。
"""
import win32com.client, os, sys

DOC = r"D:\D\教學資料\物理科\國中物理\講義\講義15_時間與運動合冊.doc"

wdAlignParagraphCenter = 1
wdAlignParagraphLeft   = 0
wdPageBreak            = 7
wdStory                = 6
wdFormatDocument       = 0

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
sys.stdout.reconfigure(encoding="utf-8")

# ── 自由落體章節內容 ──────────────────────────────────────────────────
# (text, font_size, align, bold)
CONTENT = [
    # ── 章節標題 ──
    ("§自由落體",                                                          18, wdAlignParagraphCenter, False),

    # ── a. 意義 ──
    ("a. 自由落體的意義",                                                   14, wdAlignParagraphLeft,   False),
    ("物體只受重力作用，由靜止開始落下的運動，稱為自由落體。",                  14, wdAlignParagraphLeft,   False),
    ("自由落體為初速為零的（          ）加速度直線運動。",                      14, wdAlignParagraphLeft,   False),
    ("地表附近，重力加速度 g ≒ 9.8 m/s²，計算時常取 g = 10 m/s²，方向垂直向下。",
                                                                           14, wdAlignParagraphLeft,   False),
    ("※自由落體的條件：初速度為零、只受重力、忽略空氣阻力。",                   14, wdAlignParagraphLeft,   False),

    # ── b. 公式 ──
    ("b. 自由落體公式",                                                      14, wdAlignParagraphLeft,   False),
    ("（令 Vo = 0，g = 10 m/s²，S：位移，t：時間，V：末速度）",               14, wdAlignParagraphLeft,   False),
    ("①  速度公式：V = gt",                                                  14, wdAlignParagraphLeft,   False),
    ("②  位移公式：S = ½ g t²",                                              14, wdAlignParagraphLeft,   False),
    ("③  速度位移：V² = 2gS",                                                14, wdAlignParagraphLeft,   False),
    ("④  第 t 秒的位移：St = g(2t－1)/2",                                    14, wdAlignParagraphLeft,   False),

    # ── c. 各秒位移特性 ──
    ("c. 各秒位移特性",                                                       14, wdAlignParagraphLeft,   False),
    ("※ 各秒位移比（第1秒：第2秒：第3秒…） = 1 : 3 : 5 : …（奇數比）",        14, wdAlignParagraphLeft,   False),
    ("※ 前 N 秒總位移比（前1秒：前2秒：前3秒…） = 1 : 4 : 9 : …（完全平方比）",
                                                                           14, wdAlignParagraphLeft,   False),

    # ── 例題 ──
    ("例題1.（　）自由落下的物體，重力加速度取 g = 10 m/s²，落下3秒後的速度為多少？"
     "ˉ(A) 10 m/sˉ(B) 20 m/sˉ(C) 30 m/sˉ(D) 45 m/s。",                   14, wdAlignParagraphLeft,   False),

    ("例題2.（　）一物體由靜止自由落下，g = 10 m/s²，落下5秒所通過的距離為？"
     "ˉ(A) 25 mˉ(B) 50 mˉ(C) 100 mˉ(D) 125 m。",                         14, wdAlignParagraphLeft,   False),

    ("例題3.（　）自由落體運動中，下列敘述何者正確？"
     "ˉ(A) 速度大小固定不變ˉ(B) 加速度方向向上ˉ(C) 加速度大小保持固定ˉ(D) 每秒位移相同。",
                                                                           14, wdAlignParagraphLeft,   False),

    ("例題4. 一物體自靜止開始自由落下，g = 10 m/s²，試求：",                  14, wdAlignParagraphLeft,   False),
    ("(落下第 3 秒末的速度為ˉˉˉˉ m/s。",                                    14, wdAlignParagraphLeft,   False),
    ("(前 3 秒的總位移為ˉˉˉˉ m。",                                          14, wdAlignParagraphLeft,   False),
    ("(第 3 秒內的位移為ˉˉˉˉ m。",                                          14, wdAlignParagraphLeft,   False),
    ("(落地速度為 30 m/s，則落下高度為ˉˉˉˉ m。",                             14, wdAlignParagraphLeft,   False),

    # ── 隨堂演練 ──
    ("§隨堂演練",                                                            16, wdAlignParagraphLeft,   False),

    ("1.（　）以下何者不屬於自由落體運動？"
     "ˉ(A) 真空中由靜止落下的鐵球ˉ(B) 真空中由靜止落下的紙片"
     "ˉ(C) 空氣中由靜止落下的羽毛ˉ(D) 空氣中由靜止落下的石頭。",              14, wdAlignParagraphLeft,   False),

    ("2.（　）一物體自由落下，g = 10 m/s²，落下 4 秒後的速度大小為？"
     "ˉ(A) 10 m/sˉ(B) 20 m/sˉ(C) 40 m/sˉ(D) 80 m/s。",                   14, wdAlignParagraphLeft,   False),

    ("3.（　）自由落體第 1、2、3 … 秒末的速度比為？"
     "ˉ(A) 1：1：1ˉ(B) 1：2：3ˉ(C) 1：4：9ˉ(D) 1：3：5。",               14, wdAlignParagraphLeft,   False),

    ("4.（　）自由落體第 1、2、3 秒內的位移比為？"
     "ˉ(A) 1：2：3ˉ(B) 1：3：5ˉ(C) 1：4：9ˉ(D) 1：2：4。",               14, wdAlignParagraphLeft,   False),

    ("5.（　）前 1 秒與前 2 秒的位移比為？"
     "ˉ(A) 1：2ˉ(B) 1：3ˉ(C) 1：4ˉ(D) 1：5。",                          14, wdAlignParagraphLeft,   False),

    ("6.（　）一物體自 5 m 高處由靜止自由落下，g = 10 m/s²，落地時的速度為？"
     "ˉ(A) 5 m/sˉ(B) 10 m/sˉ(C) 50 m/sˉ(D) 100 m/s。",                  14, wdAlignParagraphLeft,   False),

    ("7. 一物體自屋頂由靜止自由落下，g = 10 m/s²，屋頂距地面 80 m，試求：",   14, wdAlignParagraphLeft,   False),
    ("(落地所需時間為ˉˉˉˉ秒。",                                              14, wdAlignParagraphLeft,   False),
    ("(落地時的速度為ˉˉˉˉ m/s。",                                           14, wdAlignParagraphLeft,   False),
    ("(第 3 秒末的速度為ˉˉˉˉ m/s，位置距地面ˉˉˉˉ m。",                      14, wdAlignParagraphLeft,   False),
]

try:
    doc = word.Documents.Open(os.path.abspath(DOC))

    # 移到文件最末
    sel = word.Selection
    sel.EndKey(Unit=wdStory)

    # 插入分頁
    sel.InsertBreak(Type=wdPageBreak)

    # 逐段新增
    for text, sz, align, bold in CONTENT:
        sel.EndKey(Unit=wdStory)
        para = sel.Paragraphs(1)
        para.Style = doc.Styles("內文")
        rng = sel.Range
        rng.Font.Size = sz
        rng.Font.Bold = bold
        rng.ParagraphFormat.Alignment = align
        sel.TypeText(text)
        sel.TypeParagraph()
        print(f"  +{sz}pt {text[:40]}")

    doc.SaveAs2(os.path.abspath(DOC), FileFormat=wdFormatDocument)
    doc.Close(SaveChanges=False)
    print(f"\n完成！→ {DOC}")
finally:
    word.Quit()
