"""
用 Word COM 取得各關鍵段落的精確頁碼
"""
import sys, win32com.client, os
sys.stdout.reconfigure(encoding='utf-8')

doc_path = r"D:\D\114設備組\設備組內控修訂.docx"

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(os.path.abspath(doc_path))

# 重新計算頁碼
doc.Repaginate()

# 關鍵搜尋字串
keywords = [
    ("◎檔案及設備之安全作業", "P202"),
    ("◎硬體及系統軟體之使用及維護作業", "P256"),
    ("◎系統復原計畫及測試作業", "P313"),
    ("◎資訊安全之檢查作業", "P361"),
    ("3.4.異地備援", "P247"),
    ("2.3.3.", "P279"),
    ("由不同單位人員參加成立緊急應變小組", "P328"),
    ("重大事故硬體或軟體復原", "P334"),
    ("設備組人員應將測試結果", "P339"),
    ("是否規劃由不同單位人員參加成立緊急應變小組", "P345"),
    ("本校郵件伺服器是否設置防火牆", "P375"),
    ("3.5.", "P377"),
]

find = doc.Content.Find
results = {}
for keyword, label in keywords:
    find.ClearFormatting()
    find.Text = keyword
    find.Execute()
    if find.Found:
        rng = doc.Content.Find
        # Use Range to get page
        rng2 = doc.Range(0, doc.Content.End)
        rng2.Find.ClearFormatting()
        rng2.Find.Text = keyword
        rng2.Find.Execute()
        if rng2.Find.Found:
            page = rng2.Information(3)  # wdActiveEndPageNumber = 3
            results[label] = (keyword[:20], page)
            print(f"{label}: '{keyword[:25]}' → 第{page}頁")

doc.Close(False)
word.Quit()
