"""用 Word COM 正確取得各段落的頁碼（footer 顯示的頁碼）"""
import sys, win32com.client, os
sys.stdout.reconfigure(encoding='utf-8')

doc_path = os.path.abspath(r"D:\D\114設備組\設備組內控修訂.docx")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(doc_path)
doc.Repaginate()

wdActiveEndPageNumber = 3
wdNumberOfPagesInDocument = 4

total = doc.ComputeStatistics(2)  # wdStatisticPages
print(f"文件總頁數: {total}")
print(f"pgNumType start: 從文件設定看")
print()

# 搜尋每個關鍵字並取得頁碼
searches = [
    ("◎檔案及設備之安全作業",         "◎檔案及設備之安全作業"),
    ("◎硬體及系統軟體之使用及維護作業", "◎硬體及系統軟體之使用及維護作業"),
    ("◎系統復原計畫及測試作業",         "◎系統復原計畫及測試作業"),
    ("◎資訊安全之檢查作業",             "◎資訊安全之檢查作業"),
    ("3.4.異地備援",                     "3.4.異地備援"),
    ("2.3.3.",                           "2.3.3."),
    ("由不同單位人員參加成立緊急應變小組", "2.2 緊急應變"),
    ("重大事故硬體或軟體復原，應由",     "2.2 設備→庶務"),
    ("設備組人員應將測試結果",           "2.3 復原測試"),
    ("是否規劃由不同單位人員",           "3.2 緊急應變控制"),
    ("本校郵件伺服器是否設置防火牆",     "3.3 郵件防火牆"),
    ("3.5.設備組人員是否定期檢視郵件",   "3.5 設備組→網管"),
]

for text, label in searches:
    rng = doc.Range(0, doc.Content.End)
    rng.Find.ClearFormatting()
    rng.Find.Text = text
    rng.Find.Forward = True
    rng.Find.Wrap = 1  # wdFindStop
    found = rng.Find.Execute()
    if found:
        page_logical = rng.Information(wdActiveEndPageNumber)
        print(f"[{label}] 第{page_logical}頁")
    else:
        print(f"[{label}] 找不到")

doc.Close(False)
word.Quit()
