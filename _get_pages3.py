"""補充查詢特定段落頁碼（物理頁+165=頁碼）"""
import sys, win32com.client, os
sys.stdout.reconfigure(encoding='utf-8')

doc_path = os.path.abspath(r"D:\D\114設備組\設備組內控修訂.docx")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(doc_path)
doc.Repaginate()

OFFSET = 165  # physical + 165 = logical page

searches = [
    ("2.3.3.電子計算機中心",  "2.3.3 電子計算機中心→設備組"),
    ("設備組人員應將測試結果",  "P339 復原測試說明"),
    ("設備組人員或維修外包廠商", "P339 修正後"),
    ("3.5.",                  "3.5（含設備組or網管）"),
    ("網管人員是否定期",        "3.5 網管"),
    ("定期檢視郵件伺服器",      "3.5 郵件"),
    ("3.4.1.",                "3.4.1"),
    ("3.4.2.",                "3.4.2"),
    ("5.1.",                  "5.1 依據文件"),
]

for text, label in searches:
    rng = doc.Range(0, doc.Content.End)
    rng.Find.ClearFormatting()
    rng.Find.Text = text
    rng.Find.Forward = True
    rng.Find.Wrap = 1
    found = rng.Find.Execute()
    if found:
        phys = rng.Information(3)
        logical = phys + OFFSET
        # 取得段落文字
        para_text = rng.Paragraphs(1).Range.Text.strip()[:50]
        print(f"[{label}] 物理第{phys}頁 → 第{logical}頁 | {para_text}")
    else:
        print(f"[{label}] 找不到")

doc.Close(False)
word.Quit()
