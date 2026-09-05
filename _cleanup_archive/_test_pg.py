"""確認 Word COM 頁碼是否為邏輯頁碼（含 pgNumType offset）"""
import sys, win32com.client, os
sys.stdout.reconfigure(encoding='utf-8')

doc_path = os.path.abspath(r"D:\D\114設備組\設備組內控修訂.docx")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(doc_path)
doc.Repaginate()

# 第1段（第一頁）的頁碼
p1 = doc.Paragraphs(1).Range
pg1 = p1.Information(3)  # wdActiveEndPageNumber
print(f"第1段（第一頁）頁碼 = {pg1}")

# 最後一段
last = doc.Paragraphs(doc.Paragraphs.Count).Range
pglast = last.Information(3)
print(f"最後一段頁碼 = {pglast}")
print(f"文件共 {doc.Paragraphs.Count} 段, {doc.ComputeStatistics(2)} 頁")

doc.Close(False)
word.Quit()
