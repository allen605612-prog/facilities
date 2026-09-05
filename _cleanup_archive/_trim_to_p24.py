"""
刪除第 25 頁的段落（習題 7 全部），讓自由落體章節收在第 24 頁。
刪除段落索引 415～418（從後往前刪，避免索引位移）。
"""
import win32com.client, os, sys

DOC = r"D:\D\教學資料\物理科\國中物理\講義\講義15_時間與運動合冊.doc"
wdFormatDocument = 0

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
sys.stdout.reconfigure(encoding="utf-8")

try:
    doc = word.Documents.Open(os.path.abspath(DOC))

    # 從後往前刪，索引不會位移
    paras = list(doc.Paragraphs)
    for idx in [418, 417, 416, 415]:
        t = paras[idx].Range.Text.strip()
        paras[idx].Range.Delete()
        print(f"  已刪段落 [{idx}]：{t[:60]}")

    # 確認新頁數
    doc.Repaginate()
    pages = doc.ComputeStatistics(2)
    print(f"\n現在共 {pages} 頁")

    doc.SaveAs2(os.path.abspath(DOC), FileFormat=wdFormatDocument)
    doc.Close(SaveChanges=False)
    print(f"完成！→ {DOC}")
finally:
    word.Quit()
