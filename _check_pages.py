import win32com.client, os, sys

DOC = r"D:\D\教學資料\物理科\國中物理\講義\講義15_時間與運動合冊.doc"
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
sys.stdout.reconfigure(encoding="utf-8")

try:
    doc = word.Documents.Open(os.path.abspath(DOC))
    doc.Repaginate()

    total_pages = doc.ComputeStatistics(2)  # wdStatisticPages = 2
    print(f"總頁數：{total_pages}")

    # 找出每段所在頁碼
    paras = list(doc.Paragraphs)
    total = len(paras)
    print(f"總段落數：{total}\n")

    # 印出第 20 頁以後的段落（頁碼 + 內容）
    print("=== 第 20 頁之後各段落 ===")
    for i, para in enumerate(paras):
        t = para.Range.Text.strip()
        if not t:
            continue
        page = para.Range.Information(3)  # wdActiveEndPageNumber = 3
        if page >= 20:
            print(f"[p{page:02d}][{i:03d}] {t[:70]}")

    doc.Close(SaveChanges=False)
finally:
    word.Quit()
