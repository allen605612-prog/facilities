import win32com.client, os, sys, re

DOC = r"D:\D\教學資料\物理科\國中物理\講義\講義15_時間與運動合冊.doc"
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
sys.stdout.reconfigure(encoding="utf-8")
pat15x = re.compile(r'15-[1-9]')

try:
    doc = word.Documents.Open(os.path.abspath(DOC))

    # 改用 doc.Range 直接定位刪除，完全繞開 Find/Replace 問題
    changed = 0
    for para in doc.Paragraphs:
        t = para.Range.Text
        if not pat15x.search(t):
            continue
        orig = t.strip()
        para_start = para.Range.Start

        # 逐一找到每個 15-X 並刪除（從後往前，避免位移問題）
        matches = list(re.finditer(r'15-[1-9] ?', t))
        for m in reversed(matches):
            abs_start = para_start + m.start()
            abs_end   = para_start + m.end()
            sub = doc.Range(abs_start, abs_end)
            print(f"  刪除位置 [{abs_start},{abs_end}] 文字={repr(sub.Text)}")
            sub.Delete()
            changed += 1

        after = para.Range.Text.strip()
        print(f"  段落：「{orig}」→「{after}」")

    print(f"\n共刪除 {changed} 個 15-X")

    doc.Close(SaveChanges=False)
finally:
    word.Quit()
