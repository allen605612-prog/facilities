import win32com.client, os, sys

DOC = r"D:\D\教學資料\物理科\國中物理\講義\講義15_時間與運動合冊.doc"
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
sys.stdout.reconfigure(encoding="utf-8")
try:
    doc = word.Documents.Open(os.path.abspath(DOC))
    print(f"總段落數：{doc.Paragraphs.Count}")
    for i, para in enumerate(doc.Paragraphs):
        t = para.Range.Text
        ts = t.strip()
        if ts:
            print(f"[{i:03d}] {repr(ts[:50])}")
            print(f"       hex={ts[:10].encode('utf-16-le').hex()}")
        if i > 10:
            print("...")
            break
    # 也找含 "15" 的段落
    print("\n--- 含「15」的段落 ---")
    for i, para in enumerate(doc.Paragraphs):
        t = para.Range.Text
        if "15" in t:
            print(f"[{i:03d}] {repr(t.strip()[:60])}")
    doc.Close(SaveChanges=False)
finally:
    word.Quit()
