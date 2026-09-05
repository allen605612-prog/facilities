import win32com.client, os

DOC = r"D:\D\教學資料\物理科\國中物理\講義\講義15_時間與運動合冊.doc"
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(os.path.abspath(DOC))
    for i, para in enumerate(doc.Paragraphs):
        t = para.Range.Text.strip()
        if t and ("15" in t or i < 5):
            style = para.Style.NameLocal
            lvl   = para.OutlineLevel
            print(f"[{i:03d}] lvl={lvl} style={style!r:20s} | {t[:60]}")
        if i > 80:
            print("...")
            break
    doc.Close(SaveChanges=False)
finally:
    word.Quit()
