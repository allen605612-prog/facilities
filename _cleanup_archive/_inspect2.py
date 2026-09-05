import win32com.client, os

DOC = r"D:\D\教學資料\物理科\國中物理\講義\講義15_時間與運動合冊.doc"
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(os.path.abspath(DOC))
    hits = []
    for i, para in enumerate(doc.Paragraphs):
        t = para.Range.Text
        if "15" in t and len(t.strip()) < 30:
            hits.append((i, t.strip()))
    doc.Close(SaveChanges=False)
finally:
    word.Quit()

import sys
sys.stdout.reconfigure(encoding="utf-8")
for i, t in hits:
    print(f"[{i:03d}] {repr(t)}")
    print(f"       hex: {t.encode('utf-16-le').hex()}")
