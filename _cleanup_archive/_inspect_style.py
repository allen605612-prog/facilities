import win32com.client, os, sys

DOC = r"D:\D\教學資料\物理科\國中物理\講義\講義15_時間與運動合冊.doc"
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
sys.stdout.reconfigure(encoding="utf-8")

try:
    doc = word.Documents.Open(os.path.abspath(DOC))
    total = doc.Paragraphs.Count
    print(f"總段落數：{total}\n")

    # 印最後 60 段（加速度章末尾）
    print("=== 最後 60 段 ===")
    paras = list(doc.Paragraphs)
    for para in paras[max(0, total-60):]:
        t = para.Range.Text.strip()
        if not t:
            continue
        style = para.Style.NameLocal
        sz    = para.Range.Font.Size
        bold  = para.Range.Font.Bold
        print(f"style={style!r:12s} sz={sz:5} bold={bold} | {t[:70]}")

    # 印前 20 段（時間章開頭）觀察格式
    print("\n=== 加速度章開頭（239開始）===")
    for para in paras[239:260]:
        t = para.Range.Text.strip()
        if not t:
            continue
        style = para.Style.NameLocal
        sz    = para.Range.Font.Size
        bold  = para.Range.Font.Bold
        align = para.Alignment
        indent_left  = para.LeftIndent
        indent_first = para.FirstLineIndent
        print(f"style={style!r:12s} sz={sz:5} bold={bold} align={align} | {t[:60]}")

    doc.Close(SaveChanges=False)
finally:
    word.Quit()
