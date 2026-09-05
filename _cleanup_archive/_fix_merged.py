"""
修正合併讀義：移除標題 15-X + 加置中頁碼
"""
import win32com.client, os, sys, re

DOC = r"D:\D\教學資料\物理科\國中物理\講義\講義15_時間與運動合冊.doc"

wdAlignParagraphCenter = 1
wdHeaderFooterPrimary  = 1
wdFieldPage            = 33
wdFormatDocument       = 0

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
sys.stdout.reconfigure(encoding="utf-8")

# 含「15-數字」的章節標題模式（不含後方空格，以免多吃）
pat15x = re.compile(r'15-[1-9]')

try:
    doc = word.Documents.Open(os.path.abspath(DOC))
    print(f"已開啟，共 {doc.Paragraphs.Count} 段")

    # 先找出第四章標題（供確認）
    print("\n--- 含「加速」或「15-4」的段落 ---")
    for i, para in enumerate(doc.Paragraphs):
        t = para.Range.Text
        if "加速" in t or "15-4" in t:
            print(f"  [{i:03d}] {repr(t.strip()[:60])}")

    # ── 1. 移除 15-X（Word Range 位置 = para.Range.Start + 1 + str_index）
    changed = 0
    for para in doc.Paragraphs:
        t = para.Range.Text
        if not pat15x.search(t):
            continue
        orig = t.strip()
        para_start = para.Range.Start

        # 從後往前刪，避免位移問題
        matches = list(re.finditer(r'15-[1-9] ?', t))
        for m in reversed(matches):
            # +1 是因為 Word Range.Start 是字元「前」的位置
            abs_start = para_start + m.start() + 1
            abs_end   = para_start + m.end()   + 1
            sub = doc.Range(abs_start, abs_end)
            sub.Delete()

        after = para.Range.Text.strip()
        if after != orig:
            print(f"  「{orig}」→「{after}」")
            changed += 1

    print(f"\n共修改 {changed} 個標題")

    # ── 2. 置中頁碼 ───────────────────────────────────────────────────
    for i, section in enumerate(doc.Sections, 1):
        footer = section.Footers(wdHeaderFooterPrimary)
        footer.LinkToPrevious = False
        footer.Range.Delete()
        footer.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        footer.Range.Fields.Add(Range=footer.Range, Type=wdFieldPage)
        print(f"Section {i}：頁碼設定完成")

    doc.SaveAs2(os.path.abspath(DOC), FileFormat=wdFormatDocument)
    doc.Close(SaveChanges=False)
    print(f"\n完成！→ {DOC}")
finally:
    word.Quit()
