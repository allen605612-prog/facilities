import win32com.client
import os

FILES = [
    r"D:\D\教學資料\物理科\國中物理\講義\講義15-1時間.doc",
    r"D:\D\教學資料\物理科\國中物理\講義\講義15-2位置和位移.doc",
    r"D:\D\教學資料\物理科\國中物理\講義\講義15-3速度.doc",
    r"D:\D\教學資料\物理科\國中物理\講義\講義15-4加速度.doc",
]
OUTPUT = r"D:\D\教學資料\物理科\國中物理\講義\講義15_時間與運動合冊.doc"

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

try:
    doc = word.Documents.Open(os.path.abspath(FILES[0]))
    print(f"[1/4] 開啟：{os.path.basename(FILES[0])}")

    for i, path in enumerate(FILES[1:], start=2):
        sel = word.Selection
        sel.EndKey(Unit=6)          # wdStory：移到最末
        sel.InsertBreak(Type=7)     # wdPageBreak：插入分頁
        sel.InsertFile(
            FileName=os.path.abspath(path),
            ConfirmConversions=False,
            Link=False,
            Attachment=False,
        )
        print(f"[{i}/4] 插入：{os.path.basename(path)}")

    doc.SaveAs2(os.path.abspath(OUTPUT), FileFormat=0)  # 0 = wdFormatDocument (.doc)
    doc.Close(SaveChanges=False)
    print(f"\n完成！→ {OUTPUT}")
finally:
    word.Quit()
