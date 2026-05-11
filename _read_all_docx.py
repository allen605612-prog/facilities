import docx, sys
sys.stdout.reconfigure(encoding='utf-8')

files = [
    r"C:\Users\user\allen\GEMINI建議.docx",
    r"C:\Users\user\allen\_66屆科展得獎.docx",
]
for fp in files:
    print(f"\n{'='*50}")
    print(f"檔案：{fp}")
    print('='*50)
    try:
        doc = docx.Document(fp)
        for p in doc.paragraphs:
            if p.text.strip():
                print(p.text.strip())
    except Exception as e:
        print(f"錯誤：{e}")
