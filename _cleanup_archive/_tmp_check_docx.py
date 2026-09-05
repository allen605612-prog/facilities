import docx
doc = docx.Document(r'C:\Users\user\allen\DF模擬器\科展 (2).docx')
for p in doc.paragraphs[:25]:
    if p.text.strip():
        print(p.text.strip())
