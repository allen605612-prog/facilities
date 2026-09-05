import os, re, glob

html_files = glob.glob('設備組網頁/*.html') + glob.glob('*.html')
fixed = []

for path in html_files:
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    original = content

    # 1. target="_blank" 缺少 rel="noopener noreferrer" → 補上
    def fix_blank(m):
        tag = m.group(0)
        if 'noopener' in tag:
            return tag
        if 'rel=' in tag:
            tag = re.sub(r'rel="([^"]+)"', r'rel="\1 noopener noreferrer"', tag)
        else:
            tag = tag.replace('target="_blank"', 'target="_blank" rel="noopener noreferrer"')
        return tag

    content = re.sub(r'<a\s[^>]*target="_blank"[^>]*>', fix_blank, content)

    # 2. Google Fonts stylesheet → 補上 crossorigin="anonymous"
    def fix_gfonts(m):
        tag = m.group(0)
        if 'crossorigin' in tag:
            return tag
        return tag.replace('rel="stylesheet"', 'crossorigin="anonymous" rel="stylesheet"')

    content = re.sub(r'<link\s[^>]*fonts\.googleapis\.com[^>]*>', fix_gfonts, content)

    if content != original:
        with open(path, 'w', encoding='utf-8') as f:
            f.write(content)
        fixed.append(path)
        print(f'已修正：{path}')

print(f'\n共修正 {len(fixed)} 個檔案')
