import fitz  # pymupdf

PDF_FILES = [
    r"C:\Users\user\Downloads\QZ00313561.pdf",
    r"C:\Users\user\Downloads\QZ00191189.pdf",
    r"C:\Users\user\Downloads\QZ00438651.pdf",
    r"C:\Users\user\Downloads\QZ00594691.pdf",
    r"C:\Users\user\Downloads\QZ00717454.pdf",
    r"C:\Users\user\Downloads\QZ00849192.pdf",
    r"C:\Users\user\Downloads\QZ01000069.pdf",
    r"C:\Users\user\Downloads\QZ01065948.pdf",
]
OUTPUT = r"C:\Users\user\Downloads\invoices_A4_v2.pdf"

A4_W, A4_H = 595.28, 841.89
COLS, ROWS = 4, 2
MARGIN = 0  # 緊密排列，無間距

cell_w = A4_W / COLS
cell_h = A4_H / ROWS


def get_content_rect(page):
    """Union of all text-block and drawing bounding boxes, with 5pt padding."""
    rects = []
    for b in page.get_text("blocks"):          # (x0,y0,x1,y1,text,…)
        r = fitz.Rect(b[0], b[1], b[2], b[3])
        if not r.is_empty:
            rects.append(r)
    for d in page.get_drawings():
        r = d.get("rect")
        if r and not r.is_empty:
            rects.append(r)

    if not rects:
        return page.rect

    combined = rects[0]
    for r in rects[1:]:
        combined |= r

    pad = 5
    return fitz.Rect(
        max(0, combined.x0 - pad),
        max(0, combined.y0 - pad),
        min(page.rect.width,  combined.x1 + pad),
        min(page.rect.height, combined.y1 + pad),
    )


out_doc = fitz.open()
out_page = out_doc.new_page(width=A4_W, height=A4_H)
out_page.draw_rect(out_page.rect, color=(1, 1, 1), fill=(1, 1, 1))  # white bg

for i, path in enumerate(PDF_FILES):
    row, col = divmod(i, COLS)

    doc = fitz.open(path)
    page = doc[0]
    clip = get_content_rect(page)

    cw, ch = clip.width, clip.height
    scale = min(cell_w / cw, cell_h / ch)
    sw, sh = cw * scale, ch * scale

    cell_x = col * cell_w
    cell_y = row * cell_h
    ox = cell_x + (cell_w - sw) / 2
    oy = cell_y + (cell_h - sh) / 2

    target = fitz.Rect(ox, oy, ox + sw, oy + sh)
    out_page.show_pdf_page(target, doc, 0, clip=clip)

    name = path.split("\\")[-1]
    print(f"[{i+1}/8] {name}  clip={clip}  scale={scale:.3f}")
    doc.close()

out_doc.save(OUTPUT, garbage=4, deflate=True)
out_doc.close()
print(f"\n完成！→ {OUTPUT}")
