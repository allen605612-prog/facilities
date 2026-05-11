"""
將 PNG 海報放大至 A1 尺寸 300 DPI 並輸出 PDF
A1: 594mm × 841mm @ 300 DPI = 7016 × 9933 px
"""
from PIL import Image
import img2pdf
import os

A1_W_MM = 594
A1_H_MM = 841
DPI = 300
A1_W_PX = int(A1_W_MM * DPI / 25.4)  # 7016
A1_H_PX = int(A1_H_MM * DPI / 25.4)  # 9933

files = [
    "poster_66_01_數學科第一名.png",
    "poster_66_02_數學科第二名.png",
    "poster_66_03_生物科第二名.png",
    "poster_66_04_生物科第三名.png",
    "poster_66_05_生活應用科學一第四名.png",
    "poster_66_06_生活應用科學二第三名.png",
    "poster_66_07_生活應用科學三第四名.png",
    "poster_award.png",
]

os.makedirs("a1_output", exist_ok=True)

out_pngs = []

for fname in files:
    if not os.path.exists(fname):
        print(f"[SKIP] {fname} 不存在")
        continue

    img = Image.open(fname).convert("RGB")
    orig_w, orig_h = img.size
    print(f"處理: {fname}  原始: {orig_w}×{orig_h}")

    # 等比縮放填滿 A1（超出部分從中心裁切，不留白邊）
    scale = max(A1_W_PX / orig_w, A1_H_PX / orig_h)
    new_w = int(orig_w * scale)
    new_h = int(orig_h * scale)

    img_scaled = img.resize((new_w, new_h), Image.LANCZOS)

    # 從中心裁切至精確 A1 尺寸
    crop_x = (new_w - A1_W_PX) // 2
    crop_y = (new_h - A1_H_PX) // 2
    canvas = img_scaled.crop((crop_x, crop_y, crop_x + A1_W_PX, crop_y + A1_H_PX))

    # 設定 DPI 並儲存高解析度 PNG（pnginfo 帶 DPI）
    out_png = os.path.join("a1_output", fname)
    canvas.save(out_png, dpi=(DPI, DPI))
    print(f"  → PNG: {out_png}  ({A1_W_PX}×{A1_H_PX} @ {DPI} DPI)")
    out_pngs.append(out_png)

    # 個別 PDF
    stem = os.path.splitext(fname)[0]
    single_pdf = os.path.join("a1_output", stem + ".pdf")
    layout = img2pdf.get_layout_fun(
        (img2pdf.mm_to_pt(A1_W_MM), img2pdf.mm_to_pt(A1_H_MM))
    )
    with open(single_pdf, "wb") as f:
        f.write(img2pdf.convert(out_png, layout_fun=layout))
    print(f"  → PDF: {single_pdf}")

# 合併 PDF
if out_pngs:
    merged_pdf = os.path.join("a1_output", "海報集_A1_300dpi.pdf")
    layout = img2pdf.get_layout_fun(
        (img2pdf.mm_to_pt(A1_W_MM), img2pdf.mm_to_pt(A1_H_MM))
    )
    with open(merged_pdf, "wb") as f:
        f.write(img2pdf.convert(out_pngs, layout_fun=layout))
    print(f"\n✓ 合併 PDF → {merged_pdf}")

print("\n全部完成！")
