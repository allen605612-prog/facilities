import img2pdf
from pathlib import Path

A1_W_MM, A1_H_MM = 594, 841
outdir = Path(r"C:\Users\user\allen\a1_output")

pngs = [
    "poster_66_01_數學科第一名.png",
    "poster_66_02_數學科第二名.png",
    "poster_66_03_生物科第二名.png",
    "poster_66_04_生物科第三名.png",
    "poster_66_05_生活與應用科學科(一)第四名.png",
    "poster_66_06_生活與應用科學科(二)第三名.png",
    "poster_66_07_生活與應用科學科(三)第四名.png",
    "poster_award.png",
]

layout = img2pdf.get_layout_fun(
    (img2pdf.mm_to_pt(A1_W_MM), img2pdf.mm_to_pt(A1_H_MM))
)
files = [str(outdir / p) for p in pngs]
merged = outdir / "海報集_第66屆_A1_300dpi_完整版.pdf"
with open(merged, "wb") as f:
    f.write(img2pdf.convert(files, layout_fun=layout))
print(f"merged: {merged}")
