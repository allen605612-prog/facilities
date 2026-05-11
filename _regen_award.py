import sys
sys.path.insert(0, r"C:\Users\user\allen")
from gen_award_posters import make_poster
from pathlib import Path

award = {
    "subject": "環境學科",
    "rank": "優等",
    "names": [
        ("高一心班", ["宋彥霖"]),
        ("高一意班", ["鐘宥昕"]),
        ("高一正班", ["廖柃柃"]),
    ],
    "teacher": "張祐誠",
}
out = Path(r"C:\Users\user\allen\a1_output\poster_award.png")
make_poster(award, "66", "第四區", "分區高中科展", out, seed=42)
print("done")
