import glob
import os
import re
import pandas as pd
from datetime import datetime

def natural_key(s):
    parts = re.split(r'(\d+)', str(s))
    return [int(p) if p.isdigit() else p.lower() for p in parts]

DOWNLOADS = r"C:\Users\user\Downloads"
SURFACE_REF = r"C:\Users\user\Downloads\113學年度surfacego名稱與機碼.xlsx"
NOW = datetime.now()
YEAR_MONTH = NOW.strftime("%Y-%m")  # e.g. "2026-03"

# 找最新的載具清單檔案
pattern = os.path.join(DOWNLOADS, "*_載具清單_*.xlsx")
files = glob.glob(pattern)
if not files:
    print("找不到載具清單檔案，請確認下載資料夾。")
    exit(1)

latest_file = max(files, key=os.path.getmtime)
print(f"讀取載具清單：{os.path.basename(latest_file)}")

df = pd.read_excel(latest_file)

if "最後連線" not in df.columns:
    print("找不到「最後連線」欄位，請確認檔案格式。")
    exit(1)

# 轉換最後連線為 datetime
df["最後連線_dt"] = pd.to_datetime(df["最後連線"], errors="coerce")

# 篩選條件：空白 或 不在當月
is_empty = df["最後連線_dt"].isna()
not_this_month = df["最後連線_dt"].dt.to_period("M") != YEAR_MONTH

offline = df[is_empty | not_this_month].copy()
offline.loc[is_empty[is_empty].index, "未上網原因"] = "從未連線"
offline.loc[(~is_empty & not_this_month)[~is_empty & not_this_month].index, "未上網原因"] = "當月未連線"
offline = offline.drop(columns=["最後連線_dt"])

print(f"當月未上網平板：共 {len(offline)} 台")

# ── Surface GO 位置對照 ──────────────────────────────────────────
print(f"讀取 Surface GO 參考檔：{os.path.basename(SURFACE_REF)}")
ref = pd.read_excel(SURFACE_REF, sheet_name=0, usecols=["學校平板編號末三碼", "機器編碼全部", "位置"])
ref = ref.rename(columns={"機器編碼全部": "serialNumber", "學校平板編號末三碼": "Surface GO 編號"})
ref["serialNumber"] = ref["serialNumber"].astype(str).str.strip()
offline["serialNumber"] = offline["serialNumber"].astype(str).str.strip()

# 合併位置資訊
offline = offline.merge(ref, on="serialNumber", how="left")

DROP_COLS = ["縣市", "學校代碼", "學校名稱"]

# Surface GO 工作表：有比對到位置的（即 Surface GO 裝置）
surface_go = offline[offline["位置"].notna()].copy()
surface_go = surface_go.drop(columns=[c for c in DROP_COLS if c in surface_go.columns])
surface_go["Surface GO 編號"] = pd.to_numeric(surface_go["Surface GO 編號"], errors="coerce")
surface_go = surface_go.sort_values("Surface GO 編號")

# 依 OS 工作表：只保留 iOS 裝置，依載具名稱自然排序
by_os = offline[offline["作業系統"].astype(str).str.contains("iOS", case=False, na=False)].copy()
by_os = by_os.drop(columns=[c for c in DROP_COLS if c in by_os.columns])
by_os = by_os.iloc[sorted(range(len(by_os)), key=lambda i: natural_key(by_os["載具名稱"].iloc[i]))]

# ── 輸出 Excel ───────────────────────────────────────────────────
output_file = os.path.join(DOWNLOADS, f"{NOW.strftime('%Y%m%d')}_未上網平板清單.xlsx")

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    surface_go.to_excel(writer, sheet_name="Surface GO 位置", index=False)
    by_os.to_excel(writer, sheet_name="依OS分類", index=False)

    for sheet_name in ["Surface GO 位置", "依OS分類"]:
        ws = writer.sheets[sheet_name]
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = "portrait"
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True

print(f"\n  Surface GO 未上網：{len(surface_go)} 台")
print(f"  全部未上網（依OS）：{len(by_os)} 台")
print(f"\n輸出檔案：{output_file}")
print("工作表：「Surface GO 位置」、「依OS分類」")
