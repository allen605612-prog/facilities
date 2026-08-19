# 班級座號姓名彙整 SOP

## 用途
使用者會不定期提供「高中名條」與「國中名條」兩份 xlsx 檔（通常放在 `C:\Users\user\Downloads`，檔名格式類似 `115上高中名條.xlsx` / `115上國中名條.xlsx`，學年度數字會變動）。
只要使用者丟出這兩個檔案並要求「擷取資料與排列」，直接依照本文件的邏輯執行，不需要再詢問細節。

## 輸入檔案結構

### 高中名條.xlsx
- 單一工作表，橫向並排三個年級區塊：**高三、高二、高一**（由左到右）。
- 每個年級區塊固定 5 欄：`學號、組別、班級、座號、姓名`（後面可能還有備註/學群欄，忽略即可）。
- 資料從第 3 列開始（第 1 列是年級標題合併儲存格，第 2 列是欄名）。
- 高三、高二區塊欄位起始位置：高三從第 1 欄開始，高二區塊在其右側間隔一欄備註欄，高一同理。實務上用「班級」欄出現非空值來定位區塊，不要死記固定欄號，因為備註欄數可能變動。
- 班級名稱固定為：**正、心、誠、意、修、身**（各年級都一樣，順序即排班依據）。

### 國中名條.xlsx
- 同樣單一工作表，橫向並排三個年級區塊：**國一、國二、國三**（由左到右）。
- 每個年級區塊有兩個「班級」欄（重複欄名，其中一欄是數字組別、另一欄是中文班級名），欄位為：`學號、班級(數字)、班級(中文)、座號、姓名`（後面可能有備註欄）。
- 資料從第 3 列開始。
- 班級名稱固定為：**正、心、誠、意、修、身、齊、音甲、音乙、美甲、美乙**（各年級都一樣，順序即排班依據）。

## 讀取原則
- 一律用 `openpyxl.load_workbook(path, data_only=True)` 讀值，不要用 pandas 讀（合併表頭與多區塊排列 pandas 不好處理）。
- 不要假設固定欄索引，改用「班級（中文）」欄與「姓名」欄同時非空」來判斷一筆有效資料；用工作表第 1 列的年級標題文字找出每個區塊的起始欄，再往右找「班級」「座號」「姓名」欄。
- 座號轉成整數（原始可能是 float 或字串）。

## 排序規則（重點）
1. **年級順序（由前到後）**：高三 → 高二 → 高一 → 國三 → 國二 → 國一
2. **高中班級順序**：正、心、誠、意、修、身
3. **國中班級順序**：正、心、誠、意、修、身、齊、音甲、音乙、美甲、美乙
4. **班級編號**：按照上述「年級 × 班級」的完整順序，從 1 開始依序編號（高三正 = 1，高三心 = 2 ... 高三身 = 6，高二正 = 7 ...，依此類推）。
   - 驗證錯誤基準：**國一美乙 = 51**（若重新計算不是 51，表示年級或班級數量/順序算錯，要先檢查）。
5. 同一班內再依**座號由小到大**排序。
6. 最終總排序 key = (班級編號, 座號)。

## 輸出檔案
- 檔名：`{學年度}_全校班級座號姓名彙整.xlsx`（學年度字串直接取自輸入檔名前綴，例如輸入是 `115上高中名條.xlsx` 則輸出 `115上_全校班級座號姓名彙整.xlsx`）。
- 存放位置：與輸入檔案相同的資料夾（通常是 `Downloads`）。
- 欄位（依序）：`班級編號、年級、班級、座號、姓名`。
- 格式：
  - 全部字型 Arial。
  - 標題列粗體、置中。
  - 內容置中。
  - 凍結首列（`freeze_panes = "A2"`）。
  - 欄寬約：班級編號10、年級8、班級8、座號8、姓名12。

## 執行方式
- 用 `uv run --with openpyxl python <script>.py` 執行（不要用系統 pip/python，依使用者慣例優先用 uv/uvx）。
- 腳本寫在暫存 scratchpad 目錄即可，不需要留在專案資料夾。
- 完成後跑一次 sanity check：印出「高三正」與「國一美乙」對應的班級編號，確認分別是 1 與 51。

## 參考實作（Python，可直接套用調整路徑）

```python
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

hs_path = r"C:\Users\user\Downloads\115上高中名條.xlsx"
jh_path = r"C:\Users\user\Downloads\115上國中名條.xlsx"
out_path = r"C:\Users\user\Downloads\115上_全校班級座號姓名彙整.xlsx"

grade_order = ["高三", "高二", "高一", "國三", "國二", "國一"]
hs_classes = ["正", "心", "誠", "意", "修", "身"]
jh_classes = ["正", "心", "誠", "意", "修", "身", "齊", "音甲", "音乙", "美甲", "美乙"]

class_no = {}
n = 1
for g in grade_order:
    classes = hs_classes if g in ("高三", "高二", "高一") else jh_classes
    for c in classes:
        class_no[(g, c)] = n
        n += 1

def find_blocks(ws, grade_names):
    """依第1列的年級標題文字，找出每個年級區塊的起始欄，回傳 {grade: start_col_idx(0-based)}"""
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    blocks = {}
    for idx, v in enumerate(header):
        if v in grade_names:
            blocks[v] = idx
    return blocks

def extract(path, grade_names, class_names):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    header2 = [str(x) if x is not None else "" for x in rows[1]]
    starts = find_blocks(ws, grade_names)
    records = []
    for grade, start in starts.items():
        # 在 header2 中，從 start 起往右找 班級 / 座號 / 姓名 欄（取最後一個"班級"欄為中文班級名欄）
        seg = header2[start:start + 10]
        class_idx = [start + i for i, h in enumerate(seg) if h == "班級"]
        cname_idx = class_idx[-1] if class_idx else None
        seat_idx = start + seg.index("座號") if "座號" in seg else None
        name_idx = start + seg.index("姓名") if "姓名" in seg else None
        for r in rows[2:]:
            cname = r[cname_idx] if cname_idx is not None and cname_idx < len(r) else None
            seat = r[seat_idx] if seat_idx is not None and seat_idx < len(r) else None
            sname = r[name_idx] if name_idx is not None and name_idx < len(r) else None
            if not sname or not cname or cname not in class_names:
                continue
            records.append((class_no[(grade, cname)], grade, cname, int(seat), sname))
    return records

records = []
records += extract(hs_path, ["高三", "高二", "高一"], hs_classes)
records += extract(jh_path, ["國一", "國二", "國三"], jh_classes)
records.sort(key=lambda x: (x[0], x[3]))

out_wb = openpyxl.Workbook()
out_ws = out_wb.active
out_ws.title = "班級座號姓名"
out_ws.append(["班級編號", "年級", "班級", "座號", "姓名"])
for cell in out_ws[1]:
    cell.font = Font(name="Arial", bold=True)
    cell.alignment = Alignment(horizontal="center")
for rec in records:
    out_ws.append(rec)
for row in out_ws.iter_rows(min_row=2):
    for cell in row:
        cell.font = Font(name="Arial")
        cell.alignment = Alignment(horizontal="center")
for i, w in enumerate([10, 8, 8, 8, 12], start=1):
    out_ws.column_dimensions[get_column_letter(i)].width = w
out_ws.freeze_panes = "A2"
out_wb.save(out_path)

print("total records:", len(records))
print(class_no[("高三", "正")], class_no[("國一", "美乙")])  # 應為 1 51
```

> 注意：`find_blocks` 是通用化寫法，實際欄位排列若跟今年（115上）不同，仍要先用 openpyxl 印出前 2~3 列確認結構，再套用上面邏輯，不要盲目假設欄號。
