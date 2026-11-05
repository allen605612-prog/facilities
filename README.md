# allen 工作資料夾

正心中學設備組與自然科教學的工作資料夾，包含設備組網頁、出題與講義、得獎海報、內控文件產生腳本等。

> ⚠️ 學生成績、名條等個資與私人檔案一律不上傳，規則見 `.gitignore`。

## 資料夾索引

| 位置 | 內容 |
|---|---|
| `設備組網頁/` | 正式上線的設備組網頁（GitHub Pages：facilities repo） |
| `設備組新網頁/` | 設備組網頁改版中的新版本 |
| 根目錄 `*.html` | 設備組各頁面（管理規則、競賽、科展、報修、表單下載等） |
| `設備組工具/` | 競賽得獎名單 → 校務系統匯入總表產生器 |
| `出題/`、`_exam_fig/` | 段考練習題庫與題目附圖 |
| `國二自然_*`、`國三自然_*` | 綜合練習、模擬測驗、段考練習卷（docx + pdf） |
| `海報_66屆/`、`a1_output/` | 第 66 屆科展得獎海報（A1 300 DPI） |
| `第66屆科展*` | 科展得獎橫幅（活潑版、照片版）與榮譽海報 |
| `DF模擬器/` | 科展 DF 模擬器程式與成果 |
| `day30-signup/` | AI 工作坊活動報名系統（Supabase × Vercel） |

## 主要腳本

| 腳本 | 用途 |
|---|---|
| `gen_exam_papers.py`、`gen_exam_docx.py`、`_gen_ch*_*.js/py`、`_gen_h5_*` | 產生考卷、練習卷與附圖 |
| `gen_answer_keys.py`、`gen_answer_sheets.py` | 產生答案卷、答案卡 |
| `gen_award_posters*.py`、`make_*poster.py`、`upscale_to_a1.py` | 得獎海報產生與放大為 A1 |
| `_gen_banner_*.py`、`_gen_poster_science.py` | 科展橫幅、海報 |
| `gen_ict_*.js`、`gen_forms_4x.js`、`gen_drp.js` | 設備組內控文件、表單、主機備援計畫（DRP） |
| `gen_lab_form.js` | 實驗室借用申請表 |
| `check_offline_tablets.py` | 未上網平板清單檢查 |
| `班級名條彙整_SOP.md` | 高中／國中名條彙整步驟 |

## 執行環境

- Python：用 `uv run` / `uvx` 執行
- Node.js：`npm install` 後執行 `node <腳本>.js`（`node_modules/` 不上傳）
