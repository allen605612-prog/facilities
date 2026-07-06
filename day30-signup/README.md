# Day 30 作業：活動報名系統（Supabase × Vercel）

「AI 工具實作工作坊」活動報名系統——公開表單收資料、管理者 Google 登入看資料。

## 架構

| 角色 | 服務 | 說明 |
|------|------|------|
| 前端 | Vercel | 靜態 HTML 兩頁：`index.html`（公開報名表單）、`admin.html`（管理者後台） |
| 後端資料庫 | Supabase (PostgreSQL) | `survey_responses` 存報名資料、`admins` 存管理者名單 |
| 登入 | Supabase Auth + Google OAuth | 只有管理者需要登入；用 `is_admin()` 函式核對 email |

## 資料表與權限（RLS）

- `survey_responses`：報名資料（姓名、Email、電話、場次、人數、備註）
  - `anon` 可 INSERT（任何人都能報名）
  - 只有 `admins` 名單內的登入者可 SELECT
- `admins`：管理者 email 名單
  - 登入者只能查到自己的那一列
- `is_admin()`：`security definer` 函式，判斷目前登入者 email 是否在 `admins` 表中

## 使用流程

1. 使用者打開首頁填表送出 → 資料寫入 Supabase `survey_responses`
2. 管理者開 `/admin.html` → 點「使用 Google 登入」
3. 登入後系統呼叫 `is_admin()` 核對名單：
   - 是管理者 → 顯示所有報名資料（筆數、總人數、明細表）
   - 不是 → 顯示「沒有存取權限」

## 新增管理者

在 Supabase SQL Editor 執行：

```sql
insert into public.admins (email) values ('someone@gmail.com');
```
