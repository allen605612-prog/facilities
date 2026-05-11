import sys, io, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

REPAIR_URL = "https://sites.google.com/shsh.ylc.edu.tw/equipment-repair/%E5%A1%AB%E5%AF%AB%E5%A0%B1%E4%BF%AE%E8%A1%A8%E5%96%AE?authuser=0"

# ── 讀取 智慧黑板管理.html 作為樣板 ──────────────────────────────
with open("智慧黑板管理.html", encoding="utf-8") as f:
    tmpl = f.read()

# 替換 title
new_page = tmpl.replace(
    "<title>智慧黑板管理 — 正心中學教務處</title>",
    "<title>設備報修系統 — 正心中學教務處</title>"
)
# 替換 active 麵包屑、h1
new_page = new_page.replace(
    '<div class="hero-breadcrumb"><a href="設備組.html">設備組首頁</a><span>›</span><a href="設備組.html">設備組</a><span>›</span>智慧黑板管理</div>',
    '<div class="hero-breadcrumb"><a href="設備組.html">設備組首頁</a><span>›</span><a href="設備組.html">設備組</a><span>›</span>設備報修系統</div>'
)
new_page = new_page.replace(
    "<h1>🖥️ 智慧黑板管理</h1>",
    "<h1>🔧 設備報修系統</h1>"
)
# 替換 active-sub（dropdown 中）
new_page = new_page.replace(
    '<a href="智慧黑板管理.html" class="active-sub">智慧黑板管理</a>',
    '<a href="智慧黑板管理.html">智慧黑板管理</a><a href="設備報修系統.html" class="active-sub">設備報修系統</a>'
)
# 替換 sidebar active
new_page = new_page.replace(
    '<a href="智慧黑板管理.html" class="active">智慧黑板管理</a>',
    '<a href="智慧黑板管理.html">智慧黑板管理</a><a href="設備報修系統.html" class="active">設備報修系統</a>'
)
# 替換主內容
old_main = """    <div class="section-label"><h2>智慧黑板管理</h2></div>
    <div class="content-card">
      <p class="section-note">智慧黑板相關管理辦法與申請，請洽設備組，分機 109。</p>
      <p style="color:var(--text-muted);font-size:.9rem;padding:32px 0;text-align:center;">相關文件待上傳，請洽設備組。</p>
</div>
    </div>"""
new_main = f"""    <div class="section-label"><h2>設備報修系統</h2></div>
    <div class="content-card">
      <p class="section-note">如需申請設備報修，請點擊下方連結填寫報修表單，送出後設備組將盡快安排處理。</p>
      <div style="text-align:center;padding:36px 0;">
        <a href="{REPAIR_URL}" target="_blank"
           style="display:inline-flex;align-items:center;gap:10px;
                  padding:14px 32px;border-radius:10px;
                  background:linear-gradient(135deg,#0369a1,#0ea5e9);
                  color:#fff;font-size:1rem;font-weight:600;
                  text-decoration:none;letter-spacing:.05em;
                  box-shadow:0 4px 16px rgba(3,105,161,0.3);
                  transition:transform .2s,box-shadow .2s;"
           onmouseover="this.style.transform='translateY(-2px)';this.style.boxShadow='0 6px 22px rgba(3,105,161,0.4)'"
           onmouseout="this.style.transform='';this.style.boxShadow='0 4px 16px rgba(3,105,161,0.3)'">
          🔧 填寫設備報修表單
        </a>
        <p style="margin-top:16px;font-size:.82rem;color:var(--text-muted);">點擊後將跳轉至 Google Sites 報修表單頁面</p>
      </div>
    </div>"""
new_page = new_page.replace(old_main, new_main)

with open("設備報修系統.html", "w", encoding="utf-8") as f:
    f.write(new_page)
print("created: 設備報修系統.html")

# ── 更新其他 9 個頁面：側欄 + 下拉選單 ──────────────────────────
files = [
    "設備組.html", "設備組成員.html", "智慧黑板管理.html", "奧林匹亞競賽.html",
    "活動一覽表.html", "教學設備管理規則.html", "專科教室管理規則.html",
    "科展消息.html", "數學與自然科能力競賽.html"
]

OLD_SIDEBAR = '<a href="智慧黑板管理.html">智慧黑板管理</a>'
NEW_SIDEBAR = '<a href="智慧黑板管理.html">智慧黑板管理</a><a href="設備報修系統.html">設備報修系統</a>'

OLD_DROPDOWN = '<a href="智慧黑板管理.html">智慧黑板管理</a>'
# dropdown 有時候 active-sub 在上面，用更完整字串
OLD_DROPDOWN2 = '<a href="智慧黑板管理.html" class="active-sub">智慧黑板管理</a>'
NEW_DROPDOWN2 = '<a href="智慧黑板管理.html" class="active-sub">智慧黑板管理</a><a href="設備報修系統.html">設備報修系統</a>'

for fname in files:
    with open(fname, encoding="utf-8") as f:
        html = f.read()

    # sidebar：在 智慧黑板管理 後插入（非 active 版本）
    if OLD_SIDEBAR in html and NEW_SIDEBAR not in html:
        html = html.replace(OLD_SIDEBAR, NEW_SIDEBAR, 1)

    # dropdown（active-sub 版本，只在 智慧黑板管理.html 本頁）
    if OLD_DROPDOWN2 in html:
        html = html.replace(OLD_DROPDOWN2, NEW_DROPDOWN2, 1)

    with open(fname, "w", encoding="utf-8") as f:
        f.write(html)
    print("updated: " + fname)

# ── 設備組.html：新增卡片 ──────────────────────────────────────
with open("設備組.html", encoding="utf-8") as f:
    main_html = f.read()

OLD_LAST_CARD = '<div class="card-icon">📐</div>\n        <div class="card-title">數學與自然科能力競賽</div>\n        <div class="card-arrow">前往 →</div>\n      </a>'
NEW_LAST_CARD = f'<div class="card-icon">📐</div>\n        <div class="card-title">數學與自然科能力競賽</div>\n        <div class="card-arrow">前往 →</div>\n      </a>\n\n      <a class="card" href="設備報修系統.html">\n        <div class="card-icon">🔧</div>\n        <div class="card-title">設備報修系統</div>\n        <div class="card-arrow">前往 →</div>\n      </a>'

if OLD_LAST_CARD in main_html:
    main_html = main_html.replace(OLD_LAST_CARD, NEW_LAST_CARD)
    with open("設備組.html", "w", encoding="utf-8") as f:
        f.write(main_html)
    print("card added to: 設備組.html")

print("all done")
