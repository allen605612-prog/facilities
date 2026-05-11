"""
科展防禦答辯模擬器 — Gradio 網頁互動版
功能：
  1. 讀取科展 PDF 報告
  2. 以三個評審維度（研究動機 / 操作與控制變因 / 數據與結論）逐一提問
  3. 接收學生回答後追問，形成完整的答辯模擬對話
  4. 對話記憶（Memory）在整個 Gradio Session 中完整保留
使用方式：執行 `python app.py` 後開啟瀏覽器連至本機網址
"""

import os
import pdfplumber
import anthropic
import gradio as gr

# ── 常數設定 ──────────────────────────────────────────────────────────────────

# PDF 檔案路徑（與 app.py 同目錄）
PDF_檔案路徑 = os.path.join(os.path.dirname(__file__), "科展.pdf")

# 使用的 Claude 模型
CLAUDE_模型 = "claude-opus-4-6"

# API 金鑰（請勿將此金鑰提交至公開版本庫）
ANTHROPIC_API金鑰 = os.environ.get("ANTHROPIC_API_KEY", "")

# 三個評審維度的名稱
評審維度列表 = [
    "研究動機的合理性",
    "操作變因與控制變因的嚴謹度",
    "實驗數據與結論的連結",
]

# 系統提示詞（System Prompt）— 定義 Claude 的評審角色與提問規則
SYSTEM_PROMPT = """你是一位嚴謹但具啟發性的全國中小學科展評審委員。
學生正在進行專題口頭答辯，你負責提問並評估其研究能力。

【評審準則】
1. 每次只提出一個具體問題，不得一次列出多個問題。
2. 問題必須引用報告中的具體數字、圖表名稱、實驗步驟或結論句子。
3. 語氣專業、客觀，帶有教育引導意味；避免過於刁難，重在啟發思考。
4. 依序聚焦三個維度：
   - 第一維度：研究動機的合理性（為何做此研究？動機是否充分？）
   - 第二維度：操作變因與控制變因的嚴謹度（實驗設計是否嚴謹？）
   - 第三維度：實驗數據與結論的連結（數據是否支撐結論？推論邏輯是否正確？）
5. 學生回答後，若回答不完整，請追問一次再進入下一維度。
6. 全程使用繁體中文。
7. 在每個問題前標示目前所在維度，格式：【維度 X／3：＿＿＿＿】

【開場規則】
首先主動提出第一維度的第一個問題，不要等學生先說話。"""


# ── 核心函數 ──────────────────────────────────────────────────────────────────

def 讀取PDF文字(pdf路徑: str) -> str:
    """
    使用 pdfplumber 開啟指定的 PDF 檔案，
    逐頁提取純文字並合併為單一字串後回傳。
    """
    所有頁面文字 = []
    with pdfplumber.open(pdf路徑) as pdf檔案:
        for 頁面 in pdf檔案.pages:
            頁面文字 = 頁面.extract_text()
            if 頁面文字:
                所有頁面文字.append(頁面文字)
    return "\n\n".join(所有頁面文字)


def 建立初始對話訊息(科展報告文字: str) -> list[dict]:
    """
    建立對話的起始 user 訊息，將科展報告全文嵌入其中。
    """
    初始使用者訊息 = (
        "以下是學生的科展報告全文，請仔細閱讀後，"
        "依照你的評審角色，立刻提出第一個答辯問題。\n\n"
        f"【科展報告全文】\n{科展報告文字}"
    )
    return [{"role": "user", "content": 初始使用者訊息}]


def 呼叫Claude並取得回覆(對話紀錄: list[dict]) -> str:
    """
    將目前的完整對話紀錄送至 Claude API，取得評審的下一個回應。
    每次呼叫時建立新的客戶端實例（客戶端本身無狀態，記憶完全由對話紀錄承載）。
    """
    客戶端 = anthropic.Anthropic(api_key=ANTHROPIC_API金鑰)
    回應 = 客戶端.messages.create(
        model=CLAUDE_模型,
        max_tokens=1024,
        system=SYSTEM_PROMPT,
        messages=對話紀錄,
    )
    return 回應.content[0].text


# ── Gradio 事件處理函數 ────────────────────────────────────────────────────────

def 載入初始對話() -> tuple[list[dict], list[dict]]:
    """
    Gradio 頁面載入時執行：
      1. 讀取 PDF
      2. 呼叫 Claude 取得第一個評審問題
      3. 回傳初始聊天記錄與對話紀錄狀態
    """
    科展報告文字 = 讀取PDF文字(PDF_檔案路徑)
    對話紀錄 = 建立初始對話訊息(科展報告文字)

    # 取得評審的第一個問題
    評審問題 = 呼叫Claude並取得回覆(對話紀錄)

    # 將評審問題加入對話紀錄（Anthropic API 格式，用於記憶）
    對話紀錄.append({"role": "assistant", "content": 評審問題})

    # Gradio Chatbot 顯示格式（type="messages"）
    chatbot顯示 = [{"role": "assistant", "content": 評審問題}]

    return chatbot顯示, 對話紀錄


def 處理學生回答(
    學生輸入: str,
    chatbot顯示: list[dict],
    對話紀錄: list[dict],
) -> tuple[list[dict], list[dict], str]:
    """
    學生送出回答時執行：
      1. 將學生回答加入對話紀錄（記憶保留）
      2. 呼叫 Claude 取得評審回應
      3. 更新 Chatbot 顯示與對話紀錄狀態
      4. 清空輸入框

    Returns:
        (更新後的 chatbot 顯示, 更新後的對話紀錄, 清空的輸入框文字)
    """
    學生輸入 = 學生輸入.strip()
    if not 學生輸入:
        # 空白輸入不處理
        return chatbot顯示, 對話紀錄, ""

    # 將學生回答同時加入兩份紀錄
    對話紀錄.append({"role": "user", "content": 學生輸入})
    chatbot顯示.append({"role": "user", "content": 學生輸入})

    # 呼叫 Claude（完整對話紀錄作為記憶傳入）
    評審回應 = 呼叫Claude並取得回覆(對話紀錄)

    # 將評審回應加入兩份紀錄
    對話紀錄.append({"role": "assistant", "content": 評審回應})
    chatbot顯示.append({"role": "assistant", "content": 評審回應})

    return chatbot顯示, 對話紀錄, ""


def 重新開始() -> tuple[list[dict], list[dict], str]:
    """
    清空所有狀態並重新載入第一個評審問題。
    """
    chatbot顯示, 對話紀錄 = 載入初始對話()
    return chatbot顯示, 對話紀錄, ""


# ── Gradio 介面建構 ────────────────────────────────────────────────────────────

維度說明文字 = "\n".join(
    f"{i}. {名稱}" for i, 名稱 in enumerate(評審維度列表, start=1)
)

with gr.Blocks(title="科展防禦答辯模擬器") as demo:

    gr.Markdown(
        f"""
        # 科展防禦答辯模擬器 v2.0
        ### 三維度互動式口試訓練系統
        評審將依序就以下三個維度提問：
        {維度說明文字}
        ---
        """
    )

    # 聊天記錄顯示區（type="messages" 支援完整對話記憶）
    chatbot = gr.Chatbot(label="答辯模擬對話")

    # 對話紀錄狀態（儲存 Anthropic API 格式的完整記憶，不含 PDF 內文以節省顯示空間）
    # 注意：此 State 存放完整的 messages list，是記憶功能的核心
    對話紀錄_state = gr.State([])

    with gr.Row():
        輸入框 = gr.Textbox(
            placeholder="請在此輸入你的回答，按 Enter 或點擊「送出」...",
            label="你的回答",
            lines=3,
            scale=4,
        )
        送出按鈕 = gr.Button("送出回答", variant="primary", scale=1)

    重新開始按鈕 = gr.Button("🔄 重新開始模擬", variant="secondary")

    gr.Markdown(
        "_提示：每輪對話的完整記憶會自動保留，評審將根據你先前所有的回答決定下一個問題。_"
    )

    # ── 事件綁定 ──────────────────────────────────────────────────────────────

    # 頁面載入時自動取得第一個評審問題
    demo.load(
        fn=載入初始對話,
        outputs=[chatbot, 對話紀錄_state],
    )

    # 按下送出按鈕
    送出按鈕.click(
        fn=處理學生回答,
        inputs=[輸入框, chatbot, 對話紀錄_state],
        outputs=[chatbot, 對話紀錄_state, 輸入框],
    )

    # 在輸入框按 Enter 鍵
    輸入框.submit(
        fn=處理學生回答,
        inputs=[輸入框, chatbot, 對話紀錄_state],
        outputs=[chatbot, 對話紀錄_state, 輸入框],
    )

    # 重新開始按鈕
    重新開始按鈕.click(
        fn=重新開始,
        outputs=[chatbot, 對話紀錄_state, 輸入框],
    )


# ── 程式進入點 ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    demo.launch()
