# 現場照片勘驗生成器

AI 輔助司法現場勘驗文字產製系統，使用 Google Gemini Vision API 自動分析現場照片並產出勘驗文字，最終匯出 docx 報告。

## 系統畫面

> 上傳現場照片 → AI 自動產製勘驗文字 → 匯出 Word 報告

## 功能說明

| 功能 | 說明 |
|------|------|
| AI 自動辨識 | 呼叫 Gemini Vision API 分析照片，產出司法文書格式勘驗文字 |
| 提示詞補充 | 可針對單張照片補充指示（如：著重描述後保險桿受損）|
| 照片排序 | ↑↓ 調整照片順序，影響報告中的編排 |
| 批次生成 | 一鍵對所有照片批次執行 AI 分析 |
| docx 匯出 | 含標題、案號、日期、照片、勘驗文字的完整 Word 文件，字型使用標楷體 |

## 技術架構

```
使用者瀏覽器（Streamlit）
        ↓ 上傳照片
   app.py 後端處理
        ↓ Base64 圖片 + Prompt
  Google Gemini Vision API
        ↓ 繁體中文勘驗描述
   python-docx 產生報告
        ↓
   使用者下載 docx
```

## 快速啟動

### 本機執行

**步驟一：安裝套件**
```bash
pip install -r requirements.txt
```

**步驟二：啟動系統**
```bash
# Windows
python -m streamlit run app.py

# Mac / Linux
streamlit run app.py
```

瀏覽器自動開啟 http://localhost:8501

### Streamlit Cloud 部署

1. 前往 [https://streamlit.io/cloud](https://streamlit.io/cloud)
2. 用 GitHub 帳號登入
3. 點「New app」→ 選此 repo
4. Main file 選 `app.py`
5. 設定 Secret：`GEMINI_API_KEY = "AIza..."`
6. 點「Deploy」

## 使用方式

1. 在左側 Sidebar 輸入 **Gemini API Key**（格式：`AIza...`）
2. 填入**案號**（例：115重訴135）
3. **上傳現場照片**（支援 jpg / png / webp，多張）
4. 點擊「✨ AI 生成勘驗文字」或「⚡ 全部照片 AI 生成」
5. 確認、修改文字後點「📄 產生匯出 docx」
6. 點擊「⬇ 下載 docx 檔案」儲存報告

## 取得 Gemini API Key

1. 前往 [https://aistudio.google.com](https://aistudio.google.com)
2. 用 Google 帳號登入
3. 點左側「Get API key」→「Create API key」
4. 複製 Key（格式：`AIza...`）

> 免費方案：gemini-2.5-flash-lite 每日 1,000 次請求，Demo 使用完全足夠。

## 系統需求

- Python 3.9+
- 網路連線（呼叫 Gemini API）
- Gemini API Key（免費申請）

## 套件清單

```
streamlit>=1.32.0
google-generativeai>=0.8.0
python-docx>=1.1.0
Pillow>=10.0.0
```

## 費用估算

以 gemini-2.5-flash-lite 免費方案，每日可分析 1,000 張照片，一般業務使用完全免費。
如需更高用量或正式環境，可升級 Google AI Studio 付費方案。

## 未來規劃

- [ ] 本地端 Ollama 版本（不依賴雲端，資安更佳）
- [ ] PDF 匯出支援
- [ ] 案件資料庫整合
- [ ] 多人協作版本

## 開發單位

臺灣高雄地方檢察署 資訊室