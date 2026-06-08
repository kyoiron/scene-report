# 現場照片勘驗生成器

AI 輔助司法現場勘驗文字產製系統，使用 Claude Vision API 自動分析現場照片並產出勘驗文字，最終匯出 docx 報告。

## 快速啟動

### 步驟一：安裝套件

```bash
pip install -r requirements.txt
```

### 步驟二：啟動系統

```bash
streamlit run app.py
```

瀏覽器會自動開啟 http://localhost:8501

### 步驟三：使用方式

1. 在左側 Sidebar 輸入 **Anthropic API Key**（sk-ant-...）
2. 填入**案號**（例：115重訴135）
3. **上傳現場照片**（支援 jpg / png / webp，多張）
4. 點擊「✨ AI 生成勘驗文字」或「⚡ 全部照片 AI 生成」
5. 確認、修改文字後點「📄 產生匯出 docx」
6. 點擊「⬇ 下載 docx 檔案」儲存報告

## 功能說明

| 功能 | 說明 |
|------|------|
| AI 自動辨識 | 呼叫 Claude Vision API 分析照片，產出司法文書格式勘驗文字 |
| 提示詞補充 | 可針對單張照片補充指示（如：著重描述後保險桿受損）|
| 照片排序 | ↑↓ 調整照片順序，影響報告中的編排 |
| 批次生成 | 一鍵對所有照片批次執行 AI 分析 |
| docx 匯出 | 含標題、案號、日期、照片、勘驗文字的完整 Word 文件 |

## 系統需求

- Python 3.9+
- 網路連線（呼叫 Anthropic API）
- Anthropic API Key

## 費用估算

每張現場照片約消耗 1,600–2,400 tokens（圖片大小而異）。
以 Claude Sonnet 定價，每張照片 AI 分析費用約 NT$0.3–0.5，極為低廉。
