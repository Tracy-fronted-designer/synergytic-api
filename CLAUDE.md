# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 專案概覽

**信瑞企管 AI 企劃書自動化系統**

使用者在 Notion 輸入自然語言需求 → n8n 協調 → 從 OneDrive 搜尋相關舊企劃書 → Claude AI 分析 → 產出新企劃書（PPT）。

這個 repo（synergytic-api）是整個系統的 **FastAPI 服務**，負責 PPT 產出與 Excel 索引建立，由 n8n 透過 HTTP 呼叫。部署目標為 Render（免費方案）。

## Environment

Python 3.12。使用 `venv312`（不要用 `venv`，那是 Python 3.14，pydantic-core 會編譯失敗）：

```powershell
.\venv312\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Commands

```powershell
# 啟動開發伺服器（含 Swagger UI：http://localhost:8000/docs）
uvicorn main:app --reload
```

## API 端點

| 端點 | 方法 | 說明 |
|---|---|---|
| `/health` | GET | 確認服務存活 |
| `/scan-proposals` | POST | 掃描 OneDrive 所有企劃書 |
| `/generate-index` | POST | 產出 Excel 索引檔 |
| `/generate-pptx` | POST | 依內容產出 PPT 企劃書 |

## 系統架構（n8n Workflows）

**Workflow 1 — 搜尋企劃書**
Notion 新增記錄（Not started）→ Claude 解析客戶名稱／主題／講師 → 從 OneDrive 221 個客戶資料夾找相關資料夾與檔案 → 搜尋結果寫回 Notion（狀態改為「選擇中」）

**Workflow 2 — 產出企劃書**
Notion 狀態改為「產出中」→ Graph API 下載選定企劃書（轉 PDF）→ Claude 讀取分析 → 產出新企劃書文字（每 1500 字切段）→ 寫入 Notion 頁面

**Workflow 3 — 建立企劃書索引（進行中）**
掃描所有企劃書 PPT/PDF → Claude 解析內容 → 整理成 Excel 索引（欄位：客戶名稱、產業別、課程主題、痛點標籤、適合對象、講師、課程天數、課程模式、摘要、檔案名稱、檔案ID、建立日期）

## 重要 ID

| 項目 | ID |
|---|---|
| OneDrive 企劃書根目錄 | `42B13569E5D6C5E7!533` |
| Notion API Key | `ntn_568275460094J5FuEvebPLJAyjdkg5ll7Ujd1UMpo4IasxAz` |
| Azure App Client ID | `02ffc977-e0ef-45e0-92f4-c88d10becd78` |

OneDrive PDF 下載 URL 格式：
```
https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content?format=pdf
```

## Claude API

- 模型：`claude-sonnet-4-6`
- 方案：Tier 2（450K input tokens/分鐘）
- n8n Cloud：https://synergytic.app.n8n.cloud

## Stack

- **FastAPI 0.111** + **Pydantic v2** — API 框架與資料驗證
- **python-pptx** — 產出 PPT 企劃書
- **openpyxl** — 建立 Excel 索引
- **httpx** — 呼叫 Microsoft Graph API / OneDrive
- **uvicorn** — ASGI 伺服器
