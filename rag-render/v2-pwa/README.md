# 多策略 RAG 文件問答（PWA + RWD 版）

FastAPI 後端 + 原生 HTML/JS 前端，可安裝到手機桌面或電腦，手機 / 平板 / 桌面都能用。

- **PWA**：`manifest.json`、Service Worker（外殼離線可開啟、離線提示）、可安裝圖示（含 maskable 與 iOS）
- **RWD**：手機為單欄對話 + 滑出式設定面板；寬度 ≥ 900px 為左側設定 + 右側對話。支援安全區域（瀏海 / Home 指示列）、深色模式、`prefers-reduced-motion`
- **多使用者**：每個瀏覽器有獨立的工作階段與文件，互不覆蓋；閒置超過 6 小時自動清理
- **BYOK**：使用者可在畫面輸入自己的 Groq Key（只存在自己的瀏覽器），也可由伺服器環境變數提供
- 回答中的 `[片段n]` 會變成可點的標籤，點開直接看原文依據
- 8 種檢索策略、Word / CSV 匯出，與 Gradio 版相同

## 專案結構

```
main.py           FastAPI：API + 靜態檔
rag_core.py       RAG 核心（載入、檢索、生成、匯出）
static/
  index.html      介面（RWD）
  sw.js           Service Worker
  manifest.json   PWA manifest
  icons/          192 / 512 / maskable / apple-touch-icon
make_icons.py     重新產生圖示（選用，需 Pillow）
requirements.txt
render.yaml       Render Blueprint（選用）
```

## 本機執行

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
export GROQ_API_KEY=gsk_xxx        # 選填；不設的話在畫面輸入 Key
uvicorn main:app --reload --port 8000   # 開啟 http://localhost:8000
```

`localhost` 可直接測試 Service Worker 與安裝；其他網址需要 HTTPS（Render 預設提供）。

## 上傳到 GitHub

```bash
git init
git add .
git commit -m "init: RAG PWA app"
git branch -M main
git remote add origin https://github.com/<你的帳號>/<repo 名稱>.git
git push -u origin main
```

## 部署到 Render

**方式 A：Blueprint** New → Blueprint → 選 repo（讀取 `render.yaml`）。

**方式 B：手動建立 Web Service**

| 欄位 | 值 |
|---|---|
| Runtime | Python 3 |
| Build Command | `pip install -r requirements.txt && python -c "from sentence_transformers import SentenceTransformer; SentenceTransformer('intfloat/multilingual-e5-small')"` |
| Start Command | `uvicorn main:app --host 0.0.0.0 --port $PORT` |
| Health Check Path | `/api/health` |
| Instance Type | 建議 Starter（2 GB）以上 |

環境變數：

| 名稱 | 說明 |
|---|---|
| `PYTHON_VERSION` | `3.11.9` |
| `HF_HOME` | `/opt/render/project/src/.hf_cache` |
| `GROQ_API_KEY` | 選填。伺服器端預設 Key |
| `ALLOW_SERVER_KEY` | 預設 `1`：沒輸入 Key 的使用者會用伺服器 Key。**公開網站建議設 `0`**，避免別人消耗你的額度 |
| `MAX_UPLOAD_MB` | 上傳大小上限，預設 20 |
| `SESSION_TTL_HOURS` | 閒置多久清理文件，預設 6 |
| `EMBED_MODEL` | 選填，預設 `intfloat/multilingual-e5-small` |

服務啟動後會先開始監聽，嵌入模型在背景載入；右上角標籤顯示「引擎啟動中 → 引擎就緒」。

## 安裝成 App

- **Android / Chrome / Edge**：右上角出現「安裝」按鈕，或瀏覽器選單 → 安裝應用程式
- **iPhone / iPad（Safari）**：分享 → 加入主畫面
- 離線時可開啟外殼並查看本機保存的歷史問答；提問與上傳需要網路

## API

| 路徑 | 說明 |
|---|---|
| `GET /api/health` | 引擎狀態 |
| `GET /api/config` | 策略與預設模型 |
| `POST /api/models` | 驗證 Key 並列出可用模型（Header `X-Groq-Key`） |
| `POST /api/upload` | 表單：`session`、`file` |
| `POST /api/ask` | JSON：`session`、`q`、`strategy`、`k`、`model` |
| `POST /api/export/{docx\|csv}` | JSON：`records`、`with_src` |

## 注意事項

- 向量庫為記憶體內，Render 重啟或重新部署後需重新上傳文件（前端保留的歷史問答仍在瀏覽器）
- 更新前端後，把 `static/sw.js` 的 `CACHE` 版本號 +1，使用者才會拿到新外殼
- 掃描版 PDF 沒有文字層，需先 OCR
