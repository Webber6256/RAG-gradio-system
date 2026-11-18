# 多策略 RAG 文件問答（Gradio 版）

上傳 PDF / DOCX，選擇檢索策略後直接提問。Groq（LLM）× ChromaDB（向量庫）× Gradio（介面）。

支援 8 種檢索策略：混合搜尋（RRF）、語意搜尋、關鍵字（字元 n-gram TF-IDF）、重新排序、多查詢、壓縮、父子文檔、HyDE。可匯出問答紀錄為 Word / CSV。

## 專案結構

```
app.py            主程式（單檔）
requirements.txt  相依套件（使用 CPU 版 PyTorch）
render.yaml       Render Blueprint（選用）
.gitignore
```

## 本機執行

```bash
python -m venv .venv && source .venv/bin/activate     # Windows：.venv\Scripts\activate
pip install -r requirements.txt
export GROQ_API_KEY=gsk_xxx                            # Windows PowerShell：$env:GROQ_API_KEY="gsk_xxx"
python app.py                                          # 開啟 http://localhost:7860
```

## 上傳到 GitHub

```bash
git init
git add .
git commit -m "init: RAG gradio app"
git branch -M main
git remote add origin https://github.com/<你的帳號>/<repo 名稱>.git
git push -u origin main
```

## 部署到 Render

**方式 A：Blueprint（最快）** Render Dashboard → New → Blueprint → 選這個 repo，讀取 `render.yaml`，在畫面上填入 `GROQ_API_KEY`。

**方式 B：手動建立 Web Service**

| 欄位 | 值 |
|---|---|
| Runtime | Python 3 |
| Build Command | `pip install -r requirements.txt && python -c "from sentence_transformers import SentenceTransformer; SentenceTransformer('intfloat/multilingual-e5-small')"` |
| Start Command | `python app.py` |
| Instance Type | 建議 Starter（2 GB）以上 |

Environment Variables：

| 名稱 | 說明 |
|---|---|
| `GROQ_API_KEY` | 你的 Groq Key（console.groq.com 取得） |
| `PYTHON_VERSION` | `3.11.9` |
| `HF_HOME` | `/opt/render/project/src/.hf_cache`（讓嵌入模型在 build 階段下載） |
| `EMBED_MODEL` | 選填，預設 `intfloat/multilingual-e5-small` |

## 注意事項

- **記憶體**：torch + 嵌入模型常駐約 1 GB，Render 免費 512 MB 方案很可能啟動失敗。想省記憶體可維持 e5-small，不要換 e5-base / bge-m3。
- **單人使用設計**：這個版本的文件與設定是全域共用的，同時有多人使用會互相覆蓋文件，也會共用你設在環境變數的 Key。公開給多人使用請改用 `v2-pwa`。
- **檔案不持久**：Render 的磁碟在重新部署或重啟後會清空，向量庫與上傳內容都不會保留。
- **冷啟動**：付費方案不休眠；免費方案閒置後喚醒需要重新載入模型。
- 掃描版 PDF 沒有文字層，需先 OCR。
