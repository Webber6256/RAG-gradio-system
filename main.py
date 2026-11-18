"""FastAPI 後端 + PWA 靜態前端（static/）。啟動：uvicorn main:app --host 0.0.0.0 --port $PORT"""
import os, re, tempfile, threading, time
from pathlib import Path

from fastapi import FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from starlette.background import BackgroundTask

import rag_core as rc

BASE = Path(__file__).parent
SERVER_KEY = os.getenv("GROQ_API_KEY", "").strip()
ALLOW_SERVER_KEY = os.getenv("ALLOW_SERVER_KEY", "1") == "1"   # 設 0：只接受使用者自己的 Key
MAX_MB = int(os.getenv("MAX_UPLOAD_MB", "20"))
SESSION_TTL = int(os.getenv("SESSION_TTL_HOURS", "6")) * 3600

app = FastAPI(title="RAG 文件問答", docs_url=None, redoc_url=None)
rag = rc.RAG()
threading.Thread(target=rag.warmup, daemon=True).start()

_docs, _docs_lock = {}, threading.Lock()
_SID = re.compile(r"^[A-Za-z0-9_-]{8,64}$")


def get_doc(session: str) -> rc.Doc:
    if not _SID.match(session or ""): raise HTTPException(400, "無效的 session")
    with _docs_lock:
        doc = _docs.get(session)
        if doc is None: doc = _docs[session] = rc.Doc()
        return doc


def sweep():  # 清掉閒置過久的使用者文件，釋放記憶體與向量庫
    now = time.time()
    with _docs_lock:
        old = [s for s, d in _docs.items() if now - d.last > SESSION_TTL]
        for s in old: rag.drop(_docs.pop(s))


def pick_key(header_key):
    return (header_key or "").strip() or (SERVER_KEY if ALLOW_SERVER_KEY else "")


def need_ready():
    if rag.error: raise HTTPException(500, f"嵌入模型載入失敗：{rag.error}")
    if not rag.ready.is_set(): raise HTTPException(503, "嵌入模型載入中，請稍候幾秒再試")


@app.get("/api/health")
def health():
    return {"ready": rag.ready.is_set(), "error": rag.error, "device": rag.device, "embed_model": rc.EMBED_MODEL}


@app.get("/api/config")
def config():
    return {"strategies": [{"id": k, "name": v[0], "desc": v[1]} for k, v in rc.STRATEGIES.items()],
            "models": rc.MODELS, "server_key": bool(SERVER_KEY and ALLOW_SERVER_KEY), "max_mb": MAX_MB}


@app.post("/api/models")
def models(x_groq_key: str | None = Header(default=None)):
    key = pick_key(x_groq_key)
    if not key: return {"ok": False, "message": "⚠ 未設定 Key", "models": []}
    ms = rc.LLM(key, "").list_models()
    if not ms: return {"ok": False, "message": "✗ Key 無效或無法取得模型", "models": []}
    return {"ok": True, "message": f"✓ Key 有效 · {len(ms)} 個模型", "models": ms}


@app.post("/api/upload")
def upload(session: str = Form(...), file: UploadFile = File(...)):
    need_ready(); sweep()
    doc = get_doc(session)
    ext = Path(file.filename or "").suffix.lower()
    if ext not in (".pdf", ".docx"): raise HTTPException(400, "僅支援 PDF / DOCX")
    data = file.file.read(MAX_MB * 1024 * 1024 + 1)
    if len(data) > MAX_MB * 1024 * 1024: raise HTTPException(413, f"檔案超過 {MAX_MB} MB")
    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as t: t.write(data)
    try: msg = rag.load(doc, t.name, Path(file.filename).name)
    finally: os.unlink(t.name)
    return {"ok": msg.startswith("✓"), "message": msg, "name": doc.name, "chunks": len(doc.chunks)}


class AskIn(BaseModel):
    session: str
    q: str = Field(min_length=1, max_length=2000)
    strategy: str = "hybrid"
    k: int = Field(4, ge=1, le=10)
    model: str = ""


@app.post("/api/ask")
def ask(body: AskIn, x_groq_key: str | None = Header(default=None)):
    need_ready()
    if body.strategy not in rc.STRATEGIES: raise HTTPException(400, "未知的檢索策略")
    doc = get_doc(body.session)
    llm = rc.LLM(pick_key(x_groq_key), body.model.strip())
    try:
        return rag.answer(doc, llm, body.q.strip(), body.strategy, body.k)
    except Exception as e:
        raise HTTPException(500, f"{type(e).__name__}: {e}")


class ExportIn(BaseModel):
    records: list[dict] = Field(min_length=1, max_length=300)
    with_src: bool = True


@app.post("/api/export/{fmt}")
def export(fmt: str, body: ExportIn):
    if fmt == "docx":
        path, mt = rc.export_docx(body.records, body.with_src), "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    elif fmt == "csv":
        path, mt = rc.export_csv(body.records), "text/csv"
    else: raise HTTPException(404, "不支援的格式")
    return FileResponse(path, media_type=mt, filename=f"rag-records.{fmt}", background=BackgroundTask(os.unlink, path))


@app.middleware("http")
async def no_cache_shell(request, call_next):  # 外殼與 Service Worker 每次向伺服器驗證，確保更新能被偵測
    resp = await call_next(request)
    if request.url.path in ("/", "/index.html", "/sw.js", "/manifest.json"):
        resp.headers["Cache-Control"] = "no-cache"
    return resp


app.mount("/", StaticFiles(directory=BASE / "static", html=True), name="static")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.getenv("PORT", "8000")))
