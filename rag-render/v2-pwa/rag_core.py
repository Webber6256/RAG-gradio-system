"""RAG 核心：文件載入、八種檢索策略、Groq 生成、Word / CSV 匯出（與 Web 框架無關）"""
import csv, os, re, tempfile, threading, time, uuid
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path

import chromadb
import numpy as np
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.shared import Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph
from groq import Groq
from pypdf import PdfReader
from sklearn.feature_extraction.text import TfidfVectorizer

# ── 設定 ──
EMBED_MODEL = os.getenv("EMBED_MODEL", "intfloat/multilingual-e5-small")  # 可換 multilingual-e5-base、BAAI/bge-m3
MODELS = ["openai/gpt-oss-20b", "openai/gpt-oss-120b", "qwen/qwen3.8-27b"]
CHUNK, OVERLAP, CHILD, CHILD_OVERLAP, MAX_CTX = 500, 100, 200, 40, 7000

STRATEGIES = {
    "hybrid": ("⚡ 混合搜尋", "語意 + 關鍵字，RRF 融合排名（推薦）"),
    "semantic": ("🔍 語意搜尋", "向量相似度檢索"),
    "tfidf": ("📊 關鍵字", "字元 n-gram，適合專有名詞、條號、數字"),
    "rerank": ("🎯 重新排序", "LLM 對候選片段批次評分後重排"),
    "multi_query": ("🔄 多查詢", "LLM 改寫 3 個角度問題聯合檢索"),
    "compress": ("✂️ 壓縮", "LLM 抽取每段最相關的句子"),
    "parent_child": ("📂 父子文檔", "小片段定位，回傳完整大片段"),
    "hyde": ("💡 HyDE", "先生成假設答案再檢索"),
}
_SENT = re.compile(r"(?<=[。！？!?；;])|\n+")


def split_text(text, size, overlap):
    sents = [s.strip() for s in _SENT.split(re.sub(r"[ \t　\xa0]+", " ", text)) if s and s.strip()]
    chunks, buf, fresh = [], "", 0
    for p in (s[i:i + size] for s in sents for i in range(0, len(s), size)):
        if fresh and len(buf) + len(p) + 1 > size:
            chunks.append(buf); buf, fresh = (buf[-overlap:].lstrip() if overlap else ""), 0
        buf = f"{buf} {p}" if buf else p; fresh += len(p)
    return chunks + ([buf.strip()] if fresh else [])


class LLM:
    """每次請求建立一個，Key 由使用者（瀏覽器）帶入，伺服器不保存。"""

    def __init__(self, key, model):
        self.client = Groq(api_key=key.strip()) if key and key.strip() else None
        self.model = model or MODELS[0]

    def list_models(self):
        if not self.client: return []
        try:
            ids = [m.id for m in self.client.models.list().data if getattr(m, "active", True) is not False]
        except Exception: return []
        return sorted(i for i in ids if not any(s in i.lower() for s in ("whisper", "tts", "orpheus", "guard", "playai")))

    def __call__(self, prompt, max_tokens=300, temp=0.3, system=None, raise_err=False):
        if not self.client: return None
        kw = dict(model=self.model, temperature=temp, max_completion_tokens=max_tokens,
                  messages=([{"role": "system", "content": system}] if system else []) + [{"role": "user", "content": prompt}])
        if "gpt-oss" in self.model:  # 推理模型：推理 token 也計入上限
            kw.update(max_completion_tokens=max(max_tokens * 3, 1024), reasoning_effort="low", include_reasoning=False)
        try:
            try: r = self.client.chat.completions.create(**kw)
            except Exception as e:
                if "reasoning" not in str(e).lower(): raise
                kw.pop("reasoning_effort", None); kw.pop("include_reasoning", None)
                r = self.client.chat.completions.create(**kw)
            return re.sub(r"<think>.*?</think>", "", r.choices[0].message.content or "", flags=re.S).strip()
        except Exception:
            if raise_err: raise
            return None


class Doc:
    """單一使用者（瀏覽器工作階段）目前載入的文件。"""

    def __init__(self):
        self.sid, self.name, self.chunks, self.labels = None, "", [], []
        self.n_child, self.tfidf, self.tm = 0, None, None
        self.lock, self.last = threading.Lock(), time.time()


class RAG:
    def __init__(self):
        self.ready, self.error, self.device = threading.Event(), None, "cpu"
        self.lock = threading.Lock()

    def warmup(self):  # 背景載入，讓服務先開始監聽埠
        try:
            import torch
            from sentence_transformers import SentenceTransformer
            self.device = "cuda" if torch.cuda.is_available() else "cpu"
            self.embedder = SentenceTransformer(EMBED_MODEL, device=self.device)
            self.embedder.max_seq_length = min(512, self.embedder.max_seq_length or 512)
            self.e5 = "e5" in EMBED_MODEL.lower()
            self.db = chromadb.EphemeralClient()
            slug = re.sub(r"[^a-zA-Z0-9]+", "_", EMBED_MODEL.split("/")[-1])[:40]
            self.col = self.db.get_or_create_collection(f"p_{slug}", metadata={"hnsw:space": "cosine"})
            self.ccol = self.db.get_or_create_collection(f"c_{slug}", metadata={"hnsw:space": "cosine"})
            self.ready.set()
        except Exception as e:
            self.error = f"{type(e).__name__}: {e}"

    # ── 文件載入 ──
    def drop(self, doc):
        if doc.sid:
            for c in (self.col, self.ccol):
                try: c.delete(where={"sid": doc.sid})
                except Exception: pass
        doc.sid, doc.chunks, doc.labels, doc.n_child, doc.tfidf, doc.tm = None, [], [], 0, None, None

    def load(self, doc, path, name):
        t0, path = time.time(), Path(path)
        ext = path.suffix.lower()
        if ext not in (".pdf", ".docx"): return "✗ 僅支援 PDF / DOCX"
        with doc.lock:
            try:
                segs = self._pdf(path) if ext == ".pdf" else self._docx(path)
                pairs = [(c, lab) for lab, txt in segs for c in split_text(txt, CHUNK, OVERLAP)]
                if not pairs: return "✗ 無可擷取文字（可能是掃描檔，需先 OCR）"
                self.drop(doc)
                sid = f"{int(time.time())}_{uuid.uuid4().hex[:6]}"
                chunks, labels = [p[0] for p in pairs], [p[1] for p in pairs]
                self._add(self.col, [f"{sid}_{i}" for i in range(len(pairs))], chunks,
                          [{"sid": sid, "idx": i} for i in range(len(pairs))])
                kids = [(f"{sid}_{p}_{c}", t, p) for p, par in enumerate(chunks)
                        for c, t in enumerate(split_text(par, CHILD, CHILD_OVERLAP))]
                self._add(self.ccol, [k[0] for k in kids], [k[1] for k in kids], [{"sid": sid, "idx": k[2]} for k in kids])
                tfidf = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4), sublinear_tf=True, max_features=50000)
                doc.tm, doc.tfidf = tfidf.fit_transform(chunks), tfidf
                doc.sid, doc.name, doc.chunks, doc.labels, doc.n_child = sid, name or path.name, chunks, labels, len(kids)
                return f"✓ {doc.name}\n{len(chunks)} 片段 · {time.time() - t0:.1f}s"
            except Exception as e:
                return f"✗ 載入失敗：{type(e).__name__}: {e}"

    def _pdf(self, path):
        return [(f"第 {i} 頁", t) for i, p in enumerate(PdfReader(str(path)).pages, 1) if (t := p.extract_text() or "").strip()]

    def _docx(self, path):
        doc, segs, head, buf, nt = Document(str(path)), [], "前言", [], 0
        flush = lambda: (segs.append((head, "\n".join(buf))), buf.clear()) if buf else None
        for el in doc.element.body.iterchildren():
            if isinstance(el, CT_P):
                para = Paragraph(el, doc); txt = para.text.strip()
                if not txt: continue
                style = (para.style.name if para.style is not None else "") or ""
                if style.lower().startswith(("heading", "title")) or style.startswith(("標題", "标题")):
                    flush(); head = txt[:40]
                buf.append(txt)
            elif isinstance(el, CT_Tbl):
                nt += 1; rows = []
                for row in Table(el, doc).rows:
                    cells = []
                    for c in row.cells:
                        t = re.sub(r"\s+", " ", c.text).strip()
                        if t and (not cells or cells[-1] != t): cells.append(t)
                    if cells: rows.append(" | ".join(cells))
                if rows: flush(); segs.append((f"表格 {nt}（{head}）", "\n".join(rows)))
        flush()
        return segs

    # ── 向量與檢索 ──
    def _enc(self, texts, query=False):
        if self.e5: texts = [("query: " if query else "passage: ") + t for t in texts]
        with self.lock:
            return self.embedder.encode(texts, batch_size=64, normalize_embeddings=True, show_progress_bar=False).tolist()

    def _add(self, col, ids, docs, metas, b=2000):
        for i in range(0, len(docs), b):
            col.add(ids=ids[i:i + b], documents=docs[i:i + b], metadatas=metas[i:i + b], embeddings=self._enc(docs[i:i + b]))

    def vec(self, doc, qs, k, child=False):
        total = doc.n_child if child else len(doc.chunks)
        if not total: return [[] for _ in qs]
        res = (self.ccol if child else self.col).query(query_embeddings=self._enc(qs, True), n_results=max(1, min(k, total)),
                                                       where={"sid": doc.sid}, include=["metadatas"])
        return [[int(m["idx"]) for m in ms] for ms in res["metadatas"]]

    def kw(self, doc, q, k):
        s = (doc.tm @ doc.tfidf.transform([q]).T).toarray().ravel()
        return [int(i) for i in np.argsort(-s)[:k] if s[i] > 0]

    @staticmethod
    def rrf(lists, k):
        sc = {}
        for lst in lists:
            for r, i in enumerate(dict.fromkeys(lst)): sc[i] = sc.get(i, 0) + 1 / (61 + r)
        return sorted(sc, key=sc.get, reverse=True)[:k]

    def retrieve(self, doc, llm, q, strategy, k):  # 回傳 [(片段索引, 文字)]
        hybrid = lambda n: self.rrf([self.vec(doc, [q], n * 2)[0], self.kw(doc, q, n * 2)], n)
        if strategy == "semantic": idx = self.vec(doc, [q], k)[0]
        elif strategy == "tfidf": idx = self.kw(doc, q, k)
        elif strategy == "parent_child": idx = list(dict.fromkeys(self.vec(doc, [q], k * 4, True)[0]))[:k]
        elif strategy == "multi_query":
            r = llm(f"將以下問題改寫成 3 個角度不同的繁體中文問題，每行一題，不加編號：\n{q}", 200, 0.7) or ""
            extra = [e for e in (l.strip("-•*·0123456789.、)） ") for l in r.splitlines()) if len(e) > 3][:3]
            idx = self.rrf(self.vec(doc, [q] + extra, k * 2), k)
        elif strategy == "hyde":
            h = llm(f"請針對以下問題寫一段約 100 字、語氣像文件原文的假設性答案（繁體中文）：\n{q}", 250, 0.7)
            idx = self.rrf(self.vec(doc, [q, h] if h else [q], k * 2), k)
        elif strategy == "rerank":
            cands = hybrid(min(max(k * 3, 6), 12)); idx = cands[:k]
            if llm.client and len(cands) > 1:
                lst = "\n\n".join(f"[{n}] {doc.chunks[i][:400]}" for n, i in enumerate(cands, 1))
                r = llm(f"問題：{q}\n\n候選片段：\n{lst}\n\n評估每段與問題的相關度（0-10），只輸出「編號:分數」每行一個。",
                        15 * len(cands), 0) or ""
                sc = {int(a): float(b) for a, b in re.findall(r"\[?(\d+)\]?\s*[:：]\s*(\d+(?:\.\d+)?)", r)}
                idx = [cands[n] for n in sorted(range(len(cands)), key=lambda n: (-sc.get(n + 1, -1), n))[:k]]
        else: idx = hybrid(k)
        hits = [(i, doc.chunks[i]) for i in idx]
        if strategy == "compress" and llm.client and hits:
            def comp(h):
                r = llm(f"從以下文本逐字擷取與問題「{q}」最相關的 1-3 句，不改寫不解釋；完全無關請輸出「無」。\n\n{h[1]}", 250, 0)
                return (h[0], r if r and r != "無" else h[1][:300])
            with ThreadPoolExecutor(min(4, len(hits))) as ex: hits = list(ex.map(comp, hits))
        return hits

    def answer(self, doc, llm, q, strategy, k, model_name=None):
        doc.last = time.time()
        meta = {"strategy": STRATEGIES[strategy][0], "model": llm.model, "doc": doc.name}
        if not doc.chunks:
            return {"answer": "請先上傳 PDF 或 DOCX 文件。", "sources": [], "meta": meta, "state": "no_doc"}
        t0 = time.time()
        hits = self.retrieve(doc, llm, q, strategy, int(k)); t1 = time.time()
        sources, used = [], 0
        for n, (i, txt) in enumerate(hits, 1):
            b = f"[片段{n}｜{doc.labels[i]}]\n{txt}"
            if used and used + len(b) > MAX_CTX: break
            sources.append({"n": n, "label": doc.labels[i], "text": txt}); used += len(b)
        meta["retrieve_s"] = round(t1 - t0, 2)
        if not hits:
            return {"answer": "找不到相關片段，請換個問法或策略。", "sources": [], "meta": meta, "state": "empty"}
        if not llm.client:
            return {"answer": "⚠ 尚未設定 Groq API Key（檢索結果請看下方片段）", "sources": sources, "meta": meta, "state": "no_key"}
        ctx = "\n\n".join(f"[片段{s['n']}｜{s['label']}]\n{s['text']}" for s in sources)
        try:
            ans = llm(f"請只根據上下文回答。若上下文沒有相關資訊，請說「文件中未提及」，不要編造。\n"
                      f"每個重點後以 [片段n] 標註依據，使用繁體中文條列整理。\n\n上下文：\n{ctx}\n\n問題：{q}",
                      1024, 0.2, "你是嚴謹的文件問答助手，只依據提供的上下文作答。", True)
        except Exception as e:
            return {"answer": f"✗ 生成失敗：{type(e).__name__}: {e}", "sources": sources, "meta": meta, "state": "error"}
        meta["generate_s"] = round(time.time() - t1, 2)
        return {"answer": ans or "（模型未回傳內容）", "sources": sources, "meta": meta, "state": "ok"}


# ── 匯出：Word / CSV ──
FIELDS = ["時間", "文件", "問題", "回答", "策略", "模型", "參考片段"]


def _out_path(ext):
    return os.path.join(tempfile.gettempdir(), f"RAG_{datetime.now():%Y%m%d_%H%M%S}_{uuid.uuid4().hex[:4]}.{ext}")


def export_csv(records):
    path = _out_path("csv")
    with open(path, "w", newline="", encoding="utf-8-sig") as f:  # utf-8-sig：Excel 開啟中文不亂碼
        w = csv.DictWriter(f, fieldnames=FIELDS, restval="", extrasaction="ignore"); w.writeheader(); w.writerows(records)
    return path


def _md_runs(par, text):  # 將 **粗體** 轉成 Word 粗體
    for i, part in enumerate(re.split(r"\*\*(.+?)\*\*", text)):
        if part: par.add_run(part).bold = bool(i % 2)


def export_docx(records, with_src=True):
    g = lambda r, k: str(r.get(k, "") or "")
    doc = Document()
    st = doc.styles["Normal"]; st.font.name, st.font.size = "Microsoft JhengHei", Pt(11)
    st.element.rPr.rFonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia", "Microsoft JhengHei")
    doc.add_heading("RAG 文件問答紀錄", 0)
    meta = doc.add_paragraph(f"文件：{g(records[-1], '文件')}　·　匯出時間：{datetime.now():%Y-%m-%d %H:%M}　·　共 {len(records)} 題")
    meta.runs[0].font.color.rgb = RGBColor(0x6B, 0x72, 0x80)
    for n, r in enumerate(records, 1):
        doc.add_heading(f"Q{n}．{g(r, '問題')}", level=1)
        info = doc.add_paragraph(f"{g(r, '時間')}　·　{g(r, '策略')}　·　{g(r, '模型')}")
        info.runs[0].font.size, info.runs[0].font.color.rgb = Pt(9), RGBColor(0x6B, 0x72, 0x80)
        for line in g(r, "回答").splitlines():
            line = line.rstrip()
            if not line.strip(): continue
            if m := re.match(r"^\s*#{1,6}\s+(.*)", line): doc.add_heading(m.group(1).replace("**", ""), level=2); continue
            if m := re.match(r"^(\s*)[-*•]\s+(.*)", line):
                par = doc.add_paragraph(style="List Bullet 2" if len(m.group(1)) >= 2 else "List Bullet"); _md_runs(par, m.group(2))
            elif m := re.match(r"^\s*\d+[.)]\s+(.*)", line):
                par = doc.add_paragraph(style="List Number"); _md_runs(par, m.group(1))
            else: _md_runs(doc.add_paragraph(), line)
        if with_src and g(r, "參考片段"):
            h = doc.add_paragraph(); run = h.add_run("參考片段"); run.bold, run.font.size = True, Pt(9)
            for blk in g(r, "參考片段").split("\n\n"):
                p = doc.add_paragraph(blk.strip()); p.paragraph_format.left_indent = Pt(12)
                for run in p.runs: run.font.size, run.font.color.rgb = Pt(8.5), RGBColor(0x55, 0x5B, 0x66)
    path = _out_path("docx"); doc.save(path)
    return path
