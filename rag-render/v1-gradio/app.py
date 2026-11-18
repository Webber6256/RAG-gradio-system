# ═══ 多策略 RAG 文件問答系統 · Groq + ChromaDB + Gradio（Render / 本機部署版）═══
# 與 Colab 版的差異：移除 pip 安裝與 Colab Secrets；Key 讀環境變數 GROQ_API_KEY；
# 監聽 0.0.0.0:$PORT；不使用 share 連結。
import csv, os, re, tempfile, time, uuid
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
import chromadb, numpy as np, torch, gradio as gr
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.shared import Pt, RGBColor
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from groq import Groq
from pypdf import PdfReader
from sentence_transformers import SentenceTransformer
from sklearn.feature_extraction.text import TfidfVectorizer

# ── 設定 ──
EMBED_MODEL = os.getenv("EMBED_MODEL", "intfloat/multilingual-e5-small")   # 可換 multilingual-e5-base、BAAI/bge-m3
MODELS = ["openai/gpt-oss-20b", "openai/gpt-oss-120b", "qwen/qwen3.8-27b"]
CHUNK, OVERLAP, CHILD, CHILD_OVERLAP, MAX_CTX = 500, 100, 200, 40, 7000
DEVICE = "cuda" if torch.cuda.is_available() else "cpu"
GROQ_KEY = os.getenv("GROQ_API_KEY", "").strip()
CHROMA_PATH = os.getenv("CHROMA_PATH", "./chroma_db")

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


class RAG:
    def __init__(self):
        self.embedder = SentenceTransformer(EMBED_MODEL, device=DEVICE)
        self.embedder.max_seq_length = min(512, self.embedder.max_seq_length or 512)
        self.e5 = "e5" in EMBED_MODEL.lower()
        self.db = chromadb.PersistentClient(path=CHROMA_PATH)
        slug = re.sub(r"[^a-zA-Z0-9]+", "_", EMBED_MODEL.split("/")[-1])[:40]
        self.col = self.db.get_or_create_collection(f"p_{slug}", metadata={"hnsw:space": "cosine"})
        self.ccol = self.db.get_or_create_collection(f"c_{slug}", metadata={"hnsw:space": "cosine"})
        self.client, self.model, self.sid, self.name = None, MODELS[0], None, ""
        self.chunks, self.labels, self.n_child, self.tfidf, self.tm = [], [], 0, None, None

    # ── Groq ──
    def set_key(self, key):
        self.client = Groq(api_key=key.strip()) if key and key.strip() else None

    def list_models(self):
        if not self.client: return []
        try:
            ids = [m.id for m in self.client.models.list().data if getattr(m, "active", True) is not False]
        except Exception: return []
        return sorted(i for i in ids if not any(s in i.lower() for s in ("whisper", "tts", "orpheus", "guard", "playai")))

    def llm(self, prompt, max_tokens=300, temp=0.3, system=None, raise_err=False):
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

    # ── 文件載入 ──
    def load(self, path):
        t0, path = time.time(), Path(path)
        if path.suffix.lower() not in (".pdf", ".docx"): return "✗ 僅支援 PDF / DOCX"
        try:
            segs = self._pdf(path) if path.suffix.lower() == ".pdf" else self._docx(path)
            pairs = [(c, lab) for lab, txt in segs for c in split_text(txt, CHUNK, OVERLAP)]
            if not pairs: return "✗ 無可擷取文字（可能是掃描檔，需先 OCR）"
            if self.sid:
                for c in (self.col, self.ccol): c.delete(where={"sid": self.sid})
            self.sid, self.name = f"{int(time.time())}_{uuid.uuid4().hex[:6]}", path.name
            self.chunks, self.labels = [p[0] for p in pairs], [p[1] for p in pairs]
            self._add(self.col, [f"{self.sid}_{i}" for i in range(len(pairs))], self.chunks,
                      [{"sid": self.sid, "idx": i} for i in range(len(pairs))])
            kids = [(f"{self.sid}_{p}_{c}", t, p) for p, par in enumerate(self.chunks)
                    for c, t in enumerate(split_text(par, CHILD, CHILD_OVERLAP))]
            self.n_child = len(kids)
            self._add(self.ccol, [k[0] for k in kids], [k[1] for k in kids], [{"sid": self.sid, "idx": k[2]} for k in kids])
            self.tfidf = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4), sublinear_tf=True, max_features=50000)
            self.tm = self.tfidf.fit_transform(self.chunks)
            return f"✓ {self.name}\n{len(self.chunks)} 片段 · {time.time() - t0:.1f}s"
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
        return self.embedder.encode(texts, batch_size=64, normalize_embeddings=True, show_progress_bar=False).tolist()

    def _add(self, col, ids, docs, metas, b=2000):
        for i in range(0, len(docs), b):
            col.add(ids=ids[i:i + b], documents=docs[i:i + b], metadatas=metas[i:i + b], embeddings=self._enc(docs[i:i + b]))

    def vec(self, qs, k, child=False):
        total = self.n_child if child else len(self.chunks)
        if not total: return [[] for _ in qs]
        res = (self.ccol if child else self.col).query(query_embeddings=self._enc(qs, True), n_results=max(1, min(k, total)),
                                                       where={"sid": self.sid}, include=["metadatas"])
        return [[int(m["idx"]) for m in ms] for ms in res["metadatas"]]

    def kw(self, q, k):
        s = (self.tm @ self.tfidf.transform([q]).T).toarray().ravel()
        return [int(i) for i in np.argsort(-s)[:k] if s[i] > 0]

    @staticmethod
    def rrf(lists, k):
        sc = {}
        for lst in lists:
            for r, i in enumerate(dict.fromkeys(lst)): sc[i] = sc.get(i, 0) + 1 / (61 + r)
        return sorted(sc, key=sc.get, reverse=True)[:k]

    def retrieve(self, q, strategy, k):  # 回傳 [(片段索引, 文字)]
        hybrid = lambda n: self.rrf([self.vec([q], n * 2)[0], self.kw(q, n * 2)], n)
        if strategy == "semantic": idx = self.vec([q], k)[0]
        elif strategy == "tfidf": idx = self.kw(q, k)
        elif strategy == "parent_child": idx = list(dict.fromkeys(self.vec([q], k * 4, True)[0]))[:k]
        elif strategy == "multi_query":
            r = self.llm(f"將以下問題改寫成 3 個角度不同的繁體中文問題，每行一題，不加編號：\n{q}", 200, 0.7) or ""
            extra = [e for e in (l.strip("-•*·0123456789.、)） ") for l in r.splitlines()) if len(e) > 3][:3]
            idx = self.rrf(self.vec([q] + extra, k * 2), k)
        elif strategy == "hyde":
            h = self.llm(f"請針對以下問題寫一段約 100 字、語氣像文件原文的假設性答案（繁體中文）：\n{q}", 250, 0.7)
            idx = self.rrf(self.vec([q, h] if h else [q], k * 2), k)
        elif strategy == "rerank":
            cands = hybrid(min(max(k * 3, 6), 12)); idx = cands[:k]
            if self.client and len(cands) > 1:
                lst = "\n\n".join(f"[{n}] {self.chunks[i][:400]}" for n, i in enumerate(cands, 1))
                r = self.llm(f"問題：{q}\n\n候選片段：\n{lst}\n\n評估每段與問題的相關度（0-10），只輸出「編號:分數」每行一個。",
                             15 * len(cands), 0) or ""
                sc = {int(a): float(b) for a, b in re.findall(r"\[?(\d+)\]?\s*[:：]\s*(\d+(?:\.\d+)?)", r)}
                idx = [cands[n] for n in sorted(range(len(cands)), key=lambda n: (-sc.get(n + 1, -1), n))[:k]]
        else: idx = hybrid(k)
        hits = [(i, self.chunks[i]) for i in idx]
        if strategy == "compress" and self.client and hits:
            def comp(h):
                r = self.llm(f"從以下文本逐字擷取與問題「{q}」最相關的 1-3 句，不改寫不解釋；完全無關請輸出「無」。\n\n{h[1]}", 250, 0)
                return (h[0], r if r and r != "無" else h[1][:300])
            with ThreadPoolExecutor(min(4, len(hits))) as ex: hits = list(ex.map(comp, hits))
        return hits

    def answer(self, q, strategy, k, model):
        if not self.chunks: return "請先上傳 PDF 或 DOCX 文件。", ""
        self.model, t0 = model or MODELS[0], time.time()
        hits = self.retrieve(q, strategy, int(k)); t1 = time.time()
        blocks, used = [], 0
        for n, (i, txt) in enumerate(hits, 1):
            b = f"[片段{n}｜{self.labels[i]}]\n{txt}"
            if used and used + len(b) > MAX_CTX: break
            blocks.append(b); used += len(b)
        ctx = "\n\n".join(blocks)
        src = f"📄 {self.name} · {STRATEGIES[strategy][0]} · {len(blocks)} 片段 · 檢索 {t1 - t0:.2f}s\n{'─' * 40}\n\n{ctx}"
        if not hits: return "找不到相關片段，請換個問法或策略。", src
        if not self.client: return "⚠ 尚未設定 Groq API Key（檢索結果請看下方片段）", src
        try:
            ans = self.llm(f"請只根據上下文回答。若上下文沒有相關資訊，請說「文件中未提及」，不要編造。\n"
                           f"每個重點後以 [片段n] 標註依據，使用繁體中文條列整理。\n\n上下文：\n{ctx}\n\n問題：{q}",
                           1024, 0.2, "你是嚴謹的文件問答助手，只依據提供的上下文作答。", True)
        except Exception as e:
            return f"✗ 生成失敗：{type(e).__name__}: {e}", src
        return f"{ans or '（模型未回傳內容）'}\n\n<sub>{self.model} · 檢索 {t1 - t0:.1f}s · 生成 {time.time() - t1:.1f}s</sub>", src


# ── Gradio 介面 ──
print(f"載入嵌入模型（{DEVICE}）…")
rag = RAG(); rag.set_key(GROQ_KEY)
init_models = rag.list_models() or MODELS
GR6 = int(gr.__version__.split(".")[0]) >= 6
CSS = """
.gradio-container{max-width:1200px!important;margin:auto}
#hdr{padding:18px 4px 6px}#hdr h1{font-size:26px;margin:0;font-weight:700}
#hdr p{color:#6b7280;margin:4px 0 0;font-size:14px}
.side{border-radius:14px!important}
#send{min-width:90px}
footer{display:none!important}
"""
THEME = gr.themes.Soft(primary_hue="emerald", neutral_hue="slate", radius_size="lg",
                       font=[gr.themes.GoogleFont("Noto Sans TC"), "sans-serif"])


def on_key(key, cur):
    rag.set_key(key)
    if not (key or "").strip(): return "⚠ 未設定 Key", gr.update()
    ms = rag.list_models()
    if not ms: return "✗ Key 無效或無法取得模型", gr.update()
    return f"✓ Key 有效 · {len(ms)} 個模型", gr.update(choices=ms, value=cur if cur in ms else (MODELS[0] if MODELS[0] in ms else ms[0]))


def on_upload(f):
    return "⚠ 請選擇檔案" if f is None else rag.load(f if isinstance(f, str) else f.name)


def on_ask(q, history, records, strategy, k, model):
    q = (q or "").strip()
    if not q: return "", history, gr.update(), records
    ans, src = rag.answer(q, strategy, k, model)
    records = records + [{"時間": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "文件": rag.name, "問題": q,
                          "回答": re.sub(r"\n*<sub>.*?</sub>", "", ans, flags=re.S).strip(),
                          "策略": STRATEGIES[strategy][0], "模型": rag.model,
                          "參考片段": src.split("─" * 40, 1)[-1].strip()}]
    return "", history + [{"role": "user", "content": q}, {"role": "assistant", "content": ans}], src, records


# ── 匯出：Word / CSV ──
def _out_path(ext):
    return os.path.join(tempfile.gettempdir(), f"RAG問答紀錄_{datetime.now():%Y%m%d_%H%M%S}.{ext}")


def export_csv(records):
    if not records: gr.Warning("目前沒有問答紀錄可匯出"); return None
    path = _out_path("csv")
    with open(path, "w", newline="", encoding="utf-8-sig") as f:  # utf-8-sig：Excel 開啟中文不亂碼
        w = csv.DictWriter(f, fieldnames=list(records[0])); w.writeheader(); w.writerows(records)
    return path


def _md_runs(par, text):  # 將 **粗體** 轉成 Word 粗體
    for i, part in enumerate(re.split(r"\*\*(.+?)\*\*", text)):
        if part: par.add_run(part).bold = bool(i % 2)


def export_docx(records, with_src=True):
    if not records: gr.Warning("目前沒有問答紀錄可匯出"); return None
    doc = Document()
    st = doc.styles["Normal"]; st.font.name, st.font.size = "Microsoft JhengHei", Pt(11)
    st.element.rPr.rFonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia", "Microsoft JhengHei")
    doc.add_heading("RAG 文件問答紀錄", 0)
    meta = doc.add_paragraph(f"文件：{records[-1]['文件']}　·　匯出時間：{datetime.now():%Y-%m-%d %H:%M}　·　共 {len(records)} 題")
    meta.runs[0].font.color.rgb = RGBColor(0x6B, 0x72, 0x80)
    for n, r in enumerate(records, 1):
        doc.add_heading(f"Q{n}．{r['問題']}", level=1)
        info = doc.add_paragraph(f"{r['時間']}　·　{r['策略']}　·　{r['模型']}")
        info.runs[0].font.size, info.runs[0].font.color.rgb = Pt(9), RGBColor(0x6B, 0x72, 0x80)
        for line in r["回答"].splitlines():
            line = line.rstrip()
            if not line.strip(): continue
            if m := re.match(r"^\s*#{1,6}\s+(.*)", line): doc.add_heading(m.group(1).replace("**", ""), level=2); continue
            if m := re.match(r"^(\s*)[-*•]\s+(.*)", line):
                par = doc.add_paragraph(style="List Bullet 2" if len(m.group(1)) >= 2 else "List Bullet"); _md_runs(par, m.group(2))
            elif m := re.match(r"^\s*\d+[.)]\s+(.*)", line):
                par = doc.add_paragraph(style="List Number"); _md_runs(par, m.group(1))
            else: _md_runs(doc.add_paragraph(), line)
        if with_src and r["參考片段"]:
            h = doc.add_paragraph(); run = h.add_run("參考片段"); run.bold, run.font.size = True, Pt(9)
            for blk in r["參考片段"].split("\n\n"):
                p = doc.add_paragraph(blk.strip()); p.paragraph_format.left_indent = Pt(12)
                for run in p.runs: run.font.size, run.font.color.rgb = Pt(8.5), RGBColor(0x55, 0x5B, 0x66)
    path = _out_path("docx"); doc.save(path)
    return path


blocks_kw = {"title": "RAG 文件問答"} | ({} if GR6 else {"css": CSS, "theme": THEME})
with gr.Blocks(**blocks_kw) as demo:
    gr.HTML("<div id='hdr'><h1>📚 RAG 文件問答</h1><p>上傳 PDF / DOCX，選擇檢索策略，直接提問 · Groq × ChromaDB</p></div>")
    with gr.Row(equal_height=False):
        with gr.Column(scale=1, min_width=300, elem_classes="side"):
            with gr.Accordion("🔑 Groq API Key", open=not GROQ_KEY):
                key = gr.Textbox(value=GROQ_KEY, type="password", show_label=False, placeholder="gsk_…（Enter 套用）")
                key_status = gr.Markdown("✓ 已從環境變數 GROQ_API_KEY 載入" if GROQ_KEY else "⚠ 尚未設定")
            model = gr.Dropdown(init_models, value=MODELS[0] if MODELS[0] in init_models else init_models[0],
                                allow_custom_value=True, label="模型")
            file = gr.File(label="上傳文件", file_types=[".pdf", ".docx"], type="filepath", height=120)
            status = gr.Textbox(show_label=False, interactive=False, lines=2, placeholder="尚未載入文件")
            strategy = gr.Radio([(v[0], k) for k, v in STRATEGIES.items()], value="hybrid", label="檢索策略")
            desc = gr.Markdown(f"<small>{STRATEGIES['hybrid'][1]}</small>")
            topk = gr.Slider(1, 10, 4, step=1, label="片段數量 Top-K")
        with gr.Column(scale=2):
            chat = gr.Chatbot(height=520, show_label=False, **({} if GR6 else {"type": "messages"}),
                              placeholder="**上傳文件後開始提問**<br>例如：這份文件的重點是什麼？")
            with gr.Row():
                q = gr.Textbox(show_label=False, placeholder="輸入問題，Enter 送出", scale=6, container=False)
                send = gr.Button("送出", variant="primary", scale=1, elem_id="send")
            with gr.Row():
                gr.Examples([["這份文件的重點是什麼？"], ["文件提到哪些規定或限制？"], ["有哪些風險或注意事項？"]], q, label=None)
                clear = gr.Button("🗑 清除對話", size="sm", variant="secondary")
            with gr.Accordion("🔎 檢索到的片段", open=False):
                src = gr.Textbox(show_label=False, lines=12, interactive=False)
            with gr.Row():
                with_src = gr.Checkbox(True, label="Word 包含參考片段", scale=2)
                btn_docx = gr.Button("📄 匯出 Word", size="sm", scale=1)
                btn_csv = gr.Button("📊 匯出 CSV", size="sm", scale=1)
            out_file = gr.File(label="下載檔案", visible=False)
    records = gr.State([])

    key.submit(on_key, [key, model], [key_status, model]); key.blur(on_key, [key, model], [key_status, model])
    file.upload(on_upload, file, status)
    strategy.change(lambda s: f"<small>{STRATEGIES[s][1]}</small>", strategy, desc)
    for ev in (q.submit, send.click): ev(on_ask, [q, chat, records, strategy, topk, model], [q, chat, src, records])
    clear.click(lambda: ([], "", [], gr.update(value=None, visible=False)), None, [chat, src, records, out_file])
    show = lambda p: gr.update(value=p, visible=bool(p))
    btn_docx.click(lambda r, w: show(export_docx(r, w)), [records, with_src], out_file)
    btn_csv.click(lambda r: show(export_csv(r)), records, out_file)

if __name__ == "__main__":
    demo.queue().launch(server_name="0.0.0.0", server_port=int(os.getenv("PORT", "7860")), show_error=True,
                        **({"css": CSS, "theme": THEME} if GR6 else {}))
