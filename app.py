"""
Humanizer agent - FastAPI app (runs on Vercel or locally with `python app.py`).

GET  /               the tool page
GET  /api/health     is Gemini configured
POST /api/prepare    multipart: source=paste|docx|gdoc, text | file | gdoc_url
                     -> blocks, chunks, "before" scores
POST /api/rewrite    JSON: {heading, paras:[{id,text,tokens}], settings} -> rewritten paragraphs (one chunk)
POST /api/score      JSON: {text} -> scores
POST /api/download   multipart: source, file | gdoc_url | blocks, final (JSON {id: text}) -> .docx

The browser sends one /api/rewrite call per section (a few in parallel), so no single request runs long.
"""
from __future__ import annotations

import json
import os
import re
from pathlib import Path

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, Response

import humanizer as hz
from portal_gate import PortalGateMiddleware

try:  # local runs; on Vercel the variables come from Project Settings
    from dotenv import load_dotenv
    load_dotenv()
except Exception:  # pragma: no cover
    pass

BASE = Path(__file__).resolve().parent
MAX_UPLOAD = 4_000_000  # Vercel request body limit is ~4.5 MB
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

app = FastAPI(title="Humanizer", docs_url=None, redoc_url=None, openapi_url=None)
app.add_middleware(PortalGateMiddleware)


def gemini_key() -> str:
    return os.getenv("GEMINI_API_KEY", "").strip()


def gemini_model() -> str:
    return os.getenv("GEMINI_MODEL", "").strip() or "gemini-3.6-flash"


def err(status: int, message: str) -> JSONResponse:
    return JSONResponse({"error": message}, status_code=status)


@app.get("/", response_class=HTMLResponse)
def page():
    html = (BASE / "static" / "index.html").read_text(encoding="utf-8")
    url = os.getenv("PORTAL_URL", "").strip().rstrip("/")
    html = html.replace("{{LOGO_HREF}}", f"{url}/dashboard" if url else "/")
    html = html.replace("{{PORTAL_URL}}", url).replace("{{BACK_HIDDEN}}", "" if url else "hidden")
    return HTMLResponse(html, headers={"Cache-Control": "no-store"})


@app.get("/api/health")
def health():
    return {"ok": True, "gemini": bool(gemini_key()), "model": gemini_model(), "max_words": hz.MAX_INPUT_WORDS}


async def _source_bytes(source: str, file: UploadFile | None, gdoc_url: str) -> bytes:
    if source == "docx":
        if file is None:
            raise hz.HumanizerError("Choose a Word (.docx) file first.")
        if not (file.filename or "").lower().endswith(".docx"):
            raise hz.HumanizerError("Only .docx files are supported. In Word use File → Save As → .docx.")
        data = await file.read()
        if len(data) > MAX_UPLOAD:
            raise hz.HumanizerError("This file is larger than 4 MB (usually because of images). "
                                    "Remove the images or paste the text instead.")
        return data
    if source == "gdoc":
        return hz.fetch_google_doc(gdoc_url)
    raise hz.HumanizerError("Unknown input type.")


@app.post("/api/prepare")
async def prepare(source: str = Form("paste"), text: str = Form(""), gdoc_url: str = Form(""),
                  file: UploadFile | None = File(None)):
    try:
        if source == "paste":
            blocks = hz.read_pasted(text)
        else:
            blocks = hz.read_docx(await _source_bytes(source, file, gdoc_url))
    except hz.HumanizerError as exc:
        return err(400, str(exc))
    chunks = hz.plan_chunks(blocks)
    body = " ".join(b.plain() for b in blocks if b.kind == "para")
    return {
        "source": source,
        "blocks": [b.to_dict() for b in blocks],
        "chunks": chunks,
        "words": hz.count_words(body),
        "scores": hz.score_text(body),
        "ai_phrases": hz.AI_PHRASES,
        "gemini": bool(gemini_key()),
    }


@app.post("/api/rewrite")
async def rewrite(request: Request):
    if not gemini_key():
        return err(503, "GEMINI_API_KEY is not set. Add it in Vercel → Settings → Environment Variables.")
    try:
        data = await request.json()
        paras = data.get("paras") or []
        if not paras or len(paras) > 12:
            return err(400, "Send between 1 and 12 paragraphs per request.")
        texts = [str(p.get("text") or "") for p in paras]
        tokens = [{str(k): str(v) for k, v in (p.get("tokens") or {}).items()} for p in paras]
        if sum(hz.count_words(hz.expand_tokens(t, k)) for t, k in zip(texts, tokens)) > 900:
            return err(400, "This section is too long for one request.")
        settings = hz.Settings.from_dict(data.get("settings") or {})
        res = hz.rewrite_chunk(texts, tokens, settings, str(data.get("heading") or "")[:200],
                               gemini_key(), gemini_model())
    except hz.HumanizerError as exc:
        return err(502, str(exc))
    except (ValueError, TypeError, AttributeError):
        return err(400, "Bad request.")
    res["ids"] = [p.get("id") for p in paras]
    res["cost_usd"] = hz.estimate_cost_usd(res["usage"]["input_tokens"], res["usage"]["output_tokens"])
    return res


@app.post("/api/score")
async def score(request: Request):
    try:
        data = await request.json()
    except ValueError:
        return err(400, "Bad request.")
    return hz.score_text(str(data.get("text") or "")[:200_000])


@app.post("/api/download")
async def download(source: str = Form("paste"), final: str = Form("{}"), blocks: str = Form("[]"),
                   gdoc_url: str = Form(""), name: str = Form("humanized"), file: UploadFile | None = File(None)):
    try:
        final_by_id = {str(k): str(v) for k, v in json.loads(final or "{}").items()}
        if source == "paste":
            data = hz.build_docx_from_blocks(json.loads(blocks or "[]"), final_by_id)
        else:
            data = hz.build_docx(await _source_bytes(source, file, gdoc_url), final_by_id)
    except hz.HumanizerError as exc:
        return err(400, str(exc))
    except (ValueError, TypeError):
        return err(400, "Bad request.")
    base = re.sub(r"[^A-Za-z0-9 _-]+", "", Path(name).stem).strip()[:60] or "humanized"
    return Response(data, media_type=DOCX_MIME,
                    headers={"Content-Disposition": f'attachment; filename="{base} - humanized.docx"',
                             "Cache-Control": "no-store"})


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=int(os.getenv("PORT", "8010")))
