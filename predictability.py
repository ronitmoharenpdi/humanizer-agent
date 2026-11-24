"""
predictability.py - scores how predictable every sentence is, the way AI detectors roughly do.

A small free model (GPT-2, quantized ONNX version, ~127 MB) reads the text and, for every word, checks how expected
it was. Very predictable sentences look AI-written. The Humanizer uses this to find the sentences that still sound
like AI and sends only those back to Gemini.

* Runs on Vercel: onnxruntime + tokenizers + numpy (no PyTorch). The model is downloaded once per server start into
  /tmp from Hugging Face (PRED_MODEL_URL / PRED_TOKENIZER_URL), then reused.
* Fail-safe: if anything goes wrong (download, memory, library), available() returns False and the Humanizer simply
  skips this step. It never breaks a rewrite.
* Set PREDICTABILITY=off in Vercel to turn it off.
"""
from __future__ import annotations

import math
import os
import re
import threading
import urllib.request

MODEL_URL = os.getenv("PRED_MODEL_URL",
                      "https://huggingface.co/Xenova/gpt2/resolve/main/onnx/decoder_model_quantized.onnx")
TOKENIZER_URL = os.getenv("PRED_TOKENIZER_URL", "https://huggingface.co/Xenova/gpt2/resolve/main/tokenizer.json")
CACHE_DIR = os.getenv("PRED_CACHE_DIR", "/tmp/predictability")
MAX_TOKENS = 1024

# A sentence "looks AI-written" when it is very predictable (same cut-offs as predictability_check.py).
AI_PPL = float(os.getenv("PRED_AI_PPL", "25"))
AI_TOP10 = float(os.getenv("PRED_AI_TOP10", "0.80"))
MIN_TOKENS = 4          # shorter sentences can't be judged

_lock = threading.Lock()
_state = {"session": None, "tokenizer": None, "error": None}

ABBREV = re.compile(r"(?:\b[A-Za-z]\.[A-Za-z]\.|\b(?:e\.g|i\.e|etc|vs|Mr|Mrs|Ms|Dr|Prof|Rs|No|approx|Inc|Ltd|Co|St)\.)$",
                    re.I)


def enabled() -> bool:
    return os.getenv("PREDICTABILITY", "on").strip().lower() not in ("off", "0", "false", "no")


def _download(url: str, path: str) -> None:
    if os.path.exists(path) and os.path.getsize(path) > 0:
        return
    tmp = path + ".part"
    req = urllib.request.Request(url, headers={"User-Agent": "humanizer-agent"})
    with urllib.request.urlopen(req, timeout=120) as r, open(tmp, "wb") as f:
        while True:
            chunk = r.read(1 << 20)
            if not chunk:
                break
            f.write(chunk)
    os.replace(tmp, path)


def _load() -> bool:
    """Load the model once. Returns False (and remembers why) if it can't be used."""
    if _state["session"] is not None:
        return True
    if _state["error"]:
        return False
    with _lock:
        if _state["session"] is not None:
            return True
        try:
            import onnxruntime as ort
            from tokenizers import Tokenizer
            os.makedirs(CACHE_DIR, exist_ok=True)
            model_path = os.path.join(CACHE_DIR, "gpt2_quantized.onnx")
            tok_path = os.path.join(CACHE_DIR, "tokenizer.json")
            _download(TOKENIZER_URL, tok_path)
            _download(MODEL_URL, model_path)
            opts = ort.SessionOptions()
            opts.intra_op_num_threads = 2
            _state["session"] = ort.InferenceSession(model_path, sess_options=opts, providers=["CPUExecutionProvider"])
            _state["tokenizer"] = Tokenizer.from_file(tok_path)
            return True
        except Exception as exc:  # noqa: BLE001 - any failure means "skip this step"
            _state["error"] = f"{type(exc).__name__}: {exc}"[:300]
            return False


def available() -> bool:
    return enabled() and _load()


def status() -> dict:
    return {"enabled": enabled(), "loaded": _state["session"] is not None, "error": _state["error"]}


def _token_scores(text: str) -> list:
    """[(char_start, logprob, rank)] for every token except the first."""
    import numpy as np
    sess, tok = _state["session"], _state["tokenizer"]
    enc = tok.encode(text)
    ids, offsets = enc.ids, enc.offsets
    if len(ids) < 2:
        return []
    names = {i.name for i in sess.get_inputs()}
    out, start, done_upto = [], 0, 1
    stride = MAX_TOKENS // 2
    while True:
        end = min(start + MAX_TOKENS, len(ids))
        win = np.array([ids[start:end]], dtype=np.int64)
        feed = {"input_ids": win}
        if "attention_mask" in names:
            feed["attention_mask"] = np.ones_like(win)
        if "position_ids" in names:
            feed["position_ids"] = np.arange(win.shape[1], dtype=np.int64)[None, :]
        logits = sess.run(None, feed)[0][0]                              # (len, vocab), float32
        for pos in range(max(done_upto, start + 1), end):
            row = logits[pos - start - 1]
            mx = float(row.max())
            lse = mx + math.log(float(np.exp(row - mx).sum()))
            val = float(row[ids[pos]])
            out.append((offsets[pos][0], val - lse, int((row > val).sum())))
        done_upto = end
        if end >= len(ids):
            return out
        start = end - stride


def sentence_spans(text: str) -> list:
    """(start, end) of each sentence; 'a.m.', 'e.g.', 'Dr.', 'Rs.' don't end a sentence."""
    spans = []
    for m in re.finditer(r"[^\n]+", text):
        base, line = m.start(), m.group(0)
        start = 0
        for end in re.finditer(r"[.!?]+(?=\s|$)", line):
            if end.end() < len(line.rstrip()) and ABBREV.search(line[: end.end()]):
                continue
            spans.append((base + start, base + end.end()))
            start = end.end()
        if line[start:].strip():
            spans.append((base + start, base + len(line)))
    out = []
    for a, b in spans:
        seg = text[a:b]
        if re.search(r"\w", seg):
            out.append((a + len(seg) - len(seg.lstrip()), b))
    return out


def summarise(text: str, tokens: list) -> list:
    """Per sentence: {"start", "end", "text", "ppl", "top10", "ai"} from token scores."""
    rows = []
    for a, b in sentence_spans(text):
        toks = [t for t in tokens if a <= t[0] < b]
        if len(toks) < MIN_TOKENS:
            continue
        ppl = math.exp(-sum(t[1] for t in toks) / len(toks))
        top10 = sum(1 for t in toks if t[2] < 10) / len(toks)
        rows.append({"start": a, "end": b, "text": text[a:b].strip(), "ppl": round(ppl, 1), "top10": round(top10, 2),
                     "ai": ppl < AI_PPL or top10 >= AI_TOP10})
    return rows


def score_sentences(text: str) -> list:
    """Scores every sentence of a paragraph. Returns [] when the model isn't available."""
    if not text.strip() or not available():
        return []
    try:
        return summarise(text, _token_scores(text))
    except Exception as exc:  # noqa: BLE001
        _state["error"] = f"scoring failed: {type(exc).__name__}: {exc}"[:300]
        return []


def paragraph_score(rows: list) -> float:
    """Average predictability of a paragraph (higher = less predictable = more human-like)."""
    return sum(r["ppl"] for r in rows) / len(rows) if rows else 0.0
