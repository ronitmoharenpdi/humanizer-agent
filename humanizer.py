"""
humanizer.py - core engine of the Humanizer agent (no web code here).

Pipeline
  1. read_input()      pasted text / .docx / Google Doc  ->  list of blocks (heading | para)
                       Links and bold phrases inside a paragraph become protected tokens ⟦1⟧, ⟦2⟧ …
                       so the rewrite can move them but never change them.
  2. plan_chunks()     group body paragraphs into ~350-word chunks (one AI call each)
  3. rewrite_chunk()   Gemini rewrites a chunk under strict rules, then check_chunk() verifies
                       tokens, keywords, numbers and length; one retry with the problems listed,
                       otherwise the original paragraph is kept and flagged.
  4. score_text()      readability and "AI feel" measures before / after (plain code, no AI)
  5. build_docx()      writes the chosen paragraphs back into the original Word file,
                       keeping headings, styles, links and bold phrases.
"""
from __future__ import annotations

import copy
import io
import json
import os
import re
import statistics
import urllib.error
import urllib.request
from dataclasses import dataclass, field

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --------------------------------------------------------------------------- #
# Settings
# --------------------------------------------------------------------------- #

GEMINI_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
CHUNK_WORDS = 350
MAX_INPUT_WORDS = int(os.getenv("MAX_INPUT_WORDS", "6000"))
LENGTH_RANGE = (0.85, 1.15)  # rewritten chunk must stay within ±15 % of the original

TONES = {
    "professional": "professional and clear, like an experienced industry writer",
    "conversational": "conversational and warm, like explaining to a smart friend",
    "friendly": "friendly and simple, easy for anyone to follow",
    "expert": "confident and expert, precise and authoritative without jargon overload",
}
STRENGTHS = {
    "light": "Light polish: keep most sentences; fix only robotic phrasing, clichés and monotone rhythm.",
    "balanced": "Balanced: rewrite sentences that sound generic or robotic; keep the structure of each paragraph.",
    "strong": "Strong: freely restructure sentences inside each paragraph for natural flow, keeping every point.",
}
SPELLING = {
    "us": "American English spelling (color, optimize)",
    "uk": "British English spelling (colour, optimise)",
    "in": "Indian English (British spelling, natural for Indian readers; use lakh/crore only if the original does)",
}

# Words and phrases that make text feel machine-written (kept short and specific on purpose).
AI_PHRASES = [
    "delve", "delves", "delving", "leverage", "leveraging", "seamless", "seamlessly", "robust", "elevate",
    "unlock", "unleash", "harness", "navigate the", "landscape", "realm", "tapestry", "testament to",
    "game-changer", "game changer", "cutting-edge", "in today's fast-paced", "in today's digital",
    "in the ever-evolving", "ever-evolving", "ever-changing", "it's important to note", "it is important to note",
    "it's worth noting", "it is worth noting", "in conclusion", "to sum up", "in summary", "furthermore",
    "moreover", "additionally", "embark", "journey", "crucial", "pivotal", "paramount", "myriad", "plethora",
    "foster", "bolster", "streamline", "empower", "revolutionize", "revolutionise", "transformative",
    "a wide range of", "plays a crucial role", "plays a vital role", "look no further", "whether you're",
    "dive into", "dive deep", "deep dive", "let's explore", "in this article", "in this blog", "comprehensive guide",
    "ultimately", "notably", "vibrant", "meticulous", "meticulously", "holistic", "synergy", "paradigm",
    "navigating", "intricacies", "nuanced", "underscore", "underscores", "boasts",
]
_PHRASE_RE = re.compile(
    r"(?<![\w-])(" + "|".join(re.escape(p) for p in sorted(AI_PHRASES, key=len, reverse=True)) + r")(?![\w-])",
    re.I,
)
TOKEN_RE = re.compile(r"⟦(\d+)⟧")
NUMBER_RE = re.compile(r"(?<![\w.])\d[\d,]*(?:\.\d+)?%?")
PASSIVE_RE = re.compile(r"\b(?:is|are|was|were|be|been|being)\s+(?:\w+ly\s+)?\w+(?:ed|en)\b", re.I)


class HumanizerError(Exception):
    """A message that can be shown to the user as is."""


# --------------------------------------------------------------------------- #
# 1. Reading input
# --------------------------------------------------------------------------- #

@dataclass
class Block:
    id: int
    kind: str            # "heading" | "para" | "skip" (kept as is: tables, lists of links, very short lines)
    text: str            # text with ⟦n⟧ tokens for protected pieces
    tokens: dict = field(default_factory=dict)   # "1" -> visible text of the protected piece
    level: int = 0       # heading level, 0 for body

    def plain(self) -> str:
        return expand_tokens(self.text, self.tokens)

    def to_dict(self) -> dict:
        return {"id": self.id, "kind": self.kind, "text": self.text, "tokens": self.tokens,
                "level": self.level, "plain": self.plain(), "words": count_words(self.plain())}


def expand_tokens(text: str, tokens: dict) -> str:
    return TOKEN_RE.sub(lambda m: tokens.get(m.group(1), ""), text)


def count_words(text: str) -> int:
    return len(re.findall(r"[A-Za-z0-9’']+", text or ""))


def _looks_like_heading(line: str) -> int:
    """Pasted text: '# Title' style or a short line without end punctuation. Returns level or 0."""
    m = re.match(r"^(#{1,6})\s+\S", line)
    if m:
        return len(m.group(1))
    words = line.split()
    if 0 < len(words) <= 12 and not re.search(r"[.!?:;,]$", line) and line[:1].isupper() and len(line) < 90:
        return 2
    return 0


def read_pasted(text: str) -> list:
    text = (text or "").replace("\r\n", "\n").strip()
    if not text:
        raise HumanizerError("Paste some text or upload a file first.")
    blocks, buf = [], []

    def flush():
        if buf:
            para = " ".join(s.strip() for s in buf).strip()
            if para:
                blocks.append(Block(len(blocks), "para", para))
            buf.clear()

    for raw in text.split("\n"):
        line = raw.strip()
        if not line:
            flush()
            continue
        level = _looks_like_heading(line)
        bullet = re.match(r"^([-*•]|\d+[.)])\s+", line)
        if level and not bullet:
            flush()
            blocks.append(Block(len(blocks), "heading", re.sub(r"^#{1,6}\s+", "", line), level=level))
        elif bullet:
            flush()
            blocks.append(Block(len(blocks), "para", line))
        else:
            buf.append(line)
    flush()
    return _finish(blocks)


def _finish(blocks: list) -> list:
    for b in blocks:  # very short lines (captions, labels) are not worth rewriting
        if b.kind == "para" and count_words(b.plain()) < 8:
            b.kind = "skip"
    words = sum(count_words(b.plain()) for b in blocks if b.kind == "para")
    if words == 0:
        raise HumanizerError("No paragraphs to rewrite were found (only headings or very short lines).")
    if words > MAX_INPUT_WORDS:
        raise HumanizerError(f"This text has {words:,} words. The limit is {MAX_INPUT_WORDS:,} per run; split it into parts.")
    return blocks


# ---- Word documents -------------------------------------------------------- #

W_P, W_R, W_T, W_HYPERLINK = qn("w:p"), qn("w:r"), qn("w:t"), qn("w:hyperlink")
W_RPR, W_PPR = qn("w:rPr"), qn("w:pPr")
SKIP_RUN_CHILDREN = {qn("w:drawing"), qn("w:object"), qn("w:pict"), qn("w:fldChar"), qn("w:instrText")}


def _heading_level(p, heading_ids: set) -> int:
    ppr = p.find(W_PPR)
    if ppr is None:
        return 0
    ps = ppr.find(qn("w:pStyle"))
    sid = (ps.get(qn("w:val")) or "") if ps is not None else ""
    if sid.lower() == "title":
        return 1
    if sid in heading_ids or sid.lower().startswith("heading"):
        m = re.search(r"(\d+)$", sid)
        return int(m.group(1)) if m else 2
    lvl = ppr.find(qn("w:outlineLvl"))
    if lvl is not None and (lvl.get(qn("w:val")) or "").isdigit():
        return int(lvl.get(qn("w:val"))) + 1
    return 0


def _heading_style_ids(doc) -> set:
    ids = set()
    for st in doc.styles:
        name = (getattr(st, "name", "") or "").lower()
        if name.startswith("heading") or name == "title":
            ids.add(st.style_id)
    return ids


def _run_text(r) -> str:
    out = []
    for child in r:
        if child.tag == W_T:
            out.append(child.text or "")
        elif child.tag == qn("w:tab"):
            out.append("\t")
        elif child.tag in (qn("w:br"), qn("w:cr")):
            out.append(" ")
    return "".join(out)


def _is_emphasis(r) -> bool:
    rpr = r.find(W_RPR)
    if rpr is None:
        return False
    for tag in ("w:b", "w:i", "w:u"):
        el = rpr.find(qn(tag))
        if el is not None and (el.get(qn("w:val")) or "true").lower() not in ("0", "false", "none"):
            return True
    return False


def _body_paragraphs(doc) -> list:
    """Top-level paragraphs in order (tables, text boxes, headers/footers are left untouched)."""
    return [el for el in doc.element.body.iterchildren() if el.tag == W_P]


def _para_blocks(p, idx: int, heading_ids: set) -> Block:
    level = _heading_level(p, heading_ids)
    tokens, parts, n = {}, [], 0
    special = False
    for child in p:
        if child.tag == W_HYPERLINK:
            text = "".join(_run_text(r) for r in child.iter(W_R))
            if text.strip():
                n += 1
                tokens[str(n)] = text
                parts.append(f"⟦{n}⟧")
        elif child.tag == W_R:
            if any(c.tag in SKIP_RUN_CHILDREN for c in child):
                special = True  # images, fields: keep this paragraph exactly as it is
            text = _run_text(child)
            if not text:
                continue
            if _is_emphasis(child) and text.strip():
                n += 1
                tokens[str(n)] = text
                parts.append(f"⟦{n}⟧")
            else:
                parts.append(text)
        elif child.tag not in (W_PPR, qn("w:bookmarkStart"), qn("w:bookmarkEnd"), qn("w:proofErr")):
            special = True  # content controls, inserted/deleted revisions, etc.
    text = re.sub(r"\s+", " ", "".join(parts)).strip()
    if level:
        return Block(idx, "heading", text, tokens, level)
    numbered = p.find(W_PPR) is not None and p.find(W_PPR).find(qn("w:numPr")) is not None
    kind = "skip" if (special or not text) else "para"
    b = Block(idx, kind, text, tokens)
    if numbered and kind == "para" and count_words(b.plain()) < 8:
        b.kind = "skip"
    return b


def read_docx(data: bytes) -> list:
    try:
        doc = Document(io.BytesIO(data))
    except Exception as exc:  # noqa: BLE001
        raise HumanizerError("This Word file could not be opened. Save it as .docx and try again.") from exc
    hids = _heading_style_ids(doc)
    blocks = [_para_blocks(p, i, hids) for i, p in enumerate(_body_paragraphs(doc))]
    return _finish(blocks)


GDOC_RE = re.compile(r"docs\.google\.com/document/d/([A-Za-z0-9_-]{10,})")
NOT_PUBLIC_MSG = ("Google didn't share this document. In Google Docs click Share → General access → "
                  "'Anyone with the link' (Viewer), then try again.")


def gdoc_id(url: str) -> str:
    m = GDOC_RE.search(url or "")
    if not m:
        raise HumanizerError("That doesn't look like a Google Doc link (docs.google.com/document/d/…).")
    return m.group(1)


def fetch_google_doc(url: str, timeout: int = 40) -> bytes:
    doc_id = gdoc_id(url)
    req = urllib.request.Request(f"https://docs.google.com/document/d/{doc_id}/export?format=docx",
                                 headers={"User-Agent": "Mozilla/5.0 (NP-Humanizer)"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = resp.read(20_000_001)
    except urllib.error.HTTPError as exc:
        raise HumanizerError(NOT_PUBLIC_MSG if exc.code in (401, 403, 404) else
                             f"Google returned an error ({exc.code}) for this document.") from exc
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        raise HumanizerError("Could not reach Google Docs. Try again in a moment.") from exc
    if not data.startswith(b"PK"):
        raise HumanizerError(NOT_PUBLIC_MSG)
    return data


# --------------------------------------------------------------------------- #
# 2. Chunks
# --------------------------------------------------------------------------- #

def plan_chunks(blocks: list, max_words: int = CHUNK_WORDS) -> list:
    """Groups of consecutive body paragraphs (never across a heading), ~max_words each.
    Each chunk: {"ids": [...], "heading": nearest heading above}."""
    chunks, cur, words, heading = [], [], 0, ""

    def close():
        nonlocal cur, words
        if cur:
            chunks.append({"ids": cur, "heading": heading})
        cur, words = [], 0

    for b in blocks:
        if b.kind == "heading":
            close()
            heading = b.plain()
        elif b.kind == "para":
            w = count_words(b.plain())
            if cur and words + w > max_words:
                close()
            cur.append(b.id)
            words += w
    close()
    return chunks


# --------------------------------------------------------------------------- #
# 3. Rewrite with Gemini + checks
# --------------------------------------------------------------------------- #

@dataclass
class Settings:
    tone: str = "professional"
    strength: str = "balanced"
    spelling: str = "in"
    audience: str = ""
    keywords: list = field(default_factory=list)
    voice_sample: str = ""

    @classmethod
    def from_dict(cls, d: dict) -> "Settings":
        d = d or {}
        kws = d.get("keywords") or []
        if isinstance(kws, str):
            kws = [k for k in re.split(r"[,\n]", kws)]
        kws = [re.sub(r"\s+", " ", k).strip() for k in kws if k and k.strip()][:30]
        return cls(
            tone=d.get("tone") if d.get("tone") in TONES else "professional",
            strength=d.get("strength") if d.get("strength") in STRENGTHS else "balanced",
            spelling=d.get("spelling") if d.get("spelling") in SPELLING else "in",
            audience=str(d.get("audience") or "")[:200].strip(),
            keywords=kws,
            voice_sample=str(d.get("voice_sample") or "")[:3000].strip(),
        )


SYSTEM_PROMPT = """You are a senior human editor at a content agency. You rewrite AI-sounding draft paragraphs so they read like a skilled human writer wrote them, for real readers. You improve writing quality only: you never add new facts, statistics, quotes, names, examples that claim to be real, or claims that are not already in the text.

How good human writing sounds:
- Mix short and longer sentences. Vary how sentences start. Not every paragraph needs three parts.
- Plain, specific words. Say "use" not "leverage", "help" not "empower", "important" not "pivotal".
- Active voice where it reads better. Direct address ("you") is fine when it fits the tone.
- Natural transitions or none at all. Avoid stacking "Moreover / Furthermore / Additionally".
- No filler openers or closers ("In today's fast-paced world", "In conclusion", "It's important to note that", "Whether you're X or Y").
- No hype: delve, unlock, elevate, seamless, robust, game-changer, cutting-edge, tapestry, realm, landscape, journey, embark, harness, navigate, testament, plethora, myriad.
- Avoid em dashes; use commas or full stops. Don't overuse colons or rhetorical questions.

Hard rules (breaking any of these makes the answer unusable):
1. Return exactly the same number of paragraphs, in the same order. Paragraph N of the answer replaces paragraph N of the input. Never merge, split, drop or add paragraphs.
2. Every placeholder like ⟦1⟧, ⟦2⟧ stands for a link or bold phrase. Keep each placeholder exactly once, unchanged, in the same paragraph. You may move it inside the sentence and change the words around it so the sentence still reads naturally with the placeholder's text in place.
3. Keep every number, percentage, date, price, name, brand and product exactly as written.
4. If a "must-keep phrase" appears in an input paragraph, it must appear in your version of that paragraph at least once, written exactly the same (any letter case). Repeating it less often is fine.
5. Keep each paragraph's meaning and every point it makes. Cutting filler may make it shorter; never pad it to look longer.
6. If a paragraph is a list item (starts with "-", "•" or "1."), keep that marker at the start.
7. Output only JSON: {"paragraphs": ["...", "..."]}."""


def build_prompt(paras: list, tokens_by_para: list, settings: Settings, heading: str, problems: str = "") -> str:
    lines = [
        f"Tone: {TONES[settings.tone]}.",
        f"How much to change: {STRENGTHS[settings.strength]}",
        f"Spelling: {SPELLING[settings.spelling]}.",
    ]
    if settings.audience:
        lines.append(f"Readers: {settings.audience}.")
    if settings.keywords:
        lines.append("Must-keep phrases (SEO keywords): " + "; ".join(f'"{k}"' for k in settings.keywords))
    if settings.voice_sample:
        lines.append("Match the voice of this sample from the brand's own writing (style only, don't copy its content):\n"
                     f'"""{settings.voice_sample}"""')
    if heading:
        lines.append(f'These paragraphs sit under the heading: "{heading}".')
    placeholder_notes = []
    for i, toks in enumerate(tokens_by_para, 1):
        for k, v in toks.items():
            placeholder_notes.append(f"paragraph {i}: ⟦{k}⟧ = \"{v}\"")
    if placeholder_notes:
        lines.append("Placeholders (keep each exactly once; the quoted text is what the reader will see there):\n"
                     + "\n".join(placeholder_notes))
    if problems:
        lines.append("Your previous answer broke these rules. Fix them:\n" + problems)
    lines.append(f"Rewrite these {len(paras)} paragraph(s). Return JSON with exactly {len(paras)} items:")
    lines.append(json.dumps({"paragraphs": paras}, ensure_ascii=False, indent=1))
    return "\n\n".join(lines)


def call_gemini(prompt: str, api_key: str, model: str, timeout: float = 55.0) -> tuple:
    """Returns (text, usage dict). Raises HumanizerError with a user-friendly message."""
    body = {
        "systemInstruction": {"parts": [{"text": SYSTEM_PROMPT}]},
        "contents": [{"role": "user", "parts": [{"text": prompt}]}],
        "generationConfig": {"responseMimeType": "application/json", "maxOutputTokens": 16384},
    }
    req = urllib.request.Request(
        GEMINI_ENDPOINT.format(model=model), data=json.dumps(body).encode("utf-8"),
        headers={"Content-Type": "application/json", "x-goog-api-key": api_key}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        try:
            msg = json.loads(exc.read().decode("utf-8")).get("error", {}).get("message", "")
        except Exception:  # noqa: BLE001
            msg = ""
        if exc.code in (401, 403):
            raise HumanizerError("Gemini rejected the API key. Check GEMINI_API_KEY in Vercel.") from exc
        if exc.code == 404:
            raise HumanizerError(f"Gemini model '{model}' isn't available for this key. Check GEMINI_MODEL.") from exc
        if exc.code == 429:
            raise HumanizerError("Gemini's rate limit or quota was reached. Wait a minute and try again.") from exc
        raise HumanizerError(f"Gemini error ({exc.code}){': ' + msg[:160] if msg else ''}. Please retry.") from exc
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        raise HumanizerError("Couldn't reach Gemini (timeout or network). Please retry.") from exc
    cand = (data.get("candidates") or [{}])[0]
    if cand.get("finishReason") in ("SAFETY", "RECITATION", "PROHIBITED_CONTENT", "BLOCKLIST", "SPII"):
        raise HumanizerError(f"Gemini declined this section ({cand.get('finishReason')}).")
    text = "".join(p.get("text", "") for p in ((cand.get("content") or {}).get("parts") or []) if not p.get("thought"))
    u = data.get("usageMetadata") or {}
    usage = {"input_tokens": int(u.get("promptTokenCount", 0) or 0),
             "output_tokens": int(u.get("candidatesTokenCount", 0) or 0) + int(u.get("thoughtsTokenCount", 0) or 0)}
    return text, usage


def parse_paragraphs(text: str, expected: int):
    try:
        obj = json.loads(text)
    except ValueError:
        m = re.search(r"\{.*\}", text or "", re.S)
        if not m:
            return None
        try:
            obj = json.loads(m.group(0))
        except ValueError:
            return None
    paras = obj.get("paragraphs") if isinstance(obj, dict) else obj
    if not isinstance(paras, list) or len(paras) != expected or not all(isinstance(p, str) for p in paras):
        return None
    return [re.sub(r"\s+", " ", p).strip() for p in paras]


def _numbers(text: str) -> list:
    return sorted(n.rstrip(".,") for n in NUMBER_RE.findall(text))


def check_paragraph(original: str, rewritten: str, tokens: dict, keywords: list) -> list:
    """Problems with one rewritten paragraph (empty list = OK). Works on text with ⟦n⟧ tokens."""
    problems = []
    if not rewritten.strip():
        return ["the paragraph is empty"]
    for k in tokens:
        c = rewritten.count(f"⟦{k}⟧")
        if c != 1:
            problems.append(f"placeholder ⟦{k}⟧ appears {c} times (must be exactly once)")
    extra = set(TOKEN_RE.findall(rewritten)) - set(tokens)
    if extra:
        problems.append("unknown placeholders " + ", ".join(f"⟦{e}⟧" for e in sorted(extra)))
    o_plain, r_plain = expand_tokens(original, tokens), expand_tokens(rewritten, tokens)
    missing_nums = [n for n in set(_numbers(o_plain)) if _numbers(r_plain).count(n) < _numbers(o_plain).count(n)]
    if missing_nums:
        problems.append("numbers changed or missing: " + ", ".join(missing_nums[:6]))
    for kw in keywords:
        pat = re.compile(r"(?<!\w)" + re.escape(kw) + r"(?!\w)", re.I)
        if pat.search(o_plain) and not pat.search(r_plain):
            problems.append(f'must-keep phrase "{kw}" is missing')
    ow, rw = count_words(o_plain), count_words(r_plain)
    if ow >= 25 and not (0.5 <= rw / max(ow, 1) <= 1.4):
        problems.append(f"length changed too much ({ow} → {rw} words)")
    return problems


def rewrite_chunk(paras: list, tokens_by_para: list, settings: Settings, heading: str,
                  api_key: str, model: str, caller=call_gemini) -> dict:
    """Rewrite one chunk. Returns {"paragraphs": [...], "status": [...], "notes": [...], "usage": {...}}.
    status per paragraph: "rewritten" | "kept" (original kept because the rewrite broke a rule)."""
    usage = {"input_tokens": 0, "output_tokens": 0, "ai_calls": 0}
    best = None
    problems_text = ""
    for attempt in range(2):
        text, u = caller(build_prompt(paras, tokens_by_para, settings, heading, problems_text), api_key, model)
        usage["input_tokens"] += u["input_tokens"]
        usage["output_tokens"] += u["output_tokens"]
        usage["ai_calls"] += 1
        out = parse_paragraphs(text, len(paras))
        if out is None:
            problems_text = f"- The answer must be JSON {{\"paragraphs\": [...]}} with exactly {len(paras)} strings."
            continue
        per = [check_paragraph(o, r, t, settings.keywords) for o, r, t in zip(paras, out, tokens_by_para)]
        if best is None or sum(map(len, per)) < sum(map(len, best[1])):
            best = (out, per)
        if not any(per):
            break
        problems_text = "\n".join(f"- paragraph {i}: {'; '.join(p)}" for i, p in enumerate(per, 1) if p)
    final, status, notes = [], [], []
    if best is None:
        return {"paragraphs": list(paras), "status": ["kept"] * len(paras),
                "notes": ["The AI answer couldn't be read, so this section was left unchanged."], "usage": usage}
    for i, (orig, new, probs) in enumerate(zip(paras, best[0], best[1])):
        if probs:
            final.append(orig)
            status.append("kept")
            notes.append(f"Kept the original of one paragraph: {probs[0]}.")
        else:
            final.append(new)
            status.append("rewritten")
    return {"paragraphs": final, "status": status, "notes": notes, "usage": usage}


# --------------------------------------------------------------------------- #
# 4. Scores (plain code, explainable)
# --------------------------------------------------------------------------- #

def _sentences(text: str) -> list:
    text = re.sub(r"\s+", " ", text or "").strip()
    parts = re.split(r"(?<![A-Z][a-z]\.)(?<!\b[A-Za-z]\.)(?<=[.!?])\s+(?=\S)", text)
    return [s for s in parts if count_words(s) > 0]


def _syllables(word: str) -> int:
    w = re.sub(r"[^a-z]", "", word.lower())
    if not w:
        return 0
    if len(w) <= 3:
        return 1
    w = re.sub(r"(?:[^laeiouy]es|ed|[^laeiouy]e)$", "", w)
    w = re.sub(r"^y", "", w)
    return max(1, len(re.findall(r"[aeiouy]{1,2}", w)))


def find_ai_phrases(text: str) -> list:
    return [m.group(0) for m in _PHRASE_RE.finditer(text or "")]


def score_text(text: str) -> dict:
    sents = _sentences(text)
    words = re.findall(r"[A-Za-z0-9’']+", text or "")
    n_w, n_s = len(words), max(len(sents), 1)
    syl = sum(_syllables(w) for w in words)
    lengths = [count_words(s) for s in sents] or [0]
    mean = statistics.mean(lengths)
    sd = statistics.pstdev(lengths) if len(lengths) > 1 else 0.0
    flesch = 206.835 - 1.015 * (n_w / n_s) - 84.6 * (syl / max(n_w, 1)) if n_w else 0.0
    phrases = find_ai_phrases(text)
    passive = len(PASSIVE_RE.findall(text or ""))
    per_100 = lambda x: round(100.0 * x / max(n_w, 1), 2)  # noqa: E731
    return {
        "words": n_w,
        "sentences": len(sents),
        "readability": round(max(0.0, min(100.0, flesch)), 1),                 # higher = easier
        "avg_sentence": round(mean, 1),
        "rhythm": round(sd / mean, 2) if mean else 0.0,                         # higher = more varied
        "ai_phrases": len(phrases),
        "ai_phrases_per_100": per_100(len(phrases)),
        "ai_phrase_list": sorted({p.lower() for p in phrases}),
        "passive_pct": round(100.0 * passive / n_s, 1),
        "dashes": (text or "").count("—") + len(re.findall(r"\s–\s", text or "")),
    }


def readability_label(score: float) -> str:
    return ("Very easy" if score >= 80 else "Easy" if score >= 70 else "Fairly easy" if score >= 60
            else "Standard" if score >= 50 else "Fairly difficult" if score >= 40 else "Difficult")


# --------------------------------------------------------------------------- #
# 5. Writing the result back into Word
# --------------------------------------------------------------------------- #

def _new_run(text: str, rpr):
    r = OxmlElement("w:r")
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    r.append(t)
    return r


def _rewrite_paragraph(p, new_text: str) -> None:
    """Replace the paragraph's text, keeping its style, its links and bold phrases (tokens)."""
    base_rpr, pieces, n = None, {}, 0
    for child in list(p):
        if child.tag == W_HYPERLINK:
            if "".join(_run_text(r) for r in child.iter(W_R)).strip():
                n += 1
                pieces[str(n)] = child
        elif child.tag == W_R:
            text = _run_text(child)
            if not text:
                continue
            if _is_emphasis(child) and text.strip():
                n += 1
                pieces[str(n)] = child
            elif base_rpr is None:
                base_rpr = child.find(W_RPR)
    keep_tags = {W_PPR}
    for child in list(p):
        if child.tag not in keep_tags:
            p.remove(child)
    pos = 0
    for m in TOKEN_RE.finditer(new_text):
        if m.start() > pos:
            p.append(_new_run(new_text[pos:m.start()], base_rpr))
        piece = pieces.get(m.group(1))
        if piece is not None:
            p.append(piece)
        pos = m.end()
    if pos < len(new_text):
        p.append(_new_run(new_text[pos:], base_rpr))


def build_docx(original: bytes, final_by_id: dict) -> bytes:
    """final_by_id: {block id: new text with tokens}. Other paragraphs stay exactly as they were."""
    doc = Document(io.BytesIO(original))
    paras = _body_paragraphs(doc)
    for bid, text in final_by_id.items():
        i = int(bid)
        if 0 <= i < len(paras) and isinstance(text, str) and text.strip():
            _rewrite_paragraph(paras[i], text)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def build_docx_from_blocks(blocks: list, final_by_id: dict) -> bytes:
    """For pasted text: a clean Word file with headings and paragraphs."""
    doc = Document()
    for b in blocks:
        text = expand_tokens(final_by_id.get(str(b["id"]), b["text"]), b.get("tokens") or {})
        if b["kind"] == "heading":
            doc.add_heading(text, level=min(max(int(b.get("level") or 2), 1), 4))
        else:
            doc.add_paragraph(text)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def estimate_cost_usd(input_tokens: int, output_tokens: int) -> float:
    pin = float(os.getenv("PRICE_INPUT_PER_M", "0.30") or 0.30)
    pout = float(os.getenv("PRICE_OUTPUT_PER_M", "2.50") or 2.50)
    return round(input_tokens / 1e6 * pin + output_tokens / 1e6 * pout, 4)

