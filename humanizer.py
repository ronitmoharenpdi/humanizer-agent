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
     style_check()     AI-style check (NLP): AI phrases, flat rhythm, repeated openers, long sentences, em dashes.
                       Guides the rewrite (hints before, one targeted retry after) and marks what's left.
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

import predictability as pred
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --------------------------------------------------------------------------- #
# Settings
# --------------------------------------------------------------------------- #

GEMINI_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
CHUNK_WORDS = 350
MAX_CHUNK_PARAS = 10   # the API accepts up to 12 paragraphs per request
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
            if cur and (words + w > max_words or len(cur) >= MAX_CHUNK_PARAS):
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


def build_prompt(paras: list, tokens_by_para: list, settings: Settings, heading: str, problems: str = "",
                 hints: str = "", style_fix: str = "") -> str:
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
    if hints:
        lines.append("Our AI-style checker found these problems in the input. Fix every one of them (a flagged word may stay "
                     "only if it is a must-keep phrase or the correct technical term, e.g. 'customer journey'):\n" + hints)
    if problems:
        lines.append("Your previous answer broke these rules. Fix them:\n" + problems)
    if style_fix:
        lines.append("Your previous answer still sounds AI-written in places. Our checker found the problems below in your "
                     "previous version of these paragraphs. Fix them, and keep everything else that was good:\n" + style_fix)
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
    """Rewrite one chunk. Returns {"paragraphs", "status", "notes", "usage", "style"}.
    status per paragraph: "rewritten" | "kept" (original kept because the rewrite broke a safety rule).

    1. The AI-style checker looks at the input and Gemini gets the exact list of problems to fix.
    2. Safety checks (placeholders, numbers, keywords, length). Broken → one retry; still broken → original kept.
    3. The AI-style checker looks at the rewrite. Problems left → one more targeted try (STYLE_RETRIES).
       For every paragraph the safe version with the fewest AI-style problems is used."""
    usage = {"input_tokens": 0, "output_tokens": 0, "ai_calls": 0}
    n = len(paras)
    kws = settings.keywords
    before = [style_check(p, kws) for p in paras]
    hints = style_hints(paras, kws)
    best = [None] * n            # (style count, text) of the best SAFE version per paragraph
    worst_safety = [None] * n    # fewest safety problems seen, for the "kept" note
    problems_text, style_fix = "", ""
    safety_tries, style_tries = 0, 0
    last = None
    while True:
        text, u = caller(build_prompt(paras, tokens_by_para, settings, heading, problems_text, hints, style_fix),
                         api_key, model)
        usage["input_tokens"] += u["input_tokens"]
        usage["output_tokens"] += u["output_tokens"]
        usage["ai_calls"] += 1
        out = parse_paragraphs(text, n)
        if out is None:
            problems_text, style_fix = f"- The answer must be JSON {{\"paragraphs\": [...]}} with exactly {n} strings.", ""
            safety_tries += 1
            if safety_tries > 1:
                break
            continue
        last = out
        per = [check_paragraph(o, r, t, kws) for o, r, t in zip(paras, out, tokens_by_para)]
        styles = [style_check(r, kws) for r in out]
        for i in range(n):
            if per[i]:
                if worst_safety[i] is None or len(per[i]) < len(worst_safety[i]):
                    worst_safety[i] = per[i]
            elif best[i] is None or styles[i]["count"] < best[i][0]:
                best[i] = (styles[i]["count"], out[i])
        unsafe = [i for i in range(n) if per[i] and best[i] is None]      # no safe version of these yet
        if unsafe and safety_tries < 1:
            safety_tries += 1
            problems_text = "\n".join(f"- paragraph {i + 1}: {'; '.join(per[i])}" for i in unsafe)
            style_fix = ""
            continue
        left = [i for i in range(n) if best[i] and best[i][0] > 0]
        if left and style_tries < STYLE_RETRIES:
            style_tries += 1
            problems_text = ""
            style_fix = "\n".join(f"- paragraph {i + 1} (your version: \"{best[i][1]}\"): "
                                   + "; ".join(style_check(best[i][1], kws)["problems"]) for i in left)
            continue
        break

    if last is None and not any(best):
        return {"paragraphs": list(paras), "status": ["kept"] * n,
                "notes": ["The AI answer couldn't be read, so this section was left unchanged."], "usage": usage,
                "style": [_style_out(b) for b in before], "style_before": sum(b["count"] for b in before)}
    pred_info = _predictability_pass(paras, tokens_by_para, settings, heading, best, api_key, model, caller, usage)

    final, status, notes, style = [], [], [], []
    for i in range(n):
        if best[i] is None:
            final.append(paras[i])
            status.append("kept")
            notes.append(f"Kept the original of one paragraph: {(worst_safety[i] or ['the rewrite broke a rule'])[0]}.")
            style.append(_style_out(before[i]))
        else:
            final.append(best[i][1])
            status.append("rewritten")
            style.append(_style_out(style_check(best[i][1], kws)))
    return {"paragraphs": final, "status": status, "notes": notes, "usage": usage,
            "style": style, "style_before": sum(b["count"] for b in before), "predictability": pred_info}


# --------------------------------------------------------------------------- #
# 3b. Predictability pass: sentences a small model (GPT-2) finds too predictable get rewritten
# --------------------------------------------------------------------------- #

PRED_ROUNDS = int(os.getenv("PRED_ROUNDS", "2"))   # extra AI calls per section at most

PRED_GUIDE = """Our detector (a language model that measures how predictable each word is) says the sentences below still
read as AI-written: every word is the most obvious next word. Rewrite ONLY those sentences so they are less predictable,
the way an experienced writer explaining this to a friend would put it:
- choose less obvious but natural words; drop stock phrases ("helps you", "the right", "it is important", "peace of
  mind", "in just a few minutes", "an excellent choice", "ensures", "a wide range");
- change the structure: lead with the point, use a short punchy sentence, a question, or a contrast ("not X, but Y");
- an everyday comparison or picture is welcome (e.g. "takes about as long as ordering lunch online"), but NEVER add new
  facts, numbers, names, statistics, quotes or claims;
- keep the meaning, every number, every placeholder ⟦n⟧ and every must-keep phrase exactly;
- leave all other sentences of the paragraph exactly as they are."""


def _predictability_pass(paras, tokens_by_para, settings, heading, best, api_key, model, caller, usage) -> dict:
    """Up to PRED_ROUNDS targeted rewrites of the sentences GPT-2 finds most predictable.
    A new version of a paragraph is used only if it passes the safety checks, has no more AI-style problems and is
    less predictable than before. Updates best[i] in place. Never raises: problems just skip this step."""
    info = {"used": False, "rounds": 0, "ai_sentences_before": 0, "ai_sentences_after": 0}
    if not pred.enabled():
        return info
    kws = settings.keywords
    scored = {}
    for i, b in enumerate(best):
        if b is not None:
            rows = pred.score_sentences(expand_tokens(b[1], tokens_by_para[i]))
            if rows:
                scored[i] = rows
    if not scored:
        info["error"] = pred.status().get("error")
        return info
    info["used"] = True
    info["ai_sentences_before"] = sum(r["ai"] for rows in scored.values() for r in rows)
    for _ in range(PRED_ROUNDS):
        targets = [i for i, rows in scored.items() if any(r["ai"] for r in rows)]
        if not targets:
            break
        info["rounds"] += 1
        current = [best[i][1] for i in targets]
        lines = [f"Tone: {TONES[settings.tone]}.", f"Spelling: {SPELLING[settings.spelling]}."]
        if kws:
            lines.append("Must-keep phrases (SEO keywords): " + "; ".join(f'"{k}"' for k in kws))
        if heading:
            lines.append(f'These paragraphs sit under the heading: "{heading}".')
        notes = [f"paragraph {k}: ⟦{t}⟧ = \"{v}\"" for k, i in enumerate(targets, 1) for t, v in tokens_by_para[i].items()]
        if notes:
            lines.append("Placeholders (keep each exactly once):\n" + "\n".join(notes))
        lines.append(PRED_GUIDE)
        flagged = []
        for k, i in enumerate(targets, 1):
            for r in scored[i]:
                if r["ai"]:
                    flagged.append(f'- paragraph {k}: "{r["text"]}"')
        lines.append("Sentences to rewrite (placeholders are shown as their visible text):\n" + "\n".join(flagged))
        lines.append(f"Return JSON with exactly {len(targets)} paragraphs, in this order:")
        lines.append(json.dumps({"paragraphs": current}, ensure_ascii=False, indent=1))
        try:
            text, u = caller("\n\n".join(lines), api_key, model)
        except HumanizerError:
            break
        usage["input_tokens"] += u["input_tokens"]
        usage["output_tokens"] += u["output_tokens"]
        usage["ai_calls"] += 1
        out = parse_paragraphs(text, len(targets))
        if out is None:
            continue
        for k, i in enumerate(targets):
            new = out[k]
            if check_paragraph(paras[i], new, tokens_by_para[i], kws):
                continue                                        # breaks a safety rule → keep the current version
            new_style = style_check(new, kws)["count"]
            if new_style > best[i][0]:
                continue                                        # brought AI phrases back
            rows = pred.score_sentences(expand_tokens(new, tokens_by_para[i]))
            if not rows:
                continue
            old_rows = scored[i]
            if (sum(r["ai"] for r in rows) < sum(r["ai"] for r in old_rows)
                    or (sum(r["ai"] for r in rows) == sum(r["ai"] for r in old_rows)
                        and pred.paragraph_score(rows) > pred.paragraph_score(old_rows))):
                best[i] = (new_style, new)
                scored[i] = rows
    info["ai_sentences_after"] = sum(r["ai"] for rows in scored.values() for r in rows)
    return info


def _style_out(c: dict) -> dict:
    return {"count": c["count"], "flat": c["flat"], "problems": c["problems"], "flags": c["flags"]}


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


# --------------------------------------------------------------------------- #
# 4b. AI-style check (NLP, plain code): finds what makes a paragraph sound AI-written,
#     tells Gemini exactly what to fix, and checks the rewrite again.
# --------------------------------------------------------------------------- #

FLAT_RHYTHM = 0.25     # sentence-length variation below this (std / mean) = "every sentence the same size"
LONG_SENTENCE = 35     # words
STYLE_RETRIES = 1      # extra AI calls per section when AI-style problems are left (only when safety checks pass)


def _opener(sentence: str) -> str:
    m = re.match(r"[\W\d_]*([A-Za-z’']+)", sentence)
    return m.group(1).lower() if m else ""


def style_check(text: str, keywords: list | None = None) -> dict:
    """AI-style problems in one paragraph (text may contain ⟦n⟧ placeholders, which are ignored).
    Returns {"problems": [readable lines], "flags": [{"sentence", "why"}], "count": n, "flat": bool}."""
    clean = TOKEN_RE.sub("", text or "")
    kws = [k.lower() for k in (keywords or [])]
    sents = _sentences(text)
    problems, flags = [], []
    count = 0

    def flag(sentence, why):
        for f in flags:
            if f["sentence"] == sentence:
                if why not in f["why"]:
                    f["why"] += "; " + why
                return
        flags.append({"sentence": sentence, "why": why})

    # 1. typical AI words and phrases (not when they are part of a must-keep keyword)
    found = []
    for sent in sents:
        for m in _PHRASE_RE.finditer(TOKEN_RE.sub("", sent)):
            p = m.group(0).lower()
            if any(p in k for k in kws):
                continue
            found.append(p)
            flag(sent, f'AI phrase "{m.group(0)}"')
    if found:
        uniq = sorted(set(found))
        problems.append("AI-sounding words/phrases: " + ", ".join(f'"{p}"' for p in uniq)
                        + " (replace with plain words or cut them)")
        count += len(found)

    # 2. flat rhythm: 3+ sentences of nearly the same length
    lengths = [count_words(TOKEN_RE.sub("x", s)) for s in sents]
    flat = False
    if len(lengths) >= 3:
        mean = statistics.mean(lengths)
        if mean >= 8 and statistics.pstdev(lengths) / mean < FLAT_RHYTHM:
            flat = True
            problems.append(f"all {len(lengths)} sentences are about the same length ({min(lengths)}–{max(lengths)} words): "
                            "mix in a short sentence and a longer one")
            count += 1

    # 3. sentences in a row that start with the same word ("This… This…")
    openers = [_opener(s) for s in sents]
    repeated = sorted({o for a, o in zip(openers, openers[1:]) if o and o == a})
    if repeated:
        for i in range(1, len(sents)):
            if openers[i] and openers[i] == openers[i - 1]:
                flag(sents[i], f'starts with "{sents[i].split()[0]}" like the sentence before')
        problems.append("sentences in a row start with the same word: " + ", ".join(f'"{o}"' for o in repeated)
                        + " (vary how they start)")
        count += len(repeated)

    # 4. very long sentences and em dashes
    for sent, n in zip(sents, lengths):
        if n > LONG_SENTENCE:
            flag(sent, f"very long sentence ({n} words)")
            count += 1
    if any(n > LONG_SENTENCE for n in lengths):
        problems.append(f"a sentence is longer than {LONG_SENTENCE} words: split it")
    dashes = clean.count("—") + len(re.findall(r"\s–\s", clean))
    if dashes:
        for sent in sents:
            if "—" in sent or re.search(r"\s–\s", sent):
                flag(sent, "em dash")
        problems.append("em dashes (—): use a comma or a full stop instead")
        count += dashes

    return {"problems": problems, "flags": flags, "count": count, "flat": flat}


def style_hints(paras: list, keywords: list) -> str:
    """What the checker found in the input, as instructions for Gemini."""
    lines = []
    for i, p in enumerate(paras, 1):
        probs = style_check(p, keywords)["problems"]
        if probs:
            lines.append(f"- paragraph {i}: " + "; ".join(probs))
    return "\n".join(lines)


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

