# Humanizer (NP Digital Proto)

Rewrites AI-sounding drafts so they read like a skilled human wrote them, for the content team.
It improves writing quality only: facts, numbers, names, SEO keywords, links and headings are never changed.

## What it does
1. **Input**: paste text (a paragraph or a full article), upload a **Word (.docx)** file, or a **Google Doc link**
   (shared as *Anyone with the link → Viewer*).
2. **Must-keep keywords** (optional). Everything else uses fixed defaults: professional tone, balanced rewrite,
   Indian English (the engine still supports other tones / levels if they are needed again).
3. **Rewrite** with Google Gemini, one section (~350 words) per request, three at a time.
4. **Safety checks (code, not AI)** on every paragraph: links and bold phrases kept exactly once, every number kept,
   must-keep keywords still present (at least once), length not cut by more than half. If a check fails the AI gets one retry with the problems listed;
   if it still fails, **the original paragraph is kept** and marked "Original kept".
5. **AI-style check (NLP, plain code)**: `style_check()` finds what makes a paragraph sound AI-written:
   typical AI words/phrases (the `AI_PHRASES` list; must-keep keywords are ignored), flat rhythm (3+ sentences of
   nearly the same length), sentences in a row starting with the same word, sentences over 35 words, and em dashes.
   - **Before** the rewrite, Gemini gets the exact list of problems per paragraph, not just "rewrite this".
   - **After** the rewrite, the same check runs again. Problems left → **one more targeted try** (`STYLE_RETRIES`)
     with Gemini's previous version and what's still wrong. Per paragraph the safe version with the fewest problems is used.
   - Safety checks always come first: an AI-style retry can never replace a safe paragraph with an unsafe one.
6. **Result**: only the humanized text (headings kept). A short note appears if any paragraph was left as written.
   The **AI-style check bar** shows problems in the draft → problems left; sentences still flagged are **highlighted**
   (hover for the reason; "Hide highlights" turns them off). Copy and download never include the highlights.
7. **Output**: copy the text, or download **.docx**. For Word / Google Doc input the original file comes back with only the
   chosen paragraphs replaced: styles, headings, links, bold phrases, images and tables stay as they were.

Not rewritten on purpose: headings, very short lines (< 8 words), paragraphs with images/fields, tables.

## Files
| File | What |
|---|---|
| `humanizer.py` | Engine: reading input, chunks, Gemini prompt + checks, scores, writing .docx |
| `app.py` | FastAPI app: `/`, `/api/health`, `/api/prepare`, `/api/rewrite`, `/api/score`, `/api/download` |
| `static/index.html` | The page (Proto top bar, two panels) |
| `portal_gate.py` | Portal lock (same as the other agents) |
| `tests/` | `pytest` tests, `run_mock.py` = local run with a fake Gemini for UI testing |

## Environment variables (Vercel → Project → Settings → Environment Variables)
| Name | Value |
|---|---|
| `GEMINI_API_KEY` | Google Gemini key (required) |
| `GEMINI_MODEL` | `gemini-3.6-flash` (default) |
| `AGENT_LAUNCH_SECRET` | same value as the portal (turns the portal lock on) |
| `AGENT_SLUG` | `humanizer` |
| `PORTAL_URL` | `https://npdi-proto.vercel.app` |
| `MAX_INPUT_WORDS` | optional, default 6000 |
| `PRICE_INPUT_PER_M` / `PRICE_OUTPUT_PER_M` | optional, USD per 1M tokens for the cost line |

## Run locally
```
python -m venv .venv
.venv\Scripts\activate        (Mac/Linux: source .venv/bin/activate)
pip install -r requirements.txt
copy .env.example .env         then add GEMINI_API_KEY
python app.py                  → http://127.0.0.1:8010
```
Tests: `pip install pytest` then `python -m pytest -q tests`.

## Deploy on Vercel
Import the GitHub repo in Vercel (Framework preset: Other), add the variables above, deploy.
Then add the portal card (slug `humanizer`, `HUMANIZER_URL` in the portal project).

## Limits
- Up to 6,000 words per run; Word files up to 4 MB (Vercel's upload limit; large images make files bigger).
- English only. The tool doesn't claim to "beat AI detectors"; it improves readability and natural style.
