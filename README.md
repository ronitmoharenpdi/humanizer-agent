# Humanizer (NP Digital Proto)

Rewrites AI-sounding drafts so they read like a skilled human wrote them, for the content team.
It improves writing quality only: facts, numbers, names, SEO keywords, links and headings are never changed.

## What it does
1. **Input**: paste text (a paragraph or a full article), upload a **Word (.docx)** file, or a **Google Doc link**
   (shared as *Anyone with the link → Viewer*).
2. **Settings**: tone (Professional / Conversational / Friendly / Expert), how much to change
   (Light / Balanced / Strong), English (Indian / US / UK), must-keep keywords, audience, optional brand-voice sample.
3. **Rewrite** with Google Gemini, one section (~350 words) per request, three at a time.
4. **Safety checks (code, not AI)** on every paragraph: links and bold phrases kept exactly once, every number kept,
   must-keep keywords kept, length within range. If a check fails the AI gets one retry with the problems listed;
   if it still fails, **the original paragraph is kept** and marked "Original kept".
5. **Review**: side-by-side Original vs Rewritten (AI-sounding phrases in red, changed wording in green,
   links/bold in orange), per-paragraph **Use original / Use rewrite**, Final-text view.
6. **Scores before → after** (plain code, explainable): readability (Flesch), AI-cliché phrases, sentence rhythm
   (variety in sentence length), average sentence length, passive voice.
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
