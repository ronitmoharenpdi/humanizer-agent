"""Run the app locally with a fake Gemini (for UI testing): python tests/run_mock.py"""
import json, os, re, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault("GEMINI_API_KEY", "mock")
import humanizer as hz

SWAPS = [("In today's fast-paced world, ", ""), ("leveraging", "using"), ("leverage", "use"), ("Moreover, ", ""),
         ("Furthermore, ", ""), ("seamlessly", "simply"), ("robust", "solid"), ("crucial", "important"),
         ("delve into", "get into"), ("it is important to note that ", ""), ("It is important to note that ", "")]

def fake(prompt, key, model, **kw):
    payload, _ = json.JSONDecoder().raw_decode(prompt[prompt.rindex('{\n "paragraphs"'):])
    out = []
    for p in payload["paragraphs"]:
        q = p
        for a, b in SWAPS:
            q = q.replace(a, b)
        q = q[:1].upper() + q[1:]
        out.append(q)
    return json.dumps({"paragraphs": out}), {"input_tokens": 900, "output_tokens": 400}

hz.call_gemini = fake
import app as A
for _name in ("rewrite_chunk", "draft_chunk", "finalize_chunk"):
    _orig = getattr(hz, _name)
    setattr(hz, _name, (lambda f: lambda *a, **k: f(*a, caller=fake, **{x: y for x, y in k.items() if x != "caller"}))(_orig))
import uvicorn
uvicorn.run(A.app, host="127.0.0.1", port=8011, log_level="warning")
