import io
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import docx  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

import humanizer as h  # noqa: E402

LONG = ("In today's fast-paced world, leveraging a robust demat account is crucial for investors. "
        "Moreover, it seamlessly stores 100% of your shares electronically since 1996.")


def _add_link(par, text, url):
    part = par.part
    rid = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                         is_external=True)
    hl = OxmlElement("w:hyperlink")
    hl.set(qn("r:id"), rid)
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    r.append(t)
    hl.append(r)
    par._p.append(hl)


def make_docx() -> bytes:
    d = docx.Document()
    d.add_heading("What is a Demat Account", level=1)
    p = d.add_paragraph("Furthermore, a ")
    _add_link(p, "demat account", "https://example.com/demat")
    p.add_run(" helps you delve into investing. ")
    b = p.add_run("Open it online")
    b.bold = True
    p.add_run(" in 10 minutes and it is important to note that charges are low for everyone who invests.")
    d.add_paragraph("Short caption")
    d.add_paragraph(LONG)
    out = io.BytesIO()
    d.save(out)
    return out.getvalue()


def test_read_docx_tokens_and_kinds():
    blocks = h.read_docx(make_docx())
    kinds = [b.kind for b in blocks]
    assert kinds == ["heading", "para", "skip", "para"]
    para = blocks[1]
    assert para.tokens == {"1": "demat account", "2": "Open it online"}
    assert "⟦1⟧" in para.text and "⟦2⟧" in para.text
    assert para.plain().startswith("Furthermore, a demat account helps")


def test_build_docx_keeps_link_and_bold():
    data = make_docx()
    blocks = h.read_docx(data)
    new = "A ⟦1⟧ lets you start investing. ⟦2⟧ in 10 minutes, and the charges stay low for every investor."
    out = h.build_docx(data, {str(blocks[1].id): new})
    d = docx.Document(io.BytesIO(out))
    p = d.paragraphs[1]
    assert p.text == "A demat account lets you start investing. Open it online in 10 minutes, and the charges stay low for every investor."
    assert p._p.findall(qn("w:hyperlink")), "link kept"
    assert any(r.bold and r.text == "Open it online" for r in p.runs), "bold kept"
    assert d.paragraphs[0].text == "What is a Demat Account" and d.paragraphs[0].style.name.startswith("Heading")
    assert d.paragraphs[3].text == LONG, "untouched paragraph unchanged"


def test_read_pasted_headings_and_lists():
    text = "# Demat Guide\n\n" + LONG + "\n\nWhy It Matters\n" + LONG + "\n- " + LONG + "\nOk"
    blocks = h.read_pasted(text)
    assert [b.kind for b in blocks] == ["heading", "para", "heading", "para", "para", "heading"]


def test_check_paragraph_catches_problems():
    orig = "A ⟦1⟧ costs 20% less in 2026 for every new investor who opens one through a broker today."
    toks = {"1": "demat account"}
    assert h.check_paragraph(orig, "A ⟦1⟧ costs 20% less in 2026 for each new investor opening one with a broker today.", toks, ["demat account"]) == []
    probs = h.check_paragraph(orig, "It costs 25% less in 2026 for new investors opening one with a broker this year.", toks, ["demat account"])
    joined = " | ".join(probs)
    assert "⟦1⟧" in joined and "20%" in joined


def test_rewrite_chunk_retry_then_keep():
    paras = ["A ⟦1⟧ costs 20% less for every new investor who opens one through a broker in India today."]
    toks = [{"1": "demat account"}]
    answers = iter([
        json.dumps({"paragraphs": ["It costs less for new investors."]}),             # breaks rules
        json.dumps({"paragraphs": ["A ⟦1⟧ now costs 20% less for each new investor opening one via an Indian broker."]}),
    ])
    fake = lambda prompt, key, model: (next(answers), {"input_tokens": 100, "output_tokens": 50})  # noqa: E731
    res = h.rewrite_chunk(paras, toks, h.Settings(), "Costs", "k", "m", caller=fake)
    assert res["status"] == ["rewritten"] and res["usage"]["ai_calls"] == 2

    bad = lambda prompt, key, model: (json.dumps({"paragraphs": ["Totally different."]}), {"input_tokens": 1, "output_tokens": 1})  # noqa: E731
    res = h.rewrite_chunk(paras, toks, h.Settings(), "", "k", "m", caller=bad)
    assert res["status"] == ["kept"] and res["paragraphs"] == paras


def test_scores_improve_on_plainer_text():
    before = h.score_text(LONG * 3)
    after = h.score_text("A demat account holds your shares in electronic form. It's simple. "
                         "You can open one online in about 10 minutes, and most brokers charge little for it.")
    assert before["ai_phrases"] > 0 and after["ai_phrases"] == 0
    assert after["readability"] > before["readability"]


def test_plan_chunks_respects_headings():
    blocks = h.read_docx(make_docx())
    chunks = h.plan_chunks(blocks)
    assert all(c["heading"] == "What is a Demat Account" for c in chunks)
    assert [i for c in chunks for i in c["ids"]] == [1, 3]


# --------------------------------------------------------------------------- #
# AI-style check (NLP)
# --------------------------------------------------------------------------- #

AI_PARA = ("In today's fast-paced world, leveraging a robust demat account is crucial for every new investor in India. "
           "This account stores all of your shares in a safe electronic form at all times. "
           "This makes it easy to buy and sell shares online whenever the market is open.")
HUMAN_PARA = ("A demat account holds your shares in electronic form. It's simple. You can open one online in about "
              "ten minutes, and most brokers in India charge very little for it.")


def test_style_check_finds_ai_patterns():
    r = h.style_check(AI_PARA, [])
    joined = " ".join(r["problems"])
    assert "crucial" in joined and "same length" in joined and '"this"' in joined
    assert r["flat"] and r["count"] >= 5
    assert h.style_check(HUMAN_PARA, [])["count"] == 0


def test_style_check_respects_keywords_and_placeholders():
    assert h.style_check("Read our comprehensive guide to opening a demat account in India today.", ["comprehensive guide"])["count"] == 0
    assert h.style_check("Read ⟦1⟧ before you open a demat account with any broker in India.", [])["count"] == 0


def test_hints_go_into_the_prompt():
    prompt = h.build_prompt([AI_PARA], [{}], h.Settings(), "", hints=h.style_hints([AI_PARA], []))
    assert "AI-style checker found" in prompt and "crucial" in prompt


def test_style_retry_uses_best_version():
    still_ai = AI_PARA.replace("In today's fast-paced world, leveraging", "Using")
    answers = iter([json.dumps({"paragraphs": [still_ai]}), json.dumps({"paragraphs": [HUMAN_PARA]})])
    prompts = []

    def fake(prompt, key, model):
        prompts.append(prompt)
        return next(answers), {"input_tokens": 10, "output_tokens": 5}

    res = h.rewrite_chunk([AI_PARA], [{}], h.Settings(), "", "k", "m", caller=fake)
    assert res["paragraphs"] == [HUMAN_PARA] and res["status"] == ["rewritten"]
    assert res["usage"]["ai_calls"] == 2 and "still sounds AI-written" in prompts[1]
    assert res["style"][0]["count"] == 0 and res["style_before"] >= 5


def test_style_retry_never_beats_safety_and_keeps_better_one():
    worse = AI_PARA.replace("crucial", "important") + " Moreover, it is pivotal."   # safe but still AI-ish
    answers = iter([json.dumps({"paragraphs": [worse]}), json.dumps({"paragraphs": ["Short."]})])  # 2nd breaks length
    fake = lambda prompt, key, model: (next(answers), {"input_tokens": 1, "output_tokens": 1})  # noqa: E731
    res = h.rewrite_chunk([AI_PARA], [{}], h.Settings(), "", "k", "m", caller=fake)
    assert res["paragraphs"] == [worse] and res["status"] == ["rewritten"] and res["usage"]["ai_calls"] == 2


def test_no_style_retry_when_clean():
    fake = lambda prompt, key, model: (json.dumps({"paragraphs": [HUMAN_PARA]}), {"input_tokens": 1, "output_tokens": 1})  # noqa: E731
    res = h.rewrite_chunk([AI_PARA], [{}], h.Settings(), "", "k", "m", caller=fake)
    assert res["usage"]["ai_calls"] == 1 and res["style"][0]["flags"] == []


# --------------------------------------------------------------------------- #
# Predictability pass (GPT-2 scorer replaced by a stand-in)
# --------------------------------------------------------------------------- #

def _fake_scores(text):
    """Stand-in for GPT-2: sentences containing 'simple' or 'right' are 'predictable'."""
    import predictability as p
    rows = []
    for a, b in p.sentence_spans(text):
        t = text[a:b]
        ai = bool(re.search(r"\b(simple|right)\b", t, re.I))
        rows.append({"start": a, "end": b, "text": t.strip(), "ppl": 10.0 if ai else 60.0, "top10": 0.9 if ai else 0.5, "ai": ai})
    return rows


import re  # noqa: E402

ORIG = "Starting a SIP is simple and convenient. You can begin with ₹500 per month through most fund houses in India today."
REWRITE_1 = "Starting a SIP is simple and convenient. You can begin with ₹500 per month through most fund houses in India."
REWRITE_2 = "Setting one up takes about as long as ordering lunch online. You can begin with ₹500 per month through most fund houses in India."


def test_predictability_pass_rewrites_flagged_sentences(monkeypatch):
    monkeypatch.setattr(h.pred, "score_sentences", _fake_scores)
    monkeypatch.setattr(h.pred, "enabled", lambda: True)
    answers = iter([json.dumps({"paragraphs": [REWRITE_1]}), json.dumps({"paragraphs": [REWRITE_2]})])
    prompts = []

    def fake(prompt, key, model):
        prompts.append(prompt)
        return next(answers), {"input_tokens": 10, "output_tokens": 5}

    res = h.rewrite_chunk([ORIG], [{}], h.Settings(), "", "k", "m", caller=fake)
    assert res["paragraphs"] == [REWRITE_2]
    assert "Sentences to rewrite" in prompts[1] and "Starting a SIP is simple" in prompts[1]
    assert res["predictability"] == {"used": True, "rounds": 1, "ai_sentences_before": 1, "ai_sentences_after": 0}
    assert res["usage"]["ai_calls"] == 2


def test_predictability_never_breaks_safety(monkeypatch):
    monkeypatch.setattr(h.pred, "score_sentences", _fake_scores)
    monkeypatch.setattr(h.pred, "enabled", lambda: True)
    unsafe = "Setting one up takes about as long as ordering lunch online. You can begin with a small amount every month."  # ₹500 lost
    answers = iter([json.dumps({"paragraphs": [REWRITE_1]})] + [json.dumps({"paragraphs": [unsafe]})] * 2)
    fake = lambda prompt, key, model: (next(answers), {"input_tokens": 1, "output_tokens": 1})  # noqa: E731
    res = h.rewrite_chunk([ORIG], [{}], h.Settings(), "", "k", "m", caller=fake)
    assert res["paragraphs"] == [REWRITE_1] and res["predictability"]["rounds"] == 2


def test_predictability_skipped_when_model_missing(monkeypatch):
    monkeypatch.setattr(h.pred, "score_sentences", lambda text: [])
    fake = lambda prompt, key, model: (json.dumps({"paragraphs": [REWRITE_1]}), {"input_tokens": 1, "output_tokens": 1})  # noqa: E731
    res = h.rewrite_chunk([ORIG], [{}], h.Settings(), "", "k", "m", caller=fake)
    assert res["paragraphs"] == [REWRITE_1] and res["usage"]["ai_calls"] == 1 and res["predictability"]["used"] is False
