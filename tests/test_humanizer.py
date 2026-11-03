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


def test_rewrite_chunk_draft_retry_then_final():
    paras = ["A ⟦1⟧ costs 20% less for every new investor who opens one through a broker in India today."]
    toks = [{"1": "demat account"}]
    good = "A ⟦1⟧ now costs 20% less for each new investor opening one via an Indian broker."
    answers = iter([
        json.dumps({"tells": [], "paragraphs": ["It costs less for new investors."]}),   # draft breaks rules
        json.dumps({"tells": [], "paragraphs": [good]}),                                   # draft retry OK
        json.dumps({"remaining": [], "paragraphs": [good.replace("now ", "")]}),          # final OK
    ])
    systems = []

    def fake(prompt, key, model, **kw):
        systems.append("step 1" if "step 1 of 2" in kw.get("system", "") else "step 2")
        return next(answers), {"input_tokens": 100, "output_tokens": 50}

    res = h.rewrite_chunk(paras, toks, h.Settings(), "Costs", "k", "m", caller=fake)
    assert res["status"] == ["rewritten"] and res["usage"]["ai_calls"] == 3
    assert res["paragraphs"] == [good.replace("now ", "")]
    assert systems == ["step 1", "step 1", "step 2"]


def test_final_falls_back_to_draft_then_original():
    paras = ["A ⟦1⟧ costs 20% less for every new investor who opens one through a broker in India today."]
    toks = [{"1": "demat account"}]
    draft = ["A ⟦1⟧ costs 20% less for each new investor who opens one with an Indian broker."]
    bad = lambda prompt, key, model, **kw: (json.dumps({"paragraphs": ["Totally different."]}), {"input_tokens": 1, "output_tokens": 1})  # noqa: E731
    res = h.finalize_chunk(paras, draft, toks, h.Settings(), "", "k", "m", caller=bad)
    assert res["paragraphs"] == draft and res["status"] == ["rewritten"]
    res = h.rewrite_chunk(paras, toks, h.Settings(), "", "k", "m", caller=bad)
    assert res["status"] == ["kept"] and res["paragraphs"] == paras


def test_checks_dashes_and_invented_numbers():
    o = "Most brokers let you open an account online and the charges stay low for every new investor today."
    assert h.check_paragraph(o, "Most brokers let you open one online — and charges stay low for new investors.", {}, [])
    assert any("new numbers" in p for p in h.check_paragraph(o, "Most brokers let you open one online in 5 minutes, and charges stay low.", {}, []))
    assert h.check_paragraph(o, "Most brokers let you open an account online, and charges stay low for new investors.", {}, []) == []


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
