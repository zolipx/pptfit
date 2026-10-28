"""Every guarantee must catch its own kind of damage, and the original must pass itself.

Runs on the synthetic deck only (bench/synth.py), never on real decks.
    python -m pytest -q tests
"""
import copy, json, os, sys, tempfile

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "bench"))
import synth  # noqa: E402
from pptfit import extract as X, layout as L, render as R, textcheck as T, units as U  # noqa: E402
from pptfit.ops import Deck  # noqa: E402

WORK = tempfile.mkdtemp(prefix="pptfit_test_")


@pytest.fixture(scope="module")
def deck():
    path = synth.build(WORK)
    tr, missing = synth.translations(path)
    assert not missing
    o = X.extract(path)
    pdf = R.to_pdf(path)
    cons = {s["index"]: L.constraints(s["shapes"], R.page_words(pdf, s["index"] - 1, o["slide_w"], o["slide_h"]),
                                      o["slide_w"], o["slide_h"]) for s in o["slides"]}
    return {"path": path, "tr": tr, "orig": o, "cons": cons}


def codes(deck, path, slide):
    o = deck["orig"]
    pdf = R.render_slide(path, slide)
    cand = X.extract(path)["slides"][slide - 1]["shapes"]
    v, _ = L.check(deck["cons"][slide], cand, R.page_words(pdf, 0, o["slide_w"], o["slide_h"]))
    return {x["code"] for x in v}


def mutate(deck, name, fn):
    d = Deck(deck["path"])
    fn(d)
    out = os.path.join(WORK, name + ".pptx")
    d.save(out)
    return out


def set_text(d, eid, text):
    tx = d.idx[eid]._element.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}txBody")
    U.set_paragraph_text(tx.findall("{http://schemas.openxmlformats.org/drawingml/2006/main}p")[0], text)


def ids(deck, slide, pred):
    return [e["id"] for e in deck["orig"]["slides"][slide - 1]["shapes"] if pred(e)]


def test_original_passes_itself(deck):
    for s in deck["orig"]["slides"]:
        assert codes(deck, deck["path"], s["index"]) == set(), s["index"]


def test_ink_out_of_box(deck):
    box = deck["cons"][1]["groups"][0][0]                     # a flowchart box
    p = mutate(deck, "ink_out", lambda d: set_text(d, box, "Customer orders at the kiosk and pays for every cup"))
    assert "INK_OUT" in codes(deck, p, 1)


def test_title_wraps(deck):
    t = deck["cons"][2]["titles"][0]
    p = mutate(deck, "lines", lambda d: set_text(d, t, "Highlights of the new flagship store"))
    assert codes(deck, p, 2) & {"LINES", "INK_HIT"}


def test_group_size(deck):
    g = deck["cons"][1]["groups"][0]
    p = mutate(deck, "group", lambda d: d.scale_font(g[0], 0.8, 18))
    assert "GROUP_SIZE" in codes(deck, p, 1)


def test_picture_moved(deck):
    pic = ids(deck, 3, lambda e: e["kind"] == "picture" and not e.get("inherited"))[0]
    p = mutate(deck, "moved", lambda d: d.move(pic, 10, 0))
    assert "MOVED" in codes(deck, p, 3)


def test_text_into_footer(deck):
    texts = [e for e in deck["orig"]["slides"][2]["shapes"] if e["kind"] == "text"]
    body = max(texts, key=lambda e: e["bbox"][3])["id"]       # the tall body column above the footer
    p = mutate(deck, "footer", lambda d: set_text(d, body, "\n".join(["Line %d of a very long description" % i for i in range(30)])))
    assert codes(deck, p, 3) & {"INK_HIT", "INK_BG", "OFF_SLIDE"}


def test_box_hits_neighbour(deck):
    a, b = deck["cons"][1]["groups"][0][:2]
    ea = next(e for e in deck["orig"]["slides"][0]["shapes"] if e["id"] == a)
    eb = next(e for e in deck["orig"]["slides"][0]["shapes"] if e["id"] == b)
    p = mutate(deck, "hit", lambda d: d.move(a, eb["bbox"][0] - ea["bbox"][0], eb["bbox"][1] - ea["bbox"][1]))
    assert codes(deck, p, 1) & {"BOX_HIT", "INK_HIT"}


def test_shrink_on_overflow_is_seen(deck):
    box = ids(deck, 6, lambda e: e.get("autofit") == "norm")[0]
    p = mutate(deck, "autofit", lambda d: set_text(d, box, "Opening hours 7:00 to 22:00 with 36 seats " * 12))
    assert codes(deck, p, 6) & {"FONT_FLOOR", "INK_HIT", "INK_BG", "OFF_SLIDE"}


def test_accuracy_checks(deck):
    u = next(x for x in deck["tr"] if len(x["concepts"]) >= 2)
    assert not T.check_unit(u, u["full"])
    dropped = u["full"].lower().replace(u["concepts"][0]["en"].lower(), "")
    assert any(v["code"] == "CONCEPT" for v in T.check_unit(u, dropped))
    assert any(v["code"] == "LEFTOVER" for v in T.check_unit(u, u["src"]))
    assert any(v["code"] == "DUP_WORD" for v in T.check_unit(u, "cup cup " + u["full"]))
