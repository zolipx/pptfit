"""Stress test on many decks: pseudo-translate each, fit it, record what happened.

python bench/run_corpus.py outdir deck1.pptx deck2.pptx ...   ->  outdir/<name>.json per deck
A crash is a result too: it is recorded with its traceback instead of stopping the run.
"""
import json, os, sys, tempfile, time, traceback

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.dirname(__file__))
import pseudo  # noqa: E402
from pptfit.fit import Fitter  # noqa: E402

PER_SLIDE = int(os.environ.get("PPTFIT_SLIDE_BUDGET", "40"))
PER_DECK = int(os.environ.get("PPTFIT_DECK_BUDGET", "360"))
MAX_SLIDES = int(os.environ.get("PPTFIT_MAX_SLIDES", "12"))

out = sys.argv[1]
os.makedirs(out, exist_ok=True)
for deck in sys.argv[2:]:
    name = os.path.splitext(os.path.basename(deck))[0]
    rec, t0 = {"deck": os.path.basename(deck)}, time.time()
    try:
        from pptx import Presentation
        Presentation(deck)
    except Exception as e:  # noqa: BLE001  corrupt or not a pptx: the input is invalid, not the tool
        rec.update({"invalid": repr(e)[:200], "seconds": round(time.time() - t0)})
        json.dump(rec, open(os.path.join(out, name + ".json"), "w"), indent=1)
        print(json.dumps(rec), flush=True)
        continue
    try:
        tr = pseudo.translations(deck)
        f = Fitter(deck, tr, [], workdir=os.path.join(tempfile.gettempdir(), "pptfit_work", name), log=lambda m: None)
        slides = [s["index"] for s in f.orig["slides"]][:MAX_SLIDES]
        rep = f.finish(f.fit_slides(slides, None, PER_SLIDE, PER_DECK), os.path.join(tempfile.gettempdir(), "pptfit_work", name + "_final.pptx"))
        from pptfit import report  # side-by-side images of every fitted slide, so the numbers can be checked by eye
        report.make(deck, os.path.join(tempfile.gettempdir(), "pptfit_work", name + "_final.pptx"), rep, os.path.join(out, name), f.sw, f.sh, dpi=60)
        scales = list(rep["state"]["scale"].values())
        rec.update({
            "slides": len(slides), "units": len(tr),
            "dirty_before": sum(1 for r in rep["slides"].values() if r["before"]),
            "dirty_after": sum(1 for r in rep["slides"].values() if r["after"]),
            "violations_before": sum(r["before"] or 0 for r in rep["slides"].values()),
            "violations_after": sum(len(r["after"]) for r in rep["slides"].values()),
            "codes_after": sorted({v["code"] for r in rep["slides"].values() for v in r["after"]}),
            "left": {str(n): [[v["code"], v["id"], v.get("other")] for v in r["after"]] for n, r in rep["slides"].items() if r["after"]},
            "master_units": sum(1 for u in tr if u["id"][0] in "LM"),  # layout and master text, translated but never moved by the fitter
            "min_scale": min(scales) if scales else 1.0, "edits": sum(len(r["history"]) - 1 for r in rep["slides"].values() if r["history"]),
        })
    except Exception as e:  # noqa: BLE001
        rec.update({"crash": repr(e)[:300], "trace": traceback.format_exc()[-1500:]})
    rec["seconds"] = round(time.time() - t0)
    json.dump(rec, open(os.path.join(out, name + ".json"), "w"), indent=1)
    print(json.dumps({k: v for k, v in rec.items() if k != "trace"}), flush=True)
