"""python bench/run_deck.py deck.pptx translations.json glossary.json outdir"""
import json, os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from pptfit.fit import Fitter
from pptfit import report

deck, tr, gl, out = sys.argv[1:5]
os.makedirs(out, exist_ok=True)
f = Fitter(deck, json.load(open(tr)), json.load(open(gl)), workdir=os.path.join(out, "work"),
           log=lambda m: print(m, flush=True))
rep = f.run(os.path.join(out, "final.pptx"))
json.dump(rep, open(os.path.join(out, "report.json"), "w"), ensure_ascii=False, indent=1, default=str)
report.make(deck, os.path.join(out, "final.pptx"), rep, out, f.sw, f.sh)
print(open(os.path.join(out, "summary.md")).read(), flush=True)
