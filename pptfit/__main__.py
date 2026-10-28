"""Command line.

    python -m pptfit units deck.pptx > units.json
        every piece of text to translate, with ids
    python -m pptfit check-text translations.json [glossary.json]
        accuracy checks on the translation (exit 1 if anything fails)
    python -m pptfit fit deck.pptx translations.json [glossary.json] outdir
        build the English deck and fit it: outdir/final.pptx, cmp/*.png, summary.md, report.json
    python -m pptfit fit-part deck.pptx translations.json glossary.json 1,5,9 part.json
        fit only these slides (for running parts in parallel)
    python -m pptfit merge deck.pptx translations.json glossary.json outdir part1.json part2.json ...
        combine parts and write the same outputs as `fit`
"""
import json, os, sys


def main(argv):
    if len(argv) < 2 or argv[0] not in ("units", "check-text", "fit", "fit-part", "merge"):
        print(__doc__)
        return 2
    cmd, args = argv[0], argv[1:]
    if cmd == "units":
        from .units import make_units
        json.dump(make_units(args[0]), sys.stdout, ensure_ascii=False, indent=1)
        return 0
    if cmd == "check-text":
        from .textcheck import check_all
        tr = json.load(open(args[0]))
        gl = json.load(open(args[1])) if len(args) > 1 else []
        v = check_all(tr, {}, gl) + check_all(tr, {u["id"]: "short" for u in tr if u.get("short")}, gl)
        for x in v:
            print(json.dumps(x, ensure_ascii=False))
        print("%d problems" % len(v), file=sys.stderr)
        return 1 if v else 0
    if cmd in ("fit-part", "merge"):
        from .fit import Fitter
        from . import report
        deck, tr, gl = args[0], json.load(open(args[1])), json.load(open(args[2]))
        if cmd == "fit-part":
            f = Fitter(deck, tr, gl, log=lambda m: print(m, flush=True))
            part = f.fit_slides([int(x) for x in args[3].split(",") if x])
            json.dump(part, open(args[4], "w"), ensure_ascii=False, default=str)
            return 0
        out = args[3]
        os.makedirs(out, exist_ok=True)
        f = Fitter(deck, tr, gl, workdir=os.path.join(out, "work"), log=lambda m: print(m, flush=True))
        rep = f.finish(f.merge([json.load(open(p)) for p in args[4:]]), os.path.join(out, "final.pptx"))
        json.dump(rep, open(os.path.join(out, "report.json"), "w"), ensure_ascii=False, indent=1, default=str)
        print(open(report.make(deck, os.path.join(out, "final.pptx"), rep, out, f.sw, f.sh)).read())
        return 0
    if cmd == "fit":
        from .fit import Fitter
        from . import report
        deck, tr = args[0], json.load(open(args[1]))
        gl = json.load(open(args[2])) if len(args) > 3 else []
        out = args[-1]
        os.makedirs(out, exist_ok=True)
        f = Fitter(deck, tr, gl, workdir=os.path.join(out, "work"), log=lambda m: print(m, flush=True))
        rep = f.run(os.path.join(out, "final.pptx"))
        json.dump(rep, open(os.path.join(out, "report.json"), "w"), ensure_ascii=False, indent=1, default=str)
        print(open(report.make(deck, os.path.join(out, "final.pptx"), rep, out, f.sw, f.sh)).read())
        return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
