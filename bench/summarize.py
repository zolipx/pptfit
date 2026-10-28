"""python bench/summarize.py resultsdir > summary.md  - one table for the whole corpus."""
import glob, json, os, sys

rows = [json.load(open(p)) for p in sorted(glob.glob(os.path.join(sys.argv[1], "**", "*.json"), recursive=True))]
ok = [r for r in rows if "crash" not in r]
crash = [r for r in rows if "crash" in r]
slides = sum(r["slides"] for r in ok)
dirty_b = sum(r["dirty_before"] for r in ok); dirty_a = sum(r["dirty_after"] for r in ok)
print("# Corpus run\n")
print("decks %d, crashed %d, slides %d" % (len(rows), len(crash), slides))
print("slides with problems: %d before -> %d after (%.0f%% of slides clean after)\n" % (dirty_b, dirty_a, 100 * (slides - dirty_a) / max(slides, 1)))
print("| deck | slides | dirty before | dirty after | violations | left | min font | edits | s |")
print("|---|---:|---:|---:|---|---|---:|---:|---:|")
for r in ok:
    print("| %s | %d | %d | %d | %d -> %d | %s | %.2f | %d | %d |" % (r["deck"], r["slides"], r["dirty_before"], r["dirty_after"],
          r["violations_before"], r["violations_after"], " ".join(r["codes_after"]) or "-", r["min_scale"], r["edits"], r["seconds"]))
if crash:
    print("\n## Crashes\n")
    for r in crash:
        print("- %s: %s" % (r["deck"], r["crash"]))
