"""Pseudo-localisation: a fake translation of any deck that only makes the text longer.

python bench/pseudo.py deck.pptx out_translations.json [expansion=1.4]

Each unit is extended with its own words until it is `expansion` times longer ("short"
is 1.15x), so line wrapping behaves like real text and numbers stay intact. This lets
the fitter be stress-tested on any public deck without translating anything.
CJK sources are rendered as Latin filler of the typical CJK->English width.
"""
import json, os, re, sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from pptfit import units as U

CJK = re.compile(r"[㐀-鿿]")
FILLER = "system data process quality control station module output review update".split()


def stretch(text, factor):
    lines = []
    for line in text.split("\n"):
        fields = re.findall(r"\{F\d+\}", line)  # fields stay, once each
        line = re.sub(r"\s*\{F\d+\}\s*", " ", line).strip() if fields else line
        lines.append(_stretch_line(line, factor) + "".join(" " + f for f in fields))
    return "\n".join(lines)


def _stretch_line(line, factor):
    if CJK.search(line):  # one CJK character is about two Latin letters wide; English needs ~3.5
        n = len(CJK.findall(line))
        words, k = [], 0
        while sum(len(w) + 1 for w in words) < n * 3.5 * factor / 1.4:
            words.append(FILLER[k % len(FILLER)]); k += 1
        return " ".join(words).capitalize()
    words = line.split()
    if not words:
        return line
    out, k = list(words), 0
    while len(" ".join(out)) < factor * len(line):
        out.append(re.sub(r"\d", "", words[k % len(words)]).strip(".,;:") or "more")
        k += 1
    return " ".join(out)


def translations(deck, expansion=1.4):
    return [{"id": u["id"], "src": u["src"], "full": stretch(u["src"], expansion),
             "short": stretch(u["src"], 1 + (expansion - 1) * 0.4), "concepts": []}
            for u in U.make_units(deck)]


if __name__ == "__main__":
    deck, out = sys.argv[1], sys.argv[2]
    factor = float(sys.argv[3]) if len(sys.argv) > 3 else 1.4
    tr = translations(deck, factor)
    json.dump(tr, open(out, "w"), ensure_ascii=False, indent=1)
    print("units", len(tr))
