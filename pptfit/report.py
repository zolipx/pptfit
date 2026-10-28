"""Evidence for a human: side-by-side pages and a short summary.

    make(src_pptx, final_pptx, report_dict, outdir) -> outdir/cmp/s{N}.png, outdir/summary.md
Remaining violations are outlined in red on the right-hand page.
"""
import os
from collections import Counter

from PIL import Image, ImageDraw

from . import render as R


def make(src_pptx, final_pptx, rep, outdir, slide_w, slide_h, dpi=96):
    cmp_dir = os.path.join(outdir, "cmp")
    os.makedirs(cmp_dir, exist_ok=True)
    a_pdf, b_pdf = R.to_pdf(src_pptx), R.to_pdf(final_pptx)
    scale = dpi / 72.0
    lines = ["| slide | naive | final | moves | left |", "|---|---:|---:|---|---|"]
    for n, r in sorted(rep["slides"].items(), key=lambda kv: int(kv[0])):
        n = int(n)
        a = R.page_png(a_pdf, n - 1, os.path.join(cmp_dir, "a.png"), dpi)
        b = R.page_png(b_pdf, n - 1, os.path.join(cmp_dir, "b.png"), dpi)
        ia, ib = Image.open(a).convert("RGB"), Image.open(b).convert("RGB")
        draw = ImageDraw.Draw(ib)
        boxes = rep.get("boxes", {}).get(str(n), {})
        for v in r["after"]:
            bx = boxes.get(v["id"])
            if bx:
                x, y, w, h = [c * scale for c in bx]
                draw.rectangle([x - 2, y - 2, x + w + 2, y + h + 2], outline=(220, 0, 0), width=3)
        out = Image.new("RGB", (ia.width + ib.width + 12, max(ia.height, ib.height)), "white")
        out.paste(ia, (0, 0)); out.paste(ib, (ia.width + 12, 0))
        out.save(os.path.join(cmp_dir, "s%d.png" % n))
        moves = Counter(h["step"] for h in r.get("history", [])[1:])
        left = Counter(v["code"] for v in r["after"])
        lines.append("| %d | %s | %d | %s | %s |" % (n, r.get("before"), len(r["after"]),
                     " ".join("%s×%d" % kv for kv in moves.items()) or "-", " ".join("%s×%d" % kv for kv in left.items()) or "-"))
    for f in ("a.png", "b.png"):
        os.remove(os.path.join(cmp_dir, f))
    tv = Counter(v["code"] for v in rep.get("text_violations", []))
    lines += ["", "text checks: " + (" ".join("%s×%d" % kv for kv in tv.items()) or "all pass"),
              "needs human: " + (", ".join(str(x) for x in rep.get("needs_human", [])) or "none"),
              "seconds: %s" % rep.get("seconds")]
    open(os.path.join(outdir, "summary.md"), "w").write("\n".join(lines) + "\n")
    return os.path.join(outdir, "summary.md")
