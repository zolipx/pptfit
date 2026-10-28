"""The fitter: make a translated deck honour every guarantee of the original.

State = which variant each unit uses + per-element font scale, insets, growth,
shift, and line trims. Every candidate is rebuilt from the original deck, so
edits never accumulate drift, and every candidate is judged by rendering its slide.

Per slide: list violations -> take the offending element (or its look-alike
group) with the most violations -> render all its moves in parallel -> keep the
best if it beats the current state, else try the next offender -> repeat until
clean, stuck, or out of time.

Moves (extension point: add one in `moves()`):
    short   a shorter variant that still passes every accuracy check
    insets  tighten the text frame margins
    spacing tighten line and paragraph spacing (never below 1.1x the font size)
    shrink  any lower font step, never below the floor (titles: deck pass only)
    grow    enlarge the box into free space; lines in the way are trimmed in the same move
    shift   move the box into free space (expensive: the design moves)
    trim    shorten a line that the text runs into
    narrow  pull a box that sticks out of the slide back inside

Layout comes first. Slides still broken after the accurate moves get a second, lossy
pass: `short`/`short2` wordings that lost a detail and fonts down to 70% become
allowed. Every such compromise is written into the slide notes for a human to
check, because fixing a word is easy and fixing a layout is not.
"""
import copy, itertools, os, re, tempfile, threading, time
from concurrent.futures import ThreadPoolExecutor

from . import extract as X, layout as L, render as R, textcheck as T, units as U
from .ops import Deck

STEPS = [1.0, 0.95, 0.9, 0.85, 0.8]
LOSSY_FLOOR = 0.7  # second phase only, and always reported
GAP = 3.0          # pt kept free around a grown box
WORKERS = max(1, min(4, os.cpu_count() or 1))


def slide_of(uid):
    """Slide number of a unit or element id; None for layout and master units (L3_7#p0, M1_4#p0)."""
    m = re.match(r"s(\d+)[_#]", uid)
    return int(m.group(1)) if m else None


def new_state():
    return {"choice": {}, "scale": {}, "spacing": {}, "insets": [], "grow": {}, "shift": {}, "trim": {}}


SEVERITY = {"INK_OUT": 2, "OFF_SLIDE": 2, "BOX_HIT": 2, "MOVED": 2}  # how painful for a person to repair; others 1


def severity(v):
    """Text over text is unreadable (3); text out of its box or the slide, boxes overlapping, a moved picture
    break the layout (2); text touching a shape or line, a wrap, a split word, a size mismatch are cosmetic (1)."""
    if v["code"] == "INK_HIT":
        return 3 if v.get("on") == "text" else 1
    return SEVERITY.get(v["code"], 1)


def badness(viol):
    return sum(severity(v) for v in viol)


def cost(st, titles=()):
    c = sum((1 - f) * 10 for f in st["scale"].values()) + sum((1 - f) * 4 for f in st.get("spacing", {}).values())
    c += sum((abs(g[0]) + abs(g[1])) / 60 for g in st["grow"].values())
    c += sum((abs(d[0]) + abs(d[1])) / 6 * (2 if e in titles else 1) for e, d in st["shift"].items())
    c += 0.5 * sum(1 for v in st["choice"].values() if v != "full") + 0.2 * len(st["insets"])
    c += sum(d / 80 for _, d in st["trim"].values())
    return round(c, 3)


class Fitter:
    def __init__(self, src_pptx, translations, glossary=(), workdir=None, log=print, latin="Arial"):
        self.src, self.tr, self.gl, self.log, self.latin = src_pptx, translations, glossary, log, latin
        self.work = workdir or tempfile.mkdtemp(prefix="fit_")
        os.makedirs(self.work, exist_ok=True)
        self.orig = X.extract(src_pptx)
        self.sw, self.sh = self.orig["slide_w"], self.orig["slide_h"]
        pdf = R.to_pdf(src_pptx)
        self.cons = {}
        for s in self.orig["slides"]:
            words = R.page_words(pdf, s["index"] - 1, self.sw, self.sh)
            self.cons[s["index"]] = L.constraints(s["shapes"], words, self.sw, self.sh)
        bg = R.to_pdf(R.background_only(src_pptx, os.path.join(self.work, "bg.pptx")))
        for s in self.orig["slides"]:  # what the text sat on, as rendered from background/master only
            con = self.cons[s["index"]]
            rgb = R.page_rgb(bg, s["index"] - 1, L.BG_DPI)
            con["bg_scale"] = rgb.shape[1] / self.sw
            con["bg"] = L.bg_regions(rgb, con["ink"], con["bg_scale"])
        self.els = {e["id"]: e for s in self.orig["slides"] for e in s["shapes"]}
        self.titles = {t for c in self.cons.values() for t in c["titles"]}
        self.units_of = {}
        for u in translations:
            self.units_of.setdefault(u["id"].split("#")[0].split("/r")[0], []).append(u)
        self.allowed = {u["id"]: T.allowed_variants(u, glossary) or ["full"] for u in translations}
        self.any_variant = {u["id"]: [k for k in ("full", "short", "short2") if u.get(k)] for u in translations}
        self.mixed = {u["id"] for u in U.make_units(src_pptx) if u.get("mixed")}  # bold or coloured words inside a paragraph
        self.lossy = False  # second phase: layout first, wording may lose detail, font may go to 70%
        probe = Deck(src_pptx)
        self.spaced = {e: probe.has_spacing(e) for e in probe.idx}  # elements with explicit line/paragraph spacing
        self.counter = itertools.count()
        self.lock = threading.Lock()

    # ---------- build & evaluate ----------
    def build(self, st, out):
        tmp = out + ".tr.pptx"
        U.apply(self.src, self.tr, tmp, st["choice"])
        d = Deck(tmp)
        d.autofit_off()
        d.latin_font(self.latin)
        for eid, f in st["scale"].items():
            d.scale_font(eid, f, self.base_size(eid))
        for eid, f in st.get("spacing", {}).items():
            d.line_spacing(eid, f, self.base_size(eid))
        for eid in st["insets"]:
            d.set_insets(eid, 1.8)
        for eid, (dw, dh, ax, ay) in st["grow"].items():
            d.resize(eid, dw, dh, ax, ay)
        for eid, (dx, dy) in st["shift"].items():
            d.move(eid, dx, dy)
        for eid, (end, dd) in st["trim"].items():
            d.trim_line(eid, end, dd)
        d.save(out)
        os.remove(tmp)
        return out

    def base_size(self, eid):
        for con in self.cons.values():
            if con["size"].get(eid):
                return con["size"][eid]
        runs = [r for p in self.els[eid].get("paragraphs", []) for r in p["runs"]]
        return runs[0]["size"] if runs else 18.0

    def expected(self, st):
        return {eid: " ".join(u.get(st["choice"].get(u["id"], "full")) or u["full"] for u in us)
                for eid, us in self.units_of.items()}

    def evaluate(self, st, slide):
        path = os.path.join(self.work, "cand_%d.pptx" % next(self.counter))
        self.build(st, path)
        pdf = R.render_slide(path, slide)
        words = R.page_words(pdf, 0, self.sw, self.sh)
        cand = X.extract(path)["slides"][slide - 1]["shapes"]
        os.remove(path)
        frozen = set(st["trim"]) | set(st["grow"]) | set(st["shift"])
        floors = (LOSSY_FLOOR, LOSSY_FLOOR) if self.lossy else (L.FLOOR_BODY, L.FLOOR_TITLE)
        v, info = L.check(self.cons[slide], cand, words, self.expected(st), frozen, floors)
        return v, info, {e["id"]: e for e in cand}

    # ---------- moves ----------
    def target_of(self, eid, slide):
        for g in self.cons[slide]["groups"]:
            if eid in g:
                return tuple(g)
        return (eid,)

    def free(self, box, others):
        """Free room around `box` in each direction, given other boxes (pt)."""
        x0, y0, x1, y1 = box[0], box[1], box[0] + box[2], box[1] + box[3]
        room = {"left": x0, "right": self.sw - x1, "up": y0, "down": self.sh - y1}
        for o in others:
            ox0, oy0, ox1, oy1 = o[0], o[1], o[0] + o[2], o[1] + o[3]
            if oy1 > y0 and oy0 < y1:
                if ox1 <= x0: room["left"] = min(room["left"], x0 - ox1)
                if ox0 >= x1: room["right"] = min(room["right"], ox0 - x1)
            if ox1 > x0 and ox0 < x1:
                if oy1 <= y0: room["up"] = min(room["up"], y0 - oy1)
                if oy0 >= y1: room["down"] = min(room["down"], oy0 - y1)
        return {k: max(v - GAP, 0.0) for k, v in room.items()}

    @staticmethod
    def _trimmable(o):
        return o["kind"] == "line" and not o.get("inherited")

    def _trim_for(self, st, line, box):
        """Trim needed so `line` stops overlapping `box` (None if over the 50% cap)."""
        b = line["bbox"]
        if L.inter([b[0], b[1], max(b[2], 1), max(b[3], 1)], box) <= 0:
            return None
        if b[2] >= b[3]:
            left_side = box[0] + box[2] / 2 < b[0] + b[2] / 2
            end = "left" if left_side else "right"
            d = (box[0] + box[2] - b[0] if left_side else b[0] + b[2] - box[0]) + 4
        else:
            top_side = box[1] + box[3] / 2 < b[1] + b[3] / 2
            end = "top" if top_side else "bottom"
            d = (box[1] + box[3] - b[1] if top_side else b[1] + b[3] - box[1]) + 4
        prev = st["trim"].get(line["id"], (end, 0))[1]
        full = max(self.els[line["id"]]["bbox"][2], self.els[line["id"]]["bbox"][3])
        return (end, prev + d) if 0 < prev + d <= 0.5 * full else None

    def moves(self, st, target, slide, viol, info, cand):
        con = self.cons[slide]
        mine = [x for x in viol if x["id"] in target]
        need = max([x["amount"] for x in mine] + [2.0]) + 2.0
        is_title = any(e in self.titles for e in target)
        out = []
        units = [u for e in target for u in self.units_of.get(e, [])]
        variants = self.any_variant if self.lossy else self.allowed
        for alt in ("short", "short2"):
            sw = {u["id"]: alt for u in units if alt in variants[u["id"]] and st["choice"].get(u["id"], "full") != alt}
            if sw:
                n = copy.deepcopy(st); n["choice"].update(sw); out.append(("short", n)); break
        if any(e not in st["insets"] for e in target):
            n = copy.deepcopy(st); n["insets"].extend(e for e in target if e not in n["insets"]); out.append(("insets", n))
        if any(self.spaced.get(e) for e in target):
            cur = min(st.get("spacing", {}).get(e, 1.0) for e in target)
            for f in (f for f in (0.9, 0.8, 0.7) if f < cur - 1e-6):
                n = copy.deepcopy(st)
                for e in target:
                    n.setdefault("spacing", {})[e] = f
                out.append(("spacing", n))
        if not is_title or self.lossy:
            cur = min(st["scale"].get(e, 1.0) for e in target)
            steps = STEPS + [0.75, 0.7] if self.lossy else STEPS
            for f in (f for f in steps if f < cur - 1e-6 and (f >= 0.9 or not is_title)):  # every remaining step: one step may look worse than two; titles stay >= 90% so they keep matching each other
                n = copy.deepcopy(st)
                for e in target:
                    n["scale"][e] = f
                out.append(("shrink", n))
        boxes = {e: cand[e]["bbox"] for e in target if e in cand}
        if boxes:
            def others(skip_lines):
                o = [x["bbox"] for x in cand.values() if x["id"] not in target and L.is_visible(x)
                     and not (skip_lines and self._trimmable(x))
                     and not any(L.excess(boxes[e], x["bbox"]) <= 1 for e in boxes)]
                return (o + [b for k, b in info["inks"].items() if k not in target]
                        + [b for k, b in con["ink"].items() if k not in target])  # never into where another text sat in the original
            e0 = self.els[next(iter(boxes))]
            ax = {"center": "center", "right": "right"}.get((e0.get("paragraphs") or [{}])[0].get("align"), "left")
            ay = "center" if e0.get("anchor") == "center" else "top"
            for skip_lines in (False, True):
                rooms = [self.free(b, others(skip_lines)) for b in boxes.values()]
                room = {k: min(r[k] for r in rooms) for k in rooms[0]}
                horiz = {"center": 2 * min(room["left"], room["right"]), "right": room["left"], "left": room["right"]}[ax]
                vert = 2 * min(room["up"], room["down"]) if ay == "center" else room["down"]
                wmax = max(b[2] for b in boxes.values()); hmax = max(b[3] for b in boxes.values())
                sizes = {(min(x, horiz), 0) for x in (2 * need, 0.25 * wmax, 0.5 * wmax, wmax)}
                sizes |= {(0, min(y, vert)) for y in ((2 if ay == "center" else 1) * need, 0.25 * hmax, 0.5 * hmax)}
                for dw, dh in sorted(sizes):
                    if dw < 1 and dh < 1:
                        continue
                    n = copy.deepcopy(st)
                    ok = True
                    for e, b in boxes.items():
                        g = n["grow"].setdefault(e, [0, 0, ax, ay])
                        g[0] += dw; g[1] += dh
                        grown = [b[0] - {"left": 0, "center": dw / 2, "right": dw}[ax], b[1] - (dh / 2 if ay == "center" else 0), b[2] + dw, b[3] + dh]
                        if skip_lines:  # lines in the way get trimmed in the same move
                            for x in cand.values():
                                if self._trimmable(x) and L.inter([x["bbox"][0], x["bbox"][1], max(x["bbox"][2], 1), max(x["bbox"][3], 1)], grown) > 0:
                                    t = self._trim_for(n, x, grown)
                                    if t is None:
                                        ok = False
                                    else:
                                        n["trim"][x["id"]] = t
                    if ok:
                        out.append(("grow+trim" if skip_lines and n["trim"] != st["trim"] else "grow", n))
            # a box that sticks out of the slide: pull its outer edge back inside, text rewraps within the page
            over_r = max(b[0] + b[2] - (self.sw - GAP) for b in boxes.values())
            over_l = max(GAP - b[0] for b in boxes.values())
            if over_r > 1 or over_l > 1:
                n = copy.deepcopy(st)
                for e, b in boxes.items():
                    g = n["grow"].setdefault(e, [0, 0, "left", ay])
                    if over_r > 1:
                        g[0] -= over_r; g[2] = "left"
                    if over_l > 1:
                        g[0] -= over_l; g[2] = "right"
                out.append(("narrow", n))
            for dirn, (dx, dy) in {"up": (0, -1), "down": (0, 1), "left": (-1, 0), "right": (1, 0)}.items():
                if room[dirn] >= need:
                    n = copy.deepcopy(st)
                    for e in boxes:
                        d = n["shift"].setdefault(e, [0, 0]); d[0] += dx * need; d[1] += dy * need
                    out.append(("shift-" + dirn, n))
        for x in mine:  # the text runs into a line: shorten the line
            o = cand.get(x.get("other"))
            if o and self._trimmable(o) and x["id"] in info["inks"]:
                t = self._trim_for(st, o, info["inks"][x["id"]])
                if t:
                    n = copy.deepcopy(st); n["trim"][o["id"]] = t; out.append(("trim", n))
        return out

    # ---------- search ----------
    def score(self, viol, st, info=None):
        """Less severe violations first (never trade a cosmetic flaw for text over text), then less total
        overshoot (partial progress counts), then lower cost; a lonely short word on the last line of a
        paragraph costs a little extra."""
        widows = (info or {}).get("widows", 0)
        return (badness(viol), round(sum(x.get("amount", 1) for x in viol), 1), cost(st, self.titles) + 0.3 * widows)

    def fit_slide(self, st, slide, budget=300, max_iter=15):
        t0 = time.time()
        viol, info, cand = self.evaluate(st, slide)
        history = [{"step": "start", "violations": len(viol), "cost": cost(st, self.titles)}]
        pool = ThreadPoolExecutor(WORKERS)
        for _ in range(max_iter):
            if not viol or time.time() - t0 > budget:
                break
            weight = {}
            for x in viol:
                eid = x["id"]
                if self.els.get(eid, {}).get("inherited"):  # master/layout shapes are read-only: act on the other party
                    eid = x.get("other")
                    if not eid or self.els.get(eid, {}).get("inherited") or eid not in self.els:
                        continue
                t = self.target_of(eid, slide)
                weight[t] = weight.get(t, 0) + 1 + x.get("amount", 0) / 100
            improved = False
            for t in sorted(weight, key=weight.get, reverse=True):
                cands = self.moves(st, t, slide, viol, info, cand)
                if not cands or time.time() - t0 > budget:
                    continue
                results = list(pool.map(lambda c: (c[0], c[1], self.evaluate(c[1], slide)), cands))
                name, n, (v2, i2, c2) = min(results, key=lambda r: self.score(r[2][0], r[1], r[2][1]))
                if self.score(v2, n, i2) < self.score(viol, st, info):
                    st, viol, info, cand = n, v2, i2, c2
                    history.append({"step": name, "target": list(t), "violations": len(viol), "cost": cost(st, self.titles)})
                    self.log("  s%d %s %s -> %d violations" % (slide, name, ",".join(t), len(viol)))
                    improved = True
                    break
            if not improved:
                break
        pool.shutdown()
        return st, viol, history

    def fit_slides(self, slides=None, st=None, budget_per_slide=180, total_budget=2400):
        """Fit some slides (all by default). Slides are independent, so parts can run in parallel
        and be merged: states of different slides never touch the same keys."""
        t0 = time.time()
        st = st or new_state()
        report = {"slides": {}}
        for s in self.orig["slides"]:
            n = s["index"]
            if slides and n not in slides:
                continue
            v0, _, _ = self.evaluate(st, n)
            left = total_budget - (time.time() - t0)
            if left <= 0:
                report["slides"][n] = {"before": len(v0), "after": v0, "history": [], "skipped": True}
                continue
            t1 = time.time()
            st, v, hist = self.fit_slide(st, n, min(budget_per_slide, left))
            if v:  # layout first: let wording lose detail and fonts go to 70%, but say so in the notes
                self.lossy = True
                st2, v2, hist2 = self.fit_slide(st, n, max(budget_per_slide - (time.time() - t1), 30))
                self.lossy = False
                if (badness(v2), sum(x.get("amount", 1) for x in v2)) < (badness(v), sum(x.get("amount", 1) for x in v)):
                    st, v, hist = st2, v2, hist + [dict(h, step="lossy " + h["step"]) for h in hist2[1:]]
            report["slides"][n] = {"before": len(v0), "after": v, "history": hist}
            self.log("slide %d: %d -> %d" % (n, len(v0), len(v)))
        report["state"] = st
        report["seconds"] = round(time.time() - t0)
        return report

    @staticmethod
    def merge(parts):
        """Combine reports of disjoint slide sets into one."""
        st, report = new_state(), {"slides": {}, "seconds": 0}
        for p in parts:
            for k, v in p["state"].items():
                if isinstance(v, dict):
                    st.setdefault(k, {}).update(v)
                else:
                    st.setdefault(k, []).extend(x for x in v if x not in st[k])
            report["slides"].update({int(n): r for n, r in p["slides"].items()})
            report["seconds"] = max(report["seconds"], p.get("seconds", 0))
        st["trim"] = {k: tuple(v) for k, v in st["trim"].items()}
        report["state"] = st
        return report

    def finish(self, report, out_pptx):
        """Deck-level title pass, the final file, and notes for the human who finishes the job."""
        st = self.title_pass(self.master_pass(report["state"], report), report)
        report["boxes"] = {}
        for n, r in report["slides"].items():  # boxes of what is still wrong, for the evidence images
            if r["after"]:
                _, info, cand = self.evaluate(st, n)
                report["boxes"][str(n)] = {x["id"]: info["inks"].get(x["id"]) or cand.get(x["id"], {}).get("bbox") for x in r["after"]}
        self.build(st, out_pptx)
        notes = {n: [advice(x, self.els) for x in r["after"]] for n, r in report["slides"].items() if r["after"]}
        for u in self.tr:  # wording that lost detail to save the layout: easy for a human to touch up
            v = st["choice"].get(u["id"], "full")
            slide = slide_of(u["id"])
            if slide and T.check_unit(u, u.get(v) or "", self.gl):
                notes.setdefault(slide, []).append("wording shortened to fit, check the meaning: '%s' -> '%s'"
                                                   % (u["src"][:40], (u.get(v) or "")[:60]))
            if slide and u["id"] in self.mixed:
                notes.setdefault(slide, []).append("part of this text was bold or coloured, now one style: '%s'"
                                                   % (u.get(v) or u.get("full") or "")[:60])
        for eid, f in st["scale"].items():
            if f < L.FLOOR_BODY - 1e-6 and eid in self.els:
                slide = slide_of(eid)
                notes.setdefault(slide, []).append("font reduced to %d%% to fit: '%s'" % (round(f * 100), (self.els[eid].get("text") or "")[:30].replace("\n", " ")))
        report["leftover"] = {}  # what still reads in the source language on the rendered result, whatever its origin
        pdf = R.to_pdf(out_pptx)
        for s in self.orig["slides"]:
            left = [w["text"] for w in R.page_words(pdf, s["index"] - 1, self.sw, self.sh) if L.CJK.search(w["text"])]
            if left:
                report["leftover"][s["index"]] = left
                notes.setdefault(s["index"], []).append("text still in the source language: %s" % " ".join(left)[:80])
        if notes:
            d = Deck(out_pptx)
            for n, lines in notes.items():
                d.add_note(n, ["[pptfit] Please check this slide:"] + ["- " + l for l in lines])
            d.save(out_pptx)
        report["state"] = st
        report["advice"] = notes
        report["text_violations"] = T.check_all(self.tr, st["choice"], self.gl)
        report["needs_human"] = sorted(n for n, r in report["slides"].items() if r["after"])
        report["review"] = sorted(notes)
        return report

    def run(self, out_pptx, budget_per_slide=180, total_budget=2400):
        return self.finish(self.fit_slides(None, None, budget_per_slide, total_budget), out_pptx)

    def master_pass(self, st, report):
        """Layout and master text is shared by many slides and never moved slide by slide: where it still
        breaks, try its shorter wordings, then smaller font steps, judged on all those slides at once."""
        homes = {}
        for s, r in report["slides"].items():
            for x in r["after"]:
                home = self.els.get(x["id"], {}).get("home")
                if home and self.units_of.get(home):
                    homes.setdefault(home, {}).setdefault(s, x["id"])
        for home, where in homes.items():  # where: {slide: this shape's id on that slide}
            alias = next(iter(where.values()))
            units = self.units_of[home]
            best, best_n = (st, None), sum(badness(report["slides"][s]["after"]) for s in where)
            steps = [("choice", {u["id"]: alt for u in units if alt in self.any_variant[u["id"]]}) for alt in ("short", "short2")]
            steps += [("scale", f) for f in (0.9, 0.8, 0.7)]
            for kind, v in steps:  # each step on top of the best so far
                if kind == "choice" and not v:
                    continue
                n = copy.deepcopy(best[0])
                if kind == "choice":
                    n["choice"].update(v)
                else:
                    n["scale"][alias] = v
                results = {s: self.evaluate(n, s)[0] for s in where}
                k = sum(badness(x) for x in results.values())
                if k < best_n:
                    best, best_n = (n, results), k
                if best_n == 0:
                    break
            if best[1]:
                st = best[0]
                for s, v in best[1].items():
                    report["slides"][s]["after"] = v
        return st

    def title_pass(self, st, report):
        """Titles that shared a size keep sharing it: if any of them still breaks, shrink the whole class."""
        for members in L.title_sizes(self.cons):
            members = [(s, t) for s, t in members if s in report["slides"]]  # only slides that were fitted
            if len(members) < 2 or not any(any(x["id"] == t for x in report["slides"][s]["after"]) for s, t in members):
                continue
            for f in (0.95, 0.9):
                n = copy.deepcopy(st)
                for _, t in members:
                    n["scale"][t] = f
                results = {s: self.evaluate(n, s)[0] for s in {s for s, _ in members}}
                if sum(badness(v) for v in results.values()) < sum(badness(report["slides"][s]["after"]) for s in results):
                    st = n
                    for s, v in results.items():
                        report["slides"][s]["after"] = v
                    if not any(any(x["id"] == t for x in results[s]) for s, t in members):
                        break
        return st


ADVICE = {
    "INK_OUT": "text sticks out of its box/panel by {amount} pt: shorten the text, enlarge the box, or reduce the font",
    "INK_HIT": "text touches {other}: move or shorten the text, or move {other}",
    "INK_BG": "text runs over a background area (footer bar, panel edge): shorten the text or reduce the font",
    "OFF_SLIDE": "text runs off the slide by {amount} pt",
    "BOX_HIT": "box overlaps {other}",
    "FONT_FLOOR": "font is below 80% of the original size",
    "LINES": "a one-line text now wraps: widen its box or use a shorter wording",
    "BROKEN_WORD": "the word '{other}' is split across lines: widen the box",
    "GROUP_SIZE": "look-alike boxes ({other}) no longer share one font size",
    "GROUP_BOX": "look-alike boxes ({other}) no longer share one size",
    "MOVED": "this element moved",
}


def advice(v, els):
    """One line a human can act on: which element, what is wrong, what to try."""
    e = els.get(v["id"], {})
    label = (e.get("text") or e.get("name") or v["id"]).strip().replace("\n", " ")[:30]
    other = els.get(v.get("other"), {}).get("name") or v.get("other") or ""
    return "'%s': %s" % (label, ADVICE.get(v["code"], v["code"]).format(amount=v.get("amount"), other=other))
