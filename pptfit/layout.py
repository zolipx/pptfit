"""Ink, constraints and checks. The original deck is the specification.

1. ink: which rendered word belongs to which element (position + text match).
2. constraints(original): what the original guarantees, per slide
   - container: the smallest visible shape/picture that holds an element's ink
   - obstacles: visible things the ink did NOT touch (pictures, lines, bars,
     other text ink); visible boxes that did not overlap each other
   - size: rendered font size; groups: look-alike boxes; titles
3. check(candidate): every guarantee that broke is a violation.

Constraints come from the original itself, so the original passes by
construction and nothing deck-specific is hard-coded.

Extension point: add a new kind of guarantee by recording it in `constraints()`
and testing it in `check()` under a new violation code.
"""
import re
from collections import Counter, defaultdict
from statistics import median

TEXTY = ("text", "table", "chart", "smartart")
VISIBLE_KINDS = ("picture", "line", "chart", "table", "smartart")
TOL = 1.0          # pt
LINE_MIN = 1.5     # pt: a line's box has zero height or width; overlap tests give it this stroke thickness
GRAZE = 0.1        # an original that only grazed an obstacle (<= 10% of its word area) may not overlap it more
BACKDROP = 0.15    # an obstacle covering >= 15% of the slide is background art, judged against the whole text block
FLOOR_BODY, FLOOR_TITLE = 0.8, 0.9


# ---------- geometry ----------
def inter(a, b):
    w = min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0])
    h = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
    return w * h if w > 0 and h > 0 else 0.0


def obox(e):
    """Box for overlap tests: lines get a stroke thickness instead of a zero side."""
    b = e["bbox"]
    if e["kind"] != "line":
        return b
    w, h = max(b[2], LINE_MIN), max(b[3], LINE_MIN)
    return [b[0] - (w - b[2]) / 2, b[1] - (h - b[3]) / 2, w, h]


def word_hit(ws, b):
    """Overlap of each word with box b, summed (pt^2). Word by word: the union box of a ragged text grazes corners its words never touch."""
    return sum(inter(w["box"], b) for w in ws)


def excess(inner, outer):
    """How far `inner` sticks out of `outer` (max over the four sides)."""
    return max(outer[0] - inner[0], inner[0] + inner[2] - outer[0] - outer[2],
               outer[1] - inner[1], inner[1] + inner[3] - outer[1] - outer[3], 0.0)


def union(boxes):
    x0 = min(b[0] for b in boxes); y0 = min(b[1] for b in boxes)
    x1 = max(b[0] + b[2] for b in boxes); y1 = max(b[1] + b[3] for b in boxes)
    return [x0, y0, x1 - x0, y1 - y0]


def area(b):
    return max(b[2], 0) * max(b[3], 0)


def dist(pt, b):
    dx = max(b[0] - pt[0], 0, pt[0] - b[0] - b[2])
    dy = max(b[1] - pt[1], 0, pt[1] - b[1] - b[3])
    return (dx * dx + dy * dy) ** 0.5


# ---------- text helpers ----------
CJK = re.compile(r"[㐀-鿿豈-﫿]")


def element_text(e):
    t = e.get("text") or ""
    t += " " + " ".join(c.get("text", "") for c in e.get("cells", []))
    t += " " + " ".join(e.get("chart_texts", []))
    return t


def tokens(text):
    out = []
    for w in re.split(r"[^\w㐀-鿿'-]+", text.lower()):
        if w:
            out.append(w)
            out += [p for p in w.split("-") if p and p != w]  # a break at a hyphen is a legal break
    return out


def is_visible(e):
    return not e.get("background") and (e.get("visible_fill") or e.get("visible_line") or e["kind"] in VISIBLE_KINDS)


# ---------- 1. ink ----------
def assign(words, els):
    """{element id: [word, ...]} using position first, text match second, line majority last."""
    cands = [e for e in els if e["kind"] in TEXTY]
    vocab = {e["id"]: element_text(e).lower() for e in cands}
    owner = {}
    for i, w in enumerate(words):
        c = (w["box"][0] + w["box"][2] / 2, w["box"][1] + w["box"][3] / 2)
        tok = w["text"].lower().strip(".,;:!?()[]（）：，。；")
        pad = 4 + 0.8 * w["box"][3]
        scored = []
        for e in cands:
            d = dist(c, e["bbox"])
            match = bool(tok) and tok in vocab[e["id"]]
            if d <= pad or (match and d <= 4 * w["box"][3]):
                scored.append((not match, d, area(e["bbox"]), e["id"]))
        if scored:
            owner[i] = min(scored)[3]
    lines = defaultdict(list)
    for i, w in enumerate(words):
        lines[w["block"]].append(i)
    for idxs in lines.values():  # an overflowing line stays with the frame most of its words belong to
        votes = Counter(owner[i] for i in idxs if i in owner)
        if votes:
            top, n = votes.most_common(1)[0]
            if n > len(idxs) / 2:
                for i in idxs:
                    w = words[i]
                    tok = w["text"].lower().strip(".,;:!?()[]（）：，。；")
                    if owner.get(i) != top and not (tok and tok in vocab.get(owner.get(i), "")):
                        owner[i] = top
    out = defaultdict(list)
    for i, eid in owner.items():
        out[eid].append(words[i])
    return out


def ink_boxes(owned):
    return {eid: union([w["box"] for w in ws]) for eid, ws in owned.items() if ws}


def line_count(ws):
    """Distinct text lines among an element's words (by vertical centre)."""
    ys = sorted(w["box"][1] + w["box"][3] / 2 for w in ws)
    if not ys:
        return 0
    h = median(w["box"][3] for w in ws)
    n, last = 1, ys[0]
    for y in ys[1:]:
        if y - last > 0.5 * h:
            n, last = n + 1, y
    return n


def rendered_size(ws):
    sizes = [w["size"] for w in ws if w.get("size")]
    return round(median(sizes), 2) if sizes else None


# ---------- 2. constraints ----------
def _same_look(a, b):
    if a["kind"] != b["kind"] or a.get("prst") != b.get("prst") or a.get("fill") != b.get("fill") or a.get("line") != b.get("line"):
        return False
    if a.get("is_title") or b.get("is_title"):
        return False
    return all(abs(a["bbox"][k] - b["bbox"][k]) <= 0.1 * max(a["bbox"][k], b["bbox"][k], 1) for k in (2, 3))


def groups_of(els):
    texty = [e for e in els if e["kind"] == "text" and (e.get("text") or "").strip()]
    groups, seen = [], set()
    for e in texty:
        if e["id"] in seen:
            continue
        g = [x for x in texty if x["id"] not in seen and _same_look(e, x)]
        if len(g) >= 2:
            groups.append([x["id"] for x in g])
            seen.update(x["id"] for x in g)
    return groups


def constraints(slide_els, words, slide_w, slide_h):
    owned = assign(words, slide_els)
    inks = ink_boxes(owned)
    by_id = {e["id"]: e for e in slide_els}
    vis = [e for e in slide_els if is_visible(e)]
    con = {"ink": inks, "container": {}, "ink_clear": {}, "box_clear": [], "size": {}, "groups": groups_of(slide_els),
           "bbox": {e["id"]: e["bbox"] for e in slide_els}, "titles": [], "slide": [0, 0, slide_w, slide_h]}
    for eid, ink in inks.items():
        e = by_id[eid]
        con["size"][eid] = rendered_size(owned[eid])
        con.setdefault("lines", {})[eid] = line_count(owned[eid])
        if e.get("is_title"):
            con["titles"].append(eid)
        holders = [v for v in vis if excess(ink, v["bbox"]) <= 2 and v["kind"] not in ("line",)
                   and (v["id"] == eid or area(v["bbox"]) > area(ink))]
        if holders:
            h = min(holders, key=lambda v: area(v["bbox"]))
            con["container"][eid] = h["id"]
            con.setdefault("container_excess", {})[eid] = excess(ink, h["bbox"])  # the original may already poke out a little
        clear = []
        for v in vis:
            if v["id"] in (eid, con["container"].get(eid)) or excess(e["bbox"], v["bbox"]) <= 1:
                continue  # its own box, its container, or something it sits on
            if area(v["bbox"]) >= BACKDROP * slide_w * slide_h:  # text routinely runs across background art and watermarks
                if inter(ink, v["bbox"]) <= TOL:
                    clear.append(["backdrop", v["id"], 0.0])
                continue
            hit0 = word_hit(owned[eid], obox(v))  # a title whose last glyph box grazes the start of a rule
            if hit0 <= TOL or hit0 <= GRAZE * sum(area(w["box"]) for w in owned[eid]):
                clear.append(["box", v["id"], round(hit0, 1)])
        for oid, oink in inks.items():
            if oid != eid and inter(ink, oink) <= TOL:
                clear.append(["ink", oid])
        con["ink_clear"][eid] = clear
    for i, a in enumerate(vis):  # visible boxes that did not overlap must keep apart
        for b in vis[i + 1:]:
            if inter(a["bbox"], b["bbox"]) <= TOL and excess(a["bbox"], b["bbox"]) > 1 and excess(b["bbox"], a["bbox"]) > 1:
                con["box_clear"].append([a["id"], b["id"]])
    return con


# ---------- 3. check ----------
def widows(owned):
    """Multi-line texts whose last line is a single short word (looks broken; soft cost only)."""
    n = 0
    for ws in owned.values():
        if line_count(ws) < 2:
            continue
        last_y = max(w["box"][1] for w in ws)
        last = [w for w in ws if w["box"][1] > last_y - 0.5 * w["box"][3]]
        if len(last) == 1 and len(last[0]["text"]) <= 6:
            n += 1
    return n


def check(con, cand_els, words, expected_text=None, frozen_ids=(), floors=(FLOOR_BODY, FLOOR_TITLE)):
    """Violations of one candidate slide: [{code, id, other, amount}].

    expected_text: {element id: text it should show} for the broken-word check.
    frozen_ids: elements allowed to change geometry on purpose (e.g. trimmed lines).
    """
    owned = assign(words, cand_els)
    inks = ink_boxes(owned)
    box = {e["id"]: e["bbox"] for e in cand_els}
    obx = {e["id"]: obox(e) for e in cand_els}
    by_kind = {e["id"]: e["kind"] for e in cand_els}
    v = []
    for eid, ink in inks.items():
        cid = con["container"].get(eid)
        allow = TOL + con.get("container_excess", {}).get(eid, 0.0)
        if cid and cid in box and excess(ink, box[cid]) > allow:
            v.append({"code": "INK_OUT", "id": eid, "other": cid, "amount": round(excess(ink, box[cid]), 1)})
        for c in con["ink_clear"].get(eid, []):
            kind, oid, before = c[0], c[1], (c[2] if len(c) > 2 else 0.0)
            ob = obx.get(oid) if kind in ("box", "backdrop") else inks.get(oid)
            hit = (word_hit(owned[eid], ob) if kind == "box" else inter(ink, ob)) if ob else 0.0
            if hit > 1.25 * before + TOL:
                v.append({"code": "INK_HIT", "id": eid, "other": oid, "amount": round((hit - before) ** 0.5, 1),
                          "on": "text" if kind == "ink" else "shape"})
        if excess(ink, con["slide"]) > TOL:
            v.append({"code": "OFF_SLIDE", "id": eid, "amount": round(excess(ink, con["slide"]), 1)})
        own_layout = by_kind.get(eid) in ("chart", "smartart")  # these lay out their own labels; fix wording, not geometry
        if not own_layout and con.get("lines", {}).get(eid) == 1 and not con["container"].get(eid) and line_count(owned[eid]) > 1:
            v.append({"code": "LINES", "id": eid, "amount": float(line_count(owned[eid]) - 1)})  # free-standing one-liners stay one line
        osz, csz = con["size"].get(eid), rendered_size(owned[eid])
        if osz and csz:
            floor = osz * (floors[1] if eid in con["titles"] else floors[0])
            if csz < floor - 0.1:
                v.append({"code": "FONT_FLOOR", "id": eid, "amount": round(floor - csz, 1)})
        if expected_text and eid in expected_text and not own_layout:
            want = set(tokens(expected_text[eid]))
            for w in owned[eid]:
                t = (tokens(w["text"]) or [""])[0]
                if t and not CJK.search(t) and want and t not in want and any(x != t and (x.startswith(t) or x.endswith(t)) for x in want):
                    v.append({"code": "BROKEN_WORD", "id": eid, "other": t, "amount": 1})
    for a, b in con["box_clear"]:
        if a in box and b in box and inter(box[a], box[b]) > TOL:
            v.append({"code": "BOX_HIT", "id": a, "other": b, "amount": round(inter(box[a], box[b]) ** 0.5, 1)})
    sizes = {eid: rendered_size(ws) for eid, ws in owned.items()}
    for g in con["groups"]:
        s = [sizes[x] for x in g if sizes.get(x)]
        if s and max(s) - min(s) > 0.5:
            v.append({"code": "GROUP_SIZE", "id": g[0], "other": ",".join(g), "amount": round(max(s) - min(s), 1)})
        dims = [box[x][2:] for x in g if x in box]
        odims = [con["bbox"][x][2:] for x in g]
        if dims and all(abs(d[0] - odims[0][0]) <= 1 and abs(d[1] - odims[0][1]) <= 1 for d in odims):
            if max(d[0] for d in dims) - min(d[0] for d in dims) > 1 or max(d[1] for d in dims) - min(d[1] for d in dims) > 1:
                v.append({"code": "GROUP_BOX", "id": g[0], "other": ",".join(g), "amount": 1})
    for e in cand_els:
        if e["kind"] in ("picture", "line", "other") and not e.get("inherited") and e["id"] in con["bbox"] and e["id"] not in frozen_ids:
            if max(abs(p - q) for p, q in zip(e["bbox"], con["bbox"][e["id"]])) > 0.5:
                v.append({"code": "MOVED", "id": e["id"], "amount": 1})
    if con.get("bg"):
        v += check_bg(con["bg"], owned, con["bg_scale"])
    return v, {"inks": inks, "sizes": sizes, "widows": widows(owned)}


def title_sizes(cons):
    """Deck-level: titles that shared a rendered size must keep sharing it."""
    classes = defaultdict(list)
    for s, con in cons.items():
        for t in con["titles"]:
            if con["size"].get(t):
                classes[round(con["size"][t] * 2) / 2].append((s, t))
    return [members for members in classes.values() if len(members) > 1]


# ---------- background regions (things baked into the slide background or master) ----------
BG_DPI = 36            # 0.5 px per pt is plenty for "is this text over a footer bar"
BG_TOL = 28            # max channel difference that still counts as the same background
BG_NOISY = 18          # skip the check where the original text sat on a busy background (photo)


def bg_regions(bg_rgb, inks, scale):
    """{element id: bool mask} - the connected area of the background colour each text sat on.

    bg_rgb: HxWx3 uint8 render of the slide with every slide shape removed (background,
    master and layout only). scale: pixels per pt.
    """
    import numpy as np
    from scipy import ndimage
    out = {}
    H, W, _ = bg_rgb.shape
    img = bg_rgb.astype(int)
    for eid, b in inks.items():
        x0, y0 = max(int(b[0] * scale), 0), max(int(b[1] * scale), 0)
        x1, y1 = min(int((b[0] + b[2]) * scale) + 1, W), min(int((b[1] + b[3]) * scale) + 1, H)
        patch = img[y0:y1, x0:x1].reshape(-1, 3)
        if len(patch) == 0 or patch.std(axis=0).max() > BG_NOISY:
            continue
        home = np.median(patch, axis=0)
        mask = np.abs(img - home).max(axis=2) <= BG_TOL
        lab, _ = ndimage.label(mask)
        ids = np.unique(lab[y0:y1, x0:x1]); ids = ids[ids > 0]
        if len(ids):
            out[eid] = np.isin(lab, ids)
    return out


def check_bg(regions, owned, scale):
    """INK_BG: words that left the background area their text originally sat on."""
    v = []
    for eid, mask in regions.items():
        H, W = mask.shape
        total = 0.0  # summed over words, so pushing fewer words out counts as progress
        for w in owned.get(eid, []):
            b = w["box"]
            x0, y0 = max(int(b[0] * scale), 0), max(int(b[1] * scale), 0)
            x1, y1 = min(int((b[0] + b[2]) * scale) + 1, W), min(int((b[1] + b[3]) * scale) + 1, H)
            if x1 <= x0 or y1 <= y0:
                total += b[3]
                continue
            outside = 1.0 - mask[y0:y1, x0:x1].mean()
            if outside > 0.15:
                total += outside * b[3]
        if total:
            v.append({"code": "INK_BG", "id": eid, "amount": round(total, 1)})
    return v
