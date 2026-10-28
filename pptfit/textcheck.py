"""Checks on the translation itself (no rendering needed).

Accuracy is a hard constraint: a unit may only use a variant that passes here.
    LEFTOVER   source-language characters left in the text
    NUMBERS    a number from the source is missing
    GLOSSARY   a glossary term in the source is not rendered with its fixed English
    CONCEPT    a concept from the unit's concept table is missing
    DUP_WORD   the same word twice in a row ("cup cup")
"""
import re

CJK = re.compile(r"[㐀-鿿豈-﫿]")
NUM = re.compile(r"\d+(?:[.,]\d+)*%?")


def _has(text, term):
    """Term present as a word or a word stem."""
    tail = "[A-Za-z]*" if len(term) >= 4 else "[A-Za-z]{0,2}"  # stems: "spec" -> "specifications"; "cap" must not match "capacity"
    return re.search(r"(?<![A-Za-z])" + re.escape(term.lower()) + tail + r"(?![A-Za-z])", text.lower()) is not None


def check_unit(u, text, glossary=(), target_cjk=False):
    v = []
    if not target_cjk and CJK.search(text):
        v.append({"code": "LEFTOVER", "id": u["id"]})
    for n in NUM.findall(u["src"]):
        if n not in text:
            v.append({"code": "NUMBERS", "id": u["id"], "other": n})
    for g in glossary:
        if g["zh"] and g["zh"] in u["src"]:
            if not any(_has(text, t) for t in [g["en"]] + list(g.get("abbr") or []) if t):
                v.append({"code": "GLOSSARY", "id": u["id"], "other": g["en"]})
    abbr = {g["en"].lower(): [a for a in g.get("abbr") or []] for g in glossary}
    for c in u.get("concepts") or []:
        en = (c.get("en") or "").strip()
        if not en or not any(_has(text, t) for t in [en] + abbr.get(en.lower(), [])):
            v.append({"code": "CONCEPT", "id": u["id"], "other": c.get("zh")})
    words = re.findall(r"[A-Za-z][A-Za-z'-]*", text.lower())
    for a, b in zip(words, words[1:]):
        if a == b and len(a) > 2:
            v.append({"code": "DUP_WORD", "id": u["id"], "other": a})
    return v


def check_all(translations, choice, glossary=(), target_cjk=False):
    """Violations for the chosen variant of every unit."""
    out = []
    for u in translations:
        text = u.get(choice.get(u["id"], "full")) or u.get("full") or ""
        out += check_unit(u, text, glossary, target_cjk)
    return out


def allowed_variants(u, glossary=(), target_cjk=False):
    """Variants of a unit that keep every concept: the fitter may only pick these."""
    return [k for k in ("full", "short", "short2") if u.get(k) and not check_unit(u, u[k], glossary, target_cjk)]
