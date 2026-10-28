# pptfit — translate a PPT and keep its layout

Reference code for agents. A translated deck must keep what the original
guaranteed: text inside the boxes and panels it sat in, clear of the pictures,
lines and bars it did not touch, look-alike boxes alike, titles alike, fonts not
shrunk below a floor, and the meaning complete. The original deck is the
specification; nothing in this code is specific to one deck.

## Setup

```
apt-get install -y --no-install-recommends libreoffice-impress fonts-liberation fonts-crosextra-carlito fonts-noto-cjk
pip install python-pptx pymupdf pillow numpy scipy
```

In restricted sandboxes, if apt fails with a setgroups error, add
`-o APT::Sandbox::User=root -o Dir::Cache::archives=$PWD/tmp/apt/archives`.

## Workflow

1. `python -m pptfit units deck.pptx > units.json` lists every piece of text: shape
   paragraphs, table cells, chart and SmartArt strings, and notes.
2. **The agent translates.** It writes `glossary.json` (`[{zh, en, abbr:[]}]`, one
   English rendering per term) and `translations.json`, one record per unit:
   `{id, src, concepts:[{zh, en}], full, short, back}`.
   - `full`: a faithful translation; `short`: the same meaning, worded tighter;
     `back`: `short` translated back, for review.
   - `concepts`: every content word of `src` with its English. `full` and `short` must
     contain every concept.
   - `short2` (optional, for labels, flow boxes, titles and other tight spots): the
     tightest wording a reader still understands; it may drop secondary words. It is
     used only when the layout cannot be saved otherwise, and every use is listed in
     the slide notes for review.
   - Only standard abbreviations (Assy, Qty, HV). Never invent one (no "M" for
     manual). Keep the title casing style of the deck.
3. `python -m pptfit check-text translations.json glossary.json` until it prints
   0 problems.
4. `python -m pptfit fit deck.pptx translations.json glossary.json out/` produces
   `out/final.pptx`, `out/cmp/s{N}.png` (original | result, anything still wrong is
   outlined in red), `out/summary.md` and `out/report.json`.
5. Look at every `cmp` image yourself; the checks can miss things. Judge by what a person
   would have to repair: layout (shapes, groups, flow charts, callouts, alignment, text off
   pictures, lines and bars) is slow to repair and must be right; wording, punctuation or a
   label to retype takes seconds. Slides the accurate moves cannot fix get a lossy pass
   (shorter wordings, fonts down to 70%), and every compromise is written into that slide's
   notes as "[pptfit] Please check this slide". For layout problems still visible, add a
   tighter `short2` for the units involved or extend the code (see below) and run step 4
   again. For the quick kind, add a line to the slide notes:
   `python -m pptfit note out/final.pptx N "what to check"`.

## How it works

| module | role |
|---|---|
| `extract.py` | every element with absolute geometry (groups expanded), text, sizes, fill and line, plus non-placeholder shapes of the layout and master (footer bars, logos, rules) |
| `units.py` | translation units; writes a chosen variant back into the deck |
| `render.py` | LibreOffice → PDF → words with boxes and rendered sizes; single-slide and background-only renders |
| `layout.py` | the original's guarantees per slide, and the check of a candidate against them |
| `textcheck.py` | accuracy checks: leftover source text, numbers, glossary, concepts, repeated words |
| `ops.py` | edits by element id: font scale, insets, resize, move, trim a line, autofit off |
| `fit.py` | search: for each offending element (or look-alike group) try moves, render, keep the best |
| `report.py` | side-by-side images and summary |

**Guarantees, checked on the render (`layout.py`)**

| code | meaning |
|---|---|
| INK_OUT | text left the visible box, panel or picture that held it |
| INK_HIT | text touched a picture, line, bar or other text it did not touch before (word by word; a line counts with its stroke) |
| INK_BG | text left the background area it sat on (bars and panels drawn into the master or background image) |
| BOX_HIT | a visible box now overlaps one it did not overlap |
| OFF_SLIDE | text outside the slide |
| FONT_FLOOR | rendered size below 80% of the original (titles 90%) |
| GROUP_SIZE, GROUP_BOX | look-alike boxes no longer share font size or box size |
| BROKEN_WORD | a word split across lines (a break at a hyphen is fine) |
| LINES | a free-standing one-line text (title, caption) now wraps |
| MOVED | a picture, line or other non-text element moved without a reason |

**Moves (`fit.py`)**: a shorter accurate variant, tighter insets, tighter line spacing
(never below 1.1x the font size), any lower font step (to the floor), growing the box into free space (several sizes; lines in the way are trimmed
in the same move), shifting the box (expensive), trimming a line. Look-alike boxes always
move together. Titles do not shrink per slide; a deck-level pass shrinks all titles of one
size together if needed. Candidates are scored by violation count, then total overshoot,
then cost (how far the result drifts from the original).

## Extending

- New kind of guarantee: record it in `layout.constraints()`, test it in `layout.check()`
  under a new code. The original must pass it against itself.
- New kind of move: add an edit to `ops.Deck`, a state field and its application in
  `Fitter.build()`, a candidate in `Fitter.moves()`, and its price in `cost()`.
- Always confirm on several decks (`bench/synth.py` builds a deck with common layouts)
  and by looking at the `cmp` images, not only at the counts.

## Limits

- Latin text set in a CJK font (SimSun, Microsoft YaHei, ...) is switched to Arial
  (`Fitter(latin=...)`): those fonts render English badly. Latin fonts are kept.
- Renders with LibreOffice; PowerPoint wraps text slightly differently. A few points of
  tolerance are built in, but check the final file in PowerPoint.
- Text inside images is not translated. Chart text is translated but its layout is not
  fitted; use `short` for chart strings.

## CI and benchmarks (GitHub Actions)

| workflow | what it does |
|---|---|
| `ci` | on every push: the guarantee tests and the synthetic deck |
| `corpus` | manual: public test decks (Apache POI test data), pseudo-translated (`bench/pseudo.py`, text 1.4x longer), fitted in parallel chunks; one summary table of slides fixed, slides left for a human, crashes |
| `fit` | manual: one deck given by temporary download links (deck, translations, glossary), fitted slide-parallel, then merged; result as a workflow artifact |

No deck is ever committed: real decks may be private, and public ones are downloaded at run time.
