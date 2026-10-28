# Task prompt template

Fill in the `{...}` fields and send the text below to any agent that has a Linux
sandbox (Grok Build, ChatGPT, Claude, ...). Everything about *how* lives in this
repository; the prompt only says *what* to do this time.

---

You are an unattended agent with a Linux sandbox. Translate the PowerPoint deck at
{DECK_URL} (sha256 {DECK_SHA256}) from {SOURCE_LANGUAGE} into {TARGET_LANGUAGE} and keep
its layout, using the reference code at {REPO_URL}.

The goal is a deck a person can finish in a few minutes, not a perfect one.
- Must be right (slow for a person to repair): positions and sizes of shapes, groups, flow
  charts, callouts and their leader lines, alignment; text inside the boxes and panels it sat
  in, and off the pictures, lines and bars it did not touch.
- May stay imperfect (a person fixes it in seconds): a wording a little long or short,
  punctuation, a label to retype. Note these for the reviewer instead of chasing them.

1. `git clone {REPO_URL}` and follow "Setup" in its README.
2. Download the deck and check the sha256.
3. Follow "Workflow" steps 1–4 of the README exactly: list the units, write
   `glossary.json` and `translations.json` yourself under the README's rules (give a
   `short2` for every unit in a tight spot: labels, flow boxes, titles), run
   `check-text` until it reports 0 problems, run `fit`.
4. Review every slide yourself: open each `out/cmp/s{N}.png` (original left, result right).
   The checks miss things, so never call a slide fine because `summary.md` says so.
   - Slow for a person to repair: fix it with a tighter `short2` for the units involved or by
     extending the code at the documented extension points, then run `fit` again. Never
     weaken a check to make it pass.
   - Fixed by a person in seconds: `python -m pptfit note out/final.pptx N "<what to check>"`.
   Stop when only the second kind is left, or after two extra `fit` runs; note what remains.
5. Deliver: zip `out/final.pptx`, `out/cmp/`, `out/summary.md`, `translations.json` and
   `glossary.json`, upload it with `{UPLOAD_COMMAND}`, and reply with the link, the
   contents of `summary.md`, and one line per slide: what you saw and what you left for the
   reviewer.

Working rules: every message includes a tool call until the delivery is done (a message
with text only ends your turn); run commands longer than a minute with `nohup` and poll
their log; write files in pieces of at most 250 lines.
