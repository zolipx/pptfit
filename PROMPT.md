# Task prompt template

Fill in the `{...}` fields and send the text below to any agent that can run code
(Grok Build, ChatGPT, Claude, ...). Everything about *how* lives in this
repository; the prompt only says *what* to do this time.

- By hand: attach the .pptx to the message; `{DECK}` = `attached to this message`. Agents
  without internet access (e.g. ChatGPT's sandbox): also attach the repository as a zip,
  https://github.com/zolipx/pptfit/archive/refs/heads/master.zip
- Unattended tests: `{DECK}` = `at <url> (sha256 <hash>)`.
The result comes back twice: attached to the reply (where the chat allows files) and as a
download link (uguu.se, kept 3 hours).

---

You are an unattended agent that can run code. Translate the PowerPoint deck
{DECK} from {SOURCE_LANGUAGE} into {TARGET_LANGUAGE} and keep its layout, using the
reference code at {REPO_URL}.

The goal is a deck a person can finish in a few minutes, not a perfect one.
- Must be right (slow for a person to repair): positions and sizes of shapes, groups, flow
  charts, callouts and their leader lines, alignment; text inside the boxes and panels it sat
  in, and off the pictures, lines and bars it did not touch.
- May stay imperfect (a person fixes it in seconds): a wording a little long or short,
  punctuation, a label to retype. Note these for the reviewer instead of chasing them.

1. Get the code: `git clone {REPO_URL}`, or unzip the attached copy if you cannot reach it.
   Follow "Setup" in its README.
2. Get the deck: the attached file, or download the link and check its sha256.
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
   `glossary.json`, then send it both ways: attach the zip to your reply if this chat accepts
   files, and also upload it with `curl -sS -F "files[]=@<file>" https://uguu.se/upload`
   (JSON reply, the link is in files[0].url; .zip is refused, so base64 it into a .txt first)
   and put the link in your reply (if uploading is blocked, attach only and say so). Reply also with the contents of `summary.md` and one
   line per slide: what you saw and what you left for the reviewer.

Working rules: every message includes a tool call until the delivery is done (a message
with text only ends your turn); run commands longer than a minute in the background and
poll their log; write files in pieces of at most 250 lines.
