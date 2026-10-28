# Task prompt template

Fill in the `{...}` fields and send the text below to any agent that has a Linux
sandbox (Grok Build, ChatGPT, Claude, ...). Everything about *how* lives in this
repository; the prompt only says *what* to do this time.

---

You are an unattended agent with a Linux sandbox. Translate the PowerPoint deck at
{DECK_URL} (sha256 {DECK_SHA256}) from {SOURCE_LANGUAGE} into {TARGET_LANGUAGE} and keep
its layout, using the reference code at {REPO_URL}.

1. `git clone {REPO_URL}` and follow "Setup" in its README.
2. Download the deck and check the sha256.
3. Follow "Workflow" steps 1–5 of the README exactly: list the units, write
   `glossary.json` and `translations.json` yourself under the README's rules (give a
   `short2` for every unit in a tight spot: labels, flow boxes, titles), run
   `check-text` until it reports 0 problems, run `fit`, then open every image in `out/cmp/`.
4. For each slide listed under "needs human", add an accurate `short2` variant for the
   units involved or extend the code at the documented extension points, then run `fit`
   again. Never weaken a check to make it pass.
5. Deliver: zip `out/final.pptx`, `out/cmp/`, `out/summary.md`, `translations.json` and
   `glossary.json`, upload it with `{UPLOAD_COMMAND}`, and reply with the link and the
   contents of `summary.md`.

Working rules: every message includes a tool call until the delivery is done (a message
with text only ends your turn); run commands longer than a minute with `nohup` and poll
their log; write files in pieces of at most 250 lines.
