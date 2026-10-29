# Task prompt template

Send the text below `---` to an agent that can run code, after replacing each `{...}` with
the option that fits your situation. Everything about *how* lives in this repository; the
prompt only says *what* to do this time.

| field | situation | replace with |
|---|---|---|
| `{CODE}` | the agent can reach the internet | `` clone `https://github.com/zolipx/pptfit` `` |
| | it cannot (turn its network access on if its settings allow, or) | attach the repository zip, https://github.com/zolipx/pptfit/archive/refs/heads/master.zip , and write `unzip the attached repository zip` |
| `{DECK}` | you attach the .pptx | `attached to this message` |
| | the agent downloads it | `at <url> (sha256 <hash>)` |
| `{SEND}` | default: both ways | `attach the zip to your reply and also upload it with <upload command> and put the link in your reply; if one way is blocked, use the other and say so` |
| | the agent cannot attach files | `upload it with <upload command> and put the link in your reply` |
| | the agent is offline | `attach the zip to your reply` |
| `{SOURCE_LANGUAGE}`, `{TARGET_LANGUAGE}` | | e.g. `Chinese`, `English` |

An upload command that needs no account: `` `curl -sS -F "files[]=@<file>" https://uguu.se/upload` (the link is in files[0].url, kept 3 hours; .zip files are refused, so base64 the zip into a .txt first) ``.

---

You are an unattended agent that can run code. Translate the PowerPoint deck {DECK} from
{SOURCE_LANGUAGE} into {TARGET_LANGUAGE} and keep its layout, using the reference code
pptfit.

The goal is a deck a person can finish in a few minutes, not a perfect one.
- Must be right (slow for a person to repair): positions and sizes of shapes, groups, flow
  charts, callouts and their leader lines, alignment; text inside the boxes and panels it sat
  in, and off the pictures, lines and bars it did not touch.
- May stay imperfect (a person fixes it in seconds): a wording a little long or short,
  punctuation, a label to retype. Note these for the reviewer instead of chasing them.

1. Get the code: {CODE}. Follow "Setup" in its README.
2. Get the deck; if a sha256 is given, check it.
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
   `glossary.json`, then {SEND}. Reply also with the contents of `summary.md` and one line
   per slide: what you saw and what you left for the reviewer.

Working rules: every message includes a tool call until the delivery is done (a message
with text only ends your turn); run commands longer than a minute in the background and
poll their log; write files in pieces of at most 250 lines.
