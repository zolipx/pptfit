# Task prompt

Attach the .pptx to a message, paste the prompt below unchanged, and send it to any agent that
can run code. The prompt is the same for every deck and every agent; it handles the cases
itself. Everything about *how* lives in this repository.

```text
You are an unattended agent that can run code. Translate the PowerPoint deck in this message
into English and keep its layout, using the reference code pptfit.

The goal is a deck a person can finish in a few minutes, not a perfect one.
- Must be right (slow for a person to repair): positions and sizes of shapes, groups, flow
  charts, callouts and their leader lines, alignment; text inside the boxes and panels it sat
  in, and off the pictures, lines and bars it did not touch.
- May stay imperfect (a person fixes it in seconds): a wording a little long or short,
  punctuation, a label to retype. Note these for the reviewer instead of chasing them.

1. Get the code: git clone https://github.com/zolipx/pptfit . If you cannot reach it, use the
   repository zip attached to this message. If there is none, stop and ask the person to turn
   on your internet access or to attach
   https://github.com/zolipx/pptfit/archive/refs/heads/master.zip . Then follow "Setup" in its
   README.
2. Get the deck: the .pptx attached to this message; if the message gives a link instead,
   download it, and check the sha256 if one is given.
3. Follow "Workflow" steps 1-4 of the README exactly: list the units, write glossary.json and
   translations.json yourself under the README's rules (give a short2 for every unit in a
   tight spot: labels, flow boxes, titles), run check-text until it reports 0 problems, run
   fit.
4. Review every slide yourself: open each out/cmp/s<N>.png (original left, result right).
   The checks miss things, so never call a slide fine because summary.md says so.
   - Slow for a person to repair: fix it with a tighter short2 for the units involved or by
     extending the code at the documented extension points, then run fit again. Never weaken
     a check to make it pass.
   - Fixed by a person in seconds: python -m pptfit note out/final.pptx <N> "<what to check>"
   Stop when only the second kind is left, or after two extra fit runs; note what remains.
5. Deliver: zip out/final.pptx, out/cmp/, out/summary.md, translations.json and
   glossary.json. Send it both ways: attach the zip to your reply, and also upload it with
   curl -sS -F "files[]=@<file>" https://uguu.se/upload (the link is in files[0].url, kept
   3 hours; .zip files are refused, so base64 the zip into a .txt first) and put the link in
   your reply. If one way is blocked, use the other and say so. Reply also with the contents
   of summary.md and one line per slide: what you saw and what you left for the reviewer.

Working rules: every message includes a tool call until the delivery is done (a message with
text only ends your turn); run commands longer than a minute in the background and poll their
log; write files in pieces of at most 250 lines.
```
