# Edit Quality Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make Word track-changes redlines atomic (a one-word change redlines only that word) and stop the agent leaking internal paragraph IDs into chat.

**Architecture:** A new pure, unit-tested module `word-diff.js` computes a word-level diff (LCS) and groups it into change "hunks". `office_replace_text` (in `tools-word.js`) stops doing "delete-all-`find` + insert-all-`replace`" and instead applies only the changed hunks onto per-word sub-ranges obtained via `Word.Range.getTextRanges`. A whole-span fallback guarantees no regression. #1 is a prompt-only reinforcement.

**Tech Stack:** Vanilla ESM browser modules (`taskpane/shared/*.js`), Office.js Word API, `node:test`, plain markdown prompt.

**Branch:** `dev/edit-quality` (already created).
**Spec:** `docs/superpowers/specs/2026-05-30-edit-quality-design.md`

---

## File Structure

- **Create** `taskpane/shared/word-diff.js` — pure (no DOM): `tokenize`, `wordDiff` (segment list), `diffHunks` (application-ready hunks). Single responsibility: turn (oldText,newText) into a minimal, word-level edit description.
- **Create** `tests/word-diff.test.mjs` — `node:test` unit coverage.
- **Modify** `taskpane/shared/tools-word.js` — rewrite the *application* inside `toolReplaceText` to apply `diffHunks` onto `getTextRanges` sub-ranges, with a whole-span fallback. Interface unchanged.
- **Modify** `daemon/system-prompt-word.md` — strengthen the "Referring to paragraphs in chat" rule (#1).

**Important constraints for every task:**
- Work on branch `dev/edit-quality`. Do NOT switch/create branches.
- NEVER stage `package-lock.json` (it has an unrelated working-tree change). Stage only the files each task names.
- Tests run via `npm test` (`node --test tests/*.test.mjs`). Format via `npx prettier`.

---

## Task 1: Pure word-diff module (`word-diff.js`)

**Files:**
- Create: `taskpane/shared/word-diff.js`
- Test: `tests/word-diff.test.mjs`

- [ ] **Step 1: Write the failing tests**

Create `tests/word-diff.test.mjs`:

```js
// Unit tests for taskpane/shared/word-diff.js — pure word-level diff used to
// make track-changes redlines minimal. No DOM, no Office.

import { test } from "node:test";
import assert from "node:assert/strict";

import { tokenize, wordDiff, diffHunks } from "../taskpane/shared/word-diff.js";

test("tokenize splits on whitespace, keeps punctuation attached", () => {
  assert.deepEqual(tokenize("The widget comprises a housing."), [
    "The", "widget", "comprises", "a", "housing.",
  ]);
  assert.deepEqual(tokenize("  spaced   out \n"), ["spaced", "out"]);
  assert.deepEqual(tokenize(""), []);
});

test("wordDiff: identical text → all equal", () => {
  assert.deepEqual(wordDiff("a b c", "a b c"), [{ op: "equal", tokens: ["a", "b", "c"] }]);
});

test("wordDiff: single word changed → delete then insert", () => {
  assert.deepEqual(wordDiff("The widget comprises", "The apparatus comprises"), [
    { op: "equal", tokens: ["The"] },
    { op: "delete", tokens: ["widget"] },
    { op: "insert", tokens: ["apparatus"] },
    { op: "equal", tokens: ["comprises"] },
  ]);
});

test("wordDiff: single word deleted (the reported 'a' case)", () => {
  assert.deepEqual(wordDiff("comprises a housing", "comprises housing"), [
    { op: "equal", tokens: ["comprises"] },
    { op: "delete", tokens: ["a"] },
    { op: "equal", tokens: ["housing"] },
  ]);
});

test("wordDiff: single word inserted", () => {
  assert.deepEqual(wordDiff("comprises housing", "comprises a housing"), [
    { op: "equal", tokens: ["comprises"] },
    { op: "insert", tokens: ["a"] },
    { op: "equal", tokens: ["housing"] },
  ]);
});

test("diffHunks: identical → no hunks", () => {
  assert.deepEqual(diffHunks("a b c", "a b c"), []);
});

test("diffHunks: word change → one replace hunk at the right index", () => {
  assert.deepEqual(diffHunks("The widget comprises", "The apparatus comprises"), [
    { oldStart: 1, oldCount: 1, insertTokens: ["apparatus"] },
  ]);
});

test("diffHunks: word deletion → delete hunk (no inserts)", () => {
  assert.deepEqual(diffHunks("comprises a housing", "comprises housing"), [
    { oldStart: 1, oldCount: 1, insertTokens: [] },
  ]);
});

test("diffHunks: pure insertion → zero-width hunk", () => {
  assert.deepEqual(diffHunks("comprises housing", "comprises a housing"), [
    { oldStart: 1, oldCount: 0, insertTokens: ["a"] },
  ]);
});

test("diffHunks: two separate changes → two hunks, indices into OLD tokens", () => {
  // old: The big red fox   new: The small red cat
  assert.deepEqual(diffHunks("The big red fox", "The small red cat"), [
    { oldStart: 1, oldCount: 1, insertTokens: ["small"] },
    { oldStart: 3, oldCount: 1, insertTokens: ["cat"] },
  ]);
});

test("diffHunks: trailing-word change", () => {
  assert.deepEqual(diffHunks("housing.", "housing,"), [
    { oldStart: 0, oldCount: 1, insertTokens: ["housing,"] },
  ]);
});
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npm test`
Expected: FAIL — `Cannot find module '../taskpane/shared/word-diff.js'`. Existing tests still pass.

- [ ] **Step 3: Write the implementation**

Create `taskpane/shared/word-diff.js`:

```js
// Pure word-level diff for minimal track-changes redlines.
//
// office_replace_text used to delete the entire `find` span and insert the
// entire `replace` span, so changing one word redlined the whole sentence.
// This module computes a word-level diff so the consumer (tools-word.js) can
// redline only the words that actually changed. No DOM — unit-testable under
// `node --test`. Tokenization matches Word's getTextRanges([" "]) split:
// whitespace-delimited, punctuation stays attached to its word.

// Whitespace-delimited tokens; empties dropped. "housing." is one token.
export function tokenize(text) {
  return String(text)
    .split(/\s+/)
    .filter((t) => t.length > 0);
}

// Longest-common-subsequence diff over two token arrays → ordered segments:
//   { op: "equal" | "delete" | "insert", tokens: string[] }
// "delete" = in old only; "insert" = in new only. A word change surfaces as a
// delete segment immediately followed by an insert segment.
export function wordDiff(oldText, newText) {
  const a = tokenize(oldText);
  const b = tokenize(newText);
  const n = a.length;
  const m = b.length;
  // dp[i][j] = LCS length of a[i:] and b[j:].
  const dp = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      dp[i][j] =
        a[i] === b[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
    }
  }
  const raw = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) {
      raw.push(["equal", a[i]]);
      i++;
      j++;
    } else if (dp[i + 1][j] >= dp[i][j + 1]) {
      raw.push(["delete", a[i]]);
      i++;
    } else {
      raw.push(["insert", b[j]]);
      j++;
    }
  }
  while (i < n) {
    raw.push(["delete", a[i]]);
    i++;
  }
  while (j < m) {
    raw.push(["insert", b[j]]);
    j++;
  }
  // Coalesce consecutive same-op tokens into one segment.
  const segments = [];
  for (const [op, tok] of raw) {
    const last = segments[segments.length - 1];
    if (last && last.op === op) last.tokens.push(tok);
    else segments.push({ op, tokens: [tok] });
  }
  return segments;
}

// Group the segment list into application-ready change hunks. Each hunk:
//   { oldStart, oldCount, insertTokens }
// where oldStart/oldCount index into the OLD token array (i.e. the per-word
// ranges the consumer gets from the matched range), and insertTokens are the
// new words for that hunk. A pure deletion has insertTokens = []; a pure
// insertion has oldCount = 0. Equal runs produce no hunk.
export function diffHunks(oldText, newText) {
  const segs = wordDiff(oldText, newText);
  const hunks = [];
  let oldIdx = 0;
  let k = 0;
  while (k < segs.length) {
    if (segs[k].op === "equal") {
      oldIdx += segs[k].tokens.length;
      k++;
      continue;
    }
    const oldStart = oldIdx;
    let oldCount = 0;
    const insertTokens = [];
    while (k < segs.length && segs[k].op !== "equal") {
      if (segs[k].op === "delete") {
        oldCount += segs[k].tokens.length;
        oldIdx += segs[k].tokens.length;
      } else {
        insertTokens.push(...segs[k].tokens);
      }
      k++;
    }
    hunks.push({ oldStart, oldCount, insertTokens });
  }
  return hunks;
}
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `npm test`
Expected: PASS — all `word-diff.test.mjs` tests green; existing tests still green.

- [ ] **Step 5: Syntax + format**

Run: `node --check taskpane/shared/word-diff.js`
Run: `npx prettier --check taskpane/shared/word-diff.js tests/word-diff.test.mjs` (run `--write` then re-check if needed).

- [ ] **Step 6: Commit**

```bash
git add taskpane/shared/word-diff.js tests/word-diff.test.mjs
git commit -m "Add pure word-diff module for minimal redlines (TDD)"
```

---

## Task 2: Office.js `getTextRanges` spike (manual, in Word)

**Purpose:** the one real unknown is how `Word.Range.getTextRanges` splits a range and handles spacing, and how tracked sub-edits read. This task gathers that evidence BEFORE writing the application code in Task 3. No production code is committed here — it's a throwaway probe + written findings.

**Files:**
- Temporary: `taskpane/shared/_spike-gettextranges.js` (deleted at end of task — do NOT commit)

- [ ] **Step 1: Add a temporary probe to a tool you can trigger**

The simplest way to run Office.js code is to temporarily hijack an existing tool. In `taskpane/shared/tools-word.js`, at the very top of `toolGetSelection` (it's called whenever the agent reads the selection), add a one-shot probe that runs against the user's selected paragraph and logs to the webview console. Add this block as the first lines inside the `Word.run` callback of `toolGetSelection`:

```js
// TEMP SPIKE — remove after Task 2. Logs getTextRanges behavior for the
// selection's first paragraph so we can design the word-diff application.
try {
  const sel = context.document.getSelection();
  const para = sel.paragraphs.getFirst();
  const r = para.getRange();
  const wrTrim = r.getTextRanges([" "], true);
  const wrKeep = r.getTextRanges([" "], false);
  wrTrim.load("items/text");
  wrKeep.load("items/text");
  await context.sync();
  console.warn("[spike] trim=true tokens:", JSON.stringify(wrTrim.items.map((x) => x.text)));
  console.warn("[spike] trim=false tokens:", JSON.stringify(wrKeep.items.map((x) => x.text)));
} catch (e) {
  console.warn("[spike] getTextRanges probe failed:", e && e.message);
}
```

- [ ] **Step 2: Run it in Word and capture output**

Run: `npm start`, open Word, open the task pane DevTools console (right-click the pane → Inspect, or the host's webview devtools), select a sentence like `The widget comprises a housing.`, and ask the agent anything that reads the selection (e.g. "what's selected?"). Read the two `[spike]` console lines.

- [ ] **Step 3: Answer these questions in writing (paste answers into the commit message of Task 3, and into a comment block at the top of the application code):**

1. With `trimSpacing=true`, is there exactly one range per whitespace-delimited token, and does `item.text` equal the token WITHOUT surrounding spaces? (Expected: yes.)
2. With `trimSpacing=false`, does each `item.text` include a trailing (or leading) space? Which side? (This decides how deletions avoid double-spaces.)
3. Does the token count from `getTextRanges([" "], true)` equal `tokenize(text).length` from `word-diff.js` for a punctuated sentence? (If not, note the discrepancy — Task 3's count-guard will fall back in those cases.)

- [ ] **Step 4: Decide the spacing recipe**

Based on the answers, record the chosen recipe for Task 3 (one of):
- **(R1) trim=true + manual space fix:** operate on trimmed word ranges; after a deletion, delete one adjacent space; after an insertion, add one space. OR
- **(R2) trim=false:** each range carries its delimiter space, so deleting a range deletes its space too; replacement text carries matching spacing.

Write one or two sentences naming the recipe and why. This is the input to Task 3.

- [ ] **Step 5: Remove the probe**

Delete the TEMP SPIKE block you added to `toolGetSelection`. Confirm `git diff taskpane/shared/tools-word.js` is empty (no probe left behind). Do not commit anything in this task.

Run: `git status --short` → expect only the pre-existing `package-lock.json` modification (and nothing in tools-word.js).

---

## Task 3: Apply the diff in `office_replace_text` (with fallback)

**Files:**
- Modify: `taskpane/shared/tools-word.js` — the body of `toolReplaceText` (currently the `for (const p of probes)` apply loop, ~lines 543-558).

**Context:** `toolReplaceText` already resolves target paragraphs and runs `paragraph.search(find, …)` to get match ranges (the `probes` array, each `{ id, results }` where `results.items` are the match ranges). Today it does, per match: `r.insertText(replace, Word.InsertLocation.replace)`. We replace ONLY that application step. Everything above it (validation, `withTrackChanges`, target resolution, search) stays. Use the spacing recipe chosen in Task 2.

- [ ] **Step 1: Import the diff helper**

At the top of `taskpane/shared/tools-word.js`, add (after the existing `/* global Word */` header comment block, alongside any other imports — if there are none, add it as the first `import`):

```js
import { tokenize, diffHunks } from "./word-diff.js";
```

- [ ] **Step 2: Add a helper that applies hunks to one matched range, with fallback**

Add this function near `toolReplaceText` in `tools-word.js`. It encodes recipe **R2 (trim=false)** as the reference; if Task 2 chose R1, adjust the marked spacing lines accordingly. The count-guard + try/catch guarantee a safe fallback to today's whole-span replace.

```js
// Apply a word-level diff of `find`→`replace` onto a single matched range so
// the track-changes redline touches only the words that actually changed,
// instead of striking the whole match. Returns true on success; false if it
// bailed (caller then does the whole-span fallback). See word-diff.js.
//
// Task 2 spike recipe: R2 — getTextRanges([" "], false) returns one range per
// word INCLUDING its trailing delimiter space, so deleting a word range also
// removes its space (no double-spacing), and replacement text re-adds a single
// trailing space to match. (If the spike chose R1, swap the marked lines.)
async function applyWordDiffToRange(context, matchRange, find, replace) {
  const hunks = diffHunks(find, replace);
  if (hunks.length === 0) return true; // nothing actually changed

  const wordRanges = matchRange.getTextRanges([" "], false); // R2: keep spacing
  wordRanges.load("items");
  await context.sync();

  const items = wordRanges.items;
  // Count-guard: if Word's tokenization disagrees with ours, bail to fallback.
  if (items.length !== tokenize(find).length) return false;

  // Apply hunks in REVERSE order so edits don't shift the indices of
  // not-yet-applied (earlier) ranges.
  for (let h = hunks.length - 1; h >= 0; h--) {
    const { oldStart, oldCount, insertTokens } = hunks[h];
    const insText = insertTokens.join(" ");
    if (oldCount > 0) {
      // Replace or delete a span of existing word ranges.
      const span =
        oldCount === 1
          ? items[oldStart]
          : items[oldStart].expandTo(items[oldStart + oldCount - 1]);
      // R2: deleted ranges carried their trailing spaces; for a replacement,
      // re-add a single trailing space so the following word stays separated.
      const replacement = insertTokens.length > 0 ? insText + " " : "";
      span.insertText(replacement, Word.InsertLocation.replace);
    } else {
      // Pure insertion before the word currently at oldStart (or at the end).
      const text = insText + " ";
      if (oldStart < items.length) {
        items[oldStart].insertText(text, Word.InsertLocation.before);
      } else {
        items[items.length - 1].insertText(" " + insText, Word.InsertLocation.after);
      }
    }
  }
  await context.sync();
  return true;
}
```

- [ ] **Step 3: Use the helper in `toolReplaceText`'s apply loop**

Replace the current apply loop (the block that reads):

```js
      let totalReplaced = 0;
      const perParagraph = [];
      for (const p of probes) {
        const matches = p.results.items.length;
        for (const r of p.results.items) {
          r.insertText(replace, Word.InsertLocation.replace);
        }
        perParagraph.push({ paragraph_id: p.id, replacements: matches });
        totalReplaced += matches;
      }
      await context.sync();
```

with:

```js
      let totalReplaced = 0;
      const perParagraph = [];
      for (const p of probes) {
        const matches = p.results.items.length;
        for (const r of p.results.items) {
          let minimal = false;
          try {
            minimal = await applyWordDiffToRange(context, r, find, replace);
          } catch (err) {
            console.warn("[tools-word] word-diff apply failed; using whole-span replace:", err && err.message);
            minimal = false;
          }
          if (!minimal) {
            // Fallback: today's behavior — whole-span replace. Never worse.
            r.insertText(replace, Word.InsertLocation.replace);
          }
        }
        perParagraph.push({ paragraph_id: p.id, replacements: matches });
        totalReplaced += matches;
      }
      await context.sync();
```

(Note: `applyWordDiffToRange` does its own `context.sync()`; the trailing `await context.sync()` after the loop is harmless and covers the fallback `insertText` calls.)

- [ ] **Step 4: Syntax check + existing tests**

Run: `node --check taskpane/shared/tools-word.js`
Run: `npm test`  (expect all green — this file isn't unit-tested, but the suite must not regress)
Run: `npx prettier --check taskpane/shared/tools-word.js` (run `--write` then re-check if needed).

- [ ] **Step 5: Commit**

```bash
git add taskpane/shared/tools-word.js
git commit -m "Apply word-level diff in office_replace_text for atomic redlines

Redline only the words that changed (via getTextRanges) instead of the whole
find/replace span. Count-guard + try/catch fall back to today's whole-span
replace so behavior never regresses. Spike recipe: <paste Task 2 recipe>."
```

---

## Task 4: #1 — strengthen the paragraph-ID prompt rule

**Files:**
- Modify: `daemon/system-prompt-word.md` — the "## Referring to paragraphs in chat" section.

- [ ] **Step 1: Edit the section**

Find the existing section heading and its first paragraph:

```
## Referring to paragraphs in chat

Paragraph tool responses carry an internal `id` (a `uniqueLocalId` hex string or an index like `p7`). That `id` is the stable handle for tool calls — keep using it in `paragraph_ids`, `ids`, etc. **But never show it to the user in chat.** "I'll fix p CD84E50D" is opaque — the user can't map it to anything on the page.
```

Replace ONLY that paragraph (leave the bullet examples below it intact) with:

```
## Referring to paragraphs in chat

**Hard rule, every turn: never write an internal paragraph `id` in chat.** Tool responses carry an `id` (a `uniqueLocalId` hex string like `2DE04F83`, or an index like `p7`). It is the stable handle for tool calls — keep using it in `paragraph_ids`, `ids`, etc. — but it is **meaningless to the user**, who cannot map `2DE04F83` to anything on the page. Writing "I fixed 2DE04F83" or "highlighted p42" is a bug, not a status update.

You lose nothing by omitting it: the `id` stays in your tool-call arguments and transcript, so on later turns you still know exactly which paragraph you touched. The user just never sees the raw handle. Always refer to a paragraph by something they can see (see below).
```

- [ ] **Step 2: Verify**

Run: `npx prettier --check daemon/system-prompt-word.md` (run `--write` then re-check if needed; if prettier doesn't target it, skip — it's prose).
Confirm the ✅ bullet examples that previously followed the paragraph are still present and unchanged (read the section).

- [ ] **Step 3: Commit**

```bash
git add daemon/system-prompt-word.md
git commit -m "Strengthen rule against showing internal paragraph ids in chat"
```

---

## Task 5: Manual acceptance in Word

No code changes — the live verification of the redline behavior (the `getTextRanges` application can only be confirmed in Word).

- [ ] **Step 1: Launch**

Run: `npm start`, open Word, open a doc with track changes ON (Setup → Track changes: Always). Use a FRESH workspace folder if a prior session might be poisoned (see the OpenRouter-resume note).

- [ ] **Step 2: One-word deletion (the reported case)**

Select/point at a sentence like `The device comprises a housing.` and ask: "delete the word 'a' in this sentence." 
Expected redline: only **a** struck through; spacing stays single; the rest of the sentence is untouched (NOT a full-sentence strike+reinsert).

- [ ] **Step 3: One-word substitution**

Ask: "change 'device' to 'apparatus' here." 
Expected: only **device** struck + **apparatus** inserted; everything else untouched.

- [ ] **Step 4: Mixed edit in one sentence**

Ask: "in this sentence change 'device' to 'apparatus' and delete 'a'." 
Expected: two minimal redlines (the word swap and the deletion), nothing else struck.

- [ ] **Step 5: "Fix all" across matches**

In a doc where a word is misspelled in several paragraphs, ask: "fix every 'recieve' to 'receive'." 
Expected: each occurrence redlined as a minimal word swap; count reported.

- [ ] **Step 6: Track changes OFF**

Set Track changes: Never; repeat Step 3. Expected: clean final text, single spaces, no stray artifacts.

- [ ] **Step 7: ID leakage (#1)**

Ask for a multi-paragraph review/edit and read the chat. Expected: the agent refers to paragraphs by quoted text or heading, NOT by `2DE04F83`/`p7` handles.

- [ ] **Step 8: Commit any prettier fixups**

```bash
git status   # expect clean if all prior tasks committed cleanly
```

---

## Notes for the implementer

- **Branch `dev/edit-quality`; never stage `package-lock.json`.**
- Task 2 (spike) **gates** Task 3's exact spacing lines — do Task 2 first and record the recipe.
- The **fallback** is load-bearing: if anything about `getTextRanges` is off, `office_replace_text` must still work (whole-span replace), just less minimally. Verify the fallback path doesn't double-apply (it only runs when `applyWordDiffToRange` returns false / throws).
- `office_replace_text`'s public interface and return shape (`total_replacements`, `per_paragraph`) are unchanged.
- The taskpane has no daemon `diag()`; use `console.warn` for the fallback log (as written).

## Out of scope (recorded)
- **Approach B — `office_edit_paragraph(id, new_text)`:** future simplification; would reuse `word-diff.js` directly. Not in this plan.
