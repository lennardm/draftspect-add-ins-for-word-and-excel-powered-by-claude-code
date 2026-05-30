# Edit Quality (atomic redlines + no leaked paragraph IDs) — Design

**Date:** 2026-05-30
**Status:** Approved design, ready for implementation plan
**Scope:** Two independent Word-editing quality fixes in Draftspect:
1. **#1 — Stop leaking internal paragraph IDs in chat** (prompt-only).
2. **#2 — Make track-changes redlines atomic** (word-level diff inside `office_replace_text`).

---

## Background & motivation

While using the merged planning panel, the user noticed two editing-quality issues in Word:

1. The agent refers to internal paragraph identifiers in chat (e.g. *"I fixed p2DE04F83"*) as if the user can act on them. The `daemon/system-prompt-word.md` already forbids this ("Referring to paragraphs in chat" → never show the `id`, refer by quoted snippet/heading), so this is the **model violating an existing rule**, not a missing rule.

2. A tiny change — e.g. deleting a single "a" — produces a redline that **strikes the whole sentence and re-inserts the corrected one**, making it hard to see what actually changed.

Root cause of #2 (confirmed by reading `taskpane/shared/tools-word.js → toolReplaceText`): the tool does `paragraph.search(find)` then `range.insertText(replace, Word.InsertLocation.replace)` — so with track-changes on, **the entire `find` span is deleted and the entire `replace` span inserted**. The redline granularity is exactly whatever the model puts in `find`/`replace`. And the model is *pushed* toward large spans because `search` replaces **all** matches: to delete one "a" it cannot pass `find:"a"` (that hits every "a"), so it passes a long, unique string (the whole sentence). So the bloat is partly forced by the tool's match-all design — a structural problem, not just model laziness.

## Decisions (from brainstorming)

- **Redline granularity: word-level.** Whole changed words are struck/inserted (`widget`→`apparatus` = delete `widget` + insert `apparatus`; `a`→`the` = delete `a` + insert `the`). Matches Word's own Compare and legal-redline convention. Not character-level.
- **#2 approach: word-diff *inside* `office_replace_text`** (not a new tool, not prompt-only). The model's interface is unchanged; the tool minimizes the redline internally.
- **#1 approach: prompt-only.** The id stays in the tool-call args (so the model keeps its location grounding on resume regardless); a blind display-strip is rejected because deleting an id token mid-sentence mangles grammar ("I fixed by tightening it").
- **Enabling API:** `Word.Range.getTextRanges([" "])` returns per-word sub-ranges of a range, giving precise word addressing without the "search matches every 'a'" problem.

---

## #2 — Word-diff inside `office_replace_text`

### Component 1: pure diff module — `taskpane/shared/word-diff.js`

A DOM-free, unit-testable module (same pattern as `todos.js` / `paths.js`).

```
wordDiff(oldText, newText) -> editScript
```

- **Tokenize** both strings on spaces into tokens (a token is a run of non-space characters; punctuation stays attached to its word, e.g. `housing.` is one token — consistent with how `getTextRanges([" "])` splits the live range).
- **Diff** the two token sequences with a longest-common-subsequence algorithm.
- **Emit** a minimal edit script as an ordered list of runs:
  - `{ op: "equal", tokens: [...] }`
  - `{ op: "delete", tokens: [...] }`
  - `{ op: "insert", tokens: [...] }`
  - (a word "change" is naturally represented as an adjacent `delete` + `insert`.)
- Identical input → all-`equal` script (caller treats as no-op).

This module owns all the algorithmic logic and is the only part with heavy unit tests.

### Component 2: application — `taskpane/shared/tools-word.js → toolReplaceText`

Replace the current "insert-text-replace over the whole match" with minimal per-word edits:

1. As today, locate matches: `paragraph.search(find, { matchCase, matchWholeWord })` → match ranges. (Multi-match behavior preserved — see below.)
2. For **each** match range:
   a. Compute `editScript = wordDiff(find, replace)`.
   b. If the script is all-`equal` (no real change), skip the match.
   c. Get per-word sub-ranges of the match via `matchRange.getTextRanges([" "], /* trimSpacing */ true)`. These align 1:1 with the tokenization of `find`.
   d. Walk the edit script and apply, as **tracked** edits, only the changed words:
      - `equal` token → leave its sub-range untouched.
      - `delete` token → delete that token's sub-range (and one adjacent separating space so the result has no double/leading/trailing space).
      - `insert` token → insert the new token text at the correct word boundary (with a separating space as needed).
   e. **Whitespace discipline:** after applying, the paragraph must read with single spaces, no doubled or orphaned spaces around edited words. The implementation must normalize the space adjacent to each delete/insert (this is the fiddly part — covered by the spike + manual tests).
3. Track-changes is controlled by the existing `withTrackChanges(context, track_changes, …)` wrapper — unchanged.

### Safety net (no regression)

If the per-word application cannot be performed cleanly for a given match — e.g. `getTextRanges` returns an unexpected shape, token counts don't align, or any error is thrown mid-apply — **fall back to the current behavior for that match**: `matchRange.insertText(replace, Word.InsertLocation.replace)`. Worst case is therefore exactly today's redline (whole-span), never worse. Log the fallback via the existing `diag()` channel so it's observable.

### Multi-match behavior (preserved)

The tool still supports a `find` that matches multiple places (e.g. "fix the same misspelling across every paragraph"). Each match independently gets the minimal word-diff treatment. The return shape (`total_replacements`, `per_paragraph`) is unchanged.

### Interface (unchanged)

`paragraph_ids`, `find`, `replace`, `match_case`, `whole_word`, `track_changes` — all unchanged. No new parameters. Minimization is automatic and internal. (No opt-out flag — YAGNI; the fallback covers the only failure mode.)

---

## #1 — Prompt-only reinforcement

Edit `daemon/system-prompt-word.md`, "Referring to paragraphs in chat" section:

- Lead with the rule emphatically and add a concrete ❌ example: *"❌ 'I fixed p2DE04F83' / 'highlighted p42' — the user can't map that to anything on the page."*
- Restate the ✅ alternatives already present (quoted snippet / heading / visible number).
- Add one sentence on *why it's safe to omit*: the `id` remains in the tool-call arguments, so the model keeps its location grounding on resume — it loses nothing by keeping the id out of prose.

No code, no display-side stripping (rejected: deleting an id token mid-sentence mangles grammar). Accept that prompt adherence is not 100%; this is the correct shape and lowest-risk.

---

## Testing

### Unit (pure `wordDiff`) — `tests/word-diff.test.mjs`
- Identical strings → all-`equal` (no-op).
- Single word changed: `"The widget comprises"` → `"The apparatus comprises"` → only `widget`→`apparatus`.
- Single word **deleted** (the reported case): `"comprises a housing"` → `"comprises housing"` → only `a` deleted.
- Single word inserted: `"comprises housing"` → `"comprises a housing"`.
- Multiple separate hunks in one string.
- Punctuation-attached tokens: `"housing."` → `"housing,"` (word-level: token replaced).
- Leading and trailing edits (first/last token changed).
- Whitespace-only / no-op difference → all-`equal`.

### Manual (Word) — `getTextRanges` application
This is the part that needs live Office verification (the spike, then acceptance):
1. **Spike first:** confirm `getTextRanges([" "], true)` returns one range per space-delimited token for a representative paragraph, and that deleting/replacing a single sub-range with track-changes on produces a clean single-word redline. Validate spacing behavior.
2. Delete one "a" mid-sentence → redline shows only "a" struck, spacing correct, rest untouched.
3. Swap one word → delete old word + insert new word only.
4. Mixed edit (a swap + a deletion in one `find`/`replace`) → only those words redlined.
5. "Fix all" across multiple matches → each minimally redlined.
6. Track-changes OFF → clean final text, single spaces, no artifacts.
7. Force the fallback (if feasible) → behaves as today (whole-span), no crash.

---

## Risks

- **`getTextRanges` semantics** (punctuation, spacing, interaction with tracked changes) are the one real unknown — spike before building the application layer. The whole-span **fallback** ensures any unhandled case degrades to today's behavior rather than breaking.
- **#1** is inherently best-effort (prompt adherence). Acceptable per decision.

## Out of scope (recorded for the future)

- **Approach B — `office_edit_paragraph(id, new_text)`:** model passes the full intended paragraph, tool word-diffs against current text. Could later *simplify* the surgical-vs-paragraph prompt decision tree wholesale (the agent always just supplies intended text; redlines are always minimal). Deferred — more than the reported pain needs now, and it carries a re-emit/transcription-drift cost for small edits. The `wordDiff` module built here would be directly reusable for it.

## Files likely touched (for the plan; not exhaustive)

- Create `taskpane/shared/word-diff.js` — pure diff module.
- Create `tests/word-diff.test.mjs` — unit tests.
- Modify `taskpane/shared/tools-word.js` — `toolReplaceText` application + fallback.
- Modify `daemon/system-prompt-word.md` — #1 reinforcement.
