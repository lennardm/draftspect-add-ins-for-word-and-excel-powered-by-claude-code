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
      dp[i][j] = a[i] === b[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
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
