// Unit tests for taskpane/shared/word-diff.js — pure word-level diff used to
// make track-changes redlines minimal. No DOM, no Office.

import { test } from "node:test";
import assert from "node:assert/strict";

import { tokenize, wordDiff, diffHunks } from "../taskpane/shared/word-diff.js";

test("tokenize splits on whitespace, keeps punctuation attached", () => {
  assert.deepEqual(tokenize("The widget comprises a housing."), [
    "The",
    "widget",
    "comprises",
    "a",
    "housing.",
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

test("wordDiff: fully different inputs → all-delete then all-insert", () => {
  assert.deepEqual(wordDiff("a b", "x y"), [
    { op: "delete", tokens: ["a", "b"] },
    { op: "insert", tokens: ["x", "y"] },
  ]);
});

test("diffHunks: empty → empty inputs → no hunks", () => {
  assert.deepEqual(diffHunks("", ""), []);
});

test("diffHunks: insertion from empty / deletion to empty", () => {
  assert.deepEqual(diffHunks("", "new words"), [
    { oldStart: 0, oldCount: 0, insertTokens: ["new", "words"] },
  ]);
  assert.deepEqual(diffHunks("old words", ""), [{ oldStart: 0, oldCount: 2, insertTokens: [] }]);
});
