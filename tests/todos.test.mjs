// Unit tests for taskpane/shared/todos.js — pure normalization + coalescing of
// TodoWrite data. No DOM, no Office, no SDK.

import { test } from "node:test";
import assert from "node:assert/strict";

import {
  isTodoWrite,
  normalizeTodos,
  coalesceTodos,
  todoProgress,
  todoLabel,
} from "../taskpane/shared/todos.js";

test("isTodoWrite matches only the planning tool", () => {
  assert.equal(isTodoWrite("TodoWrite"), true);
  assert.equal(isTodoWrite("Read"), false);
  assert.equal(isTodoWrite(undefined), false);
  assert.equal(isTodoWrite(null), false);
});

test("normalizeTodos keeps valid items and fills activeForm", () => {
  const out = normalizeTodos({
    todos: [
      { content: "Read §3", status: "completed", activeForm: "Reading §3" },
      { content: "Redraft intro", status: "in_progress", activeForm: "Redrafting intro" },
      { content: "Fix citations", status: "pending", activeForm: "" },
    ],
  });
  assert.deepEqual(out, [
    { content: "Read §3", status: "completed", activeForm: "Reading §3" },
    { content: "Redraft intro", status: "in_progress", activeForm: "Redrafting intro" },
    // activeForm falls back to content when blank.
    { content: "Fix citations", status: "pending", activeForm: "Fix citations" },
  ]);
});

test("normalizeTodos drops malformed entries and coerces unknown status", () => {
  const out = normalizeTodos({
    todos: [
      null,
      { status: "pending" }, // no content → dropped
      { content: "", status: "pending" }, // empty content → dropped
      { content: "Do thing", status: "weird" }, // unknown status → pending
    ],
  });
  assert.deepEqual(out, [{ content: "Do thing", status: "pending", activeForm: "Do thing" }]);
});

test("normalizeTodos returns [] for missing/!array input", () => {
  assert.deepEqual(normalizeTodos(undefined), []);
  assert.deepEqual(normalizeTodos({}), []);
  assert.deepEqual(normalizeTodos({ todos: "nope" }), []);
});

test("coalesceTodos returns null when no TodoWrite event exists", () => {
  assert.equal(
    coalesceTodos([
      { kind: "user", text: "hi" },
      { kind: "tool", name: "Read", input: { file: "a" } },
    ]),
    null,
  );
  assert.equal(coalesceTodos([]), null);
  assert.equal(coalesceTodos(undefined), null);
});

test("coalesceTodos picks the LAST TodoWrite, ignoring others", () => {
  const out = coalesceTodos([
    null,
    { kind: "tool", name: "TodoWrite", input: { todos: [{ content: "A", status: "pending" }] } },
    { kind: "assistant", text: "working" },
    {
      kind: "tool",
      name: "TodoWrite",
      input: {
        todos: [
          { content: "A", status: "completed" },
          { content: "B", status: "in_progress" },
        ],
      },
    },
  ]);
  assert.deepEqual(out, [
    { content: "A", status: "completed", activeForm: "A" },
    { content: "B", status: "in_progress", activeForm: "B" },
  ]);
});

test("coalesceTodos returns [] when the last TodoWrite emptied the list", () => {
  assert.deepEqual(
    coalesceTodos([
      { kind: "tool", name: "TodoWrite", input: { todos: [{ content: "A", status: "pending" }] } },
      { kind: "tool", name: "TodoWrite", input: { todos: [] } },
    ]),
    [],
  );
});

test("todoProgress counts completed vs total", () => {
  assert.deepEqual(todoProgress([]), { done: 0, total: 0 });
  assert.deepEqual(todoProgress(null), { done: 0, total: 0 });
  assert.deepEqual(
    todoProgress([
      { content: "A", status: "completed", activeForm: "A" },
      { content: "B", status: "in_progress", activeForm: "B" },
      { content: "C", status: "pending", activeForm: "C" },
    ]),
    { done: 1, total: 3 },
  );
});

test("todoLabel shows activeForm only while in_progress", () => {
  assert.equal(
    todoLabel({ content: "Fix it", status: "in_progress", activeForm: "Fixing it" }),
    "Fixing it",
  );
  assert.equal(
    todoLabel({ content: "Fix it", status: "completed", activeForm: "Fixing it" }),
    "Fix it",
  );
  assert.equal(todoLabel(null), "");
});
