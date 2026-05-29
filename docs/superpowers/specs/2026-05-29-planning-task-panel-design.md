# Planning & Task Panel (D1) — Design

**Date:** 2026-05-29
**Status:** Approved design, ready for implementation plan
**Scope:** Surface the agent's `TodoWrite` planning as a live checklist in the
Draftspect taskpane, and nudge the agent to plan multi-step document work.

---

## Background & motivation

Today each open document has exactly one continuous chat thread (the daemon
runs one agent loop per `(host, document)` pane, resuming the SDK session by
`(host, workspace folder)`). The original request was "multiple chats" for
**task/topic organization** — keeping each piece of work focused instead of one
ever-growing thread.

While exploring that, we found the heavyweight interpretation (browser-style
tabs, one background agent loop per open chat) implies extensive new
infrastructure — a chat registry, a fan-out bridge protocol, and N concurrent
agent loops — i.e. "extremely heavy use" for the underlying need.

The underlying need is better served by capabilities the Agent SDK **already
provides first-class**:

- **Planning / task tracking:** `TodoWrite` (plus a `TaskCreate/Get/Update/List`
  toolset) — confirmed in `node_modules/@anthropic-ai/claude-agent-sdk/sdk-tools.d.ts`.
- **Subagents:** an `agents` option + Task/Agent tool, with background execution,
  `task_progress` events, and a `taskBudget` cost cap — confirmed in `sdk.d.ts`.

The daemon already runs with `systemPrompt: { preset: "claude_code" }` and **no**
`allowedTools` restriction (only `disallowedTools` for the flickery `word-mcp`
tools), so `TodoWrite` is **very likely already available to the agent right
now** — it is simply not surfaced in the taskpane UI nor encouraged by the
system prompt.

**This spec covers D1 only** (the planning & task panel). Subagents (D2) are
recorded as a future direction (see below) but not built here.

## The actual problem

> "Working on a document means several distinct tasks; I want to keep each one's
> context focused, run/track several pieces of work, and find past work later —
> without one giant messy thread, and without burning tokens for no reason."

D1 addresses the "see and track multi-step work in a focused way" part with the
lightest possible mechanism.

---

## Design

### Guiding principle

Do almost all the work in the **taskpane**; keep the daemon dumb. The daemon
already forwards everything we need. One chat, one loop — unchanged.

### 1. Data flow (mostly already exists)

The agent calls `TodoWrite` with the **full updated list** each time:

```ts
TodoWrite({ todos: [{ content: string,
                      status: "pending" | "in_progress" | "completed",
                      activeForm: string }] })
```

- **Live path:** the daemon already forwards tool calls as
  `assistant_event { event: "tool_use_announce", tool: "TodoWrite", input }`
  (`daemon/index.mjs`, `handleAgentMessage`, the `assistant`/`tool_use` branch).
  **No daemon change required.**
- **Replay path:** `daemon/transcript.mjs` (`eventsFromLine`) already emits
  `{ kind: "tool", name: "TodoWrite", input }` for each `tool_use` block.
  **No daemon change required.**

All new logic lives in `taskpane/shared/taskpane.js`: special-case the
`TodoWrite` tool name instead of rendering it as a generic 🔧 tool bubble.

> Optional cleanup (not required): the daemon could emit a typed
> `assistant_event { event: "todo_update", todos }` so the taskpane doesn't
> string-match a tool name in two places. Deferred — taskpane-only keeps the
> blast radius minimal. The implementation plan may revisit.

### 2. The plan panel (UI)

- **Placement:** a collapsible **"Plan" panel pinned just above the composer**,
  inside the Chat tab's `agent-status-row`/composer region — always visible while
  the agent works, so it does not scroll away with the message list.
- **Per item:** a status glyph (`☐` pending · `◐`/spinner `in_progress` · `☑`
  completed) and a label — `activeForm` while `in_progress`, else `content` —
  plus an overall `N/M` completed count in the header.
- **Collapsible:** header click toggles expanded/collapsed; default expanded
  while any item is `in_progress`, collapsible by the user. Collapsed state shows
  just the header + `N/M`.
- **Not** gated by the "Show diagnostics in chat" setting — this is a primary
  feature, not a diagnostic. (Render it outside the diagnostics-gated message
  area so the existing CSS gate does not hide it.)
- Plain text labels (no markdown rendering needed).

### 3. Coalescing (the core logic)

The agent calls `TodoWrite` repeatedly; each call is the **complete current
list**. The panel must **update in place**, never stack one bubble per call.

- A single pure helper drives both the live and replay paths:

  ```
  coalesceTodos(events) -> finalTodoState | null
  ```

  For replay: scan all events, return the **last** `TodoWrite` input's `todos`
  (or `null` if none / empty). Earlier `TodoWrite` events are skipped (not
  rendered as tool bubbles); every other tool renders unchanged.

  For live: each incoming `TodoWrite` replaces the panel's contents with the new
  `todos`.

- A new plan started in a later turn simply replaces the panel (the new
  `TodoWrite` carries a fresh list).

### 4. System-prompt nudge

Add a short section to `daemon/system-prompt.md` (shared base — applies to both
Word and Excel). Intent:

- For genuinely multi-step requests (multi-section redrafts, research-then-edit,
  multi-sheet/multi-range operations) lay out a brief plan with `TodoWrite` and
  keep it updated as work proceeds.
- Skip planning for trivial single-step asks — avoid over-planning.

Kept deliberately light so it does not turn every one-line request into a
ceremony.

### 5. Lifecycle & edge cases

- **Empty list:** `todos: []` → hide the panel.
- **All completed:** keep the panel visible (dimmed) so the user sees the final
  result; it is replaced when a new plan starts or rebuilt on chat/workspace
  switch via the normal `transcript_replay`.
- **Workspace switch / reopen:** the existing replay flow rebuilds the panel from
  the last `TodoWrite` in the resumed transcript — open plans survive a reopen
  with no extra persistence.
- **Many calls per turn:** handled by coalescing (§3).
- **Interleaving with other tools:** other tool announces still render in the
  message stream as today; only `TodoWrite` is diverted to the panel.

### 6. Risk & first implementation step

**Verify at runtime** (step 1 of the plan) that the SDK actually emits
`TodoWrite` tool calls under the current `claude_code` preset:

- Expected: it does, because no `allowedTools` allowlist is set (all preset tools
  allowed) and `TodoWrite` is a base preset tool.
- `todoFeatureEnabled` (sdk.d.ts) is a **terminal-UI display flag** (it sits among
  `showTurnDuration` / `showMessageTimestamps`), not a tool gate — irrelevant to
  our own render.
- Contingency: if `TodoWrite` is not emitted, investigate the preset toolset; a
  targeted `allowedTools` including `TodoWrite` (alongside the office tools) is a
  fallback, but note that switching from disallow-only to an allowlist is a
  behavior change that must preserve all currently-available tools — handle with
  care.

### 7. Testing

- **Unit:** `coalesceTodos(events)` is pure — test empty, single, multi-call,
  all-completed, and "new plan replaces old" cases.
- **Manual (Word + Excel):**
  1. Multi-step request → checklist appears pinned above the composer and updates
     in place (items flip pending → in_progress → completed).
  2. No duplicate 🔧 `TodoWrite` bubbles appear in the message stream.
  3. Close & reopen the panel/document → plan is rebuilt from replay.
  4. Trivial single-step request → no spurious plan panel.
  5. Panel visible regardless of the "Show diagnostics in chat" setting.

---

## Out of scope (recorded for the future)

- **D2 — Subagents (future wish).** Let the main agent delegate isolated
  sub-jobs to subagents, each with its own context window; optional background
  execution with `task_progress` rendered in the pane and a `taskBudget` cost
  cap. SDK-native via the `agents` option + Task tool. Naturally bounds
  parallelism to what the agent decides a task needs, and tears subagents down
  when done — far lighter than user-managed background top-level chats. To be
  designed in its own spec.
- **D3 — Multiple chats.** New-chat + history list (single active loop), or the
  heavyweight background-loop tab design. Not pursued; D1 (and later D2) address
  the underlying need at a fraction of the cost.

## Files likely touched (for the plan; not exhaustive)

- `taskpane/shared/taskpane.js` — TodoWrite special-casing, coalescing helper,
  panel render/update, replay integration.
- `taskpane/word/index.html` + `taskpane/excel/index.html` — the pinned panel
  markup above the composer (shared structure; both hosts use one `taskpane.js`).
- `taskpane/shared/styles.css` — panel styling, status glyphs, collapsed state.
- `daemon/system-prompt.md` — the planning nudge.
- Tests — `coalesceTodos` unit coverage.
