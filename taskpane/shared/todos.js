// Pure helpers for the planning/task panel.
//
// The agent drives planning via the SDK's `TodoWrite` tool. Each call carries
// the COMPLETE, current todo list (not a delta), forwarded to the taskpane as a
// tool call. These helpers normalize that raw input and coalesce a stream of
// transcript events down to the single latest list, so the live render path and
// the replay render path share one source of truth.
//
// No DOM here — keep it pure so it is unit-testable under `node --test`.

const VALID_STATUS = new Set(["pending", "in_progress", "completed"]);

// True iff a forwarded tool name is the planning tool.
export function isTodoWrite(name) {
  return name === "TodoWrite";
}

// Normalize a TodoWrite tool input ({ todos: [...] }) into a clean array of
// { content, status, activeForm }. Drops entries with no content; coerces an
// unknown status to "pending"; falls back activeForm → content when blank.
// Returns [] when there is nothing usable.
export function normalizeTodos(input) {
  const list = input && Array.isArray(input.todos) ? input.todos : [];
  const out = [];
  for (const t of list) {
    if (!t || typeof t !== "object") continue;
    const content = typeof t.content === "string" ? t.content : "";
    if (!content) continue;
    const status = VALID_STATUS.has(t.status) ? t.status : "pending";
    const activeForm = typeof t.activeForm === "string" && t.activeForm ? t.activeForm : content;
    out.push({ content, status, activeForm });
  }
  return out;
}

// Reduce a list of transcript replay events to the LATEST todo list, or null if
// the conversation contains no TodoWrite call. Replay events look like
// { kind: "tool", name, input }; anything else is ignored.
//   null = never planned (leave the panel as-is / hidden)
//   []   = planned then emptied (hide the panel)
export function coalesceTodos(events) {
  let latest = null;
  for (const ev of events || []) {
    if (ev && ev.kind === "tool" && isTodoWrite(ev.name)) {
      latest = normalizeTodos(ev.input);
    }
  }
  return latest;
}

// { done, total } completion counts for a normalized todo list.
export function todoProgress(todos) {
  const list = Array.isArray(todos) ? todos : [];
  const done = list.filter((t) => t.status === "completed").length;
  return { done, total: list.length };
}

// The label to show for an item: the present-continuous activeForm while it is
// in progress, otherwise its content. Accepts a normalized todo or a raw SDK
// todo — the fallbacks keep it safe if activeForm is missing on raw input.
export function todoLabel(todo) {
  if (!todo || typeof todo !== "object") return "";
  return todo.status === "in_progress" ? todo.activeForm || todo.content || "" : todo.content || "";
}
