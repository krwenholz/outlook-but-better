// Action table: what each action means and how it executes. Pure data.
//
// Strategies (PLAN.md, Keymap layer):
//   passthrough – Outlook's Gmail shortcut mode already does it; let the key through.
//   drive       – click a control found through the selector registry. React honours click().
//                 (Dispatching Outlook's own shortcut as a synthetic key event did nothing on
//                 the live site in Nov 2026, so there is no keystroke strategy.)
//   toggle      – click a toggle button (Read / Unread) only if the selection isn't already in
//                 the wanted state.
//   extension   – Overlook's own UI (palette, help).
//   todo        – not decided yet; the key is swallowed and a hint says so.
//
// Outlook's Gmail-mode keys come from its own ribbon tooltips in the Oct 2026 capture:
// C new mail, # delete, E archive, V move, ⇧I/⇧U read/unread, B snooze, L categorize,
// R/A/F reply/reply all/forward, ⌘Z undo.

import type { SelectorName } from "../selectors.ts";

export type ExtensionCommand = "help" | "palette-go";

export type Strategy =
  | { kind: "passthrough" }
  | { kind: "drive"; target: SelectorName }
  | { kind: "toggle"; target: SelectorName; want: "read" | "unread" }
  | { kind: "extension"; command: ExtensionCommand }
  | { kind: "todo" };

export interface Action {
  description: string;
  strategy: Strategy;
}

const passthrough: Strategy = { kind: "passthrough" };
const todo: Strategy = { kind: "todo" };
const drive = (target: SelectorName): Strategy => ({ kind: "drive", target });
const extension = (command: ExtensionCommand): Strategy => ({ kind: "extension", command });
const readState = (want: "read" | "unread"): Strategy => ({
  kind: "toggle",
  target: "readUnreadButton",
  want,
});

export const ACTIONS = {
  compose: { description: "Compose", strategy: passthrough },
  search: { description: "Focus search", strategy: passthrough },
  help: { description: "Show keyboard shortcuts", strategy: extension("help") },
  next: { description: "Next conversation", strategy: passthrough },
  previous: { description: "Previous conversation", strategy: passthrough },
  open: { description: "Open conversation", strategy: passthrough },
  select: { description: "Select or deselect", strategy: passthrough },
  selectRange: { description: "Select range since last selection", strategy: todo },
  selectAll: { description: "Select all", strategy: todo },
  selectNone: { description: "Select none", strategy: todo },
  archive: { description: "Archive", strategy: drive("archiveButton") },
  archivePrevious: { description: "Archive, then open previous", strategy: todo },
  archiveNext: { description: "Archive, then open next", strategy: todo },
  delete: { description: "Delete", strategy: drive("deleteButton") },
  junk: { description: "Report junk", strategy: passthrough },
  // Outlook's own "Move to" picker has folder search, which covers Fastmail's `m` until the
  // palette can move messages itself (docs/decisions.md).
  move: { description: "Move to folder…", strategy: drive("moveButton") },
  goTo: { description: "Go to folder or app…", strategy: extension("palette-go") },
  snooze: { description: "Snooze", strategy: passthrough },
  markRead: { description: "Mark read", strategy: readState("read") },
  markUnread: { description: "Mark unread", strategy: readState("unread") },
  reply: { description: "Reply", strategy: passthrough },
  replyAll: { description: "Reply all", strategy: passthrough },
  forward: { description: "Forward", strategy: passthrough },
  undo: { description: "Undo", strategy: drive("undoButton") },
  refresh: { description: "Refresh", strategy: drive("syncButton") },
  goMail: { description: "Go to Mail", strategy: drive("mailApp") },
  goCalendar: { description: "Go to Calendar", strategy: drive("calendarApp") },
  goContacts: { description: "Go to Contacts", strategy: drive("peopleApp") },
  goSettings: { description: "Go to Settings", strategy: drive("settingsButton") },
} as const satisfies Record<string, Action>;

export type ActionId = keyof typeof ACTIONS;
