// Action table: what each action means and how it executes. Pure data.
//
// Strategies (PLAN.md, Keymap layer):
//   passthrough – Outlook's Gmail shortcut mode already does it; let the key through.
//   translate   – dispatch Outlook's own shortcut. Only works if Outlook honours synthetic
//                 (isTrusted: false) events; unknown until the Phase 0 spike.
//   drive       – click a control found through the selector registry.
//   extension   – Overlook's own UI (palette, help).
//   todo        – not decided yet; the key is swallowed and a hint says so.
//
// PROVISIONAL: strategies and Gmail-mode keys below are best guesses. Phase 1 replaces them
// with what was verified against the live page, and fills in PLAN.md's mechanism column.

import type { SelectorName } from "../selectors.ts";

export type ExtensionCommand = "help" | "palette-go" | "palette-move";

export type Strategy =
  | { kind: "passthrough" }
  | { kind: "translate"; key: string }
  | { kind: "drive"; target: SelectorName }
  | { kind: "extension"; command: ExtensionCommand }
  | { kind: "todo" };

export interface Action {
  description: string;
  strategy: Strategy;
}

const passthrough: Strategy = { kind: "passthrough" };
const todo: Strategy = { kind: "todo" };
const translate = (key: string): Strategy => ({ kind: "translate", key });
const drive = (target: SelectorName): Strategy => ({ kind: "drive", target });
const extension = (command: ExtensionCommand): Strategy => ({ kind: "extension", command });

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
  archive: { description: "Archive", strategy: translate("e") },
  archivePrevious: { description: "Archive, then open previous", strategy: todo },
  archiveNext: { description: "Archive, then open next", strategy: todo },
  delete: { description: "Delete", strategy: translate("#") },
  junk: { description: "Report junk", strategy: passthrough },
  move: { description: "Move to folder…", strategy: extension("palette-move") },
  goTo: { description: "Go to folder…", strategy: extension("palette-go") },
  snooze: { description: "Snooze", strategy: drive("snoozeButton") },
  pin: { description: "Toggle flag", strategy: translate("s") },
  markRead: { description: "Mark read", strategy: translate("I") },
  markUnread: { description: "Mark unread", strategy: translate("U") },
  reply: { description: "Reply", strategy: passthrough },
  replyAll: { description: "Reply all", strategy: passthrough },
  forward: { description: "Forward", strategy: passthrough },
  undo: { description: "Undo", strategy: passthrough },
  refresh: { description: "Refresh", strategy: todo },
  goMail: { description: "Go to Mail", strategy: todo },
  goCalendar: { description: "Go to Calendar", strategy: todo },
  goContacts: { description: "Go to Contacts", strategy: todo },
  goSettings: { description: "Go to Settings", strategy: todo },
} as const satisfies Record<string, Action>;

export type ActionId = keyof typeof ACTIONS;
