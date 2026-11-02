// Action table: what each action means and how it executes. Pure data.
//
// Strategies (PLAN.md, Keymap layer):
//   passthrough – Outlook's Gmail shortcut mode already does it; let the key through.
//   translate   – dispatch Outlook's own shortcut. Only works if Outlook honours synthetic
//                 (isTrusted: false) events; still unverified on the live site.
//   drive       – click a control found through the selector registry. React honours click(),
//                 so this is the most reliable option; `key` is a translate fallback for when
//                 the control is absent (e.g. collapsed into the ribbon's overflow menu).
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
  | { kind: "translate"; key: string }
  | { kind: "drive"; target: SelectorName; key?: string }
  | { kind: "extension"; command: ExtensionCommand }
  | { kind: "todo" };

export interface Action {
  description: string;
  strategy: Strategy;
}

const passthrough: Strategy = { kind: "passthrough" };
const todo: Strategy = { kind: "todo" };
const translate = (key: string): Strategy => ({ kind: "translate", key });
const drive = (target: SelectorName, key?: string): Strategy => ({ kind: "drive", target, key });
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
  archive: { description: "Archive", strategy: drive("archiveButton", "e") },
  archivePrevious: { description: "Archive, then open previous", strategy: todo },
  archiveNext: { description: "Archive, then open next", strategy: todo },
  delete: { description: "Delete", strategy: drive("deleteButton", "#") },
  junk: { description: "Report junk", strategy: passthrough },
  // Outlook's own "Move to" picker has folder search, which covers Fastmail's `m` until the
  // palette can move messages itself (docs/decisions.md).
  move: { description: "Move to folder…", strategy: drive("moveButton", "v") },
  goTo: { description: "Go to folder…", strategy: extension("palette-go") },
  snooze: { description: "Snooze", strategy: passthrough },
  markRead: { description: "Mark read", strategy: translate("I") },
  markUnread: { description: "Mark unread", strategy: translate("U") },
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
