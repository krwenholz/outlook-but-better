// The Fastmail-style binding set (PLAN.md, Shortcut map). Keys are KeyboardEvent.key values.

import type { ActionId } from "./actions.ts";
import type { Binding, Context } from "./engine.ts";

const MAIL: readonly Context[] = ["list", "reading"];
const LIST: readonly Context[] = ["list"];
const READING: readonly Context[] = ["reading"];

const bind = (keys: string, action: ActionId, contexts = MAIL): Binding<ActionId> => ({
  keys: keys.split(" "),
  action,
  contexts,
});

export const BINDINGS: readonly Binding<ActionId>[] = [
  bind("c", "compose"),
  bind("/", "search"),
  bind("?", "help"),
  bind("j", "next"),
  bind("k", "previous"),
  bind("o", "open", LIST),
  bind("Enter", "open", LIST),
  bind("x", "select", LIST),
  bind("X", "selectRange", LIST),
  bind("* a", "selectAll", LIST),
  bind("* n", "selectNone", LIST),
  bind("y", "archive"),
  bind("]", "archivePrevious", READING),
  bind("[", "archiveNext", READING),
  bind("d", "delete"),
  bind("!", "junk"),
  bind("m", "move"),
  bind("g", "goTo"),
  bind("b", "snooze"),
  bind("s", "pin"),
  bind(". r", "markRead"),
  bind(". u", "markUnread"),
  bind("r", "reply", READING),
  bind("a", "replyAll", READING),
  bind("f", "forward", READING),
  bind("z", "undo"),
  bind("u", "refresh"),
  bind("G m", "goMail"),
  bind("G c", "goCalendar"),
  bind("G a", "goContacts"),
  bind("G s", "goSettings"),
];
