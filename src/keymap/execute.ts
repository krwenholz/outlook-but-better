// Turns an action into DOM effects. The only keymap file that touches the page.

import { openFolderPalette } from "../palette/palette.ts";
import {
  checkedRows,
  cursorRow,
  isUnreadRow,
  resolve,
  resolveAll,
  rowCheckbox,
  selectedRows,
} from "../selectors.ts";
import { showHelp, toast } from "../ui/overlay.ts";
import { ACTIONS, type ActionId, type ExtensionCommand } from "./actions.ts";
import { BINDINGS } from "./bindings.ts";
import type { Context } from "./engine.ts";

/** Which bindings apply, judged from where focus is. */
export function currentContext(): Context {
  const reading = resolve("readingPane");
  const focused = document.activeElement;
  return reading && focused && reading.contains(focused) ? "reading" : "list";
}

export interface ExecuteOptions {
  /** Take over x and j/k so ticks survive moving (Fastmail-style). Off: Outlook's own keys. */
  multiSelect: boolean;
}

/** Moves the cursor one row without selecting it, so ticked rows stay ticked. */
function moveCursor(step: 1 | -1): boolean {
  // Nothing ticked: Outlook's own j/k (which select as they go) are what you want.
  if (checkedRows().length === 0) return false;
  const rows = resolveAll("messageRow");
  const at = cursorRow();
  const next = rows[(at ? rows.indexOf(at) : -1) + step];
  if (next instanceof HTMLElement) {
    next.focus();
    next.scrollIntoView?.({ block: "nearest" });
  }
  return true; // at either end the key is swallowed rather than collapsing the selection
}

/** Ticks or unticks the cursor row, as clicking its tick box does. */
function toggleSelect(): boolean {
  const row = cursorRow();
  const box = row && rowCheckbox(row);
  if (!box) return false;
  box.click();
  return true;
}

/** Returns false when the key should be left for Outlook. */
function runExtension(command: ExtensionCommand, options: ExecuteOptions): boolean {
  switch (command) {
    case "help":
      showHelp(
        BINDINGS.map((b) => ({
          keys: b.keys.join(" "),
          description: ACTIONS[b.action].description,
        })),
      );
      return true;
    case "palette-go":
      openFolderPalette();
      return true;
    case "toggle-select":
      return options.multiSelect && toggleSelect();
    case "cursor-next":
      return options.multiSelect && moveCursor(1);
    case "cursor-previous":
      return options.multiSelect && moveCursor(-1);
  }
}

/** Run an action. Returns false when the key should be left for Outlook (passthrough). */
export function execute(id: ActionId, options: ExecuteOptions = { multiSelect: true }): boolean {
  const { strategy, description } = ACTIONS[id];
  switch (strategy.kind) {
    case "passthrough":
      return false;
    case "drive": {
      const target = resolve(strategy.target);
      if (target instanceof HTMLElement) target.click();
      else toast(`${description}: couldn't find the button`);
      return true;
    }
    case "toggle": {
      const rows = selectedRows();
      if (rows.length === 0) {
        toast(`${description}: select a message first`);
        return true;
      }
      const wantUnread = strategy.want === "unread";
      if (rows.every((row) => isUnreadRow(row) === wantUnread)) return true;
      const target = resolve(strategy.target);
      if (target instanceof HTMLElement) target.click();
      else toast(`${description}: couldn't find the button`);
      return true;
    }
    case "extension":
      return runExtension(strategy.command, options);
    case "todo":
      toast(`${description} isn't implemented yet`);
      return true;
  }
}
