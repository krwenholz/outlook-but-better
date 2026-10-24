// Turns an action into DOM effects. The only keymap file that touches the page.

import { openFolderPalette } from "../palette/palette.ts";
import { resolve } from "../selectors.ts";
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

/** Dispatch Outlook's own shortcut. Our listener ignores these because isTrusted is false. */
function dispatchKey(key: string): void {
  const target = document.activeElement ?? document.body;
  const shiftKey = key.length === 1 && key !== key.toLowerCase();
  const init: KeyboardEventInit = {
    key,
    shiftKey,
    bubbles: true,
    cancelable: true,
    composed: true,
  };
  target.dispatchEvent(new KeyboardEvent("keydown", init));
  target.dispatchEvent(new KeyboardEvent("keyup", init));
}

function runExtension(command: ExtensionCommand): void {
  switch (command) {
    case "help":
      showHelp(
        BINDINGS.map((b) => ({
          keys: b.keys.join(" "),
          description: ACTIONS[b.action].description,
        })),
      );
      return;
    case "palette-go":
      openFolderPalette();
      return;
  }
}

/** Run an action. Returns false when the key should be left for Outlook (passthrough). */
export function execute(id: ActionId): boolean {
  const { strategy, description } = ACTIONS[id];
  switch (strategy.kind) {
    case "passthrough":
      return false;
    case "translate":
      dispatchKey(strategy.key);
      return true;
    case "drive": {
      const target = resolve(strategy.target);
      if (target instanceof HTMLElement) target.click();
      else if (strategy.key) dispatchKey(strategy.key);
      else toast(`${description}: couldn't find the button`);
      return true;
    }
    case "extension":
      runExtension(strategy.command);
      return true;
    case "todo":
      toast(`${description} isn't implemented yet`);
      return true;
  }
}
