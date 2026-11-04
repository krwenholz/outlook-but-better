// OBB's own UI lives in one shadow root, so Outlook's CSS cannot reach it and ours cannot
// leak into Outlook. Everything here is plain DOM; no framework.

import css from "./overlay.css";

let shadow: ShadowRoot | null = null;

export function overlayRoot(): ShadowRoot {
  if (shadow) return shadow;
  const host = document.createElement("overlook-root");
  document.documentElement.append(host);
  shadow = host.attachShadow({ mode: "open" });
  // Keystrokes typed into our UI must not reach Outlook's document-level shortcut handlers,
  // which would see the host element (not an input) as the target.
  for (const type of ["keydown", "keypress", "keyup"]) {
    host.addEventListener(type, (e) => e.stopPropagation());
  }
  const style = document.createElement("style");
  style.textContent = css;
  shadow.append(style);
  return shadow;
}

/** True when the event came from inside OBB's own UI. */
export function isFromOverlay(event: Event): boolean {
  return shadow !== null && event.composedPath().includes(shadow.host);
}

/** True while a modal (palette or help) is open; it owns the keyboard until closed. */
export function isModalOpen(): boolean {
  return shadow?.querySelector(".backdrop") != null;
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className: string, text?: string) {
  const node = document.createElement(tag);
  node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function singleton(className: string): HTMLDivElement {
  const root = overlayRoot();
  const existing = root.querySelector<HTMLDivElement>(`.${className}`);
  if (existing) return existing;
  const node = el("div", className);
  root.append(node);
  return node;
}

let toastTimer: ReturnType<typeof setTimeout> | undefined;

export function toast(message: string, ms = 2500): void {
  const node = singleton("toast");
  node.textContent = message;
  node.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    node.hidden = true;
  }, ms);
}

/** The pending-prefix hint, e.g. "G … m Mail · c Calendar". Pass null to hide. */
export function showHint(text: string | null): void {
  const node = singleton("hint");
  node.hidden = text === null;
  node.textContent = text ?? "";
}

export function showHealthBadge(broken: readonly string[], onDismiss: () => void): void {
  const node = singleton("health");
  node.replaceChildren();
  if (broken.length === 0) {
    node.hidden = true;
    return;
  }
  node.hidden = false;
  node.title = `Not found: ${broken.join(", ")}. See the console for details.`;
  const close = el("button", "close", "×");
  close.setAttribute("aria-label", "Dismiss");
  close.addEventListener("click", () => {
    node.hidden = true;
    onDismiss();
  });
  node.append(el("span", "", `OBB: ${broken.length} selector(s) broken`), close);
}

export function showHelp(rows: readonly { keys: string; description: string }[]): void {
  const root = overlayRoot();
  const backdrop = el("div", "backdrop");
  const panel = el("div", "panel help");
  panel.append(el("h2", "", "Keyboard shortcuts"));
  const table = el("table", "");
  for (const { keys, description } of rows) {
    const tr = el("tr", "");
    const k = el("td", "keys");
    k.append(el("kbd", "", keys));
    tr.append(k, el("td", "", description));
    table.append(tr);
  }
  panel.append(table);
  backdrop.append(panel);
  root.append(backdrop);

  const close = () => {
    backdrop.remove();
    window.removeEventListener("keydown", onKey, true);
  };
  const onKey = (e: KeyboardEvent) => {
    e.preventDefault();
    e.stopPropagation();
    close();
  };
  backdrop.addEventListener("click", (e) => {
    if (e.target === backdrop) close();
  });
  window.addEventListener("keydown", onKey, true);
}

export { el };
