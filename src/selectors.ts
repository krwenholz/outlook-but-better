// The only file allowed to know what Outlook's DOM looks like.
//
// Anchor preference (see PLAN.md): structural IDs > ARIA roles/labels > semantic data-* > text.
// Never use generated class names; if one is unavoidable, mark it `// FRAGILE:`.
//
// STATUS: every candidate below is an unverified guess until the Phase 0 DOM survey
// (docs/dom-survey.md) confirms it. aria-label values assume an English UI.

export type SelectorName = keyof typeof SELECTORS;

export interface SelectorEntry {
  /** Candidate CSS selectors, most stable first. The first one that matches wins. */
  candidates: readonly string[];
  /** "always": should exist on every mail view, so a miss is a breakage. "sometimes": contextual. */
  expect: "always" | "sometimes";
  /** Name of another entry to resolve inside, instead of the whole document. */
  within?: string;
  /** Token added to the element's `data-overlook` attribute so CSS can target it. */
  tag?: string;
  /** Tag every match rather than only the first. */
  all?: boolean;
}

export const SELECTORS = {
  // --- Structure ---
  messageList: {
    candidates: ["#MailList", '[role="listbox"][aria-label*="Message list" i]'],
    expect: "always",
    tag: "message-list",
  },
  readingPane: {
    candidates: ["#ReadingPaneContainerId", '[role="main"][aria-label*="Reading pane" i]'],
    expect: "sometimes",
    tag: "reading-pane",
  },
  folderPane: {
    candidates: ['[role="tree"][aria-label*="folder" i]', '[role="navigation"] [role="tree"]'],
    expect: "always",
    tag: "folder-pane",
  },
  folderItem: {
    candidates: ['[role="treeitem"]'],
    expect: "always",
    within: "folderPane",
    tag: "folder-item",
    all: true,
  },
  searchBox: {
    candidates: ["#topSearchInput", 'input[aria-label="Search"]', '[role="search"] input'],
    expect: "always",
  },

  // --- Clutter (each hideable entry gets a toggle in settings.ts) ---
  appRail: {
    candidates: ["#LeftRail", '[role="navigation"][aria-label="App bar"]'],
    expect: "sometimes",
    tag: "app-rail",
  },
  ribbon: {
    candidates: ["#RibbonRoot", '[role="toolbar"][aria-label="Ribbon"]'],
    expect: "sometimes",
    tag: "ribbon",
  },
  copilot: {
    candidates: ['button[aria-label*="Copilot" i]', '[role="region"][aria-label*="Copilot" i]'],
    expect: "sometimes",
    tag: "copilot",
    all: true,
  },
  promoBanner: {
    candidates: ['[role="region"][aria-label*="banner" i]'],
    expect: "sometimes",
    tag: "promo-banner",
    all: true,
  },
  externalTag: {
    candidates: ['[title="External sender"]', '[aria-label="External"]'],
    expect: "sometimes",
    tag: "external-tag",
    all: true,
  },

  // --- Controls the keymap drives ---
  composeButton: {
    candidates: ['button[aria-label="New mail"]', 'button[aria-label="New email"]'],
    expect: "always",
  },
  archiveButton: { candidates: ['button[aria-label="Archive"]'], expect: "sometimes" },
  deleteButton: { candidates: ['button[aria-label="Delete"]'], expect: "sometimes" },
  junkButton: {
    candidates: ['button[aria-label="Report"]', 'button[aria-label="Junk"]'],
    expect: "sometimes",
  },
  moveButton: { candidates: ['button[aria-label="Move to"]'], expect: "sometimes" },
  snoozeButton: { candidates: ['button[aria-label="Snooze"]'], expect: "sometimes" },
  flagButton: { candidates: ['button[aria-label="Flag / Unflag"]'], expect: "sometimes" },
  readUnreadButton: { candidates: ['button[aria-label="Read / Unread"]'], expect: "sometimes" },
  undoButton: { candidates: ['button[aria-label="Undo"]'], expect: "sometimes" },
} as const satisfies Record<string, SelectorEntry>;

const entry = (name: SelectorName): SelectorEntry => SELECTORS[name];

function scope(name: SelectorName, root: ParentNode): ParentNode | null {
  const parent = entry(name).within;
  return parent ? resolve(parent as SelectorName, root) : root;
}

/** The first element matching `name`, trying candidates in priority order. */
export function resolve(name: SelectorName, root: ParentNode = document): Element | null {
  const within = scope(name, root);
  if (!within) return null;
  for (const candidate of entry(name).candidates) {
    const el = within.querySelector(candidate);
    if (el) return el;
  }
  return null;
}

/** Every element matched by the first candidate that matches anything. */
export function resolveAll(name: SelectorName, root: ParentNode = document): Element[] {
  const within = scope(name, root);
  if (!within) return [];
  for (const candidate of entry(name).candidates) {
    const els = within.querySelectorAll(candidate);
    if (els.length > 0) return [...els];
  }
  return [];
}

/** Add `data-overlook` tokens to every tagged entry so stylesheets never touch Outlook's DOM shape. */
export function tagElements(root: ParentNode = document): void {
  for (const name of Object.keys(SELECTORS) as SelectorName[]) {
    const { tag, all } = entry(name);
    if (!tag) continue;
    const els = all ? resolveAll(name, root) : [resolve(name, root)].filter((e) => e !== null);
    for (const el of els) {
      const tokens = new Set((el.getAttribute("data-overlook") ?? "").split(" ").filter(Boolean));
      if (tokens.has(tag)) continue;
      tokens.add(tag);
      el.setAttribute("data-overlook", [...tokens].join(" "));
    }
  }
}

export interface HealthResult {
  name: SelectorName;
  expect: SelectorEntry["expect"];
  /** Index of the candidate that matched, or -1. A non-zero index means a fallback is in use. */
  candidate: number;
  count: number;
}

export function healthCheck(root: ParentNode = document): HealthResult[] {
  return (Object.keys(SELECTORS) as SelectorName[]).map((name) => {
    const { candidates, expect } = entry(name);
    const within = scope(name, root);
    let candidate = -1;
    let count = 0;
    if (within) {
      candidate = candidates.findIndex((c) => within.querySelector(c) !== null);
      if (candidate >= 0) count = within.querySelectorAll(candidates[candidate] ?? "").length;
    }
    return { name, expect, candidate, count };
  });
}

export interface FolderNode {
  name: string;
  /** Ancestor names joined with "/", e.g. "Inbox/Receipts". */
  path: string;
  element: Element;
}

/**
 * Folders currently rendered in the folder pane, with paths built from aria-level.
 * UNVERIFIED: collapsed subfolders may be absent from the DOM, and the visible text may include
 * unread counts. The Phase 1 palette spike decides whether this approach survives.
 */
export function readFolderTree(root: ParentNode = document): FolderNode[] {
  const stack: string[] = [];
  return resolveAll("folderItem", root).map((element) => {
    const level = Number(element.getAttribute("aria-level") ?? "1");
    const name = (element.getAttribute("title") ?? element.textContent ?? "").trim();
    stack.length = Math.max(level - 1, 0);
    stack.push(name);
    return { name, path: stack.filter(Boolean).join("/"), element };
  });
}

/** Entries expected on every mail view that did not resolve. Non-empty means something broke. */
export function brokenEntries(results: HealthResult[]): SelectorName[] {
  return results.filter((r) => r.expect === "always" && r.candidate < 0).map((r) => r.name);
}
