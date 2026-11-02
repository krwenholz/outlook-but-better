// The only file allowed to know what Outlook's DOM looks like.
//
// Anchor preference (see PLAN.md): structural IDs > ARIA roles/labels > semantic data-* > text.
// Never use generated class names; if one is unavoidable, mark it `// FRAGILE:`.
//
// Verified against a saved page from Oct 2026 (test/fixtures/inbox.html); see docs/dom-survey.md.
// Entries marked UNVERIFIED were not present in that capture. aria-labels assume an English UI.

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
  /** Match every candidate (a list of distinct things), not just the first that matches. */
  union?: boolean;
}

export const SELECTORS = {
  // --- Structure ---
  titleBar: { candidates: ["#OwaTitleBar"], expect: "always", tag: "title-bar" },
  searchBox: { candidates: ["#topSearchInput", '[role="search"] input'], expect: "always" },
  appRail: {
    candidates: ["#LeftRail", '[role="navigation"][aria-label="left-rail-appbar"]'],
    expect: "always",
    tag: "app-rail",
  },
  ribbon: {
    candidates: ['[data-app-section="Ribbon"]', '[role="region"][aria-label="Ribbon"]'],
    expect: "sometimes",
    tag: "ribbon",
  },
  ribbonTabs: {
    candidates: ['[role="tablist"][aria-label="Ribbon tabs"]'],
    expect: "sometimes",
    within: "ribbon",
    tag: "ribbon-tabs",
  },
  folderPane: {
    candidates: [
      '[data-app-section="NavigationPane"]',
      '[role="region"][aria-label="Navigation pane"]',
    ],
    expect: "always",
    tag: "folder-pane",
  },
  folderItem: {
    candidates: ['[role="treeitem"][data-folder-name]', '[role="treeitem"]'],
    expect: "always",
    within: "folderPane",
    tag: "folder-item",
    all: true,
  },
  listPane: {
    candidates: ['[data-app-section="MessageList"]'],
    expect: "always",
    tag: "list-pane",
  },
  messageList: {
    candidates: ["#MailList", '[role="listbox"][aria-label="Message list"]'],
    expect: "always",
    tag: "message-list",
  },
  messageRow: {
    candidates: ['[role="option"][data-convid]', '[role="option"]'],
    expect: "sometimes",
    within: "messageList",
    tag: "message-row",
    all: true,
  },
  groupHeader: {
    candidates: ['[role="button"][id^="groupHeader"]'],
    expect: "sometimes",
    within: "messageList",
    tag: "group-header",
    all: true,
  },
  focusedTabs: {
    candidates: ['[role="tablist"]'],
    expect: "sometimes",
    within: "listPane",
    tag: "focused-tabs",
  },
  readingPane: {
    candidates: [
      '[data-app-section="MailReadCompose"]',
      "#ReadingPaneContainerId",
      '[role="main"][aria-label="Reading Pane"]',
    ],
    expect: "sometimes",
    tag: "reading-pane",
  },

  messageBody: {
    candidates: ['[role="document"][aria-label="Message body"]'],
    expect: "sometimes",
    within: "readingPane",
    tag: "message-body",
    all: true,
  },

  infoBar: {
    // Notices above a message ("You replied on…"); their text sits in an element whose id starts
    // with "infoBar". Outlook paints them with inline colours.
    candidates: ['[role="group"]:has(> * > [id^="infoBar"])'],
    expect: "sometimes",
    within: "readingPane",
    tag: "info-bar",
    all: true,
  },

  // --- Clutter (each hideable entry gets a toggle in settings.ts) ---
  upsellBanner: {
    candidates: ['[data-app-section="UpsellBannerSection"]'],
    expect: "sometimes",
    tag: "upsell",
    all: true,
  },
  headerExtras: {
    candidates: ["#owaTeamsBtn_container", "#owaTimePanelBtn_container"],
    expect: "sometimes",
    tag: "header-extra",
    union: true,
  },
  appRailExtras: {
    candidates: ["To Do", "Newsletters", "Org Explorer", "OneDrive"].map(
      (app) => `[role="none"][aria-label="${app}"]`,
    ),
    expect: "sometimes",
    within: "appRail",
    tag: "rail-extra",
    union: true,
  },
  ribbonExtras: {
    candidates: [
      "Share to Teams",
      "Sweep",
      "Assign policy",
      "Viva Insights",
      "More apps",
      "Quick steps",
      "More quick steps",
    ].map((label) => `[aria-label="${label}"]`),
    expect: "sometimes",
    within: "ribbon",
    tag: "ribbon-extra",
    union: true,
  },
  rowAvatar: {
    candidates: ['[role="img"][id^="avatar-"]'],
    expect: "sometimes",
    within: "messageList",
    tag: "row-avatar",
    all: true,
  },
  rowHoverActions: {
    candidates: [
      "Customize Actions",
      "Flag this message",
      "Keep this message at the top of your folder",
    ].map((label) => `button[aria-label="${label}"]`),
    expect: "sometimes",
    within: "messageList",
    tag: "row-action",
    union: true,
  },
  appBanner: {
    // Fluent's static "fui-MessageBar" class is part of its public styling API, not generated;
    // nothing else marks these. Scoped to bars across the top of the app, not inside messages.
    candidates: ['#mainApp > div > [role="group"].fui-MessageBar'],
    expect: "sometimes",
    tag: "app-banner",
    all: true,
  },
  copilot: {
    // UNVERIFIED: no Copilot surfaces in the Oct 2026 capture.
    candidates: ['button[aria-label*="Copilot" i]', '[role="region"][aria-label*="Copilot" i]'],
    expect: "sometimes",
    tag: "copilot",
    union: true,
  },
  externalTag: {
    // UNVERIFIED: no External tags in the Oct 2026 capture.
    candidates: ['[title="External sender"]', '[aria-label="External"]'],
    expect: "sometimes",
    tag: "external-tag",
    union: true,
  },

  // --- Controls the keymap drives ---
  composeButton: { candidates: ['button[aria-label="New mail"]'], expect: "always" },
  archiveButton: { candidates: ['button[aria-label="Archive"]'], expect: "sometimes" },
  deleteButton: { candidates: ['button[aria-label="Delete"]'], expect: "sometimes" },
  // UNVERIFIED: this tenant shows a "Phish Alert" add-in instead of Outlook's Report menu.
  junkButton: {
    candidates: ['button[aria-label="Report"]', 'button[aria-label="Junk"]'],
    expect: "sometimes",
  },
  moveButton: { candidates: ['button[aria-label="Move to"]'], expect: "sometimes" },
  snoozeButton: { candidates: ['button[aria-label="Snooze"]'], expect: "sometimes" },
  flagButton: { candidates: ['button[aria-label="Flag / Unflag"]'], expect: "sometimes" },
  readUnreadButton: { candidates: ['button[aria-label="Read / Unread"]'], expect: "sometimes" },
  undoButton: { candidates: ['button[aria-label="Undo"]'], expect: "sometimes" },
  syncButton: { candidates: ["#owaSyncButton_container"], expect: "sometimes" },
  settingsButton: { candidates: ["#owaSettingsBtn_container"], expect: "always" },
  mailApp: { candidates: ['button[aria-label="Mail"]'], within: "appRail", expect: "always" },
  calendarApp: {
    candidates: ['button[aria-label="Calendar"]'],
    within: "appRail",
    expect: "always",
  },
  peopleApp: { candidates: ['button[aria-label="People"]'], within: "appRail", expect: "always" },
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

/** Every match: for union entries, of all candidates; otherwise of the first candidate that matches. */
export function resolveAll(name: SelectorName, root: ParentNode = document): Element[] {
  const within = scope(name, root);
  if (!within) return [];
  const { candidates, union } = entry(name);
  if (union) return [...within.querySelectorAll(candidates.join(", "))];
  for (const candidate of candidates) {
    const els = within.querySelectorAll(candidate);
    if (els.length > 0) return [...els];
  }
  return [];
}

/** Add `data-overlook` tokens to every tagged entry so stylesheets never touch Outlook's DOM shape. */
export function tagElements(root: ParentNode = document): void {
  for (const name of Object.keys(SELECTORS) as SelectorName[]) {
    const { tag, all, union } = entry(name);
    if (!tag) continue;
    const els =
      all || union ? resolveAll(name, root) : [resolve(name, root)].filter((e) => e !== null);
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
    const { candidates, expect, union } = entry(name);
    const within = scope(name, root);
    let candidate = -1;
    let count = 0;
    if (within) {
      candidate = candidates.findIndex((c) => within.querySelector(c) !== null);
      if (union && candidate >= 0) count = resolveAll(name, root).length;
      else if (candidate >= 0) count = within.querySelectorAll(candidates[candidate] ?? "").length;
    }
    // A union entry lists distinct things, so matching a later one is not a fallback.
    if (union && candidate > 0) candidate = 0;
    return { name, expect, candidate, count };
  });
}

export interface FolderNode {
  name: string;
  /** Ancestor names joined with "/", e.g. "Inbox/Receipts". */
  path: string;
  element: Element;
}

/** "Inbox - 9 items (1 unread)" -> "Inbox". The title attribute carries counts. */
function folderName(element: Element): string {
  const title = (element.getAttribute("title") ?? element.textContent ?? "").trim();
  return title.replace(/\s+-\s+\d+\s+items?(\s+\(\d+\s+unread\))?$/, "");
}

/**
 * Folders currently rendered in the folder pane, with paths built from aria-level. Folders
 * under Favorites are skipped (they duplicate the mailbox tree), as are the top-level mailbox
 * roots, whose names are account addresses.
 * UNVERIFIED: collapsed subfolders may be absent from the DOM.
 */
export function readFolderTree(root: ParentNode = document): FolderNode[] {
  const stack: string[] = [];
  const folders: FolderNode[] = [];
  let inFavorites = false;
  for (const element of resolveAll("folderItem", root)) {
    const level = Number(element.getAttribute("aria-level") ?? "1");
    const name = folderName(element);
    stack.length = Math.max(level - 1, 0);
    stack.push(name);
    if (level === 1) {
      inFavorites = element.getAttribute("data-folder-name") === "favorites";
      continue;
    }
    if (inFavorites) continue;
    folders.push({ name, path: stack.slice(1).filter(Boolean).join("/"), element });
  }
  return folders;
}

/** Outlook's current light/dark mode, from the palette variables it writes on <html>. */
export function outlookColorScheme(html: Element): "light" | "dark" {
  const style = html.getAttribute("style") ?? "";
  return /--native-scrollbar-color-scheme:\s*dark/.test(style) ? "dark" : "light";
}

/** Entries expected on every mail view that did not resolve. Non-empty means something broke. */
export function brokenEntries(results: HealthResult[]): SelectorName[] {
  return results.filter((r) => r.expect === "always" && r.candidate < 0).map((r) => r.name);
}
