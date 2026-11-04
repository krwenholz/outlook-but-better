"use strict";
(() => {
  // src/keymap/actions.ts
  var passthrough = { kind: "passthrough" };
  var todo = { kind: "todo" };
  var drive = (target) => ({ kind: "drive", target });
  var extension = (command) => ({ kind: "extension", command });
  var readState = (want) => ({
    kind: "toggle",
    target: "readUnreadButton",
    want
  });
  var ACTIONS = {
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
    move: { description: "Move to folder\u2026", strategy: drive("moveButton") },
    goTo: { description: "Go to folder or app\u2026", strategy: extension("palette-go") },
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
    goSettings: { description: "Go to Settings", strategy: drive("settingsButton") }
  };

  // src/keymap/bindings.ts
  var MAIL = ["list", "reading"];
  var LIST = ["list"];
  var READING = ["reading"];
  var bind = (keys, action, contexts = MAIL) => ({
    keys: keys.split(" "),
    action,
    contexts
  });
  var BINDINGS = [
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
    bind("G s", "goSettings")
  ];

  // src/keymap/engine.ts
  var DEFAULT_TIMEOUT_MS = 1500;
  var startsWith = (keys, prefix) => prefix.length <= keys.length && prefix.every((k, i) => keys[i] === k);
  var SequenceEngine = class {
    #bindings;
    #timeoutMs;
    #prefix = [];
    #lastAt = 0;
    constructor(bindings, timeoutMs = DEFAULT_TIMEOUT_MS) {
      this.#bindings = bindings;
      this.#timeoutMs = timeoutMs;
    }
    get pending() {
      return this.#prefix;
    }
    reset() {
      this.#prefix = [];
    }
    feed(key, context, now) {
      if (this.#prefix.length > 0 && now - this.#lastAt > this.#timeoutMs) this.reset();
      const wasPending = this.#prefix.length > 0;
      const seq = [...this.#prefix, key];
      const active = this.#bindings.filter((b) => b.contexts.includes(context));
      const longer = active.filter((b) => b.keys.length > seq.length && startsWith(b.keys, seq));
      if (longer.length > 0) {
        this.#prefix = seq;
        this.#lastAt = now;
        return { kind: "pending", prefix: seq, next: longer };
      }
      this.reset();
      const exact = active.find((b) => b.keys.length === seq.length && startsWith(b.keys, seq));
      if (exact) return { kind: "match", action: exact.action };
      return wasPending ? { kind: "cancelled" } : { kind: "none" };
    }
  };
  var MODIFIER_KEYS = /* @__PURE__ */ new Set(["Shift", "Control", "Alt", "Meta", "CapsLock", "Dead"]);
  function keyFromEvent(event) {
    if (event.ctrlKey || event.metaKey || event.altKey) return null;
    if (MODIFIER_KEYS.has(event.key)) return null;
    return event.key;
  }

  // src/selectors.ts
  var SELECTORS = {
    // --- Structure ---
    titleBar: { candidates: ["#OwaTitleBar"], expect: "always", tag: "title-bar" },
    searchBox: { candidates: ["#topSearchInput", '[role="search"] input'], expect: "always" },
    appRail: {
      candidates: ["#LeftRail", '[role="navigation"][aria-label="left-rail-appbar"]'],
      expect: "always",
      tag: "app-rail"
    },
    ribbon: {
      candidates: ['[data-app-section="Ribbon"]', '[role="region"][aria-label="Ribbon"]'],
      expect: "sometimes",
      tag: "ribbon"
    },
    ribbonTabs: {
      candidates: ['[role="tablist"][aria-label="Ribbon tabs"]'],
      expect: "sometimes",
      within: "ribbon",
      tag: "ribbon-tabs"
    },
    folderPane: {
      candidates: [
        '[data-app-section="NavigationPane"]',
        '[role="region"][aria-label="Navigation pane"]'
      ],
      expect: "always",
      tag: "folder-pane"
    },
    folderItem: {
      candidates: ['[role="treeitem"][data-folder-name]', '[role="treeitem"]'],
      expect: "always",
      within: "folderPane",
      tag: "folder-item",
      all: true
    },
    listPane: {
      candidates: ['[data-app-section="MessageList"]'],
      expect: "always",
      tag: "list-pane"
    },
    messageList: {
      candidates: ["#MailList", '[role="listbox"][aria-label="Message list"]'],
      expect: "always",
      tag: "message-list"
    },
    messageRow: {
      candidates: ['[role="option"][data-convid]', '[role="option"]'],
      expect: "sometimes",
      within: "messageList",
      tag: "message-row",
      all: true
    },
    groupHeader: {
      candidates: ['[role="button"][id^="groupHeader"]'],
      expect: "sometimes",
      within: "messageList",
      tag: "group-header",
      all: true
    },
    focusedTabs: {
      candidates: ['[role="tablist"]'],
      expect: "sometimes",
      within: "listPane",
      tag: "focused-tabs"
    },
    readingPane: {
      candidates: [
        '[data-app-section="MailReadCompose"]',
        "#ReadingPaneContainerId",
        '[role="main"][aria-label="Reading Pane"]'
      ],
      expect: "sometimes",
      tag: "reading-pane"
    },
    messageBody: {
      candidates: ['[role="document"][aria-label="Message body"]'],
      expect: "sometimes",
      within: "readingPane",
      tag: "message-body",
      all: true
    },
    infoBar: {
      // Notices above a message ("You replied on…"); their text sits in an element whose id starts
      // with "infoBar". Outlook paints them with inline colours.
      candidates: ['[role="group"]:has(> * > [id^="infoBar"])'],
      expect: "sometimes",
      within: "readingPane",
      tag: "info-bar",
      all: true
    },
    // --- Clutter (each hideable entry gets a toggle in settings.ts) ---
    upsellBanner: {
      candidates: ['[data-app-section="UpsellBannerSection"]'],
      expect: "sometimes",
      tag: "upsell",
      all: true
    },
    headerExtras: {
      candidates: ["#owaTeamsBtn_container", "#owaTimePanelBtn_container"],
      expect: "sometimes",
      tag: "header-extra",
      union: true
    },
    appRailExtras: {
      candidates: ["To Do", "Newsletters", "Org Explorer", "OneDrive"].map(
        (app) => `[role="none"][aria-label="${app}"]`
      ),
      expect: "sometimes",
      within: "appRail",
      tag: "rail-extra",
      union: true
    },
    ribbonExtras: {
      candidates: [
        "Share to Teams",
        "Sweep",
        "Assign policy",
        "Viva Insights",
        "More apps",
        "Quick steps",
        "More quick steps"
      ].map((label) => `[aria-label="${label}"]`),
      expect: "sometimes",
      within: "ribbon",
      tag: "ribbon-extra",
      union: true
    },
    rowAvatar: {
      candidates: ['[role="img"][id^="avatar-"]'],
      expect: "sometimes",
      within: "messageList",
      tag: "row-avatar",
      all: true
    },
    rowHoverActions: {
      candidates: [
        "Customize Actions",
        "Flag this message",
        "Keep this message at the top of your folder"
      ].map((label) => `button[aria-label="${label}"]`),
      expect: "sometimes",
      within: "messageList",
      tag: "row-action",
      union: true
    },
    appBanner: {
      // Fluent's static "fui-MessageBar" class is part of its public styling API, not generated;
      // nothing else marks these. Scoped to bars across the top of the app, not inside messages.
      candidates: ['#mainApp > div > [role="group"].fui-MessageBar'],
      expect: "sometimes",
      tag: "app-banner",
      all: true
    },
    copilot: {
      // UNVERIFIED: no Copilot surfaces in the Oct 2026 capture.
      candidates: ['button[aria-label*="Copilot" i]', '[role="region"][aria-label*="Copilot" i]'],
      expect: "sometimes",
      tag: "copilot",
      union: true
    },
    externalTag: {
      // UNVERIFIED: no External tags in the Oct 2026 capture.
      candidates: ['[title="External sender"]', '[aria-label="External"]'],
      expect: "sometimes",
      tag: "external-tag",
      union: true
    },
    // --- Controls the keymap drives ---
    composeButton: { candidates: ['button[aria-label="New mail"]'], expect: "always" },
    archiveButton: { candidates: ['button[aria-label="Archive"]'], expect: "sometimes" },
    deleteButton: { candidates: ['button[aria-label="Delete"]'], expect: "sometimes" },
    // UNVERIFIED: this tenant shows a "Phish Alert" add-in instead of Outlook's Report menu.
    junkButton: {
      candidates: ['button[aria-label="Report"]', 'button[aria-label="Junk"]'],
      expect: "sometimes"
    },
    moveButton: { candidates: ['button[aria-label="Move to"]'], expect: "sometimes" },
    snoozeButton: { candidates: ['button[aria-label="Snooze"]'], expect: "sometimes" },
    flagButton: { candidates: ['button[aria-label="Flag / Unflag"]'], expect: "sometimes" },
    readUnreadButton: { candidates: ['button[aria-label="Read / Unread"]'], expect: "sometimes" },
    undoButton: { candidates: ['button[aria-label="Undo"]'], expect: "sometimes" },
    syncButton: { candidates: ["#owaSyncButton_container"], expect: "sometimes" },
    settingsButton: { candidates: ["#owaSettingsBtn_container"], expect: "always" },
    // App rail buttons. Fallbacks are the Microsoft 365 app ids Outlook puts in data-tid, the same
    // across the Oct and Nov 2026 captures. The rail renders after the message list.
    mailApp: {
      candidates: ['button[aria-label="Mail"]', '[data-tid="ddea774c-382b-47d7-aab5-adc2139a802b"]'],
      within: "appRail",
      expect: "always"
    },
    calendarApp: {
      candidates: [
        'button[aria-label="Calendar"]',
        '[data-tid="8cbeb86f-83e1-43b5-aaba-cd3514322f0b"]'
      ],
      within: "appRail",
      expect: "always"
    },
    peopleApp: {
      candidates: [
        'button[aria-label="People"]',
        '[data-tid="355fbd79-3ba2-4554-8f2d-0300fde91f30"]'
      ],
      within: "appRail",
      expect: "always"
    }
  };
  var entry = (name) => SELECTORS[name];
  function scope(name, root) {
    const parent = entry(name).within;
    return parent ? resolve(parent, root) : root;
  }
  function resolve(name, root = document) {
    const within = scope(name, root);
    if (!within) return null;
    for (const candidate of entry(name).candidates) {
      const el2 = within.querySelector(candidate);
      if (el2) return el2;
    }
    return null;
  }
  function resolveAll(name, root = document) {
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
  function tagElements(root = document) {
    for (const name of Object.keys(SELECTORS)) {
      const { tag, all, union } = entry(name);
      if (!tag) continue;
      const els = all || union ? resolveAll(name, root) : [resolve(name, root)].filter((e) => e !== null);
      for (const el2 of els) {
        const tokens = new Set((el2.getAttribute("data-overlook") ?? "").split(" ").filter(Boolean));
        if (tokens.has(tag)) continue;
        tokens.add(tag);
        el2.setAttribute("data-overlook", [...tokens].join(" "));
      }
    }
  }
  function healthCheck(root = document) {
    return Object.keys(SELECTORS).map((name) => {
      const { candidates, expect, union } = entry(name);
      const within = scope(name, root);
      let candidate = -1;
      let count = 0;
      if (within) {
        candidate = candidates.findIndex((c) => within.querySelector(c) !== null);
        if (union && candidate >= 0) count = resolveAll(name, root).length;
        else if (candidate >= 0) count = within.querySelectorAll(candidates[candidate] ?? "").length;
      }
      if (union && candidate > 0) candidate = 0;
      return { name, expect, candidate, count };
    });
  }
  function folderName(element) {
    const title = (element.getAttribute("title") ?? element.textContent ?? "").trim();
    return title.replace(/\s+-\s+\d+\s+items?(\s+\(\d+\s+unread\))?$/, "");
  }
  function readFolderTree(root = document) {
    const stack = [];
    const folders = [];
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
  function selectedRows(root = document) {
    return resolveAll("messageRow", root).filter(
      (row) => row.getAttribute("aria-selected") === "true"
    );
  }
  function isUnreadRow(row) {
    return /^Unread\b/.test(row.getAttribute("aria-label") ?? "");
  }
  function outlookColorScheme(html) {
    const style = html.getAttribute("style") ?? "";
    return /--native-scrollbar-color-scheme:\s*dark/.test(style) ? "dark" : "light";
  }
  function brokenEntries(results) {
    return results.filter((r) => r.expect === "always" && r.candidate < 0).map((r) => r.name);
  }

  // src/ui/overlay.css
  var overlay_default = ':host {\n  all: initial;\n  --ol-bg: #fff;\n  --ol-fg: #1f1f1f;\n  --ol-muted: #6b6b6b;\n  --ol-accent: #2f6fde;\n  --ol-border: rgb(0 0 0 / 12%);\n  --ol-shadow: 0 8px 32px rgb(0 0 0 / 18%);\n  font:\n    13px / 1.4 system-ui,\n    -apple-system,\n    "Segoe UI",\n    sans-serif;\n}\n\n@media (prefers-color-scheme: dark) {\n  :host {\n    --ol-bg: #262626;\n    --ol-fg: #ececec;\n    --ol-muted: #a0a0a0;\n    --ol-accent: #7aa7ff;\n    --ol-border: rgb(255 255 255 / 14%);\n  }\n}\n\n[hidden] {\n  display: none !important;\n}\n\n.toast,\n.hint,\n.health {\n  position: fixed;\n  z-index: 2147483647;\n  background: var(--ol-bg);\n  color: var(--ol-fg);\n  border: 1px solid var(--ol-border);\n  border-radius: 6px;\n  box-shadow: var(--ol-shadow);\n  padding: 6px 10px;\n}\n\n.toast {\n  bottom: 24px;\n  left: 50%;\n  transform: translateX(-50%);\n}\n\n.hint {\n  bottom: 24px;\n  right: 24px;\n  font-family: ui-monospace, "SF Mono", Menlo, monospace;\n}\n\n.health {\n  top: 8px;\n  right: 8px;\n  display: flex;\n  gap: 8px;\n  align-items: center;\n  border-color: #d97706;\n}\n\n.close {\n  all: unset;\n  cursor: pointer;\n  color: var(--ol-muted);\n  padding: 0 2px;\n}\n\n.backdrop {\n  position: fixed;\n  inset: 0;\n  z-index: 2147483646;\n  background: rgb(0 0 0 / 25%);\n  display: flex;\n  justify-content: center;\n  align-items: flex-start;\n  padding-top: 12vh;\n}\n\n.panel {\n  background: var(--ol-bg);\n  color: var(--ol-fg);\n  border-radius: 10px;\n  box-shadow: var(--ol-shadow);\n  width: min(520px, 90vw);\n  max-height: 70vh;\n  overflow: auto;\n}\n\n.help {\n  padding: 16px 20px;\n}\n\n.help h2 {\n  font-size: 15px;\n  margin: 0 0 12px;\n}\n\n.help table {\n  border-collapse: collapse;\n  width: 100%;\n}\n\n.help td {\n  padding: 3px 0;\n}\n\n.keys {\n  width: 7em;\n}\n\nkbd {\n  font-family: ui-monospace, "SF Mono", Menlo, monospace;\n  border: 1px solid var(--ol-border);\n  border-radius: 4px;\n  padding: 0 5px;\n}\n\n.palette input {\n  all: unset;\n  box-sizing: border-box;\n  width: 100%;\n  padding: 12px 16px;\n  font-size: 15px;\n  border-bottom: 1px solid var(--ol-border);\n}\n\n.palette ul {\n  list-style: none;\n  margin: 0;\n  padding: 4px 0;\n}\n\n.palette li {\n  padding: 5px 16px;\n  cursor: pointer;\n}\n\n.palette li[aria-selected="true"] {\n  background: color-mix(in srgb, var(--ol-accent) 16%, transparent);\n}\n\n.palette li .path {\n  color: var(--ol-muted);\n  margin-left: 6px;\n}\n\n.palette mark {\n  background: none;\n  color: var(--ol-accent);\n  font-weight: 600;\n}\n\n.palette .empty {\n  color: var(--ol-muted);\n  cursor: default;\n}\n';

  // src/ui/overlay.ts
  var shadow = null;
  function overlayRoot() {
    if (shadow) return shadow;
    const host = document.createElement("overlook-root");
    document.documentElement.append(host);
    shadow = host.attachShadow({ mode: "open" });
    for (const type of ["keydown", "keypress", "keyup"]) {
      host.addEventListener(type, (e) => e.stopPropagation());
    }
    const style = document.createElement("style");
    style.textContent = overlay_default;
    shadow.append(style);
    return shadow;
  }
  function isFromOverlay(event) {
    return shadow !== null && event.composedPath().includes(shadow.host);
  }
  function isModalOpen() {
    return shadow?.querySelector(".backdrop") != null;
  }
  function el(tag, className, text) {
    const node = document.createElement(tag);
    node.className = className;
    if (text !== void 0) node.textContent = text;
    return node;
  }
  function singleton(className) {
    const root = overlayRoot();
    const existing = root.querySelector(`.${className}`);
    if (existing) return existing;
    const node = el("div", className);
    root.append(node);
    return node;
  }
  var toastTimer;
  function toast(message, ms = 2500) {
    const node = singleton("toast");
    node.textContent = message;
    node.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => {
      node.hidden = true;
    }, ms);
  }
  function showHint(text) {
    const node = singleton("hint");
    node.hidden = text === null;
    node.textContent = text ?? "";
  }
  function showHealthBadge(broken, onDismiss) {
    const node = singleton("health");
    node.replaceChildren();
    if (broken.length === 0) {
      node.hidden = true;
      return;
    }
    node.hidden = false;
    node.title = `Not found: ${broken.join(", ")}. See the console for details.`;
    const close = el("button", "close", "\xD7");
    close.setAttribute("aria-label", "Dismiss");
    close.addEventListener("click", () => {
      node.hidden = true;
      onDismiss();
    });
    node.append(el("span", "", `OBB: ${broken.length} selector(s) broken`), close);
  }
  function showHelp(rows) {
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
    const onKey = (e) => {
      e.preventDefault();
      e.stopPropagation();
      close();
    };
    backdrop.addEventListener("click", (e) => {
      if (e.target === backdrop) close();
    });
    window.addEventListener("keydown", onKey, true);
  }

  // src/palette/fuzzy.ts
  var isBoundary = (text, i) => i === 0 || /[\s/\\._-]/.test(text[i - 1] ?? "");
  function isSubsequence(query, text, from) {
    let at = from;
    for (const ch of query) {
      at = text.indexOf(ch, at) + 1;
      if (at === 0) return false;
    }
    return true;
  }
  function score(query, text) {
    const q = query.toLowerCase();
    const t = text.toLowerCase();
    if (q.length === 0) return { score: 0, indices: [] };
    const indices = [];
    let total = 0;
    let from = 0;
    for (const [n, ch] of [...q].entries()) {
      const rest = q.slice(n + 1);
      let at = -1;
      for (let i = t.indexOf(ch, from); i >= 0; i = t.indexOf(ch, i + 1)) {
        if (!isSubsequence(rest, t, i + 1)) break;
        if (at < 0) at = i;
        if (isBoundary(text, i)) {
          at = i;
          break;
        }
      }
      if (at < 0) return null;
      const prev = indices.at(-1);
      total += 1;
      if (isBoundary(text, at)) total += 8;
      if (prev !== void 0 && at === prev + 1) total += 5;
      if (prev !== void 0) total -= Math.min(at - prev - 1, 3);
      indices.push(at);
      from = at + 1;
    }
    total -= text.length * 0.01;
    return { score: total, indices };
  }
  function fuzzyFilter(query, items, text) {
    const matches = [];
    for (const item of items) {
      const result = score(query, text(item));
      if (result) matches.push({ item, ...result });
    }
    return matches.sort((a, b) => b.score - a.score);
  }

  // src/palette/palette.ts
  var MAX_RESULTS = 12;
  function highlight(text, indices) {
    const frag = document.createDocumentFragment();
    const hit = new Set(indices);
    for (const [i, ch] of [...text].entries()) {
      frag.append(hit.has(i) ? el("mark", "", ch) : document.createTextNode(ch));
    }
    return frag;
  }
  var APPS = [
    ["Mail", "mailApp"],
    ["Calendar", "calendarApp"],
    ["People", "peopleApp"]
  ];
  function appTargets() {
    return APPS.flatMap(([name, target]) => {
      const element = resolve(target);
      return element ? [{ name, path: name, element }] : [];
    });
  }
  function openFolderPalette() {
    const folders = [...appTargets(), ...readFolderTree()];
    if (folders.length === 0) {
      toast("OBB couldn't find the folder pane");
      return;
    }
    const root = overlayRoot();
    const backdrop = el("div", "backdrop");
    const panel = el("div", "panel palette");
    const input = el("input", "");
    input.placeholder = "Go to folder, or Mail, Calendar, People\u2026";
    input.setAttribute("aria-label", input.placeholder);
    const list = el("ul", "");
    list.setAttribute("role", "listbox");
    panel.append(input, list);
    backdrop.append(panel);
    root.append(backdrop);
    let matches = [];
    let selected = 0;
    const close = () => backdrop.remove();
    const commit = () => {
      const match = matches[selected];
      close();
      if (match) match.item.element.click();
    };
    const render = () => {
      matches = fuzzyFilter(input.value, folders, (f) => f.path).slice(0, MAX_RESULTS);
      selected = Math.min(selected, Math.max(matches.length - 1, 0));
      list.replaceChildren(
        ...matches.map((m, i) => {
          const li = el("li", "");
          li.setAttribute("role", "option");
          li.setAttribute("aria-selected", String(i === selected));
          li.append(highlight(m.item.path, m.indices));
          li.addEventListener("mousedown", (e) => {
            e.preventDefault();
            selected = i;
            commit();
          });
          return li;
        })
      );
      if (matches.length === 0) list.append(el("li", "empty", "No matching folders"));
    };
    input.addEventListener("input", () => {
      selected = 0;
      render();
    });
    input.addEventListener("keydown", (e) => {
      e.stopPropagation();
      if (e.key === "Escape") close();
      else if (e.key === "Enter") commit();
      else if (e.key === "ArrowDown" || e.ctrlKey && e.key === "n") selected++;
      else if (e.key === "ArrowUp" || e.ctrlKey && e.key === "p")
        selected = Math.max(selected - 1, 0);
      else return;
      e.preventDefault();
      render();
    });
    backdrop.addEventListener("mousedown", (e) => {
      if (e.target === backdrop) close();
    });
    render();
    input.focus();
  }

  // src/keymap/execute.ts
  function currentContext() {
    const reading = resolve("readingPane");
    const focused = document.activeElement;
    return reading && focused && reading.contains(focused) ? "reading" : "list";
  }
  function runExtension(command) {
    switch (command) {
      case "help":
        showHelp(
          BINDINGS.map((b) => ({
            keys: b.keys.join(" "),
            description: ACTIONS[b.action].description
          }))
        );
        return;
      case "palette-go":
        openFolderPalette();
        return;
    }
  }
  function execute(id) {
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
        runExtension(strategy.command);
        return true;
      case "todo":
        toast(`${description} isn't implemented yet`);
        return true;
    }
  }

  // src/settings.ts
  var FEATURES = {
    healthIndicator: {
      label: "Show an indicator when Outlook's page changes break OBB",
      group: "General",
      default: true
    },
    keymap: { label: "Fastmail-style keyboard shortcuts", group: "Keyboard", default: true },
    "hide-upsell": { label: "Upgrade and 'try the new\u2026' banners", group: "Declutter", default: true },
    "hide-app-banners": {
      label: "Notices across the top of Outlook (pop-up, sign-in and storage reminders)",
      group: "Declutter",
      default: false
    },
    "hide-info-bars": {
      label: "Notices above an open message ('You replied on\u2026', 'Extra line breaks removed')",
      group: "Declutter",
      default: false
    },
    "hide-copilot": { label: "Copilot buttons and panels", group: "Declutter", default: true },
    "hide-header-extras": {
      label: "Teams and My Day buttons in the header",
      group: "Declutter",
      default: true
    },
    "hide-rail-extras": {
      label: "To Do, Newsletters, Org Explorer and OneDrive in the app rail",
      group: "Declutter",
      default: true
    },
    "hide-ribbon-extras": {
      label: "Teams, Sweep, Quick steps, Viva and add-in buttons in the ribbon",
      group: "Declutter",
      default: true
    },
    "hide-external-tag": { label: "'External' sender tags", group: "Declutter", default: true },
    "hide-row-actions": {
      label: "Flag and pin icons that appear when hovering a message",
      group: "Declutter",
      default: false
    },
    "hide-avatars": {
      label: "Sender pictures in the message list",
      group: "Declutter",
      default: false
    },
    "hide-group-headers": {
      label: "Date group headers ('This Month') in the message list",
      group: "Declutter",
      default: false
    },
    "hide-focused-tabs": { label: "Focused / Other tabs", group: "Declutter", default: false },
    "hide-ribbon-tabs": {
      label: "Ribbon tabs (Home, View, Help)",
      group: "Declutter",
      default: false
    },
    "hide-app-rail": { label: "App rail (left icon bar)", group: "Declutter", default: false },
    "hide-ribbon": { label: "Ribbon", group: "Declutter", default: false },
    density: { label: "Tighter, denser layout", group: "Look", default: true }
  };
  var CHOICES = {
    palette: {
      label: "Colors",
      group: "Look",
      options: {
        outlook: "Outlook's own",
        solarized: "Solarized",
        base16: "Base16 Default",
        fastmail: "Fastmail",
        paper: "Paper"
      },
      default: "outlook"
    },
    mode: {
      label: "Light or dark",
      group: "Look",
      options: { auto: "Follow Outlook", light: "Light", dark: "Dark" },
      default: "auto"
    },
    layout: {
      label: "Layout",
      group: "Look",
      options: {
        outlook: "Outlook's own",
        calm: "Calm three-pane",
        fastmail: "Fastmail (list first)",
        focus: "Focus (single column)"
      },
      default: "outlook"
    }
  };
  var FEATURE_IDS = Object.keys(FEATURES);
  var CHOICE_IDS = Object.keys(CHOICES);
  function defaultSettings() {
    return {
      ...Object.fromEntries(FEATURE_IDS.map((id) => [id, FEATURES[id].default])),
      ...Object.fromEntries(CHOICE_IDS.map((id) => [id, CHOICES[id].default]))
    };
  }
  function normalize(stored) {
    const settings2 = defaultSettings();
    for (const id of FEATURE_IDS) {
      if (typeof stored[id] === "boolean") settings2[id] = stored[id];
    }
    for (const id of CHOICE_IDS) {
      const value = stored[id];
      if (typeof value === "string" && value in CHOICES[id].options) settings2[id] = value;
    }
    return settings2;
  }
  function featureTokens(settings2) {
    return FEATURE_IDS.filter((id) => settings2[id]).join(" ");
  }
  function applyToDocument(settings2, root) {
    root.setAttribute("data-overlook-features", featureTokens(settings2));
    for (const id of CHOICE_IDS) root.setAttribute(`data-overlook-${id}`, settings2[id]);
  }
  async function loadSettings() {
    return normalize(await chrome.storage.sync.get(null));
  }
  function onSettingsChanged(callback) {
    chrome.storage.onChanged.addListener((_changes, area) => {
      if (area === "sync") void loadSettings().then(callback);
    });
  }

  // src/content.ts
  var LOG = "[obb]";
  var APP_READY_TIMEOUT_MS = 2e4;
  var HEALTH_SETTLE_MS = 1e3;
  var HEALTH_RETRY_MS = [3e3, 1e4];
  var settings;
  var healthDismissed = false;
  function applySettings(next) {
    settings = next;
    applyToDocument(settings, document.documentElement);
    syncColorScheme();
    if (!settings.healthIndicator) showHealthBadge([], () => {
    });
  }
  function syncReadingPane() {
    const state = resolve("readingPane") ? "shown" : "none";
    if (document.documentElement.dataset.overlookReadingPane !== state) {
      document.documentElement.dataset.overlookReadingPane = state;
    }
  }
  function syncColorScheme() {
    const scheme = outlookColorScheme(document.documentElement);
    if (document.documentElement.getAttribute("data-overlook-outlook-scheme") !== scheme) {
      document.documentElement.setAttribute("data-overlook-outlook-scheme", scheme);
    }
  }
  function reportHealth(results) {
    const broken = brokenEntries(results);
    const fallbacks = results.filter((r) => r.candidate > 0).map((r) => r.name);
    document.documentElement.dataset.overlookHealth = broken.length ? "broken" : "ok";
    console.groupCollapsed(`${LOG} health: ${broken.length ? `${broken.length} broken` : "ok"}`);
    console.table(results);
    console.groupEnd();
    if (broken.length) console.warn(`${LOG} selectors not found:`, broken.join(", "));
    if (fallbacks.length) console.info(`${LOG} using fallback selectors for:`, fallbacks.join(", "));
    if (settings.healthIndicator && !healthDismissed) {
      showHealthBadge(broken, () => {
        healthDismissed = true;
      });
    }
  }
  var healthTimer;
  function scheduleHealthCheck() {
    clearTimeout(healthTimer);
    if (!location.pathname.includes("/mail")) return;
    const attempt = (n) => {
      const results = healthCheck();
      const retryIn = HEALTH_RETRY_MS[n];
      if (brokenEntries(results).length === 0 || retryIn === void 0) reportHealth(results);
      else healthTimer = setTimeout(() => attempt(n + 1), retryIn);
    };
    healthTimer = setTimeout(() => attempt(0), HEALTH_SETTLE_MS);
  }
  function whenAppReady() {
    return new Promise((done) => {
      if (resolve("messageList")) return done();
      const observer = new MutationObserver(() => {
        if (resolve("messageList")) finish();
      });
      const timer = setTimeout(finish, APP_READY_TIMEOUT_MS);
      function finish() {
        observer.disconnect();
        clearTimeout(timer);
        done();
      }
      observer.observe(document.body, { childList: true, subtree: true });
    });
  }
  function watchPage() {
    let scheduled = false;
    let lastUrl = location.href;
    new MutationObserver(() => {
      if (scheduled) return;
      scheduled = true;
      requestAnimationFrame(() => {
        scheduled = false;
        tagElements();
        syncColorScheme();
        syncReadingPane();
        if (location.href !== lastUrl) {
          lastUrl = location.href;
          scheduleHealthCheck();
        }
      });
    }).observe(document.body, { childList: true, subtree: true });
  }
  var orphanNoticeShown = false;
  function isOrphaned() {
    if (chrome.runtime?.id) return false;
    if (!orphanNoticeShown) {
      orphanNoticeShown = true;
      toast("OBB was updated. Reload this tab to use the new version.", 8e3);
    }
    return true;
  }
  function isEditable(target) {
    if (!(target instanceof HTMLElement)) return false;
    return target.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(target.tagName);
  }
  function mountKeymap() {
    const engine = new SequenceEngine(BINDINGS);
    let hintTimer;
    window.addEventListener(
      "keydown",
      (event) => {
        if (!settings.keymap || !event.isTrusted || isFromOverlay(event) || isModalOpen()) return;
        if (isOrphaned()) return;
        if (isEditable(event.composedPath()[0])) return;
        const key = keyFromEvent(event);
        if (key === null) return;
        const result = engine.feed(key, currentContext(), event.timeStamp);
        if (result.kind === "none") return;
        showHint(
          result.kind === "pending" ? `${result.prefix.join(" ")} \u2026 ${result.next.map((b) => `${b.keys.at(-1)} ${ACTIONS[b.action].description}`).join(" \xB7 ")}` : null
        );
        clearTimeout(hintTimer);
        if (result.kind === "pending")
          hintTimer = setTimeout(() => showHint(null), DEFAULT_TIMEOUT_MS);
        if (result.kind === "match" && !execute(result.action)) return;
        event.preventDefault();
        event.stopImmediatePropagation();
      },
      true
    );
  }
  async function main() {
    applySettings(await loadSettings());
    onSettingsChanged(applySettings);
    mountKeymap();
    await whenAppReady();
    tagElements();
    syncReadingPane();
    scheduleHealthCheck();
    watchPage();
  }
  void main();
})();
//# sourceMappingURL=data:application/json;base64,ewogICJ2ZXJzaW9uIjogMywKICAic291cmNlcyI6IFsiLi4vc3JjL2tleW1hcC9hY3Rpb25zLnRzIiwgIi4uL3NyYy9rZXltYXAvYmluZGluZ3MudHMiLCAiLi4vc3JjL2tleW1hcC9lbmdpbmUudHMiLCAiLi4vc3JjL3NlbGVjdG9ycy50cyIsICIuLi9zcmMvdWkvb3ZlcmxheS5jc3MiLCAiLi4vc3JjL3VpL292ZXJsYXkudHMiLCAiLi4vc3JjL3BhbGV0dGUvZnV6enkudHMiLCAiLi4vc3JjL3BhbGV0dGUvcGFsZXR0ZS50cyIsICIuLi9zcmMva2V5bWFwL2V4ZWN1dGUudHMiLCAiLi4vc3JjL3NldHRpbmdzLnRzIiwgIi4uL3NyYy9jb250ZW50LnRzIl0sCiAgInNvdXJjZXNDb250ZW50IjogWyIvLyBBY3Rpb24gdGFibGU6IHdoYXQgZWFjaCBhY3Rpb24gbWVhbnMgYW5kIGhvdyBpdCBleGVjdXRlcy4gUHVyZSBkYXRhLlxuLy9cbi8vIFN0cmF0ZWdpZXMgKFBMQU4ubWQsIEtleW1hcCBsYXllcik6XG4vLyAgIHBhc3N0aHJvdWdoIFx1MjAxMyBPdXRsb29rJ3MgR21haWwgc2hvcnRjdXQgbW9kZSBhbHJlYWR5IGRvZXMgaXQ7IGxldCB0aGUga2V5IHRocm91Z2guXG4vLyAgIGRyaXZlICAgICAgIFx1MjAxMyBjbGljayBhIGNvbnRyb2wgZm91bmQgdGhyb3VnaCB0aGUgc2VsZWN0b3IgcmVnaXN0cnkuIFJlYWN0IGhvbm91cnMgY2xpY2soKS5cbi8vICAgICAgICAgICAgICAgICAoRGlzcGF0Y2hpbmcgT3V0bG9vaydzIG93biBzaG9ydGN1dCBhcyBhIHN5bnRoZXRpYyBrZXkgZXZlbnQgZGlkIG5vdGhpbmcgb25cbi8vICAgICAgICAgICAgICAgICB0aGUgbGl2ZSBzaXRlIGluIE5vdiAyMDI2LCBzbyB0aGVyZSBpcyBubyBrZXlzdHJva2Ugc3RyYXRlZ3kuKVxuLy8gICB0b2dnbGUgICAgICBcdTIwMTMgY2xpY2sgYSB0b2dnbGUgYnV0dG9uIChSZWFkIC8gVW5yZWFkKSBvbmx5IGlmIHRoZSBzZWxlY3Rpb24gaXNuJ3QgYWxyZWFkeSBpblxuLy8gICAgICAgICAgICAgICAgIHRoZSB3YW50ZWQgc3RhdGUuXG4vLyAgIGV4dGVuc2lvbiAgIFx1MjAxMyBPQkIncyBvd24gVUkgKHBhbGV0dGUsIGhlbHApLlxuLy8gICB0b2RvICAgICAgICBcdTIwMTMgbm90IGRlY2lkZWQgeWV0OyB0aGUga2V5IGlzIHN3YWxsb3dlZCBhbmQgYSBoaW50IHNheXMgc28uXG4vL1xuLy8gT3V0bG9vaydzIEdtYWlsLW1vZGUga2V5cyBjb21lIGZyb20gaXRzIG93biByaWJib24gdG9vbHRpcHMgaW4gdGhlIE9jdCAyMDI2IGNhcHR1cmU6XG4vLyBDIG5ldyBtYWlsLCAjIGRlbGV0ZSwgRSBhcmNoaXZlLCBWIG1vdmUsIFx1MjFFN0kvXHUyMUU3VSByZWFkL3VucmVhZCwgQiBzbm9vemUsIEwgY2F0ZWdvcml6ZSxcbi8vIFIvQS9GIHJlcGx5L3JlcGx5IGFsbC9mb3J3YXJkLCBcdTIzMThaIHVuZG8uXG5cbmltcG9ydCB0eXBlIHsgU2VsZWN0b3JOYW1lIH0gZnJvbSBcIi4uL3NlbGVjdG9ycy50c1wiO1xuXG5leHBvcnQgdHlwZSBFeHRlbnNpb25Db21tYW5kID0gXCJoZWxwXCIgfCBcInBhbGV0dGUtZ29cIjtcblxuZXhwb3J0IHR5cGUgU3RyYXRlZ3kgPVxuICB8IHsga2luZDogXCJwYXNzdGhyb3VnaFwiIH1cbiAgfCB7IGtpbmQ6IFwiZHJpdmVcIjsgdGFyZ2V0OiBTZWxlY3Rvck5hbWUgfVxuICB8IHsga2luZDogXCJ0b2dnbGVcIjsgdGFyZ2V0OiBTZWxlY3Rvck5hbWU7IHdhbnQ6IFwicmVhZFwiIHwgXCJ1bnJlYWRcIiB9XG4gIHwgeyBraW5kOiBcImV4dGVuc2lvblwiOyBjb21tYW5kOiBFeHRlbnNpb25Db21tYW5kIH1cbiAgfCB7IGtpbmQ6IFwidG9kb1wiIH07XG5cbmV4cG9ydCBpbnRlcmZhY2UgQWN0aW9uIHtcbiAgZGVzY3JpcHRpb246IHN0cmluZztcbiAgc3RyYXRlZ3k6IFN0cmF0ZWd5O1xufVxuXG5jb25zdCBwYXNzdGhyb3VnaDogU3RyYXRlZ3kgPSB7IGtpbmQ6IFwicGFzc3Rocm91Z2hcIiB9O1xuY29uc3QgdG9kbzogU3RyYXRlZ3kgPSB7IGtpbmQ6IFwidG9kb1wiIH07XG5jb25zdCBkcml2ZSA9ICh0YXJnZXQ6IFNlbGVjdG9yTmFtZSk6IFN0cmF0ZWd5ID0+ICh7IGtpbmQ6IFwiZHJpdmVcIiwgdGFyZ2V0IH0pO1xuY29uc3QgZXh0ZW5zaW9uID0gKGNvbW1hbmQ6IEV4dGVuc2lvbkNvbW1hbmQpOiBTdHJhdGVneSA9PiAoeyBraW5kOiBcImV4dGVuc2lvblwiLCBjb21tYW5kIH0pO1xuY29uc3QgcmVhZFN0YXRlID0gKHdhbnQ6IFwicmVhZFwiIHwgXCJ1bnJlYWRcIik6IFN0cmF0ZWd5ID0+ICh7XG4gIGtpbmQ6IFwidG9nZ2xlXCIsXG4gIHRhcmdldDogXCJyZWFkVW5yZWFkQnV0dG9uXCIsXG4gIHdhbnQsXG59KTtcblxuZXhwb3J0IGNvbnN0IEFDVElPTlMgPSB7XG4gIGNvbXBvc2U6IHsgZGVzY3JpcHRpb246IFwiQ29tcG9zZVwiLCBzdHJhdGVneTogcGFzc3Rocm91Z2ggfSxcbiAgc2VhcmNoOiB7IGRlc2NyaXB0aW9uOiBcIkZvY3VzIHNlYXJjaFwiLCBzdHJhdGVneTogcGFzc3Rocm91Z2ggfSxcbiAgaGVscDogeyBkZXNjcmlwdGlvbjogXCJTaG93IGtleWJvYXJkIHNob3J0Y3V0c1wiLCBzdHJhdGVneTogZXh0ZW5zaW9uKFwiaGVscFwiKSB9LFxuICBuZXh0OiB7IGRlc2NyaXB0aW9uOiBcIk5leHQgY29udmVyc2F0aW9uXCIsIHN0cmF0ZWd5OiBwYXNzdGhyb3VnaCB9LFxuICBwcmV2aW91czogeyBkZXNjcmlwdGlvbjogXCJQcmV2aW91cyBjb252ZXJzYXRpb25cIiwgc3RyYXRlZ3k6IHBhc3N0aHJvdWdoIH0sXG4gIG9wZW46IHsgZGVzY3JpcHRpb246IFwiT3BlbiBjb252ZXJzYXRpb25cIiwgc3RyYXRlZ3k6IHBhc3N0aHJvdWdoIH0sXG4gIHNlbGVjdDogeyBkZXNjcmlwdGlvbjogXCJTZWxlY3Qgb3IgZGVzZWxlY3RcIiwgc3RyYXRlZ3k6IHBhc3N0aHJvdWdoIH0sXG4gIHNlbGVjdFJhbmdlOiB7IGRlc2NyaXB0aW9uOiBcIlNlbGVjdCByYW5nZSBzaW5jZSBsYXN0IHNlbGVjdGlvblwiLCBzdHJhdGVneTogdG9kbyB9LFxuICBzZWxlY3RBbGw6IHsgZGVzY3JpcHRpb246IFwiU2VsZWN0IGFsbFwiLCBzdHJhdGVneTogdG9kbyB9LFxuICBzZWxlY3ROb25lOiB7IGRlc2NyaXB0aW9uOiBcIlNlbGVjdCBub25lXCIsIHN0cmF0ZWd5OiB0b2RvIH0sXG4gIGFyY2hpdmU6IHsgZGVzY3JpcHRpb246IFwiQXJjaGl2ZVwiLCBzdHJhdGVneTogZHJpdmUoXCJhcmNoaXZlQnV0dG9uXCIpIH0sXG4gIGFyY2hpdmVQcmV2aW91czogeyBkZXNjcmlwdGlvbjogXCJBcmNoaXZlLCB0aGVuIG9wZW4gcHJldmlvdXNcIiwgc3RyYXRlZ3k6IHRvZG8gfSxcbiAgYXJjaGl2ZU5leHQ6IHsgZGVzY3JpcHRpb246IFwiQXJjaGl2ZSwgdGhlbiBvcGVuIG5leHRcIiwgc3RyYXRlZ3k6IHRvZG8gfSxcbiAgZGVsZXRlOiB7IGRlc2NyaXB0aW9uOiBcIkRlbGV0ZVwiLCBzdHJhdGVneTogZHJpdmUoXCJkZWxldGVCdXR0b25cIikgfSxcbiAganVuazogeyBkZXNjcmlwdGlvbjogXCJSZXBvcnQganVua1wiLCBzdHJhdGVneTogcGFzc3Rocm91Z2ggfSxcbiAgLy8gT3V0bG9vaydzIG93biBcIk1vdmUgdG9cIiBwaWNrZXIgaGFzIGZvbGRlciBzZWFyY2gsIHdoaWNoIGNvdmVycyBGYXN0bWFpbCdzIGBtYCB1bnRpbCB0aGVcbiAgLy8gcGFsZXR0ZSBjYW4gbW92ZSBtZXNzYWdlcyBpdHNlbGYgKGRvY3MvZGVjaXNpb25zLm1kKS5cbiAgbW92ZTogeyBkZXNjcmlwdGlvbjogXCJNb3ZlIHRvIGZvbGRlclx1MjAyNlwiLCBzdHJhdGVneTogZHJpdmUoXCJtb3ZlQnV0dG9uXCIpIH0sXG4gIGdvVG86IHsgZGVzY3JpcHRpb246IFwiR28gdG8gZm9sZGVyIG9yIGFwcFx1MjAyNlwiLCBzdHJhdGVneTogZXh0ZW5zaW9uKFwicGFsZXR0ZS1nb1wiKSB9LFxuICBzbm9vemU6IHsgZGVzY3JpcHRpb246IFwiU25vb3plXCIsIHN0cmF0ZWd5OiBwYXNzdGhyb3VnaCB9LFxuICBtYXJrUmVhZDogeyBkZXNjcmlwdGlvbjogXCJNYXJrIHJlYWRcIiwgc3RyYXRlZ3k6IHJlYWRTdGF0ZShcInJlYWRcIikgfSxcbiAgbWFya1VucmVhZDogeyBkZXNjcmlwdGlvbjogXCJNYXJrIHVucmVhZFwiLCBzdHJhdGVneTogcmVhZFN0YXRlKFwidW5yZWFkXCIpIH0sXG4gIHJlcGx5OiB7IGRlc2NyaXB0aW9uOiBcIlJlcGx5XCIsIHN0cmF0ZWd5OiBwYXNzdGhyb3VnaCB9LFxuICByZXBseUFsbDogeyBkZXNjcmlwdGlvbjogXCJSZXBseSBhbGxcIiwgc3RyYXRlZ3k6IHBhc3N0aHJvdWdoIH0sXG4gIGZvcndhcmQ6IHsgZGVzY3JpcHRpb246IFwiRm9yd2FyZFwiLCBzdHJhdGVneTogcGFzc3Rocm91Z2ggfSxcbiAgdW5kbzogeyBkZXNjcmlwdGlvbjogXCJVbmRvXCIsIHN0cmF0ZWd5OiBkcml2ZShcInVuZG9CdXR0b25cIikgfSxcbiAgcmVmcmVzaDogeyBkZXNjcmlwdGlvbjogXCJSZWZyZXNoXCIsIHN0cmF0ZWd5OiBkcml2ZShcInN5bmNCdXR0b25cIikgfSxcbiAgZ29NYWlsOiB7IGRlc2NyaXB0aW9uOiBcIkdvIHRvIE1haWxcIiwgc3RyYXRlZ3k6IGRyaXZlKFwibWFpbEFwcFwiKSB9LFxuICBnb0NhbGVuZGFyOiB7IGRlc2NyaXB0aW9uOiBcIkdvIHRvIENhbGVuZGFyXCIsIHN0cmF0ZWd5OiBkcml2ZShcImNhbGVuZGFyQXBwXCIpIH0sXG4gIGdvQ29udGFjdHM6IHsgZGVzY3JpcHRpb246IFwiR28gdG8gQ29udGFjdHNcIiwgc3RyYXRlZ3k6IGRyaXZlKFwicGVvcGxlQXBwXCIpIH0sXG4gIGdvU2V0dGluZ3M6IHsgZGVzY3JpcHRpb246IFwiR28gdG8gU2V0dGluZ3NcIiwgc3RyYXRlZ3k6IGRyaXZlKFwic2V0dGluZ3NCdXR0b25cIikgfSxcbn0gYXMgY29uc3Qgc2F0aXNmaWVzIFJlY29yZDxzdHJpbmcsIEFjdGlvbj47XG5cbmV4cG9ydCB0eXBlIEFjdGlvbklkID0ga2V5b2YgdHlwZW9mIEFDVElPTlM7XG4iLCAiLy8gVGhlIEZhc3RtYWlsLXN0eWxlIGJpbmRpbmcgc2V0IChQTEFOLm1kLCBTaG9ydGN1dCBtYXApLiBLZXlzIGFyZSBLZXlib2FyZEV2ZW50LmtleSB2YWx1ZXMuXG5cbmltcG9ydCB0eXBlIHsgQWN0aW9uSWQgfSBmcm9tIFwiLi9hY3Rpb25zLnRzXCI7XG5pbXBvcnQgdHlwZSB7IEJpbmRpbmcsIENvbnRleHQgfSBmcm9tIFwiLi9lbmdpbmUudHNcIjtcblxuY29uc3QgTUFJTDogcmVhZG9ubHkgQ29udGV4dFtdID0gW1wibGlzdFwiLCBcInJlYWRpbmdcIl07XG5jb25zdCBMSVNUOiByZWFkb25seSBDb250ZXh0W10gPSBbXCJsaXN0XCJdO1xuY29uc3QgUkVBRElORzogcmVhZG9ubHkgQ29udGV4dFtdID0gW1wicmVhZGluZ1wiXTtcblxuY29uc3QgYmluZCA9IChrZXlzOiBzdHJpbmcsIGFjdGlvbjogQWN0aW9uSWQsIGNvbnRleHRzID0gTUFJTCk6IEJpbmRpbmc8QWN0aW9uSWQ+ID0+ICh7XG4gIGtleXM6IGtleXMuc3BsaXQoXCIgXCIpLFxuICBhY3Rpb24sXG4gIGNvbnRleHRzLFxufSk7XG5cbmV4cG9ydCBjb25zdCBCSU5ESU5HUzogcmVhZG9ubHkgQmluZGluZzxBY3Rpb25JZD5bXSA9IFtcbiAgYmluZChcImNcIiwgXCJjb21wb3NlXCIpLFxuICBiaW5kKFwiL1wiLCBcInNlYXJjaFwiKSxcbiAgYmluZChcIj9cIiwgXCJoZWxwXCIpLFxuICBiaW5kKFwialwiLCBcIm5leHRcIiksXG4gIGJpbmQoXCJrXCIsIFwicHJldmlvdXNcIiksXG4gIGJpbmQoXCJvXCIsIFwib3BlblwiLCBMSVNUKSxcbiAgYmluZChcIkVudGVyXCIsIFwib3BlblwiLCBMSVNUKSxcbiAgYmluZChcInhcIiwgXCJzZWxlY3RcIiwgTElTVCksXG4gIGJpbmQoXCJYXCIsIFwic2VsZWN0UmFuZ2VcIiwgTElTVCksXG4gIGJpbmQoXCIqIGFcIiwgXCJzZWxlY3RBbGxcIiwgTElTVCksXG4gIGJpbmQoXCIqIG5cIiwgXCJzZWxlY3ROb25lXCIsIExJU1QpLFxuICBiaW5kKFwieVwiLCBcImFyY2hpdmVcIiksXG4gIGJpbmQoXCJdXCIsIFwiYXJjaGl2ZVByZXZpb3VzXCIsIFJFQURJTkcpLFxuICBiaW5kKFwiW1wiLCBcImFyY2hpdmVOZXh0XCIsIFJFQURJTkcpLFxuICBiaW5kKFwiZFwiLCBcImRlbGV0ZVwiKSxcbiAgYmluZChcIiFcIiwgXCJqdW5rXCIpLFxuICBiaW5kKFwibVwiLCBcIm1vdmVcIiksXG4gIGJpbmQoXCJnXCIsIFwiZ29Ub1wiKSxcbiAgYmluZChcImJcIiwgXCJzbm9vemVcIiksXG4gIGJpbmQoXCIuIHJcIiwgXCJtYXJrUmVhZFwiKSxcbiAgYmluZChcIi4gdVwiLCBcIm1hcmtVbnJlYWRcIiksXG4gIGJpbmQoXCJyXCIsIFwicmVwbHlcIiwgUkVBRElORyksXG4gIGJpbmQoXCJhXCIsIFwicmVwbHlBbGxcIiwgUkVBRElORyksXG4gIGJpbmQoXCJmXCIsIFwiZm9yd2FyZFwiLCBSRUFESU5HKSxcbiAgYmluZChcInpcIiwgXCJ1bmRvXCIpLFxuICBiaW5kKFwidVwiLCBcInJlZnJlc2hcIiksXG4gIGJpbmQoXCJHIG1cIiwgXCJnb01haWxcIiksXG4gIGJpbmQoXCJHIGNcIiwgXCJnb0NhbGVuZGFyXCIpLFxuICBiaW5kKFwiRyBhXCIsIFwiZ29Db250YWN0c1wiKSxcbiAgYmluZChcIkcgc1wiLCBcImdvU2V0dGluZ3NcIiksXG5dO1xuIiwgIi8vIFB1cmUga2V5LXNlcXVlbmNlIGVuZ2luZS4gTm8gRE9NIGFjY2Vzcywgc28gaXQgaXMgZnVsbHkgdW5pdC10ZXN0YWJsZS5cbi8vXG4vLyBLZXlzIGFyZSBLZXlib2FyZEV2ZW50LmtleSB2YWx1ZXMsIHdoaWNoIGFscmVhZHkgcmVmbGVjdCBTaGlmdCAoXCJHXCIsIFwiKlwiLCBcIiFcIiwgXCI/XCIpLlxuLy8gQSBiaW5kaW5nIGxpa2UgRmFzdG1haWwncyBcIlNoaWZ0K0cgdGhlbiBNXCIgaXMgd3JpdHRlbiBbXCJHXCIsIFwibVwiXS5cblxuZXhwb3J0IHR5cGUgQ29udGV4dCA9IFwibGlzdFwiIHwgXCJyZWFkaW5nXCIgfCBcInBhbGV0dGVcIjtcblxuZXhwb3J0IGludGVyZmFjZSBCaW5kaW5nPEEgZXh0ZW5kcyBzdHJpbmcgPSBzdHJpbmc+IHtcbiAga2V5czogcmVhZG9ubHkgc3RyaW5nW107XG4gIGFjdGlvbjogQTtcbiAgY29udGV4dHM6IHJlYWRvbmx5IENvbnRleHRbXTtcbn1cblxuZXhwb3J0IHR5cGUgRmVlZFJlc3VsdDxBIGV4dGVuZHMgc3RyaW5nPiA9XG4gIHwgeyBraW5kOiBcIm1hdGNoXCI7IGFjdGlvbjogQSB9XG4gIHwgeyBraW5kOiBcInBlbmRpbmdcIjsgcHJlZml4OiByZWFkb25seSBzdHJpbmdbXTsgbmV4dDogcmVhZG9ubHkgQmluZGluZzxBPltdIH1cbiAgLyoqIEEgcHJlZml4IHdhcyBwZW5kaW5nIGFuZCB0aGlzIGtleSBjb21wbGV0ZWQgbm90aGluZy4gVGhlIGtleSBzaG91bGQgYmUgc3dhbGxvd2VkLiAqL1xuICB8IHsga2luZDogXCJjYW5jZWxsZWRcIiB9XG4gIHwgeyBraW5kOiBcIm5vbmVcIiB9O1xuXG5leHBvcnQgY29uc3QgREVGQVVMVF9USU1FT1VUX01TID0gMTUwMDtcblxuY29uc3Qgc3RhcnRzV2l0aCA9IChrZXlzOiByZWFkb25seSBzdHJpbmdbXSwgcHJlZml4OiByZWFkb25seSBzdHJpbmdbXSkgPT5cbiAgcHJlZml4Lmxlbmd0aCA8PSBrZXlzLmxlbmd0aCAmJiBwcmVmaXguZXZlcnkoKGssIGkpID0+IGtleXNbaV0gPT09IGspO1xuXG5leHBvcnQgY2xhc3MgU2VxdWVuY2VFbmdpbmU8QSBleHRlbmRzIHN0cmluZz4ge1xuICByZWFkb25seSAjYmluZGluZ3M6IHJlYWRvbmx5IEJpbmRpbmc8QT5bXTtcbiAgcmVhZG9ubHkgI3RpbWVvdXRNczogbnVtYmVyO1xuICAjcHJlZml4OiBzdHJpbmdbXSA9IFtdO1xuICAjbGFzdEF0ID0gMDtcblxuICBjb25zdHJ1Y3RvcihiaW5kaW5nczogcmVhZG9ubHkgQmluZGluZzxBPltdLCB0aW1lb3V0TXMgPSBERUZBVUxUX1RJTUVPVVRfTVMpIHtcbiAgICB0aGlzLiNiaW5kaW5ncyA9IGJpbmRpbmdzO1xuICAgIHRoaXMuI3RpbWVvdXRNcyA9IHRpbWVvdXRNcztcbiAgfVxuXG4gIGdldCBwZW5kaW5nKCk6IHJlYWRvbmx5IHN0cmluZ1tdIHtcbiAgICByZXR1cm4gdGhpcy4jcHJlZml4O1xuICB9XG5cbiAgcmVzZXQoKTogdm9pZCB7XG4gICAgdGhpcy4jcHJlZml4ID0gW107XG4gIH1cblxuICBmZWVkKGtleTogc3RyaW5nLCBjb250ZXh0OiBDb250ZXh0LCBub3c6IG51bWJlcik6IEZlZWRSZXN1bHQ8QT4ge1xuICAgIGlmICh0aGlzLiNwcmVmaXgubGVuZ3RoID4gMCAmJiBub3cgLSB0aGlzLiNsYXN0QXQgPiB0aGlzLiN0aW1lb3V0TXMpIHRoaXMucmVzZXQoKTtcbiAgICBjb25zdCB3YXNQZW5kaW5nID0gdGhpcy4jcHJlZml4Lmxlbmd0aCA+IDA7XG4gICAgY29uc3Qgc2VxID0gWy4uLnRoaXMuI3ByZWZpeCwga2V5XTtcbiAgICBjb25zdCBhY3RpdmUgPSB0aGlzLiNiaW5kaW5ncy5maWx0ZXIoKGIpID0+IGIuY29udGV4dHMuaW5jbHVkZXMoY29udGV4dCkpO1xuICAgIGNvbnN0IGxvbmdlciA9IGFjdGl2ZS5maWx0ZXIoKGIpID0+IGIua2V5cy5sZW5ndGggPiBzZXEubGVuZ3RoICYmIHN0YXJ0c1dpdGgoYi5rZXlzLCBzZXEpKTtcblxuICAgIGlmIChsb25nZXIubGVuZ3RoID4gMCkge1xuICAgICAgdGhpcy4jcHJlZml4ID0gc2VxO1xuICAgICAgdGhpcy4jbGFzdEF0ID0gbm93O1xuICAgICAgcmV0dXJuIHsga2luZDogXCJwZW5kaW5nXCIsIHByZWZpeDogc2VxLCBuZXh0OiBsb25nZXIgfTtcbiAgICB9XG5cbiAgICB0aGlzLnJlc2V0KCk7XG4gICAgY29uc3QgZXhhY3QgPSBhY3RpdmUuZmluZCgoYikgPT4gYi5rZXlzLmxlbmd0aCA9PT0gc2VxLmxlbmd0aCAmJiBzdGFydHNXaXRoKGIua2V5cywgc2VxKSk7XG4gICAgaWYgKGV4YWN0KSByZXR1cm4geyBraW5kOiBcIm1hdGNoXCIsIGFjdGlvbjogZXhhY3QuYWN0aW9uIH07XG4gICAgcmV0dXJuIHdhc1BlbmRpbmcgPyB7IGtpbmQ6IFwiY2FuY2VsbGVkXCIgfSA6IHsga2luZDogXCJub25lXCIgfTtcbiAgfVxufVxuXG4vKiogQmluZGluZ3MgdGhhdCBzaGFkb3cgZWFjaCBvdGhlciAoc2FtZSBrZXlzLCBvciBvbmUgYSBwcmVmaXggb2YgYW5vdGhlcikgaW4gYSBzaGFyZWQgY29udGV4dC4gKi9cbmV4cG9ydCBmdW5jdGlvbiBmaW5kQ29uZmxpY3RzKGJpbmRpbmdzOiByZWFkb25seSBCaW5kaW5nW10pOiBzdHJpbmdbXSB7XG4gIGNvbnN0IGNvbmZsaWN0czogc3RyaW5nW10gPSBbXTtcbiAgYmluZGluZ3MuZm9yRWFjaCgoYSwgaSkgPT4ge1xuICAgIGZvciAoY29uc3QgYiBvZiBiaW5kaW5ncy5zbGljZShpICsgMSkpIHtcbiAgICAgIGNvbnN0IHNoYXJlZCA9IGEuY29udGV4dHMuc29tZSgoYykgPT4gYi5jb250ZXh0cy5pbmNsdWRlcyhjKSk7XG4gICAgICBjb25zdCBbc2hvcnQsIGxvbmddID0gYS5rZXlzLmxlbmd0aCA8PSBiLmtleXMubGVuZ3RoID8gW2EsIGJdIDogW2IsIGFdO1xuICAgICAgaWYgKHNoYXJlZCAmJiBzdGFydHNXaXRoKGxvbmcua2V5cywgc2hvcnQua2V5cykpIHtcbiAgICAgICAgY29uZmxpY3RzLnB1c2goXG4gICAgICAgICAgYCR7c2hvcnQua2V5cy5qb2luKFwiIFwiKX0gKCR7c2hvcnQuYWN0aW9ufSkgdnMgJHtsb25nLmtleXMuam9pbihcIiBcIil9ICgke2xvbmcuYWN0aW9ufSlgLFxuICAgICAgICApO1xuICAgICAgfVxuICAgIH1cbiAgfSk7XG4gIHJldHVybiBjb25mbGljdHM7XG59XG5cbmludGVyZmFjZSBLZXlMaWtlIHtcbiAga2V5OiBzdHJpbmc7XG4gIGN0cmxLZXk6IGJvb2xlYW47XG4gIG1ldGFLZXk6IGJvb2xlYW47XG4gIGFsdEtleTogYm9vbGVhbjtcbn1cblxuY29uc3QgTU9ESUZJRVJfS0VZUyA9IG5ldyBTZXQoW1wiU2hpZnRcIiwgXCJDb250cm9sXCIsIFwiQWx0XCIsIFwiTWV0YVwiLCBcIkNhcHNMb2NrXCIsIFwiRGVhZFwiXSk7XG5cbi8qKiBUaGUgZW5naW5lIGtleSBmb3IgYW4gZXZlbnQsIG9yIG51bGwgd2hlbiBpdCBiZWxvbmdzIHRvIHRoZSBicm93c2VyL09TIChDdHJsLCBDbWQsIEFsdCBjaG9yZHMpLiAqL1xuZXhwb3J0IGZ1bmN0aW9uIGtleUZyb21FdmVudChldmVudDogS2V5TGlrZSk6IHN0cmluZyB8IG51bGwge1xuICBpZiAoZXZlbnQuY3RybEtleSB8fCBldmVudC5tZXRhS2V5IHx8IGV2ZW50LmFsdEtleSkgcmV0dXJuIG51bGw7XG4gIGlmIChNT0RJRklFUl9LRVlTLmhhcyhldmVudC5rZXkpKSByZXR1cm4gbnVsbDtcbiAgcmV0dXJuIGV2ZW50LmtleTtcbn1cbiIsICIvLyBUaGUgb25seSBmaWxlIGFsbG93ZWQgdG8ga25vdyB3aGF0IE91dGxvb2sncyBET00gbG9va3MgbGlrZS5cbi8vXG4vLyBBbmNob3IgcHJlZmVyZW5jZSAoc2VlIFBMQU4ubWQpOiBzdHJ1Y3R1cmFsIElEcyA+IEFSSUEgcm9sZXMvbGFiZWxzID4gc2VtYW50aWMgZGF0YS0qID4gdGV4dC5cbi8vIE5ldmVyIHVzZSBnZW5lcmF0ZWQgY2xhc3MgbmFtZXM7IGlmIG9uZSBpcyB1bmF2b2lkYWJsZSwgbWFyayBpdCBgLy8gRlJBR0lMRTpgLlxuLy9cbi8vIFZlcmlmaWVkIGFnYWluc3QgYSBzYXZlZCBwYWdlIGZyb20gT2N0IDIwMjYgKHRlc3QvZml4dHVyZXMvaW5ib3guaHRtbCk7IHNlZSBkb2NzL2RvbS1zdXJ2ZXkubWQuXG4vLyBFbnRyaWVzIG1hcmtlZCBVTlZFUklGSUVEIHdlcmUgbm90IHByZXNlbnQgaW4gdGhhdCBjYXB0dXJlLiBhcmlhLWxhYmVscyBhc3N1bWUgYW4gRW5nbGlzaCBVSS5cblxuZXhwb3J0IHR5cGUgU2VsZWN0b3JOYW1lID0ga2V5b2YgdHlwZW9mIFNFTEVDVE9SUztcblxuZXhwb3J0IGludGVyZmFjZSBTZWxlY3RvckVudHJ5IHtcbiAgLyoqIENhbmRpZGF0ZSBDU1Mgc2VsZWN0b3JzLCBtb3N0IHN0YWJsZSBmaXJzdC4gVGhlIGZpcnN0IG9uZSB0aGF0IG1hdGNoZXMgd2lucy4gKi9cbiAgY2FuZGlkYXRlczogcmVhZG9ubHkgc3RyaW5nW107XG4gIC8qKiBcImFsd2F5c1wiOiBzaG91bGQgZXhpc3Qgb24gZXZlcnkgbWFpbCB2aWV3LCBzbyBhIG1pc3MgaXMgYSBicmVha2FnZS4gXCJzb21ldGltZXNcIjogY29udGV4dHVhbC4gKi9cbiAgZXhwZWN0OiBcImFsd2F5c1wiIHwgXCJzb21ldGltZXNcIjtcbiAgLyoqIE5hbWUgb2YgYW5vdGhlciBlbnRyeSB0byByZXNvbHZlIGluc2lkZSwgaW5zdGVhZCBvZiB0aGUgd2hvbGUgZG9jdW1lbnQuICovXG4gIHdpdGhpbj86IHN0cmluZztcbiAgLyoqIFRva2VuIGFkZGVkIHRvIHRoZSBlbGVtZW50J3MgYGRhdGEtb3Zlcmxvb2tgIGF0dHJpYnV0ZSBzbyBDU1MgY2FuIHRhcmdldCBpdC4gKi9cbiAgdGFnPzogc3RyaW5nO1xuICAvKiogVGFnIGV2ZXJ5IG1hdGNoIHJhdGhlciB0aGFuIG9ubHkgdGhlIGZpcnN0LiAqL1xuICBhbGw/OiBib29sZWFuO1xuICAvKiogTWF0Y2ggZXZlcnkgY2FuZGlkYXRlIChhIGxpc3Qgb2YgZGlzdGluY3QgdGhpbmdzKSwgbm90IGp1c3QgdGhlIGZpcnN0IHRoYXQgbWF0Y2hlcy4gKi9cbiAgdW5pb24/OiBib29sZWFuO1xufVxuXG5leHBvcnQgY29uc3QgU0VMRUNUT1JTID0ge1xuICAvLyAtLS0gU3RydWN0dXJlIC0tLVxuICB0aXRsZUJhcjogeyBjYW5kaWRhdGVzOiBbXCIjT3dhVGl0bGVCYXJcIl0sIGV4cGVjdDogXCJhbHdheXNcIiwgdGFnOiBcInRpdGxlLWJhclwiIH0sXG4gIHNlYXJjaEJveDogeyBjYW5kaWRhdGVzOiBbXCIjdG9wU2VhcmNoSW5wdXRcIiwgJ1tyb2xlPVwic2VhcmNoXCJdIGlucHV0J10sIGV4cGVjdDogXCJhbHdheXNcIiB9LFxuICBhcHBSYWlsOiB7XG4gICAgY2FuZGlkYXRlczogW1wiI0xlZnRSYWlsXCIsICdbcm9sZT1cIm5hdmlnYXRpb25cIl1bYXJpYS1sYWJlbD1cImxlZnQtcmFpbC1hcHBiYXJcIl0nXSxcbiAgICBleHBlY3Q6IFwiYWx3YXlzXCIsXG4gICAgdGFnOiBcImFwcC1yYWlsXCIsXG4gIH0sXG4gIHJpYmJvbjoge1xuICAgIGNhbmRpZGF0ZXM6IFsnW2RhdGEtYXBwLXNlY3Rpb249XCJSaWJib25cIl0nLCAnW3JvbGU9XCJyZWdpb25cIl1bYXJpYS1sYWJlbD1cIlJpYmJvblwiXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB0YWc6IFwicmliYm9uXCIsXG4gIH0sXG4gIHJpYmJvblRhYnM6IHtcbiAgICBjYW5kaWRhdGVzOiBbJ1tyb2xlPVwidGFibGlzdFwiXVthcmlhLWxhYmVsPVwiUmliYm9uIHRhYnNcIl0nXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgd2l0aGluOiBcInJpYmJvblwiLFxuICAgIHRhZzogXCJyaWJib24tdGFic1wiLFxuICB9LFxuICBmb2xkZXJQYW5lOiB7XG4gICAgY2FuZGlkYXRlczogW1xuICAgICAgJ1tkYXRhLWFwcC1zZWN0aW9uPVwiTmF2aWdhdGlvblBhbmVcIl0nLFxuICAgICAgJ1tyb2xlPVwicmVnaW9uXCJdW2FyaWEtbGFiZWw9XCJOYXZpZ2F0aW9uIHBhbmVcIl0nLFxuICAgIF0sXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICAgIHRhZzogXCJmb2xkZXItcGFuZVwiLFxuICB9LFxuICBmb2xkZXJJdGVtOiB7XG4gICAgY2FuZGlkYXRlczogWydbcm9sZT1cInRyZWVpdGVtXCJdW2RhdGEtZm9sZGVyLW5hbWVdJywgJ1tyb2xlPVwidHJlZWl0ZW1cIl0nXSxcbiAgICBleHBlY3Q6IFwiYWx3YXlzXCIsXG4gICAgd2l0aGluOiBcImZvbGRlclBhbmVcIixcbiAgICB0YWc6IFwiZm9sZGVyLWl0ZW1cIixcbiAgICBhbGw6IHRydWUsXG4gIH0sXG4gIGxpc3RQYW5lOiB7XG4gICAgY2FuZGlkYXRlczogWydbZGF0YS1hcHAtc2VjdGlvbj1cIk1lc3NhZ2VMaXN0XCJdJ10sXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICAgIHRhZzogXCJsaXN0LXBhbmVcIixcbiAgfSxcbiAgbWVzc2FnZUxpc3Q6IHtcbiAgICBjYW5kaWRhdGVzOiBbXCIjTWFpbExpc3RcIiwgJ1tyb2xlPVwibGlzdGJveFwiXVthcmlhLWxhYmVsPVwiTWVzc2FnZSBsaXN0XCJdJ10sXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICAgIHRhZzogXCJtZXNzYWdlLWxpc3RcIixcbiAgfSxcbiAgbWVzc2FnZVJvdzoge1xuICAgIGNhbmRpZGF0ZXM6IFsnW3JvbGU9XCJvcHRpb25cIl1bZGF0YS1jb252aWRdJywgJ1tyb2xlPVwib3B0aW9uXCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJtZXNzYWdlTGlzdFwiLFxuICAgIHRhZzogXCJtZXNzYWdlLXJvd1wiLFxuICAgIGFsbDogdHJ1ZSxcbiAgfSxcbiAgZ3JvdXBIZWFkZXI6IHtcbiAgICBjYW5kaWRhdGVzOiBbJ1tyb2xlPVwiYnV0dG9uXCJdW2lkXj1cImdyb3VwSGVhZGVyXCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJtZXNzYWdlTGlzdFwiLFxuICAgIHRhZzogXCJncm91cC1oZWFkZXJcIixcbiAgICBhbGw6IHRydWUsXG4gIH0sXG4gIGZvY3VzZWRUYWJzOiB7XG4gICAgY2FuZGlkYXRlczogWydbcm9sZT1cInRhYmxpc3RcIl0nXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgd2l0aGluOiBcImxpc3RQYW5lXCIsXG4gICAgdGFnOiBcImZvY3VzZWQtdGFic1wiLFxuICB9LFxuICByZWFkaW5nUGFuZToge1xuICAgIGNhbmRpZGF0ZXM6IFtcbiAgICAgICdbZGF0YS1hcHAtc2VjdGlvbj1cIk1haWxSZWFkQ29tcG9zZVwiXScsXG4gICAgICBcIiNSZWFkaW5nUGFuZUNvbnRhaW5lcklkXCIsXG4gICAgICAnW3JvbGU9XCJtYWluXCJdW2FyaWEtbGFiZWw9XCJSZWFkaW5nIFBhbmVcIl0nLFxuICAgIF0sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHRhZzogXCJyZWFkaW5nLXBhbmVcIixcbiAgfSxcblxuICBtZXNzYWdlQm9keToge1xuICAgIGNhbmRpZGF0ZXM6IFsnW3JvbGU9XCJkb2N1bWVudFwiXVthcmlhLWxhYmVsPVwiTWVzc2FnZSBib2R5XCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJyZWFkaW5nUGFuZVwiLFxuICAgIHRhZzogXCJtZXNzYWdlLWJvZHlcIixcbiAgICBhbGw6IHRydWUsXG4gIH0sXG5cbiAgaW5mb0Jhcjoge1xuICAgIC8vIE5vdGljZXMgYWJvdmUgYSBtZXNzYWdlIChcIllvdSByZXBsaWVkIG9uXHUyMDI2XCIpOyB0aGVpciB0ZXh0IHNpdHMgaW4gYW4gZWxlbWVudCB3aG9zZSBpZCBzdGFydHNcbiAgICAvLyB3aXRoIFwiaW5mb0JhclwiLiBPdXRsb29rIHBhaW50cyB0aGVtIHdpdGggaW5saW5lIGNvbG91cnMuXG4gICAgY2FuZGlkYXRlczogWydbcm9sZT1cImdyb3VwXCJdOmhhcyg+ICogPiBbaWRePVwiaW5mb0JhclwiXSknXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgd2l0aGluOiBcInJlYWRpbmdQYW5lXCIsXG4gICAgdGFnOiBcImluZm8tYmFyXCIsXG4gICAgYWxsOiB0cnVlLFxuICB9LFxuXG4gIC8vIC0tLSBDbHV0dGVyIChlYWNoIGhpZGVhYmxlIGVudHJ5IGdldHMgYSB0b2dnbGUgaW4gc2V0dGluZ3MudHMpIC0tLVxuICB1cHNlbGxCYW5uZXI6IHtcbiAgICBjYW5kaWRhdGVzOiBbJ1tkYXRhLWFwcC1zZWN0aW9uPVwiVXBzZWxsQmFubmVyU2VjdGlvblwiXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB0YWc6IFwidXBzZWxsXCIsXG4gICAgYWxsOiB0cnVlLFxuICB9LFxuICBoZWFkZXJFeHRyYXM6IHtcbiAgICBjYW5kaWRhdGVzOiBbXCIjb3dhVGVhbXNCdG5fY29udGFpbmVyXCIsIFwiI293YVRpbWVQYW5lbEJ0bl9jb250YWluZXJcIl0sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHRhZzogXCJoZWFkZXItZXh0cmFcIixcbiAgICB1bmlvbjogdHJ1ZSxcbiAgfSxcbiAgYXBwUmFpbEV4dHJhczoge1xuICAgIGNhbmRpZGF0ZXM6IFtcIlRvIERvXCIsIFwiTmV3c2xldHRlcnNcIiwgXCJPcmcgRXhwbG9yZXJcIiwgXCJPbmVEcml2ZVwiXS5tYXAoXG4gICAgICAoYXBwKSA9PiBgW3JvbGU9XCJub25lXCJdW2FyaWEtbGFiZWw9XCIke2FwcH1cIl1gLFxuICAgICksXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJhcHBSYWlsXCIsXG4gICAgdGFnOiBcInJhaWwtZXh0cmFcIixcbiAgICB1bmlvbjogdHJ1ZSxcbiAgfSxcbiAgcmliYm9uRXh0cmFzOiB7XG4gICAgY2FuZGlkYXRlczogW1xuICAgICAgXCJTaGFyZSB0byBUZWFtc1wiLFxuICAgICAgXCJTd2VlcFwiLFxuICAgICAgXCJBc3NpZ24gcG9saWN5XCIsXG4gICAgICBcIlZpdmEgSW5zaWdodHNcIixcbiAgICAgIFwiTW9yZSBhcHBzXCIsXG4gICAgICBcIlF1aWNrIHN0ZXBzXCIsXG4gICAgICBcIk1vcmUgcXVpY2sgc3RlcHNcIixcbiAgICBdLm1hcCgobGFiZWwpID0+IGBbYXJpYS1sYWJlbD1cIiR7bGFiZWx9XCJdYCksXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJyaWJib25cIixcbiAgICB0YWc6IFwicmliYm9uLWV4dHJhXCIsXG4gICAgdW5pb246IHRydWUsXG4gIH0sXG4gIHJvd0F2YXRhcjoge1xuICAgIGNhbmRpZGF0ZXM6IFsnW3JvbGU9XCJpbWdcIl1baWRePVwiYXZhdGFyLVwiXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB3aXRoaW46IFwibWVzc2FnZUxpc3RcIixcbiAgICB0YWc6IFwicm93LWF2YXRhclwiLFxuICAgIGFsbDogdHJ1ZSxcbiAgfSxcbiAgcm93SG92ZXJBY3Rpb25zOiB7XG4gICAgY2FuZGlkYXRlczogW1xuICAgICAgXCJDdXN0b21pemUgQWN0aW9uc1wiLFxuICAgICAgXCJGbGFnIHRoaXMgbWVzc2FnZVwiLFxuICAgICAgXCJLZWVwIHRoaXMgbWVzc2FnZSBhdCB0aGUgdG9wIG9mIHlvdXIgZm9sZGVyXCIsXG4gICAgXS5tYXAoKGxhYmVsKSA9PiBgYnV0dG9uW2FyaWEtbGFiZWw9XCIke2xhYmVsfVwiXWApLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB3aXRoaW46IFwibWVzc2FnZUxpc3RcIixcbiAgICB0YWc6IFwicm93LWFjdGlvblwiLFxuICAgIHVuaW9uOiB0cnVlLFxuICB9LFxuICBhcHBCYW5uZXI6IHtcbiAgICAvLyBGbHVlbnQncyBzdGF0aWMgXCJmdWktTWVzc2FnZUJhclwiIGNsYXNzIGlzIHBhcnQgb2YgaXRzIHB1YmxpYyBzdHlsaW5nIEFQSSwgbm90IGdlbmVyYXRlZDtcbiAgICAvLyBub3RoaW5nIGVsc2UgbWFya3MgdGhlc2UuIFNjb3BlZCB0byBiYXJzIGFjcm9zcyB0aGUgdG9wIG9mIHRoZSBhcHAsIG5vdCBpbnNpZGUgbWVzc2FnZXMuXG4gICAgY2FuZGlkYXRlczogWycjbWFpbkFwcCA+IGRpdiA+IFtyb2xlPVwiZ3JvdXBcIl0uZnVpLU1lc3NhZ2VCYXInXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgdGFnOiBcImFwcC1iYW5uZXJcIixcbiAgICBhbGw6IHRydWUsXG4gIH0sXG4gIGNvcGlsb3Q6IHtcbiAgICAvLyBVTlZFUklGSUVEOiBubyBDb3BpbG90IHN1cmZhY2VzIGluIHRoZSBPY3QgMjAyNiBjYXB0dXJlLlxuICAgIGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWwqPVwiQ29waWxvdFwiIGldJywgJ1tyb2xlPVwicmVnaW9uXCJdW2FyaWEtbGFiZWwqPVwiQ29waWxvdFwiIGldJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHRhZzogXCJjb3BpbG90XCIsXG4gICAgdW5pb246IHRydWUsXG4gIH0sXG4gIGV4dGVybmFsVGFnOiB7XG4gICAgLy8gVU5WRVJJRklFRDogbm8gRXh0ZXJuYWwgdGFncyBpbiB0aGUgT2N0IDIwMjYgY2FwdHVyZS5cbiAgICBjYW5kaWRhdGVzOiBbJ1t0aXRsZT1cIkV4dGVybmFsIHNlbmRlclwiXScsICdbYXJpYS1sYWJlbD1cIkV4dGVybmFsXCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHRhZzogXCJleHRlcm5hbC10YWdcIixcbiAgICB1bmlvbjogdHJ1ZSxcbiAgfSxcblxuICAvLyAtLS0gQ29udHJvbHMgdGhlIGtleW1hcCBkcml2ZXMgLS0tXG4gIGNvbXBvc2VCdXR0b246IHsgY2FuZGlkYXRlczogWydidXR0b25bYXJpYS1sYWJlbD1cIk5ldyBtYWlsXCJdJ10sIGV4cGVjdDogXCJhbHdheXNcIiB9LFxuICBhcmNoaXZlQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJBcmNoaXZlXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICBkZWxldGVCdXR0b246IHsgY2FuZGlkYXRlczogWydidXR0b25bYXJpYS1sYWJlbD1cIkRlbGV0ZVwiXSddLCBleHBlY3Q6IFwic29tZXRpbWVzXCIgfSxcbiAgLy8gVU5WRVJJRklFRDogdGhpcyB0ZW5hbnQgc2hvd3MgYSBcIlBoaXNoIEFsZXJ0XCIgYWRkLWluIGluc3RlYWQgb2YgT3V0bG9vaydzIFJlcG9ydCBtZW51LlxuICBqdW5rQnV0dG9uOiB7XG4gICAgY2FuZGlkYXRlczogWydidXR0b25bYXJpYS1sYWJlbD1cIlJlcG9ydFwiXScsICdidXR0b25bYXJpYS1sYWJlbD1cIkp1bmtcIl0nXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gIH0sXG4gIG1vdmVCdXR0b246IHsgY2FuZGlkYXRlczogWydidXR0b25bYXJpYS1sYWJlbD1cIk1vdmUgdG9cIl0nXSwgZXhwZWN0OiBcInNvbWV0aW1lc1wiIH0sXG4gIHNub296ZUJ1dHRvbjogeyBjYW5kaWRhdGVzOiBbJ2J1dHRvblthcmlhLWxhYmVsPVwiU25vb3plXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICBmbGFnQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJGbGFnIC8gVW5mbGFnXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICByZWFkVW5yZWFkQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJSZWFkIC8gVW5yZWFkXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICB1bmRvQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJVbmRvXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICBzeW5jQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFtcIiNvd2FTeW5jQnV0dG9uX2NvbnRhaW5lclwiXSwgZXhwZWN0OiBcInNvbWV0aW1lc1wiIH0sXG4gIHNldHRpbmdzQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFtcIiNvd2FTZXR0aW5nc0J0bl9jb250YWluZXJcIl0sIGV4cGVjdDogXCJhbHdheXNcIiB9LFxuICAvLyBBcHAgcmFpbCBidXR0b25zLiBGYWxsYmFja3MgYXJlIHRoZSBNaWNyb3NvZnQgMzY1IGFwcCBpZHMgT3V0bG9vayBwdXRzIGluIGRhdGEtdGlkLCB0aGUgc2FtZVxuICAvLyBhY3Jvc3MgdGhlIE9jdCBhbmQgTm92IDIwMjYgY2FwdHVyZXMuIFRoZSByYWlsIHJlbmRlcnMgYWZ0ZXIgdGhlIG1lc3NhZ2UgbGlzdC5cbiAgbWFpbEFwcDoge1xuICAgIGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJNYWlsXCJdJywgJ1tkYXRhLXRpZD1cImRkZWE3NzRjLTM4MmItNDdkNy1hYWI1LWFkYzIxMzlhODAyYlwiXSddLFxuICAgIHdpdGhpbjogXCJhcHBSYWlsXCIsXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICB9LFxuICBjYWxlbmRhckFwcDoge1xuICAgIGNhbmRpZGF0ZXM6IFtcbiAgICAgICdidXR0b25bYXJpYS1sYWJlbD1cIkNhbGVuZGFyXCJdJyxcbiAgICAgICdbZGF0YS10aWQ9XCI4Y2JlYjg2Zi04M2UxLTQzYjUtYWFiYS1jZDM1MTQzMjJmMGJcIl0nLFxuICAgIF0sXG4gICAgd2l0aGluOiBcImFwcFJhaWxcIixcbiAgICBleHBlY3Q6IFwiYWx3YXlzXCIsXG4gIH0sXG4gIHBlb3BsZUFwcDoge1xuICAgIGNhbmRpZGF0ZXM6IFtcbiAgICAgICdidXR0b25bYXJpYS1sYWJlbD1cIlBlb3BsZVwiXScsXG4gICAgICAnW2RhdGEtdGlkPVwiMzU1ZmJkNzktM2JhMi00NTU0LThmMmQtMDMwMGZkZTkxZjMwXCJdJyxcbiAgICBdLFxuICAgIHdpdGhpbjogXCJhcHBSYWlsXCIsXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICB9LFxufSBhcyBjb25zdCBzYXRpc2ZpZXMgUmVjb3JkPHN0cmluZywgU2VsZWN0b3JFbnRyeT47XG5cbmNvbnN0IGVudHJ5ID0gKG5hbWU6IFNlbGVjdG9yTmFtZSk6IFNlbGVjdG9yRW50cnkgPT4gU0VMRUNUT1JTW25hbWVdO1xuXG5mdW5jdGlvbiBzY29wZShuYW1lOiBTZWxlY3Rvck5hbWUsIHJvb3Q6IFBhcmVudE5vZGUpOiBQYXJlbnROb2RlIHwgbnVsbCB7XG4gIGNvbnN0IHBhcmVudCA9IGVudHJ5KG5hbWUpLndpdGhpbjtcbiAgcmV0dXJuIHBhcmVudCA/IHJlc29sdmUocGFyZW50IGFzIFNlbGVjdG9yTmFtZSwgcm9vdCkgOiByb290O1xufVxuXG4vKiogVGhlIGZpcnN0IGVsZW1lbnQgbWF0Y2hpbmcgYG5hbWVgLCB0cnlpbmcgY2FuZGlkYXRlcyBpbiBwcmlvcml0eSBvcmRlci4gKi9cbmV4cG9ydCBmdW5jdGlvbiByZXNvbHZlKG5hbWU6IFNlbGVjdG9yTmFtZSwgcm9vdDogUGFyZW50Tm9kZSA9IGRvY3VtZW50KTogRWxlbWVudCB8IG51bGwge1xuICBjb25zdCB3aXRoaW4gPSBzY29wZShuYW1lLCByb290KTtcbiAgaWYgKCF3aXRoaW4pIHJldHVybiBudWxsO1xuICBmb3IgKGNvbnN0IGNhbmRpZGF0ZSBvZiBlbnRyeShuYW1lKS5jYW5kaWRhdGVzKSB7XG4gICAgY29uc3QgZWwgPSB3aXRoaW4ucXVlcnlTZWxlY3RvcihjYW5kaWRhdGUpO1xuICAgIGlmIChlbCkgcmV0dXJuIGVsO1xuICB9XG4gIHJldHVybiBudWxsO1xufVxuXG4vKiogRXZlcnkgbWF0Y2g6IGZvciB1bmlvbiBlbnRyaWVzLCBvZiBhbGwgY2FuZGlkYXRlczsgb3RoZXJ3aXNlIG9mIHRoZSBmaXJzdCBjYW5kaWRhdGUgdGhhdCBtYXRjaGVzLiAqL1xuZXhwb3J0IGZ1bmN0aW9uIHJlc29sdmVBbGwobmFtZTogU2VsZWN0b3JOYW1lLCByb290OiBQYXJlbnROb2RlID0gZG9jdW1lbnQpOiBFbGVtZW50W10ge1xuICBjb25zdCB3aXRoaW4gPSBzY29wZShuYW1lLCByb290KTtcbiAgaWYgKCF3aXRoaW4pIHJldHVybiBbXTtcbiAgY29uc3QgeyBjYW5kaWRhdGVzLCB1bmlvbiB9ID0gZW50cnkobmFtZSk7XG4gIGlmICh1bmlvbikgcmV0dXJuIFsuLi53aXRoaW4ucXVlcnlTZWxlY3RvckFsbChjYW5kaWRhdGVzLmpvaW4oXCIsIFwiKSldO1xuICBmb3IgKGNvbnN0IGNhbmRpZGF0ZSBvZiBjYW5kaWRhdGVzKSB7XG4gICAgY29uc3QgZWxzID0gd2l0aGluLnF1ZXJ5U2VsZWN0b3JBbGwoY2FuZGlkYXRlKTtcbiAgICBpZiAoZWxzLmxlbmd0aCA+IDApIHJldHVybiBbLi4uZWxzXTtcbiAgfVxuICByZXR1cm4gW107XG59XG5cbi8qKiBBZGQgYGRhdGEtb3Zlcmxvb2tgIHRva2VucyB0byBldmVyeSB0YWdnZWQgZW50cnkgc28gc3R5bGVzaGVldHMgbmV2ZXIgdG91Y2ggT3V0bG9vaydzIERPTSBzaGFwZS4gKi9cbmV4cG9ydCBmdW5jdGlvbiB0YWdFbGVtZW50cyhyb290OiBQYXJlbnROb2RlID0gZG9jdW1lbnQpOiB2b2lkIHtcbiAgZm9yIChjb25zdCBuYW1lIG9mIE9iamVjdC5rZXlzKFNFTEVDVE9SUykgYXMgU2VsZWN0b3JOYW1lW10pIHtcbiAgICBjb25zdCB7IHRhZywgYWxsLCB1bmlvbiB9ID0gZW50cnkobmFtZSk7XG4gICAgaWYgKCF0YWcpIGNvbnRpbnVlO1xuICAgIGNvbnN0IGVscyA9XG4gICAgICBhbGwgfHwgdW5pb24gPyByZXNvbHZlQWxsKG5hbWUsIHJvb3QpIDogW3Jlc29sdmUobmFtZSwgcm9vdCldLmZpbHRlcigoZSkgPT4gZSAhPT0gbnVsbCk7XG4gICAgZm9yIChjb25zdCBlbCBvZiBlbHMpIHtcbiAgICAgIGNvbnN0IHRva2VucyA9IG5ldyBTZXQoKGVsLmdldEF0dHJpYnV0ZShcImRhdGEtb3Zlcmxvb2tcIikgPz8gXCJcIikuc3BsaXQoXCIgXCIpLmZpbHRlcihCb29sZWFuKSk7XG4gICAgICBpZiAodG9rZW5zLmhhcyh0YWcpKSBjb250aW51ZTtcbiAgICAgIHRva2Vucy5hZGQodGFnKTtcbiAgICAgIGVsLnNldEF0dHJpYnV0ZShcImRhdGEtb3Zlcmxvb2tcIiwgWy4uLnRva2Vuc10uam9pbihcIiBcIikpO1xuICAgIH1cbiAgfVxufVxuXG5leHBvcnQgaW50ZXJmYWNlIEhlYWx0aFJlc3VsdCB7XG4gIG5hbWU6IFNlbGVjdG9yTmFtZTtcbiAgZXhwZWN0OiBTZWxlY3RvckVudHJ5W1wiZXhwZWN0XCJdO1xuICAvKiogSW5kZXggb2YgdGhlIGNhbmRpZGF0ZSB0aGF0IG1hdGNoZWQsIG9yIC0xLiBBIG5vbi16ZXJvIGluZGV4IG1lYW5zIGEgZmFsbGJhY2sgaXMgaW4gdXNlLiAqL1xuICBjYW5kaWRhdGU6IG51bWJlcjtcbiAgY291bnQ6IG51bWJlcjtcbn1cblxuZXhwb3J0IGZ1bmN0aW9uIGhlYWx0aENoZWNrKHJvb3Q6IFBhcmVudE5vZGUgPSBkb2N1bWVudCk6IEhlYWx0aFJlc3VsdFtdIHtcbiAgcmV0dXJuIChPYmplY3Qua2V5cyhTRUxFQ1RPUlMpIGFzIFNlbGVjdG9yTmFtZVtdKS5tYXAoKG5hbWUpID0+IHtcbiAgICBjb25zdCB7IGNhbmRpZGF0ZXMsIGV4cGVjdCwgdW5pb24gfSA9IGVudHJ5KG5hbWUpO1xuICAgIGNvbnN0IHdpdGhpbiA9IHNjb3BlKG5hbWUsIHJvb3QpO1xuICAgIGxldCBjYW5kaWRhdGUgPSAtMTtcbiAgICBsZXQgY291bnQgPSAwO1xuICAgIGlmICh3aXRoaW4pIHtcbiAgICAgIGNhbmRpZGF0ZSA9IGNhbmRpZGF0ZXMuZmluZEluZGV4KChjKSA9PiB3aXRoaW4ucXVlcnlTZWxlY3RvcihjKSAhPT0gbnVsbCk7XG4gICAgICBpZiAodW5pb24gJiYgY2FuZGlkYXRlID49IDApIGNvdW50ID0gcmVzb2x2ZUFsbChuYW1lLCByb290KS5sZW5ndGg7XG4gICAgICBlbHNlIGlmIChjYW5kaWRhdGUgPj0gMCkgY291bnQgPSB3aXRoaW4ucXVlcnlTZWxlY3RvckFsbChjYW5kaWRhdGVzW2NhbmRpZGF0ZV0gPz8gXCJcIikubGVuZ3RoO1xuICAgIH1cbiAgICAvLyBBIHVuaW9uIGVudHJ5IGxpc3RzIGRpc3RpbmN0IHRoaW5ncywgc28gbWF0Y2hpbmcgYSBsYXRlciBvbmUgaXMgbm90IGEgZmFsbGJhY2suXG4gICAgaWYgKHVuaW9uICYmIGNhbmRpZGF0ZSA+IDApIGNhbmRpZGF0ZSA9IDA7XG4gICAgcmV0dXJuIHsgbmFtZSwgZXhwZWN0LCBjYW5kaWRhdGUsIGNvdW50IH07XG4gIH0pO1xufVxuXG5leHBvcnQgaW50ZXJmYWNlIEZvbGRlck5vZGUge1xuICBuYW1lOiBzdHJpbmc7XG4gIC8qKiBBbmNlc3RvciBuYW1lcyBqb2luZWQgd2l0aCBcIi9cIiwgZS5nLiBcIkluYm94L1JlY2VpcHRzXCIuICovXG4gIHBhdGg6IHN0cmluZztcbiAgZWxlbWVudDogRWxlbWVudDtcbn1cblxuLyoqIFwiSW5ib3ggLSA5IGl0ZW1zICgxIHVucmVhZClcIiAtPiBcIkluYm94XCIuIFRoZSB0aXRsZSBhdHRyaWJ1dGUgY2FycmllcyBjb3VudHMuICovXG5mdW5jdGlvbiBmb2xkZXJOYW1lKGVsZW1lbnQ6IEVsZW1lbnQpOiBzdHJpbmcge1xuICBjb25zdCB0aXRsZSA9IChlbGVtZW50LmdldEF0dHJpYnV0ZShcInRpdGxlXCIpID8/IGVsZW1lbnQudGV4dENvbnRlbnQgPz8gXCJcIikudHJpbSgpO1xuICByZXR1cm4gdGl0bGUucmVwbGFjZSgvXFxzKy1cXHMrXFxkK1xccytpdGVtcz8oXFxzK1xcKFxcZCtcXHMrdW5yZWFkXFwpKT8kLywgXCJcIik7XG59XG5cbi8qKlxuICogRm9sZGVycyBjdXJyZW50bHkgcmVuZGVyZWQgaW4gdGhlIGZvbGRlciBwYW5lLCB3aXRoIHBhdGhzIGJ1aWx0IGZyb20gYXJpYS1sZXZlbC4gRm9sZGVyc1xuICogdW5kZXIgRmF2b3JpdGVzIGFyZSBza2lwcGVkICh0aGV5IGR1cGxpY2F0ZSB0aGUgbWFpbGJveCB0cmVlKSwgYXMgYXJlIHRoZSB0b3AtbGV2ZWwgbWFpbGJveFxuICogcm9vdHMsIHdob3NlIG5hbWVzIGFyZSBhY2NvdW50IGFkZHJlc3Nlcy5cbiAqIFVOVkVSSUZJRUQ6IGNvbGxhcHNlZCBzdWJmb2xkZXJzIG1heSBiZSBhYnNlbnQgZnJvbSB0aGUgRE9NLlxuICovXG5leHBvcnQgZnVuY3Rpb24gcmVhZEZvbGRlclRyZWUocm9vdDogUGFyZW50Tm9kZSA9IGRvY3VtZW50KTogRm9sZGVyTm9kZVtdIHtcbiAgY29uc3Qgc3RhY2s6IHN0cmluZ1tdID0gW107XG4gIGNvbnN0IGZvbGRlcnM6IEZvbGRlck5vZGVbXSA9IFtdO1xuICBsZXQgaW5GYXZvcml0ZXMgPSBmYWxzZTtcbiAgZm9yIChjb25zdCBlbGVtZW50IG9mIHJlc29sdmVBbGwoXCJmb2xkZXJJdGVtXCIsIHJvb3QpKSB7XG4gICAgY29uc3QgbGV2ZWwgPSBOdW1iZXIoZWxlbWVudC5nZXRBdHRyaWJ1dGUoXCJhcmlhLWxldmVsXCIpID8/IFwiMVwiKTtcbiAgICBjb25zdCBuYW1lID0gZm9sZGVyTmFtZShlbGVtZW50KTtcbiAgICBzdGFjay5sZW5ndGggPSBNYXRoLm1heChsZXZlbCAtIDEsIDApO1xuICAgIHN0YWNrLnB1c2gobmFtZSk7XG4gICAgaWYgKGxldmVsID09PSAxKSB7XG4gICAgICBpbkZhdm9yaXRlcyA9IGVsZW1lbnQuZ2V0QXR0cmlidXRlKFwiZGF0YS1mb2xkZXItbmFtZVwiKSA9PT0gXCJmYXZvcml0ZXNcIjtcbiAgICAgIGNvbnRpbnVlO1xuICAgIH1cbiAgICBpZiAoaW5GYXZvcml0ZXMpIGNvbnRpbnVlO1xuICAgIGZvbGRlcnMucHVzaCh7IG5hbWUsIHBhdGg6IHN0YWNrLnNsaWNlKDEpLmZpbHRlcihCb29sZWFuKS5qb2luKFwiL1wiKSwgZWxlbWVudCB9KTtcbiAgfVxuICByZXR1cm4gZm9sZGVycztcbn1cblxuLyoqIE1lc3NhZ2Ugcm93cyBjdXJyZW50bHkgc2VsZWN0ZWQgKG9yIG9wZW4pIGluIHRoZSBsaXN0LiAqL1xuZXhwb3J0IGZ1bmN0aW9uIHNlbGVjdGVkUm93cyhyb290OiBQYXJlbnROb2RlID0gZG9jdW1lbnQpOiBFbGVtZW50W10ge1xuICByZXR1cm4gcmVzb2x2ZUFsbChcIm1lc3NhZ2VSb3dcIiwgcm9vdCkuZmlsdGVyKFxuICAgIChyb3cpID0+IHJvdy5nZXRBdHRyaWJ1dGUoXCJhcmlhLXNlbGVjdGVkXCIpID09PSBcInRydWVcIixcbiAgKTtcbn1cblxuLyoqIE91dGxvb2sgc3RhcnRzIGFuIHVucmVhZCByb3cncyBhY2Nlc3NpYmxlIG5hbWUgd2l0aCBcIlVucmVhZFwiIChFbmdsaXNoIFVJKS4gKi9cbmV4cG9ydCBmdW5jdGlvbiBpc1VucmVhZFJvdyhyb3c6IEVsZW1lbnQpOiBib29sZWFuIHtcbiAgcmV0dXJuIC9eVW5yZWFkXFxiLy50ZXN0KHJvdy5nZXRBdHRyaWJ1dGUoXCJhcmlhLWxhYmVsXCIpID8/IFwiXCIpO1xufVxuXG4vKiogT3V0bG9vaydzIGN1cnJlbnQgbGlnaHQvZGFyayBtb2RlLCBmcm9tIHRoZSBwYWxldHRlIHZhcmlhYmxlcyBpdCB3cml0ZXMgb24gPGh0bWw+LiAqL1xuZXhwb3J0IGZ1bmN0aW9uIG91dGxvb2tDb2xvclNjaGVtZShodG1sOiBFbGVtZW50KTogXCJsaWdodFwiIHwgXCJkYXJrXCIge1xuICBjb25zdCBzdHlsZSA9IGh0bWwuZ2V0QXR0cmlidXRlKFwic3R5bGVcIikgPz8gXCJcIjtcbiAgcmV0dXJuIC8tLW5hdGl2ZS1zY3JvbGxiYXItY29sb3Itc2NoZW1lOlxccypkYXJrLy50ZXN0KHN0eWxlKSA/IFwiZGFya1wiIDogXCJsaWdodFwiO1xufVxuXG4vKiogRW50cmllcyBleHBlY3RlZCBvbiBldmVyeSBtYWlsIHZpZXcgdGhhdCBkaWQgbm90IHJlc29sdmUuIE5vbi1lbXB0eSBtZWFucyBzb21ldGhpbmcgYnJva2UuICovXG5leHBvcnQgZnVuY3Rpb24gYnJva2VuRW50cmllcyhyZXN1bHRzOiBIZWFsdGhSZXN1bHRbXSk6IFNlbGVjdG9yTmFtZVtdIHtcbiAgcmV0dXJuIHJlc3VsdHMuZmlsdGVyKChyKSA9PiByLmV4cGVjdCA9PT0gXCJhbHdheXNcIiAmJiByLmNhbmRpZGF0ZSA8IDApLm1hcCgocikgPT4gci5uYW1lKTtcbn1cbiIsICI6aG9zdCB7XG4gIGFsbDogaW5pdGlhbDtcbiAgLS1vbC1iZzogI2ZmZjtcbiAgLS1vbC1mZzogIzFmMWYxZjtcbiAgLS1vbC1tdXRlZDogIzZiNmI2YjtcbiAgLS1vbC1hY2NlbnQ6ICMyZjZmZGU7XG4gIC0tb2wtYm9yZGVyOiByZ2IoMCAwIDAgLyAxMiUpO1xuICAtLW9sLXNoYWRvdzogMCA4cHggMzJweCByZ2IoMCAwIDAgLyAxOCUpO1xuICBmb250OlxuICAgIDEzcHggLyAxLjQgc3lzdGVtLXVpLFxuICAgIC1hcHBsZS1zeXN0ZW0sXG4gICAgXCJTZWdvZSBVSVwiLFxuICAgIHNhbnMtc2VyaWY7XG59XG5cbkBtZWRpYSAocHJlZmVycy1jb2xvci1zY2hlbWU6IGRhcmspIHtcbiAgOmhvc3Qge1xuICAgIC0tb2wtYmc6ICMyNjI2MjY7XG4gICAgLS1vbC1mZzogI2VjZWNlYztcbiAgICAtLW9sLW11dGVkOiAjYTBhMGEwO1xuICAgIC0tb2wtYWNjZW50OiAjN2FhN2ZmO1xuICAgIC0tb2wtYm9yZGVyOiByZ2IoMjU1IDI1NSAyNTUgLyAxNCUpO1xuICB9XG59XG5cbltoaWRkZW5dIHtcbiAgZGlzcGxheTogbm9uZSAhaW1wb3J0YW50O1xufVxuXG4udG9hc3QsXG4uaGludCxcbi5oZWFsdGgge1xuICBwb3NpdGlvbjogZml4ZWQ7XG4gIHotaW5kZXg6IDIxNDc0ODM2NDc7XG4gIGJhY2tncm91bmQ6IHZhcigtLW9sLWJnKTtcbiAgY29sb3I6IHZhcigtLW9sLWZnKTtcbiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tb2wtYm9yZGVyKTtcbiAgYm9yZGVyLXJhZGl1czogNnB4O1xuICBib3gtc2hhZG93OiB2YXIoLS1vbC1zaGFkb3cpO1xuICBwYWRkaW5nOiA2cHggMTBweDtcbn1cblxuLnRvYXN0IHtcbiAgYm90dG9tOiAyNHB4O1xuICBsZWZ0OiA1MCU7XG4gIHRyYW5zZm9ybTogdHJhbnNsYXRlWCgtNTAlKTtcbn1cblxuLmhpbnQge1xuICBib3R0b206IDI0cHg7XG4gIHJpZ2h0OiAyNHB4O1xuICBmb250LWZhbWlseTogdWktbW9ub3NwYWNlLCBcIlNGIE1vbm9cIiwgTWVubG8sIG1vbm9zcGFjZTtcbn1cblxuLmhlYWx0aCB7XG4gIHRvcDogOHB4O1xuICByaWdodDogOHB4O1xuICBkaXNwbGF5OiBmbGV4O1xuICBnYXA6IDhweDtcbiAgYWxpZ24taXRlbXM6IGNlbnRlcjtcbiAgYm9yZGVyLWNvbG9yOiAjZDk3NzA2O1xufVxuXG4uY2xvc2Uge1xuICBhbGw6IHVuc2V0O1xuICBjdXJzb3I6IHBvaW50ZXI7XG4gIGNvbG9yOiB2YXIoLS1vbC1tdXRlZCk7XG4gIHBhZGRpbmc6IDAgMnB4O1xufVxuXG4uYmFja2Ryb3Age1xuICBwb3NpdGlvbjogZml4ZWQ7XG4gIGluc2V0OiAwO1xuICB6LWluZGV4OiAyMTQ3NDgzNjQ2O1xuICBiYWNrZ3JvdW5kOiByZ2IoMCAwIDAgLyAyNSUpO1xuICBkaXNwbGF5OiBmbGV4O1xuICBqdXN0aWZ5LWNvbnRlbnQ6IGNlbnRlcjtcbiAgYWxpZ24taXRlbXM6IGZsZXgtc3RhcnQ7XG4gIHBhZGRpbmctdG9wOiAxMnZoO1xufVxuXG4ucGFuZWwge1xuICBiYWNrZ3JvdW5kOiB2YXIoLS1vbC1iZyk7XG4gIGNvbG9yOiB2YXIoLS1vbC1mZyk7XG4gIGJvcmRlci1yYWRpdXM6IDEwcHg7XG4gIGJveC1zaGFkb3c6IHZhcigtLW9sLXNoYWRvdyk7XG4gIHdpZHRoOiBtaW4oNTIwcHgsIDkwdncpO1xuICBtYXgtaGVpZ2h0OiA3MHZoO1xuICBvdmVyZmxvdzogYXV0bztcbn1cblxuLmhlbHAge1xuICBwYWRkaW5nOiAxNnB4IDIwcHg7XG59XG5cbi5oZWxwIGgyIHtcbiAgZm9udC1zaXplOiAxNXB4O1xuICBtYXJnaW46IDAgMCAxMnB4O1xufVxuXG4uaGVscCB0YWJsZSB7XG4gIGJvcmRlci1jb2xsYXBzZTogY29sbGFwc2U7XG4gIHdpZHRoOiAxMDAlO1xufVxuXG4uaGVscCB0ZCB7XG4gIHBhZGRpbmc6IDNweCAwO1xufVxuXG4ua2V5cyB7XG4gIHdpZHRoOiA3ZW07XG59XG5cbmtiZCB7XG4gIGZvbnQtZmFtaWx5OiB1aS1tb25vc3BhY2UsIFwiU0YgTW9ub1wiLCBNZW5sbywgbW9ub3NwYWNlO1xuICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1vbC1ib3JkZXIpO1xuICBib3JkZXItcmFkaXVzOiA0cHg7XG4gIHBhZGRpbmc6IDAgNXB4O1xufVxuXG4ucGFsZXR0ZSBpbnB1dCB7XG4gIGFsbDogdW5zZXQ7XG4gIGJveC1zaXppbmc6IGJvcmRlci1ib3g7XG4gIHdpZHRoOiAxMDAlO1xuICBwYWRkaW5nOiAxMnB4IDE2cHg7XG4gIGZvbnQtc2l6ZTogMTVweDtcbiAgYm9yZGVyLWJvdHRvbTogMXB4IHNvbGlkIHZhcigtLW9sLWJvcmRlcik7XG59XG5cbi5wYWxldHRlIHVsIHtcbiAgbGlzdC1zdHlsZTogbm9uZTtcbiAgbWFyZ2luOiAwO1xuICBwYWRkaW5nOiA0cHggMDtcbn1cblxuLnBhbGV0dGUgbGkge1xuICBwYWRkaW5nOiA1cHggMTZweDtcbiAgY3Vyc29yOiBwb2ludGVyO1xufVxuXG4ucGFsZXR0ZSBsaVthcmlhLXNlbGVjdGVkPVwidHJ1ZVwiXSB7XG4gIGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1vbC1hY2NlbnQpIDE2JSwgdHJhbnNwYXJlbnQpO1xufVxuXG4ucGFsZXR0ZSBsaSAucGF0aCB7XG4gIGNvbG9yOiB2YXIoLS1vbC1tdXRlZCk7XG4gIG1hcmdpbi1sZWZ0OiA2cHg7XG59XG5cbi5wYWxldHRlIG1hcmsge1xuICBiYWNrZ3JvdW5kOiBub25lO1xuICBjb2xvcjogdmFyKC0tb2wtYWNjZW50KTtcbiAgZm9udC13ZWlnaHQ6IDYwMDtcbn1cblxuLnBhbGV0dGUgLmVtcHR5IHtcbiAgY29sb3I6IHZhcigtLW9sLW11dGVkKTtcbiAgY3Vyc29yOiBkZWZhdWx0O1xufVxuIiwgIi8vIE9CQidzIG93biBVSSBsaXZlcyBpbiBvbmUgc2hhZG93IHJvb3QsIHNvIE91dGxvb2sncyBDU1MgY2Fubm90IHJlYWNoIGl0IGFuZCBvdXJzIGNhbm5vdFxuLy8gbGVhayBpbnRvIE91dGxvb2suIEV2ZXJ5dGhpbmcgaGVyZSBpcyBwbGFpbiBET007IG5vIGZyYW1ld29yay5cblxuaW1wb3J0IGNzcyBmcm9tIFwiLi9vdmVybGF5LmNzc1wiO1xuXG5sZXQgc2hhZG93OiBTaGFkb3dSb290IHwgbnVsbCA9IG51bGw7XG5cbmV4cG9ydCBmdW5jdGlvbiBvdmVybGF5Um9vdCgpOiBTaGFkb3dSb290IHtcbiAgaWYgKHNoYWRvdykgcmV0dXJuIHNoYWRvdztcbiAgY29uc3QgaG9zdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoXCJvdmVybG9vay1yb290XCIpO1xuICBkb2N1bWVudC5kb2N1bWVudEVsZW1lbnQuYXBwZW5kKGhvc3QpO1xuICBzaGFkb3cgPSBob3N0LmF0dGFjaFNoYWRvdyh7IG1vZGU6IFwib3BlblwiIH0pO1xuICAvLyBLZXlzdHJva2VzIHR5cGVkIGludG8gb3VyIFVJIG11c3Qgbm90IHJlYWNoIE91dGxvb2sncyBkb2N1bWVudC1sZXZlbCBzaG9ydGN1dCBoYW5kbGVycyxcbiAgLy8gd2hpY2ggd291bGQgc2VlIHRoZSBob3N0IGVsZW1lbnQgKG5vdCBhbiBpbnB1dCkgYXMgdGhlIHRhcmdldC5cbiAgZm9yIChjb25zdCB0eXBlIG9mIFtcImtleWRvd25cIiwgXCJrZXlwcmVzc1wiLCBcImtleXVwXCJdKSB7XG4gICAgaG9zdC5hZGRFdmVudExpc3RlbmVyKHR5cGUsIChlKSA9PiBlLnN0b3BQcm9wYWdhdGlvbigpKTtcbiAgfVxuICBjb25zdCBzdHlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoXCJzdHlsZVwiKTtcbiAgc3R5bGUudGV4dENvbnRlbnQgPSBjc3M7XG4gIHNoYWRvdy5hcHBlbmQoc3R5bGUpO1xuICByZXR1cm4gc2hhZG93O1xufVxuXG4vKiogVHJ1ZSB3aGVuIHRoZSBldmVudCBjYW1lIGZyb20gaW5zaWRlIE9CQidzIG93biBVSS4gKi9cbmV4cG9ydCBmdW5jdGlvbiBpc0Zyb21PdmVybGF5KGV2ZW50OiBFdmVudCk6IGJvb2xlYW4ge1xuICByZXR1cm4gc2hhZG93ICE9PSBudWxsICYmIGV2ZW50LmNvbXBvc2VkUGF0aCgpLmluY2x1ZGVzKHNoYWRvdy5ob3N0KTtcbn1cblxuLyoqIFRydWUgd2hpbGUgYSBtb2RhbCAocGFsZXR0ZSBvciBoZWxwKSBpcyBvcGVuOyBpdCBvd25zIHRoZSBrZXlib2FyZCB1bnRpbCBjbG9zZWQuICovXG5leHBvcnQgZnVuY3Rpb24gaXNNb2RhbE9wZW4oKTogYm9vbGVhbiB7XG4gIHJldHVybiBzaGFkb3c/LnF1ZXJ5U2VsZWN0b3IoXCIuYmFja2Ryb3BcIikgIT0gbnVsbDtcbn1cblxuZnVuY3Rpb24gZWw8SyBleHRlbmRzIGtleW9mIEhUTUxFbGVtZW50VGFnTmFtZU1hcD4odGFnOiBLLCBjbGFzc05hbWU6IHN0cmluZywgdGV4dD86IHN0cmluZykge1xuICBjb25zdCBub2RlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCh0YWcpO1xuICBub2RlLmNsYXNzTmFtZSA9IGNsYXNzTmFtZTtcbiAgaWYgKHRleHQgIT09IHVuZGVmaW5lZCkgbm9kZS50ZXh0Q29udGVudCA9IHRleHQ7XG4gIHJldHVybiBub2RlO1xufVxuXG5mdW5jdGlvbiBzaW5nbGV0b24oY2xhc3NOYW1lOiBzdHJpbmcpOiBIVE1MRGl2RWxlbWVudCB7XG4gIGNvbnN0IHJvb3QgPSBvdmVybGF5Um9vdCgpO1xuICBjb25zdCBleGlzdGluZyA9IHJvb3QucXVlcnlTZWxlY3RvcjxIVE1MRGl2RWxlbWVudD4oYC4ke2NsYXNzTmFtZX1gKTtcbiAgaWYgKGV4aXN0aW5nKSByZXR1cm4gZXhpc3Rpbmc7XG4gIGNvbnN0IG5vZGUgPSBlbChcImRpdlwiLCBjbGFzc05hbWUpO1xuICByb290LmFwcGVuZChub2RlKTtcbiAgcmV0dXJuIG5vZGU7XG59XG5cbmxldCB0b2FzdFRpbWVyOiBSZXR1cm5UeXBlPHR5cGVvZiBzZXRUaW1lb3V0PiB8IHVuZGVmaW5lZDtcblxuZXhwb3J0IGZ1bmN0aW9uIHRvYXN0KG1lc3NhZ2U6IHN0cmluZywgbXMgPSAyNTAwKTogdm9pZCB7XG4gIGNvbnN0IG5vZGUgPSBzaW5nbGV0b24oXCJ0b2FzdFwiKTtcbiAgbm9kZS50ZXh0Q29udGVudCA9IG1lc3NhZ2U7XG4gIG5vZGUuaGlkZGVuID0gZmFsc2U7XG4gIGNsZWFyVGltZW91dCh0b2FzdFRpbWVyKTtcbiAgdG9hc3RUaW1lciA9IHNldFRpbWVvdXQoKCkgPT4ge1xuICAgIG5vZGUuaGlkZGVuID0gdHJ1ZTtcbiAgfSwgbXMpO1xufVxuXG4vKiogVGhlIHBlbmRpbmctcHJlZml4IGhpbnQsIGUuZy4gXCJHIFx1MjAyNiBtIE1haWwgXHUwMEI3IGMgQ2FsZW5kYXJcIi4gUGFzcyBudWxsIHRvIGhpZGUuICovXG5leHBvcnQgZnVuY3Rpb24gc2hvd0hpbnQodGV4dDogc3RyaW5nIHwgbnVsbCk6IHZvaWQge1xuICBjb25zdCBub2RlID0gc2luZ2xldG9uKFwiaGludFwiKTtcbiAgbm9kZS5oaWRkZW4gPSB0ZXh0ID09PSBudWxsO1xuICBub2RlLnRleHRDb250ZW50ID0gdGV4dCA/PyBcIlwiO1xufVxuXG5leHBvcnQgZnVuY3Rpb24gc2hvd0hlYWx0aEJhZGdlKGJyb2tlbjogcmVhZG9ubHkgc3RyaW5nW10sIG9uRGlzbWlzczogKCkgPT4gdm9pZCk6IHZvaWQge1xuICBjb25zdCBub2RlID0gc2luZ2xldG9uKFwiaGVhbHRoXCIpO1xuICBub2RlLnJlcGxhY2VDaGlsZHJlbigpO1xuICBpZiAoYnJva2VuLmxlbmd0aCA9PT0gMCkge1xuICAgIG5vZGUuaGlkZGVuID0gdHJ1ZTtcbiAgICByZXR1cm47XG4gIH1cbiAgbm9kZS5oaWRkZW4gPSBmYWxzZTtcbiAgbm9kZS50aXRsZSA9IGBOb3QgZm91bmQ6ICR7YnJva2VuLmpvaW4oXCIsIFwiKX0uIFNlZSB0aGUgY29uc29sZSBmb3IgZGV0YWlscy5gO1xuICBjb25zdCBjbG9zZSA9IGVsKFwiYnV0dG9uXCIsIFwiY2xvc2VcIiwgXCJcdTAwRDdcIik7XG4gIGNsb3NlLnNldEF0dHJpYnV0ZShcImFyaWEtbGFiZWxcIiwgXCJEaXNtaXNzXCIpO1xuICBjbG9zZS5hZGRFdmVudExpc3RlbmVyKFwiY2xpY2tcIiwgKCkgPT4ge1xuICAgIG5vZGUuaGlkZGVuID0gdHJ1ZTtcbiAgICBvbkRpc21pc3MoKTtcbiAgfSk7XG4gIG5vZGUuYXBwZW5kKGVsKFwic3BhblwiLCBcIlwiLCBgT0JCOiAke2Jyb2tlbi5sZW5ndGh9IHNlbGVjdG9yKHMpIGJyb2tlbmApLCBjbG9zZSk7XG59XG5cbmV4cG9ydCBmdW5jdGlvbiBzaG93SGVscChyb3dzOiByZWFkb25seSB7IGtleXM6IHN0cmluZzsgZGVzY3JpcHRpb246IHN0cmluZyB9W10pOiB2b2lkIHtcbiAgY29uc3Qgcm9vdCA9IG92ZXJsYXlSb290KCk7XG4gIGNvbnN0IGJhY2tkcm9wID0gZWwoXCJkaXZcIiwgXCJiYWNrZHJvcFwiKTtcbiAgY29uc3QgcGFuZWwgPSBlbChcImRpdlwiLCBcInBhbmVsIGhlbHBcIik7XG4gIHBhbmVsLmFwcGVuZChlbChcImgyXCIsIFwiXCIsIFwiS2V5Ym9hcmQgc2hvcnRjdXRzXCIpKTtcbiAgY29uc3QgdGFibGUgPSBlbChcInRhYmxlXCIsIFwiXCIpO1xuICBmb3IgKGNvbnN0IHsga2V5cywgZGVzY3JpcHRpb24gfSBvZiByb3dzKSB7XG4gICAgY29uc3QgdHIgPSBlbChcInRyXCIsIFwiXCIpO1xuICAgIGNvbnN0IGsgPSBlbChcInRkXCIsIFwia2V5c1wiKTtcbiAgICBrLmFwcGVuZChlbChcImtiZFwiLCBcIlwiLCBrZXlzKSk7XG4gICAgdHIuYXBwZW5kKGssIGVsKFwidGRcIiwgXCJcIiwgZGVzY3JpcHRpb24pKTtcbiAgICB0YWJsZS5hcHBlbmQodHIpO1xuICB9XG4gIHBhbmVsLmFwcGVuZCh0YWJsZSk7XG4gIGJhY2tkcm9wLmFwcGVuZChwYW5lbCk7XG4gIHJvb3QuYXBwZW5kKGJhY2tkcm9wKTtcblxuICBjb25zdCBjbG9zZSA9ICgpID0+IHtcbiAgICBiYWNrZHJvcC5yZW1vdmUoKTtcbiAgICB3aW5kb3cucmVtb3ZlRXZlbnRMaXN0ZW5lcihcImtleWRvd25cIiwgb25LZXksIHRydWUpO1xuICB9O1xuICBjb25zdCBvbktleSA9IChlOiBLZXlib2FyZEV2ZW50KSA9PiB7XG4gICAgZS5wcmV2ZW50RGVmYXVsdCgpO1xuICAgIGUuc3RvcFByb3BhZ2F0aW9uKCk7XG4gICAgY2xvc2UoKTtcbiAgfTtcbiAgYmFja2Ryb3AuYWRkRXZlbnRMaXN0ZW5lcihcImNsaWNrXCIsIChlKSA9PiB7XG4gICAgaWYgKGUudGFyZ2V0ID09PSBiYWNrZHJvcCkgY2xvc2UoKTtcbiAgfSk7XG4gIHdpbmRvdy5hZGRFdmVudExpc3RlbmVyKFwia2V5ZG93blwiLCBvbktleSwgdHJ1ZSk7XG59XG5cbmV4cG9ydCB7IGVsIH07XG4iLCAiLy8gUHVyZSBmdXp6eSBtYXRjaGVyIGZvciB0aGUgZm9sZGVyIHBhbGV0dGU6IGNhc2UtaW5zZW5zaXRpdmUgc3Vic2VxdWVuY2UgbWF0Y2hpbmcsIHNjb3JlZCB0b1xuLy8gZmF2b3VyIHdvcmQgc3RhcnRzIGFuZCBjb25zZWN1dGl2ZSBydW5zLCB0aGUgd2F5IEZhc3RtYWlsJ3MgYW5kIG1vc3QgZWRpdG9ycycgcGlja2VycyBmZWVsLlxuXG5leHBvcnQgaW50ZXJmYWNlIEZ1enp5TWF0Y2g8VD4ge1xuICBpdGVtOiBUO1xuICBzY29yZTogbnVtYmVyO1xuICAvKiogSW5kaWNlcyBpbnRvIHRoZSBpdGVtJ3MgdGV4dCB0aGF0IG1hdGNoZWQsIGZvciBoaWdobGlnaHRpbmcuICovXG4gIGluZGljZXM6IG51bWJlcltdO1xufVxuXG5jb25zdCBpc0JvdW5kYXJ5ID0gKHRleHQ6IHN0cmluZywgaTogbnVtYmVyKSA9PiBpID09PSAwIHx8IC9bXFxzL1xcXFwuXy1dLy50ZXN0KHRleHRbaSAtIDFdID8/IFwiXCIpO1xuXG5mdW5jdGlvbiBpc1N1YnNlcXVlbmNlKHF1ZXJ5OiBzdHJpbmcsIHRleHQ6IHN0cmluZywgZnJvbTogbnVtYmVyKTogYm9vbGVhbiB7XG4gIGxldCBhdCA9IGZyb207XG4gIGZvciAoY29uc3QgY2ggb2YgcXVlcnkpIHtcbiAgICBhdCA9IHRleHQuaW5kZXhPZihjaCwgYXQpICsgMTtcbiAgICBpZiAoYXQgPT09IDApIHJldHVybiBmYWxzZTtcbiAgfVxuICByZXR1cm4gdHJ1ZTtcbn1cblxuLyoqIFNjb3JlIGBxdWVyeWAgYWdhaW5zdCBgdGV4dGAsIG9yIG51bGwgaWYgaXQgaXMgbm90IGEgc3Vic2VxdWVuY2UuIEhpZ2hlciBpcyBiZXR0ZXIuICovXG5leHBvcnQgZnVuY3Rpb24gc2NvcmUocXVlcnk6IHN0cmluZywgdGV4dDogc3RyaW5nKTogeyBzY29yZTogbnVtYmVyOyBpbmRpY2VzOiBudW1iZXJbXSB9IHwgbnVsbCB7XG4gIGNvbnN0IHEgPSBxdWVyeS50b0xvd2VyQ2FzZSgpO1xuICBjb25zdCB0ID0gdGV4dC50b0xvd2VyQ2FzZSgpO1xuICBpZiAocS5sZW5ndGggPT09IDApIHJldHVybiB7IHNjb3JlOiAwLCBpbmRpY2VzOiBbXSB9O1xuXG4gIGNvbnN0IGluZGljZXM6IG51bWJlcltdID0gW107XG4gIGxldCB0b3RhbCA9IDA7XG4gIGxldCBmcm9tID0gMDtcbiAgZm9yIChjb25zdCBbbiwgY2hdIG9mIFsuLi5xXS5lbnRyaWVzKCkpIHtcbiAgICAvLyBQcmVmZXIgdGhlIG5leHQgb2NjdXJyZW5jZSBhdCBhIHdvcmQgYm91bmRhcnksIGFzIGxvbmcgYXMgdGhlIHJlc3Qgb2YgdGhlIHF1ZXJ5IHN0aWxsXG4gICAgLy8gZml0cyBhZnRlciBpdDsgb3RoZXJ3aXNlIHRha2UgdGhlIG5lYXJlc3Qgb25lLlxuICAgIGNvbnN0IHJlc3QgPSBxLnNsaWNlKG4gKyAxKTtcbiAgICBsZXQgYXQgPSAtMTtcbiAgICBmb3IgKGxldCBpID0gdC5pbmRleE9mKGNoLCBmcm9tKTsgaSA+PSAwOyBpID0gdC5pbmRleE9mKGNoLCBpICsgMSkpIHtcbiAgICAgIGlmICghaXNTdWJzZXF1ZW5jZShyZXN0LCB0LCBpICsgMSkpIGJyZWFrO1xuICAgICAgaWYgKGF0IDwgMCkgYXQgPSBpO1xuICAgICAgaWYgKGlzQm91bmRhcnkodGV4dCwgaSkpIHtcbiAgICAgICAgYXQgPSBpO1xuICAgICAgICBicmVhaztcbiAgICAgIH1cbiAgICB9XG4gICAgaWYgKGF0IDwgMCkgcmV0dXJuIG51bGw7XG4gICAgY29uc3QgcHJldiA9IGluZGljZXMuYXQoLTEpO1xuICAgIHRvdGFsICs9IDE7XG4gICAgaWYgKGlzQm91bmRhcnkodGV4dCwgYXQpKSB0b3RhbCArPSA4O1xuICAgIGlmIChwcmV2ICE9PSB1bmRlZmluZWQgJiYgYXQgPT09IHByZXYgKyAxKSB0b3RhbCArPSA1O1xuICAgIGlmIChwcmV2ICE9PSB1bmRlZmluZWQpIHRvdGFsIC09IE1hdGgubWluKGF0IC0gcHJldiAtIDEsIDMpO1xuICAgIGluZGljZXMucHVzaChhdCk7XG4gICAgZnJvbSA9IGF0ICsgMTtcbiAgfVxuICAvLyBTaG9ydGVyIHRhcmdldHMgd2luIHRpZXMsIHNvIFwiSW5ib3hcIiBiZWF0cyBcIkluYm94L1JlY2VpcHRzLzIwMTlcIiBmb3IgXCJpbmJcIi5cbiAgdG90YWwgLT0gdGV4dC5sZW5ndGggKiAwLjAxO1xuICByZXR1cm4geyBzY29yZTogdG90YWwsIGluZGljZXMgfTtcbn1cblxuZXhwb3J0IGZ1bmN0aW9uIGZ1enp5RmlsdGVyPFQ+KFxuICBxdWVyeTogc3RyaW5nLFxuICBpdGVtczogcmVhZG9ubHkgVFtdLFxuICB0ZXh0OiAoaXRlbTogVCkgPT4gc3RyaW5nLFxuKTogRnV6enlNYXRjaDxUPltdIHtcbiAgY29uc3QgbWF0Y2hlczogRnV6enlNYXRjaDxUPltdID0gW107XG4gIGZvciAoY29uc3QgaXRlbSBvZiBpdGVtcykge1xuICAgIGNvbnN0IHJlc3VsdCA9IHNjb3JlKHF1ZXJ5LCB0ZXh0KGl0ZW0pKTtcbiAgICBpZiAocmVzdWx0KSBtYXRjaGVzLnB1c2goeyBpdGVtLCAuLi5yZXN1bHQgfSk7XG4gIH1cbiAgcmV0dXJuIG1hdGNoZXMuc29ydCgoYSwgYikgPT4gYi5zY29yZSAtIGEuc2NvcmUpO1xufVxuIiwgIi8vIEZvbGRlciBwYWxldHRlOiBGYXN0bWFpbCdzIGBnYCAoZ28gdG8gZm9sZGVyKSB0eXBlLWFoZWFkIHBpY2tlciwgcGx1cyBPdXRsb29rJ3MgYXBwcy5cbi8vIE1vdmluZyAoYG1gKSB1c2VzIE91dGxvb2sncyBvd24gXCJNb3ZlIHRvXCIgcGlja2VyIGZvciBub3c7IHNlZSBkb2NzL2RlY2lzaW9ucy5tZC5cblxuaW1wb3J0IHsgdHlwZSBGb2xkZXJOb2RlLCByZWFkRm9sZGVyVHJlZSwgcmVzb2x2ZSwgdHlwZSBTZWxlY3Rvck5hbWUgfSBmcm9tIFwiLi4vc2VsZWN0b3JzLnRzXCI7XG5pbXBvcnQgeyBlbCwgb3ZlcmxheVJvb3QsIHRvYXN0IH0gZnJvbSBcIi4uL3VpL292ZXJsYXkudHNcIjtcbmltcG9ydCB7IHR5cGUgRnV6enlNYXRjaCwgZnV6enlGaWx0ZXIgfSBmcm9tIFwiLi9mdXp6eS50c1wiO1xuXG5jb25zdCBNQVhfUkVTVUxUUyA9IDEyO1xuXG5mdW5jdGlvbiBoaWdobGlnaHQodGV4dDogc3RyaW5nLCBpbmRpY2VzOiByZWFkb25seSBudW1iZXJbXSk6IERvY3VtZW50RnJhZ21lbnQge1xuICBjb25zdCBmcmFnID0gZG9jdW1lbnQuY3JlYXRlRG9jdW1lbnRGcmFnbWVudCgpO1xuICBjb25zdCBoaXQgPSBuZXcgU2V0KGluZGljZXMpO1xuICBmb3IgKGNvbnN0IFtpLCBjaF0gb2YgWy4uLnRleHRdLmVudHJpZXMoKSkge1xuICAgIGZyYWcuYXBwZW5kKGhpdC5oYXMoaSkgPyBlbChcIm1hcmtcIiwgXCJcIiwgY2gpIDogZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUoY2gpKTtcbiAgfVxuICByZXR1cm4gZnJhZztcbn1cblxuY29uc3QgQVBQUzogW3N0cmluZywgU2VsZWN0b3JOYW1lXVtdID0gW1xuICBbXCJNYWlsXCIsIFwibWFpbEFwcFwiXSxcbiAgW1wiQ2FsZW5kYXJcIiwgXCJjYWxlbmRhckFwcFwiXSxcbiAgW1wiUGVvcGxlXCIsIFwicGVvcGxlQXBwXCJdLFxuXTtcblxuLyoqIE91dGxvb2sncyBhcHBzLCBsaXN0ZWQgYWxvbmdzaWRlIGZvbGRlcnMgc28gYGdgIHRoZW4gXCJjYWxcIiBzd2l0Y2hlcyB0byBDYWxlbmRhci4gKi9cbmZ1bmN0aW9uIGFwcFRhcmdldHMoKTogRm9sZGVyTm9kZVtdIHtcbiAgcmV0dXJuIEFQUFMuZmxhdE1hcCgoW25hbWUsIHRhcmdldF0pID0+IHtcbiAgICBjb25zdCBlbGVtZW50ID0gcmVzb2x2ZSh0YXJnZXQpO1xuICAgIHJldHVybiBlbGVtZW50ID8gW3sgbmFtZSwgcGF0aDogbmFtZSwgZWxlbWVudCB9XSA6IFtdO1xuICB9KTtcbn1cblxuZXhwb3J0IGZ1bmN0aW9uIG9wZW5Gb2xkZXJQYWxldHRlKCk6IHZvaWQge1xuICBjb25zdCBmb2xkZXJzID0gWy4uLmFwcFRhcmdldHMoKSwgLi4ucmVhZEZvbGRlclRyZWUoKV07XG4gIGlmIChmb2xkZXJzLmxlbmd0aCA9PT0gMCkge1xuICAgIHRvYXN0KFwiT0JCIGNvdWxkbid0IGZpbmQgdGhlIGZvbGRlciBwYW5lXCIpO1xuICAgIHJldHVybjtcbiAgfVxuXG4gIGNvbnN0IHJvb3QgPSBvdmVybGF5Um9vdCgpO1xuICBjb25zdCBiYWNrZHJvcCA9IGVsKFwiZGl2XCIsIFwiYmFja2Ryb3BcIik7XG4gIGNvbnN0IHBhbmVsID0gZWwoXCJkaXZcIiwgXCJwYW5lbCBwYWxldHRlXCIpO1xuICBjb25zdCBpbnB1dCA9IGVsKFwiaW5wdXRcIiwgXCJcIik7XG4gIGlucHV0LnBsYWNlaG9sZGVyID0gXCJHbyB0byBmb2xkZXIsIG9yIE1haWwsIENhbGVuZGFyLCBQZW9wbGVcdTIwMjZcIjtcbiAgaW5wdXQuc2V0QXR0cmlidXRlKFwiYXJpYS1sYWJlbFwiLCBpbnB1dC5wbGFjZWhvbGRlcik7XG4gIGNvbnN0IGxpc3QgPSBlbChcInVsXCIsIFwiXCIpO1xuICBsaXN0LnNldEF0dHJpYnV0ZShcInJvbGVcIiwgXCJsaXN0Ym94XCIpO1xuICBwYW5lbC5hcHBlbmQoaW5wdXQsIGxpc3QpO1xuICBiYWNrZHJvcC5hcHBlbmQocGFuZWwpO1xuICByb290LmFwcGVuZChiYWNrZHJvcCk7XG5cbiAgbGV0IG1hdGNoZXM6IEZ1enp5TWF0Y2g8Rm9sZGVyTm9kZT5bXSA9IFtdO1xuICBsZXQgc2VsZWN0ZWQgPSAwO1xuXG4gIGNvbnN0IGNsb3NlID0gKCkgPT4gYmFja2Ryb3AucmVtb3ZlKCk7XG4gIGNvbnN0IGNvbW1pdCA9ICgpID0+IHtcbiAgICBjb25zdCBtYXRjaCA9IG1hdGNoZXNbc2VsZWN0ZWRdO1xuICAgIGNsb3NlKCk7XG4gICAgaWYgKG1hdGNoKSAobWF0Y2guaXRlbS5lbGVtZW50IGFzIEhUTUxFbGVtZW50KS5jbGljaygpO1xuICB9O1xuXG4gIGNvbnN0IHJlbmRlciA9ICgpID0+IHtcbiAgICBtYXRjaGVzID0gZnV6enlGaWx0ZXIoaW5wdXQudmFsdWUsIGZvbGRlcnMsIChmKSA9PiBmLnBhdGgpLnNsaWNlKDAsIE1BWF9SRVNVTFRTKTtcbiAgICBzZWxlY3RlZCA9IE1hdGgubWluKHNlbGVjdGVkLCBNYXRoLm1heChtYXRjaGVzLmxlbmd0aCAtIDEsIDApKTtcbiAgICBsaXN0LnJlcGxhY2VDaGlsZHJlbihcbiAgICAgIC4uLm1hdGNoZXMubWFwKChtLCBpKSA9PiB7XG4gICAgICAgIGNvbnN0IGxpID0gZWwoXCJsaVwiLCBcIlwiKTtcbiAgICAgICAgbGkuc2V0QXR0cmlidXRlKFwicm9sZVwiLCBcIm9wdGlvblwiKTtcbiAgICAgICAgbGkuc2V0QXR0cmlidXRlKFwiYXJpYS1zZWxlY3RlZFwiLCBTdHJpbmcoaSA9PT0gc2VsZWN0ZWQpKTtcbiAgICAgICAgbGkuYXBwZW5kKGhpZ2hsaWdodChtLml0ZW0ucGF0aCwgbS5pbmRpY2VzKSk7XG4gICAgICAgIGxpLmFkZEV2ZW50TGlzdGVuZXIoXCJtb3VzZWRvd25cIiwgKGUpID0+IHtcbiAgICAgICAgICBlLnByZXZlbnREZWZhdWx0KCk7XG4gICAgICAgICAgc2VsZWN0ZWQgPSBpO1xuICAgICAgICAgIGNvbW1pdCgpO1xuICAgICAgICB9KTtcbiAgICAgICAgcmV0dXJuIGxpO1xuICAgICAgfSksXG4gICAgKTtcbiAgICBpZiAobWF0Y2hlcy5sZW5ndGggPT09IDApIGxpc3QuYXBwZW5kKGVsKFwibGlcIiwgXCJlbXB0eVwiLCBcIk5vIG1hdGNoaW5nIGZvbGRlcnNcIikpO1xuICB9O1xuXG4gIGlucHV0LmFkZEV2ZW50TGlzdGVuZXIoXCJpbnB1dFwiLCAoKSA9PiB7XG4gICAgc2VsZWN0ZWQgPSAwO1xuICAgIHJlbmRlcigpO1xuICB9KTtcbiAgaW5wdXQuYWRkRXZlbnRMaXN0ZW5lcihcImtleWRvd25cIiwgKGUpID0+IHtcbiAgICBlLnN0b3BQcm9wYWdhdGlvbigpO1xuICAgIGlmIChlLmtleSA9PT0gXCJFc2NhcGVcIikgY2xvc2UoKTtcbiAgICBlbHNlIGlmIChlLmtleSA9PT0gXCJFbnRlclwiKSBjb21taXQoKTtcbiAgICBlbHNlIGlmIChlLmtleSA9PT0gXCJBcnJvd0Rvd25cIiB8fCAoZS5jdHJsS2V5ICYmIGUua2V5ID09PSBcIm5cIikpIHNlbGVjdGVkKys7XG4gICAgZWxzZSBpZiAoZS5rZXkgPT09IFwiQXJyb3dVcFwiIHx8IChlLmN0cmxLZXkgJiYgZS5rZXkgPT09IFwicFwiKSlcbiAgICAgIHNlbGVjdGVkID0gTWF0aC5tYXgoc2VsZWN0ZWQgLSAxLCAwKTtcbiAgICBlbHNlIHJldHVybjtcbiAgICBlLnByZXZlbnREZWZhdWx0KCk7XG4gICAgcmVuZGVyKCk7XG4gIH0pO1xuICBiYWNrZHJvcC5hZGRFdmVudExpc3RlbmVyKFwibW91c2Vkb3duXCIsIChlKSA9PiB7XG4gICAgaWYgKGUudGFyZ2V0ID09PSBiYWNrZHJvcCkgY2xvc2UoKTtcbiAgfSk7XG5cbiAgcmVuZGVyKCk7XG4gIGlucHV0LmZvY3VzKCk7XG59XG4iLCAiLy8gVHVybnMgYW4gYWN0aW9uIGludG8gRE9NIGVmZmVjdHMuIFRoZSBvbmx5IGtleW1hcCBmaWxlIHRoYXQgdG91Y2hlcyB0aGUgcGFnZS5cblxuaW1wb3J0IHsgb3BlbkZvbGRlclBhbGV0dGUgfSBmcm9tIFwiLi4vcGFsZXR0ZS9wYWxldHRlLnRzXCI7XG5pbXBvcnQgeyBpc1VucmVhZFJvdywgcmVzb2x2ZSwgc2VsZWN0ZWRSb3dzIH0gZnJvbSBcIi4uL3NlbGVjdG9ycy50c1wiO1xuaW1wb3J0IHsgc2hvd0hlbHAsIHRvYXN0IH0gZnJvbSBcIi4uL3VpL292ZXJsYXkudHNcIjtcbmltcG9ydCB7IEFDVElPTlMsIHR5cGUgQWN0aW9uSWQsIHR5cGUgRXh0ZW5zaW9uQ29tbWFuZCB9IGZyb20gXCIuL2FjdGlvbnMudHNcIjtcbmltcG9ydCB7IEJJTkRJTkdTIH0gZnJvbSBcIi4vYmluZGluZ3MudHNcIjtcbmltcG9ydCB0eXBlIHsgQ29udGV4dCB9IGZyb20gXCIuL2VuZ2luZS50c1wiO1xuXG4vKiogV2hpY2ggYmluZGluZ3MgYXBwbHksIGp1ZGdlZCBmcm9tIHdoZXJlIGZvY3VzIGlzLiAqL1xuZXhwb3J0IGZ1bmN0aW9uIGN1cnJlbnRDb250ZXh0KCk6IENvbnRleHQge1xuICBjb25zdCByZWFkaW5nID0gcmVzb2x2ZShcInJlYWRpbmdQYW5lXCIpO1xuICBjb25zdCBmb2N1c2VkID0gZG9jdW1lbnQuYWN0aXZlRWxlbWVudDtcbiAgcmV0dXJuIHJlYWRpbmcgJiYgZm9jdXNlZCAmJiByZWFkaW5nLmNvbnRhaW5zKGZvY3VzZWQpID8gXCJyZWFkaW5nXCIgOiBcImxpc3RcIjtcbn1cblxuZnVuY3Rpb24gcnVuRXh0ZW5zaW9uKGNvbW1hbmQ6IEV4dGVuc2lvbkNvbW1hbmQpOiB2b2lkIHtcbiAgc3dpdGNoIChjb21tYW5kKSB7XG4gICAgY2FzZSBcImhlbHBcIjpcbiAgICAgIHNob3dIZWxwKFxuICAgICAgICBCSU5ESU5HUy5tYXAoKGIpID0+ICh7XG4gICAgICAgICAga2V5czogYi5rZXlzLmpvaW4oXCIgXCIpLFxuICAgICAgICAgIGRlc2NyaXB0aW9uOiBBQ1RJT05TW2IuYWN0aW9uXS5kZXNjcmlwdGlvbixcbiAgICAgICAgfSkpLFxuICAgICAgKTtcbiAgICAgIHJldHVybjtcbiAgICBjYXNlIFwicGFsZXR0ZS1nb1wiOlxuICAgICAgb3BlbkZvbGRlclBhbGV0dGUoKTtcbiAgICAgIHJldHVybjtcbiAgfVxufVxuXG4vKiogUnVuIGFuIGFjdGlvbi4gUmV0dXJucyBmYWxzZSB3aGVuIHRoZSBrZXkgc2hvdWxkIGJlIGxlZnQgZm9yIE91dGxvb2sgKHBhc3N0aHJvdWdoKS4gKi9cbmV4cG9ydCBmdW5jdGlvbiBleGVjdXRlKGlkOiBBY3Rpb25JZCk6IGJvb2xlYW4ge1xuICBjb25zdCB7IHN0cmF0ZWd5LCBkZXNjcmlwdGlvbiB9ID0gQUNUSU9OU1tpZF07XG4gIHN3aXRjaCAoc3RyYXRlZ3kua2luZCkge1xuICAgIGNhc2UgXCJwYXNzdGhyb3VnaFwiOlxuICAgICAgcmV0dXJuIGZhbHNlO1xuICAgIGNhc2UgXCJkcml2ZVwiOiB7XG4gICAgICBjb25zdCB0YXJnZXQgPSByZXNvbHZlKHN0cmF0ZWd5LnRhcmdldCk7XG4gICAgICBpZiAodGFyZ2V0IGluc3RhbmNlb2YgSFRNTEVsZW1lbnQpIHRhcmdldC5jbGljaygpO1xuICAgICAgZWxzZSB0b2FzdChgJHtkZXNjcmlwdGlvbn06IGNvdWxkbid0IGZpbmQgdGhlIGJ1dHRvbmApO1xuICAgICAgcmV0dXJuIHRydWU7XG4gICAgfVxuICAgIGNhc2UgXCJ0b2dnbGVcIjoge1xuICAgICAgY29uc3Qgcm93cyA9IHNlbGVjdGVkUm93cygpO1xuICAgICAgaWYgKHJvd3MubGVuZ3RoID09PSAwKSB7XG4gICAgICAgIHRvYXN0KGAke2Rlc2NyaXB0aW9ufTogc2VsZWN0IGEgbWVzc2FnZSBmaXJzdGApO1xuICAgICAgICByZXR1cm4gdHJ1ZTtcbiAgICAgIH1cbiAgICAgIGNvbnN0IHdhbnRVbnJlYWQgPSBzdHJhdGVneS53YW50ID09PSBcInVucmVhZFwiO1xuICAgICAgaWYgKHJvd3MuZXZlcnkoKHJvdykgPT4gaXNVbnJlYWRSb3cocm93KSA9PT0gd2FudFVucmVhZCkpIHJldHVybiB0cnVlO1xuICAgICAgY29uc3QgdGFyZ2V0ID0gcmVzb2x2ZShzdHJhdGVneS50YXJnZXQpO1xuICAgICAgaWYgKHRhcmdldCBpbnN0YW5jZW9mIEhUTUxFbGVtZW50KSB0YXJnZXQuY2xpY2soKTtcbiAgICAgIGVsc2UgdG9hc3QoYCR7ZGVzY3JpcHRpb259OiBjb3VsZG4ndCBmaW5kIHRoZSBidXR0b25gKTtcbiAgICAgIHJldHVybiB0cnVlO1xuICAgIH1cbiAgICBjYXNlIFwiZXh0ZW5zaW9uXCI6XG4gICAgICBydW5FeHRlbnNpb24oc3RyYXRlZ3kuY29tbWFuZCk7XG4gICAgICByZXR1cm4gdHJ1ZTtcbiAgICBjYXNlIFwidG9kb1wiOlxuICAgICAgdG9hc3QoYCR7ZGVzY3JpcHRpb259IGlzbid0IGltcGxlbWVudGVkIHlldGApO1xuICAgICAgcmV0dXJuIHRydWU7XG4gIH1cbn1cbiIsICIvLyBVc2VyIHByZWZlcmVuY2VzLiBTdG9yZWQgaW4gY2hyb21lLnN0b3JhZ2Uuc3luYzsgbm90aGluZyBlbHNlIGlzIGV2ZXIgcGVyc2lzdGVkLlxuLy9cbi8vIEVhY2ggZW5hYmxlZCBmZWF0dXJlIGlkIGlzIHdyaXR0ZW4gYXMgYSB0b2tlbiBvbiA8aHRtbCBkYXRhLW92ZXJsb29rLWZlYXR1cmVzPVwiLi4uXCI+LCBhbmQgZWFjaFxuLy8gY2hvaWNlIGFzIGl0cyBvd24gYXR0cmlidXRlICg8aHRtbCBkYXRhLW92ZXJsb29rLXBhbGV0dGU9XCJzb2xhcml6ZWRcIj4pLCBzbyBzdHlsZXNoZWV0cyBnYXRlXG4vLyB0aGVtc2VsdmVzIHdpdGggYFtkYXRhLW92ZXJsb29rLWZlYXR1cmVzfj1cImRlbnNpdHlcIl1gIG9yIGBbZGF0YS1vdmVybG9vay1wYWxldHRlPVwic29sYXJpemVkXCJdYC5cblxuZXhwb3J0IHR5cGUgRmVhdHVyZUlkID0ga2V5b2YgdHlwZW9mIEZFQVRVUkVTO1xuZXhwb3J0IHR5cGUgQ2hvaWNlSWQgPSBrZXlvZiB0eXBlb2YgQ0hPSUNFUztcbmV4cG9ydCB0eXBlIFNldHRpbmdzID0gUmVjb3JkPEZlYXR1cmVJZCwgYm9vbGVhbj4gJiB7IFtLIGluIENob2ljZUlkXTogQ2hvaWNlVmFsdWU8Sz4gfTtcbmV4cG9ydCB0eXBlIENob2ljZVZhbHVlPEsgZXh0ZW5kcyBDaG9pY2VJZD4gPSBrZXlvZiAodHlwZW9mIENIT0lDRVMpW0tdW1wib3B0aW9uc1wiXTtcblxudHlwZSBHcm91cCA9IFwiR2VuZXJhbFwiIHwgXCJLZXlib2FyZFwiIHwgXCJEZWNsdXR0ZXJcIiB8IFwiTG9va1wiO1xuXG5pbnRlcmZhY2UgRmVhdHVyZSB7XG4gIGxhYmVsOiBzdHJpbmc7XG4gIGdyb3VwOiBHcm91cDtcbiAgZGVmYXVsdDogYm9vbGVhbjtcbn1cblxuaW50ZXJmYWNlIENob2ljZSB7XG4gIGxhYmVsOiBzdHJpbmc7XG4gIGdyb3VwOiBHcm91cDtcbiAgb3B0aW9uczogUmVjb3JkPHN0cmluZywgc3RyaW5nPjtcbiAgZGVmYXVsdDogc3RyaW5nO1xufVxuXG5leHBvcnQgY29uc3QgRkVBVFVSRVMgPSB7XG4gIGhlYWx0aEluZGljYXRvcjoge1xuICAgIGxhYmVsOiBcIlNob3cgYW4gaW5kaWNhdG9yIHdoZW4gT3V0bG9vaydzIHBhZ2UgY2hhbmdlcyBicmVhayBPQkJcIixcbiAgICBncm91cDogXCJHZW5lcmFsXCIsXG4gICAgZGVmYXVsdDogdHJ1ZSxcbiAgfSxcbiAga2V5bWFwOiB7IGxhYmVsOiBcIkZhc3RtYWlsLXN0eWxlIGtleWJvYXJkIHNob3J0Y3V0c1wiLCBncm91cDogXCJLZXlib2FyZFwiLCBkZWZhdWx0OiB0cnVlIH0sXG4gIFwiaGlkZS11cHNlbGxcIjogeyBsYWJlbDogXCJVcGdyYWRlIGFuZCAndHJ5IHRoZSBuZXdcdTIwMjYnIGJhbm5lcnNcIiwgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsIGRlZmF1bHQ6IHRydWUgfSxcbiAgXCJoaWRlLWFwcC1iYW5uZXJzXCI6IHtcbiAgICBsYWJlbDogXCJOb3RpY2VzIGFjcm9zcyB0aGUgdG9wIG9mIE91dGxvb2sgKHBvcC11cCwgc2lnbi1pbiBhbmQgc3RvcmFnZSByZW1pbmRlcnMpXCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogZmFsc2UsXG4gIH0sXG4gIFwiaGlkZS1pbmZvLWJhcnNcIjoge1xuICAgIGxhYmVsOiBcIk5vdGljZXMgYWJvdmUgYW4gb3BlbiBtZXNzYWdlICgnWW91IHJlcGxpZWQgb25cdTIwMjYnLCAnRXh0cmEgbGluZSBicmVha3MgcmVtb3ZlZCcpXCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogZmFsc2UsXG4gIH0sXG4gIFwiaGlkZS1jb3BpbG90XCI6IHsgbGFiZWw6IFwiQ29waWxvdCBidXR0b25zIGFuZCBwYW5lbHNcIiwgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsIGRlZmF1bHQ6IHRydWUgfSxcbiAgXCJoaWRlLWhlYWRlci1leHRyYXNcIjoge1xuICAgIGxhYmVsOiBcIlRlYW1zIGFuZCBNeSBEYXkgYnV0dG9ucyBpbiB0aGUgaGVhZGVyXCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogdHJ1ZSxcbiAgfSxcbiAgXCJoaWRlLXJhaWwtZXh0cmFzXCI6IHtcbiAgICBsYWJlbDogXCJUbyBEbywgTmV3c2xldHRlcnMsIE9yZyBFeHBsb3JlciBhbmQgT25lRHJpdmUgaW4gdGhlIGFwcCByYWlsXCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogdHJ1ZSxcbiAgfSxcbiAgXCJoaWRlLXJpYmJvbi1leHRyYXNcIjoge1xuICAgIGxhYmVsOiBcIlRlYW1zLCBTd2VlcCwgUXVpY2sgc3RlcHMsIFZpdmEgYW5kIGFkZC1pbiBidXR0b25zIGluIHRoZSByaWJib25cIixcbiAgICBncm91cDogXCJEZWNsdXR0ZXJcIixcbiAgICBkZWZhdWx0OiB0cnVlLFxuICB9LFxuICBcImhpZGUtZXh0ZXJuYWwtdGFnXCI6IHsgbGFiZWw6IFwiJ0V4dGVybmFsJyBzZW5kZXIgdGFnc1wiLCBncm91cDogXCJEZWNsdXR0ZXJcIiwgZGVmYXVsdDogdHJ1ZSB9LFxuICBcImhpZGUtcm93LWFjdGlvbnNcIjoge1xuICAgIGxhYmVsOiBcIkZsYWcgYW5kIHBpbiBpY29ucyB0aGF0IGFwcGVhciB3aGVuIGhvdmVyaW5nIGEgbWVzc2FnZVwiLFxuICAgIGdyb3VwOiBcIkRlY2x1dHRlclwiLFxuICAgIGRlZmF1bHQ6IGZhbHNlLFxuICB9LFxuICBcImhpZGUtYXZhdGFyc1wiOiB7XG4gICAgbGFiZWw6IFwiU2VuZGVyIHBpY3R1cmVzIGluIHRoZSBtZXNzYWdlIGxpc3RcIixcbiAgICBncm91cDogXCJEZWNsdXR0ZXJcIixcbiAgICBkZWZhdWx0OiBmYWxzZSxcbiAgfSxcbiAgXCJoaWRlLWdyb3VwLWhlYWRlcnNcIjoge1xuICAgIGxhYmVsOiBcIkRhdGUgZ3JvdXAgaGVhZGVycyAoJ1RoaXMgTW9udGgnKSBpbiB0aGUgbWVzc2FnZSBsaXN0XCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogZmFsc2UsXG4gIH0sXG4gIFwiaGlkZS1mb2N1c2VkLXRhYnNcIjogeyBsYWJlbDogXCJGb2N1c2VkIC8gT3RoZXIgdGFic1wiLCBncm91cDogXCJEZWNsdXR0ZXJcIiwgZGVmYXVsdDogZmFsc2UgfSxcbiAgXCJoaWRlLXJpYmJvbi10YWJzXCI6IHtcbiAgICBsYWJlbDogXCJSaWJib24gdGFicyAoSG9tZSwgVmlldywgSGVscClcIixcbiAgICBncm91cDogXCJEZWNsdXR0ZXJcIixcbiAgICBkZWZhdWx0OiBmYWxzZSxcbiAgfSxcbiAgXCJoaWRlLWFwcC1yYWlsXCI6IHsgbGFiZWw6IFwiQXBwIHJhaWwgKGxlZnQgaWNvbiBiYXIpXCIsIGdyb3VwOiBcIkRlY2x1dHRlclwiLCBkZWZhdWx0OiBmYWxzZSB9LFxuICBcImhpZGUtcmliYm9uXCI6IHsgbGFiZWw6IFwiUmliYm9uXCIsIGdyb3VwOiBcIkRlY2x1dHRlclwiLCBkZWZhdWx0OiBmYWxzZSB9LFxuICBkZW5zaXR5OiB7IGxhYmVsOiBcIlRpZ2h0ZXIsIGRlbnNlciBsYXlvdXRcIiwgZ3JvdXA6IFwiTG9va1wiLCBkZWZhdWx0OiB0cnVlIH0sXG59IGFzIGNvbnN0IHNhdGlzZmllcyBSZWNvcmQ8c3RyaW5nLCBGZWF0dXJlPjtcblxuZXhwb3J0IGNvbnN0IENIT0lDRVMgPSB7XG4gIHBhbGV0dGU6IHtcbiAgICBsYWJlbDogXCJDb2xvcnNcIixcbiAgICBncm91cDogXCJMb29rXCIsXG4gICAgb3B0aW9uczoge1xuICAgICAgb3V0bG9vazogXCJPdXRsb29rJ3Mgb3duXCIsXG4gICAgICBzb2xhcml6ZWQ6IFwiU29sYXJpemVkXCIsXG4gICAgICBiYXNlMTY6IFwiQmFzZTE2IERlZmF1bHRcIixcbiAgICAgIGZhc3RtYWlsOiBcIkZhc3RtYWlsXCIsXG4gICAgICBwYXBlcjogXCJQYXBlclwiLFxuICAgIH0sXG4gICAgZGVmYXVsdDogXCJvdXRsb29rXCIsXG4gIH0sXG4gIG1vZGU6IHtcbiAgICBsYWJlbDogXCJMaWdodCBvciBkYXJrXCIsXG4gICAgZ3JvdXA6IFwiTG9va1wiLFxuICAgIG9wdGlvbnM6IHsgYXV0bzogXCJGb2xsb3cgT3V0bG9va1wiLCBsaWdodDogXCJMaWdodFwiLCBkYXJrOiBcIkRhcmtcIiB9LFxuICAgIGRlZmF1bHQ6IFwiYXV0b1wiLFxuICB9LFxuICBsYXlvdXQ6IHtcbiAgICBsYWJlbDogXCJMYXlvdXRcIixcbiAgICBncm91cDogXCJMb29rXCIsXG4gICAgb3B0aW9uczoge1xuICAgICAgb3V0bG9vazogXCJPdXRsb29rJ3Mgb3duXCIsXG4gICAgICBjYWxtOiBcIkNhbG0gdGhyZWUtcGFuZVwiLFxuICAgICAgZmFzdG1haWw6IFwiRmFzdG1haWwgKGxpc3QgZmlyc3QpXCIsXG4gICAgICBmb2N1czogXCJGb2N1cyAoc2luZ2xlIGNvbHVtbilcIixcbiAgICB9LFxuICAgIGRlZmF1bHQ6IFwib3V0bG9va1wiLFxuICB9LFxufSBhcyBjb25zdCBzYXRpc2ZpZXMgUmVjb3JkPHN0cmluZywgQ2hvaWNlPjtcblxuZXhwb3J0IGNvbnN0IEZFQVRVUkVfSURTID0gT2JqZWN0LmtleXMoRkVBVFVSRVMpIGFzIEZlYXR1cmVJZFtdO1xuZXhwb3J0IGNvbnN0IENIT0lDRV9JRFMgPSBPYmplY3Qua2V5cyhDSE9JQ0VTKSBhcyBDaG9pY2VJZFtdO1xuXG5leHBvcnQgZnVuY3Rpb24gZGVmYXVsdFNldHRpbmdzKCk6IFNldHRpbmdzIHtcbiAgcmV0dXJuIHtcbiAgICAuLi5PYmplY3QuZnJvbUVudHJpZXMoRkVBVFVSRV9JRFMubWFwKChpZCkgPT4gW2lkLCBGRUFUVVJFU1tpZF0uZGVmYXVsdF0pKSxcbiAgICAuLi5PYmplY3QuZnJvbUVudHJpZXMoQ0hPSUNFX0lEUy5tYXAoKGlkKSA9PiBbaWQsIENIT0lDRVNbaWRdLmRlZmF1bHRdKSksXG4gIH0gYXMgU2V0dGluZ3M7XG59XG5cbi8qKiBNZXJnZSB3aGF0ZXZlciBpcyBzdG9yZWQgb3ZlciB0aGUgZGVmYXVsdHMsIGlnbm9yaW5nIHVua25vd24gb3IgbWlzdHlwZWQga2V5cy4gKi9cbmV4cG9ydCBmdW5jdGlvbiBub3JtYWxpemUoc3RvcmVkOiBSZWNvcmQ8c3RyaW5nLCB1bmtub3duPik6IFNldHRpbmdzIHtcbiAgY29uc3Qgc2V0dGluZ3MgPSBkZWZhdWx0U2V0dGluZ3MoKSBhcyBSZWNvcmQ8c3RyaW5nLCB1bmtub3duPjtcbiAgZm9yIChjb25zdCBpZCBvZiBGRUFUVVJFX0lEUykge1xuICAgIGlmICh0eXBlb2Ygc3RvcmVkW2lkXSA9PT0gXCJib29sZWFuXCIpIHNldHRpbmdzW2lkXSA9IHN0b3JlZFtpZF07XG4gIH1cbiAgZm9yIChjb25zdCBpZCBvZiBDSE9JQ0VfSURTKSB7XG4gICAgY29uc3QgdmFsdWUgPSBzdG9yZWRbaWRdO1xuICAgIGlmICh0eXBlb2YgdmFsdWUgPT09IFwic3RyaW5nXCIgJiYgdmFsdWUgaW4gQ0hPSUNFU1tpZF0ub3B0aW9ucykgc2V0dGluZ3NbaWRdID0gdmFsdWU7XG4gIH1cbiAgcmV0dXJuIHNldHRpbmdzIGFzIFNldHRpbmdzO1xufVxuXG5leHBvcnQgZnVuY3Rpb24gZmVhdHVyZVRva2VucyhzZXR0aW5nczogU2V0dGluZ3MpOiBzdHJpbmcge1xuICByZXR1cm4gRkVBVFVSRV9JRFMuZmlsdGVyKChpZCkgPT4gc2V0dGluZ3NbaWRdKS5qb2luKFwiIFwiKTtcbn1cblxuLyoqIFdyaXRlcyB0aGUgc2V0dGluZ3Mgb250byA8aHRtbD4gYXMgdGhlIGF0dHJpYnV0ZXMgdGhlIHN0eWxlc2hlZXRzIGtleSBvZmYuICovXG5leHBvcnQgZnVuY3Rpb24gYXBwbHlUb0RvY3VtZW50KHNldHRpbmdzOiBTZXR0aW5ncywgcm9vdDogSFRNTEVsZW1lbnQpOiB2b2lkIHtcbiAgcm9vdC5zZXRBdHRyaWJ1dGUoXCJkYXRhLW92ZXJsb29rLWZlYXR1cmVzXCIsIGZlYXR1cmVUb2tlbnMoc2V0dGluZ3MpKTtcbiAgZm9yIChjb25zdCBpZCBvZiBDSE9JQ0VfSURTKSByb290LnNldEF0dHJpYnV0ZShgZGF0YS1vdmVybG9vay0ke2lkfWAsIHNldHRpbmdzW2lkXSk7XG59XG5cbmV4cG9ydCBhc3luYyBmdW5jdGlvbiBsb2FkU2V0dGluZ3MoKTogUHJvbWlzZTxTZXR0aW5ncz4ge1xuICByZXR1cm4gbm9ybWFsaXplKGF3YWl0IGNocm9tZS5zdG9yYWdlLnN5bmMuZ2V0KG51bGwpKTtcbn1cblxuZXhwb3J0IGFzeW5jIGZ1bmN0aW9uIHNhdmVTZXR0aW5nKFxuICBpZDogRmVhdHVyZUlkIHwgQ2hvaWNlSWQsXG4gIHZhbHVlOiBib29sZWFuIHwgc3RyaW5nLFxuKTogUHJvbWlzZTx2b2lkPiB7XG4gIGF3YWl0IGNocm9tZS5zdG9yYWdlLnN5bmMuc2V0KHsgW2lkXTogdmFsdWUgfSk7XG59XG5cbmV4cG9ydCBmdW5jdGlvbiBvblNldHRpbmdzQ2hhbmdlZChjYWxsYmFjazogKHNldHRpbmdzOiBTZXR0aW5ncykgPT4gdm9pZCk6IHZvaWQge1xuICBjaHJvbWUuc3RvcmFnZS5vbkNoYW5nZWQuYWRkTGlzdGVuZXIoKF9jaGFuZ2VzLCBhcmVhKSA9PiB7XG4gICAgaWYgKGFyZWEgPT09IFwic3luY1wiKSB2b2lkIGxvYWRTZXR0aW5ncygpLnRoZW4oY2FsbGJhY2spO1xuICB9KTtcbn1cbiIsICIvLyBFbnRyeSBwb2ludDogYXBwbGllcyBzZXR0aW5ncywgdGFncyBlbGVtZW50cyBmb3IgQ1NTLCBydW5zIHRoZSBoZWFsdGggY2hlY2ssIG1vdW50cyB0aGUga2V5bWFwLlxuXG5pbXBvcnQgeyBBQ1RJT05TIH0gZnJvbSBcIi4va2V5bWFwL2FjdGlvbnMudHNcIjtcbmltcG9ydCB7IEJJTkRJTkdTIH0gZnJvbSBcIi4va2V5bWFwL2JpbmRpbmdzLnRzXCI7XG5pbXBvcnQgeyBERUZBVUxUX1RJTUVPVVRfTVMsIGtleUZyb21FdmVudCwgU2VxdWVuY2VFbmdpbmUgfSBmcm9tIFwiLi9rZXltYXAvZW5naW5lLnRzXCI7XG5pbXBvcnQgeyBjdXJyZW50Q29udGV4dCwgZXhlY3V0ZSB9IGZyb20gXCIuL2tleW1hcC9leGVjdXRlLnRzXCI7XG5pbXBvcnQge1xuICBicm9rZW5FbnRyaWVzLFxuICB0eXBlIEhlYWx0aFJlc3VsdCxcbiAgaGVhbHRoQ2hlY2ssXG4gIG91dGxvb2tDb2xvclNjaGVtZSxcbiAgcmVzb2x2ZSxcbiAgdGFnRWxlbWVudHMsXG59IGZyb20gXCIuL3NlbGVjdG9ycy50c1wiO1xuaW1wb3J0IHsgYXBwbHlUb0RvY3VtZW50LCBsb2FkU2V0dGluZ3MsIG9uU2V0dGluZ3NDaGFuZ2VkLCB0eXBlIFNldHRpbmdzIH0gZnJvbSBcIi4vc2V0dGluZ3MudHNcIjtcbmltcG9ydCB7IGlzRnJvbU92ZXJsYXksIGlzTW9kYWxPcGVuLCBzaG93SGVhbHRoQmFkZ2UsIHNob3dIaW50LCB0b2FzdCB9IGZyb20gXCIuL3VpL292ZXJsYXkudHNcIjtcblxuY29uc3QgTE9HID0gXCJbb2JiXVwiO1xuY29uc3QgQVBQX1JFQURZX1RJTUVPVVRfTVMgPSAyMF8wMDA7XG5jb25zdCBIRUFMVEhfU0VUVExFX01TID0gMV8wMDA7XG4vKiogV2FpdHMgYmV0d2VlbiByZS1jaGVja3Mgb2YgYSBmYWlsZWQgaGVhbHRoIGNoZWNrIGJlZm9yZSBpdCBpcyByZXBvcnRlZC4gKi9cbmNvbnN0IEhFQUxUSF9SRVRSWV9NUyA9IFszXzAwMCwgMTBfMDAwXTtcblxubGV0IHNldHRpbmdzOiBTZXR0aW5ncztcbmxldCBoZWFsdGhEaXNtaXNzZWQgPSBmYWxzZTtcblxuZnVuY3Rpb24gYXBwbHlTZXR0aW5ncyhuZXh0OiBTZXR0aW5ncyk6IHZvaWQge1xuICBzZXR0aW5ncyA9IG5leHQ7XG4gIGFwcGx5VG9Eb2N1bWVudChzZXR0aW5ncywgZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50KTtcbiAgc3luY0NvbG9yU2NoZW1lKCk7XG4gIGlmICghc2V0dGluZ3MuaGVhbHRoSW5kaWNhdG9yKSBzaG93SGVhbHRoQmFkZ2UoW10sICgpID0+IHt9KTtcbn1cblxuLyoqIFJlY29yZHMgd2hldGhlciBPdXRsb29rIHNob3dzIGEgcmVhZGluZyBwYW5lLCBzbyBsYXlvdXRzIG9ubHkgcmVzaGFwZSBwYW5lcyB3aGVuIGl0IGRvZXNuJ3QuICovXG5mdW5jdGlvbiBzeW5jUmVhZGluZ1BhbmUoKTogdm9pZCB7XG4gIGNvbnN0IHN0YXRlID0gcmVzb2x2ZShcInJlYWRpbmdQYW5lXCIpID8gXCJzaG93blwiIDogXCJub25lXCI7XG4gIGlmIChkb2N1bWVudC5kb2N1bWVudEVsZW1lbnQuZGF0YXNldC5vdmVybG9va1JlYWRpbmdQYW5lICE9PSBzdGF0ZSkge1xuICAgIGRvY3VtZW50LmRvY3VtZW50RWxlbWVudC5kYXRhc2V0Lm92ZXJsb29rUmVhZGluZ1BhbmUgPSBzdGF0ZTtcbiAgfVxufVxuXG4vKiogTWlycm9ycyBPdXRsb29rJ3Mgb3duIGxpZ2h0L2RhcmsgbW9kZSBvbnRvIDxodG1sPiBzbyBcIkZvbGxvdyBPdXRsb29rXCIgcGFsZXR0ZXMgY2FuIG1hdGNoIGl0LiAqL1xuZnVuY3Rpb24gc3luY0NvbG9yU2NoZW1lKCk6IHZvaWQge1xuICBjb25zdCBzY2hlbWUgPSBvdXRsb29rQ29sb3JTY2hlbWUoZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50KTtcbiAgaWYgKGRvY3VtZW50LmRvY3VtZW50RWxlbWVudC5nZXRBdHRyaWJ1dGUoXCJkYXRhLW92ZXJsb29rLW91dGxvb2stc2NoZW1lXCIpICE9PSBzY2hlbWUpIHtcbiAgICBkb2N1bWVudC5kb2N1bWVudEVsZW1lbnQuc2V0QXR0cmlidXRlKFwiZGF0YS1vdmVybG9vay1vdXRsb29rLXNjaGVtZVwiLCBzY2hlbWUpO1xuICB9XG59XG5cbmZ1bmN0aW9uIHJlcG9ydEhlYWx0aChyZXN1bHRzOiBIZWFsdGhSZXN1bHRbXSk6IHZvaWQge1xuICBjb25zdCBicm9rZW4gPSBicm9rZW5FbnRyaWVzKHJlc3VsdHMpO1xuICBjb25zdCBmYWxsYmFja3MgPSByZXN1bHRzLmZpbHRlcigocikgPT4gci5jYW5kaWRhdGUgPiAwKS5tYXAoKHIpID0+IHIubmFtZSk7XG4gIGRvY3VtZW50LmRvY3VtZW50RWxlbWVudC5kYXRhc2V0Lm92ZXJsb29rSGVhbHRoID0gYnJva2VuLmxlbmd0aCA/IFwiYnJva2VuXCIgOiBcIm9rXCI7XG4gIGNvbnNvbGUuZ3JvdXBDb2xsYXBzZWQoYCR7TE9HfSBoZWFsdGg6ICR7YnJva2VuLmxlbmd0aCA/IGAke2Jyb2tlbi5sZW5ndGh9IGJyb2tlbmAgOiBcIm9rXCJ9YCk7XG4gIGNvbnNvbGUudGFibGUocmVzdWx0cyk7XG4gIGNvbnNvbGUuZ3JvdXBFbmQoKTtcbiAgaWYgKGJyb2tlbi5sZW5ndGgpIGNvbnNvbGUud2FybihgJHtMT0d9IHNlbGVjdG9ycyBub3QgZm91bmQ6YCwgYnJva2VuLmpvaW4oXCIsIFwiKSk7XG4gIGlmIChmYWxsYmFja3MubGVuZ3RoKSBjb25zb2xlLmluZm8oYCR7TE9HfSB1c2luZyBmYWxsYmFjayBzZWxlY3RvcnMgZm9yOmAsIGZhbGxiYWNrcy5qb2luKFwiLCBcIikpO1xuICBpZiAoc2V0dGluZ3MuaGVhbHRoSW5kaWNhdG9yICYmICFoZWFsdGhEaXNtaXNzZWQpIHtcbiAgICBzaG93SGVhbHRoQmFkZ2UoYnJva2VuLCAoKSA9PiB7XG4gICAgICBoZWFsdGhEaXNtaXNzZWQgPSB0cnVlO1xuICAgIH0pO1xuICB9XG59XG5cbmxldCBoZWFsdGhUaW1lcjogUmV0dXJuVHlwZTx0eXBlb2Ygc2V0VGltZW91dD4gfCB1bmRlZmluZWQ7XG5cbi8qKlxuICogQ2hlY2tzIHRoZSByZWdpc3RyeSBvbmNlIE91dGxvb2sgaGFzIHNldHRsZWQuIFBhcnRzIG9mIHRoZSBwYWdlICh0aGUgYXBwIHJhaWwsIGZvciBvbmUpIHJlbmRlclxuICogYWZ0ZXIgdGhlIG1lc3NhZ2UgbGlzdCwgc28gYSBtaXNzIGlzIHJlLWNoZWNrZWQgYSBmZXcgdGltZXMgYmVmb3JlIGl0IGlzIHJlcG9ydGVkLlxuICovXG5mdW5jdGlvbiBzY2hlZHVsZUhlYWx0aENoZWNrKCk6IHZvaWQge1xuICBjbGVhclRpbWVvdXQoaGVhbHRoVGltZXIpO1xuICAvLyBPbmx5IHRoZSBtYWlsIHZpZXdzIGFyZSBzdXJ2ZXllZDsgQ2FsZW5kYXIgYW5kIFBlb3BsZSBsZWdpdGltYXRlbHkgbGFjayBtYWlsIGVsZW1lbnRzLlxuICBpZiAoIWxvY2F0aW9uLnBhdGhuYW1lLmluY2x1ZGVzKFwiL21haWxcIikpIHJldHVybjtcbiAgY29uc3QgYXR0ZW1wdCA9IChuOiBudW1iZXIpID0+IHtcbiAgICBjb25zdCByZXN1bHRzID0gaGVhbHRoQ2hlY2soKTtcbiAgICBjb25zdCByZXRyeUluID0gSEVBTFRIX1JFVFJZX01TW25dO1xuICAgIGlmIChicm9rZW5FbnRyaWVzKHJlc3VsdHMpLmxlbmd0aCA9PT0gMCB8fCByZXRyeUluID09PSB1bmRlZmluZWQpIHJlcG9ydEhlYWx0aChyZXN1bHRzKTtcbiAgICBlbHNlIGhlYWx0aFRpbWVyID0gc2V0VGltZW91dCgoKSA9PiBhdHRlbXB0KG4gKyAxKSwgcmV0cnlJbik7XG4gIH07XG4gIGhlYWx0aFRpbWVyID0gc2V0VGltZW91dCgoKSA9PiBhdHRlbXB0KDApLCBIRUFMVEhfU0VUVExFX01TKTtcbn1cblxuLyoqIFJlc29sdmVzIG9uY2UgT3V0bG9vaydzIG1haWwgVUkgaGFzIHJlbmRlcmVkLCBvciBhZnRlciBhIHRpbWVvdXQgc28gdGhlIGNoZWNrIHN0aWxsIHJlcG9ydHMuICovXG5mdW5jdGlvbiB3aGVuQXBwUmVhZHkoKTogUHJvbWlzZTx2b2lkPiB7XG4gIHJldHVybiBuZXcgUHJvbWlzZSgoZG9uZSkgPT4ge1xuICAgIGlmIChyZXNvbHZlKFwibWVzc2FnZUxpc3RcIikpIHJldHVybiBkb25lKCk7XG4gICAgY29uc3Qgb2JzZXJ2ZXIgPSBuZXcgTXV0YXRpb25PYnNlcnZlcigoKSA9PiB7XG4gICAgICBpZiAocmVzb2x2ZShcIm1lc3NhZ2VMaXN0XCIpKSBmaW5pc2goKTtcbiAgICB9KTtcbiAgICBjb25zdCB0aW1lciA9IHNldFRpbWVvdXQoZmluaXNoLCBBUFBfUkVBRFlfVElNRU9VVF9NUyk7XG4gICAgZnVuY3Rpb24gZmluaXNoKCkge1xuICAgICAgb2JzZXJ2ZXIuZGlzY29ubmVjdCgpO1xuICAgICAgY2xlYXJUaW1lb3V0KHRpbWVyKTtcbiAgICAgIGRvbmUoKTtcbiAgICB9XG4gICAgb2JzZXJ2ZXIub2JzZXJ2ZShkb2N1bWVudC5ib2R5LCB7IGNoaWxkTGlzdDogdHJ1ZSwgc3VidHJlZTogdHJ1ZSB9KTtcbiAgfSk7XG59XG5cbi8qKiBSZS10YWcgb24gRE9NIGNodXJuIChSZWFjdCByZS1yZW5kZXJzKSwgYW5kIHJlLWNoZWNrIGhlYWx0aCB3aGVuIHRoZSBTUEEgbmF2aWdhdGVzLiAqL1xuZnVuY3Rpb24gd2F0Y2hQYWdlKCk6IHZvaWQge1xuICBsZXQgc2NoZWR1bGVkID0gZmFsc2U7XG4gIGxldCBsYXN0VXJsID0gbG9jYXRpb24uaHJlZjtcbiAgbmV3IE11dGF0aW9uT2JzZXJ2ZXIoKCkgPT4ge1xuICAgIGlmIChzY2hlZHVsZWQpIHJldHVybjtcbiAgICBzY2hlZHVsZWQgPSB0cnVlO1xuICAgIHJlcXVlc3RBbmltYXRpb25GcmFtZSgoKSA9PiB7XG4gICAgICBzY2hlZHVsZWQgPSBmYWxzZTtcbiAgICAgIHRhZ0VsZW1lbnRzKCk7XG4gICAgICBzeW5jQ29sb3JTY2hlbWUoKTtcbiAgICAgIHN5bmNSZWFkaW5nUGFuZSgpO1xuICAgICAgaWYgKGxvY2F0aW9uLmhyZWYgIT09IGxhc3RVcmwpIHtcbiAgICAgICAgbGFzdFVybCA9IGxvY2F0aW9uLmhyZWY7XG4gICAgICAgIHNjaGVkdWxlSGVhbHRoQ2hlY2soKTtcbiAgICAgIH1cbiAgICB9KTtcbiAgfSkub2JzZXJ2ZShkb2N1bWVudC5ib2R5LCB7IGNoaWxkTGlzdDogdHJ1ZSwgc3VidHJlZTogdHJ1ZSB9KTtcbn1cblxubGV0IG9ycGhhbk5vdGljZVNob3duID0gZmFsc2U7XG5cbi8qKlxuICogQWZ0ZXIgT0JCIGlzIHJlbG9hZGVkIG9yIHVwZGF0ZWQsIHRoaXMgY29weSBvZiB0aGUgY29udGVudCBzY3JpcHQga2VlcHMgcnVubmluZyBpbiB0YWJzIHRoYXRcbiAqIHdlcmUgYWxyZWFkeSBvcGVuLCBjdXQgb2ZmIGZyb20gdGhlIGV4dGVuc2lvbi4gU2F5IHNvIG9uY2UgaW5zdGVhZCBvZiBoYWxmLXdvcmtpbmcuXG4gKi9cbmZ1bmN0aW9uIGlzT3JwaGFuZWQoKTogYm9vbGVhbiB7XG4gIGlmIChjaHJvbWUucnVudGltZT8uaWQpIHJldHVybiBmYWxzZTtcbiAgaWYgKCFvcnBoYW5Ob3RpY2VTaG93bikge1xuICAgIG9ycGhhbk5vdGljZVNob3duID0gdHJ1ZTtcbiAgICB0b2FzdChcIk9CQiB3YXMgdXBkYXRlZC4gUmVsb2FkIHRoaXMgdGFiIHRvIHVzZSB0aGUgbmV3IHZlcnNpb24uXCIsIDgwMDApO1xuICB9XG4gIHJldHVybiB0cnVlO1xufVxuXG5mdW5jdGlvbiBpc0VkaXRhYmxlKHRhcmdldDogRXZlbnRUYXJnZXQgfCB1bmRlZmluZWQpOiBib29sZWFuIHtcbiAgaWYgKCEodGFyZ2V0IGluc3RhbmNlb2YgSFRNTEVsZW1lbnQpKSByZXR1cm4gZmFsc2U7XG4gIHJldHVybiB0YXJnZXQuaXNDb250ZW50RWRpdGFibGUgfHwgL14oSU5QVVR8VEVYVEFSRUF8U0VMRUNUKSQvLnRlc3QodGFyZ2V0LnRhZ05hbWUpO1xufVxuXG5mdW5jdGlvbiBtb3VudEtleW1hcCgpOiB2b2lkIHtcbiAgY29uc3QgZW5naW5lID0gbmV3IFNlcXVlbmNlRW5naW5lKEJJTkRJTkdTKTtcbiAgbGV0IGhpbnRUaW1lcjogUmV0dXJuVHlwZTx0eXBlb2Ygc2V0VGltZW91dD4gfCB1bmRlZmluZWQ7XG4gIHdpbmRvdy5hZGRFdmVudExpc3RlbmVyKFxuICAgIFwia2V5ZG93blwiLFxuICAgIChldmVudCkgPT4ge1xuICAgICAgaWYgKCFzZXR0aW5ncy5rZXltYXAgfHwgIWV2ZW50LmlzVHJ1c3RlZCB8fCBpc0Zyb21PdmVybGF5KGV2ZW50KSB8fCBpc01vZGFsT3BlbigpKSByZXR1cm47XG4gICAgICBpZiAoaXNPcnBoYW5lZCgpKSByZXR1cm47XG4gICAgICBpZiAoaXNFZGl0YWJsZShldmVudC5jb21wb3NlZFBhdGgoKVswXSkpIHJldHVybjtcbiAgICAgIGNvbnN0IGtleSA9IGtleUZyb21FdmVudChldmVudCk7XG4gICAgICBpZiAoa2V5ID09PSBudWxsKSByZXR1cm47XG5cbiAgICAgIGNvbnN0IHJlc3VsdCA9IGVuZ2luZS5mZWVkKGtleSwgY3VycmVudENvbnRleHQoKSwgZXZlbnQudGltZVN0YW1wKTtcbiAgICAgIGlmIChyZXN1bHQua2luZCA9PT0gXCJub25lXCIpIHJldHVybjtcbiAgICAgIHNob3dIaW50KFxuICAgICAgICByZXN1bHQua2luZCA9PT0gXCJwZW5kaW5nXCJcbiAgICAgICAgICA/IGAke3Jlc3VsdC5wcmVmaXguam9pbihcIiBcIil9IFx1MjAyNiAke3Jlc3VsdC5uZXh0XG4gICAgICAgICAgICAgIC5tYXAoKGIpID0+IGAke2Iua2V5cy5hdCgtMSl9ICR7QUNUSU9OU1tiLmFjdGlvbl0uZGVzY3JpcHRpb259YClcbiAgICAgICAgICAgICAgLmpvaW4oXCIgXHUwMEI3IFwiKX1gXG4gICAgICAgICAgOiBudWxsLFxuICAgICAgKTtcbiAgICAgIGNsZWFyVGltZW91dChoaW50VGltZXIpO1xuICAgICAgaWYgKHJlc3VsdC5raW5kID09PSBcInBlbmRpbmdcIilcbiAgICAgICAgaGludFRpbWVyID0gc2V0VGltZW91dCgoKSA9PiBzaG93SGludChudWxsKSwgREVGQVVMVF9USU1FT1VUX01TKTtcbiAgICAgIGlmIChyZXN1bHQua2luZCA9PT0gXCJtYXRjaFwiICYmICFleGVjdXRlKHJlc3VsdC5hY3Rpb24pKSByZXR1cm47XG4gICAgICBldmVudC5wcmV2ZW50RGVmYXVsdCgpO1xuICAgICAgZXZlbnQuc3RvcEltbWVkaWF0ZVByb3BhZ2F0aW9uKCk7XG4gICAgfSxcbiAgICB0cnVlLFxuICApO1xufVxuXG5hc3luYyBmdW5jdGlvbiBtYWluKCk6IFByb21pc2U8dm9pZD4ge1xuICBhcHBseVNldHRpbmdzKGF3YWl0IGxvYWRTZXR0aW5ncygpKTtcbiAgb25TZXR0aW5nc0NoYW5nZWQoYXBwbHlTZXR0aW5ncyk7XG4gIG1vdW50S2V5bWFwKCk7XG4gIGF3YWl0IHdoZW5BcHBSZWFkeSgpO1xuICB0YWdFbGVtZW50cygpO1xuICBzeW5jUmVhZGluZ1BhbmUoKTtcbiAgc2NoZWR1bGVIZWFsdGhDaGVjaygpO1xuICB3YXRjaFBhZ2UoKTtcbn1cblxudm9pZCBtYWluKCk7XG4iXSwKICAibWFwcGluZ3MiOiAiOzs7QUFnQ0EsTUFBTSxjQUF3QixFQUFFLE1BQU0sY0FBYztBQUNwRCxNQUFNLE9BQWlCLEVBQUUsTUFBTSxPQUFPO0FBQ3RDLE1BQU0sUUFBUSxDQUFDLFlBQW9DLEVBQUUsTUFBTSxTQUFTLE9BQU87QUFDM0UsTUFB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QUFFLFdBQVcsTUFBTSxTQUFTLEtBQUssQ0FBQztBQUFBLEVBQzlEO0FBRUEsTUFBSSxvQkFBb0I7QUFNeEIsV0FBUyxhQUFzQjtBQUM3QixRQUFJLE9BQU8sU0FBUyxHQUFJLFFBQU87QUFDL0IsUUFBSSxDQUFDLG1CQUFtQjtBQUN0QiwwQkFBb0I7QUFDcEIsWUFBTSw0REFBNEQsR0FBSTtBQUFBLElBQ3hFO0FBQ0EsV0FBTztBQUFBLEVBQ1Q7QUFFQSxXQUFTLFdBQVcsUUFBMEM7QUFDNUQsUUFBSSxFQUFFLGtCQUFrQixhQUFjLFFBQU87QUFDN0MsV0FBTyxPQUFPLHFCQUFxQiw0QkFBNEIsS0FBSyxPQUFPLE9BQU87QUFBQSxFQUNwRjtBQUVBLFdBQVMsY0FBb0I7QUFDM0IsVUFBTSxTQUFTLElBQUksZUFBZSxRQUFRO0FBQzFDLFFBQUk7QUFDSixXQUFPO0FBQUEsTUFDTDtBQUFBLE1BQ0EsQ0FBQyxVQUFVO0FBQ1QsWUFBSSxDQUFDLFNBQVMsVUFBVSxDQUFDLE1BQU0sYUFBYSxjQUFjLEtBQUssS0FBSyxZQUFZLEVBQUc7QUFDbkYsWUFBSSxXQUFXLEVBQUc7QUFDbEIsWUFBSSxXQUFXLE1BQU0sYUFBYSxFQUFFLENBQUMsQ0FBQyxFQUFHO0FBQ3pDLGNBQU0sTUFBTSxhQUFhLEtBQUs7QUFDOUIsWUFBSSxRQUFRLEtBQU07QUFFbEIsY0FBTSxTQUFTLE9BQU8sS0FBSyxLQUFLLGVBQWUsR0FBRyxNQUFNLFNBQVM7QUFDakUsWUFBSSxPQUFPLFNBQVMsT0FBUTtBQUM1QjtBQUFBLFVBQ0UsT0FBTyxTQUFTLFlBQ1osR0FBRyxPQUFPLE9BQU8sS0FBSyxHQUFHLENBQUMsV0FBTSxPQUFPLEtBQ3BDLElBQUksQ0FBQyxNQUFNLEdBQUcsRUFBRSxLQUFLLEdBQUcsRUFBRSxDQUFDLElBQUksUUFBUSxFQUFFLE1BQU0sRUFBRSxXQUFXLEVBQUUsRUFDOUQsS0FBSyxRQUFLLENBQUMsS0FDZDtBQUFBLFFBQ047QUFDQSxxQkFBYSxTQUFTO0FBQ3RCLFlBQUksT0FBTyxTQUFTO0FBQ2xCLHNCQUFZLFdBQVcsTUFBTSxTQUFTLElBQUksR0FBRyxrQkFBa0I7QUFDakUsWUFBSSxPQUFPLFNBQVMsV0FBVyxDQUFDLFFBQVEsT0FBTyxNQUFNLEVBQUc7QUFDeEQsY0FBTSxlQUFlO0FBQ3JCLGNBQU0seUJBQXlCO0FBQUEsTUFDakM7QUFBQSxNQUNBO0FBQUEsSUFDRjtBQUFBLEVBQ0Y7QUFFQSxpQkFBZSxPQUFzQjtBQUNuQyxrQkFBYyxNQUFNLGFBQWEsQ0FBQztBQUNsQyxzQkFBa0IsYUFBYTtBQUMvQixnQkFBWTtBQUNaLFVBQU0sYUFBYTtBQUNuQixnQkFBWTtBQUNaLG9CQUFnQjtBQUNoQix3QkFBb0I7QUFDcEIsY0FBVTtBQUFBLEVBQ1o7QUFFQSxPQUFLLEtBQUs7IiwKICAibmFtZXMiOiBbImVsIiwgInNldHRpbmdzIl0KfQo=
