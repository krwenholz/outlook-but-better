"use strict";
(() => {
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
  function rowCheckbox(row) {
    return row.querySelector('[role="checkbox"]');
  }
  function checkedRows(root = document) {
    return resolveAll("messageRow", root).filter(
      (row) => rowCheckbox(row)?.getAttribute("aria-checked") === "true"
    );
  }
  function cursorRow(root = document) {
    const rows = resolveAll("messageRow", root);
    const active = root.activeElement;
    return rows.find((row) => active && row.contains(active)) ?? rows.find((row) => row.getAttribute("tabindex") === "0") ?? selectedRows(root)[0] ?? null;
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

  // src/capture.ts
  var CAPTURE_MESSAGE = "obb-capture";
  var OBB_ATTRIBUTE = /^data-overlook/;
  var RUNTIME_STYLE_MARKER = "data-overlook-runtime";
  function runtimeCss(doc) {
    const sheets = [...doc.styleSheets, ...doc.adoptedStyleSheets ?? []].filter(
      (sheet) => !("href" in sheet && sheet.href)
    );
    return sheets.flatMap((sheet) => {
      try {
        return [...sheet.cssRules].map((rule) => rule.cssText);
      } catch {
        return [];
      }
    }).filter((rule) => !rule.includes("data-overlook")).join("\n");
  }
  function captureFilename(now = /* @__PURE__ */ new Date()) {
    const stamp = now.toISOString().slice(0, 16).replace(/\D/g, "");
    return `obb-capture-${stamp}.html`;
  }
  function capturePage(doc = document, loc = location, version = "dev", now = /* @__PURE__ */ new Date()) {
    const live = doc.documentElement.dataset;
    const results = healthCheck(doc);
    const meta = {
      capturedAt: now.toISOString(),
      obbVersion: version,
      route: loc.pathname.split("/").slice(0, 3).join("/"),
      viewport: `${doc.defaultView?.innerWidth ?? 0}x${doc.defaultView?.innerHeight ?? 0}`,
      settings: {
        features: live.overlookFeatures,
        palette: live.overlookPalette,
        mode: live.overlookMode,
        layout: live.overlookLayout,
        outlookScheme: live.overlookOutlookScheme,
        readingPane: live.overlookReadingPane
      },
      health: {
        broken: brokenEntries(results),
        entries: results.map(({ name, candidate, count }) => ({ name, candidate, count }))
      }
    };
    const root = doc.documentElement.cloneNode(true);
    for (const el2 of root.querySelectorAll("overlook-root")) el2.remove();
    for (const el2 of [root, ...root.querySelectorAll("*")]) {
      for (const { name } of [...el2.attributes]) {
        if (OBB_ATTRIBUTE.test(name)) el2.removeAttribute(name);
      }
    }
    const head = root.querySelector("head");
    const style = doc.createElement("style");
    style.setAttribute(RUNTIME_STYLE_MARKER, "");
    style.textContent = runtimeCss(doc);
    head?.append(style);
    const script = doc.createElement("script");
    script.type = "application/json";
    script.id = "obb-capture-meta";
    script.textContent = JSON.stringify(meta, null, 2).replaceAll("<", "\\u003c");
    head?.append(script);
    return { html: `<!doctype html>
${root.outerHTML}`, name: captureFilename(now), meta };
  }

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
    next: { description: "Next conversation", strategy: extension("cursor-next") },
    previous: { description: "Previous conversation", strategy: extension("cursor-previous") },
    open: { description: "Open conversation", strategy: passthrough },
    select: { description: "Select or deselect", strategy: extension("toggle-select") },
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

  // src/ui/overlay.css
  var overlay_default = ':host {\n  all: initial;\n  --ol-bg: #fff;\n  --ol-fg: #1f1f1f;\n  --ol-muted: #6b6b6b;\n  --ol-accent: #2f6fde;\n  --ol-border: rgb(0 0 0 / 12%);\n  --ol-shadow: 0 8px 32px rgb(0 0 0 / 18%);\n  font:\n    13px / 1.4 system-ui,\n    -apple-system,\n    "Segoe UI",\n    sans-serif;\n}\n\n@media (prefers-color-scheme: dark) {\n  :host {\n    --ol-bg: #262626;\n    --ol-fg: #ececec;\n    --ol-muted: #a0a0a0;\n    --ol-accent: #7aa7ff;\n    --ol-border: rgb(255 255 255 / 14%);\n  }\n}\n\n[hidden] {\n  display: none !important;\n}\n\n.toast,\n.hint,\n.health {\n  position: fixed;\n  z-index: 2147483647;\n  background: var(--ol-bg);\n  color: var(--ol-fg);\n  border: 1px solid var(--ol-border);\n  border-radius: 6px;\n  box-shadow: var(--ol-shadow);\n  padding: 6px 10px;\n}\n\n.toast {\n  bottom: 24px;\n  left: 50%;\n  transform: translateX(-50%);\n}\n\n.toast.notice {\n  display: flex;\n  align-items: center;\n  gap: 12px;\n  padding: 8px 10px 8px 14px;\n}\n\n.action {\n  all: unset;\n  cursor: pointer;\n  padding: 4px 12px;\n  border-radius: 999px;\n  background: var(--ol-accent);\n  color: #fff;\n  font-weight: 600;\n}\n\n.action:focus-visible,\n.close:focus-visible {\n  outline: 2px solid var(--ol-accent);\n  outline-offset: 2px;\n}\n\n.hint {\n  bottom: 24px;\n  right: 24px;\n  font-family: ui-monospace, "SF Mono", Menlo, monospace;\n}\n\n.health {\n  top: 8px;\n  right: 8px;\n  display: flex;\n  gap: 8px;\n  align-items: center;\n  border-color: #d97706;\n}\n\n.close {\n  all: unset;\n  cursor: pointer;\n  color: var(--ol-muted);\n  padding: 0 2px;\n}\n\n.backdrop {\n  position: fixed;\n  inset: 0;\n  z-index: 2147483646;\n  background: rgb(0 0 0 / 25%);\n  display: flex;\n  justify-content: center;\n  align-items: flex-start;\n  padding-top: 12vh;\n}\n\n.panel {\n  background: var(--ol-bg);\n  color: var(--ol-fg);\n  border-radius: 10px;\n  box-shadow: var(--ol-shadow);\n  width: min(520px, 90vw);\n  max-height: 70vh;\n  overflow: auto;\n}\n\n.help {\n  padding: 16px 20px;\n}\n\n.help h2 {\n  font-size: 15px;\n  margin: 0 0 12px;\n}\n\n.help table {\n  border-collapse: collapse;\n  width: 100%;\n}\n\n.help td {\n  padding: 3px 0;\n}\n\n.keys {\n  width: 7em;\n}\n\nkbd {\n  font-family: ui-monospace, "SF Mono", Menlo, monospace;\n  border: 1px solid var(--ol-border);\n  border-radius: 4px;\n  padding: 0 5px;\n}\n\n.palette input {\n  all: unset;\n  box-sizing: border-box;\n  width: 100%;\n  padding: 12px 16px;\n  font-size: 15px;\n  border-bottom: 1px solid var(--ol-border);\n}\n\n.palette ul {\n  list-style: none;\n  margin: 0;\n  padding: 4px 0;\n}\n\n.palette li {\n  padding: 5px 16px;\n  cursor: pointer;\n}\n\n.palette li[aria-selected="true"] {\n  background: color-mix(in srgb, var(--ol-accent) 16%, transparent);\n}\n\n.palette li .path {\n  color: var(--ol-muted);\n  margin-left: 6px;\n}\n\n.palette mark {\n  background: none;\n  color: var(--ol-accent);\n  font-weight: 600;\n}\n\n.palette .empty {\n  color: var(--ol-muted);\n  cursor: default;\n}\n';

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
    node.classList.remove("notice");
    node.textContent = message;
    node.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => {
      node.hidden = true;
    }, ms);
  }
  function notice(message, action, onAction) {
    const node = singleton("toast");
    clearTimeout(toastTimer);
    const button = el("button", "action", action);
    button.addEventListener("click", onAction);
    const close = el("button", "close", "\xD7");
    close.setAttribute("aria-label", "Dismiss");
    close.addEventListener("click", () => {
      node.hidden = true;
    });
    node.replaceChildren(el("span", "", message), button, close);
    node.classList.add("notice");
    node.hidden = false;
    button.focus();
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
  function moveCursor(step) {
    if (checkedRows().length === 0) return false;
    const rows = resolveAll("messageRow");
    const at = cursorRow();
    const next = rows[(at ? rows.indexOf(at) : -1) + step];
    if (next instanceof HTMLElement) {
      next.focus();
      next.scrollIntoView?.({ block: "nearest" });
    }
    return true;
  }
  function toggleSelect() {
    const row = cursorRow();
    const box = row && rowCheckbox(row);
    if (!box) return false;
    box.click();
    return true;
  }
  function runExtension(command, options) {
    switch (command) {
      case "help":
        showHelp(
          BINDINGS.map((b) => ({
            keys: b.keys.join(" "),
            description: ACTIONS[b.action].description
          }))
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
  function execute(id, options = { multiSelect: true }) {
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

  // src/settings.ts
  var FEATURES = {
    healthIndicator: {
      label: "Show an indicator when Outlook's page changes break OBB",
      group: "General",
      default: true
    },
    keymap: { label: "Fastmail-style keyboard shortcuts", group: "Keyboard", default: true },
    multiSelect: {
      label: "x ticks a message; j and k move without losing ticks",
      group: "Keyboard",
      default: true
    },
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
      notice("OBB was updated.", "Reload tab", () => location.reload());
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
        if (result.kind === "match" && !execute(result.action, { multiSelect: settings.multiSelect }))
          return;
        event.preventDefault();
        event.stopImmediatePropagation();
      },
      true
    );
  }
  function serveCaptureRequests() {
    chrome.runtime.onMessage.addListener((message, _sender, reply) => {
      if (message?.type !== CAPTURE_MESSAGE) return;
      const { html, name } = capturePage(document, location, chrome.runtime.getManifest().version);
      const url = URL.createObjectURL(new Blob([html], { type: "text/html" }));
      const link = document.createElement("a");
      link.href = url;
      link.download = name;
      document.body.append(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1e4);
      reply({ ok: true, name, bytes: html.length });
    });
  }
  async function main() {
    serveCaptureRequests();
    applySettings(await loadSettings());
    onSettingsChanged(applySettings);
    mountKeymap();
    document.addEventListener("visibilitychange", () => {
      if (!document.hidden) isOrphaned();
    });
    await whenAppReady();
    tagElements();
    syncReadingPane();
    scheduleHealthCheck();
    watchPage();
  }
  void main();
})();
//# sourceMappingURL=data:application/json;base64,ewogICJ2ZXJzaW9uIjogMywKICAic291cmNlcyI6IFsiLi4vc3JjL3NlbGVjdG9ycy50cyIsICIuLi9zcmMvY2FwdHVyZS50cyIsICIuLi9zcmMva2V5bWFwL2FjdGlvbnMudHMiLCAiLi4vc3JjL2tleW1hcC9iaW5kaW5ncy50cyIsICIuLi9zcmMva2V5bWFwL2VuZ2luZS50cyIsICIuLi9zcmMvdWkvb3ZlcmxheS5jc3MiLCAiLi4vc3JjL3VpL292ZXJsYXkudHMiLCAiLi4vc3JjL3BhbGV0dGUvZnV6enkudHMiLCAiLi4vc3JjL3BhbGV0dGUvcGFsZXR0ZS50cyIsICIuLi9zcmMva2V5bWFwL2V4ZWN1dGUudHMiLCAiLi4vc3JjL3NldHRpbmdzLnRzIiwgIi4uL3NyYy9jb250ZW50LnRzIl0sCiAgInNvdXJjZXNDb250ZW50IjogWyIvLyBUaGUgb25seSBmaWxlIGFsbG93ZWQgdG8ga25vdyB3aGF0IE91dGxvb2sncyBET00gbG9va3MgbGlrZS5cbi8vXG4vLyBBbmNob3IgcHJlZmVyZW5jZSAoc2VlIFBMQU4ubWQpOiBzdHJ1Y3R1cmFsIElEcyA+IEFSSUEgcm9sZXMvbGFiZWxzID4gc2VtYW50aWMgZGF0YS0qID4gdGV4dC5cbi8vIE5ldmVyIHVzZSBnZW5lcmF0ZWQgY2xhc3MgbmFtZXM7IGlmIG9uZSBpcyB1bmF2b2lkYWJsZSwgbWFyayBpdCBgLy8gRlJBR0lMRTpgLlxuLy9cbi8vIFZlcmlmaWVkIGFnYWluc3QgYSBzYXZlZCBwYWdlIGZyb20gT2N0IDIwMjYgKHRlc3QvZml4dHVyZXMvaW5ib3guaHRtbCk7IHNlZSBkb2NzL2RvbS1zdXJ2ZXkubWQuXG4vLyBFbnRyaWVzIG1hcmtlZCBVTlZFUklGSUVEIHdlcmUgbm90IHByZXNlbnQgaW4gdGhhdCBjYXB0dXJlLiBhcmlhLWxhYmVscyBhc3N1bWUgYW4gRW5nbGlzaCBVSS5cblxuZXhwb3J0IHR5cGUgU2VsZWN0b3JOYW1lID0ga2V5b2YgdHlwZW9mIFNFTEVDVE9SUztcblxuZXhwb3J0IGludGVyZmFjZSBTZWxlY3RvckVudHJ5IHtcbiAgLyoqIENhbmRpZGF0ZSBDU1Mgc2VsZWN0b3JzLCBtb3N0IHN0YWJsZSBmaXJzdC4gVGhlIGZpcnN0IG9uZSB0aGF0IG1hdGNoZXMgd2lucy4gKi9cbiAgY2FuZGlkYXRlczogcmVhZG9ubHkgc3RyaW5nW107XG4gIC8qKiBcImFsd2F5c1wiOiBzaG91bGQgZXhpc3Qgb24gZXZlcnkgbWFpbCB2aWV3LCBzbyBhIG1pc3MgaXMgYSBicmVha2FnZS4gXCJzb21ldGltZXNcIjogY29udGV4dHVhbC4gKi9cbiAgZXhwZWN0OiBcImFsd2F5c1wiIHwgXCJzb21ldGltZXNcIjtcbiAgLyoqIE5hbWUgb2YgYW5vdGhlciBlbnRyeSB0byByZXNvbHZlIGluc2lkZSwgaW5zdGVhZCBvZiB0aGUgd2hvbGUgZG9jdW1lbnQuICovXG4gIHdpdGhpbj86IHN0cmluZztcbiAgLyoqIFRva2VuIGFkZGVkIHRvIHRoZSBlbGVtZW50J3MgYGRhdGEtb3Zlcmxvb2tgIGF0dHJpYnV0ZSBzbyBDU1MgY2FuIHRhcmdldCBpdC4gKi9cbiAgdGFnPzogc3RyaW5nO1xuICAvKiogVGFnIGV2ZXJ5IG1hdGNoIHJhdGhlciB0aGFuIG9ubHkgdGhlIGZpcnN0LiAqL1xuICBhbGw/OiBib29sZWFuO1xuICAvKiogTWF0Y2ggZXZlcnkgY2FuZGlkYXRlIChhIGxpc3Qgb2YgZGlzdGluY3QgdGhpbmdzKSwgbm90IGp1c3QgdGhlIGZpcnN0IHRoYXQgbWF0Y2hlcy4gKi9cbiAgdW5pb24/OiBib29sZWFuO1xufVxuXG5leHBvcnQgY29uc3QgU0VMRUNUT1JTID0ge1xuICAvLyAtLS0gU3RydWN0dXJlIC0tLVxuICB0aXRsZUJhcjogeyBjYW5kaWRhdGVzOiBbXCIjT3dhVGl0bGVCYXJcIl0sIGV4cGVjdDogXCJhbHdheXNcIiwgdGFnOiBcInRpdGxlLWJhclwiIH0sXG4gIHNlYXJjaEJveDogeyBjYW5kaWRhdGVzOiBbXCIjdG9wU2VhcmNoSW5wdXRcIiwgJ1tyb2xlPVwic2VhcmNoXCJdIGlucHV0J10sIGV4cGVjdDogXCJhbHdheXNcIiB9LFxuICBhcHBSYWlsOiB7XG4gICAgY2FuZGlkYXRlczogW1wiI0xlZnRSYWlsXCIsICdbcm9sZT1cIm5hdmlnYXRpb25cIl1bYXJpYS1sYWJlbD1cImxlZnQtcmFpbC1hcHBiYXJcIl0nXSxcbiAgICBleHBlY3Q6IFwiYWx3YXlzXCIsXG4gICAgdGFnOiBcImFwcC1yYWlsXCIsXG4gIH0sXG4gIHJpYmJvbjoge1xuICAgIGNhbmRpZGF0ZXM6IFsnW2RhdGEtYXBwLXNlY3Rpb249XCJSaWJib25cIl0nLCAnW3JvbGU9XCJyZWdpb25cIl1bYXJpYS1sYWJlbD1cIlJpYmJvblwiXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB0YWc6IFwicmliYm9uXCIsXG4gIH0sXG4gIHJpYmJvblRhYnM6IHtcbiAgICBjYW5kaWRhdGVzOiBbJ1tyb2xlPVwidGFibGlzdFwiXVthcmlhLWxhYmVsPVwiUmliYm9uIHRhYnNcIl0nXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgd2l0aGluOiBcInJpYmJvblwiLFxuICAgIHRhZzogXCJyaWJib24tdGFic1wiLFxuICB9LFxuICBmb2xkZXJQYW5lOiB7XG4gICAgY2FuZGlkYXRlczogW1xuICAgICAgJ1tkYXRhLWFwcC1zZWN0aW9uPVwiTmF2aWdhdGlvblBhbmVcIl0nLFxuICAgICAgJ1tyb2xlPVwicmVnaW9uXCJdW2FyaWEtbGFiZWw9XCJOYXZpZ2F0aW9uIHBhbmVcIl0nLFxuICAgIF0sXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICAgIHRhZzogXCJmb2xkZXItcGFuZVwiLFxuICB9LFxuICBmb2xkZXJJdGVtOiB7XG4gICAgY2FuZGlkYXRlczogWydbcm9sZT1cInRyZWVpdGVtXCJdW2RhdGEtZm9sZGVyLW5hbWVdJywgJ1tyb2xlPVwidHJlZWl0ZW1cIl0nXSxcbiAgICBleHBlY3Q6IFwiYWx3YXlzXCIsXG4gICAgd2l0aGluOiBcImZvbGRlclBhbmVcIixcbiAgICB0YWc6IFwiZm9sZGVyLWl0ZW1cIixcbiAgICBhbGw6IHRydWUsXG4gIH0sXG4gIGxpc3RQYW5lOiB7XG4gICAgY2FuZGlkYXRlczogWydbZGF0YS1hcHAtc2VjdGlvbj1cIk1lc3NhZ2VMaXN0XCJdJ10sXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICAgIHRhZzogXCJsaXN0LXBhbmVcIixcbiAgfSxcbiAgbWVzc2FnZUxpc3Q6IHtcbiAgICBjYW5kaWRhdGVzOiBbXCIjTWFpbExpc3RcIiwgJ1tyb2xlPVwibGlzdGJveFwiXVthcmlhLWxhYmVsPVwiTWVzc2FnZSBsaXN0XCJdJ10sXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICAgIHRhZzogXCJtZXNzYWdlLWxpc3RcIixcbiAgfSxcbiAgbWVzc2FnZVJvdzoge1xuICAgIGNhbmRpZGF0ZXM6IFsnW3JvbGU9XCJvcHRpb25cIl1bZGF0YS1jb252aWRdJywgJ1tyb2xlPVwib3B0aW9uXCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJtZXNzYWdlTGlzdFwiLFxuICAgIHRhZzogXCJtZXNzYWdlLXJvd1wiLFxuICAgIGFsbDogdHJ1ZSxcbiAgfSxcbiAgZ3JvdXBIZWFkZXI6IHtcbiAgICBjYW5kaWRhdGVzOiBbJ1tyb2xlPVwiYnV0dG9uXCJdW2lkXj1cImdyb3VwSGVhZGVyXCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJtZXNzYWdlTGlzdFwiLFxuICAgIHRhZzogXCJncm91cC1oZWFkZXJcIixcbiAgICBhbGw6IHRydWUsXG4gIH0sXG4gIGZvY3VzZWRUYWJzOiB7XG4gICAgY2FuZGlkYXRlczogWydbcm9sZT1cInRhYmxpc3RcIl0nXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgd2l0aGluOiBcImxpc3RQYW5lXCIsXG4gICAgdGFnOiBcImZvY3VzZWQtdGFic1wiLFxuICB9LFxuICByZWFkaW5nUGFuZToge1xuICAgIGNhbmRpZGF0ZXM6IFtcbiAgICAgICdbZGF0YS1hcHAtc2VjdGlvbj1cIk1haWxSZWFkQ29tcG9zZVwiXScsXG4gICAgICBcIiNSZWFkaW5nUGFuZUNvbnRhaW5lcklkXCIsXG4gICAgICAnW3JvbGU9XCJtYWluXCJdW2FyaWEtbGFiZWw9XCJSZWFkaW5nIFBhbmVcIl0nLFxuICAgIF0sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHRhZzogXCJyZWFkaW5nLXBhbmVcIixcbiAgfSxcblxuICBtZXNzYWdlQm9keToge1xuICAgIGNhbmRpZGF0ZXM6IFsnW3JvbGU9XCJkb2N1bWVudFwiXVthcmlhLWxhYmVsPVwiTWVzc2FnZSBib2R5XCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJyZWFkaW5nUGFuZVwiLFxuICAgIHRhZzogXCJtZXNzYWdlLWJvZHlcIixcbiAgICBhbGw6IHRydWUsXG4gIH0sXG5cbiAgaW5mb0Jhcjoge1xuICAgIC8vIE5vdGljZXMgYWJvdmUgYSBtZXNzYWdlIChcIllvdSByZXBsaWVkIG9uXHUyMDI2XCIpOyB0aGVpciB0ZXh0IHNpdHMgaW4gYW4gZWxlbWVudCB3aG9zZSBpZCBzdGFydHNcbiAgICAvLyB3aXRoIFwiaW5mb0JhclwiLiBPdXRsb29rIHBhaW50cyB0aGVtIHdpdGggaW5saW5lIGNvbG91cnMuXG4gICAgY2FuZGlkYXRlczogWydbcm9sZT1cImdyb3VwXCJdOmhhcyg+ICogPiBbaWRePVwiaW5mb0JhclwiXSknXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgd2l0aGluOiBcInJlYWRpbmdQYW5lXCIsXG4gICAgdGFnOiBcImluZm8tYmFyXCIsXG4gICAgYWxsOiB0cnVlLFxuICB9LFxuXG4gIC8vIC0tLSBDbHV0dGVyIChlYWNoIGhpZGVhYmxlIGVudHJ5IGdldHMgYSB0b2dnbGUgaW4gc2V0dGluZ3MudHMpIC0tLVxuICB1cHNlbGxCYW5uZXI6IHtcbiAgICBjYW5kaWRhdGVzOiBbJ1tkYXRhLWFwcC1zZWN0aW9uPVwiVXBzZWxsQmFubmVyU2VjdGlvblwiXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB0YWc6IFwidXBzZWxsXCIsXG4gICAgYWxsOiB0cnVlLFxuICB9LFxuICBoZWFkZXJFeHRyYXM6IHtcbiAgICBjYW5kaWRhdGVzOiBbXCIjb3dhVGVhbXNCdG5fY29udGFpbmVyXCIsIFwiI293YVRpbWVQYW5lbEJ0bl9jb250YWluZXJcIl0sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHRhZzogXCJoZWFkZXItZXh0cmFcIixcbiAgICB1bmlvbjogdHJ1ZSxcbiAgfSxcbiAgYXBwUmFpbEV4dHJhczoge1xuICAgIGNhbmRpZGF0ZXM6IFtcIlRvIERvXCIsIFwiTmV3c2xldHRlcnNcIiwgXCJPcmcgRXhwbG9yZXJcIiwgXCJPbmVEcml2ZVwiXS5tYXAoXG4gICAgICAoYXBwKSA9PiBgW3JvbGU9XCJub25lXCJdW2FyaWEtbGFiZWw9XCIke2FwcH1cIl1gLFxuICAgICksXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJhcHBSYWlsXCIsXG4gICAgdGFnOiBcInJhaWwtZXh0cmFcIixcbiAgICB1bmlvbjogdHJ1ZSxcbiAgfSxcbiAgcmliYm9uRXh0cmFzOiB7XG4gICAgY2FuZGlkYXRlczogW1xuICAgICAgXCJTaGFyZSB0byBUZWFtc1wiLFxuICAgICAgXCJTd2VlcFwiLFxuICAgICAgXCJBc3NpZ24gcG9saWN5XCIsXG4gICAgICBcIlZpdmEgSW5zaWdodHNcIixcbiAgICAgIFwiTW9yZSBhcHBzXCIsXG4gICAgICBcIlF1aWNrIHN0ZXBzXCIsXG4gICAgICBcIk1vcmUgcXVpY2sgc3RlcHNcIixcbiAgICBdLm1hcCgobGFiZWwpID0+IGBbYXJpYS1sYWJlbD1cIiR7bGFiZWx9XCJdYCksXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJyaWJib25cIixcbiAgICB0YWc6IFwicmliYm9uLWV4dHJhXCIsXG4gICAgdW5pb246IHRydWUsXG4gIH0sXG4gIHJvd0F2YXRhcjoge1xuICAgIGNhbmRpZGF0ZXM6IFsnW3JvbGU9XCJpbWdcIl1baWRePVwiYXZhdGFyLVwiXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB3aXRoaW46IFwibWVzc2FnZUxpc3RcIixcbiAgICB0YWc6IFwicm93LWF2YXRhclwiLFxuICAgIGFsbDogdHJ1ZSxcbiAgfSxcbiAgcm93SG92ZXJBY3Rpb25zOiB7XG4gICAgY2FuZGlkYXRlczogW1xuICAgICAgXCJDdXN0b21pemUgQWN0aW9uc1wiLFxuICAgICAgXCJGbGFnIHRoaXMgbWVzc2FnZVwiLFxuICAgICAgXCJLZWVwIHRoaXMgbWVzc2FnZSBhdCB0aGUgdG9wIG9mIHlvdXIgZm9sZGVyXCIsXG4gICAgXS5tYXAoKGxhYmVsKSA9PiBgYnV0dG9uW2FyaWEtbGFiZWw9XCIke2xhYmVsfVwiXWApLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB3aXRoaW46IFwibWVzc2FnZUxpc3RcIixcbiAgICB0YWc6IFwicm93LWFjdGlvblwiLFxuICAgIHVuaW9uOiB0cnVlLFxuICB9LFxuICBhcHBCYW5uZXI6IHtcbiAgICAvLyBGbHVlbnQncyBzdGF0aWMgXCJmdWktTWVzc2FnZUJhclwiIGNsYXNzIGlzIHBhcnQgb2YgaXRzIHB1YmxpYyBzdHlsaW5nIEFQSSwgbm90IGdlbmVyYXRlZDtcbiAgICAvLyBub3RoaW5nIGVsc2UgbWFya3MgdGhlc2UuIFNjb3BlZCB0byBiYXJzIGFjcm9zcyB0aGUgdG9wIG9mIHRoZSBhcHAsIG5vdCBpbnNpZGUgbWVzc2FnZXMuXG4gICAgY2FuZGlkYXRlczogWycjbWFpbkFwcCA+IGRpdiA+IFtyb2xlPVwiZ3JvdXBcIl0uZnVpLU1lc3NhZ2VCYXInXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgdGFnOiBcImFwcC1iYW5uZXJcIixcbiAgICBhbGw6IHRydWUsXG4gIH0sXG4gIGNvcGlsb3Q6IHtcbiAgICAvLyBVTlZFUklGSUVEOiBubyBDb3BpbG90IHN1cmZhY2VzIGluIHRoZSBPY3QgMjAyNiBjYXB0dXJlLlxuICAgIGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWwqPVwiQ29waWxvdFwiIGldJywgJ1tyb2xlPVwicmVnaW9uXCJdW2FyaWEtbGFiZWwqPVwiQ29waWxvdFwiIGldJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHRhZzogXCJjb3BpbG90XCIsXG4gICAgdW5pb246IHRydWUsXG4gIH0sXG4gIGV4dGVybmFsVGFnOiB7XG4gICAgLy8gVU5WRVJJRklFRDogbm8gRXh0ZXJuYWwgdGFncyBpbiB0aGUgT2N0IDIwMjYgY2FwdHVyZS5cbiAgICBjYW5kaWRhdGVzOiBbJ1t0aXRsZT1cIkV4dGVybmFsIHNlbmRlclwiXScsICdbYXJpYS1sYWJlbD1cIkV4dGVybmFsXCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHRhZzogXCJleHRlcm5hbC10YWdcIixcbiAgICB1bmlvbjogdHJ1ZSxcbiAgfSxcblxuICAvLyAtLS0gQ29udHJvbHMgdGhlIGtleW1hcCBkcml2ZXMgLS0tXG4gIGNvbXBvc2VCdXR0b246IHsgY2FuZGlkYXRlczogWydidXR0b25bYXJpYS1sYWJlbD1cIk5ldyBtYWlsXCJdJ10sIGV4cGVjdDogXCJhbHdheXNcIiB9LFxuICBhcmNoaXZlQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJBcmNoaXZlXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICBkZWxldGVCdXR0b246IHsgY2FuZGlkYXRlczogWydidXR0b25bYXJpYS1sYWJlbD1cIkRlbGV0ZVwiXSddLCBleHBlY3Q6IFwic29tZXRpbWVzXCIgfSxcbiAgLy8gVU5WRVJJRklFRDogdGhpcyB0ZW5hbnQgc2hvd3MgYSBcIlBoaXNoIEFsZXJ0XCIgYWRkLWluIGluc3RlYWQgb2YgT3V0bG9vaydzIFJlcG9ydCBtZW51LlxuICBqdW5rQnV0dG9uOiB7XG4gICAgY2FuZGlkYXRlczogWydidXR0b25bYXJpYS1sYWJlbD1cIlJlcG9ydFwiXScsICdidXR0b25bYXJpYS1sYWJlbD1cIkp1bmtcIl0nXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gIH0sXG4gIG1vdmVCdXR0b246IHsgY2FuZGlkYXRlczogWydidXR0b25bYXJpYS1sYWJlbD1cIk1vdmUgdG9cIl0nXSwgZXhwZWN0OiBcInNvbWV0aW1lc1wiIH0sXG4gIHNub296ZUJ1dHRvbjogeyBjYW5kaWRhdGVzOiBbJ2J1dHRvblthcmlhLWxhYmVsPVwiU25vb3plXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICBmbGFnQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJGbGFnIC8gVW5mbGFnXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICByZWFkVW5yZWFkQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJSZWFkIC8gVW5yZWFkXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICB1bmRvQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJVbmRvXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICBzeW5jQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFtcIiNvd2FTeW5jQnV0dG9uX2NvbnRhaW5lclwiXSwgZXhwZWN0OiBcInNvbWV0aW1lc1wiIH0sXG4gIHNldHRpbmdzQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFtcIiNvd2FTZXR0aW5nc0J0bl9jb250YWluZXJcIl0sIGV4cGVjdDogXCJhbHdheXNcIiB9LFxuICAvLyBBcHAgcmFpbCBidXR0b25zLiBGYWxsYmFja3MgYXJlIHRoZSBNaWNyb3NvZnQgMzY1IGFwcCBpZHMgT3V0bG9vayBwdXRzIGluIGRhdGEtdGlkLCB0aGUgc2FtZVxuICAvLyBhY3Jvc3MgdGhlIE9jdCBhbmQgTm92IDIwMjYgY2FwdHVyZXMuIFRoZSByYWlsIHJlbmRlcnMgYWZ0ZXIgdGhlIG1lc3NhZ2UgbGlzdC5cbiAgbWFpbEFwcDoge1xuICAgIGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJNYWlsXCJdJywgJ1tkYXRhLXRpZD1cImRkZWE3NzRjLTM4MmItNDdkNy1hYWI1LWFkYzIxMzlhODAyYlwiXSddLFxuICAgIHdpdGhpbjogXCJhcHBSYWlsXCIsXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICB9LFxuICBjYWxlbmRhckFwcDoge1xuICAgIGNhbmRpZGF0ZXM6IFtcbiAgICAgICdidXR0b25bYXJpYS1sYWJlbD1cIkNhbGVuZGFyXCJdJyxcbiAgICAgICdbZGF0YS10aWQ9XCI4Y2JlYjg2Zi04M2UxLTQzYjUtYWFiYS1jZDM1MTQzMjJmMGJcIl0nLFxuICAgIF0sXG4gICAgd2l0aGluOiBcImFwcFJhaWxcIixcbiAgICBleHBlY3Q6IFwiYWx3YXlzXCIsXG4gIH0sXG4gIHBlb3BsZUFwcDoge1xuICAgIGNhbmRpZGF0ZXM6IFtcbiAgICAgICdidXR0b25bYXJpYS1sYWJlbD1cIlBlb3BsZVwiXScsXG4gICAgICAnW2RhdGEtdGlkPVwiMzU1ZmJkNzktM2JhMi00NTU0LThmMmQtMDMwMGZkZTkxZjMwXCJdJyxcbiAgICBdLFxuICAgIHdpdGhpbjogXCJhcHBSYWlsXCIsXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICB9LFxufSBhcyBjb25zdCBzYXRpc2ZpZXMgUmVjb3JkPHN0cmluZywgU2VsZWN0b3JFbnRyeT47XG5cbmNvbnN0IGVudHJ5ID0gKG5hbWU6IFNlbGVjdG9yTmFtZSk6IFNlbGVjdG9yRW50cnkgPT4gU0VMRUNUT1JTW25hbWVdO1xuXG5mdW5jdGlvbiBzY29wZShuYW1lOiBTZWxlY3Rvck5hbWUsIHJvb3Q6IFBhcmVudE5vZGUpOiBQYXJlbnROb2RlIHwgbnVsbCB7XG4gIGNvbnN0IHBhcmVudCA9IGVudHJ5KG5hbWUpLndpdGhpbjtcbiAgcmV0dXJuIHBhcmVudCA/IHJlc29sdmUocGFyZW50IGFzIFNlbGVjdG9yTmFtZSwgcm9vdCkgOiByb290O1xufVxuXG4vKiogVGhlIGZpcnN0IGVsZW1lbnQgbWF0Y2hpbmcgYG5hbWVgLCB0cnlpbmcgY2FuZGlkYXRlcyBpbiBwcmlvcml0eSBvcmRlci4gKi9cbmV4cG9ydCBmdW5jdGlvbiByZXNvbHZlKG5hbWU6IFNlbGVjdG9yTmFtZSwgcm9vdDogUGFyZW50Tm9kZSA9IGRvY3VtZW50KTogRWxlbWVudCB8IG51bGwge1xuICBjb25zdCB3aXRoaW4gPSBzY29wZShuYW1lLCByb290KTtcbiAgaWYgKCF3aXRoaW4pIHJldHVybiBudWxsO1xuICBmb3IgKGNvbnN0IGNhbmRpZGF0ZSBvZiBlbnRyeShuYW1lKS5jYW5kaWRhdGVzKSB7XG4gICAgY29uc3QgZWwgPSB3aXRoaW4ucXVlcnlTZWxlY3RvcihjYW5kaWRhdGUpO1xuICAgIGlmIChlbCkgcmV0dXJuIGVsO1xuICB9XG4gIHJldHVybiBudWxsO1xufVxuXG4vKiogRXZlcnkgbWF0Y2g6IGZvciB1bmlvbiBlbnRyaWVzLCBvZiBhbGwgY2FuZGlkYXRlczsgb3RoZXJ3aXNlIG9mIHRoZSBmaXJzdCBjYW5kaWRhdGUgdGhhdCBtYXRjaGVzLiAqL1xuZXhwb3J0IGZ1bmN0aW9uIHJlc29sdmVBbGwobmFtZTogU2VsZWN0b3JOYW1lLCByb290OiBQYXJlbnROb2RlID0gZG9jdW1lbnQpOiBFbGVtZW50W10ge1xuICBjb25zdCB3aXRoaW4gPSBzY29wZShuYW1lLCByb290KTtcbiAgaWYgKCF3aXRoaW4pIHJldHVybiBbXTtcbiAgY29uc3QgeyBjYW5kaWRhdGVzLCB1bmlvbiB9ID0gZW50cnkobmFtZSk7XG4gIGlmICh1bmlvbikgcmV0dXJuIFsuLi53aXRoaW4ucXVlcnlTZWxlY3RvckFsbChjYW5kaWRhdGVzLmpvaW4oXCIsIFwiKSldO1xuICBmb3IgKGNvbnN0IGNhbmRpZGF0ZSBvZiBjYW5kaWRhdGVzKSB7XG4gICAgY29uc3QgZWxzID0gd2l0aGluLnF1ZXJ5U2VsZWN0b3JBbGwoY2FuZGlkYXRlKTtcbiAgICBpZiAoZWxzLmxlbmd0aCA+IDApIHJldHVybiBbLi4uZWxzXTtcbiAgfVxuICByZXR1cm4gW107XG59XG5cbi8qKiBBZGQgYGRhdGEtb3Zlcmxvb2tgIHRva2VucyB0byBldmVyeSB0YWdnZWQgZW50cnkgc28gc3R5bGVzaGVldHMgbmV2ZXIgdG91Y2ggT3V0bG9vaydzIERPTSBzaGFwZS4gKi9cbmV4cG9ydCBmdW5jdGlvbiB0YWdFbGVtZW50cyhyb290OiBQYXJlbnROb2RlID0gZG9jdW1lbnQpOiB2b2lkIHtcbiAgZm9yIChjb25zdCBuYW1lIG9mIE9iamVjdC5rZXlzKFNFTEVDVE9SUykgYXMgU2VsZWN0b3JOYW1lW10pIHtcbiAgICBjb25zdCB7IHRhZywgYWxsLCB1bmlvbiB9ID0gZW50cnkobmFtZSk7XG4gICAgaWYgKCF0YWcpIGNvbnRpbnVlO1xuICAgIGNvbnN0IGVscyA9XG4gICAgICBhbGwgfHwgdW5pb24gPyByZXNvbHZlQWxsKG5hbWUsIHJvb3QpIDogW3Jlc29sdmUobmFtZSwgcm9vdCldLmZpbHRlcigoZSkgPT4gZSAhPT0gbnVsbCk7XG4gICAgZm9yIChjb25zdCBlbCBvZiBlbHMpIHtcbiAgICAgIGNvbnN0IHRva2VucyA9IG5ldyBTZXQoKGVsLmdldEF0dHJpYnV0ZShcImRhdGEtb3Zlcmxvb2tcIikgPz8gXCJcIikuc3BsaXQoXCIgXCIpLmZpbHRlcihCb29sZWFuKSk7XG4gICAgICBpZiAodG9rZW5zLmhhcyh0YWcpKSBjb250aW51ZTtcbiAgICAgIHRva2Vucy5hZGQodGFnKTtcbiAgICAgIGVsLnNldEF0dHJpYnV0ZShcImRhdGEtb3Zlcmxvb2tcIiwgWy4uLnRva2Vuc10uam9pbihcIiBcIikpO1xuICAgIH1cbiAgfVxufVxuXG5leHBvcnQgaW50ZXJmYWNlIEhlYWx0aFJlc3VsdCB7XG4gIG5hbWU6IFNlbGVjdG9yTmFtZTtcbiAgZXhwZWN0OiBTZWxlY3RvckVudHJ5W1wiZXhwZWN0XCJdO1xuICAvKiogSW5kZXggb2YgdGhlIGNhbmRpZGF0ZSB0aGF0IG1hdGNoZWQsIG9yIC0xLiBBIG5vbi16ZXJvIGluZGV4IG1lYW5zIGEgZmFsbGJhY2sgaXMgaW4gdXNlLiAqL1xuICBjYW5kaWRhdGU6IG51bWJlcjtcbiAgY291bnQ6IG51bWJlcjtcbn1cblxuZXhwb3J0IGZ1bmN0aW9uIGhlYWx0aENoZWNrKHJvb3Q6IFBhcmVudE5vZGUgPSBkb2N1bWVudCk6IEhlYWx0aFJlc3VsdFtdIHtcbiAgcmV0dXJuIChPYmplY3Qua2V5cyhTRUxFQ1RPUlMpIGFzIFNlbGVjdG9yTmFtZVtdKS5tYXAoKG5hbWUpID0+IHtcbiAgICBjb25zdCB7IGNhbmRpZGF0ZXMsIGV4cGVjdCwgdW5pb24gfSA9IGVudHJ5KG5hbWUpO1xuICAgIGNvbnN0IHdpdGhpbiA9IHNjb3BlKG5hbWUsIHJvb3QpO1xuICAgIGxldCBjYW5kaWRhdGUgPSAtMTtcbiAgICBsZXQgY291bnQgPSAwO1xuICAgIGlmICh3aXRoaW4pIHtcbiAgICAgIGNhbmRpZGF0ZSA9IGNhbmRpZGF0ZXMuZmluZEluZGV4KChjKSA9PiB3aXRoaW4ucXVlcnlTZWxlY3RvcihjKSAhPT0gbnVsbCk7XG4gICAgICBpZiAodW5pb24gJiYgY2FuZGlkYXRlID49IDApIGNvdW50ID0gcmVzb2x2ZUFsbChuYW1lLCByb290KS5sZW5ndGg7XG4gICAgICBlbHNlIGlmIChjYW5kaWRhdGUgPj0gMCkgY291bnQgPSB3aXRoaW4ucXVlcnlTZWxlY3RvckFsbChjYW5kaWRhdGVzW2NhbmRpZGF0ZV0gPz8gXCJcIikubGVuZ3RoO1xuICAgIH1cbiAgICAvLyBBIHVuaW9uIGVudHJ5IGxpc3RzIGRpc3RpbmN0IHRoaW5ncywgc28gbWF0Y2hpbmcgYSBsYXRlciBvbmUgaXMgbm90IGEgZmFsbGJhY2suXG4gICAgaWYgKHVuaW9uICYmIGNhbmRpZGF0ZSA+IDApIGNhbmRpZGF0ZSA9IDA7XG4gICAgcmV0dXJuIHsgbmFtZSwgZXhwZWN0LCBjYW5kaWRhdGUsIGNvdW50IH07XG4gIH0pO1xufVxuXG5leHBvcnQgaW50ZXJmYWNlIEZvbGRlck5vZGUge1xuICBuYW1lOiBzdHJpbmc7XG4gIC8qKiBBbmNlc3RvciBuYW1lcyBqb2luZWQgd2l0aCBcIi9cIiwgZS5nLiBcIkluYm94L1JlY2VpcHRzXCIuICovXG4gIHBhdGg6IHN0cmluZztcbiAgZWxlbWVudDogRWxlbWVudDtcbn1cblxuLyoqIFwiSW5ib3ggLSA5IGl0ZW1zICgxIHVucmVhZClcIiAtPiBcIkluYm94XCIuIFRoZSB0aXRsZSBhdHRyaWJ1dGUgY2FycmllcyBjb3VudHMuICovXG5mdW5jdGlvbiBmb2xkZXJOYW1lKGVsZW1lbnQ6IEVsZW1lbnQpOiBzdHJpbmcge1xuICBjb25zdCB0aXRsZSA9IChlbGVtZW50LmdldEF0dHJpYnV0ZShcInRpdGxlXCIpID8/IGVsZW1lbnQudGV4dENvbnRlbnQgPz8gXCJcIikudHJpbSgpO1xuICByZXR1cm4gdGl0bGUucmVwbGFjZSgvXFxzKy1cXHMrXFxkK1xccytpdGVtcz8oXFxzK1xcKFxcZCtcXHMrdW5yZWFkXFwpKT8kLywgXCJcIik7XG59XG5cbi8qKlxuICogRm9sZGVycyBjdXJyZW50bHkgcmVuZGVyZWQgaW4gdGhlIGZvbGRlciBwYW5lLCB3aXRoIHBhdGhzIGJ1aWx0IGZyb20gYXJpYS1sZXZlbC4gRm9sZGVyc1xuICogdW5kZXIgRmF2b3JpdGVzIGFyZSBza2lwcGVkICh0aGV5IGR1cGxpY2F0ZSB0aGUgbWFpbGJveCB0cmVlKSwgYXMgYXJlIHRoZSB0b3AtbGV2ZWwgbWFpbGJveFxuICogcm9vdHMsIHdob3NlIG5hbWVzIGFyZSBhY2NvdW50IGFkZHJlc3Nlcy5cbiAqIFVOVkVSSUZJRUQ6IGNvbGxhcHNlZCBzdWJmb2xkZXJzIG1heSBiZSBhYnNlbnQgZnJvbSB0aGUgRE9NLlxuICovXG5leHBvcnQgZnVuY3Rpb24gcmVhZEZvbGRlclRyZWUocm9vdDogUGFyZW50Tm9kZSA9IGRvY3VtZW50KTogRm9sZGVyTm9kZVtdIHtcbiAgY29uc3Qgc3RhY2s6IHN0cmluZ1tdID0gW107XG4gIGNvbnN0IGZvbGRlcnM6IEZvbGRlck5vZGVbXSA9IFtdO1xuICBsZXQgaW5GYXZvcml0ZXMgPSBmYWxzZTtcbiAgZm9yIChjb25zdCBlbGVtZW50IG9mIHJlc29sdmVBbGwoXCJmb2xkZXJJdGVtXCIsIHJvb3QpKSB7XG4gICAgY29uc3QgbGV2ZWwgPSBOdW1iZXIoZWxlbWVudC5nZXRBdHRyaWJ1dGUoXCJhcmlhLWxldmVsXCIpID8/IFwiMVwiKTtcbiAgICBjb25zdCBuYW1lID0gZm9sZGVyTmFtZShlbGVtZW50KTtcbiAgICBzdGFjay5sZW5ndGggPSBNYXRoLm1heChsZXZlbCAtIDEsIDApO1xuICAgIHN0YWNrLnB1c2gobmFtZSk7XG4gICAgaWYgKGxldmVsID09PSAxKSB7XG4gICAgICBpbkZhdm9yaXRlcyA9IGVsZW1lbnQuZ2V0QXR0cmlidXRlKFwiZGF0YS1mb2xkZXItbmFtZVwiKSA9PT0gXCJmYXZvcml0ZXNcIjtcbiAgICAgIGNvbnRpbnVlO1xuICAgIH1cbiAgICBpZiAoaW5GYXZvcml0ZXMpIGNvbnRpbnVlO1xuICAgIGZvbGRlcnMucHVzaCh7IG5hbWUsIHBhdGg6IHN0YWNrLnNsaWNlKDEpLmZpbHRlcihCb29sZWFuKS5qb2luKFwiL1wiKSwgZWxlbWVudCB9KTtcbiAgfVxuICByZXR1cm4gZm9sZGVycztcbn1cblxuLyoqIE1lc3NhZ2Ugcm93cyBjdXJyZW50bHkgc2VsZWN0ZWQgKG9yIG9wZW4pIGluIHRoZSBsaXN0LiAqL1xuZXhwb3J0IGZ1bmN0aW9uIHNlbGVjdGVkUm93cyhyb290OiBQYXJlbnROb2RlID0gZG9jdW1lbnQpOiBFbGVtZW50W10ge1xuICByZXR1cm4gcmVzb2x2ZUFsbChcIm1lc3NhZ2VSb3dcIiwgcm9vdCkuZmlsdGVyKFxuICAgIChyb3cpID0+IHJvdy5nZXRBdHRyaWJ1dGUoXCJhcmlhLXNlbGVjdGVkXCIpID09PSBcInRydWVcIixcbiAgKTtcbn1cblxuLyoqIEEgcm93J3MgdGljayBib3ggKFwiU2VsZWN0IGEgY29udmVyc2F0aW9uXCIpOyBpdHMgYXJpYS1jaGVja2VkIHNheXMgd2hldGhlciBpdCBpcyB0aWNrZWQuICovXG5leHBvcnQgZnVuY3Rpb24gcm93Q2hlY2tib3gocm93OiBFbGVtZW50KTogSFRNTEVsZW1lbnQgfCBudWxsIHtcbiAgcmV0dXJuIHJvdy5xdWVyeVNlbGVjdG9yPEhUTUxFbGVtZW50PignW3JvbGU9XCJjaGVja2JveFwiXScpO1xufVxuXG4vKiogUm93cyB3aG9zZSB0aWNrIGJveCBpcyB0aWNrZWQ6IHRoZSBtdWx0aS1zZWxlY3Rpb24sIGFzIG9wcG9zZWQgdG8gdGhlIHNpbmdsZSBvcGVuIHJvdy4gKi9cbmV4cG9ydCBmdW5jdGlvbiBjaGVja2VkUm93cyhyb290OiBQYXJlbnROb2RlID0gZG9jdW1lbnQpOiBFbGVtZW50W10ge1xuICByZXR1cm4gcmVzb2x2ZUFsbChcIm1lc3NhZ2VSb3dcIiwgcm9vdCkuZmlsdGVyKFxuICAgIChyb3cpID0+IHJvd0NoZWNrYm94KHJvdyk/LmdldEF0dHJpYnV0ZShcImFyaWEtY2hlY2tlZFwiKSA9PT0gXCJ0cnVlXCIsXG4gICk7XG59XG5cbi8qKlxuICogVGhlIHJvdyB0aGUga2V5Ym9hcmQgY3Vyc29yIGlzIG9uOiB0aGUgb25lIGhvbGRpbmcgZm9jdXMsIGVsc2UgT3V0bG9vaydzIHJvdmluZy10YWJpbmRleCByb3csXG4gKiBlbHNlIHRoZSBzZWxlY3RlZCBvbmUuXG4gKi9cbmV4cG9ydCBmdW5jdGlvbiBjdXJzb3JSb3cocm9vdDogRG9jdW1lbnQgPSBkb2N1bWVudCk6IEVsZW1lbnQgfCBudWxsIHtcbiAgY29uc3Qgcm93cyA9IHJlc29sdmVBbGwoXCJtZXNzYWdlUm93XCIsIHJvb3QpO1xuICBjb25zdCBhY3RpdmUgPSByb290LmFjdGl2ZUVsZW1lbnQ7XG4gIHJldHVybiAoXG4gICAgcm93cy5maW5kKChyb3cpID0+IGFjdGl2ZSAmJiByb3cuY29udGFpbnMoYWN0aXZlKSkgPz9cbiAgICByb3dzLmZpbmQoKHJvdykgPT4gcm93LmdldEF0dHJpYnV0ZShcInRhYmluZGV4XCIpID09PSBcIjBcIikgPz9cbiAgICBzZWxlY3RlZFJvd3Mocm9vdClbMF0gPz9cbiAgICBudWxsXG4gICk7XG59XG5cbi8qKiBPdXRsb29rIHN0YXJ0cyBhbiB1bnJlYWQgcm93J3MgYWNjZXNzaWJsZSBuYW1lIHdpdGggXCJVbnJlYWRcIiAoRW5nbGlzaCBVSSkuICovXG5leHBvcnQgZnVuY3Rpb24gaXNVbnJlYWRSb3cocm93OiBFbGVtZW50KTogYm9vbGVhbiB7XG4gIHJldHVybiAvXlVucmVhZFxcYi8udGVzdChyb3cuZ2V0QXR0cmlidXRlKFwiYXJpYS1sYWJlbFwiKSA/PyBcIlwiKTtcbn1cblxuLyoqIE91dGxvb2sncyBjdXJyZW50IGxpZ2h0L2RhcmsgbW9kZSwgZnJvbSB0aGUgcGFsZXR0ZSB2YXJpYWJsZXMgaXQgd3JpdGVzIG9uIDxodG1sPi4gKi9cbmV4cG9ydCBmdW5jdGlvbiBvdXRsb29rQ29sb3JTY2hlbWUoaHRtbDogRWxlbWVudCk6IFwibGlnaHRcIiB8IFwiZGFya1wiIHtcbiAgY29uc3Qgc3R5bGUgPSBodG1sLmdldEF0dHJpYnV0ZShcInN0eWxlXCIpID8/IFwiXCI7XG4gIHJldHVybiAvLS1uYXRpdmUtc2Nyb2xsYmFyLWNvbG9yLXNjaGVtZTpcXHMqZGFyay8udGVzdChzdHlsZSkgPyBcImRhcmtcIiA6IFwibGlnaHRcIjtcbn1cblxuLyoqIEVudHJpZXMgZXhwZWN0ZWQgb24gZXZlcnkgbWFpbCB2aWV3IHRoYXQgZGlkIG5vdCByZXNvbHZlLiBOb24tZW1wdHkgbWVhbnMgc29tZXRoaW5nIGJyb2tlLiAqL1xuZXhwb3J0IGZ1bmN0aW9uIGJyb2tlbkVudHJpZXMocmVzdWx0czogSGVhbHRoUmVzdWx0W10pOiBTZWxlY3Rvck5hbWVbXSB7XG4gIHJldHVybiByZXN1bHRzLmZpbHRlcigocikgPT4gci5leHBlY3QgPT09IFwiYWx3YXlzXCIgJiYgci5jYW5kaWRhdGUgPCAwKS5tYXAoKHIpID0+IHIubmFtZSk7XG59XG4iLCAiLy8gRGV2ZWxvcGVyIGNhcHR1cmU6IHNlcmlhbGlzZXMgdGhlIGxpdmUgT3V0bG9vayBwYWdlIGludG8gb25lIEhUTUwgZmlsZSwgc28gYSBmaXh0dXJlIGNhbiBiZVxuLy8gbWFkZSAoZG9jcy9jYXB0dXJpbmctcGFnZXMubWQpIHdpdGhvdXQgRGV2VG9vbHMuIFNhdmVkIGxvY2FsbHkgYnkgdGhlIGNhbGxlcjsgbmV2ZXIgc2VudFxuLy8gYW55d2hlcmUuIFRoZSBmaWxlIGhvbGRzIHJlYWwgbWFpbCwgc28gaXQgZ29lcyB0aHJvdWdoIHNjcmlwdHMvc2NydWItZml4dHVyZS50cyBiZWZvcmUgaXRcbi8vIGdvZXMgYW55d2hlcmUgbmVhciB0aGUgcmVwby5cblxuaW1wb3J0IHsgYnJva2VuRW50cmllcywgaGVhbHRoQ2hlY2sgfSBmcm9tIFwiLi9zZWxlY3RvcnMudHNcIjtcblxuLyoqIE1lc3NhZ2UgdGhlIHBvcHVwIHNlbmRzIHRvIHRoZSBjb250ZW50IHNjcmlwdCBvZiB0aGUgYWN0aXZlIHRhYi4gKi9cbmV4cG9ydCBjb25zdCBDQVBUVVJFX01FU1NBR0UgPSBcIm9iYi1jYXB0dXJlXCI7XG5cbmNvbnN0IE9CQl9BVFRSSUJVVEUgPSAvXmRhdGEtb3Zlcmxvb2svO1xuLyoqIFRoZSBzY3J1YmJlciAoc2NyaXB0cy9zY3J1Yi1maXh0dXJlLnRzKSBsb29rcyBmb3IgdGhpcyBtYXJrZXIgdG8gc3BsaXQgb3V0IHJ1bnRpbWUgc3R5bGVzLiAqL1xuZXhwb3J0IGNvbnN0IFJVTlRJTUVfU1RZTEVfTUFSS0VSID0gXCJkYXRhLW92ZXJsb29rLXJ1bnRpbWVcIjtcblxuZXhwb3J0IGludGVyZmFjZSBDYXB0dXJlTWV0YSB7XG4gIGNhcHR1cmVkQXQ6IHN0cmluZztcbiAgb2JiVmVyc2lvbjogc3RyaW5nO1xuICAvKiogRmlyc3QgdHdvIHBhdGggc2VnbWVudHMgb25seTogZGVlcGVyIG9uZXMgY2FycnkgbWVzc2FnZSBpZHMuICovXG4gIHJvdXRlOiBzdHJpbmc7XG4gIHZpZXdwb3J0OiBzdHJpbmc7XG4gIHNldHRpbmdzOiBSZWNvcmQ8c3RyaW5nLCBzdHJpbmcgfCB1bmRlZmluZWQ+O1xuICBoZWFsdGg6IHsgYnJva2VuOiBzdHJpbmdbXTsgZW50cmllczogeyBuYW1lOiBzdHJpbmc7IGNhbmRpZGF0ZTogbnVtYmVyOyBjb3VudDogbnVtYmVyIH1bXSB9O1xufVxuXG5leHBvcnQgaW50ZXJmYWNlIENhcHR1cmUge1xuICBodG1sOiBzdHJpbmc7XG4gIG5hbWU6IHN0cmluZztcbiAgbWV0YTogQ2FwdHVyZU1ldGE7XG59XG5cbi8qKiBTdHlsZSBydWxlcyBPdXRsb29rIGluamVjdGVkIGF0IHJ1bnRpbWUgKENTUy1pbi1KUyksIG1pbnVzIHRoZSBvbmVzIE9CQiBpdHNlbGYgYWRkZWQuICovXG5mdW5jdGlvbiBydW50aW1lQ3NzKGRvYzogRG9jdW1lbnQpOiBzdHJpbmcge1xuICBjb25zdCBzaGVldHMgPSBbLi4uZG9jLnN0eWxlU2hlZXRzLCAuLi4oZG9jLmFkb3B0ZWRTdHlsZVNoZWV0cyA/PyBbXSldLmZpbHRlcihcbiAgICAoc2hlZXQpID0+ICEoXCJocmVmXCIgaW4gc2hlZXQgJiYgc2hlZXQuaHJlZiksXG4gICk7XG4gIHJldHVybiBzaGVldHNcbiAgICAuZmxhdE1hcCgoc2hlZXQpID0+IHtcbiAgICAgIHRyeSB7XG4gICAgICAgIHJldHVybiBbLi4uc2hlZXQuY3NzUnVsZXNdLm1hcCgocnVsZSkgPT4gcnVsZS5jc3NUZXh0KTtcbiAgICAgIH0gY2F0Y2gge1xuICAgICAgICByZXR1cm4gW107IC8vIGNyb3NzLW9yaWdpbiBzaGVldFxuICAgICAgfVxuICAgIH0pXG4gICAgLmZpbHRlcigocnVsZSkgPT4gIXJ1bGUuaW5jbHVkZXMoXCJkYXRhLW92ZXJsb29rXCIpKVxuICAgIC5qb2luKFwiXFxuXCIpO1xufVxuXG5leHBvcnQgZnVuY3Rpb24gY2FwdHVyZUZpbGVuYW1lKG5vdyA9IG5ldyBEYXRlKCkpOiBzdHJpbmcge1xuICBjb25zdCBzdGFtcCA9IG5vdy50b0lTT1N0cmluZygpLnNsaWNlKDAsIDE2KS5yZXBsYWNlKC9cXEQvZywgXCJcIik7XG4gIHJldHVybiBgb2JiLWNhcHR1cmUtJHtzdGFtcH0uaHRtbGA7XG59XG5cbmV4cG9ydCBmdW5jdGlvbiBjYXB0dXJlUGFnZShcbiAgZG9jOiBEb2N1bWVudCA9IGRvY3VtZW50LFxuICBsb2M6IFBpY2s8TG9jYXRpb24sIFwicGF0aG5hbWVcIj4gPSBsb2NhdGlvbixcbiAgdmVyc2lvbiA9IFwiZGV2XCIsXG4gIG5vdyA9IG5ldyBEYXRlKCksXG4pOiBDYXB0dXJlIHtcbiAgY29uc3QgbGl2ZSA9IGRvYy5kb2N1bWVudEVsZW1lbnQuZGF0YXNldDtcbiAgY29uc3QgcmVzdWx0cyA9IGhlYWx0aENoZWNrKGRvYyk7XG4gIGNvbnN0IG1ldGE6IENhcHR1cmVNZXRhID0ge1xuICAgIGNhcHR1cmVkQXQ6IG5vdy50b0lTT1N0cmluZygpLFxuICAgIG9iYlZlcnNpb246IHZlcnNpb24sXG4gICAgcm91dGU6IGxvYy5wYXRobmFtZS5zcGxpdChcIi9cIikuc2xpY2UoMCwgMykuam9pbihcIi9cIiksXG4gICAgdmlld3BvcnQ6IGAke2RvYy5kZWZhdWx0Vmlldz8uaW5uZXJXaWR0aCA/PyAwfXgke2RvYy5kZWZhdWx0Vmlldz8uaW5uZXJIZWlnaHQgPz8gMH1gLFxuICAgIHNldHRpbmdzOiB7XG4gICAgICBmZWF0dXJlczogbGl2ZS5vdmVybG9va0ZlYXR1cmVzLFxuICAgICAgcGFsZXR0ZTogbGl2ZS5vdmVybG9va1BhbGV0dGUsXG4gICAgICBtb2RlOiBsaXZlLm92ZXJsb29rTW9kZSxcbiAgICAgIGxheW91dDogbGl2ZS5vdmVybG9va0xheW91dCxcbiAgICAgIG91dGxvb2tTY2hlbWU6IGxpdmUub3Zlcmxvb2tPdXRsb29rU2NoZW1lLFxuICAgICAgcmVhZGluZ1BhbmU6IGxpdmUub3Zlcmxvb2tSZWFkaW5nUGFuZSxcbiAgICB9LFxuICAgIGhlYWx0aDoge1xuICAgICAgYnJva2VuOiBicm9rZW5FbnRyaWVzKHJlc3VsdHMpLFxuICAgICAgZW50cmllczogcmVzdWx0cy5tYXAoKHsgbmFtZSwgY2FuZGlkYXRlLCBjb3VudCB9KSA9PiAoeyBuYW1lLCBjYW5kaWRhdGUsIGNvdW50IH0pKSxcbiAgICB9LFxuICB9O1xuXG4gIC8vIFdvcmsgb24gYSBjb3B5LCB3aXRoIE9CQidzIG93biBtYXJrcyByZW1vdmVkLCBzbyB0aGUgY2FwdHVyZSBpcyBzdG9jayBPdXRsb29rLlxuICBjb25zdCByb290ID0gZG9jLmRvY3VtZW50RWxlbWVudC5jbG9uZU5vZGUodHJ1ZSkgYXMgSFRNTEVsZW1lbnQ7XG4gIGZvciAoY29uc3QgZWwgb2Ygcm9vdC5xdWVyeVNlbGVjdG9yQWxsKFwib3Zlcmxvb2stcm9vdFwiKSkgZWwucmVtb3ZlKCk7XG4gIGZvciAoY29uc3QgZWwgb2YgW3Jvb3QsIC4uLnJvb3QucXVlcnlTZWxlY3RvckFsbChcIipcIildKSB7XG4gICAgZm9yIChjb25zdCB7IG5hbWUgfSBvZiBbLi4uZWwuYXR0cmlidXRlc10pIHtcbiAgICAgIGlmIChPQkJfQVRUUklCVVRFLnRlc3QobmFtZSkpIGVsLnJlbW92ZUF0dHJpYnV0ZShuYW1lKTtcbiAgICB9XG4gIH1cblxuICBjb25zdCBoZWFkID0gcm9vdC5xdWVyeVNlbGVjdG9yKFwiaGVhZFwiKTtcbiAgY29uc3Qgc3R5bGUgPSBkb2MuY3JlYXRlRWxlbWVudChcInN0eWxlXCIpO1xuICBzdHlsZS5zZXRBdHRyaWJ1dGUoUlVOVElNRV9TVFlMRV9NQVJLRVIsIFwiXCIpO1xuICBzdHlsZS50ZXh0Q29udGVudCA9IHJ1bnRpbWVDc3MoZG9jKTtcbiAgaGVhZD8uYXBwZW5kKHN0eWxlKTtcbiAgY29uc3Qgc2NyaXB0ID0gZG9jLmNyZWF0ZUVsZW1lbnQoXCJzY3JpcHRcIik7XG4gIHNjcmlwdC50eXBlID0gXCJhcHBsaWNhdGlvbi9qc29uXCI7XG4gIHNjcmlwdC5pZCA9IFwib2JiLWNhcHR1cmUtbWV0YVwiO1xuICBzY3JpcHQudGV4dENvbnRlbnQgPSBKU09OLnN0cmluZ2lmeShtZXRhLCBudWxsLCAyKS5yZXBsYWNlQWxsKFwiPFwiLCBcIlxcXFx1MDAzY1wiKTtcbiAgaGVhZD8uYXBwZW5kKHNjcmlwdCk7XG5cbiAgcmV0dXJuIHsgaHRtbDogYDwhZG9jdHlwZSBodG1sPlxcbiR7cm9vdC5vdXRlckhUTUx9YCwgbmFtZTogY2FwdHVyZUZpbGVuYW1lKG5vdyksIG1ldGEgfTtcbn1cbiIsICIvLyBBY3Rpb24gdGFibGU6IHdoYXQgZWFjaCBhY3Rpb24gbWVhbnMgYW5kIGhvdyBpdCBleGVjdXRlcy4gUHVyZSBkYXRhLlxuLy9cbi8vIFN0cmF0ZWdpZXMgKFBMQU4ubWQsIEtleW1hcCBsYXllcik6XG4vLyAgIHBhc3N0aHJvdWdoIFx1MjAxMyBPdXRsb29rJ3MgR21haWwgc2hvcnRjdXQgbW9kZSBhbHJlYWR5IGRvZXMgaXQ7IGxldCB0aGUga2V5IHRocm91Z2guXG4vLyAgIGRyaXZlICAgICAgIFx1MjAxMyBjbGljayBhIGNvbnRyb2wgZm91bmQgdGhyb3VnaCB0aGUgc2VsZWN0b3IgcmVnaXN0cnkuIFJlYWN0IGhvbm91cnMgY2xpY2soKS5cbi8vICAgICAgICAgICAgICAgICAoRGlzcGF0Y2hpbmcgT3V0bG9vaydzIG93biBzaG9ydGN1dCBhcyBhIHN5bnRoZXRpYyBrZXkgZXZlbnQgZGlkIG5vdGhpbmcgb25cbi8vICAgICAgICAgICAgICAgICB0aGUgbGl2ZSBzaXRlIGluIE5vdiAyMDI2LCBzbyB0aGVyZSBpcyBubyBrZXlzdHJva2Ugc3RyYXRlZ3kuKVxuLy8gICB0b2dnbGUgICAgICBcdTIwMTMgY2xpY2sgYSB0b2dnbGUgYnV0dG9uIChSZWFkIC8gVW5yZWFkKSBvbmx5IGlmIHRoZSBzZWxlY3Rpb24gaXNuJ3QgYWxyZWFkeSBpblxuLy8gICAgICAgICAgICAgICAgIHRoZSB3YW50ZWQgc3RhdGUuXG4vLyAgIGV4dGVuc2lvbiAgIFx1MjAxMyBPQkIncyBvd24gYmVoYXZpb3VyIChwYWxldHRlLCBoZWxwLCB0aGUgbXVsdGktc2VsZWN0IGN1cnNvcikuIEEgY29tbWFuZCBtYXlcbi8vICAgICAgICAgICAgICAgICBkZWNsaW5lLCBpbiB3aGljaCBjYXNlIHRoZSBrZXkgcGFzc2VzIHRocm91Z2ggdG8gT3V0bG9vay5cbi8vICAgdG9kbyAgICAgICAgXHUyMDEzIG5vdCBkZWNpZGVkIHlldDsgdGhlIGtleSBpcyBzd2FsbG93ZWQgYW5kIGEgaGludCBzYXlzIHNvLlxuLy9cbi8vIE91dGxvb2sncyBHbWFpbC1tb2RlIGtleXMgY29tZSBmcm9tIGl0cyBvd24gcmliYm9uIHRvb2x0aXBzIGluIHRoZSBPY3QgMjAyNiBjYXB0dXJlOlxuLy8gQyBuZXcgbWFpbCwgIyBkZWxldGUsIEUgYXJjaGl2ZSwgViBtb3ZlLCBcdTIxRTdJL1x1MjFFN1UgcmVhZC91bnJlYWQsIEIgc25vb3plLCBMIGNhdGVnb3JpemUsXG4vLyBSL0EvRiByZXBseS9yZXBseSBhbGwvZm9yd2FyZCwgXHUyMzE4WiB1bmRvLlxuXG5pbXBvcnQgdHlwZSB7IFNlbGVjdG9yTmFtZSB9IGZyb20gXCIuLi9zZWxlY3RvcnMudHNcIjtcblxuZXhwb3J0IHR5cGUgRXh0ZW5zaW9uQ29tbWFuZCA9XG4gIHwgXCJoZWxwXCJcbiAgfCBcInBhbGV0dGUtZ29cIlxuICB8IFwidG9nZ2xlLXNlbGVjdFwiXG4gIHwgXCJjdXJzb3ItbmV4dFwiXG4gIHwgXCJjdXJzb3ItcHJldmlvdXNcIjtcblxuZXhwb3J0IHR5cGUgU3RyYXRlZ3kgPVxuICB8IHsga2luZDogXCJwYXNzdGhyb3VnaFwiIH1cbiAgfCB7IGtpbmQ6IFwiZHJpdmVcIjsgdGFyZ2V0OiBTZWxlY3Rvck5hbWUgfVxuICB8IHsga2luZDogXCJ0b2dnbGVcIjsgdGFyZ2V0OiBTZWxlY3Rvck5hbWU7IHdhbnQ6IFwicmVhZFwiIHwgXCJ1bnJlYWRcIiB9XG4gIHwgeyBraW5kOiBcImV4dGVuc2lvblwiOyBjb21tYW5kOiBFeHRlbnNpb25Db21tYW5kIH1cbiAgfCB7IGtpbmQ6IFwidG9kb1wiIH07XG5cbmV4cG9ydCBpbnRlcmZhY2UgQWN0aW9uIHtcbiAgZGVzY3JpcHRpb246IHN0cmluZztcbiAgc3RyYXRlZ3k6IFN0cmF0ZWd5O1xufVxuXG5jb25zdCBwYXNzdGhyb3VnaDogU3RyYXRlZ3kgPSB7IGtpbmQ6IFwicGFzc3Rocm91Z2hcIiB9O1xuY29uc3QgdG9kbzogU3RyYXRlZ3kgPSB7IGtpbmQ6IFwidG9kb1wiIH07XG5jb25zdCBkcml2ZSA9ICh0YXJnZXQ6IFNlbGVjdG9yTmFtZSk6IFN0cmF0ZWd5ID0+ICh7IGtpbmQ6IFwiZHJpdmVcIiwgdGFyZ2V0IH0pO1xuY29uc3QgZXh0ZW5zaW9uID0gKGNvbW1hbmQ6IEV4dGVuc2lvbkNvbW1hbmQpOiBTdHJhdGVneSA9PiAoeyBraW5kOiBcImV4dGVuc2lvblwiLCBjb21tYW5kIH0pO1xuY29uc3QgcmVhZFN0YXRlID0gKHdhbnQ6IFwicmVhZFwiIHwgXCJ1bnJlYWRcIik6IFN0cmF0ZWd5ID0+ICh7XG4gIGtpbmQ6IFwidG9nZ2xlXCIsXG4gIHRhcmdldDogXCJyZWFkVW5yZWFkQnV0dG9uXCIsXG4gIHdhbnQsXG59KTtcblxuZXhwb3J0IGNvbnN0IEFDVElPTlMgPSB7XG4gIGNvbXBvc2U6IHsgZGVzY3JpcHRpb246IFwiQ29tcG9zZVwiLCBzdHJhdGVneTogcGFzc3Rocm91Z2ggfSxcbiAgc2VhcmNoOiB7IGRlc2NyaXB0aW9uOiBcIkZvY3VzIHNlYXJjaFwiLCBzdHJhdGVneTogcGFzc3Rocm91Z2ggfSxcbiAgaGVscDogeyBkZXNjcmlwdGlvbjogXCJTaG93IGtleWJvYXJkIHNob3J0Y3V0c1wiLCBzdHJhdGVneTogZXh0ZW5zaW9uKFwiaGVscFwiKSB9LFxuICBuZXh0OiB7IGRlc2NyaXB0aW9uOiBcIk5leHQgY29udmVyc2F0aW9uXCIsIHN0cmF0ZWd5OiBleHRlbnNpb24oXCJjdXJzb3ItbmV4dFwiKSB9LFxuICBwcmV2aW91czogeyBkZXNjcmlwdGlvbjogXCJQcmV2aW91cyBjb252ZXJzYXRpb25cIiwgc3RyYXRlZ3k6IGV4dGVuc2lvbihcImN1cnNvci1wcmV2aW91c1wiKSB9LFxuICBvcGVuOiB7IGRlc2NyaXB0aW9uOiBcIk9wZW4gY29udmVyc2F0aW9uXCIsIHN0cmF0ZWd5OiBwYXNzdGhyb3VnaCB9LFxuICBzZWxlY3Q6IHsgZGVzY3JpcHRpb246IFwiU2VsZWN0IG9yIGRlc2VsZWN0XCIsIHN0cmF0ZWd5OiBleHRlbnNpb24oXCJ0b2dnbGUtc2VsZWN0XCIpIH0sXG4gIHNlbGVjdFJhbmdlOiB7IGRlc2NyaXB0aW9uOiBcIlNlbGVjdCByYW5nZSBzaW5jZSBsYXN0IHNlbGVjdGlvblwiLCBzdHJhdGVneTogdG9kbyB9LFxuICBzZWxlY3RBbGw6IHsgZGVzY3JpcHRpb246IFwiU2VsZWN0IGFsbFwiLCBzdHJhdGVneTogdG9kbyB9LFxuICBzZWxlY3ROb25lOiB7IGRlc2NyaXB0aW9uOiBcIlNlbGVjdCBub25lXCIsIHN0cmF0ZWd5OiB0b2RvIH0sXG4gIGFyY2hpdmU6IHsgZGVzY3JpcHRpb246IFwiQXJjaGl2ZVwiLCBzdHJhdGVneTogZHJpdmUoXCJhcmNoaXZlQnV0dG9uXCIpIH0sXG4gIGFyY2hpdmVQcmV2aW91czogeyBkZXNjcmlwdGlvbjogXCJBcmNoaXZlLCB0aGVuIG9wZW4gcHJldmlvdXNcIiwgc3RyYXRlZ3k6IHRvZG8gfSxcbiAgYXJjaGl2ZU5leHQ6IHsgZGVzY3JpcHRpb246IFwiQXJjaGl2ZSwgdGhlbiBvcGVuIG5leHRcIiwgc3RyYXRlZ3k6IHRvZG8gfSxcbiAgZGVsZXRlOiB7IGRlc2NyaXB0aW9uOiBcIkRlbGV0ZVwiLCBzdHJhdGVneTogZHJpdmUoXCJkZWxldGVCdXR0b25cIikgfSxcbiAganVuazogeyBkZXNjcmlwdGlvbjogXCJSZXBvcnQganVua1wiLCBzdHJhdGVneTogcGFzc3Rocm91Z2ggfSxcbiAgLy8gT3V0bG9vaydzIG93biBcIk1vdmUgdG9cIiBwaWNrZXIgaGFzIGZvbGRlciBzZWFyY2gsIHdoaWNoIGNvdmVycyBGYXN0bWFpbCdzIGBtYCB1bnRpbCB0aGVcbiAgLy8gcGFsZXR0ZSBjYW4gbW92ZSBtZXNzYWdlcyBpdHNlbGYgKGRvY3MvZGVjaXNpb25zLm1kKS5cbiAgbW92ZTogeyBkZXNjcmlwdGlvbjogXCJNb3ZlIHRvIGZvbGRlclx1MjAyNlwiLCBzdHJhdGVneTogZHJpdmUoXCJtb3ZlQnV0dG9uXCIpIH0sXG4gIGdvVG86IHsgZGVzY3JpcHRpb246IFwiR28gdG8gZm9sZGVyIG9yIGFwcFx1MjAyNlwiLCBzdHJhdGVneTogZXh0ZW5zaW9uKFwicGFsZXR0ZS1nb1wiKSB9LFxuICBzbm9vemU6IHsgZGVzY3JpcHRpb246IFwiU25vb3plXCIsIHN0cmF0ZWd5OiBwYXNzdGhyb3VnaCB9LFxuICBtYXJrUmVhZDogeyBkZXNjcmlwdGlvbjogXCJNYXJrIHJlYWRcIiwgc3RyYXRlZ3k6IHJlYWRTdGF0ZShcInJlYWRcIikgfSxcbiAgbWFya1VucmVhZDogeyBkZXNjcmlwdGlvbjogXCJNYXJrIHVucmVhZFwiLCBzdHJhdGVneTogcmVhZFN0YXRlKFwidW5yZWFkXCIpIH0sXG4gIHJlcGx5OiB7IGRlc2NyaXB0aW9uOiBcIlJlcGx5XCIsIHN0cmF0ZWd5OiBwYXNzdGhyb3VnaCB9LFxuICByZXBseUFsbDogeyBkZXNjcmlwdGlvbjogXCJSZXBseSBhbGxcIiwgc3RyYXRlZ3k6IHBhc3N0aHJvdWdoIH0sXG4gIGZvcndhcmQ6IHsgZGVzY3JpcHRpb246IFwiRm9yd2FyZFwiLCBzdHJhdGVneTogcGFzc3Rocm91Z2ggfSxcbiAgdW5kbzogeyBkZXNjcmlwdGlvbjogXCJVbmRvXCIsIHN0cmF0ZWd5OiBkcml2ZShcInVuZG9CdXR0b25cIikgfSxcbiAgcmVmcmVzaDogeyBkZXNjcmlwdGlvbjogXCJSZWZyZXNoXCIsIHN0cmF0ZWd5OiBkcml2ZShcInN5bmNCdXR0b25cIikgfSxcbiAgZ29NYWlsOiB7IGRlc2NyaXB0aW9uOiBcIkdvIHRvIE1haWxcIiwgc3RyYXRlZ3k6IGRyaXZlKFwibWFpbEFwcFwiKSB9LFxuICBnb0NhbGVuZGFyOiB7IGRlc2NyaXB0aW9uOiBcIkdvIHRvIENhbGVuZGFyXCIsIHN0cmF0ZWd5OiBkcml2ZShcImNhbGVuZGFyQXBwXCIpIH0sXG4gIGdvQ29udGFjdHM6IHsgZGVzY3JpcHRpb246IFwiR28gdG8gQ29udGFjdHNcIiwgc3RyYXRlZ3k6IGRyaXZlKFwicGVvcGxlQXBwXCIpIH0sXG4gIGdvU2V0dGluZ3M6IHsgZGVzY3JpcHRpb246IFwiR28gdG8gU2V0dGluZ3NcIiwgc3RyYXRlZ3k6IGRyaXZlKFwic2V0dGluZ3NCdXR0b25cIikgfSxcbn0gYXMgY29uc3Qgc2F0aXNmaWVzIFJlY29yZDxzdHJpbmcsIEFjdGlvbj47XG5cbmV4cG9ydCB0eXBlIEFjdGlvbklkID0ga2V5b2YgdHlwZW9mIEFDVElPTlM7XG4iLCAiLy8gVGhlIEZhc3RtYWlsLXN0eWxlIGJpbmRpbmcgc2V0IChQTEFOLm1kLCBTaG9ydGN1dCBtYXApLiBLZXlzIGFyZSBLZXlib2FyZEV2ZW50LmtleSB2YWx1ZXMuXG5cbmltcG9ydCB0eXBlIHsgQWN0aW9uSWQgfSBmcm9tIFwiLi9hY3Rpb25zLnRzXCI7XG5pbXBvcnQgdHlwZSB7IEJpbmRpbmcsIENvbnRleHQgfSBmcm9tIFwiLi9lbmdpbmUudHNcIjtcblxuY29uc3QgTUFJTDogcmVhZG9ubHkgQ29udGV4dFtdID0gW1wibGlzdFwiLCBcInJlYWRpbmdcIl07XG5jb25zdCBMSVNUOiByZWFkb25seSBDb250ZXh0W10gPSBbXCJsaXN0XCJdO1xuY29uc3QgUkVBRElORzogcmVhZG9ubHkgQ29udGV4dFtdID0gW1wicmVhZGluZ1wiXTtcblxuY29uc3QgYmluZCA9IChrZXlzOiBzdHJpbmcsIGFjdGlvbjogQWN0aW9uSWQsIGNvbnRleHRzID0gTUFJTCk6IEJpbmRpbmc8QWN0aW9uSWQ+ID0+ICh7XG4gIGtleXM6IGtleXMuc3BsaXQoXCIgXCIpLFxuICBhY3Rpb24sXG4gIGNvbnRleHRzLFxufSk7XG5cbmV4cG9ydCBjb25zdCBCSU5ESU5HUzogcmVhZG9ubHkgQmluZGluZzxBY3Rpb25JZD5bXSA9IFtcbiAgYmluZChcImNcIiwgXCJjb21wb3NlXCIpLFxuICBiaW5kKFwiL1wiLCBcInNlYXJjaFwiKSxcbiAgYmluZChcIj9cIiwgXCJoZWxwXCIpLFxuICBiaW5kKFwialwiLCBcIm5leHRcIiksXG4gIGJpbmQoXCJrXCIsIFwicHJldmlvdXNcIiksXG4gIGJpbmQoXCJvXCIsIFwib3BlblwiLCBMSVNUKSxcbiAgYmluZChcIkVudGVyXCIsIFwib3BlblwiLCBMSVNUKSxcbiAgYmluZChcInhcIiwgXCJzZWxlY3RcIiwgTElTVCksXG4gIGJpbmQoXCJYXCIsIFwic2VsZWN0UmFuZ2VcIiwgTElTVCksXG4gIGJpbmQoXCIqIGFcIiwgXCJzZWxlY3RBbGxcIiwgTElTVCksXG4gIGJpbmQoXCIqIG5cIiwgXCJzZWxlY3ROb25lXCIsIExJU1QpLFxuICBiaW5kKFwieVwiLCBcImFyY2hpdmVcIiksXG4gIGJpbmQoXCJdXCIsIFwiYXJjaGl2ZVByZXZpb3VzXCIsIFJFQURJTkcpLFxuICBiaW5kKFwiW1wiLCBcImFyY2hpdmVOZXh0XCIsIFJFQURJTkcpLFxuICBiaW5kKFwiZFwiLCBcImRlbGV0ZVwiKSxcbiAgYmluZChcIiFcIiwgXCJqdW5rXCIpLFxuICBiaW5kKFwibVwiLCBcIm1vdmVcIiksXG4gIGJpbmQoXCJnXCIsIFwiZ29Ub1wiKSxcbiAgYmluZChcImJcIiwgXCJzbm9vemVcIiksXG4gIGJpbmQoXCIuIHJcIiwgXCJtYXJrUmVhZFwiKSxcbiAgYmluZChcIi4gdVwiLCBcIm1hcmtVbnJlYWRcIiksXG4gIGJpbmQoXCJyXCIsIFwicmVwbHlcIiwgUkVBRElORyksXG4gIGJpbmQoXCJhXCIsIFwicmVwbHlBbGxcIiwgUkVBRElORyksXG4gIGJpbmQoXCJmXCIsIFwiZm9yd2FyZFwiLCBSRUFESU5HKSxcbiAgYmluZChcInpcIiwgXCJ1bmRvXCIpLFxuICBiaW5kKFwidVwiLCBcInJlZnJlc2hcIiksXG4gIGJpbmQoXCJHIG1cIiwgXCJnb01haWxcIiksXG4gIGJpbmQoXCJHIGNcIiwgXCJnb0NhbGVuZGFyXCIpLFxuICBiaW5kKFwiRyBhXCIsIFwiZ29Db250YWN0c1wiKSxcbiAgYmluZChcIkcgc1wiLCBcImdvU2V0dGluZ3NcIiksXG5dO1xuIiwgIi8vIFB1cmUga2V5LXNlcXVlbmNlIGVuZ2luZS4gTm8gRE9NIGFjY2Vzcywgc28gaXQgaXMgZnVsbHkgdW5pdC10ZXN0YWJsZS5cbi8vXG4vLyBLZXlzIGFyZSBLZXlib2FyZEV2ZW50LmtleSB2YWx1ZXMsIHdoaWNoIGFscmVhZHkgcmVmbGVjdCBTaGlmdCAoXCJHXCIsIFwiKlwiLCBcIiFcIiwgXCI/XCIpLlxuLy8gQSBiaW5kaW5nIGxpa2UgRmFzdG1haWwncyBcIlNoaWZ0K0cgdGhlbiBNXCIgaXMgd3JpdHRlbiBbXCJHXCIsIFwibVwiXS5cblxuZXhwb3J0IHR5cGUgQ29udGV4dCA9IFwibGlzdFwiIHwgXCJyZWFkaW5nXCIgfCBcInBhbGV0dGVcIjtcblxuZXhwb3J0IGludGVyZmFjZSBCaW5kaW5nPEEgZXh0ZW5kcyBzdHJpbmcgPSBzdHJpbmc+IHtcbiAga2V5czogcmVhZG9ubHkgc3RyaW5nW107XG4gIGFjdGlvbjogQTtcbiAgY29udGV4dHM6IHJlYWRvbmx5IENvbnRleHRbXTtcbn1cblxuZXhwb3J0IHR5cGUgRmVlZFJlc3VsdDxBIGV4dGVuZHMgc3RyaW5nPiA9XG4gIHwgeyBraW5kOiBcIm1hdGNoXCI7IGFjdGlvbjogQSB9XG4gIHwgeyBraW5kOiBcInBlbmRpbmdcIjsgcHJlZml4OiByZWFkb25seSBzdHJpbmdbXTsgbmV4dDogcmVhZG9ubHkgQmluZGluZzxBPltdIH1cbiAgLyoqIEEgcHJlZml4IHdhcyBwZW5kaW5nIGFuZCB0aGlzIGtleSBjb21wbGV0ZWQgbm90aGluZy4gVGhlIGtleSBzaG91bGQgYmUgc3dhbGxvd2VkLiAqL1xuICB8IHsga2luZDogXCJjYW5jZWxsZWRcIiB9XG4gIHwgeyBraW5kOiBcIm5vbmVcIiB9O1xuXG5leHBvcnQgY29uc3QgREVGQVVMVF9USU1FT1VUX01TID0gMTUwMDtcblxuY29uc3Qgc3RhcnRzV2l0aCA9IChrZXlzOiByZWFkb25seSBzdHJpbmdbXSwgcHJlZml4OiByZWFkb25seSBzdHJpbmdbXSkgPT5cbiAgcHJlZml4Lmxlbmd0aCA8PSBrZXlzLmxlbmd0aCAmJiBwcmVmaXguZXZlcnkoKGssIGkpID0+IGtleXNbaV0gPT09IGspO1xuXG5leHBvcnQgY2xhc3MgU2VxdWVuY2VFbmdpbmU8QSBleHRlbmRzIHN0cmluZz4ge1xuICByZWFkb25seSAjYmluZGluZ3M6IHJlYWRvbmx5IEJpbmRpbmc8QT5bXTtcbiAgcmVhZG9ubHkgI3RpbWVvdXRNczogbnVtYmVyO1xuICAjcHJlZml4OiBzdHJpbmdbXSA9IFtdO1xuICAjbGFzdEF0ID0gMDtcblxuICBjb25zdHJ1Y3RvcihiaW5kaW5nczogcmVhZG9ubHkgQmluZGluZzxBPltdLCB0aW1lb3V0TXMgPSBERUZBVUxUX1RJTUVPVVRfTVMpIHtcbiAgICB0aGlzLiNiaW5kaW5ncyA9IGJpbmRpbmdzO1xuICAgIHRoaXMuI3RpbWVvdXRNcyA9IHRpbWVvdXRNcztcbiAgfVxuXG4gIGdldCBwZW5kaW5nKCk6IHJlYWRvbmx5IHN0cmluZ1tdIHtcbiAgICByZXR1cm4gdGhpcy4jcHJlZml4O1xuICB9XG5cbiAgcmVzZXQoKTogdm9pZCB7XG4gICAgdGhpcy4jcHJlZml4ID0gW107XG4gIH1cblxuICBmZWVkKGtleTogc3RyaW5nLCBjb250ZXh0OiBDb250ZXh0LCBub3c6IG51bWJlcik6IEZlZWRSZXN1bHQ8QT4ge1xuICAgIGlmICh0aGlzLiNwcmVmaXgubGVuZ3RoID4gMCAmJiBub3cgLSB0aGlzLiNsYXN0QXQgPiB0aGlzLiN0aW1lb3V0TXMpIHRoaXMucmVzZXQoKTtcbiAgICBjb25zdCB3YXNQZW5kaW5nID0gdGhpcy4jcHJlZml4Lmxlbmd0aCA+IDA7XG4gICAgY29uc3Qgc2VxID0gWy4uLnRoaXMuI3ByZWZpeCwga2V5XTtcbiAgICBjb25zdCBhY3RpdmUgPSB0aGlzLiNiaW5kaW5ncy5maWx0ZXIoKGIpID0+IGIuY29udGV4dHMuaW5jbHVkZXMoY29udGV4dCkpO1xuICAgIGNvbnN0IGxvbmdlciA9IGFjdGl2ZS5maWx0ZXIoKGIpID0+IGIua2V5cy5sZW5ndGggPiBzZXEubGVuZ3RoICYmIHN0YXJ0c1dpdGgoYi5rZXlzLCBzZXEpKTtcblxuICAgIGlmIChsb25nZXIubGVuZ3RoID4gMCkge1xuICAgICAgdGhpcy4jcHJlZml4ID0gc2VxO1xuICAgICAgdGhpcy4jbGFzdEF0ID0gbm93O1xuICAgICAgcmV0dXJuIHsga2luZDogXCJwZW5kaW5nXCIsIHByZWZpeDogc2VxLCBuZXh0OiBsb25nZXIgfTtcbiAgICB9XG5cbiAgICB0aGlzLnJlc2V0KCk7XG4gICAgY29uc3QgZXhhY3QgPSBhY3RpdmUuZmluZCgoYikgPT4gYi5rZXlzLmxlbmd0aCA9PT0gc2VxLmxlbmd0aCAmJiBzdGFydHNXaXRoKGIua2V5cywgc2VxKSk7XG4gICAgaWYgKGV4YWN0KSByZXR1cm4geyBraW5kOiBcIm1hdGNoXCIsIGFjdGlvbjogZXhhY3QuYWN0aW9uIH07XG4gICAgcmV0dXJuIHdhc1BlbmRpbmcgPyB7IGtpbmQ6IFwiY2FuY2VsbGVkXCIgfSA6IHsga2luZDogXCJub25lXCIgfTtcbiAgfVxufVxuXG4vKiogQmluZGluZ3MgdGhhdCBzaGFkb3cgZWFjaCBvdGhlciAoc2FtZSBrZXlzLCBvciBvbmUgYSBwcmVmaXggb2YgYW5vdGhlcikgaW4gYSBzaGFyZWQgY29udGV4dC4gKi9cbmV4cG9ydCBmdW5jdGlvbiBmaW5kQ29uZmxpY3RzKGJpbmRpbmdzOiByZWFkb25seSBCaW5kaW5nW10pOiBzdHJpbmdbXSB7XG4gIGNvbnN0IGNvbmZsaWN0czogc3RyaW5nW10gPSBbXTtcbiAgYmluZGluZ3MuZm9yRWFjaCgoYSwgaSkgPT4ge1xuICAgIGZvciAoY29uc3QgYiBvZiBiaW5kaW5ncy5zbGljZShpICsgMSkpIHtcbiAgICAgIGNvbnN0IHNoYXJlZCA9IGEuY29udGV4dHMuc29tZSgoYykgPT4gYi5jb250ZXh0cy5pbmNsdWRlcyhjKSk7XG4gICAgICBjb25zdCBbc2hvcnQsIGxvbmddID0gYS5rZXlzLmxlbmd0aCA8PSBiLmtleXMubGVuZ3RoID8gW2EsIGJdIDogW2IsIGFdO1xuICAgICAgaWYgKHNoYXJlZCAmJiBzdGFydHNXaXRoKGxvbmcua2V5cywgc2hvcnQua2V5cykpIHtcbiAgICAgICAgY29uZmxpY3RzLnB1c2goXG4gICAgICAgICAgYCR7c2hvcnQua2V5cy5qb2luKFwiIFwiKX0gKCR7c2hvcnQuYWN0aW9ufSkgdnMgJHtsb25nLmtleXMuam9pbihcIiBcIil9ICgke2xvbmcuYWN0aW9ufSlgLFxuICAgICAgICApO1xuICAgICAgfVxuICAgIH1cbiAgfSk7XG4gIHJldHVybiBjb25mbGljdHM7XG59XG5cbmludGVyZmFjZSBLZXlMaWtlIHtcbiAga2V5OiBzdHJpbmc7XG4gIGN0cmxLZXk6IGJvb2xlYW47XG4gIG1ldGFLZXk6IGJvb2xlYW47XG4gIGFsdEtleTogYm9vbGVhbjtcbn1cblxuY29uc3QgTU9ESUZJRVJfS0VZUyA9IG5ldyBTZXQoW1wiU2hpZnRcIiwgXCJDb250cm9sXCIsIFwiQWx0XCIsIFwiTWV0YVwiLCBcIkNhcHNMb2NrXCIsIFwiRGVhZFwiXSk7XG5cbi8qKiBUaGUgZW5naW5lIGtleSBmb3IgYW4gZXZlbnQsIG9yIG51bGwgd2hlbiBpdCBiZWxvbmdzIHRvIHRoZSBicm93c2VyL09TIChDdHJsLCBDbWQsIEFsdCBjaG9yZHMpLiAqL1xuZXhwb3J0IGZ1bmN0aW9uIGtleUZyb21FdmVudChldmVudDogS2V5TGlrZSk6IHN0cmluZyB8IG51bGwge1xuICBpZiAoZXZlbnQuY3RybEtleSB8fCBldmVudC5tZXRhS2V5IHx8IGV2ZW50LmFsdEtleSkgcmV0dXJuIG51bGw7XG4gIGlmIChNT0RJRklFUl9LRVlTLmhhcyhldmVudC5rZXkpKSByZXR1cm4gbnVsbDtcbiAgcmV0dXJuIGV2ZW50LmtleTtcbn1cbiIsICI6aG9zdCB7XG4gIGFsbDogaW5pdGlhbDtcbiAgLS1vbC1iZzogI2ZmZjtcbiAgLS1vbC1mZzogIzFmMWYxZjtcbiAgLS1vbC1tdXRlZDogIzZiNmI2YjtcbiAgLS1vbC1hY2NlbnQ6ICMyZjZmZGU7XG4gIC0tb2wtYm9yZGVyOiByZ2IoMCAwIDAgLyAxMiUpO1xuICAtLW9sLXNoYWRvdzogMCA4cHggMzJweCByZ2IoMCAwIDAgLyAxOCUpO1xuICBmb250OlxuICAgIDEzcHggLyAxLjQgc3lzdGVtLXVpLFxuICAgIC1hcHBsZS1zeXN0ZW0sXG4gICAgXCJTZWdvZSBVSVwiLFxuICAgIHNhbnMtc2VyaWY7XG59XG5cbkBtZWRpYSAocHJlZmVycy1jb2xvci1zY2hlbWU6IGRhcmspIHtcbiAgOmhvc3Qge1xuICAgIC0tb2wtYmc6ICMyNjI2MjY7XG4gICAgLS1vbC1mZzogI2VjZWNlYztcbiAgICAtLW9sLW11dGVkOiAjYTBhMGEwO1xuICAgIC0tb2wtYWNjZW50OiAjN2FhN2ZmO1xuICAgIC0tb2wtYm9yZGVyOiByZ2IoMjU1IDI1NSAyNTUgLyAxNCUpO1xuICB9XG59XG5cbltoaWRkZW5dIHtcbiAgZGlzcGxheTogbm9uZSAhaW1wb3J0YW50O1xufVxuXG4udG9hc3QsXG4uaGludCxcbi5oZWFsdGgge1xuICBwb3NpdGlvbjogZml4ZWQ7XG4gIHotaW5kZXg6IDIxNDc0ODM2NDc7XG4gIGJhY2tncm91bmQ6IHZhcigtLW9sLWJnKTtcbiAgY29sb3I6IHZhcigtLW9sLWZnKTtcbiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tb2wtYm9yZGVyKTtcbiAgYm9yZGVyLXJhZGl1czogNnB4O1xuICBib3gtc2hhZG93OiB2YXIoLS1vbC1zaGFkb3cpO1xuICBwYWRkaW5nOiA2cHggMTBweDtcbn1cblxuLnRvYXN0IHtcbiAgYm90dG9tOiAyNHB4O1xuICBsZWZ0OiA1MCU7XG4gIHRyYW5zZm9ybTogdHJhbnNsYXRlWCgtNTAlKTtcbn1cblxuLnRvYXN0Lm5vdGljZSB7XG4gIGRpc3BsYXk6IGZsZXg7XG4gIGFsaWduLWl0ZW1zOiBjZW50ZXI7XG4gIGdhcDogMTJweDtcbiAgcGFkZGluZzogOHB4IDEwcHggOHB4IDE0cHg7XG59XG5cbi5hY3Rpb24ge1xuICBhbGw6IHVuc2V0O1xuICBjdXJzb3I6IHBvaW50ZXI7XG4gIHBhZGRpbmc6IDRweCAxMnB4O1xuICBib3JkZXItcmFkaXVzOiA5OTlweDtcbiAgYmFja2dyb3VuZDogdmFyKC0tb2wtYWNjZW50KTtcbiAgY29sb3I6ICNmZmY7XG4gIGZvbnQtd2VpZ2h0OiA2MDA7XG59XG5cbi5hY3Rpb246Zm9jdXMtdmlzaWJsZSxcbi5jbG9zZTpmb2N1cy12aXNpYmxlIHtcbiAgb3V0bGluZTogMnB4IHNvbGlkIHZhcigtLW9sLWFjY2VudCk7XG4gIG91dGxpbmUtb2Zmc2V0OiAycHg7XG59XG5cbi5oaW50IHtcbiAgYm90dG9tOiAyNHB4O1xuICByaWdodDogMjRweDtcbiAgZm9udC1mYW1pbHk6IHVpLW1vbm9zcGFjZSwgXCJTRiBNb25vXCIsIE1lbmxvLCBtb25vc3BhY2U7XG59XG5cbi5oZWFsdGgge1xuICB0b3A6IDhweDtcbiAgcmlnaHQ6IDhweDtcbiAgZGlzcGxheTogZmxleDtcbiAgZ2FwOiA4cHg7XG4gIGFsaWduLWl0ZW1zOiBjZW50ZXI7XG4gIGJvcmRlci1jb2xvcjogI2Q5NzcwNjtcbn1cblxuLmNsb3NlIHtcbiAgYWxsOiB1bnNldDtcbiAgY3Vyc29yOiBwb2ludGVyO1xuICBjb2xvcjogdmFyKC0tb2wtbXV0ZWQpO1xuICBwYWRkaW5nOiAwIDJweDtcbn1cblxuLmJhY2tkcm9wIHtcbiAgcG9zaXRpb246IGZpeGVkO1xuICBpbnNldDogMDtcbiAgei1pbmRleDogMjE0NzQ4MzY0NjtcbiAgYmFja2dyb3VuZDogcmdiKDAgMCAwIC8gMjUlKTtcbiAgZGlzcGxheTogZmxleDtcbiAganVzdGlmeS1jb250ZW50OiBjZW50ZXI7XG4gIGFsaWduLWl0ZW1zOiBmbGV4LXN0YXJ0O1xuICBwYWRkaW5nLXRvcDogMTJ2aDtcbn1cblxuLnBhbmVsIHtcbiAgYmFja2dyb3VuZDogdmFyKC0tb2wtYmcpO1xuICBjb2xvcjogdmFyKC0tb2wtZmcpO1xuICBib3JkZXItcmFkaXVzOiAxMHB4O1xuICBib3gtc2hhZG93OiB2YXIoLS1vbC1zaGFkb3cpO1xuICB3aWR0aDogbWluKDUyMHB4LCA5MHZ3KTtcbiAgbWF4LWhlaWdodDogNzB2aDtcbiAgb3ZlcmZsb3c6IGF1dG87XG59XG5cbi5oZWxwIHtcbiAgcGFkZGluZzogMTZweCAyMHB4O1xufVxuXG4uaGVscCBoMiB7XG4gIGZvbnQtc2l6ZTogMTVweDtcbiAgbWFyZ2luOiAwIDAgMTJweDtcbn1cblxuLmhlbHAgdGFibGUge1xuICBib3JkZXItY29sbGFwc2U6IGNvbGxhcHNlO1xuICB3aWR0aDogMTAwJTtcbn1cblxuLmhlbHAgdGQge1xuICBwYWRkaW5nOiAzcHggMDtcbn1cblxuLmtleXMge1xuICB3aWR0aDogN2VtO1xufVxuXG5rYmQge1xuICBmb250LWZhbWlseTogdWktbW9ub3NwYWNlLCBcIlNGIE1vbm9cIiwgTWVubG8sIG1vbm9zcGFjZTtcbiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tb2wtYm9yZGVyKTtcbiAgYm9yZGVyLXJhZGl1czogNHB4O1xuICBwYWRkaW5nOiAwIDVweDtcbn1cblxuLnBhbGV0dGUgaW5wdXQge1xuICBhbGw6IHVuc2V0O1xuICBib3gtc2l6aW5nOiBib3JkZXItYm94O1xuICB3aWR0aDogMTAwJTtcbiAgcGFkZGluZzogMTJweCAxNnB4O1xuICBmb250LXNpemU6IDE1cHg7XG4gIGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1vbC1ib3JkZXIpO1xufVxuXG4ucGFsZXR0ZSB1bCB7XG4gIGxpc3Qtc3R5bGU6IG5vbmU7XG4gIG1hcmdpbjogMDtcbiAgcGFkZGluZzogNHB4IDA7XG59XG5cbi5wYWxldHRlIGxpIHtcbiAgcGFkZGluZzogNXB4IDE2cHg7XG4gIGN1cnNvcjogcG9pbnRlcjtcbn1cblxuLnBhbGV0dGUgbGlbYXJpYS1zZWxlY3RlZD1cInRydWVcIl0ge1xuICBiYWNrZ3JvdW5kOiBjb2xvci1taXgoaW4gc3JnYiwgdmFyKC0tb2wtYWNjZW50KSAxNiUsIHRyYW5zcGFyZW50KTtcbn1cblxuLnBhbGV0dGUgbGkgLnBhdGgge1xuICBjb2xvcjogdmFyKC0tb2wtbXV0ZWQpO1xuICBtYXJnaW4tbGVmdDogNnB4O1xufVxuXG4ucGFsZXR0ZSBtYXJrIHtcbiAgYmFja2dyb3VuZDogbm9uZTtcbiAgY29sb3I6IHZhcigtLW9sLWFjY2VudCk7XG4gIGZvbnQtd2VpZ2h0OiA2MDA7XG59XG5cbi5wYWxldHRlIC5lbXB0eSB7XG4gIGNvbG9yOiB2YXIoLS1vbC1tdXRlZCk7XG4gIGN1cnNvcjogZGVmYXVsdDtcbn1cbiIsICIvLyBPQkIncyBvd24gVUkgbGl2ZXMgaW4gb25lIHNoYWRvdyByb290LCBzbyBPdXRsb29rJ3MgQ1NTIGNhbm5vdCByZWFjaCBpdCBhbmQgb3VycyBjYW5ub3Rcbi8vIGxlYWsgaW50byBPdXRsb29rLiBFdmVyeXRoaW5nIGhlcmUgaXMgcGxhaW4gRE9NOyBubyBmcmFtZXdvcmsuXG5cbmltcG9ydCBjc3MgZnJvbSBcIi4vb3ZlcmxheS5jc3NcIjtcblxubGV0IHNoYWRvdzogU2hhZG93Um9vdCB8IG51bGwgPSBudWxsO1xuXG5leHBvcnQgZnVuY3Rpb24gb3ZlcmxheVJvb3QoKTogU2hhZG93Um9vdCB7XG4gIGlmIChzaGFkb3cpIHJldHVybiBzaGFkb3c7XG4gIGNvbnN0IGhvc3QgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KFwib3Zlcmxvb2stcm9vdFwiKTtcbiAgZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50LmFwcGVuZChob3N0KTtcbiAgc2hhZG93ID0gaG9zdC5hdHRhY2hTaGFkb3coeyBtb2RlOiBcIm9wZW5cIiB9KTtcbiAgLy8gS2V5c3Ryb2tlcyB0eXBlZCBpbnRvIG91ciBVSSBtdXN0IG5vdCByZWFjaCBPdXRsb29rJ3MgZG9jdW1lbnQtbGV2ZWwgc2hvcnRjdXQgaGFuZGxlcnMsXG4gIC8vIHdoaWNoIHdvdWxkIHNlZSB0aGUgaG9zdCBlbGVtZW50IChub3QgYW4gaW5wdXQpIGFzIHRoZSB0YXJnZXQuXG4gIGZvciAoY29uc3QgdHlwZSBvZiBbXCJrZXlkb3duXCIsIFwia2V5cHJlc3NcIiwgXCJrZXl1cFwiXSkge1xuICAgIGhvc3QuYWRkRXZlbnRMaXN0ZW5lcih0eXBlLCAoZSkgPT4gZS5zdG9wUHJvcGFnYXRpb24oKSk7XG4gIH1cbiAgY29uc3Qgc3R5bGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KFwic3R5bGVcIik7XG4gIHN0eWxlLnRleHRDb250ZW50ID0gY3NzO1xuICBzaGFkb3cuYXBwZW5kKHN0eWxlKTtcbiAgcmV0dXJuIHNoYWRvdztcbn1cblxuLyoqIFRydWUgd2hlbiB0aGUgZXZlbnQgY2FtZSBmcm9tIGluc2lkZSBPQkIncyBvd24gVUkuICovXG5leHBvcnQgZnVuY3Rpb24gaXNGcm9tT3ZlcmxheShldmVudDogRXZlbnQpOiBib29sZWFuIHtcbiAgcmV0dXJuIHNoYWRvdyAhPT0gbnVsbCAmJiBldmVudC5jb21wb3NlZFBhdGgoKS5pbmNsdWRlcyhzaGFkb3cuaG9zdCk7XG59XG5cbi8qKiBUcnVlIHdoaWxlIGEgbW9kYWwgKHBhbGV0dGUgb3IgaGVscCkgaXMgb3BlbjsgaXQgb3ducyB0aGUga2V5Ym9hcmQgdW50aWwgY2xvc2VkLiAqL1xuZXhwb3J0IGZ1bmN0aW9uIGlzTW9kYWxPcGVuKCk6IGJvb2xlYW4ge1xuICByZXR1cm4gc2hhZG93Py5xdWVyeVNlbGVjdG9yKFwiLmJhY2tkcm9wXCIpICE9IG51bGw7XG59XG5cbmZ1bmN0aW9uIGVsPEsgZXh0ZW5kcyBrZXlvZiBIVE1MRWxlbWVudFRhZ05hbWVNYXA+KHRhZzogSywgY2xhc3NOYW1lOiBzdHJpbmcsIHRleHQ/OiBzdHJpbmcpIHtcbiAgY29uc3Qgbm9kZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQodGFnKTtcbiAgbm9kZS5jbGFzc05hbWUgPSBjbGFzc05hbWU7XG4gIGlmICh0ZXh0ICE9PSB1bmRlZmluZWQpIG5vZGUudGV4dENvbnRlbnQgPSB0ZXh0O1xuICByZXR1cm4gbm9kZTtcbn1cblxuZnVuY3Rpb24gc2luZ2xldG9uKGNsYXNzTmFtZTogc3RyaW5nKTogSFRNTERpdkVsZW1lbnQge1xuICBjb25zdCByb290ID0gb3ZlcmxheVJvb3QoKTtcbiAgY29uc3QgZXhpc3RpbmcgPSByb290LnF1ZXJ5U2VsZWN0b3I8SFRNTERpdkVsZW1lbnQ+KGAuJHtjbGFzc05hbWV9YCk7XG4gIGlmIChleGlzdGluZykgcmV0dXJuIGV4aXN0aW5nO1xuICBjb25zdCBub2RlID0gZWwoXCJkaXZcIiwgY2xhc3NOYW1lKTtcbiAgcm9vdC5hcHBlbmQobm9kZSk7XG4gIHJldHVybiBub2RlO1xufVxuXG5sZXQgdG9hc3RUaW1lcjogUmV0dXJuVHlwZTx0eXBlb2Ygc2V0VGltZW91dD4gfCB1bmRlZmluZWQ7XG5cbmV4cG9ydCBmdW5jdGlvbiB0b2FzdChtZXNzYWdlOiBzdHJpbmcsIG1zID0gMjUwMCk6IHZvaWQge1xuICBjb25zdCBub2RlID0gc2luZ2xldG9uKFwidG9hc3RcIik7XG4gIG5vZGUuY2xhc3NMaXN0LnJlbW92ZShcIm5vdGljZVwiKTtcbiAgbm9kZS50ZXh0Q29udGVudCA9IG1lc3NhZ2U7XG4gIG5vZGUuaGlkZGVuID0gZmFsc2U7XG4gIGNsZWFyVGltZW91dCh0b2FzdFRpbWVyKTtcbiAgdG9hc3RUaW1lciA9IHNldFRpbWVvdXQoKCkgPT4ge1xuICAgIG5vZGUuaGlkZGVuID0gdHJ1ZTtcbiAgfSwgbXMpO1xufVxuXG4vKiogQSB0b2FzdCB0aGF0IHN0YXlzIHVudGlsIGFjdGVkIG9uOiBhIG1lc3NhZ2UsIG9uZSBhY3Rpb24gYnV0dG9uLCBhbmQgYSBkaXNtaXNzIFx1MDBENy4gKi9cbmV4cG9ydCBmdW5jdGlvbiBub3RpY2UobWVzc2FnZTogc3RyaW5nLCBhY3Rpb246IHN0cmluZywgb25BY3Rpb246ICgpID0+IHZvaWQpOiB2b2lkIHtcbiAgY29uc3Qgbm9kZSA9IHNpbmdsZXRvbihcInRvYXN0XCIpO1xuICBjbGVhclRpbWVvdXQodG9hc3RUaW1lcik7XG4gIGNvbnN0IGJ1dHRvbiA9IGVsKFwiYnV0dG9uXCIsIFwiYWN0aW9uXCIsIGFjdGlvbik7XG4gIGJ1dHRvbi5hZGRFdmVudExpc3RlbmVyKFwiY2xpY2tcIiwgb25BY3Rpb24pO1xuICBjb25zdCBjbG9zZSA9IGVsKFwiYnV0dG9uXCIsIFwiY2xvc2VcIiwgXCJcdTAwRDdcIik7XG4gIGNsb3NlLnNldEF0dHJpYnV0ZShcImFyaWEtbGFiZWxcIiwgXCJEaXNtaXNzXCIpO1xuICBjbG9zZS5hZGRFdmVudExpc3RlbmVyKFwiY2xpY2tcIiwgKCkgPT4ge1xuICAgIG5vZGUuaGlkZGVuID0gdHJ1ZTtcbiAgfSk7XG4gIG5vZGUucmVwbGFjZUNoaWxkcmVuKGVsKFwic3BhblwiLCBcIlwiLCBtZXNzYWdlKSwgYnV0dG9uLCBjbG9zZSk7XG4gIG5vZGUuY2xhc3NMaXN0LmFkZChcIm5vdGljZVwiKTtcbiAgbm9kZS5oaWRkZW4gPSBmYWxzZTtcbiAgYnV0dG9uLmZvY3VzKCk7XG59XG5cbi8qKiBUaGUgcGVuZGluZy1wcmVmaXggaGludCwgZS5nLiBcIkcgXHUyMDI2IG0gTWFpbCBcdTAwQjcgYyBDYWxlbmRhclwiLiBQYXNzIG51bGwgdG8gaGlkZS4gKi9cbmV4cG9ydCBmdW5jdGlvbiBzaG93SGludCh0ZXh0OiBzdHJpbmcgfCBudWxsKTogdm9pZCB7XG4gIGNvbnN0IG5vZGUgPSBzaW5nbGV0b24oXCJoaW50XCIpO1xuICBub2RlLmhpZGRlbiA9IHRleHQgPT09IG51bGw7XG4gIG5vZGUudGV4dENvbnRlbnQgPSB0ZXh0ID8/IFwiXCI7XG59XG5cbmV4cG9ydCBmdW5jdGlvbiBzaG93SGVhbHRoQmFkZ2UoYnJva2VuOiByZWFkb25seSBzdHJpbmdbXSwgb25EaXNtaXNzOiAoKSA9PiB2b2lkKTogdm9pZCB7XG4gIGNvbnN0IG5vZGUgPSBzaW5nbGV0b24oXCJoZWFsdGhcIik7XG4gIG5vZGUucmVwbGFjZUNoaWxkcmVuKCk7XG4gIGlmIChicm9rZW4ubGVuZ3RoID09PSAwKSB7XG4gICAgbm9kZS5oaWRkZW4gPSB0cnVlO1xuICAgIHJldHVybjtcbiAgfVxuICBub2RlLmhpZGRlbiA9IGZhbHNlO1xuICBub2RlLnRpdGxlID0gYE5vdCBmb3VuZDogJHticm9rZW4uam9pbihcIiwgXCIpfS4gU2VlIHRoZSBjb25zb2xlIGZvciBkZXRhaWxzLmA7XG4gIGNvbnN0IGNsb3NlID0gZWwoXCJidXR0b25cIiwgXCJjbG9zZVwiLCBcIlx1MDBEN1wiKTtcbiAgY2xvc2Uuc2V0QXR0cmlidXRlKFwiYXJpYS1sYWJlbFwiLCBcIkRpc21pc3NcIik7XG4gIGNsb3NlLmFkZEV2ZW50TGlzdGVuZXIoXCJjbGlja1wiLCAoKSA9PiB7XG4gICAgbm9kZS5oaWRkZW4gPSB0cnVlO1xuICAgIG9uRGlzbWlzcygpO1xuICB9KTtcbiAgbm9kZS5hcHBlbmQoZWwoXCJzcGFuXCIsIFwiXCIsIGBPQkI6ICR7YnJva2VuLmxlbmd0aH0gc2VsZWN0b3IocykgYnJva2VuYCksIGNsb3NlKTtcbn1cblxuZXhwb3J0IGZ1bmN0aW9uIHNob3dIZWxwKHJvd3M6IHJlYWRvbmx5IHsga2V5czogc3RyaW5nOyBkZXNjcmlwdGlvbjogc3RyaW5nIH1bXSk6IHZvaWQge1xuICBjb25zdCByb290ID0gb3ZlcmxheVJvb3QoKTtcbiAgY29uc3QgYmFja2Ryb3AgPSBlbChcImRpdlwiLCBcImJhY2tkcm9wXCIpO1xuICBjb25zdCBwYW5lbCA9IGVsKFwiZGl2XCIsIFwicGFuZWwgaGVscFwiKTtcbiAgcGFuZWwuYXBwZW5kKGVsKFwiaDJcIiwgXCJcIiwgXCJLZXlib2FyZCBzaG9ydGN1dHNcIikpO1xuICBjb25zdCB0YWJsZSA9IGVsKFwidGFibGVcIiwgXCJcIik7XG4gIGZvciAoY29uc3QgeyBrZXlzLCBkZXNjcmlwdGlvbiB9IG9mIHJvd3MpIHtcbiAgICBjb25zdCB0ciA9IGVsKFwidHJcIiwgXCJcIik7XG4gICAgY29uc3QgayA9IGVsKFwidGRcIiwgXCJrZXlzXCIpO1xuICAgIGsuYXBwZW5kKGVsKFwia2JkXCIsIFwiXCIsIGtleXMpKTtcbiAgICB0ci5hcHBlbmQoaywgZWwoXCJ0ZFwiLCBcIlwiLCBkZXNjcmlwdGlvbikpO1xuICAgIHRhYmxlLmFwcGVuZCh0cik7XG4gIH1cbiAgcGFuZWwuYXBwZW5kKHRhYmxlKTtcbiAgYmFja2Ryb3AuYXBwZW5kKHBhbmVsKTtcbiAgcm9vdC5hcHBlbmQoYmFja2Ryb3ApO1xuXG4gIGNvbnN0IGNsb3NlID0gKCkgPT4ge1xuICAgIGJhY2tkcm9wLnJlbW92ZSgpO1xuICAgIHdpbmRvdy5yZW1vdmVFdmVudExpc3RlbmVyKFwia2V5ZG93blwiLCBvbktleSwgdHJ1ZSk7XG4gIH07XG4gIGNvbnN0IG9uS2V5ID0gKGU6IEtleWJvYXJkRXZlbnQpID0+IHtcbiAgICBlLnByZXZlbnREZWZhdWx0KCk7XG4gICAgZS5zdG9wUHJvcGFnYXRpb24oKTtcbiAgICBjbG9zZSgpO1xuICB9O1xuICBiYWNrZHJvcC5hZGRFdmVudExpc3RlbmVyKFwiY2xpY2tcIiwgKGUpID0+IHtcbiAgICBpZiAoZS50YXJnZXQgPT09IGJhY2tkcm9wKSBjbG9zZSgpO1xuICB9KTtcbiAgd2luZG93LmFkZEV2ZW50TGlzdGVuZXIoXCJrZXlkb3duXCIsIG9uS2V5LCB0cnVlKTtcbn1cblxuZXhwb3J0IHsgZWwgfTtcbiIsICIvLyBQdXJlIGZ1enp5IG1hdGNoZXIgZm9yIHRoZSBmb2xkZXIgcGFsZXR0ZTogY2FzZS1pbnNlbnNpdGl2ZSBzdWJzZXF1ZW5jZSBtYXRjaGluZywgc2NvcmVkIHRvXG4vLyBmYXZvdXIgd29yZCBzdGFydHMgYW5kIGNvbnNlY3V0aXZlIHJ1bnMsIHRoZSB3YXkgRmFzdG1haWwncyBhbmQgbW9zdCBlZGl0b3JzJyBwaWNrZXJzIGZlZWwuXG5cbmV4cG9ydCBpbnRlcmZhY2UgRnV6enlNYXRjaDxUPiB7XG4gIGl0ZW06IFQ7XG4gIHNjb3JlOiBudW1iZXI7XG4gIC8qKiBJbmRpY2VzIGludG8gdGhlIGl0ZW0ncyB0ZXh0IHRoYXQgbWF0Y2hlZCwgZm9yIGhpZ2hsaWdodGluZy4gKi9cbiAgaW5kaWNlczogbnVtYmVyW107XG59XG5cbmNvbnN0IGlzQm91bmRhcnkgPSAodGV4dDogc3RyaW5nLCBpOiBudW1iZXIpID0+IGkgPT09IDAgfHwgL1tcXHMvXFxcXC5fLV0vLnRlc3QodGV4dFtpIC0gMV0gPz8gXCJcIik7XG5cbmZ1bmN0aW9uIGlzU3Vic2VxdWVuY2UocXVlcnk6IHN0cmluZywgdGV4dDogc3RyaW5nLCBmcm9tOiBudW1iZXIpOiBib29sZWFuIHtcbiAgbGV0IGF0ID0gZnJvbTtcbiAgZm9yIChjb25zdCBjaCBvZiBxdWVyeSkge1xuICAgIGF0ID0gdGV4dC5pbmRleE9mKGNoLCBhdCkgKyAxO1xuICAgIGlmIChhdCA9PT0gMCkgcmV0dXJuIGZhbHNlO1xuICB9XG4gIHJldHVybiB0cnVlO1xufVxuXG4vKiogU2NvcmUgYHF1ZXJ5YCBhZ2FpbnN0IGB0ZXh0YCwgb3IgbnVsbCBpZiBpdCBpcyBub3QgYSBzdWJzZXF1ZW5jZS4gSGlnaGVyIGlzIGJldHRlci4gKi9cbmV4cG9ydCBmdW5jdGlvbiBzY29yZShxdWVyeTogc3RyaW5nLCB0ZXh0OiBzdHJpbmcpOiB7IHNjb3JlOiBudW1iZXI7IGluZGljZXM6IG51bWJlcltdIH0gfCBudWxsIHtcbiAgY29uc3QgcSA9IHF1ZXJ5LnRvTG93ZXJDYXNlKCk7XG4gIGNvbnN0IHQgPSB0ZXh0LnRvTG93ZXJDYXNlKCk7XG4gIGlmIChxLmxlbmd0aCA9PT0gMCkgcmV0dXJuIHsgc2NvcmU6IDAsIGluZGljZXM6IFtdIH07XG5cbiAgY29uc3QgaW5kaWNlczogbnVtYmVyW10gPSBbXTtcbiAgbGV0IHRvdGFsID0gMDtcbiAgbGV0IGZyb20gPSAwO1xuICBmb3IgKGNvbnN0IFtuLCBjaF0gb2YgWy4uLnFdLmVudHJpZXMoKSkge1xuICAgIC8vIFByZWZlciB0aGUgbmV4dCBvY2N1cnJlbmNlIGF0IGEgd29yZCBib3VuZGFyeSwgYXMgbG9uZyBhcyB0aGUgcmVzdCBvZiB0aGUgcXVlcnkgc3RpbGxcbiAgICAvLyBmaXRzIGFmdGVyIGl0OyBvdGhlcndpc2UgdGFrZSB0aGUgbmVhcmVzdCBvbmUuXG4gICAgY29uc3QgcmVzdCA9IHEuc2xpY2UobiArIDEpO1xuICAgIGxldCBhdCA9IC0xO1xuICAgIGZvciAobGV0IGkgPSB0LmluZGV4T2YoY2gsIGZyb20pOyBpID49IDA7IGkgPSB0LmluZGV4T2YoY2gsIGkgKyAxKSkge1xuICAgICAgaWYgKCFpc1N1YnNlcXVlbmNlKHJlc3QsIHQsIGkgKyAxKSkgYnJlYWs7XG4gICAgICBpZiAoYXQgPCAwKSBhdCA9IGk7XG4gICAgICBpZiAoaXNCb3VuZGFyeSh0ZXh0LCBpKSkge1xuICAgICAgICBhdCA9IGk7XG4gICAgICAgIGJyZWFrO1xuICAgICAgfVxuICAgIH1cbiAgICBpZiAoYXQgPCAwKSByZXR1cm4gbnVsbDtcbiAgICBjb25zdCBwcmV2ID0gaW5kaWNlcy5hdCgtMSk7XG4gICAgdG90YWwgKz0gMTtcbiAgICBpZiAoaXNCb3VuZGFyeSh0ZXh0LCBhdCkpIHRvdGFsICs9IDg7XG4gICAgaWYgKHByZXYgIT09IHVuZGVmaW5lZCAmJiBhdCA9PT0gcHJldiArIDEpIHRvdGFsICs9IDU7XG4gICAgaWYgKHByZXYgIT09IHVuZGVmaW5lZCkgdG90YWwgLT0gTWF0aC5taW4oYXQgLSBwcmV2IC0gMSwgMyk7XG4gICAgaW5kaWNlcy5wdXNoKGF0KTtcbiAgICBmcm9tID0gYXQgKyAxO1xuICB9XG4gIC8vIFNob3J0ZXIgdGFyZ2V0cyB3aW4gdGllcywgc28gXCJJbmJveFwiIGJlYXRzIFwiSW5ib3gvUmVjZWlwdHMvMjAxOVwiIGZvciBcImluYlwiLlxuICB0b3RhbCAtPSB0ZXh0Lmxlbmd0aCAqIDAuMDE7XG4gIHJldHVybiB7IHNjb3JlOiB0b3RhbCwgaW5kaWNlcyB9O1xufVxuXG5leHBvcnQgZnVuY3Rpb24gZnV6enlGaWx0ZXI8VD4oXG4gIHF1ZXJ5OiBzdHJpbmcsXG4gIGl0ZW1zOiByZWFkb25seSBUW10sXG4gIHRleHQ6IChpdGVtOiBUKSA9PiBzdHJpbmcsXG4pOiBGdXp6eU1hdGNoPFQ+W10ge1xuICBjb25zdCBtYXRjaGVzOiBGdXp6eU1hdGNoPFQ+W10gPSBbXTtcbiAgZm9yIChjb25zdCBpdGVtIG9mIGl0ZW1zKSB7XG4gICAgY29uc3QgcmVzdWx0ID0gc2NvcmUocXVlcnksIHRleHQoaXRlbSkpO1xuICAgIGlmIChyZXN1bHQpIG1hdGNoZXMucHVzaCh7IGl0ZW0sIC4uLnJlc3VsdCB9KTtcbiAgfVxuICByZXR1cm4gbWF0Y2hlcy5zb3J0KChhLCBiKSA9PiBiLnNjb3JlIC0gYS5zY29yZSk7XG59XG4iLCAiLy8gRm9sZGVyIHBhbGV0dGU6IEZhc3RtYWlsJ3MgYGdgIChnbyB0byBmb2xkZXIpIHR5cGUtYWhlYWQgcGlja2VyLCBwbHVzIE91dGxvb2sncyBhcHBzLlxuLy8gTW92aW5nIChgbWApIHVzZXMgT3V0bG9vaydzIG93biBcIk1vdmUgdG9cIiBwaWNrZXIgZm9yIG5vdzsgc2VlIGRvY3MvZGVjaXNpb25zLm1kLlxuXG5pbXBvcnQgeyB0eXBlIEZvbGRlck5vZGUsIHJlYWRGb2xkZXJUcmVlLCByZXNvbHZlLCB0eXBlIFNlbGVjdG9yTmFtZSB9IGZyb20gXCIuLi9zZWxlY3RvcnMudHNcIjtcbmltcG9ydCB7IGVsLCBvdmVybGF5Um9vdCwgdG9hc3QgfSBmcm9tIFwiLi4vdWkvb3ZlcmxheS50c1wiO1xuaW1wb3J0IHsgdHlwZSBGdXp6eU1hdGNoLCBmdXp6eUZpbHRlciB9IGZyb20gXCIuL2Z1enp5LnRzXCI7XG5cbmNvbnN0IE1BWF9SRVNVTFRTID0gMTI7XG5cbmZ1bmN0aW9uIGhpZ2hsaWdodCh0ZXh0OiBzdHJpbmcsIGluZGljZXM6IHJlYWRvbmx5IG51bWJlcltdKTogRG9jdW1lbnRGcmFnbWVudCB7XG4gIGNvbnN0IGZyYWcgPSBkb2N1bWVudC5jcmVhdGVEb2N1bWVudEZyYWdtZW50KCk7XG4gIGNvbnN0IGhpdCA9IG5ldyBTZXQoaW5kaWNlcyk7XG4gIGZvciAoY29uc3QgW2ksIGNoXSBvZiBbLi4udGV4dF0uZW50cmllcygpKSB7XG4gICAgZnJhZy5hcHBlbmQoaGl0LmhhcyhpKSA/IGVsKFwibWFya1wiLCBcIlwiLCBjaCkgOiBkb2N1bWVudC5jcmVhdGVUZXh0Tm9kZShjaCkpO1xuICB9XG4gIHJldHVybiBmcmFnO1xufVxuXG5jb25zdCBBUFBTOiBbc3RyaW5nLCBTZWxlY3Rvck5hbWVdW10gPSBbXG4gIFtcIk1haWxcIiwgXCJtYWlsQXBwXCJdLFxuICBbXCJDYWxlbmRhclwiLCBcImNhbGVuZGFyQXBwXCJdLFxuICBbXCJQZW9wbGVcIiwgXCJwZW9wbGVBcHBcIl0sXG5dO1xuXG4vKiogT3V0bG9vaydzIGFwcHMsIGxpc3RlZCBhbG9uZ3NpZGUgZm9sZGVycyBzbyBgZ2AgdGhlbiBcImNhbFwiIHN3aXRjaGVzIHRvIENhbGVuZGFyLiAqL1xuZnVuY3Rpb24gYXBwVGFyZ2V0cygpOiBGb2xkZXJOb2RlW10ge1xuICByZXR1cm4gQVBQUy5mbGF0TWFwKChbbmFtZSwgdGFyZ2V0XSkgPT4ge1xuICAgIGNvbnN0IGVsZW1lbnQgPSByZXNvbHZlKHRhcmdldCk7XG4gICAgcmV0dXJuIGVsZW1lbnQgPyBbeyBuYW1lLCBwYXRoOiBuYW1lLCBlbGVtZW50IH1dIDogW107XG4gIH0pO1xufVxuXG5leHBvcnQgZnVuY3Rpb24gb3BlbkZvbGRlclBhbGV0dGUoKTogdm9pZCB7XG4gIGNvbnN0IGZvbGRlcnMgPSBbLi4uYXBwVGFyZ2V0cygpLCAuLi5yZWFkRm9sZGVyVHJlZSgpXTtcbiAgaWYgKGZvbGRlcnMubGVuZ3RoID09PSAwKSB7XG4gICAgdG9hc3QoXCJPQkIgY291bGRuJ3QgZmluZCB0aGUgZm9sZGVyIHBhbmVcIik7XG4gICAgcmV0dXJuO1xuICB9XG5cbiAgY29uc3Qgcm9vdCA9IG92ZXJsYXlSb290KCk7XG4gIGNvbnN0IGJhY2tkcm9wID0gZWwoXCJkaXZcIiwgXCJiYWNrZHJvcFwiKTtcbiAgY29uc3QgcGFuZWwgPSBlbChcImRpdlwiLCBcInBhbmVsIHBhbGV0dGVcIik7XG4gIGNvbnN0IGlucHV0ID0gZWwoXCJpbnB1dFwiLCBcIlwiKTtcbiAgaW5wdXQucGxhY2Vob2xkZXIgPSBcIkdvIHRvIGZvbGRlciwgb3IgTWFpbCwgQ2FsZW5kYXIsIFBlb3BsZVx1MjAyNlwiO1xuICBpbnB1dC5zZXRBdHRyaWJ1dGUoXCJhcmlhLWxhYmVsXCIsIGlucHV0LnBsYWNlaG9sZGVyKTtcbiAgY29uc3QgbGlzdCA9IGVsKFwidWxcIiwgXCJcIik7XG4gIGxpc3Quc2V0QXR0cmlidXRlKFwicm9sZVwiLCBcImxpc3Rib3hcIik7XG4gIHBhbmVsLmFwcGVuZChpbnB1dCwgbGlzdCk7XG4gIGJhY2tkcm9wLmFwcGVuZChwYW5lbCk7XG4gIHJvb3QuYXBwZW5kKGJhY2tkcm9wKTtcblxuICBsZXQgbWF0Y2hlczogRnV6enlNYXRjaDxGb2xkZXJOb2RlPltdID0gW107XG4gIGxldCBzZWxlY3RlZCA9IDA7XG5cbiAgY29uc3QgY2xvc2UgPSAoKSA9PiBiYWNrZHJvcC5yZW1vdmUoKTtcbiAgY29uc3QgY29tbWl0ID0gKCkgPT4ge1xuICAgIGNvbnN0IG1hdGNoID0gbWF0Y2hlc1tzZWxlY3RlZF07XG4gICAgY2xvc2UoKTtcbiAgICBpZiAobWF0Y2gpIChtYXRjaC5pdGVtLmVsZW1lbnQgYXMgSFRNTEVsZW1lbnQpLmNsaWNrKCk7XG4gIH07XG5cbiAgY29uc3QgcmVuZGVyID0gKCkgPT4ge1xuICAgIG1hdGNoZXMgPSBmdXp6eUZpbHRlcihpbnB1dC52YWx1ZSwgZm9sZGVycywgKGYpID0+IGYucGF0aCkuc2xpY2UoMCwgTUFYX1JFU1VMVFMpO1xuICAgIHNlbGVjdGVkID0gTWF0aC5taW4oc2VsZWN0ZWQsIE1hdGgubWF4KG1hdGNoZXMubGVuZ3RoIC0gMSwgMCkpO1xuICAgIGxpc3QucmVwbGFjZUNoaWxkcmVuKFxuICAgICAgLi4ubWF0Y2hlcy5tYXAoKG0sIGkpID0+IHtcbiAgICAgICAgY29uc3QgbGkgPSBlbChcImxpXCIsIFwiXCIpO1xuICAgICAgICBsaS5zZXRBdHRyaWJ1dGUoXCJyb2xlXCIsIFwib3B0aW9uXCIpO1xuICAgICAgICBsaS5zZXRBdHRyaWJ1dGUoXCJhcmlhLXNlbGVjdGVkXCIsIFN0cmluZyhpID09PSBzZWxlY3RlZCkpO1xuICAgICAgICBsaS5hcHBlbmQoaGlnaGxpZ2h0KG0uaXRlbS5wYXRoLCBtLmluZGljZXMpKTtcbiAgICAgICAgbGkuYWRkRXZlbnRMaXN0ZW5lcihcIm1vdXNlZG93blwiLCAoZSkgPT4ge1xuICAgICAgICAgIGUucHJldmVudERlZmF1bHQoKTtcbiAgICAgICAgICBzZWxlY3RlZCA9IGk7XG4gICAgICAgICAgY29tbWl0KCk7XG4gICAgICAgIH0pO1xuICAgICAgICByZXR1cm4gbGk7XG4gICAgICB9KSxcbiAgICApO1xuICAgIGlmIChtYXRjaGVzLmxlbmd0aCA9PT0gMCkgbGlzdC5hcHBlbmQoZWwoXCJsaVwiLCBcImVtcHR5XCIsIFwiTm8gbWF0Y2hpbmcgZm9sZGVyc1wiKSk7XG4gIH07XG5cbiAgaW5wdXQuYWRkRXZlbnRMaXN0ZW5lcihcImlucHV0XCIsICgpID0+IHtcbiAgICBzZWxlY3RlZCA9IDA7XG4gICAgcmVuZGVyKCk7XG4gIH0pO1xuICBpbnB1dC5hZGRFdmVudExpc3RlbmVyKFwia2V5ZG93blwiLCAoZSkgPT4ge1xuICAgIGUuc3RvcFByb3BhZ2F0aW9uKCk7XG4gICAgaWYgKGUua2V5ID09PSBcIkVzY2FwZVwiKSBjbG9zZSgpO1xuICAgIGVsc2UgaWYgKGUua2V5ID09PSBcIkVudGVyXCIpIGNvbW1pdCgpO1xuICAgIGVsc2UgaWYgKGUua2V5ID09PSBcIkFycm93RG93blwiIHx8IChlLmN0cmxLZXkgJiYgZS5rZXkgPT09IFwiblwiKSkgc2VsZWN0ZWQrKztcbiAgICBlbHNlIGlmIChlLmtleSA9PT0gXCJBcnJvd1VwXCIgfHwgKGUuY3RybEtleSAmJiBlLmtleSA9PT0gXCJwXCIpKVxuICAgICAgc2VsZWN0ZWQgPSBNYXRoLm1heChzZWxlY3RlZCAtIDEsIDApO1xuICAgIGVsc2UgcmV0dXJuO1xuICAgIGUucHJldmVudERlZmF1bHQoKTtcbiAgICByZW5kZXIoKTtcbiAgfSk7XG4gIGJhY2tkcm9wLmFkZEV2ZW50TGlzdGVuZXIoXCJtb3VzZWRvd25cIiwgKGUpID0+IHtcbiAgICBpZiAoZS50YXJnZXQgPT09IGJhY2tkcm9wKSBjbG9zZSgpO1xuICB9KTtcblxuICByZW5kZXIoKTtcbiAgaW5wdXQuZm9jdXMoKTtcbn1cbiIsICIvLyBUdXJucyBhbiBhY3Rpb24gaW50byBET00gZWZmZWN0cy4gVGhlIG9ubHkga2V5bWFwIGZpbGUgdGhhdCB0b3VjaGVzIHRoZSBwYWdlLlxuXG5pbXBvcnQgeyBvcGVuRm9sZGVyUGFsZXR0ZSB9IGZyb20gXCIuLi9wYWxldHRlL3BhbGV0dGUudHNcIjtcbmltcG9ydCB7XG4gIGNoZWNrZWRSb3dzLFxuICBjdXJzb3JSb3csXG4gIGlzVW5yZWFkUm93LFxuICByZXNvbHZlLFxuICByZXNvbHZlQWxsLFxuICByb3dDaGVja2JveCxcbiAgc2VsZWN0ZWRSb3dzLFxufSBmcm9tIFwiLi4vc2VsZWN0b3JzLnRzXCI7XG5pbXBvcnQgeyBzaG93SGVscCwgdG9hc3QgfSBmcm9tIFwiLi4vdWkvb3ZlcmxheS50c1wiO1xuaW1wb3J0IHsgQUNUSU9OUywgdHlwZSBBY3Rpb25JZCwgdHlwZSBFeHRlbnNpb25Db21tYW5kIH0gZnJvbSBcIi4vYWN0aW9ucy50c1wiO1xuaW1wb3J0IHsgQklORElOR1MgfSBmcm9tIFwiLi9iaW5kaW5ncy50c1wiO1xuaW1wb3J0IHR5cGUgeyBDb250ZXh0IH0gZnJvbSBcIi4vZW5naW5lLnRzXCI7XG5cbi8qKiBXaGljaCBiaW5kaW5ncyBhcHBseSwganVkZ2VkIGZyb20gd2hlcmUgZm9jdXMgaXMuICovXG5leHBvcnQgZnVuY3Rpb24gY3VycmVudENvbnRleHQoKTogQ29udGV4dCB7XG4gIGNvbnN0IHJlYWRpbmcgPSByZXNvbHZlKFwicmVhZGluZ1BhbmVcIik7XG4gIGNvbnN0IGZvY3VzZWQgPSBkb2N1bWVudC5hY3RpdmVFbGVtZW50O1xuICByZXR1cm4gcmVhZGluZyAmJiBmb2N1c2VkICYmIHJlYWRpbmcuY29udGFpbnMoZm9jdXNlZCkgPyBcInJlYWRpbmdcIiA6IFwibGlzdFwiO1xufVxuXG5leHBvcnQgaW50ZXJmYWNlIEV4ZWN1dGVPcHRpb25zIHtcbiAgLyoqIFRha2Ugb3ZlciB4IGFuZCBqL2sgc28gdGlja3Mgc3Vydml2ZSBtb3ZpbmcgKEZhc3RtYWlsLXN0eWxlKS4gT2ZmOiBPdXRsb29rJ3Mgb3duIGtleXMuICovXG4gIG11bHRpU2VsZWN0OiBib29sZWFuO1xufVxuXG4vKiogTW92ZXMgdGhlIGN1cnNvciBvbmUgcm93IHdpdGhvdXQgc2VsZWN0aW5nIGl0LCBzbyB0aWNrZWQgcm93cyBzdGF5IHRpY2tlZC4gKi9cbmZ1bmN0aW9uIG1vdmVDdXJzb3Ioc3RlcDogMSB8IC0xKTogYm9vbGVhbiB7XG4gIC8vIE5vdGhpbmcgdGlja2VkOiBPdXRsb29rJ3Mgb3duIGovayAod2hpY2ggc2VsZWN0IGFzIHRoZXkgZ28pIGFyZSB3aGF0IHlvdSB3YW50LlxuICBpZiAoY2hlY2tlZFJvd3MoKS5sZW5ndGggPT09IDApIHJldHVybiBmYWxzZTtcbiAgY29uc3Qgcm93cyA9IHJlc29sdmVBbGwoXCJtZXNzYWdlUm93XCIpO1xuICBjb25zdCBhdCA9IGN1cnNvclJvdygpO1xuICBjb25zdCBuZXh0ID0gcm93c1soYXQgPyByb3dzLmluZGV4T2YoYXQpIDogLTEpICsgc3RlcF07XG4gIGlmIChuZXh0IGluc3RhbmNlb2YgSFRNTEVsZW1lbnQpIHtcbiAgICBuZXh0LmZvY3VzKCk7XG4gICAgbmV4dC5zY3JvbGxJbnRvVmlldz8uKHsgYmxvY2s6IFwibmVhcmVzdFwiIH0pO1xuICB9XG4gIHJldHVybiB0cnVlOyAvLyBhdCBlaXRoZXIgZW5kIHRoZSBrZXkgaXMgc3dhbGxvd2VkIHJhdGhlciB0aGFuIGNvbGxhcHNpbmcgdGhlIHNlbGVjdGlvblxufVxuXG4vKiogVGlja3Mgb3IgdW50aWNrcyB0aGUgY3Vyc29yIHJvdywgYXMgY2xpY2tpbmcgaXRzIHRpY2sgYm94IGRvZXMuICovXG5mdW5jdGlvbiB0b2dnbGVTZWxlY3QoKTogYm9vbGVhbiB7XG4gIGNvbnN0IHJvdyA9IGN1cnNvclJvdygpO1xuICBjb25zdCBib3ggPSByb3cgJiYgcm93Q2hlY2tib3gocm93KTtcbiAgaWYgKCFib3gpIHJldHVybiBmYWxzZTtcbiAgYm94LmNsaWNrKCk7XG4gIHJldHVybiB0cnVlO1xufVxuXG4vKiogUmV0dXJucyBmYWxzZSB3aGVuIHRoZSBrZXkgc2hvdWxkIGJlIGxlZnQgZm9yIE91dGxvb2suICovXG5mdW5jdGlvbiBydW5FeHRlbnNpb24oY29tbWFuZDogRXh0ZW5zaW9uQ29tbWFuZCwgb3B0aW9uczogRXhlY3V0ZU9wdGlvbnMpOiBib29sZWFuIHtcbiAgc3dpdGNoIChjb21tYW5kKSB7XG4gICAgY2FzZSBcImhlbHBcIjpcbiAgICAgIHNob3dIZWxwKFxuICAgICAgICBCSU5ESU5HUy5tYXAoKGIpID0+ICh7XG4gICAgICAgICAga2V5czogYi5rZXlzLmpvaW4oXCIgXCIpLFxuICAgICAgICAgIGRlc2NyaXB0aW9uOiBBQ1RJT05TW2IuYWN0aW9uXS5kZXNjcmlwdGlvbixcbiAgICAgICAgfSkpLFxuICAgICAgKTtcbiAgICAgIHJldHVybiB0cnVlO1xuICAgIGNhc2UgXCJwYWxldHRlLWdvXCI6XG4gICAgICBvcGVuRm9sZGVyUGFsZXR0ZSgpO1xuICAgICAgcmV0dXJuIHRydWU7XG4gICAgY2FzZSBcInRvZ2dsZS1zZWxlY3RcIjpcbiAgICAgIHJldHVybiBvcHRpb25zLm11bHRpU2VsZWN0ICYmIHRvZ2dsZVNlbGVjdCgpO1xuICAgIGNhc2UgXCJjdXJzb3ItbmV4dFwiOlxuICAgICAgcmV0dXJuIG9wdGlvbnMubXVsdGlTZWxlY3QgJiYgbW92ZUN1cnNvcigxKTtcbiAgICBjYXNlIFwiY3Vyc29yLXByZXZpb3VzXCI6XG4gICAgICByZXR1cm4gb3B0aW9ucy5tdWx0aVNlbGVjdCAmJiBtb3ZlQ3Vyc29yKC0xKTtcbiAgfVxufVxuXG4vKiogUnVuIGFuIGFjdGlvbi4gUmV0dXJucyBmYWxzZSB3aGVuIHRoZSBrZXkgc2hvdWxkIGJlIGxlZnQgZm9yIE91dGxvb2sgKHBhc3N0aHJvdWdoKS4gKi9cbmV4cG9ydCBmdW5jdGlvbiBleGVjdXRlKGlkOiBBY3Rpb25JZCwgb3B0aW9uczogRXhlY3V0ZU9wdGlvbnMgPSB7IG11bHRpU2VsZWN0OiB0cnVlIH0pOiBib29sZWFuIHtcbiAgY29uc3QgeyBzdHJhdGVneSwgZGVzY3JpcHRpb24gfSA9IEFDVElPTlNbaWRdO1xuICBzd2l0Y2ggKHN0cmF0ZWd5LmtpbmQpIHtcbiAgICBjYXNlIFwicGFzc3Rocm91Z2hcIjpcbiAgICAgIHJldHVybiBmYWxzZTtcbiAgICBjYXNlIFwiZHJpdmVcIjoge1xuICAgICAgY29uc3QgdGFyZ2V0ID0gcmVzb2x2ZShzdHJhdGVneS50YXJnZXQpO1xuICAgICAgaWYgKHRhcmdldCBpbnN0YW5jZW9mIEhUTUxFbGVtZW50KSB0YXJnZXQuY2xpY2soKTtcbiAgICAgIGVsc2UgdG9hc3QoYCR7ZGVzY3JpcHRpb259OiBjb3VsZG4ndCBmaW5kIHRoZSBidXR0b25gKTtcbiAgICAgIHJldHVybiB0cnVlO1xuICAgIH1cbiAgICBjYXNlIFwidG9nZ2xlXCI6IHtcbiAgICAgIGNvbnN0IHJvd3MgPSBzZWxlY3RlZFJvd3MoKTtcbiAgICAgIGlmIChyb3dzLmxlbmd0aCA9PT0gMCkge1xuICAgICAgICB0b2FzdChgJHtkZXNjcmlwdGlvbn06IHNlbGVjdCBhIG1lc3NhZ2UgZmlyc3RgKTtcbiAgICAgICAgcmV0dXJuIHRydWU7XG4gICAgICB9XG4gICAgICBjb25zdCB3YW50VW5yZWFkID0gc3RyYXRlZ3kud2FudCA9PT0gXCJ1bnJlYWRcIjtcbiAgICAgIGlmIChyb3dzLmV2ZXJ5KChyb3cpID0+IGlzVW5yZWFkUm93KHJvdykgPT09IHdhbnRVbnJlYWQpKSByZXR1cm4gdHJ1ZTtcbiAgICAgIGNvbnN0IHRhcmdldCA9IHJlc29sdmUoc3RyYXRlZ3kudGFyZ2V0KTtcbiAgICAgIGlmICh0YXJnZXQgaW5zdGFuY2VvZiBIVE1MRWxlbWVudCkgdGFyZ2V0LmNsaWNrKCk7XG4gICAgICBlbHNlIHRvYXN0KGAke2Rlc2NyaXB0aW9ufTogY291bGRuJ3QgZmluZCB0aGUgYnV0dG9uYCk7XG4gICAgICByZXR1cm4gdHJ1ZTtcbiAgICB9XG4gICAgY2FzZSBcImV4dGVuc2lvblwiOlxuICAgICAgcmV0dXJuIHJ1bkV4dGVuc2lvbihzdHJhdGVneS5jb21tYW5kLCBvcHRpb25zKTtcbiAgICBjYXNlIFwidG9kb1wiOlxuICAgICAgdG9hc3QoYCR7ZGVzY3JpcHRpb259IGlzbid0IGltcGxlbWVudGVkIHlldGApO1xuICAgICAgcmV0dXJuIHRydWU7XG4gIH1cbn1cbiIsICIvLyBVc2VyIHByZWZlcmVuY2VzLiBTdG9yZWQgaW4gY2hyb21lLnN0b3JhZ2Uuc3luYzsgbm90aGluZyBlbHNlIGlzIGV2ZXIgcGVyc2lzdGVkLlxuLy9cbi8vIEVhY2ggZW5hYmxlZCBmZWF0dXJlIGlkIGlzIHdyaXR0ZW4gYXMgYSB0b2tlbiBvbiA8aHRtbCBkYXRhLW92ZXJsb29rLWZlYXR1cmVzPVwiLi4uXCI+LCBhbmQgZWFjaFxuLy8gY2hvaWNlIGFzIGl0cyBvd24gYXR0cmlidXRlICg8aHRtbCBkYXRhLW92ZXJsb29rLXBhbGV0dGU9XCJzb2xhcml6ZWRcIj4pLCBzbyBzdHlsZXNoZWV0cyBnYXRlXG4vLyB0aGVtc2VsdmVzIHdpdGggYFtkYXRhLW92ZXJsb29rLWZlYXR1cmVzfj1cImRlbnNpdHlcIl1gIG9yIGBbZGF0YS1vdmVybG9vay1wYWxldHRlPVwic29sYXJpemVkXCJdYC5cblxuZXhwb3J0IHR5cGUgRmVhdHVyZUlkID0ga2V5b2YgdHlwZW9mIEZFQVRVUkVTO1xuZXhwb3J0IHR5cGUgQ2hvaWNlSWQgPSBrZXlvZiB0eXBlb2YgQ0hPSUNFUztcbmV4cG9ydCB0eXBlIFNldHRpbmdzID0gUmVjb3JkPEZlYXR1cmVJZCwgYm9vbGVhbj4gJiB7IFtLIGluIENob2ljZUlkXTogQ2hvaWNlVmFsdWU8Sz4gfTtcbmV4cG9ydCB0eXBlIENob2ljZVZhbHVlPEsgZXh0ZW5kcyBDaG9pY2VJZD4gPSBrZXlvZiAodHlwZW9mIENIT0lDRVMpW0tdW1wib3B0aW9uc1wiXTtcblxudHlwZSBHcm91cCA9IFwiR2VuZXJhbFwiIHwgXCJLZXlib2FyZFwiIHwgXCJEZWNsdXR0ZXJcIiB8IFwiTG9va1wiO1xuXG5pbnRlcmZhY2UgRmVhdHVyZSB7XG4gIGxhYmVsOiBzdHJpbmc7XG4gIGdyb3VwOiBHcm91cDtcbiAgZGVmYXVsdDogYm9vbGVhbjtcbn1cblxuaW50ZXJmYWNlIENob2ljZSB7XG4gIGxhYmVsOiBzdHJpbmc7XG4gIGdyb3VwOiBHcm91cDtcbiAgb3B0aW9uczogUmVjb3JkPHN0cmluZywgc3RyaW5nPjtcbiAgZGVmYXVsdDogc3RyaW5nO1xufVxuXG5leHBvcnQgY29uc3QgRkVBVFVSRVMgPSB7XG4gIGhlYWx0aEluZGljYXRvcjoge1xuICAgIGxhYmVsOiBcIlNob3cgYW4gaW5kaWNhdG9yIHdoZW4gT3V0bG9vaydzIHBhZ2UgY2hhbmdlcyBicmVhayBPQkJcIixcbiAgICBncm91cDogXCJHZW5lcmFsXCIsXG4gICAgZGVmYXVsdDogdHJ1ZSxcbiAgfSxcbiAga2V5bWFwOiB7IGxhYmVsOiBcIkZhc3RtYWlsLXN0eWxlIGtleWJvYXJkIHNob3J0Y3V0c1wiLCBncm91cDogXCJLZXlib2FyZFwiLCBkZWZhdWx0OiB0cnVlIH0sXG4gIG11bHRpU2VsZWN0OiB7XG4gICAgbGFiZWw6IFwieCB0aWNrcyBhIG1lc3NhZ2U7IGogYW5kIGsgbW92ZSB3aXRob3V0IGxvc2luZyB0aWNrc1wiLFxuICAgIGdyb3VwOiBcIktleWJvYXJkXCIsXG4gICAgZGVmYXVsdDogdHJ1ZSxcbiAgfSxcbiAgXCJoaWRlLXVwc2VsbFwiOiB7IGxhYmVsOiBcIlVwZ3JhZGUgYW5kICd0cnkgdGhlIG5ld1x1MjAyNicgYmFubmVyc1wiLCBncm91cDogXCJEZWNsdXR0ZXJcIiwgZGVmYXVsdDogdHJ1ZSB9LFxuICBcImhpZGUtYXBwLWJhbm5lcnNcIjoge1xuICAgIGxhYmVsOiBcIk5vdGljZXMgYWNyb3NzIHRoZSB0b3Agb2YgT3V0bG9vayAocG9wLXVwLCBzaWduLWluIGFuZCBzdG9yYWdlIHJlbWluZGVycylcIixcbiAgICBncm91cDogXCJEZWNsdXR0ZXJcIixcbiAgICBkZWZhdWx0OiBmYWxzZSxcbiAgfSxcbiAgXCJoaWRlLWluZm8tYmFyc1wiOiB7XG4gICAgbGFiZWw6IFwiTm90aWNlcyBhYm92ZSBhbiBvcGVuIG1lc3NhZ2UgKCdZb3UgcmVwbGllZCBvblx1MjAyNicsICdFeHRyYSBsaW5lIGJyZWFrcyByZW1vdmVkJylcIixcbiAgICBncm91cDogXCJEZWNsdXR0ZXJcIixcbiAgICBkZWZhdWx0OiBmYWxzZSxcbiAgfSxcbiAgXCJoaWRlLWNvcGlsb3RcIjogeyBsYWJlbDogXCJDb3BpbG90IGJ1dHRvbnMgYW5kIHBhbmVsc1wiLCBncm91cDogXCJEZWNsdXR0ZXJcIiwgZGVmYXVsdDogdHJ1ZSB9LFxuICBcImhpZGUtaGVhZGVyLWV4dHJhc1wiOiB7XG4gICAgbGFiZWw6IFwiVGVhbXMgYW5kIE15IERheSBidXR0b25zIGluIHRoZSBoZWFkZXJcIixcbiAgICBncm91cDogXCJEZWNsdXR0ZXJcIixcbiAgICBkZWZhdWx0OiB0cnVlLFxuICB9LFxuICBcImhpZGUtcmFpbC1leHRyYXNcIjoge1xuICAgIGxhYmVsOiBcIlRvIERvLCBOZXdzbGV0dGVycywgT3JnIEV4cGxvcmVyIGFuZCBPbmVEcml2ZSBpbiB0aGUgYXBwIHJhaWxcIixcbiAgICBncm91cDogXCJEZWNsdXR0ZXJcIixcbiAgICBkZWZhdWx0OiB0cnVlLFxuICB9LFxuICBcImhpZGUtcmliYm9uLWV4dHJhc1wiOiB7XG4gICAgbGFiZWw6IFwiVGVhbXMsIFN3ZWVwLCBRdWljayBzdGVwcywgVml2YSBhbmQgYWRkLWluIGJ1dHRvbnMgaW4gdGhlIHJpYmJvblwiLFxuICAgIGdyb3VwOiBcIkRlY2x1dHRlclwiLFxuICAgIGRlZmF1bHQ6IHRydWUsXG4gIH0sXG4gIFwiaGlkZS1leHRlcm5hbC10YWdcIjogeyBsYWJlbDogXCInRXh0ZXJuYWwnIHNlbmRlciB0YWdzXCIsIGdyb3VwOiBcIkRlY2x1dHRlclwiLCBkZWZhdWx0OiB0cnVlIH0sXG4gIFwiaGlkZS1yb3ctYWN0aW9uc1wiOiB7XG4gICAgbGFiZWw6IFwiRmxhZyBhbmQgcGluIGljb25zIHRoYXQgYXBwZWFyIHdoZW4gaG92ZXJpbmcgYSBtZXNzYWdlXCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogZmFsc2UsXG4gIH0sXG4gIFwiaGlkZS1hdmF0YXJzXCI6IHtcbiAgICBsYWJlbDogXCJTZW5kZXIgcGljdHVyZXMgaW4gdGhlIG1lc3NhZ2UgbGlzdFwiLFxuICAgIGdyb3VwOiBcIkRlY2x1dHRlclwiLFxuICAgIGRlZmF1bHQ6IGZhbHNlLFxuICB9LFxuICBcImhpZGUtZm9jdXNlZC10YWJzXCI6IHsgbGFiZWw6IFwiRm9jdXNlZCAvIE90aGVyIHRhYnNcIiwgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsIGRlZmF1bHQ6IGZhbHNlIH0sXG4gIFwiaGlkZS1yaWJib24tdGFic1wiOiB7XG4gICAgbGFiZWw6IFwiUmliYm9uIHRhYnMgKEhvbWUsIFZpZXcsIEhlbHApXCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogZmFsc2UsXG4gIH0sXG4gIFwiaGlkZS1hcHAtcmFpbFwiOiB7IGxhYmVsOiBcIkFwcCByYWlsIChsZWZ0IGljb24gYmFyKVwiLCBncm91cDogXCJEZWNsdXR0ZXJcIiwgZGVmYXVsdDogZmFsc2UgfSxcbiAgXCJoaWRlLXJpYmJvblwiOiB7IGxhYmVsOiBcIlJpYmJvblwiLCBncm91cDogXCJEZWNsdXR0ZXJcIiwgZGVmYXVsdDogZmFsc2UgfSxcbiAgZGVuc2l0eTogeyBsYWJlbDogXCJUaWdodGVyLCBkZW5zZXIgbGF5b3V0XCIsIGdyb3VwOiBcIkxvb2tcIiwgZGVmYXVsdDogdHJ1ZSB9LFxufSBhcyBjb25zdCBzYXRpc2ZpZXMgUmVjb3JkPHN0cmluZywgRmVhdHVyZT47XG5cbmV4cG9ydCBjb25zdCBDSE9JQ0VTID0ge1xuICBwYWxldHRlOiB7XG4gICAgbGFiZWw6IFwiQ29sb3JzXCIsXG4gICAgZ3JvdXA6IFwiTG9va1wiLFxuICAgIG9wdGlvbnM6IHtcbiAgICAgIG91dGxvb2s6IFwiT3V0bG9vaydzIG93blwiLFxuICAgICAgc29sYXJpemVkOiBcIlNvbGFyaXplZFwiLFxuICAgICAgYmFzZTE2OiBcIkJhc2UxNiBEZWZhdWx0XCIsXG4gICAgICBmYXN0bWFpbDogXCJGYXN0bWFpbFwiLFxuICAgICAgcGFwZXI6IFwiUGFwZXJcIixcbiAgICB9LFxuICAgIGRlZmF1bHQ6IFwib3V0bG9va1wiLFxuICB9LFxuICBtb2RlOiB7XG4gICAgbGFiZWw6IFwiTGlnaHQgb3IgZGFya1wiLFxuICAgIGdyb3VwOiBcIkxvb2tcIixcbiAgICBvcHRpb25zOiB7IGF1dG86IFwiRm9sbG93IE91dGxvb2tcIiwgbGlnaHQ6IFwiTGlnaHRcIiwgZGFyazogXCJEYXJrXCIgfSxcbiAgICBkZWZhdWx0OiBcImF1dG9cIixcbiAgfSxcbiAgbGF5b3V0OiB7XG4gICAgbGFiZWw6IFwiTGF5b3V0XCIsXG4gICAgZ3JvdXA6IFwiTG9va1wiLFxuICAgIG9wdGlvbnM6IHtcbiAgICAgIG91dGxvb2s6IFwiT3V0bG9vaydzIG93blwiLFxuICAgICAgY2FsbTogXCJDYWxtIHRocmVlLXBhbmVcIixcbiAgICAgIGZhc3RtYWlsOiBcIkZhc3RtYWlsIChsaXN0IGZpcnN0KVwiLFxuICAgICAgZm9jdXM6IFwiRm9jdXMgKHNpbmdsZSBjb2x1bW4pXCIsXG4gICAgfSxcbiAgICBkZWZhdWx0OiBcIm91dGxvb2tcIixcbiAgfSxcbn0gYXMgY29uc3Qgc2F0aXNmaWVzIFJlY29yZDxzdHJpbmcsIENob2ljZT47XG5cbmV4cG9ydCBjb25zdCBGRUFUVVJFX0lEUyA9IE9iamVjdC5rZXlzKEZFQVRVUkVTKSBhcyBGZWF0dXJlSWRbXTtcbmV4cG9ydCBjb25zdCBDSE9JQ0VfSURTID0gT2JqZWN0LmtleXMoQ0hPSUNFUykgYXMgQ2hvaWNlSWRbXTtcblxuZXhwb3J0IGZ1bmN0aW9uIGRlZmF1bHRTZXR0aW5ncygpOiBTZXR0aW5ncyB7XG4gIHJldHVybiB7XG4gICAgLi4uT2JqZWN0LmZyb21FbnRyaWVzKEZFQVRVUkVfSURTLm1hcCgoaWQpID0+IFtpZCwgRkVBVFVSRVNbaWRdLmRlZmF1bHRdKSksXG4gICAgLi4uT2JqZWN0LmZyb21FbnRyaWVzKENIT0lDRV9JRFMubWFwKChpZCkgPT4gW2lkLCBDSE9JQ0VTW2lkXS5kZWZhdWx0XSkpLFxuICB9IGFzIFNldHRpbmdzO1xufVxuXG4vKiogTWVyZ2Ugd2hhdGV2ZXIgaXMgc3RvcmVkIG92ZXIgdGhlIGRlZmF1bHRzLCBpZ25vcmluZyB1bmtub3duIG9yIG1pc3R5cGVkIGtleXMuICovXG5leHBvcnQgZnVuY3Rpb24gbm9ybWFsaXplKHN0b3JlZDogUmVjb3JkPHN0cmluZywgdW5rbm93bj4pOiBTZXR0aW5ncyB7XG4gIGNvbnN0IHNldHRpbmdzID0gZGVmYXVsdFNldHRpbmdzKCkgYXMgUmVjb3JkPHN0cmluZywgdW5rbm93bj47XG4gIGZvciAoY29uc3QgaWQgb2YgRkVBVFVSRV9JRFMpIHtcbiAgICBpZiAodHlwZW9mIHN0b3JlZFtpZF0gPT09IFwiYm9vbGVhblwiKSBzZXR0aW5nc1tpZF0gPSBzdG9yZWRbaWRdO1xuICB9XG4gIGZvciAoY29uc3QgaWQgb2YgQ0hPSUNFX0lEUykge1xuICAgIGNvbnN0IHZhbHVlID0gc3RvcmVkW2lkXTtcbiAgICBpZiAodHlwZW9mIHZhbHVlID09PSBcInN0cmluZ1wiICYmIHZhbHVlIGluIENIT0lDRVNbaWRdLm9wdGlvbnMpIHNldHRpbmdzW2lkXSA9IHZhbHVlO1xuICB9XG4gIHJldHVybiBzZXR0aW5ncyBhcyBTZXR0aW5ncztcbn1cblxuZXhwb3J0IGZ1bmN0aW9uIGZlYXR1cmVUb2tlbnMoc2V0dGluZ3M6IFNldHRpbmdzKTogc3RyaW5nIHtcbiAgcmV0dXJuIEZFQVRVUkVfSURTLmZpbHRlcigoaWQpID0+IHNldHRpbmdzW2lkXSkuam9pbihcIiBcIik7XG59XG5cbi8qKiBXcml0ZXMgdGhlIHNldHRpbmdzIG9udG8gPGh0bWw+IGFzIHRoZSBhdHRyaWJ1dGVzIHRoZSBzdHlsZXNoZWV0cyBrZXkgb2ZmLiAqL1xuZXhwb3J0IGZ1bmN0aW9uIGFwcGx5VG9Eb2N1bWVudChzZXR0aW5nczogU2V0dGluZ3MsIHJvb3Q6IEhUTUxFbGVtZW50KTogdm9pZCB7XG4gIHJvb3Quc2V0QXR0cmlidXRlKFwiZGF0YS1vdmVybG9vay1mZWF0dXJlc1wiLCBmZWF0dXJlVG9rZW5zKHNldHRpbmdzKSk7XG4gIGZvciAoY29uc3QgaWQgb2YgQ0hPSUNFX0lEUykgcm9vdC5zZXRBdHRyaWJ1dGUoYGRhdGEtb3Zlcmxvb2stJHtpZH1gLCBzZXR0aW5nc1tpZF0pO1xufVxuXG5leHBvcnQgYXN5bmMgZnVuY3Rpb24gbG9hZFNldHRpbmdzKCk6IFByb21pc2U8U2V0dGluZ3M+IHtcbiAgcmV0dXJuIG5vcm1hbGl6ZShhd2FpdCBjaHJvbWUuc3RvcmFnZS5zeW5jLmdldChudWxsKSk7XG59XG5cbmV4cG9ydCBhc3luYyBmdW5jdGlvbiBzYXZlU2V0dGluZyhcbiAgaWQ6IEZlYXR1cmVJZCB8IENob2ljZUlkLFxuICB2YWx1ZTogYm9vbGVhbiB8IHN0cmluZyxcbik6IFByb21pc2U8dm9pZD4ge1xuICBhd2FpdCBjaHJvbWUuc3RvcmFnZS5zeW5jLnNldCh7IFtpZF06IHZhbHVlIH0pO1xufVxuXG5leHBvcnQgZnVuY3Rpb24gb25TZXR0aW5nc0NoYW5nZWQoY2FsbGJhY2s6IChzZXR0aW5nczogU2V0dGluZ3MpID0+IHZvaWQpOiB2b2lkIHtcbiAgY2hyb21lLnN0b3JhZ2Uub25DaGFuZ2VkLmFkZExpc3RlbmVyKChfY2hhbmdlcywgYXJlYSkgPT4ge1xuICAgIGlmIChhcmVhID09PSBcInN5bmNcIikgdm9pZCBsb2FkU2V0dGluZ3MoKS50aGVuKGNhbGxiYWNrKTtcbiAgfSk7XG59XG4iLCAiLy8gRW50cnkgcG9pbnQ6IGFwcGxpZXMgc2V0dGluZ3MsIHRhZ3MgZWxlbWVudHMgZm9yIENTUywgcnVucyB0aGUgaGVhbHRoIGNoZWNrLCBtb3VudHMgdGhlIGtleW1hcC5cblxuaW1wb3J0IHsgQ0FQVFVSRV9NRVNTQUdFLCBjYXB0dXJlUGFnZSB9IGZyb20gXCIuL2NhcHR1cmUudHNcIjtcbmltcG9ydCB7IEFDVElPTlMgfSBmcm9tIFwiLi9rZXltYXAvYWN0aW9ucy50c1wiO1xuaW1wb3J0IHsgQklORElOR1MgfSBmcm9tIFwiLi9rZXltYXAvYmluZGluZ3MudHNcIjtcbmltcG9ydCB7IERFRkFVTFRfVElNRU9VVF9NUywga2V5RnJvbUV2ZW50LCBTZXF1ZW5jZUVuZ2luZSB9IGZyb20gXCIuL2tleW1hcC9lbmdpbmUudHNcIjtcbmltcG9ydCB7IGN1cnJlbnRDb250ZXh0LCBleGVjdXRlIH0gZnJvbSBcIi4va2V5bWFwL2V4ZWN1dGUudHNcIjtcbmltcG9ydCB7XG4gIGJyb2tlbkVudHJpZXMsXG4gIHR5cGUgSGVhbHRoUmVzdWx0LFxuICBoZWFsdGhDaGVjayxcbiAgb3V0bG9va0NvbG9yU2NoZW1lLFxuICByZXNvbHZlLFxuICB0YWdFbGVtZW50cyxcbn0gZnJvbSBcIi4vc2VsZWN0b3JzLnRzXCI7XG5pbXBvcnQgeyBhcHBseVRvRG9jdW1lbnQsIGxvYWRTZXR0aW5ncywgb25TZXR0aW5nc0NoYW5nZWQsIHR5cGUgU2V0dGluZ3MgfSBmcm9tIFwiLi9zZXR0aW5ncy50c1wiO1xuaW1wb3J0IHsgaXNGcm9tT3ZlcmxheSwgaXNNb2RhbE9wZW4sIG5vdGljZSwgc2hvd0hlYWx0aEJhZGdlLCBzaG93SGludCB9IGZyb20gXCIuL3VpL292ZXJsYXkudHNcIjtcblxuY29uc3QgTE9HID0gXCJbb2JiXVwiO1xuY29uc3QgQVBQX1JFQURZX1RJTUVPVVRfTVMgPSAyMF8wMDA7XG5jb25zdCBIRUFMVEhfU0VUVExFX01TID0gMV8wMDA7XG4vKiogV2FpdHMgYmV0d2VlbiByZS1jaGVja3Mgb2YgYSBmYWlsZWQgaGVhbHRoIGNoZWNrIGJlZm9yZSBpdCBpcyByZXBvcnRlZC4gKi9cbmNvbnN0IEhFQUxUSF9SRVRSWV9NUyA9IFszXzAwMCwgMTBfMDAwXTtcblxubGV0IHNldHRpbmdzOiBTZXR0aW5ncztcbmxldCBoZWFsdGhEaXNtaXNzZWQgPSBmYWxzZTtcblxuZnVuY3Rpb24gYXBwbHlTZXR0aW5ncyhuZXh0OiBTZXR0aW5ncyk6IHZvaWQge1xuICBzZXR0aW5ncyA9IG5leHQ7XG4gIGFwcGx5VG9Eb2N1bWVudChzZXR0aW5ncywgZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50KTtcbiAgc3luY0NvbG9yU2NoZW1lKCk7XG4gIGlmICghc2V0dGluZ3MuaGVhbHRoSW5kaWNhdG9yKSBzaG93SGVhbHRoQmFkZ2UoW10sICgpID0+IHt9KTtcbn1cblxuLyoqIFJlY29yZHMgd2hldGhlciBPdXRsb29rIHNob3dzIGEgcmVhZGluZyBwYW5lLCBzbyBsYXlvdXRzIG9ubHkgcmVzaGFwZSBwYW5lcyB3aGVuIGl0IGRvZXNuJ3QuICovXG5mdW5jdGlvbiBzeW5jUmVhZGluZ1BhbmUoKTogdm9pZCB7XG4gIGNvbnN0IHN0YXRlID0gcmVzb2x2ZShcInJlYWRpbmdQYW5lXCIpID8gXCJzaG93blwiIDogXCJub25lXCI7XG4gIGlmIChkb2N1bWVudC5kb2N1bWVudEVsZW1lbnQuZGF0YXNldC5vdmVybG9va1JlYWRpbmdQYW5lICE9PSBzdGF0ZSkge1xuICAgIGRvY3VtZW50LmRvY3VtZW50RWxlbWVudC5kYXRhc2V0Lm92ZXJsb29rUmVhZGluZ1BhbmUgPSBzdGF0ZTtcbiAgfVxufVxuXG4vKiogTWlycm9ycyBPdXRsb29rJ3Mgb3duIGxpZ2h0L2RhcmsgbW9kZSBvbnRvIDxodG1sPiBzbyBcIkZvbGxvdyBPdXRsb29rXCIgcGFsZXR0ZXMgY2FuIG1hdGNoIGl0LiAqL1xuZnVuY3Rpb24gc3luY0NvbG9yU2NoZW1lKCk6IHZvaWQge1xuICBjb25zdCBzY2hlbWUgPSBvdXRsb29rQ29sb3JTY2hlbWUoZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50KTtcbiAgaWYgKGRvY3VtZW50LmRvY3VtZW50RWxlbWVudC5nZXRBdHRyaWJ1dGUoXCJkYXRhLW92ZXJsb29rLW91dGxvb2stc2NoZW1lXCIpICE9PSBzY2hlbWUpIHtcbiAgICBkb2N1bWVudC5kb2N1bWVudEVsZW1lbnQuc2V0QXR0cmlidXRlKFwiZGF0YS1vdmVybG9vay1vdXRsb29rLXNjaGVtZVwiLCBzY2hlbWUpO1xuICB9XG59XG5cbmZ1bmN0aW9uIHJlcG9ydEhlYWx0aChyZXN1bHRzOiBIZWFsdGhSZXN1bHRbXSk6IHZvaWQge1xuICBjb25zdCBicm9rZW4gPSBicm9rZW5FbnRyaWVzKHJlc3VsdHMpO1xuICBjb25zdCBmYWxsYmFja3MgPSByZXN1bHRzLmZpbHRlcigocikgPT4gci5jYW5kaWRhdGUgPiAwKS5tYXAoKHIpID0+IHIubmFtZSk7XG4gIGRvY3VtZW50LmRvY3VtZW50RWxlbWVudC5kYXRhc2V0Lm92ZXJsb29rSGVhbHRoID0gYnJva2VuLmxlbmd0aCA/IFwiYnJva2VuXCIgOiBcIm9rXCI7XG4gIGNvbnNvbGUuZ3JvdXBDb2xsYXBzZWQoYCR7TE9HfSBoZWFsdGg6ICR7YnJva2VuLmxlbmd0aCA/IGAke2Jyb2tlbi5sZW5ndGh9IGJyb2tlbmAgOiBcIm9rXCJ9YCk7XG4gIGNvbnNvbGUudGFibGUocmVzdWx0cyk7XG4gIGNvbnNvbGUuZ3JvdXBFbmQoKTtcbiAgaWYgKGJyb2tlbi5sZW5ndGgpIGNvbnNvbGUud2FybihgJHtMT0d9IHNlbGVjdG9ycyBub3QgZm91bmQ6YCwgYnJva2VuLmpvaW4oXCIsIFwiKSk7XG4gIGlmIChmYWxsYmFja3MubGVuZ3RoKSBjb25zb2xlLmluZm8oYCR7TE9HfSB1c2luZyBmYWxsYmFjayBzZWxlY3RvcnMgZm9yOmAsIGZhbGxiYWNrcy5qb2luKFwiLCBcIikpO1xuICBpZiAoc2V0dGluZ3MuaGVhbHRoSW5kaWNhdG9yICYmICFoZWFsdGhEaXNtaXNzZWQpIHtcbiAgICBzaG93SGVhbHRoQmFkZ2UoYnJva2VuLCAoKSA9PiB7XG4gICAgICBoZWFsdGhEaXNtaXNzZWQgPSB0cnVlO1xuICAgIH0pO1xuICB9XG59XG5cbmxldCBoZWFsdGhUaW1lcjogUmV0dXJuVHlwZTx0eXBlb2Ygc2V0VGltZW91dD4gfCB1bmRlZmluZWQ7XG5cbi8qKlxuICogQ2hlY2tzIHRoZSByZWdpc3RyeSBvbmNlIE91dGxvb2sgaGFzIHNldHRsZWQuIFBhcnRzIG9mIHRoZSBwYWdlICh0aGUgYXBwIHJhaWwsIGZvciBvbmUpIHJlbmRlclxuICogYWZ0ZXIgdGhlIG1lc3NhZ2UgbGlzdCwgc28gYSBtaXNzIGlzIHJlLWNoZWNrZWQgYSBmZXcgdGltZXMgYmVmb3JlIGl0IGlzIHJlcG9ydGVkLlxuICovXG5mdW5jdGlvbiBzY2hlZHVsZUhlYWx0aENoZWNrKCk6IHZvaWQge1xuICBjbGVhclRpbWVvdXQoaGVhbHRoVGltZXIpO1xuICAvLyBPbmx5IHRoZSBtYWlsIHZpZXdzIGFyZSBzdXJ2ZXllZDsgQ2FsZW5kYXIgYW5kIFBlb3BsZSBsZWdpdGltYXRlbHkgbGFjayBtYWlsIGVsZW1lbnRzLlxuICBpZiAoIWxvY2F0aW9uLnBhdGhuYW1lLmluY2x1ZGVzKFwiL21haWxcIikpIHJldHVybjtcbiAgY29uc3QgYXR0ZW1wdCA9IChuOiBudW1iZXIpID0+IHtcbiAgICBjb25zdCByZXN1bHRzID0gaGVhbHRoQ2hlY2soKTtcbiAgICBjb25zdCByZXRyeUluID0gSEVBTFRIX1JFVFJZX01TW25dO1xuICAgIGlmIChicm9rZW5FbnRyaWVzKHJlc3VsdHMpLmxlbmd0aCA9PT0gMCB8fCByZXRyeUluID09PSB1bmRlZmluZWQpIHJlcG9ydEhlYWx0aChyZXN1bHRzKTtcbiAgICBlbHNlIGhlYWx0aFRpbWVyID0gc2V0VGltZW91dCgoKSA9PiBhdHRlbXB0KG4gKyAxKSwgcmV0cnlJbik7XG4gIH07XG4gIGhlYWx0aFRpbWVyID0gc2V0VGltZW91dCgoKSA9PiBhdHRlbXB0KDApLCBIRUFMVEhfU0VUVExFX01TKTtcbn1cblxuLyoqIFJlc29sdmVzIG9uY2UgT3V0bG9vaydzIG1haWwgVUkgaGFzIHJlbmRlcmVkLCBvciBhZnRlciBhIHRpbWVvdXQgc28gdGhlIGNoZWNrIHN0aWxsIHJlcG9ydHMuICovXG5mdW5jdGlvbiB3aGVuQXBwUmVhZHkoKTogUHJvbWlzZTx2b2lkPiB7XG4gIHJldHVybiBuZXcgUHJvbWlzZSgoZG9uZSkgPT4ge1xuICAgIGlmIChyZXNvbHZlKFwibWVzc2FnZUxpc3RcIikpIHJldHVybiBkb25lKCk7XG4gICAgY29uc3Qgb2JzZXJ2ZXIgPSBuZXcgTXV0YXRpb25PYnNlcnZlcigoKSA9PiB7XG4gICAgICBpZiAocmVzb2x2ZShcIm1lc3NhZ2VMaXN0XCIpKSBmaW5pc2goKTtcbiAgICB9KTtcbiAgICBjb25zdCB0aW1lciA9IHNldFRpbWVvdXQoZmluaXNoLCBBUFBfUkVBRFlfVElNRU9VVF9NUyk7XG4gICAgZnVuY3Rpb24gZmluaXNoKCkge1xuICAgICAgb2JzZXJ2ZXIuZGlzY29ubmVjdCgpO1xuICAgICAgY2xlYXJUaW1lb3V0KHRpbWVyKTtcbiAgICAgIGRvbmUoKTtcbiAgICB9XG4gICAgb2JzZXJ2ZXIub2JzZXJ2ZShkb2N1bWVudC5ib2R5LCB7IGNoaWxkTGlzdDogdHJ1ZSwgc3VidHJlZTogdHJ1ZSB9KTtcbiAgfSk7XG59XG5cbi8qKiBSZS10YWcgb24gRE9NIGNodXJuIChSZWFjdCByZS1yZW5kZXJzKSwgYW5kIHJlLWNoZWNrIGhlYWx0aCB3aGVuIHRoZSBTUEEgbmF2aWdhdGVzLiAqL1xuZnVuY3Rpb24gd2F0Y2hQYWdlKCk6IHZvaWQge1xuICBsZXQgc2NoZWR1bGVkID0gZmFsc2U7XG4gIGxldCBsYXN0VXJsID0gbG9jYXRpb24uaHJlZjtcbiAgbmV3IE11dGF0aW9uT2JzZXJ2ZXIoKCkgPT4ge1xuICAgIGlmIChzY2hlZHVsZWQpIHJldHVybjtcbiAgICBzY2hlZHVsZWQgPSB0cnVlO1xuICAgIHJlcXVlc3RBbmltYXRpb25GcmFtZSgoKSA9PiB7XG4gICAgICBzY2hlZHVsZWQgPSBmYWxzZTtcbiAgICAgIHRhZ0VsZW1lbnRzKCk7XG4gICAgICBzeW5jQ29sb3JTY2hlbWUoKTtcbiAgICAgIHN5bmNSZWFkaW5nUGFuZSgpO1xuICAgICAgaWYgKGxvY2F0aW9uLmhyZWYgIT09IGxhc3RVcmwpIHtcbiAgICAgICAgbGFzdFVybCA9IGxvY2F0aW9uLmhyZWY7XG4gICAgICAgIHNjaGVkdWxlSGVhbHRoQ2hlY2soKTtcbiAgICAgIH1cbiAgICB9KTtcbiAgfSkub2JzZXJ2ZShkb2N1bWVudC5ib2R5LCB7IGNoaWxkTGlzdDogdHJ1ZSwgc3VidHJlZTogdHJ1ZSB9KTtcbn1cblxubGV0IG9ycGhhbk5vdGljZVNob3duID0gZmFsc2U7XG5cbi8qKlxuICogQWZ0ZXIgT0JCIGlzIHJlbG9hZGVkIG9yIHVwZGF0ZWQsIHRoaXMgY29weSBvZiB0aGUgY29udGVudCBzY3JpcHQga2VlcHMgcnVubmluZyBpbiB0YWJzIHRoYXRcbiAqIHdlcmUgYWxyZWFkeSBvcGVuLCBjdXQgb2ZmIGZyb20gdGhlIGV4dGVuc2lvbi4gU2F5IHNvIG9uY2UgaW5zdGVhZCBvZiBoYWxmLXdvcmtpbmcuXG4gKi9cbmZ1bmN0aW9uIGlzT3JwaGFuZWQoKTogYm9vbGVhbiB7XG4gIGlmIChjaHJvbWUucnVudGltZT8uaWQpIHJldHVybiBmYWxzZTtcbiAgaWYgKCFvcnBoYW5Ob3RpY2VTaG93bikge1xuICAgIG9ycGhhbk5vdGljZVNob3duID0gdHJ1ZTtcbiAgICBub3RpY2UoXCJPQkIgd2FzIHVwZGF0ZWQuXCIsIFwiUmVsb2FkIHRhYlwiLCAoKSA9PiBsb2NhdGlvbi5yZWxvYWQoKSk7XG4gIH1cbiAgcmV0dXJuIHRydWU7XG59XG5cbmZ1bmN0aW9uIGlzRWRpdGFibGUodGFyZ2V0OiBFdmVudFRhcmdldCB8IHVuZGVmaW5lZCk6IGJvb2xlYW4ge1xuICBpZiAoISh0YXJnZXQgaW5zdGFuY2VvZiBIVE1MRWxlbWVudCkpIHJldHVybiBmYWxzZTtcbiAgcmV0dXJuIHRhcmdldC5pc0NvbnRlbnRFZGl0YWJsZSB8fCAvXihJTlBVVHxURVhUQVJFQXxTRUxFQ1QpJC8udGVzdCh0YXJnZXQudGFnTmFtZSk7XG59XG5cbmZ1bmN0aW9uIG1vdW50S2V5bWFwKCk6IHZvaWQge1xuICBjb25zdCBlbmdpbmUgPSBuZXcgU2VxdWVuY2VFbmdpbmUoQklORElOR1MpO1xuICBsZXQgaGludFRpbWVyOiBSZXR1cm5UeXBlPHR5cGVvZiBzZXRUaW1lb3V0PiB8IHVuZGVmaW5lZDtcbiAgd2luZG93LmFkZEV2ZW50TGlzdGVuZXIoXG4gICAgXCJrZXlkb3duXCIsXG4gICAgKGV2ZW50KSA9PiB7XG4gICAgICBpZiAoIXNldHRpbmdzLmtleW1hcCB8fCAhZXZlbnQuaXNUcnVzdGVkIHx8IGlzRnJvbU92ZXJsYXkoZXZlbnQpIHx8IGlzTW9kYWxPcGVuKCkpIHJldHVybjtcbiAgICAgIGlmIChpc09ycGhhbmVkKCkpIHJldHVybjtcbiAgICAgIGlmIChpc0VkaXRhYmxlKGV2ZW50LmNvbXBvc2VkUGF0aCgpWzBdKSkgcmV0dXJuO1xuICAgICAgY29uc3Qga2V5ID0ga2V5RnJvbUV2ZW50KGV2ZW50KTtcbiAgICAgIGlmIChrZXkgPT09IG51bGwpIHJldHVybjtcblxuICAgICAgY29uc3QgcmVzdWx0ID0gZW5naW5lLmZlZWQoa2V5LCBjdXJyZW50Q29udGV4dCgpLCBldmVudC50aW1lU3RhbXApO1xuICAgICAgaWYgKHJlc3VsdC5raW5kID09PSBcIm5vbmVcIikgcmV0dXJuO1xuICAgICAgc2hvd0hpbnQoXG4gICAgICAgIHJlc3VsdC5raW5kID09PSBcInBlbmRpbmdcIlxuICAgICAgICAgID8gYCR7cmVzdWx0LnByZWZpeC5qb2luKFwiIFwiKX0gXHUyMDI2ICR7cmVzdWx0Lm5leHRcbiAgICAgICAgICAgICAgLm1hcCgoYikgPT4gYCR7Yi5rZXlzLmF0KC0xKX0gJHtBQ1RJT05TW2IuYWN0aW9uXS5kZXNjcmlwdGlvbn1gKVxuICAgICAgICAgICAgICAuam9pbihcIiBcdTAwQjcgXCIpfWBcbiAgICAgICAgICA6IG51bGwsXG4gICAgICApO1xuICAgICAgY2xlYXJUaW1lb3V0KGhpbnRUaW1lcik7XG4gICAgICBpZiAocmVzdWx0LmtpbmQgPT09IFwicGVuZGluZ1wiKVxuICAgICAgICBoaW50VGltZXIgPSBzZXRUaW1lb3V0KCgpID0+IHNob3dIaW50KG51bGwpLCBERUZBVUxUX1RJTUVPVVRfTVMpO1xuICAgICAgaWYgKHJlc3VsdC5raW5kID09PSBcIm1hdGNoXCIgJiYgIWV4ZWN1dGUocmVzdWx0LmFjdGlvbiwgeyBtdWx0aVNlbGVjdDogc2V0dGluZ3MubXVsdGlTZWxlY3QgfSkpXG4gICAgICAgIHJldHVybjtcbiAgICAgIGV2ZW50LnByZXZlbnREZWZhdWx0KCk7XG4gICAgICBldmVudC5zdG9wSW1tZWRpYXRlUHJvcGFnYXRpb24oKTtcbiAgICB9LFxuICAgIHRydWUsXG4gICk7XG59XG5cbi8qKiBEZXZlbG9wZXIgdG9vbHM6IHRoZSBwb3B1cCBhc2tzIHRoaXMgdGFiIHRvIHNhdmUgYSBjYXB0dXJlIG9mIHRoZSBwYWdlIChzcmMvY2FwdHVyZS50cykuICovXG5mdW5jdGlvbiBzZXJ2ZUNhcHR1cmVSZXF1ZXN0cygpOiB2b2lkIHtcbiAgY2hyb21lLnJ1bnRpbWUub25NZXNzYWdlLmFkZExpc3RlbmVyKChtZXNzYWdlLCBfc2VuZGVyLCByZXBseSkgPT4ge1xuICAgIGlmIChtZXNzYWdlPy50eXBlICE9PSBDQVBUVVJFX01FU1NBR0UpIHJldHVybjtcbiAgICBjb25zdCB7IGh0bWwsIG5hbWUgfSA9IGNhcHR1cmVQYWdlKGRvY3VtZW50LCBsb2NhdGlvbiwgY2hyb21lLnJ1bnRpbWUuZ2V0TWFuaWZlc3QoKS52ZXJzaW9uKTtcbiAgICBjb25zdCB1cmwgPSBVUkwuY3JlYXRlT2JqZWN0VVJMKG5ldyBCbG9iKFtodG1sXSwgeyB0eXBlOiBcInRleHQvaHRtbFwiIH0pKTtcbiAgICBjb25zdCBsaW5rID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudChcImFcIik7XG4gICAgbGluay5ocmVmID0gdXJsO1xuICAgIGxpbmsuZG93bmxvYWQgPSBuYW1lO1xuICAgIGRvY3VtZW50LmJvZHkuYXBwZW5kKGxpbmspO1xuICAgIGxpbmsuY2xpY2soKTtcbiAgICBsaW5rLnJlbW92ZSgpO1xuICAgIHNldFRpbWVvdXQoKCkgPT4gVVJMLnJldm9rZU9iamVjdFVSTCh1cmwpLCAxMF8wMDApO1xuICAgIHJlcGx5KHsgb2s6IHRydWUsIG5hbWUsIGJ5dGVzOiBodG1sLmxlbmd0aCB9KTtcbiAgfSk7XG59XG5cbmFzeW5jIGZ1bmN0aW9uIG1haW4oKTogUHJvbWlzZTx2b2lkPiB7XG4gIHNlcnZlQ2FwdHVyZVJlcXVlc3RzKCk7XG4gIGFwcGx5U2V0dGluZ3MoYXdhaXQgbG9hZFNldHRpbmdzKCkpO1xuICBvblNldHRpbmdzQ2hhbmdlZChhcHBseVNldHRpbmdzKTtcbiAgbW91bnRLZXltYXAoKTtcbiAgLy8gQ29taW5nIGJhY2sgdG8gdGhpcyB0YWIgYWZ0ZXIgcmVsb2FkaW5nIE9CQiBmcm9tIGl0cyBwb3B1cDogb2ZmZXIgdGhlIHJlZnJlc2ggcmlnaHQgYXdheS5cbiAgZG9jdW1lbnQuYWRkRXZlbnRMaXN0ZW5lcihcInZpc2liaWxpdHljaGFuZ2VcIiwgKCkgPT4ge1xuICAgIGlmICghZG9jdW1lbnQuaGlkZGVuKSBpc09ycGhhbmVkKCk7XG4gIH0pO1xuICBhd2FpdCB3aGVuQXBwUmVhZHkoKTtcbiAgdGFnRWxlbWVudHMoKTtcbiAgc3luY1JlYWRpbmdQYW5lKCk7XG4gIHNjaGVkdWxlSGVhbHRoQ2hlY2soKTtcbiAgd2F0Y2hQYWdlKCk7XG59XG5cbnZvaWQgbWFpbigpO1xuIl0sCiAgIm1hcHBpbmdzIjogIjs7O0FBeUJPLE1BQU0sWUFBWTtBQUFBO0FBQUEsSUFFdkIsVUFBVSxFQUFFLFlBQVksQ0FBQyxjQUFjLEdBQUcsUUFBUSxVQUFVLEtBQUssWUFBWTtBQUFBLElBQzdFLFdBQVcsRUFBRSxZQUFZLENBQUMsbUJBQW1CLHVCQUF1QixHQUFHLFFBQVEsU0FBUztBQUFBLElBQ3hGLFNBQVM7QUFBQSxNQUNQLFlBQVksQ0FBQyxhQUFhLG9EQUFvRDtBQUFBLE1BQzlFLFFBQVE7QUFBQSxNQUNSLEtBQUs7QUFBQSxJQUNQO0FBQUEsSUFDQSxRQUFRO0FBQUEsTUFDTixZQUFZLENBQUMsK0JBQStCLHNDQUFzQztBQUFBLE1BQ2xGLFFBQVE7QUFBQSxNQUNSLEtBQUs7QUFBQSxJQUNQO0FBQUEsSUFDQSxZQUFZO0FBQUEsTUFDVixZQUFZLENBQUMsNENBQTRDO0FBQUEsTUFDekQsUUFBUTtBQUFBLE1BQ1IsUUFBUTtBQUFBLE1BQ1IsS0FBSztB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O0FBQ2pFLFlBQUksT0FBTyxTQUFTLE9BQVE7QUFDNUI7QUFBQSxVQUNFLE9BQU8sU0FBUyxZQUNaLEdBQUcsT0FBTyxPQUFPLEtBQUssR0FBRyxDQUFDLFdBQU0sT0FBTyxLQUNwQyxJQUFJLENBQUMsTUFBTSxHQUFHLEVBQUUsS0FBSyxHQUFHLEVBQUUsQ0FBQyxJQUFJLFFBQVEsRUFBRSxNQUFNLEVBQUUsV0FBVyxFQUFFLEVBQzlELEtBQUssUUFBSyxDQUFDLEtBQ2Q7QUFBQSxRQUNOO0FBQ0EscUJBQWEsU0FBUztBQUN0QixZQUFJLE9BQU8sU0FBUztBQUNsQixzQkFBWSxXQUFXLE1BQU0sU0FBUyxJQUFJLEdBQUcsa0JBQWtCO0FBQ2pFLFlBQUksT0FBTyxTQUFTLFdBQVcsQ0FBQyxRQUFRLE9BQU8sUUFBUSxFQUFFLGFBQWEsU0FBUyxZQUFZLENBQUM7QUFDMUY7QUFDRixjQUFNLGVBQWU7QUFDckIsY0FBTSx5QkFBeUI7QUFBQSxNQUNqQztBQUFBLE1BQ0E7QUFBQSxJQUNGO0FBQUEsRUFDRjtBQUdBLFdBQVMsdUJBQTZCO0FBQ3BDLFdBQU8sUUFBUSxVQUFVLFlBQVksQ0FBQyxTQUFTLFNBQVMsVUFBVTtBQUNoRSxVQUFJLFNBQVMsU0FBUyxnQkFBaUI7QUFDdkMsWUFBTSxFQUFFLE1BQU0sS0FBSyxJQUFJLFlBQVksVUFBVSxVQUFVLE9BQU8sUUFBUSxZQUFZLEVBQUUsT0FBTztBQUMzRixZQUFNLE1BQU0sSUFBSSxnQkFBZ0IsSUFBSSxLQUFLLENBQUMsSUFBSSxHQUFHLEVBQUUsTUFBTSxZQUFZLENBQUMsQ0FBQztBQUN2RSxZQUFNLE9BQU8sU0FBUyxjQUFjLEdBQUc7QUFDdkMsV0FBSyxPQUFPO0FBQ1osV0FBSyxXQUFXO0FBQ2hCLGVBQVMsS0FBSyxPQUFPLElBQUk7QUFDekIsV0FBSyxNQUFNO0FBQ1gsV0FBSyxPQUFPO0FBQ1osaUJBQVcsTUFBTSxJQUFJLGdCQUFnQixHQUFHLEdBQUcsR0FBTTtBQUNqRCxZQUFNLEVBQUUsSUFBSSxNQUFNLE1BQU0sT0FBTyxLQUFLLE9BQU8sQ0FBQztBQUFBLElBQzlDLENBQUM7QUFBQSxFQUNIO0FBRUEsaUJBQWUsT0FBc0I7QUFDbkMseUJBQXFCO0FBQ3JCLGtCQUFjLE1BQU0sYUFBYSxDQUFDO0FBQ2xDLHNCQUFrQixhQUFhO0FBQy9CLGdCQUFZO0FBRVosYUFBUyxpQkFBaUIsb0JBQW9CLE1BQU07QUFDbEQsVUFBSSxDQUFDLFNBQVMsT0FBUSxZQUFXO0FBQUEsSUFDbkMsQ0FBQztBQUNELFVBQU0sYUFBYTtBQUNuQixnQkFBWTtBQUNaLG9CQUFnQjtBQUNoQix3QkFBb0I7QUFDcEIsY0FBVTtBQUFBLEVBQ1o7QUFFQSxPQUFLLEtBQUs7IiwKICAibmFtZXMiOiBbImVsIiwgImVsIiwgInNldHRpbmdzIl0KfQo=
