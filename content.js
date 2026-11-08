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
        if (result.kind === "match" && !execute(result.action)) return;
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
//# sourceMappingURL=data:application/json;base64,ewogICJ2ZXJzaW9uIjogMywKICAic291cmNlcyI6IFsiLi4vc3JjL3NlbGVjdG9ycy50cyIsICIuLi9zcmMvY2FwdHVyZS50cyIsICIuLi9zcmMva2V5bWFwL2FjdGlvbnMudHMiLCAiLi4vc3JjL2tleW1hcC9iaW5kaW5ncy50cyIsICIuLi9zcmMva2V5bWFwL2VuZ2luZS50cyIsICIuLi9zcmMvdWkvb3ZlcmxheS5jc3MiLCAiLi4vc3JjL3VpL292ZXJsYXkudHMiLCAiLi4vc3JjL3BhbGV0dGUvZnV6enkudHMiLCAiLi4vc3JjL3BhbGV0dGUvcGFsZXR0ZS50cyIsICIuLi9zcmMva2V5bWFwL2V4ZWN1dGUudHMiLCAiLi4vc3JjL3NldHRpbmdzLnRzIiwgIi4uL3NyYy9jb250ZW50LnRzIl0sCiAgInNvdXJjZXNDb250ZW50IjogWyIvLyBUaGUgb25seSBmaWxlIGFsbG93ZWQgdG8ga25vdyB3aGF0IE91dGxvb2sncyBET00gbG9va3MgbGlrZS5cbi8vXG4vLyBBbmNob3IgcHJlZmVyZW5jZSAoc2VlIFBMQU4ubWQpOiBzdHJ1Y3R1cmFsIElEcyA+IEFSSUEgcm9sZXMvbGFiZWxzID4gc2VtYW50aWMgZGF0YS0qID4gdGV4dC5cbi8vIE5ldmVyIHVzZSBnZW5lcmF0ZWQgY2xhc3MgbmFtZXM7IGlmIG9uZSBpcyB1bmF2b2lkYWJsZSwgbWFyayBpdCBgLy8gRlJBR0lMRTpgLlxuLy9cbi8vIFZlcmlmaWVkIGFnYWluc3QgYSBzYXZlZCBwYWdlIGZyb20gT2N0IDIwMjYgKHRlc3QvZml4dHVyZXMvaW5ib3guaHRtbCk7IHNlZSBkb2NzL2RvbS1zdXJ2ZXkubWQuXG4vLyBFbnRyaWVzIG1hcmtlZCBVTlZFUklGSUVEIHdlcmUgbm90IHByZXNlbnQgaW4gdGhhdCBjYXB0dXJlLiBhcmlhLWxhYmVscyBhc3N1bWUgYW4gRW5nbGlzaCBVSS5cblxuZXhwb3J0IHR5cGUgU2VsZWN0b3JOYW1lID0ga2V5b2YgdHlwZW9mIFNFTEVDVE9SUztcblxuZXhwb3J0IGludGVyZmFjZSBTZWxlY3RvckVudHJ5IHtcbiAgLyoqIENhbmRpZGF0ZSBDU1Mgc2VsZWN0b3JzLCBtb3N0IHN0YWJsZSBmaXJzdC4gVGhlIGZpcnN0IG9uZSB0aGF0IG1hdGNoZXMgd2lucy4gKi9cbiAgY2FuZGlkYXRlczogcmVhZG9ubHkgc3RyaW5nW107XG4gIC8qKiBcImFsd2F5c1wiOiBzaG91bGQgZXhpc3Qgb24gZXZlcnkgbWFpbCB2aWV3LCBzbyBhIG1pc3MgaXMgYSBicmVha2FnZS4gXCJzb21ldGltZXNcIjogY29udGV4dHVhbC4gKi9cbiAgZXhwZWN0OiBcImFsd2F5c1wiIHwgXCJzb21ldGltZXNcIjtcbiAgLyoqIE5hbWUgb2YgYW5vdGhlciBlbnRyeSB0byByZXNvbHZlIGluc2lkZSwgaW5zdGVhZCBvZiB0aGUgd2hvbGUgZG9jdW1lbnQuICovXG4gIHdpdGhpbj86IHN0cmluZztcbiAgLyoqIFRva2VuIGFkZGVkIHRvIHRoZSBlbGVtZW50J3MgYGRhdGEtb3Zlcmxvb2tgIGF0dHJpYnV0ZSBzbyBDU1MgY2FuIHRhcmdldCBpdC4gKi9cbiAgdGFnPzogc3RyaW5nO1xuICAvKiogVGFnIGV2ZXJ5IG1hdGNoIHJhdGhlciB0aGFuIG9ubHkgdGhlIGZpcnN0LiAqL1xuICBhbGw/OiBib29sZWFuO1xuICAvKiogTWF0Y2ggZXZlcnkgY2FuZGlkYXRlIChhIGxpc3Qgb2YgZGlzdGluY3QgdGhpbmdzKSwgbm90IGp1c3QgdGhlIGZpcnN0IHRoYXQgbWF0Y2hlcy4gKi9cbiAgdW5pb24/OiBib29sZWFuO1xufVxuXG5leHBvcnQgY29uc3QgU0VMRUNUT1JTID0ge1xuICAvLyAtLS0gU3RydWN0dXJlIC0tLVxuICB0aXRsZUJhcjogeyBjYW5kaWRhdGVzOiBbXCIjT3dhVGl0bGVCYXJcIl0sIGV4cGVjdDogXCJhbHdheXNcIiwgdGFnOiBcInRpdGxlLWJhclwiIH0sXG4gIHNlYXJjaEJveDogeyBjYW5kaWRhdGVzOiBbXCIjdG9wU2VhcmNoSW5wdXRcIiwgJ1tyb2xlPVwic2VhcmNoXCJdIGlucHV0J10sIGV4cGVjdDogXCJhbHdheXNcIiB9LFxuICBhcHBSYWlsOiB7XG4gICAgY2FuZGlkYXRlczogW1wiI0xlZnRSYWlsXCIsICdbcm9sZT1cIm5hdmlnYXRpb25cIl1bYXJpYS1sYWJlbD1cImxlZnQtcmFpbC1hcHBiYXJcIl0nXSxcbiAgICBleHBlY3Q6IFwiYWx3YXlzXCIsXG4gICAgdGFnOiBcImFwcC1yYWlsXCIsXG4gIH0sXG4gIHJpYmJvbjoge1xuICAgIGNhbmRpZGF0ZXM6IFsnW2RhdGEtYXBwLXNlY3Rpb249XCJSaWJib25cIl0nLCAnW3JvbGU9XCJyZWdpb25cIl1bYXJpYS1sYWJlbD1cIlJpYmJvblwiXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB0YWc6IFwicmliYm9uXCIsXG4gIH0sXG4gIHJpYmJvblRhYnM6IHtcbiAgICBjYW5kaWRhdGVzOiBbJ1tyb2xlPVwidGFibGlzdFwiXVthcmlhLWxhYmVsPVwiUmliYm9uIHRhYnNcIl0nXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgd2l0aGluOiBcInJpYmJvblwiLFxuICAgIHRhZzogXCJyaWJib24tdGFic1wiLFxuICB9LFxuICBmb2xkZXJQYW5lOiB7XG4gICAgY2FuZGlkYXRlczogW1xuICAgICAgJ1tkYXRhLWFwcC1zZWN0aW9uPVwiTmF2aWdhdGlvblBhbmVcIl0nLFxuICAgICAgJ1tyb2xlPVwicmVnaW9uXCJdW2FyaWEtbGFiZWw9XCJOYXZpZ2F0aW9uIHBhbmVcIl0nLFxuICAgIF0sXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICAgIHRhZzogXCJmb2xkZXItcGFuZVwiLFxuICB9LFxuICBmb2xkZXJJdGVtOiB7XG4gICAgY2FuZGlkYXRlczogWydbcm9sZT1cInRyZWVpdGVtXCJdW2RhdGEtZm9sZGVyLW5hbWVdJywgJ1tyb2xlPVwidHJlZWl0ZW1cIl0nXSxcbiAgICBleHBlY3Q6IFwiYWx3YXlzXCIsXG4gICAgd2l0aGluOiBcImZvbGRlclBhbmVcIixcbiAgICB0YWc6IFwiZm9sZGVyLWl0ZW1cIixcbiAgICBhbGw6IHRydWUsXG4gIH0sXG4gIGxpc3RQYW5lOiB7XG4gICAgY2FuZGlkYXRlczogWydbZGF0YS1hcHAtc2VjdGlvbj1cIk1lc3NhZ2VMaXN0XCJdJ10sXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICAgIHRhZzogXCJsaXN0LXBhbmVcIixcbiAgfSxcbiAgbWVzc2FnZUxpc3Q6IHtcbiAgICBjYW5kaWRhdGVzOiBbXCIjTWFpbExpc3RcIiwgJ1tyb2xlPVwibGlzdGJveFwiXVthcmlhLWxhYmVsPVwiTWVzc2FnZSBsaXN0XCJdJ10sXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICAgIHRhZzogXCJtZXNzYWdlLWxpc3RcIixcbiAgfSxcbiAgbWVzc2FnZVJvdzoge1xuICAgIGNhbmRpZGF0ZXM6IFsnW3JvbGU9XCJvcHRpb25cIl1bZGF0YS1jb252aWRdJywgJ1tyb2xlPVwib3B0aW9uXCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJtZXNzYWdlTGlzdFwiLFxuICAgIHRhZzogXCJtZXNzYWdlLXJvd1wiLFxuICAgIGFsbDogdHJ1ZSxcbiAgfSxcbiAgZ3JvdXBIZWFkZXI6IHtcbiAgICBjYW5kaWRhdGVzOiBbJ1tyb2xlPVwiYnV0dG9uXCJdW2lkXj1cImdyb3VwSGVhZGVyXCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJtZXNzYWdlTGlzdFwiLFxuICAgIHRhZzogXCJncm91cC1oZWFkZXJcIixcbiAgICBhbGw6IHRydWUsXG4gIH0sXG4gIGZvY3VzZWRUYWJzOiB7XG4gICAgY2FuZGlkYXRlczogWydbcm9sZT1cInRhYmxpc3RcIl0nXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgd2l0aGluOiBcImxpc3RQYW5lXCIsXG4gICAgdGFnOiBcImZvY3VzZWQtdGFic1wiLFxuICB9LFxuICByZWFkaW5nUGFuZToge1xuICAgIGNhbmRpZGF0ZXM6IFtcbiAgICAgICdbZGF0YS1hcHAtc2VjdGlvbj1cIk1haWxSZWFkQ29tcG9zZVwiXScsXG4gICAgICBcIiNSZWFkaW5nUGFuZUNvbnRhaW5lcklkXCIsXG4gICAgICAnW3JvbGU9XCJtYWluXCJdW2FyaWEtbGFiZWw9XCJSZWFkaW5nIFBhbmVcIl0nLFxuICAgIF0sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHRhZzogXCJyZWFkaW5nLXBhbmVcIixcbiAgfSxcblxuICBtZXNzYWdlQm9keToge1xuICAgIGNhbmRpZGF0ZXM6IFsnW3JvbGU9XCJkb2N1bWVudFwiXVthcmlhLWxhYmVsPVwiTWVzc2FnZSBib2R5XCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJyZWFkaW5nUGFuZVwiLFxuICAgIHRhZzogXCJtZXNzYWdlLWJvZHlcIixcbiAgICBhbGw6IHRydWUsXG4gIH0sXG5cbiAgaW5mb0Jhcjoge1xuICAgIC8vIE5vdGljZXMgYWJvdmUgYSBtZXNzYWdlIChcIllvdSByZXBsaWVkIG9uXHUyMDI2XCIpOyB0aGVpciB0ZXh0IHNpdHMgaW4gYW4gZWxlbWVudCB3aG9zZSBpZCBzdGFydHNcbiAgICAvLyB3aXRoIFwiaW5mb0JhclwiLiBPdXRsb29rIHBhaW50cyB0aGVtIHdpdGggaW5saW5lIGNvbG91cnMuXG4gICAgY2FuZGlkYXRlczogWydbcm9sZT1cImdyb3VwXCJdOmhhcyg+ICogPiBbaWRePVwiaW5mb0JhclwiXSknXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgd2l0aGluOiBcInJlYWRpbmdQYW5lXCIsXG4gICAgdGFnOiBcImluZm8tYmFyXCIsXG4gICAgYWxsOiB0cnVlLFxuICB9LFxuXG4gIC8vIC0tLSBDbHV0dGVyIChlYWNoIGhpZGVhYmxlIGVudHJ5IGdldHMgYSB0b2dnbGUgaW4gc2V0dGluZ3MudHMpIC0tLVxuICB1cHNlbGxCYW5uZXI6IHtcbiAgICBjYW5kaWRhdGVzOiBbJ1tkYXRhLWFwcC1zZWN0aW9uPVwiVXBzZWxsQmFubmVyU2VjdGlvblwiXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB0YWc6IFwidXBzZWxsXCIsXG4gICAgYWxsOiB0cnVlLFxuICB9LFxuICBoZWFkZXJFeHRyYXM6IHtcbiAgICBjYW5kaWRhdGVzOiBbXCIjb3dhVGVhbXNCdG5fY29udGFpbmVyXCIsIFwiI293YVRpbWVQYW5lbEJ0bl9jb250YWluZXJcIl0sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHRhZzogXCJoZWFkZXItZXh0cmFcIixcbiAgICB1bmlvbjogdHJ1ZSxcbiAgfSxcbiAgYXBwUmFpbEV4dHJhczoge1xuICAgIGNhbmRpZGF0ZXM6IFtcIlRvIERvXCIsIFwiTmV3c2xldHRlcnNcIiwgXCJPcmcgRXhwbG9yZXJcIiwgXCJPbmVEcml2ZVwiXS5tYXAoXG4gICAgICAoYXBwKSA9PiBgW3JvbGU9XCJub25lXCJdW2FyaWEtbGFiZWw9XCIke2FwcH1cIl1gLFxuICAgICksXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJhcHBSYWlsXCIsXG4gICAgdGFnOiBcInJhaWwtZXh0cmFcIixcbiAgICB1bmlvbjogdHJ1ZSxcbiAgfSxcbiAgcmliYm9uRXh0cmFzOiB7XG4gICAgY2FuZGlkYXRlczogW1xuICAgICAgXCJTaGFyZSB0byBUZWFtc1wiLFxuICAgICAgXCJTd2VlcFwiLFxuICAgICAgXCJBc3NpZ24gcG9saWN5XCIsXG4gICAgICBcIlZpdmEgSW5zaWdodHNcIixcbiAgICAgIFwiTW9yZSBhcHBzXCIsXG4gICAgICBcIlF1aWNrIHN0ZXBzXCIsXG4gICAgICBcIk1vcmUgcXVpY2sgc3RlcHNcIixcbiAgICBdLm1hcCgobGFiZWwpID0+IGBbYXJpYS1sYWJlbD1cIiR7bGFiZWx9XCJdYCksXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJyaWJib25cIixcbiAgICB0YWc6IFwicmliYm9uLWV4dHJhXCIsXG4gICAgdW5pb246IHRydWUsXG4gIH0sXG4gIHJvd0F2YXRhcjoge1xuICAgIGNhbmRpZGF0ZXM6IFsnW3JvbGU9XCJpbWdcIl1baWRePVwiYXZhdGFyLVwiXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB3aXRoaW46IFwibWVzc2FnZUxpc3RcIixcbiAgICB0YWc6IFwicm93LWF2YXRhclwiLFxuICAgIGFsbDogdHJ1ZSxcbiAgfSxcbiAgcm93SG92ZXJBY3Rpb25zOiB7XG4gICAgY2FuZGlkYXRlczogW1xuICAgICAgXCJDdXN0b21pemUgQWN0aW9uc1wiLFxuICAgICAgXCJGbGFnIHRoaXMgbWVzc2FnZVwiLFxuICAgICAgXCJLZWVwIHRoaXMgbWVzc2FnZSBhdCB0aGUgdG9wIG9mIHlvdXIgZm9sZGVyXCIsXG4gICAgXS5tYXAoKGxhYmVsKSA9PiBgYnV0dG9uW2FyaWEtbGFiZWw9XCIke2xhYmVsfVwiXWApLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB3aXRoaW46IFwibWVzc2FnZUxpc3RcIixcbiAgICB0YWc6IFwicm93LWFjdGlvblwiLFxuICAgIHVuaW9uOiB0cnVlLFxuICB9LFxuICBhcHBCYW5uZXI6IHtcbiAgICAvLyBGbHVlbnQncyBzdGF0aWMgXCJmdWktTWVzc2FnZUJhclwiIGNsYXNzIGlzIHBhcnQgb2YgaXRzIHB1YmxpYyBzdHlsaW5nIEFQSSwgbm90IGdlbmVyYXRlZDtcbiAgICAvLyBub3RoaW5nIGVsc2UgbWFya3MgdGhlc2UuIFNjb3BlZCB0byBiYXJzIGFjcm9zcyB0aGUgdG9wIG9mIHRoZSBhcHAsIG5vdCBpbnNpZGUgbWVzc2FnZXMuXG4gICAgY2FuZGlkYXRlczogWycjbWFpbkFwcCA+IGRpdiA+IFtyb2xlPVwiZ3JvdXBcIl0uZnVpLU1lc3NhZ2VCYXInXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgdGFnOiBcImFwcC1iYW5uZXJcIixcbiAgICBhbGw6IHRydWUsXG4gIH0sXG4gIGNvcGlsb3Q6IHtcbiAgICAvLyBVTlZFUklGSUVEOiBubyBDb3BpbG90IHN1cmZhY2VzIGluIHRoZSBPY3QgMjAyNiBjYXB0dXJlLlxuICAgIGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWwqPVwiQ29waWxvdFwiIGldJywgJ1tyb2xlPVwicmVnaW9uXCJdW2FyaWEtbGFiZWwqPVwiQ29waWxvdFwiIGldJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHRhZzogXCJjb3BpbG90XCIsXG4gICAgdW5pb246IHRydWUsXG4gIH0sXG4gIGV4dGVybmFsVGFnOiB7XG4gICAgLy8gVU5WRVJJRklFRDogbm8gRXh0ZXJuYWwgdGFncyBpbiB0aGUgT2N0IDIwMjYgY2FwdHVyZS5cbiAgICBjYW5kaWRhdGVzOiBbJ1t0aXRsZT1cIkV4dGVybmFsIHNlbmRlclwiXScsICdbYXJpYS1sYWJlbD1cIkV4dGVybmFsXCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHRhZzogXCJleHRlcm5hbC10YWdcIixcbiAgICB1bmlvbjogdHJ1ZSxcbiAgfSxcblxuICAvLyAtLS0gQ29udHJvbHMgdGhlIGtleW1hcCBkcml2ZXMgLS0tXG4gIGNvbXBvc2VCdXR0b246IHsgY2FuZGlkYXRlczogWydidXR0b25bYXJpYS1sYWJlbD1cIk5ldyBtYWlsXCJdJ10sIGV4cGVjdDogXCJhbHdheXNcIiB9LFxuICBhcmNoaXZlQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJBcmNoaXZlXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICBkZWxldGVCdXR0b246IHsgY2FuZGlkYXRlczogWydidXR0b25bYXJpYS1sYWJlbD1cIkRlbGV0ZVwiXSddLCBleHBlY3Q6IFwic29tZXRpbWVzXCIgfSxcbiAgLy8gVU5WRVJJRklFRDogdGhpcyB0ZW5hbnQgc2hvd3MgYSBcIlBoaXNoIEFsZXJ0XCIgYWRkLWluIGluc3RlYWQgb2YgT3V0bG9vaydzIFJlcG9ydCBtZW51LlxuICBqdW5rQnV0dG9uOiB7XG4gICAgY2FuZGlkYXRlczogWydidXR0b25bYXJpYS1sYWJlbD1cIlJlcG9ydFwiXScsICdidXR0b25bYXJpYS1sYWJlbD1cIkp1bmtcIl0nXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gIH0sXG4gIG1vdmVCdXR0b246IHsgY2FuZGlkYXRlczogWydidXR0b25bYXJpYS1sYWJlbD1cIk1vdmUgdG9cIl0nXSwgZXhwZWN0OiBcInNvbWV0aW1lc1wiIH0sXG4gIHNub296ZUJ1dHRvbjogeyBjYW5kaWRhdGVzOiBbJ2J1dHRvblthcmlhLWxhYmVsPVwiU25vb3plXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICBmbGFnQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJGbGFnIC8gVW5mbGFnXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICByZWFkVW5yZWFkQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJSZWFkIC8gVW5yZWFkXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICB1bmRvQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJVbmRvXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICBzeW5jQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFtcIiNvd2FTeW5jQnV0dG9uX2NvbnRhaW5lclwiXSwgZXhwZWN0OiBcInNvbWV0aW1lc1wiIH0sXG4gIHNldHRpbmdzQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFtcIiNvd2FTZXR0aW5nc0J0bl9jb250YWluZXJcIl0sIGV4cGVjdDogXCJhbHdheXNcIiB9LFxuICAvLyBBcHAgcmFpbCBidXR0b25zLiBGYWxsYmFja3MgYXJlIHRoZSBNaWNyb3NvZnQgMzY1IGFwcCBpZHMgT3V0bG9vayBwdXRzIGluIGRhdGEtdGlkLCB0aGUgc2FtZVxuICAvLyBhY3Jvc3MgdGhlIE9jdCBhbmQgTm92IDIwMjYgY2FwdHVyZXMuIFRoZSByYWlsIHJlbmRlcnMgYWZ0ZXIgdGhlIG1lc3NhZ2UgbGlzdC5cbiAgbWFpbEFwcDoge1xuICAgIGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJNYWlsXCJdJywgJ1tkYXRhLXRpZD1cImRkZWE3NzRjLTM4MmItNDdkNy1hYWI1LWFkYzIxMzlhODAyYlwiXSddLFxuICAgIHdpdGhpbjogXCJhcHBSYWlsXCIsXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICB9LFxuICBjYWxlbmRhckFwcDoge1xuICAgIGNhbmRpZGF0ZXM6IFtcbiAgICAgICdidXR0b25bYXJpYS1sYWJlbD1cIkNhbGVuZGFyXCJdJyxcbiAgICAgICdbZGF0YS10aWQ9XCI4Y2JlYjg2Zi04M2UxLTQzYjUtYWFiYS1jZDM1MTQzMjJmMGJcIl0nLFxuICAgIF0sXG4gICAgd2l0aGluOiBcImFwcFJhaWxcIixcbiAgICBleHBlY3Q6IFwiYWx3YXlzXCIsXG4gIH0sXG4gIHBlb3BsZUFwcDoge1xuICAgIGNhbmRpZGF0ZXM6IFtcbiAgICAgICdidXR0b25bYXJpYS1sYWJlbD1cIlBlb3BsZVwiXScsXG4gICAgICAnW2RhdGEtdGlkPVwiMzU1ZmJkNzktM2JhMi00NTU0LThmMmQtMDMwMGZkZTkxZjMwXCJdJyxcbiAgICBdLFxuICAgIHdpdGhpbjogXCJhcHBSYWlsXCIsXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICB9LFxufSBhcyBjb25zdCBzYXRpc2ZpZXMgUmVjb3JkPHN0cmluZywgU2VsZWN0b3JFbnRyeT47XG5cbmNvbnN0IGVudHJ5ID0gKG5hbWU6IFNlbGVjdG9yTmFtZSk6IFNlbGVjdG9yRW50cnkgPT4gU0VMRUNUT1JTW25hbWVdO1xuXG5mdW5jdGlvbiBzY29wZShuYW1lOiBTZWxlY3Rvck5hbWUsIHJvb3Q6IFBhcmVudE5vZGUpOiBQYXJlbnROb2RlIHwgbnVsbCB7XG4gIGNvbnN0IHBhcmVudCA9IGVudHJ5KG5hbWUpLndpdGhpbjtcbiAgcmV0dXJuIHBhcmVudCA/IHJlc29sdmUocGFyZW50IGFzIFNlbGVjdG9yTmFtZSwgcm9vdCkgOiByb290O1xufVxuXG4vKiogVGhlIGZpcnN0IGVsZW1lbnQgbWF0Y2hpbmcgYG5hbWVgLCB0cnlpbmcgY2FuZGlkYXRlcyBpbiBwcmlvcml0eSBvcmRlci4gKi9cbmV4cG9ydCBmdW5jdGlvbiByZXNvbHZlKG5hbWU6IFNlbGVjdG9yTmFtZSwgcm9vdDogUGFyZW50Tm9kZSA9IGRvY3VtZW50KTogRWxlbWVudCB8IG51bGwge1xuICBjb25zdCB3aXRoaW4gPSBzY29wZShuYW1lLCByb290KTtcbiAgaWYgKCF3aXRoaW4pIHJldHVybiBudWxsO1xuICBmb3IgKGNvbnN0IGNhbmRpZGF0ZSBvZiBlbnRyeShuYW1lKS5jYW5kaWRhdGVzKSB7XG4gICAgY29uc3QgZWwgPSB3aXRoaW4ucXVlcnlTZWxlY3RvcihjYW5kaWRhdGUpO1xuICAgIGlmIChlbCkgcmV0dXJuIGVsO1xuICB9XG4gIHJldHVybiBudWxsO1xufVxuXG4vKiogRXZlcnkgbWF0Y2g6IGZvciB1bmlvbiBlbnRyaWVzLCBvZiBhbGwgY2FuZGlkYXRlczsgb3RoZXJ3aXNlIG9mIHRoZSBmaXJzdCBjYW5kaWRhdGUgdGhhdCBtYXRjaGVzLiAqL1xuZXhwb3J0IGZ1bmN0aW9uIHJlc29sdmVBbGwobmFtZTogU2VsZWN0b3JOYW1lLCByb290OiBQYXJlbnROb2RlID0gZG9jdW1lbnQpOiBFbGVtZW50W10ge1xuICBjb25zdCB3aXRoaW4gPSBzY29wZShuYW1lLCByb290KTtcbiAgaWYgKCF3aXRoaW4pIHJldHVybiBbXTtcbiAgY29uc3QgeyBjYW5kaWRhdGVzLCB1bmlvbiB9ID0gZW50cnkobmFtZSk7XG4gIGlmICh1bmlvbikgcmV0dXJuIFsuLi53aXRoaW4ucXVlcnlTZWxlY3RvckFsbChjYW5kaWRhdGVzLmpvaW4oXCIsIFwiKSldO1xuICBmb3IgKGNvbnN0IGNhbmRpZGF0ZSBvZiBjYW5kaWRhdGVzKSB7XG4gICAgY29uc3QgZWxzID0gd2l0aGluLnF1ZXJ5U2VsZWN0b3JBbGwoY2FuZGlkYXRlKTtcbiAgICBpZiAoZWxzLmxlbmd0aCA+IDApIHJldHVybiBbLi4uZWxzXTtcbiAgfVxuICByZXR1cm4gW107XG59XG5cbi8qKiBBZGQgYGRhdGEtb3Zlcmxvb2tgIHRva2VucyB0byBldmVyeSB0YWdnZWQgZW50cnkgc28gc3R5bGVzaGVldHMgbmV2ZXIgdG91Y2ggT3V0bG9vaydzIERPTSBzaGFwZS4gKi9cbmV4cG9ydCBmdW5jdGlvbiB0YWdFbGVtZW50cyhyb290OiBQYXJlbnROb2RlID0gZG9jdW1lbnQpOiB2b2lkIHtcbiAgZm9yIChjb25zdCBuYW1lIG9mIE9iamVjdC5rZXlzKFNFTEVDVE9SUykgYXMgU2VsZWN0b3JOYW1lW10pIHtcbiAgICBjb25zdCB7IHRhZywgYWxsLCB1bmlvbiB9ID0gZW50cnkobmFtZSk7XG4gICAgaWYgKCF0YWcpIGNvbnRpbnVlO1xuICAgIGNvbnN0IGVscyA9XG4gICAgICBhbGwgfHwgdW5pb24gPyByZXNvbHZlQWxsKG5hbWUsIHJvb3QpIDogW3Jlc29sdmUobmFtZSwgcm9vdCldLmZpbHRlcigoZSkgPT4gZSAhPT0gbnVsbCk7XG4gICAgZm9yIChjb25zdCBlbCBvZiBlbHMpIHtcbiAgICAgIGNvbnN0IHRva2VucyA9IG5ldyBTZXQoKGVsLmdldEF0dHJpYnV0ZShcImRhdGEtb3Zlcmxvb2tcIikgPz8gXCJcIikuc3BsaXQoXCIgXCIpLmZpbHRlcihCb29sZWFuKSk7XG4gICAgICBpZiAodG9rZW5zLmhhcyh0YWcpKSBjb250aW51ZTtcbiAgICAgIHRva2Vucy5hZGQodGFnKTtcbiAgICAgIGVsLnNldEF0dHJpYnV0ZShcImRhdGEtb3Zlcmxvb2tcIiwgWy4uLnRva2Vuc10uam9pbihcIiBcIikpO1xuICAgIH1cbiAgfVxufVxuXG5leHBvcnQgaW50ZXJmYWNlIEhlYWx0aFJlc3VsdCB7XG4gIG5hbWU6IFNlbGVjdG9yTmFtZTtcbiAgZXhwZWN0OiBTZWxlY3RvckVudHJ5W1wiZXhwZWN0XCJdO1xuICAvKiogSW5kZXggb2YgdGhlIGNhbmRpZGF0ZSB0aGF0IG1hdGNoZWQsIG9yIC0xLiBBIG5vbi16ZXJvIGluZGV4IG1lYW5zIGEgZmFsbGJhY2sgaXMgaW4gdXNlLiAqL1xuICBjYW5kaWRhdGU6IG51bWJlcjtcbiAgY291bnQ6IG51bWJlcjtcbn1cblxuZXhwb3J0IGZ1bmN0aW9uIGhlYWx0aENoZWNrKHJvb3Q6IFBhcmVudE5vZGUgPSBkb2N1bWVudCk6IEhlYWx0aFJlc3VsdFtdIHtcbiAgcmV0dXJuIChPYmplY3Qua2V5cyhTRUxFQ1RPUlMpIGFzIFNlbGVjdG9yTmFtZVtdKS5tYXAoKG5hbWUpID0+IHtcbiAgICBjb25zdCB7IGNhbmRpZGF0ZXMsIGV4cGVjdCwgdW5pb24gfSA9IGVudHJ5KG5hbWUpO1xuICAgIGNvbnN0IHdpdGhpbiA9IHNjb3BlKG5hbWUsIHJvb3QpO1xuICAgIGxldCBjYW5kaWRhdGUgPSAtMTtcbiAgICBsZXQgY291bnQgPSAwO1xuICAgIGlmICh3aXRoaW4pIHtcbiAgICAgIGNhbmRpZGF0ZSA9IGNhbmRpZGF0ZXMuZmluZEluZGV4KChjKSA9PiB3aXRoaW4ucXVlcnlTZWxlY3RvcihjKSAhPT0gbnVsbCk7XG4gICAgICBpZiAodW5pb24gJiYgY2FuZGlkYXRlID49IDApIGNvdW50ID0gcmVzb2x2ZUFsbChuYW1lLCByb290KS5sZW5ndGg7XG4gICAgICBlbHNlIGlmIChjYW5kaWRhdGUgPj0gMCkgY291bnQgPSB3aXRoaW4ucXVlcnlTZWxlY3RvckFsbChjYW5kaWRhdGVzW2NhbmRpZGF0ZV0gPz8gXCJcIikubGVuZ3RoO1xuICAgIH1cbiAgICAvLyBBIHVuaW9uIGVudHJ5IGxpc3RzIGRpc3RpbmN0IHRoaW5ncywgc28gbWF0Y2hpbmcgYSBsYXRlciBvbmUgaXMgbm90IGEgZmFsbGJhY2suXG4gICAgaWYgKHVuaW9uICYmIGNhbmRpZGF0ZSA+IDApIGNhbmRpZGF0ZSA9IDA7XG4gICAgcmV0dXJuIHsgbmFtZSwgZXhwZWN0LCBjYW5kaWRhdGUsIGNvdW50IH07XG4gIH0pO1xufVxuXG5leHBvcnQgaW50ZXJmYWNlIEZvbGRlck5vZGUge1xuICBuYW1lOiBzdHJpbmc7XG4gIC8qKiBBbmNlc3RvciBuYW1lcyBqb2luZWQgd2l0aCBcIi9cIiwgZS5nLiBcIkluYm94L1JlY2VpcHRzXCIuICovXG4gIHBhdGg6IHN0cmluZztcbiAgZWxlbWVudDogRWxlbWVudDtcbn1cblxuLyoqIFwiSW5ib3ggLSA5IGl0ZW1zICgxIHVucmVhZClcIiAtPiBcIkluYm94XCIuIFRoZSB0aXRsZSBhdHRyaWJ1dGUgY2FycmllcyBjb3VudHMuICovXG5mdW5jdGlvbiBmb2xkZXJOYW1lKGVsZW1lbnQ6IEVsZW1lbnQpOiBzdHJpbmcge1xuICBjb25zdCB0aXRsZSA9IChlbGVtZW50LmdldEF0dHJpYnV0ZShcInRpdGxlXCIpID8/IGVsZW1lbnQudGV4dENvbnRlbnQgPz8gXCJcIikudHJpbSgpO1xuICByZXR1cm4gdGl0bGUucmVwbGFjZSgvXFxzKy1cXHMrXFxkK1xccytpdGVtcz8oXFxzK1xcKFxcZCtcXHMrdW5yZWFkXFwpKT8kLywgXCJcIik7XG59XG5cbi8qKlxuICogRm9sZGVycyBjdXJyZW50bHkgcmVuZGVyZWQgaW4gdGhlIGZvbGRlciBwYW5lLCB3aXRoIHBhdGhzIGJ1aWx0IGZyb20gYXJpYS1sZXZlbC4gRm9sZGVyc1xuICogdW5kZXIgRmF2b3JpdGVzIGFyZSBza2lwcGVkICh0aGV5IGR1cGxpY2F0ZSB0aGUgbWFpbGJveCB0cmVlKSwgYXMgYXJlIHRoZSB0b3AtbGV2ZWwgbWFpbGJveFxuICogcm9vdHMsIHdob3NlIG5hbWVzIGFyZSBhY2NvdW50IGFkZHJlc3Nlcy5cbiAqIFVOVkVSSUZJRUQ6IGNvbGxhcHNlZCBzdWJmb2xkZXJzIG1heSBiZSBhYnNlbnQgZnJvbSB0aGUgRE9NLlxuICovXG5leHBvcnQgZnVuY3Rpb24gcmVhZEZvbGRlclRyZWUocm9vdDogUGFyZW50Tm9kZSA9IGRvY3VtZW50KTogRm9sZGVyTm9kZVtdIHtcbiAgY29uc3Qgc3RhY2s6IHN0cmluZ1tdID0gW107XG4gIGNvbnN0IGZvbGRlcnM6IEZvbGRlck5vZGVbXSA9IFtdO1xuICBsZXQgaW5GYXZvcml0ZXMgPSBmYWxzZTtcbiAgZm9yIChjb25zdCBlbGVtZW50IG9mIHJlc29sdmVBbGwoXCJmb2xkZXJJdGVtXCIsIHJvb3QpKSB7XG4gICAgY29uc3QgbGV2ZWwgPSBOdW1iZXIoZWxlbWVudC5nZXRBdHRyaWJ1dGUoXCJhcmlhLWxldmVsXCIpID8/IFwiMVwiKTtcbiAgICBjb25zdCBuYW1lID0gZm9sZGVyTmFtZShlbGVtZW50KTtcbiAgICBzdGFjay5sZW5ndGggPSBNYXRoLm1heChsZXZlbCAtIDEsIDApO1xuICAgIHN0YWNrLnB1c2gobmFtZSk7XG4gICAgaWYgKGxldmVsID09PSAxKSB7XG4gICAgICBpbkZhdm9yaXRlcyA9IGVsZW1lbnQuZ2V0QXR0cmlidXRlKFwiZGF0YS1mb2xkZXItbmFtZVwiKSA9PT0gXCJmYXZvcml0ZXNcIjtcbiAgICAgIGNvbnRpbnVlO1xuICAgIH1cbiAgICBpZiAoaW5GYXZvcml0ZXMpIGNvbnRpbnVlO1xuICAgIGZvbGRlcnMucHVzaCh7IG5hbWUsIHBhdGg6IHN0YWNrLnNsaWNlKDEpLmZpbHRlcihCb29sZWFuKS5qb2luKFwiL1wiKSwgZWxlbWVudCB9KTtcbiAgfVxuICByZXR1cm4gZm9sZGVycztcbn1cblxuLyoqIE1lc3NhZ2Ugcm93cyBjdXJyZW50bHkgc2VsZWN0ZWQgKG9yIG9wZW4pIGluIHRoZSBsaXN0LiAqL1xuZXhwb3J0IGZ1bmN0aW9uIHNlbGVjdGVkUm93cyhyb290OiBQYXJlbnROb2RlID0gZG9jdW1lbnQpOiBFbGVtZW50W10ge1xuICByZXR1cm4gcmVzb2x2ZUFsbChcIm1lc3NhZ2VSb3dcIiwgcm9vdCkuZmlsdGVyKFxuICAgIChyb3cpID0+IHJvdy5nZXRBdHRyaWJ1dGUoXCJhcmlhLXNlbGVjdGVkXCIpID09PSBcInRydWVcIixcbiAgKTtcbn1cblxuLyoqIE91dGxvb2sgc3RhcnRzIGFuIHVucmVhZCByb3cncyBhY2Nlc3NpYmxlIG5hbWUgd2l0aCBcIlVucmVhZFwiIChFbmdsaXNoIFVJKS4gKi9cbmV4cG9ydCBmdW5jdGlvbiBpc1VucmVhZFJvdyhyb3c6IEVsZW1lbnQpOiBib29sZWFuIHtcbiAgcmV0dXJuIC9eVW5yZWFkXFxiLy50ZXN0KHJvdy5nZXRBdHRyaWJ1dGUoXCJhcmlhLWxhYmVsXCIpID8/IFwiXCIpO1xufVxuXG4vKiogT3V0bG9vaydzIGN1cnJlbnQgbGlnaHQvZGFyayBtb2RlLCBmcm9tIHRoZSBwYWxldHRlIHZhcmlhYmxlcyBpdCB3cml0ZXMgb24gPGh0bWw+LiAqL1xuZXhwb3J0IGZ1bmN0aW9uIG91dGxvb2tDb2xvclNjaGVtZShodG1sOiBFbGVtZW50KTogXCJsaWdodFwiIHwgXCJkYXJrXCIge1xuICBjb25zdCBzdHlsZSA9IGh0bWwuZ2V0QXR0cmlidXRlKFwic3R5bGVcIikgPz8gXCJcIjtcbiAgcmV0dXJuIC8tLW5hdGl2ZS1zY3JvbGxiYXItY29sb3Itc2NoZW1lOlxccypkYXJrLy50ZXN0KHN0eWxlKSA/IFwiZGFya1wiIDogXCJsaWdodFwiO1xufVxuXG4vKiogRW50cmllcyBleHBlY3RlZCBvbiBldmVyeSBtYWlsIHZpZXcgdGhhdCBkaWQgbm90IHJlc29sdmUuIE5vbi1lbXB0eSBtZWFucyBzb21ldGhpbmcgYnJva2UuICovXG5leHBvcnQgZnVuY3Rpb24gYnJva2VuRW50cmllcyhyZXN1bHRzOiBIZWFsdGhSZXN1bHRbXSk6IFNlbGVjdG9yTmFtZVtdIHtcbiAgcmV0dXJuIHJlc3VsdHMuZmlsdGVyKChyKSA9PiByLmV4cGVjdCA9PT0gXCJhbHdheXNcIiAmJiByLmNhbmRpZGF0ZSA8IDApLm1hcCgocikgPT4gci5uYW1lKTtcbn1cbiIsICIvLyBEZXZlbG9wZXIgY2FwdHVyZTogc2VyaWFsaXNlcyB0aGUgbGl2ZSBPdXRsb29rIHBhZ2UgaW50byBvbmUgSFRNTCBmaWxlLCBzbyBhIGZpeHR1cmUgY2FuIGJlXG4vLyBtYWRlIChkb2NzL2NhcHR1cmluZy1wYWdlcy5tZCkgd2l0aG91dCBEZXZUb29scy4gU2F2ZWQgbG9jYWxseSBieSB0aGUgY2FsbGVyOyBuZXZlciBzZW50XG4vLyBhbnl3aGVyZS4gVGhlIGZpbGUgaG9sZHMgcmVhbCBtYWlsLCBzbyBpdCBnb2VzIHRocm91Z2ggc2NyaXB0cy9zY3J1Yi1maXh0dXJlLnRzIGJlZm9yZSBpdFxuLy8gZ29lcyBhbnl3aGVyZSBuZWFyIHRoZSByZXBvLlxuXG5pbXBvcnQgeyBicm9rZW5FbnRyaWVzLCBoZWFsdGhDaGVjayB9IGZyb20gXCIuL3NlbGVjdG9ycy50c1wiO1xuXG4vKiogTWVzc2FnZSB0aGUgcG9wdXAgc2VuZHMgdG8gdGhlIGNvbnRlbnQgc2NyaXB0IG9mIHRoZSBhY3RpdmUgdGFiLiAqL1xuZXhwb3J0IGNvbnN0IENBUFRVUkVfTUVTU0FHRSA9IFwib2JiLWNhcHR1cmVcIjtcblxuY29uc3QgT0JCX0FUVFJJQlVURSA9IC9eZGF0YS1vdmVybG9vay87XG4vKiogVGhlIHNjcnViYmVyIChzY3JpcHRzL3NjcnViLWZpeHR1cmUudHMpIGxvb2tzIGZvciB0aGlzIG1hcmtlciB0byBzcGxpdCBvdXQgcnVudGltZSBzdHlsZXMuICovXG5leHBvcnQgY29uc3QgUlVOVElNRV9TVFlMRV9NQVJLRVIgPSBcImRhdGEtb3Zlcmxvb2stcnVudGltZVwiO1xuXG5leHBvcnQgaW50ZXJmYWNlIENhcHR1cmVNZXRhIHtcbiAgY2FwdHVyZWRBdDogc3RyaW5nO1xuICBvYmJWZXJzaW9uOiBzdHJpbmc7XG4gIC8qKiBGaXJzdCB0d28gcGF0aCBzZWdtZW50cyBvbmx5OiBkZWVwZXIgb25lcyBjYXJyeSBtZXNzYWdlIGlkcy4gKi9cbiAgcm91dGU6IHN0cmluZztcbiAgdmlld3BvcnQ6IHN0cmluZztcbiAgc2V0dGluZ3M6IFJlY29yZDxzdHJpbmcsIHN0cmluZyB8IHVuZGVmaW5lZD47XG4gIGhlYWx0aDogeyBicm9rZW46IHN0cmluZ1tdOyBlbnRyaWVzOiB7IG5hbWU6IHN0cmluZzsgY2FuZGlkYXRlOiBudW1iZXI7IGNvdW50OiBudW1iZXIgfVtdIH07XG59XG5cbmV4cG9ydCBpbnRlcmZhY2UgQ2FwdHVyZSB7XG4gIGh0bWw6IHN0cmluZztcbiAgbmFtZTogc3RyaW5nO1xuICBtZXRhOiBDYXB0dXJlTWV0YTtcbn1cblxuLyoqIFN0eWxlIHJ1bGVzIE91dGxvb2sgaW5qZWN0ZWQgYXQgcnVudGltZSAoQ1NTLWluLUpTKSwgbWludXMgdGhlIG9uZXMgT0JCIGl0c2VsZiBhZGRlZC4gKi9cbmZ1bmN0aW9uIHJ1bnRpbWVDc3MoZG9jOiBEb2N1bWVudCk6IHN0cmluZyB7XG4gIGNvbnN0IHNoZWV0cyA9IFsuLi5kb2Muc3R5bGVTaGVldHMsIC4uLihkb2MuYWRvcHRlZFN0eWxlU2hlZXRzID8/IFtdKV0uZmlsdGVyKFxuICAgIChzaGVldCkgPT4gIShcImhyZWZcIiBpbiBzaGVldCAmJiBzaGVldC5ocmVmKSxcbiAgKTtcbiAgcmV0dXJuIHNoZWV0c1xuICAgIC5mbGF0TWFwKChzaGVldCkgPT4ge1xuICAgICAgdHJ5IHtcbiAgICAgICAgcmV0dXJuIFsuLi5zaGVldC5jc3NSdWxlc10ubWFwKChydWxlKSA9PiBydWxlLmNzc1RleHQpO1xuICAgICAgfSBjYXRjaCB7XG4gICAgICAgIHJldHVybiBbXTsgLy8gY3Jvc3Mtb3JpZ2luIHNoZWV0XG4gICAgICB9XG4gICAgfSlcbiAgICAuZmlsdGVyKChydWxlKSA9PiAhcnVsZS5pbmNsdWRlcyhcImRhdGEtb3Zlcmxvb2tcIikpXG4gICAgLmpvaW4oXCJcXG5cIik7XG59XG5cbmV4cG9ydCBmdW5jdGlvbiBjYXB0dXJlRmlsZW5hbWUobm93ID0gbmV3IERhdGUoKSk6IHN0cmluZyB7XG4gIGNvbnN0IHN0YW1wID0gbm93LnRvSVNPU3RyaW5nKCkuc2xpY2UoMCwgMTYpLnJlcGxhY2UoL1xcRC9nLCBcIlwiKTtcbiAgcmV0dXJuIGBvYmItY2FwdHVyZS0ke3N0YW1wfS5odG1sYDtcbn1cblxuZXhwb3J0IGZ1bmN0aW9uIGNhcHR1cmVQYWdlKFxuICBkb2M6IERvY3VtZW50ID0gZG9jdW1lbnQsXG4gIGxvYzogUGljazxMb2NhdGlvbiwgXCJwYXRobmFtZVwiPiA9IGxvY2F0aW9uLFxuICB2ZXJzaW9uID0gXCJkZXZcIixcbiAgbm93ID0gbmV3IERhdGUoKSxcbik6IENhcHR1cmUge1xuICBjb25zdCBsaXZlID0gZG9jLmRvY3VtZW50RWxlbWVudC5kYXRhc2V0O1xuICBjb25zdCByZXN1bHRzID0gaGVhbHRoQ2hlY2soZG9jKTtcbiAgY29uc3QgbWV0YTogQ2FwdHVyZU1ldGEgPSB7XG4gICAgY2FwdHVyZWRBdDogbm93LnRvSVNPU3RyaW5nKCksXG4gICAgb2JiVmVyc2lvbjogdmVyc2lvbixcbiAgICByb3V0ZTogbG9jLnBhdGhuYW1lLnNwbGl0KFwiL1wiKS5zbGljZSgwLCAzKS5qb2luKFwiL1wiKSxcbiAgICB2aWV3cG9ydDogYCR7ZG9jLmRlZmF1bHRWaWV3Py5pbm5lcldpZHRoID8/IDB9eCR7ZG9jLmRlZmF1bHRWaWV3Py5pbm5lckhlaWdodCA/PyAwfWAsXG4gICAgc2V0dGluZ3M6IHtcbiAgICAgIGZlYXR1cmVzOiBsaXZlLm92ZXJsb29rRmVhdHVyZXMsXG4gICAgICBwYWxldHRlOiBsaXZlLm92ZXJsb29rUGFsZXR0ZSxcbiAgICAgIG1vZGU6IGxpdmUub3Zlcmxvb2tNb2RlLFxuICAgICAgbGF5b3V0OiBsaXZlLm92ZXJsb29rTGF5b3V0LFxuICAgICAgb3V0bG9va1NjaGVtZTogbGl2ZS5vdmVybG9va091dGxvb2tTY2hlbWUsXG4gICAgICByZWFkaW5nUGFuZTogbGl2ZS5vdmVybG9va1JlYWRpbmdQYW5lLFxuICAgIH0sXG4gICAgaGVhbHRoOiB7XG4gICAgICBicm9rZW46IGJyb2tlbkVudHJpZXMocmVzdWx0cyksXG4gICAgICBlbnRyaWVzOiByZXN1bHRzLm1hcCgoeyBuYW1lLCBjYW5kaWRhdGUsIGNvdW50IH0pID0+ICh7IG5hbWUsIGNhbmRpZGF0ZSwgY291bnQgfSkpLFxuICAgIH0sXG4gIH07XG5cbiAgLy8gV29yayBvbiBhIGNvcHksIHdpdGggT0JCJ3Mgb3duIG1hcmtzIHJlbW92ZWQsIHNvIHRoZSBjYXB0dXJlIGlzIHN0b2NrIE91dGxvb2suXG4gIGNvbnN0IHJvb3QgPSBkb2MuZG9jdW1lbnRFbGVtZW50LmNsb25lTm9kZSh0cnVlKSBhcyBIVE1MRWxlbWVudDtcbiAgZm9yIChjb25zdCBlbCBvZiByb290LnF1ZXJ5U2VsZWN0b3JBbGwoXCJvdmVybG9vay1yb290XCIpKSBlbC5yZW1vdmUoKTtcbiAgZm9yIChjb25zdCBlbCBvZiBbcm9vdCwgLi4ucm9vdC5xdWVyeVNlbGVjdG9yQWxsKFwiKlwiKV0pIHtcbiAgICBmb3IgKGNvbnN0IHsgbmFtZSB9IG9mIFsuLi5lbC5hdHRyaWJ1dGVzXSkge1xuICAgICAgaWYgKE9CQl9BVFRSSUJVVEUudGVzdChuYW1lKSkgZWwucmVtb3ZlQXR0cmlidXRlKG5hbWUpO1xuICAgIH1cbiAgfVxuXG4gIGNvbnN0IGhlYWQgPSByb290LnF1ZXJ5U2VsZWN0b3IoXCJoZWFkXCIpO1xuICBjb25zdCBzdHlsZSA9IGRvYy5jcmVhdGVFbGVtZW50KFwic3R5bGVcIik7XG4gIHN0eWxlLnNldEF0dHJpYnV0ZShSVU5USU1FX1NUWUxFX01BUktFUiwgXCJcIik7XG4gIHN0eWxlLnRleHRDb250ZW50ID0gcnVudGltZUNzcyhkb2MpO1xuICBoZWFkPy5hcHBlbmQoc3R5bGUpO1xuICBjb25zdCBzY3JpcHQgPSBkb2MuY3JlYXRlRWxlbWVudChcInNjcmlwdFwiKTtcbiAgc2NyaXB0LnR5cGUgPSBcImFwcGxpY2F0aW9uL2pzb25cIjtcbiAgc2NyaXB0LmlkID0gXCJvYmItY2FwdHVyZS1tZXRhXCI7XG4gIHNjcmlwdC50ZXh0Q29udGVudCA9IEpTT04uc3RyaW5naWZ5KG1ldGEsIG51bGwsIDIpLnJlcGxhY2VBbGwoXCI8XCIsIFwiXFxcXHUwMDNjXCIpO1xuICBoZWFkPy5hcHBlbmQoc2NyaXB0KTtcblxuICByZXR1cm4geyBodG1sOiBgPCFkb2N0eXBlIGh0bWw+XFxuJHtyb290Lm91dGVySFRNTH1gLCBuYW1lOiBjYXB0dXJlRmlsZW5hbWUobm93KSwgbWV0YSB9O1xufVxuIiwgIi8vIEFjdGlvbiB0YWJsZTogd2hhdCBlYWNoIGFjdGlvbiBtZWFucyBhbmQgaG93IGl0IGV4ZWN1dGVzLiBQdXJlIGRhdGEuXG4vL1xuLy8gU3RyYXRlZ2llcyAoUExBTi5tZCwgS2V5bWFwIGxheWVyKTpcbi8vICAgcGFzc3Rocm91Z2ggXHUyMDEzIE91dGxvb2sncyBHbWFpbCBzaG9ydGN1dCBtb2RlIGFscmVhZHkgZG9lcyBpdDsgbGV0IHRoZSBrZXkgdGhyb3VnaC5cbi8vICAgZHJpdmUgICAgICAgXHUyMDEzIGNsaWNrIGEgY29udHJvbCBmb3VuZCB0aHJvdWdoIHRoZSBzZWxlY3RvciByZWdpc3RyeS4gUmVhY3QgaG9ub3VycyBjbGljaygpLlxuLy8gICAgICAgICAgICAgICAgIChEaXNwYXRjaGluZyBPdXRsb29rJ3Mgb3duIHNob3J0Y3V0IGFzIGEgc3ludGhldGljIGtleSBldmVudCBkaWQgbm90aGluZyBvblxuLy8gICAgICAgICAgICAgICAgIHRoZSBsaXZlIHNpdGUgaW4gTm92IDIwMjYsIHNvIHRoZXJlIGlzIG5vIGtleXN0cm9rZSBzdHJhdGVneS4pXG4vLyAgIHRvZ2dsZSAgICAgIFx1MjAxMyBjbGljayBhIHRvZ2dsZSBidXR0b24gKFJlYWQgLyBVbnJlYWQpIG9ubHkgaWYgdGhlIHNlbGVjdGlvbiBpc24ndCBhbHJlYWR5IGluXG4vLyAgICAgICAgICAgICAgICAgdGhlIHdhbnRlZCBzdGF0ZS5cbi8vICAgZXh0ZW5zaW9uICAgXHUyMDEzIE9CQidzIG93biBVSSAocGFsZXR0ZSwgaGVscCkuXG4vLyAgIHRvZG8gICAgICAgIFx1MjAxMyBub3QgZGVjaWRlZCB5ZXQ7IHRoZSBrZXkgaXMgc3dhbGxvd2VkIGFuZCBhIGhpbnQgc2F5cyBzby5cbi8vXG4vLyBPdXRsb29rJ3MgR21haWwtbW9kZSBrZXlzIGNvbWUgZnJvbSBpdHMgb3duIHJpYmJvbiB0b29sdGlwcyBpbiB0aGUgT2N0IDIwMjYgY2FwdHVyZTpcbi8vIEMgbmV3IG1haWwsICMgZGVsZXRlLCBFIGFyY2hpdmUsIFYgbW92ZSwgXHUyMUU3SS9cdTIxRTdVIHJlYWQvdW5yZWFkLCBCIHNub296ZSwgTCBjYXRlZ29yaXplLFxuLy8gUi9BL0YgcmVwbHkvcmVwbHkgYWxsL2ZvcndhcmQsIFx1MjMxOFogdW5kby5cblxuaW1wb3J0IHR5cGUgeyBTZWxlY3Rvck5hbWUgfSBmcm9tIFwiLi4vc2VsZWN0b3JzLnRzXCI7XG5cbmV4cG9ydCB0eXBlIEV4dGVuc2lvbkNvbW1hbmQgPSBcImhlbHBcIiB8IFwicGFsZXR0ZS1nb1wiO1xuXG5leHBvcnQgdHlwZSBTdHJhdGVneSA9XG4gIHwgeyBraW5kOiBcInBhc3N0aHJvdWdoXCIgfVxuICB8IHsga2luZDogXCJkcml2ZVwiOyB0YXJnZXQ6IFNlbGVjdG9yTmFtZSB9XG4gIHwgeyBraW5kOiBcInRvZ2dsZVwiOyB0YXJnZXQ6IFNlbGVjdG9yTmFtZTsgd2FudDogXCJyZWFkXCIgfCBcInVucmVhZFwiIH1cbiAgfCB7IGtpbmQ6IFwiZXh0ZW5zaW9uXCI7IGNvbW1hbmQ6IEV4dGVuc2lvbkNvbW1hbmQgfVxuICB8IHsga2luZDogXCJ0b2RvXCIgfTtcblxuZXhwb3J0IGludGVyZmFjZSBBY3Rpb24ge1xuICBkZXNjcmlwdGlvbjogc3RyaW5nO1xuICBzdHJhdGVneTogU3RyYXRlZ3k7XG59XG5cbmNvbnN0IHBhc3N0aHJvdWdoOiBTdHJhdGVneSA9IHsga2luZDogXCJwYXNzdGhyb3VnaFwiIH07XG5jb25zdCB0b2RvOiBTdHJhdGVneSA9IHsga2luZDogXCJ0b2RvXCIgfTtcbmNvbnN0IGRyaXZlID0gKHRhcmdldDogU2VsZWN0b3JOYW1lKTogU3RyYXRlZ3kgPT4gKHsga2luZDogXCJkcml2ZVwiLCB0YXJnZXQgfSk7XG5jb25zdCBleHRlbnNpb24gPSAoY29tbWFuZDogRXh0ZW5zaW9uQ29tbWFuZCk6IFN0cmF0ZWd5ID0+ICh7IGtpbmQ6IFwiZXh0ZW5zaW9uXCIsIGNvbW1hbmQgfSk7XG5jb25zdCByZWFkU3RhdGUgPSAod2FudDogXCJyZWFkXCIgfCBcInVucmVhZFwiKTogU3RyYXRlZ3kgPT4gKHtcbiAga2luZDogXCJ0b2dnbGVcIixcbiAgdGFyZ2V0OiBcInJlYWRVbnJlYWRCdXR0b25cIixcbiAgd2FudCxcbn0pO1xuXG5leHBvcnQgY29uc3QgQUNUSU9OUyA9IHtcbiAgY29tcG9zZTogeyBkZXNjcmlwdGlvbjogXCJDb21wb3NlXCIsIHN0cmF0ZWd5OiBwYXNzdGhyb3VnaCB9LFxuICBzZWFyY2g6IHsgZGVzY3JpcHRpb246IFwiRm9jdXMgc2VhcmNoXCIsIHN0cmF0ZWd5OiBwYXNzdGhyb3VnaCB9LFxuICBoZWxwOiB7IGRlc2NyaXB0aW9uOiBcIlNob3cga2V5Ym9hcmQgc2hvcnRjdXRzXCIsIHN0cmF0ZWd5OiBleHRlbnNpb24oXCJoZWxwXCIpIH0sXG4gIG5leHQ6IHsgZGVzY3JpcHRpb246IFwiTmV4dCBjb252ZXJzYXRpb25cIiwgc3RyYXRlZ3k6IHBhc3N0aHJvdWdoIH0sXG4gIHByZXZpb3VzOiB7IGRlc2NyaXB0aW9uOiBcIlByZXZpb3VzIGNvbnZlcnNhdGlvblwiLCBzdHJhdGVneTogcGFzc3Rocm91Z2ggfSxcbiAgb3BlbjogeyBkZXNjcmlwdGlvbjogXCJPcGVuIGNvbnZlcnNhdGlvblwiLCBzdHJhdGVneTogcGFzc3Rocm91Z2ggfSxcbiAgc2VsZWN0OiB7IGRlc2NyaXB0aW9uOiBcIlNlbGVjdCBvciBkZXNlbGVjdFwiLCBzdHJhdGVneTogcGFzc3Rocm91Z2ggfSxcbiAgc2VsZWN0UmFuZ2U6IHsgZGVzY3JpcHRpb246IFwiU2VsZWN0IHJhbmdlIHNpbmNlIGxhc3Qgc2VsZWN0aW9uXCIsIHN0cmF0ZWd5OiB0b2RvIH0sXG4gIHNlbGVjdEFsbDogeyBkZXNjcmlwdGlvbjogXCJTZWxlY3QgYWxsXCIsIHN0cmF0ZWd5OiB0b2RvIH0sXG4gIHNlbGVjdE5vbmU6IHsgZGVzY3JpcHRpb246IFwiU2VsZWN0IG5vbmVcIiwgc3RyYXRlZ3k6IHRvZG8gfSxcbiAgYXJjaGl2ZTogeyBkZXNjcmlwdGlvbjogXCJBcmNoaXZlXCIsIHN0cmF0ZWd5OiBkcml2ZShcImFyY2hpdmVCdXR0b25cIikgfSxcbiAgYXJjaGl2ZVByZXZpb3VzOiB7IGRlc2NyaXB0aW9uOiBcIkFyY2hpdmUsIHRoZW4gb3BlbiBwcmV2aW91c1wiLCBzdHJhdGVneTogdG9kbyB9LFxuICBhcmNoaXZlTmV4dDogeyBkZXNjcmlwdGlvbjogXCJBcmNoaXZlLCB0aGVuIG9wZW4gbmV4dFwiLCBzdHJhdGVneTogdG9kbyB9LFxuICBkZWxldGU6IHsgZGVzY3JpcHRpb246IFwiRGVsZXRlXCIsIHN0cmF0ZWd5OiBkcml2ZShcImRlbGV0ZUJ1dHRvblwiKSB9LFxuICBqdW5rOiB7IGRlc2NyaXB0aW9uOiBcIlJlcG9ydCBqdW5rXCIsIHN0cmF0ZWd5OiBwYXNzdGhyb3VnaCB9LFxuICAvLyBPdXRsb29rJ3Mgb3duIFwiTW92ZSB0b1wiIHBpY2tlciBoYXMgZm9sZGVyIHNlYXJjaCwgd2hpY2ggY292ZXJzIEZhc3RtYWlsJ3MgYG1gIHVudGlsIHRoZVxuICAvLyBwYWxldHRlIGNhbiBtb3ZlIG1lc3NhZ2VzIGl0c2VsZiAoZG9jcy9kZWNpc2lvbnMubWQpLlxuICBtb3ZlOiB7IGRlc2NyaXB0aW9uOiBcIk1vdmUgdG8gZm9sZGVyXHUyMDI2XCIsIHN0cmF0ZWd5OiBkcml2ZShcIm1vdmVCdXR0b25cIikgfSxcbiAgZ29UbzogeyBkZXNjcmlwdGlvbjogXCJHbyB0byBmb2xkZXIgb3IgYXBwXHUyMDI2XCIsIHN0cmF0ZWd5OiBleHRlbnNpb24oXCJwYWxldHRlLWdvXCIpIH0sXG4gIHNub296ZTogeyBkZXNjcmlwdGlvbjogXCJTbm9vemVcIiwgc3RyYXRlZ3k6IHBhc3N0aHJvdWdoIH0sXG4gIG1hcmtSZWFkOiB7IGRlc2NyaXB0aW9uOiBcIk1hcmsgcmVhZFwiLCBzdHJhdGVneTogcmVhZFN0YXRlKFwicmVhZFwiKSB9LFxuICBtYXJrVW5yZWFkOiB7IGRlc2NyaXB0aW9uOiBcIk1hcmsgdW5yZWFkXCIsIHN0cmF0ZWd5OiByZWFkU3RhdGUoXCJ1bnJlYWRcIikgfSxcbiAgcmVwbHk6IHsgZGVzY3JpcHRpb246IFwiUmVwbHlcIiwgc3RyYXRlZ3k6IHBhc3N0aHJvdWdoIH0sXG4gIHJlcGx5QWxsOiB7IGRlc2NyaXB0aW9uOiBcIlJlcGx5IGFsbFwiLCBzdHJhdGVneTogcGFzc3Rocm91Z2ggfSxcbiAgZm9yd2FyZDogeyBkZXNjcmlwdGlvbjogXCJGb3J3YXJkXCIsIHN0cmF0ZWd5OiBwYXNzdGhyb3VnaCB9LFxuICB1bmRvOiB7IGRlc2NyaXB0aW9uOiBcIlVuZG9cIiwgc3RyYXRlZ3k6IGRyaXZlKFwidW5kb0J1dHRvblwiKSB9LFxuICByZWZyZXNoOiB7IGRlc2NyaXB0aW9uOiBcIlJlZnJlc2hcIiwgc3RyYXRlZ3k6IGRyaXZlKFwic3luY0J1dHRvblwiKSB9LFxuICBnb01haWw6IHsgZGVzY3JpcHRpb246IFwiR28gdG8gTWFpbFwiLCBzdHJhdGVneTogZHJpdmUoXCJtYWlsQXBwXCIpIH0sXG4gIGdvQ2FsZW5kYXI6IHsgZGVzY3JpcHRpb246IFwiR28gdG8gQ2FsZW5kYXJcIiwgc3RyYXRlZ3k6IGRyaXZlKFwiY2FsZW5kYXJBcHBcIikgfSxcbiAgZ29Db250YWN0czogeyBkZXNjcmlwdGlvbjogXCJHbyB0byBDb250YWN0c1wiLCBzdHJhdGVneTogZHJpdmUoXCJwZW9wbGVBcHBcIikgfSxcbiAgZ29TZXR0aW5nczogeyBkZXNjcmlwdGlvbjogXCJHbyB0byBTZXR0aW5nc1wiLCBzdHJhdGVneTogZHJpdmUoXCJzZXR0aW5nc0J1dHRvblwiKSB9LFxufSBhcyBjb25zdCBzYXRpc2ZpZXMgUmVjb3JkPHN0cmluZywgQWN0aW9uPjtcblxuZXhwb3J0IHR5cGUgQWN0aW9uSWQgPSBrZXlvZiB0eXBlb2YgQUNUSU9OUztcbiIsICIvLyBUaGUgRmFzdG1haWwtc3R5bGUgYmluZGluZyBzZXQgKFBMQU4ubWQsIFNob3J0Y3V0IG1hcCkuIEtleXMgYXJlIEtleWJvYXJkRXZlbnQua2V5IHZhbHVlcy5cblxuaW1wb3J0IHR5cGUgeyBBY3Rpb25JZCB9IGZyb20gXCIuL2FjdGlvbnMudHNcIjtcbmltcG9ydCB0eXBlIHsgQmluZGluZywgQ29udGV4dCB9IGZyb20gXCIuL2VuZ2luZS50c1wiO1xuXG5jb25zdCBNQUlMOiByZWFkb25seSBDb250ZXh0W10gPSBbXCJsaXN0XCIsIFwicmVhZGluZ1wiXTtcbmNvbnN0IExJU1Q6IHJlYWRvbmx5IENvbnRleHRbXSA9IFtcImxpc3RcIl07XG5jb25zdCBSRUFESU5HOiByZWFkb25seSBDb250ZXh0W10gPSBbXCJyZWFkaW5nXCJdO1xuXG5jb25zdCBiaW5kID0gKGtleXM6IHN0cmluZywgYWN0aW9uOiBBY3Rpb25JZCwgY29udGV4dHMgPSBNQUlMKTogQmluZGluZzxBY3Rpb25JZD4gPT4gKHtcbiAga2V5czoga2V5cy5zcGxpdChcIiBcIiksXG4gIGFjdGlvbixcbiAgY29udGV4dHMsXG59KTtcblxuZXhwb3J0IGNvbnN0IEJJTkRJTkdTOiByZWFkb25seSBCaW5kaW5nPEFjdGlvbklkPltdID0gW1xuICBiaW5kKFwiY1wiLCBcImNvbXBvc2VcIiksXG4gIGJpbmQoXCIvXCIsIFwic2VhcmNoXCIpLFxuICBiaW5kKFwiP1wiLCBcImhlbHBcIiksXG4gIGJpbmQoXCJqXCIsIFwibmV4dFwiKSxcbiAgYmluZChcImtcIiwgXCJwcmV2aW91c1wiKSxcbiAgYmluZChcIm9cIiwgXCJvcGVuXCIsIExJU1QpLFxuICBiaW5kKFwiRW50ZXJcIiwgXCJvcGVuXCIsIExJU1QpLFxuICBiaW5kKFwieFwiLCBcInNlbGVjdFwiLCBMSVNUKSxcbiAgYmluZChcIlhcIiwgXCJzZWxlY3RSYW5nZVwiLCBMSVNUKSxcbiAgYmluZChcIiogYVwiLCBcInNlbGVjdEFsbFwiLCBMSVNUKSxcbiAgYmluZChcIiogblwiLCBcInNlbGVjdE5vbmVcIiwgTElTVCksXG4gIGJpbmQoXCJ5XCIsIFwiYXJjaGl2ZVwiKSxcbiAgYmluZChcIl1cIiwgXCJhcmNoaXZlUHJldmlvdXNcIiwgUkVBRElORyksXG4gIGJpbmQoXCJbXCIsIFwiYXJjaGl2ZU5leHRcIiwgUkVBRElORyksXG4gIGJpbmQoXCJkXCIsIFwiZGVsZXRlXCIpLFxuICBiaW5kKFwiIVwiLCBcImp1bmtcIiksXG4gIGJpbmQoXCJtXCIsIFwibW92ZVwiKSxcbiAgYmluZChcImdcIiwgXCJnb1RvXCIpLFxuICBiaW5kKFwiYlwiLCBcInNub296ZVwiKSxcbiAgYmluZChcIi4gclwiLCBcIm1hcmtSZWFkXCIpLFxuICBiaW5kKFwiLiB1XCIsIFwibWFya1VucmVhZFwiKSxcbiAgYmluZChcInJcIiwgXCJyZXBseVwiLCBSRUFESU5HKSxcbiAgYmluZChcImFcIiwgXCJyZXBseUFsbFwiLCBSRUFESU5HKSxcbiAgYmluZChcImZcIiwgXCJmb3J3YXJkXCIsIFJFQURJTkcpLFxuICBiaW5kKFwielwiLCBcInVuZG9cIiksXG4gIGJpbmQoXCJ1XCIsIFwicmVmcmVzaFwiKSxcbiAgYmluZChcIkcgbVwiLCBcImdvTWFpbFwiKSxcbiAgYmluZChcIkcgY1wiLCBcImdvQ2FsZW5kYXJcIiksXG4gIGJpbmQoXCJHIGFcIiwgXCJnb0NvbnRhY3RzXCIpLFxuICBiaW5kKFwiRyBzXCIsIFwiZ29TZXR0aW5nc1wiKSxcbl07XG4iLCAiLy8gUHVyZSBrZXktc2VxdWVuY2UgZW5naW5lLiBObyBET00gYWNjZXNzLCBzbyBpdCBpcyBmdWxseSB1bml0LXRlc3RhYmxlLlxuLy9cbi8vIEtleXMgYXJlIEtleWJvYXJkRXZlbnQua2V5IHZhbHVlcywgd2hpY2ggYWxyZWFkeSByZWZsZWN0IFNoaWZ0IChcIkdcIiwgXCIqXCIsIFwiIVwiLCBcIj9cIikuXG4vLyBBIGJpbmRpbmcgbGlrZSBGYXN0bWFpbCdzIFwiU2hpZnQrRyB0aGVuIE1cIiBpcyB3cml0dGVuIFtcIkdcIiwgXCJtXCJdLlxuXG5leHBvcnQgdHlwZSBDb250ZXh0ID0gXCJsaXN0XCIgfCBcInJlYWRpbmdcIiB8IFwicGFsZXR0ZVwiO1xuXG5leHBvcnQgaW50ZXJmYWNlIEJpbmRpbmc8QSBleHRlbmRzIHN0cmluZyA9IHN0cmluZz4ge1xuICBrZXlzOiByZWFkb25seSBzdHJpbmdbXTtcbiAgYWN0aW9uOiBBO1xuICBjb250ZXh0czogcmVhZG9ubHkgQ29udGV4dFtdO1xufVxuXG5leHBvcnQgdHlwZSBGZWVkUmVzdWx0PEEgZXh0ZW5kcyBzdHJpbmc+ID1cbiAgfCB7IGtpbmQ6IFwibWF0Y2hcIjsgYWN0aW9uOiBBIH1cbiAgfCB7IGtpbmQ6IFwicGVuZGluZ1wiOyBwcmVmaXg6IHJlYWRvbmx5IHN0cmluZ1tdOyBuZXh0OiByZWFkb25seSBCaW5kaW5nPEE+W10gfVxuICAvKiogQSBwcmVmaXggd2FzIHBlbmRpbmcgYW5kIHRoaXMga2V5IGNvbXBsZXRlZCBub3RoaW5nLiBUaGUga2V5IHNob3VsZCBiZSBzd2FsbG93ZWQuICovXG4gIHwgeyBraW5kOiBcImNhbmNlbGxlZFwiIH1cbiAgfCB7IGtpbmQ6IFwibm9uZVwiIH07XG5cbmV4cG9ydCBjb25zdCBERUZBVUxUX1RJTUVPVVRfTVMgPSAxNTAwO1xuXG5jb25zdCBzdGFydHNXaXRoID0gKGtleXM6IHJlYWRvbmx5IHN0cmluZ1tdLCBwcmVmaXg6IHJlYWRvbmx5IHN0cmluZ1tdKSA9PlxuICBwcmVmaXgubGVuZ3RoIDw9IGtleXMubGVuZ3RoICYmIHByZWZpeC5ldmVyeSgoaywgaSkgPT4ga2V5c1tpXSA9PT0gayk7XG5cbmV4cG9ydCBjbGFzcyBTZXF1ZW5jZUVuZ2luZTxBIGV4dGVuZHMgc3RyaW5nPiB7XG4gIHJlYWRvbmx5ICNiaW5kaW5nczogcmVhZG9ubHkgQmluZGluZzxBPltdO1xuICByZWFkb25seSAjdGltZW91dE1zOiBudW1iZXI7XG4gICNwcmVmaXg6IHN0cmluZ1tdID0gW107XG4gICNsYXN0QXQgPSAwO1xuXG4gIGNvbnN0cnVjdG9yKGJpbmRpbmdzOiByZWFkb25seSBCaW5kaW5nPEE+W10sIHRpbWVvdXRNcyA9IERFRkFVTFRfVElNRU9VVF9NUykge1xuICAgIHRoaXMuI2JpbmRpbmdzID0gYmluZGluZ3M7XG4gICAgdGhpcy4jdGltZW91dE1zID0gdGltZW91dE1zO1xuICB9XG5cbiAgZ2V0IHBlbmRpbmcoKTogcmVhZG9ubHkgc3RyaW5nW10ge1xuICAgIHJldHVybiB0aGlzLiNwcmVmaXg7XG4gIH1cblxuICByZXNldCgpOiB2b2lkIHtcbiAgICB0aGlzLiNwcmVmaXggPSBbXTtcbiAgfVxuXG4gIGZlZWQoa2V5OiBzdHJpbmcsIGNvbnRleHQ6IENvbnRleHQsIG5vdzogbnVtYmVyKTogRmVlZFJlc3VsdDxBPiB7XG4gICAgaWYgKHRoaXMuI3ByZWZpeC5sZW5ndGggPiAwICYmIG5vdyAtIHRoaXMuI2xhc3RBdCA+IHRoaXMuI3RpbWVvdXRNcykgdGhpcy5yZXNldCgpO1xuICAgIGNvbnN0IHdhc1BlbmRpbmcgPSB0aGlzLiNwcmVmaXgubGVuZ3RoID4gMDtcbiAgICBjb25zdCBzZXEgPSBbLi4udGhpcy4jcHJlZml4LCBrZXldO1xuICAgIGNvbnN0IGFjdGl2ZSA9IHRoaXMuI2JpbmRpbmdzLmZpbHRlcigoYikgPT4gYi5jb250ZXh0cy5pbmNsdWRlcyhjb250ZXh0KSk7XG4gICAgY29uc3QgbG9uZ2VyID0gYWN0aXZlLmZpbHRlcigoYikgPT4gYi5rZXlzLmxlbmd0aCA+IHNlcS5sZW5ndGggJiYgc3RhcnRzV2l0aChiLmtleXMsIHNlcSkpO1xuXG4gICAgaWYgKGxvbmdlci5sZW5ndGggPiAwKSB7XG4gICAgICB0aGlzLiNwcmVmaXggPSBzZXE7XG4gICAgICB0aGlzLiNsYXN0QXQgPSBub3c7XG4gICAgICByZXR1cm4geyBraW5kOiBcInBlbmRpbmdcIiwgcHJlZml4OiBzZXEsIG5leHQ6IGxvbmdlciB9O1xuICAgIH1cblxuICAgIHRoaXMucmVzZXQoKTtcbiAgICBjb25zdCBleGFjdCA9IGFjdGl2ZS5maW5kKChiKSA9PiBiLmtleXMubGVuZ3RoID09PSBzZXEubGVuZ3RoICYmIHN0YXJ0c1dpdGgoYi5rZXlzLCBzZXEpKTtcbiAgICBpZiAoZXhhY3QpIHJldHVybiB7IGtpbmQ6IFwibWF0Y2hcIiwgYWN0aW9uOiBleGFjdC5hY3Rpb24gfTtcbiAgICByZXR1cm4gd2FzUGVuZGluZyA/IHsga2luZDogXCJjYW5jZWxsZWRcIiB9IDogeyBraW5kOiBcIm5vbmVcIiB9O1xuICB9XG59XG5cbi8qKiBCaW5kaW5ncyB0aGF0IHNoYWRvdyBlYWNoIG90aGVyIChzYW1lIGtleXMsIG9yIG9uZSBhIHByZWZpeCBvZiBhbm90aGVyKSBpbiBhIHNoYXJlZCBjb250ZXh0LiAqL1xuZXhwb3J0IGZ1bmN0aW9uIGZpbmRDb25mbGljdHMoYmluZGluZ3M6IHJlYWRvbmx5IEJpbmRpbmdbXSk6IHN0cmluZ1tdIHtcbiAgY29uc3QgY29uZmxpY3RzOiBzdHJpbmdbXSA9IFtdO1xuICBiaW5kaW5ncy5mb3JFYWNoKChhLCBpKSA9PiB7XG4gICAgZm9yIChjb25zdCBiIG9mIGJpbmRpbmdzLnNsaWNlKGkgKyAxKSkge1xuICAgICAgY29uc3Qgc2hhcmVkID0gYS5jb250ZXh0cy5zb21lKChjKSA9PiBiLmNvbnRleHRzLmluY2x1ZGVzKGMpKTtcbiAgICAgIGNvbnN0IFtzaG9ydCwgbG9uZ10gPSBhLmtleXMubGVuZ3RoIDw9IGIua2V5cy5sZW5ndGggPyBbYSwgYl0gOiBbYiwgYV07XG4gICAgICBpZiAoc2hhcmVkICYmIHN0YXJ0c1dpdGgobG9uZy5rZXlzLCBzaG9ydC5rZXlzKSkge1xuICAgICAgICBjb25mbGljdHMucHVzaChcbiAgICAgICAgICBgJHtzaG9ydC5rZXlzLmpvaW4oXCIgXCIpfSAoJHtzaG9ydC5hY3Rpb259KSB2cyAke2xvbmcua2V5cy5qb2luKFwiIFwiKX0gKCR7bG9uZy5hY3Rpb259KWAsXG4gICAgICAgICk7XG4gICAgICB9XG4gICAgfVxuICB9KTtcbiAgcmV0dXJuIGNvbmZsaWN0cztcbn1cblxuaW50ZXJmYWNlIEtleUxpa2Uge1xuICBrZXk6IHN0cmluZztcbiAgY3RybEtleTogYm9vbGVhbjtcbiAgbWV0YUtleTogYm9vbGVhbjtcbiAgYWx0S2V5OiBib29sZWFuO1xufVxuXG5jb25zdCBNT0RJRklFUl9LRVlTID0gbmV3IFNldChbXCJTaGlmdFwiLCBcIkNvbnRyb2xcIiwgXCJBbHRcIiwgXCJNZXRhXCIsIFwiQ2Fwc0xvY2tcIiwgXCJEZWFkXCJdKTtcblxuLyoqIFRoZSBlbmdpbmUga2V5IGZvciBhbiBldmVudCwgb3IgbnVsbCB3aGVuIGl0IGJlbG9uZ3MgdG8gdGhlIGJyb3dzZXIvT1MgKEN0cmwsIENtZCwgQWx0IGNob3JkcykuICovXG5leHBvcnQgZnVuY3Rpb24ga2V5RnJvbUV2ZW50KGV2ZW50OiBLZXlMaWtlKTogc3RyaW5nIHwgbnVsbCB7XG4gIGlmIChldmVudC5jdHJsS2V5IHx8IGV2ZW50Lm1ldGFLZXkgfHwgZXZlbnQuYWx0S2V5KSByZXR1cm4gbnVsbDtcbiAgaWYgKE1PRElGSUVSX0tFWVMuaGFzKGV2ZW50LmtleSkpIHJldHVybiBudWxsO1xuICByZXR1cm4gZXZlbnQua2V5O1xufVxuIiwgIjpob3N0IHtcbiAgYWxsOiBpbml0aWFsO1xuICAtLW9sLWJnOiAjZmZmO1xuICAtLW9sLWZnOiAjMWYxZjFmO1xuICAtLW9sLW11dGVkOiAjNmI2YjZiO1xuICAtLW9sLWFjY2VudDogIzJmNmZkZTtcbiAgLS1vbC1ib3JkZXI6IHJnYigwIDAgMCAvIDEyJSk7XG4gIC0tb2wtc2hhZG93OiAwIDhweCAzMnB4IHJnYigwIDAgMCAvIDE4JSk7XG4gIGZvbnQ6XG4gICAgMTNweCAvIDEuNCBzeXN0ZW0tdWksXG4gICAgLWFwcGxlLXN5c3RlbSxcbiAgICBcIlNlZ29lIFVJXCIsXG4gICAgc2Fucy1zZXJpZjtcbn1cblxuQG1lZGlhIChwcmVmZXJzLWNvbG9yLXNjaGVtZTogZGFyaykge1xuICA6aG9zdCB7XG4gICAgLS1vbC1iZzogIzI2MjYyNjtcbiAgICAtLW9sLWZnOiAjZWNlY2VjO1xuICAgIC0tb2wtbXV0ZWQ6ICNhMGEwYTA7XG4gICAgLS1vbC1hY2NlbnQ6ICM3YWE3ZmY7XG4gICAgLS1vbC1ib3JkZXI6IHJnYigyNTUgMjU1IDI1NSAvIDE0JSk7XG4gIH1cbn1cblxuW2hpZGRlbl0ge1xuICBkaXNwbGF5OiBub25lICFpbXBvcnRhbnQ7XG59XG5cbi50b2FzdCxcbi5oaW50LFxuLmhlYWx0aCB7XG4gIHBvc2l0aW9uOiBmaXhlZDtcbiAgei1pbmRleDogMjE0NzQ4MzY0NztcbiAgYmFja2dyb3VuZDogdmFyKC0tb2wtYmcpO1xuICBjb2xvcjogdmFyKC0tb2wtZmcpO1xuICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1vbC1ib3JkZXIpO1xuICBib3JkZXItcmFkaXVzOiA2cHg7XG4gIGJveC1zaGFkb3c6IHZhcigtLW9sLXNoYWRvdyk7XG4gIHBhZGRpbmc6IDZweCAxMHB4O1xufVxuXG4udG9hc3Qge1xuICBib3R0b206IDI0cHg7XG4gIGxlZnQ6IDUwJTtcbiAgdHJhbnNmb3JtOiB0cmFuc2xhdGVYKC01MCUpO1xufVxuXG4udG9hc3Qubm90aWNlIHtcbiAgZGlzcGxheTogZmxleDtcbiAgYWxpZ24taXRlbXM6IGNlbnRlcjtcbiAgZ2FwOiAxMnB4O1xuICBwYWRkaW5nOiA4cHggMTBweCA4cHggMTRweDtcbn1cblxuLmFjdGlvbiB7XG4gIGFsbDogdW5zZXQ7XG4gIGN1cnNvcjogcG9pbnRlcjtcbiAgcGFkZGluZzogNHB4IDEycHg7XG4gIGJvcmRlci1yYWRpdXM6IDk5OXB4O1xuICBiYWNrZ3JvdW5kOiB2YXIoLS1vbC1hY2NlbnQpO1xuICBjb2xvcjogI2ZmZjtcbiAgZm9udC13ZWlnaHQ6IDYwMDtcbn1cblxuLmFjdGlvbjpmb2N1cy12aXNpYmxlLFxuLmNsb3NlOmZvY3VzLXZpc2libGUge1xuICBvdXRsaW5lOiAycHggc29saWQgdmFyKC0tb2wtYWNjZW50KTtcbiAgb3V0bGluZS1vZmZzZXQ6IDJweDtcbn1cblxuLmhpbnQge1xuICBib3R0b206IDI0cHg7XG4gIHJpZ2h0OiAyNHB4O1xuICBmb250LWZhbWlseTogdWktbW9ub3NwYWNlLCBcIlNGIE1vbm9cIiwgTWVubG8sIG1vbm9zcGFjZTtcbn1cblxuLmhlYWx0aCB7XG4gIHRvcDogOHB4O1xuICByaWdodDogOHB4O1xuICBkaXNwbGF5OiBmbGV4O1xuICBnYXA6IDhweDtcbiAgYWxpZ24taXRlbXM6IGNlbnRlcjtcbiAgYm9yZGVyLWNvbG9yOiAjZDk3NzA2O1xufVxuXG4uY2xvc2Uge1xuICBhbGw6IHVuc2V0O1xuICBjdXJzb3I6IHBvaW50ZXI7XG4gIGNvbG9yOiB2YXIoLS1vbC1tdXRlZCk7XG4gIHBhZGRpbmc6IDAgMnB4O1xufVxuXG4uYmFja2Ryb3Age1xuICBwb3NpdGlvbjogZml4ZWQ7XG4gIGluc2V0OiAwO1xuICB6LWluZGV4OiAyMTQ3NDgzNjQ2O1xuICBiYWNrZ3JvdW5kOiByZ2IoMCAwIDAgLyAyNSUpO1xuICBkaXNwbGF5OiBmbGV4O1xuICBqdXN0aWZ5LWNvbnRlbnQ6IGNlbnRlcjtcbiAgYWxpZ24taXRlbXM6IGZsZXgtc3RhcnQ7XG4gIHBhZGRpbmctdG9wOiAxMnZoO1xufVxuXG4ucGFuZWwge1xuICBiYWNrZ3JvdW5kOiB2YXIoLS1vbC1iZyk7XG4gIGNvbG9yOiB2YXIoLS1vbC1mZyk7XG4gIGJvcmRlci1yYWRpdXM6IDEwcHg7XG4gIGJveC1zaGFkb3c6IHZhcigtLW9sLXNoYWRvdyk7XG4gIHdpZHRoOiBtaW4oNTIwcHgsIDkwdncpO1xuICBtYXgtaGVpZ2h0OiA3MHZoO1xuICBvdmVyZmxvdzogYXV0bztcbn1cblxuLmhlbHAge1xuICBwYWRkaW5nOiAxNnB4IDIwcHg7XG59XG5cbi5oZWxwIGgyIHtcbiAgZm9udC1zaXplOiAxNXB4O1xuICBtYXJnaW46IDAgMCAxMnB4O1xufVxuXG4uaGVscCB0YWJsZSB7XG4gIGJvcmRlci1jb2xsYXBzZTogY29sbGFwc2U7XG4gIHdpZHRoOiAxMDAlO1xufVxuXG4uaGVscCB0ZCB7XG4gIHBhZGRpbmc6IDNweCAwO1xufVxuXG4ua2V5cyB7XG4gIHdpZHRoOiA3ZW07XG59XG5cbmtiZCB7XG4gIGZvbnQtZmFtaWx5OiB1aS1tb25vc3BhY2UsIFwiU0YgTW9ub1wiLCBNZW5sbywgbW9ub3NwYWNlO1xuICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1vbC1ib3JkZXIpO1xuICBib3JkZXItcmFkaXVzOiA0cHg7XG4gIHBhZGRpbmc6IDAgNXB4O1xufVxuXG4ucGFsZXR0ZSBpbnB1dCB7XG4gIGFsbDogdW5zZXQ7XG4gIGJveC1zaXppbmc6IGJvcmRlci1ib3g7XG4gIHdpZHRoOiAxMDAlO1xuICBwYWRkaW5nOiAxMnB4IDE2cHg7XG4gIGZvbnQtc2l6ZTogMTVweDtcbiAgYm9yZGVyLWJvdHRvbTogMXB4IHNvbGlkIHZhcigtLW9sLWJvcmRlcik7XG59XG5cbi5wYWxldHRlIHVsIHtcbiAgbGlzdC1zdHlsZTogbm9uZTtcbiAgbWFyZ2luOiAwO1xuICBwYWRkaW5nOiA0cHggMDtcbn1cblxuLnBhbGV0dGUgbGkge1xuICBwYWRkaW5nOiA1cHggMTZweDtcbiAgY3Vyc29yOiBwb2ludGVyO1xufVxuXG4ucGFsZXR0ZSBsaVthcmlhLXNlbGVjdGVkPVwidHJ1ZVwiXSB7XG4gIGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1vbC1hY2NlbnQpIDE2JSwgdHJhbnNwYXJlbnQpO1xufVxuXG4ucGFsZXR0ZSBsaSAucGF0aCB7XG4gIGNvbG9yOiB2YXIoLS1vbC1tdXRlZCk7XG4gIG1hcmdpbi1sZWZ0OiA2cHg7XG59XG5cbi5wYWxldHRlIG1hcmsge1xuICBiYWNrZ3JvdW5kOiBub25lO1xuICBjb2xvcjogdmFyKC0tb2wtYWNjZW50KTtcbiAgZm9udC13ZWlnaHQ6IDYwMDtcbn1cblxuLnBhbGV0dGUgLmVtcHR5IHtcbiAgY29sb3I6IHZhcigtLW9sLW11dGVkKTtcbiAgY3Vyc29yOiBkZWZhdWx0O1xufVxuIiwgIi8vIE9CQidzIG93biBVSSBsaXZlcyBpbiBvbmUgc2hhZG93IHJvb3QsIHNvIE91dGxvb2sncyBDU1MgY2Fubm90IHJlYWNoIGl0IGFuZCBvdXJzIGNhbm5vdFxuLy8gbGVhayBpbnRvIE91dGxvb2suIEV2ZXJ5dGhpbmcgaGVyZSBpcyBwbGFpbiBET007IG5vIGZyYW1ld29yay5cblxuaW1wb3J0IGNzcyBmcm9tIFwiLi9vdmVybGF5LmNzc1wiO1xuXG5sZXQgc2hhZG93OiBTaGFkb3dSb290IHwgbnVsbCA9IG51bGw7XG5cbmV4cG9ydCBmdW5jdGlvbiBvdmVybGF5Um9vdCgpOiBTaGFkb3dSb290IHtcbiAgaWYgKHNoYWRvdykgcmV0dXJuIHNoYWRvdztcbiAgY29uc3QgaG9zdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoXCJvdmVybG9vay1yb290XCIpO1xuICBkb2N1bWVudC5kb2N1bWVudEVsZW1lbnQuYXBwZW5kKGhvc3QpO1xuICBzaGFkb3cgPSBob3N0LmF0dGFjaFNoYWRvdyh7IG1vZGU6IFwib3BlblwiIH0pO1xuICAvLyBLZXlzdHJva2VzIHR5cGVkIGludG8gb3VyIFVJIG11c3Qgbm90IHJlYWNoIE91dGxvb2sncyBkb2N1bWVudC1sZXZlbCBzaG9ydGN1dCBoYW5kbGVycyxcbiAgLy8gd2hpY2ggd291bGQgc2VlIHRoZSBob3N0IGVsZW1lbnQgKG5vdCBhbiBpbnB1dCkgYXMgdGhlIHRhcmdldC5cbiAgZm9yIChjb25zdCB0eXBlIG9mIFtcImtleWRvd25cIiwgXCJrZXlwcmVzc1wiLCBcImtleXVwXCJdKSB7XG4gICAgaG9zdC5hZGRFdmVudExpc3RlbmVyKHR5cGUsIChlKSA9PiBlLnN0b3BQcm9wYWdhdGlvbigpKTtcbiAgfVxuICBjb25zdCBzdHlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoXCJzdHlsZVwiKTtcbiAgc3R5bGUudGV4dENvbnRlbnQgPSBjc3M7XG4gIHNoYWRvdy5hcHBlbmQoc3R5bGUpO1xuICByZXR1cm4gc2hhZG93O1xufVxuXG4vKiogVHJ1ZSB3aGVuIHRoZSBldmVudCBjYW1lIGZyb20gaW5zaWRlIE9CQidzIG93biBVSS4gKi9cbmV4cG9ydCBmdW5jdGlvbiBpc0Zyb21PdmVybGF5KGV2ZW50OiBFdmVudCk6IGJvb2xlYW4ge1xuICByZXR1cm4gc2hhZG93ICE9PSBudWxsICYmIGV2ZW50LmNvbXBvc2VkUGF0aCgpLmluY2x1ZGVzKHNoYWRvdy5ob3N0KTtcbn1cblxuLyoqIFRydWUgd2hpbGUgYSBtb2RhbCAocGFsZXR0ZSBvciBoZWxwKSBpcyBvcGVuOyBpdCBvd25zIHRoZSBrZXlib2FyZCB1bnRpbCBjbG9zZWQuICovXG5leHBvcnQgZnVuY3Rpb24gaXNNb2RhbE9wZW4oKTogYm9vbGVhbiB7XG4gIHJldHVybiBzaGFkb3c/LnF1ZXJ5U2VsZWN0b3IoXCIuYmFja2Ryb3BcIikgIT0gbnVsbDtcbn1cblxuZnVuY3Rpb24gZWw8SyBleHRlbmRzIGtleW9mIEhUTUxFbGVtZW50VGFnTmFtZU1hcD4odGFnOiBLLCBjbGFzc05hbWU6IHN0cmluZywgdGV4dD86IHN0cmluZykge1xuICBjb25zdCBub2RlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCh0YWcpO1xuICBub2RlLmNsYXNzTmFtZSA9IGNsYXNzTmFtZTtcbiAgaWYgKHRleHQgIT09IHVuZGVmaW5lZCkgbm9kZS50ZXh0Q29udGVudCA9IHRleHQ7XG4gIHJldHVybiBub2RlO1xufVxuXG5mdW5jdGlvbiBzaW5nbGV0b24oY2xhc3NOYW1lOiBzdHJpbmcpOiBIVE1MRGl2RWxlbWVudCB7XG4gIGNvbnN0IHJvb3QgPSBvdmVybGF5Um9vdCgpO1xuICBjb25zdCBleGlzdGluZyA9IHJvb3QucXVlcnlTZWxlY3RvcjxIVE1MRGl2RWxlbWVudD4oYC4ke2NsYXNzTmFtZX1gKTtcbiAgaWYgKGV4aXN0aW5nKSByZXR1cm4gZXhpc3Rpbmc7XG4gIGNvbnN0IG5vZGUgPSBlbChcImRpdlwiLCBjbGFzc05hbWUpO1xuICByb290LmFwcGVuZChub2RlKTtcbiAgcmV0dXJuIG5vZGU7XG59XG5cbmxldCB0b2FzdFRpbWVyOiBSZXR1cm5UeXBlPHR5cGVvZiBzZXRUaW1lb3V0PiB8IHVuZGVmaW5lZDtcblxuZXhwb3J0IGZ1bmN0aW9uIHRvYXN0KG1lc3NhZ2U6IHN0cmluZywgbXMgPSAyNTAwKTogdm9pZCB7XG4gIGNvbnN0IG5vZGUgPSBzaW5nbGV0b24oXCJ0b2FzdFwiKTtcbiAgbm9kZS5jbGFzc0xpc3QucmVtb3ZlKFwibm90aWNlXCIpO1xuICBub2RlLnRleHRDb250ZW50ID0gbWVzc2FnZTtcbiAgbm9kZS5oaWRkZW4gPSBmYWxzZTtcbiAgY2xlYXJUaW1lb3V0KHRvYXN0VGltZXIpO1xuICB0b2FzdFRpbWVyID0gc2V0VGltZW91dCgoKSA9PiB7XG4gICAgbm9kZS5oaWRkZW4gPSB0cnVlO1xuICB9LCBtcyk7XG59XG5cbi8qKiBBIHRvYXN0IHRoYXQgc3RheXMgdW50aWwgYWN0ZWQgb246IGEgbWVzc2FnZSwgb25lIGFjdGlvbiBidXR0b24sIGFuZCBhIGRpc21pc3MgXHUwMEQ3LiAqL1xuZXhwb3J0IGZ1bmN0aW9uIG5vdGljZShtZXNzYWdlOiBzdHJpbmcsIGFjdGlvbjogc3RyaW5nLCBvbkFjdGlvbjogKCkgPT4gdm9pZCk6IHZvaWQge1xuICBjb25zdCBub2RlID0gc2luZ2xldG9uKFwidG9hc3RcIik7XG4gIGNsZWFyVGltZW91dCh0b2FzdFRpbWVyKTtcbiAgY29uc3QgYnV0dG9uID0gZWwoXCJidXR0b25cIiwgXCJhY3Rpb25cIiwgYWN0aW9uKTtcbiAgYnV0dG9uLmFkZEV2ZW50TGlzdGVuZXIoXCJjbGlja1wiLCBvbkFjdGlvbik7XG4gIGNvbnN0IGNsb3NlID0gZWwoXCJidXR0b25cIiwgXCJjbG9zZVwiLCBcIlx1MDBEN1wiKTtcbiAgY2xvc2Uuc2V0QXR0cmlidXRlKFwiYXJpYS1sYWJlbFwiLCBcIkRpc21pc3NcIik7XG4gIGNsb3NlLmFkZEV2ZW50TGlzdGVuZXIoXCJjbGlja1wiLCAoKSA9PiB7XG4gICAgbm9kZS5oaWRkZW4gPSB0cnVlO1xuICB9KTtcbiAgbm9kZS5yZXBsYWNlQ2hpbGRyZW4oZWwoXCJzcGFuXCIsIFwiXCIsIG1lc3NhZ2UpLCBidXR0b24sIGNsb3NlKTtcbiAgbm9kZS5jbGFzc0xpc3QuYWRkKFwibm90aWNlXCIpO1xuICBub2RlLmhpZGRlbiA9IGZhbHNlO1xuICBidXR0b24uZm9jdXMoKTtcbn1cblxuLyoqIFRoZSBwZW5kaW5nLXByZWZpeCBoaW50LCBlLmcuIFwiRyBcdTIwMjYgbSBNYWlsIFx1MDBCNyBjIENhbGVuZGFyXCIuIFBhc3MgbnVsbCB0byBoaWRlLiAqL1xuZXhwb3J0IGZ1bmN0aW9uIHNob3dIaW50KHRleHQ6IHN0cmluZyB8IG51bGwpOiB2b2lkIHtcbiAgY29uc3Qgbm9kZSA9IHNpbmdsZXRvbihcImhpbnRcIik7XG4gIG5vZGUuaGlkZGVuID0gdGV4dCA9PT0gbnVsbDtcbiAgbm9kZS50ZXh0Q29udGVudCA9IHRleHQgPz8gXCJcIjtcbn1cblxuZXhwb3J0IGZ1bmN0aW9uIHNob3dIZWFsdGhCYWRnZShicm9rZW46IHJlYWRvbmx5IHN0cmluZ1tdLCBvbkRpc21pc3M6ICgpID0+IHZvaWQpOiB2b2lkIHtcbiAgY29uc3Qgbm9kZSA9IHNpbmdsZXRvbihcImhlYWx0aFwiKTtcbiAgbm9kZS5yZXBsYWNlQ2hpbGRyZW4oKTtcbiAgaWYgKGJyb2tlbi5sZW5ndGggPT09IDApIHtcbiAgICBub2RlLmhpZGRlbiA9IHRydWU7XG4gICAgcmV0dXJuO1xuICB9XG4gIG5vZGUuaGlkZGVuID0gZmFsc2U7XG4gIG5vZGUudGl0bGUgPSBgTm90IGZvdW5kOiAke2Jyb2tlbi5qb2luKFwiLCBcIil9LiBTZWUgdGhlIGNvbnNvbGUgZm9yIGRldGFpbHMuYDtcbiAgY29uc3QgY2xvc2UgPSBlbChcImJ1dHRvblwiLCBcImNsb3NlXCIsIFwiXHUwMEQ3XCIpO1xuICBjbG9zZS5zZXRBdHRyaWJ1dGUoXCJhcmlhLWxhYmVsXCIsIFwiRGlzbWlzc1wiKTtcbiAgY2xvc2UuYWRkRXZlbnRMaXN0ZW5lcihcImNsaWNrXCIsICgpID0+IHtcbiAgICBub2RlLmhpZGRlbiA9IHRydWU7XG4gICAgb25EaXNtaXNzKCk7XG4gIH0pO1xuICBub2RlLmFwcGVuZChlbChcInNwYW5cIiwgXCJcIiwgYE9CQjogJHticm9rZW4ubGVuZ3RofSBzZWxlY3RvcihzKSBicm9rZW5gKSwgY2xvc2UpO1xufVxuXG5leHBvcnQgZnVuY3Rpb24gc2hvd0hlbHAocm93czogcmVhZG9ubHkgeyBrZXlzOiBzdHJpbmc7IGRlc2NyaXB0aW9uOiBzdHJpbmcgfVtdKTogdm9pZCB7XG4gIGNvbnN0IHJvb3QgPSBvdmVybGF5Um9vdCgpO1xuICBjb25zdCBiYWNrZHJvcCA9IGVsKFwiZGl2XCIsIFwiYmFja2Ryb3BcIik7XG4gIGNvbnN0IHBhbmVsID0gZWwoXCJkaXZcIiwgXCJwYW5lbCBoZWxwXCIpO1xuICBwYW5lbC5hcHBlbmQoZWwoXCJoMlwiLCBcIlwiLCBcIktleWJvYXJkIHNob3J0Y3V0c1wiKSk7XG4gIGNvbnN0IHRhYmxlID0gZWwoXCJ0YWJsZVwiLCBcIlwiKTtcbiAgZm9yIChjb25zdCB7IGtleXMsIGRlc2NyaXB0aW9uIH0gb2Ygcm93cykge1xuICAgIGNvbnN0IHRyID0gZWwoXCJ0clwiLCBcIlwiKTtcbiAgICBjb25zdCBrID0gZWwoXCJ0ZFwiLCBcImtleXNcIik7XG4gICAgay5hcHBlbmQoZWwoXCJrYmRcIiwgXCJcIiwga2V5cykpO1xuICAgIHRyLmFwcGVuZChrLCBlbChcInRkXCIsIFwiXCIsIGRlc2NyaXB0aW9uKSk7XG4gICAgdGFibGUuYXBwZW5kKHRyKTtcbiAgfVxuICBwYW5lbC5hcHBlbmQodGFibGUpO1xuICBiYWNrZHJvcC5hcHBlbmQocGFuZWwpO1xuICByb290LmFwcGVuZChiYWNrZHJvcCk7XG5cbiAgY29uc3QgY2xvc2UgPSAoKSA9PiB7XG4gICAgYmFja2Ryb3AucmVtb3ZlKCk7XG4gICAgd2luZG93LnJlbW92ZUV2ZW50TGlzdGVuZXIoXCJrZXlkb3duXCIsIG9uS2V5LCB0cnVlKTtcbiAgfTtcbiAgY29uc3Qgb25LZXkgPSAoZTogS2V5Ym9hcmRFdmVudCkgPT4ge1xuICAgIGUucHJldmVudERlZmF1bHQoKTtcbiAgICBlLnN0b3BQcm9wYWdhdGlvbigpO1xuICAgIGNsb3NlKCk7XG4gIH07XG4gIGJhY2tkcm9wLmFkZEV2ZW50TGlzdGVuZXIoXCJjbGlja1wiLCAoZSkgPT4ge1xuICAgIGlmIChlLnRhcmdldCA9PT0gYmFja2Ryb3ApIGNsb3NlKCk7XG4gIH0pO1xuICB3aW5kb3cuYWRkRXZlbnRMaXN0ZW5lcihcImtleWRvd25cIiwgb25LZXksIHRydWUpO1xufVxuXG5leHBvcnQgeyBlbCB9O1xuIiwgIi8vIFB1cmUgZnV6enkgbWF0Y2hlciBmb3IgdGhlIGZvbGRlciBwYWxldHRlOiBjYXNlLWluc2Vuc2l0aXZlIHN1YnNlcXVlbmNlIG1hdGNoaW5nLCBzY29yZWQgdG9cbi8vIGZhdm91ciB3b3JkIHN0YXJ0cyBhbmQgY29uc2VjdXRpdmUgcnVucywgdGhlIHdheSBGYXN0bWFpbCdzIGFuZCBtb3N0IGVkaXRvcnMnIHBpY2tlcnMgZmVlbC5cblxuZXhwb3J0IGludGVyZmFjZSBGdXp6eU1hdGNoPFQ+IHtcbiAgaXRlbTogVDtcbiAgc2NvcmU6IG51bWJlcjtcbiAgLyoqIEluZGljZXMgaW50byB0aGUgaXRlbSdzIHRleHQgdGhhdCBtYXRjaGVkLCBmb3IgaGlnaGxpZ2h0aW5nLiAqL1xuICBpbmRpY2VzOiBudW1iZXJbXTtcbn1cblxuY29uc3QgaXNCb3VuZGFyeSA9ICh0ZXh0OiBzdHJpbmcsIGk6IG51bWJlcikgPT4gaSA9PT0gMCB8fCAvW1xccy9cXFxcLl8tXS8udGVzdCh0ZXh0W2kgLSAxXSA/PyBcIlwiKTtcblxuZnVuY3Rpb24gaXNTdWJzZXF1ZW5jZShxdWVyeTogc3RyaW5nLCB0ZXh0OiBzdHJpbmcsIGZyb206IG51bWJlcik6IGJvb2xlYW4ge1xuICBsZXQgYXQgPSBmcm9tO1xuICBmb3IgKGNvbnN0IGNoIG9mIHF1ZXJ5KSB7XG4gICAgYXQgPSB0ZXh0LmluZGV4T2YoY2gsIGF0KSArIDE7XG4gICAgaWYgKGF0ID09PSAwKSByZXR1cm4gZmFsc2U7XG4gIH1cbiAgcmV0dXJuIHRydWU7XG59XG5cbi8qKiBTY29yZSBgcXVlcnlgIGFnYWluc3QgYHRleHRgLCBvciBudWxsIGlmIGl0IGlzIG5vdCBhIHN1YnNlcXVlbmNlLiBIaWdoZXIgaXMgYmV0dGVyLiAqL1xuZXhwb3J0IGZ1bmN0aW9uIHNjb3JlKHF1ZXJ5OiBzdHJpbmcsIHRleHQ6IHN0cmluZyk6IHsgc2NvcmU6IG51bWJlcjsgaW5kaWNlczogbnVtYmVyW10gfSB8IG51bGwge1xuICBjb25zdCBxID0gcXVlcnkudG9Mb3dlckNhc2UoKTtcbiAgY29uc3QgdCA9IHRleHQudG9Mb3dlckNhc2UoKTtcbiAgaWYgKHEubGVuZ3RoID09PSAwKSByZXR1cm4geyBzY29yZTogMCwgaW5kaWNlczogW10gfTtcblxuICBjb25zdCBpbmRpY2VzOiBudW1iZXJbXSA9IFtdO1xuICBsZXQgdG90YWwgPSAwO1xuICBsZXQgZnJvbSA9IDA7XG4gIGZvciAoY29uc3QgW24sIGNoXSBvZiBbLi4ucV0uZW50cmllcygpKSB7XG4gICAgLy8gUHJlZmVyIHRoZSBuZXh0IG9jY3VycmVuY2UgYXQgYSB3b3JkIGJvdW5kYXJ5LCBhcyBsb25nIGFzIHRoZSByZXN0IG9mIHRoZSBxdWVyeSBzdGlsbFxuICAgIC8vIGZpdHMgYWZ0ZXIgaXQ7IG90aGVyd2lzZSB0YWtlIHRoZSBuZWFyZXN0IG9uZS5cbiAgICBjb25zdCByZXN0ID0gcS5zbGljZShuICsgMSk7XG4gICAgbGV0IGF0ID0gLTE7XG4gICAgZm9yIChsZXQgaSA9IHQuaW5kZXhPZihjaCwgZnJvbSk7IGkgPj0gMDsgaSA9IHQuaW5kZXhPZihjaCwgaSArIDEpKSB7XG4gICAgICBpZiAoIWlzU3Vic2VxdWVuY2UocmVzdCwgdCwgaSArIDEpKSBicmVhaztcbiAgICAgIGlmIChhdCA8IDApIGF0ID0gaTtcbiAgICAgIGlmIChpc0JvdW5kYXJ5KHRleHQsIGkpKSB7XG4gICAgICAgIGF0ID0gaTtcbiAgICAgICAgYnJlYWs7XG4gICAgICB9XG4gICAgfVxuICAgIGlmIChhdCA8IDApIHJldHVybiBudWxsO1xuICAgIGNvbnN0IHByZXYgPSBpbmRpY2VzLmF0KC0xKTtcbiAgICB0b3RhbCArPSAxO1xuICAgIGlmIChpc0JvdW5kYXJ5KHRleHQsIGF0KSkgdG90YWwgKz0gODtcbiAgICBpZiAocHJldiAhPT0gdW5kZWZpbmVkICYmIGF0ID09PSBwcmV2ICsgMSkgdG90YWwgKz0gNTtcbiAgICBpZiAocHJldiAhPT0gdW5kZWZpbmVkKSB0b3RhbCAtPSBNYXRoLm1pbihhdCAtIHByZXYgLSAxLCAzKTtcbiAgICBpbmRpY2VzLnB1c2goYXQpO1xuICAgIGZyb20gPSBhdCArIDE7XG4gIH1cbiAgLy8gU2hvcnRlciB0YXJnZXRzIHdpbiB0aWVzLCBzbyBcIkluYm94XCIgYmVhdHMgXCJJbmJveC9SZWNlaXB0cy8yMDE5XCIgZm9yIFwiaW5iXCIuXG4gIHRvdGFsIC09IHRleHQubGVuZ3RoICogMC4wMTtcbiAgcmV0dXJuIHsgc2NvcmU6IHRvdGFsLCBpbmRpY2VzIH07XG59XG5cbmV4cG9ydCBmdW5jdGlvbiBmdXp6eUZpbHRlcjxUPihcbiAgcXVlcnk6IHN0cmluZyxcbiAgaXRlbXM6IHJlYWRvbmx5IFRbXSxcbiAgdGV4dDogKGl0ZW06IFQpID0+IHN0cmluZyxcbik6IEZ1enp5TWF0Y2g8VD5bXSB7XG4gIGNvbnN0IG1hdGNoZXM6IEZ1enp5TWF0Y2g8VD5bXSA9IFtdO1xuICBmb3IgKGNvbnN0IGl0ZW0gb2YgaXRlbXMpIHtcbiAgICBjb25zdCByZXN1bHQgPSBzY29yZShxdWVyeSwgdGV4dChpdGVtKSk7XG4gICAgaWYgKHJlc3VsdCkgbWF0Y2hlcy5wdXNoKHsgaXRlbSwgLi4ucmVzdWx0IH0pO1xuICB9XG4gIHJldHVybiBtYXRjaGVzLnNvcnQoKGEsIGIpID0+IGIuc2NvcmUgLSBhLnNjb3JlKTtcbn1cbiIsICIvLyBGb2xkZXIgcGFsZXR0ZTogRmFzdG1haWwncyBgZ2AgKGdvIHRvIGZvbGRlcikgdHlwZS1haGVhZCBwaWNrZXIsIHBsdXMgT3V0bG9vaydzIGFwcHMuXG4vLyBNb3ZpbmcgKGBtYCkgdXNlcyBPdXRsb29rJ3Mgb3duIFwiTW92ZSB0b1wiIHBpY2tlciBmb3Igbm93OyBzZWUgZG9jcy9kZWNpc2lvbnMubWQuXG5cbmltcG9ydCB7IHR5cGUgRm9sZGVyTm9kZSwgcmVhZEZvbGRlclRyZWUsIHJlc29sdmUsIHR5cGUgU2VsZWN0b3JOYW1lIH0gZnJvbSBcIi4uL3NlbGVjdG9ycy50c1wiO1xuaW1wb3J0IHsgZWwsIG92ZXJsYXlSb290LCB0b2FzdCB9IGZyb20gXCIuLi91aS9vdmVybGF5LnRzXCI7XG5pbXBvcnQgeyB0eXBlIEZ1enp5TWF0Y2gsIGZ1enp5RmlsdGVyIH0gZnJvbSBcIi4vZnV6enkudHNcIjtcblxuY29uc3QgTUFYX1JFU1VMVFMgPSAxMjtcblxuZnVuY3Rpb24gaGlnaGxpZ2h0KHRleHQ6IHN0cmluZywgaW5kaWNlczogcmVhZG9ubHkgbnVtYmVyW10pOiBEb2N1bWVudEZyYWdtZW50IHtcbiAgY29uc3QgZnJhZyA9IGRvY3VtZW50LmNyZWF0ZURvY3VtZW50RnJhZ21lbnQoKTtcbiAgY29uc3QgaGl0ID0gbmV3IFNldChpbmRpY2VzKTtcbiAgZm9yIChjb25zdCBbaSwgY2hdIG9mIFsuLi50ZXh0XS5lbnRyaWVzKCkpIHtcbiAgICBmcmFnLmFwcGVuZChoaXQuaGFzKGkpID8gZWwoXCJtYXJrXCIsIFwiXCIsIGNoKSA6IGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKGNoKSk7XG4gIH1cbiAgcmV0dXJuIGZyYWc7XG59XG5cbmNvbnN0IEFQUFM6IFtzdHJpbmcsIFNlbGVjdG9yTmFtZV1bXSA9IFtcbiAgW1wiTWFpbFwiLCBcIm1haWxBcHBcIl0sXG4gIFtcIkNhbGVuZGFyXCIsIFwiY2FsZW5kYXJBcHBcIl0sXG4gIFtcIlBlb3BsZVwiLCBcInBlb3BsZUFwcFwiXSxcbl07XG5cbi8qKiBPdXRsb29rJ3MgYXBwcywgbGlzdGVkIGFsb25nc2lkZSBmb2xkZXJzIHNvIGBnYCB0aGVuIFwiY2FsXCIgc3dpdGNoZXMgdG8gQ2FsZW5kYXIuICovXG5mdW5jdGlvbiBhcHBUYXJnZXRzKCk6IEZvbGRlck5vZGVbXSB7XG4gIHJldHVybiBBUFBTLmZsYXRNYXAoKFtuYW1lLCB0YXJnZXRdKSA9PiB7XG4gICAgY29uc3QgZWxlbWVudCA9IHJlc29sdmUodGFyZ2V0KTtcbiAgICByZXR1cm4gZWxlbWVudCA/IFt7IG5hbWUsIHBhdGg6IG5hbWUsIGVsZW1lbnQgfV0gOiBbXTtcbiAgfSk7XG59XG5cbmV4cG9ydCBmdW5jdGlvbiBvcGVuRm9sZGVyUGFsZXR0ZSgpOiB2b2lkIHtcbiAgY29uc3QgZm9sZGVycyA9IFsuLi5hcHBUYXJnZXRzKCksIC4uLnJlYWRGb2xkZXJUcmVlKCldO1xuICBpZiAoZm9sZGVycy5sZW5ndGggPT09IDApIHtcbiAgICB0b2FzdChcIk9CQiBjb3VsZG4ndCBmaW5kIHRoZSBmb2xkZXIgcGFuZVwiKTtcbiAgICByZXR1cm47XG4gIH1cblxuICBjb25zdCByb290ID0gb3ZlcmxheVJvb3QoKTtcbiAgY29uc3QgYmFja2Ryb3AgPSBlbChcImRpdlwiLCBcImJhY2tkcm9wXCIpO1xuICBjb25zdCBwYW5lbCA9IGVsKFwiZGl2XCIsIFwicGFuZWwgcGFsZXR0ZVwiKTtcbiAgY29uc3QgaW5wdXQgPSBlbChcImlucHV0XCIsIFwiXCIpO1xuICBpbnB1dC5wbGFjZWhvbGRlciA9IFwiR28gdG8gZm9sZGVyLCBvciBNYWlsLCBDYWxlbmRhciwgUGVvcGxlXHUyMDI2XCI7XG4gIGlucHV0LnNldEF0dHJpYnV0ZShcImFyaWEtbGFiZWxcIiwgaW5wdXQucGxhY2Vob2xkZXIpO1xuICBjb25zdCBsaXN0ID0gZWwoXCJ1bFwiLCBcIlwiKTtcbiAgbGlzdC5zZXRBdHRyaWJ1dGUoXCJyb2xlXCIsIFwibGlzdGJveFwiKTtcbiAgcGFuZWwuYXBwZW5kKGlucHV0LCBsaXN0KTtcbiAgYmFja2Ryb3AuYXBwZW5kKHBhbmVsKTtcbiAgcm9vdC5hcHBlbmQoYmFja2Ryb3ApO1xuXG4gIGxldCBtYXRjaGVzOiBGdXp6eU1hdGNoPEZvbGRlck5vZGU+W10gPSBbXTtcbiAgbGV0IHNlbGVjdGVkID0gMDtcblxuICBjb25zdCBjbG9zZSA9ICgpID0+IGJhY2tkcm9wLnJlbW92ZSgpO1xuICBjb25zdCBjb21taXQgPSAoKSA9PiB7XG4gICAgY29uc3QgbWF0Y2ggPSBtYXRjaGVzW3NlbGVjdGVkXTtcbiAgICBjbG9zZSgpO1xuICAgIGlmIChtYXRjaCkgKG1hdGNoLml0ZW0uZWxlbWVudCBhcyBIVE1MRWxlbWVudCkuY2xpY2soKTtcbiAgfTtcblxuICBjb25zdCByZW5kZXIgPSAoKSA9PiB7XG4gICAgbWF0Y2hlcyA9IGZ1enp5RmlsdGVyKGlucHV0LnZhbHVlLCBmb2xkZXJzLCAoZikgPT4gZi5wYXRoKS5zbGljZSgwLCBNQVhfUkVTVUxUUyk7XG4gICAgc2VsZWN0ZWQgPSBNYXRoLm1pbihzZWxlY3RlZCwgTWF0aC5tYXgobWF0Y2hlcy5sZW5ndGggLSAxLCAwKSk7XG4gICAgbGlzdC5yZXBsYWNlQ2hpbGRyZW4oXG4gICAgICAuLi5tYXRjaGVzLm1hcCgobSwgaSkgPT4ge1xuICAgICAgICBjb25zdCBsaSA9IGVsKFwibGlcIiwgXCJcIik7XG4gICAgICAgIGxpLnNldEF0dHJpYnV0ZShcInJvbGVcIiwgXCJvcHRpb25cIik7XG4gICAgICAgIGxpLnNldEF0dHJpYnV0ZShcImFyaWEtc2VsZWN0ZWRcIiwgU3RyaW5nKGkgPT09IHNlbGVjdGVkKSk7XG4gICAgICAgIGxpLmFwcGVuZChoaWdobGlnaHQobS5pdGVtLnBhdGgsIG0uaW5kaWNlcykpO1xuICAgICAgICBsaS5hZGRFdmVudExpc3RlbmVyKFwibW91c2Vkb3duXCIsIChlKSA9PiB7XG4gICAgICAgICAgZS5wcmV2ZW50RGVmYXVsdCgpO1xuICAgICAgICAgIHNlbGVjdGVkID0gaTtcbiAgICAgICAgICBjb21taXQoKTtcbiAgICAgICAgfSk7XG4gICAgICAgIHJldHVybiBsaTtcbiAgICAgIH0pLFxuICAgICk7XG4gICAgaWYgKG1hdGNoZXMubGVuZ3RoID09PSAwKSBsaXN0LmFwcGVuZChlbChcImxpXCIsIFwiZW1wdHlcIiwgXCJObyBtYXRjaGluZyBmb2xkZXJzXCIpKTtcbiAgfTtcblxuICBpbnB1dC5hZGRFdmVudExpc3RlbmVyKFwiaW5wdXRcIiwgKCkgPT4ge1xuICAgIHNlbGVjdGVkID0gMDtcbiAgICByZW5kZXIoKTtcbiAgfSk7XG4gIGlucHV0LmFkZEV2ZW50TGlzdGVuZXIoXCJrZXlkb3duXCIsIChlKSA9PiB7XG4gICAgZS5zdG9wUHJvcGFnYXRpb24oKTtcbiAgICBpZiAoZS5rZXkgPT09IFwiRXNjYXBlXCIpIGNsb3NlKCk7XG4gICAgZWxzZSBpZiAoZS5rZXkgPT09IFwiRW50ZXJcIikgY29tbWl0KCk7XG4gICAgZWxzZSBpZiAoZS5rZXkgPT09IFwiQXJyb3dEb3duXCIgfHwgKGUuY3RybEtleSAmJiBlLmtleSA9PT0gXCJuXCIpKSBzZWxlY3RlZCsrO1xuICAgIGVsc2UgaWYgKGUua2V5ID09PSBcIkFycm93VXBcIiB8fCAoZS5jdHJsS2V5ICYmIGUua2V5ID09PSBcInBcIikpXG4gICAgICBzZWxlY3RlZCA9IE1hdGgubWF4KHNlbGVjdGVkIC0gMSwgMCk7XG4gICAgZWxzZSByZXR1cm47XG4gICAgZS5wcmV2ZW50RGVmYXVsdCgpO1xuICAgIHJlbmRlcigpO1xuICB9KTtcbiAgYmFja2Ryb3AuYWRkRXZlbnRMaXN0ZW5lcihcIm1vdXNlZG93blwiLCAoZSkgPT4ge1xuICAgIGlmIChlLnRhcmdldCA9PT0gYmFja2Ryb3ApIGNsb3NlKCk7XG4gIH0pO1xuXG4gIHJlbmRlcigpO1xuICBpbnB1dC5mb2N1cygpO1xufVxuIiwgIi8vIFR1cm5zIGFuIGFjdGlvbiBpbnRvIERPTSBlZmZlY3RzLiBUaGUgb25seSBrZXltYXAgZmlsZSB0aGF0IHRvdWNoZXMgdGhlIHBhZ2UuXG5cbmltcG9ydCB7IG9wZW5Gb2xkZXJQYWxldHRlIH0gZnJvbSBcIi4uL3BhbGV0dGUvcGFsZXR0ZS50c1wiO1xuaW1wb3J0IHsgaXNVbnJlYWRSb3csIHJlc29sdmUsIHNlbGVjdGVkUm93cyB9IGZyb20gXCIuLi9zZWxlY3RvcnMudHNcIjtcbmltcG9ydCB7IHNob3dIZWxwLCB0b2FzdCB9IGZyb20gXCIuLi91aS9vdmVybGF5LnRzXCI7XG5pbXBvcnQgeyBBQ1RJT05TLCB0eXBlIEFjdGlvbklkLCB0eXBlIEV4dGVuc2lvbkNvbW1hbmQgfSBmcm9tIFwiLi9hY3Rpb25zLnRzXCI7XG5pbXBvcnQgeyBCSU5ESU5HUyB9IGZyb20gXCIuL2JpbmRpbmdzLnRzXCI7XG5pbXBvcnQgdHlwZSB7IENvbnRleHQgfSBmcm9tIFwiLi9lbmdpbmUudHNcIjtcblxuLyoqIFdoaWNoIGJpbmRpbmdzIGFwcGx5LCBqdWRnZWQgZnJvbSB3aGVyZSBmb2N1cyBpcy4gKi9cbmV4cG9ydCBmdW5jdGlvbiBjdXJyZW50Q29udGV4dCgpOiBDb250ZXh0IHtcbiAgY29uc3QgcmVhZGluZyA9IHJlc29sdmUoXCJyZWFkaW5nUGFuZVwiKTtcbiAgY29uc3QgZm9jdXNlZCA9IGRvY3VtZW50LmFjdGl2ZUVsZW1lbnQ7XG4gIHJldHVybiByZWFkaW5nICYmIGZvY3VzZWQgJiYgcmVhZGluZy5jb250YWlucyhmb2N1c2VkKSA/IFwicmVhZGluZ1wiIDogXCJsaXN0XCI7XG59XG5cbmZ1bmN0aW9uIHJ1bkV4dGVuc2lvbihjb21tYW5kOiBFeHRlbnNpb25Db21tYW5kKTogdm9pZCB7XG4gIHN3aXRjaCAoY29tbWFuZCkge1xuICAgIGNhc2UgXCJoZWxwXCI6XG4gICAgICBzaG93SGVscChcbiAgICAgICAgQklORElOR1MubWFwKChiKSA9PiAoe1xuICAgICAgICAgIGtleXM6IGIua2V5cy5qb2luKFwiIFwiKSxcbiAgICAgICAgICBkZXNjcmlwdGlvbjogQUNUSU9OU1tiLmFjdGlvbl0uZGVzY3JpcHRpb24sXG4gICAgICAgIH0pKSxcbiAgICAgICk7XG4gICAgICByZXR1cm47XG4gICAgY2FzZSBcInBhbGV0dGUtZ29cIjpcbiAgICAgIG9wZW5Gb2xkZXJQYWxldHRlKCk7XG4gICAgICByZXR1cm47XG4gIH1cbn1cblxuLyoqIFJ1biBhbiBhY3Rpb24uIFJldHVybnMgZmFsc2Ugd2hlbiB0aGUga2V5IHNob3VsZCBiZSBsZWZ0IGZvciBPdXRsb29rIChwYXNzdGhyb3VnaCkuICovXG5leHBvcnQgZnVuY3Rpb24gZXhlY3V0ZShpZDogQWN0aW9uSWQpOiBib29sZWFuIHtcbiAgY29uc3QgeyBzdHJhdGVneSwgZGVzY3JpcHRpb24gfSA9IEFDVElPTlNbaWRdO1xuICBzd2l0Y2ggKHN0cmF0ZWd5LmtpbmQpIHtcbiAgICBjYXNlIFwicGFzc3Rocm91Z2hcIjpcbiAgICAgIHJldHVybiBmYWxzZTtcbiAgICBjYXNlIFwiZHJpdmVcIjoge1xuICAgICAgY29uc3QgdGFyZ2V0ID0gcmVzb2x2ZShzdHJhdGVneS50YXJnZXQpO1xuICAgICAgaWYgKHRhcmdldCBpbnN0YW5jZW9mIEhUTUxFbGVtZW50KSB0YXJnZXQuY2xpY2soKTtcbiAgICAgIGVsc2UgdG9hc3QoYCR7ZGVzY3JpcHRpb259OiBjb3VsZG4ndCBmaW5kIHRoZSBidXR0b25gKTtcbiAgICAgIHJldHVybiB0cnVlO1xuICAgIH1cbiAgICBjYXNlIFwidG9nZ2xlXCI6IHtcbiAgICAgIGNvbnN0IHJvd3MgPSBzZWxlY3RlZFJvd3MoKTtcbiAgICAgIGlmIChyb3dzLmxlbmd0aCA9PT0gMCkge1xuICAgICAgICB0b2FzdChgJHtkZXNjcmlwdGlvbn06IHNlbGVjdCBhIG1lc3NhZ2UgZmlyc3RgKTtcbiAgICAgICAgcmV0dXJuIHRydWU7XG4gICAgICB9XG4gICAgICBjb25zdCB3YW50VW5yZWFkID0gc3RyYXRlZ3kud2FudCA9PT0gXCJ1bnJlYWRcIjtcbiAgICAgIGlmIChyb3dzLmV2ZXJ5KChyb3cpID0+IGlzVW5yZWFkUm93KHJvdykgPT09IHdhbnRVbnJlYWQpKSByZXR1cm4gdHJ1ZTtcbiAgICAgIGNvbnN0IHRhcmdldCA9IHJlc29sdmUoc3RyYXRlZ3kudGFyZ2V0KTtcbiAgICAgIGlmICh0YXJnZXQgaW5zdGFuY2VvZiBIVE1MRWxlbWVudCkgdGFyZ2V0LmNsaWNrKCk7XG4gICAgICBlbHNlIHRvYXN0KGAke2Rlc2NyaXB0aW9ufTogY291bGRuJ3QgZmluZCB0aGUgYnV0dG9uYCk7XG4gICAgICByZXR1cm4gdHJ1ZTtcbiAgICB9XG4gICAgY2FzZSBcImV4dGVuc2lvblwiOlxuICAgICAgcnVuRXh0ZW5zaW9uKHN0cmF0ZWd5LmNvbW1hbmQpO1xuICAgICAgcmV0dXJuIHRydWU7XG4gICAgY2FzZSBcInRvZG9cIjpcbiAgICAgIHRvYXN0KGAke2Rlc2NyaXB0aW9ufSBpc24ndCBpbXBsZW1lbnRlZCB5ZXRgKTtcbiAgICAgIHJldHVybiB0cnVlO1xuICB9XG59XG4iLCAiLy8gVXNlciBwcmVmZXJlbmNlcy4gU3RvcmVkIGluIGNocm9tZS5zdG9yYWdlLnN5bmM7IG5vdGhpbmcgZWxzZSBpcyBldmVyIHBlcnNpc3RlZC5cbi8vXG4vLyBFYWNoIGVuYWJsZWQgZmVhdHVyZSBpZCBpcyB3cml0dGVuIGFzIGEgdG9rZW4gb24gPGh0bWwgZGF0YS1vdmVybG9vay1mZWF0dXJlcz1cIi4uLlwiPiwgYW5kIGVhY2hcbi8vIGNob2ljZSBhcyBpdHMgb3duIGF0dHJpYnV0ZSAoPGh0bWwgZGF0YS1vdmVybG9vay1wYWxldHRlPVwic29sYXJpemVkXCI+KSwgc28gc3R5bGVzaGVldHMgZ2F0ZVxuLy8gdGhlbXNlbHZlcyB3aXRoIGBbZGF0YS1vdmVybG9vay1mZWF0dXJlc349XCJkZW5zaXR5XCJdYCBvciBgW2RhdGEtb3Zlcmxvb2stcGFsZXR0ZT1cInNvbGFyaXplZFwiXWAuXG5cbmV4cG9ydCB0eXBlIEZlYXR1cmVJZCA9IGtleW9mIHR5cGVvZiBGRUFUVVJFUztcbmV4cG9ydCB0eXBlIENob2ljZUlkID0ga2V5b2YgdHlwZW9mIENIT0lDRVM7XG5leHBvcnQgdHlwZSBTZXR0aW5ncyA9IFJlY29yZDxGZWF0dXJlSWQsIGJvb2xlYW4+ICYgeyBbSyBpbiBDaG9pY2VJZF06IENob2ljZVZhbHVlPEs+IH07XG5leHBvcnQgdHlwZSBDaG9pY2VWYWx1ZTxLIGV4dGVuZHMgQ2hvaWNlSWQ+ID0ga2V5b2YgKHR5cGVvZiBDSE9JQ0VTKVtLXVtcIm9wdGlvbnNcIl07XG5cbnR5cGUgR3JvdXAgPSBcIkdlbmVyYWxcIiB8IFwiS2V5Ym9hcmRcIiB8IFwiRGVjbHV0dGVyXCIgfCBcIkxvb2tcIjtcblxuaW50ZXJmYWNlIEZlYXR1cmUge1xuICBsYWJlbDogc3RyaW5nO1xuICBncm91cDogR3JvdXA7XG4gIGRlZmF1bHQ6IGJvb2xlYW47XG59XG5cbmludGVyZmFjZSBDaG9pY2Uge1xuICBsYWJlbDogc3RyaW5nO1xuICBncm91cDogR3JvdXA7XG4gIG9wdGlvbnM6IFJlY29yZDxzdHJpbmcsIHN0cmluZz47XG4gIGRlZmF1bHQ6IHN0cmluZztcbn1cblxuZXhwb3J0IGNvbnN0IEZFQVRVUkVTID0ge1xuICBoZWFsdGhJbmRpY2F0b3I6IHtcbiAgICBsYWJlbDogXCJTaG93IGFuIGluZGljYXRvciB3aGVuIE91dGxvb2sncyBwYWdlIGNoYW5nZXMgYnJlYWsgT0JCXCIsXG4gICAgZ3JvdXA6IFwiR2VuZXJhbFwiLFxuICAgIGRlZmF1bHQ6IHRydWUsXG4gIH0sXG4gIGtleW1hcDogeyBsYWJlbDogXCJGYXN0bWFpbC1zdHlsZSBrZXlib2FyZCBzaG9ydGN1dHNcIiwgZ3JvdXA6IFwiS2V5Ym9hcmRcIiwgZGVmYXVsdDogdHJ1ZSB9LFxuICBcImhpZGUtdXBzZWxsXCI6IHsgbGFiZWw6IFwiVXBncmFkZSBhbmQgJ3RyeSB0aGUgbmV3XHUyMDI2JyBiYW5uZXJzXCIsIGdyb3VwOiBcIkRlY2x1dHRlclwiLCBkZWZhdWx0OiB0cnVlIH0sXG4gIFwiaGlkZS1hcHAtYmFubmVyc1wiOiB7XG4gICAgbGFiZWw6IFwiTm90aWNlcyBhY3Jvc3MgdGhlIHRvcCBvZiBPdXRsb29rIChwb3AtdXAsIHNpZ24taW4gYW5kIHN0b3JhZ2UgcmVtaW5kZXJzKVwiLFxuICAgIGdyb3VwOiBcIkRlY2x1dHRlclwiLFxuICAgIGRlZmF1bHQ6IGZhbHNlLFxuICB9LFxuICBcImhpZGUtaW5mby1iYXJzXCI6IHtcbiAgICBsYWJlbDogXCJOb3RpY2VzIGFib3ZlIGFuIG9wZW4gbWVzc2FnZSAoJ1lvdSByZXBsaWVkIG9uXHUyMDI2JywgJ0V4dHJhIGxpbmUgYnJlYWtzIHJlbW92ZWQnKVwiLFxuICAgIGdyb3VwOiBcIkRlY2x1dHRlclwiLFxuICAgIGRlZmF1bHQ6IGZhbHNlLFxuICB9LFxuICBcImhpZGUtY29waWxvdFwiOiB7IGxhYmVsOiBcIkNvcGlsb3QgYnV0dG9ucyBhbmQgcGFuZWxzXCIsIGdyb3VwOiBcIkRlY2x1dHRlclwiLCBkZWZhdWx0OiB0cnVlIH0sXG4gIFwiaGlkZS1oZWFkZXItZXh0cmFzXCI6IHtcbiAgICBsYWJlbDogXCJUZWFtcyBhbmQgTXkgRGF5IGJ1dHRvbnMgaW4gdGhlIGhlYWRlclwiLFxuICAgIGdyb3VwOiBcIkRlY2x1dHRlclwiLFxuICAgIGRlZmF1bHQ6IHRydWUsXG4gIH0sXG4gIFwiaGlkZS1yYWlsLWV4dHJhc1wiOiB7XG4gICAgbGFiZWw6IFwiVG8gRG8sIE5ld3NsZXR0ZXJzLCBPcmcgRXhwbG9yZXIgYW5kIE9uZURyaXZlIGluIHRoZSBhcHAgcmFpbFwiLFxuICAgIGdyb3VwOiBcIkRlY2x1dHRlclwiLFxuICAgIGRlZmF1bHQ6IHRydWUsXG4gIH0sXG4gIFwiaGlkZS1yaWJib24tZXh0cmFzXCI6IHtcbiAgICBsYWJlbDogXCJUZWFtcywgU3dlZXAsIFF1aWNrIHN0ZXBzLCBWaXZhIGFuZCBhZGQtaW4gYnV0dG9ucyBpbiB0aGUgcmliYm9uXCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogdHJ1ZSxcbiAgfSxcbiAgXCJoaWRlLWV4dGVybmFsLXRhZ1wiOiB7IGxhYmVsOiBcIidFeHRlcm5hbCcgc2VuZGVyIHRhZ3NcIiwgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsIGRlZmF1bHQ6IHRydWUgfSxcbiAgXCJoaWRlLXJvdy1hY3Rpb25zXCI6IHtcbiAgICBsYWJlbDogXCJGbGFnIGFuZCBwaW4gaWNvbnMgdGhhdCBhcHBlYXIgd2hlbiBob3ZlcmluZyBhIG1lc3NhZ2VcIixcbiAgICBncm91cDogXCJEZWNsdXR0ZXJcIixcbiAgICBkZWZhdWx0OiBmYWxzZSxcbiAgfSxcbiAgXCJoaWRlLWF2YXRhcnNcIjoge1xuICAgIGxhYmVsOiBcIlNlbmRlciBwaWN0dXJlcyBpbiB0aGUgbWVzc2FnZSBsaXN0XCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogZmFsc2UsXG4gIH0sXG4gIFwiaGlkZS1ncm91cC1oZWFkZXJzXCI6IHtcbiAgICBsYWJlbDogXCJEYXRlIGdyb3VwIGhlYWRlcnMgKCdUaGlzIE1vbnRoJykgaW4gdGhlIG1lc3NhZ2UgbGlzdFwiLFxuICAgIGdyb3VwOiBcIkRlY2x1dHRlclwiLFxuICAgIGRlZmF1bHQ6IGZhbHNlLFxuICB9LFxuICBcImhpZGUtZm9jdXNlZC10YWJzXCI6IHsgbGFiZWw6IFwiRm9jdXNlZCAvIE90aGVyIHRhYnNcIiwgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsIGRlZmF1bHQ6IGZhbHNlIH0sXG4gIFwiaGlkZS1yaWJib24tdGFic1wiOiB7XG4gICAgbGFiZWw6IFwiUmliYm9uIHRhYnMgKEhvbWUsIFZpZXcsIEhlbHApXCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogZmFsc2UsXG4gIH0sXG4gIFwiaGlkZS1hcHAtcmFpbFwiOiB7IGxhYmVsOiBcIkFwcCByYWlsIChsZWZ0IGljb24gYmFyKVwiLCBncm91cDogXCJEZWNsdXR0ZXJcIiwgZGVmYXVsdDogZmFsc2UgfSxcbiAgXCJoaWRlLXJpYmJvblwiOiB7IGxhYmVsOiBcIlJpYmJvblwiLCBncm91cDogXCJEZWNsdXR0ZXJcIiwgZGVmYXVsdDogZmFsc2UgfSxcbiAgZGVuc2l0eTogeyBsYWJlbDogXCJUaWdodGVyLCBkZW5zZXIgbGF5b3V0XCIsIGdyb3VwOiBcIkxvb2tcIiwgZGVmYXVsdDogdHJ1ZSB9LFxufSBhcyBjb25zdCBzYXRpc2ZpZXMgUmVjb3JkPHN0cmluZywgRmVhdHVyZT47XG5cbmV4cG9ydCBjb25zdCBDSE9JQ0VTID0ge1xuICBwYWxldHRlOiB7XG4gICAgbGFiZWw6IFwiQ29sb3JzXCIsXG4gICAgZ3JvdXA6IFwiTG9va1wiLFxuICAgIG9wdGlvbnM6IHtcbiAgICAgIG91dGxvb2s6IFwiT3V0bG9vaydzIG93blwiLFxuICAgICAgc29sYXJpemVkOiBcIlNvbGFyaXplZFwiLFxuICAgICAgYmFzZTE2OiBcIkJhc2UxNiBEZWZhdWx0XCIsXG4gICAgICBmYXN0bWFpbDogXCJGYXN0bWFpbFwiLFxuICAgICAgcGFwZXI6IFwiUGFwZXJcIixcbiAgICB9LFxuICAgIGRlZmF1bHQ6IFwib3V0bG9va1wiLFxuICB9LFxuICBtb2RlOiB7XG4gICAgbGFiZWw6IFwiTGlnaHQgb3IgZGFya1wiLFxuICAgIGdyb3VwOiBcIkxvb2tcIixcbiAgICBvcHRpb25zOiB7IGF1dG86IFwiRm9sbG93IE91dGxvb2tcIiwgbGlnaHQ6IFwiTGlnaHRcIiwgZGFyazogXCJEYXJrXCIgfSxcbiAgICBkZWZhdWx0OiBcImF1dG9cIixcbiAgfSxcbiAgbGF5b3V0OiB7XG4gICAgbGFiZWw6IFwiTGF5b3V0XCIsXG4gICAgZ3JvdXA6IFwiTG9va1wiLFxuICAgIG9wdGlvbnM6IHtcbiAgICAgIG91dGxvb2s6IFwiT3V0bG9vaydzIG93blwiLFxuICAgICAgY2FsbTogXCJDYWxtIHRocmVlLXBhbmVcIixcbiAgICAgIGZhc3RtYWlsOiBcIkZhc3RtYWlsIChsaXN0IGZpcnN0KVwiLFxuICAgICAgZm9jdXM6IFwiRm9jdXMgKHNpbmdsZSBjb2x1bW4pXCIsXG4gICAgfSxcbiAgICBkZWZhdWx0OiBcIm91dGxvb2tcIixcbiAgfSxcbn0gYXMgY29uc3Qgc2F0aXNmaWVzIFJlY29yZDxzdHJpbmcsIENob2ljZT47XG5cbmV4cG9ydCBjb25zdCBGRUFUVVJFX0lEUyA9IE9iamVjdC5rZXlzKEZFQVRVUkVTKSBhcyBGZWF0dXJlSWRbXTtcbmV4cG9ydCBjb25zdCBDSE9JQ0VfSURTID0gT2JqZWN0LmtleXMoQ0hPSUNFUykgYXMgQ2hvaWNlSWRbXTtcblxuZXhwb3J0IGZ1bmN0aW9uIGRlZmF1bHRTZXR0aW5ncygpOiBTZXR0aW5ncyB7XG4gIHJldHVybiB7XG4gICAgLi4uT2JqZWN0LmZyb21FbnRyaWVzKEZFQVRVUkVfSURTLm1hcCgoaWQpID0+IFtpZCwgRkVBVFVSRVNbaWRdLmRlZmF1bHRdKSksXG4gICAgLi4uT2JqZWN0LmZyb21FbnRyaWVzKENIT0lDRV9JRFMubWFwKChpZCkgPT4gW2lkLCBDSE9JQ0VTW2lkXS5kZWZhdWx0XSkpLFxuICB9IGFzIFNldHRpbmdzO1xufVxuXG4vKiogTWVyZ2Ugd2hhdGV2ZXIgaXMgc3RvcmVkIG92ZXIgdGhlIGRlZmF1bHRzLCBpZ25vcmluZyB1bmtub3duIG9yIG1pc3R5cGVkIGtleXMuICovXG5leHBvcnQgZnVuY3Rpb24gbm9ybWFsaXplKHN0b3JlZDogUmVjb3JkPHN0cmluZywgdW5rbm93bj4pOiBTZXR0aW5ncyB7XG4gIGNvbnN0IHNldHRpbmdzID0gZGVmYXVsdFNldHRpbmdzKCkgYXMgUmVjb3JkPHN0cmluZywgdW5rbm93bj47XG4gIGZvciAoY29uc3QgaWQgb2YgRkVBVFVSRV9JRFMpIHtcbiAgICBpZiAodHlwZW9mIHN0b3JlZFtpZF0gPT09IFwiYm9vbGVhblwiKSBzZXR0aW5nc1tpZF0gPSBzdG9yZWRbaWRdO1xuICB9XG4gIGZvciAoY29uc3QgaWQgb2YgQ0hPSUNFX0lEUykge1xuICAgIGNvbnN0IHZhbHVlID0gc3RvcmVkW2lkXTtcbiAgICBpZiAodHlwZW9mIHZhbHVlID09PSBcInN0cmluZ1wiICYmIHZhbHVlIGluIENIT0lDRVNbaWRdLm9wdGlvbnMpIHNldHRpbmdzW2lkXSA9IHZhbHVlO1xuICB9XG4gIHJldHVybiBzZXR0aW5ncyBhcyBTZXR0aW5ncztcbn1cblxuZXhwb3J0IGZ1bmN0aW9uIGZlYXR1cmVUb2tlbnMoc2V0dGluZ3M6IFNldHRpbmdzKTogc3RyaW5nIHtcbiAgcmV0dXJuIEZFQVRVUkVfSURTLmZpbHRlcigoaWQpID0+IHNldHRpbmdzW2lkXSkuam9pbihcIiBcIik7XG59XG5cbi8qKiBXcml0ZXMgdGhlIHNldHRpbmdzIG9udG8gPGh0bWw+IGFzIHRoZSBhdHRyaWJ1dGVzIHRoZSBzdHlsZXNoZWV0cyBrZXkgb2ZmLiAqL1xuZXhwb3J0IGZ1bmN0aW9uIGFwcGx5VG9Eb2N1bWVudChzZXR0aW5nczogU2V0dGluZ3MsIHJvb3Q6IEhUTUxFbGVtZW50KTogdm9pZCB7XG4gIHJvb3Quc2V0QXR0cmlidXRlKFwiZGF0YS1vdmVybG9vay1mZWF0dXJlc1wiLCBmZWF0dXJlVG9rZW5zKHNldHRpbmdzKSk7XG4gIGZvciAoY29uc3QgaWQgb2YgQ0hPSUNFX0lEUykgcm9vdC5zZXRBdHRyaWJ1dGUoYGRhdGEtb3Zlcmxvb2stJHtpZH1gLCBzZXR0aW5nc1tpZF0pO1xufVxuXG5leHBvcnQgYXN5bmMgZnVuY3Rpb24gbG9hZFNldHRpbmdzKCk6IFByb21pc2U8U2V0dGluZ3M+IHtcbiAgcmV0dXJuIG5vcm1hbGl6ZShhd2FpdCBjaHJvbWUuc3RvcmFnZS5zeW5jLmdldChudWxsKSk7XG59XG5cbmV4cG9ydCBhc3luYyBmdW5jdGlvbiBzYXZlU2V0dGluZyhcbiAgaWQ6IEZlYXR1cmVJZCB8IENob2ljZUlkLFxuICB2YWx1ZTogYm9vbGVhbiB8IHN0cmluZyxcbik6IFByb21pc2U8dm9pZD4ge1xuICBhd2FpdCBjaHJvbWUuc3RvcmFnZS5zeW5jLnNldCh7IFtpZF06IHZhbHVlIH0pO1xufVxuXG5leHBvcnQgZnVuY3Rpb24gb25TZXR0aW5nc0NoYW5nZWQoY2FsbGJhY2s6IChzZXR0aW5nczogU2V0dGluZ3MpID0+IHZvaWQpOiB2b2lkIHtcbiAgY2hyb21lLnN0b3JhZ2Uub25DaGFuZ2VkLmFkZExpc3RlbmVyKChfY2hhbmdlcywgYXJlYSkgPT4ge1xuICAgIGlmIChhcmVhID09PSBcInN5bmNcIikgdm9pZCBsb2FkU2V0dGluZ3MoKS50aGVuKGNhbGxiYWNrKTtcbiAgfSk7XG59XG4iLCAiLy8gRW50cnkgcG9pbnQ6IGFwcGxpZXMgc2V0dGluZ3MsIHRhZ3MgZWxlbWVudHMgZm9yIENTUywgcnVucyB0aGUgaGVhbHRoIGNoZWNrLCBtb3VudHMgdGhlIGtleW1hcC5cblxuaW1wb3J0IHsgQ0FQVFVSRV9NRVNTQUdFLCBjYXB0dXJlUGFnZSB9IGZyb20gXCIuL2NhcHR1cmUudHNcIjtcbmltcG9ydCB7IEFDVElPTlMgfSBmcm9tIFwiLi9rZXltYXAvYWN0aW9ucy50c1wiO1xuaW1wb3J0IHsgQklORElOR1MgfSBmcm9tIFwiLi9rZXltYXAvYmluZGluZ3MudHNcIjtcbmltcG9ydCB7IERFRkFVTFRfVElNRU9VVF9NUywga2V5RnJvbUV2ZW50LCBTZXF1ZW5jZUVuZ2luZSB9IGZyb20gXCIuL2tleW1hcC9lbmdpbmUudHNcIjtcbmltcG9ydCB7IGN1cnJlbnRDb250ZXh0LCBleGVjdXRlIH0gZnJvbSBcIi4va2V5bWFwL2V4ZWN1dGUudHNcIjtcbmltcG9ydCB7XG4gIGJyb2tlbkVudHJpZXMsXG4gIHR5cGUgSGVhbHRoUmVzdWx0LFxuICBoZWFsdGhDaGVjayxcbiAgb3V0bG9va0NvbG9yU2NoZW1lLFxuICByZXNvbHZlLFxuICB0YWdFbGVtZW50cyxcbn0gZnJvbSBcIi4vc2VsZWN0b3JzLnRzXCI7XG5pbXBvcnQgeyBhcHBseVRvRG9jdW1lbnQsIGxvYWRTZXR0aW5ncywgb25TZXR0aW5nc0NoYW5nZWQsIHR5cGUgU2V0dGluZ3MgfSBmcm9tIFwiLi9zZXR0aW5ncy50c1wiO1xuaW1wb3J0IHsgaXNGcm9tT3ZlcmxheSwgaXNNb2RhbE9wZW4sIG5vdGljZSwgc2hvd0hlYWx0aEJhZGdlLCBzaG93SGludCB9IGZyb20gXCIuL3VpL292ZXJsYXkudHNcIjtcblxuY29uc3QgTE9HID0gXCJbb2JiXVwiO1xuY29uc3QgQVBQX1JFQURZX1RJTUVPVVRfTVMgPSAyMF8wMDA7XG5jb25zdCBIRUFMVEhfU0VUVExFX01TID0gMV8wMDA7XG4vKiogV2FpdHMgYmV0d2VlbiByZS1jaGVja3Mgb2YgYSBmYWlsZWQgaGVhbHRoIGNoZWNrIGJlZm9yZSBpdCBpcyByZXBvcnRlZC4gKi9cbmNvbnN0IEhFQUxUSF9SRVRSWV9NUyA9IFszXzAwMCwgMTBfMDAwXTtcblxubGV0IHNldHRpbmdzOiBTZXR0aW5ncztcbmxldCBoZWFsdGhEaXNtaXNzZWQgPSBmYWxzZTtcblxuZnVuY3Rpb24gYXBwbHlTZXR0aW5ncyhuZXh0OiBTZXR0aW5ncyk6IHZvaWQge1xuICBzZXR0aW5ncyA9IG5leHQ7XG4gIGFwcGx5VG9Eb2N1bWVudChzZXR0aW5ncywgZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50KTtcbiAgc3luY0NvbG9yU2NoZW1lKCk7XG4gIGlmICghc2V0dGluZ3MuaGVhbHRoSW5kaWNhdG9yKSBzaG93SGVhbHRoQmFkZ2UoW10sICgpID0+IHt9KTtcbn1cblxuLyoqIFJlY29yZHMgd2hldGhlciBPdXRsb29rIHNob3dzIGEgcmVhZGluZyBwYW5lLCBzbyBsYXlvdXRzIG9ubHkgcmVzaGFwZSBwYW5lcyB3aGVuIGl0IGRvZXNuJ3QuICovXG5mdW5jdGlvbiBzeW5jUmVhZGluZ1BhbmUoKTogdm9pZCB7XG4gIGNvbnN0IHN0YXRlID0gcmVzb2x2ZShcInJlYWRpbmdQYW5lXCIpID8gXCJzaG93blwiIDogXCJub25lXCI7XG4gIGlmIChkb2N1bWVudC5kb2N1bWVudEVsZW1lbnQuZGF0YXNldC5vdmVybG9va1JlYWRpbmdQYW5lICE9PSBzdGF0ZSkge1xuICAgIGRvY3VtZW50LmRvY3VtZW50RWxlbWVudC5kYXRhc2V0Lm92ZXJsb29rUmVhZGluZ1BhbmUgPSBzdGF0ZTtcbiAgfVxufVxuXG4vKiogTWlycm9ycyBPdXRsb29rJ3Mgb3duIGxpZ2h0L2RhcmsgbW9kZSBvbnRvIDxodG1sPiBzbyBcIkZvbGxvdyBPdXRsb29rXCIgcGFsZXR0ZXMgY2FuIG1hdGNoIGl0LiAqL1xuZnVuY3Rpb24gc3luY0NvbG9yU2NoZW1lKCk6IHZvaWQge1xuICBjb25zdCBzY2hlbWUgPSBvdXRsb29rQ29sb3JTY2hlbWUoZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50KTtcbiAgaWYgKGRvY3VtZW50LmRvY3VtZW50RWxlbWVudC5nZXRBdHRyaWJ1dGUoXCJkYXRhLW92ZXJsb29rLW91dGxvb2stc2NoZW1lXCIpICE9PSBzY2hlbWUpIHtcbiAgICBkb2N1bWVudC5kb2N1bWVudEVsZW1lbnQuc2V0QXR0cmlidXRlKFwiZGF0YS1vdmVybG9vay1vdXRsb29rLXNjaGVtZVwiLCBzY2hlbWUpO1xuICB9XG59XG5cbmZ1bmN0aW9uIHJlcG9ydEhlYWx0aChyZXN1bHRzOiBIZWFsdGhSZXN1bHRbXSk6IHZvaWQge1xuICBjb25zdCBicm9rZW4gPSBicm9rZW5FbnRyaWVzKHJlc3VsdHMpO1xuICBjb25zdCBmYWxsYmFja3MgPSByZXN1bHRzLmZpbHRlcigocikgPT4gci5jYW5kaWRhdGUgPiAwKS5tYXAoKHIpID0+IHIubmFtZSk7XG4gIGRvY3VtZW50LmRvY3VtZW50RWxlbWVudC5kYXRhc2V0Lm92ZXJsb29rSGVhbHRoID0gYnJva2VuLmxlbmd0aCA/IFwiYnJva2VuXCIgOiBcIm9rXCI7XG4gIGNvbnNvbGUuZ3JvdXBDb2xsYXBzZWQoYCR7TE9HfSBoZWFsdGg6ICR7YnJva2VuLmxlbmd0aCA/IGAke2Jyb2tlbi5sZW5ndGh9IGJyb2tlbmAgOiBcIm9rXCJ9YCk7XG4gIGNvbnNvbGUudGFibGUocmVzdWx0cyk7XG4gIGNvbnNvbGUuZ3JvdXBFbmQoKTtcbiAgaWYgKGJyb2tlbi5sZW5ndGgpIGNvbnNvbGUud2FybihgJHtMT0d9IHNlbGVjdG9ycyBub3QgZm91bmQ6YCwgYnJva2VuLmpvaW4oXCIsIFwiKSk7XG4gIGlmIChmYWxsYmFja3MubGVuZ3RoKSBjb25zb2xlLmluZm8oYCR7TE9HfSB1c2luZyBmYWxsYmFjayBzZWxlY3RvcnMgZm9yOmAsIGZhbGxiYWNrcy5qb2luKFwiLCBcIikpO1xuICBpZiAoc2V0dGluZ3MuaGVhbHRoSW5kaWNhdG9yICYmICFoZWFsdGhEaXNtaXNzZWQpIHtcbiAgICBzaG93SGVhbHRoQmFkZ2UoYnJva2VuLCAoKSA9PiB7XG4gICAgICBoZWFsdGhEaXNtaXNzZWQgPSB0cnVlO1xuICAgIH0pO1xuICB9XG59XG5cbmxldCBoZWFsdGhUaW1lcjogUmV0dXJuVHlwZTx0eXBlb2Ygc2V0VGltZW91dD4gfCB1bmRlZmluZWQ7XG5cbi8qKlxuICogQ2hlY2tzIHRoZSByZWdpc3RyeSBvbmNlIE91dGxvb2sgaGFzIHNldHRsZWQuIFBhcnRzIG9mIHRoZSBwYWdlICh0aGUgYXBwIHJhaWwsIGZvciBvbmUpIHJlbmRlclxuICogYWZ0ZXIgdGhlIG1lc3NhZ2UgbGlzdCwgc28gYSBtaXNzIGlzIHJlLWNoZWNrZWQgYSBmZXcgdGltZXMgYmVmb3JlIGl0IGlzIHJlcG9ydGVkLlxuICovXG5mdW5jdGlvbiBzY2hlZHVsZUhlYWx0aENoZWNrKCk6IHZvaWQge1xuICBjbGVhclRpbWVvdXQoaGVhbHRoVGltZXIpO1xuICAvLyBPbmx5IHRoZSBtYWlsIHZpZXdzIGFyZSBzdXJ2ZXllZDsgQ2FsZW5kYXIgYW5kIFBlb3BsZSBsZWdpdGltYXRlbHkgbGFjayBtYWlsIGVsZW1lbnRzLlxuICBpZiAoIWxvY2F0aW9uLnBhdGhuYW1lLmluY2x1ZGVzKFwiL21haWxcIikpIHJldHVybjtcbiAgY29uc3QgYXR0ZW1wdCA9IChuOiBudW1iZXIpID0+IHtcbiAgICBjb25zdCByZXN1bHRzID0gaGVhbHRoQ2hlY2soKTtcbiAgICBjb25zdCByZXRyeUluID0gSEVBTFRIX1JFVFJZX01TW25dO1xuICAgIGlmIChicm9rZW5FbnRyaWVzKHJlc3VsdHMpLmxlbmd0aCA9PT0gMCB8fCByZXRyeUluID09PSB1bmRlZmluZWQpIHJlcG9ydEhlYWx0aChyZXN1bHRzKTtcbiAgICBlbHNlIGhlYWx0aFRpbWVyID0gc2V0VGltZW91dCgoKSA9PiBhdHRlbXB0KG4gKyAxKSwgcmV0cnlJbik7XG4gIH07XG4gIGhlYWx0aFRpbWVyID0gc2V0VGltZW91dCgoKSA9PiBhdHRlbXB0KDApLCBIRUFMVEhfU0VUVExFX01TKTtcbn1cblxuLyoqIFJlc29sdmVzIG9uY2UgT3V0bG9vaydzIG1haWwgVUkgaGFzIHJlbmRlcmVkLCBvciBhZnRlciBhIHRpbWVvdXQgc28gdGhlIGNoZWNrIHN0aWxsIHJlcG9ydHMuICovXG5mdW5jdGlvbiB3aGVuQXBwUmVhZHkoKTogUHJvbWlzZTx2b2lkPiB7XG4gIHJldHVybiBuZXcgUHJvbWlzZSgoZG9uZSkgPT4ge1xuICAgIGlmIChyZXNvbHZlKFwibWVzc2FnZUxpc3RcIikpIHJldHVybiBkb25lKCk7XG4gICAgY29uc3Qgb2JzZXJ2ZXIgPSBuZXcgTXV0YXRpb25PYnNlcnZlcigoKSA9PiB7XG4gICAgICBpZiAocmVzb2x2ZShcIm1lc3NhZ2VMaXN0XCIpKSBmaW5pc2goKTtcbiAgICB9KTtcbiAgICBjb25zdCB0aW1lciA9IHNldFRpbWVvdXQoZmluaXNoLCBBUFBfUkVBRFlfVElNRU9VVF9NUyk7XG4gICAgZnVuY3Rpb24gZmluaXNoKCkge1xuICAgICAgb2JzZXJ2ZXIuZGlzY29ubmVjdCgpO1xuICAgICAgY2xlYXJUaW1lb3V0KHRpbWVyKTtcbiAgICAgIGRvbmUoKTtcbiAgICB9XG4gICAgb2JzZXJ2ZXIub2JzZXJ2ZShkb2N1bWVudC5ib2R5LCB7IGNoaWxkTGlzdDogdHJ1ZSwgc3VidHJlZTogdHJ1ZSB9KTtcbiAgfSk7XG59XG5cbi8qKiBSZS10YWcgb24gRE9NIGNodXJuIChSZWFjdCByZS1yZW5kZXJzKSwgYW5kIHJlLWNoZWNrIGhlYWx0aCB3aGVuIHRoZSBTUEEgbmF2aWdhdGVzLiAqL1xuZnVuY3Rpb24gd2F0Y2hQYWdlKCk6IHZvaWQge1xuICBsZXQgc2NoZWR1bGVkID0gZmFsc2U7XG4gIGxldCBsYXN0VXJsID0gbG9jYXRpb24uaHJlZjtcbiAgbmV3IE11dGF0aW9uT2JzZXJ2ZXIoKCkgPT4ge1xuICAgIGlmIChzY2hlZHVsZWQpIHJldHVybjtcbiAgICBzY2hlZHVsZWQgPSB0cnVlO1xuICAgIHJlcXVlc3RBbmltYXRpb25GcmFtZSgoKSA9PiB7XG4gICAgICBzY2hlZHVsZWQgPSBmYWxzZTtcbiAgICAgIHRhZ0VsZW1lbnRzKCk7XG4gICAgICBzeW5jQ29sb3JTY2hlbWUoKTtcbiAgICAgIHN5bmNSZWFkaW5nUGFuZSgpO1xuICAgICAgaWYgKGxvY2F0aW9uLmhyZWYgIT09IGxhc3RVcmwpIHtcbiAgICAgICAgbGFzdFVybCA9IGxvY2F0aW9uLmhyZWY7XG4gICAgICAgIHNjaGVkdWxlSGVhbHRoQ2hlY2soKTtcbiAgICAgIH1cbiAgICB9KTtcbiAgfSkub2JzZXJ2ZShkb2N1bWVudC5ib2R5LCB7IGNoaWxkTGlzdDogdHJ1ZSwgc3VidHJlZTogdHJ1ZSB9KTtcbn1cblxubGV0IG9ycGhhbk5vdGljZVNob3duID0gZmFsc2U7XG5cbi8qKlxuICogQWZ0ZXIgT0JCIGlzIHJlbG9hZGVkIG9yIHVwZGF0ZWQsIHRoaXMgY29weSBvZiB0aGUgY29udGVudCBzY3JpcHQga2VlcHMgcnVubmluZyBpbiB0YWJzIHRoYXRcbiAqIHdlcmUgYWxyZWFkeSBvcGVuLCBjdXQgb2ZmIGZyb20gdGhlIGV4dGVuc2lvbi4gU2F5IHNvIG9uY2UgaW5zdGVhZCBvZiBoYWxmLXdvcmtpbmcuXG4gKi9cbmZ1bmN0aW9uIGlzT3JwaGFuZWQoKTogYm9vbGVhbiB7XG4gIGlmIChjaHJvbWUucnVudGltZT8uaWQpIHJldHVybiBmYWxzZTtcbiAgaWYgKCFvcnBoYW5Ob3RpY2VTaG93bikge1xuICAgIG9ycGhhbk5vdGljZVNob3duID0gdHJ1ZTtcbiAgICBub3RpY2UoXCJPQkIgd2FzIHVwZGF0ZWQuXCIsIFwiUmVsb2FkIHRhYlwiLCAoKSA9PiBsb2NhdGlvbi5yZWxvYWQoKSk7XG4gIH1cbiAgcmV0dXJuIHRydWU7XG59XG5cbmZ1bmN0aW9uIGlzRWRpdGFibGUodGFyZ2V0OiBFdmVudFRhcmdldCB8IHVuZGVmaW5lZCk6IGJvb2xlYW4ge1xuICBpZiAoISh0YXJnZXQgaW5zdGFuY2VvZiBIVE1MRWxlbWVudCkpIHJldHVybiBmYWxzZTtcbiAgcmV0dXJuIHRhcmdldC5pc0NvbnRlbnRFZGl0YWJsZSB8fCAvXihJTlBVVHxURVhUQVJFQXxTRUxFQ1QpJC8udGVzdCh0YXJnZXQudGFnTmFtZSk7XG59XG5cbmZ1bmN0aW9uIG1vdW50S2V5bWFwKCk6IHZvaWQge1xuICBjb25zdCBlbmdpbmUgPSBuZXcgU2VxdWVuY2VFbmdpbmUoQklORElOR1MpO1xuICBsZXQgaGludFRpbWVyOiBSZXR1cm5UeXBlPHR5cGVvZiBzZXRUaW1lb3V0PiB8IHVuZGVmaW5lZDtcbiAgd2luZG93LmFkZEV2ZW50TGlzdGVuZXIoXG4gICAgXCJrZXlkb3duXCIsXG4gICAgKGV2ZW50KSA9PiB7XG4gICAgICBpZiAoIXNldHRpbmdzLmtleW1hcCB8fCAhZXZlbnQuaXNUcnVzdGVkIHx8IGlzRnJvbU92ZXJsYXkoZXZlbnQpIHx8IGlzTW9kYWxPcGVuKCkpIHJldHVybjtcbiAgICAgIGlmIChpc09ycGhhbmVkKCkpIHJldHVybjtcbiAgICAgIGlmIChpc0VkaXRhYmxlKGV2ZW50LmNvbXBvc2VkUGF0aCgpWzBdKSkgcmV0dXJuO1xuICAgICAgY29uc3Qga2V5ID0ga2V5RnJvbUV2ZW50KGV2ZW50KTtcbiAgICAgIGlmIChrZXkgPT09IG51bGwpIHJldHVybjtcblxuICAgICAgY29uc3QgcmVzdWx0ID0gZW5naW5lLmZlZWQoa2V5LCBjdXJyZW50Q29udGV4dCgpLCBldmVudC50aW1lU3RhbXApO1xuICAgICAgaWYgKHJlc3VsdC5raW5kID09PSBcIm5vbmVcIikgcmV0dXJuO1xuICAgICAgc2hvd0hpbnQoXG4gICAgICAgIHJlc3VsdC5raW5kID09PSBcInBlbmRpbmdcIlxuICAgICAgICAgID8gYCR7cmVzdWx0LnByZWZpeC5qb2luKFwiIFwiKX0gXHUyMDI2ICR7cmVzdWx0Lm5leHRcbiAgICAgICAgICAgICAgLm1hcCgoYikgPT4gYCR7Yi5rZXlzLmF0KC0xKX0gJHtBQ1RJT05TW2IuYWN0aW9uXS5kZXNjcmlwdGlvbn1gKVxuICAgICAgICAgICAgICAuam9pbihcIiBcdTAwQjcgXCIpfWBcbiAgICAgICAgICA6IG51bGwsXG4gICAgICApO1xuICAgICAgY2xlYXJUaW1lb3V0KGhpbnRUaW1lcik7XG4gICAgICBpZiAocmVzdWx0LmtpbmQgPT09IFwicGVuZGluZ1wiKVxuICAgICAgICBoaW50VGltZXIgPSBzZXRUaW1lb3V0KCgpID0+IHNob3dIaW50KG51bGwpLCBERUZBVUxUX1RJTUVPVVRfTVMpO1xuICAgICAgaWYgKHJlc3VsdC5raW5kID09PSBcIm1hdGNoXCIgJiYgIWV4ZWN1dGUocmVzdWx0LmFjdGlvbikpIHJldHVybjtcbiAgICAgIGV2ZW50LnByZXZlbnREZWZhdWx0KCk7XG4gICAgICBldmVudC5zdG9wSW1tZWRpYXRlUHJvcGFnYXRpb24oKTtcbiAgICB9LFxuICAgIHRydWUsXG4gICk7XG59XG5cbi8qKiBEZXZlbG9wZXIgdG9vbHM6IHRoZSBwb3B1cCBhc2tzIHRoaXMgdGFiIHRvIHNhdmUgYSBjYXB0dXJlIG9mIHRoZSBwYWdlIChzcmMvY2FwdHVyZS50cykuICovXG5mdW5jdGlvbiBzZXJ2ZUNhcHR1cmVSZXF1ZXN0cygpOiB2b2lkIHtcbiAgY2hyb21lLnJ1bnRpbWUub25NZXNzYWdlLmFkZExpc3RlbmVyKChtZXNzYWdlLCBfc2VuZGVyLCByZXBseSkgPT4ge1xuICAgIGlmIChtZXNzYWdlPy50eXBlICE9PSBDQVBUVVJFX01FU1NBR0UpIHJldHVybjtcbiAgICBjb25zdCB7IGh0bWwsIG5hbWUgfSA9IGNhcHR1cmVQYWdlKGRvY3VtZW50LCBsb2NhdGlvbiwgY2hyb21lLnJ1bnRpbWUuZ2V0TWFuaWZlc3QoKS52ZXJzaW9uKTtcbiAgICBjb25zdCB1cmwgPSBVUkwuY3JlYXRlT2JqZWN0VVJMKG5ldyBCbG9iKFtodG1sXSwgeyB0eXBlOiBcInRleHQvaHRtbFwiIH0pKTtcbiAgICBjb25zdCBsaW5rID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudChcImFcIik7XG4gICAgbGluay5ocmVmID0gdXJsO1xuICAgIGxpbmsuZG93bmxvYWQgPSBuYW1lO1xuICAgIGRvY3VtZW50LmJvZHkuYXBwZW5kKGxpbmspO1xuICAgIGxpbmsuY2xpY2soKTtcbiAgICBsaW5rLnJlbW92ZSgpO1xuICAgIHNldFRpbWVvdXQoKCkgPT4gVVJMLnJldm9rZU9iamVjdFVSTCh1cmwpLCAxMF8wMDApO1xuICAgIHJlcGx5KHsgb2s6IHRydWUsIG5hbWUsIGJ5dGVzOiBodG1sLmxlbmd0aCB9KTtcbiAgfSk7XG59XG5cbmFzeW5jIGZ1bmN0aW9uIG1haW4oKTogUHJvbWlzZTx2b2lkPiB7XG4gIHNlcnZlQ2FwdHVyZVJlcXVlc3RzKCk7XG4gIGFwcGx5U2V0dGluZ3MoYXdhaXQgbG9hZFNldHRpbmdzKCkpO1xuICBvblNldHRpbmdzQ2hhbmdlZChhcHBseVNldHRpbmdzKTtcbiAgbW91bnRLZXltYXAoKTtcbiAgLy8gQ29taW5nIGJhY2sgdG8gdGhpcyB0YWIgYWZ0ZXIgcmVsb2FkaW5nIE9CQiBmcm9tIGl0cyBwb3B1cDogb2ZmZXIgdGhlIHJlZnJlc2ggcmlnaHQgYXdheS5cbiAgZG9jdW1lbnQuYWRkRXZlbnRMaXN0ZW5lcihcInZpc2liaWxpdHljaGFuZ2VcIiwgKCkgPT4ge1xuICAgIGlmICghZG9jdW1lbnQuaGlkZGVuKSBpc09ycGhhbmVkKCk7XG4gIH0pO1xuICBhd2FpdCB3aGVuQXBwUmVhZHkoKTtcbiAgdGFnRWxlbWVudHMoKTtcbiAgc3luY1JlYWRpbmdQYW5lKCk7XG4gIHNjaGVkdWxlSGVhbHRoQ2hlY2soKTtcbiAgd2F0Y2hQYWdlKCk7XG59XG5cbnZvaWQgbWFpbigpO1xuIl0sCiAgIm1hcHBpbmdzIjogIjs7O0FBeUJPLE1BQU0sWUFBWTtBQUFBO0FBQUEsSUFFdkIsVUFBVSxFQUFFLFlBQVksQ0FBQyxjQUFjLEdBQUcsUUFBUSxVQUFVLEtBQUssWUFBWTtBQUFBLElBQzdFLFdBQVcsRUFBRSxZQUFZLENBQUMsbUJBQW1CLHVCQUF1QixHQUFHLFFBQVEsU0FBUztBQUFBLElBQ3hGLFNBQVM7QUFBQSxNQUNQLFlBQVksQ0FBQyxhQUFhLG9EQUFvRDtBQUFBLE1BQzlFLFFBQVE7QUFBQSxNQUNSLEtBQUs7QUFBQSxJQUNQO0FBQUEsSUFDQSxRQUFRO0FBQUEsTUFDTixZQUFZLENBQUMsK0JBQStCLHNDQUFzQztBQUFBLE1BQ2xGLFFBQVE7QUFBQSxNQUNSLEtBQUs7QUFBQSxJQUNQO0FBQUEsSUFDQSxZQUFZO0FBQUEsTUFDVixZQUFZLENBQUMsNENBQTRDO0FBQUEsTUFDekQsUUFBUTtBQUFBLE1BQ1IsUUFBUTtBQUFBLE1BQ1IsS0FBSztBQUFBLElBQ1A7QUFBQSxJQUNBLFlBQVk7QUFBQSxNQUNWLFlBQVk7QUFBQSxRQUNWO0FBQUEsUUFDQTtBQUFBLE1BQ0Y7QUFBQSxNQUNB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
