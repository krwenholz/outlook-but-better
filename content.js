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
//# sourceMappingURL=data:application/json;base64,ewogICJ2ZXJzaW9uIjogMywKICAic291cmNlcyI6IFsiLi4vc3JjL3NlbGVjdG9ycy50cyIsICIuLi9zcmMvY2FwdHVyZS50cyIsICIuLi9zcmMva2V5bWFwL2FjdGlvbnMudHMiLCAiLi4vc3JjL2tleW1hcC9iaW5kaW5ncy50cyIsICIuLi9zcmMva2V5bWFwL2VuZ2luZS50cyIsICIuLi9zcmMvdWkvb3ZlcmxheS5jc3MiLCAiLi4vc3JjL3VpL292ZXJsYXkudHMiLCAiLi4vc3JjL3BhbGV0dGUvZnV6enkudHMiLCAiLi4vc3JjL3BhbGV0dGUvcGFsZXR0ZS50cyIsICIuLi9zcmMva2V5bWFwL2V4ZWN1dGUudHMiLCAiLi4vc3JjL3NldHRpbmdzLnRzIiwgIi4uL3NyYy9jb250ZW50LnRzIl0sCiAgInNvdXJjZXNDb250ZW50IjogWyIvLyBUaGUgb25seSBmaWxlIGFsbG93ZWQgdG8ga25vdyB3aGF0IE91dGxvb2sncyBET00gbG9va3MgbGlrZS5cbi8vXG4vLyBBbmNob3IgcHJlZmVyZW5jZSAoc2VlIFBMQU4ubWQpOiBzdHJ1Y3R1cmFsIElEcyA+IEFSSUEgcm9sZXMvbGFiZWxzID4gc2VtYW50aWMgZGF0YS0qID4gdGV4dC5cbi8vIE5ldmVyIHVzZSBnZW5lcmF0ZWQgY2xhc3MgbmFtZXM7IGlmIG9uZSBpcyB1bmF2b2lkYWJsZSwgbWFyayBpdCBgLy8gRlJBR0lMRTpgLlxuLy9cbi8vIFZlcmlmaWVkIGFnYWluc3QgYSBzYXZlZCBwYWdlIGZyb20gT2N0IDIwMjYgKHRlc3QvZml4dHVyZXMvaW5ib3guaHRtbCk7IHNlZSBkb2NzL2RvbS1zdXJ2ZXkubWQuXG4vLyBFbnRyaWVzIG1hcmtlZCBVTlZFUklGSUVEIHdlcmUgbm90IHByZXNlbnQgaW4gdGhhdCBjYXB0dXJlLiBhcmlhLWxhYmVscyBhc3N1bWUgYW4gRW5nbGlzaCBVSS5cblxuZXhwb3J0IHR5cGUgU2VsZWN0b3JOYW1lID0ga2V5b2YgdHlwZW9mIFNFTEVDVE9SUztcblxuZXhwb3J0IGludGVyZmFjZSBTZWxlY3RvckVudHJ5IHtcbiAgLyoqIENhbmRpZGF0ZSBDU1Mgc2VsZWN0b3JzLCBtb3N0IHN0YWJsZSBmaXJzdC4gVGhlIGZpcnN0IG9uZSB0aGF0IG1hdGNoZXMgd2lucy4gKi9cbiAgY2FuZGlkYXRlczogcmVhZG9ubHkgc3RyaW5nW107XG4gIC8qKiBcImFsd2F5c1wiOiBzaG91bGQgZXhpc3Qgb24gZXZlcnkgbWFpbCB2aWV3LCBzbyBhIG1pc3MgaXMgYSBicmVha2FnZS4gXCJzb21ldGltZXNcIjogY29udGV4dHVhbC4gKi9cbiAgZXhwZWN0OiBcImFsd2F5c1wiIHwgXCJzb21ldGltZXNcIjtcbiAgLyoqIE5hbWUgb2YgYW5vdGhlciBlbnRyeSB0byByZXNvbHZlIGluc2lkZSwgaW5zdGVhZCBvZiB0aGUgd2hvbGUgZG9jdW1lbnQuICovXG4gIHdpdGhpbj86IHN0cmluZztcbiAgLyoqIFRva2VuIGFkZGVkIHRvIHRoZSBlbGVtZW50J3MgYGRhdGEtb3Zlcmxvb2tgIGF0dHJpYnV0ZSBzbyBDU1MgY2FuIHRhcmdldCBpdC4gKi9cbiAgdGFnPzogc3RyaW5nO1xuICAvKiogVGFnIGV2ZXJ5IG1hdGNoIHJhdGhlciB0aGFuIG9ubHkgdGhlIGZpcnN0LiAqL1xuICBhbGw/OiBib29sZWFuO1xuICAvKiogTWF0Y2ggZXZlcnkgY2FuZGlkYXRlIChhIGxpc3Qgb2YgZGlzdGluY3QgdGhpbmdzKSwgbm90IGp1c3QgdGhlIGZpcnN0IHRoYXQgbWF0Y2hlcy4gKi9cbiAgdW5pb24/OiBib29sZWFuO1xufVxuXG5leHBvcnQgY29uc3QgU0VMRUNUT1JTID0ge1xuICAvLyAtLS0gU3RydWN0dXJlIC0tLVxuICB0aXRsZUJhcjogeyBjYW5kaWRhdGVzOiBbXCIjT3dhVGl0bGVCYXJcIl0sIGV4cGVjdDogXCJhbHdheXNcIiwgdGFnOiBcInRpdGxlLWJhclwiIH0sXG4gIHNlYXJjaEJveDogeyBjYW5kaWRhdGVzOiBbXCIjdG9wU2VhcmNoSW5wdXRcIiwgJ1tyb2xlPVwic2VhcmNoXCJdIGlucHV0J10sIGV4cGVjdDogXCJhbHdheXNcIiB9LFxuICBhcHBSYWlsOiB7XG4gICAgY2FuZGlkYXRlczogW1wiI0xlZnRSYWlsXCIsICdbcm9sZT1cIm5hdmlnYXRpb25cIl1bYXJpYS1sYWJlbD1cImxlZnQtcmFpbC1hcHBiYXJcIl0nXSxcbiAgICBleHBlY3Q6IFwiYWx3YXlzXCIsXG4gICAgdGFnOiBcImFwcC1yYWlsXCIsXG4gIH0sXG4gIHJpYmJvbjoge1xuICAgIGNhbmRpZGF0ZXM6IFsnW2RhdGEtYXBwLXNlY3Rpb249XCJSaWJib25cIl0nLCAnW3JvbGU9XCJyZWdpb25cIl1bYXJpYS1sYWJlbD1cIlJpYmJvblwiXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB0YWc6IFwicmliYm9uXCIsXG4gIH0sXG4gIHJpYmJvblRhYnM6IHtcbiAgICBjYW5kaWRhdGVzOiBbJ1tyb2xlPVwidGFibGlzdFwiXVthcmlhLWxhYmVsPVwiUmliYm9uIHRhYnNcIl0nXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgd2l0aGluOiBcInJpYmJvblwiLFxuICAgIHRhZzogXCJyaWJib24tdGFic1wiLFxuICB9LFxuICBmb2xkZXJQYW5lOiB7XG4gICAgY2FuZGlkYXRlczogW1xuICAgICAgJ1tkYXRhLWFwcC1zZWN0aW9uPVwiTmF2aWdhdGlvblBhbmVcIl0nLFxuICAgICAgJ1tyb2xlPVwicmVnaW9uXCJdW2FyaWEtbGFiZWw9XCJOYXZpZ2F0aW9uIHBhbmVcIl0nLFxuICAgIF0sXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICAgIHRhZzogXCJmb2xkZXItcGFuZVwiLFxuICB9LFxuICBmb2xkZXJJdGVtOiB7XG4gICAgY2FuZGlkYXRlczogWydbcm9sZT1cInRyZWVpdGVtXCJdW2RhdGEtZm9sZGVyLW5hbWVdJywgJ1tyb2xlPVwidHJlZWl0ZW1cIl0nXSxcbiAgICBleHBlY3Q6IFwiYWx3YXlzXCIsXG4gICAgd2l0aGluOiBcImZvbGRlclBhbmVcIixcbiAgICB0YWc6IFwiZm9sZGVyLWl0ZW1cIixcbiAgICBhbGw6IHRydWUsXG4gIH0sXG4gIGxpc3RQYW5lOiB7XG4gICAgY2FuZGlkYXRlczogWydbZGF0YS1hcHAtc2VjdGlvbj1cIk1lc3NhZ2VMaXN0XCJdJ10sXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICAgIHRhZzogXCJsaXN0LXBhbmVcIixcbiAgfSxcbiAgbWVzc2FnZUxpc3Q6IHtcbiAgICBjYW5kaWRhdGVzOiBbXCIjTWFpbExpc3RcIiwgJ1tyb2xlPVwibGlzdGJveFwiXVthcmlhLWxhYmVsPVwiTWVzc2FnZSBsaXN0XCJdJ10sXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICAgIHRhZzogXCJtZXNzYWdlLWxpc3RcIixcbiAgfSxcbiAgbWVzc2FnZVJvdzoge1xuICAgIGNhbmRpZGF0ZXM6IFsnW3JvbGU9XCJvcHRpb25cIl1bZGF0YS1jb252aWRdJywgJ1tyb2xlPVwib3B0aW9uXCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJtZXNzYWdlTGlzdFwiLFxuICAgIHRhZzogXCJtZXNzYWdlLXJvd1wiLFxuICAgIGFsbDogdHJ1ZSxcbiAgfSxcbiAgZ3JvdXBIZWFkZXI6IHtcbiAgICBjYW5kaWRhdGVzOiBbJ1tyb2xlPVwiYnV0dG9uXCJdW2lkXj1cImdyb3VwSGVhZGVyXCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJtZXNzYWdlTGlzdFwiLFxuICAgIHRhZzogXCJncm91cC1oZWFkZXJcIixcbiAgICBhbGw6IHRydWUsXG4gIH0sXG4gIGZvY3VzZWRUYWJzOiB7XG4gICAgY2FuZGlkYXRlczogWydbcm9sZT1cInRhYmxpc3RcIl0nXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgd2l0aGluOiBcImxpc3RQYW5lXCIsXG4gICAgdGFnOiBcImZvY3VzZWQtdGFic1wiLFxuICB9LFxuICByZWFkaW5nUGFuZToge1xuICAgIGNhbmRpZGF0ZXM6IFtcbiAgICAgICdbZGF0YS1hcHAtc2VjdGlvbj1cIk1haWxSZWFkQ29tcG9zZVwiXScsXG4gICAgICBcIiNSZWFkaW5nUGFuZUNvbnRhaW5lcklkXCIsXG4gICAgICAnW3JvbGU9XCJtYWluXCJdW2FyaWEtbGFiZWw9XCJSZWFkaW5nIFBhbmVcIl0nLFxuICAgIF0sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHRhZzogXCJyZWFkaW5nLXBhbmVcIixcbiAgfSxcblxuICBtZXNzYWdlQm9keToge1xuICAgIGNhbmRpZGF0ZXM6IFsnW3JvbGU9XCJkb2N1bWVudFwiXVthcmlhLWxhYmVsPVwiTWVzc2FnZSBib2R5XCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJyZWFkaW5nUGFuZVwiLFxuICAgIHRhZzogXCJtZXNzYWdlLWJvZHlcIixcbiAgICBhbGw6IHRydWUsXG4gIH0sXG5cbiAgaW5mb0Jhcjoge1xuICAgIC8vIE5vdGljZXMgYWJvdmUgYSBtZXNzYWdlIChcIllvdSByZXBsaWVkIG9uXHUyMDI2XCIpOyB0aGVpciB0ZXh0IHNpdHMgaW4gYW4gZWxlbWVudCB3aG9zZSBpZCBzdGFydHNcbiAgICAvLyB3aXRoIFwiaW5mb0JhclwiLiBPdXRsb29rIHBhaW50cyB0aGVtIHdpdGggaW5saW5lIGNvbG91cnMuXG4gICAgY2FuZGlkYXRlczogWydbcm9sZT1cImdyb3VwXCJdOmhhcyg+ICogPiBbaWRePVwiaW5mb0JhclwiXSknXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgd2l0aGluOiBcInJlYWRpbmdQYW5lXCIsXG4gICAgdGFnOiBcImluZm8tYmFyXCIsXG4gICAgYWxsOiB0cnVlLFxuICB9LFxuXG4gIC8vIC0tLSBDbHV0dGVyIChlYWNoIGhpZGVhYmxlIGVudHJ5IGdldHMgYSB0b2dnbGUgaW4gc2V0dGluZ3MudHMpIC0tLVxuICB1cHNlbGxCYW5uZXI6IHtcbiAgICBjYW5kaWRhdGVzOiBbJ1tkYXRhLWFwcC1zZWN0aW9uPVwiVXBzZWxsQmFubmVyU2VjdGlvblwiXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB0YWc6IFwidXBzZWxsXCIsXG4gICAgYWxsOiB0cnVlLFxuICB9LFxuICBoZWFkZXJFeHRyYXM6IHtcbiAgICBjYW5kaWRhdGVzOiBbXCIjb3dhVGVhbXNCdG5fY29udGFpbmVyXCIsIFwiI293YVRpbWVQYW5lbEJ0bl9jb250YWluZXJcIl0sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHRhZzogXCJoZWFkZXItZXh0cmFcIixcbiAgICB1bmlvbjogdHJ1ZSxcbiAgfSxcbiAgYXBwUmFpbEV4dHJhczoge1xuICAgIGNhbmRpZGF0ZXM6IFtcIlRvIERvXCIsIFwiTmV3c2xldHRlcnNcIiwgXCJPcmcgRXhwbG9yZXJcIiwgXCJPbmVEcml2ZVwiXS5tYXAoXG4gICAgICAoYXBwKSA9PiBgW3JvbGU9XCJub25lXCJdW2FyaWEtbGFiZWw9XCIke2FwcH1cIl1gLFxuICAgICksXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJhcHBSYWlsXCIsXG4gICAgdGFnOiBcInJhaWwtZXh0cmFcIixcbiAgICB1bmlvbjogdHJ1ZSxcbiAgfSxcbiAgcmliYm9uRXh0cmFzOiB7XG4gICAgY2FuZGlkYXRlczogW1xuICAgICAgXCJTaGFyZSB0byBUZWFtc1wiLFxuICAgICAgXCJTd2VlcFwiLFxuICAgICAgXCJBc3NpZ24gcG9saWN5XCIsXG4gICAgICBcIlZpdmEgSW5zaWdodHNcIixcbiAgICAgIFwiTW9yZSBhcHBzXCIsXG4gICAgICBcIlF1aWNrIHN0ZXBzXCIsXG4gICAgICBcIk1vcmUgcXVpY2sgc3RlcHNcIixcbiAgICBdLm1hcCgobGFiZWwpID0+IGBbYXJpYS1sYWJlbD1cIiR7bGFiZWx9XCJdYCksXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJyaWJib25cIixcbiAgICB0YWc6IFwicmliYm9uLWV4dHJhXCIsXG4gICAgdW5pb246IHRydWUsXG4gIH0sXG4gIHJvd0F2YXRhcjoge1xuICAgIGNhbmRpZGF0ZXM6IFsnW3JvbGU9XCJpbWdcIl1baWRePVwiYXZhdGFyLVwiXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB3aXRoaW46IFwibWVzc2FnZUxpc3RcIixcbiAgICB0YWc6IFwicm93LWF2YXRhclwiLFxuICAgIGFsbDogdHJ1ZSxcbiAgfSxcbiAgcm93SG92ZXJBY3Rpb25zOiB7XG4gICAgY2FuZGlkYXRlczogW1xuICAgICAgXCJDdXN0b21pemUgQWN0aW9uc1wiLFxuICAgICAgXCJGbGFnIHRoaXMgbWVzc2FnZVwiLFxuICAgICAgXCJLZWVwIHRoaXMgbWVzc2FnZSBhdCB0aGUgdG9wIG9mIHlvdXIgZm9sZGVyXCIsXG4gICAgXS5tYXAoKGxhYmVsKSA9PiBgYnV0dG9uW2FyaWEtbGFiZWw9XCIke2xhYmVsfVwiXWApLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB3aXRoaW46IFwibWVzc2FnZUxpc3RcIixcbiAgICB0YWc6IFwicm93LWFjdGlvblwiLFxuICAgIHVuaW9uOiB0cnVlLFxuICB9LFxuICBhcHBCYW5uZXI6IHtcbiAgICAvLyBGbHVlbnQncyBzdGF0aWMgXCJmdWktTWVzc2FnZUJhclwiIGNsYXNzIGlzIHBhcnQgb2YgaXRzIHB1YmxpYyBzdHlsaW5nIEFQSSwgbm90IGdlbmVyYXRlZDtcbiAgICAvLyBub3RoaW5nIGVsc2UgbWFya3MgdGhlc2UuIFNjb3BlZCB0byBiYXJzIGFjcm9zcyB0aGUgdG9wIG9mIHRoZSBhcHAsIG5vdCBpbnNpZGUgbWVzc2FnZXMuXG4gICAgY2FuZGlkYXRlczogWycjbWFpbkFwcCA+IGRpdiA+IFtyb2xlPVwiZ3JvdXBcIl0uZnVpLU1lc3NhZ2VCYXInXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgdGFnOiBcImFwcC1iYW5uZXJcIixcbiAgICBhbGw6IHRydWUsXG4gIH0sXG4gIGNvcGlsb3Q6IHtcbiAgICAvLyBVTlZFUklGSUVEOiBubyBDb3BpbG90IHN1cmZhY2VzIGluIHRoZSBPY3QgMjAyNiBjYXB0dXJlLlxuICAgIGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWwqPVwiQ29waWxvdFwiIGldJywgJ1tyb2xlPVwicmVnaW9uXCJdW2FyaWEtbGFiZWwqPVwiQ29waWxvdFwiIGldJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHRhZzogXCJjb3BpbG90XCIsXG4gICAgdW5pb246IHRydWUsXG4gIH0sXG4gIGV4dGVybmFsVGFnOiB7XG4gICAgLy8gVU5WRVJJRklFRDogbm8gRXh0ZXJuYWwgdGFncyBpbiB0aGUgT2N0IDIwMjYgY2FwdHVyZS5cbiAgICBjYW5kaWRhdGVzOiBbJ1t0aXRsZT1cIkV4dGVybmFsIHNlbmRlclwiXScsICdbYXJpYS1sYWJlbD1cIkV4dGVybmFsXCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHRhZzogXCJleHRlcm5hbC10YWdcIixcbiAgICB1bmlvbjogdHJ1ZSxcbiAgfSxcblxuICAvLyAtLS0gQ29udHJvbHMgdGhlIGtleW1hcCBkcml2ZXMgLS0tXG4gIGNvbXBvc2VCdXR0b246IHsgY2FuZGlkYXRlczogWydidXR0b25bYXJpYS1sYWJlbD1cIk5ldyBtYWlsXCJdJ10sIGV4cGVjdDogXCJhbHdheXNcIiB9LFxuICBhcmNoaXZlQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJBcmNoaXZlXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICBkZWxldGVCdXR0b246IHsgY2FuZGlkYXRlczogWydidXR0b25bYXJpYS1sYWJlbD1cIkRlbGV0ZVwiXSddLCBleHBlY3Q6IFwic29tZXRpbWVzXCIgfSxcbiAgLy8gVU5WRVJJRklFRDogdGhpcyB0ZW5hbnQgc2hvd3MgYSBcIlBoaXNoIEFsZXJ0XCIgYWRkLWluIGluc3RlYWQgb2YgT3V0bG9vaydzIFJlcG9ydCBtZW51LlxuICBqdW5rQnV0dG9uOiB7XG4gICAgY2FuZGlkYXRlczogWydidXR0b25bYXJpYS1sYWJlbD1cIlJlcG9ydFwiXScsICdidXR0b25bYXJpYS1sYWJlbD1cIkp1bmtcIl0nXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gIH0sXG4gIG1vdmVCdXR0b246IHsgY2FuZGlkYXRlczogWydidXR0b25bYXJpYS1sYWJlbD1cIk1vdmUgdG9cIl0nXSwgZXhwZWN0OiBcInNvbWV0aW1lc1wiIH0sXG4gIHNub296ZUJ1dHRvbjogeyBjYW5kaWRhdGVzOiBbJ2J1dHRvblthcmlhLWxhYmVsPVwiU25vb3plXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICBmbGFnQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJGbGFnIC8gVW5mbGFnXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICByZWFkVW5yZWFkQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJSZWFkIC8gVW5yZWFkXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICB1bmRvQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJVbmRvXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICBzeW5jQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFtcIiNvd2FTeW5jQnV0dG9uX2NvbnRhaW5lclwiXSwgZXhwZWN0OiBcInNvbWV0aW1lc1wiIH0sXG4gIHNldHRpbmdzQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFtcIiNvd2FTZXR0aW5nc0J0bl9jb250YWluZXJcIl0sIGV4cGVjdDogXCJhbHdheXNcIiB9LFxuICAvLyBBcHAgcmFpbCBidXR0b25zLiBGYWxsYmFja3MgYXJlIHRoZSBNaWNyb3NvZnQgMzY1IGFwcCBpZHMgT3V0bG9vayBwdXRzIGluIGRhdGEtdGlkLCB0aGUgc2FtZVxuICAvLyBhY3Jvc3MgdGhlIE9jdCBhbmQgTm92IDIwMjYgY2FwdHVyZXMuIFRoZSByYWlsIHJlbmRlcnMgYWZ0ZXIgdGhlIG1lc3NhZ2UgbGlzdC5cbiAgbWFpbEFwcDoge1xuICAgIGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJNYWlsXCJdJywgJ1tkYXRhLXRpZD1cImRkZWE3NzRjLTM4MmItNDdkNy1hYWI1LWFkYzIxMzlhODAyYlwiXSddLFxuICAgIHdpdGhpbjogXCJhcHBSYWlsXCIsXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICB9LFxuICBjYWxlbmRhckFwcDoge1xuICAgIGNhbmRpZGF0ZXM6IFtcbiAgICAgICdidXR0b25bYXJpYS1sYWJlbD1cIkNhbGVuZGFyXCJdJyxcbiAgICAgICdbZGF0YS10aWQ9XCI4Y2JlYjg2Zi04M2UxLTQzYjUtYWFiYS1jZDM1MTQzMjJmMGJcIl0nLFxuICAgIF0sXG4gICAgd2l0aGluOiBcImFwcFJhaWxcIixcbiAgICBleHBlY3Q6IFwiYWx3YXlzXCIsXG4gIH0sXG4gIHBlb3BsZUFwcDoge1xuICAgIGNhbmRpZGF0ZXM6IFtcbiAgICAgICdidXR0b25bYXJpYS1sYWJlbD1cIlBlb3BsZVwiXScsXG4gICAgICAnW2RhdGEtdGlkPVwiMzU1ZmJkNzktM2JhMi00NTU0LThmMmQtMDMwMGZkZTkxZjMwXCJdJyxcbiAgICBdLFxuICAgIHdpdGhpbjogXCJhcHBSYWlsXCIsXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICB9LFxufSBhcyBjb25zdCBzYXRpc2ZpZXMgUmVjb3JkPHN0cmluZywgU2VsZWN0b3JFbnRyeT47XG5cbmNvbnN0IGVudHJ5ID0gKG5hbWU6IFNlbGVjdG9yTmFtZSk6IFNlbGVjdG9yRW50cnkgPT4gU0VMRUNUT1JTW25hbWVdO1xuXG5mdW5jdGlvbiBzY29wZShuYW1lOiBTZWxlY3Rvck5hbWUsIHJvb3Q6IFBhcmVudE5vZGUpOiBQYXJlbnROb2RlIHwgbnVsbCB7XG4gIGNvbnN0IHBhcmVudCA9IGVudHJ5KG5hbWUpLndpdGhpbjtcbiAgcmV0dXJuIHBhcmVudCA/IHJlc29sdmUocGFyZW50IGFzIFNlbGVjdG9yTmFtZSwgcm9vdCkgOiByb290O1xufVxuXG4vKiogVGhlIGZpcnN0IGVsZW1lbnQgbWF0Y2hpbmcgYG5hbWVgLCB0cnlpbmcgY2FuZGlkYXRlcyBpbiBwcmlvcml0eSBvcmRlci4gKi9cbmV4cG9ydCBmdW5jdGlvbiByZXNvbHZlKG5hbWU6IFNlbGVjdG9yTmFtZSwgcm9vdDogUGFyZW50Tm9kZSA9IGRvY3VtZW50KTogRWxlbWVudCB8IG51bGwge1xuICBjb25zdCB3aXRoaW4gPSBzY29wZShuYW1lLCByb290KTtcbiAgaWYgKCF3aXRoaW4pIHJldHVybiBudWxsO1xuICBmb3IgKGNvbnN0IGNhbmRpZGF0ZSBvZiBlbnRyeShuYW1lKS5jYW5kaWRhdGVzKSB7XG4gICAgY29uc3QgZWwgPSB3aXRoaW4ucXVlcnlTZWxlY3RvcihjYW5kaWRhdGUpO1xuICAgIGlmIChlbCkgcmV0dXJuIGVsO1xuICB9XG4gIHJldHVybiBudWxsO1xufVxuXG4vKiogRXZlcnkgbWF0Y2g6IGZvciB1bmlvbiBlbnRyaWVzLCBvZiBhbGwgY2FuZGlkYXRlczsgb3RoZXJ3aXNlIG9mIHRoZSBmaXJzdCBjYW5kaWRhdGUgdGhhdCBtYXRjaGVzLiAqL1xuZXhwb3J0IGZ1bmN0aW9uIHJlc29sdmVBbGwobmFtZTogU2VsZWN0b3JOYW1lLCByb290OiBQYXJlbnROb2RlID0gZG9jdW1lbnQpOiBFbGVtZW50W10ge1xuICBjb25zdCB3aXRoaW4gPSBzY29wZShuYW1lLCByb290KTtcbiAgaWYgKCF3aXRoaW4pIHJldHVybiBbXTtcbiAgY29uc3QgeyBjYW5kaWRhdGVzLCB1bmlvbiB9ID0gZW50cnkobmFtZSk7XG4gIGlmICh1bmlvbikgcmV0dXJuIFsuLi53aXRoaW4ucXVlcnlTZWxlY3RvckFsbChjYW5kaWRhdGVzLmpvaW4oXCIsIFwiKSldO1xuICBmb3IgKGNvbnN0IGNhbmRpZGF0ZSBvZiBjYW5kaWRhdGVzKSB7XG4gICAgY29uc3QgZWxzID0gd2l0aGluLnF1ZXJ5U2VsZWN0b3JBbGwoY2FuZGlkYXRlKTtcbiAgICBpZiAoZWxzLmxlbmd0aCA+IDApIHJldHVybiBbLi4uZWxzXTtcbiAgfVxuICByZXR1cm4gW107XG59XG5cbi8qKiBBZGQgYGRhdGEtb3Zlcmxvb2tgIHRva2VucyB0byBldmVyeSB0YWdnZWQgZW50cnkgc28gc3R5bGVzaGVldHMgbmV2ZXIgdG91Y2ggT3V0bG9vaydzIERPTSBzaGFwZS4gKi9cbmV4cG9ydCBmdW5jdGlvbiB0YWdFbGVtZW50cyhyb290OiBQYXJlbnROb2RlID0gZG9jdW1lbnQpOiB2b2lkIHtcbiAgZm9yIChjb25zdCBuYW1lIG9mIE9iamVjdC5rZXlzKFNFTEVDVE9SUykgYXMgU2VsZWN0b3JOYW1lW10pIHtcbiAgICBjb25zdCB7IHRhZywgYWxsLCB1bmlvbiB9ID0gZW50cnkobmFtZSk7XG4gICAgaWYgKCF0YWcpIGNvbnRpbnVlO1xuICAgIGNvbnN0IGVscyA9XG4gICAgICBhbGwgfHwgdW5pb24gPyByZXNvbHZlQWxsKG5hbWUsIHJvb3QpIDogW3Jlc29sdmUobmFtZSwgcm9vdCldLmZpbHRlcigoZSkgPT4gZSAhPT0gbnVsbCk7XG4gICAgZm9yIChjb25zdCBlbCBvZiBlbHMpIHtcbiAgICAgIGNvbnN0IHRva2VucyA9IG5ldyBTZXQoKGVsLmdldEF0dHJpYnV0ZShcImRhdGEtb3Zlcmxvb2tcIikgPz8gXCJcIikuc3BsaXQoXCIgXCIpLmZpbHRlcihCb29sZWFuKSk7XG4gICAgICBpZiAodG9rZW5zLmhhcyh0YWcpKSBjb250aW51ZTtcbiAgICAgIHRva2Vucy5hZGQodGFnKTtcbiAgICAgIGVsLnNldEF0dHJpYnV0ZShcImRhdGEtb3Zlcmxvb2tcIiwgWy4uLnRva2Vuc10uam9pbihcIiBcIikpO1xuICAgIH1cbiAgfVxufVxuXG5leHBvcnQgaW50ZXJmYWNlIEhlYWx0aFJlc3VsdCB7XG4gIG5hbWU6IFNlbGVjdG9yTmFtZTtcbiAgZXhwZWN0OiBTZWxlY3RvckVudHJ5W1wiZXhwZWN0XCJdO1xuICAvKiogSW5kZXggb2YgdGhlIGNhbmRpZGF0ZSB0aGF0IG1hdGNoZWQsIG9yIC0xLiBBIG5vbi16ZXJvIGluZGV4IG1lYW5zIGEgZmFsbGJhY2sgaXMgaW4gdXNlLiAqL1xuICBjYW5kaWRhdGU6IG51bWJlcjtcbiAgY291bnQ6IG51bWJlcjtcbn1cblxuZXhwb3J0IGZ1bmN0aW9uIGhlYWx0aENoZWNrKHJvb3Q6IFBhcmVudE5vZGUgPSBkb2N1bWVudCk6IEhlYWx0aFJlc3VsdFtdIHtcbiAgcmV0dXJuIChPYmplY3Qua2V5cyhTRUxFQ1RPUlMpIGFzIFNlbGVjdG9yTmFtZVtdKS5tYXAoKG5hbWUpID0+IHtcbiAgICBjb25zdCB7IGNhbmRpZGF0ZXMsIGV4cGVjdCwgdW5pb24gfSA9IGVudHJ5KG5hbWUpO1xuICAgIGNvbnN0IHdpdGhpbiA9IHNjb3BlKG5hbWUsIHJvb3QpO1xuICAgIGxldCBjYW5kaWRhdGUgPSAtMTtcbiAgICBsZXQgY291bnQgPSAwO1xuICAgIGlmICh3aXRoaW4pIHtcbiAgICAgIGNhbmRpZGF0ZSA9IGNhbmRpZGF0ZXMuZmluZEluZGV4KChjKSA9PiB3aXRoaW4ucXVlcnlTZWxlY3RvcihjKSAhPT0gbnVsbCk7XG4gICAgICBpZiAodW5pb24gJiYgY2FuZGlkYXRlID49IDApIGNvdW50ID0gcmVzb2x2ZUFsbChuYW1lLCByb290KS5sZW5ndGg7XG4gICAgICBlbHNlIGlmIChjYW5kaWRhdGUgPj0gMCkgY291bnQgPSB3aXRoaW4ucXVlcnlTZWxlY3RvckFsbChjYW5kaWRhdGVzW2NhbmRpZGF0ZV0gPz8gXCJcIikubGVuZ3RoO1xuICAgIH1cbiAgICAvLyBBIHVuaW9uIGVudHJ5IGxpc3RzIGRpc3RpbmN0IHRoaW5ncywgc28gbWF0Y2hpbmcgYSBsYXRlciBvbmUgaXMgbm90IGEgZmFsbGJhY2suXG4gICAgaWYgKHVuaW9uICYmIGNhbmRpZGF0ZSA+IDApIGNhbmRpZGF0ZSA9IDA7XG4gICAgcmV0dXJuIHsgbmFtZSwgZXhwZWN0LCBjYW5kaWRhdGUsIGNvdW50IH07XG4gIH0pO1xufVxuXG5leHBvcnQgaW50ZXJmYWNlIEZvbGRlck5vZGUge1xuICBuYW1lOiBzdHJpbmc7XG4gIC8qKiBBbmNlc3RvciBuYW1lcyBqb2luZWQgd2l0aCBcIi9cIiwgZS5nLiBcIkluYm94L1JlY2VpcHRzXCIuICovXG4gIHBhdGg6IHN0cmluZztcbiAgZWxlbWVudDogRWxlbWVudDtcbn1cblxuLyoqIFwiSW5ib3ggLSA5IGl0ZW1zICgxIHVucmVhZClcIiAtPiBcIkluYm94XCIuIFRoZSB0aXRsZSBhdHRyaWJ1dGUgY2FycmllcyBjb3VudHMuICovXG5mdW5jdGlvbiBmb2xkZXJOYW1lKGVsZW1lbnQ6IEVsZW1lbnQpOiBzdHJpbmcge1xuICBjb25zdCB0aXRsZSA9IChlbGVtZW50LmdldEF0dHJpYnV0ZShcInRpdGxlXCIpID8/IGVsZW1lbnQudGV4dENvbnRlbnQgPz8gXCJcIikudHJpbSgpO1xuICByZXR1cm4gdGl0bGUucmVwbGFjZSgvXFxzKy1cXHMrXFxkK1xccytpdGVtcz8oXFxzK1xcKFxcZCtcXHMrdW5yZWFkXFwpKT8kLywgXCJcIik7XG59XG5cbi8qKlxuICogRm9sZGVycyBjdXJyZW50bHkgcmVuZGVyZWQgaW4gdGhlIGZvbGRlciBwYW5lLCB3aXRoIHBhdGhzIGJ1aWx0IGZyb20gYXJpYS1sZXZlbC4gRm9sZGVyc1xuICogdW5kZXIgRmF2b3JpdGVzIGFyZSBza2lwcGVkICh0aGV5IGR1cGxpY2F0ZSB0aGUgbWFpbGJveCB0cmVlKSwgYXMgYXJlIHRoZSB0b3AtbGV2ZWwgbWFpbGJveFxuICogcm9vdHMsIHdob3NlIG5hbWVzIGFyZSBhY2NvdW50IGFkZHJlc3Nlcy5cbiAqIFVOVkVSSUZJRUQ6IGNvbGxhcHNlZCBzdWJmb2xkZXJzIG1heSBiZSBhYnNlbnQgZnJvbSB0aGUgRE9NLlxuICovXG5leHBvcnQgZnVuY3Rpb24gcmVhZEZvbGRlclRyZWUocm9vdDogUGFyZW50Tm9kZSA9IGRvY3VtZW50KTogRm9sZGVyTm9kZVtdIHtcbiAgY29uc3Qgc3RhY2s6IHN0cmluZ1tdID0gW107XG4gIGNvbnN0IGZvbGRlcnM6IEZvbGRlck5vZGVbXSA9IFtdO1xuICBsZXQgaW5GYXZvcml0ZXMgPSBmYWxzZTtcbiAgZm9yIChjb25zdCBlbGVtZW50IG9mIHJlc29sdmVBbGwoXCJmb2xkZXJJdGVtXCIsIHJvb3QpKSB7XG4gICAgY29uc3QgbGV2ZWwgPSBOdW1iZXIoZWxlbWVudC5nZXRBdHRyaWJ1dGUoXCJhcmlhLWxldmVsXCIpID8/IFwiMVwiKTtcbiAgICBjb25zdCBuYW1lID0gZm9sZGVyTmFtZShlbGVtZW50KTtcbiAgICBzdGFjay5sZW5ndGggPSBNYXRoLm1heChsZXZlbCAtIDEsIDApO1xuICAgIHN0YWNrLnB1c2gobmFtZSk7XG4gICAgaWYgKGxldmVsID09PSAxKSB7XG4gICAgICBpbkZhdm9yaXRlcyA9IGVsZW1lbnQuZ2V0QXR0cmlidXRlKFwiZGF0YS1mb2xkZXItbmFtZVwiKSA9PT0gXCJmYXZvcml0ZXNcIjtcbiAgICAgIGNvbnRpbnVlO1xuICAgIH1cbiAgICBpZiAoaW5GYXZvcml0ZXMpIGNvbnRpbnVlO1xuICAgIGZvbGRlcnMucHVzaCh7IG5hbWUsIHBhdGg6IHN0YWNrLnNsaWNlKDEpLmZpbHRlcihCb29sZWFuKS5qb2luKFwiL1wiKSwgZWxlbWVudCB9KTtcbiAgfVxuICByZXR1cm4gZm9sZGVycztcbn1cblxuLyoqIE1lc3NhZ2Ugcm93cyBjdXJyZW50bHkgc2VsZWN0ZWQgKG9yIG9wZW4pIGluIHRoZSBsaXN0LiAqL1xuZXhwb3J0IGZ1bmN0aW9uIHNlbGVjdGVkUm93cyhyb290OiBQYXJlbnROb2RlID0gZG9jdW1lbnQpOiBFbGVtZW50W10ge1xuICByZXR1cm4gcmVzb2x2ZUFsbChcIm1lc3NhZ2VSb3dcIiwgcm9vdCkuZmlsdGVyKFxuICAgIChyb3cpID0+IHJvdy5nZXRBdHRyaWJ1dGUoXCJhcmlhLXNlbGVjdGVkXCIpID09PSBcInRydWVcIixcbiAgKTtcbn1cblxuLyoqIE91dGxvb2sgc3RhcnRzIGFuIHVucmVhZCByb3cncyBhY2Nlc3NpYmxlIG5hbWUgd2l0aCBcIlVucmVhZFwiIChFbmdsaXNoIFVJKS4gKi9cbmV4cG9ydCBmdW5jdGlvbiBpc1VucmVhZFJvdyhyb3c6IEVsZW1lbnQpOiBib29sZWFuIHtcbiAgcmV0dXJuIC9eVW5yZWFkXFxiLy50ZXN0KHJvdy5nZXRBdHRyaWJ1dGUoXCJhcmlhLWxhYmVsXCIpID8/IFwiXCIpO1xufVxuXG4vKiogT3V0bG9vaydzIGN1cnJlbnQgbGlnaHQvZGFyayBtb2RlLCBmcm9tIHRoZSBwYWxldHRlIHZhcmlhYmxlcyBpdCB3cml0ZXMgb24gPGh0bWw+LiAqL1xuZXhwb3J0IGZ1bmN0aW9uIG91dGxvb2tDb2xvclNjaGVtZShodG1sOiBFbGVtZW50KTogXCJsaWdodFwiIHwgXCJkYXJrXCIge1xuICBjb25zdCBzdHlsZSA9IGh0bWwuZ2V0QXR0cmlidXRlKFwic3R5bGVcIikgPz8gXCJcIjtcbiAgcmV0dXJuIC8tLW5hdGl2ZS1zY3JvbGxiYXItY29sb3Itc2NoZW1lOlxccypkYXJrLy50ZXN0KHN0eWxlKSA/IFwiZGFya1wiIDogXCJsaWdodFwiO1xufVxuXG4vKiogRW50cmllcyBleHBlY3RlZCBvbiBldmVyeSBtYWlsIHZpZXcgdGhhdCBkaWQgbm90IHJlc29sdmUuIE5vbi1lbXB0eSBtZWFucyBzb21ldGhpbmcgYnJva2UuICovXG5leHBvcnQgZnVuY3Rpb24gYnJva2VuRW50cmllcyhyZXN1bHRzOiBIZWFsdGhSZXN1bHRbXSk6IFNlbGVjdG9yTmFtZVtdIHtcbiAgcmV0dXJuIHJlc3VsdHMuZmlsdGVyKChyKSA9PiByLmV4cGVjdCA9PT0gXCJhbHdheXNcIiAmJiByLmNhbmRpZGF0ZSA8IDApLm1hcCgocikgPT4gci5uYW1lKTtcbn1cbiIsICIvLyBEZXZlbG9wZXIgY2FwdHVyZTogc2VyaWFsaXNlcyB0aGUgbGl2ZSBPdXRsb29rIHBhZ2UgaW50byBvbmUgSFRNTCBmaWxlLCBzbyBhIGZpeHR1cmUgY2FuIGJlXG4vLyBtYWRlIChkb2NzL2NhcHR1cmluZy1wYWdlcy5tZCkgd2l0aG91dCBEZXZUb29scy4gU2F2ZWQgbG9jYWxseSBieSB0aGUgY2FsbGVyOyBuZXZlciBzZW50XG4vLyBhbnl3aGVyZS4gVGhlIGZpbGUgaG9sZHMgcmVhbCBtYWlsLCBzbyBpdCBnb2VzIHRocm91Z2ggc2NyaXB0cy9zY3J1Yi1maXh0dXJlLnRzIGJlZm9yZSBpdFxuLy8gZ29lcyBhbnl3aGVyZSBuZWFyIHRoZSByZXBvLlxuXG5pbXBvcnQgeyBicm9rZW5FbnRyaWVzLCBoZWFsdGhDaGVjayB9IGZyb20gXCIuL3NlbGVjdG9ycy50c1wiO1xuXG4vKiogTWVzc2FnZSB0aGUgcG9wdXAgc2VuZHMgdG8gdGhlIGNvbnRlbnQgc2NyaXB0IG9mIHRoZSBhY3RpdmUgdGFiLiAqL1xuZXhwb3J0IGNvbnN0IENBUFRVUkVfTUVTU0FHRSA9IFwib2JiLWNhcHR1cmVcIjtcblxuY29uc3QgT0JCX0FUVFJJQlVURSA9IC9eZGF0YS1vdmVybG9vay87XG4vKiogVGhlIHNjcnViYmVyIChzY3JpcHRzL3NjcnViLWZpeHR1cmUudHMpIGxvb2tzIGZvciB0aGlzIG1hcmtlciB0byBzcGxpdCBvdXQgcnVudGltZSBzdHlsZXMuICovXG5leHBvcnQgY29uc3QgUlVOVElNRV9TVFlMRV9NQVJLRVIgPSBcImRhdGEtb3Zlcmxvb2stcnVudGltZVwiO1xuXG5leHBvcnQgaW50ZXJmYWNlIENhcHR1cmVNZXRhIHtcbiAgY2FwdHVyZWRBdDogc3RyaW5nO1xuICBvYmJWZXJzaW9uOiBzdHJpbmc7XG4gIC8qKiBGaXJzdCB0d28gcGF0aCBzZWdtZW50cyBvbmx5OiBkZWVwZXIgb25lcyBjYXJyeSBtZXNzYWdlIGlkcy4gKi9cbiAgcm91dGU6IHN0cmluZztcbiAgdmlld3BvcnQ6IHN0cmluZztcbiAgc2V0dGluZ3M6IFJlY29yZDxzdHJpbmcsIHN0cmluZyB8IHVuZGVmaW5lZD47XG4gIGhlYWx0aDogeyBicm9rZW46IHN0cmluZ1tdOyBlbnRyaWVzOiB7IG5hbWU6IHN0cmluZzsgY2FuZGlkYXRlOiBudW1iZXI7IGNvdW50OiBudW1iZXIgfVtdIH07XG59XG5cbmV4cG9ydCBpbnRlcmZhY2UgQ2FwdHVyZSB7XG4gIGh0bWw6IHN0cmluZztcbiAgbmFtZTogc3RyaW5nO1xuICBtZXRhOiBDYXB0dXJlTWV0YTtcbn1cblxuLyoqIFN0eWxlIHJ1bGVzIE91dGxvb2sgaW5qZWN0ZWQgYXQgcnVudGltZSAoQ1NTLWluLUpTKSwgbWludXMgdGhlIG9uZXMgT0JCIGl0c2VsZiBhZGRlZC4gKi9cbmZ1bmN0aW9uIHJ1bnRpbWVDc3MoZG9jOiBEb2N1bWVudCk6IHN0cmluZyB7XG4gIGNvbnN0IHNoZWV0cyA9IFsuLi5kb2Muc3R5bGVTaGVldHMsIC4uLihkb2MuYWRvcHRlZFN0eWxlU2hlZXRzID8/IFtdKV0uZmlsdGVyKFxuICAgIChzaGVldCkgPT4gIShcImhyZWZcIiBpbiBzaGVldCAmJiBzaGVldC5ocmVmKSxcbiAgKTtcbiAgcmV0dXJuIHNoZWV0c1xuICAgIC5mbGF0TWFwKChzaGVldCkgPT4ge1xuICAgICAgdHJ5IHtcbiAgICAgICAgcmV0dXJuIFsuLi5zaGVldC5jc3NSdWxlc10ubWFwKChydWxlKSA9PiBydWxlLmNzc1RleHQpO1xuICAgICAgfSBjYXRjaCB7XG4gICAgICAgIHJldHVybiBbXTsgLy8gY3Jvc3Mtb3JpZ2luIHNoZWV0XG4gICAgICB9XG4gICAgfSlcbiAgICAuZmlsdGVyKChydWxlKSA9PiAhcnVsZS5pbmNsdWRlcyhcImRhdGEtb3Zlcmxvb2tcIikpXG4gICAgLmpvaW4oXCJcXG5cIik7XG59XG5cbmV4cG9ydCBmdW5jdGlvbiBjYXB0dXJlRmlsZW5hbWUobm93ID0gbmV3IERhdGUoKSk6IHN0cmluZyB7XG4gIGNvbnN0IHN0YW1wID0gbm93LnRvSVNPU3RyaW5nKCkuc2xpY2UoMCwgMTYpLnJlcGxhY2UoL1xcRC9nLCBcIlwiKTtcbiAgcmV0dXJuIGBvYmItY2FwdHVyZS0ke3N0YW1wfS5odG1sYDtcbn1cblxuZXhwb3J0IGZ1bmN0aW9uIGNhcHR1cmVQYWdlKFxuICBkb2M6IERvY3VtZW50ID0gZG9jdW1lbnQsXG4gIGxvYzogUGljazxMb2NhdGlvbiwgXCJwYXRobmFtZVwiPiA9IGxvY2F0aW9uLFxuICB2ZXJzaW9uID0gXCJkZXZcIixcbiAgbm93ID0gbmV3IERhdGUoKSxcbik6IENhcHR1cmUge1xuICBjb25zdCBsaXZlID0gZG9jLmRvY3VtZW50RWxlbWVudC5kYXRhc2V0O1xuICBjb25zdCByZXN1bHRzID0gaGVhbHRoQ2hlY2soZG9jKTtcbiAgY29uc3QgbWV0YTogQ2FwdHVyZU1ldGEgPSB7XG4gICAgY2FwdHVyZWRBdDogbm93LnRvSVNPU3RyaW5nKCksXG4gICAgb2JiVmVyc2lvbjogdmVyc2lvbixcbiAgICByb3V0ZTogbG9jLnBhdGhuYW1lLnNwbGl0KFwiL1wiKS5zbGljZSgwLCAzKS5qb2luKFwiL1wiKSxcbiAgICB2aWV3cG9ydDogYCR7ZG9jLmRlZmF1bHRWaWV3Py5pbm5lcldpZHRoID8/IDB9eCR7ZG9jLmRlZmF1bHRWaWV3Py5pbm5lckhlaWdodCA/PyAwfWAsXG4gICAgc2V0dGluZ3M6IHtcbiAgICAgIGZlYXR1cmVzOiBsaXZlLm92ZXJsb29rRmVhdHVyZXMsXG4gICAgICBwYWxldHRlOiBsaXZlLm92ZXJsb29rUGFsZXR0ZSxcbiAgICAgIG1vZGU6IGxpdmUub3Zlcmxvb2tNb2RlLFxuICAgICAgbGF5b3V0OiBsaXZlLm92ZXJsb29rTGF5b3V0LFxuICAgICAgb3V0bG9va1NjaGVtZTogbGl2ZS5vdmVybG9va091dGxvb2tTY2hlbWUsXG4gICAgICByZWFkaW5nUGFuZTogbGl2ZS5vdmVybG9va1JlYWRpbmdQYW5lLFxuICAgIH0sXG4gICAgaGVhbHRoOiB7XG4gICAgICBicm9rZW46IGJyb2tlbkVudHJpZXMocmVzdWx0cyksXG4gICAgICBlbnRyaWVzOiByZXN1bHRzLm1hcCgoeyBuYW1lLCBjYW5kaWRhdGUsIGNvdW50IH0pID0+ICh7IG5hbWUsIGNhbmRpZGF0ZSwgY291bnQgfSkpLFxuICAgIH0sXG4gIH07XG5cbiAgLy8gV29yayBvbiBhIGNvcHksIHdpdGggT0JCJ3Mgb3duIG1hcmtzIHJlbW92ZWQsIHNvIHRoZSBjYXB0dXJlIGlzIHN0b2NrIE91dGxvb2suXG4gIGNvbnN0IHJvb3QgPSBkb2MuZG9jdW1lbnRFbGVtZW50LmNsb25lTm9kZSh0cnVlKSBhcyBIVE1MRWxlbWVudDtcbiAgZm9yIChjb25zdCBlbCBvZiByb290LnF1ZXJ5U2VsZWN0b3JBbGwoXCJvdmVybG9vay1yb290XCIpKSBlbC5yZW1vdmUoKTtcbiAgZm9yIChjb25zdCBlbCBvZiBbcm9vdCwgLi4ucm9vdC5xdWVyeVNlbGVjdG9yQWxsKFwiKlwiKV0pIHtcbiAgICBmb3IgKGNvbnN0IHsgbmFtZSB9IG9mIFsuLi5lbC5hdHRyaWJ1dGVzXSkge1xuICAgICAgaWYgKE9CQl9BVFRSSUJVVEUudGVzdChuYW1lKSkgZWwucmVtb3ZlQXR0cmlidXRlKG5hbWUpO1xuICAgIH1cbiAgfVxuXG4gIGNvbnN0IGhlYWQgPSByb290LnF1ZXJ5U2VsZWN0b3IoXCJoZWFkXCIpO1xuICBjb25zdCBzdHlsZSA9IGRvYy5jcmVhdGVFbGVtZW50KFwic3R5bGVcIik7XG4gIHN0eWxlLnNldEF0dHJpYnV0ZShSVU5USU1FX1NUWUxFX01BUktFUiwgXCJcIik7XG4gIHN0eWxlLnRleHRDb250ZW50ID0gcnVudGltZUNzcyhkb2MpO1xuICBoZWFkPy5hcHBlbmQoc3R5bGUpO1xuICBjb25zdCBzY3JpcHQgPSBkb2MuY3JlYXRlRWxlbWVudChcInNjcmlwdFwiKTtcbiAgc2NyaXB0LnR5cGUgPSBcImFwcGxpY2F0aW9uL2pzb25cIjtcbiAgc2NyaXB0LmlkID0gXCJvYmItY2FwdHVyZS1tZXRhXCI7XG4gIHNjcmlwdC50ZXh0Q29udGVudCA9IEpTT04uc3RyaW5naWZ5KG1ldGEsIG51bGwsIDIpLnJlcGxhY2VBbGwoXCI8XCIsIFwiXFxcXHUwMDNjXCIpO1xuICBoZWFkPy5hcHBlbmQoc2NyaXB0KTtcblxuICByZXR1cm4geyBodG1sOiBgPCFkb2N0eXBlIGh0bWw+XFxuJHtyb290Lm91dGVySFRNTH1gLCBuYW1lOiBjYXB0dXJlRmlsZW5hbWUobm93KSwgbWV0YSB9O1xufVxuIiwgIi8vIEFjdGlvbiB0YWJsZTogd2hhdCBlYWNoIGFjdGlvbiBtZWFucyBhbmQgaG93IGl0IGV4ZWN1dGVzLiBQdXJlIGRhdGEuXG4vL1xuLy8gU3RyYXRlZ2llcyAoUExBTi5tZCwgS2V5bWFwIGxheWVyKTpcbi8vICAgcGFzc3Rocm91Z2ggXHUyMDEzIE91dGxvb2sncyBHbWFpbCBzaG9ydGN1dCBtb2RlIGFscmVhZHkgZG9lcyBpdDsgbGV0IHRoZSBrZXkgdGhyb3VnaC5cbi8vICAgZHJpdmUgICAgICAgXHUyMDEzIGNsaWNrIGEgY29udHJvbCBmb3VuZCB0aHJvdWdoIHRoZSBzZWxlY3RvciByZWdpc3RyeS4gUmVhY3QgaG9ub3VycyBjbGljaygpLlxuLy8gICAgICAgICAgICAgICAgIChEaXNwYXRjaGluZyBPdXRsb29rJ3Mgb3duIHNob3J0Y3V0IGFzIGEgc3ludGhldGljIGtleSBldmVudCBkaWQgbm90aGluZyBvblxuLy8gICAgICAgICAgICAgICAgIHRoZSBsaXZlIHNpdGUgaW4gTm92IDIwMjYsIHNvIHRoZXJlIGlzIG5vIGtleXN0cm9rZSBzdHJhdGVneS4pXG4vLyAgIHRvZ2dsZSAgICAgIFx1MjAxMyBjbGljayBhIHRvZ2dsZSBidXR0b24gKFJlYWQgLyBVbnJlYWQpIG9ubHkgaWYgdGhlIHNlbGVjdGlvbiBpc24ndCBhbHJlYWR5IGluXG4vLyAgICAgICAgICAgICAgICAgdGhlIHdhbnRlZCBzdGF0ZS5cbi8vICAgZXh0ZW5zaW9uICAgXHUyMDEzIE9CQidzIG93biBVSSAocGFsZXR0ZSwgaGVscCkuXG4vLyAgIHRvZG8gICAgICAgIFx1MjAxMyBub3QgZGVjaWRlZCB5ZXQ7IHRoZSBrZXkgaXMgc3dhbGxvd2VkIGFuZCBhIGhpbnQgc2F5cyBzby5cbi8vXG4vLyBPdXRsb29rJ3MgR21haWwtbW9kZSBrZXlzIGNvbWUgZnJvbSBpdHMgb3duIHJpYmJvbiB0b29sdGlwcyBpbiB0aGUgT2N0IDIwMjYgY2FwdHVyZTpcbi8vIEMgbmV3IG1haWwsICMgZGVsZXRlLCBFIGFyY2hpdmUsIFYgbW92ZSwgXHUyMUU3SS9cdTIxRTdVIHJlYWQvdW5yZWFkLCBCIHNub296ZSwgTCBjYXRlZ29yaXplLFxuLy8gUi9BL0YgcmVwbHkvcmVwbHkgYWxsL2ZvcndhcmQsIFx1MjMxOFogdW5kby5cblxuaW1wb3J0IHR5cGUgeyBTZWxlY3Rvck5hbWUgfSBmcm9tIFwiLi4vc2VsZWN0b3JzLnRzXCI7XG5cbmV4cG9ydCB0eXBlIEV4dGVuc2lvbkNvbW1hbmQgPSBcImhlbHBcIiB8IFwicGFsZXR0ZS1nb1wiO1xuXG5leHBvcnQgdHlwZSBTdHJhdGVneSA9XG4gIHwgeyBraW5kOiBcInBhc3N0aHJvdWdoXCIgfVxuICB8IHsga2luZDogXCJkcml2ZVwiOyB0YXJnZXQ6IFNlbGVjdG9yTmFtZSB9XG4gIHwgeyBraW5kOiBcInRvZ2dsZVwiOyB0YXJnZXQ6IFNlbGVjdG9yTmFtZTsgd2FudDogXCJyZWFkXCIgfCBcInVucmVhZFwiIH1cbiAgfCB7IGtpbmQ6IFwiZXh0ZW5zaW9uXCI7IGNvbW1hbmQ6IEV4dGVuc2lvbkNvbW1hbmQgfVxuICB8IHsga2luZDogXCJ0b2RvXCIgfTtcblxuZXhwb3J0IGludGVyZmFjZSBBY3Rpb24ge1xuICBkZXNjcmlwdGlvbjogc3RyaW5nO1xuICBzdHJhdGVneTogU3RyYXRlZ3k7XG59XG5cbmNvbnN0IHBhc3N0aHJvdWdoOiBTdHJhdGVneSA9IHsga2luZDogXCJwYXNzdGhyb3VnaFwiIH07XG5jb25zdCB0b2RvOiBTdHJhdGVneSA9IHsga2luZDogXCJ0b2RvXCIgfTtcbmNvbnN0IGRyaXZlID0gKHRhcmdldDogU2VsZWN0b3JOYW1lKTogU3RyYXRlZ3kgPT4gKHsga2luZDogXCJkcml2ZVwiLCB0YXJnZXQgfSk7XG5jb25zdCBleHRlbnNpb24gPSAoY29tbWFuZDogRXh0ZW5zaW9uQ29tbWFuZCk6IFN0cmF0ZWd5ID0+ICh7IGtpbmQ6IFwiZXh0ZW5zaW9uXCIsIGNvbW1hbmQgfSk7XG5jb25zdCByZWFkU3RhdGUgPSAod2FudDogXCJyZWFkXCIgfCBcInVucmVhZFwiKTogU3RyYXRlZ3kgPT4gKHtcbiAga2luZDogXCJ0b2dnbGVcIixcbiAgdGFyZ2V0OiBcInJlYWRVbnJlYWRCdXR0b25cIixcbiAgd2FudCxcbn0pO1xuXG5leHBvcnQgY29uc3QgQUNUSU9OUyA9IHtcbiAgY29tcG9zZTogeyBkZXNjcmlwdGlvbjogXCJDb21wb3NlXCIsIHN0cmF0ZWd5OiBwYXNzdGhyb3VnaCB9LFxuICBzZWFyY2g6IHsgZGVzY3JpcHRpb246IFwiRm9jdXMgc2VhcmNoXCIsIHN0cmF0ZWd5OiBwYXNzdGhyb3VnaCB9LFxuICBoZWxwOiB7IGRlc2NyaXB0aW9uOiBcIlNob3cga2V5Ym9hcmQgc2hvcnRjdXRzXCIsIHN0cmF0ZWd5OiBleHRlbnNpb24oXCJoZWxwXCIpIH0sXG4gIG5leHQ6IHsgZGVzY3JpcHRpb246IFwiTmV4dCBjb252ZXJzYXRpb25cIiwgc3RyYXRlZ3k6IHBhc3N0aHJvdWdoIH0sXG4gIHByZXZpb3VzOiB7IGRlc2NyaXB0aW9uOiBcIlByZXZpb3VzIGNvbnZlcnNhdGlvblwiLCBzdHJhdGVneTogcGFzc3Rocm91Z2ggfSxcbiAgb3BlbjogeyBkZXNjcmlwdGlvbjogXCJPcGVuIGNvbnZlcnNhdGlvblwiLCBzdHJhdGVneTogcGFzc3Rocm91Z2ggfSxcbiAgc2VsZWN0OiB7IGRlc2NyaXB0aW9uOiBcIlNlbGVjdCBvciBkZXNlbGVjdFwiLCBzdHJhdGVneTogcGFzc3Rocm91Z2ggfSxcbiAgc2VsZWN0UmFuZ2U6IHsgZGVzY3JpcHRpb246IFwiU2VsZWN0IHJhbmdlIHNpbmNlIGxhc3Qgc2VsZWN0aW9uXCIsIHN0cmF0ZWd5OiB0b2RvIH0sXG4gIHNlbGVjdEFsbDogeyBkZXNjcmlwdGlvbjogXCJTZWxlY3QgYWxsXCIsIHN0cmF0ZWd5OiB0b2RvIH0sXG4gIHNlbGVjdE5vbmU6IHsgZGVzY3JpcHRpb246IFwiU2VsZWN0IG5vbmVcIiwgc3RyYXRlZ3k6IHRvZG8gfSxcbiAgYXJjaGl2ZTogeyBkZXNjcmlwdGlvbjogXCJBcmNoaXZlXCIsIHN0cmF0ZWd5OiBkcml2ZShcImFyY2hpdmVCdXR0b25cIikgfSxcbiAgYXJjaGl2ZVByZXZpb3VzOiB7IGRlc2NyaXB0aW9uOiBcIkFyY2hpdmUsIHRoZW4gb3BlbiBwcmV2aW91c1wiLCBzdHJhdGVneTogdG9kbyB9LFxuICBhcmNoaXZlTmV4dDogeyBkZXNjcmlwdGlvbjogXCJBcmNoaXZlLCB0aGVuIG9wZW4gbmV4dFwiLCBzdHJhdGVneTogdG9kbyB9LFxuICBkZWxldGU6IHsgZGVzY3JpcHRpb246IFwiRGVsZXRlXCIsIHN0cmF0ZWd5OiBkcml2ZShcImRlbGV0ZUJ1dHRvblwiKSB9LFxuICBqdW5rOiB7IGRlc2NyaXB0aW9uOiBcIlJlcG9ydCBqdW5rXCIsIHN0cmF0ZWd5OiBwYXNzdGhyb3VnaCB9LFxuICAvLyBPdXRsb29rJ3Mgb3duIFwiTW92ZSB0b1wiIHBpY2tlciBoYXMgZm9sZGVyIHNlYXJjaCwgd2hpY2ggY292ZXJzIEZhc3RtYWlsJ3MgYG1gIHVudGlsIHRoZVxuICAvLyBwYWxldHRlIGNhbiBtb3ZlIG1lc3NhZ2VzIGl0c2VsZiAoZG9jcy9kZWNpc2lvbnMubWQpLlxuICBtb3ZlOiB7IGRlc2NyaXB0aW9uOiBcIk1vdmUgdG8gZm9sZGVyXHUyMDI2XCIsIHN0cmF0ZWd5OiBkcml2ZShcIm1vdmVCdXR0b25cIikgfSxcbiAgZ29UbzogeyBkZXNjcmlwdGlvbjogXCJHbyB0byBmb2xkZXIgb3IgYXBwXHUyMDI2XCIsIHN0cmF0ZWd5OiBleHRlbnNpb24oXCJwYWxldHRlLWdvXCIpIH0sXG4gIHNub296ZTogeyBkZXNjcmlwdGlvbjogXCJTbm9vemVcIiwgc3RyYXRlZ3k6IHBhc3N0aHJvdWdoIH0sXG4gIG1hcmtSZWFkOiB7IGRlc2NyaXB0aW9uOiBcIk1hcmsgcmVhZFwiLCBzdHJhdGVneTogcmVhZFN0YXRlKFwicmVhZFwiKSB9LFxuICBtYXJrVW5yZWFkOiB7IGRlc2NyaXB0aW9uOiBcIk1hcmsgdW5yZWFkXCIsIHN0cmF0ZWd5OiByZWFkU3RhdGUoXCJ1bnJlYWRcIikgfSxcbiAgcmVwbHk6IHsgZGVzY3JpcHRpb246IFwiUmVwbHlcIiwgc3RyYXRlZ3k6IHBhc3N0aHJvdWdoIH0sXG4gIHJlcGx5QWxsOiB7IGRlc2NyaXB0aW9uOiBcIlJlcGx5IGFsbFwiLCBzdHJhdGVneTogcGFzc3Rocm91Z2ggfSxcbiAgZm9yd2FyZDogeyBkZXNjcmlwdGlvbjogXCJGb3J3YXJkXCIsIHN0cmF0ZWd5OiBwYXNzdGhyb3VnaCB9LFxuICB1bmRvOiB7IGRlc2NyaXB0aW9uOiBcIlVuZG9cIiwgc3RyYXRlZ3k6IGRyaXZlKFwidW5kb0J1dHRvblwiKSB9LFxuICByZWZyZXNoOiB7IGRlc2NyaXB0aW9uOiBcIlJlZnJlc2hcIiwgc3RyYXRlZ3k6IGRyaXZlKFwic3luY0J1dHRvblwiKSB9LFxuICBnb01haWw6IHsgZGVzY3JpcHRpb246IFwiR28gdG8gTWFpbFwiLCBzdHJhdGVneTogZHJpdmUoXCJtYWlsQXBwXCIpIH0sXG4gIGdvQ2FsZW5kYXI6IHsgZGVzY3JpcHRpb246IFwiR28gdG8gQ2FsZW5kYXJcIiwgc3RyYXRlZ3k6IGRyaXZlKFwiY2FsZW5kYXJBcHBcIikgfSxcbiAgZ29Db250YWN0czogeyBkZXNjcmlwdGlvbjogXCJHbyB0byBDb250YWN0c1wiLCBzdHJhdGVneTogZHJpdmUoXCJwZW9wbGVBcHBcIikgfSxcbiAgZ29TZXR0aW5nczogeyBkZXNjcmlwdGlvbjogXCJHbyB0byBTZXR0aW5nc1wiLCBzdHJhdGVneTogZHJpdmUoXCJzZXR0aW5nc0J1dHRvblwiKSB9LFxufSBhcyBjb25zdCBzYXRpc2ZpZXMgUmVjb3JkPHN0cmluZywgQWN0aW9uPjtcblxuZXhwb3J0IHR5cGUgQWN0aW9uSWQgPSBrZXlvZiB0eXBlb2YgQUNUSU9OUztcbiIsICIvLyBUaGUgRmFzdG1haWwtc3R5bGUgYmluZGluZyBzZXQgKFBMQU4ubWQsIFNob3J0Y3V0IG1hcCkuIEtleXMgYXJlIEtleWJvYXJkRXZlbnQua2V5IHZhbHVlcy5cblxuaW1wb3J0IHR5cGUgeyBBY3Rpb25JZCB9IGZyb20gXCIuL2FjdGlvbnMudHNcIjtcbmltcG9ydCB0eXBlIHsgQmluZGluZywgQ29udGV4dCB9IGZyb20gXCIuL2VuZ2luZS50c1wiO1xuXG5jb25zdCBNQUlMOiByZWFkb25seSBDb250ZXh0W10gPSBbXCJsaXN0XCIsIFwicmVhZGluZ1wiXTtcbmNvbnN0IExJU1Q6IHJlYWRvbmx5IENvbnRleHRbXSA9IFtcImxpc3RcIl07XG5jb25zdCBSRUFESU5HOiByZWFkb25seSBDb250ZXh0W10gPSBbXCJyZWFkaW5nXCJdO1xuXG5jb25zdCBiaW5kID0gKGtleXM6IHN0cmluZywgYWN0aW9uOiBBY3Rpb25JZCwgY29udGV4dHMgPSBNQUlMKTogQmluZGluZzxBY3Rpb25JZD4gPT4gKHtcbiAga2V5czoga2V5cy5zcGxpdChcIiBcIiksXG4gIGFjdGlvbixcbiAgY29udGV4dHMsXG59KTtcblxuZXhwb3J0IGNvbnN0IEJJTkRJTkdTOiByZWFkb25seSBCaW5kaW5nPEFjdGlvbklkPltdID0gW1xuICBiaW5kKFwiY1wiLCBcImNvbXBvc2VcIiksXG4gIGJpbmQoXCIvXCIsIFwic2VhcmNoXCIpLFxuICBiaW5kKFwiP1wiLCBcImhlbHBcIiksXG4gIGJpbmQoXCJqXCIsIFwibmV4dFwiKSxcbiAgYmluZChcImtcIiwgXCJwcmV2aW91c1wiKSxcbiAgYmluZChcIm9cIiwgXCJvcGVuXCIsIExJU1QpLFxuICBiaW5kKFwiRW50ZXJcIiwgXCJvcGVuXCIsIExJU1QpLFxuICBiaW5kKFwieFwiLCBcInNlbGVjdFwiLCBMSVNUKSxcbiAgYmluZChcIlhcIiwgXCJzZWxlY3RSYW5nZVwiLCBMSVNUKSxcbiAgYmluZChcIiogYVwiLCBcInNlbGVjdEFsbFwiLCBMSVNUKSxcbiAgYmluZChcIiogblwiLCBcInNlbGVjdE5vbmVcIiwgTElTVCksXG4gIGJpbmQoXCJ5XCIsIFwiYXJjaGl2ZVwiKSxcbiAgYmluZChcIl1cIiwgXCJhcmNoaXZlUHJldmlvdXNcIiwgUkVBRElORyksXG4gIGJpbmQoXCJbXCIsIFwiYXJjaGl2ZU5leHRcIiwgUkVBRElORyksXG4gIGJpbmQoXCJkXCIsIFwiZGVsZXRlXCIpLFxuICBiaW5kKFwiIVwiLCBcImp1bmtcIiksXG4gIGJpbmQoXCJtXCIsIFwibW92ZVwiKSxcbiAgYmluZChcImdcIiwgXCJnb1RvXCIpLFxuICBiaW5kKFwiYlwiLCBcInNub296ZVwiKSxcbiAgYmluZChcIi4gclwiLCBcIm1hcmtSZWFkXCIpLFxuICBiaW5kKFwiLiB1XCIsIFwibWFya1VucmVhZFwiKSxcbiAgYmluZChcInJcIiwgXCJyZXBseVwiLCBSRUFESU5HKSxcbiAgYmluZChcImFcIiwgXCJyZXBseUFsbFwiLCBSRUFESU5HKSxcbiAgYmluZChcImZcIiwgXCJmb3J3YXJkXCIsIFJFQURJTkcpLFxuICBiaW5kKFwielwiLCBcInVuZG9cIiksXG4gIGJpbmQoXCJ1XCIsIFwicmVmcmVzaFwiKSxcbiAgYmluZChcIkcgbVwiLCBcImdvTWFpbFwiKSxcbiAgYmluZChcIkcgY1wiLCBcImdvQ2FsZW5kYXJcIiksXG4gIGJpbmQoXCJHIGFcIiwgXCJnb0NvbnRhY3RzXCIpLFxuICBiaW5kKFwiRyBzXCIsIFwiZ29TZXR0aW5nc1wiKSxcbl07XG4iLCAiLy8gUHVyZSBrZXktc2VxdWVuY2UgZW5naW5lLiBObyBET00gYWNjZXNzLCBzbyBpdCBpcyBmdWxseSB1bml0LXRlc3RhYmxlLlxuLy9cbi8vIEtleXMgYXJlIEtleWJvYXJkRXZlbnQua2V5IHZhbHVlcywgd2hpY2ggYWxyZWFkeSByZWZsZWN0IFNoaWZ0IChcIkdcIiwgXCIqXCIsIFwiIVwiLCBcIj9cIikuXG4vLyBBIGJpbmRpbmcgbGlrZSBGYXN0bWFpbCdzIFwiU2hpZnQrRyB0aGVuIE1cIiBpcyB3cml0dGVuIFtcIkdcIiwgXCJtXCJdLlxuXG5leHBvcnQgdHlwZSBDb250ZXh0ID0gXCJsaXN0XCIgfCBcInJlYWRpbmdcIiB8IFwicGFsZXR0ZVwiO1xuXG5leHBvcnQgaW50ZXJmYWNlIEJpbmRpbmc8QSBleHRlbmRzIHN0cmluZyA9IHN0cmluZz4ge1xuICBrZXlzOiByZWFkb25seSBzdHJpbmdbXTtcbiAgYWN0aW9uOiBBO1xuICBjb250ZXh0czogcmVhZG9ubHkgQ29udGV4dFtdO1xufVxuXG5leHBvcnQgdHlwZSBGZWVkUmVzdWx0PEEgZXh0ZW5kcyBzdHJpbmc+ID1cbiAgfCB7IGtpbmQ6IFwibWF0Y2hcIjsgYWN0aW9uOiBBIH1cbiAgfCB7IGtpbmQ6IFwicGVuZGluZ1wiOyBwcmVmaXg6IHJlYWRvbmx5IHN0cmluZ1tdOyBuZXh0OiByZWFkb25seSBCaW5kaW5nPEE+W10gfVxuICAvKiogQSBwcmVmaXggd2FzIHBlbmRpbmcgYW5kIHRoaXMga2V5IGNvbXBsZXRlZCBub3RoaW5nLiBUaGUga2V5IHNob3VsZCBiZSBzd2FsbG93ZWQuICovXG4gIHwgeyBraW5kOiBcImNhbmNlbGxlZFwiIH1cbiAgfCB7IGtpbmQ6IFwibm9uZVwiIH07XG5cbmV4cG9ydCBjb25zdCBERUZBVUxUX1RJTUVPVVRfTVMgPSAxNTAwO1xuXG5jb25zdCBzdGFydHNXaXRoID0gKGtleXM6IHJlYWRvbmx5IHN0cmluZ1tdLCBwcmVmaXg6IHJlYWRvbmx5IHN0cmluZ1tdKSA9PlxuICBwcmVmaXgubGVuZ3RoIDw9IGtleXMubGVuZ3RoICYmIHByZWZpeC5ldmVyeSgoaywgaSkgPT4ga2V5c1tpXSA9PT0gayk7XG5cbmV4cG9ydCBjbGFzcyBTZXF1ZW5jZUVuZ2luZTxBIGV4dGVuZHMgc3RyaW5nPiB7XG4gIHJlYWRvbmx5ICNiaW5kaW5nczogcmVhZG9ubHkgQmluZGluZzxBPltdO1xuICByZWFkb25seSAjdGltZW91dE1zOiBudW1iZXI7XG4gICNwcmVmaXg6IHN0cmluZ1tdID0gW107XG4gICNsYXN0QXQgPSAwO1xuXG4gIGNvbnN0cnVjdG9yKGJpbmRpbmdzOiByZWFkb25seSBCaW5kaW5nPEE+W10sIHRpbWVvdXRNcyA9IERFRkFVTFRfVElNRU9VVF9NUykge1xuICAgIHRoaXMuI2JpbmRpbmdzID0gYmluZGluZ3M7XG4gICAgdGhpcy4jdGltZW91dE1zID0gdGltZW91dE1zO1xuICB9XG5cbiAgZ2V0IHBlbmRpbmcoKTogcmVhZG9ubHkgc3RyaW5nW10ge1xuICAgIHJldHVybiB0aGlzLiNwcmVmaXg7XG4gIH1cblxuICByZXNldCgpOiB2b2lkIHtcbiAgICB0aGlzLiNwcmVmaXggPSBbXTtcbiAgfVxuXG4gIGZlZWQoa2V5OiBzdHJpbmcsIGNvbnRleHQ6IENvbnRleHQsIG5vdzogbnVtYmVyKTogRmVlZFJlc3VsdDxBPiB7XG4gICAgaWYgKHRoaXMuI3ByZWZpeC5sZW5ndGggPiAwICYmIG5vdyAtIHRoaXMuI2xhc3RBdCA+IHRoaXMuI3RpbWVvdXRNcykgdGhpcy5yZXNldCgpO1xuICAgIGNvbnN0IHdhc1BlbmRpbmcgPSB0aGlzLiNwcmVmaXgubGVuZ3RoID4gMDtcbiAgICBjb25zdCBzZXEgPSBbLi4udGhpcy4jcHJlZml4LCBrZXldO1xuICAgIGNvbnN0IGFjdGl2ZSA9IHRoaXMuI2JpbmRpbmdzLmZpbHRlcigoYikgPT4gYi5jb250ZXh0cy5pbmNsdWRlcyhjb250ZXh0KSk7XG4gICAgY29uc3QgbG9uZ2VyID0gYWN0aXZlLmZpbHRlcigoYikgPT4gYi5rZXlzLmxlbmd0aCA+IHNlcS5sZW5ndGggJiYgc3RhcnRzV2l0aChiLmtleXMsIHNlcSkpO1xuXG4gICAgaWYgKGxvbmdlci5sZW5ndGggPiAwKSB7XG4gICAgICB0aGlzLiNwcmVmaXggPSBzZXE7XG4gICAgICB0aGlzLiNsYXN0QXQgPSBub3c7XG4gICAgICByZXR1cm4geyBraW5kOiBcInBlbmRpbmdcIiwgcHJlZml4OiBzZXEsIG5leHQ6IGxvbmdlciB9O1xuICAgIH1cblxuICAgIHRoaXMucmVzZXQoKTtcbiAgICBjb25zdCBleGFjdCA9IGFjdGl2ZS5maW5kKChiKSA9PiBiLmtleXMubGVuZ3RoID09PSBzZXEubGVuZ3RoICYmIHN0YXJ0c1dpdGgoYi5rZXlzLCBzZXEpKTtcbiAgICBpZiAoZXhhY3QpIHJldHVybiB7IGtpbmQ6IFwibWF0Y2hcIiwgYWN0aW9uOiBleGFjdC5hY3Rpb24gfTtcbiAgICByZXR1cm4gd2FzUGVuZGluZyA/IHsga2luZDogXCJjYW5jZWxsZWRcIiB9IDogeyBraW5kOiBcIm5vbmVcIiB9O1xuICB9XG59XG5cbi8qKiBCaW5kaW5ncyB0aGF0IHNoYWRvdyBlYWNoIG90aGVyIChzYW1lIGtleXMsIG9yIG9uZSBhIHByZWZpeCBvZiBhbm90aGVyKSBpbiBhIHNoYXJlZCBjb250ZXh0LiAqL1xuZXhwb3J0IGZ1bmN0aW9uIGZpbmRDb25mbGljdHMoYmluZGluZ3M6IHJlYWRvbmx5IEJpbmRpbmdbXSk6IHN0cmluZ1tdIHtcbiAgY29uc3QgY29uZmxpY3RzOiBzdHJpbmdbXSA9IFtdO1xuICBiaW5kaW5ncy5mb3JFYWNoKChhLCBpKSA9PiB7XG4gICAgZm9yIChjb25zdCBiIG9mIGJpbmRpbmdzLnNsaWNlKGkgKyAxKSkge1xuICAgICAgY29uc3Qgc2hhcmVkID0gYS5jb250ZXh0cy5zb21lKChjKSA9PiBiLmNvbnRleHRzLmluY2x1ZGVzKGMpKTtcbiAgICAgIGNvbnN0IFtzaG9ydCwgbG9uZ10gPSBhLmtleXMubGVuZ3RoIDw9IGIua2V5cy5sZW5ndGggPyBbYSwgYl0gOiBbYiwgYV07XG4gICAgICBpZiAoc2hhcmVkICYmIHN0YXJ0c1dpdGgobG9uZy5rZXlzLCBzaG9ydC5rZXlzKSkge1xuICAgICAgICBjb25mbGljdHMucHVzaChcbiAgICAgICAgICBgJHtzaG9ydC5rZXlzLmpvaW4oXCIgXCIpfSAoJHtzaG9ydC5hY3Rpb259KSB2cyAke2xvbmcua2V5cy5qb2luKFwiIFwiKX0gKCR7bG9uZy5hY3Rpb259KWAsXG4gICAgICAgICk7XG4gICAgICB9XG4gICAgfVxuICB9KTtcbiAgcmV0dXJuIGNvbmZsaWN0cztcbn1cblxuaW50ZXJmYWNlIEtleUxpa2Uge1xuICBrZXk6IHN0cmluZztcbiAgY3RybEtleTogYm9vbGVhbjtcbiAgbWV0YUtleTogYm9vbGVhbjtcbiAgYWx0S2V5OiBib29sZWFuO1xufVxuXG5jb25zdCBNT0RJRklFUl9LRVlTID0gbmV3IFNldChbXCJTaGlmdFwiLCBcIkNvbnRyb2xcIiwgXCJBbHRcIiwgXCJNZXRhXCIsIFwiQ2Fwc0xvY2tcIiwgXCJEZWFkXCJdKTtcblxuLyoqIFRoZSBlbmdpbmUga2V5IGZvciBhbiBldmVudCwgb3IgbnVsbCB3aGVuIGl0IGJlbG9uZ3MgdG8gdGhlIGJyb3dzZXIvT1MgKEN0cmwsIENtZCwgQWx0IGNob3JkcykuICovXG5leHBvcnQgZnVuY3Rpb24ga2V5RnJvbUV2ZW50KGV2ZW50OiBLZXlMaWtlKTogc3RyaW5nIHwgbnVsbCB7XG4gIGlmIChldmVudC5jdHJsS2V5IHx8IGV2ZW50Lm1ldGFLZXkgfHwgZXZlbnQuYWx0S2V5KSByZXR1cm4gbnVsbDtcbiAgaWYgKE1PRElGSUVSX0tFWVMuaGFzKGV2ZW50LmtleSkpIHJldHVybiBudWxsO1xuICByZXR1cm4gZXZlbnQua2V5O1xufVxuIiwgIjpob3N0IHtcbiAgYWxsOiBpbml0aWFsO1xuICAtLW9sLWJnOiAjZmZmO1xuICAtLW9sLWZnOiAjMWYxZjFmO1xuICAtLW9sLW11dGVkOiAjNmI2YjZiO1xuICAtLW9sLWFjY2VudDogIzJmNmZkZTtcbiAgLS1vbC1ib3JkZXI6IHJnYigwIDAgMCAvIDEyJSk7XG4gIC0tb2wtc2hhZG93OiAwIDhweCAzMnB4IHJnYigwIDAgMCAvIDE4JSk7XG4gIGZvbnQ6XG4gICAgMTNweCAvIDEuNCBzeXN0ZW0tdWksXG4gICAgLWFwcGxlLXN5c3RlbSxcbiAgICBcIlNlZ29lIFVJXCIsXG4gICAgc2Fucy1zZXJpZjtcbn1cblxuQG1lZGlhIChwcmVmZXJzLWNvbG9yLXNjaGVtZTogZGFyaykge1xuICA6aG9zdCB7XG4gICAgLS1vbC1iZzogIzI2MjYyNjtcbiAgICAtLW9sLWZnOiAjZWNlY2VjO1xuICAgIC0tb2wtbXV0ZWQ6ICNhMGEwYTA7XG4gICAgLS1vbC1hY2NlbnQ6ICM3YWE3ZmY7XG4gICAgLS1vbC1ib3JkZXI6IHJnYigyNTUgMjU1IDI1NSAvIDE0JSk7XG4gIH1cbn1cblxuW2hpZGRlbl0ge1xuICBkaXNwbGF5OiBub25lICFpbXBvcnRhbnQ7XG59XG5cbi50b2FzdCxcbi5oaW50LFxuLmhlYWx0aCB7XG4gIHBvc2l0aW9uOiBmaXhlZDtcbiAgei1pbmRleDogMjE0NzQ4MzY0NztcbiAgYmFja2dyb3VuZDogdmFyKC0tb2wtYmcpO1xuICBjb2xvcjogdmFyKC0tb2wtZmcpO1xuICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1vbC1ib3JkZXIpO1xuICBib3JkZXItcmFkaXVzOiA2cHg7XG4gIGJveC1zaGFkb3c6IHZhcigtLW9sLXNoYWRvdyk7XG4gIHBhZGRpbmc6IDZweCAxMHB4O1xufVxuXG4udG9hc3Qge1xuICBib3R0b206IDI0cHg7XG4gIGxlZnQ6IDUwJTtcbiAgdHJhbnNmb3JtOiB0cmFuc2xhdGVYKC01MCUpO1xufVxuXG4udG9hc3Qubm90aWNlIHtcbiAgZGlzcGxheTogZmxleDtcbiAgYWxpZ24taXRlbXM6IGNlbnRlcjtcbiAgZ2FwOiAxMnB4O1xuICBwYWRkaW5nOiA4cHggMTBweCA4cHggMTRweDtcbn1cblxuLmFjdGlvbiB7XG4gIGFsbDogdW5zZXQ7XG4gIGN1cnNvcjogcG9pbnRlcjtcbiAgcGFkZGluZzogNHB4IDEycHg7XG4gIGJvcmRlci1yYWRpdXM6IDk5OXB4O1xuICBiYWNrZ3JvdW5kOiB2YXIoLS1vbC1hY2NlbnQpO1xuICBjb2xvcjogI2ZmZjtcbiAgZm9udC13ZWlnaHQ6IDYwMDtcbn1cblxuLmFjdGlvbjpmb2N1cy12aXNpYmxlLFxuLmNsb3NlOmZvY3VzLXZpc2libGUge1xuICBvdXRsaW5lOiAycHggc29saWQgdmFyKC0tb2wtYWNjZW50KTtcbiAgb3V0bGluZS1vZmZzZXQ6IDJweDtcbn1cblxuLmhpbnQge1xuICBib3R0b206IDI0cHg7XG4gIHJpZ2h0OiAyNHB4O1xuICBmb250LWZhbWlseTogdWktbW9ub3NwYWNlLCBcIlNGIE1vbm9cIiwgTWVubG8sIG1vbm9zcGFjZTtcbn1cblxuLmhlYWx0aCB7XG4gIHRvcDogOHB4O1xuICByaWdodDogOHB4O1xuICBkaXNwbGF5OiBmbGV4O1xuICBnYXA6IDhweDtcbiAgYWxpZ24taXRlbXM6IGNlbnRlcjtcbiAgYm9yZGVyLWNvbG9yOiAjZDk3NzA2O1xufVxuXG4uY2xvc2Uge1xuICBhbGw6IHVuc2V0O1xuICBjdXJzb3I6IHBvaW50ZXI7XG4gIGNvbG9yOiB2YXIoLS1vbC1tdXRlZCk7XG4gIHBhZGRpbmc6IDAgMnB4O1xufVxuXG4uYmFja2Ryb3Age1xuICBwb3NpdGlvbjogZml4ZWQ7XG4gIGluc2V0OiAwO1xuICB6LWluZGV4OiAyMTQ3NDgzNjQ2O1xuICBiYWNrZ3JvdW5kOiByZ2IoMCAwIDAgLyAyNSUpO1xuICBkaXNwbGF5OiBmbGV4O1xuICBqdXN0aWZ5LWNvbnRlbnQ6IGNlbnRlcjtcbiAgYWxpZ24taXRlbXM6IGZsZXgtc3RhcnQ7XG4gIHBhZGRpbmctdG9wOiAxMnZoO1xufVxuXG4ucGFuZWwge1xuICBiYWNrZ3JvdW5kOiB2YXIoLS1vbC1iZyk7XG4gIGNvbG9yOiB2YXIoLS1vbC1mZyk7XG4gIGJvcmRlci1yYWRpdXM6IDEwcHg7XG4gIGJveC1zaGFkb3c6IHZhcigtLW9sLXNoYWRvdyk7XG4gIHdpZHRoOiBtaW4oNTIwcHgsIDkwdncpO1xuICBtYXgtaGVpZ2h0OiA3MHZoO1xuICBvdmVyZmxvdzogYXV0bztcbn1cblxuLmhlbHAge1xuICBwYWRkaW5nOiAxNnB4IDIwcHg7XG59XG5cbi5oZWxwIGgyIHtcbiAgZm9udC1zaXplOiAxNXB4O1xuICBtYXJnaW46IDAgMCAxMnB4O1xufVxuXG4uaGVscCB0YWJsZSB7XG4gIGJvcmRlci1jb2xsYXBzZTogY29sbGFwc2U7XG4gIHdpZHRoOiAxMDAlO1xufVxuXG4uaGVscCB0ZCB7XG4gIHBhZGRpbmc6IDNweCAwO1xufVxuXG4ua2V5cyB7XG4gIHdpZHRoOiA3ZW07XG59XG5cbmtiZCB7XG4gIGZvbnQtZmFtaWx5OiB1aS1tb25vc3BhY2UsIFwiU0YgTW9ub1wiLCBNZW5sbywgbW9ub3NwYWNlO1xuICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1vbC1ib3JkZXIpO1xuICBib3JkZXItcmFkaXVzOiA0cHg7XG4gIHBhZGRpbmc6IDAgNXB4O1xufVxuXG4ucGFsZXR0ZSBpbnB1dCB7XG4gIGFsbDogdW5zZXQ7XG4gIGJveC1zaXppbmc6IGJvcmRlci1ib3g7XG4gIHdpZHRoOiAxMDAlO1xuICBwYWRkaW5nOiAxMnB4IDE2cHg7XG4gIGZvbnQtc2l6ZTogMTVweDtcbiAgYm9yZGVyLWJvdHRvbTogMXB4IHNvbGlkIHZhcigtLW9sLWJvcmRlcik7XG59XG5cbi5wYWxldHRlIHVsIHtcbiAgbGlzdC1zdHlsZTogbm9uZTtcbiAgbWFyZ2luOiAwO1xuICBwYWRkaW5nOiA0cHggMDtcbn1cblxuLnBhbGV0dGUgbGkge1xuICBwYWRkaW5nOiA1cHggMTZweDtcbiAgY3Vyc29yOiBwb2ludGVyO1xufVxuXG4ucGFsZXR0ZSBsaVthcmlhLXNlbGVjdGVkPVwidHJ1ZVwiXSB7XG4gIGJhY2tncm91bmQ6IGNvbG9yLW1peChpbiBzcmdiLCB2YXIoLS1vbC1hY2NlbnQpIDE2JSwgdHJhbnNwYXJlbnQpO1xufVxuXG4ucGFsZXR0ZSBsaSAucGF0aCB7XG4gIGNvbG9yOiB2YXIoLS1vbC1tdXRlZCk7XG4gIG1hcmdpbi1sZWZ0OiA2cHg7XG59XG5cbi5wYWxldHRlIG1hcmsge1xuICBiYWNrZ3JvdW5kOiBub25lO1xuICBjb2xvcjogdmFyKC0tb2wtYWNjZW50KTtcbiAgZm9udC13ZWlnaHQ6IDYwMDtcbn1cblxuLnBhbGV0dGUgLmVtcHR5IHtcbiAgY29sb3I6IHZhcigtLW9sLW11dGVkKTtcbiAgY3Vyc29yOiBkZWZhdWx0O1xufVxuIiwgIi8vIE9CQidzIG93biBVSSBsaXZlcyBpbiBvbmUgc2hhZG93IHJvb3QsIHNvIE91dGxvb2sncyBDU1MgY2Fubm90IHJlYWNoIGl0IGFuZCBvdXJzIGNhbm5vdFxuLy8gbGVhayBpbnRvIE91dGxvb2suIEV2ZXJ5dGhpbmcgaGVyZSBpcyBwbGFpbiBET007IG5vIGZyYW1ld29yay5cblxuaW1wb3J0IGNzcyBmcm9tIFwiLi9vdmVybGF5LmNzc1wiO1xuXG5sZXQgc2hhZG93OiBTaGFkb3dSb290IHwgbnVsbCA9IG51bGw7XG5cbmV4cG9ydCBmdW5jdGlvbiBvdmVybGF5Um9vdCgpOiBTaGFkb3dSb290IHtcbiAgaWYgKHNoYWRvdykgcmV0dXJuIHNoYWRvdztcbiAgY29uc3QgaG9zdCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoXCJvdmVybG9vay1yb290XCIpO1xuICBkb2N1bWVudC5kb2N1bWVudEVsZW1lbnQuYXBwZW5kKGhvc3QpO1xuICBzaGFkb3cgPSBob3N0LmF0dGFjaFNoYWRvdyh7IG1vZGU6IFwib3BlblwiIH0pO1xuICAvLyBLZXlzdHJva2VzIHR5cGVkIGludG8gb3VyIFVJIG11c3Qgbm90IHJlYWNoIE91dGxvb2sncyBkb2N1bWVudC1sZXZlbCBzaG9ydGN1dCBoYW5kbGVycyxcbiAgLy8gd2hpY2ggd291bGQgc2VlIHRoZSBob3N0IGVsZW1lbnQgKG5vdCBhbiBpbnB1dCkgYXMgdGhlIHRhcmdldC5cbiAgZm9yIChjb25zdCB0eXBlIG9mIFtcImtleWRvd25cIiwgXCJrZXlwcmVzc1wiLCBcImtleXVwXCJdKSB7XG4gICAgaG9zdC5hZGRFdmVudExpc3RlbmVyKHR5cGUsIChlKSA9PiBlLnN0b3BQcm9wYWdhdGlvbigpKTtcbiAgfVxuICBjb25zdCBzdHlsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoXCJzdHlsZVwiKTtcbiAgc3R5bGUudGV4dENvbnRlbnQgPSBjc3M7XG4gIHNoYWRvdy5hcHBlbmQoc3R5bGUpO1xuICByZXR1cm4gc2hhZG93O1xufVxuXG4vKiogVHJ1ZSB3aGVuIHRoZSBldmVudCBjYW1lIGZyb20gaW5zaWRlIE9CQidzIG93biBVSS4gKi9cbmV4cG9ydCBmdW5jdGlvbiBpc0Zyb21PdmVybGF5KGV2ZW50OiBFdmVudCk6IGJvb2xlYW4ge1xuICByZXR1cm4gc2hhZG93ICE9PSBudWxsICYmIGV2ZW50LmNvbXBvc2VkUGF0aCgpLmluY2x1ZGVzKHNoYWRvdy5ob3N0KTtcbn1cblxuLyoqIFRydWUgd2hpbGUgYSBtb2RhbCAocGFsZXR0ZSBvciBoZWxwKSBpcyBvcGVuOyBpdCBvd25zIHRoZSBrZXlib2FyZCB1bnRpbCBjbG9zZWQuICovXG5leHBvcnQgZnVuY3Rpb24gaXNNb2RhbE9wZW4oKTogYm9vbGVhbiB7XG4gIHJldHVybiBzaGFkb3c/LnF1ZXJ5U2VsZWN0b3IoXCIuYmFja2Ryb3BcIikgIT0gbnVsbDtcbn1cblxuZnVuY3Rpb24gZWw8SyBleHRlbmRzIGtleW9mIEhUTUxFbGVtZW50VGFnTmFtZU1hcD4odGFnOiBLLCBjbGFzc05hbWU6IHN0cmluZywgdGV4dD86IHN0cmluZykge1xuICBjb25zdCBub2RlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCh0YWcpO1xuICBub2RlLmNsYXNzTmFtZSA9IGNsYXNzTmFtZTtcbiAgaWYgKHRleHQgIT09IHVuZGVmaW5lZCkgbm9kZS50ZXh0Q29udGVudCA9IHRleHQ7XG4gIHJldHVybiBub2RlO1xufVxuXG5mdW5jdGlvbiBzaW5nbGV0b24oY2xhc3NOYW1lOiBzdHJpbmcpOiBIVE1MRGl2RWxlbWVudCB7XG4gIGNvbnN0IHJvb3QgPSBvdmVybGF5Um9vdCgpO1xuICBjb25zdCBleGlzdGluZyA9IHJvb3QucXVlcnlTZWxlY3RvcjxIVE1MRGl2RWxlbWVudD4oYC4ke2NsYXNzTmFtZX1gKTtcbiAgaWYgKGV4aXN0aW5nKSByZXR1cm4gZXhpc3Rpbmc7XG4gIGNvbnN0IG5vZGUgPSBlbChcImRpdlwiLCBjbGFzc05hbWUpO1xuICByb290LmFwcGVuZChub2RlKTtcbiAgcmV0dXJuIG5vZGU7XG59XG5cbmxldCB0b2FzdFRpbWVyOiBSZXR1cm5UeXBlPHR5cGVvZiBzZXRUaW1lb3V0PiB8IHVuZGVmaW5lZDtcblxuZXhwb3J0IGZ1bmN0aW9uIHRvYXN0KG1lc3NhZ2U6IHN0cmluZywgbXMgPSAyNTAwKTogdm9pZCB7XG4gIGNvbnN0IG5vZGUgPSBzaW5nbGV0b24oXCJ0b2FzdFwiKTtcbiAgbm9kZS5jbGFzc0xpc3QucmVtb3ZlKFwibm90aWNlXCIpO1xuICBub2RlLnRleHRDb250ZW50ID0gbWVzc2FnZTtcbiAgbm9kZS5oaWRkZW4gPSBmYWxzZTtcbiAgY2xlYXJUaW1lb3V0KHRvYXN0VGltZXIpO1xuICB0b2FzdFRpbWVyID0gc2V0VGltZW91dCgoKSA9PiB7XG4gICAgbm9kZS5oaWRkZW4gPSB0cnVlO1xuICB9LCBtcyk7XG59XG5cbi8qKiBBIHRvYXN0IHRoYXQgc3RheXMgdW50aWwgYWN0ZWQgb246IGEgbWVzc2FnZSwgb25lIGFjdGlvbiBidXR0b24sIGFuZCBhIGRpc21pc3MgXHUwMEQ3LiAqL1xuZXhwb3J0IGZ1bmN0aW9uIG5vdGljZShtZXNzYWdlOiBzdHJpbmcsIGFjdGlvbjogc3RyaW5nLCBvbkFjdGlvbjogKCkgPT4gdm9pZCk6IHZvaWQge1xuICBjb25zdCBub2RlID0gc2luZ2xldG9uKFwidG9hc3RcIik7XG4gIGNsZWFyVGltZW91dCh0b2FzdFRpbWVyKTtcbiAgY29uc3QgYnV0dG9uID0gZWwoXCJidXR0b25cIiwgXCJhY3Rpb25cIiwgYWN0aW9uKTtcbiAgYnV0dG9uLmFkZEV2ZW50TGlzdGVuZXIoXCJjbGlja1wiLCBvbkFjdGlvbik7XG4gIGNvbnN0IGNsb3NlID0gZWwoXCJidXR0b25cIiwgXCJjbG9zZVwiLCBcIlx1MDBEN1wiKTtcbiAgY2xvc2Uuc2V0QXR0cmlidXRlKFwiYXJpYS1sYWJlbFwiLCBcIkRpc21pc3NcIik7XG4gIGNsb3NlLmFkZEV2ZW50TGlzdGVuZXIoXCJjbGlja1wiLCAoKSA9PiB7XG4gICAgbm9kZS5oaWRkZW4gPSB0cnVlO1xuICB9KTtcbiAgbm9kZS5yZXBsYWNlQ2hpbGRyZW4oZWwoXCJzcGFuXCIsIFwiXCIsIG1lc3NhZ2UpLCBidXR0b24sIGNsb3NlKTtcbiAgbm9kZS5jbGFzc0xpc3QuYWRkKFwibm90aWNlXCIpO1xuICBub2RlLmhpZGRlbiA9IGZhbHNlO1xuICBidXR0b24uZm9jdXMoKTtcbn1cblxuLyoqIFRoZSBwZW5kaW5nLXByZWZpeCBoaW50LCBlLmcuIFwiRyBcdTIwMjYgbSBNYWlsIFx1MDBCNyBjIENhbGVuZGFyXCIuIFBhc3MgbnVsbCB0byBoaWRlLiAqL1xuZXhwb3J0IGZ1bmN0aW9uIHNob3dIaW50KHRleHQ6IHN0cmluZyB8IG51bGwpOiB2b2lkIHtcbiAgY29uc3Qgbm9kZSA9IHNpbmdsZXRvbihcImhpbnRcIik7XG4gIG5vZGUuaGlkZGVuID0gdGV4dCA9PT0gbnVsbDtcbiAgbm9kZS50ZXh0Q29udGVudCA9IHRleHQgPz8gXCJcIjtcbn1cblxuZXhwb3J0IGZ1bmN0aW9uIHNob3dIZWFsdGhCYWRnZShicm9rZW46IHJlYWRvbmx5IHN0cmluZ1tdLCBvbkRpc21pc3M6ICgpID0+IHZvaWQpOiB2b2lkIHtcbiAgY29uc3Qgbm9kZSA9IHNpbmdsZXRvbihcImhlYWx0aFwiKTtcbiAgbm9kZS5yZXBsYWNlQ2hpbGRyZW4oKTtcbiAgaWYgKGJyb2tlbi5sZW5ndGggPT09IDApIHtcbiAgICBub2RlLmhpZGRlbiA9IHRydWU7XG4gICAgcmV0dXJuO1xuICB9XG4gIG5vZGUuaGlkZGVuID0gZmFsc2U7XG4gIG5vZGUudGl0bGUgPSBgTm90IGZvdW5kOiAke2Jyb2tlbi5qb2luKFwiLCBcIil9LiBTZWUgdGhlIGNvbnNvbGUgZm9yIGRldGFpbHMuYDtcbiAgY29uc3QgY2xvc2UgPSBlbChcImJ1dHRvblwiLCBcImNsb3NlXCIsIFwiXHUwMEQ3XCIpO1xuICBjbG9zZS5zZXRBdHRyaWJ1dGUoXCJhcmlhLWxhYmVsXCIsIFwiRGlzbWlzc1wiKTtcbiAgY2xvc2UuYWRkRXZlbnRMaXN0ZW5lcihcImNsaWNrXCIsICgpID0+IHtcbiAgICBub2RlLmhpZGRlbiA9IHRydWU7XG4gICAgb25EaXNtaXNzKCk7XG4gIH0pO1xuICBub2RlLmFwcGVuZChlbChcInNwYW5cIiwgXCJcIiwgYE9CQjogJHticm9rZW4ubGVuZ3RofSBzZWxlY3RvcihzKSBicm9rZW5gKSwgY2xvc2UpO1xufVxuXG5leHBvcnQgZnVuY3Rpb24gc2hvd0hlbHAocm93czogcmVhZG9ubHkgeyBrZXlzOiBzdHJpbmc7IGRlc2NyaXB0aW9uOiBzdHJpbmcgfVtdKTogdm9pZCB7XG4gIGNvbnN0IHJvb3QgPSBvdmVybGF5Um9vdCgpO1xuICBjb25zdCBiYWNrZHJvcCA9IGVsKFwiZGl2XCIsIFwiYmFja2Ryb3BcIik7XG4gIGNvbnN0IHBhbmVsID0gZWwoXCJkaXZcIiwgXCJwYW5lbCBoZWxwXCIpO1xuICBwYW5lbC5hcHBlbmQoZWwoXCJoMlwiLCBcIlwiLCBcIktleWJvYXJkIHNob3J0Y3V0c1wiKSk7XG4gIGNvbnN0IHRhYmxlID0gZWwoXCJ0YWJsZVwiLCBcIlwiKTtcbiAgZm9yIChjb25zdCB7IGtleXMsIGRlc2NyaXB0aW9uIH0gb2Ygcm93cykge1xuICAgIGNvbnN0IHRyID0gZWwoXCJ0clwiLCBcIlwiKTtcbiAgICBjb25zdCBrID0gZWwoXCJ0ZFwiLCBcImtleXNcIik7XG4gICAgay5hcHBlbmQoZWwoXCJrYmRcIiwgXCJcIiwga2V5cykpO1xuICAgIHRyLmFwcGVuZChrLCBlbChcInRkXCIsIFwiXCIsIGRlc2NyaXB0aW9uKSk7XG4gICAgdGFibGUuYXBwZW5kKHRyKTtcbiAgfVxuICBwYW5lbC5hcHBlbmQodGFibGUpO1xuICBiYWNrZHJvcC5hcHBlbmQocGFuZWwpO1xuICByb290LmFwcGVuZChiYWNrZHJvcCk7XG5cbiAgY29uc3QgY2xvc2UgPSAoKSA9PiB7XG4gICAgYmFja2Ryb3AucmVtb3ZlKCk7XG4gICAgd2luZG93LnJlbW92ZUV2ZW50TGlzdGVuZXIoXCJrZXlkb3duXCIsIG9uS2V5LCB0cnVlKTtcbiAgfTtcbiAgY29uc3Qgb25LZXkgPSAoZTogS2V5Ym9hcmRFdmVudCkgPT4ge1xuICAgIGUucHJldmVudERlZmF1bHQoKTtcbiAgICBlLnN0b3BQcm9wYWdhdGlvbigpO1xuICAgIGNsb3NlKCk7XG4gIH07XG4gIGJhY2tkcm9wLmFkZEV2ZW50TGlzdGVuZXIoXCJjbGlja1wiLCAoZSkgPT4ge1xuICAgIGlmIChlLnRhcmdldCA9PT0gYmFja2Ryb3ApIGNsb3NlKCk7XG4gIH0pO1xuICB3aW5kb3cuYWRkRXZlbnRMaXN0ZW5lcihcImtleWRvd25cIiwgb25LZXksIHRydWUpO1xufVxuXG5leHBvcnQgeyBlbCB9O1xuIiwgIi8vIFB1cmUgZnV6enkgbWF0Y2hlciBmb3IgdGhlIGZvbGRlciBwYWxldHRlOiBjYXNlLWluc2Vuc2l0aXZlIHN1YnNlcXVlbmNlIG1hdGNoaW5nLCBzY29yZWQgdG9cbi8vIGZhdm91ciB3b3JkIHN0YXJ0cyBhbmQgY29uc2VjdXRpdmUgcnVucywgdGhlIHdheSBGYXN0bWFpbCdzIGFuZCBtb3N0IGVkaXRvcnMnIHBpY2tlcnMgZmVlbC5cblxuZXhwb3J0IGludGVyZmFjZSBGdXp6eU1hdGNoPFQ+IHtcbiAgaXRlbTogVDtcbiAgc2NvcmU6IG51bWJlcjtcbiAgLyoqIEluZGljZXMgaW50byB0aGUgaXRlbSdzIHRleHQgdGhhdCBtYXRjaGVkLCBmb3IgaGlnaGxpZ2h0aW5nLiAqL1xuICBpbmRpY2VzOiBudW1iZXJbXTtcbn1cblxuY29uc3QgaXNCb3VuZGFyeSA9ICh0ZXh0OiBzdHJpbmcsIGk6IG51bWJlcikgPT4gaSA9PT0gMCB8fCAvW1xccy9cXFxcLl8tXS8udGVzdCh0ZXh0W2kgLSAxXSA/PyBcIlwiKTtcblxuZnVuY3Rpb24gaXNTdWJzZXF1ZW5jZShxdWVyeTogc3RyaW5nLCB0ZXh0OiBzdHJpbmcsIGZyb206IG51bWJlcik6IGJvb2xlYW4ge1xuICBsZXQgYXQgPSBmcm9tO1xuICBmb3IgKGNvbnN0IGNoIG9mIHF1ZXJ5KSB7XG4gICAgYXQgPSB0ZXh0LmluZGV4T2YoY2gsIGF0KSArIDE7XG4gICAgaWYgKGF0ID09PSAwKSByZXR1cm4gZmFsc2U7XG4gIH1cbiAgcmV0dXJuIHRydWU7XG59XG5cbi8qKiBTY29yZSBgcXVlcnlgIGFnYWluc3QgYHRleHRgLCBvciBudWxsIGlmIGl0IGlzIG5vdCBhIHN1YnNlcXVlbmNlLiBIaWdoZXIgaXMgYmV0dGVyLiAqL1xuZXhwb3J0IGZ1bmN0aW9uIHNjb3JlKHF1ZXJ5OiBzdHJpbmcsIHRleHQ6IHN0cmluZyk6IHsgc2NvcmU6IG51bWJlcjsgaW5kaWNlczogbnVtYmVyW10gfSB8IG51bGwge1xuICBjb25zdCBxID0gcXVlcnkudG9Mb3dlckNhc2UoKTtcbiAgY29uc3QgdCA9IHRleHQudG9Mb3dlckNhc2UoKTtcbiAgaWYgKHEubGVuZ3RoID09PSAwKSByZXR1cm4geyBzY29yZTogMCwgaW5kaWNlczogW10gfTtcblxuICBjb25zdCBpbmRpY2VzOiBudW1iZXJbXSA9IFtdO1xuICBsZXQgdG90YWwgPSAwO1xuICBsZXQgZnJvbSA9IDA7XG4gIGZvciAoY29uc3QgW24sIGNoXSBvZiBbLi4ucV0uZW50cmllcygpKSB7XG4gICAgLy8gUHJlZmVyIHRoZSBuZXh0IG9jY3VycmVuY2UgYXQgYSB3b3JkIGJvdW5kYXJ5LCBhcyBsb25nIGFzIHRoZSByZXN0IG9mIHRoZSBxdWVyeSBzdGlsbFxuICAgIC8vIGZpdHMgYWZ0ZXIgaXQ7IG90aGVyd2lzZSB0YWtlIHRoZSBuZWFyZXN0IG9uZS5cbiAgICBjb25zdCByZXN0ID0gcS5zbGljZShuICsgMSk7XG4gICAgbGV0IGF0ID0gLTE7XG4gICAgZm9yIChsZXQgaSA9IHQuaW5kZXhPZihjaCwgZnJvbSk7IGkgPj0gMDsgaSA9IHQuaW5kZXhPZihjaCwgaSArIDEpKSB7XG4gICAgICBpZiAoIWlzU3Vic2VxdWVuY2UocmVzdCwgdCwgaSArIDEpKSBicmVhaztcbiAgICAgIGlmIChhdCA8IDApIGF0ID0gaTtcbiAgICAgIGlmIChpc0JvdW5kYXJ5KHRleHQsIGkpKSB7XG4gICAgICAgIGF0ID0gaTtcbiAgICAgICAgYnJlYWs7XG4gICAgICB9XG4gICAgfVxuICAgIGlmIChhdCA8IDApIHJldHVybiBudWxsO1xuICAgIGNvbnN0IHByZXYgPSBpbmRpY2VzLmF0KC0xKTtcbiAgICB0b3RhbCArPSAxO1xuICAgIGlmIChpc0JvdW5kYXJ5KHRleHQsIGF0KSkgdG90YWwgKz0gODtcbiAgICBpZiAocHJldiAhPT0gdW5kZWZpbmVkICYmIGF0ID09PSBwcmV2ICsgMSkgdG90YWwgKz0gNTtcbiAgICBpZiAocHJldiAhPT0gdW5kZWZpbmVkKSB0b3RhbCAtPSBNYXRoLm1pbihhdCAtIHByZXYgLSAxLCAzKTtcbiAgICBpbmRpY2VzLnB1c2goYXQpO1xuICAgIGZyb20gPSBhdCArIDE7XG4gIH1cbiAgLy8gU2hvcnRlciB0YXJnZXRzIHdpbiB0aWVzLCBzbyBcIkluYm94XCIgYmVhdHMgXCJJbmJveC9SZWNlaXB0cy8yMDE5XCIgZm9yIFwiaW5iXCIuXG4gIHRvdGFsIC09IHRleHQubGVuZ3RoICogMC4wMTtcbiAgcmV0dXJuIHsgc2NvcmU6IHRvdGFsLCBpbmRpY2VzIH07XG59XG5cbmV4cG9ydCBmdW5jdGlvbiBmdXp6eUZpbHRlcjxUPihcbiAgcXVlcnk6IHN0cmluZyxcbiAgaXRlbXM6IHJlYWRvbmx5IFRbXSxcbiAgdGV4dDogKGl0ZW06IFQpID0+IHN0cmluZyxcbik6IEZ1enp5TWF0Y2g8VD5bXSB7XG4gIGNvbnN0IG1hdGNoZXM6IEZ1enp5TWF0Y2g8VD5bXSA9IFtdO1xuICBmb3IgKGNvbnN0IGl0ZW0gb2YgaXRlbXMpIHtcbiAgICBjb25zdCByZXN1bHQgPSBzY29yZShxdWVyeSwgdGV4dChpdGVtKSk7XG4gICAgaWYgKHJlc3VsdCkgbWF0Y2hlcy5wdXNoKHsgaXRlbSwgLi4ucmVzdWx0IH0pO1xuICB9XG4gIHJldHVybiBtYXRjaGVzLnNvcnQoKGEsIGIpID0+IGIuc2NvcmUgLSBhLnNjb3JlKTtcbn1cbiIsICIvLyBGb2xkZXIgcGFsZXR0ZTogRmFzdG1haWwncyBgZ2AgKGdvIHRvIGZvbGRlcikgdHlwZS1haGVhZCBwaWNrZXIsIHBsdXMgT3V0bG9vaydzIGFwcHMuXG4vLyBNb3ZpbmcgKGBtYCkgdXNlcyBPdXRsb29rJ3Mgb3duIFwiTW92ZSB0b1wiIHBpY2tlciBmb3Igbm93OyBzZWUgZG9jcy9kZWNpc2lvbnMubWQuXG5cbmltcG9ydCB7IHR5cGUgRm9sZGVyTm9kZSwgcmVhZEZvbGRlclRyZWUsIHJlc29sdmUsIHR5cGUgU2VsZWN0b3JOYW1lIH0gZnJvbSBcIi4uL3NlbGVjdG9ycy50c1wiO1xuaW1wb3J0IHsgZWwsIG92ZXJsYXlSb290LCB0b2FzdCB9IGZyb20gXCIuLi91aS9vdmVybGF5LnRzXCI7XG5pbXBvcnQgeyB0eXBlIEZ1enp5TWF0Y2gsIGZ1enp5RmlsdGVyIH0gZnJvbSBcIi4vZnV6enkudHNcIjtcblxuY29uc3QgTUFYX1JFU1VMVFMgPSAxMjtcblxuZnVuY3Rpb24gaGlnaGxpZ2h0KHRleHQ6IHN0cmluZywgaW5kaWNlczogcmVhZG9ubHkgbnVtYmVyW10pOiBEb2N1bWVudEZyYWdtZW50IHtcbiAgY29uc3QgZnJhZyA9IGRvY3VtZW50LmNyZWF0ZURvY3VtZW50RnJhZ21lbnQoKTtcbiAgY29uc3QgaGl0ID0gbmV3IFNldChpbmRpY2VzKTtcbiAgZm9yIChjb25zdCBbaSwgY2hdIG9mIFsuLi50ZXh0XS5lbnRyaWVzKCkpIHtcbiAgICBmcmFnLmFwcGVuZChoaXQuaGFzKGkpID8gZWwoXCJtYXJrXCIsIFwiXCIsIGNoKSA6IGRvY3VtZW50LmNyZWF0ZVRleHROb2RlKGNoKSk7XG4gIH1cbiAgcmV0dXJuIGZyYWc7XG59XG5cbmNvbnN0IEFQUFM6IFtzdHJpbmcsIFNlbGVjdG9yTmFtZV1bXSA9IFtcbiAgW1wiTWFpbFwiLCBcIm1haWxBcHBcIl0sXG4gIFtcIkNhbGVuZGFyXCIsIFwiY2FsZW5kYXJBcHBcIl0sXG4gIFtcIlBlb3BsZVwiLCBcInBlb3BsZUFwcFwiXSxcbl07XG5cbi8qKiBPdXRsb29rJ3MgYXBwcywgbGlzdGVkIGFsb25nc2lkZSBmb2xkZXJzIHNvIGBnYCB0aGVuIFwiY2FsXCIgc3dpdGNoZXMgdG8gQ2FsZW5kYXIuICovXG5mdW5jdGlvbiBhcHBUYXJnZXRzKCk6IEZvbGRlck5vZGVbXSB7XG4gIHJldHVybiBBUFBTLmZsYXRNYXAoKFtuYW1lLCB0YXJnZXRdKSA9PiB7XG4gICAgY29uc3QgZWxlbWVudCA9IHJlc29sdmUodGFyZ2V0KTtcbiAgICByZXR1cm4gZWxlbWVudCA/IFt7IG5hbWUsIHBhdGg6IG5hbWUsIGVsZW1lbnQgfV0gOiBbXTtcbiAgfSk7XG59XG5cbmV4cG9ydCBmdW5jdGlvbiBvcGVuRm9sZGVyUGFsZXR0ZSgpOiB2b2lkIHtcbiAgY29uc3QgZm9sZGVycyA9IFsuLi5hcHBUYXJnZXRzKCksIC4uLnJlYWRGb2xkZXJUcmVlKCldO1xuICBpZiAoZm9sZGVycy5sZW5ndGggPT09IDApIHtcbiAgICB0b2FzdChcIk9CQiBjb3VsZG4ndCBmaW5kIHRoZSBmb2xkZXIgcGFuZVwiKTtcbiAgICByZXR1cm47XG4gIH1cblxuICBjb25zdCByb290ID0gb3ZlcmxheVJvb3QoKTtcbiAgY29uc3QgYmFja2Ryb3AgPSBlbChcImRpdlwiLCBcImJhY2tkcm9wXCIpO1xuICBjb25zdCBwYW5lbCA9IGVsKFwiZGl2XCIsIFwicGFuZWwgcGFsZXR0ZVwiKTtcbiAgY29uc3QgaW5wdXQgPSBlbChcImlucHV0XCIsIFwiXCIpO1xuICBpbnB1dC5wbGFjZWhvbGRlciA9IFwiR28gdG8gZm9sZGVyLCBvciBNYWlsLCBDYWxlbmRhciwgUGVvcGxlXHUyMDI2XCI7XG4gIGlucHV0LnNldEF0dHJpYnV0ZShcImFyaWEtbGFiZWxcIiwgaW5wdXQucGxhY2Vob2xkZXIpO1xuICBjb25zdCBsaXN0ID0gZWwoXCJ1bFwiLCBcIlwiKTtcbiAgbGlzdC5zZXRBdHRyaWJ1dGUoXCJyb2xlXCIsIFwibGlzdGJveFwiKTtcbiAgcGFuZWwuYXBwZW5kKGlucHV0LCBsaXN0KTtcbiAgYmFja2Ryb3AuYXBwZW5kKHBhbmVsKTtcbiAgcm9vdC5hcHBlbmQoYmFja2Ryb3ApO1xuXG4gIGxldCBtYXRjaGVzOiBGdXp6eU1hdGNoPEZvbGRlck5vZGU+W10gPSBbXTtcbiAgbGV0IHNlbGVjdGVkID0gMDtcblxuICBjb25zdCBjbG9zZSA9ICgpID0+IGJhY2tkcm9wLnJlbW92ZSgpO1xuICBjb25zdCBjb21taXQgPSAoKSA9PiB7XG4gICAgY29uc3QgbWF0Y2ggPSBtYXRjaGVzW3NlbGVjdGVkXTtcbiAgICBjbG9zZSgpO1xuICAgIGlmIChtYXRjaCkgKG1hdGNoLml0ZW0uZWxlbWVudCBhcyBIVE1MRWxlbWVudCkuY2xpY2soKTtcbiAgfTtcblxuICBjb25zdCByZW5kZXIgPSAoKSA9PiB7XG4gICAgbWF0Y2hlcyA9IGZ1enp5RmlsdGVyKGlucHV0LnZhbHVlLCBmb2xkZXJzLCAoZikgPT4gZi5wYXRoKS5zbGljZSgwLCBNQVhfUkVTVUxUUyk7XG4gICAgc2VsZWN0ZWQgPSBNYXRoLm1pbihzZWxlY3RlZCwgTWF0aC5tYXgobWF0Y2hlcy5sZW5ndGggLSAxLCAwKSk7XG4gICAgbGlzdC5yZXBsYWNlQ2hpbGRyZW4oXG4gICAgICAuLi5tYXRjaGVzLm1hcCgobSwgaSkgPT4ge1xuICAgICAgICBjb25zdCBsaSA9IGVsKFwibGlcIiwgXCJcIik7XG4gICAgICAgIGxpLnNldEF0dHJpYnV0ZShcInJvbGVcIiwgXCJvcHRpb25cIik7XG4gICAgICAgIGxpLnNldEF0dHJpYnV0ZShcImFyaWEtc2VsZWN0ZWRcIiwgU3RyaW5nKGkgPT09IHNlbGVjdGVkKSk7XG4gICAgICAgIGxpLmFwcGVuZChoaWdobGlnaHQobS5pdGVtLnBhdGgsIG0uaW5kaWNlcykpO1xuICAgICAgICBsaS5hZGRFdmVudExpc3RlbmVyKFwibW91c2Vkb3duXCIsIChlKSA9PiB7XG4gICAgICAgICAgZS5wcmV2ZW50RGVmYXVsdCgpO1xuICAgICAgICAgIHNlbGVjdGVkID0gaTtcbiAgICAgICAgICBjb21taXQoKTtcbiAgICAgICAgfSk7XG4gICAgICAgIHJldHVybiBsaTtcbiAgICAgIH0pLFxuICAgICk7XG4gICAgaWYgKG1hdGNoZXMubGVuZ3RoID09PSAwKSBsaXN0LmFwcGVuZChlbChcImxpXCIsIFwiZW1wdHlcIiwgXCJObyBtYXRjaGluZyBmb2xkZXJzXCIpKTtcbiAgfTtcblxuICBpbnB1dC5hZGRFdmVudExpc3RlbmVyKFwiaW5wdXRcIiwgKCkgPT4ge1xuICAgIHNlbGVjdGVkID0gMDtcbiAgICByZW5kZXIoKTtcbiAgfSk7XG4gIGlucHV0LmFkZEV2ZW50TGlzdGVuZXIoXCJrZXlkb3duXCIsIChlKSA9PiB7XG4gICAgZS5zdG9wUHJvcGFnYXRpb24oKTtcbiAgICBpZiAoZS5rZXkgPT09IFwiRXNjYXBlXCIpIGNsb3NlKCk7XG4gICAgZWxzZSBpZiAoZS5rZXkgPT09IFwiRW50ZXJcIikgY29tbWl0KCk7XG4gICAgZWxzZSBpZiAoZS5rZXkgPT09IFwiQXJyb3dEb3duXCIgfHwgKGUuY3RybEtleSAmJiBlLmtleSA9PT0gXCJuXCIpKSBzZWxlY3RlZCsrO1xuICAgIGVsc2UgaWYgKGUua2V5ID09PSBcIkFycm93VXBcIiB8fCAoZS5jdHJsS2V5ICYmIGUua2V5ID09PSBcInBcIikpXG4gICAgICBzZWxlY3RlZCA9IE1hdGgubWF4KHNlbGVjdGVkIC0gMSwgMCk7XG4gICAgZWxzZSByZXR1cm47XG4gICAgZS5wcmV2ZW50RGVmYXVsdCgpO1xuICAgIHJlbmRlcigpO1xuICB9KTtcbiAgYmFja2Ryb3AuYWRkRXZlbnRMaXN0ZW5lcihcIm1vdXNlZG93blwiLCAoZSkgPT4ge1xuICAgIGlmIChlLnRhcmdldCA9PT0gYmFja2Ryb3ApIGNsb3NlKCk7XG4gIH0pO1xuXG4gIHJlbmRlcigpO1xuICBpbnB1dC5mb2N1cygpO1xufVxuIiwgIi8vIFR1cm5zIGFuIGFjdGlvbiBpbnRvIERPTSBlZmZlY3RzLiBUaGUgb25seSBrZXltYXAgZmlsZSB0aGF0IHRvdWNoZXMgdGhlIHBhZ2UuXG5cbmltcG9ydCB7IG9wZW5Gb2xkZXJQYWxldHRlIH0gZnJvbSBcIi4uL3BhbGV0dGUvcGFsZXR0ZS50c1wiO1xuaW1wb3J0IHsgaXNVbnJlYWRSb3csIHJlc29sdmUsIHNlbGVjdGVkUm93cyB9IGZyb20gXCIuLi9zZWxlY3RvcnMudHNcIjtcbmltcG9ydCB7IHNob3dIZWxwLCB0b2FzdCB9IGZyb20gXCIuLi91aS9vdmVybGF5LnRzXCI7XG5pbXBvcnQgeyBBQ1RJT05TLCB0eXBlIEFjdGlvbklkLCB0eXBlIEV4dGVuc2lvbkNvbW1hbmQgfSBmcm9tIFwiLi9hY3Rpb25zLnRzXCI7XG5pbXBvcnQgeyBCSU5ESU5HUyB9IGZyb20gXCIuL2JpbmRpbmdzLnRzXCI7XG5pbXBvcnQgdHlwZSB7IENvbnRleHQgfSBmcm9tIFwiLi9lbmdpbmUudHNcIjtcblxuLyoqIFdoaWNoIGJpbmRpbmdzIGFwcGx5LCBqdWRnZWQgZnJvbSB3aGVyZSBmb2N1cyBpcy4gKi9cbmV4cG9ydCBmdW5jdGlvbiBjdXJyZW50Q29udGV4dCgpOiBDb250ZXh0IHtcbiAgY29uc3QgcmVhZGluZyA9IHJlc29sdmUoXCJyZWFkaW5nUGFuZVwiKTtcbiAgY29uc3QgZm9jdXNlZCA9IGRvY3VtZW50LmFjdGl2ZUVsZW1lbnQ7XG4gIHJldHVybiByZWFkaW5nICYmIGZvY3VzZWQgJiYgcmVhZGluZy5jb250YWlucyhmb2N1c2VkKSA/IFwicmVhZGluZ1wiIDogXCJsaXN0XCI7XG59XG5cbmZ1bmN0aW9uIHJ1bkV4dGVuc2lvbihjb21tYW5kOiBFeHRlbnNpb25Db21tYW5kKTogdm9pZCB7XG4gIHN3aXRjaCAoY29tbWFuZCkge1xuICAgIGNhc2UgXCJoZWxwXCI6XG4gICAgICBzaG93SGVscChcbiAgICAgICAgQklORElOR1MubWFwKChiKSA9PiAoe1xuICAgICAgICAgIGtleXM6IGIua2V5cy5qb2luKFwiIFwiKSxcbiAgICAgICAgICBkZXNjcmlwdGlvbjogQUNUSU9OU1tiLmFjdGlvbl0uZGVzY3JpcHRpb24sXG4gICAgICAgIH0pKSxcbiAgICAgICk7XG4gICAgICByZXR1cm47XG4gICAgY2FzZSBcInBhbGV0dGUtZ29cIjpcbiAgICAgIG9wZW5Gb2xkZXJQYWxldHRlKCk7XG4gICAgICByZXR1cm47XG4gIH1cbn1cblxuLyoqIFJ1biBhbiBhY3Rpb24uIFJldHVybnMgZmFsc2Ugd2hlbiB0aGUga2V5IHNob3VsZCBiZSBsZWZ0IGZvciBPdXRsb29rIChwYXNzdGhyb3VnaCkuICovXG5leHBvcnQgZnVuY3Rpb24gZXhlY3V0ZShpZDogQWN0aW9uSWQpOiBib29sZWFuIHtcbiAgY29uc3QgeyBzdHJhdGVneSwgZGVzY3JpcHRpb24gfSA9IEFDVElPTlNbaWRdO1xuICBzd2l0Y2ggKHN0cmF0ZWd5LmtpbmQpIHtcbiAgICBjYXNlIFwicGFzc3Rocm91Z2hcIjpcbiAgICAgIHJldHVybiBmYWxzZTtcbiAgICBjYXNlIFwiZHJpdmVcIjoge1xuICAgICAgY29uc3QgdGFyZ2V0ID0gcmVzb2x2ZShzdHJhdGVneS50YXJnZXQpO1xuICAgICAgaWYgKHRhcmdldCBpbnN0YW5jZW9mIEhUTUxFbGVtZW50KSB0YXJnZXQuY2xpY2soKTtcbiAgICAgIGVsc2UgdG9hc3QoYCR7ZGVzY3JpcHRpb259OiBjb3VsZG4ndCBmaW5kIHRoZSBidXR0b25gKTtcbiAgICAgIHJldHVybiB0cnVlO1xuICAgIH1cbiAgICBjYXNlIFwidG9nZ2xlXCI6IHtcbiAgICAgIGNvbnN0IHJvd3MgPSBzZWxlY3RlZFJvd3MoKTtcbiAgICAgIGlmIChyb3dzLmxlbmd0aCA9PT0gMCkge1xuICAgICAgICB0b2FzdChgJHtkZXNjcmlwdGlvbn06IHNlbGVjdCBhIG1lc3NhZ2UgZmlyc3RgKTtcbiAgICAgICAgcmV0dXJuIHRydWU7XG4gICAgICB9XG4gICAgICBjb25zdCB3YW50VW5yZWFkID0gc3RyYXRlZ3kud2FudCA9PT0gXCJ1bnJlYWRcIjtcbiAgICAgIGlmIChyb3dzLmV2ZXJ5KChyb3cpID0+IGlzVW5yZWFkUm93KHJvdykgPT09IHdhbnRVbnJlYWQpKSByZXR1cm4gdHJ1ZTtcbiAgICAgIGNvbnN0IHRhcmdldCA9IHJlc29sdmUoc3RyYXRlZ3kudGFyZ2V0KTtcbiAgICAgIGlmICh0YXJnZXQgaW5zdGFuY2VvZiBIVE1MRWxlbWVudCkgdGFyZ2V0LmNsaWNrKCk7XG4gICAgICBlbHNlIHRvYXN0KGAke2Rlc2NyaXB0aW9ufTogY291bGRuJ3QgZmluZCB0aGUgYnV0dG9uYCk7XG4gICAgICByZXR1cm4gdHJ1ZTtcbiAgICB9XG4gICAgY2FzZSBcImV4dGVuc2lvblwiOlxuICAgICAgcnVuRXh0ZW5zaW9uKHN0cmF0ZWd5LmNvbW1hbmQpO1xuICAgICAgcmV0dXJuIHRydWU7XG4gICAgY2FzZSBcInRvZG9cIjpcbiAgICAgIHRvYXN0KGAke2Rlc2NyaXB0aW9ufSBpc24ndCBpbXBsZW1lbnRlZCB5ZXRgKTtcbiAgICAgIHJldHVybiB0cnVlO1xuICB9XG59XG4iLCAiLy8gVXNlciBwcmVmZXJlbmNlcy4gU3RvcmVkIGluIGNocm9tZS5zdG9yYWdlLnN5bmM7IG5vdGhpbmcgZWxzZSBpcyBldmVyIHBlcnNpc3RlZC5cbi8vXG4vLyBFYWNoIGVuYWJsZWQgZmVhdHVyZSBpZCBpcyB3cml0dGVuIGFzIGEgdG9rZW4gb24gPGh0bWwgZGF0YS1vdmVybG9vay1mZWF0dXJlcz1cIi4uLlwiPiwgYW5kIGVhY2hcbi8vIGNob2ljZSBhcyBpdHMgb3duIGF0dHJpYnV0ZSAoPGh0bWwgZGF0YS1vdmVybG9vay1wYWxldHRlPVwic29sYXJpemVkXCI+KSwgc28gc3R5bGVzaGVldHMgZ2F0ZVxuLy8gdGhlbXNlbHZlcyB3aXRoIGBbZGF0YS1vdmVybG9vay1mZWF0dXJlc349XCJkZW5zaXR5XCJdYCBvciBgW2RhdGEtb3Zlcmxvb2stcGFsZXR0ZT1cInNvbGFyaXplZFwiXWAuXG5cbmV4cG9ydCB0eXBlIEZlYXR1cmVJZCA9IGtleW9mIHR5cGVvZiBGRUFUVVJFUztcbmV4cG9ydCB0eXBlIENob2ljZUlkID0ga2V5b2YgdHlwZW9mIENIT0lDRVM7XG5leHBvcnQgdHlwZSBTZXR0aW5ncyA9IFJlY29yZDxGZWF0dXJlSWQsIGJvb2xlYW4+ICYgeyBbSyBpbiBDaG9pY2VJZF06IENob2ljZVZhbHVlPEs+IH07XG5leHBvcnQgdHlwZSBDaG9pY2VWYWx1ZTxLIGV4dGVuZHMgQ2hvaWNlSWQ+ID0ga2V5b2YgKHR5cGVvZiBDSE9JQ0VTKVtLXVtcIm9wdGlvbnNcIl07XG5cbnR5cGUgR3JvdXAgPSBcIkdlbmVyYWxcIiB8IFwiS2V5Ym9hcmRcIiB8IFwiRGVjbHV0dGVyXCIgfCBcIkxvb2tcIjtcblxuaW50ZXJmYWNlIEZlYXR1cmUge1xuICBsYWJlbDogc3RyaW5nO1xuICBncm91cDogR3JvdXA7XG4gIGRlZmF1bHQ6IGJvb2xlYW47XG59XG5cbmludGVyZmFjZSBDaG9pY2Uge1xuICBsYWJlbDogc3RyaW5nO1xuICBncm91cDogR3JvdXA7XG4gIG9wdGlvbnM6IFJlY29yZDxzdHJpbmcsIHN0cmluZz47XG4gIGRlZmF1bHQ6IHN0cmluZztcbn1cblxuZXhwb3J0IGNvbnN0IEZFQVRVUkVTID0ge1xuICBoZWFsdGhJbmRpY2F0b3I6IHtcbiAgICBsYWJlbDogXCJTaG93IGFuIGluZGljYXRvciB3aGVuIE91dGxvb2sncyBwYWdlIGNoYW5nZXMgYnJlYWsgT0JCXCIsXG4gICAgZ3JvdXA6IFwiR2VuZXJhbFwiLFxuICAgIGRlZmF1bHQ6IHRydWUsXG4gIH0sXG4gIGtleW1hcDogeyBsYWJlbDogXCJGYXN0bWFpbC1zdHlsZSBrZXlib2FyZCBzaG9ydGN1dHNcIiwgZ3JvdXA6IFwiS2V5Ym9hcmRcIiwgZGVmYXVsdDogdHJ1ZSB9LFxuICBcImhpZGUtdXBzZWxsXCI6IHsgbGFiZWw6IFwiVXBncmFkZSBhbmQgJ3RyeSB0aGUgbmV3XHUyMDI2JyBiYW5uZXJzXCIsIGdyb3VwOiBcIkRlY2x1dHRlclwiLCBkZWZhdWx0OiB0cnVlIH0sXG4gIFwiaGlkZS1hcHAtYmFubmVyc1wiOiB7XG4gICAgbGFiZWw6IFwiTm90aWNlcyBhY3Jvc3MgdGhlIHRvcCBvZiBPdXRsb29rIChwb3AtdXAsIHNpZ24taW4gYW5kIHN0b3JhZ2UgcmVtaW5kZXJzKVwiLFxuICAgIGdyb3VwOiBcIkRlY2x1dHRlclwiLFxuICAgIGRlZmF1bHQ6IGZhbHNlLFxuICB9LFxuICBcImhpZGUtaW5mby1iYXJzXCI6IHtcbiAgICBsYWJlbDogXCJOb3RpY2VzIGFib3ZlIGFuIG9wZW4gbWVzc2FnZSAoJ1lvdSByZXBsaWVkIG9uXHUyMDI2JywgJ0V4dHJhIGxpbmUgYnJlYWtzIHJlbW92ZWQnKVwiLFxuICAgIGdyb3VwOiBcIkRlY2x1dHRlclwiLFxuICAgIGRlZmF1bHQ6IGZhbHNlLFxuICB9LFxuICBcImhpZGUtY29waWxvdFwiOiB7IGxhYmVsOiBcIkNvcGlsb3QgYnV0dG9ucyBhbmQgcGFuZWxzXCIsIGdyb3VwOiBcIkRlY2x1dHRlclwiLCBkZWZhdWx0OiB0cnVlIH0sXG4gIFwiaGlkZS1oZWFkZXItZXh0cmFzXCI6IHtcbiAgICBsYWJlbDogXCJUZWFtcyBhbmQgTXkgRGF5IGJ1dHRvbnMgaW4gdGhlIGhlYWRlclwiLFxuICAgIGdyb3VwOiBcIkRlY2x1dHRlclwiLFxuICAgIGRlZmF1bHQ6IHRydWUsXG4gIH0sXG4gIFwiaGlkZS1yYWlsLWV4dHJhc1wiOiB7XG4gICAgbGFiZWw6IFwiVG8gRG8sIE5ld3NsZXR0ZXJzLCBPcmcgRXhwbG9yZXIgYW5kIE9uZURyaXZlIGluIHRoZSBhcHAgcmFpbFwiLFxuICAgIGdyb3VwOiBcIkRlY2x1dHRlclwiLFxuICAgIGRlZmF1bHQ6IHRydWUsXG4gIH0sXG4gIFwiaGlkZS1yaWJib24tZXh0cmFzXCI6IHtcbiAgICBsYWJlbDogXCJUZWFtcywgU3dlZXAsIFF1aWNrIHN0ZXBzLCBWaXZhIGFuZCBhZGQtaW4gYnV0dG9ucyBpbiB0aGUgcmliYm9uXCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogdHJ1ZSxcbiAgfSxcbiAgXCJoaWRlLWV4dGVybmFsLXRhZ1wiOiB7IGxhYmVsOiBcIidFeHRlcm5hbCcgc2VuZGVyIHRhZ3NcIiwgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsIGRlZmF1bHQ6IHRydWUgfSxcbiAgXCJoaWRlLXJvdy1hY3Rpb25zXCI6IHtcbiAgICBsYWJlbDogXCJGbGFnIGFuZCBwaW4gaWNvbnMgdGhhdCBhcHBlYXIgd2hlbiBob3ZlcmluZyBhIG1lc3NhZ2VcIixcbiAgICBncm91cDogXCJEZWNsdXR0ZXJcIixcbiAgICBkZWZhdWx0OiBmYWxzZSxcbiAgfSxcbiAgXCJoaWRlLWF2YXRhcnNcIjoge1xuICAgIGxhYmVsOiBcIlNlbmRlciBwaWN0dXJlcyBpbiB0aGUgbWVzc2FnZSBsaXN0XCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogZmFsc2UsXG4gIH0sXG4gIFwiaGlkZS1mb2N1c2VkLXRhYnNcIjogeyBsYWJlbDogXCJGb2N1c2VkIC8gT3RoZXIgdGFic1wiLCBncm91cDogXCJEZWNsdXR0ZXJcIiwgZGVmYXVsdDogZmFsc2UgfSxcbiAgXCJoaWRlLXJpYmJvbi10YWJzXCI6IHtcbiAgICBsYWJlbDogXCJSaWJib24gdGFicyAoSG9tZSwgVmlldywgSGVscClcIixcbiAgICBncm91cDogXCJEZWNsdXR0ZXJcIixcbiAgICBkZWZhdWx0OiBmYWxzZSxcbiAgfSxcbiAgXCJoaWRlLWFwcC1yYWlsXCI6IHsgbGFiZWw6IFwiQXBwIHJhaWwgKGxlZnQgaWNvbiBiYXIpXCIsIGdyb3VwOiBcIkRlY2x1dHRlclwiLCBkZWZhdWx0OiBmYWxzZSB9LFxuICBcImhpZGUtcmliYm9uXCI6IHsgbGFiZWw6IFwiUmliYm9uXCIsIGdyb3VwOiBcIkRlY2x1dHRlclwiLCBkZWZhdWx0OiBmYWxzZSB9LFxuICBkZW5zaXR5OiB7IGxhYmVsOiBcIlRpZ2h0ZXIsIGRlbnNlciBsYXlvdXRcIiwgZ3JvdXA6IFwiTG9va1wiLCBkZWZhdWx0OiB0cnVlIH0sXG59IGFzIGNvbnN0IHNhdGlzZmllcyBSZWNvcmQ8c3RyaW5nLCBGZWF0dXJlPjtcblxuZXhwb3J0IGNvbnN0IENIT0lDRVMgPSB7XG4gIHBhbGV0dGU6IHtcbiAgICBsYWJlbDogXCJDb2xvcnNcIixcbiAgICBncm91cDogXCJMb29rXCIsXG4gICAgb3B0aW9uczoge1xuICAgICAgb3V0bG9vazogXCJPdXRsb29rJ3Mgb3duXCIsXG4gICAgICBzb2xhcml6ZWQ6IFwiU29sYXJpemVkXCIsXG4gICAgICBiYXNlMTY6IFwiQmFzZTE2IERlZmF1bHRcIixcbiAgICAgIGZhc3RtYWlsOiBcIkZhc3RtYWlsXCIsXG4gICAgICBwYXBlcjogXCJQYXBlclwiLFxuICAgIH0sXG4gICAgZGVmYXVsdDogXCJvdXRsb29rXCIsXG4gIH0sXG4gIG1vZGU6IHtcbiAgICBsYWJlbDogXCJMaWdodCBvciBkYXJrXCIsXG4gICAgZ3JvdXA6IFwiTG9va1wiLFxuICAgIG9wdGlvbnM6IHsgYXV0bzogXCJGb2xsb3cgT3V0bG9va1wiLCBsaWdodDogXCJMaWdodFwiLCBkYXJrOiBcIkRhcmtcIiB9LFxuICAgIGRlZmF1bHQ6IFwiYXV0b1wiLFxuICB9LFxuICBsYXlvdXQ6IHtcbiAgICBsYWJlbDogXCJMYXlvdXRcIixcbiAgICBncm91cDogXCJMb29rXCIsXG4gICAgb3B0aW9uczoge1xuICAgICAgb3V0bG9vazogXCJPdXRsb29rJ3Mgb3duXCIsXG4gICAgICBjYWxtOiBcIkNhbG0gdGhyZWUtcGFuZVwiLFxuICAgICAgZmFzdG1haWw6IFwiRmFzdG1haWwgKGxpc3QgZmlyc3QpXCIsXG4gICAgICBmb2N1czogXCJGb2N1cyAoc2luZ2xlIGNvbHVtbilcIixcbiAgICB9LFxuICAgIGRlZmF1bHQ6IFwib3V0bG9va1wiLFxuICB9LFxufSBhcyBjb25zdCBzYXRpc2ZpZXMgUmVjb3JkPHN0cmluZywgQ2hvaWNlPjtcblxuZXhwb3J0IGNvbnN0IEZFQVRVUkVfSURTID0gT2JqZWN0LmtleXMoRkVBVFVSRVMpIGFzIEZlYXR1cmVJZFtdO1xuZXhwb3J0IGNvbnN0IENIT0lDRV9JRFMgPSBPYmplY3Qua2V5cyhDSE9JQ0VTKSBhcyBDaG9pY2VJZFtdO1xuXG5leHBvcnQgZnVuY3Rpb24gZGVmYXVsdFNldHRpbmdzKCk6IFNldHRpbmdzIHtcbiAgcmV0dXJuIHtcbiAgICAuLi5PYmplY3QuZnJvbUVudHJpZXMoRkVBVFVSRV9JRFMubWFwKChpZCkgPT4gW2lkLCBGRUFUVVJFU1tpZF0uZGVmYXVsdF0pKSxcbiAgICAuLi5PYmplY3QuZnJvbUVudHJpZXMoQ0hPSUNFX0lEUy5tYXAoKGlkKSA9PiBbaWQsIENIT0lDRVNbaWRdLmRlZmF1bHRdKSksXG4gIH0gYXMgU2V0dGluZ3M7XG59XG5cbi8qKiBNZXJnZSB3aGF0ZXZlciBpcyBzdG9yZWQgb3ZlciB0aGUgZGVmYXVsdHMsIGlnbm9yaW5nIHVua25vd24gb3IgbWlzdHlwZWQga2V5cy4gKi9cbmV4cG9ydCBmdW5jdGlvbiBub3JtYWxpemUoc3RvcmVkOiBSZWNvcmQ8c3RyaW5nLCB1bmtub3duPik6IFNldHRpbmdzIHtcbiAgY29uc3Qgc2V0dGluZ3MgPSBkZWZhdWx0U2V0dGluZ3MoKSBhcyBSZWNvcmQ8c3RyaW5nLCB1bmtub3duPjtcbiAgZm9yIChjb25zdCBpZCBvZiBGRUFUVVJFX0lEUykge1xuICAgIGlmICh0eXBlb2Ygc3RvcmVkW2lkXSA9PT0gXCJib29sZWFuXCIpIHNldHRpbmdzW2lkXSA9IHN0b3JlZFtpZF07XG4gIH1cbiAgZm9yIChjb25zdCBpZCBvZiBDSE9JQ0VfSURTKSB7XG4gICAgY29uc3QgdmFsdWUgPSBzdG9yZWRbaWRdO1xuICAgIGlmICh0eXBlb2YgdmFsdWUgPT09IFwic3RyaW5nXCIgJiYgdmFsdWUgaW4gQ0hPSUNFU1tpZF0ub3B0aW9ucykgc2V0dGluZ3NbaWRdID0gdmFsdWU7XG4gIH1cbiAgcmV0dXJuIHNldHRpbmdzIGFzIFNldHRpbmdzO1xufVxuXG5leHBvcnQgZnVuY3Rpb24gZmVhdHVyZVRva2VucyhzZXR0aW5nczogU2V0dGluZ3MpOiBzdHJpbmcge1xuICByZXR1cm4gRkVBVFVSRV9JRFMuZmlsdGVyKChpZCkgPT4gc2V0dGluZ3NbaWRdKS5qb2luKFwiIFwiKTtcbn1cblxuLyoqIFdyaXRlcyB0aGUgc2V0dGluZ3Mgb250byA8aHRtbD4gYXMgdGhlIGF0dHJpYnV0ZXMgdGhlIHN0eWxlc2hlZXRzIGtleSBvZmYuICovXG5leHBvcnQgZnVuY3Rpb24gYXBwbHlUb0RvY3VtZW50KHNldHRpbmdzOiBTZXR0aW5ncywgcm9vdDogSFRNTEVsZW1lbnQpOiB2b2lkIHtcbiAgcm9vdC5zZXRBdHRyaWJ1dGUoXCJkYXRhLW92ZXJsb29rLWZlYXR1cmVzXCIsIGZlYXR1cmVUb2tlbnMoc2V0dGluZ3MpKTtcbiAgZm9yIChjb25zdCBpZCBvZiBDSE9JQ0VfSURTKSByb290LnNldEF0dHJpYnV0ZShgZGF0YS1vdmVybG9vay0ke2lkfWAsIHNldHRpbmdzW2lkXSk7XG59XG5cbmV4cG9ydCBhc3luYyBmdW5jdGlvbiBsb2FkU2V0dGluZ3MoKTogUHJvbWlzZTxTZXR0aW5ncz4ge1xuICByZXR1cm4gbm9ybWFsaXplKGF3YWl0IGNocm9tZS5zdG9yYWdlLnN5bmMuZ2V0KG51bGwpKTtcbn1cblxuZXhwb3J0IGFzeW5jIGZ1bmN0aW9uIHNhdmVTZXR0aW5nKFxuICBpZDogRmVhdHVyZUlkIHwgQ2hvaWNlSWQsXG4gIHZhbHVlOiBib29sZWFuIHwgc3RyaW5nLFxuKTogUHJvbWlzZTx2b2lkPiB7XG4gIGF3YWl0IGNocm9tZS5zdG9yYWdlLnN5bmMuc2V0KHsgW2lkXTogdmFsdWUgfSk7XG59XG5cbmV4cG9ydCBmdW5jdGlvbiBvblNldHRpbmdzQ2hhbmdlZChjYWxsYmFjazogKHNldHRpbmdzOiBTZXR0aW5ncykgPT4gdm9pZCk6IHZvaWQge1xuICBjaHJvbWUuc3RvcmFnZS5vbkNoYW5nZWQuYWRkTGlzdGVuZXIoKF9jaGFuZ2VzLCBhcmVhKSA9PiB7XG4gICAgaWYgKGFyZWEgPT09IFwic3luY1wiKSB2b2lkIGxvYWRTZXR0aW5ncygpLnRoZW4oY2FsbGJhY2spO1xuICB9KTtcbn1cbiIsICIvLyBFbnRyeSBwb2ludDogYXBwbGllcyBzZXR0aW5ncywgdGFncyBlbGVtZW50cyBmb3IgQ1NTLCBydW5zIHRoZSBoZWFsdGggY2hlY2ssIG1vdW50cyB0aGUga2V5bWFwLlxuXG5pbXBvcnQgeyBDQVBUVVJFX01FU1NBR0UsIGNhcHR1cmVQYWdlIH0gZnJvbSBcIi4vY2FwdHVyZS50c1wiO1xuaW1wb3J0IHsgQUNUSU9OUyB9IGZyb20gXCIuL2tleW1hcC9hY3Rpb25zLnRzXCI7XG5pbXBvcnQgeyBCSU5ESU5HUyB9IGZyb20gXCIuL2tleW1hcC9iaW5kaW5ncy50c1wiO1xuaW1wb3J0IHsgREVGQVVMVF9USU1FT1VUX01TLCBrZXlGcm9tRXZlbnQsIFNlcXVlbmNlRW5naW5lIH0gZnJvbSBcIi4va2V5bWFwL2VuZ2luZS50c1wiO1xuaW1wb3J0IHsgY3VycmVudENvbnRleHQsIGV4ZWN1dGUgfSBmcm9tIFwiLi9rZXltYXAvZXhlY3V0ZS50c1wiO1xuaW1wb3J0IHtcbiAgYnJva2VuRW50cmllcyxcbiAgdHlwZSBIZWFsdGhSZXN1bHQsXG4gIGhlYWx0aENoZWNrLFxuICBvdXRsb29rQ29sb3JTY2hlbWUsXG4gIHJlc29sdmUsXG4gIHRhZ0VsZW1lbnRzLFxufSBmcm9tIFwiLi9zZWxlY3RvcnMudHNcIjtcbmltcG9ydCB7IGFwcGx5VG9Eb2N1bWVudCwgbG9hZFNldHRpbmdzLCBvblNldHRpbmdzQ2hhbmdlZCwgdHlwZSBTZXR0aW5ncyB9IGZyb20gXCIuL3NldHRpbmdzLnRzXCI7XG5pbXBvcnQgeyBpc0Zyb21PdmVybGF5LCBpc01vZGFsT3Blbiwgbm90aWNlLCBzaG93SGVhbHRoQmFkZ2UsIHNob3dIaW50IH0gZnJvbSBcIi4vdWkvb3ZlcmxheS50c1wiO1xuXG5jb25zdCBMT0cgPSBcIltvYmJdXCI7XG5jb25zdCBBUFBfUkVBRFlfVElNRU9VVF9NUyA9IDIwXzAwMDtcbmNvbnN0IEhFQUxUSF9TRVRUTEVfTVMgPSAxXzAwMDtcbi8qKiBXYWl0cyBiZXR3ZWVuIHJlLWNoZWNrcyBvZiBhIGZhaWxlZCBoZWFsdGggY2hlY2sgYmVmb3JlIGl0IGlzIHJlcG9ydGVkLiAqL1xuY29uc3QgSEVBTFRIX1JFVFJZX01TID0gWzNfMDAwLCAxMF8wMDBdO1xuXG5sZXQgc2V0dGluZ3M6IFNldHRpbmdzO1xubGV0IGhlYWx0aERpc21pc3NlZCA9IGZhbHNlO1xuXG5mdW5jdGlvbiBhcHBseVNldHRpbmdzKG5leHQ6IFNldHRpbmdzKTogdm9pZCB7XG4gIHNldHRpbmdzID0gbmV4dDtcbiAgYXBwbHlUb0RvY3VtZW50KHNldHRpbmdzLCBkb2N1bWVudC5kb2N1bWVudEVsZW1lbnQpO1xuICBzeW5jQ29sb3JTY2hlbWUoKTtcbiAgaWYgKCFzZXR0aW5ncy5oZWFsdGhJbmRpY2F0b3IpIHNob3dIZWFsdGhCYWRnZShbXSwgKCkgPT4ge30pO1xufVxuXG4vKiogUmVjb3JkcyB3aGV0aGVyIE91dGxvb2sgc2hvd3MgYSByZWFkaW5nIHBhbmUsIHNvIGxheW91dHMgb25seSByZXNoYXBlIHBhbmVzIHdoZW4gaXQgZG9lc24ndC4gKi9cbmZ1bmN0aW9uIHN5bmNSZWFkaW5nUGFuZSgpOiB2b2lkIHtcbiAgY29uc3Qgc3RhdGUgPSByZXNvbHZlKFwicmVhZGluZ1BhbmVcIikgPyBcInNob3duXCIgOiBcIm5vbmVcIjtcbiAgaWYgKGRvY3VtZW50LmRvY3VtZW50RWxlbWVudC5kYXRhc2V0Lm92ZXJsb29rUmVhZGluZ1BhbmUgIT09IHN0YXRlKSB7XG4gICAgZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50LmRhdGFzZXQub3Zlcmxvb2tSZWFkaW5nUGFuZSA9IHN0YXRlO1xuICB9XG59XG5cbi8qKiBNaXJyb3JzIE91dGxvb2sncyBvd24gbGlnaHQvZGFyayBtb2RlIG9udG8gPGh0bWw+IHNvIFwiRm9sbG93IE91dGxvb2tcIiBwYWxldHRlcyBjYW4gbWF0Y2ggaXQuICovXG5mdW5jdGlvbiBzeW5jQ29sb3JTY2hlbWUoKTogdm9pZCB7XG4gIGNvbnN0IHNjaGVtZSA9IG91dGxvb2tDb2xvclNjaGVtZShkb2N1bWVudC5kb2N1bWVudEVsZW1lbnQpO1xuICBpZiAoZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50LmdldEF0dHJpYnV0ZShcImRhdGEtb3Zlcmxvb2stb3V0bG9vay1zY2hlbWVcIikgIT09IHNjaGVtZSkge1xuICAgIGRvY3VtZW50LmRvY3VtZW50RWxlbWVudC5zZXRBdHRyaWJ1dGUoXCJkYXRhLW92ZXJsb29rLW91dGxvb2stc2NoZW1lXCIsIHNjaGVtZSk7XG4gIH1cbn1cblxuZnVuY3Rpb24gcmVwb3J0SGVhbHRoKHJlc3VsdHM6IEhlYWx0aFJlc3VsdFtdKTogdm9pZCB7XG4gIGNvbnN0IGJyb2tlbiA9IGJyb2tlbkVudHJpZXMocmVzdWx0cyk7XG4gIGNvbnN0IGZhbGxiYWNrcyA9IHJlc3VsdHMuZmlsdGVyKChyKSA9PiByLmNhbmRpZGF0ZSA+IDApLm1hcCgocikgPT4gci5uYW1lKTtcbiAgZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50LmRhdGFzZXQub3Zlcmxvb2tIZWFsdGggPSBicm9rZW4ubGVuZ3RoID8gXCJicm9rZW5cIiA6IFwib2tcIjtcbiAgY29uc29sZS5ncm91cENvbGxhcHNlZChgJHtMT0d9IGhlYWx0aDogJHticm9rZW4ubGVuZ3RoID8gYCR7YnJva2VuLmxlbmd0aH0gYnJva2VuYCA6IFwib2tcIn1gKTtcbiAgY29uc29sZS50YWJsZShyZXN1bHRzKTtcbiAgY29uc29sZS5ncm91cEVuZCgpO1xuICBpZiAoYnJva2VuLmxlbmd0aCkgY29uc29sZS53YXJuKGAke0xPR30gc2VsZWN0b3JzIG5vdCBmb3VuZDpgLCBicm9rZW4uam9pbihcIiwgXCIpKTtcbiAgaWYgKGZhbGxiYWNrcy5sZW5ndGgpIGNvbnNvbGUuaW5mbyhgJHtMT0d9IHVzaW5nIGZhbGxiYWNrIHNlbGVjdG9ycyBmb3I6YCwgZmFsbGJhY2tzLmpvaW4oXCIsIFwiKSk7XG4gIGlmIChzZXR0aW5ncy5oZWFsdGhJbmRpY2F0b3IgJiYgIWhlYWx0aERpc21pc3NlZCkge1xuICAgIHNob3dIZWFsdGhCYWRnZShicm9rZW4sICgpID0+IHtcbiAgICAgIGhlYWx0aERpc21pc3NlZCA9IHRydWU7XG4gICAgfSk7XG4gIH1cbn1cblxubGV0IGhlYWx0aFRpbWVyOiBSZXR1cm5UeXBlPHR5cGVvZiBzZXRUaW1lb3V0PiB8IHVuZGVmaW5lZDtcblxuLyoqXG4gKiBDaGVja3MgdGhlIHJlZ2lzdHJ5IG9uY2UgT3V0bG9vayBoYXMgc2V0dGxlZC4gUGFydHMgb2YgdGhlIHBhZ2UgKHRoZSBhcHAgcmFpbCwgZm9yIG9uZSkgcmVuZGVyXG4gKiBhZnRlciB0aGUgbWVzc2FnZSBsaXN0LCBzbyBhIG1pc3MgaXMgcmUtY2hlY2tlZCBhIGZldyB0aW1lcyBiZWZvcmUgaXQgaXMgcmVwb3J0ZWQuXG4gKi9cbmZ1bmN0aW9uIHNjaGVkdWxlSGVhbHRoQ2hlY2soKTogdm9pZCB7XG4gIGNsZWFyVGltZW91dChoZWFsdGhUaW1lcik7XG4gIC8vIE9ubHkgdGhlIG1haWwgdmlld3MgYXJlIHN1cnZleWVkOyBDYWxlbmRhciBhbmQgUGVvcGxlIGxlZ2l0aW1hdGVseSBsYWNrIG1haWwgZWxlbWVudHMuXG4gIGlmICghbG9jYXRpb24ucGF0aG5hbWUuaW5jbHVkZXMoXCIvbWFpbFwiKSkgcmV0dXJuO1xuICBjb25zdCBhdHRlbXB0ID0gKG46IG51bWJlcikgPT4ge1xuICAgIGNvbnN0IHJlc3VsdHMgPSBoZWFsdGhDaGVjaygpO1xuICAgIGNvbnN0IHJldHJ5SW4gPSBIRUFMVEhfUkVUUllfTVNbbl07XG4gICAgaWYgKGJyb2tlbkVudHJpZXMocmVzdWx0cykubGVuZ3RoID09PSAwIHx8IHJldHJ5SW4gPT09IHVuZGVmaW5lZCkgcmVwb3J0SGVhbHRoKHJlc3VsdHMpO1xuICAgIGVsc2UgaGVhbHRoVGltZXIgPSBzZXRUaW1lb3V0KCgpID0+IGF0dGVtcHQobiArIDEpLCByZXRyeUluKTtcbiAgfTtcbiAgaGVhbHRoVGltZXIgPSBzZXRUaW1lb3V0KCgpID0+IGF0dGVtcHQoMCksIEhFQUxUSF9TRVRUTEVfTVMpO1xufVxuXG4vKiogUmVzb2x2ZXMgb25jZSBPdXRsb29rJ3MgbWFpbCBVSSBoYXMgcmVuZGVyZWQsIG9yIGFmdGVyIGEgdGltZW91dCBzbyB0aGUgY2hlY2sgc3RpbGwgcmVwb3J0cy4gKi9cbmZ1bmN0aW9uIHdoZW5BcHBSZWFkeSgpOiBQcm9taXNlPHZvaWQ+IHtcbiAgcmV0dXJuIG5ldyBQcm9taXNlKChkb25lKSA9PiB7XG4gICAgaWYgKHJlc29sdmUoXCJtZXNzYWdlTGlzdFwiKSkgcmV0dXJuIGRvbmUoKTtcbiAgICBjb25zdCBvYnNlcnZlciA9IG5ldyBNdXRhdGlvbk9ic2VydmVyKCgpID0+IHtcbiAgICAgIGlmIChyZXNvbHZlKFwibWVzc2FnZUxpc3RcIikpIGZpbmlzaCgpO1xuICAgIH0pO1xuICAgIGNvbnN0IHRpbWVyID0gc2V0VGltZW91dChmaW5pc2gsIEFQUF9SRUFEWV9USU1FT1VUX01TKTtcbiAgICBmdW5jdGlvbiBmaW5pc2goKSB7XG4gICAgICBvYnNlcnZlci5kaXNjb25uZWN0KCk7XG4gICAgICBjbGVhclRpbWVvdXQodGltZXIpO1xuICAgICAgZG9uZSgpO1xuICAgIH1cbiAgICBvYnNlcnZlci5vYnNlcnZlKGRvY3VtZW50LmJvZHksIHsgY2hpbGRMaXN0OiB0cnVlLCBzdWJ0cmVlOiB0cnVlIH0pO1xuICB9KTtcbn1cblxuLyoqIFJlLXRhZyBvbiBET00gY2h1cm4gKFJlYWN0IHJlLXJlbmRlcnMpLCBhbmQgcmUtY2hlY2sgaGVhbHRoIHdoZW4gdGhlIFNQQSBuYXZpZ2F0ZXMuICovXG5mdW5jdGlvbiB3YXRjaFBhZ2UoKTogdm9pZCB7XG4gIGxldCBzY2hlZHVsZWQgPSBmYWxzZTtcbiAgbGV0IGxhc3RVcmwgPSBsb2NhdGlvbi5ocmVmO1xuICBuZXcgTXV0YXRpb25PYnNlcnZlcigoKSA9PiB7XG4gICAgaWYgKHNjaGVkdWxlZCkgcmV0dXJuO1xuICAgIHNjaGVkdWxlZCA9IHRydWU7XG4gICAgcmVxdWVzdEFuaW1hdGlvbkZyYW1lKCgpID0+IHtcbiAgICAgIHNjaGVkdWxlZCA9IGZhbHNlO1xuICAgICAgdGFnRWxlbWVudHMoKTtcbiAgICAgIHN5bmNDb2xvclNjaGVtZSgpO1xuICAgICAgc3luY1JlYWRpbmdQYW5lKCk7XG4gICAgICBpZiAobG9jYXRpb24uaHJlZiAhPT0gbGFzdFVybCkge1xuICAgICAgICBsYXN0VXJsID0gbG9jYXRpb24uaHJlZjtcbiAgICAgICAgc2NoZWR1bGVIZWFsdGhDaGVjaygpO1xuICAgICAgfVxuICAgIH0pO1xuICB9KS5vYnNlcnZlKGRvY3VtZW50LmJvZHksIHsgY2hpbGRMaXN0OiB0cnVlLCBzdWJ0cmVlOiB0cnVlIH0pO1xufVxuXG5sZXQgb3JwaGFuTm90aWNlU2hvd24gPSBmYWxzZTtcblxuLyoqXG4gKiBBZnRlciBPQkIgaXMgcmVsb2FkZWQgb3IgdXBkYXRlZCwgdGhpcyBjb3B5IG9mIHRoZSBjb250ZW50IHNjcmlwdCBrZWVwcyBydW5uaW5nIGluIHRhYnMgdGhhdFxuICogd2VyZSBhbHJlYWR5IG9wZW4sIGN1dCBvZmYgZnJvbSB0aGUgZXh0ZW5zaW9uLiBTYXkgc28gb25jZSBpbnN0ZWFkIG9mIGhhbGYtd29ya2luZy5cbiAqL1xuZnVuY3Rpb24gaXNPcnBoYW5lZCgpOiBib29sZWFuIHtcbiAgaWYgKGNocm9tZS5ydW50aW1lPy5pZCkgcmV0dXJuIGZhbHNlO1xuICBpZiAoIW9ycGhhbk5vdGljZVNob3duKSB7XG4gICAgb3JwaGFuTm90aWNlU2hvd24gPSB0cnVlO1xuICAgIG5vdGljZShcIk9CQiB3YXMgdXBkYXRlZC5cIiwgXCJSZWxvYWQgdGFiXCIsICgpID0+IGxvY2F0aW9uLnJlbG9hZCgpKTtcbiAgfVxuICByZXR1cm4gdHJ1ZTtcbn1cblxuZnVuY3Rpb24gaXNFZGl0YWJsZSh0YXJnZXQ6IEV2ZW50VGFyZ2V0IHwgdW5kZWZpbmVkKTogYm9vbGVhbiB7XG4gIGlmICghKHRhcmdldCBpbnN0YW5jZW9mIEhUTUxFbGVtZW50KSkgcmV0dXJuIGZhbHNlO1xuICByZXR1cm4gdGFyZ2V0LmlzQ29udGVudEVkaXRhYmxlIHx8IC9eKElOUFVUfFRFWFRBUkVBfFNFTEVDVCkkLy50ZXN0KHRhcmdldC50YWdOYW1lKTtcbn1cblxuZnVuY3Rpb24gbW91bnRLZXltYXAoKTogdm9pZCB7XG4gIGNvbnN0IGVuZ2luZSA9IG5ldyBTZXF1ZW5jZUVuZ2luZShCSU5ESU5HUyk7XG4gIGxldCBoaW50VGltZXI6IFJldHVyblR5cGU8dHlwZW9mIHNldFRpbWVvdXQ+IHwgdW5kZWZpbmVkO1xuICB3aW5kb3cuYWRkRXZlbnRMaXN0ZW5lcihcbiAgICBcImtleWRvd25cIixcbiAgICAoZXZlbnQpID0+IHtcbiAgICAgIGlmICghc2V0dGluZ3Mua2V5bWFwIHx8ICFldmVudC5pc1RydXN0ZWQgfHwgaXNGcm9tT3ZlcmxheShldmVudCkgfHwgaXNNb2RhbE9wZW4oKSkgcmV0dXJuO1xuICAgICAgaWYgKGlzT3JwaGFuZWQoKSkgcmV0dXJuO1xuICAgICAgaWYgKGlzRWRpdGFibGUoZXZlbnQuY29tcG9zZWRQYXRoKClbMF0pKSByZXR1cm47XG4gICAgICBjb25zdCBrZXkgPSBrZXlGcm9tRXZlbnQoZXZlbnQpO1xuICAgICAgaWYgKGtleSA9PT0gbnVsbCkgcmV0dXJuO1xuXG4gICAgICBjb25zdCByZXN1bHQgPSBlbmdpbmUuZmVlZChrZXksIGN1cnJlbnRDb250ZXh0KCksIGV2ZW50LnRpbWVTdGFtcCk7XG4gICAgICBpZiAocmVzdWx0LmtpbmQgPT09IFwibm9uZVwiKSByZXR1cm47XG4gICAgICBzaG93SGludChcbiAgICAgICAgcmVzdWx0LmtpbmQgPT09IFwicGVuZGluZ1wiXG4gICAgICAgICAgPyBgJHtyZXN1bHQucHJlZml4LmpvaW4oXCIgXCIpfSBcdTIwMjYgJHtyZXN1bHQubmV4dFxuICAgICAgICAgICAgICAubWFwKChiKSA9PiBgJHtiLmtleXMuYXQoLTEpfSAke0FDVElPTlNbYi5hY3Rpb25dLmRlc2NyaXB0aW9ufWApXG4gICAgICAgICAgICAgIC5qb2luKFwiIFx1MDBCNyBcIil9YFxuICAgICAgICAgIDogbnVsbCxcbiAgICAgICk7XG4gICAgICBjbGVhclRpbWVvdXQoaGludFRpbWVyKTtcbiAgICAgIGlmIChyZXN1bHQua2luZCA9PT0gXCJwZW5kaW5nXCIpXG4gICAgICAgIGhpbnRUaW1lciA9IHNldFRpbWVvdXQoKCkgPT4gc2hvd0hpbnQobnVsbCksIERFRkFVTFRfVElNRU9VVF9NUyk7XG4gICAgICBpZiAocmVzdWx0LmtpbmQgPT09IFwibWF0Y2hcIiAmJiAhZXhlY3V0ZShyZXN1bHQuYWN0aW9uKSkgcmV0dXJuO1xuICAgICAgZXZlbnQucHJldmVudERlZmF1bHQoKTtcbiAgICAgIGV2ZW50LnN0b3BJbW1lZGlhdGVQcm9wYWdhdGlvbigpO1xuICAgIH0sXG4gICAgdHJ1ZSxcbiAgKTtcbn1cblxuLyoqIERldmVsb3BlciB0b29sczogdGhlIHBvcHVwIGFza3MgdGhpcyB0YWIgdG8gc2F2ZSBhIGNhcHR1cmUgb2YgdGhlIHBhZ2UgKHNyYy9jYXB0dXJlLnRzKS4gKi9cbmZ1bmN0aW9uIHNlcnZlQ2FwdHVyZVJlcXVlc3RzKCk6IHZvaWQge1xuICBjaHJvbWUucnVudGltZS5vbk1lc3NhZ2UuYWRkTGlzdGVuZXIoKG1lc3NhZ2UsIF9zZW5kZXIsIHJlcGx5KSA9PiB7XG4gICAgaWYgKG1lc3NhZ2U/LnR5cGUgIT09IENBUFRVUkVfTUVTU0FHRSkgcmV0dXJuO1xuICAgIGNvbnN0IHsgaHRtbCwgbmFtZSB9ID0gY2FwdHVyZVBhZ2UoZG9jdW1lbnQsIGxvY2F0aW9uLCBjaHJvbWUucnVudGltZS5nZXRNYW5pZmVzdCgpLnZlcnNpb24pO1xuICAgIGNvbnN0IHVybCA9IFVSTC5jcmVhdGVPYmplY3RVUkwobmV3IEJsb2IoW2h0bWxdLCB7IHR5cGU6IFwidGV4dC9odG1sXCIgfSkpO1xuICAgIGNvbnN0IGxpbmsgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KFwiYVwiKTtcbiAgICBsaW5rLmhyZWYgPSB1cmw7XG4gICAgbGluay5kb3dubG9hZCA9IG5hbWU7XG4gICAgZG9jdW1lbnQuYm9keS5hcHBlbmQobGluayk7XG4gICAgbGluay5jbGljaygpO1xuICAgIGxpbmsucmVtb3ZlKCk7XG4gICAgc2V0VGltZW91dCgoKSA9PiBVUkwucmV2b2tlT2JqZWN0VVJMKHVybCksIDEwXzAwMCk7XG4gICAgcmVwbHkoeyBvazogdHJ1ZSwgbmFtZSwgYnl0ZXM6IGh0bWwubGVuZ3RoIH0pO1xuICB9KTtcbn1cblxuYXN5bmMgZnVuY3Rpb24gbWFpbigpOiBQcm9taXNlPHZvaWQ+IHtcbiAgc2VydmVDYXB0dXJlUmVxdWVzdHMoKTtcbiAgYXBwbHlTZXR0aW5ncyhhd2FpdCBsb2FkU2V0dGluZ3MoKSk7XG4gIG9uU2V0dGluZ3NDaGFuZ2VkKGFwcGx5U2V0dGluZ3MpO1xuICBtb3VudEtleW1hcCgpO1xuICAvLyBDb21pbmcgYmFjayB0byB0aGlzIHRhYiBhZnRlciByZWxvYWRpbmcgT0JCIGZyb20gaXRzIHBvcHVwOiBvZmZlciB0aGUgcmVmcmVzaCByaWdodCBhd2F5LlxuICBkb2N1bWVudC5hZGRFdmVudExpc3RlbmVyKFwidmlzaWJpbGl0eWNoYW5nZVwiLCAoKSA9PiB7XG4gICAgaWYgKCFkb2N1bWVudC5oaWRkZW4pIGlzT3JwaGFuZWQoKTtcbiAgfSk7XG4gIGF3YWl0IHdoZW5BcHBSZWFkeSgpO1xuICB0YWdFbGVtZW50cygpO1xuICBzeW5jUmVhZGluZ1BhbmUoKTtcbiAgc2NoZWR1bGVIZWFsdGhDaGVjaygpO1xuICB3YXRjaFBhZ2UoKTtcbn1cblxudm9pZCBtYWluKCk7XG4iXSwKICAibWFwcGluZ3MiOiAiOzs7QUF5Qk8sTUFBTSxZQUFZO0FBQUE7QUFBQSxJQUV2QixVQUFVLEVBQUUsWUFBWSxDQUFDLGNBQWMsR0FBRyxRQUFRLFVBQVUsS0FBSyxZQUFZO0FBQUEsSUFDN0UsV0FBVyxFQUFFLFlBQVksQ0FBQyxtQkFBbUIsdUJBQXVCLEdBQUcsUUFBUSxTQUFTO0FBQUEsSUFDeEYsU0FBUztBQUFBLE1BQ1AsWUFBWSxDQUFDLGFBQWEsb0RBQW9EO0FBQUEsTUFDOUUsUUFBUTtBQUFBLE1BQ1IsS0FBSztBQUFBLElBQ1A7QUFBQSxJQUNBLFFBQVE7QUFBQSxNQUNOLFlBQVksQ0FBQywrQkFBK0Isc0NBQXNDO0FBQUEsTUFDbEYsUUFBUTtBQUFBLE1BQ1IsS0FBSztBQUFBLElBQ1A7QUFBQSxJQUNBLFlBQVk7QUFBQSxNQUNWLFlBQVksQ0FBQyw0Q0FBNEM7QUFBQSxNQUN6RCxRQUFRO0FBQUEsTUFDUixRQUFRO0FBQUEsTUFDUixLQUFLO0FBQUEsSUFDUDtBQUFBLElBQ0EsWUFBWTtBQUFBLE1BQ1YsWUFBWTtBQUFBLFFBQ1Y7QUFBQSxRQUNBO0FBQUEsTUFDRjtBQUFBLE1BQ0EsUUFBUTtBQUFBLE1BQ1IsS0FBSztBQUFBLElBQ1A7QUFBQSxJQUNBLFlBQVk7QUFBQSxNQUNWLFlBQVksQ0FBQyx1Q0FBdUMsbUJBQW1CO0FBQUEsTUFDdkUsUUFBUTtBQUFBLE1BQ1IsUUFBUTtBQUFBLE1BQ1IsS0FBSztBQUFBLE1BQ0wsS0FBSztBQUFBLElBQ1A7QUFB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
