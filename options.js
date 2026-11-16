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

  // src/capture.ts
  var CAPTURE_MESSAGE = "obb-capture";

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
    const settings = defaultSettings();
    for (const id of FEATURE_IDS) {
      if (typeof stored[id] === "boolean") settings[id] = stored[id];
    }
    for (const id of CHOICE_IDS) {
      const value = stored[id];
      if (typeof value === "string" && value in CHOICES[id].options) settings[id] = value;
    }
    return settings;
  }
  async function loadSettings() {
    return normalize(await chrome.storage.sync.get(null));
  }
  async function saveSetting(id, value) {
    await chrome.storage.sync.set({ [id]: value });
  }
  function onSettingsChanged(callback) {
    chrome.storage.onChanged.addListener((_changes, area) => {
      if (area === "sync") void loadSettings().then(callback);
    });
  }

  // src/options/options.ts
  var SHORT_LAYOUT_NAMES = {
    outlook: "Outlook",
    calm: "Calm",
    fastmail: "Fastmail",
    focus: "Focus"
  };
  var SHORT_PALETTE_NAMES = {
    outlook: "Outlook",
    solarized: "Solarized",
    base16: "Base16",
    fastmail: "Fastmail",
    paper: "Paper"
  };
  function rows(x, y, width) {
    return [0, 1, 2, 3, 4].map((i) => {
      const top = y + i * 6.5;
      const subject = i === 0 ? "hi" : "row";
      const subjectWidth = i === 0 ? width * 0.55 : width;
      return `<rect class="${subject}" x="${x}" y="${top}" width="${subjectWidth}" height="2.5" rx="1"/><rect class="row" x="${x}" y="${top + 3}" width="${width * 0.75}" height="1.5" rx="0.75"/>`;
    }).join("");
  }
  var LAYOUT_SKETCHES = {
    outlook: `<rect class="chrome" width="64" height="5"/>
    <rect class="pane" x="0" y="5" width="5" height="39"/>
    <rect class="pane" x="6" y="7" width="12" height="35" rx="1"/>
    ${rows(20, 8, 15)}<rect class="pane" x="37" y="7" width="25" height="35" rx="1"/>`,
    calm: `<rect class="chrome" width="64" height="5"/>
    <rect class="pane" x="2" y="8" width="13" height="33" rx="2"/>
    ${rows(18, 9, 16)}<rect class="pane" x="37" y="8" width="25" height="33" rx="2"/>`,
    fastmail: `<rect class="chrome" width="64" height="5"/>
    <rect class="pane" x="2" y="8" width="13" height="33" rx="2"/>
    <rect class="hi" x="4" y="10" width="9" height="3" rx="1.5"/>${rows(18, 9, 44)}`,
    focus: `<rect class="chrome" width="64" height="5"/>${rows(16, 9, 32)}`
  };
  var html = document.documentElement;
  var prefersDark = matchMedia("(prefers-color-scheme: dark)");
  function effectiveMode(settings) {
    if (settings.mode === "light" || settings.mode === "dark") return settings.mode;
    return prefersDark.matches ? "dark" : "light";
  }
  function applyTheme(settings) {
    html.dataset.overlookPalette = settings.palette;
    html.dataset.overlookMode = settings.mode;
    html.dataset.overlookOutlookScheme = prefersDark.matches ? "dark" : "light";
  }
  function el(tag, attrs = {}, ...children) {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) node.setAttribute(k, v);
    node.append(...children);
    return node;
  }
  function section(title, ...children) {
    return el("section", {}, el("h2", {}, title), ...children);
  }
  function choiceGroup(id, settings, className, render2) {
    const group = el("div", {
      class: className,
      role: "radiogroup",
      "aria-label": CHOICES[id].label
    });
    for (const [value, label] of Object.entries(CHOICES[id].options)) {
      const button = render2(value, label);
      button.type = "button";
      button.setAttribute("role", "radio");
      button.setAttribute("aria-checked", String(settings[id] === value));
      button.addEventListener("click", () => void saveSetting(id, value));
      group.append(button);
    }
    return group;
  }
  function swatch(settings) {
    return (value, label) => {
      const preview = el(
        "span",
        {
          class: "preview",
          "data-overlook-swatch": "",
          "data-overlook-palette": value,
          "data-overlook-mode": effectiveMode(settings)
        },
        el("span", { class: "bar" }),
        el("span", { class: "side" }),
        el(
          "span",
          { class: "body" },
          el("span", { class: "line short" }),
          el("span", { class: "line" }),
          el("span", { class: "line faint" }),
          el("span", { class: "line" })
        )
      );
      return el(
        "button",
        { class: "swatch", title: label },
        preview,
        SHORT_PALETTE_NAMES[value] ?? label
      );
    };
  }
  function layoutCard(value, label) {
    const button = el("button", { class: "layout", title: label });
    button.innerHTML = `<svg viewBox="0 0 64 44" aria-hidden="true">${LAYOUT_SKETCHES[value] ?? ""}</svg>`;
    button.append(SHORT_LAYOUT_NAMES[value] ?? label);
    return button;
  }
  function segment(value, label) {
    return el("button", { "data-value": value }, label);
  }
  function toggle(id, settings) {
    const input = el("input", { type: "checkbox", id: `toggle-${id}` });
    input.checked = settings[id];
    input.addEventListener("change", () => void saveSetting(id, input.checked));
    return el(
      "label",
      { class: "toggle", for: `toggle-${id}` },
      el("span", {}, FEATURES[id].label),
      input
    );
  }
  function toggles(ids, settings) {
    return el("div", { class: "toggles" }, ...ids.map((id) => toggle(id, settings)));
  }
  var featuresIn = (group) => FEATURE_IDS.filter((id) => FEATURES[id].group === group);
  var declutterOpen = false;
  var developerOpen = false;
  var captureStatus = "";
  async function requestCapture() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (tab?.id === void 0) return "Open an Outlook tab first.";
    try {
      const reply = await chrome.tabs.sendMessage(tab.id, { type: CAPTURE_MESSAGE });
      return `Saved ${reply.name} (${Math.round(reply.bytes / 1024)} KB) to Downloads.`;
    } catch {
      return "No answer. Use this on an Outlook tab, and reload that tab if OBB was just updated.";
    }
  }
  function developerSection() {
    const status = el("p", { class: "hint", role: "status" }, captureStatus);
    const button = el("button", { type: "button", class: "pill" }, "Save page capture");
    button.addEventListener("click", () => {
      void requestCapture().then((text) => {
        captureStatus = text;
        status.textContent = text;
      });
    });
    const details = el(
      "details",
      {},
      el("summary", {}, "Developer"),
      el(
        "p",
        { class: "hint" },
        "Saves this Outlook tab as an HTML file for fixing or testing OBB. It contains your real mail and stays on this computer; share it only with whoever is helping you fix OBB."
      ),
      button,
      status
    );
    details.open = developerOpen;
    details.addEventListener("toggle", () => {
      developerOpen = details.open;
    });
    return el("section", {}, details);
  }
  function render(settings) {
    applyTheme(settings);
    const main2 = document.getElementById("settings");
    if (!main2) return;
    const layoutHint = settings.layout === "fastmail" || settings.layout === "focus" ? [
      el(
        "p",
        { class: "hint" },
        "Pair with Outlook's Settings \u2192 Mail \u2192 Layout \u2192 Reading pane: Hide."
      )
    ] : [];
    const declutter = featuresIn("Declutter");
    const hidden = declutter.filter((id) => settings[id]).length;
    const details = el(
      "details",
      {},
      el(
        "summary",
        {},
        "Declutter",
        el("span", { class: "count" }, `${hidden} of ${declutter.length} hidden`)
      ),
      toggles(declutter, settings)
    );
    details.open = declutterOpen;
    details.addEventListener("toggle", () => {
      declutterOpen = details.open;
    });
    main2.replaceChildren(
      section("Colors", choiceGroup("palette", settings, "swatches", swatch(settings))),
      section("Light or dark", choiceGroup("mode", settings, "segmented", segment)),
      section("Layout", choiceGroup("layout", settings, "layouts", layoutCard), ...layoutHint),
      section("Keyboard", toggles(featuresIn("Keyboard"), settings)),
      el("section", {}, details),
      section("General", toggles(featuresIn("General"), settings)),
      developerSection()
    );
  }
  async function main() {
    const version = document.getElementById("version");
    if (version) version.textContent = `v${chrome.runtime.getManifest().version}`;
    document.getElementById("reload")?.addEventListener("click", () => chrome.runtime.reload());
    let settings = await loadSettings();
    render(settings);
    onSettingsChanged((next) => {
      settings = next;
      render(settings);
    });
    prefersDark.addEventListener("change", () => render(settings));
  }
  void main();
})();
//# sourceMappingURL=data:application/json;base64,ewogICJ2ZXJzaW9uIjogMywKICAic291cmNlcyI6IFsiLi4vc3JjL3NlbGVjdG9ycy50cyIsICIuLi9zcmMvY2FwdHVyZS50cyIsICIuLi9zcmMvc2V0dGluZ3MudHMiLCAiLi4vc3JjL29wdGlvbnMvb3B0aW9ucy50cyJdLAogICJzb3VyY2VzQ29udGVudCI6IFsiLy8gVGhlIG9ubHkgZmlsZSBhbGxvd2VkIHRvIGtub3cgd2hhdCBPdXRsb29rJ3MgRE9NIGxvb2tzIGxpa2UuXG4vL1xuLy8gQW5jaG9yIHByZWZlcmVuY2UgKHNlZSBQTEFOLm1kKTogc3RydWN0dXJhbCBJRHMgPiBBUklBIHJvbGVzL2xhYmVscyA+IHNlbWFudGljIGRhdGEtKiA+IHRleHQuXG4vLyBOZXZlciB1c2UgZ2VuZXJhdGVkIGNsYXNzIG5hbWVzOyBpZiBvbmUgaXMgdW5hdm9pZGFibGUsIG1hcmsgaXQgYC8vIEZSQUdJTEU6YC5cbi8vXG4vLyBWZXJpZmllZCBhZ2FpbnN0IGEgc2F2ZWQgcGFnZSBmcm9tIE9jdCAyMDI2ICh0ZXN0L2ZpeHR1cmVzL2luYm94Lmh0bWwpOyBzZWUgZG9jcy9kb20tc3VydmV5Lm1kLlxuLy8gRW50cmllcyBtYXJrZWQgVU5WRVJJRklFRCB3ZXJlIG5vdCBwcmVzZW50IGluIHRoYXQgY2FwdHVyZS4gYXJpYS1sYWJlbHMgYXNzdW1lIGFuIEVuZ2xpc2ggVUkuXG5cbmV4cG9ydCB0eXBlIFNlbGVjdG9yTmFtZSA9IGtleW9mIHR5cGVvZiBTRUxFQ1RPUlM7XG5cbmV4cG9ydCBpbnRlcmZhY2UgU2VsZWN0b3JFbnRyeSB7XG4gIC8qKiBDYW5kaWRhdGUgQ1NTIHNlbGVjdG9ycywgbW9zdCBzdGFibGUgZmlyc3QuIFRoZSBmaXJzdCBvbmUgdGhhdCBtYXRjaGVzIHdpbnMuICovXG4gIGNhbmRpZGF0ZXM6IHJlYWRvbmx5IHN0cmluZ1tdO1xuICAvKiogXCJhbHdheXNcIjogc2hvdWxkIGV4aXN0IG9uIGV2ZXJ5IG1haWwgdmlldywgc28gYSBtaXNzIGlzIGEgYnJlYWthZ2UuIFwic29tZXRpbWVzXCI6IGNvbnRleHR1YWwuICovXG4gIGV4cGVjdDogXCJhbHdheXNcIiB8IFwic29tZXRpbWVzXCI7XG4gIC8qKiBOYW1lIG9mIGFub3RoZXIgZW50cnkgdG8gcmVzb2x2ZSBpbnNpZGUsIGluc3RlYWQgb2YgdGhlIHdob2xlIGRvY3VtZW50LiAqL1xuICB3aXRoaW4/OiBzdHJpbmc7XG4gIC8qKiBUb2tlbiBhZGRlZCB0byB0aGUgZWxlbWVudCdzIGBkYXRhLW92ZXJsb29rYCBhdHRyaWJ1dGUgc28gQ1NTIGNhbiB0YXJnZXQgaXQuICovXG4gIHRhZz86IHN0cmluZztcbiAgLyoqIFRhZyBldmVyeSBtYXRjaCByYXRoZXIgdGhhbiBvbmx5IHRoZSBmaXJzdC4gKi9cbiAgYWxsPzogYm9vbGVhbjtcbiAgLyoqIE1hdGNoIGV2ZXJ5IGNhbmRpZGF0ZSAoYSBsaXN0IG9mIGRpc3RpbmN0IHRoaW5ncyksIG5vdCBqdXN0IHRoZSBmaXJzdCB0aGF0IG1hdGNoZXMuICovXG4gIHVuaW9uPzogYm9vbGVhbjtcbn1cblxuZXhwb3J0IGNvbnN0IFNFTEVDVE9SUyA9IHtcbiAgLy8gLS0tIFN0cnVjdHVyZSAtLS1cbiAgdGl0bGVCYXI6IHsgY2FuZGlkYXRlczogW1wiI093YVRpdGxlQmFyXCJdLCBleHBlY3Q6IFwiYWx3YXlzXCIsIHRhZzogXCJ0aXRsZS1iYXJcIiB9LFxuICBzZWFyY2hCb3g6IHsgY2FuZGlkYXRlczogW1wiI3RvcFNlYXJjaElucHV0XCIsICdbcm9sZT1cInNlYXJjaFwiXSBpbnB1dCddLCBleHBlY3Q6IFwiYWx3YXlzXCIgfSxcbiAgYXBwUmFpbDoge1xuICAgIGNhbmRpZGF0ZXM6IFtcIiNMZWZ0UmFpbFwiLCAnW3JvbGU9XCJuYXZpZ2F0aW9uXCJdW2FyaWEtbGFiZWw9XCJsZWZ0LXJhaWwtYXBwYmFyXCJdJ10sXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICAgIHRhZzogXCJhcHAtcmFpbFwiLFxuICB9LFxuICByaWJib246IHtcbiAgICBjYW5kaWRhdGVzOiBbJ1tkYXRhLWFwcC1zZWN0aW9uPVwiUmliYm9uXCJdJywgJ1tyb2xlPVwicmVnaW9uXCJdW2FyaWEtbGFiZWw9XCJSaWJib25cIl0nXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgdGFnOiBcInJpYmJvblwiLFxuICB9LFxuICByaWJib25UYWJzOiB7XG4gICAgY2FuZGlkYXRlczogWydbcm9sZT1cInRhYmxpc3RcIl1bYXJpYS1sYWJlbD1cIlJpYmJvbiB0YWJzXCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJyaWJib25cIixcbiAgICB0YWc6IFwicmliYm9uLXRhYnNcIixcbiAgfSxcbiAgZm9sZGVyUGFuZToge1xuICAgIGNhbmRpZGF0ZXM6IFtcbiAgICAgICdbZGF0YS1hcHAtc2VjdGlvbj1cIk5hdmlnYXRpb25QYW5lXCJdJyxcbiAgICAgICdbcm9sZT1cInJlZ2lvblwiXVthcmlhLWxhYmVsPVwiTmF2aWdhdGlvbiBwYW5lXCJdJyxcbiAgICBdLFxuICAgIGV4cGVjdDogXCJhbHdheXNcIixcbiAgICB0YWc6IFwiZm9sZGVyLXBhbmVcIixcbiAgfSxcbiAgZm9sZGVySXRlbToge1xuICAgIGNhbmRpZGF0ZXM6IFsnW3JvbGU9XCJ0cmVlaXRlbVwiXVtkYXRhLWZvbGRlci1uYW1lXScsICdbcm9sZT1cInRyZWVpdGVtXCJdJ10sXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICAgIHdpdGhpbjogXCJmb2xkZXJQYW5lXCIsXG4gICAgdGFnOiBcImZvbGRlci1pdGVtXCIsXG4gICAgYWxsOiB0cnVlLFxuICB9LFxuICBsaXN0UGFuZToge1xuICAgIGNhbmRpZGF0ZXM6IFsnW2RhdGEtYXBwLXNlY3Rpb249XCJNZXNzYWdlTGlzdFwiXSddLFxuICAgIGV4cGVjdDogXCJhbHdheXNcIixcbiAgICB0YWc6IFwibGlzdC1wYW5lXCIsXG4gIH0sXG4gIG1lc3NhZ2VMaXN0OiB7XG4gICAgY2FuZGlkYXRlczogW1wiI01haWxMaXN0XCIsICdbcm9sZT1cImxpc3Rib3hcIl1bYXJpYS1sYWJlbD1cIk1lc3NhZ2UgbGlzdFwiXSddLFxuICAgIGV4cGVjdDogXCJhbHdheXNcIixcbiAgICB0YWc6IFwibWVzc2FnZS1saXN0XCIsXG4gIH0sXG4gIG1lc3NhZ2VSb3c6IHtcbiAgICBjYW5kaWRhdGVzOiBbJ1tyb2xlPVwib3B0aW9uXCJdW2RhdGEtY29udmlkXScsICdbcm9sZT1cIm9wdGlvblwiXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB3aXRoaW46IFwibWVzc2FnZUxpc3RcIixcbiAgICB0YWc6IFwibWVzc2FnZS1yb3dcIixcbiAgICBhbGw6IHRydWUsXG4gIH0sXG4gIGdyb3VwSGVhZGVyOiB7XG4gICAgY2FuZGlkYXRlczogWydbcm9sZT1cImJ1dHRvblwiXVtpZF49XCJncm91cEhlYWRlclwiXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB3aXRoaW46IFwibWVzc2FnZUxpc3RcIixcbiAgICB0YWc6IFwiZ3JvdXAtaGVhZGVyXCIsXG4gICAgYWxsOiB0cnVlLFxuICB9LFxuICBmb2N1c2VkVGFiczoge1xuICAgIGNhbmRpZGF0ZXM6IFsnW3JvbGU9XCJ0YWJsaXN0XCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJsaXN0UGFuZVwiLFxuICAgIHRhZzogXCJmb2N1c2VkLXRhYnNcIixcbiAgfSxcbiAgcmVhZGluZ1BhbmU6IHtcbiAgICBjYW5kaWRhdGVzOiBbXG4gICAgICAnW2RhdGEtYXBwLXNlY3Rpb249XCJNYWlsUmVhZENvbXBvc2VcIl0nLFxuICAgICAgXCIjUmVhZGluZ1BhbmVDb250YWluZXJJZFwiLFxuICAgICAgJ1tyb2xlPVwibWFpblwiXVthcmlhLWxhYmVsPVwiUmVhZGluZyBQYW5lXCJdJyxcbiAgICBdLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB0YWc6IFwicmVhZGluZy1wYW5lXCIsXG4gIH0sXG5cbiAgbWVzc2FnZUJvZHk6IHtcbiAgICBjYW5kaWRhdGVzOiBbJ1tyb2xlPVwiZG9jdW1lbnRcIl1bYXJpYS1sYWJlbD1cIk1lc3NhZ2UgYm9keVwiXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB3aXRoaW46IFwicmVhZGluZ1BhbmVcIixcbiAgICB0YWc6IFwibWVzc2FnZS1ib2R5XCIsXG4gICAgYWxsOiB0cnVlLFxuICB9LFxuXG4gIGluZm9CYXI6IHtcbiAgICAvLyBOb3RpY2VzIGFib3ZlIGEgbWVzc2FnZSAoXCJZb3UgcmVwbGllZCBvblx1MjAyNlwiKTsgdGhlaXIgdGV4dCBzaXRzIGluIGFuIGVsZW1lbnQgd2hvc2UgaWQgc3RhcnRzXG4gICAgLy8gd2l0aCBcImluZm9CYXJcIi4gT3V0bG9vayBwYWludHMgdGhlbSB3aXRoIGlubGluZSBjb2xvdXJzLlxuICAgIGNhbmRpZGF0ZXM6IFsnW3JvbGU9XCJncm91cFwiXTpoYXMoPiAqID4gW2lkXj1cImluZm9CYXJcIl0pJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJyZWFkaW5nUGFuZVwiLFxuICAgIHRhZzogXCJpbmZvLWJhclwiLFxuICAgIGFsbDogdHJ1ZSxcbiAgfSxcblxuICAvLyAtLS0gQ2x1dHRlciAoZWFjaCBoaWRlYWJsZSBlbnRyeSBnZXRzIGEgdG9nZ2xlIGluIHNldHRpbmdzLnRzKSAtLS1cbiAgdXBzZWxsQmFubmVyOiB7XG4gICAgY2FuZGlkYXRlczogWydbZGF0YS1hcHAtc2VjdGlvbj1cIlVwc2VsbEJhbm5lclNlY3Rpb25cIl0nXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgdGFnOiBcInVwc2VsbFwiLFxuICAgIGFsbDogdHJ1ZSxcbiAgfSxcbiAgaGVhZGVyRXh0cmFzOiB7XG4gICAgY2FuZGlkYXRlczogW1wiI293YVRlYW1zQnRuX2NvbnRhaW5lclwiLCBcIiNvd2FUaW1lUGFuZWxCdG5fY29udGFpbmVyXCJdLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB0YWc6IFwiaGVhZGVyLWV4dHJhXCIsXG4gICAgdW5pb246IHRydWUsXG4gIH0sXG4gIGFwcFJhaWxFeHRyYXM6IHtcbiAgICBjYW5kaWRhdGVzOiBbXCJUbyBEb1wiLCBcIk5ld3NsZXR0ZXJzXCIsIFwiT3JnIEV4cGxvcmVyXCIsIFwiT25lRHJpdmVcIl0ubWFwKFxuICAgICAgKGFwcCkgPT4gYFtyb2xlPVwibm9uZVwiXVthcmlhLWxhYmVsPVwiJHthcHB9XCJdYCxcbiAgICApLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB3aXRoaW46IFwiYXBwUmFpbFwiLFxuICAgIHRhZzogXCJyYWlsLWV4dHJhXCIsXG4gICAgdW5pb246IHRydWUsXG4gIH0sXG4gIHJpYmJvbkV4dHJhczoge1xuICAgIGNhbmRpZGF0ZXM6IFtcbiAgICAgIFwiU2hhcmUgdG8gVGVhbXNcIixcbiAgICAgIFwiU3dlZXBcIixcbiAgICAgIFwiQXNzaWduIHBvbGljeVwiLFxuICAgICAgXCJWaXZhIEluc2lnaHRzXCIsXG4gICAgICBcIk1vcmUgYXBwc1wiLFxuICAgICAgXCJRdWljayBzdGVwc1wiLFxuICAgICAgXCJNb3JlIHF1aWNrIHN0ZXBzXCIsXG4gICAgXS5tYXAoKGxhYmVsKSA9PiBgW2FyaWEtbGFiZWw9XCIke2xhYmVsfVwiXWApLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB3aXRoaW46IFwicmliYm9uXCIsXG4gICAgdGFnOiBcInJpYmJvbi1leHRyYVwiLFxuICAgIHVuaW9uOiB0cnVlLFxuICB9LFxuICByb3dBdmF0YXI6IHtcbiAgICBjYW5kaWRhdGVzOiBbJ1tyb2xlPVwiaW1nXCJdW2lkXj1cImF2YXRhci1cIl0nXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgd2l0aGluOiBcIm1lc3NhZ2VMaXN0XCIsXG4gICAgdGFnOiBcInJvdy1hdmF0YXJcIixcbiAgICBhbGw6IHRydWUsXG4gIH0sXG4gIHJvd0hvdmVyQWN0aW9uczoge1xuICAgIGNhbmRpZGF0ZXM6IFtcbiAgICAgIFwiQ3VzdG9taXplIEFjdGlvbnNcIixcbiAgICAgIFwiRmxhZyB0aGlzIG1lc3NhZ2VcIixcbiAgICAgIFwiS2VlcCB0aGlzIG1lc3NhZ2UgYXQgdGhlIHRvcCBvZiB5b3VyIGZvbGRlclwiLFxuICAgIF0ubWFwKChsYWJlbCkgPT4gYGJ1dHRvblthcmlhLWxhYmVsPVwiJHtsYWJlbH1cIl1gKSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgd2l0aGluOiBcIm1lc3NhZ2VMaXN0XCIsXG4gICAgdGFnOiBcInJvdy1hY3Rpb25cIixcbiAgICB1bmlvbjogdHJ1ZSxcbiAgfSxcbiAgYXBwQmFubmVyOiB7XG4gICAgLy8gRmx1ZW50J3Mgc3RhdGljIFwiZnVpLU1lc3NhZ2VCYXJcIiBjbGFzcyBpcyBwYXJ0IG9mIGl0cyBwdWJsaWMgc3R5bGluZyBBUEksIG5vdCBnZW5lcmF0ZWQ7XG4gICAgLy8gbm90aGluZyBlbHNlIG1hcmtzIHRoZXNlLiBTY29wZWQgdG8gYmFycyBhY3Jvc3MgdGhlIHRvcCBvZiB0aGUgYXBwLCBub3QgaW5zaWRlIG1lc3NhZ2VzLlxuICAgIGNhbmRpZGF0ZXM6IFsnI21haW5BcHAgPiBkaXYgPiBbcm9sZT1cImdyb3VwXCJdLmZ1aS1NZXNzYWdlQmFyJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHRhZzogXCJhcHAtYmFubmVyXCIsXG4gICAgYWxsOiB0cnVlLFxuICB9LFxuICBjb3BpbG90OiB7XG4gICAgLy8gVU5WRVJJRklFRDogbm8gQ29waWxvdCBzdXJmYWNlcyBpbiB0aGUgT2N0IDIwMjYgY2FwdHVyZS5cbiAgICBjYW5kaWRhdGVzOiBbJ2J1dHRvblthcmlhLWxhYmVsKj1cIkNvcGlsb3RcIiBpXScsICdbcm9sZT1cInJlZ2lvblwiXVthcmlhLWxhYmVsKj1cIkNvcGlsb3RcIiBpXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB0YWc6IFwiY29waWxvdFwiLFxuICAgIHVuaW9uOiB0cnVlLFxuICB9LFxuICBleHRlcm5hbFRhZzoge1xuICAgIC8vIFVOVkVSSUZJRUQ6IG5vIEV4dGVybmFsIHRhZ3MgaW4gdGhlIE9jdCAyMDI2IGNhcHR1cmUuXG4gICAgY2FuZGlkYXRlczogWydbdGl0bGU9XCJFeHRlcm5hbCBzZW5kZXJcIl0nLCAnW2FyaWEtbGFiZWw9XCJFeHRlcm5hbFwiXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB0YWc6IFwiZXh0ZXJuYWwtdGFnXCIsXG4gICAgdW5pb246IHRydWUsXG4gIH0sXG5cbiAgLy8gLS0tIENvbnRyb2xzIHRoZSBrZXltYXAgZHJpdmVzIC0tLVxuICBjb21wb3NlQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJOZXcgbWFpbFwiXSddLCBleHBlY3Q6IFwiYWx3YXlzXCIgfSxcbiAgYXJjaGl2ZUJ1dHRvbjogeyBjYW5kaWRhdGVzOiBbJ2J1dHRvblthcmlhLWxhYmVsPVwiQXJjaGl2ZVwiXSddLCBleHBlY3Q6IFwic29tZXRpbWVzXCIgfSxcbiAgZGVsZXRlQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJEZWxldGVcIl0nXSwgZXhwZWN0OiBcInNvbWV0aW1lc1wiIH0sXG4gIC8vIFVOVkVSSUZJRUQ6IHRoaXMgdGVuYW50IHNob3dzIGEgXCJQaGlzaCBBbGVydFwiIGFkZC1pbiBpbnN0ZWFkIG9mIE91dGxvb2sncyBSZXBvcnQgbWVudS5cbiAganVua0J1dHRvbjoge1xuICAgIGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJSZXBvcnRcIl0nLCAnYnV0dG9uW2FyaWEtbGFiZWw9XCJKdW5rXCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICB9LFxuICBtb3ZlQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJNb3ZlIHRvXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICBzbm9vemVCdXR0b246IHsgY2FuZGlkYXRlczogWydidXR0b25bYXJpYS1sYWJlbD1cIlNub296ZVwiXSddLCBleHBlY3Q6IFwic29tZXRpbWVzXCIgfSxcbiAgZmxhZ0J1dHRvbjogeyBjYW5kaWRhdGVzOiBbJ2J1dHRvblthcmlhLWxhYmVsPVwiRmxhZyAvIFVuZmxhZ1wiXSddLCBleHBlY3Q6IFwic29tZXRpbWVzXCIgfSxcbiAgcmVhZFVucmVhZEJ1dHRvbjogeyBjYW5kaWRhdGVzOiBbJ2J1dHRvblthcmlhLWxhYmVsPVwiUmVhZCAvIFVucmVhZFwiXSddLCBleHBlY3Q6IFwic29tZXRpbWVzXCIgfSxcbiAgdW5kb0J1dHRvbjogeyBjYW5kaWRhdGVzOiBbJ2J1dHRvblthcmlhLWxhYmVsPVwiVW5kb1wiXSddLCBleHBlY3Q6IFwic29tZXRpbWVzXCIgfSxcbiAgc3luY0J1dHRvbjogeyBjYW5kaWRhdGVzOiBbXCIjb3dhU3luY0J1dHRvbl9jb250YWluZXJcIl0sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICBzZXR0aW5nc0J1dHRvbjogeyBjYW5kaWRhdGVzOiBbXCIjb3dhU2V0dGluZ3NCdG5fY29udGFpbmVyXCJdLCBleHBlY3Q6IFwiYWx3YXlzXCIgfSxcbiAgLy8gQXBwIHJhaWwgYnV0dG9ucy4gRmFsbGJhY2tzIGFyZSB0aGUgTWljcm9zb2Z0IDM2NSBhcHAgaWRzIE91dGxvb2sgcHV0cyBpbiBkYXRhLXRpZCwgdGhlIHNhbWVcbiAgLy8gYWNyb3NzIHRoZSBPY3QgYW5kIE5vdiAyMDI2IGNhcHR1cmVzLiBUaGUgcmFpbCByZW5kZXJzIGFmdGVyIHRoZSBtZXNzYWdlIGxpc3QuXG4gIG1haWxBcHA6IHtcbiAgICBjYW5kaWRhdGVzOiBbJ2J1dHRvblthcmlhLWxhYmVsPVwiTWFpbFwiXScsICdbZGF0YS10aWQ9XCJkZGVhNzc0Yy0zODJiLTQ3ZDctYWFiNS1hZGMyMTM5YTgwMmJcIl0nXSxcbiAgICB3aXRoaW46IFwiYXBwUmFpbFwiLFxuICAgIGV4cGVjdDogXCJhbHdheXNcIixcbiAgfSxcbiAgY2FsZW5kYXJBcHA6IHtcbiAgICBjYW5kaWRhdGVzOiBbXG4gICAgICAnYnV0dG9uW2FyaWEtbGFiZWw9XCJDYWxlbmRhclwiXScsXG4gICAgICAnW2RhdGEtdGlkPVwiOGNiZWI4NmYtODNlMS00M2I1LWFhYmEtY2QzNTE0MzIyZjBiXCJdJyxcbiAgICBdLFxuICAgIHdpdGhpbjogXCJhcHBSYWlsXCIsXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICB9LFxuICBwZW9wbGVBcHA6IHtcbiAgICBjYW5kaWRhdGVzOiBbXG4gICAgICAnYnV0dG9uW2FyaWEtbGFiZWw9XCJQZW9wbGVcIl0nLFxuICAgICAgJ1tkYXRhLXRpZD1cIjM1NWZiZDc5LTNiYTItNDU1NC04ZjJkLTAzMDBmZGU5MWYzMFwiXScsXG4gICAgXSxcbiAgICB3aXRoaW46IFwiYXBwUmFpbFwiLFxuICAgIGV4cGVjdDogXCJhbHdheXNcIixcbiAgfSxcbn0gYXMgY29uc3Qgc2F0aXNmaWVzIFJlY29yZDxzdHJpbmcsIFNlbGVjdG9yRW50cnk+O1xuXG5jb25zdCBlbnRyeSA9IChuYW1lOiBTZWxlY3Rvck5hbWUpOiBTZWxlY3RvckVudHJ5ID0+IFNFTEVDVE9SU1tuYW1lXTtcblxuZnVuY3Rpb24gc2NvcGUobmFtZTogU2VsZWN0b3JOYW1lLCByb290OiBQYXJlbnROb2RlKTogUGFyZW50Tm9kZSB8IG51bGwge1xuICBjb25zdCBwYXJlbnQgPSBlbnRyeShuYW1lKS53aXRoaW47XG4gIHJldHVybiBwYXJlbnQgPyByZXNvbHZlKHBhcmVudCBhcyBTZWxlY3Rvck5hbWUsIHJvb3QpIDogcm9vdDtcbn1cblxuLyoqIFRoZSBmaXJzdCBlbGVtZW50IG1hdGNoaW5nIGBuYW1lYCwgdHJ5aW5nIGNhbmRpZGF0ZXMgaW4gcHJpb3JpdHkgb3JkZXIuICovXG5leHBvcnQgZnVuY3Rpb24gcmVzb2x2ZShuYW1lOiBTZWxlY3Rvck5hbWUsIHJvb3Q6IFBhcmVudE5vZGUgPSBkb2N1bWVudCk6IEVsZW1lbnQgfCBudWxsIHtcbiAgY29uc3Qgd2l0aGluID0gc2NvcGUobmFtZSwgcm9vdCk7XG4gIGlmICghd2l0aGluKSByZXR1cm4gbnVsbDtcbiAgZm9yIChjb25zdCBjYW5kaWRhdGUgb2YgZW50cnkobmFtZSkuY2FuZGlkYXRlcykge1xuICAgIGNvbnN0IGVsID0gd2l0aGluLnF1ZXJ5U2VsZWN0b3IoY2FuZGlkYXRlKTtcbiAgICBpZiAoZWwpIHJldHVybiBlbDtcbiAgfVxuICByZXR1cm4gbnVsbDtcbn1cblxuLyoqIEV2ZXJ5IG1hdGNoOiBmb3IgdW5pb24gZW50cmllcywgb2YgYWxsIGNhbmRpZGF0ZXM7IG90aGVyd2lzZSBvZiB0aGUgZmlyc3QgY2FuZGlkYXRlIHRoYXQgbWF0Y2hlcy4gKi9cbmV4cG9ydCBmdW5jdGlvbiByZXNvbHZlQWxsKG5hbWU6IFNlbGVjdG9yTmFtZSwgcm9vdDogUGFyZW50Tm9kZSA9IGRvY3VtZW50KTogRWxlbWVudFtdIHtcbiAgY29uc3Qgd2l0aGluID0gc2NvcGUobmFtZSwgcm9vdCk7XG4gIGlmICghd2l0aGluKSByZXR1cm4gW107XG4gIGNvbnN0IHsgY2FuZGlkYXRlcywgdW5pb24gfSA9IGVudHJ5KG5hbWUpO1xuICBpZiAodW5pb24pIHJldHVybiBbLi4ud2l0aGluLnF1ZXJ5U2VsZWN0b3JBbGwoY2FuZGlkYXRlcy5qb2luKFwiLCBcIikpXTtcbiAgZm9yIChjb25zdCBjYW5kaWRhdGUgb2YgY2FuZGlkYXRlcykge1xuICAgIGNvbnN0IGVscyA9IHdpdGhpbi5xdWVyeVNlbGVjdG9yQWxsKGNhbmRpZGF0ZSk7XG4gICAgaWYgKGVscy5sZW5ndGggPiAwKSByZXR1cm4gWy4uLmVsc107XG4gIH1cbiAgcmV0dXJuIFtdO1xufVxuXG4vKiogQWRkIGBkYXRhLW92ZXJsb29rYCB0b2tlbnMgdG8gZXZlcnkgdGFnZ2VkIGVudHJ5IHNvIHN0eWxlc2hlZXRzIG5ldmVyIHRvdWNoIE91dGxvb2sncyBET00gc2hhcGUuICovXG5leHBvcnQgZnVuY3Rpb24gdGFnRWxlbWVudHMocm9vdDogUGFyZW50Tm9kZSA9IGRvY3VtZW50KTogdm9pZCB7XG4gIGZvciAoY29uc3QgbmFtZSBvZiBPYmplY3Qua2V5cyhTRUxFQ1RPUlMpIGFzIFNlbGVjdG9yTmFtZVtdKSB7XG4gICAgY29uc3QgeyB0YWcsIGFsbCwgdW5pb24gfSA9IGVudHJ5KG5hbWUpO1xuICAgIGlmICghdGFnKSBjb250aW51ZTtcbiAgICBjb25zdCBlbHMgPVxuICAgICAgYWxsIHx8IHVuaW9uID8gcmVzb2x2ZUFsbChuYW1lLCByb290KSA6IFtyZXNvbHZlKG5hbWUsIHJvb3QpXS5maWx0ZXIoKGUpID0+IGUgIT09IG51bGwpO1xuICAgIGZvciAoY29uc3QgZWwgb2YgZWxzKSB7XG4gICAgICBjb25zdCB0b2tlbnMgPSBuZXcgU2V0KChlbC5nZXRBdHRyaWJ1dGUoXCJkYXRhLW92ZXJsb29rXCIpID8/IFwiXCIpLnNwbGl0KFwiIFwiKS5maWx0ZXIoQm9vbGVhbikpO1xuICAgICAgaWYgKHRva2Vucy5oYXModGFnKSkgY29udGludWU7XG4gICAgICB0b2tlbnMuYWRkKHRhZyk7XG4gICAgICBlbC5zZXRBdHRyaWJ1dGUoXCJkYXRhLW92ZXJsb29rXCIsIFsuLi50b2tlbnNdLmpvaW4oXCIgXCIpKTtcbiAgICB9XG4gIH1cbn1cblxuZXhwb3J0IGludGVyZmFjZSBIZWFsdGhSZXN1bHQge1xuICBuYW1lOiBTZWxlY3Rvck5hbWU7XG4gIGV4cGVjdDogU2VsZWN0b3JFbnRyeVtcImV4cGVjdFwiXTtcbiAgLyoqIEluZGV4IG9mIHRoZSBjYW5kaWRhdGUgdGhhdCBtYXRjaGVkLCBvciAtMS4gQSBub24temVybyBpbmRleCBtZWFucyBhIGZhbGxiYWNrIGlzIGluIHVzZS4gKi9cbiAgY2FuZGlkYXRlOiBudW1iZXI7XG4gIGNvdW50OiBudW1iZXI7XG59XG5cbmV4cG9ydCBmdW5jdGlvbiBoZWFsdGhDaGVjayhyb290OiBQYXJlbnROb2RlID0gZG9jdW1lbnQpOiBIZWFsdGhSZXN1bHRbXSB7XG4gIHJldHVybiAoT2JqZWN0LmtleXMoU0VMRUNUT1JTKSBhcyBTZWxlY3Rvck5hbWVbXSkubWFwKChuYW1lKSA9PiB7XG4gICAgY29uc3QgeyBjYW5kaWRhdGVzLCBleHBlY3QsIHVuaW9uIH0gPSBlbnRyeShuYW1lKTtcbiAgICBjb25zdCB3aXRoaW4gPSBzY29wZShuYW1lLCByb290KTtcbiAgICBsZXQgY2FuZGlkYXRlID0gLTE7XG4gICAgbGV0IGNvdW50ID0gMDtcbiAgICBpZiAod2l0aGluKSB7XG4gICAgICBjYW5kaWRhdGUgPSBjYW5kaWRhdGVzLmZpbmRJbmRleCgoYykgPT4gd2l0aGluLnF1ZXJ5U2VsZWN0b3IoYykgIT09IG51bGwpO1xuICAgICAgaWYgKHVuaW9uICYmIGNhbmRpZGF0ZSA+PSAwKSBjb3VudCA9IHJlc29sdmVBbGwobmFtZSwgcm9vdCkubGVuZ3RoO1xuICAgICAgZWxzZSBpZiAoY2FuZGlkYXRlID49IDApIGNvdW50ID0gd2l0aGluLnF1ZXJ5U2VsZWN0b3JBbGwoY2FuZGlkYXRlc1tjYW5kaWRhdGVdID8/IFwiXCIpLmxlbmd0aDtcbiAgICB9XG4gICAgLy8gQSB1bmlvbiBlbnRyeSBsaXN0cyBkaXN0aW5jdCB0aGluZ3MsIHNvIG1hdGNoaW5nIGEgbGF0ZXIgb25lIGlzIG5vdCBhIGZhbGxiYWNrLlxuICAgIGlmICh1bmlvbiAmJiBjYW5kaWRhdGUgPiAwKSBjYW5kaWRhdGUgPSAwO1xuICAgIHJldHVybiB7IG5hbWUsIGV4cGVjdCwgY2FuZGlkYXRlLCBjb3VudCB9O1xuICB9KTtcbn1cblxuZXhwb3J0IGludGVyZmFjZSBGb2xkZXJOb2RlIHtcbiAgbmFtZTogc3RyaW5nO1xuICAvKiogQW5jZXN0b3IgbmFtZXMgam9pbmVkIHdpdGggXCIvXCIsIGUuZy4gXCJJbmJveC9SZWNlaXB0c1wiLiAqL1xuICBwYXRoOiBzdHJpbmc7XG4gIGVsZW1lbnQ6IEVsZW1lbnQ7XG59XG5cbi8qKiBcIkluYm94IC0gOSBpdGVtcyAoMSB1bnJlYWQpXCIgLT4gXCJJbmJveFwiLiBUaGUgdGl0bGUgYXR0cmlidXRlIGNhcnJpZXMgY291bnRzLiAqL1xuZnVuY3Rpb24gZm9sZGVyTmFtZShlbGVtZW50OiBFbGVtZW50KTogc3RyaW5nIHtcbiAgY29uc3QgdGl0bGUgPSAoZWxlbWVudC5nZXRBdHRyaWJ1dGUoXCJ0aXRsZVwiKSA/PyBlbGVtZW50LnRleHRDb250ZW50ID8/IFwiXCIpLnRyaW0oKTtcbiAgcmV0dXJuIHRpdGxlLnJlcGxhY2UoL1xccystXFxzK1xcZCtcXHMraXRlbXM/KFxccytcXChcXGQrXFxzK3VucmVhZFxcKSk/JC8sIFwiXCIpO1xufVxuXG4vKipcbiAqIEZvbGRlcnMgY3VycmVudGx5IHJlbmRlcmVkIGluIHRoZSBmb2xkZXIgcGFuZSwgd2l0aCBwYXRocyBidWlsdCBmcm9tIGFyaWEtbGV2ZWwuIEZvbGRlcnNcbiAqIHVuZGVyIEZhdm9yaXRlcyBhcmUgc2tpcHBlZCAodGhleSBkdXBsaWNhdGUgdGhlIG1haWxib3ggdHJlZSksIGFzIGFyZSB0aGUgdG9wLWxldmVsIG1haWxib3hcbiAqIHJvb3RzLCB3aG9zZSBuYW1lcyBhcmUgYWNjb3VudCBhZGRyZXNzZXMuXG4gKiBVTlZFUklGSUVEOiBjb2xsYXBzZWQgc3ViZm9sZGVycyBtYXkgYmUgYWJzZW50IGZyb20gdGhlIERPTS5cbiAqL1xuZXhwb3J0IGZ1bmN0aW9uIHJlYWRGb2xkZXJUcmVlKHJvb3Q6IFBhcmVudE5vZGUgPSBkb2N1bWVudCk6IEZvbGRlck5vZGVbXSB7XG4gIGNvbnN0IHN0YWNrOiBzdHJpbmdbXSA9IFtdO1xuICBjb25zdCBmb2xkZXJzOiBGb2xkZXJOb2RlW10gPSBbXTtcbiAgbGV0IGluRmF2b3JpdGVzID0gZmFsc2U7XG4gIGZvciAoY29uc3QgZWxlbWVudCBvZiByZXNvbHZlQWxsKFwiZm9sZGVySXRlbVwiLCByb290KSkge1xuICAgIGNvbnN0IGxldmVsID0gTnVtYmVyKGVsZW1lbnQuZ2V0QXR0cmlidXRlKFwiYXJpYS1sZXZlbFwiKSA/PyBcIjFcIik7XG4gICAgY29uc3QgbmFtZSA9IGZvbGRlck5hbWUoZWxlbWVudCk7XG4gICAgc3RhY2subGVuZ3RoID0gTWF0aC5tYXgobGV2ZWwgLSAxLCAwKTtcbiAgICBzdGFjay5wdXNoKG5hbWUpO1xuICAgIGlmIChsZXZlbCA9PT0gMSkge1xuICAgICAgaW5GYXZvcml0ZXMgPSBlbGVtZW50LmdldEF0dHJpYnV0ZShcImRhdGEtZm9sZGVyLW5hbWVcIikgPT09IFwiZmF2b3JpdGVzXCI7XG4gICAgICBjb250aW51ZTtcbiAgICB9XG4gICAgaWYgKGluRmF2b3JpdGVzKSBjb250aW51ZTtcbiAgICBmb2xkZXJzLnB1c2goeyBuYW1lLCBwYXRoOiBzdGFjay5zbGljZSgxKS5maWx0ZXIoQm9vbGVhbikuam9pbihcIi9cIiksIGVsZW1lbnQgfSk7XG4gIH1cbiAgcmV0dXJuIGZvbGRlcnM7XG59XG5cbi8qKiBNZXNzYWdlIHJvd3MgY3VycmVudGx5IHNlbGVjdGVkIChvciBvcGVuKSBpbiB0aGUgbGlzdC4gKi9cbmV4cG9ydCBmdW5jdGlvbiBzZWxlY3RlZFJvd3Mocm9vdDogUGFyZW50Tm9kZSA9IGRvY3VtZW50KTogRWxlbWVudFtdIHtcbiAgcmV0dXJuIHJlc29sdmVBbGwoXCJtZXNzYWdlUm93XCIsIHJvb3QpLmZpbHRlcihcbiAgICAocm93KSA9PiByb3cuZ2V0QXR0cmlidXRlKFwiYXJpYS1zZWxlY3RlZFwiKSA9PT0gXCJ0cnVlXCIsXG4gICk7XG59XG5cbi8qKiBBIHJvdydzIHRpY2sgYm94IChcIlNlbGVjdCBhIGNvbnZlcnNhdGlvblwiKTsgaXRzIGFyaWEtY2hlY2tlZCBzYXlzIHdoZXRoZXIgaXQgaXMgdGlja2VkLiAqL1xuZXhwb3J0IGZ1bmN0aW9uIHJvd0NoZWNrYm94KHJvdzogRWxlbWVudCk6IEhUTUxFbGVtZW50IHwgbnVsbCB7XG4gIHJldHVybiByb3cucXVlcnlTZWxlY3RvcjxIVE1MRWxlbWVudD4oJ1tyb2xlPVwiY2hlY2tib3hcIl0nKTtcbn1cblxuLyoqIFJvd3Mgd2hvc2UgdGljayBib3ggaXMgdGlja2VkOiB0aGUgbXVsdGktc2VsZWN0aW9uLCBhcyBvcHBvc2VkIHRvIHRoZSBzaW5nbGUgb3BlbiByb3cuICovXG5leHBvcnQgZnVuY3Rpb24gY2hlY2tlZFJvd3Mocm9vdDogUGFyZW50Tm9kZSA9IGRvY3VtZW50KTogRWxlbWVudFtdIHtcbiAgcmV0dXJuIHJlc29sdmVBbGwoXCJtZXNzYWdlUm93XCIsIHJvb3QpLmZpbHRlcihcbiAgICAocm93KSA9PiByb3dDaGVja2JveChyb3cpPy5nZXRBdHRyaWJ1dGUoXCJhcmlhLWNoZWNrZWRcIikgPT09IFwidHJ1ZVwiLFxuICApO1xufVxuXG4vKipcbiAqIFRoZSByb3cgdGhlIGtleWJvYXJkIGN1cnNvciBpcyBvbjogdGhlIG9uZSBob2xkaW5nIGZvY3VzLCBlbHNlIE91dGxvb2sncyByb3ZpbmctdGFiaW5kZXggcm93LFxuICogZWxzZSB0aGUgc2VsZWN0ZWQgb25lLlxuICovXG5leHBvcnQgZnVuY3Rpb24gY3Vyc29yUm93KHJvb3Q6IERvY3VtZW50ID0gZG9jdW1lbnQpOiBFbGVtZW50IHwgbnVsbCB7XG4gIGNvbnN0IHJvd3MgPSByZXNvbHZlQWxsKFwibWVzc2FnZVJvd1wiLCByb290KTtcbiAgY29uc3QgYWN0aXZlID0gcm9vdC5hY3RpdmVFbGVtZW50O1xuICByZXR1cm4gKFxuICAgIHJvd3MuZmluZCgocm93KSA9PiBhY3RpdmUgJiYgcm93LmNvbnRhaW5zKGFjdGl2ZSkpID8/XG4gICAgcm93cy5maW5kKChyb3cpID0+IHJvdy5nZXRBdHRyaWJ1dGUoXCJ0YWJpbmRleFwiKSA9PT0gXCIwXCIpID8/XG4gICAgc2VsZWN0ZWRSb3dzKHJvb3QpWzBdID8/XG4gICAgbnVsbFxuICApO1xufVxuXG4vKiogT3V0bG9vayBzdGFydHMgYW4gdW5yZWFkIHJvdydzIGFjY2Vzc2libGUgbmFtZSB3aXRoIFwiVW5yZWFkXCIgKEVuZ2xpc2ggVUkpLiAqL1xuZXhwb3J0IGZ1bmN0aW9uIGlzVW5yZWFkUm93KHJvdzogRWxlbWVudCk6IGJvb2xlYW4ge1xuICByZXR1cm4gL15VbnJlYWRcXGIvLnRlc3Qocm93LmdldEF0dHJpYnV0ZShcImFyaWEtbGFiZWxcIikgPz8gXCJcIik7XG59XG5cbi8qKiBPdXRsb29rJ3MgY3VycmVudCBsaWdodC9kYXJrIG1vZGUsIGZyb20gdGhlIHBhbGV0dGUgdmFyaWFibGVzIGl0IHdyaXRlcyBvbiA8aHRtbD4uICovXG5leHBvcnQgZnVuY3Rpb24gb3V0bG9va0NvbG9yU2NoZW1lKGh0bWw6IEVsZW1lbnQpOiBcImxpZ2h0XCIgfCBcImRhcmtcIiB7XG4gIGNvbnN0IHN0eWxlID0gaHRtbC5nZXRBdHRyaWJ1dGUoXCJzdHlsZVwiKSA/PyBcIlwiO1xuICByZXR1cm4gLy0tbmF0aXZlLXNjcm9sbGJhci1jb2xvci1zY2hlbWU6XFxzKmRhcmsvLnRlc3Qoc3R5bGUpID8gXCJkYXJrXCIgOiBcImxpZ2h0XCI7XG59XG5cbi8qKiBFbnRyaWVzIGV4cGVjdGVkIG9uIGV2ZXJ5IG1haWwgdmlldyB0aGF0IGRpZCBub3QgcmVzb2x2ZS4gTm9uLWVtcHR5IG1lYW5zIHNvbWV0aGluZyBicm9rZS4gKi9cbmV4cG9ydCBmdW5jdGlvbiBicm9rZW5FbnRyaWVzKHJlc3VsdHM6IEhlYWx0aFJlc3VsdFtdKTogU2VsZWN0b3JOYW1lW10ge1xuICByZXR1cm4gcmVzdWx0cy5maWx0ZXIoKHIpID0+IHIuZXhwZWN0ID09PSBcImFsd2F5c1wiICYmIHIuY2FuZGlkYXRlIDwgMCkubWFwKChyKSA9PiByLm5hbWUpO1xufVxuIiwgIi8vIERldmVsb3BlciBjYXB0dXJlOiBzZXJpYWxpc2VzIHRoZSBsaXZlIE91dGxvb2sgcGFnZSBpbnRvIG9uZSBIVE1MIGZpbGUsIHNvIGEgZml4dHVyZSBjYW4gYmVcbi8vIG1hZGUgKGRvY3MvY2FwdHVyaW5nLXBhZ2VzLm1kKSB3aXRob3V0IERldlRvb2xzLiBTYXZlZCBsb2NhbGx5IGJ5IHRoZSBjYWxsZXI7IG5ldmVyIHNlbnRcbi8vIGFueXdoZXJlLiBUaGUgZmlsZSBob2xkcyByZWFsIG1haWwsIHNvIGl0IGdvZXMgdGhyb3VnaCBzY3JpcHRzL3NjcnViLWZpeHR1cmUudHMgYmVmb3JlIGl0XG4vLyBnb2VzIGFueXdoZXJlIG5lYXIgdGhlIHJlcG8uXG5cbmltcG9ydCB7IGJyb2tlbkVudHJpZXMsIGhlYWx0aENoZWNrIH0gZnJvbSBcIi4vc2VsZWN0b3JzLnRzXCI7XG5cbi8qKiBNZXNzYWdlIHRoZSBwb3B1cCBzZW5kcyB0byB0aGUgY29udGVudCBzY3JpcHQgb2YgdGhlIGFjdGl2ZSB0YWIuICovXG5leHBvcnQgY29uc3QgQ0FQVFVSRV9NRVNTQUdFID0gXCJvYmItY2FwdHVyZVwiO1xuXG5jb25zdCBPQkJfQVRUUklCVVRFID0gL15kYXRhLW92ZXJsb29rLztcbi8qKiBUaGUgc2NydWJiZXIgKHNjcmlwdHMvc2NydWItZml4dHVyZS50cykgbG9va3MgZm9yIHRoaXMgbWFya2VyIHRvIHNwbGl0IG91dCBydW50aW1lIHN0eWxlcy4gKi9cbmV4cG9ydCBjb25zdCBSVU5USU1FX1NUWUxFX01BUktFUiA9IFwiZGF0YS1vdmVybG9vay1ydW50aW1lXCI7XG5cbmV4cG9ydCBpbnRlcmZhY2UgQ2FwdHVyZU1ldGEge1xuICBjYXB0dXJlZEF0OiBzdHJpbmc7XG4gIG9iYlZlcnNpb246IHN0cmluZztcbiAgLyoqIEZpcnN0IHR3byBwYXRoIHNlZ21lbnRzIG9ubHk6IGRlZXBlciBvbmVzIGNhcnJ5IG1lc3NhZ2UgaWRzLiAqL1xuICByb3V0ZTogc3RyaW5nO1xuICB2aWV3cG9ydDogc3RyaW5nO1xuICBzZXR0aW5nczogUmVjb3JkPHN0cmluZywgc3RyaW5nIHwgdW5kZWZpbmVkPjtcbiAgaGVhbHRoOiB7IGJyb2tlbjogc3RyaW5nW107IGVudHJpZXM6IHsgbmFtZTogc3RyaW5nOyBjYW5kaWRhdGU6IG51bWJlcjsgY291bnQ6IG51bWJlciB9W10gfTtcbn1cblxuZXhwb3J0IGludGVyZmFjZSBDYXB0dXJlIHtcbiAgaHRtbDogc3RyaW5nO1xuICBuYW1lOiBzdHJpbmc7XG4gIG1ldGE6IENhcHR1cmVNZXRhO1xufVxuXG4vKiogU3R5bGUgcnVsZXMgT3V0bG9vayBpbmplY3RlZCBhdCBydW50aW1lIChDU1MtaW4tSlMpLCBtaW51cyB0aGUgb25lcyBPQkIgaXRzZWxmIGFkZGVkLiAqL1xuZnVuY3Rpb24gcnVudGltZUNzcyhkb2M6IERvY3VtZW50KTogc3RyaW5nIHtcbiAgY29uc3Qgc2hlZXRzID0gWy4uLmRvYy5zdHlsZVNoZWV0cywgLi4uKGRvYy5hZG9wdGVkU3R5bGVTaGVldHMgPz8gW10pXS5maWx0ZXIoXG4gICAgKHNoZWV0KSA9PiAhKFwiaHJlZlwiIGluIHNoZWV0ICYmIHNoZWV0LmhyZWYpLFxuICApO1xuICByZXR1cm4gc2hlZXRzXG4gICAgLmZsYXRNYXAoKHNoZWV0KSA9PiB7XG4gICAgICB0cnkge1xuICAgICAgICByZXR1cm4gWy4uLnNoZWV0LmNzc1J1bGVzXS5tYXAoKHJ1bGUpID0+IHJ1bGUuY3NzVGV4dCk7XG4gICAgICB9IGNhdGNoIHtcbiAgICAgICAgcmV0dXJuIFtdOyAvLyBjcm9zcy1vcmlnaW4gc2hlZXRcbiAgICAgIH1cbiAgICB9KVxuICAgIC5maWx0ZXIoKHJ1bGUpID0+ICFydWxlLmluY2x1ZGVzKFwiZGF0YS1vdmVybG9va1wiKSlcbiAgICAuam9pbihcIlxcblwiKTtcbn1cblxuZXhwb3J0IGZ1bmN0aW9uIGNhcHR1cmVGaWxlbmFtZShub3cgPSBuZXcgRGF0ZSgpKTogc3RyaW5nIHtcbiAgY29uc3Qgc3RhbXAgPSBub3cudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxNikucmVwbGFjZSgvXFxEL2csIFwiXCIpO1xuICByZXR1cm4gYG9iYi1jYXB0dXJlLSR7c3RhbXB9Lmh0bWxgO1xufVxuXG5leHBvcnQgZnVuY3Rpb24gY2FwdHVyZVBhZ2UoXG4gIGRvYzogRG9jdW1lbnQgPSBkb2N1bWVudCxcbiAgbG9jOiBQaWNrPExvY2F0aW9uLCBcInBhdGhuYW1lXCI+ID0gbG9jYXRpb24sXG4gIHZlcnNpb24gPSBcImRldlwiLFxuICBub3cgPSBuZXcgRGF0ZSgpLFxuKTogQ2FwdHVyZSB7XG4gIGNvbnN0IGxpdmUgPSBkb2MuZG9jdW1lbnRFbGVtZW50LmRhdGFzZXQ7XG4gIGNvbnN0IHJlc3VsdHMgPSBoZWFsdGhDaGVjayhkb2MpO1xuICBjb25zdCBtZXRhOiBDYXB0dXJlTWV0YSA9IHtcbiAgICBjYXB0dXJlZEF0OiBub3cudG9JU09TdHJpbmcoKSxcbiAgICBvYmJWZXJzaW9uOiB2ZXJzaW9uLFxuICAgIHJvdXRlOiBsb2MucGF0aG5hbWUuc3BsaXQoXCIvXCIpLnNsaWNlKDAsIDMpLmpvaW4oXCIvXCIpLFxuICAgIHZpZXdwb3J0OiBgJHtkb2MuZGVmYXVsdFZpZXc/LmlubmVyV2lkdGggPz8gMH14JHtkb2MuZGVmYXVsdFZpZXc/LmlubmVySGVpZ2h0ID8/IDB9YCxcbiAgICBzZXR0aW5nczoge1xuICAgICAgZmVhdHVyZXM6IGxpdmUub3Zlcmxvb2tGZWF0dXJlcyxcbiAgICAgIHBhbGV0dGU6IGxpdmUub3Zlcmxvb2tQYWxldHRlLFxuICAgICAgbW9kZTogbGl2ZS5vdmVybG9va01vZGUsXG4gICAgICBsYXlvdXQ6IGxpdmUub3Zlcmxvb2tMYXlvdXQsXG4gICAgICBvdXRsb29rU2NoZW1lOiBsaXZlLm92ZXJsb29rT3V0bG9va1NjaGVtZSxcbiAgICAgIHJlYWRpbmdQYW5lOiBsaXZlLm92ZXJsb29rUmVhZGluZ1BhbmUsXG4gICAgfSxcbiAgICBoZWFsdGg6IHtcbiAgICAgIGJyb2tlbjogYnJva2VuRW50cmllcyhyZXN1bHRzKSxcbiAgICAgIGVudHJpZXM6IHJlc3VsdHMubWFwKCh7IG5hbWUsIGNhbmRpZGF0ZSwgY291bnQgfSkgPT4gKHsgbmFtZSwgY2FuZGlkYXRlLCBjb3VudCB9KSksXG4gICAgfSxcbiAgfTtcblxuICAvLyBXb3JrIG9uIGEgY29weSwgd2l0aCBPQkIncyBvd24gbWFya3MgcmVtb3ZlZCwgc28gdGhlIGNhcHR1cmUgaXMgc3RvY2sgT3V0bG9vay5cbiAgY29uc3Qgcm9vdCA9IGRvYy5kb2N1bWVudEVsZW1lbnQuY2xvbmVOb2RlKHRydWUpIGFzIEhUTUxFbGVtZW50O1xuICBmb3IgKGNvbnN0IGVsIG9mIHJvb3QucXVlcnlTZWxlY3RvckFsbChcIm92ZXJsb29rLXJvb3RcIikpIGVsLnJlbW92ZSgpO1xuICBmb3IgKGNvbnN0IGVsIG9mIFtyb290LCAuLi5yb290LnF1ZXJ5U2VsZWN0b3JBbGwoXCIqXCIpXSkge1xuICAgIGZvciAoY29uc3QgeyBuYW1lIH0gb2YgWy4uLmVsLmF0dHJpYnV0ZXNdKSB7XG4gICAgICBpZiAoT0JCX0FUVFJJQlVURS50ZXN0KG5hbWUpKSBlbC5yZW1vdmVBdHRyaWJ1dGUobmFtZSk7XG4gICAgfVxuICB9XG5cbiAgY29uc3QgaGVhZCA9IHJvb3QucXVlcnlTZWxlY3RvcihcImhlYWRcIik7XG4gIGNvbnN0IHN0eWxlID0gZG9jLmNyZWF0ZUVsZW1lbnQoXCJzdHlsZVwiKTtcbiAgc3R5bGUuc2V0QXR0cmlidXRlKFJVTlRJTUVfU1RZTEVfTUFSS0VSLCBcIlwiKTtcbiAgc3R5bGUudGV4dENvbnRlbnQgPSBydW50aW1lQ3NzKGRvYyk7XG4gIGhlYWQ/LmFwcGVuZChzdHlsZSk7XG4gIGNvbnN0IHNjcmlwdCA9IGRvYy5jcmVhdGVFbGVtZW50KFwic2NyaXB0XCIpO1xuICBzY3JpcHQudHlwZSA9IFwiYXBwbGljYXRpb24vanNvblwiO1xuICBzY3JpcHQuaWQgPSBcIm9iYi1jYXB0dXJlLW1ldGFcIjtcbiAgc2NyaXB0LnRleHRDb250ZW50ID0gSlNPTi5zdHJpbmdpZnkobWV0YSwgbnVsbCwgMikucmVwbGFjZUFsbChcIjxcIiwgXCJcXFxcdTAwM2NcIik7XG4gIGhlYWQ/LmFwcGVuZChzY3JpcHQpO1xuXG4gIHJldHVybiB7IGh0bWw6IGA8IWRvY3R5cGUgaHRtbD5cXG4ke3Jvb3Qub3V0ZXJIVE1MfWAsIG5hbWU6IGNhcHR1cmVGaWxlbmFtZShub3cpLCBtZXRhIH07XG59XG4iLCAiLy8gVXNlciBwcmVmZXJlbmNlcy4gU3RvcmVkIGluIGNocm9tZS5zdG9yYWdlLnN5bmM7IG5vdGhpbmcgZWxzZSBpcyBldmVyIHBlcnNpc3RlZC5cbi8vXG4vLyBFYWNoIGVuYWJsZWQgZmVhdHVyZSBpZCBpcyB3cml0dGVuIGFzIGEgdG9rZW4gb24gPGh0bWwgZGF0YS1vdmVybG9vay1mZWF0dXJlcz1cIi4uLlwiPiwgYW5kIGVhY2hcbi8vIGNob2ljZSBhcyBpdHMgb3duIGF0dHJpYnV0ZSAoPGh0bWwgZGF0YS1vdmVybG9vay1wYWxldHRlPVwic29sYXJpemVkXCI+KSwgc28gc3R5bGVzaGVldHMgZ2F0ZVxuLy8gdGhlbXNlbHZlcyB3aXRoIGBbZGF0YS1vdmVybG9vay1mZWF0dXJlc349XCJkZW5zaXR5XCJdYCBvciBgW2RhdGEtb3Zlcmxvb2stcGFsZXR0ZT1cInNvbGFyaXplZFwiXWAuXG5cbmV4cG9ydCB0eXBlIEZlYXR1cmVJZCA9IGtleW9mIHR5cGVvZiBGRUFUVVJFUztcbmV4cG9ydCB0eXBlIENob2ljZUlkID0ga2V5b2YgdHlwZW9mIENIT0lDRVM7XG5leHBvcnQgdHlwZSBTZXR0aW5ncyA9IFJlY29yZDxGZWF0dXJlSWQsIGJvb2xlYW4+ICYgeyBbSyBpbiBDaG9pY2VJZF06IENob2ljZVZhbHVlPEs+IH07XG5leHBvcnQgdHlwZSBDaG9pY2VWYWx1ZTxLIGV4dGVuZHMgQ2hvaWNlSWQ+ID0ga2V5b2YgKHR5cGVvZiBDSE9JQ0VTKVtLXVtcIm9wdGlvbnNcIl07XG5cbnR5cGUgR3JvdXAgPSBcIkdlbmVyYWxcIiB8IFwiS2V5Ym9hcmRcIiB8IFwiRGVjbHV0dGVyXCIgfCBcIkxvb2tcIjtcblxuaW50ZXJmYWNlIEZlYXR1cmUge1xuICBsYWJlbDogc3RyaW5nO1xuICBncm91cDogR3JvdXA7XG4gIGRlZmF1bHQ6IGJvb2xlYW47XG59XG5cbmludGVyZmFjZSBDaG9pY2Uge1xuICBsYWJlbDogc3RyaW5nO1xuICBncm91cDogR3JvdXA7XG4gIG9wdGlvbnM6IFJlY29yZDxzdHJpbmcsIHN0cmluZz47XG4gIGRlZmF1bHQ6IHN0cmluZztcbn1cblxuZXhwb3J0IGNvbnN0IEZFQVRVUkVTID0ge1xuICBoZWFsdGhJbmRpY2F0b3I6IHtcbiAgICBsYWJlbDogXCJTaG93IGFuIGluZGljYXRvciB3aGVuIE91dGxvb2sncyBwYWdlIGNoYW5nZXMgYnJlYWsgT0JCXCIsXG4gICAgZ3JvdXA6IFwiR2VuZXJhbFwiLFxuICAgIGRlZmF1bHQ6IHRydWUsXG4gIH0sXG4gIGtleW1hcDogeyBsYWJlbDogXCJGYXN0bWFpbC1zdHlsZSBrZXlib2FyZCBzaG9ydGN1dHNcIiwgZ3JvdXA6IFwiS2V5Ym9hcmRcIiwgZGVmYXVsdDogdHJ1ZSB9LFxuICBtdWx0aVNlbGVjdDoge1xuICAgIGxhYmVsOiBcInggdGlja3MgYSBtZXNzYWdlOyBqIGFuZCBrIG1vdmUgd2l0aG91dCBsb3NpbmcgdGlja3NcIixcbiAgICBncm91cDogXCJLZXlib2FyZFwiLFxuICAgIGRlZmF1bHQ6IHRydWUsXG4gIH0sXG4gIFwiaGlkZS11cHNlbGxcIjogeyBsYWJlbDogXCJVcGdyYWRlIGFuZCAndHJ5IHRoZSBuZXdcdTIwMjYnIGJhbm5lcnNcIiwgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsIGRlZmF1bHQ6IHRydWUgfSxcbiAgXCJoaWRlLWFwcC1iYW5uZXJzXCI6IHtcbiAgICBsYWJlbDogXCJOb3RpY2VzIGFjcm9zcyB0aGUgdG9wIG9mIE91dGxvb2sgKHBvcC11cCwgc2lnbi1pbiBhbmQgc3RvcmFnZSByZW1pbmRlcnMpXCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogZmFsc2UsXG4gIH0sXG4gIFwiaGlkZS1pbmZvLWJhcnNcIjoge1xuICAgIGxhYmVsOiBcIk5vdGljZXMgYWJvdmUgYW4gb3BlbiBtZXNzYWdlICgnWW91IHJlcGxpZWQgb25cdTIwMjYnLCAnRXh0cmEgbGluZSBicmVha3MgcmVtb3ZlZCcpXCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogZmFsc2UsXG4gIH0sXG4gIFwiaGlkZS1jb3BpbG90XCI6IHsgbGFiZWw6IFwiQ29waWxvdCBidXR0b25zIGFuZCBwYW5lbHNcIiwgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsIGRlZmF1bHQ6IHRydWUgfSxcbiAgXCJoaWRlLWhlYWRlci1leHRyYXNcIjoge1xuICAgIGxhYmVsOiBcIlRlYW1zIGFuZCBNeSBEYXkgYnV0dG9ucyBpbiB0aGUgaGVhZGVyXCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogdHJ1ZSxcbiAgfSxcbiAgXCJoaWRlLXJhaWwtZXh0cmFzXCI6IHtcbiAgICBsYWJlbDogXCJUbyBEbywgTmV3c2xldHRlcnMsIE9yZyBFeHBsb3JlciBhbmQgT25lRHJpdmUgaW4gdGhlIGFwcCByYWlsXCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogdHJ1ZSxcbiAgfSxcbiAgXCJoaWRlLXJpYmJvbi1leHRyYXNcIjoge1xuICAgIGxhYmVsOiBcIlRlYW1zLCBTd2VlcCwgUXVpY2sgc3RlcHMsIFZpdmEgYW5kIGFkZC1pbiBidXR0b25zIGluIHRoZSByaWJib25cIixcbiAgICBncm91cDogXCJEZWNsdXR0ZXJcIixcbiAgICBkZWZhdWx0OiB0cnVlLFxuICB9LFxuICBcImhpZGUtZXh0ZXJuYWwtdGFnXCI6IHsgbGFiZWw6IFwiJ0V4dGVybmFsJyBzZW5kZXIgdGFnc1wiLCBncm91cDogXCJEZWNsdXR0ZXJcIiwgZGVmYXVsdDogdHJ1ZSB9LFxuICBcImhpZGUtcm93LWFjdGlvbnNcIjoge1xuICAgIGxhYmVsOiBcIkZsYWcgYW5kIHBpbiBpY29ucyB0aGF0IGFwcGVhciB3aGVuIGhvdmVyaW5nIGEgbWVzc2FnZVwiLFxuICAgIGdyb3VwOiBcIkRlY2x1dHRlclwiLFxuICAgIGRlZmF1bHQ6IGZhbHNlLFxuICB9LFxuICBcImhpZGUtYXZhdGFyc1wiOiB7XG4gICAgbGFiZWw6IFwiU2VuZGVyIHBpY3R1cmVzIGluIHRoZSBtZXNzYWdlIGxpc3RcIixcbiAgICBncm91cDogXCJEZWNsdXR0ZXJcIixcbiAgICBkZWZhdWx0OiBmYWxzZSxcbiAgfSxcbiAgXCJoaWRlLWZvY3VzZWQtdGFic1wiOiB7IGxhYmVsOiBcIkZvY3VzZWQgLyBPdGhlciB0YWJzXCIsIGdyb3VwOiBcIkRlY2x1dHRlclwiLCBkZWZhdWx0OiBmYWxzZSB9LFxuICBcImhpZGUtcmliYm9uLXRhYnNcIjoge1xuICAgIGxhYmVsOiBcIlJpYmJvbiB0YWJzIChIb21lLCBWaWV3LCBIZWxwKVwiLFxuICAgIGdyb3VwOiBcIkRlY2x1dHRlclwiLFxuICAgIGRlZmF1bHQ6IGZhbHNlLFxuICB9LFxuICBcImhpZGUtYXBwLXJhaWxcIjogeyBsYWJlbDogXCJBcHAgcmFpbCAobGVmdCBpY29uIGJhcilcIiwgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsIGRlZmF1bHQ6IGZhbHNlIH0sXG4gIFwiaGlkZS1yaWJib25cIjogeyBsYWJlbDogXCJSaWJib25cIiwgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsIGRlZmF1bHQ6IGZhbHNlIH0sXG4gIGRlbnNpdHk6IHsgbGFiZWw6IFwiVGlnaHRlciwgZGVuc2VyIGxheW91dFwiLCBncm91cDogXCJMb29rXCIsIGRlZmF1bHQ6IHRydWUgfSxcbn0gYXMgY29uc3Qgc2F0aXNmaWVzIFJlY29yZDxzdHJpbmcsIEZlYXR1cmU+O1xuXG5leHBvcnQgY29uc3QgQ0hPSUNFUyA9IHtcbiAgcGFsZXR0ZToge1xuICAgIGxhYmVsOiBcIkNvbG9yc1wiLFxuICAgIGdyb3VwOiBcIkxvb2tcIixcbiAgICBvcHRpb25zOiB7XG4gICAgICBvdXRsb29rOiBcIk91dGxvb2sncyBvd25cIixcbiAgICAgIHNvbGFyaXplZDogXCJTb2xhcml6ZWRcIixcbiAgICAgIGJhc2UxNjogXCJCYXNlMTYgRGVmYXVsdFwiLFxuICAgICAgZmFzdG1haWw6IFwiRmFzdG1haWxcIixcbiAgICAgIHBhcGVyOiBcIlBhcGVyXCIsXG4gICAgfSxcbiAgICBkZWZhdWx0OiBcIm91dGxvb2tcIixcbiAgfSxcbiAgbW9kZToge1xuICAgIGxhYmVsOiBcIkxpZ2h0IG9yIGRhcmtcIixcbiAgICBncm91cDogXCJMb29rXCIsXG4gICAgb3B0aW9uczogeyBhdXRvOiBcIkZvbGxvdyBPdXRsb29rXCIsIGxpZ2h0OiBcIkxpZ2h0XCIsIGRhcms6IFwiRGFya1wiIH0sXG4gICAgZGVmYXVsdDogXCJhdXRvXCIsXG4gIH0sXG4gIGxheW91dDoge1xuICAgIGxhYmVsOiBcIkxheW91dFwiLFxuICAgIGdyb3VwOiBcIkxvb2tcIixcbiAgICBvcHRpb25zOiB7XG4gICAgICBvdXRsb29rOiBcIk91dGxvb2sncyBvd25cIixcbiAgICAgIGNhbG06IFwiQ2FsbSB0aHJlZS1wYW5lXCIsXG4gICAgICBmYXN0bWFpbDogXCJGYXN0bWFpbCAobGlzdCBmaXJzdClcIixcbiAgICAgIGZvY3VzOiBcIkZvY3VzIChzaW5nbGUgY29sdW1uKVwiLFxuICAgIH0sXG4gICAgZGVmYXVsdDogXCJvdXRsb29rXCIsXG4gIH0sXG59IGFzIGNvbnN0IHNhdGlzZmllcyBSZWNvcmQ8c3RyaW5nLCBDaG9pY2U+O1xuXG5leHBvcnQgY29uc3QgRkVBVFVSRV9JRFMgPSBPYmplY3Qua2V5cyhGRUFUVVJFUykgYXMgRmVhdHVyZUlkW107XG5leHBvcnQgY29uc3QgQ0hPSUNFX0lEUyA9IE9iamVjdC5rZXlzKENIT0lDRVMpIGFzIENob2ljZUlkW107XG5cbmV4cG9ydCBmdW5jdGlvbiBkZWZhdWx0U2V0dGluZ3MoKTogU2V0dGluZ3Mge1xuICByZXR1cm4ge1xuICAgIC4uLk9iamVjdC5mcm9tRW50cmllcyhGRUFUVVJFX0lEUy5tYXAoKGlkKSA9PiBbaWQsIEZFQVRVUkVTW2lkXS5kZWZhdWx0XSkpLFxuICAgIC4uLk9iamVjdC5mcm9tRW50cmllcyhDSE9JQ0VfSURTLm1hcCgoaWQpID0+IFtpZCwgQ0hPSUNFU1tpZF0uZGVmYXVsdF0pKSxcbiAgfSBhcyBTZXR0aW5ncztcbn1cblxuLyoqIE1lcmdlIHdoYXRldmVyIGlzIHN0b3JlZCBvdmVyIHRoZSBkZWZhdWx0cywgaWdub3JpbmcgdW5rbm93biBvciBtaXN0eXBlZCBrZXlzLiAqL1xuZXhwb3J0IGZ1bmN0aW9uIG5vcm1hbGl6ZShzdG9yZWQ6IFJlY29yZDxzdHJpbmcsIHVua25vd24+KTogU2V0dGluZ3Mge1xuICBjb25zdCBzZXR0aW5ncyA9IGRlZmF1bHRTZXR0aW5ncygpIGFzIFJlY29yZDxzdHJpbmcsIHVua25vd24+O1xuICBmb3IgKGNvbnN0IGlkIG9mIEZFQVRVUkVfSURTKSB7XG4gICAgaWYgKHR5cGVvZiBzdG9yZWRbaWRdID09PSBcImJvb2xlYW5cIikgc2V0dGluZ3NbaWRdID0gc3RvcmVkW2lkXTtcbiAgfVxuICBmb3IgKGNvbnN0IGlkIG9mIENIT0lDRV9JRFMpIHtcbiAgICBjb25zdCB2YWx1ZSA9IHN0b3JlZFtpZF07XG4gICAgaWYgKHR5cGVvZiB2YWx1ZSA9PT0gXCJzdHJpbmdcIiAmJiB2YWx1ZSBpbiBDSE9JQ0VTW2lkXS5vcHRpb25zKSBzZXR0aW5nc1tpZF0gPSB2YWx1ZTtcbiAgfVxuICByZXR1cm4gc2V0dGluZ3MgYXMgU2V0dGluZ3M7XG59XG5cbmV4cG9ydCBmdW5jdGlvbiBmZWF0dXJlVG9rZW5zKHNldHRpbmdzOiBTZXR0aW5ncyk6IHN0cmluZyB7XG4gIHJldHVybiBGRUFUVVJFX0lEUy5maWx0ZXIoKGlkKSA9PiBzZXR0aW5nc1tpZF0pLmpvaW4oXCIgXCIpO1xufVxuXG4vKiogV3JpdGVzIHRoZSBzZXR0aW5ncyBvbnRvIDxodG1sPiBhcyB0aGUgYXR0cmlidXRlcyB0aGUgc3R5bGVzaGVldHMga2V5IG9mZi4gKi9cbmV4cG9ydCBmdW5jdGlvbiBhcHBseVRvRG9jdW1lbnQoc2V0dGluZ3M6IFNldHRpbmdzLCByb290OiBIVE1MRWxlbWVudCk6IHZvaWQge1xuICByb290LnNldEF0dHJpYnV0ZShcImRhdGEtb3Zlcmxvb2stZmVhdHVyZXNcIiwgZmVhdHVyZVRva2VucyhzZXR0aW5ncykpO1xuICBmb3IgKGNvbnN0IGlkIG9mIENIT0lDRV9JRFMpIHJvb3Quc2V0QXR0cmlidXRlKGBkYXRhLW92ZXJsb29rLSR7aWR9YCwgc2V0dGluZ3NbaWRdKTtcbn1cblxuZXhwb3J0IGFzeW5jIGZ1bmN0aW9uIGxvYWRTZXR0aW5ncygpOiBQcm9taXNlPFNldHRpbmdzPiB7XG4gIHJldHVybiBub3JtYWxpemUoYXdhaXQgY2hyb21lLnN0b3JhZ2Uuc3luYy5nZXQobnVsbCkpO1xufVxuXG5leHBvcnQgYXN5bmMgZnVuY3Rpb24gc2F2ZVNldHRpbmcoXG4gIGlkOiBGZWF0dXJlSWQgfCBDaG9pY2VJZCxcbiAgdmFsdWU6IGJvb2xlYW4gfCBzdHJpbmcsXG4pOiBQcm9taXNlPHZvaWQ+IHtcbiAgYXdhaXQgY2hyb21lLnN0b3JhZ2Uuc3luYy5zZXQoeyBbaWRdOiB2YWx1ZSB9KTtcbn1cblxuZXhwb3J0IGZ1bmN0aW9uIG9uU2V0dGluZ3NDaGFuZ2VkKGNhbGxiYWNrOiAoc2V0dGluZ3M6IFNldHRpbmdzKSA9PiB2b2lkKTogdm9pZCB7XG4gIGNocm9tZS5zdG9yYWdlLm9uQ2hhbmdlZC5hZGRMaXN0ZW5lcigoX2NoYW5nZXMsIGFyZWEpID0+IHtcbiAgICBpZiAoYXJlYSA9PT0gXCJzeW5jXCIpIHZvaWQgbG9hZFNldHRpbmdzKCkudGhlbihjYWxsYmFjayk7XG4gIH0pO1xufVxuIiwgIi8vIFRoZSBvcHRpb25zIHBhZ2UsIGFsc28gdGhlIHRvb2xiYXIgcG9wdXAuIENoYW5nZXMgc2F2ZSBpbW1lZGlhdGVseSBhbmQgYXBwbHkgbGl2ZSBpbiBPdXRsb29rLlxuXG5pbXBvcnQgeyBDQVBUVVJFX01FU1NBR0UgfSBmcm9tIFwiLi4vY2FwdHVyZS50c1wiO1xuaW1wb3J0IHtcbiAgQ0hPSUNFUyxcbiAgdHlwZSBDaG9pY2VJZCxcbiAgRkVBVFVSRV9JRFMsXG4gIEZFQVRVUkVTLFxuICB0eXBlIEZlYXR1cmVJZCxcbiAgbG9hZFNldHRpbmdzLFxuICBvblNldHRpbmdzQ2hhbmdlZCxcbiAgdHlwZSBTZXR0aW5ncyxcbiAgc2F2ZVNldHRpbmcsXG59IGZyb20gXCIuLi9zZXR0aW5ncy50c1wiO1xuXG5jb25zdCBTSE9SVF9MQVlPVVRfTkFNRVM6IFJlY29yZDxzdHJpbmcsIHN0cmluZz4gPSB7XG4gIG91dGxvb2s6IFwiT3V0bG9va1wiLFxuICBjYWxtOiBcIkNhbG1cIixcbiAgZmFzdG1haWw6IFwiRmFzdG1haWxcIixcbiAgZm9jdXM6IFwiRm9jdXNcIixcbn07XG5cbmNvbnN0IFNIT1JUX1BBTEVUVEVfTkFNRVM6IFJlY29yZDxzdHJpbmcsIHN0cmluZz4gPSB7XG4gIG91dGxvb2s6IFwiT3V0bG9va1wiLFxuICBzb2xhcml6ZWQ6IFwiU29sYXJpemVkXCIsXG4gIGJhc2UxNjogXCJCYXNlMTZcIixcbiAgZmFzdG1haWw6IFwiRmFzdG1haWxcIixcbiAgcGFwZXI6IFwiUGFwZXJcIixcbn07XG5cbi8qKiBUd28gbGluZXMgb2YgYSBtZXNzYWdlIHJvdyAoc3ViamVjdCwgcHJldmlldykgYXQgdGhlIGdpdmVuIHBvc2l0aW9uIGluIGEgc2tldGNoLiAqL1xuZnVuY3Rpb24gcm93cyh4OiBudW1iZXIsIHk6IG51bWJlciwgd2lkdGg6IG51bWJlcik6IHN0cmluZyB7XG4gIHJldHVybiBbMCwgMSwgMiwgMywgNF1cbiAgICAubWFwKChpKSA9PiB7XG4gICAgICBjb25zdCB0b3AgPSB5ICsgaSAqIDYuNTtcbiAgICAgIGNvbnN0IHN1YmplY3QgPSBpID09PSAwID8gXCJoaVwiIDogXCJyb3dcIjtcbiAgICAgIGNvbnN0IHN1YmplY3RXaWR0aCA9IGkgPT09IDAgPyB3aWR0aCAqIDAuNTUgOiB3aWR0aDtcbiAgICAgIHJldHVybiAoXG4gICAgICAgIGA8cmVjdCBjbGFzcz1cIiR7c3ViamVjdH1cIiB4PVwiJHt4fVwiIHk9XCIke3RvcH1cIiB3aWR0aD1cIiR7c3ViamVjdFdpZHRofVwiIGhlaWdodD1cIjIuNVwiIHJ4PVwiMVwiLz5gICtcbiAgICAgICAgYDxyZWN0IGNsYXNzPVwicm93XCIgeD1cIiR7eH1cIiB5PVwiJHt0b3AgKyAzfVwiIHdpZHRoPVwiJHt3aWR0aCAqIDAuNzV9XCIgaGVpZ2h0PVwiMS41XCIgcng9XCIwLjc1XCIvPmBcbiAgICAgICk7XG4gICAgfSlcbiAgICAuam9pbihcIlwiKTtcbn1cblxuLy8gV2lyZWZyYW1lcyBvZiBlYWNoIGxheW91dDogdmlld0JveCA2NFx1MDBENzQ0OyBjaHJvbWUgYmFyLCB0aGVuIHBhbmVzIGFuZCBtZXNzYWdlIHJvd3MuXG5jb25zdCBMQVlPVVRfU0tFVENIRVM6IFJlY29yZDxzdHJpbmcsIHN0cmluZz4gPSB7XG4gIG91dGxvb2s6IGA8cmVjdCBjbGFzcz1cImNocm9tZVwiIHdpZHRoPVwiNjRcIiBoZWlnaHQ9XCI1XCIvPlxuICAgIDxyZWN0IGNsYXNzPVwicGFuZVwiIHg9XCIwXCIgeT1cIjVcIiB3aWR0aD1cIjVcIiBoZWlnaHQ9XCIzOVwiLz5cbiAgICA8cmVjdCBjbGFzcz1cInBhbmVcIiB4PVwiNlwiIHk9XCI3XCIgd2lkdGg9XCIxMlwiIGhlaWdodD1cIjM1XCIgcng9XCIxXCIvPlxuICAgICR7cm93cygyMCwgOCwgMTUpfTxyZWN0IGNsYXNzPVwicGFuZVwiIHg9XCIzN1wiIHk9XCI3XCIgd2lkdGg9XCIyNVwiIGhlaWdodD1cIjM1XCIgcng9XCIxXCIvPmAsXG4gIGNhbG06IGA8cmVjdCBjbGFzcz1cImNocm9tZVwiIHdpZHRoPVwiNjRcIiBoZWlnaHQ9XCI1XCIvPlxuICAgIDxyZWN0IGNsYXNzPVwicGFuZVwiIHg9XCIyXCIgeT1cIjhcIiB3aWR0aD1cIjEzXCIgaGVpZ2h0PVwiMzNcIiByeD1cIjJcIi8+XG4gICAgJHtyb3dzKDE4LCA5LCAxNil9PHJlY3QgY2xhc3M9XCJwYW5lXCIgeD1cIjM3XCIgeT1cIjhcIiB3aWR0aD1cIjI1XCIgaGVpZ2h0PVwiMzNcIiByeD1cIjJcIi8+YCxcbiAgZmFzdG1haWw6IGA8cmVjdCBjbGFzcz1cImNocm9tZVwiIHdpZHRoPVwiNjRcIiBoZWlnaHQ9XCI1XCIvPlxuICAgIDxyZWN0IGNsYXNzPVwicGFuZVwiIHg9XCIyXCIgeT1cIjhcIiB3aWR0aD1cIjEzXCIgaGVpZ2h0PVwiMzNcIiByeD1cIjJcIi8+XG4gICAgPHJlY3QgY2xhc3M9XCJoaVwiIHg9XCI0XCIgeT1cIjEwXCIgd2lkdGg9XCI5XCIgaGVpZ2h0PVwiM1wiIHJ4PVwiMS41XCIvPiR7cm93cygxOCwgOSwgNDQpfWAsXG4gIGZvY3VzOiBgPHJlY3QgY2xhc3M9XCJjaHJvbWVcIiB3aWR0aD1cIjY0XCIgaGVpZ2h0PVwiNVwiLz4ke3Jvd3MoMTYsIDksIDMyKX1gLFxufTtcblxuY29uc3QgaHRtbCA9IGRvY3VtZW50LmRvY3VtZW50RWxlbWVudDtcbmNvbnN0IHByZWZlcnNEYXJrID0gbWF0Y2hNZWRpYShcIihwcmVmZXJzLWNvbG9yLXNjaGVtZTogZGFyaylcIik7XG5cbi8qKiBUaGUgbGlnaHQvZGFyayBtb2RlIGluIGVmZmVjdDogXCJhdXRvXCIgZm9sbG93cyB0aGUgc3lzdGVtIGhlcmUsIHN0YW5kaW5nIGluIGZvciBPdXRsb29rLiAqL1xuZnVuY3Rpb24gZWZmZWN0aXZlTW9kZShzZXR0aW5nczogU2V0dGluZ3MpOiBcImxpZ2h0XCIgfCBcImRhcmtcIiB7XG4gIGlmIChzZXR0aW5ncy5tb2RlID09PSBcImxpZ2h0XCIgfHwgc2V0dGluZ3MubW9kZSA9PT0gXCJkYXJrXCIpIHJldHVybiBzZXR0aW5ncy5tb2RlO1xuICByZXR1cm4gcHJlZmVyc0RhcmsubWF0Y2hlcyA/IFwiZGFya1wiIDogXCJsaWdodFwiO1xufVxuXG4vKiogRHJlc3NlcyB0aGUgcGFnZSBpbiB0aGUgY2hvc2VuIHBhbGV0dGUgKHBhbGV0dGVzLmNzcyBrZXlzIG9mZiB0aGVzZSBhdHRyaWJ1dGVzKS4gKi9cbmZ1bmN0aW9uIGFwcGx5VGhlbWUoc2V0dGluZ3M6IFNldHRpbmdzKTogdm9pZCB7XG4gIGh0bWwuZGF0YXNldC5vdmVybG9va1BhbGV0dGUgPSBzZXR0aW5ncy5wYWxldHRlO1xuICBodG1sLmRhdGFzZXQub3Zlcmxvb2tNb2RlID0gc2V0dGluZ3MubW9kZTtcbiAgaHRtbC5kYXRhc2V0Lm92ZXJsb29rT3V0bG9va1NjaGVtZSA9IHByZWZlcnNEYXJrLm1hdGNoZXMgPyBcImRhcmtcIiA6IFwibGlnaHRcIjtcbn1cblxuZnVuY3Rpb24gZWw8SyBleHRlbmRzIGtleW9mIEhUTUxFbGVtZW50VGFnTmFtZU1hcD4oXG4gIHRhZzogSyxcbiAgYXR0cnM6IFJlY29yZDxzdHJpbmcsIHN0cmluZz4gPSB7fSxcbiAgLi4uY2hpbGRyZW46IChOb2RlIHwgc3RyaW5nKVtdXG4pOiBIVE1MRWxlbWVudFRhZ05hbWVNYXBbS10ge1xuICBjb25zdCBub2RlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCh0YWcpO1xuICBmb3IgKGNvbnN0IFtrLCB2XSBvZiBPYmplY3QuZW50cmllcyhhdHRycykpIG5vZGUuc2V0QXR0cmlidXRlKGssIHYpO1xuICBub2RlLmFwcGVuZCguLi5jaGlsZHJlbik7XG4gIHJldHVybiBub2RlO1xufVxuXG5mdW5jdGlvbiBzZWN0aW9uKHRpdGxlOiBzdHJpbmcsIC4uLmNoaWxkcmVuOiBOb2RlW10pOiBIVE1MRWxlbWVudCB7XG4gIHJldHVybiBlbChcInNlY3Rpb25cIiwge30sIGVsKFwiaDJcIiwge30sIHRpdGxlKSwgLi4uY2hpbGRyZW4pO1xufVxuXG4vKiogTXV0dWFsbHkgZXhjbHVzaXZlIGJ1dHRvbnMgZm9yIGEgY2hvaWNlOyBgcmVuZGVyYCBkcmF3cyBlYWNoIG9wdGlvbi4gKi9cbmZ1bmN0aW9uIGNob2ljZUdyb3VwKFxuICBpZDogQ2hvaWNlSWQsXG4gIHNldHRpbmdzOiBTZXR0aW5ncyxcbiAgY2xhc3NOYW1lOiBzdHJpbmcsXG4gIHJlbmRlcjogKHZhbHVlOiBzdHJpbmcsIGxhYmVsOiBzdHJpbmcpID0+IEhUTUxCdXR0b25FbGVtZW50LFxuKTogSFRNTEVsZW1lbnQge1xuICBjb25zdCBncm91cCA9IGVsKFwiZGl2XCIsIHtcbiAgICBjbGFzczogY2xhc3NOYW1lLFxuICAgIHJvbGU6IFwicmFkaW9ncm91cFwiLFxuICAgIFwiYXJpYS1sYWJlbFwiOiBDSE9JQ0VTW2lkXS5sYWJlbCxcbiAgfSk7XG4gIGZvciAoY29uc3QgW3ZhbHVlLCBsYWJlbF0gb2YgT2JqZWN0LmVudHJpZXMoQ0hPSUNFU1tpZF0ub3B0aW9ucykpIHtcbiAgICBjb25zdCBidXR0b24gPSByZW5kZXIodmFsdWUsIGxhYmVsKTtcbiAgICBidXR0b24udHlwZSA9IFwiYnV0dG9uXCI7XG4gICAgYnV0dG9uLnNldEF0dHJpYnV0ZShcInJvbGVcIiwgXCJyYWRpb1wiKTtcbiAgICBidXR0b24uc2V0QXR0cmlidXRlKFwiYXJpYS1jaGVja2VkXCIsIFN0cmluZyhzZXR0aW5nc1tpZF0gPT09IHZhbHVlKSk7XG4gICAgYnV0dG9uLmFkZEV2ZW50TGlzdGVuZXIoXCJjbGlja1wiLCAoKSA9PiB2b2lkIHNhdmVTZXR0aW5nKGlkLCB2YWx1ZSkpO1xuICAgIGdyb3VwLmFwcGVuZChidXR0b24pO1xuICB9XG4gIHJldHVybiBncm91cDtcbn1cblxuLyoqIEEgdGlueSB3aW5kb3cgcGFpbnRlZCBpbiBpdHMgb3duIHBhbGV0dGUsIGluIHRoZSBtb2RlIGN1cnJlbnRseSBpbiBlZmZlY3QuICovXG5mdW5jdGlvbiBzd2F0Y2goc2V0dGluZ3M6IFNldHRpbmdzKSB7XG4gIHJldHVybiAodmFsdWU6IHN0cmluZywgbGFiZWw6IHN0cmluZyk6IEhUTUxCdXR0b25FbGVtZW50ID0+IHtcbiAgICBjb25zdCBwcmV2aWV3ID0gZWwoXG4gICAgICBcInNwYW5cIixcbiAgICAgIHtcbiAgICAgICAgY2xhc3M6IFwicHJldmlld1wiLFxuICAgICAgICBcImRhdGEtb3Zlcmxvb2stc3dhdGNoXCI6IFwiXCIsXG4gICAgICAgIFwiZGF0YS1vdmVybG9vay1wYWxldHRlXCI6IHZhbHVlLFxuICAgICAgICBcImRhdGEtb3Zlcmxvb2stbW9kZVwiOiBlZmZlY3RpdmVNb2RlKHNldHRpbmdzKSxcbiAgICAgIH0sXG4gICAgICBlbChcInNwYW5cIiwgeyBjbGFzczogXCJiYXJcIiB9KSxcbiAgICAgIGVsKFwic3BhblwiLCB7IGNsYXNzOiBcInNpZGVcIiB9KSxcbiAgICAgIGVsKFxuICAgICAgICBcInNwYW5cIixcbiAgICAgICAgeyBjbGFzczogXCJib2R5XCIgfSxcbiAgICAgICAgZWwoXCJzcGFuXCIsIHsgY2xhc3M6IFwibGluZSBzaG9ydFwiIH0pLFxuICAgICAgICBlbChcInNwYW5cIiwgeyBjbGFzczogXCJsaW5lXCIgfSksXG4gICAgICAgIGVsKFwic3BhblwiLCB7IGNsYXNzOiBcImxpbmUgZmFpbnRcIiB9KSxcbiAgICAgICAgZWwoXCJzcGFuXCIsIHsgY2xhc3M6IFwibGluZVwiIH0pLFxuICAgICAgKSxcbiAgICApO1xuICAgIHJldHVybiBlbChcbiAgICAgIFwiYnV0dG9uXCIsXG4gICAgICB7IGNsYXNzOiBcInN3YXRjaFwiLCB0aXRsZTogbGFiZWwgfSxcbiAgICAgIHByZXZpZXcsXG4gICAgICBTSE9SVF9QQUxFVFRFX05BTUVTW3ZhbHVlXSA/PyBsYWJlbCxcbiAgICApO1xuICB9O1xufVxuXG5mdW5jdGlvbiBsYXlvdXRDYXJkKHZhbHVlOiBzdHJpbmcsIGxhYmVsOiBzdHJpbmcpOiBIVE1MQnV0dG9uRWxlbWVudCB7XG4gIGNvbnN0IGJ1dHRvbiA9IGVsKFwiYnV0dG9uXCIsIHsgY2xhc3M6IFwibGF5b3V0XCIsIHRpdGxlOiBsYWJlbCB9KTtcbiAgYnV0dG9uLmlubmVySFRNTCA9IGA8c3ZnIHZpZXdCb3g9XCIwIDAgNjQgNDRcIiBhcmlhLWhpZGRlbj1cInRydWVcIj4ke0xBWU9VVF9TS0VUQ0hFU1t2YWx1ZV0gPz8gXCJcIn08L3N2Zz5gO1xuICBidXR0b24uYXBwZW5kKFNIT1JUX0xBWU9VVF9OQU1FU1t2YWx1ZV0gPz8gbGFiZWwpO1xuICByZXR1cm4gYnV0dG9uO1xufVxuXG5mdW5jdGlvbiBzZWdtZW50KHZhbHVlOiBzdHJpbmcsIGxhYmVsOiBzdHJpbmcpOiBIVE1MQnV0dG9uRWxlbWVudCB7XG4gIHJldHVybiBlbChcImJ1dHRvblwiLCB7IFwiZGF0YS12YWx1ZVwiOiB2YWx1ZSB9LCBsYWJlbCk7XG59XG5cbmZ1bmN0aW9uIHRvZ2dsZShpZDogRmVhdHVyZUlkLCBzZXR0aW5nczogU2V0dGluZ3MpOiBIVE1MTGFiZWxFbGVtZW50IHtcbiAgY29uc3QgaW5wdXQgPSBlbChcImlucHV0XCIsIHsgdHlwZTogXCJjaGVja2JveFwiLCBpZDogYHRvZ2dsZS0ke2lkfWAgfSk7XG4gIGlucHV0LmNoZWNrZWQgPSBzZXR0aW5nc1tpZF07XG4gIGlucHV0LmFkZEV2ZW50TGlzdGVuZXIoXCJjaGFuZ2VcIiwgKCkgPT4gdm9pZCBzYXZlU2V0dGluZyhpZCwgaW5wdXQuY2hlY2tlZCkpO1xuICByZXR1cm4gZWwoXG4gICAgXCJsYWJlbFwiLFxuICAgIHsgY2xhc3M6IFwidG9nZ2xlXCIsIGZvcjogYHRvZ2dsZS0ke2lkfWAgfSxcbiAgICBlbChcInNwYW5cIiwge30sIEZFQVRVUkVTW2lkXS5sYWJlbCksXG4gICAgaW5wdXQsXG4gICk7XG59XG5cbmZ1bmN0aW9uIHRvZ2dsZXMoaWRzOiBGZWF0dXJlSWRbXSwgc2V0dGluZ3M6IFNldHRpbmdzKTogSFRNTEVsZW1lbnQge1xuICByZXR1cm4gZWwoXCJkaXZcIiwgeyBjbGFzczogXCJ0b2dnbGVzXCIgfSwgLi4uaWRzLm1hcCgoaWQpID0+IHRvZ2dsZShpZCwgc2V0dGluZ3MpKSk7XG59XG5cbmNvbnN0IGZlYXR1cmVzSW4gPSAoZ3JvdXA6IHN0cmluZykgPT4gRkVBVFVSRV9JRFMuZmlsdGVyKChpZCkgPT4gRkVBVFVSRVNbaWRdLmdyb3VwID09PSBncm91cCk7XG5cbmxldCBkZWNsdXR0ZXJPcGVuID0gZmFsc2U7XG5sZXQgZGV2ZWxvcGVyT3BlbiA9IGZhbHNlO1xubGV0IGNhcHR1cmVTdGF0dXMgPSBcIlwiO1xuXG4vKiogQXNrcyB0aGUgT3V0bG9vayB0YWIgaW4gZnJvbnQgdG8gc2F2ZSBhIGNhcHR1cmUgb2YgaXRzZWxmLiBObyBwZXJtaXNzaW9uIG5lZWRlZDogdGhlIGNvbnRlbnRcbiAqIHNjcmlwdCB0aGF0IGFuc3dlcnMgaXMgYWxyZWFkeSBvbiBPdXRsb29rJ3MgaG9zdHMuICovXG5hc3luYyBmdW5jdGlvbiByZXF1ZXN0Q2FwdHVyZSgpOiBQcm9taXNlPHN0cmluZz4ge1xuICBjb25zdCBbdGFiXSA9IGF3YWl0IGNocm9tZS50YWJzLnF1ZXJ5KHsgYWN0aXZlOiB0cnVlLCBjdXJyZW50V2luZG93OiB0cnVlIH0pO1xuICBpZiAodGFiPy5pZCA9PT0gdW5kZWZpbmVkKSByZXR1cm4gXCJPcGVuIGFuIE91dGxvb2sgdGFiIGZpcnN0LlwiO1xuICB0cnkge1xuICAgIGNvbnN0IHJlcGx5ID0gKGF3YWl0IGNocm9tZS50YWJzLnNlbmRNZXNzYWdlKHRhYi5pZCwgeyB0eXBlOiBDQVBUVVJFX01FU1NBR0UgfSkpIGFzIHtcbiAgICAgIG5hbWU6IHN0cmluZztcbiAgICAgIGJ5dGVzOiBudW1iZXI7XG4gICAgfTtcbiAgICByZXR1cm4gYFNhdmVkICR7cmVwbHkubmFtZX0gKCR7TWF0aC5yb3VuZChyZXBseS5ieXRlcyAvIDEwMjQpfSBLQikgdG8gRG93bmxvYWRzLmA7XG4gIH0gY2F0Y2gge1xuICAgIHJldHVybiBcIk5vIGFuc3dlci4gVXNlIHRoaXMgb24gYW4gT3V0bG9vayB0YWIsIGFuZCByZWxvYWQgdGhhdCB0YWIgaWYgT0JCIHdhcyBqdXN0IHVwZGF0ZWQuXCI7XG4gIH1cbn1cblxuZnVuY3Rpb24gZGV2ZWxvcGVyU2VjdGlvbigpOiBIVE1MRWxlbWVudCB7XG4gIGNvbnN0IHN0YXR1cyA9IGVsKFwicFwiLCB7IGNsYXNzOiBcImhpbnRcIiwgcm9sZTogXCJzdGF0dXNcIiB9LCBjYXB0dXJlU3RhdHVzKTtcbiAgY29uc3QgYnV0dG9uID0gZWwoXCJidXR0b25cIiwgeyB0eXBlOiBcImJ1dHRvblwiLCBjbGFzczogXCJwaWxsXCIgfSwgXCJTYXZlIHBhZ2UgY2FwdHVyZVwiKTtcbiAgYnV0dG9uLmFkZEV2ZW50TGlzdGVuZXIoXCJjbGlja1wiLCAoKSA9PiB7XG4gICAgdm9pZCByZXF1ZXN0Q2FwdHVyZSgpLnRoZW4oKHRleHQpID0+IHtcbiAgICAgIGNhcHR1cmVTdGF0dXMgPSB0ZXh0O1xuICAgICAgc3RhdHVzLnRleHRDb250ZW50ID0gdGV4dDtcbiAgICB9KTtcbiAgfSk7XG4gIGNvbnN0IGRldGFpbHMgPSBlbChcbiAgICBcImRldGFpbHNcIixcbiAgICB7fSxcbiAgICBlbChcInN1bW1hcnlcIiwge30sIFwiRGV2ZWxvcGVyXCIpLFxuICAgIGVsKFxuICAgICAgXCJwXCIsXG4gICAgICB7IGNsYXNzOiBcImhpbnRcIiB9LFxuICAgICAgXCJTYXZlcyB0aGlzIE91dGxvb2sgdGFiIGFzIGFuIEhUTUwgZmlsZSBmb3IgZml4aW5nIG9yIHRlc3RpbmcgT0JCLiBJdCBjb250YWlucyB5b3VyIHJlYWwgbWFpbCBhbmQgc3RheXMgb24gdGhpcyBjb21wdXRlcjsgc2hhcmUgaXQgb25seSB3aXRoIHdob2V2ZXIgaXMgaGVscGluZyB5b3UgZml4IE9CQi5cIixcbiAgICApLFxuICAgIGJ1dHRvbixcbiAgICBzdGF0dXMsXG4gICk7XG4gIGRldGFpbHMub3BlbiA9IGRldmVsb3Blck9wZW47XG4gIGRldGFpbHMuYWRkRXZlbnRMaXN0ZW5lcihcInRvZ2dsZVwiLCAoKSA9PiB7XG4gICAgZGV2ZWxvcGVyT3BlbiA9IGRldGFpbHMub3BlbjtcbiAgfSk7XG4gIHJldHVybiBlbChcInNlY3Rpb25cIiwge30sIGRldGFpbHMpO1xufVxuXG5mdW5jdGlvbiByZW5kZXIoc2V0dGluZ3M6IFNldHRpbmdzKTogdm9pZCB7XG4gIGFwcGx5VGhlbWUoc2V0dGluZ3MpO1xuICBjb25zdCBtYWluID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoXCJzZXR0aW5nc1wiKTtcbiAgaWYgKCFtYWluKSByZXR1cm47XG5cbiAgY29uc3QgbGF5b3V0SGludCA9XG4gICAgc2V0dGluZ3MubGF5b3V0ID09PSBcImZhc3RtYWlsXCIgfHwgc2V0dGluZ3MubGF5b3V0ID09PSBcImZvY3VzXCJcbiAgICAgID8gW1xuICAgICAgICAgIGVsKFxuICAgICAgICAgICAgXCJwXCIsXG4gICAgICAgICAgICB7IGNsYXNzOiBcImhpbnRcIiB9LFxuICAgICAgICAgICAgXCJQYWlyIHdpdGggT3V0bG9vaydzIFNldHRpbmdzIFx1MjE5MiBNYWlsIFx1MjE5MiBMYXlvdXQgXHUyMTkyIFJlYWRpbmcgcGFuZTogSGlkZS5cIixcbiAgICAgICAgICApLFxuICAgICAgICBdXG4gICAgICA6IFtdO1xuXG4gIGNvbnN0IGRlY2x1dHRlciA9IGZlYXR1cmVzSW4oXCJEZWNsdXR0ZXJcIik7XG4gIGNvbnN0IGhpZGRlbiA9IGRlY2x1dHRlci5maWx0ZXIoKGlkKSA9PiBzZXR0aW5nc1tpZF0pLmxlbmd0aDtcbiAgY29uc3QgZGV0YWlscyA9IGVsKFxuICAgIFwiZGV0YWlsc1wiLFxuICAgIHt9LFxuICAgIGVsKFxuICAgICAgXCJzdW1tYXJ5XCIsXG4gICAgICB7fSxcbiAgICAgIFwiRGVjbHV0dGVyXCIsXG4gICAgICBlbChcInNwYW5cIiwgeyBjbGFzczogXCJjb3VudFwiIH0sIGAke2hpZGRlbn0gb2YgJHtkZWNsdXR0ZXIubGVuZ3RofSBoaWRkZW5gKSxcbiAgICApLFxuICAgIHRvZ2dsZXMoZGVjbHV0dGVyLCBzZXR0aW5ncyksXG4gICk7XG4gIGRldGFpbHMub3BlbiA9IGRlY2x1dHRlck9wZW47XG4gIGRldGFpbHMuYWRkRXZlbnRMaXN0ZW5lcihcInRvZ2dsZVwiLCAoKSA9PiB7XG4gICAgZGVjbHV0dGVyT3BlbiA9IGRldGFpbHMub3BlbjtcbiAgfSk7XG5cbiAgbWFpbi5yZXBsYWNlQ2hpbGRyZW4oXG4gICAgc2VjdGlvbihcIkNvbG9yc1wiLCBjaG9pY2VHcm91cChcInBhbGV0dGVcIiwgc2V0dGluZ3MsIFwic3dhdGNoZXNcIiwgc3dhdGNoKHNldHRpbmdzKSkpLFxuICAgIHNlY3Rpb24oXCJMaWdodCBvciBkYXJrXCIsIGNob2ljZUdyb3VwKFwibW9kZVwiLCBzZXR0aW5ncywgXCJzZWdtZW50ZWRcIiwgc2VnbWVudCkpLFxuICAgIHNlY3Rpb24oXCJMYXlvdXRcIiwgY2hvaWNlR3JvdXAoXCJsYXlvdXRcIiwgc2V0dGluZ3MsIFwibGF5b3V0c1wiLCBsYXlvdXRDYXJkKSwgLi4ubGF5b3V0SGludCksXG4gICAgc2VjdGlvbihcIktleWJvYXJkXCIsIHRvZ2dsZXMoZmVhdHVyZXNJbihcIktleWJvYXJkXCIpLCBzZXR0aW5ncykpLFxuICAgIGVsKFwic2VjdGlvblwiLCB7fSwgZGV0YWlscyksXG4gICAgc2VjdGlvbihcIkdlbmVyYWxcIiwgdG9nZ2xlcyhmZWF0dXJlc0luKFwiR2VuZXJhbFwiKSwgc2V0dGluZ3MpKSxcbiAgICBkZXZlbG9wZXJTZWN0aW9uKCksXG4gICk7XG59XG5cbmFzeW5jIGZ1bmN0aW9uIG1haW4oKTogUHJvbWlzZTx2b2lkPiB7XG4gIGNvbnN0IHZlcnNpb24gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChcInZlcnNpb25cIik7XG4gIGlmICh2ZXJzaW9uKSB2ZXJzaW9uLnRleHRDb250ZW50ID0gYHYke2Nocm9tZS5ydW50aW1lLmdldE1hbmlmZXN0KCkudmVyc2lvbn1gO1xuICAvLyBQaWNrcyB1cCBuZXcgZmlsZXMgb24gZGlzayAoYWZ0ZXIgYGdpdCBwdWxsYCBpbiB0aGUgaW5zdGFsbCBmb2xkZXIpOyBzZWUgUkVBRE1FLlxuICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChcInJlbG9hZFwiKT8uYWRkRXZlbnRMaXN0ZW5lcihcImNsaWNrXCIsICgpID0+IGNocm9tZS5ydW50aW1lLnJlbG9hZCgpKTtcbiAgbGV0IHNldHRpbmdzID0gYXdhaXQgbG9hZFNldHRpbmdzKCk7XG4gIHJlbmRlcihzZXR0aW5ncyk7XG4gIG9uU2V0dGluZ3NDaGFuZ2VkKChuZXh0KSA9PiB7XG4gICAgc2V0dGluZ3MgPSBuZXh0O1xuICAgIHJlbmRlcihzZXR0aW5ncyk7XG4gIH0pO1xuICBwcmVmZXJzRGFyay5hZGRFdmVudExpc3RlbmVyKFwiY2hhbmdlXCIsICgpID0+IHJlbmRlcihzZXR0aW5ncykpO1xufVxuXG52b2lkIG1haW4oKTtcbiJdLAogICJtYXBwaW5ncyI6ICI7OztBQXlCTyxNQUFNLFlBQVk7QUFBQTtBQUFBLElBRXZCLFVBQVUsRUFBRSxZQUFZLENBQUMsY0FBYyxHQUFHLFFBQVEsVUFBVSxLQUFLLFlBQVk7QUFBQSxJQUM3RSxXQUFXLEVBQUUsWUFBWSxDQUFD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
