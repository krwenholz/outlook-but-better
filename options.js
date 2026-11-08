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
//# sourceMappingURL=data:application/json;base64,ewogICJ2ZXJzaW9uIjogMywKICAic291cmNlcyI6IFsiLi4vc3JjL3NlbGVjdG9ycy50cyIsICIuLi9zcmMvY2FwdHVyZS50cyIsICIuLi9zcmMvc2V0dGluZ3MudHMiLCAiLi4vc3JjL29wdGlvbnMvb3B0aW9ucy50cyJdLAogICJzb3VyY2VzQ29udGVudCI6IFsiLy8gVGhlIG9ubHkgZmlsZSBhbGxvd2VkIHRvIGtub3cgd2hhdCBPdXRsb29rJ3MgRE9NIGxvb2tzIGxpa2UuXG4vL1xuLy8gQW5jaG9yIHByZWZlcmVuY2UgKHNlZSBQTEFOLm1kKTogc3RydWN0dXJhbCBJRHMgPiBBUklBIHJvbGVzL2xhYmVscyA+IHNlbWFudGljIGRhdGEtKiA+IHRleHQuXG4vLyBOZXZlciB1c2UgZ2VuZXJhdGVkIGNsYXNzIG5hbWVzOyBpZiBvbmUgaXMgdW5hdm9pZGFibGUsIG1hcmsgaXQgYC8vIEZSQUdJTEU6YC5cbi8vXG4vLyBWZXJpZmllZCBhZ2FpbnN0IGEgc2F2ZWQgcGFnZSBmcm9tIE9jdCAyMDI2ICh0ZXN0L2ZpeHR1cmVzL2luYm94Lmh0bWwpOyBzZWUgZG9jcy9kb20tc3VydmV5Lm1kLlxuLy8gRW50cmllcyBtYXJrZWQgVU5WRVJJRklFRCB3ZXJlIG5vdCBwcmVzZW50IGluIHRoYXQgY2FwdHVyZS4gYXJpYS1sYWJlbHMgYXNzdW1lIGFuIEVuZ2xpc2ggVUkuXG5cbmV4cG9ydCB0eXBlIFNlbGVjdG9yTmFtZSA9IGtleW9mIHR5cGVvZiBTRUxFQ1RPUlM7XG5cbmV4cG9ydCBpbnRlcmZhY2UgU2VsZWN0b3JFbnRyeSB7XG4gIC8qKiBDYW5kaWRhdGUgQ1NTIHNlbGVjdG9ycywgbW9zdCBzdGFibGUgZmlyc3QuIFRoZSBmaXJzdCBvbmUgdGhhdCBtYXRjaGVzIHdpbnMuICovXG4gIGNhbmRpZGF0ZXM6IHJlYWRvbmx5IHN0cmluZ1tdO1xuICAvKiogXCJhbHdheXNcIjogc2hvdWxkIGV4aXN0IG9uIGV2ZXJ5IG1haWwgdmlldywgc28gYSBtaXNzIGlzIGEgYnJlYWthZ2UuIFwic29tZXRpbWVzXCI6IGNvbnRleHR1YWwuICovXG4gIGV4cGVjdDogXCJhbHdheXNcIiB8IFwic29tZXRpbWVzXCI7XG4gIC8qKiBOYW1lIG9mIGFub3RoZXIgZW50cnkgdG8gcmVzb2x2ZSBpbnNpZGUsIGluc3RlYWQgb2YgdGhlIHdob2xlIGRvY3VtZW50LiAqL1xuICB3aXRoaW4/OiBzdHJpbmc7XG4gIC8qKiBUb2tlbiBhZGRlZCB0byB0aGUgZWxlbWVudCdzIGBkYXRhLW92ZXJsb29rYCBhdHRyaWJ1dGUgc28gQ1NTIGNhbiB0YXJnZXQgaXQuICovXG4gIHRhZz86IHN0cmluZztcbiAgLyoqIFRhZyBldmVyeSBtYXRjaCByYXRoZXIgdGhhbiBvbmx5IHRoZSBmaXJzdC4gKi9cbiAgYWxsPzogYm9vbGVhbjtcbiAgLyoqIE1hdGNoIGV2ZXJ5IGNhbmRpZGF0ZSAoYSBsaXN0IG9mIGRpc3RpbmN0IHRoaW5ncyksIG5vdCBqdXN0IHRoZSBmaXJzdCB0aGF0IG1hdGNoZXMuICovXG4gIHVuaW9uPzogYm9vbGVhbjtcbn1cblxuZXhwb3J0IGNvbnN0IFNFTEVDVE9SUyA9IHtcbiAgLy8gLS0tIFN0cnVjdHVyZSAtLS1cbiAgdGl0bGVCYXI6IHsgY2FuZGlkYXRlczogW1wiI093YVRpdGxlQmFyXCJdLCBleHBlY3Q6IFwiYWx3YXlzXCIsIHRhZzogXCJ0aXRsZS1iYXJcIiB9LFxuICBzZWFyY2hCb3g6IHsgY2FuZGlkYXRlczogW1wiI3RvcFNlYXJjaElucHV0XCIsICdbcm9sZT1cInNlYXJjaFwiXSBpbnB1dCddLCBleHBlY3Q6IFwiYWx3YXlzXCIgfSxcbiAgYXBwUmFpbDoge1xuICAgIGNhbmRpZGF0ZXM6IFtcIiNMZWZ0UmFpbFwiLCAnW3JvbGU9XCJuYXZpZ2F0aW9uXCJdW2FyaWEtbGFiZWw9XCJsZWZ0LXJhaWwtYXBwYmFyXCJdJ10sXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICAgIHRhZzogXCJhcHAtcmFpbFwiLFxuICB9LFxuICByaWJib246IHtcbiAgICBjYW5kaWRhdGVzOiBbJ1tkYXRhLWFwcC1zZWN0aW9uPVwiUmliYm9uXCJdJywgJ1tyb2xlPVwicmVnaW9uXCJdW2FyaWEtbGFiZWw9XCJSaWJib25cIl0nXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgdGFnOiBcInJpYmJvblwiLFxuICB9LFxuICByaWJib25UYWJzOiB7XG4gICAgY2FuZGlkYXRlczogWydbcm9sZT1cInRhYmxpc3RcIl1bYXJpYS1sYWJlbD1cIlJpYmJvbiB0YWJzXCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJyaWJib25cIixcbiAgICB0YWc6IFwicmliYm9uLXRhYnNcIixcbiAgfSxcbiAgZm9sZGVyUGFuZToge1xuICAgIGNhbmRpZGF0ZXM6IFtcbiAgICAgICdbZGF0YS1hcHAtc2VjdGlvbj1cIk5hdmlnYXRpb25QYW5lXCJdJyxcbiAgICAgICdbcm9sZT1cInJlZ2lvblwiXVthcmlhLWxhYmVsPVwiTmF2aWdhdGlvbiBwYW5lXCJdJyxcbiAgICBdLFxuICAgIGV4cGVjdDogXCJhbHdheXNcIixcbiAgICB0YWc6IFwiZm9sZGVyLXBhbmVcIixcbiAgfSxcbiAgZm9sZGVySXRlbToge1xuICAgIGNhbmRpZGF0ZXM6IFsnW3JvbGU9XCJ0cmVlaXRlbVwiXVtkYXRhLWZvbGRlci1uYW1lXScsICdbcm9sZT1cInRyZWVpdGVtXCJdJ10sXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICAgIHdpdGhpbjogXCJmb2xkZXJQYW5lXCIsXG4gICAgdGFnOiBcImZvbGRlci1pdGVtXCIsXG4gICAgYWxsOiB0cnVlLFxuICB9LFxuICBsaXN0UGFuZToge1xuICAgIGNhbmRpZGF0ZXM6IFsnW2RhdGEtYXBwLXNlY3Rpb249XCJNZXNzYWdlTGlzdFwiXSddLFxuICAgIGV4cGVjdDogXCJhbHdheXNcIixcbiAgICB0YWc6IFwibGlzdC1wYW5lXCIsXG4gIH0sXG4gIG1lc3NhZ2VMaXN0OiB7XG4gICAgY2FuZGlkYXRlczogW1wiI01haWxMaXN0XCIsICdbcm9sZT1cImxpc3Rib3hcIl1bYXJpYS1sYWJlbD1cIk1lc3NhZ2UgbGlzdFwiXSddLFxuICAgIGV4cGVjdDogXCJhbHdheXNcIixcbiAgICB0YWc6IFwibWVzc2FnZS1saXN0XCIsXG4gIH0sXG4gIG1lc3NhZ2VSb3c6IHtcbiAgICBjYW5kaWRhdGVzOiBbJ1tyb2xlPVwib3B0aW9uXCJdW2RhdGEtY29udmlkXScsICdbcm9sZT1cIm9wdGlvblwiXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB3aXRoaW46IFwibWVzc2FnZUxpc3RcIixcbiAgICB0YWc6IFwibWVzc2FnZS1yb3dcIixcbiAgICBhbGw6IHRydWUsXG4gIH0sXG4gIGdyb3VwSGVhZGVyOiB7XG4gICAgY2FuZGlkYXRlczogWydbcm9sZT1cImJ1dHRvblwiXVtpZF49XCJncm91cEhlYWRlclwiXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB3aXRoaW46IFwibWVzc2FnZUxpc3RcIixcbiAgICB0YWc6IFwiZ3JvdXAtaGVhZGVyXCIsXG4gICAgYWxsOiB0cnVlLFxuICB9LFxuICBmb2N1c2VkVGFiczoge1xuICAgIGNhbmRpZGF0ZXM6IFsnW3JvbGU9XCJ0YWJsaXN0XCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJsaXN0UGFuZVwiLFxuICAgIHRhZzogXCJmb2N1c2VkLXRhYnNcIixcbiAgfSxcbiAgcmVhZGluZ1BhbmU6IHtcbiAgICBjYW5kaWRhdGVzOiBbXG4gICAgICAnW2RhdGEtYXBwLXNlY3Rpb249XCJNYWlsUmVhZENvbXBvc2VcIl0nLFxuICAgICAgXCIjUmVhZGluZ1BhbmVDb250YWluZXJJZFwiLFxuICAgICAgJ1tyb2xlPVwibWFpblwiXVthcmlhLWxhYmVsPVwiUmVhZGluZyBQYW5lXCJdJyxcbiAgICBdLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB0YWc6IFwicmVhZGluZy1wYW5lXCIsXG4gIH0sXG5cbiAgbWVzc2FnZUJvZHk6IHtcbiAgICBjYW5kaWRhdGVzOiBbJ1tyb2xlPVwiZG9jdW1lbnRcIl1bYXJpYS1sYWJlbD1cIk1lc3NhZ2UgYm9keVwiXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB3aXRoaW46IFwicmVhZGluZ1BhbmVcIixcbiAgICB0YWc6IFwibWVzc2FnZS1ib2R5XCIsXG4gICAgYWxsOiB0cnVlLFxuICB9LFxuXG4gIGluZm9CYXI6IHtcbiAgICAvLyBOb3RpY2VzIGFib3ZlIGEgbWVzc2FnZSAoXCJZb3UgcmVwbGllZCBvblx1MjAyNlwiKTsgdGhlaXIgdGV4dCBzaXRzIGluIGFuIGVsZW1lbnQgd2hvc2UgaWQgc3RhcnRzXG4gICAgLy8gd2l0aCBcImluZm9CYXJcIi4gT3V0bG9vayBwYWludHMgdGhlbSB3aXRoIGlubGluZSBjb2xvdXJzLlxuICAgIGNhbmRpZGF0ZXM6IFsnW3JvbGU9XCJncm91cFwiXTpoYXMoPiAqID4gW2lkXj1cImluZm9CYXJcIl0pJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHdpdGhpbjogXCJyZWFkaW5nUGFuZVwiLFxuICAgIHRhZzogXCJpbmZvLWJhclwiLFxuICAgIGFsbDogdHJ1ZSxcbiAgfSxcblxuICAvLyAtLS0gQ2x1dHRlciAoZWFjaCBoaWRlYWJsZSBlbnRyeSBnZXRzIGEgdG9nZ2xlIGluIHNldHRpbmdzLnRzKSAtLS1cbiAgdXBzZWxsQmFubmVyOiB7XG4gICAgY2FuZGlkYXRlczogWydbZGF0YS1hcHAtc2VjdGlvbj1cIlVwc2VsbEJhbm5lclNlY3Rpb25cIl0nXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgdGFnOiBcInVwc2VsbFwiLFxuICAgIGFsbDogdHJ1ZSxcbiAgfSxcbiAgaGVhZGVyRXh0cmFzOiB7XG4gICAgY2FuZGlkYXRlczogW1wiI293YVRlYW1zQnRuX2NvbnRhaW5lclwiLCBcIiNvd2FUaW1lUGFuZWxCdG5fY29udGFpbmVyXCJdLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB0YWc6IFwiaGVhZGVyLWV4dHJhXCIsXG4gICAgdW5pb246IHRydWUsXG4gIH0sXG4gIGFwcFJhaWxFeHRyYXM6IHtcbiAgICBjYW5kaWRhdGVzOiBbXCJUbyBEb1wiLCBcIk5ld3NsZXR0ZXJzXCIsIFwiT3JnIEV4cGxvcmVyXCIsIFwiT25lRHJpdmVcIl0ubWFwKFxuICAgICAgKGFwcCkgPT4gYFtyb2xlPVwibm9uZVwiXVthcmlhLWxhYmVsPVwiJHthcHB9XCJdYCxcbiAgICApLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB3aXRoaW46IFwiYXBwUmFpbFwiLFxuICAgIHRhZzogXCJyYWlsLWV4dHJhXCIsXG4gICAgdW5pb246IHRydWUsXG4gIH0sXG4gIHJpYmJvbkV4dHJhczoge1xuICAgIGNhbmRpZGF0ZXM6IFtcbiAgICAgIFwiU2hhcmUgdG8gVGVhbXNcIixcbiAgICAgIFwiU3dlZXBcIixcbiAgICAgIFwiQXNzaWduIHBvbGljeVwiLFxuICAgICAgXCJWaXZhIEluc2lnaHRzXCIsXG4gICAgICBcIk1vcmUgYXBwc1wiLFxuICAgICAgXCJRdWljayBzdGVwc1wiLFxuICAgICAgXCJNb3JlIHF1aWNrIHN0ZXBzXCIsXG4gICAgXS5tYXAoKGxhYmVsKSA9PiBgW2FyaWEtbGFiZWw9XCIke2xhYmVsfVwiXWApLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB3aXRoaW46IFwicmliYm9uXCIsXG4gICAgdGFnOiBcInJpYmJvbi1leHRyYVwiLFxuICAgIHVuaW9uOiB0cnVlLFxuICB9LFxuICByb3dBdmF0YXI6IHtcbiAgICBjYW5kaWRhdGVzOiBbJ1tyb2xlPVwiaW1nXCJdW2lkXj1cImF2YXRhci1cIl0nXSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgd2l0aGluOiBcIm1lc3NhZ2VMaXN0XCIsXG4gICAgdGFnOiBcInJvdy1hdmF0YXJcIixcbiAgICBhbGw6IHRydWUsXG4gIH0sXG4gIHJvd0hvdmVyQWN0aW9uczoge1xuICAgIGNhbmRpZGF0ZXM6IFtcbiAgICAgIFwiQ3VzdG9taXplIEFjdGlvbnNcIixcbiAgICAgIFwiRmxhZyB0aGlzIG1lc3NhZ2VcIixcbiAgICAgIFwiS2VlcCB0aGlzIG1lc3NhZ2UgYXQgdGhlIHRvcCBvZiB5b3VyIGZvbGRlclwiLFxuICAgIF0ubWFwKChsYWJlbCkgPT4gYGJ1dHRvblthcmlhLWxhYmVsPVwiJHtsYWJlbH1cIl1gKSxcbiAgICBleHBlY3Q6IFwic29tZXRpbWVzXCIsXG4gICAgd2l0aGluOiBcIm1lc3NhZ2VMaXN0XCIsXG4gICAgdGFnOiBcInJvdy1hY3Rpb25cIixcbiAgICB1bmlvbjogdHJ1ZSxcbiAgfSxcbiAgYXBwQmFubmVyOiB7XG4gICAgLy8gRmx1ZW50J3Mgc3RhdGljIFwiZnVpLU1lc3NhZ2VCYXJcIiBjbGFzcyBpcyBwYXJ0IG9mIGl0cyBwdWJsaWMgc3R5bGluZyBBUEksIG5vdCBnZW5lcmF0ZWQ7XG4gICAgLy8gbm90aGluZyBlbHNlIG1hcmtzIHRoZXNlLiBTY29wZWQgdG8gYmFycyBhY3Jvc3MgdGhlIHRvcCBvZiB0aGUgYXBwLCBub3QgaW5zaWRlIG1lc3NhZ2VzLlxuICAgIGNhbmRpZGF0ZXM6IFsnI21haW5BcHAgPiBkaXYgPiBbcm9sZT1cImdyb3VwXCJdLmZ1aS1NZXNzYWdlQmFyJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICAgIHRhZzogXCJhcHAtYmFubmVyXCIsXG4gICAgYWxsOiB0cnVlLFxuICB9LFxuICBjb3BpbG90OiB7XG4gICAgLy8gVU5WRVJJRklFRDogbm8gQ29waWxvdCBzdXJmYWNlcyBpbiB0aGUgT2N0IDIwMjYgY2FwdHVyZS5cbiAgICBjYW5kaWRhdGVzOiBbJ2J1dHRvblthcmlhLWxhYmVsKj1cIkNvcGlsb3RcIiBpXScsICdbcm9sZT1cInJlZ2lvblwiXVthcmlhLWxhYmVsKj1cIkNvcGlsb3RcIiBpXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB0YWc6IFwiY29waWxvdFwiLFxuICAgIHVuaW9uOiB0cnVlLFxuICB9LFxuICBleHRlcm5hbFRhZzoge1xuICAgIC8vIFVOVkVSSUZJRUQ6IG5vIEV4dGVybmFsIHRhZ3MgaW4gdGhlIE9jdCAyMDI2IGNhcHR1cmUuXG4gICAgY2FuZGlkYXRlczogWydbdGl0bGU9XCJFeHRlcm5hbCBzZW5kZXJcIl0nLCAnW2FyaWEtbGFiZWw9XCJFeHRlcm5hbFwiXSddLFxuICAgIGV4cGVjdDogXCJzb21ldGltZXNcIixcbiAgICB0YWc6IFwiZXh0ZXJuYWwtdGFnXCIsXG4gICAgdW5pb246IHRydWUsXG4gIH0sXG5cbiAgLy8gLS0tIENvbnRyb2xzIHRoZSBrZXltYXAgZHJpdmVzIC0tLVxuICBjb21wb3NlQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJOZXcgbWFpbFwiXSddLCBleHBlY3Q6IFwiYWx3YXlzXCIgfSxcbiAgYXJjaGl2ZUJ1dHRvbjogeyBjYW5kaWRhdGVzOiBbJ2J1dHRvblthcmlhLWxhYmVsPVwiQXJjaGl2ZVwiXSddLCBleHBlY3Q6IFwic29tZXRpbWVzXCIgfSxcbiAgZGVsZXRlQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJEZWxldGVcIl0nXSwgZXhwZWN0OiBcInNvbWV0aW1lc1wiIH0sXG4gIC8vIFVOVkVSSUZJRUQ6IHRoaXMgdGVuYW50IHNob3dzIGEgXCJQaGlzaCBBbGVydFwiIGFkZC1pbiBpbnN0ZWFkIG9mIE91dGxvb2sncyBSZXBvcnQgbWVudS5cbiAganVua0J1dHRvbjoge1xuICAgIGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJSZXBvcnRcIl0nLCAnYnV0dG9uW2FyaWEtbGFiZWw9XCJKdW5rXCJdJ10sXG4gICAgZXhwZWN0OiBcInNvbWV0aW1lc1wiLFxuICB9LFxuICBtb3ZlQnV0dG9uOiB7IGNhbmRpZGF0ZXM6IFsnYnV0dG9uW2FyaWEtbGFiZWw9XCJNb3ZlIHRvXCJdJ10sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICBzbm9vemVCdXR0b246IHsgY2FuZGlkYXRlczogWydidXR0b25bYXJpYS1sYWJlbD1cIlNub296ZVwiXSddLCBleHBlY3Q6IFwic29tZXRpbWVzXCIgfSxcbiAgZmxhZ0J1dHRvbjogeyBjYW5kaWRhdGVzOiBbJ2J1dHRvblthcmlhLWxhYmVsPVwiRmxhZyAvIFVuZmxhZ1wiXSddLCBleHBlY3Q6IFwic29tZXRpbWVzXCIgfSxcbiAgcmVhZFVucmVhZEJ1dHRvbjogeyBjYW5kaWRhdGVzOiBbJ2J1dHRvblthcmlhLWxhYmVsPVwiUmVhZCAvIFVucmVhZFwiXSddLCBleHBlY3Q6IFwic29tZXRpbWVzXCIgfSxcbiAgdW5kb0J1dHRvbjogeyBjYW5kaWRhdGVzOiBbJ2J1dHRvblthcmlhLWxhYmVsPVwiVW5kb1wiXSddLCBleHBlY3Q6IFwic29tZXRpbWVzXCIgfSxcbiAgc3luY0J1dHRvbjogeyBjYW5kaWRhdGVzOiBbXCIjb3dhU3luY0J1dHRvbl9jb250YWluZXJcIl0sIGV4cGVjdDogXCJzb21ldGltZXNcIiB9LFxuICBzZXR0aW5nc0J1dHRvbjogeyBjYW5kaWRhdGVzOiBbXCIjb3dhU2V0dGluZ3NCdG5fY29udGFpbmVyXCJdLCBleHBlY3Q6IFwiYWx3YXlzXCIgfSxcbiAgLy8gQXBwIHJhaWwgYnV0dG9ucy4gRmFsbGJhY2tzIGFyZSB0aGUgTWljcm9zb2Z0IDM2NSBhcHAgaWRzIE91dGxvb2sgcHV0cyBpbiBkYXRhLXRpZCwgdGhlIHNhbWVcbiAgLy8gYWNyb3NzIHRoZSBPY3QgYW5kIE5vdiAyMDI2IGNhcHR1cmVzLiBUaGUgcmFpbCByZW5kZXJzIGFmdGVyIHRoZSBtZXNzYWdlIGxpc3QuXG4gIG1haWxBcHA6IHtcbiAgICBjYW5kaWRhdGVzOiBbJ2J1dHRvblthcmlhLWxhYmVsPVwiTWFpbFwiXScsICdbZGF0YS10aWQ9XCJkZGVhNzc0Yy0zODJiLTQ3ZDctYWFiNS1hZGMyMTM5YTgwMmJcIl0nXSxcbiAgICB3aXRoaW46IFwiYXBwUmFpbFwiLFxuICAgIGV4cGVjdDogXCJhbHdheXNcIixcbiAgfSxcbiAgY2FsZW5kYXJBcHA6IHtcbiAgICBjYW5kaWRhdGVzOiBbXG4gICAgICAnYnV0dG9uW2FyaWEtbGFiZWw9XCJDYWxlbmRhclwiXScsXG4gICAgICAnW2RhdGEtdGlkPVwiOGNiZWI4NmYtODNlMS00M2I1LWFhYmEtY2QzNTE0MzIyZjBiXCJdJyxcbiAgICBdLFxuICAgIHdpdGhpbjogXCJhcHBSYWlsXCIsXG4gICAgZXhwZWN0OiBcImFsd2F5c1wiLFxuICB9LFxuICBwZW9wbGVBcHA6IHtcbiAgICBjYW5kaWRhdGVzOiBbXG4gICAgICAnYnV0dG9uW2FyaWEtbGFiZWw9XCJQZW9wbGVcIl0nLFxuICAgICAgJ1tkYXRhLXRpZD1cIjM1NWZiZDc5LTNiYTItNDU1NC04ZjJkLTAzMDBmZGU5MWYzMFwiXScsXG4gICAgXSxcbiAgICB3aXRoaW46IFwiYXBwUmFpbFwiLFxuICAgIGV4cGVjdDogXCJhbHdheXNcIixcbiAgfSxcbn0gYXMgY29uc3Qgc2F0aXNmaWVzIFJlY29yZDxzdHJpbmcsIFNlbGVjdG9yRW50cnk+O1xuXG5jb25zdCBlbnRyeSA9IChuYW1lOiBTZWxlY3Rvck5hbWUpOiBTZWxlY3RvckVudHJ5ID0+IFNFTEVDVE9SU1tuYW1lXTtcblxuZnVuY3Rpb24gc2NvcGUobmFtZTogU2VsZWN0b3JOYW1lLCByb290OiBQYXJlbnROb2RlKTogUGFyZW50Tm9kZSB8IG51bGwge1xuICBjb25zdCBwYXJlbnQgPSBlbnRyeShuYW1lKS53aXRoaW47XG4gIHJldHVybiBwYXJlbnQgPyByZXNvbHZlKHBhcmVudCBhcyBTZWxlY3Rvck5hbWUsIHJvb3QpIDogcm9vdDtcbn1cblxuLyoqIFRoZSBmaXJzdCBlbGVtZW50IG1hdGNoaW5nIGBuYW1lYCwgdHJ5aW5nIGNhbmRpZGF0ZXMgaW4gcHJpb3JpdHkgb3JkZXIuICovXG5leHBvcnQgZnVuY3Rpb24gcmVzb2x2ZShuYW1lOiBTZWxlY3Rvck5hbWUsIHJvb3Q6IFBhcmVudE5vZGUgPSBkb2N1bWVudCk6IEVsZW1lbnQgfCBudWxsIHtcbiAgY29uc3Qgd2l0aGluID0gc2NvcGUobmFtZSwgcm9vdCk7XG4gIGlmICghd2l0aGluKSByZXR1cm4gbnVsbDtcbiAgZm9yIChjb25zdCBjYW5kaWRhdGUgb2YgZW50cnkobmFtZSkuY2FuZGlkYXRlcykge1xuICAgIGNvbnN0IGVsID0gd2l0aGluLnF1ZXJ5U2VsZWN0b3IoY2FuZGlkYXRlKTtcbiAgICBpZiAoZWwpIHJldHVybiBlbDtcbiAgfVxuICByZXR1cm4gbnVsbDtcbn1cblxuLyoqIEV2ZXJ5IG1hdGNoOiBmb3IgdW5pb24gZW50cmllcywgb2YgYWxsIGNhbmRpZGF0ZXM7IG90aGVyd2lzZSBvZiB0aGUgZmlyc3QgY2FuZGlkYXRlIHRoYXQgbWF0Y2hlcy4gKi9cbmV4cG9ydCBmdW5jdGlvbiByZXNvbHZlQWxsKG5hbWU6IFNlbGVjdG9yTmFtZSwgcm9vdDogUGFyZW50Tm9kZSA9IGRvY3VtZW50KTogRWxlbWVudFtdIHtcbiAgY29uc3Qgd2l0aGluID0gc2NvcGUobmFtZSwgcm9vdCk7XG4gIGlmICghd2l0aGluKSByZXR1cm4gW107XG4gIGNvbnN0IHsgY2FuZGlkYXRlcywgdW5pb24gfSA9IGVudHJ5KG5hbWUpO1xuICBpZiAodW5pb24pIHJldHVybiBbLi4ud2l0aGluLnF1ZXJ5U2VsZWN0b3JBbGwoY2FuZGlkYXRlcy5qb2luKFwiLCBcIikpXTtcbiAgZm9yIChjb25zdCBjYW5kaWRhdGUgb2YgY2FuZGlkYXRlcykge1xuICAgIGNvbnN0IGVscyA9IHdpdGhpbi5xdWVyeVNlbGVjdG9yQWxsKGNhbmRpZGF0ZSk7XG4gICAgaWYgKGVscy5sZW5ndGggPiAwKSByZXR1cm4gWy4uLmVsc107XG4gIH1cbiAgcmV0dXJuIFtdO1xufVxuXG4vKiogQWRkIGBkYXRhLW92ZXJsb29rYCB0b2tlbnMgdG8gZXZlcnkgdGFnZ2VkIGVudHJ5IHNvIHN0eWxlc2hlZXRzIG5ldmVyIHRvdWNoIE91dGxvb2sncyBET00gc2hhcGUuICovXG5leHBvcnQgZnVuY3Rpb24gdGFnRWxlbWVudHMocm9vdDogUGFyZW50Tm9kZSA9IGRvY3VtZW50KTogdm9pZCB7XG4gIGZvciAoY29uc3QgbmFtZSBvZiBPYmplY3Qua2V5cyhTRUxFQ1RPUlMpIGFzIFNlbGVjdG9yTmFtZVtdKSB7XG4gICAgY29uc3QgeyB0YWcsIGFsbCwgdW5pb24gfSA9IGVudHJ5KG5hbWUpO1xuICAgIGlmICghdGFnKSBjb250aW51ZTtcbiAgICBjb25zdCBlbHMgPVxuICAgICAgYWxsIHx8IHVuaW9uID8gcmVzb2x2ZUFsbChuYW1lLCByb290KSA6IFtyZXNvbHZlKG5hbWUsIHJvb3QpXS5maWx0ZXIoKGUpID0+IGUgIT09IG51bGwpO1xuICAgIGZvciAoY29uc3QgZWwgb2YgZWxzKSB7XG4gICAgICBjb25zdCB0b2tlbnMgPSBuZXcgU2V0KChlbC5nZXRBdHRyaWJ1dGUoXCJkYXRhLW92ZXJsb29rXCIpID8/IFwiXCIpLnNwbGl0KFwiIFwiKS5maWx0ZXIoQm9vbGVhbikpO1xuICAgICAgaWYgKHRva2Vucy5oYXModGFnKSkgY29udGludWU7XG4gICAgICB0b2tlbnMuYWRkKHRhZyk7XG4gICAgICBlbC5zZXRBdHRyaWJ1dGUoXCJkYXRhLW92ZXJsb29rXCIsIFsuLi50b2tlbnNdLmpvaW4oXCIgXCIpKTtcbiAgICB9XG4gIH1cbn1cblxuZXhwb3J0IGludGVyZmFjZSBIZWFsdGhSZXN1bHQge1xuICBuYW1lOiBTZWxlY3Rvck5hbWU7XG4gIGV4cGVjdDogU2VsZWN0b3JFbnRyeVtcImV4cGVjdFwiXTtcbiAgLyoqIEluZGV4IG9mIHRoZSBjYW5kaWRhdGUgdGhhdCBtYXRjaGVkLCBvciAtMS4gQSBub24temVybyBpbmRleCBtZWFucyBhIGZhbGxiYWNrIGlzIGluIHVzZS4gKi9cbiAgY2FuZGlkYXRlOiBudW1iZXI7XG4gIGNvdW50OiBudW1iZXI7XG59XG5cbmV4cG9ydCBmdW5jdGlvbiBoZWFsdGhDaGVjayhyb290OiBQYXJlbnROb2RlID0gZG9jdW1lbnQpOiBIZWFsdGhSZXN1bHRbXSB7XG4gIHJldHVybiAoT2JqZWN0LmtleXMoU0VMRUNUT1JTKSBhcyBTZWxlY3Rvck5hbWVbXSkubWFwKChuYW1lKSA9PiB7XG4gICAgY29uc3QgeyBjYW5kaWRhdGVzLCBleHBlY3QsIHVuaW9uIH0gPSBlbnRyeShuYW1lKTtcbiAgICBjb25zdCB3aXRoaW4gPSBzY29wZShuYW1lLCByb290KTtcbiAgICBsZXQgY2FuZGlkYXRlID0gLTE7XG4gICAgbGV0IGNvdW50ID0gMDtcbiAgICBpZiAod2l0aGluKSB7XG4gICAgICBjYW5kaWRhdGUgPSBjYW5kaWRhdGVzLmZpbmRJbmRleCgoYykgPT4gd2l0aGluLnF1ZXJ5U2VsZWN0b3IoYykgIT09IG51bGwpO1xuICAgICAgaWYgKHVuaW9uICYmIGNhbmRpZGF0ZSA+PSAwKSBjb3VudCA9IHJlc29sdmVBbGwobmFtZSwgcm9vdCkubGVuZ3RoO1xuICAgICAgZWxzZSBpZiAoY2FuZGlkYXRlID49IDApIGNvdW50ID0gd2l0aGluLnF1ZXJ5U2VsZWN0b3JBbGwoY2FuZGlkYXRlc1tjYW5kaWRhdGVdID8/IFwiXCIpLmxlbmd0aDtcbiAgICB9XG4gICAgLy8gQSB1bmlvbiBlbnRyeSBsaXN0cyBkaXN0aW5jdCB0aGluZ3MsIHNvIG1hdGNoaW5nIGEgbGF0ZXIgb25lIGlzIG5vdCBhIGZhbGxiYWNrLlxuICAgIGlmICh1bmlvbiAmJiBjYW5kaWRhdGUgPiAwKSBjYW5kaWRhdGUgPSAwO1xuICAgIHJldHVybiB7IG5hbWUsIGV4cGVjdCwgY2FuZGlkYXRlLCBjb3VudCB9O1xuICB9KTtcbn1cblxuZXhwb3J0IGludGVyZmFjZSBGb2xkZXJOb2RlIHtcbiAgbmFtZTogc3RyaW5nO1xuICAvKiogQW5jZXN0b3IgbmFtZXMgam9pbmVkIHdpdGggXCIvXCIsIGUuZy4gXCJJbmJveC9SZWNlaXB0c1wiLiAqL1xuICBwYXRoOiBzdHJpbmc7XG4gIGVsZW1lbnQ6IEVsZW1lbnQ7XG59XG5cbi8qKiBcIkluYm94IC0gOSBpdGVtcyAoMSB1bnJlYWQpXCIgLT4gXCJJbmJveFwiLiBUaGUgdGl0bGUgYXR0cmlidXRlIGNhcnJpZXMgY291bnRzLiAqL1xuZnVuY3Rpb24gZm9sZGVyTmFtZShlbGVtZW50OiBFbGVtZW50KTogc3RyaW5nIHtcbiAgY29uc3QgdGl0bGUgPSAoZWxlbWVudC5nZXRBdHRyaWJ1dGUoXCJ0aXRsZVwiKSA/PyBlbGVtZW50LnRleHRDb250ZW50ID8/IFwiXCIpLnRyaW0oKTtcbiAgcmV0dXJuIHRpdGxlLnJlcGxhY2UoL1xccystXFxzK1xcZCtcXHMraXRlbXM/KFxccytcXChcXGQrXFxzK3VucmVhZFxcKSk/JC8sIFwiXCIpO1xufVxuXG4vKipcbiAqIEZvbGRlcnMgY3VycmVudGx5IHJlbmRlcmVkIGluIHRoZSBmb2xkZXIgcGFuZSwgd2l0aCBwYXRocyBidWlsdCBmcm9tIGFyaWEtbGV2ZWwuIEZvbGRlcnNcbiAqIHVuZGVyIEZhdm9yaXRlcyBhcmUgc2tpcHBlZCAodGhleSBkdXBsaWNhdGUgdGhlIG1haWxib3ggdHJlZSksIGFzIGFyZSB0aGUgdG9wLWxldmVsIG1haWxib3hcbiAqIHJvb3RzLCB3aG9zZSBuYW1lcyBhcmUgYWNjb3VudCBhZGRyZXNzZXMuXG4gKiBVTlZFUklGSUVEOiBjb2xsYXBzZWQgc3ViZm9sZGVycyBtYXkgYmUgYWJzZW50IGZyb20gdGhlIERPTS5cbiAqL1xuZXhwb3J0IGZ1bmN0aW9uIHJlYWRGb2xkZXJUcmVlKHJvb3Q6IFBhcmVudE5vZGUgPSBkb2N1bWVudCk6IEZvbGRlck5vZGVbXSB7XG4gIGNvbnN0IHN0YWNrOiBzdHJpbmdbXSA9IFtdO1xuICBjb25zdCBmb2xkZXJzOiBGb2xkZXJOb2RlW10gPSBbXTtcbiAgbGV0IGluRmF2b3JpdGVzID0gZmFsc2U7XG4gIGZvciAoY29uc3QgZWxlbWVudCBvZiByZXNvbHZlQWxsKFwiZm9sZGVySXRlbVwiLCByb290KSkge1xuICAgIGNvbnN0IGxldmVsID0gTnVtYmVyKGVsZW1lbnQuZ2V0QXR0cmlidXRlKFwiYXJpYS1sZXZlbFwiKSA/PyBcIjFcIik7XG4gICAgY29uc3QgbmFtZSA9IGZvbGRlck5hbWUoZWxlbWVudCk7XG4gICAgc3RhY2subGVuZ3RoID0gTWF0aC5tYXgobGV2ZWwgLSAxLCAwKTtcbiAgICBzdGFjay5wdXNoKG5hbWUpO1xuICAgIGlmIChsZXZlbCA9PT0gMSkge1xuICAgICAgaW5GYXZvcml0ZXMgPSBlbGVtZW50LmdldEF0dHJpYnV0ZShcImRhdGEtZm9sZGVyLW5hbWVcIikgPT09IFwiZmF2b3JpdGVzXCI7XG4gICAgICBjb250aW51ZTtcbiAgICB9XG4gICAgaWYgKGluRmF2b3JpdGVzKSBjb250aW51ZTtcbiAgICBmb2xkZXJzLnB1c2goeyBuYW1lLCBwYXRoOiBzdGFjay5zbGljZSgxKS5maWx0ZXIoQm9vbGVhbikuam9pbihcIi9cIiksIGVsZW1lbnQgfSk7XG4gIH1cbiAgcmV0dXJuIGZvbGRlcnM7XG59XG5cbi8qKiBNZXNzYWdlIHJvd3MgY3VycmVudGx5IHNlbGVjdGVkIChvciBvcGVuKSBpbiB0aGUgbGlzdC4gKi9cbmV4cG9ydCBmdW5jdGlvbiBzZWxlY3RlZFJvd3Mocm9vdDogUGFyZW50Tm9kZSA9IGRvY3VtZW50KTogRWxlbWVudFtdIHtcbiAgcmV0dXJuIHJlc29sdmVBbGwoXCJtZXNzYWdlUm93XCIsIHJvb3QpLmZpbHRlcihcbiAgICAocm93KSA9PiByb3cuZ2V0QXR0cmlidXRlKFwiYXJpYS1zZWxlY3RlZFwiKSA9PT0gXCJ0cnVlXCIsXG4gICk7XG59XG5cbi8qKiBPdXRsb29rIHN0YXJ0cyBhbiB1bnJlYWQgcm93J3MgYWNjZXNzaWJsZSBuYW1lIHdpdGggXCJVbnJlYWRcIiAoRW5nbGlzaCBVSSkuICovXG5leHBvcnQgZnVuY3Rpb24gaXNVbnJlYWRSb3cocm93OiBFbGVtZW50KTogYm9vbGVhbiB7XG4gIHJldHVybiAvXlVucmVhZFxcYi8udGVzdChyb3cuZ2V0QXR0cmlidXRlKFwiYXJpYS1sYWJlbFwiKSA/PyBcIlwiKTtcbn1cblxuLyoqIE91dGxvb2sncyBjdXJyZW50IGxpZ2h0L2RhcmsgbW9kZSwgZnJvbSB0aGUgcGFsZXR0ZSB2YXJpYWJsZXMgaXQgd3JpdGVzIG9uIDxodG1sPi4gKi9cbmV4cG9ydCBmdW5jdGlvbiBvdXRsb29rQ29sb3JTY2hlbWUoaHRtbDogRWxlbWVudCk6IFwibGlnaHRcIiB8IFwiZGFya1wiIHtcbiAgY29uc3Qgc3R5bGUgPSBodG1sLmdldEF0dHJpYnV0ZShcInN0eWxlXCIpID8/IFwiXCI7XG4gIHJldHVybiAvLS1uYXRpdmUtc2Nyb2xsYmFyLWNvbG9yLXNjaGVtZTpcXHMqZGFyay8udGVzdChzdHlsZSkgPyBcImRhcmtcIiA6IFwibGlnaHRcIjtcbn1cblxuLyoqIEVudHJpZXMgZXhwZWN0ZWQgb24gZXZlcnkgbWFpbCB2aWV3IHRoYXQgZGlkIG5vdCByZXNvbHZlLiBOb24tZW1wdHkgbWVhbnMgc29tZXRoaW5nIGJyb2tlLiAqL1xuZXhwb3J0IGZ1bmN0aW9uIGJyb2tlbkVudHJpZXMocmVzdWx0czogSGVhbHRoUmVzdWx0W10pOiBTZWxlY3Rvck5hbWVbXSB7XG4gIHJldHVybiByZXN1bHRzLmZpbHRlcigocikgPT4gci5leHBlY3QgPT09IFwiYWx3YXlzXCIgJiYgci5jYW5kaWRhdGUgPCAwKS5tYXAoKHIpID0+IHIubmFtZSk7XG59XG4iLCAiLy8gRGV2ZWxvcGVyIGNhcHR1cmU6IHNlcmlhbGlzZXMgdGhlIGxpdmUgT3V0bG9vayBwYWdlIGludG8gb25lIEhUTUwgZmlsZSwgc28gYSBmaXh0dXJlIGNhbiBiZVxuLy8gbWFkZSAoZG9jcy9jYXB0dXJpbmctcGFnZXMubWQpIHdpdGhvdXQgRGV2VG9vbHMuIFNhdmVkIGxvY2FsbHkgYnkgdGhlIGNhbGxlcjsgbmV2ZXIgc2VudFxuLy8gYW55d2hlcmUuIFRoZSBmaWxlIGhvbGRzIHJlYWwgbWFpbCwgc28gaXQgZ29lcyB0aHJvdWdoIHNjcmlwdHMvc2NydWItZml4dHVyZS50cyBiZWZvcmUgaXRcbi8vIGdvZXMgYW55d2hlcmUgbmVhciB0aGUgcmVwby5cblxuaW1wb3J0IHsgYnJva2VuRW50cmllcywgaGVhbHRoQ2hlY2sgfSBmcm9tIFwiLi9zZWxlY3RvcnMudHNcIjtcblxuLyoqIE1lc3NhZ2UgdGhlIHBvcHVwIHNlbmRzIHRvIHRoZSBjb250ZW50IHNjcmlwdCBvZiB0aGUgYWN0aXZlIHRhYi4gKi9cbmV4cG9ydCBjb25zdCBDQVBUVVJFX01FU1NBR0UgPSBcIm9iYi1jYXB0dXJlXCI7XG5cbmNvbnN0IE9CQl9BVFRSSUJVVEUgPSAvXmRhdGEtb3Zlcmxvb2svO1xuLyoqIFRoZSBzY3J1YmJlciAoc2NyaXB0cy9zY3J1Yi1maXh0dXJlLnRzKSBsb29rcyBmb3IgdGhpcyBtYXJrZXIgdG8gc3BsaXQgb3V0IHJ1bnRpbWUgc3R5bGVzLiAqL1xuZXhwb3J0IGNvbnN0IFJVTlRJTUVfU1RZTEVfTUFSS0VSID0gXCJkYXRhLW92ZXJsb29rLXJ1bnRpbWVcIjtcblxuZXhwb3J0IGludGVyZmFjZSBDYXB0dXJlTWV0YSB7XG4gIGNhcHR1cmVkQXQ6IHN0cmluZztcbiAgb2JiVmVyc2lvbjogc3RyaW5nO1xuICAvKiogRmlyc3QgdHdvIHBhdGggc2VnbWVudHMgb25seTogZGVlcGVyIG9uZXMgY2FycnkgbWVzc2FnZSBpZHMuICovXG4gIHJvdXRlOiBzdHJpbmc7XG4gIHZpZXdwb3J0OiBzdHJpbmc7XG4gIHNldHRpbmdzOiBSZWNvcmQ8c3RyaW5nLCBzdHJpbmcgfCB1bmRlZmluZWQ+O1xuICBoZWFsdGg6IHsgYnJva2VuOiBzdHJpbmdbXTsgZW50cmllczogeyBuYW1lOiBzdHJpbmc7IGNhbmRpZGF0ZTogbnVtYmVyOyBjb3VudDogbnVtYmVyIH1bXSB9O1xufVxuXG5leHBvcnQgaW50ZXJmYWNlIENhcHR1cmUge1xuICBodG1sOiBzdHJpbmc7XG4gIG5hbWU6IHN0cmluZztcbiAgbWV0YTogQ2FwdHVyZU1ldGE7XG59XG5cbi8qKiBTdHlsZSBydWxlcyBPdXRsb29rIGluamVjdGVkIGF0IHJ1bnRpbWUgKENTUy1pbi1KUyksIG1pbnVzIHRoZSBvbmVzIE9CQiBpdHNlbGYgYWRkZWQuICovXG5mdW5jdGlvbiBydW50aW1lQ3NzKGRvYzogRG9jdW1lbnQpOiBzdHJpbmcge1xuICBjb25zdCBzaGVldHMgPSBbLi4uZG9jLnN0eWxlU2hlZXRzLCAuLi4oZG9jLmFkb3B0ZWRTdHlsZVNoZWV0cyA/PyBbXSldLmZpbHRlcihcbiAgICAoc2hlZXQpID0+ICEoXCJocmVmXCIgaW4gc2hlZXQgJiYgc2hlZXQuaHJlZiksXG4gICk7XG4gIHJldHVybiBzaGVldHNcbiAgICAuZmxhdE1hcCgoc2hlZXQpID0+IHtcbiAgICAgIHRyeSB7XG4gICAgICAgIHJldHVybiBbLi4uc2hlZXQuY3NzUnVsZXNdLm1hcCgocnVsZSkgPT4gcnVsZS5jc3NUZXh0KTtcbiAgICAgIH0gY2F0Y2gge1xuICAgICAgICByZXR1cm4gW107IC8vIGNyb3NzLW9yaWdpbiBzaGVldFxuICAgICAgfVxuICAgIH0pXG4gICAgLmZpbHRlcigocnVsZSkgPT4gIXJ1bGUuaW5jbHVkZXMoXCJkYXRhLW92ZXJsb29rXCIpKVxuICAgIC5qb2luKFwiXFxuXCIpO1xufVxuXG5leHBvcnQgZnVuY3Rpb24gY2FwdHVyZUZpbGVuYW1lKG5vdyA9IG5ldyBEYXRlKCkpOiBzdHJpbmcge1xuICBjb25zdCBzdGFtcCA9IG5vdy50b0lTT1N0cmluZygpLnNsaWNlKDAsIDE2KS5yZXBsYWNlKC9cXEQvZywgXCJcIik7XG4gIHJldHVybiBgb2JiLWNhcHR1cmUtJHtzdGFtcH0uaHRtbGA7XG59XG5cbmV4cG9ydCBmdW5jdGlvbiBjYXB0dXJlUGFnZShcbiAgZG9jOiBEb2N1bWVudCA9IGRvY3VtZW50LFxuICBsb2M6IFBpY2s8TG9jYXRpb24sIFwicGF0aG5hbWVcIj4gPSBsb2NhdGlvbixcbiAgdmVyc2lvbiA9IFwiZGV2XCIsXG4gIG5vdyA9IG5ldyBEYXRlKCksXG4pOiBDYXB0dXJlIHtcbiAgY29uc3QgbGl2ZSA9IGRvYy5kb2N1bWVudEVsZW1lbnQuZGF0YXNldDtcbiAgY29uc3QgcmVzdWx0cyA9IGhlYWx0aENoZWNrKGRvYyk7XG4gIGNvbnN0IG1ldGE6IENhcHR1cmVNZXRhID0ge1xuICAgIGNhcHR1cmVkQXQ6IG5vdy50b0lTT1N0cmluZygpLFxuICAgIG9iYlZlcnNpb246IHZlcnNpb24sXG4gICAgcm91dGU6IGxvYy5wYXRobmFtZS5zcGxpdChcIi9cIikuc2xpY2UoMCwgMykuam9pbihcIi9cIiksXG4gICAgdmlld3BvcnQ6IGAke2RvYy5kZWZhdWx0Vmlldz8uaW5uZXJXaWR0aCA/PyAwfXgke2RvYy5kZWZhdWx0Vmlldz8uaW5uZXJIZWlnaHQgPz8gMH1gLFxuICAgIHNldHRpbmdzOiB7XG4gICAgICBmZWF0dXJlczogbGl2ZS5vdmVybG9va0ZlYXR1cmVzLFxuICAgICAgcGFsZXR0ZTogbGl2ZS5vdmVybG9va1BhbGV0dGUsXG4gICAgICBtb2RlOiBsaXZlLm92ZXJsb29rTW9kZSxcbiAgICAgIGxheW91dDogbGl2ZS5vdmVybG9va0xheW91dCxcbiAgICAgIG91dGxvb2tTY2hlbWU6IGxpdmUub3Zlcmxvb2tPdXRsb29rU2NoZW1lLFxuICAgICAgcmVhZGluZ1BhbmU6IGxpdmUub3Zlcmxvb2tSZWFkaW5nUGFuZSxcbiAgICB9LFxuICAgIGhlYWx0aDoge1xuICAgICAgYnJva2VuOiBicm9rZW5FbnRyaWVzKHJlc3VsdHMpLFxuICAgICAgZW50cmllczogcmVzdWx0cy5tYXAoKHsgbmFtZSwgY2FuZGlkYXRlLCBjb3VudCB9KSA9PiAoeyBuYW1lLCBjYW5kaWRhdGUsIGNvdW50IH0pKSxcbiAgICB9LFxuICB9O1xuXG4gIC8vIFdvcmsgb24gYSBjb3B5LCB3aXRoIE9CQidzIG93biBtYXJrcyByZW1vdmVkLCBzbyB0aGUgY2FwdHVyZSBpcyBzdG9jayBPdXRsb29rLlxuICBjb25zdCByb290ID0gZG9jLmRvY3VtZW50RWxlbWVudC5jbG9uZU5vZGUodHJ1ZSkgYXMgSFRNTEVsZW1lbnQ7XG4gIGZvciAoY29uc3QgZWwgb2Ygcm9vdC5xdWVyeVNlbGVjdG9yQWxsKFwib3Zlcmxvb2stcm9vdFwiKSkgZWwucmVtb3ZlKCk7XG4gIGZvciAoY29uc3QgZWwgb2YgW3Jvb3QsIC4uLnJvb3QucXVlcnlTZWxlY3RvckFsbChcIipcIildKSB7XG4gICAgZm9yIChjb25zdCB7IG5hbWUgfSBvZiBbLi4uZWwuYXR0cmlidXRlc10pIHtcbiAgICAgIGlmIChPQkJfQVRUUklCVVRFLnRlc3QobmFtZSkpIGVsLnJlbW92ZUF0dHJpYnV0ZShuYW1lKTtcbiAgICB9XG4gIH1cblxuICBjb25zdCBoZWFkID0gcm9vdC5xdWVyeVNlbGVjdG9yKFwiaGVhZFwiKTtcbiAgY29uc3Qgc3R5bGUgPSBkb2MuY3JlYXRlRWxlbWVudChcInN0eWxlXCIpO1xuICBzdHlsZS5zZXRBdHRyaWJ1dGUoUlVOVElNRV9TVFlMRV9NQVJLRVIsIFwiXCIpO1xuICBzdHlsZS50ZXh0Q29udGVudCA9IHJ1bnRpbWVDc3MoZG9jKTtcbiAgaGVhZD8uYXBwZW5kKHN0eWxlKTtcbiAgY29uc3Qgc2NyaXB0ID0gZG9jLmNyZWF0ZUVsZW1lbnQoXCJzY3JpcHRcIik7XG4gIHNjcmlwdC50eXBlID0gXCJhcHBsaWNhdGlvbi9qc29uXCI7XG4gIHNjcmlwdC5pZCA9IFwib2JiLWNhcHR1cmUtbWV0YVwiO1xuICBzY3JpcHQudGV4dENvbnRlbnQgPSBKU09OLnN0cmluZ2lmeShtZXRhLCBudWxsLCAyKS5yZXBsYWNlQWxsKFwiPFwiLCBcIlxcXFx1MDAzY1wiKTtcbiAgaGVhZD8uYXBwZW5kKHNjcmlwdCk7XG5cbiAgcmV0dXJuIHsgaHRtbDogYDwhZG9jdHlwZSBodG1sPlxcbiR7cm9vdC5vdXRlckhUTUx9YCwgbmFtZTogY2FwdHVyZUZpbGVuYW1lKG5vdyksIG1ldGEgfTtcbn1cbiIsICIvLyBVc2VyIHByZWZlcmVuY2VzLiBTdG9yZWQgaW4gY2hyb21lLnN0b3JhZ2Uuc3luYzsgbm90aGluZyBlbHNlIGlzIGV2ZXIgcGVyc2lzdGVkLlxuLy9cbi8vIEVhY2ggZW5hYmxlZCBmZWF0dXJlIGlkIGlzIHdyaXR0ZW4gYXMgYSB0b2tlbiBvbiA8aHRtbCBkYXRhLW92ZXJsb29rLWZlYXR1cmVzPVwiLi4uXCI+LCBhbmQgZWFjaFxuLy8gY2hvaWNlIGFzIGl0cyBvd24gYXR0cmlidXRlICg8aHRtbCBkYXRhLW92ZXJsb29rLXBhbGV0dGU9XCJzb2xhcml6ZWRcIj4pLCBzbyBzdHlsZXNoZWV0cyBnYXRlXG4vLyB0aGVtc2VsdmVzIHdpdGggYFtkYXRhLW92ZXJsb29rLWZlYXR1cmVzfj1cImRlbnNpdHlcIl1gIG9yIGBbZGF0YS1vdmVybG9vay1wYWxldHRlPVwic29sYXJpemVkXCJdYC5cblxuZXhwb3J0IHR5cGUgRmVhdHVyZUlkID0ga2V5b2YgdHlwZW9mIEZFQVRVUkVTO1xuZXhwb3J0IHR5cGUgQ2hvaWNlSWQgPSBrZXlvZiB0eXBlb2YgQ0hPSUNFUztcbmV4cG9ydCB0eXBlIFNldHRpbmdzID0gUmVjb3JkPEZlYXR1cmVJZCwgYm9vbGVhbj4gJiB7IFtLIGluIENob2ljZUlkXTogQ2hvaWNlVmFsdWU8Sz4gfTtcbmV4cG9ydCB0eXBlIENob2ljZVZhbHVlPEsgZXh0ZW5kcyBDaG9pY2VJZD4gPSBrZXlvZiAodHlwZW9mIENIT0lDRVMpW0tdW1wib3B0aW9uc1wiXTtcblxudHlwZSBHcm91cCA9IFwiR2VuZXJhbFwiIHwgXCJLZXlib2FyZFwiIHwgXCJEZWNsdXR0ZXJcIiB8IFwiTG9va1wiO1xuXG5pbnRlcmZhY2UgRmVhdHVyZSB7XG4gIGxhYmVsOiBzdHJpbmc7XG4gIGdyb3VwOiBHcm91cDtcbiAgZGVmYXVsdDogYm9vbGVhbjtcbn1cblxuaW50ZXJmYWNlIENob2ljZSB7XG4gIGxhYmVsOiBzdHJpbmc7XG4gIGdyb3VwOiBHcm91cDtcbiAgb3B0aW9uczogUmVjb3JkPHN0cmluZywgc3RyaW5nPjtcbiAgZGVmYXVsdDogc3RyaW5nO1xufVxuXG5leHBvcnQgY29uc3QgRkVBVFVSRVMgPSB7XG4gIGhlYWx0aEluZGljYXRvcjoge1xuICAgIGxhYmVsOiBcIlNob3cgYW4gaW5kaWNhdG9yIHdoZW4gT3V0bG9vaydzIHBhZ2UgY2hhbmdlcyBicmVhayBPQkJcIixcbiAgICBncm91cDogXCJHZW5lcmFsXCIsXG4gICAgZGVmYXVsdDogdHJ1ZSxcbiAgfSxcbiAga2V5bWFwOiB7IGxhYmVsOiBcIkZhc3RtYWlsLXN0eWxlIGtleWJvYXJkIHNob3J0Y3V0c1wiLCBncm91cDogXCJLZXlib2FyZFwiLCBkZWZhdWx0OiB0cnVlIH0sXG4gIFwiaGlkZS11cHNlbGxcIjogeyBsYWJlbDogXCJVcGdyYWRlIGFuZCAndHJ5IHRoZSBuZXdcdTIwMjYnIGJhbm5lcnNcIiwgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsIGRlZmF1bHQ6IHRydWUgfSxcbiAgXCJoaWRlLWFwcC1iYW5uZXJzXCI6IHtcbiAgICBsYWJlbDogXCJOb3RpY2VzIGFjcm9zcyB0aGUgdG9wIG9mIE91dGxvb2sgKHBvcC11cCwgc2lnbi1pbiBhbmQgc3RvcmFnZSByZW1pbmRlcnMpXCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogZmFsc2UsXG4gIH0sXG4gIFwiaGlkZS1pbmZvLWJhcnNcIjoge1xuICAgIGxhYmVsOiBcIk5vdGljZXMgYWJvdmUgYW4gb3BlbiBtZXNzYWdlICgnWW91IHJlcGxpZWQgb25cdTIwMjYnLCAnRXh0cmEgbGluZSBicmVha3MgcmVtb3ZlZCcpXCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogZmFsc2UsXG4gIH0sXG4gIFwiaGlkZS1jb3BpbG90XCI6IHsgbGFiZWw6IFwiQ29waWxvdCBidXR0b25zIGFuZCBwYW5lbHNcIiwgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsIGRlZmF1bHQ6IHRydWUgfSxcbiAgXCJoaWRlLWhlYWRlci1leHRyYXNcIjoge1xuICAgIGxhYmVsOiBcIlRlYW1zIGFuZCBNeSBEYXkgYnV0dG9ucyBpbiB0aGUgaGVhZGVyXCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogdHJ1ZSxcbiAgfSxcbiAgXCJoaWRlLXJhaWwtZXh0cmFzXCI6IHtcbiAgICBsYWJlbDogXCJUbyBEbywgTmV3c2xldHRlcnMsIE9yZyBFeHBsb3JlciBhbmQgT25lRHJpdmUgaW4gdGhlIGFwcCByYWlsXCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogdHJ1ZSxcbiAgfSxcbiAgXCJoaWRlLXJpYmJvbi1leHRyYXNcIjoge1xuICAgIGxhYmVsOiBcIlRlYW1zLCBTd2VlcCwgUXVpY2sgc3RlcHMsIFZpdmEgYW5kIGFkZC1pbiBidXR0b25zIGluIHRoZSByaWJib25cIixcbiAgICBncm91cDogXCJEZWNsdXR0ZXJcIixcbiAgICBkZWZhdWx0OiB0cnVlLFxuICB9LFxuICBcImhpZGUtZXh0ZXJuYWwtdGFnXCI6IHsgbGFiZWw6IFwiJ0V4dGVybmFsJyBzZW5kZXIgdGFnc1wiLCBncm91cDogXCJEZWNsdXR0ZXJcIiwgZGVmYXVsdDogdHJ1ZSB9LFxuICBcImhpZGUtcm93LWFjdGlvbnNcIjoge1xuICAgIGxhYmVsOiBcIkZsYWcgYW5kIHBpbiBpY29ucyB0aGF0IGFwcGVhciB3aGVuIGhvdmVyaW5nIGEgbWVzc2FnZVwiLFxuICAgIGdyb3VwOiBcIkRlY2x1dHRlclwiLFxuICAgIGRlZmF1bHQ6IGZhbHNlLFxuICB9LFxuICBcImhpZGUtYXZhdGFyc1wiOiB7XG4gICAgbGFiZWw6IFwiU2VuZGVyIHBpY3R1cmVzIGluIHRoZSBtZXNzYWdlIGxpc3RcIixcbiAgICBncm91cDogXCJEZWNsdXR0ZXJcIixcbiAgICBkZWZhdWx0OiBmYWxzZSxcbiAgfSxcbiAgXCJoaWRlLWdyb3VwLWhlYWRlcnNcIjoge1xuICAgIGxhYmVsOiBcIkRhdGUgZ3JvdXAgaGVhZGVycyAoJ1RoaXMgTW9udGgnKSBpbiB0aGUgbWVzc2FnZSBsaXN0XCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogZmFsc2UsXG4gIH0sXG4gIFwiaGlkZS1mb2N1c2VkLXRhYnNcIjogeyBsYWJlbDogXCJGb2N1c2VkIC8gT3RoZXIgdGFic1wiLCBncm91cDogXCJEZWNsdXR0ZXJcIiwgZGVmYXVsdDogZmFsc2UgfSxcbiAgXCJoaWRlLXJpYmJvbi10YWJzXCI6IHtcbiAgICBsYWJlbDogXCJSaWJib24gdGFicyAoSG9tZSwgVmlldywgSGVscClcIixcbiAgICBncm91cDogXCJEZWNsdXR0ZXJcIixcbiAgICBkZWZhdWx0OiBmYWxzZSxcbiAgfSxcbiAgXCJoaWRlLWFwcC1yYWlsXCI6IHsgbGFiZWw6IFwiQXBwIHJhaWwgKGxlZnQgaWNvbiBiYXIpXCIsIGdyb3VwOiBcIkRlY2x1dHRlclwiLCBkZWZhdWx0OiBmYWxzZSB9LFxuICBcImhpZGUtcmliYm9uXCI6IHsgbGFiZWw6IFwiUmliYm9uXCIsIGdyb3VwOiBcIkRlY2x1dHRlclwiLCBkZWZhdWx0OiBmYWxzZSB9LFxuICBkZW5zaXR5OiB7IGxhYmVsOiBcIlRpZ2h0ZXIsIGRlbnNlciBsYXlvdXRcIiwgZ3JvdXA6IFwiTG9va1wiLCBkZWZhdWx0OiB0cnVlIH0sXG59IGFzIGNvbnN0IHNhdGlzZmllcyBSZWNvcmQ8c3RyaW5nLCBGZWF0dXJlPjtcblxuZXhwb3J0IGNvbnN0IENIT0lDRVMgPSB7XG4gIHBhbGV0dGU6IHtcbiAgICBsYWJlbDogXCJDb2xvcnNcIixcbiAgICBncm91cDogXCJMb29rXCIsXG4gICAgb3B0aW9uczoge1xuICAgICAgb3V0bG9vazogXCJPdXRsb29rJ3Mgb3duXCIsXG4gICAgICBzb2xhcml6ZWQ6IFwiU29sYXJpemVkXCIsXG4gICAgICBiYXNlMTY6IFwiQmFzZTE2IERlZmF1bHRcIixcbiAgICAgIGZhc3RtYWlsOiBcIkZhc3RtYWlsXCIsXG4gICAgICBwYXBlcjogXCJQYXBlclwiLFxuICAgIH0sXG4gICAgZGVmYXVsdDogXCJvdXRsb29rXCIsXG4gIH0sXG4gIG1vZGU6IHtcbiAgICBsYWJlbDogXCJMaWdodCBvciBkYXJrXCIsXG4gICAgZ3JvdXA6IFwiTG9va1wiLFxuICAgIG9wdGlvbnM6IHsgYXV0bzogXCJGb2xsb3cgT3V0bG9va1wiLCBsaWdodDogXCJMaWdodFwiLCBkYXJrOiBcIkRhcmtcIiB9LFxuICAgIGRlZmF1bHQ6IFwiYXV0b1wiLFxuICB9LFxuICBsYXlvdXQ6IHtcbiAgICBsYWJlbDogXCJMYXlvdXRcIixcbiAgICBncm91cDogXCJMb29rXCIsXG4gICAgb3B0aW9uczoge1xuICAgICAgb3V0bG9vazogXCJPdXRsb29rJ3Mgb3duXCIsXG4gICAgICBjYWxtOiBcIkNhbG0gdGhyZWUtcGFuZVwiLFxuICAgICAgZmFzdG1haWw6IFwiRmFzdG1haWwgKGxpc3QgZmlyc3QpXCIsXG4gICAgICBmb2N1czogXCJGb2N1cyAoc2luZ2xlIGNvbHVtbilcIixcbiAgICB9LFxuICAgIGRlZmF1bHQ6IFwib3V0bG9va1wiLFxuICB9LFxufSBhcyBjb25zdCBzYXRpc2ZpZXMgUmVjb3JkPHN0cmluZywgQ2hvaWNlPjtcblxuZXhwb3J0IGNvbnN0IEZFQVRVUkVfSURTID0gT2JqZWN0LmtleXMoRkVBVFVSRVMpIGFzIEZlYXR1cmVJZFtdO1xuZXhwb3J0IGNvbnN0IENIT0lDRV9JRFMgPSBPYmplY3Qua2V5cyhDSE9JQ0VTKSBhcyBDaG9pY2VJZFtdO1xuXG5leHBvcnQgZnVuY3Rpb24gZGVmYXVsdFNldHRpbmdzKCk6IFNldHRpbmdzIHtcbiAgcmV0dXJuIHtcbiAgICAuLi5PYmplY3QuZnJvbUVudHJpZXMoRkVBVFVSRV9JRFMubWFwKChpZCkgPT4gW2lkLCBGRUFUVVJFU1tpZF0uZGVmYXVsdF0pKSxcbiAgICAuLi5PYmplY3QuZnJvbUVudHJpZXMoQ0hPSUNFX0lEUy5tYXAoKGlkKSA9PiBbaWQsIENIT0lDRVNbaWRdLmRlZmF1bHRdKSksXG4gIH0gYXMgU2V0dGluZ3M7XG59XG5cbi8qKiBNZXJnZSB3aGF0ZXZlciBpcyBzdG9yZWQgb3ZlciB0aGUgZGVmYXVsdHMsIGlnbm9yaW5nIHVua25vd24gb3IgbWlzdHlwZWQga2V5cy4gKi9cbmV4cG9ydCBmdW5jdGlvbiBub3JtYWxpemUoc3RvcmVkOiBSZWNvcmQ8c3RyaW5nLCB1bmtub3duPik6IFNldHRpbmdzIHtcbiAgY29uc3Qgc2V0dGluZ3MgPSBkZWZhdWx0U2V0dGluZ3MoKSBhcyBSZWNvcmQ8c3RyaW5nLCB1bmtub3duPjtcbiAgZm9yIChjb25zdCBpZCBvZiBGRUFUVVJFX0lEUykge1xuICAgIGlmICh0eXBlb2Ygc3RvcmVkW2lkXSA9PT0gXCJib29sZWFuXCIpIHNldHRpbmdzW2lkXSA9IHN0b3JlZFtpZF07XG4gIH1cbiAgZm9yIChjb25zdCBpZCBvZiBDSE9JQ0VfSURTKSB7XG4gICAgY29uc3QgdmFsdWUgPSBzdG9yZWRbaWRdO1xuICAgIGlmICh0eXBlb2YgdmFsdWUgPT09IFwic3RyaW5nXCIgJiYgdmFsdWUgaW4gQ0hPSUNFU1tpZF0ub3B0aW9ucykgc2V0dGluZ3NbaWRdID0gdmFsdWU7XG4gIH1cbiAgcmV0dXJuIHNldHRpbmdzIGFzIFNldHRpbmdzO1xufVxuXG5leHBvcnQgZnVuY3Rpb24gZmVhdHVyZVRva2VucyhzZXR0aW5nczogU2V0dGluZ3MpOiBzdHJpbmcge1xuICByZXR1cm4gRkVBVFVSRV9JRFMuZmlsdGVyKChpZCkgPT4gc2V0dGluZ3NbaWRdKS5qb2luKFwiIFwiKTtcbn1cblxuLyoqIFdyaXRlcyB0aGUgc2V0dGluZ3Mgb250byA8aHRtbD4gYXMgdGhlIGF0dHJpYnV0ZXMgdGhlIHN0eWxlc2hlZXRzIGtleSBvZmYuICovXG5leHBvcnQgZnVuY3Rpb24gYXBwbHlUb0RvY3VtZW50KHNldHRpbmdzOiBTZXR0aW5ncywgcm9vdDogSFRNTEVsZW1lbnQpOiB2b2lkIHtcbiAgcm9vdC5zZXRBdHRyaWJ1dGUoXCJkYXRhLW92ZXJsb29rLWZlYXR1cmVzXCIsIGZlYXR1cmVUb2tlbnMoc2V0dGluZ3MpKTtcbiAgZm9yIChjb25zdCBpZCBvZiBDSE9JQ0VfSURTKSByb290LnNldEF0dHJpYnV0ZShgZGF0YS1vdmVybG9vay0ke2lkfWAsIHNldHRpbmdzW2lkXSk7XG59XG5cbmV4cG9ydCBhc3luYyBmdW5jdGlvbiBsb2FkU2V0dGluZ3MoKTogUHJvbWlzZTxTZXR0aW5ncz4ge1xuICByZXR1cm4gbm9ybWFsaXplKGF3YWl0IGNocm9tZS5zdG9yYWdlLnN5bmMuZ2V0KG51bGwpKTtcbn1cblxuZXhwb3J0IGFzeW5jIGZ1bmN0aW9uIHNhdmVTZXR0aW5nKFxuICBpZDogRmVhdHVyZUlkIHwgQ2hvaWNlSWQsXG4gIHZhbHVlOiBib29sZWFuIHwgc3RyaW5nLFxuKTogUHJvbWlzZTx2b2lkPiB7XG4gIGF3YWl0IGNocm9tZS5zdG9yYWdlLnN5bmMuc2V0KHsgW2lkXTogdmFsdWUgfSk7XG59XG5cbmV4cG9ydCBmdW5jdGlvbiBvblNldHRpbmdzQ2hhbmdlZChjYWxsYmFjazogKHNldHRpbmdzOiBTZXR0aW5ncykgPT4gdm9pZCk6IHZvaWQge1xuICBjaHJvbWUuc3RvcmFnZS5vbkNoYW5nZWQuYWRkTGlzdGVuZXIoKF9jaGFuZ2VzLCBhcmVhKSA9PiB7XG4gICAgaWYgKGFyZWEgPT09IFwic3luY1wiKSB2b2lkIGxvYWRTZXR0aW5ncygpLnRoZW4oY2FsbGJhY2spO1xuICB9KTtcbn1cbiIsICIvLyBUaGUgb3B0aW9ucyBwYWdlLCBhbHNvIHRoZSB0b29sYmFyIHBvcHVwLiBDaGFuZ2VzIHNhdmUgaW1tZWRpYXRlbHkgYW5kIGFwcGx5IGxpdmUgaW4gT3V0bG9vay5cblxuaW1wb3J0IHsgQ0FQVFVSRV9NRVNTQUdFIH0gZnJvbSBcIi4uL2NhcHR1cmUudHNcIjtcbmltcG9ydCB7XG4gIENIT0lDRVMsXG4gIHR5cGUgQ2hvaWNlSWQsXG4gIEZFQVRVUkVfSURTLFxuICBGRUFUVVJFUyxcbiAgdHlwZSBGZWF0dXJlSWQsXG4gIGxvYWRTZXR0aW5ncyxcbiAgb25TZXR0aW5nc0NoYW5nZWQsXG4gIHR5cGUgU2V0dGluZ3MsXG4gIHNhdmVTZXR0aW5nLFxufSBmcm9tIFwiLi4vc2V0dGluZ3MudHNcIjtcblxuY29uc3QgU0hPUlRfTEFZT1VUX05BTUVTOiBSZWNvcmQ8c3RyaW5nLCBzdHJpbmc+ID0ge1xuICBvdXRsb29rOiBcIk91dGxvb2tcIixcbiAgY2FsbTogXCJDYWxtXCIsXG4gIGZhc3RtYWlsOiBcIkZhc3RtYWlsXCIsXG4gIGZvY3VzOiBcIkZvY3VzXCIsXG59O1xuXG5jb25zdCBTSE9SVF9QQUxFVFRFX05BTUVTOiBSZWNvcmQ8c3RyaW5nLCBzdHJpbmc+ID0ge1xuICBvdXRsb29rOiBcIk91dGxvb2tcIixcbiAgc29sYXJpemVkOiBcIlNvbGFyaXplZFwiLFxuICBiYXNlMTY6IFwiQmFzZTE2XCIsXG4gIGZhc3RtYWlsOiBcIkZhc3RtYWlsXCIsXG4gIHBhcGVyOiBcIlBhcGVyXCIsXG59O1xuXG4vKiogVHdvIGxpbmVzIG9mIGEgbWVzc2FnZSByb3cgKHN1YmplY3QsIHByZXZpZXcpIGF0IHRoZSBnaXZlbiBwb3NpdGlvbiBpbiBhIHNrZXRjaC4gKi9cbmZ1bmN0aW9uIHJvd3MoeDogbnVtYmVyLCB5OiBudW1iZXIsIHdpZHRoOiBudW1iZXIpOiBzdHJpbmcge1xuICByZXR1cm4gWzAsIDEsIDIsIDMsIDRdXG4gICAgLm1hcCgoaSkgPT4ge1xuICAgICAgY29uc3QgdG9wID0geSArIGkgKiA2LjU7XG4gICAgICBjb25zdCBzdWJqZWN0ID0gaSA9PT0gMCA/IFwiaGlcIiA6IFwicm93XCI7XG4gICAgICBjb25zdCBzdWJqZWN0V2lkdGggPSBpID09PSAwID8gd2lkdGggKiAwLjU1IDogd2lkdGg7XG4gICAgICByZXR1cm4gKFxuICAgICAgICBgPHJlY3QgY2xhc3M9XCIke3N1YmplY3R9XCIgeD1cIiR7eH1cIiB5PVwiJHt0b3B9XCIgd2lkdGg9XCIke3N1YmplY3RXaWR0aH1cIiBoZWlnaHQ9XCIyLjVcIiByeD1cIjFcIi8+YCArXG4gICAgICAgIGA8cmVjdCBjbGFzcz1cInJvd1wiIHg9XCIke3h9XCIgeT1cIiR7dG9wICsgM31cIiB3aWR0aD1cIiR7d2lkdGggKiAwLjc1fVwiIGhlaWdodD1cIjEuNVwiIHJ4PVwiMC43NVwiLz5gXG4gICAgICApO1xuICAgIH0pXG4gICAgLmpvaW4oXCJcIik7XG59XG5cbi8vIFdpcmVmcmFtZXMgb2YgZWFjaCBsYXlvdXQ6IHZpZXdCb3ggNjRcdTAwRDc0NDsgY2hyb21lIGJhciwgdGhlbiBwYW5lcyBhbmQgbWVzc2FnZSByb3dzLlxuY29uc3QgTEFZT1VUX1NLRVRDSEVTOiBSZWNvcmQ8c3RyaW5nLCBzdHJpbmc+ID0ge1xuICBvdXRsb29rOiBgPHJlY3QgY2xhc3M9XCJjaHJvbWVcIiB3aWR0aD1cIjY0XCIgaGVpZ2h0PVwiNVwiLz5cbiAgICA8cmVjdCBjbGFzcz1cInBhbmVcIiB4PVwiMFwiIHk9XCI1XCIgd2lkdGg9XCI1XCIgaGVpZ2h0PVwiMzlcIi8+XG4gICAgPHJlY3QgY2xhc3M9XCJwYW5lXCIgeD1cIjZcIiB5PVwiN1wiIHdpZHRoPVwiMTJcIiBoZWlnaHQ9XCIzNVwiIHJ4PVwiMVwiLz5cbiAgICAke3Jvd3MoMjAsIDgsIDE1KX08cmVjdCBjbGFzcz1cInBhbmVcIiB4PVwiMzdcIiB5PVwiN1wiIHdpZHRoPVwiMjVcIiBoZWlnaHQ9XCIzNVwiIHJ4PVwiMVwiLz5gLFxuICBjYWxtOiBgPHJlY3QgY2xhc3M9XCJjaHJvbWVcIiB3aWR0aD1cIjY0XCIgaGVpZ2h0PVwiNVwiLz5cbiAgICA8cmVjdCBjbGFzcz1cInBhbmVcIiB4PVwiMlwiIHk9XCI4XCIgd2lkdGg9XCIxM1wiIGhlaWdodD1cIjMzXCIgcng9XCIyXCIvPlxuICAgICR7cm93cygxOCwgOSwgMTYpfTxyZWN0IGNsYXNzPVwicGFuZVwiIHg9XCIzN1wiIHk9XCI4XCIgd2lkdGg9XCIyNVwiIGhlaWdodD1cIjMzXCIgcng9XCIyXCIvPmAsXG4gIGZhc3RtYWlsOiBgPHJlY3QgY2xhc3M9XCJjaHJvbWVcIiB3aWR0aD1cIjY0XCIgaGVpZ2h0PVwiNVwiLz5cbiAgICA8cmVjdCBjbGFzcz1cInBhbmVcIiB4PVwiMlwiIHk9XCI4XCIgd2lkdGg9XCIxM1wiIGhlaWdodD1cIjMzXCIgcng9XCIyXCIvPlxuICAgIDxyZWN0IGNsYXNzPVwiaGlcIiB4PVwiNFwiIHk9XCIxMFwiIHdpZHRoPVwiOVwiIGhlaWdodD1cIjNcIiByeD1cIjEuNVwiLz4ke3Jvd3MoMTgsIDksIDQ0KX1gLFxuICBmb2N1czogYDxyZWN0IGNsYXNzPVwiY2hyb21lXCIgd2lkdGg9XCI2NFwiIGhlaWdodD1cIjVcIi8+JHtyb3dzKDE2LCA5LCAzMil9YCxcbn07XG5cbmNvbnN0IGh0bWwgPSBkb2N1bWVudC5kb2N1bWVudEVsZW1lbnQ7XG5jb25zdCBwcmVmZXJzRGFyayA9IG1hdGNoTWVkaWEoXCIocHJlZmVycy1jb2xvci1zY2hlbWU6IGRhcmspXCIpO1xuXG4vKiogVGhlIGxpZ2h0L2RhcmsgbW9kZSBpbiBlZmZlY3Q6IFwiYXV0b1wiIGZvbGxvd3MgdGhlIHN5c3RlbSBoZXJlLCBzdGFuZGluZyBpbiBmb3IgT3V0bG9vay4gKi9cbmZ1bmN0aW9uIGVmZmVjdGl2ZU1vZGUoc2V0dGluZ3M6IFNldHRpbmdzKTogXCJsaWdodFwiIHwgXCJkYXJrXCIge1xuICBpZiAoc2V0dGluZ3MubW9kZSA9PT0gXCJsaWdodFwiIHx8IHNldHRpbmdzLm1vZGUgPT09IFwiZGFya1wiKSByZXR1cm4gc2V0dGluZ3MubW9kZTtcbiAgcmV0dXJuIHByZWZlcnNEYXJrLm1hdGNoZXMgPyBcImRhcmtcIiA6IFwibGlnaHRcIjtcbn1cblxuLyoqIERyZXNzZXMgdGhlIHBhZ2UgaW4gdGhlIGNob3NlbiBwYWxldHRlIChwYWxldHRlcy5jc3Mga2V5cyBvZmYgdGhlc2UgYXR0cmlidXRlcykuICovXG5mdW5jdGlvbiBhcHBseVRoZW1lKHNldHRpbmdzOiBTZXR0aW5ncyk6IHZvaWQge1xuICBodG1sLmRhdGFzZXQub3Zlcmxvb2tQYWxldHRlID0gc2V0dGluZ3MucGFsZXR0ZTtcbiAgaHRtbC5kYXRhc2V0Lm92ZXJsb29rTW9kZSA9IHNldHRpbmdzLm1vZGU7XG4gIGh0bWwuZGF0YXNldC5vdmVybG9va091dGxvb2tTY2hlbWUgPSBwcmVmZXJzRGFyay5tYXRjaGVzID8gXCJkYXJrXCIgOiBcImxpZ2h0XCI7XG59XG5cbmZ1bmN0aW9uIGVsPEsgZXh0ZW5kcyBrZXlvZiBIVE1MRWxlbWVudFRhZ05hbWVNYXA+KFxuICB0YWc6IEssXG4gIGF0dHJzOiBSZWNvcmQ8c3RyaW5nLCBzdHJpbmc+ID0ge30sXG4gIC4uLmNoaWxkcmVuOiAoTm9kZSB8IHN0cmluZylbXVxuKTogSFRNTEVsZW1lbnRUYWdOYW1lTWFwW0tdIHtcbiAgY29uc3Qgbm9kZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQodGFnKTtcbiAgZm9yIChjb25zdCBbaywgdl0gb2YgT2JqZWN0LmVudHJpZXMoYXR0cnMpKSBub2RlLnNldEF0dHJpYnV0ZShrLCB2KTtcbiAgbm9kZS5hcHBlbmQoLi4uY2hpbGRyZW4pO1xuICByZXR1cm4gbm9kZTtcbn1cblxuZnVuY3Rpb24gc2VjdGlvbih0aXRsZTogc3RyaW5nLCAuLi5jaGlsZHJlbjogTm9kZVtdKTogSFRNTEVsZW1lbnQge1xuICByZXR1cm4gZWwoXCJzZWN0aW9uXCIsIHt9LCBlbChcImgyXCIsIHt9LCB0aXRsZSksIC4uLmNoaWxkcmVuKTtcbn1cblxuLyoqIE11dHVhbGx5IGV4Y2x1c2l2ZSBidXR0b25zIGZvciBhIGNob2ljZTsgYHJlbmRlcmAgZHJhd3MgZWFjaCBvcHRpb24uICovXG5mdW5jdGlvbiBjaG9pY2VHcm91cChcbiAgaWQ6IENob2ljZUlkLFxuICBzZXR0aW5nczogU2V0dGluZ3MsXG4gIGNsYXNzTmFtZTogc3RyaW5nLFxuICByZW5kZXI6ICh2YWx1ZTogc3RyaW5nLCBsYWJlbDogc3RyaW5nKSA9PiBIVE1MQnV0dG9uRWxlbWVudCxcbik6IEhUTUxFbGVtZW50IHtcbiAgY29uc3QgZ3JvdXAgPSBlbChcImRpdlwiLCB7XG4gICAgY2xhc3M6IGNsYXNzTmFtZSxcbiAgICByb2xlOiBcInJhZGlvZ3JvdXBcIixcbiAgICBcImFyaWEtbGFiZWxcIjogQ0hPSUNFU1tpZF0ubGFiZWwsXG4gIH0pO1xuICBmb3IgKGNvbnN0IFt2YWx1ZSwgbGFiZWxdIG9mIE9iamVjdC5lbnRyaWVzKENIT0lDRVNbaWRdLm9wdGlvbnMpKSB7XG4gICAgY29uc3QgYnV0dG9uID0gcmVuZGVyKHZhbHVlLCBsYWJlbCk7XG4gICAgYnV0dG9uLnR5cGUgPSBcImJ1dHRvblwiO1xuICAgIGJ1dHRvbi5zZXRBdHRyaWJ1dGUoXCJyb2xlXCIsIFwicmFkaW9cIik7XG4gICAgYnV0dG9uLnNldEF0dHJpYnV0ZShcImFyaWEtY2hlY2tlZFwiLCBTdHJpbmcoc2V0dGluZ3NbaWRdID09PSB2YWx1ZSkpO1xuICAgIGJ1dHRvbi5hZGRFdmVudExpc3RlbmVyKFwiY2xpY2tcIiwgKCkgPT4gdm9pZCBzYXZlU2V0dGluZyhpZCwgdmFsdWUpKTtcbiAgICBncm91cC5hcHBlbmQoYnV0dG9uKTtcbiAgfVxuICByZXR1cm4gZ3JvdXA7XG59XG5cbi8qKiBBIHRpbnkgd2luZG93IHBhaW50ZWQgaW4gaXRzIG93biBwYWxldHRlLCBpbiB0aGUgbW9kZSBjdXJyZW50bHkgaW4gZWZmZWN0LiAqL1xuZnVuY3Rpb24gc3dhdGNoKHNldHRpbmdzOiBTZXR0aW5ncykge1xuICByZXR1cm4gKHZhbHVlOiBzdHJpbmcsIGxhYmVsOiBzdHJpbmcpOiBIVE1MQnV0dG9uRWxlbWVudCA9PiB7XG4gICAgY29uc3QgcHJldmlldyA9IGVsKFxuICAgICAgXCJzcGFuXCIsXG4gICAgICB7XG4gICAgICAgIGNsYXNzOiBcInByZXZpZXdcIixcbiAgICAgICAgXCJkYXRhLW92ZXJsb29rLXN3YXRjaFwiOiBcIlwiLFxuICAgICAgICBcImRhdGEtb3Zlcmxvb2stcGFsZXR0ZVwiOiB2YWx1ZSxcbiAgICAgICAgXCJkYXRhLW92ZXJsb29rLW1vZGVcIjogZWZmZWN0aXZlTW9kZShzZXR0aW5ncyksXG4gICAgICB9LFxuICAgICAgZWwoXCJzcGFuXCIsIHsgY2xhc3M6IFwiYmFyXCIgfSksXG4gICAgICBlbChcInNwYW5cIiwgeyBjbGFzczogXCJzaWRlXCIgfSksXG4gICAgICBlbChcbiAgICAgICAgXCJzcGFuXCIsXG4gICAgICAgIHsgY2xhc3M6IFwiYm9keVwiIH0sXG4gICAgICAgIGVsKFwic3BhblwiLCB7IGNsYXNzOiBcImxpbmUgc2hvcnRcIiB9KSxcbiAgICAgICAgZWwoXCJzcGFuXCIsIHsgY2xhc3M6IFwibGluZVwiIH0pLFxuICAgICAgICBlbChcInNwYW5cIiwgeyBjbGFzczogXCJsaW5lIGZhaW50XCIgfSksXG4gICAgICAgIGVsKFwic3BhblwiLCB7IGNsYXNzOiBcImxpbmVcIiB9KSxcbiAgICAgICksXG4gICAgKTtcbiAgICByZXR1cm4gZWwoXG4gICAgICBcImJ1dHRvblwiLFxuICAgICAgeyBjbGFzczogXCJzd2F0Y2hcIiwgdGl0bGU6IGxhYmVsIH0sXG4gICAgICBwcmV2aWV3LFxuICAgICAgU0hPUlRfUEFMRVRURV9OQU1FU1t2YWx1ZV0gPz8gbGFiZWwsXG4gICAgKTtcbiAgfTtcbn1cblxuZnVuY3Rpb24gbGF5b3V0Q2FyZCh2YWx1ZTogc3RyaW5nLCBsYWJlbDogc3RyaW5nKTogSFRNTEJ1dHRvbkVsZW1lbnQge1xuICBjb25zdCBidXR0b24gPSBlbChcImJ1dHRvblwiLCB7IGNsYXNzOiBcImxheW91dFwiLCB0aXRsZTogbGFiZWwgfSk7XG4gIGJ1dHRvbi5pbm5lckhUTUwgPSBgPHN2ZyB2aWV3Qm94PVwiMCAwIDY0IDQ0XCIgYXJpYS1oaWRkZW49XCJ0cnVlXCI+JHtMQVlPVVRfU0tFVENIRVNbdmFsdWVdID8/IFwiXCJ9PC9zdmc+YDtcbiAgYnV0dG9uLmFwcGVuZChTSE9SVF9MQVlPVVRfTkFNRVNbdmFsdWVdID8/IGxhYmVsKTtcbiAgcmV0dXJuIGJ1dHRvbjtcbn1cblxuZnVuY3Rpb24gc2VnbWVudCh2YWx1ZTogc3RyaW5nLCBsYWJlbDogc3RyaW5nKTogSFRNTEJ1dHRvbkVsZW1lbnQge1xuICByZXR1cm4gZWwoXCJidXR0b25cIiwgeyBcImRhdGEtdmFsdWVcIjogdmFsdWUgfSwgbGFiZWwpO1xufVxuXG5mdW5jdGlvbiB0b2dnbGUoaWQ6IEZlYXR1cmVJZCwgc2V0dGluZ3M6IFNldHRpbmdzKTogSFRNTExhYmVsRWxlbWVudCB7XG4gIGNvbnN0IGlucHV0ID0gZWwoXCJpbnB1dFwiLCB7IHR5cGU6IFwiY2hlY2tib3hcIiwgaWQ6IGB0b2dnbGUtJHtpZH1gIH0pO1xuICBpbnB1dC5jaGVja2VkID0gc2V0dGluZ3NbaWRdO1xuICBpbnB1dC5hZGRFdmVudExpc3RlbmVyKFwiY2hhbmdlXCIsICgpID0+IHZvaWQgc2F2ZVNldHRpbmcoaWQsIGlucHV0LmNoZWNrZWQpKTtcbiAgcmV0dXJuIGVsKFxuICAgIFwibGFiZWxcIixcbiAgICB7IGNsYXNzOiBcInRvZ2dsZVwiLCBmb3I6IGB0b2dnbGUtJHtpZH1gIH0sXG4gICAgZWwoXCJzcGFuXCIsIHt9LCBGRUFUVVJFU1tpZF0ubGFiZWwpLFxuICAgIGlucHV0LFxuICApO1xufVxuXG5mdW5jdGlvbiB0b2dnbGVzKGlkczogRmVhdHVyZUlkW10sIHNldHRpbmdzOiBTZXR0aW5ncyk6IEhUTUxFbGVtZW50IHtcbiAgcmV0dXJuIGVsKFwiZGl2XCIsIHsgY2xhc3M6IFwidG9nZ2xlc1wiIH0sIC4uLmlkcy5tYXAoKGlkKSA9PiB0b2dnbGUoaWQsIHNldHRpbmdzKSkpO1xufVxuXG5jb25zdCBmZWF0dXJlc0luID0gKGdyb3VwOiBzdHJpbmcpID0+IEZFQVRVUkVfSURTLmZpbHRlcigoaWQpID0+IEZFQVRVUkVTW2lkXS5ncm91cCA9PT0gZ3JvdXApO1xuXG5sZXQgZGVjbHV0dGVyT3BlbiA9IGZhbHNlO1xubGV0IGRldmVsb3Blck9wZW4gPSBmYWxzZTtcbmxldCBjYXB0dXJlU3RhdHVzID0gXCJcIjtcblxuLyoqIEFza3MgdGhlIE91dGxvb2sgdGFiIGluIGZyb250IHRvIHNhdmUgYSBjYXB0dXJlIG9mIGl0c2VsZi4gTm8gcGVybWlzc2lvbiBuZWVkZWQ6IHRoZSBjb250ZW50XG4gKiBzY3JpcHQgdGhhdCBhbnN3ZXJzIGlzIGFscmVhZHkgb24gT3V0bG9vaydzIGhvc3RzLiAqL1xuYXN5bmMgZnVuY3Rpb24gcmVxdWVzdENhcHR1cmUoKTogUHJvbWlzZTxzdHJpbmc+IHtcbiAgY29uc3QgW3RhYl0gPSBhd2FpdCBjaHJvbWUudGFicy5xdWVyeSh7IGFjdGl2ZTogdHJ1ZSwgY3VycmVudFdpbmRvdzogdHJ1ZSB9KTtcbiAgaWYgKHRhYj8uaWQgPT09IHVuZGVmaW5lZCkgcmV0dXJuIFwiT3BlbiBhbiBPdXRsb29rIHRhYiBmaXJzdC5cIjtcbiAgdHJ5IHtcbiAgICBjb25zdCByZXBseSA9IChhd2FpdCBjaHJvbWUudGFicy5zZW5kTWVzc2FnZSh0YWIuaWQsIHsgdHlwZTogQ0FQVFVSRV9NRVNTQUdFIH0pKSBhcyB7XG4gICAgICBuYW1lOiBzdHJpbmc7XG4gICAgICBieXRlczogbnVtYmVyO1xuICAgIH07XG4gICAgcmV0dXJuIGBTYXZlZCAke3JlcGx5Lm5hbWV9ICgke01hdGgucm91bmQocmVwbHkuYnl0ZXMgLyAxMDI0KX0gS0IpIHRvIERvd25sb2Fkcy5gO1xuICB9IGNhdGNoIHtcbiAgICByZXR1cm4gXCJObyBhbnN3ZXIuIFVzZSB0aGlzIG9uIGFuIE91dGxvb2sgdGFiLCBhbmQgcmVsb2FkIHRoYXQgdGFiIGlmIE9CQiB3YXMganVzdCB1cGRhdGVkLlwiO1xuICB9XG59XG5cbmZ1bmN0aW9uIGRldmVsb3BlclNlY3Rpb24oKTogSFRNTEVsZW1lbnQge1xuICBjb25zdCBzdGF0dXMgPSBlbChcInBcIiwgeyBjbGFzczogXCJoaW50XCIsIHJvbGU6IFwic3RhdHVzXCIgfSwgY2FwdHVyZVN0YXR1cyk7XG4gIGNvbnN0IGJ1dHRvbiA9IGVsKFwiYnV0dG9uXCIsIHsgdHlwZTogXCJidXR0b25cIiwgY2xhc3M6IFwicGlsbFwiIH0sIFwiU2F2ZSBwYWdlIGNhcHR1cmVcIik7XG4gIGJ1dHRvbi5hZGRFdmVudExpc3RlbmVyKFwiY2xpY2tcIiwgKCkgPT4ge1xuICAgIHZvaWQgcmVxdWVzdENhcHR1cmUoKS50aGVuKCh0ZXh0KSA9PiB7XG4gICAgICBjYXB0dXJlU3RhdHVzID0gdGV4dDtcbiAgICAgIHN0YXR1cy50ZXh0Q29udGVudCA9IHRleHQ7XG4gICAgfSk7XG4gIH0pO1xuICBjb25zdCBkZXRhaWxzID0gZWwoXG4gICAgXCJkZXRhaWxzXCIsXG4gICAge30sXG4gICAgZWwoXCJzdW1tYXJ5XCIsIHt9LCBcIkRldmVsb3BlclwiKSxcbiAgICBlbChcbiAgICAgIFwicFwiLFxuICAgICAgeyBjbGFzczogXCJoaW50XCIgfSxcbiAgICAgIFwiU2F2ZXMgdGhpcyBPdXRsb29rIHRhYiBhcyBhbiBIVE1MIGZpbGUgZm9yIGZpeGluZyBvciB0ZXN0aW5nIE9CQi4gSXQgY29udGFpbnMgeW91ciByZWFsIG1haWwgYW5kIHN0YXlzIG9uIHRoaXMgY29tcHV0ZXI7IHNoYXJlIGl0IG9ubHkgd2l0aCB3aG9ldmVyIGlzIGhlbHBpbmcgeW91IGZpeCBPQkIuXCIsXG4gICAgKSxcbiAgICBidXR0b24sXG4gICAgc3RhdHVzLFxuICApO1xuICBkZXRhaWxzLm9wZW4gPSBkZXZlbG9wZXJPcGVuO1xuICBkZXRhaWxzLmFkZEV2ZW50TGlzdGVuZXIoXCJ0b2dnbGVcIiwgKCkgPT4ge1xuICAgIGRldmVsb3Blck9wZW4gPSBkZXRhaWxzLm9wZW47XG4gIH0pO1xuICByZXR1cm4gZWwoXCJzZWN0aW9uXCIsIHt9LCBkZXRhaWxzKTtcbn1cblxuZnVuY3Rpb24gcmVuZGVyKHNldHRpbmdzOiBTZXR0aW5ncyk6IHZvaWQge1xuICBhcHBseVRoZW1lKHNldHRpbmdzKTtcbiAgY29uc3QgbWFpbiA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKFwic2V0dGluZ3NcIik7XG4gIGlmICghbWFpbikgcmV0dXJuO1xuXG4gIGNvbnN0IGxheW91dEhpbnQgPVxuICAgIHNldHRpbmdzLmxheW91dCA9PT0gXCJmYXN0bWFpbFwiIHx8IHNldHRpbmdzLmxheW91dCA9PT0gXCJmb2N1c1wiXG4gICAgICA/IFtcbiAgICAgICAgICBlbChcbiAgICAgICAgICAgIFwicFwiLFxuICAgICAgICAgICAgeyBjbGFzczogXCJoaW50XCIgfSxcbiAgICAgICAgICAgIFwiUGFpciB3aXRoIE91dGxvb2sncyBTZXR0aW5ncyBcdTIxOTIgTWFpbCBcdTIxOTIgTGF5b3V0IFx1MjE5MiBSZWFkaW5nIHBhbmU6IEhpZGUuXCIsXG4gICAgICAgICAgKSxcbiAgICAgICAgXVxuICAgICAgOiBbXTtcblxuICBjb25zdCBkZWNsdXR0ZXIgPSBmZWF0dXJlc0luKFwiRGVjbHV0dGVyXCIpO1xuICBjb25zdCBoaWRkZW4gPSBkZWNsdXR0ZXIuZmlsdGVyKChpZCkgPT4gc2V0dGluZ3NbaWRdKS5sZW5ndGg7XG4gIGNvbnN0IGRldGFpbHMgPSBlbChcbiAgICBcImRldGFpbHNcIixcbiAgICB7fSxcbiAgICBlbChcbiAgICAgIFwic3VtbWFyeVwiLFxuICAgICAge30sXG4gICAgICBcIkRlY2x1dHRlclwiLFxuICAgICAgZWwoXCJzcGFuXCIsIHsgY2xhc3M6IFwiY291bnRcIiB9LCBgJHtoaWRkZW59IG9mICR7ZGVjbHV0dGVyLmxlbmd0aH0gaGlkZGVuYCksXG4gICAgKSxcbiAgICB0b2dnbGVzKGRlY2x1dHRlciwgc2V0dGluZ3MpLFxuICApO1xuICBkZXRhaWxzLm9wZW4gPSBkZWNsdXR0ZXJPcGVuO1xuICBkZXRhaWxzLmFkZEV2ZW50TGlzdGVuZXIoXCJ0b2dnbGVcIiwgKCkgPT4ge1xuICAgIGRlY2x1dHRlck9wZW4gPSBkZXRhaWxzLm9wZW47XG4gIH0pO1xuXG4gIG1haW4ucmVwbGFjZUNoaWxkcmVuKFxuICAgIHNlY3Rpb24oXCJDb2xvcnNcIiwgY2hvaWNlR3JvdXAoXCJwYWxldHRlXCIsIHNldHRpbmdzLCBcInN3YXRjaGVzXCIsIHN3YXRjaChzZXR0aW5ncykpKSxcbiAgICBzZWN0aW9uKFwiTGlnaHQgb3IgZGFya1wiLCBjaG9pY2VHcm91cChcIm1vZGVcIiwgc2V0dGluZ3MsIFwic2VnbWVudGVkXCIsIHNlZ21lbnQpKSxcbiAgICBzZWN0aW9uKFwiTGF5b3V0XCIsIGNob2ljZUdyb3VwKFwibGF5b3V0XCIsIHNldHRpbmdzLCBcImxheW91dHNcIiwgbGF5b3V0Q2FyZCksIC4uLmxheW91dEhpbnQpLFxuICAgIHNlY3Rpb24oXCJLZXlib2FyZFwiLCB0b2dnbGVzKGZlYXR1cmVzSW4oXCJLZXlib2FyZFwiKSwgc2V0dGluZ3MpKSxcbiAgICBlbChcInNlY3Rpb25cIiwge30sIGRldGFpbHMpLFxuICAgIHNlY3Rpb24oXCJHZW5lcmFsXCIsIHRvZ2dsZXMoZmVhdHVyZXNJbihcIkdlbmVyYWxcIiksIHNldHRpbmdzKSksXG4gICAgZGV2ZWxvcGVyU2VjdGlvbigpLFxuICApO1xufVxuXG5hc3luYyBmdW5jdGlvbiBtYWluKCk6IFByb21pc2U8dm9pZD4ge1xuICBjb25zdCB2ZXJzaW9uID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoXCJ2ZXJzaW9uXCIpO1xuICBpZiAodmVyc2lvbikgdmVyc2lvbi50ZXh0Q29udGVudCA9IGB2JHtjaHJvbWUucnVudGltZS5nZXRNYW5pZmVzdCgpLnZlcnNpb259YDtcbiAgLy8gUGlja3MgdXAgbmV3IGZpbGVzIG9uIGRpc2sgKGFmdGVyIGBnaXQgcHVsbGAgaW4gdGhlIGluc3RhbGwgZm9sZGVyKTsgc2VlIFJFQURNRS5cbiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoXCJyZWxvYWRcIik/LmFkZEV2ZW50TGlzdGVuZXIoXCJjbGlja1wiLCAoKSA9PiBjaHJvbWUucnVudGltZS5yZWxvYWQoKSk7XG4gIGxldCBzZXR0aW5ncyA9IGF3YWl0IGxvYWRTZXR0aW5ncygpO1xuICByZW5kZXIoc2V0dGluZ3MpO1xuICBvblNldHRpbmdzQ2hhbmdlZCgobmV4dCkgPT4ge1xuICAgIHNldHRpbmdzID0gbmV4dDtcbiAgICByZW5kZXIoc2V0dGluZ3MpO1xuICB9KTtcbiAgcHJlZmVyc0RhcmsuYWRkRXZlbnRMaXN0ZW5lcihcImNoYW5nZVwiLCAoKSA9PiByZW5kZXIoc2V0dGluZ3MpKTtcbn1cblxudm9pZCBtYWluKCk7XG4iXSwKICAibWFwcGluZ3MiOiAiOzs7QUF5Qk8sTUFBTSxZQUFZO0FBQUE7QUFBQSxJQUV2QixVQUFVLEVBQUUsWUFBWSxDQUFDLGNBQWMsR0FBRyxRQUFRLFVBQVUsS0FBSyxZQUFZO0FBQUEsSUFDN0UsV0FBVyxFQUFFLFlBQVksQ0FBQyxtQkFBbUIsdUJBQXVCLEdBQUcsUUFBUSxTQUFTO0FBQUEsSUFDeEYsU0FBUztBQUFBLE1BQ1AsWUFBWSxDQUFDLGFBQWEsb0RBQW9EO0FBQUEsTUFDOUUsUUFBUTtBQUFBLE1BQ1IsS0FBSztBQUFBLElBQ1A7QUFBQSxJQUNBLFFBQVE7QUFBQSxNQUNOLFlBQVksQ0FBQywrQkFBK0Isc0NBQXNDO0FBQUEsTUFDbEYsUUFBUTtBQUFBLE1BQ1IsS0FBSztBQUFBLElBQ1A7QUFBQSxJQUNBLFlBQVk7QUFBQSxNQUNWLFlBQVksQ0FBQyw0Q0FBNEM7QUFBQSxNQUN6RCxRQUFRO0FBQUEsTUFDUixRQUFRO0FBQUEsTUFDUixLQUFLO0FBQUEsSUFDUDtBQUFBLElBQ0EsWUFBWTtBQUFBLE1BQ1YsWUFBWTtBQUFBLFFBQ1Y7QUFBQSxRQUNBO0FBQUEsTUFDRjtBQUFBLE1BQ0EsUUFBUTtBQUFBLE1BQ1IsS0FBSztBQUFBLElBQ1A7QUFBQSxJQUNBLFlBQVk7QUFBQSxNQUNWLFlBQVksQ0FBQyx1Q0FBdUMsbUJBQW1CO0FBQUEsTUFDdkUsUUFBUTtBQUFBLE1BQ1IsUUFBUTtBQUFBLE1BQ1IsS0FBSztBQUFBLE1BQ0wsS0FBSztBQUFBLElBQ1A7QUFBQSxJQUNBLFVBQVU7QUFBQSxNQUNSLFlBQVksQ0FBQyxrQ0FBa0M7QUFBQSxNQUMvQyxRQUFRO0FBQUEsTUFDUixLQUFLO0FBQUEsSUFDUDtBQUFBLElBQ0EsYUFBYTtBQUFBLE1BQ1gsWUFBWSxDQUFDLGFBQWEsNkNBQTZDO0FBQUEsTUFDdkUsUUFBUTtBQUFBLE1BQ1IsS0FBSztBQUFBLElBQ1A7QUFBQSxJQUNBLFlBQVk7QUFBQSxNQUNWLFlBQVksQ0FBQyxnQ0FBZ0MsaUJBQWlCO0FBQUEsTUFDOUQsUUFBUTtBQUFBLE1BQ1IsUUFBUTtBQUFBLE1BQ1IsS0FBSztBQUFBLE1BQ0wsS0FBSztBQUFBLElBQ1A7QUFBQSxJQUNBLGFBQWE7QUFBQSxNQUNYLFlBQVksQ0FBQyxvQ0FBb0M7QUFBQSxNQUNqRCxRQUFRO0FBQUEsTUFDUixRQUFRO0FBQUEsTUFDUixLQUFLO0FBQUEsTUFDTCxLQUFLO0FBQUEsSUFDUDtBQUFBLElBQ0EsYUFBYTtBQUFBLE1BQ1gsWUFBWSxDQUFDLGtCQUFrQjtBQUFBLE1BQy9CLFFBQVE7QUFBQSxNQUNSLFFBQVE7QUFBQSxNQUNSLEtBQUs7QUFBQSxJQUNQO0FBQUEsSUFDQSxhQUFhO0FBQUEsTUFDWCxZQUFZO0FBQUEsUUFDVjtBQUFBLFFBQ0E7QUFBQSxRQUNBO0FBQUEsTUFDRjtBQUFBLE1BQ0EsUUFBUTtBQUFBLE1BQ1IsS0FBSztBQUFBLElBQ1A7QUFBQSxJQUVBLGFBQWE7QUFBQSxNQUNYLFlBQVksQ0FBQyw4Q0FBOEM7QUFBQSxNQUMzRCxRQUFRO0FBQUEsTUFDUixRQUFRO0FBQUEsTUFDUixLQUFLO0FBQUEsTUFDTCxLQUFLO0FBQUEsSUFDUDtBQUFBLElBRUEsU0FBUztBQUFBO0FBQUE7QUFBQSxNQUdQLFlBQVksQ0FBQywyQ0FBMkM7QUFBQSxNQUN4RCxRQUFRO0FBQUEsTUFDUixRQUFRO0FBQUEsTUFDUixLQUFLO0FBQUEsTUFDTCxLQUFLO0FBQUEsSUFDUDtBQUFBO0FBQUEsSUFHQSxjQUFjO0FBQUEsTUFDWixZQUFZLENBQUMsMENBQTBDO0FBQUEsTUFDdkQsUUFBUTtBQUFBLE1BQ1IsS0FBSztBQUFBLE1BQ0wsS0FBSztBQUFBLElBQ1A7QUFBQSxJQUNBLGNBQWM7QUFBQSxNQUNaLFlBQVksQ0FBQywwQkFBMEIsNEJBQTRCO0FBQUEsTUFDbkUsUUFBUTtBQUFBLE1BQ1IsS0FBSztBQUFBLE1BQ0wsT0FBTztBQUFBLElBQ1Q7QUFBQSxJQUNBLGVBQWU7QUFBQSxNQUNiLFlBQVksQ0FBQyxTQUFTLGVBQWUsZ0JBQWdCLFVBQVUsRUFBRTtBQUFBLFFBQy9ELENBQUMsUUFBUSw2QkFBNkIsR0FBRztBQUFBLE1BQzNDO0FBQUEsTUFDQSxRQUFRO0FBQUEsTUFDUixRQUFRO0FBQUEsTUFDUixLQUFLO0FBQUEsTUFDTCxPQUFPO0FBQUEsSUFDVDtBQUFBLElBQ0EsY0FBYztBQUFBLE1BQ1osWUFBWTtBQUFBLFFBQ1Y7QUFBQSxRQUNBO0FBQUEsUUFDQTtBQUFBLFFBQ0E7QUFBQSxRQUNBO0FBQUEsUUFDQTtBQUFBLFFBQ0E7QUFBQSxNQUNGLEVBQUUsSUFBSSxDQUFDLFVBQVUsZ0JBQWdCLEtBQUssSUFBSTtBQUFBLE1BQzFDLFFBQVE7QUFBQSxNQUNSLFFBQVE7QUFBQSxNQUNSLEtBQUs7QUFBQSxNQUNMLE9BQU87QUFBQSxJQUNUO0FBQUEsSUFDQSxXQUFXO0FBQUEsTUFDVCxZQUFZLENBQUMsNkJBQTZCO0FBQUEsTUFDMUMsUUFBUTtBQUFBLE1BQ1IsUUFBUTtBQUFBLE1BQ1IsS0FBSztBQUFBLE1BQ0wsS0FBSztBQUFBLElBQ1A7QUFBQSxJQUNBLGlCQUFpQjtBQUFBLE1BQ2YsWUFBWTtBQUFBLFFBQ1Y7QUFBQSxRQUNBO0FBQUEsUUFDQTtBQUFBLE1BQ0YsRUFBRSxJQUFJLENBQUMsVUFBVSxzQkFBc0IsS0FBSyxJQUFJO0FBQUEsTUFDaEQsUUFBUTtBQUFBLE1BQ1IsUUFBUTtBQUFBLE1BQ1IsS0FBSztBQUFBLE1BQ0wsT0FBTztBQUFBLElBQ1Q7QUFBQSxJQUNBLFdBQVc7QUFBQTtBQUFBO0FBQUEsTUFHVCxZQUFZLENBQUMsZ0RBQWdEO0FBQUEsTUFDN0QsUUFBUTtBQUFBLE1BQ1IsS0FBSztBQUFBLE1BQ0wsS0FBSztBQUFBLElBQ1A7QUFBQSxJQUNBLFNBQVM7QUFBQTtBQUFBLE1BRVAsWUFBWSxDQUFDLG1DQUFtQywwQ0FBMEM7QUFBQSxNQUMxRixRQUFRO0FBQUEsTUFDUixLQUFLO0FBQUEsTUFDTCxPQUFPO0FBQUEsSUFDVDtBQUFBLElBQ0EsYUFBYTtBQUFBO0FBQUEsTUFFWCxZQUFZLENBQUMsNkJBQTZCLHlCQUF5QjtBQUFBLE1BQ25FLFFBQVE7QUFBQSxNQUNSLEtBQUs7QUFBQSxNQUNMLE9BQU87QUFBQSxJQUNUO0FBQUE7QUFBQSxJQUdBLGVBQWUsRUFBRSxZQUFZLENBQUMsK0JBQStCLEdBQUcsUUFBUSxTQUFTO0FBQUEsSUFDakYsZUFBZSxFQUFFLFlBQVksQ0FBQyw4QkFBOEIsR0FBRyxRQUFRLFlBQVk7QUFBQSxJQUNuRixjQUFjLEVBQUUsWUFBWSxDQUFDLDZCQUE2QixHQUFHLFFBQVEsWUFBWTtBQUFBO0FBQUEsSUFFakYsWUFBWTtBQUFBLE1BQ1YsWUFBWSxDQUFDLCtCQUErQiwyQkFBMkI7QUFBQSxNQUN2RSxRQUFRO0FBQUEsSUFDVjtBQUFBLElBQ0EsWUFBWSxFQUFFLFlBQVksQ0FBQyw4QkFBOEIsR0FBRyxRQUFRLFlBQVk7QUFBQSxJQUNoRixjQUFjLEVBQUUsWUFBWSxDQUFDLDZCQUE2QixHQUFHLFFBQVEsWUFBWTtBQUFBLElBQ2pGLFlBQVksRUFBRSxZQUFZLENBQUMsb0NBQW9DLEdBQUcsUUFBUSxZQUFZO0FBQUEsSUFDdEYsa0JBQWtCLEVBQUUsWUFBWSxDQUFDLG9DQUFvQyxHQUFHLFFBQVEsWUFBWTtBQUFBLElBQzVGLFlBQVksRUFBRSxZQUFZLENBQUMsMkJBQTJCLEdBQUcsUUFBUSxZQUFZO0FBQUEsSUFDN0UsWUFBWSxFQUFFLFlBQVksQ0FBQywwQkFBMEIsR0FBRyxRQUFRLFlBQVk7QUFBQSxJQUM1RSxnQkFBZ0IsRUFBRSxZQUFZLENBQUMsMkJBQTJCLEdBQUcsUUFBUSxTQUFTO0FBQUE7QUFBQTtBQUFBLElBRzlFLFNBQVM7QUFBQSxNQUNQLFlBQVksQ0FBQyw2QkFBNkIsbURBQW1EO0FBQUEsTUFDN0YsUUFBUTtBQUFBLE1BQ1IsUUFBUTtBQUFBLElBQ1Y7QUFBQSxJQUNBLGFBQWE7QUFBQSxNQUNYLFlBQVk7QUFBQSxRQUNWO0FBQUEsUUFDQTtBQUFBLE1BQ0Y7QUFBQSxNQUNBLFFBQVE7QUFBQSxNQUNSLFFBQVE7QUFBQSxJQUNWO0FBQUEsSUFDQSxXQUFXO0FBQUEsTUFDVCxZQUFZO0FBQUEsUUFDVjtBQUFBLFFBQ0E7QUFBQSxNQUNGO0FBQUEsTUFDQSxRQUFRO0FBQUEsTUFDUixRQUFRO0FBQUEsSUFDVjtBQUFBLEVBQ0Y7OztBQ25PTyxNQUFNLGtCQUFrQjs7O0FDa0J4QixNQUFNLFdBQVc7QUFBQSxJQUN0QixpQkFBaUI7QUFBQSxNQUNmLE9BQU87QUFBQSxNQUNQLE9BQU87QUFBQSxNQUNQLFNBQVM7QUFBQSxJQUNYO0FBQUEsSUFDQSxRQUFRLEVBQUUsT0FBTyxxQ0FBcUMsT0FBTyxZQUFZLFNBQVMsS0FBSztBQUFBLElBQ3ZGLGVBQWUsRUFBRSxPQUFPLDJDQUFzQyxPQUFPLGFBQWEsU0FBUyxLQUFLO0FBQUEsSUFDaEcsb0JBQW9CO0FBQUEsTUFDbEIsT0FBTztBQUFBLE1BQ1AsT0FBTztBQUFBLE1BQ1AsU0FBUztBQUFBLElBQ1g7QUFBQSxJQUNBLGtCQUFrQjtBQUFBLE1BQ2hCLE9BQU87QUFBQSxNQUNQLE9BQU87QUFBQSxNQUNQLFNBQVM7QUFBQSxJQUNYO0FBQUEsSUFDQSxnQkFBZ0IsRUFBRSxPQUFPLDhCQUE4QixPQUFPLGFBQWEsU0FBUyxLQUFLO0FBQUEsSUFDekYsc0JBQXNCO0FBQUEsTUFDcEIsT0FBTztBQUFBLE1BQ1AsT0FBTztBQUFBLE1BQ1AsU0FBUztBQUFBLElBQ1g7QUFBQSxJQUNBLG9CQUFvQjtBQUFBLE1BQ2xCLE9BQU87QUFBQSxNQUNQLE9BQU87QUFBQSxNQUNQLFNBQVM7QUFBQSxJQUNYO0FBQUEsSUFDQSxzQkFBc0I7QUFBQSxNQUNwQixPQUFPO0FBQUEsTUFDUCxPQUFPO0FBQUEsTUFDUCxTQUFTO0FBQUEsSUFDWDtBQUFBLElBQ0EscUJBQXFCLEVBQUUsT0FBTywwQkFBMEIsT0FBTyxhQUFhLFNBQVMsS0FBSztBQUFBLElBQzFGLG9CQUFvQjtBQUFBLE1BQ2xCLE9BQU87QUFBQSxNQUNQLE9BQU87QUFBQSxNQUNQLFNBQVM7QUFBQSxJQUNYO0FBQUEsSUFDQSxnQkFBZ0I7QUFBQSxNQUNkLE9BQU87QUFBQSxNQUNQLE9BQU87QUFBQSxNQUNQLFNBQVM7QUFBQSxJQUNYO0FBQUEsSUFDQSxzQkFBc0I7QUFBQSxNQUNwQixPQUFPO0FBQUEsTUFDUCxPQUFPO0FBQUEsTUFDUCxTQUFTO0FBQUEsSUFDWDtBQUFBLElBQ0EscUJBQXFCLEVBQUUsT0FBTyx3QkFBd0IsT0FBTyxhQUFhLFNBQVMsTUFBTTtBQUFBLElBQ3pGLG9CQUFvQjtBQUFBLE1BQ2xCLE9BQU87QUFBQSxNQUNQLE9BQU87QUFBQSxNQUNQLFNBQVM7QUFBQSxJQUNYO0FBQUEsSUFDQSxpQkFBaUIsRUFBRSxPQUFPLDRCQUE0QixPQUFPLGFBQWEsU0FBUyxNQUFNO0FBQUEsSUFDekYsZUFBZSxFQUFFLE9BQU8sVUFBVSxPQUFPLGFBQWEsU0FBUyxNQUFNO0FBQUEsSUFDckUsU0FBUyxFQUFFLE9BQU8sMEJBQTBCLE9BQU8sUUFBUSxTQUFTLEtBQUs7QUFBQSxFQUMzRTtBQUVPLE1BQU0sVUFBVTtBQUFBLElBQ3JCLFNBQVM7QUFBQSxNQUNQLE9BQU87QUFBQSxNQUNQLE9BQU87QUFBQSxNQUNQLFNBQVM7QUFBQSxRQUNQLFNBQVM7QUFBQSxRQUNULFdBQVc7QUFBQSxRQUNYLFFBQVE7QUFBQSxRQUNSLFVBQVU7QUFBQSxRQUNWLE9BQU87QUFBQSxNQUNUO0FBQUEsTUFDQSxTQUFTO0FBQUEsSUFDWDtBQUFBLElBQ0EsTUFBTTtBQUFBLE1BQ0osT0FBTztBQUFBLE1BQ1AsT0FBTztBQUFBLE1BQ1AsU0FBUyxFQUFFLE1BQU0sa0JBQWtCLE9BQU8sU0FBUyxNQUFNLE9BQU87QUFBQSxNQUNoRSxTQUFTO0FBQUEsSUFDWDtBQUFBLElBQ0EsUUFBUTtBQUFBLE1BQ04sT0FBTztBQUFBLE1BQ1AsT0FBTztBQUFBLE1BQ1AsU0FBUztBQUFBLFFBQ1AsU0FBUztBQUFBLFFBQ1QsTUFBTTtBQUFBLFFBQ04sVUFBVTtBQUFBLFFBQ1YsT0FBTztBQUFBLE1BQ1Q7QUFBQSxNQUNBLFNBQVM7QUFBQSxJQUNYO0FBQUEsRUFDRjtBQUVPLE1BQU0sY0FBYyxPQUFPLEtBQUssUUFBUTtBQUN4QyxNQUFNLGFBQWEsT0FBTyxLQUFLLE9BQU87QUFFdEMsV0FBUyxrQkFBNEI7QUFDMUMsV0FBTztBQUFBLE1BQ0wsR0FBRyxPQUFPLFlBQVksWUFBWSxJQUFJLENBQUMsT0FBTyxDQUFDLElBQUksU0FBUyxFQUFFLEVBQUUsT0FBTyxDQUFDLENBQUM7QUFBQSxNQUN6RSxHQUFHLE9BQU8sWUFBWSxXQUFXLElBQUksQ0FBQyxPQUFPLENBQUMsSUFBSSxRQUFRLEVBQUUsRUFBRSxPQUFPLENBQUMsQ0FBQztBQUFBLElBQ3pFO0FBQUEsRUFDRjtBQUdPLFdBQVMsVUFBVSxRQUEyQztBQUNuRSxVQUFNLFdBQVcsZ0JBQWdCO0FBQ2pDLGVBQVcsTUFBTSxhQUFhO0FBQzVCLFVBQUksT0FBTyxPQUFPLEVBQUUsTUFBTSxVQUFXLFVBQVMsRUFBRSxJQUFJLE9BQU8sRUFBRTtBQUFBLElBQy9EO0FBQ0EsZUFBVyxNQUFNLFlBQVk7QUFDM0IsWUFBTSxRQUFRLE9BQU8sRUFBRTtBQUN2QixVQUFJLE9BQU8sVUFBVSxZQUFZLFNBQVMsUUFBUSxFQUFFLEVBQUUsUUFBUyxVQUFTLEVBQUUsSUFBSTtBQUFBLElBQ2hGO0FBQ0EsV0FBTztBQUFBLEVBQ1Q7QUFZQSxpQkFBc0IsZUFBa0M7QUFDdEQsV0FBTyxVQUFVLE1BQU0sT0FBTyxRQUFRLEtBQUssSUFBSSxJQUFJLENBQUM7QUFBQSxFQUN0RDtBQUVBLGlCQUFzQixZQUNwQixJQUNBLE9BQ2U7QUFDZixVQUFNLE9BQU8sUUFBUSxLQUFLLElBQUksRUFBRSxDQUFDLEVBQUUsR0FBRyxNQUFNLENBQUM7QUFBQSxFQUMvQztBQUVPLFdBQVMsa0JBQWtCLFVBQThDO0FBQzlFLFdBQU8sUUFBUSxVQUFVLFlBQVksQ0FBQyxVQUFVLFNBQVM7QUFDdkQsVUFBSSxTQUFTLE9BQVEsTUFBSyxhQUFhLEVBQUUsS0FBSyxRQUFRO0FBQUEsSUFDeEQsQ0FBQztBQUFBLEVBQ0g7OztBQ3hKQSxNQUFNLHFCQUE2QztBQUFBLElBQ2pELFNBQVM7QUFBQSxJQUNULE1BQU07QUFBQSxJQUNOLFVBQVU7QUFBQSxJQUNWLE9BQU87QUFBQSxFQUNUO0FBRUEsTUFBTSxzQkFBOEM7QUFBQSxJQUNsRCxTQUFTO0FBQUEsSUFDVCxXQUFXO0FBQUEsSUFDWCxRQUFRO0FBQUEsSUFDUixVQUFVO0FBQUEsSUFDVixPQUFPO0FBQUEsRUFDVDtBQUdBLFdBQVMsS0FBSyxHQUFXLEdBQVcsT0FBdUI7QUFDekQsV0FBTyxDQUFDLEdBQUcsR0FBRyxHQUFHLEdBQUcsQ0FBQyxFQUNsQixJQUFJLENBQUMsTUFBTTtBQUNWLFlBQU0sTUFBTSxJQUFJLElBQUk7QUFDcEIsWUFBTSxVQUFVLE1BQU0sSUFBSSxPQUFPO0FBQ2pDLFlBQU0sZUFBZSxNQUFNLElBQUksUUFBUSxPQUFPO0FBQzlDLGFBQ0UsZ0JBQWdCLE9BQU8sUUFBUSxDQUFDLFFBQVEsR0FBRyxZQUFZLFlBQVksK0NBQzNDLENBQUMsUUFBUSxNQUFNLENBQUMsWUFBWSxRQUFRLElBQUk7QUFBQSxJQUVwRSxDQUFDLEVBQ0EsS0FBSyxFQUFFO0FBQUEsRUFDWjtBQUdBLE1BQU0sa0JBQTBDO0FBQUEsSUFDOUMsU0FBUztBQUFBO0FBQUE7QUFBQSxNQUdMLEtBQUssSUFBSSxHQUFHLEVBQUUsQ0FBQztBQUFBLElBQ25CLE1BQU07QUFBQTtBQUFBLE1BRUYsS0FBSyxJQUFJLEdBQUcsRUFBRSxDQUFDO0FBQUEsSUFDbkIsVUFBVTtBQUFBO0FBQUEsbUVBRXVELEtBQUssSUFBSSxHQUFHLEVBQUUsQ0FBQztBQUFBLElBQ2hGLE9BQU8sK0NBQStDLEtBQUssSUFBSSxHQUFHLEVBQUUsQ0FBQztBQUFBLEVBQ3ZFO0FBRUEsTUFBTSxPQUFPLFNBQVM7QUFDdEIsTUFBTSxjQUFjLFdBQVcsOEJBQThCO0FBRzdELFdBQVMsY0FBYyxVQUFzQztBQUMzRCxRQUFJLFNBQVMsU0FBUyxXQUFXLFNBQVMsU0FBUyxPQUFRLFFBQU8sU0FBUztBQUMzRSxXQUFPLFlBQVksVUFBVSxTQUFTO0FBQUEsRUFDeEM7QUFHQSxXQUFTLFdBQVcsVUFBMEI7QUFDNUMsU0FBSyxRQUFRLGtCQUFrQixTQUFTO0FBQ3hDLFNBQUssUUFBUSxlQUFlLFNBQVM7QUFDckMsU0FBSyxRQUFRLHdCQUF3QixZQUFZLFVBQVUsU0FBUztBQUFBLEVBQ3RFO0FBRUEsV0FBUyxHQUNQLEtBQ0EsUUFBZ0MsQ0FBQyxNQUM5QixVQUN1QjtBQUMxQixVQUFNLE9BQU8sU0FBUyxjQUFjLEdBQUc7QUFDdkMsZUFBVyxDQUFDLEdBQUcsQ0FBQyxLQUFLLE9BQU8sUUFBUSxLQUFLLEVBQUcsTUFBSyxhQUFhLEdBQUcsQ0FBQztBQUNsRSxTQUFLLE9BQU8sR0FBRyxRQUFRO0FBQ3ZCLFdBQU87QUFBQSxFQUNUO0FBRUEsV0FBUyxRQUFRLFVBQWtCLFVBQStCO0FBQ2hFLFdBQU8sR0FBRyxXQUFXLENBQUMsR0FBRyxHQUFHLE1BQU0sQ0FBQyxHQUFHLEtBQUssR0FBRyxHQUFHLFFBQVE7QUFBQSxFQUMzRDtBQUdBLFdBQVMsWUFDUCxJQUNBLFVBQ0EsV0FDQUEsU0FDYTtBQUNiLFVBQU0sUUFBUSxHQUFHLE9BQU87QUFBQSxNQUN0QixPQUFPO0FBQUEsTUFDUCxNQUFNO0FBQUEsTUFDTixjQUFjLFFBQVEsRUFBRSxFQUFFO0FBQUEsSUFDNUIsQ0FBQztBQUNELGVBQVcsQ0FBQyxPQUFPLEtBQUssS0FBSyxPQUFPLFFBQVEsUUFBUSxFQUFFLEVBQUUsT0FBTyxHQUFHO0FBQ2hFLFlBQU0sU0FBU0EsUUFBTyxPQUFPLEtBQUs7QUFDbEMsYUFBTyxPQUFPO0FBQ2QsYUFBTyxhQUFhLFFBQVEsT0FBTztBQUNuQyxhQUFPLGFBQWEsZ0JBQWdCLE9BQU8sU0FBUyxFQUFFLE1BQU0sS0FBSyxDQUFDO0FBQ2xFLGFBQU8saUJBQWlCLFNBQVMsTUFBTSxLQUFLLFlBQVksSUFBSSxLQUFLLENBQUM7QUFDbEUsWUFBTSxPQUFPLE1BQU07QUFBQSxJQUNyQjtBQUNBLFdBQU87QUFBQSxFQUNUO0FBR0EsV0FBUyxPQUFPLFVBQW9CO0FBQ2xDLFdBQU8sQ0FBQyxPQUFlLFVBQXFDO0FBQzFELFlBQU0sVUFBVTtBQUFBLFFBQ2Q7QUFBQSxRQUNBO0FBQUEsVUFDRSxPQUFPO0FBQUEsVUFDUCx3QkFBd0I7QUFBQSxVQUN4Qix5QkFBeUI7QUFBQSxVQUN6QixzQkFBc0IsY0FBYyxRQUFRO0FBQUEsUUFDOUM7QUFBQSxRQUNBLEdBQUcsUUFBUSxFQUFFLE9BQU8sTUFBTSxDQUFDO0FBQUEsUUFDM0IsR0FBRyxRQUFRLEVBQUUsT0FBTyxPQUFPLENBQUM7QUFBQSxRQUM1QjtBQUFBLFVBQ0U7QUFBQSxVQUNBLEVBQUUsT0FBTyxPQUFPO0FBQUEsVUFDaEIsR0FBRyxRQUFRLEVBQUUsT0FBTyxhQUFhLENBQUM7QUFBQSxVQUNsQyxHQUFHLFFBQVEsRUFBRSxPQUFPLE9BQU8sQ0FBQztBQUFBLFVBQzVCLEdBQUcsUUFBUSxFQUFFLE9BQU8sYUFBYSxDQUFDO0FBQUEsVUFDbEMsR0FBRyxRQUFRLEVBQUUsT0FBTyxPQUFPLENBQUM7QUFBQSxRQUM5QjtBQUFBLE1BQ0Y7QUFDQSxhQUFPO0FBQUEsUUFDTDtBQUFBLFFBQ0EsRUFBRSxPQUFPLFVBQVUsT0FBTyxNQUFNO0FBQUEsUUFDaEM7QUFBQSxRQUNBLG9CQUFvQixLQUFLLEtBQUs7QUFBQSxNQUNoQztBQUFBLElBQ0Y7QUFBQSxFQUNGO0FBRUEsV0FBUyxXQUFXLE9BQWUsT0FBa0M7QUFDbkUsVUFBTSxTQUFTLEdBQUcsVUFBVSxFQUFFLE9BQU8sVUFBVSxPQUFPLE1BQU0sQ0FBQztBQUM3RCxXQUFPLFlBQVksK0NBQStDLGdCQUFnQixLQUFLLEtBQUssRUFBRTtBQUM5RixXQUFPLE9BQU8sbUJBQW1CLEtBQUssS0FBSyxLQUFLO0FBQ2hELFdBQU87QUFBQSxFQUNUO0FBRUEsV0FBUyxRQUFRLE9BQWUsT0FBa0M7QUFDaEUsV0FBTyxHQUFHLFVBQVUsRUFBRSxjQUFjLE1BQU0sR0FBRyxLQUFLO0FBQUEsRUFDcEQ7QUFFQSxXQUFTLE9BQU8sSUFBZSxVQUFzQztBQUNuRSxVQUFNLFFBQVEsR0FBRyxTQUFTLEVBQUUsTUFBTSxZQUFZLElBQUksVUFBVSxFQUFFLEdBQUcsQ0FBQztBQUNsRSxVQUFNLFVBQVUsU0FBUyxFQUFFO0FBQzNCLFVBQU0saUJBQWlCLFVBQVUsTUFBTSxLQUFLLFlBQVksSUFBSSxNQUFNLE9BQU8sQ0FBQztBQUMxRSxXQUFPO0FBQUEsTUFDTDtBQUFBLE1BQ0EsRUFBRSxPQUFPLFVBQVUsS0FBSyxVQUFVLEVBQUUsR0FBRztBQUFBLE1BQ3ZDLEdBQUcsUUFBUSxDQUFDLEdBQUcsU0FBUyxFQUFFLEVBQUUsS0FBSztBQUFBLE1BQ2pDO0FBQUEsSUFDRjtBQUFBLEVBQ0Y7QUFFQSxXQUFTLFFBQVEsS0FBa0IsVUFBaUM7QUFDbEUsV0FBTyxHQUFHLE9BQU8sRUFBRSxPQUFPLFVBQVUsR0FBRyxHQUFHLElBQUksSUFBSSxDQUFDLE9BQU8sT0FBTyxJQUFJLFFBQVEsQ0FBQyxDQUFDO0FBQUEsRUFDakY7QUFFQSxNQUFNLGFBQWEsQ0FBQyxVQUFrQixZQUFZLE9BQU8sQ0FBQyxPQUFPLFNBQVMsRUFBRSxFQUFFLFVBQVUsS0FBSztBQUU3RixNQUFJLGdCQUFnQjtBQUNwQixNQUFJLGdCQUFnQjtBQUNwQixNQUFJLGdCQUFnQjtBQUlwQixpQkFBZSxpQkFBa0M7QUFDL0MsVUFBTSxDQUFDLEdBQUcsSUFBSSxNQUFNLE9BQU8sS0FBSyxNQUFNLEVBQUUsUUFBUSxNQUFNLGVBQWUsS0FBSyxDQUFDO0FBQzNFLFFBQUksS0FBSyxPQUFPLE9BQVcsUUFBTztBQUNsQyxRQUFJO0FBQ0YsWUFBTSxRQUFTLE1BQU0sT0FBTyxLQUFLLFlBQVksSUFBSSxJQUFJLEVBQUUsTUFBTSxnQkFBZ0IsQ0FBQztBQUk5RSxhQUFPLFNBQVMsTUFBTSxJQUFJLEtBQUssS0FBSyxNQUFNLE1BQU0sUUFBUSxJQUFJLENBQUM7QUFBQSxJQUMvRCxRQUFRO0FBQ04sYUFBTztBQUFBLElBQ1Q7QUFBQSxFQUNGO0FBRUEsV0FBUyxtQkFBZ0M7QUFDdkMsVUFBTSxTQUFTLEdBQUcsS0FBSyxFQUFFLE9BQU8sUUFBUSxNQUFNLFNBQVMsR0FBRyxhQUFhO0FBQ3ZFLFVBQU0sU0FBUyxHQUFHLFVBQVUsRUFBRSxNQUFNLFVBQVUsT0FBTyxPQUFPLEdBQUcsbUJBQW1CO0FBQ2xGLFdBQU8saUJBQWlCLFNBQVMsTUFBTTtBQUNyQyxXQUFLLGVBQWUsRUFBRSxLQUFLLENBQUMsU0FBUztBQUNuQyx3QkFBZ0I7QUFDaEIsZUFBTyxjQUFjO0FBQUEsTUFDdkIsQ0FBQztBQUFBLElBQ0gsQ0FBQztBQUNELFVBQU0sVUFBVTtBQUFBLE1BQ2Q7QUFBQSxNQUNBLENBQUM7QUFBQSxNQUNELEdBQUcsV0FBVyxDQUFDLEdBQUcsV0FBVztBQUFBLE1BQzdCO0FBQUEsUUFDRTtBQUFBLFFBQ0EsRUFBRSxPQUFPLE9BQU87QUFBQSxRQUNoQjtBQUFBLE1BQ0Y7QUFBQSxNQUNBO0FBQUEsTUFDQTtBQUFBLElBQ0Y7QUFDQSxZQUFRLE9BQU87QUFDZixZQUFRLGlCQUFpQixVQUFVLE1BQU07QUFDdkMsc0JBQWdCLFFBQVE7QUFBQSxJQUMxQixDQUFDO0FBQ0QsV0FBTyxHQUFHLFdBQVcsQ0FBQyxHQUFHLE9BQU87QUFBQSxFQUNsQztBQUVBLFdBQVMsT0FBTyxVQUEwQjtBQUN4QyxlQUFXLFFBQVE7QUFDbkIsVUFBTUMsUUFBTyxTQUFTLGVBQWUsVUFBVTtBQUMvQyxRQUFJLENBQUNBLE1BQU07QUFFWCxVQUFNLGFBQ0osU0FBUyxXQUFXLGNBQWMsU0FBUyxXQUFXLFVBQ2xEO0FBQUEsTUFDRTtBQUFBLFFBQ0U7QUFBQSxRQUNBLEVBQUUsT0FBTyxPQUFPO0FBQUEsUUFDaEI7QUFBQSxNQUNGO0FBQUEsSUFDRixJQUNBLENBQUM7QUFFUCxVQUFNLFlBQVksV0FBVyxXQUFXO0FBQ3hDLFVBQU0sU0FBUyxVQUFVLE9BQU8sQ0FBQyxPQUFPLFNBQVMsRUFBRSxDQUFDLEVBQUU7QUFDdEQsVUFBTSxVQUFVO0FBQUEsTUFDZDtBQUFBLE1BQ0EsQ0FBQztBQUFBLE1BQ0Q7QUFBQSxRQUNFO0FBQUEsUUFDQSxDQUFDO0FBQUEsUUFDRDtBQUFBLFFBQ0EsR0FBRyxRQUFRLEVBQUUsT0FBTyxRQUFRLEdBQUcsR0FBRyxNQUFNLE9BQU8sVUFBVSxNQUFNLFNBQVM7QUFBQSxNQUMxRTtBQUFBLE1BQ0EsUUFBUSxXQUFXLFFBQVE7QUFBQSxJQUM3QjtBQUNBLFlBQVEsT0FBTztBQUNmLFlBQVEsaUJBQWlCLFVBQVUsTUFBTTtBQUN2QyxzQkFBZ0IsUUFBUTtBQUFBLElBQzFCLENBQUM7QUFFRCxJQUFBQSxNQUFLO0FBQUEsTUFDSCxRQUFRLFVBQVUsWUFBWSxXQUFXLFVBQVUsWUFBWSxPQUFPLFFBQVEsQ0FBQyxDQUFDO0FBQUEsTUFDaEYsUUFBUSxpQkFBaUIsWUFBWSxRQUFRLFVBQVUsYUFBYSxPQUFPLENBQUM7QUFBQSxNQUM1RSxRQUFRLFVBQVUsWUFBWSxVQUFVLFVBQVUsV0FBVyxVQUFVLEdBQUcsR0FBRyxVQUFVO0FBQUEsTUFDdkYsUUFBUSxZQUFZLFFBQVEsV0FBVyxVQUFVLEdBQUcsUUFBUSxDQUFDO0FBQUEsTUFDN0QsR0FBRyxXQUFXLENBQUMsR0FBRyxPQUFPO0FBQUEsTUFDekIsUUFBUSxXQUFXLFFBQVEsV0FBVyxTQUFTLEdBQUcsUUFBUSxDQUFDO0FBQUEsTUFDM0QsaUJBQWlCO0FBQUEsSUFDbkI7QUFBQSxFQUNGO0FBRUEsaUJBQWUsT0FBc0I7QUFDbkMsVUFBTSxVQUFVLFNBQVMsZUFBZSxTQUFTO0FBQ2pELFFBQUksUUFBUyxTQUFRLGNBQWMsSUFBSSxPQUFPLFFBQVEsWUFBWSxFQUFFLE9BQU87QUFFM0UsYUFBUyxlQUFlLFFBQVEsR0FBRyxpQkFBaUIsU0FBUyxNQUFNLE9BQU8sUUFBUSxPQUFPLENBQUM7QUFDMUYsUUFBSSxXQUFXLE1BQU0sYUFBYTtBQUNsQyxXQUFPLFFBQVE7QUFDZixzQkFBa0IsQ0FBQyxTQUFTO0FBQzFCLGlCQUFXO0FBQ1gsYUFBTyxRQUFRO0FBQUEsSUFDakIsQ0FBQztBQUNELGdCQUFZLGlCQUFpQixVQUFVLE1BQU0sT0FBTyxRQUFRLENBQUM7QUFBQSxFQUMvRDtBQUVBLE9BQUssS0FBSzsiLAogICJuYW1lcyI6IFsicmVuZGVyIiwgIm1haW4iXQp9Cg==
