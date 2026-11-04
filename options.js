"use strict";
(() => {
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
      section("General", toggles(featuresIn("General"), settings))
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
//# sourceMappingURL=data:application/json;base64,ewogICJ2ZXJzaW9uIjogMywKICAic291cmNlcyI6IFsiLi4vc3JjL3NldHRpbmdzLnRzIiwgIi4uL3NyYy9vcHRpb25zL29wdGlvbnMudHMiXSwKICAic291cmNlc0NvbnRlbnQiOiBbIi8vIFVzZXIgcHJlZmVyZW5jZXMuIFN0b3JlZCBpbiBjaHJvbWUuc3RvcmFnZS5zeW5jOyBub3RoaW5nIGVsc2UgaXMgZXZlciBwZXJzaXN0ZWQuXG4vL1xuLy8gRWFjaCBlbmFibGVkIGZlYXR1cmUgaWQgaXMgd3JpdHRlbiBhcyBhIHRva2VuIG9uIDxodG1sIGRhdGEtb3Zlcmxvb2stZmVhdHVyZXM9XCIuLi5cIj4sIGFuZCBlYWNoXG4vLyBjaG9pY2UgYXMgaXRzIG93biBhdHRyaWJ1dGUgKDxodG1sIGRhdGEtb3Zlcmxvb2stcGFsZXR0ZT1cInNvbGFyaXplZFwiPiksIHNvIHN0eWxlc2hlZXRzIGdhdGVcbi8vIHRoZW1zZWx2ZXMgd2l0aCBgW2RhdGEtb3Zlcmxvb2stZmVhdHVyZXN+PVwiZGVuc2l0eVwiXWAgb3IgYFtkYXRhLW92ZXJsb29rLXBhbGV0dGU9XCJzb2xhcml6ZWRcIl1gLlxuXG5leHBvcnQgdHlwZSBGZWF0dXJlSWQgPSBrZXlvZiB0eXBlb2YgRkVBVFVSRVM7XG5leHBvcnQgdHlwZSBDaG9pY2VJZCA9IGtleW9mIHR5cGVvZiBDSE9JQ0VTO1xuZXhwb3J0IHR5cGUgU2V0dGluZ3MgPSBSZWNvcmQ8RmVhdHVyZUlkLCBib29sZWFuPiAmIHsgW0sgaW4gQ2hvaWNlSWRdOiBDaG9pY2VWYWx1ZTxLPiB9O1xuZXhwb3J0IHR5cGUgQ2hvaWNlVmFsdWU8SyBleHRlbmRzIENob2ljZUlkPiA9IGtleW9mICh0eXBlb2YgQ0hPSUNFUylbS11bXCJvcHRpb25zXCJdO1xuXG50eXBlIEdyb3VwID0gXCJHZW5lcmFsXCIgfCBcIktleWJvYXJkXCIgfCBcIkRlY2x1dHRlclwiIHwgXCJMb29rXCI7XG5cbmludGVyZmFjZSBGZWF0dXJlIHtcbiAgbGFiZWw6IHN0cmluZztcbiAgZ3JvdXA6IEdyb3VwO1xuICBkZWZhdWx0OiBib29sZWFuO1xufVxuXG5pbnRlcmZhY2UgQ2hvaWNlIHtcbiAgbGFiZWw6IHN0cmluZztcbiAgZ3JvdXA6IEdyb3VwO1xuICBvcHRpb25zOiBSZWNvcmQ8c3RyaW5nLCBzdHJpbmc+O1xuICBkZWZhdWx0OiBzdHJpbmc7XG59XG5cbmV4cG9ydCBjb25zdCBGRUFUVVJFUyA9IHtcbiAgaGVhbHRoSW5kaWNhdG9yOiB7XG4gICAgbGFiZWw6IFwiU2hvdyBhbiBpbmRpY2F0b3Igd2hlbiBPdXRsb29rJ3MgcGFnZSBjaGFuZ2VzIGJyZWFrIE9CQlwiLFxuICAgIGdyb3VwOiBcIkdlbmVyYWxcIixcbiAgICBkZWZhdWx0OiB0cnVlLFxuICB9LFxuICBrZXltYXA6IHsgbGFiZWw6IFwiRmFzdG1haWwtc3R5bGUga2V5Ym9hcmQgc2hvcnRjdXRzXCIsIGdyb3VwOiBcIktleWJvYXJkXCIsIGRlZmF1bHQ6IHRydWUgfSxcbiAgXCJoaWRlLXVwc2VsbFwiOiB7IGxhYmVsOiBcIlVwZ3JhZGUgYW5kICd0cnkgdGhlIG5ld1x1MjAyNicgYmFubmVyc1wiLCBncm91cDogXCJEZWNsdXR0ZXJcIiwgZGVmYXVsdDogdHJ1ZSB9LFxuICBcImhpZGUtYXBwLWJhbm5lcnNcIjoge1xuICAgIGxhYmVsOiBcIk5vdGljZXMgYWNyb3NzIHRoZSB0b3Agb2YgT3V0bG9vayAocG9wLXVwLCBzaWduLWluIGFuZCBzdG9yYWdlIHJlbWluZGVycylcIixcbiAgICBncm91cDogXCJEZWNsdXR0ZXJcIixcbiAgICBkZWZhdWx0OiBmYWxzZSxcbiAgfSxcbiAgXCJoaWRlLWluZm8tYmFyc1wiOiB7XG4gICAgbGFiZWw6IFwiTm90aWNlcyBhYm92ZSBhbiBvcGVuIG1lc3NhZ2UgKCdZb3UgcmVwbGllZCBvblx1MjAyNicsICdFeHRyYSBsaW5lIGJyZWFrcyByZW1vdmVkJylcIixcbiAgICBncm91cDogXCJEZWNsdXR0ZXJcIixcbiAgICBkZWZhdWx0OiBmYWxzZSxcbiAgfSxcbiAgXCJoaWRlLWNvcGlsb3RcIjogeyBsYWJlbDogXCJDb3BpbG90IGJ1dHRvbnMgYW5kIHBhbmVsc1wiLCBncm91cDogXCJEZWNsdXR0ZXJcIiwgZGVmYXVsdDogdHJ1ZSB9LFxuICBcImhpZGUtaGVhZGVyLWV4dHJhc1wiOiB7XG4gICAgbGFiZWw6IFwiVGVhbXMgYW5kIE15IERheSBidXR0b25zIGluIHRoZSBoZWFkZXJcIixcbiAgICBncm91cDogXCJEZWNsdXR0ZXJcIixcbiAgICBkZWZhdWx0OiB0cnVlLFxuICB9LFxuICBcImhpZGUtcmFpbC1leHRyYXNcIjoge1xuICAgIGxhYmVsOiBcIlRvIERvLCBOZXdzbGV0dGVycywgT3JnIEV4cGxvcmVyIGFuZCBPbmVEcml2ZSBpbiB0aGUgYXBwIHJhaWxcIixcbiAgICBncm91cDogXCJEZWNsdXR0ZXJcIixcbiAgICBkZWZhdWx0OiB0cnVlLFxuICB9LFxuICBcImhpZGUtcmliYm9uLWV4dHJhc1wiOiB7XG4gICAgbGFiZWw6IFwiVGVhbXMsIFN3ZWVwLCBRdWljayBzdGVwcywgVml2YSBhbmQgYWRkLWluIGJ1dHRvbnMgaW4gdGhlIHJpYmJvblwiLFxuICAgIGdyb3VwOiBcIkRlY2x1dHRlclwiLFxuICAgIGRlZmF1bHQ6IHRydWUsXG4gIH0sXG4gIFwiaGlkZS1leHRlcm5hbC10YWdcIjogeyBsYWJlbDogXCInRXh0ZXJuYWwnIHNlbmRlciB0YWdzXCIsIGdyb3VwOiBcIkRlY2x1dHRlclwiLCBkZWZhdWx0OiB0cnVlIH0sXG4gIFwiaGlkZS1yb3ctYWN0aW9uc1wiOiB7XG4gICAgbGFiZWw6IFwiRmxhZyBhbmQgcGluIGljb25zIHRoYXQgYXBwZWFyIHdoZW4gaG92ZXJpbmcgYSBtZXNzYWdlXCIsXG4gICAgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsXG4gICAgZGVmYXVsdDogZmFsc2UsXG4gIH0sXG4gIFwiaGlkZS1hdmF0YXJzXCI6IHtcbiAgICBsYWJlbDogXCJTZW5kZXIgcGljdHVyZXMgaW4gdGhlIG1lc3NhZ2UgbGlzdFwiLFxuICAgIGdyb3VwOiBcIkRlY2x1dHRlclwiLFxuICAgIGRlZmF1bHQ6IGZhbHNlLFxuICB9LFxuICBcImhpZGUtZ3JvdXAtaGVhZGVyc1wiOiB7XG4gICAgbGFiZWw6IFwiRGF0ZSBncm91cCBoZWFkZXJzICgnVGhpcyBNb250aCcpIGluIHRoZSBtZXNzYWdlIGxpc3RcIixcbiAgICBncm91cDogXCJEZWNsdXR0ZXJcIixcbiAgICBkZWZhdWx0OiBmYWxzZSxcbiAgfSxcbiAgXCJoaWRlLWZvY3VzZWQtdGFic1wiOiB7IGxhYmVsOiBcIkZvY3VzZWQgLyBPdGhlciB0YWJzXCIsIGdyb3VwOiBcIkRlY2x1dHRlclwiLCBkZWZhdWx0OiBmYWxzZSB9LFxuICBcImhpZGUtcmliYm9uLXRhYnNcIjoge1xuICAgIGxhYmVsOiBcIlJpYmJvbiB0YWJzIChIb21lLCBWaWV3LCBIZWxwKVwiLFxuICAgIGdyb3VwOiBcIkRlY2x1dHRlclwiLFxuICAgIGRlZmF1bHQ6IGZhbHNlLFxuICB9LFxuICBcImhpZGUtYXBwLXJhaWxcIjogeyBsYWJlbDogXCJBcHAgcmFpbCAobGVmdCBpY29uIGJhcilcIiwgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsIGRlZmF1bHQ6IGZhbHNlIH0sXG4gIFwiaGlkZS1yaWJib25cIjogeyBsYWJlbDogXCJSaWJib25cIiwgZ3JvdXA6IFwiRGVjbHV0dGVyXCIsIGRlZmF1bHQ6IGZhbHNlIH0sXG4gIGRlbnNpdHk6IHsgbGFiZWw6IFwiVGlnaHRlciwgZGVuc2VyIGxheW91dFwiLCBncm91cDogXCJMb29rXCIsIGRlZmF1bHQ6IHRydWUgfSxcbn0gYXMgY29uc3Qgc2F0aXNmaWVzIFJlY29yZDxzdHJpbmcsIEZlYXR1cmU+O1xuXG5leHBvcnQgY29uc3QgQ0hPSUNFUyA9IHtcbiAgcGFsZXR0ZToge1xuICAgIGxhYmVsOiBcIkNvbG9yc1wiLFxuICAgIGdyb3VwOiBcIkxvb2tcIixcbiAgICBvcHRpb25zOiB7XG4gICAgICBvdXRsb29rOiBcIk91dGxvb2sncyBvd25cIixcbiAgICAgIHNvbGFyaXplZDogXCJTb2xhcml6ZWRcIixcbiAgICAgIGJhc2UxNjogXCJCYXNlMTYgRGVmYXVsdFwiLFxuICAgICAgZmFzdG1haWw6IFwiRmFzdG1haWxcIixcbiAgICAgIHBhcGVyOiBcIlBhcGVyXCIsXG4gICAgfSxcbiAgICBkZWZhdWx0OiBcIm91dGxvb2tcIixcbiAgfSxcbiAgbW9kZToge1xuICAgIGxhYmVsOiBcIkxpZ2h0IG9yIGRhcmtcIixcbiAgICBncm91cDogXCJMb29rXCIsXG4gICAgb3B0aW9uczogeyBhdXRvOiBcIkZvbGxvdyBPdXRsb29rXCIsIGxpZ2h0OiBcIkxpZ2h0XCIsIGRhcms6IFwiRGFya1wiIH0sXG4gICAgZGVmYXVsdDogXCJhdXRvXCIsXG4gIH0sXG4gIGxheW91dDoge1xuICAgIGxhYmVsOiBcIkxheW91dFwiLFxuICAgIGdyb3VwOiBcIkxvb2tcIixcbiAgICBvcHRpb25zOiB7XG4gICAgICBvdXRsb29rOiBcIk91dGxvb2sncyBvd25cIixcbiAgICAgIGNhbG06IFwiQ2FsbSB0aHJlZS1wYW5lXCIsXG4gICAgICBmYXN0bWFpbDogXCJGYXN0bWFpbCAobGlzdCBmaXJzdClcIixcbiAgICAgIGZvY3VzOiBcIkZvY3VzIChzaW5nbGUgY29sdW1uKVwiLFxuICAgIH0sXG4gICAgZGVmYXVsdDogXCJvdXRsb29rXCIsXG4gIH0sXG59IGFzIGNvbnN0IHNhdGlzZmllcyBSZWNvcmQ8c3RyaW5nLCBDaG9pY2U+O1xuXG5leHBvcnQgY29uc3QgRkVBVFVSRV9JRFMgPSBPYmplY3Qua2V5cyhGRUFUVVJFUykgYXMgRmVhdHVyZUlkW107XG5leHBvcnQgY29uc3QgQ0hPSUNFX0lEUyA9IE9iamVjdC5rZXlzKENIT0lDRVMpIGFzIENob2ljZUlkW107XG5cbmV4cG9ydCBmdW5jdGlvbiBkZWZhdWx0U2V0dGluZ3MoKTogU2V0dGluZ3Mge1xuICByZXR1cm4ge1xuICAgIC4uLk9iamVjdC5mcm9tRW50cmllcyhGRUFUVVJFX0lEUy5tYXAoKGlkKSA9PiBbaWQsIEZFQVRVUkVTW2lkXS5kZWZhdWx0XSkpLFxuICAgIC4uLk9iamVjdC5mcm9tRW50cmllcyhDSE9JQ0VfSURTLm1hcCgoaWQpID0+IFtpZCwgQ0hPSUNFU1tpZF0uZGVmYXVsdF0pKSxcbiAgfSBhcyBTZXR0aW5ncztcbn1cblxuLyoqIE1lcmdlIHdoYXRldmVyIGlzIHN0b3JlZCBvdmVyIHRoZSBkZWZhdWx0cywgaWdub3JpbmcgdW5rbm93biBvciBtaXN0eXBlZCBrZXlzLiAqL1xuZXhwb3J0IGZ1bmN0aW9uIG5vcm1hbGl6ZShzdG9yZWQ6IFJlY29yZDxzdHJpbmcsIHVua25vd24+KTogU2V0dGluZ3Mge1xuICBjb25zdCBzZXR0aW5ncyA9IGRlZmF1bHRTZXR0aW5ncygpIGFzIFJlY29yZDxzdHJpbmcsIHVua25vd24+O1xuICBmb3IgKGNvbnN0IGlkIG9mIEZFQVRVUkVfSURTKSB7XG4gICAgaWYgKHR5cGVvZiBzdG9yZWRbaWRdID09PSBcImJvb2xlYW5cIikgc2V0dGluZ3NbaWRdID0gc3RvcmVkW2lkXTtcbiAgfVxuICBmb3IgKGNvbnN0IGlkIG9mIENIT0lDRV9JRFMpIHtcbiAgICBjb25zdCB2YWx1ZSA9IHN0b3JlZFtpZF07XG4gICAgaWYgKHR5cGVvZiB2YWx1ZSA9PT0gXCJzdHJpbmdcIiAmJiB2YWx1ZSBpbiBDSE9JQ0VTW2lkXS5vcHRpb25zKSBzZXR0aW5nc1tpZF0gPSB2YWx1ZTtcbiAgfVxuICByZXR1cm4gc2V0dGluZ3MgYXMgU2V0dGluZ3M7XG59XG5cbmV4cG9ydCBmdW5jdGlvbiBmZWF0dXJlVG9rZW5zKHNldHRpbmdzOiBTZXR0aW5ncyk6IHN0cmluZyB7XG4gIHJldHVybiBGRUFUVVJFX0lEUy5maWx0ZXIoKGlkKSA9PiBzZXR0aW5nc1tpZF0pLmpvaW4oXCIgXCIpO1xufVxuXG4vKiogV3JpdGVzIHRoZSBzZXR0aW5ncyBvbnRvIDxodG1sPiBhcyB0aGUgYXR0cmlidXRlcyB0aGUgc3R5bGVzaGVldHMga2V5IG9mZi4gKi9cbmV4cG9ydCBmdW5jdGlvbiBhcHBseVRvRG9jdW1lbnQoc2V0dGluZ3M6IFNldHRpbmdzLCByb290OiBIVE1MRWxlbWVudCk6IHZvaWQge1xuICByb290LnNldEF0dHJpYnV0ZShcImRhdGEtb3Zlcmxvb2stZmVhdHVyZXNcIiwgZmVhdHVyZVRva2VucyhzZXR0aW5ncykpO1xuICBmb3IgKGNvbnN0IGlkIG9mIENIT0lDRV9JRFMpIHJvb3Quc2V0QXR0cmlidXRlKGBkYXRhLW92ZXJsb29rLSR7aWR9YCwgc2V0dGluZ3NbaWRdKTtcbn1cblxuZXhwb3J0IGFzeW5jIGZ1bmN0aW9uIGxvYWRTZXR0aW5ncygpOiBQcm9taXNlPFNldHRpbmdzPiB7XG4gIHJldHVybiBub3JtYWxpemUoYXdhaXQgY2hyb21lLnN0b3JhZ2Uuc3luYy5nZXQobnVsbCkpO1xufVxuXG5leHBvcnQgYXN5bmMgZnVuY3Rpb24gc2F2ZVNldHRpbmcoXG4gIGlkOiBGZWF0dXJlSWQgfCBDaG9pY2VJZCxcbiAgdmFsdWU6IGJvb2xlYW4gfCBzdHJpbmcsXG4pOiBQcm9taXNlPHZvaWQ+IHtcbiAgYXdhaXQgY2hyb21lLnN0b3JhZ2Uuc3luYy5zZXQoeyBbaWRdOiB2YWx1ZSB9KTtcbn1cblxuZXhwb3J0IGZ1bmN0aW9uIG9uU2V0dGluZ3NDaGFuZ2VkKGNhbGxiYWNrOiAoc2V0dGluZ3M6IFNldHRpbmdzKSA9PiB2b2lkKTogdm9pZCB7XG4gIGNocm9tZS5zdG9yYWdlLm9uQ2hhbmdlZC5hZGRMaXN0ZW5lcigoX2NoYW5nZXMsIGFyZWEpID0+IHtcbiAgICBpZiAoYXJlYSA9PT0gXCJzeW5jXCIpIHZvaWQgbG9hZFNldHRpbmdzKCkudGhlbihjYWxsYmFjayk7XG4gIH0pO1xufVxuIiwgIi8vIFRoZSBvcHRpb25zIHBhZ2UsIGFsc28gdGhlIHRvb2xiYXIgcG9wdXAuIENoYW5nZXMgc2F2ZSBpbW1lZGlhdGVseSBhbmQgYXBwbHkgbGl2ZSBpbiBPdXRsb29rLlxuXG5pbXBvcnQge1xuICBDSE9JQ0VTLFxuICB0eXBlIENob2ljZUlkLFxuICBGRUFUVVJFX0lEUyxcbiAgRkVBVFVSRVMsXG4gIHR5cGUgRmVhdHVyZUlkLFxuICBsb2FkU2V0dGluZ3MsXG4gIG9uU2V0dGluZ3NDaGFuZ2VkLFxuICB0eXBlIFNldHRpbmdzLFxuICBzYXZlU2V0dGluZyxcbn0gZnJvbSBcIi4uL3NldHRpbmdzLnRzXCI7XG5cbmNvbnN0IFNIT1JUX0xBWU9VVF9OQU1FUzogUmVjb3JkPHN0cmluZywgc3RyaW5nPiA9IHtcbiAgb3V0bG9vazogXCJPdXRsb29rXCIsXG4gIGNhbG06IFwiQ2FsbVwiLFxuICBmYXN0bWFpbDogXCJGYXN0bWFpbFwiLFxuICBmb2N1czogXCJGb2N1c1wiLFxufTtcblxuY29uc3QgU0hPUlRfUEFMRVRURV9OQU1FUzogUmVjb3JkPHN0cmluZywgc3RyaW5nPiA9IHtcbiAgb3V0bG9vazogXCJPdXRsb29rXCIsXG4gIHNvbGFyaXplZDogXCJTb2xhcml6ZWRcIixcbiAgYmFzZTE2OiBcIkJhc2UxNlwiLFxuICBmYXN0bWFpbDogXCJGYXN0bWFpbFwiLFxuICBwYXBlcjogXCJQYXBlclwiLFxufTtcblxuLyoqIFR3byBsaW5lcyBvZiBhIG1lc3NhZ2Ugcm93IChzdWJqZWN0LCBwcmV2aWV3KSBhdCB0aGUgZ2l2ZW4gcG9zaXRpb24gaW4gYSBza2V0Y2guICovXG5mdW5jdGlvbiByb3dzKHg6IG51bWJlciwgeTogbnVtYmVyLCB3aWR0aDogbnVtYmVyKTogc3RyaW5nIHtcbiAgcmV0dXJuIFswLCAxLCAyLCAzLCA0XVxuICAgIC5tYXAoKGkpID0+IHtcbiAgICAgIGNvbnN0IHRvcCA9IHkgKyBpICogNi41O1xuICAgICAgY29uc3Qgc3ViamVjdCA9IGkgPT09IDAgPyBcImhpXCIgOiBcInJvd1wiO1xuICAgICAgY29uc3Qgc3ViamVjdFdpZHRoID0gaSA9PT0gMCA/IHdpZHRoICogMC41NSA6IHdpZHRoO1xuICAgICAgcmV0dXJuIChcbiAgICAgICAgYDxyZWN0IGNsYXNzPVwiJHtzdWJqZWN0fVwiIHg9XCIke3h9XCIgeT1cIiR7dG9wfVwiIHdpZHRoPVwiJHtzdWJqZWN0V2lkdGh9XCIgaGVpZ2h0PVwiMi41XCIgcng9XCIxXCIvPmAgK1xuICAgICAgICBgPHJlY3QgY2xhc3M9XCJyb3dcIiB4PVwiJHt4fVwiIHk9XCIke3RvcCArIDN9XCIgd2lkdGg9XCIke3dpZHRoICogMC43NX1cIiBoZWlnaHQ9XCIxLjVcIiByeD1cIjAuNzVcIi8+YFxuICAgICAgKTtcbiAgICB9KVxuICAgIC5qb2luKFwiXCIpO1xufVxuXG4vLyBXaXJlZnJhbWVzIG9mIGVhY2ggbGF5b3V0OiB2aWV3Qm94IDY0XHUwMEQ3NDQ7IGNocm9tZSBiYXIsIHRoZW4gcGFuZXMgYW5kIG1lc3NhZ2Ugcm93cy5cbmNvbnN0IExBWU9VVF9TS0VUQ0hFUzogUmVjb3JkPHN0cmluZywgc3RyaW5nPiA9IHtcbiAgb3V0bG9vazogYDxyZWN0IGNsYXNzPVwiY2hyb21lXCIgd2lkdGg9XCI2NFwiIGhlaWdodD1cIjVcIi8+XG4gICAgPHJlY3QgY2xhc3M9XCJwYW5lXCIgeD1cIjBcIiB5PVwiNVwiIHdpZHRoPVwiNVwiIGhlaWdodD1cIjM5XCIvPlxuICAgIDxyZWN0IGNsYXNzPVwicGFuZVwiIHg9XCI2XCIgeT1cIjdcIiB3aWR0aD1cIjEyXCIgaGVpZ2h0PVwiMzVcIiByeD1cIjFcIi8+XG4gICAgJHtyb3dzKDIwLCA4LCAxNSl9PHJlY3QgY2xhc3M9XCJwYW5lXCIgeD1cIjM3XCIgeT1cIjdcIiB3aWR0aD1cIjI1XCIgaGVpZ2h0PVwiMzVcIiByeD1cIjFcIi8+YCxcbiAgY2FsbTogYDxyZWN0IGNsYXNzPVwiY2hyb21lXCIgd2lkdGg9XCI2NFwiIGhlaWdodD1cIjVcIi8+XG4gICAgPHJlY3QgY2xhc3M9XCJwYW5lXCIgeD1cIjJcIiB5PVwiOFwiIHdpZHRoPVwiMTNcIiBoZWlnaHQ9XCIzM1wiIHJ4PVwiMlwiLz5cbiAgICAke3Jvd3MoMTgsIDksIDE2KX08cmVjdCBjbGFzcz1cInBhbmVcIiB4PVwiMzdcIiB5PVwiOFwiIHdpZHRoPVwiMjVcIiBoZWlnaHQ9XCIzM1wiIHJ4PVwiMlwiLz5gLFxuICBmYXN0bWFpbDogYDxyZWN0IGNsYXNzPVwiY2hyb21lXCIgd2lkdGg9XCI2NFwiIGhlaWdodD1cIjVcIi8+XG4gICAgPHJlY3QgY2xhc3M9XCJwYW5lXCIgeD1cIjJcIiB5PVwiOFwiIHdpZHRoPVwiMTNcIiBoZWlnaHQ9XCIzM1wiIHJ4PVwiMlwiLz5cbiAgICA8cmVjdCBjbGFzcz1cImhpXCIgeD1cIjRcIiB5PVwiMTBcIiB3aWR0aD1cIjlcIiBoZWlnaHQ9XCIzXCIgcng9XCIxLjVcIi8+JHtyb3dzKDE4LCA5LCA0NCl9YCxcbiAgZm9jdXM6IGA8cmVjdCBjbGFzcz1cImNocm9tZVwiIHdpZHRoPVwiNjRcIiBoZWlnaHQ9XCI1XCIvPiR7cm93cygxNiwgOSwgMzIpfWAsXG59O1xuXG5jb25zdCBodG1sID0gZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50O1xuY29uc3QgcHJlZmVyc0RhcmsgPSBtYXRjaE1lZGlhKFwiKHByZWZlcnMtY29sb3Itc2NoZW1lOiBkYXJrKVwiKTtcblxuLyoqIFRoZSBsaWdodC9kYXJrIG1vZGUgaW4gZWZmZWN0OiBcImF1dG9cIiBmb2xsb3dzIHRoZSBzeXN0ZW0gaGVyZSwgc3RhbmRpbmcgaW4gZm9yIE91dGxvb2suICovXG5mdW5jdGlvbiBlZmZlY3RpdmVNb2RlKHNldHRpbmdzOiBTZXR0aW5ncyk6IFwibGlnaHRcIiB8IFwiZGFya1wiIHtcbiAgaWYgKHNldHRpbmdzLm1vZGUgPT09IFwibGlnaHRcIiB8fCBzZXR0aW5ncy5tb2RlID09PSBcImRhcmtcIikgcmV0dXJuIHNldHRpbmdzLm1vZGU7XG4gIHJldHVybiBwcmVmZXJzRGFyay5tYXRjaGVzID8gXCJkYXJrXCIgOiBcImxpZ2h0XCI7XG59XG5cbi8qKiBEcmVzc2VzIHRoZSBwYWdlIGluIHRoZSBjaG9zZW4gcGFsZXR0ZSAocGFsZXR0ZXMuY3NzIGtleXMgb2ZmIHRoZXNlIGF0dHJpYnV0ZXMpLiAqL1xuZnVuY3Rpb24gYXBwbHlUaGVtZShzZXR0aW5nczogU2V0dGluZ3MpOiB2b2lkIHtcbiAgaHRtbC5kYXRhc2V0Lm92ZXJsb29rUGFsZXR0ZSA9IHNldHRpbmdzLnBhbGV0dGU7XG4gIGh0bWwuZGF0YXNldC5vdmVybG9va01vZGUgPSBzZXR0aW5ncy5tb2RlO1xuICBodG1sLmRhdGFzZXQub3Zlcmxvb2tPdXRsb29rU2NoZW1lID0gcHJlZmVyc0RhcmsubWF0Y2hlcyA/IFwiZGFya1wiIDogXCJsaWdodFwiO1xufVxuXG5mdW5jdGlvbiBlbDxLIGV4dGVuZHMga2V5b2YgSFRNTEVsZW1lbnRUYWdOYW1lTWFwPihcbiAgdGFnOiBLLFxuICBhdHRyczogUmVjb3JkPHN0cmluZywgc3RyaW5nPiA9IHt9LFxuICAuLi5jaGlsZHJlbjogKE5vZGUgfCBzdHJpbmcpW11cbik6IEhUTUxFbGVtZW50VGFnTmFtZU1hcFtLXSB7XG4gIGNvbnN0IG5vZGUgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KHRhZyk7XG4gIGZvciAoY29uc3QgW2ssIHZdIG9mIE9iamVjdC5lbnRyaWVzKGF0dHJzKSkgbm9kZS5zZXRBdHRyaWJ1dGUoaywgdik7XG4gIG5vZGUuYXBwZW5kKC4uLmNoaWxkcmVuKTtcbiAgcmV0dXJuIG5vZGU7XG59XG5cbmZ1bmN0aW9uIHNlY3Rpb24odGl0bGU6IHN0cmluZywgLi4uY2hpbGRyZW46IE5vZGVbXSk6IEhUTUxFbGVtZW50IHtcbiAgcmV0dXJuIGVsKFwic2VjdGlvblwiLCB7fSwgZWwoXCJoMlwiLCB7fSwgdGl0bGUpLCAuLi5jaGlsZHJlbik7XG59XG5cbi8qKiBNdXR1YWxseSBleGNsdXNpdmUgYnV0dG9ucyBmb3IgYSBjaG9pY2U7IGByZW5kZXJgIGRyYXdzIGVhY2ggb3B0aW9uLiAqL1xuZnVuY3Rpb24gY2hvaWNlR3JvdXAoXG4gIGlkOiBDaG9pY2VJZCxcbiAgc2V0dGluZ3M6IFNldHRpbmdzLFxuICBjbGFzc05hbWU6IHN0cmluZyxcbiAgcmVuZGVyOiAodmFsdWU6IHN0cmluZywgbGFiZWw6IHN0cmluZykgPT4gSFRNTEJ1dHRvbkVsZW1lbnQsXG4pOiBIVE1MRWxlbWVudCB7XG4gIGNvbnN0IGdyb3VwID0gZWwoXCJkaXZcIiwge1xuICAgIGNsYXNzOiBjbGFzc05hbWUsXG4gICAgcm9sZTogXCJyYWRpb2dyb3VwXCIsXG4gICAgXCJhcmlhLWxhYmVsXCI6IENIT0lDRVNbaWRdLmxhYmVsLFxuICB9KTtcbiAgZm9yIChjb25zdCBbdmFsdWUsIGxhYmVsXSBvZiBPYmplY3QuZW50cmllcyhDSE9JQ0VTW2lkXS5vcHRpb25zKSkge1xuICAgIGNvbnN0IGJ1dHRvbiA9IHJlbmRlcih2YWx1ZSwgbGFiZWwpO1xuICAgIGJ1dHRvbi50eXBlID0gXCJidXR0b25cIjtcbiAgICBidXR0b24uc2V0QXR0cmlidXRlKFwicm9sZVwiLCBcInJhZGlvXCIpO1xuICAgIGJ1dHRvbi5zZXRBdHRyaWJ1dGUoXCJhcmlhLWNoZWNrZWRcIiwgU3RyaW5nKHNldHRpbmdzW2lkXSA9PT0gdmFsdWUpKTtcbiAgICBidXR0b24uYWRkRXZlbnRMaXN0ZW5lcihcImNsaWNrXCIsICgpID0+IHZvaWQgc2F2ZVNldHRpbmcoaWQsIHZhbHVlKSk7XG4gICAgZ3JvdXAuYXBwZW5kKGJ1dHRvbik7XG4gIH1cbiAgcmV0dXJuIGdyb3VwO1xufVxuXG4vKiogQSB0aW55IHdpbmRvdyBwYWludGVkIGluIGl0cyBvd24gcGFsZXR0ZSwgaW4gdGhlIG1vZGUgY3VycmVudGx5IGluIGVmZmVjdC4gKi9cbmZ1bmN0aW9uIHN3YXRjaChzZXR0aW5nczogU2V0dGluZ3MpIHtcbiAgcmV0dXJuICh2YWx1ZTogc3RyaW5nLCBsYWJlbDogc3RyaW5nKTogSFRNTEJ1dHRvbkVsZW1lbnQgPT4ge1xuICAgIGNvbnN0IHByZXZpZXcgPSBlbChcbiAgICAgIFwic3BhblwiLFxuICAgICAge1xuICAgICAgICBjbGFzczogXCJwcmV2aWV3XCIsXG4gICAgICAgIFwiZGF0YS1vdmVybG9vay1zd2F0Y2hcIjogXCJcIixcbiAgICAgICAgXCJkYXRhLW92ZXJsb29rLXBhbGV0dGVcIjogdmFsdWUsXG4gICAgICAgIFwiZGF0YS1vdmVybG9vay1tb2RlXCI6IGVmZmVjdGl2ZU1vZGUoc2V0dGluZ3MpLFxuICAgICAgfSxcbiAgICAgIGVsKFwic3BhblwiLCB7IGNsYXNzOiBcImJhclwiIH0pLFxuICAgICAgZWwoXCJzcGFuXCIsIHsgY2xhc3M6IFwic2lkZVwiIH0pLFxuICAgICAgZWwoXG4gICAgICAgIFwic3BhblwiLFxuICAgICAgICB7IGNsYXNzOiBcImJvZHlcIiB9LFxuICAgICAgICBlbChcInNwYW5cIiwgeyBjbGFzczogXCJsaW5lIHNob3J0XCIgfSksXG4gICAgICAgIGVsKFwic3BhblwiLCB7IGNsYXNzOiBcImxpbmVcIiB9KSxcbiAgICAgICAgZWwoXCJzcGFuXCIsIHsgY2xhc3M6IFwibGluZSBmYWludFwiIH0pLFxuICAgICAgICBlbChcInNwYW5cIiwgeyBjbGFzczogXCJsaW5lXCIgfSksXG4gICAgICApLFxuICAgICk7XG4gICAgcmV0dXJuIGVsKFxuICAgICAgXCJidXR0b25cIixcbiAgICAgIHsgY2xhc3M6IFwic3dhdGNoXCIsIHRpdGxlOiBsYWJlbCB9LFxuICAgICAgcHJldmlldyxcbiAgICAgIFNIT1JUX1BBTEVUVEVfTkFNRVNbdmFsdWVdID8/IGxhYmVsLFxuICAgICk7XG4gIH07XG59XG5cbmZ1bmN0aW9uIGxheW91dENhcmQodmFsdWU6IHN0cmluZywgbGFiZWw6IHN0cmluZyk6IEhUTUxCdXR0b25FbGVtZW50IHtcbiAgY29uc3QgYnV0dG9uID0gZWwoXCJidXR0b25cIiwgeyBjbGFzczogXCJsYXlvdXRcIiwgdGl0bGU6IGxhYmVsIH0pO1xuICBidXR0b24uaW5uZXJIVE1MID0gYDxzdmcgdmlld0JveD1cIjAgMCA2NCA0NFwiIGFyaWEtaGlkZGVuPVwidHJ1ZVwiPiR7TEFZT1VUX1NLRVRDSEVTW3ZhbHVlXSA/PyBcIlwifTwvc3ZnPmA7XG4gIGJ1dHRvbi5hcHBlbmQoU0hPUlRfTEFZT1VUX05BTUVTW3ZhbHVlXSA/PyBsYWJlbCk7XG4gIHJldHVybiBidXR0b247XG59XG5cbmZ1bmN0aW9uIHNlZ21lbnQodmFsdWU6IHN0cmluZywgbGFiZWw6IHN0cmluZyk6IEhUTUxCdXR0b25FbGVtZW50IHtcbiAgcmV0dXJuIGVsKFwiYnV0dG9uXCIsIHsgXCJkYXRhLXZhbHVlXCI6IHZhbHVlIH0sIGxhYmVsKTtcbn1cblxuZnVuY3Rpb24gdG9nZ2xlKGlkOiBGZWF0dXJlSWQsIHNldHRpbmdzOiBTZXR0aW5ncyk6IEhUTUxMYWJlbEVsZW1lbnQge1xuICBjb25zdCBpbnB1dCA9IGVsKFwiaW5wdXRcIiwgeyB0eXBlOiBcImNoZWNrYm94XCIsIGlkOiBgdG9nZ2xlLSR7aWR9YCB9KTtcbiAgaW5wdXQuY2hlY2tlZCA9IHNldHRpbmdzW2lkXTtcbiAgaW5wdXQuYWRkRXZlbnRMaXN0ZW5lcihcImNoYW5nZVwiLCAoKSA9PiB2b2lkIHNhdmVTZXR0aW5nKGlkLCBpbnB1dC5jaGVja2VkKSk7XG4gIHJldHVybiBlbChcbiAgICBcImxhYmVsXCIsXG4gICAgeyBjbGFzczogXCJ0b2dnbGVcIiwgZm9yOiBgdG9nZ2xlLSR7aWR9YCB9LFxuICAgIGVsKFwic3BhblwiLCB7fSwgRkVBVFVSRVNbaWRdLmxhYmVsKSxcbiAgICBpbnB1dCxcbiAgKTtcbn1cblxuZnVuY3Rpb24gdG9nZ2xlcyhpZHM6IEZlYXR1cmVJZFtdLCBzZXR0aW5nczogU2V0dGluZ3MpOiBIVE1MRWxlbWVudCB7XG4gIHJldHVybiBlbChcImRpdlwiLCB7IGNsYXNzOiBcInRvZ2dsZXNcIiB9LCAuLi5pZHMubWFwKChpZCkgPT4gdG9nZ2xlKGlkLCBzZXR0aW5ncykpKTtcbn1cblxuY29uc3QgZmVhdHVyZXNJbiA9IChncm91cDogc3RyaW5nKSA9PiBGRUFUVVJFX0lEUy5maWx0ZXIoKGlkKSA9PiBGRUFUVVJFU1tpZF0uZ3JvdXAgPT09IGdyb3VwKTtcblxubGV0IGRlY2x1dHRlck9wZW4gPSBmYWxzZTtcblxuZnVuY3Rpb24gcmVuZGVyKHNldHRpbmdzOiBTZXR0aW5ncyk6IHZvaWQge1xuICBhcHBseVRoZW1lKHNldHRpbmdzKTtcbiAgY29uc3QgbWFpbiA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKFwic2V0dGluZ3NcIik7XG4gIGlmICghbWFpbikgcmV0dXJuO1xuXG4gIGNvbnN0IGxheW91dEhpbnQgPVxuICAgIHNldHRpbmdzLmxheW91dCA9PT0gXCJmYXN0bWFpbFwiIHx8IHNldHRpbmdzLmxheW91dCA9PT0gXCJmb2N1c1wiXG4gICAgICA/IFtcbiAgICAgICAgICBlbChcbiAgICAgICAgICAgIFwicFwiLFxuICAgICAgICAgICAgeyBjbGFzczogXCJoaW50XCIgfSxcbiAgICAgICAgICAgIFwiUGFpciB3aXRoIE91dGxvb2sncyBTZXR0aW5ncyBcdTIxOTIgTWFpbCBcdTIxOTIgTGF5b3V0IFx1MjE5MiBSZWFkaW5nIHBhbmU6IEhpZGUuXCIsXG4gICAgICAgICAgKSxcbiAgICAgICAgXVxuICAgICAgOiBbXTtcblxuICBjb25zdCBkZWNsdXR0ZXIgPSBmZWF0dXJlc0luKFwiRGVjbHV0dGVyXCIpO1xuICBjb25zdCBoaWRkZW4gPSBkZWNsdXR0ZXIuZmlsdGVyKChpZCkgPT4gc2V0dGluZ3NbaWRdKS5sZW5ndGg7XG4gIGNvbnN0IGRldGFpbHMgPSBlbChcbiAgICBcImRldGFpbHNcIixcbiAgICB7fSxcbiAgICBlbChcbiAgICAgIFwic3VtbWFyeVwiLFxuICAgICAge30sXG4gICAgICBcIkRlY2x1dHRlclwiLFxuICAgICAgZWwoXCJzcGFuXCIsIHsgY2xhc3M6IFwiY291bnRcIiB9LCBgJHtoaWRkZW59IG9mICR7ZGVjbHV0dGVyLmxlbmd0aH0gaGlkZGVuYCksXG4gICAgKSxcbiAgICB0b2dnbGVzKGRlY2x1dHRlciwgc2V0dGluZ3MpLFxuICApO1xuICBkZXRhaWxzLm9wZW4gPSBkZWNsdXR0ZXJPcGVuO1xuICBkZXRhaWxzLmFkZEV2ZW50TGlzdGVuZXIoXCJ0b2dnbGVcIiwgKCkgPT4ge1xuICAgIGRlY2x1dHRlck9wZW4gPSBkZXRhaWxzLm9wZW47XG4gIH0pO1xuXG4gIG1haW4ucmVwbGFjZUNoaWxkcmVuKFxuICAgIHNlY3Rpb24oXCJDb2xvcnNcIiwgY2hvaWNlR3JvdXAoXCJwYWxldHRlXCIsIHNldHRpbmdzLCBcInN3YXRjaGVzXCIsIHN3YXRjaChzZXR0aW5ncykpKSxcbiAgICBzZWN0aW9uKFwiTGlnaHQgb3IgZGFya1wiLCBjaG9pY2VHcm91cChcIm1vZGVcIiwgc2V0dGluZ3MsIFwic2VnbWVudGVkXCIsIHNlZ21lbnQpKSxcbiAgICBzZWN0aW9uKFwiTGF5b3V0XCIsIGNob2ljZUdyb3VwKFwibGF5b3V0XCIsIHNldHRpbmdzLCBcImxheW91dHNcIiwgbGF5b3V0Q2FyZCksIC4uLmxheW91dEhpbnQpLFxuICAgIHNlY3Rpb24oXCJLZXlib2FyZFwiLCB0b2dnbGVzKGZlYXR1cmVzSW4oXCJLZXlib2FyZFwiKSwgc2V0dGluZ3MpKSxcbiAgICBlbChcInNlY3Rpb25cIiwge30sIGRldGFpbHMpLFxuICAgIHNlY3Rpb24oXCJHZW5lcmFsXCIsIHRvZ2dsZXMoZmVhdHVyZXNJbihcIkdlbmVyYWxcIiksIHNldHRpbmdzKSksXG4gICk7XG59XG5cbmFzeW5jIGZ1bmN0aW9uIG1haW4oKTogUHJvbWlzZTx2b2lkPiB7XG4gIGNvbnN0IHZlcnNpb24gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChcInZlcnNpb25cIik7XG4gIGlmICh2ZXJzaW9uKSB2ZXJzaW9uLnRleHRDb250ZW50ID0gYHYke2Nocm9tZS5ydW50aW1lLmdldE1hbmlmZXN0KCkudmVyc2lvbn1gO1xuICAvLyBQaWNrcyB1cCBuZXcgZmlsZXMgb24gZGlzayAoYWZ0ZXIgYGdpdCBwdWxsYCBpbiB0aGUgaW5zdGFsbCBmb2xkZXIpOyBzZWUgUkVBRE1FLlxuICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChcInJlbG9hZFwiKT8uYWRkRXZlbnRMaXN0ZW5lcihcImNsaWNrXCIsICgpID0+IGNocm9tZS5ydW50aW1lLnJlbG9hZCgpKTtcbiAgbGV0IHNldHRpbmdzID0gYXdhaXQgbG9hZFNldHRpbmdzKCk7XG4gIHJlbmRlcihzZXR0aW5ncyk7XG4gIG9uU2V0dGluZ3NDaGFuZ2VkKChuZXh0KSA9PiB7XG4gICAgc2V0dGluZ3MgPSBuZXh0O1xuICAgIHJlbmRlcihzZXR0aW5ncyk7XG4gIH0pO1xuICBwcmVmZXJzRGFyay5hZGRFdmVudExpc3RlbmVyKFwiY2hhbmdlXCIsICgpID0+IHJlbmRlcihzZXR0aW5ncykpO1xufVxuXG52b2lkIG1haW4oKTtcbiJdLAogICJtYXBwaW5ncyI6ICI7OztBQTBCTyxNQUFNLFdBQVc7QUFBQSxJQUN0QixpQkFBaUI7QUFBQSxNQUNmLE9BQU87QUFBQSxNQUNQLE9BQU87QUFBQSxNQUNQLFNBQVM7QUFBQSxJQUNYO0FBQUEsSUFDQSxRQUFRLEVBQUUsT0FBTyxxQ0FBcUMsT0FBTyxZQUFZLFNBQVMsS0FBSztBQUFBLElBQ3ZGLGVBQWUsRUFBRSxPQUFPLDJDQUFzQyxPQUFPLGFBQWEsU0FBUyxLQUFLO0FBQUEsSUFDaEcsb0JBQW9CO0FBQUEsTUFDbEIsT0FBTztBQUFBLE1BQ1AsT0FBTztBQUFBLE1BQ1AsU0FBUztBQUFBLElBQ1g7QUFBQSxJQUNBLGtCQUFrQjtBQUFBLE1BQ2hCLE9BQU87QUFBQSxNQUNQLE9BQU87QUFBQSxNQUNQLFNBQVM7QUFBQSxJQUNYO0FBQUEsSUFDQSxnQkFBZ0IsRUFBRSxPQUFPLDhCQUE4QixPQUFPLGFBQWEsU0FBUyxLQUFLO0FBQUEsSUFDekYsc0JBQXNCO0FBQUEsTUFDcEIsT0FBTztBQUFBLE1BQ1AsT0FBTztBQUFBLE1BQ1AsU0FBUztBQUFBLElBQ1g7QUFBQSxJQUNBLG9CQUFvQjtBQUFBLE1BQ2xCLE9BQU87QUFBQSxNQUNQLE9BQU87QUFBQSxNQUNQLFNBQVM7QUFBQSxJQUNYO0FBQUEsSUFDQSxzQkFBc0I7QUFBQSxNQUNwQixPQUFPO0FBQUEsTUFDUCxPQUFPO0FBQUEsTUFDUCxTQUFTO0FBQUEsSUFDWDtBQUFBLElBQ0EscUJBQXFCLEVBQUUsT0FBTywwQkFBMEIsT0FBTyxhQUFhLFNBQVMsS0FBSztBQUFBLElBQzFGLG9CQUFvQjtBQUFBLE1BQ2xCLE9BQU87QUFBQSxNQUNQLE9BQU87QUFBQSxNQUNQLFNBQVM7QUFBQSxJQUNYO0FBQUEsSUFDQSxnQkFBZ0I7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
