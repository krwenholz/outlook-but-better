// The options page, also the toolbar popup. Changes save immediately and apply live in Outlook.

import {
  CHOICES,
  type ChoiceId,
  FEATURE_IDS,
  FEATURES,
  type FeatureId,
  loadSettings,
  onSettingsChanged,
  type Settings,
  saveSetting,
} from "../settings.ts";

const SHORT_LAYOUT_NAMES: Record<string, string> = {
  outlook: "Outlook",
  calm: "Calm",
  fastmail: "Fastmail",
  focus: "Focus",
};

const SHORT_PALETTE_NAMES: Record<string, string> = {
  outlook: "Outlook",
  solarized: "Solarized",
  base16: "Base16",
  fastmail: "Fastmail",
  paper: "Paper",
};

/** Two lines of a message row (subject, preview) at the given position in a sketch. */
function rows(x: number, y: number, width: number): string {
  return [0, 1, 2, 3, 4]
    .map((i) => {
      const top = y + i * 6.5;
      const subject = i === 0 ? "hi" : "row";
      const subjectWidth = i === 0 ? width * 0.55 : width;
      return (
        `<rect class="${subject}" x="${x}" y="${top}" width="${subjectWidth}" height="2.5" rx="1"/>` +
        `<rect class="row" x="${x}" y="${top + 3}" width="${width * 0.75}" height="1.5" rx="0.75"/>`
      );
    })
    .join("");
}

// Wireframes of each layout: viewBox 64×44; chrome bar, then panes and message rows.
const LAYOUT_SKETCHES: Record<string, string> = {
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
  focus: `<rect class="chrome" width="64" height="5"/>${rows(16, 9, 32)}`,
};

const html = document.documentElement;
const prefersDark = matchMedia("(prefers-color-scheme: dark)");

/** The light/dark mode in effect: "auto" follows the system here, standing in for Outlook. */
function effectiveMode(settings: Settings): "light" | "dark" {
  if (settings.mode === "light" || settings.mode === "dark") return settings.mode;
  return prefersDark.matches ? "dark" : "light";
}

/** Dresses the page in the chosen palette (palettes.css keys off these attributes). */
function applyTheme(settings: Settings): void {
  html.dataset.overlookPalette = settings.palette;
  html.dataset.overlookMode = settings.mode;
  html.dataset.overlookOutlookScheme = prefersDark.matches ? "dark" : "light";
}

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  attrs: Record<string, string> = {},
  ...children: (Node | string)[]
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) node.setAttribute(k, v);
  node.append(...children);
  return node;
}

function section(title: string, ...children: Node[]): HTMLElement {
  return el("section", {}, el("h2", {}, title), ...children);
}

/** Mutually exclusive buttons for a choice; `render` draws each option. */
function choiceGroup(
  id: ChoiceId,
  settings: Settings,
  className: string,
  render: (value: string, label: string) => HTMLButtonElement,
): HTMLElement {
  const group = el("div", {
    class: className,
    role: "radiogroup",
    "aria-label": CHOICES[id].label,
  });
  for (const [value, label] of Object.entries(CHOICES[id].options)) {
    const button = render(value, label);
    button.type = "button";
    button.setAttribute("role", "radio");
    button.setAttribute("aria-checked", String(settings[id] === value));
    button.addEventListener("click", () => void saveSetting(id, value));
    group.append(button);
  }
  return group;
}

/** A tiny window painted in its own palette, in the mode currently in effect. */
function swatch(settings: Settings) {
  return (value: string, label: string): HTMLButtonElement => {
    const preview = el(
      "span",
      {
        class: "preview",
        "data-overlook-swatch": "",
        "data-overlook-palette": value,
        "data-overlook-mode": effectiveMode(settings),
      },
      el("span", { class: "bar" }),
      el("span", { class: "side" }),
      el(
        "span",
        { class: "body" },
        el("span", { class: "line short" }),
        el("span", { class: "line" }),
        el("span", { class: "line faint" }),
        el("span", { class: "line" }),
      ),
    );
    return el(
      "button",
      { class: "swatch", title: label },
      preview,
      SHORT_PALETTE_NAMES[value] ?? label,
    );
  };
}

function layoutCard(value: string, label: string): HTMLButtonElement {
  const button = el("button", { class: "layout", title: label });
  button.innerHTML = `<svg viewBox="0 0 64 44" aria-hidden="true">${LAYOUT_SKETCHES[value] ?? ""}</svg>`;
  button.append(SHORT_LAYOUT_NAMES[value] ?? label);
  return button;
}

function segment(value: string, label: string): HTMLButtonElement {
  return el("button", { "data-value": value }, label);
}

function toggle(id: FeatureId, settings: Settings): HTMLLabelElement {
  const input = el("input", { type: "checkbox", id: `toggle-${id}` });
  input.checked = settings[id];
  input.addEventListener("change", () => void saveSetting(id, input.checked));
  return el(
    "label",
    { class: "toggle", for: `toggle-${id}` },
    el("span", {}, FEATURES[id].label),
    input,
  );
}

function toggles(ids: FeatureId[], settings: Settings): HTMLElement {
  return el("div", { class: "toggles" }, ...ids.map((id) => toggle(id, settings)));
}

const featuresIn = (group: string) => FEATURE_IDS.filter((id) => FEATURES[id].group === group);

let declutterOpen = false;

function render(settings: Settings): void {
  applyTheme(settings);
  const main = document.getElementById("settings");
  if (!main) return;

  const layoutHint =
    settings.layout === "fastmail" || settings.layout === "focus"
      ? [
          el(
            "p",
            { class: "hint" },
            "Pair with Outlook's Settings → Mail → Layout → Reading pane: Hide.",
          ),
        ]
      : [];

  const declutter = featuresIn("Declutter");
  const hidden = declutter.filter((id) => settings[id]).length;
  const details = el(
    "details",
    {},
    el(
      "summary",
      {},
      "Declutter",
      el("span", { class: "count" }, `${hidden} of ${declutter.length} hidden`),
    ),
    toggles(declutter, settings),
  );
  details.open = declutterOpen;
  details.addEventListener("toggle", () => {
    declutterOpen = details.open;
  });

  main.replaceChildren(
    section("Colors", choiceGroup("palette", settings, "swatches", swatch(settings))),
    section("Light or dark", choiceGroup("mode", settings, "segmented", segment)),
    section("Layout", choiceGroup("layout", settings, "layouts", layoutCard), ...layoutHint),
    section("Keyboard", toggles(featuresIn("Keyboard"), settings)),
    el("section", {}, details),
    section("General", toggles(featuresIn("General"), settings)),
  );
}

async function main(): Promise<void> {
  const version = document.getElementById("version");
  if (version) version.textContent = `v${chrome.runtime.getManifest().version}`;
  let settings = await loadSettings();
  render(settings);
  onSettingsChanged((next) => {
    settings = next;
    render(settings);
  });
  prefersDark.addEventListener("change", () => render(settings));
}

void main();
