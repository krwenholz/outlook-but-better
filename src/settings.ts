// User preferences. Stored in chrome.storage.sync; nothing else is ever persisted.
//
// Each enabled feature id is written as a token on <html data-overlook-features="...">, and each
// choice as its own attribute (<html data-overlook-palette="solarized">), so stylesheets gate
// themselves with `[data-overlook-features~="density"]` or `[data-overlook-palette="solarized"]`.

export type FeatureId = keyof typeof FEATURES;
export type ChoiceId = keyof typeof CHOICES;
export type Settings = Record<FeatureId, boolean> & { [K in ChoiceId]: ChoiceValue<K> };
export type ChoiceValue<K extends ChoiceId> = keyof (typeof CHOICES)[K]["options"];

type Group = "General" | "Keyboard" | "Declutter" | "Look";

interface Feature {
  label: string;
  group: Group;
  default: boolean;
}

interface Choice {
  label: string;
  group: Group;
  options: Record<string, string>;
  default: string;
}

export const FEATURES = {
  healthIndicator: {
    label: "Show an indicator when Outlook's page changes break Overlook",
    group: "General",
    default: true,
  },
  keymap: { label: "Fastmail-style keyboard shortcuts", group: "Keyboard", default: true },
  "hide-upsell": { label: "Upgrade and 'try the new…' banners", group: "Declutter", default: true },
  "hide-copilot": { label: "Copilot buttons and panels", group: "Declutter", default: true },
  "hide-header-extras": {
    label: "Teams and My Day buttons in the header",
    group: "Declutter",
    default: true,
  },
  "hide-rail-extras": {
    label: "To Do, Newsletters, Org Explorer and OneDrive in the app rail",
    group: "Declutter",
    default: true,
  },
  "hide-ribbon-extras": {
    label: "Teams, Sweep, Quick steps, Viva and add-in buttons in the ribbon",
    group: "Declutter",
    default: true,
  },
  "hide-external-tag": { label: "'External' sender tags", group: "Declutter", default: true },
  "hide-row-actions": {
    label: "Flag and pin icons that appear when hovering a message",
    group: "Declutter",
    default: false,
  },
  "hide-avatars": {
    label: "Sender pictures in the message list",
    group: "Declutter",
    default: false,
  },
  "hide-group-headers": {
    label: "Date group headers ('This Month') in the message list",
    group: "Declutter",
    default: false,
  },
  "hide-focused-tabs": { label: "Focused / Other tabs", group: "Declutter", default: false },
  "hide-ribbon-tabs": {
    label: "Ribbon tabs (Home, View, Help)",
    group: "Declutter",
    default: false,
  },
  "hide-app-rail": { label: "App rail (left icon bar)", group: "Declutter", default: false },
  "hide-ribbon": { label: "Ribbon", group: "Declutter", default: false },
  density: { label: "Tighter, denser layout", group: "Look", default: true },
} as const satisfies Record<string, Feature>;

export const CHOICES = {
  palette: {
    label: "Colors",
    group: "Look",
    options: {
      outlook: "Outlook's own",
      solarized: "Solarized",
      base16: "Base16 Default",
      fastmail: "Fastmail",
      paper: "Paper",
    },
    default: "outlook",
  },
  mode: {
    label: "Light or dark",
    group: "Look",
    options: { auto: "Follow Outlook", light: "Light", dark: "Dark" },
    default: "auto",
  },
  layout: {
    label: "Layout",
    group: "Look",
    options: {
      outlook: "Outlook's own",
      calm: "Calm three-pane",
      fastmail: "Fastmail (list first)",
      focus: "Focus (single column)",
    },
    default: "outlook",
  },
} as const satisfies Record<string, Choice>;

export const FEATURE_IDS = Object.keys(FEATURES) as FeatureId[];
export const CHOICE_IDS = Object.keys(CHOICES) as ChoiceId[];

export function defaultSettings(): Settings {
  return {
    ...Object.fromEntries(FEATURE_IDS.map((id) => [id, FEATURES[id].default])),
    ...Object.fromEntries(CHOICE_IDS.map((id) => [id, CHOICES[id].default])),
  } as Settings;
}

/** Merge whatever is stored over the defaults, ignoring unknown or mistyped keys. */
export function normalize(stored: Record<string, unknown>): Settings {
  const settings = defaultSettings() as Record<string, unknown>;
  for (const id of FEATURE_IDS) {
    if (typeof stored[id] === "boolean") settings[id] = stored[id];
  }
  for (const id of CHOICE_IDS) {
    const value = stored[id];
    if (typeof value === "string" && value in CHOICES[id].options) settings[id] = value;
  }
  return settings as Settings;
}

export function featureTokens(settings: Settings): string {
  return FEATURE_IDS.filter((id) => settings[id]).join(" ");
}

/** Writes the settings onto <html> as the attributes the stylesheets key off. */
export function applyToDocument(settings: Settings, root: HTMLElement): void {
  root.setAttribute("data-overlook-features", featureTokens(settings));
  for (const id of CHOICE_IDS) root.setAttribute(`data-overlook-${id}`, settings[id]);
}

export async function loadSettings(): Promise<Settings> {
  return normalize(await chrome.storage.sync.get(null));
}

export async function saveSetting(
  id: FeatureId | ChoiceId,
  value: boolean | string,
): Promise<void> {
  await chrome.storage.sync.set({ [id]: value });
}

export function onSettingsChanged(callback: (settings: Settings) => void): void {
  chrome.storage.onChanged.addListener((_changes, area) => {
    if (area === "sync") void loadSettings().then(callback);
  });
}
