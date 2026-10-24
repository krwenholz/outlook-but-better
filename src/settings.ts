// User preferences. Stored in chrome.storage.sync; nothing else is ever persisted.
//
// Each enabled feature id is written as a token on <html data-overlook-features="...">,
// so stylesheets gate themselves with `[data-overlook-features~="density"]`.

export type FeatureId = keyof typeof FEATURES;
export type Settings = Record<FeatureId, boolean>;

interface Feature {
  label: string;
  group: "General" | "Keyboard" | "Declutter" | "Look";
  default: boolean;
}

export const FEATURES = {
  healthIndicator: {
    label: "Show an indicator when Outlook's page changes break Overlook",
    group: "General",
    default: true,
  },
  keymap: { label: "Fastmail-style keyboard shortcuts", group: "Keyboard", default: true },
  "hide-copilot": { label: "Copilot buttons and panels", group: "Declutter", default: true },
  "hide-promo-banner": {
    label: "Promo and 'try the new…' banners",
    group: "Declutter",
    default: true,
  },
  "hide-external-tag": { label: "'External' sender tags", group: "Declutter", default: true },
  "hide-app-rail": { label: "App rail (left icon bar)", group: "Declutter", default: false },
  "hide-ribbon": { label: "Ribbon", group: "Declutter", default: false },
  density: { label: "Tighter, denser layout", group: "Look", default: true },
  theme: { label: "Calmer theme (work in progress)", group: "Look", default: false },
} as const satisfies Record<string, Feature>;

export const FEATURE_IDS = Object.keys(FEATURES) as FeatureId[];

export function defaultSettings(): Settings {
  return Object.fromEntries(FEATURE_IDS.map((id) => [id, FEATURES[id].default])) as Settings;
}

/** Merge whatever is stored over the defaults, ignoring unknown or mistyped keys. */
export function normalize(stored: Record<string, unknown>): Settings {
  const settings = defaultSettings();
  for (const id of FEATURE_IDS) {
    const value = stored[id];
    if (typeof value === "boolean") settings[id] = value;
  }
  return settings;
}

export function featureTokens(settings: Settings): string {
  return FEATURE_IDS.filter((id) => settings[id]).join(" ");
}

export async function loadSettings(): Promise<Settings> {
  return normalize(await chrome.storage.sync.get(null));
}

export async function saveSetting(id: FeatureId, value: boolean): Promise<void> {
  await chrome.storage.sync.set({ [id]: value });
}

export function onSettingsChanged(callback: (settings: Settings) => void): void {
  chrome.storage.onChanged.addListener((_changes, area) => {
    if (area === "sync") void loadSettings().then(callback);
  });
}
