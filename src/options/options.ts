import {
  CHOICE_IDS,
  CHOICES,
  FEATURE_IDS,
  FEATURES,
  loadSettings,
  saveSetting,
} from "../settings.ts";

function section(form: HTMLElement, title: string): void {
  const heading = document.createElement("h2");
  heading.textContent = title;
  form.append(heading);
}

async function render(): Promise<void> {
  const form = document.getElementById("settings");
  if (!form) return;
  const settings = await loadSettings();
  const groups = Map.groupBy(FEATURE_IDS, (id): string => FEATURES[id].group);
  const choiceGroups = Map.groupBy(CHOICE_IDS, (id): string => CHOICES[id].group);
  for (const group of new Set([...groups.keys(), ...choiceGroups.keys()])) {
    section(form, group);
    for (const id of choiceGroups.get(group) ?? []) {
      const label = document.createElement("label");
      const select = document.createElement("select");
      for (const [value, text] of Object.entries(CHOICES[id].options)) {
        select.append(new Option(text, value, false, settings[id] === value));
      }
      select.addEventListener("change", () => void saveSetting(id, select.value));
      label.append(`${CHOICES[id].label}: `, select);
      form.append(label);
    }
    for (const id of groups.get(group) ?? []) {
      const label = document.createElement("label");
      const box = document.createElement("input");
      box.type = "checkbox";
      box.checked = settings[id];
      box.addEventListener("change", () => void saveSetting(id, box.checked));
      label.append(box, FEATURES[id].label);
      form.append(label);
    }
  }
}

void render();
