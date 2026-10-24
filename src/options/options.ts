import { FEATURE_IDS, FEATURES, loadSettings, saveSetting } from "../settings.ts";

async function render(): Promise<void> {
  const form = document.getElementById("settings");
  if (!form) return;
  const settings = await loadSettings();
  const groups = Map.groupBy(FEATURE_IDS, (id) => FEATURES[id].group);
  for (const [group, ids] of groups) {
    const heading = document.createElement("h2");
    heading.textContent = group;
    form.append(heading);
    for (const id of ids) {
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
