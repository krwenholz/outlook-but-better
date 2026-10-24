// Renders the inbox fixture with the built extension under each look, into docs/screenshots/.
// Usage: mise run screenshots   (builds and fetches fixture assets first)

import { mkdir } from "node:fs/promises";
import { launch, openFixture } from "../test/e2e/harness.ts";

interface Look {
  name: string;
  palette: string;
  mode: string;
  layout: string;
  /** Extra feature tokens on top of the defaults. */
  features?: string[];
  /** Stand in for Outlook's "Reading pane: Hide" setting, which list-first layouts pair with. */
  noReadingPane?: boolean;
}

export const LOOKS: Look[] = [
  { name: "outlook-stock", palette: "outlook", mode: "auto", layout: "outlook" },
  { name: "solarized-light-calm", palette: "solarized", mode: "light", layout: "calm" },
  { name: "solarized-dark-calm", palette: "solarized", mode: "dark", layout: "calm" },
  {
    name: "base16-dark-fastmail",
    palette: "base16",
    mode: "dark",
    layout: "fastmail",
    noReadingPane: true,
  },
  {
    name: "base16-light-fastmail",
    palette: "base16",
    mode: "light",
    layout: "fastmail",
    noReadingPane: true,
  },
  {
    name: "fastmail-light-fastmail",
    palette: "fastmail",
    mode: "light",
    layout: "fastmail",
    noReadingPane: true,
  },
  { name: "fastmail-dark-calm", palette: "fastmail", mode: "dark", layout: "calm" },
  {
    name: "paper-light-focus",
    palette: "paper",
    mode: "light",
    layout: "focus",
    noReadingPane: true,
  },
  {
    name: "solarized-dark-focus",
    palette: "solarized",
    mode: "dark",
    layout: "focus",
    noReadingPane: true,
  },
];

const only = process.argv.slice(2);
const out = "docs/screenshots";
await mkdir(out, { recursive: true });
const context = await launch();
for (const look of LOOKS.filter((l) => only.length === 0 || only.includes(l.name))) {
  const page = await openFixture(context, "inbox");
  await page.evaluate((l) => {
    const html = document.documentElement;
    html.dataset.overlookPalette = l.palette;
    html.dataset.overlookMode = l.mode;
    html.dataset.overlookLayout = l.layout;
    if (l.name === "outlook-stock") html.removeAttribute("data-overlook-features");
    else html.dataset.overlookFeatures += ` hide-app-banners ${(l.features ?? []).join(" ")}`;
    if (l.noReadingPane) {
      const style = document.createElement("style");
      style.textContent = '[data-overlook~="reading-pane"] { display: none !important; }';
      document.head.append(style);
    }
  }, look);
  await page.waitForTimeout(400);
  await page.screenshot({ path: `${out}/${look.name}.png` });
  await page.close();
  console.log(`${out}/${look.name}.png`);
}
await context.close();
