// Renders fixtures with the built extension under each look, into docs/screenshots/.
// Usage: mise run screenshots [name…]   (builds and fetches fixture assets first)
// Fonts and icons load from Outlook's public CDN, so this needs network access.

import { mkdir } from "node:fs/promises";
import { launch, openFixture } from "../test/e2e/harness.ts";

interface Look {
  name: string;
  /** test/fixtures/<fixture>.html */
  fixture: "list" | "move-menu";
  palette: string;
  mode: string;
  layout: string;
  /** Stand in for Outlook's "Reading pane: Hide" setting, which list-first layouts pair with. */
  noReadingPane?: boolean;
}

const look = (
  name: string,
  fixture: Look["fixture"],
  palette: string,
  mode: string,
  layout: string,
): Look => ({
  name,
  fixture,
  palette,
  mode,
  layout,
  noReadingPane: layout === "fastmail" || layout === "focus",
});

export const LOOKS: Look[] = [
  look("outlook-stock", "move-menu", "outlook", "auto", "outlook"),
  look("solarized-light-calm", "move-menu", "solarized", "light", "calm"),
  look("solarized-dark-calm", "move-menu", "solarized", "dark", "calm"),
  look("fastmail-dark-calm", "move-menu", "fastmail", "dark", "calm"),
  look("fastmail-light-fastmail", "list", "fastmail", "light", "fastmail"),
  look("base16-dark-fastmail", "list", "base16", "dark", "fastmail"),
  look("base16-light-fastmail", "list", "base16", "light", "fastmail"),
  look("paper-light-focus", "list", "paper", "light", "focus"),
  look("solarized-dark-focus", "list", "solarized", "dark", "focus"),
];

const only = process.argv.slice(2);
const out = "docs/screenshots";
await mkdir(out, { recursive: true });
const context = await launch();
for (const l of LOOKS.filter((l) => only.length === 0 || only.includes(l.name))) {
  const page = await openFixture(context, l.fixture, { cdn: true });
  await page.evaluate((l) => {
    const html = document.documentElement;
    html.dataset.overlookPalette = l.palette;
    html.dataset.overlookMode = l.mode;
    html.dataset.overlookLayout = l.layout;
    if (l.name === "outlook-stock") html.removeAttribute("data-overlook-features");
    else html.dataset.overlookFeatures += " hide-app-banners";
    if (l.noReadingPane) {
      const style = document.createElement("style");
      style.textContent = '[data-overlook~="reading-pane"] { display: none !important; }';
      document.head.append(style);
    }
  }, l);
  await page.waitForLoadState("networkidle");
  await page.waitForTimeout(300);
  await page.screenshot({ path: `${out}/${l.name}.png` });
  await page.close();
  console.log(`${out}/${l.name}.png`);
}
await context.close();
