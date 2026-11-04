// Renders fixtures with the built extension under each look, into docs/screenshots/.
// Usage: mise run screenshots [name…]   (builds and fetches fixture assets first)
// Fonts and icons load from Outlook's public CDN, so this needs network access.

import { mkdir } from "node:fs/promises";
import { CHOICES } from "../src/settings.ts";
import { launch, openFixture, openOptions } from "../test/e2e/harness.ts";

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

/** "Solarized · Dark · Calm", for the frames of docs/preview.webp. */
function caption(l: Look): string {
  if (l.palette === "outlook") return "Stock Outlook";
  const palette = CHOICES.palette.options[l.palette as keyof typeof CHOICES.palette.options];
  const mode = l.mode === "dark" ? "Dark" : "Light";
  const layout = l.layout[0]?.toUpperCase() + l.layout.slice(1);
  return `${palette} · ${mode} · ${layout}`;
}

const only = process.argv.slice(2);
const out = "docs/screenshots";
// Captioned copies for scripts/preview.ts; regenerated, so not committed.
const frames = ".cache/preview-frames";
await mkdir(out, { recursive: true });
await mkdir(frames, { recursive: true });
const context = await launch();
const looks = LOOKS.filter((l) => only.length === 0 || only.includes(l.name));
for (const [i, l] of looks.entries()) {
  const page = await openFixture(context, l.fixture, { cdn: true });
  await page.evaluate((l) => {
    const html = document.documentElement;
    html.dataset.overlookPalette = l.palette;
    html.dataset.overlookMode = l.mode;
    html.dataset.overlookLayout = l.layout;
    if (l.name === "outlook-stock") html.removeAttribute("data-overlook-features");
    else html.dataset.overlookFeatures += " hide-app-banners";
    if (l.noReadingPane) {
      // Stand in for Outlook's "Reading pane: Hide": the pane goes away and Outlook stretches the
      // list into its space. (The extension itself never sets pane widths.)
      document.querySelector('[data-app-section="MailReadCompose"]')?.remove();
      const style = document.createElement("style");
      style.textContent = `
        [data-app-section="MessageList"],
        [data-app-section="MessageList"] [role="complementary"],
        #MailList { flex: 1 1 auto !important; }
        [data-app-section="MessageList"],
        [data-app-section="MessageList"] [role="complementary"],
        [data-app-section="MessageList"] [role="complementary"] > div,
        #MailList { width: auto !important; max-width: none !important; }`;
      document.head.append(style);
    }
  }, l);
  await page.waitForLoadState("networkidle");
  await page.waitForTimeout(300);
  await page.screenshot({ path: `${out}/${l.name}.png` });
  await page.evaluate((text) => {
    const pill = document.createElement("div");
    pill.textContent = text;
    pill.style.cssText = `position:fixed;left:24px;bottom:24px;z-index:2147483647;
      padding:10px 18px;border-radius:999px;font:600 22px/1 -apple-system,"Segoe UI",sans-serif;
      color:#fff;background:rgb(17 24 39 / 82%);box-shadow:0 8px 24px rgb(0 0 0 / 30%);
      letter-spacing:0.01em`;
    document.body.append(pill);
  }, caption(l));
  await page.screenshot({ path: `${frames}/${String(i).padStart(2, "0")}-${l.name}.png` });
  await page.close();
  console.log(`${out}/${l.name}.png`);
}

// The toolbar popup, in two palettes.
if (only.length === 0) {
  for (const [name, settings] of [
    ["popup-light", { palette: "fastmail", mode: "light", layout: "fastmail" }],
    ["popup-dark", { palette: "solarized", mode: "dark", layout: "calm" }],
  ] as const) {
    const page = await openOptions(context, settings);
    await page.setViewportSize({ width: 400, height: 600 });
    await page.screenshot({ path: `${out}/${name}.png`, fullPage: true });
    await page.close();
    console.log(`${out}/${name}.png`);
  }
}
await context.close();
