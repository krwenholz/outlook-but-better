// Loads the built extension into Chromium and serves a fixture at an Outlook URL, so the content
// script runs exactly as it would on the real site. No network by default: every other request is
// aborted. Also used by scripts/screenshots.ts, which lets Outlook's public CDN through for fonts.

import { existsSync } from "node:fs";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { type BrowserContext, chromium, type Page } from "playwright";

export const OUTLOOK_URL = "https://outlook.cloud.microsoft/mail/";
const OUTLOOK_CDN = "https://res.public.onecdn.static.microsoft/";
const FIXTURES = path.resolve("test/fixtures");

export async function launch(): Promise<BrowserContext> {
  const ext = path.resolve("dist");
  return chromium.launchPersistentContext("", {
    channel: "chromium",
    viewport: { width: 1600, height: 1000 },
    args: [`--disable-extensions-except=${ext}`, `--load-extension=${ext}`],
  });
}

export interface FixtureOptions {
  /** Let requests to Outlook's public CDN (fonts, icon fonts) through. Screenshots only. */
  cdn?: boolean;
}

/** Opens test/fixtures/<name>.html at the Outlook URL, with its assets/ and vendor/ alongside. */
export async function openFixture(
  context: BrowserContext,
  name: string,
  options: FixtureOptions = {},
): Promise<Page> {
  const page = await context.newPage();
  await page.route("**/*", async (route) => {
    const url = route.request().url();
    if (url === OUTLOOK_URL) {
      const html = await readFile(path.join(FIXTURES, `${name}.html`), "utf8");
      // Captures made with the DevTools snippet carry their runtime styles in vendor/; only
      // "Save page" captures need the heavier approximations.
      const shims = html.includes('href="vendor/') ? ["shim.css"] : ["shim.css", "shim-saved.css"];
      const links = shims.map((s) => `<link rel="stylesheet" href="${s}">`).join("");
      return route.fulfill({
        contentType: "text/html; charset=utf-8",
        body: html.replace("</head>", `${links}</head>`),
      });
    }
    if (url === `${OUTLOOK_URL}shim-saved.css`) {
      return route.fulfill({ path: path.join(FIXTURES, "shim-saved.css") });
    }
    if (url.startsWith(`${OUTLOOK_URL}vendor/`)) {
      const file = path.join(FIXTURES, "vendor", path.basename(new URL(url).pathname));
      return existsSync(file) ? route.fulfill({ path: file }) : route.fulfill({ status: 404 });
    }
    if (options.cdn && url.startsWith(OUTLOOK_CDN)) return route.continue();
    if (url === `${OUTLOOK_URL}shim.css`)
      return route.fulfill({ path: path.join(FIXTURES, "shim.css") });
    if (url.startsWith(`${OUTLOOK_URL}assets/`)) {
      // Fetched by `mise run fixtures:assets`; absent in CI, where pages just render unstyled.
      const file = path.join(FIXTURES, "assets", path.basename(new URL(url).pathname));
      return existsSync(file) ? route.fulfill({ path: file }) : route.fulfill({ status: 404 });
    }
    return route.abort();
  });
  await page.goto(OUTLOOK_URL);
  await page.waitForFunction(() => document.documentElement.hasAttribute("data-overlook-features"));
  return page;
}

/** Records clicks and synthetic key events on the page, so tests can see what Overlook did. */
export async function recordActions(page: Page): Promise<void> {
  await page.evaluate(() => {
    const log: string[] = [];
    (window as unknown as { overlookLog: string[] }).overlookLog = log;
    document.addEventListener(
      "click",
      (e) => {
        const el = e.target as Element;
        log.push(`click:${el.getAttribute("aria-label") ?? el.getAttribute("data-folder-name")}`);
        e.preventDefault();
      },
      true,
    );
    document.addEventListener("keydown", (e) => {
      if (!e.isTrusted) log.push(`key:${e.shiftKey ? "shift+" : ""}${e.key}`);
    });
  });
}

export async function actions(page: Page): Promise<string[]> {
  return page.evaluate(() => (window as unknown as { overlookLog: string[] }).overlookLog);
}
