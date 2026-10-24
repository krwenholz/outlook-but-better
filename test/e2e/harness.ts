// Loads the built extension into Chromium and serves a fixture at an Outlook URL, so the content
// script runs exactly as it would on the real site. No network: every other request is aborted.

import { readFile } from "node:fs/promises";
import path from "node:path";
import { type BrowserContext, chromium, type Page } from "playwright";

export const OUTLOOK_URL = "https://outlook.cloud.microsoft/mail/";
const FIXTURES = path.resolve("test/fixtures");

export async function launch(): Promise<BrowserContext> {
  const ext = path.resolve("dist");
  return chromium.launchPersistentContext("", {
    channel: "chromium",
    viewport: { width: 1600, height: 1000 },
    args: [`--disable-extensions-except=${ext}`, `--load-extension=${ext}`],
  });
}

/** Opens test/fixtures/<name>.html at the Outlook URL, with its assets/ alongside. */
export async function openFixture(context: BrowserContext, name: string): Promise<Page> {
  const page = await context.newPage();
  await page.route("**/*", async (route) => {
    const url = route.request().url();
    if (url === OUTLOOK_URL) {
      return route.fulfill({
        contentType: "text/html",
        body: await readFile(path.join(FIXTURES, `${name}.html`)),
      });
    }
    if (url.startsWith(`${OUTLOOK_URL}assets/`)) {
      const file = path.join(FIXTURES, "assets", path.basename(new URL(url).pathname));
      return route.fulfill({ path: file });
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
