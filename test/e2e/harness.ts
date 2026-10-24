// Loads the built extension into Chromium and serves arbitrary HTML at an Outlook URL, so the
// content script runs exactly as it would on the real site. No network: every request is routed.

import path from "node:path";
import { type BrowserContext, chromium, type Page } from "playwright";

export const OUTLOOK_URL = "https://outlook.cloud.microsoft/mail/";

export async function launch(): Promise<BrowserContext> {
  const ext = path.resolve("dist");
  return chromium.launchPersistentContext("", {
    channel: "chromium",
    args: [`--disable-extensions-except=${ext}`, `--load-extension=${ext}`],
  });
}

export async function openPage(context: BrowserContext, html: string): Promise<Page> {
  const page = await context.newPage();
  await page.route("**/*", (route) =>
    route.request().url() === OUTLOOK_URL
      ? route.fulfill({ contentType: "text/html", body: html })
      : route.abort(),
  );
  await page.goto(OUTLOOK_URL);
  await page.waitForFunction(() => document.documentElement.hasAttribute("data-overlook-features"));
  return page;
}
