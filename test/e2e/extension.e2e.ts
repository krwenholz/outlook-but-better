import { readdir, readFile } from "node:fs/promises";
import type { BrowserContext, Page } from "playwright";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { launch, openPage } from "./harness.ts";

let context: BrowserContext;
let page: Page;

beforeAll(async () => {
  context = await launch();
});
afterAll(async () => {
  await context?.close();
});

describe("extension on a fake Outlook page", () => {
  beforeEach(async () => {
    page = await openPage(context, await readFile("test/e2e/fake-outlook.html", "utf8"));
  });

  it("applies default settings and hides decluttered elements", async () => {
    const features = await page.evaluate(() => document.documentElement.dataset.overlookFeatures);
    expect(features).toContain("hide-copilot");
    await expect.poll(() => page.locator('[aria-label="Copilot"]').isVisible()).toBe(false);
  });

  it("translates y into Outlook's archive key as a synthetic event", async () => {
    await page.keyboard.press("y");
    expect(
      await page.evaluate(() => (window as unknown as { received: string[] }).received),
    ).toEqual(["e:false"]);
  });

  it("leaves typing in inputs alone", async () => {
    await page.locator("#topSearchInput").fill("");
    await page.locator("#topSearchInput").press("y");
    expect(await page.locator("#topSearchInput").inputValue()).toBe("y");
  });

  it("keeps palette keystrokes away from the page and from the keymap", async () => {
    await page.keyboard.press("g");
    await page.keyboard.type("yj");
    await page.keyboard.press("Escape");
    expect(
      await page.evaluate(() => (window as unknown as { received: string[] }).received),
    ).toEqual([]);
  });

  it("swallows shortcuts while the help overlay is open", async () => {
    await page.keyboard.press("?");
    await page.keyboard.press("y");
    await page.keyboard.press("y");
    expect(
      await page.evaluate(() => (window as unknown as { received: string[] }).received),
    ).toEqual(["e:false"]);
  });

  it("goes to a nested folder through the palette", async () => {
    await page.keyboard.press("g");
    await page.keyboard.type("rec");
    await page.keyboard.press("Enter");
    await expect.poll(() => page.title()).toBe("folder:Receipts");
  });
});

// Scrubbed real pages (see test/fixtures/README.md). Each must pass the health check.
const fixtures = (await readdir("test/fixtures")).filter((f) => f.endsWith(".html"));
describe.skipIf(fixtures.length === 0)("scrubbed Outlook fixtures", () => {
  it.each(fixtures)("%s has no broken selectors", async (name) => {
    const fixture = await openPage(context, await readFile(`test/fixtures/${name}`, "utf8"));
    await fixture.waitForTimeout(500);
    const badge = await fixture.evaluate(
      () =>
        document.querySelector("overlook-root")?.shadowRoot?.querySelector(".health:not([hidden])")
          ?.textContent ?? null,
    );
    expect(badge).toBeNull();
  });
});
