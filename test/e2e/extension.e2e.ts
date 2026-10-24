import { readdir } from "node:fs/promises";
import type { BrowserContext, Page } from "playwright";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { actions, launch, openFixture, recordActions } from "./harness.ts";

let context: BrowserContext;
let page: Page;

beforeAll(async () => {
  context = await launch();
});
afterAll(async () => {
  await context?.close();
});

const badge = (p: Page) =>
  p.evaluate(
    () =>
      document.querySelector("overlook-root")?.shadowRoot?.querySelector(".health:not([hidden])")
        ?.textContent ?? null,
  );

// Every scrubbed capture must pass the health check.
const fixtures = (await readdir("test/fixtures")).filter((f) => f.endsWith(".html"));
describe("scrubbed Outlook fixtures", () => {
  it.each(fixtures)("%s has no broken selectors", async (file) => {
    const fixture = await openFixture(context, file.replace(/\.html$/, ""));
    await fixture.waitForTimeout(300);
    expect(await badge(fixture)).toBeNull();
    await fixture.close();
  });
});

describe("extension on the inbox fixture", () => {
  beforeEach(async () => {
    page = await openFixture(context, "inbox");
    await recordActions(page);
  });

  it("hides decluttered elements by default", async () => {
    await expect.poll(() => page.locator("#owaTeamsBtn_container").isVisible()).toBe(false);
    await expect
      .poll(() => page.locator('[data-app-section="UpsellBannerSection"]').count())
      .toBe(1);
    expect(await page.locator('[data-overlook~="upsell"]').isVisible()).toBe(false);
  });

  it("archives by clicking Outlook's own Archive button", async () => {
    await page.keyboard.press("y");
    expect(await actions(page)).toEqual(["click:Archive"]);
  });

  it("marks read with Outlook's Shift+I shortcut", async () => {
    await page.keyboard.press(".");
    await page.keyboard.press("r");
    expect(await actions(page)).toEqual(["key:shift+I"]);
  });

  it("leaves typing in the search box alone", async () => {
    await page.locator("#topSearchInput").fill("");
    await page.locator("#topSearchInput").press("y");
    expect(await page.locator("#topSearchInput").inputValue()).toBe("y");
    expect(await actions(page)).toEqual([]);
  });

  it("goes to a folder through the palette", async () => {
    await page.keyboard.press("g");
    await page.keyboard.type("arch");
    await page.keyboard.press("Enter");
    expect(await actions(page)).toEqual(["click:archive"]);
  });

  it("keeps palette keystrokes away from the page and the keymap", async () => {
    await page.keyboard.press("g");
    await page.keyboard.type("yj");
    await page.keyboard.press("Escape");
    expect(await actions(page)).toEqual([]);
  });

  it("swallows shortcuts while the help overlay is open", async () => {
    await page.keyboard.press("?");
    await page.keyboard.press("y");
    await page.keyboard.press("y");
    expect(await actions(page)).toEqual(["click:Archive"]);
  });
});
