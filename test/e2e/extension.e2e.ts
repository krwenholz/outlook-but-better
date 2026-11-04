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
    await fixture.waitForFunction(() => document.documentElement.dataset.overlookHealth, null, {
      timeout: 20_000,
    });
    expect(await fixture.evaluate(() => document.documentElement.dataset.overlookHealth)).toBe(
      "ok",
    );
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

  it("recolours Outlook through its own palette variables", async () => {
    const folderPaneBackground = () =>
      page.evaluate(() => {
        const pane = document.querySelector('[data-overlook~="folder-pane"]');
        return pane ? getComputedStyle(pane).backgroundColor : null;
      });
    await page.evaluate(() => {
      document.documentElement.dataset.overlookPalette = "solarized";
      document.documentElement.dataset.overlookMode = "dark";
    });
    await expect.poll(folderPaneBackground).toBe("rgb(7, 54, 66)"); // Solarized base02
    await page.evaluate(() => {
      document.documentElement.dataset.overlookMode = "auto";
      document.documentElement.dataset.overlookOutlookScheme = "light";
    });
    await expect.poll(folderPaneBackground).toBe("rgb(238, 232, 213)"); // Solarized base2
  });

  it("archives by clicking Outlook's own Archive button", async () => {
    await page.keyboard.press("y");
    expect(await actions(page)).toEqual(["click:Archive"]);
  });

  it("marks an unread message read with Outlook's Read / Unread button", async () => {
    await page.keyboard.press(".");
    await page.keyboard.press("r");
    expect(await actions(page)).toEqual(["click:Read / Unread"]);
  });

  it("leaves an already-unread message alone on mark unread", async () => {
    await page.keyboard.press(".");
    await page.keyboard.press("u");
    expect(await actions(page)).toEqual([]);
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

  it("switches to Calendar through the palette", async () => {
    await page.keyboard.press("g");
    await page.keyboard.type("cal");
    await page.keyboard.press("Enter");
    expect(await actions(page)).toEqual(["click:Calendar"]);
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
