import { readdir, readFile } from "node:fs/promises";
import type { BrowserContext, Page } from "playwright";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { CAPTURE_MESSAGE } from "../../src/capture.ts";
import { actions, launch, openFixture, openOptions, recordActions } from "./harness.ts";

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

  it("leaves pane widths to Outlook while a reading pane is showing", async () => {
    const listWidth = (layout: string) =>
      page.evaluate((l) => {
        document.documentElement.dataset.overlookLayout = l;
        const pane = document.querySelector('[data-overlook~="list-pane"]');
        return pane?.getBoundingClientRect().width;
      }, layout);
    await expect
      .poll(() => page.evaluate(() => document.documentElement.dataset.overlookReadingPane))
      .toBe("shown");
    const stock = await listWidth("outlook");
    for (const layout of ["calm", "fastmail", "focus"]) expect(await listWidth(layout)).toBe(stock);
  });

  it("keeps the date group headers visible: j/k steps through them", async () => {
    const hiddenHeaders = (layout: string) =>
      page.evaluate((l) => {
        document.documentElement.dataset.overlookLayout = l;
        document.documentElement.dataset.overlookPalette = "solarized";
        const headers = [...document.querySelectorAll('[data-overlook~="group-header"]')];
        return {
          count: headers.length,
          hidden: headers.filter((h) => !(h as HTMLElement).offsetHeight).length,
        };
      }, layout);
    for (const layout of ["outlook", "calm", "fastmail", "focus"]) {
      const { count, hidden } = await hiddenHeaders(layout);
      expect(count, layout).toBeGreaterThan(0);
      expect(hidden, `${layout}: hidden headers`).toBe(0);
    }
  });

  it("lets Outlook swap each avatar for its tick box", async () => {
    // Outlook shows the tick box instead of the avatar on a selected row (and on hover). A palette
    // must not force the avatar back, or it sits on top of the box and multi-select breaks.
    const avatarDisplays = () =>
      page.evaluate(() => {
        document.documentElement.dataset.overlookPalette = "paper";
        const display = (selected: string) => {
          const row = document.querySelector(`[role="option"][aria-selected="${selected}"]`);
          const avatar = row?.querySelector('[data-overlook~="row-avatar"]');
          return avatar ? getComputedStyle(avatar).display : "missing";
        };
        return { selected: display("true"), other: display("false") };
      });
    await expect.poll(async () => (await avatarDisplays()).other).not.toBe("missing");
    expect(await avatarDisplays()).toEqual({
      selected: "none",
      other: expect.not.stringMatching(/^(none|missing)$/),
    });
  });

  it("archives by clicking Outlook's own Archive button", async () => {
    await page.keyboard.press("y");
    expect(await actions(page)).toEqual(["click:Archive"]);
  });

  it("x ticks the row; j and k then move the cursor without dropping ticks", async () => {
    const cursorIndex = () =>
      page.evaluate(() => {
        const rows = [...document.querySelectorAll('[role="option"][data-convid]')];
        return rows.findIndex((row) => row.contains(document.activeElement));
      });
    const ticked = () =>
      page.evaluate(() => {
        const rows = [...document.querySelectorAll('[role="option"][data-convid]')];
        return rows.flatMap((row, i) =>
          row.querySelector('[role="checkbox"]')?.getAttribute("aria-checked") === "true"
            ? [i]
            : [],
        );
      });
    // The fixture has no Outlook scripts, so play Outlook's part: a click on a tick box ticks it.
    await page.evaluate(() => {
      document.addEventListener("click", (e) => {
        const box = (e.target as Element).closest('[role="checkbox"]');
        if (box)
          box.setAttribute("aria-checked", String(box.getAttribute("aria-checked") !== "true"));
      });
    });

    await page.keyboard.press("x"); // the cursor starts on the selected (first) row
    expect(await ticked()).toEqual([0]);
    await page.keyboard.press("j");
    expect(await cursorIndex()).toBe(1);
    await page.keyboard.press("x");
    expect(await ticked()).toEqual([0, 1]);
    await page.keyboard.press("j");
    await page.keyboard.press("k");
    expect(await cursorIndex()).toBe(1);
    expect(await ticked()).toEqual([0, 1]); // moving never touched the ticks
    expect((await actions(page)).every((a) => a === "click:Select a conversation")).toBe(true);
  });

  it("leaves j/k to Outlook while nothing is ticked", async () => {
    await page.keyboard.press("j");
    const moved = await page.evaluate(() =>
      [...document.querySelectorAll('[role="option"]')].some((r) =>
        r.contains(document.activeElement),
      ),
    );
    expect(moved).toBe(false); // OBB did nothing; in Outlook its own j would take it from here
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

  it("saves a stock-Outlook capture when the popup asks", async () => {
    // The popup only asks the active tab; earlier tests left theirs open, and Chrome blocks a
    // burst of automatic downloads from one origin, so leave just this one.
    for (const other of context.pages()) if (other !== page) await other.close();
    const options = await openOptions(context);
    const download = page.waitForEvent("download");
    // Stands in for the popup's button: asks every tab; only the Outlook one answers.
    await options.evaluate(async (type) => {
      for (const tab of await chrome.tabs.query({})) {
        if (tab.id === undefined) continue;
        try {
          await chrome.tabs.sendMessage(tab.id, { type });
        } catch {}
      }
    }, CAPTURE_MESSAGE);
    const file = await download;
    expect(file.suggestedFilename()).toMatch(/^obb-capture-\d{12}\.html$/);
    const html = await readFile((await file.path()) ?? "", "utf8");
    expect(html).toContain('id="obb-capture-meta"');
    expect(html).toContain('id="MailList"');
    expect(html).not.toContain("overlook-root");
    expect(html).not.toMatch(/data-overlook(?!-runtime)/);
    await options.close();
  });

  it("swallows shortcuts while the help overlay is open", async () => {
    await page.keyboard.press("?");
    await page.keyboard.press("y");
    await page.keyboard.press("y");
    expect(await actions(page)).toEqual(["click:Archive"]);
  });
});

// Reloads the extension, so it gets a browser of its own.
describe("after OBB is reloaded", () => {
  it("offers to reload tabs still running the old copy", async () => {
    const own = await launch();
    try {
      const tab = await openFixture(own, "inbox");
      await recordActions(tab);
      const options = await openOptions(own);
      await options.evaluate(() => chrome.runtime.reload()).catch(() => {});
      await tab.bringToFront();
      await tab.keyboard.press("y");
      const button = () =>
        tab.evaluate(
          () =>
            document.querySelector("overlook-root")?.shadowRoot?.querySelector(".notice .action")
              ?.textContent ?? null,
        );
      await expect.poll(button).toBe("Reload tab");
      expect(await actions(tab)).toEqual([]);
    } finally {
      await own.close();
    }
  });
});
