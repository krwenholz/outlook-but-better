// @vitest-environment jsdom
import { readFileSync } from "node:fs";
import { beforeEach, describe, expect, it } from "vitest";
import { captureFilename, capturePage, RUNTIME_STYLE_MARKER } from "../src/capture.ts";

const inbox = readFileSync("test/fixtures/inbox.html", "utf8");

describe("capturePage", () => {
  beforeEach(() => {
    document.documentElement.innerHTML = inbox;
    document.documentElement.dataset.overlookPalette = "solarized";
    document.documentElement.dataset.overlookFeatures = "keymap density";
    document.body.append(document.createElement("overlook-root"));
    document.querySelector("#MailList")?.setAttribute("data-overlook", "message-list");
  });

  const take = () =>
    capturePage(
      document,
      { pathname: "/mail/id/AAQkADAwATM0MDAAMS1" },
      "9.9.9",
      new Date("2026-11-08T17:30:00Z"),
    );

  it("leaves out OBB's own marks, so the capture is stock Outlook", () => {
    const { html } = take();
    expect(html.startsWith("<!doctype html>")).toBe(true);
    expect(html).not.toContain("overlook-root");
    expect(html).not.toMatch(/data-overlook(?!-runtime)/);
    expect(html).toContain('id="MailList"');
  });

  it("keeps the runtime-styles marker the scrubber looks for", () => {
    expect(take().html).toContain(`<style ${RUNTIME_STYLE_MARKER}`);
  });

  it("records version, settings and health, with the route trimmed of message ids", () => {
    const { meta, html } = take();
    expect(meta).toMatchObject({
      obbVersion: "9.9.9",
      route: "/mail/id",
      capturedAt: "2026-11-08T17:30:00.000Z",
      settings: { palette: "solarized", features: "keymap density" },
    });
    expect(meta.health.broken).toEqual([]);
    expect(html).toContain('id="obb-capture-meta"');
    expect(html).not.toContain("AAQkADAwATM0MDAAMS1");
  });

  it("names the file by minute", () => {
    expect(captureFilename(new Date("2026-11-08T17:30:45Z"))).toBe("obb-capture-202611081730.html");
  });
});
