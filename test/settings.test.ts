// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { applyToDocument, defaultSettings, featureTokens, normalize } from "../src/settings.ts";

describe("settings", () => {
  it("merges stored values over defaults and ignores junk", () => {
    const settings = normalize({ keymap: false, density: "yes", palette: "neon", bogus: true });
    expect(settings.keymap).toBe(false);
    expect(settings.density).toBe(defaultSettings().density);
    expect(settings.palette).toBe(defaultSettings().palette);
    expect(settings).not.toHaveProperty("bogus");
  });

  it("accepts known choice values", () => {
    expect(normalize({ palette: "solarized", mode: "dark" })).toMatchObject({
      palette: "solarized",
      mode: "dark",
    });
  });

  it("emits a token per enabled feature", () => {
    const tokens = featureTokens({ ...defaultSettings(), density: true, "hide-ribbon": false });
    expect(tokens.split(" ")).toContain("density");
    expect(tokens.split(" ")).not.toContain("hide-ribbon");
  });

  it("writes features and choices onto the root element", () => {
    applyToDocument({ ...defaultSettings(), palette: "paper" }, document.documentElement);
    expect(document.documentElement.dataset.overlookPalette).toBe("paper");
    expect(document.documentElement.dataset.overlookFeatures).toContain("keymap");
  });
});
