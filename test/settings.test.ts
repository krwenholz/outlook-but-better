import { describe, expect, it } from "vitest";
import { defaultSettings, featureTokens, normalize } from "../src/settings.ts";

describe("settings", () => {
  it("merges stored values over defaults and ignores junk", () => {
    const settings = normalize({ keymap: false, theme: "yes", bogus: true });
    expect(settings.keymap).toBe(false);
    expect(settings.theme).toBe(defaultSettings().theme);
    expect(settings).not.toHaveProperty("bogus");
  });

  it("emits a token per enabled feature", () => {
    const tokens = featureTokens({ ...defaultSettings(), density: true, theme: false }).split(" ");
    expect(tokens).toContain("density");
    expect(tokens).not.toContain("theme");
  });
});
