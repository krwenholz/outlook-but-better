import { describe, expect, it } from "vitest";
import { ACTIONS } from "../src/keymap/actions.ts";
import { BINDINGS } from "../src/keymap/bindings.ts";
import { findConflicts } from "../src/keymap/engine.ts";
import { SELECTORS } from "../src/selectors.ts";

describe("binding set", () => {
  it("has no shadowed or duplicate bindings", () => {
    expect(findConflicts(BINDINGS)).toEqual([]);
  });

  it("binds every action at least once", () => {
    const bound = new Set(BINDINGS.map((b) => b.action));
    expect(Object.keys(ACTIONS).filter((id) => !bound.has(id as keyof typeof ACTIONS))).toEqual([]);
  });

  it("only drives controls that exist in the selector registry", () => {
    for (const { strategy } of Object.values(ACTIONS)) {
      if (strategy.kind === "drive") expect(SELECTORS).toHaveProperty(strategy.target);
    }
  });
});
