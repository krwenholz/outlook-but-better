import { describe, expect, it } from "vitest";
import { type Binding, findConflicts, keyFromEvent, SequenceEngine } from "../src/keymap/engine.ts";

const bindings: Binding[] = [
  { keys: ["j"], action: "next", contexts: ["list", "reading"] },
  { keys: ["r"], action: "reply", contexts: ["reading"] },
  { keys: [".", "r"], action: "markRead", contexts: ["list", "reading"] },
  { keys: ["G", "m"], action: "goMail", contexts: ["list"] },
];

describe("SequenceEngine", () => {
  it("matches single keys in the right context only", () => {
    const engine = new SequenceEngine(bindings);
    expect(engine.feed("j", "list", 0)).toEqual({ kind: "match", action: "next" });
    expect(engine.feed("r", "list", 0)).toEqual({ kind: "none" });
    expect(engine.feed("r", "reading", 0)).toEqual({ kind: "match", action: "reply" });
  });

  it("holds a prefix and completes a sequence", () => {
    const engine = new SequenceEngine(bindings);
    const first = engine.feed(".", "list", 0);
    expect(first.kind).toBe("pending");
    expect(engine.pending).toEqual(["."]);
    expect(engine.feed("r", "list", 100)).toEqual({ kind: "match", action: "markRead" });
    expect(engine.pending).toEqual([]);
  });

  it("lists the possible next keys while pending", () => {
    const engine = new SequenceEngine(bindings);
    const result = engine.feed("G", "list", 0);
    expect(result.kind === "pending" && result.next.map((b) => b.action)).toEqual(["goMail"]);
  });

  it("cancels (and swallows) an unknown key after a prefix", () => {
    const engine = new SequenceEngine(bindings);
    engine.feed(".", "list", 0);
    expect(engine.feed("q", "list", 10)).toEqual({ kind: "cancelled" });
    expect(engine.feed("j", "list", 20)).toEqual({ kind: "match", action: "next" });
  });

  it("drops a stale prefix after the timeout", () => {
    const engine = new SequenceEngine(bindings, 1000);
    engine.feed(".", "list", 0);
    expect(engine.feed("r", "reading", 5000)).toEqual({ kind: "match", action: "reply" });
  });

  it("does not offer sequences from another context", () => {
    const engine = new SequenceEngine(bindings);
    expect(engine.feed("G", "reading", 0)).toEqual({ kind: "none" });
  });
});

describe("findConflicts", () => {
  it("reports a key that shadows a longer sequence in a shared context", () => {
    expect(
      findConflicts([
        { keys: ["g"], action: "goTo", contexts: ["list"] },
        { keys: ["g", "i"], action: "inbox", contexts: ["list"] },
      ]),
    ).toHaveLength(1);
  });

  it("allows the same keys in disjoint contexts", () => {
    expect(
      findConflicts([
        { keys: ["r"], action: "reply", contexts: ["reading"] },
        { keys: ["r"], action: "other", contexts: ["list"] },
      ]),
    ).toEqual([]);
  });
});

describe("keyFromEvent", () => {
  const event = (
    key: string,
    mods: Partial<Record<"ctrlKey" | "metaKey" | "altKey", boolean>> = {},
  ) => ({
    key,
    ctrlKey: false,
    metaKey: false,
    altKey: false,
    ...mods,
  });

  it("passes plain and shifted keys through", () => {
    expect(keyFromEvent(event("j"))).toBe("j");
    expect(keyFromEvent(event("G"))).toBe("G");
  });

  it("ignores chords and bare modifiers", () => {
    expect(keyFromEvent(event("r", { metaKey: true }))).toBeNull();
    expect(keyFromEvent(event("Shift"))).toBeNull();
  });
});
