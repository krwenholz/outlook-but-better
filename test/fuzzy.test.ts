import { describe, expect, it } from "vitest";
import { fuzzyFilter, score } from "../src/palette/fuzzy.ts";

const folders = [
  "Inbox",
  "Inbox/Receipts",
  "Archive",
  "Projects/Acme",
  "Projects/Billing",
  "Sent Items",
];
const top = (query: string) => fuzzyFilter(query, folders, (f) => f)[0]?.item;

describe("fuzzy", () => {
  it("matches subsequences case-insensitively", () => {
    expect(score("rcpt", "Inbox/Receipts")).not.toBeNull();
    expect(score("zzz", "Inbox")).toBeNull();
  });

  it("returns every item for an empty query", () => {
    expect(fuzzyFilter("", folders, (f) => f)).toHaveLength(folders.length);
  });

  it("prefers word starts and shorter names", () => {
    expect(top("inb")).toBe("Inbox");
    expect(top("acme")).toBe("Projects/Acme");
    expect(top("si")).toBe("Sent Items");
    expect(top("pb")).toBe("Projects/Billing");
  });

  it("does not let a later word start steal a needed character", () => {
    expect(score("bc", "abc b")).not.toBeNull();
  });

  it("reports matched indices for highlighting", () => {
    expect(score("arc", "Archive")?.indices).toEqual([0, 1, 2]);
  });
});
