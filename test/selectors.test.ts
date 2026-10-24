// @vitest-environment jsdom
import { beforeEach, describe, expect, it } from "vitest";
import {
  brokenEntries,
  healthCheck,
  readFolderTree,
  resolve,
  SELECTORS,
  tagElements,
} from "../src/selectors.ts";

beforeEach(() => {
  document.body.innerHTML = `
    <div role="navigation">
      <div role="tree" aria-label="Folders">
        <div role="treeitem" aria-level="1" title="Inbox"></div>
        <div role="treeitem" aria-level="2" title="Receipts"></div>
        <div role="treeitem" aria-level="1" title="Archive"></div>
      </div>
    </div>
    <div role="listbox" aria-label="Message list"></div>
    <input aria-label="Search" />
    <button aria-label="New mail"></button>
    <button aria-label="Copilot"></button>
    <button aria-label="Ask Copilot"></button>
  `;
});

describe("selector registry", () => {
  it("only nests entries inside entries that exist", () => {
    for (const entry of Object.values(SELECTORS)) {
      if ("within" in entry) expect(SELECTORS).toHaveProperty(entry.within);
    }
  });

  it("falls back to later candidates", () => {
    expect(resolve("messageList")?.getAttribute("aria-label")).toBe("Message list");
    const health = healthCheck().find((r) => r.name === "messageList");
    expect(health?.candidate).toBe(1);
  });

  it("reports missing 'always' entries as broken", () => {
    expect(brokenEntries(healthCheck())).toEqual([]);
    document.querySelector("[aria-label='New mail']")?.remove();
    expect(brokenEntries(healthCheck())).toEqual(["composeButton"]);
  });

  it("tags elements idempotently, including every match for 'all' entries", () => {
    tagElements();
    tagElements();
    expect(document.querySelectorAll('[data-overlook~="copilot"]')).toHaveLength(2);
    expect(resolve("messageList")?.getAttribute("data-overlook")).toBe("message-list");
  });

  it("reads nested folder paths from aria-level", () => {
    expect(readFolderTree().map((f) => f.path)).toEqual(["Inbox", "Inbox/Receipts", "Archive"]);
  });
});
