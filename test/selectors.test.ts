// @vitest-environment jsdom
import { readFileSync } from "node:fs";
import { beforeEach, describe, expect, it } from "vitest";
import {
  brokenEntries,
  healthCheck,
  isUnreadRow,
  outlookColorScheme,
  readFolderTree,
  resolve,
  resolveAll,
  SELECTORS,
  selectedRows,
  tagElements,
} from "../src/selectors.ts";

const inbox = readFileSync("test/fixtures/inbox.html", "utf8");

describe("selector registry", () => {
  it("only nests entries inside entries that exist", () => {
    for (const entry of Object.values(SELECTORS)) {
      if ("within" in entry) expect(SELECTORS).toHaveProperty(entry.within);
    }
  });

  describe("against the inbox fixture", () => {
    beforeEach(() => {
      document.documentElement.innerHTML = inbox;
    });

    it("finds every 'always' entry", () => {
      expect(brokenEntries(healthCheck())).toEqual([]);
    });

    it("uses the preferred candidate for structural entries", () => {
      const used = new Map(healthCheck().map((r) => [r.name, r.candidate]));
      for (const name of ["messageList", "folderPane", "listPane", "readingPane", "ribbon"]) {
        expect(used.get(name as keyof typeof SELECTORS)).toBe(0);
      }
    });

    it("tags elements idempotently, including every match of union entries", () => {
      tagElements();
      tagElements();
      expect(document.querySelectorAll('[data-overlook~="message-row"]')).toHaveLength(9);
      expect(document.querySelectorAll('[data-overlook~="header-extra"]')).toHaveLength(2);
      expect(document.querySelectorAll('[data-overlook~="rail-extra"]')).toHaveLength(4);
      expect(resolve("messageList")?.getAttribute("data-overlook")).toBe("message-list");
      expect(document.querySelectorAll('[data-overlook~="info-bar"]')).toHaveLength(3);
      expect(document.querySelectorAll('[data-overlook~="app-banner"]')).toHaveLength(1);
    });

    it("reads mailbox folders without favorites, roots or item counts", () => {
      const names = readFolderTree().map((f) => f.path);
      expect(names).toContain("Inbox");
      expect(names).toContain("Archive");
      expect(names.filter((n) => n === "Inbox")).toHaveLength(1);
      expect(names.some((n) => /items/.test(n))).toBe(false);
    });

    it("reads which rows are selected and unread", () => {
      const [selected] = selectedRows();
      expect(selectedRows()).toHaveLength(1);
      expect(selected && isUnreadRow(selected)).toBe(true);
      expect(resolveAll("messageRow").filter(isUnreadRow)).toHaveLength(1);
    });

    it("reads Outlook's colour scheme", () => {
      expect(outlookColorScheme(document.documentElement)).toBe("light");
    });
  });

  it("falls back to later candidates", () => {
    document.body.innerHTML = '<div role="listbox" aria-label="Message list"></div>';
    expect(resolve("messageList")).not.toBeNull();
    expect(healthCheck().find((r) => r.name === "messageList")?.candidate).toBe(1);
  });

  it("matches every candidate of a union entry", () => {
    document.body.innerHTML =
      '<b id="owaTeamsBtn_container"></b><b id="owaTimePanelBtn_container"></b>';
    expect(resolveAll("headerExtras")).toHaveLength(2);
  });
});
