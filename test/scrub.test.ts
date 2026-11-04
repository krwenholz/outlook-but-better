import { JSDOM } from "jsdom";
import { describe, expect, it } from "vitest";
import { assertClean, fakeText, isUiText, scrubHtml } from "../scripts/scrub-fixture.ts";
import { healthCheck } from "../src/selectors.ts";

const raw = `<!doctype html><html><head><script>secret()</script><title>Inbox - Jane Doe</title>
  <link rel="stylesheet" href="./Mail%20-%20Jane_files/owa.mail.css">
  <link rel="stylesheet" href="https://cdn.example/remote.css"></head><body>
  <div id="MailList" role="listbox" aria-label="Message list">
    <div role="option" id="AQAAAAAAAQ=" aria-labelledby="AQAAAAAAAQ=" data-convid="AAQkADAwATM0MDAAMS1"
         aria-label="From Jane Doe, Subject: Merger plans at Contoso">
      <span>jane.doe@contoso.com</span><span>Merger plans for Q3, ref 4815162342</span>
      <img src="./Mail%20-%20Jane_files/avatar.jpg" alt="Jane Doe">
      <a href="https://contoso.com/secret">link</a>
      <span originalsrc="https://contoso.zoom.us/j/123?pwd=hunter2">join</span>
    </div>
  </div>
  <div role="treeitem" data-folder-name="contoso clients" title="Inbox - 9 items (1 unread)"></div>
  <button aria-label="Archive">Archive</button>
  <span>Move this message to your archive folder. (E)</span>
  <button aria-label="Copilot summary of Merger plans"></button>
  <!-- Jane's comment -->
</body></html>`;

describe("scrubHtml", () => {
  const { html, assets } = scrubHtml(raw, { assetDir: "Mail - Jane_files", denylist: ["contoso"] });

  it("removes names, addresses, subjects, links, ids, tokens and images", () => {
    for (const secret of [
      "Jane",
      "Doe",
      "contoso",
      "Merger",
      "secret",
      "AAQk",
      "AQAAAAAAAQ",
      "avatar",
      "4815162342",
      "hunter2",
      "zoom",
    ]) {
      expect(html).not.toContain(secret);
    }
  });

  it("keeps Outlook UI text and the labels the registry depends on", () => {
    const { document } = new JSDOM(html).window;
    const found = new Map(healthCheck(document).map((r) => [r.name, r.candidate >= 0]));
    expect(found.get("messageList")).toBe(true);
    expect(found.get("archiveButton")).toBe(true);
    expect(
      document.querySelector('button[aria-label*="Copilot"]')?.getAttribute("aria-label"),
    ).toBe("Copilot");
    expect(html).toContain("Move this message to your archive folder. (E)");
    expect(html).toContain('title="Inbox - 9 items (1 unread)"');
  });

  it("rewrites id references consistently", () => {
    const { document } = new JSDOM(html).window;
    const row = document.querySelector('[role="option"]');
    expect(row?.getAttribute("aria-labelledby")).toBe(row?.id);
  });

  it("keeps only local stylesheets and safe assets", () => {
    expect(assets).toEqual(["owa.mail.css"]);
    expect(html).toContain('href="assets/owa.mail.css"');
    expect(html).not.toContain("cdn.example");
  });

  it("refuses output that still contains a denylisted term", () => {
    expect(() => assertClean("hello Contoso", ["contoso"])).toThrow();
    expect(() => assertClean("flags@2x.png someone@example.com", [])).not.toThrow();
  });
});

describe("fake text", () => {
  it("is deterministic and keeps UI vocabulary", () => {
    expect(fakeText("Quarterly numbers")).toBe(fakeText("Quarterly numbers"));
    expect(fakeText("Quarterly numbers")).not.toBe("Quarterly numbers");
    expect(isUiText("Sent Items - 4 items (0 unread)")).toBe(true);
    expect(isUiText("Thu 12/17/2026 8:30 AM")).toBe(true);
    expect(isUiText("Design & Make Conference")).toBe(false);
  });

  it("keeps the state words that start a message row's label", () => {
    const faked = fakeText("Unread Has attachments Merger plans for Q3");
    expect(faked.startsWith("Unread Has attachments ")).toBe(true);
    expect(faked).not.toContain("Merger");
  });

  it("masks every digit of phone numbers, not just the long groups", () => {
    const faked = fakeText("Call +44 (0)121 766 9274 or 555-867-5309");
    expect(faked).not.toMatch(/121|766|9274|867|5309/);
    expect(fakeText("Thu 12/17/2026 8:30 AM")).toBe("Thu 12/17/2026 8:30 AM");
  });
});
