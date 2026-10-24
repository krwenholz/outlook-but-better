import { JSDOM } from "jsdom";
import { describe, expect, it } from "vitest";
import { scrubHtml } from "../scripts/scrub-fixture.ts";
import { healthCheck } from "../src/selectors.ts";

const raw = `<!doctype html><html><head><script>secret()</script><title>Inbox - Jane Doe</title></head><body>
  <div id="MailList" role="listbox">
    <div role="option" aria-label="From Jane Doe, Subject: Merger plans" data-convid="AAQkADAwATM0MDAAMS1">
      <span>jane.doe@contoso.com</span><span>Merger plans for Q3</span>
      <img src="https://contoso.com/avatar.png" alt="Jane Doe">
      <a href="https://contoso.com/secret">link</a>
    </div>
  </div>
  <button aria-label="Archive">Archive</button>
  <button aria-label="Copilot summary of Merger plans"></button>
  <!-- Jane's comment -->
</body></html>`;

describe("scrubHtml", () => {
  const { html } = scrubHtml(raw);

  it("removes names, addresses, subjects, links and images", () => {
    for (const secret of ["Jane", "Doe", "contoso", "Merger", "secret", "AAQk", "avatar"]) {
      expect(html).not.toContain(secret);
    }
  });

  it("keeps structure and the labels the registry depends on", () => {
    const { document } = new JSDOM(html).window;
    const found = new Map(healthCheck(document).map((r) => [r.name, r.candidate >= 0]));
    expect(found.get("messageList")).toBe(true);
    expect(found.get("archiveButton")).toBe(true);
    expect(found.get("copilot")).toBe(true);
    expect(
      document.querySelector('button[aria-label*="Copilot"]')?.getAttribute("aria-label"),
    ).toBe("Copilot");
  });
});
