// Turns a saved Outlook page into a fixture that is safe to commit.
// Usage: node scripts/scrub-fixture.ts fixtures-raw/inbox.html test/fixtures/inbox.html
//
// Everything that could carry mail content (text, labels, titles, links, images, data-* values)
// is masked, except attribute values the selector registry itself looks for, so the health check
// still works against the fixture. Structure, roles, IDs and classes are kept.

import { readFile, writeFile } from "node:fs/promises";
import { JSDOM } from "jsdom";
import { SELECTORS } from "../src/selectors.ts";

const DROP_ELEMENTS =
  "script, noscript, iframe, object, embed, base, link:not([rel=stylesheet]), meta:not([charset])";
const TEXT_ATTRIBUTES = new Set([
  "aria-label",
  "aria-description",
  "aria-placeholder",
  "aria-valuetext",
  "title",
  "alt",
  "placeholder",
  "value",
  "content",
  "download",
]);
const DROP_ATTRIBUTES = new Set(["src", "srcset", "poster", "background", "action", "formaction"]);
const EMAIL = /[\w.+-]+@[\w-]+(\.[\w-]+)+/g;
const SAFE_DATA_VALUE = /^[A-Za-z_-]{1,40}$/;

interface LabelRule {
  attr: string;
  value: string;
  contains: boolean;
}

/** Attribute values the registry matches on, e.g. [aria-label="Archive"] or [aria-label*="Copilot" i]. */
export function registryLabels(): LabelRule[] {
  const rules: LabelRule[] = [];
  const pattern = /\[([\w-]+)(\*?)="([^"]+)"( i)?\]/g;
  for (const { candidates } of Object.values(SELECTORS)) {
    for (const candidate of candidates) {
      for (const [, attr = "", star, value = ""] of candidate.matchAll(pattern)) {
        rules.push({ attr, value, contains: star === "*" });
      }
    }
  }
  return rules;
}

const mask = (text: string) => text.replace(/\p{L}/gu, "x").replace(/\p{N}/gu, "0");

export interface ScrubStats {
  textNodes: number;
  attributes: number;
  keptLabels: number;
  dropped: number;
}

export function scrubHtml(html: string): { html: string; stats: ScrubStats } {
  const dom = new JSDOM(html);
  const { document, NodeFilter } = dom.window;
  const stats: ScrubStats = { textNodes: 0, attributes: 0, keptLabels: 0, dropped: 0 };
  const labels = registryLabels();

  for (const node of document.querySelectorAll(DROP_ELEMENTS)) {
    node.remove();
    stats.dropped++;
  }

  const walker = document.createTreeWalker(
    document,
    NodeFilter.SHOW_TEXT | NodeFilter.SHOW_COMMENT,
  );
  const comments: Node[] = [];
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (node.nodeType === node.COMMENT_NODE) comments.push(node);
    else if (node.parentElement?.tagName !== "STYLE" && node.textContent?.trim()) {
      node.textContent = mask(node.textContent);
      stats.textNodes++;
    }
  }
  for (const comment of comments) comment.parentNode?.removeChild(comment);

  for (const element of document.querySelectorAll("*")) {
    for (const { name, value } of [...element.attributes]) {
      let next: string | null = value;
      if (DROP_ATTRIBUTES.has(name)) next = null;
      else if (name === "href") next = value.startsWith("#") ? "#" : null;
      else if (name === "style") next = value.replace(/url\([^)]*\)/g, "none");
      else if (TEXT_ATTRIBUTES.has(name)) {
        const rule = labels.find(
          (r) =>
            r.attr === name &&
            (r.contains ? value.toLowerCase().includes(r.value.toLowerCase()) : value === r.value),
        );
        // For a contains-match keep only the matched part, so nothing around it leaks.
        next = rule ? rule.value : mask(value);
        if (rule) stats.keptLabels++;
      } else if (name.startsWith("data-") && !SAFE_DATA_VALUE.test(value)) next = mask(value);
      next = next?.replace(EMAIL, "user@example.com") ?? null;
      if (next === value) continue;
      stats.attributes++;
      if (next === null) element.removeAttribute(name);
      else element.setAttribute(name, next);
    }
  }

  return { html: dom.serialize().replace(EMAIL, "user@example.com"), stats };
}

if (import.meta.main) {
  const [input, output] = process.argv.slice(2);
  if (!input || !output) {
    console.error("usage: node scripts/scrub-fixture.ts <saved-page.html> <fixture.html>");
    process.exit(2);
  }
  const { html, stats } = scrubHtml(await readFile(input, "utf8"));
  await writeFile(output, html);
  console.log(`wrote ${output}`, stats);
  console.log("Skim the output before committing: the scrubber is a safety net, not a guarantee.");
}
