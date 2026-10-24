// Runs the selector health check against a saved (or scrubbed) Outlook page.
// Usage: node scripts/health.ts test/fixtures/inbox.html

import { readFile } from "node:fs/promises";
import { JSDOM } from "jsdom";
import { brokenEntries, healthCheck, SELECTORS } from "../src/selectors.ts";

const [file] = process.argv.slice(2);
if (!file) {
  console.error("usage: node scripts/health.ts <page.html>");
  process.exit(2);
}

const { document } = new JSDOM(await readFile(file, "utf8")).window;
const results = healthCheck(document);
console.table(
  results.map((r) => ({
    name: r.name,
    expect: r.expect,
    matched: r.candidate < 0 ? "-" : SELECTORS[r.name].candidates[r.candidate],
    count: r.count,
  })),
);
const broken = brokenEntries(results);
if (broken.length) {
  console.error(`broken (expected on every mail view): ${broken.join(", ")}`);
  process.exit(1);
}
