// Downloads the Outlook stylesheets fixtures reference (test/fixtures/assets.json) from Outlook's
// public CDN into test/fixtures/assets/, which is gitignored: they are Microsoft's files, so the
// repo only records their names. Only needed for screenshots; tests run without them.

import { existsSync } from "node:fs";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { OUTLOOK_CDN as CDN } from "./scrub-fixture.ts";

const dir = "test/fixtures/assets";

const files: string[] = JSON.parse(await readFile("test/fixtures/assets.json", "utf8"));
await mkdir(dir, { recursive: true });
let failed = 0;
for (const file of files) {
  const target = path.join(dir, file);
  if (existsSync(target)) continue;
  const response = await fetch(CDN + file, { signal: AbortSignal.timeout(30_000) });
  if (!response.ok) {
    console.error(`${file}: HTTP ${response.status}`);
    failed++;
    continue;
  }
  await writeFile(target, Buffer.from(await response.arrayBuffer()));
  console.log(`fetched ${file}`);
}
if (failed) process.exit(1);
