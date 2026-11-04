// Renders icons/icon.svg into the PNG sizes Chrome needs. Usage: node scripts/icons.ts
import { readFile } from "node:fs/promises";
import { chromium } from "playwright";

const svg = await readFile("icons/icon.svg", "utf8");
const browser = await chromium.launch();
const page = await browser.newPage();
for (const size of [16, 32, 48, 128]) {
  await page.setViewportSize({ width: size, height: size });
  await page.setContent(
    `<style>*{margin:0}svg{width:${size}px;height:${size}px;display:block}</style>${svg}`,
  );
  await page.screenshot({ path: `icons/icon-${size}.png`, omitBackground: true });
}
await browser.close();
