// Developer capture: serialises the live Outlook page into one HTML file, so a fixture can be
// made (docs/capturing-pages.md) without DevTools. Saved locally by the caller; never sent
// anywhere. The file holds real mail, so it goes through scripts/scrub-fixture.ts before it
// goes anywhere near the repo.

import { brokenEntries, healthCheck } from "./selectors.ts";

/** Message the popup sends to the content script of the active tab. */
export const CAPTURE_MESSAGE = "obb-capture";

const OBB_ATTRIBUTE = /^data-overlook/;
/** The scrubber (scripts/scrub-fixture.ts) looks for this marker to split out runtime styles. */
export const RUNTIME_STYLE_MARKER = "data-overlook-runtime";

export interface CaptureMeta {
  capturedAt: string;
  obbVersion: string;
  /** First two path segments only: deeper ones carry message ids. */
  route: string;
  viewport: string;
  settings: Record<string, string | undefined>;
  health: { broken: string[]; entries: { name: string; candidate: number; count: number }[] };
}

export interface Capture {
  html: string;
  name: string;
  meta: CaptureMeta;
}

/** Style rules Outlook injected at runtime (CSS-in-JS), minus the ones OBB itself added. */
function runtimeCss(doc: Document): string {
  const sheets = [...doc.styleSheets, ...(doc.adoptedStyleSheets ?? [])].filter(
    (sheet) => !("href" in sheet && sheet.href),
  );
  return sheets
    .flatMap((sheet) => {
      try {
        return [...sheet.cssRules].map((rule) => rule.cssText);
      } catch {
        return []; // cross-origin sheet
      }
    })
    .filter((rule) => !rule.includes("data-overlook"))
    .join("\n");
}

export function captureFilename(now = new Date()): string {
  const stamp = now.toISOString().slice(0, 16).replace(/\D/g, "");
  return `obb-capture-${stamp}.html`;
}

export function capturePage(
  doc: Document = document,
  loc: Pick<Location, "pathname"> = location,
  version = "dev",
  now = new Date(),
): Capture {
  const live = doc.documentElement.dataset;
  const results = healthCheck(doc);
  const meta: CaptureMeta = {
    capturedAt: now.toISOString(),
    obbVersion: version,
    route: loc.pathname.split("/").slice(0, 3).join("/"),
    viewport: `${doc.defaultView?.innerWidth ?? 0}x${doc.defaultView?.innerHeight ?? 0}`,
    settings: {
      features: live.overlookFeatures,
      palette: live.overlookPalette,
      mode: live.overlookMode,
      layout: live.overlookLayout,
      outlookScheme: live.overlookOutlookScheme,
      readingPane: live.overlookReadingPane,
    },
    health: {
      broken: brokenEntries(results),
      entries: results.map(({ name, candidate, count }) => ({ name, candidate, count })),
    },
  };

  // Work on a copy, with OBB's own marks removed, so the capture is stock Outlook.
  const root = doc.documentElement.cloneNode(true) as HTMLElement;
  for (const el of root.querySelectorAll("overlook-root")) el.remove();
  for (const el of [root, ...root.querySelectorAll("*")]) {
    for (const { name } of [...el.attributes]) {
      if (OBB_ATTRIBUTE.test(name)) el.removeAttribute(name);
    }
  }

  const head = root.querySelector("head");
  const style = doc.createElement("style");
  style.setAttribute(RUNTIME_STYLE_MARKER, "");
  style.textContent = runtimeCss(doc);
  head?.append(style);
  const script = doc.createElement("script");
  script.type = "application/json";
  script.id = "obb-capture-meta";
  script.textContent = JSON.stringify(meta, null, 2).replaceAll("<", "\\u003c");
  head?.append(script);

  return { html: `<!doctype html>\n${root.outerHTML}`, name: captureFilename(now), meta };
}
