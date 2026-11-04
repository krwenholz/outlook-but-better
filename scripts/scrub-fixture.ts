// Turns a saved Outlook page into a fixture that is safe to commit.
// Usage: node scripts/scrub-fixture.ts <saved.html> <out-dir> [--name inbox]
//
// Writes <out-dir>/<name>.html and adds the page's stylesheets to <out-dir>/assets.json. The
// stylesheets themselves are Microsoft's, so they are never committed: `mise run fixtures:assets`
// downloads them from Outlook's public CDN into the gitignored <out-dir>/assets/.
//
// What survives: structure, roles, classes, semantic data-* values, attribute values the selector
// registry matches on, and text made up entirely of Outlook UI vocabulary ("Inbox - 9 items").
// Everything else (names, subjects, bodies, addresses, ids, links, images) becomes deterministic
// fake words of similar shape, so screenshots stay readable.
//
// Optional fixtures-raw/denylist.txt (gitignored) lists extra terms, one per line, that must not
// appear anywhere in the output, e.g. employer names. The scrub fails if any survive.

import { createHash } from "node:crypto";
import { existsSync } from "node:fs";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { parseArgs } from "node:util";
import { JSDOM } from "jsdom";
import { SELECTORS } from "../src/selectors.ts";

const DROP_ELEMENTS = [
  "script",
  "noscript",
  "iframe",
  "object",
  "embed",
  "base",
  "meta:not([charset])",
  "link:not([rel=stylesheet])",
].join(", ");

/** Attributes whose values are human-readable text. */
const TEXT_ATTRIBUTES = new Set([
  "aria-label",
  "aria-description",
  "aria-placeholder",
  "aria-valuetext",
  "aria-roledescription",
  "title",
  "alt",
  "placeholder",
  "value",
  "content",
  "label",
  "name",
  "download",
  "datetime",
]);
/** Attributes that hold element ids (space-separated), rewritten consistently with `id`. */
const ID_REF_ATTRIBUTES = new Set([
  "id",
  "for",
  "aria-labelledby",
  "aria-describedby",
  "aria-controls",
  "aria-owns",
  "aria-activedescendant",
  "aria-details",
  "aria-flowto",
  "headers",
  "list",
]);
/** Structural or presentational attributes kept verbatim. Anything unlisted is treated as text. */
const KEEP_ATTRIBUTES = new Set(
  `class role tabindex type dir lang hidden disabled draggable contenteditable spellcheck
  translate autocomplete charset rel media crossorigin colspan rowspan align valign border
  cellpadding cellspacing bgcolor color width height xmlns xmlns:xlink version viewBox d fill
  fill-rule fill-opacity clip-rule clip-path stroke stroke-width stroke-linecap stroke-linejoin
  opacity transform gradientTransform gradientUnits offset stop-color stop-opacity x y x1 x2 y1
  y2 cx cy r rx ry fx fy points preserveAspectRatio focusable mask filter`.split(/\s+/),
);
const URL_ATTRIBUTES = new Set(["src", "srcset", "poster", "background", "action", "formaction"]);

/** Outlook's public asset host (fonts, app icons); URLs there identify no one. */
const PUBLIC_CDN = "https://res.public.onecdn.static.microsoft/";
/** Where Outlook serves its stylesheets; scripts/fetch-fixture-assets.ts downloads from here. */
export const OUTLOOK_CDN = "https://res.public.onecdn.static.microsoft/owamail/hashed-v1/scripts/";
/** Saved assets the fixture may reference: Outlook's stylesheets, fetched from its CDN. */
const SAFE_ASSET = /^owa\.[\w.]+\.css$/;

const EMAIL = /[\w.+-]+@[\w-]+(\.[\w-]+)+/g;
/** "flags@2x.png" looks like an address but is an image file. */
const ASSET_FILE = /\.(png|jpe?g|gif|svg|webp|css|js|woff2?)$/i;
const HAS_EMAIL = /[\w.+-]+@[\w-]+(\.[\w-]+)+/;
const SAFE_TOKEN = /^[A-Za-z][\w-]{0,40}$/;

// Outlook's own UI words. Text made only of these (plus short numbers) is kept verbatim.
// Never add names, company or product terms that are not Outlook's own.
const UI_PHRASES = `
  accept accept the meeting access additional features add-ins app launcher apps archive
  assign policy assign an archive or retention policy to automatically archive or delete the
  selected items august july june may april march february january september october november
  december calendar categorize conversation history create a new email message delete deleted
  items decline decline the meeting drafts enhance outlook with apps expand to see delete flag
  forward more new options favorites file filter unflag flag or unflag this message focused
  follow go to groups header action menu help hide navigation pane this message until later time
  home inbox junk last month this week today yesterday older mail make rule around email authors
  mark as read unread meeting response actions items move options quick steps & notes other
  people phish alert report phishing pin unpin print promotions reply all respond ribbon tabs
  search folders for meetings files and sent settings share using teams snooze sorted by date
  sweep switch layouts sync received tags chat to do undo view viva insights my day
  notifications newsletters onedrive org explorer quick actions reactions in-place reading
  message list select jump left-rail-appbar navigation skip main content of your folder at top
  keep customize expand collapse conversation show original size close email body from mark
  read selected opens card more apps items a an the is in on am pm last action backspace
  mon tue wed thu fri sat sun monday tuesday wednesday thursday friday saturday sunday
`;
const UI_WORDS = new Set(UI_PHRASES.split(/\s+/).filter(Boolean));

const FAKE_WORDS = `
  a an at by go if in my of on to up we ash bay cab dew elm fig gum hay ivy jam kit log mud nut
  oak pea rye sea tea urn van wax yew zip acre bark cove dune fern gale hill isle jade kelp lake
  mesa nook opal pine quay reef sage tarn vale wren yarn amber birch cedar delta ember flint
  grove heron inlet juniper karst larch maple north ocean pecan quail river slate thyme umber
  violet willow aspen basalt canyon dahlia estuary falcon garnet harbor indigo jasmine kestrel
  lagoon meadow nettle orchard pebble quartz rosemary saffron tundra upland verbena walnut
  glacier hazelnut lavender magnolia sandpiper tamarack wildflower blackberry cottonwood
  marshland riverstone
`
  .split(/\s+/)
  .filter(Boolean);
const FAKE_BY_LENGTH = new Map<number, string[]>();
for (const word of FAKE_WORDS) {
  FAKE_BY_LENGTH.set(word.length, [...(FAKE_BY_LENGTH.get(word.length) ?? []), word]);
}

const hash = (text: string) => createHash("sha256").update(text).digest();

function fakeWord(word: string): string {
  if (/^\d+$/.test(word)) {
    if (word.length <= 3 || /^(19|20)\d\d$/.test(word)) return word;
    const h = hash(word);
    return [...word].map((_, i) => String((h[i % h.length] ?? 0) % 10)).join("");
  }
  const h = hash(word.toLowerCase());
  let length = Math.min(Math.max(word.length, 2), 10);
  while (!FAKE_BY_LENGTH.has(length)) length--;
  const options = FAKE_BY_LENGTH.get(length) ?? FAKE_WORDS;
  const fake = options[h.readUInt32BE(0) % options.length] ?? "lorem";
  if (word === word.toUpperCase() && /\p{L}/u.test(word)) return fake.toUpperCase();
  if (/^\p{Lu}/u.test(word)) return fake[0]?.toUpperCase() + fake.slice(1);
  return fake;
}

const WORD = /[\p{L}\p{N}'’]+/gu;

/** True when every word is Outlook UI vocabulary or a short number. */
export function isUiText(text: string): boolean {
  if (HAS_EMAIL.test(text)) return false;
  for (const [word] of text.matchAll(WORD)) {
    // Single letters (shortcut hints like "(E)"), short counts, and years carry no identity.
    if (word.length === 1 || /^(\d{1,3}|(19|20)\d\d)$/.test(word)) continue;
    if (!UI_WORDS.has(word.toLowerCase())) return false;
  }
  return true;
}

/** Digit runs split by phone-style separators ("+44 (0)121 766 9274", "555-0100"). */
const PHONE = /\+?\(?\d[\d\s().-]{5,}\d/g;

/** Every digit of a phone-like run replaced, so short groups don't survive one by one. */
function maskPhones(text: string): string {
  return text.replace(PHONE, (run) => {
    if ((run.match(/\d/g) ?? []).length < 7) return run;
    const h = hash(run);
    let i = 0;
    return run.replace(/\d/g, () => String((h[i++ % h.length] ?? 0) % 10));
  });
}

/**
 * State words Outlook puts at the start of a message row's label ("Unread Has attachments
 * Quarterly numbers…"). They identify no one, and Overlook reads them (isUnreadRow).
 */
const ROW_STATES = [
  "Unread",
  "Has attachments",
  "Flagged",
  "Pinned",
  "Replied",
  "Forwarded",
  "Collapsed",
  "Expanded",
  "Important",
  "Meeting start time",
];
const LEADING_STATES = new RegExp(`^(?:(?:${ROW_STATES.join("|")})\\s+)+`);

export function fakeText(text: string): string {
  const masked = maskPhones(text);
  if (isUiText(masked)) return masked;
  const states = masked.match(LEADING_STATES)?.[0] ?? "";
  const rest = masked.slice(states.length);
  return states + rest.replace(EMAIL, "someone@example.com").replace(WORD, fakeWord);
}

/**
 * True for values that look like server-generated identifiers ("AAQkADAw…", GUIDs, item ids)
 * rather than names a developer chose ("ReadingPaneContainerId", "avatar-r3n").
 */
export function isOpaque(value: string): boolean {
  // Plain numbers (indexes, sizes, counts) are layout data, not identifiers.
  if (/^-?\d{1,5}(\.\d+)?$/.test(value)) return false;
  if (!SAFE_TOKEN.test(value) || /\d{4,}/.test(value)) return true;
  return value
    .split(/[_\-\s]+/)
    .some((seg) => seg.length >= 6 && (/[A-Z]{2,}[a-z]/.test(seg) || /[a-z]\d|\d[a-z]/i.test(seg)));
}

function fakeId(id: string): string {
  return isOpaque(id) ? `id-${hash(id).toString("hex").slice(0, 10)}` : id;
}

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

export interface ScrubOptions {
  /** Directory name the saved page used for its assets, e.g. "Mail - Outlook_files". */
  assetDir?: string;
  /** Extra terms that must not survive (case-insensitive). */
  denylist?: readonly string[];
  /** Fixture name; the runtime styles a DevTools capture carries go to vendor/<name>.runtime.css. */
  name?: string;
}

export interface ScrubResult {
  html: string;
  /** Asset file names referenced by the output, to copy into assets/. */
  assets: string[];
  /** Styles Fluent injected at runtime, when the capture kept them (docs/capturing-pages.md). */
  runtimeCss?: string;
}

/**
 * url(...) values other than inline data and Outlook's public CDN (fonts, icon fonts). Anything
 * else (avatars, attachments, tracking pixels) could identify someone, so it becomes `none`.
 */
const PRIVATE_URL =
  /url\((?!\s*["']?(?:data:|(?:https?:)?\/\/res\.public\.onecdn\.static\.microsoft\/))[^)]*\)/g;
const stripUrls = (css: string) => css.replace(PRIVATE_URL, "none");

export function scrubHtml(html: string, options: ScrubOptions = {}): ScrubResult {
  const dom = new JSDOM(html);
  const { document, NodeFilter } = dom.window;
  const labels = registryLabels();
  const assets = new Set<string>();

  const localAsset = (url: string): string | null => {
    const decoded = decodeURIComponent(url).replace(/^\.\//, "");
    const dir = options.assetDir;
    let file: string;
    if (dir && decoded.startsWith(`${dir}/`)) file = decoded.slice(dir.length + 1);
    else if (decoded.startsWith(OUTLOOK_CDN)) file = decoded.slice(OUTLOOK_CDN.length);
    else return null;
    if (!SAFE_ASSET.test(file)) return null;
    assets.add(file);
    return `assets/${file}`;
  };

  for (const node of document.querySelectorAll(DROP_ELEMENTS)) node.remove();
  for (const link of document.querySelectorAll("link[rel=stylesheet]")) {
    const href = localAsset(link.getAttribute("href") ?? "");
    if (!href) {
      link.remove();
      continue;
    }
    link.setAttribute("href", href);
    // CORS attributes make local (file://) loads fail.
    link.removeAttribute("crossorigin");
    link.removeAttribute("integrity");
  }

  const walker = document.createTreeWalker(
    document,
    NodeFilter.SHOW_TEXT | NodeFilter.SHOW_COMMENT,
  );
  const comments: Node[] = [];
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (node.nodeType === node.COMMENT_NODE) comments.push(node);
    else if (node.parentElement?.tagName === "STYLE") {
      node.textContent = stripUrls(node.textContent ?? "");
    } else if (node.textContent?.trim()) {
      node.textContent = fakeText(node.textContent);
    }
  }
  for (const comment of comments) comment.parentNode?.removeChild(comment);
  // A <noscript> in <head> parses into stray text at the top of <body>; it never renders live.
  for (const node of [...document.body.childNodes]) {
    if (node.nodeType === node.TEXT_NODE) node.remove();
  }

  for (const element of document.querySelectorAll("*")) {
    if (element.tagName === "LINK") continue;
    for (const { name, value } of [...element.attributes]) {
      let next: string | null = value;
      if (URL_ATTRIBUTES.has(name)) next = value.startsWith(PUBLIC_CDN) ? value : localAsset(value);
      else if (name === "href") next = value.startsWith("#") ? "#" : null;
      else if (name === "style") next = stripUrls(value);
      else if (ID_REF_ATTRIBUTES.has(name)) next = value.split(/\s+/).map(fakeId).join(" ");
      else if (TEXT_ATTRIBUTES.has(name)) {
        const rule = labels.find(
          (r) =>
            r.attr === name &&
            (r.contains ? value.toLowerCase().includes(r.value.toLowerCase()) : value === r.value),
        );
        // For a contains-match keep only the matched part, so nothing around it leaks.
        next = rule ? rule.value : fakeText(value);
      } else if (name === "data-folder-name") next = fakeText(value);
      else if (name.startsWith("data-") && value !== "" && isOpaque(value)) {
        next = hash(value).toString("hex").slice(0, 10);
      } else if (
        !name.startsWith("data-") &&
        !name.startsWith("aria-") &&
        !KEEP_ATTRIBUTES.has(name)
      ) {
        next = fakeText(value);
      }
      next = next?.replace(EMAIL, "someone@example.com") ?? null;
      if (next === value) continue;
      if (next === null) element.removeAttribute(name);
      else element.setAttribute(name, next);
    }
  }

  // Runtime styles are Microsoft's; they live in vendor/ under its own notice, not in the page.
  let runtimeCss: string | undefined;
  const runtime = document.querySelector("style[data-overlook-runtime]");
  if (runtime) {
    runtimeCss = deny(runtime.textContent ?? "", options.denylist);
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = `vendor/${options.name ?? "fixture"}.runtime.css`;
    runtime.replaceWith(link);
    assertClean(runtimeCss, options.denylist ?? []);
  }

  const out = deny(dom.serialize().replace(EMAIL, "someone@example.com"), options.denylist);
  assertClean(out, options.denylist ?? []);
  return { html: out, assets: [...assets].sort(), runtimeCss };
}

/** Replaces any denylisted term (case-insensitive) with a fake word. */
function deny(text: string, denylist: readonly string[] = []): string {
  let out = text;
  for (const term of denylist) {
    const pattern = new RegExp(term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi");
    out = out.replace(pattern, (match) => fakeWord(match));
  }
  return out;
}

/** Throws if any denylisted term or email address (other than the placeholder) remains. */
export function assertClean(text: string, denylist: readonly string[]): void {
  const lower = text.toLowerCase();
  const found = denylist.filter((term) => lower.includes(term.toLowerCase()));
  const emails = [...text.matchAll(EMAIL)]
    .map((m) => m[0])
    .filter((e) => !e.endsWith("example.com") && !ASSET_FILE.test(e));
  if (found.length || emails.length) {
    throw new Error(`scrub incomplete: ${[...found, ...emails].length} sensitive term(s) remain`);
  }
}

if (import.meta.main) {
  const { values, positionals } = parseArgs({
    allowPositionals: true,
    options: {
      name: { type: "string", default: "inbox" },
      denylist: { type: "string", default: "fixtures-raw/denylist.txt" },
    },
  });
  const [input, outDir] = positionals;
  if (!input || !outDir) {
    console.error("usage: node scripts/scrub-fixture.ts <saved.html> <out-dir> [--name inbox]");
    process.exit(2);
  }
  const denylist = existsSync(values.denylist)
    ? (await readFile(values.denylist, "utf8"))
        .split("\n")
        .map((l) => l.trim())
        .filter(Boolean)
    : [];
  const raw = await readFile(input, "utf8");
  // Chrome saves assets next to the page as "<page name>_files/".
  const assetDir = raw.match(/(?:href|src)="\.\/([^"/]+_files)\//)?.[1];
  const { html, assets, runtimeCss } = scrubHtml(raw, { assetDir, denylist, name: values.name });
  if (runtimeCss !== undefined) {
    await mkdir(path.join(outDir, "vendor"), { recursive: true });
    await writeFile(path.join(outDir, "vendor", `${values.name}.runtime.css`), runtimeCss);
  }

  const manifest = path.join(outDir, "assets.json");
  const known: string[] = existsSync(manifest) ? JSON.parse(await readFile(manifest, "utf8")) : [];
  const all = [...new Set([...known, ...assets])].sort();
  await writeFile(manifest, `${JSON.stringify(all, null, 2)}\n`);
  const output = path.join(outDir, `${values.name}.html`);
  await writeFile(output, html);
  console.log(`wrote ${output}; ${assets.length} stylesheets listed in ${manifest}`);
  console.log(
    `denylist: ${denylist.length} terms. Run \`mise run fixtures:assets\` to fetch styles.`,
  );
  console.log("Skim the output before committing: the scrubber is a safety net, not a guarantee.");
}
