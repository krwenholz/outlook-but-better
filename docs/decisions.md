# Decisions

Settled choices and why, so they are not relitigated. Newest first.

## 2026-11-05: Installs and updates from the `dist` branch

Chrome only auto-updates extensions from the Web Store or ones force-installed by policy, so an
unpacked install needs some step to update. To make that step tiny:
- `mise run release` commits the build to a `dist` branch (fast-forward, never force-pushed). An
  install is a clone of that branch; an update is `git pull` plus **Reload OBB** in the popup
  (`chrome.runtime.reload()`). CI does this after every green push to main and publishes a
  GitHub Release with a zip when the manifest version is new.
- The manifest carries a public `key`, which fixes the extension id regardless of folder, so
  `chrome.storage.sync` settings survive updates and moves. The private key isn't kept; it's
  only needed to pack a .crx, which unpacked installs don't use. The Web Store rejects a
  manifest `key`, so strip it if OBB is ever listed there.
- Tabs left open across a reload keep the old content script, cut off from the extension; it
  shows a one-time "reload this tab" toast instead of half-working.

An unlisted Chrome Web Store listing would make updates fully automatic, at the cost of a
developer account and review. Revisit if this is shared with colleagues.

## 2026-11-04: Renamed to OBB

The user-facing name is OBB (outlook-but-better). Internal identifiers keep the drafting name:
`data-overlook-*` attributes, `--ol-*` tokens and `<overlook-root>`. Renaming them would only
churn every stylesheet and test, and nobody sees them.

## 2026-11-04: No synthetic keystrokes; read/unread drives the button

On the live site, `. u` (dispatching Outlook's Gmail-mode Shift+U) did nothing, so OBB no
longer sends keystrokes at all. Mark read/unread click the ribbon's Read / Unread toggle, and
only when the selected rows (judged by the "Unread" that starts their labels) aren't already in
the wanted state. Drive actions show a toast when their button is missing instead of falling
back to a key.

## 2026-11-04: Health check waits for the page to settle

The app rail renders after the message list, so the first live check reported its buttons
missing. The check now runs a second after Outlook is ready, re-checks a miss after 3s and 10s,
reports only the final result, and skips non-mail views (Calendar, People).

## 2026-11-04: Options in the toolbar popup; apps in the `g` palette

The options page doubles as the toolbar popup (`action.default_popup`; no new permission).
Mail, Calendar and People are listed in the `g` palette so switching apps is `g` "cal" Enter;
`Shift+G` then `m`/`c`/`a` still work.

## 2026-11-02: No pin binding

Kyle doesn't use Fastmail's pin, so `s` is unbound rather than mapped to Outlook's Pin or Flag.

## 2026-11-02: Runtime styles from captures live in `test/fixtures/vendor/`

DevTools-snippet captures carry Outlook's runtime (CSS-in-JS) rules, which can't be fetched from
anywhere, so they are committed, apart from the page, under `test/fixtures/vendor/` with a
NOTICE: Microsoft's work, outside the MIT license, test use only, removed on request. A notice
doesn't grant rights, but it is honest about what the files are and keeps them separable. The
large static stylesheets are still fetched from Microsoft's CDN rather than committed. Commit
`c5ad72d`, which once held those stylesheets, was purged from history on 2026-11-02.

## 2026-10-24: The repo may go public, so nothing Microsoft-owned or personal is committed

- Outlook's stylesheets are listed in `test/fixtures/assets.json` and downloaded from Microsoft's
  CDN by `mise run fixtures:assets` into a gitignored folder. Icons are dropped.
- Fixtures go through `scripts/scrub-fixture.ts`; a gitignored `fixtures-raw/denylist.txt` makes
  the scrub fail if any listed term (employer, names) survives.

## 2026-10-24: Fixture shim for screenshots

Saved pages lose Fluent's runtime CSS-in-JS rules, so only ~25% of element classes have styles.
`test/fixtures/shim.css` (fixture-only, never shipped) approximates the few that matter: pane
geometry, ribbon rows, hidden portals and tooltip text. Icon glyphs are hidden because Outlook's
icon-font codepoints don't match the public `@fluentui/react-icons` fonts. The DevTools snippet
in `docs/capturing-pages.md` captures runtime rules too, which should make the shim unnecessary.

## 2026-10-24: Palettes recolour Outlook through its own variables

Outlook paints itself from the Fabric palette on `<html>` (`--white`, `--neutral*`, `--theme*`)
and Fluent v9 tokens (`--colorNeutral*`, `--colorBrand*`). Palettes set OBB tokens
(`--ol-*`) and map them onto both with `!important`, so surfaces OBB never tags still match.
Fluent re-declares tokens on provider elements, targeted via its public `fui-FluentProvider`
class. Dark palettes show email bodies as a light "letter" because bodies carry their own colours.

## 2026-10-24: Layouts never hide the reading pane

Hiding Outlook's reading pane with CSS would make opened messages invisible. The list-first
layouts (fastmail, focus) instead pair with Outlook's own "Reading pane: Hide" setting; the
screenshot script simulates that setting.

## 2026-10-24: Drive Outlook's buttons first, keys second

React honours `element.click()`, while synthetic key events are still unverified. Actions with a
ribbon button (archive, delete, move, undo, refresh) click it and fall back to Outlook's
Gmail-mode key when the button is absent. Read/unread uses Shift+I/Shift+U, since the ribbon
only has a toggle.

## 2026-10-24: `m` opens Outlook's own "Move to" picker

Outlook's picker already has folder search. The folder palette handles `g` only; moving from
the palette can come later if Outlook's picker proves too slow.

## 2026-10-24: Scaffolded before Phase 0

PLAN.md says to finish the policy check and DOM survey first. Kyle asked for the scaffold now. All
selectors and keymap strategies are therefore marked unverified, and Phase 0 findings should
replace them rather than be bent to fit them.

## 2026-10-24: Plain esbuild, not WXT/CRXJS/Plasmo

WXT is the common default for MV3 extensions in 2026, but the build only needs to produce one
content script, one stylesheet, and an options page. A ~50-line `scripts/build.ts` keeps the
toolchain small and transparent, and makes a Stylus/userscript fallback (if Chrome policy blocks
unpacked extensions) easy to produce from the same sources. Cost: no HMR or auto-reload; reload
the extension by hand.

## 2026-10-24: Tooling

- mise pins Node and defines every command, so humans, CI and Claude sessions run the same thing.
- Node 24 runs `scripts/*.ts` directly (type stripping), so no ts-node/tsx. Source imports use
  explicit `.ts` extensions for the same reason.
- Biome for lint and format (one dependency instead of ESLint + Prettier). `!important` is
  allowed because overriding Outlook's styles needs it.
- Vitest for unit tests (jsdom for selector tests). Playwright, pinned to the version whose
  Chromium is preinstalled in Claude cloud sessions, for e2e.

## 2026-10-24: CSS targets data-overlook tokens

`selectors.ts` adds `data-overlook="message-list"`-style tokens to matched elements, and
`<html data-overlook-features="…">` lists enabled toggles. Stylesheets use only those two hooks,
so DOM knowledge stays in one file and every toggle is pure CSS. A MutationObserver re-tags on
React re-renders (batched per animation frame).

## 2026-10-24: OBB's UI lives in a shadow root

The palette, help, hints and health badge render inside one `<overlook-root>` shadow root so
Outlook's CSS cannot restyle them and ours cannot leak out. The keymap ignores events from inside
it.

## 2026-10-24: Synthetic key events are ignored by our own listener

The translate strategy dispatches Outlook's shortcut as a synthetic `KeyboardEvent`. Our listener
skips `isTrusted: false` events, so translations never loop. Whether Outlook acts on them is still
the open Phase 0 question.

## 2026-10-24: Fixture tests via a saved page

Real pages are saved to gitignored `fixtures-raw/`, scrubbed by `scripts/scrub-fixture.ts`
(masks all text and attribute values except those the registry matches on), and committed to
`test/fixtures/`. The e2e suite serves each fixture at an Outlook URL with the built extension
loaded and fails if the health check reports breakage.
