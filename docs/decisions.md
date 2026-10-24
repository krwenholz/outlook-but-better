# Decisions

Settled choices and why, so they are not relitigated. Newest first.

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

## 2026-10-24: Overlook's UI lives in a shadow root

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
