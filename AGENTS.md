# Overlook

A Manifest V3 Chrome extension that improves Outlook on the web: Fastmail-style shortcuts, less
clutter, a calmer look. It works only by modifying the page. Read `PLAN.md` for the why and the
phases, and `docs/decisions.md` before revisiting a settled choice.

## Commands (mise)

- `mise install && mise run install`: toolchain and deps (a SessionStart hook does this in cloud sessions)
- `mise run check`: lint, typecheck, unit tests, build, e2e. Run before every commit.
- `mise run dev`: rebuild `dist/` on change; load `dist/` unpacked in `chrome://extensions`
- `mise run fmt`: auto-fix lint and formatting
- `mise run scrub <raw.html> <out.html>` / `mise run health <page.html>`: fixture tooling
- `mise tasks` lists the rest

## Layout

- `src/selectors.ts`: the only file that knows Outlook's DOM. Tags matches with `data-overlook`.
- `src/styles/*.css`: target only `[data-overlook~=…]`, gated by `[data-overlook-features~=…]` on `<html>`.
- `src/keymap/`: `engine.ts` (pure), `actions.ts` (strategy table), `bindings.ts`, `execute.ts` (DOM side).
- `src/palette/`: `fuzzy.ts` (pure) and the folder picker UI.
- `src/ui/overlay.ts`: all of Overlook's own UI, inside one shadow root.
- `src/settings.ts`: feature toggles, `chrome.storage.sync`. A new toggle is one entry in `FEATURES`.
- `test/*.test.ts`: unit (vitest). `test/e2e/`: real Chromium with `dist/` loaded, pages served at an Outlook URL.

## Rules

- Selectors: IDs > ARIA roles/labels > semantic `data-*` > text. No generated class names; if
  unavoidable, comment `// FRAGILE:`. Record why in `docs/dom-survey.md`.
- No network requests, analytics, or remote code. Storage is preferences only.
- Never commit real mail content. Raw saved pages go in `fixtures-raw/` (gitignored); commit only
  scrubbed output in `test/fixtures/`.
- Ask before adding any manifest permission.
- Keep pure logic DOM-free and unit-tested. Keep the toolchain small.
- Write decisions between alternatives into `docs/decisions.md`.
