# OBB

A Manifest V3 Chrome extension that improves Outlook on the web: Fastmail-style shortcuts, less
clutter, a calmer look. It works only by modifying the page. Read `PLAN.md` for the why and the
phases, and `docs/decisions.md` before revisiting a settled choice.

## Commands (mise)

- `mise install && mise run install`: toolchain and deps (a SessionStart hook does this in cloud sessions)
- `mise run check`: lint, typecheck, unit tests, build, e2e. Run before every commit.
- `mise run dev`: rebuild `dist/` on change; load `dist/` unpacked in `chrome://extensions`
- `mise run fmt`: auto-fix lint and formatting
- `mise run screenshots [name…]` / `mise run preview`: README images and animated preview (preview needs ffmpeg)
- `mise run scrub` / `health` / `fixtures:assets`: fixture tooling (see `docs/capturing-pages.md`)
- `mise tasks` lists the rest

## Layout

- `src/selectors.ts`: the only file that knows Outlook's DOM. Tags matches with `data-overlook`.
- `src/styles/*.css`: target only `[data-overlook~=…]`, gated by `[data-overlook-features~=…]` or
  `[data-overlook-palette|mode|layout=…]` on `<html>`. Palettes map `--ol-*` onto Outlook's own variables.
- `src/keymap/`: `engine.ts` (pure), `actions.ts` (strategy table), `bindings.ts`, `execute.ts` (DOM side).
- `src/palette/`: `fuzzy.ts` (pure) and the folder picker UI.
- `src/ui/overlay.ts`: all of OBB's own UI, inside one shadow root.
- `src/settings.ts`: toggles (`FEATURES`) and choices (`CHOICES`) in `chrome.storage.sync`.
- `test/*.test.ts`: unit (vitest). `test/e2e/`: real Chromium with `dist/` loaded, fixtures served at an Outlook URL.
- `test/fixtures/`: scrubbed real pages; `shim.css` patches capture artifacts for screenshots only.

## Rules

- Selectors: IDs > ARIA roles/labels > semantic `data-*` > text. No generated class names; if
  unavoidable, comment `// FRAGILE:`. Record why in `docs/dom-survey.md`.
- No network requests, analytics, or remote code. Storage is preferences only.
- Never commit real mail content. Raw saved pages go in `fixtures-raw/` (gitignored); commit only
  scrubbed output in `test/fixtures/`.
- The repo may become public: never commit Microsoft's files (stylesheets, icons, scripts) or
  anything naming the user's employer. Fetch vendor files at build/test time instead.
- Ask before adding any manifest permission.
- Keep pure logic DOM-free and unit-tested. Keep the toolchain small.
- Write decisions between alternatives into `docs/decisions.md`.
