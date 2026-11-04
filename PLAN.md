# OBB: a Chrome extension that makes Outlook on the web tolerable

"OBB" (outlook-but-better) is the name; it was "Overlook" while this plan was drafted.

## Why this exists

Kyle's work email is moving from Gmail to Microsoft 365. The tenant almost certainly enforces a daily sign-in frequency through Conditional Access. That makes Apple Mail (via macOS Internet Accounts) a second daily sign-in on top of the browser sign-in he already does for other work apps. The plan is to drop Apple Mail and live in Outlook on the web, installed as a Chrome app so it gets its own window and Dock icon. Then fix the parts of Outlook on the web that are bad.

The three problems to solve, in priority order:
1. **Keyboard shortcuts.** Kyle's muscle memory comes from Fastmail and Gmail. Outlook's built-in Gmail shortcut mode is a reasonable base, but it lacks Fastmail's best commands, especially the fuzzy "go to folder" and "move to folder" pickers.
2. **Clutter.** Banners, promos, Copilot surfaces, callouts, "External" tags, the app rail, and the ribbon all take up space and attention.
3. **Look and feel.** The interface is blocky and low-density. It should feel closer to Fastmail: tighter, calmer, and more typographic.

Calendar is explicitly out of scope until Kyle has used it for a week or two and knows what actually bothers him.

## Constraints that shape the design

**No Microsoft Graph API.** Any app that calls Graph on Kyle's behalf needs an Entra app registration. Most enterprise tenants require admin consent for that, and a request from an individual engineer for a personal mail tool will not get approved. Everything in this extension therefore works by modifying the Outlook web page itself. It never makes its own network calls.

**The DOM is hostile.** Outlook on the web is a large React app with generated, obfuscated class names that change without notice. Existing userscripts that hide elements with deep `nth-child` selector chains break quietly and often. This extension must anchor on the most stable hooks available, in this order of preference:
1. Element IDs that Microsoft clearly treats as structural, such as `#MailList` and the folder pane container.
2. ARIA roles and `aria-label` values. Accessibility labels change far less often than class names, because changing them breaks screen-reader users and automated accessibility tests.
3. `data-*` attributes that look semantic rather than generated.
4. Text content, as a last resort.

Class names should never be used as selectors unless there is no alternative, and any such selector should be flagged in code as fragile.

**Domains.** Microsoft has moved Outlook on the web to `outlook.cloud.microsoft`, with `outlook.office.com` and `outlook.office365.com` redirecting to it. Host permissions should cover all three and nothing else.

**Corporate policy on the Mac.** If the work Mac's Chrome is managed, policy may block unpacked extensions, developer mode, or extensions outside an allowlist. This is the single biggest risk to the project, so it is checked first (Phase 0). If unpacked extensions are blocked, the fallback is to see whether Stylus and a userscript manager such as Violentmonkey are allowed. The architecture below keeps the CSS layer separable so it can ship as a plain user style if needed.

**Data handling.** The extension handles work email, so it must not send data anywhere. That means no network requests, no analytics, no remote code, and storage limited to user preferences in `chrome.storage.sync`. DOM snapshots captured for development must never be committed with real mail content in them (see Testing).

## Architecture

The extension is a Manifest V3 extension with a content script, a small options page, and no background logic unless a later phase needs it. TypeScript, bundled with esbuild or Vite. Keep the toolchain small, because the build only needs to produce a content script, a CSS file, and an options page.

It has four layers. Keeping them separate means that when Microsoft changes the page, the breakage is confined to one small, obvious place.

### 1. Selector registry (`src/selectors.ts`)

This is the only file allowed to know what Outlook's DOM looks like. Every element the extension touches has a named entry here, such as `messageList`, `readingPane`, `folderPane`, `archiveButton`, or `topBanner`, with one or more candidate selectors in priority order.

The registry also runs a **selector health check** on load and on major navigation. It reports which named entries currently resolve and which do not. Failures are logged to the console with a clear prefix, and a small, dismissible indicator appears on the page so Kyle notices when something has broken rather than discovering it through mysterious behaviour. This health check is the main defence against Microsoft's UI churn, since there is no way to run Outlook in CI.

### 2. Style layer (`src/styles/`)

Plain CSS, injected by the content script, organised into independent modules that can each be toggled from the options page:
- **Declutter:** hides banners, promos, Copilot entry points, "try the new..." callouts, External tags, and other noise. Each hidden element is its own toggle so a false positive can be switched off without disabling the rest.
- **Density:** tightens row height, padding, and whitespace in the message list, folder pane, and reading pane.
- **Theme:** a calmer visual treatment, including typography, flatter surfaces, and fewer borders, defined with CSS custom properties so it can be adjusted in one place. It must respect Outlook's own light and dark modes rather than fight them.

The CSS should reference elements through the selector registry wherever possible. One way to do that is to have the content script add stable `data-overlook="message-list"`-style attributes to the elements the registry finds, and then write all CSS against those attributes. That keeps the stylesheets readable and pushes all DOM knowledge into one file.

### 3. Keymap layer (`src/keymap/`)

A small key-sequence engine that implements Fastmail-style bindings on top of Outlook.
- **Context awareness.** Bindings depend on where the user is: the message list, an open conversation, compose, or a picker. All single-key bindings are ignored while focus is inside any editable element.
- **Sequences.** Some Fastmail commands are two keys, such as `Shift+G` then `M`, `.` then `R`, or `*` then `A`. The engine needs pending-prefix state, a short timeout, and a visible hint while a prefix is pending.
- **Execution strategies.** Each action resolves to one of three strategies, chosen per action during Phase 1:
  1. **Pass through:** the key already does the right thing in Outlook's Gmail mode, so the extension does nothing.
  2. **Translate:** the extension triggers Outlook's own shortcut for the action, for example turning `y` into Outlook's Gmail-mode archive key. This is the preferred strategy because it relies on Outlook's own behaviour. It only works if Outlook responds to synthetic keyboard events. Synthetic events have `isTrusted: false`, and some apps ignore them, so this must be tested in the Phase 1 spike.
  3. **Drive the UI:** the extension clicks the appropriate button or menu item, found through the selector registry by its `aria-label`. This is the fallback when translation does not work.
- **Pure core.** The sequence engine and the action table should be pure TypeScript with no DOM dependency, so they can be unit tested properly.

### 4. Folder palette (`src/palette/`)

This is the main new feature. Fastmail's `g` (go to folder) and `m` (move to folder) open a type-ahead picker. Outlook has nothing equivalent from the keyboard. The palette is an overlay with fuzzy matching over the folder list, arrow-key and Enter navigation, and Escape to close.

The open question is where the folder list comes from. There are two candidate approaches, and the Phase 1 spike should decide between them:
- **Read the folder pane.** Scrape folder names and their clickable elements from the folder pane. The risk is that collapsed subfolders may not be in the DOM, so the extension may need to expand the tree once or cache folder names as they appear.
- **Drive Outlook's own "Move to" dialog.** Outlook's move dialog has its own folder search. The palette could open that dialog and type into its search field. This is less pleasant but may be more robust for nested folders.

For "go to", selecting a folder clicks it in the folder pane. For "move", selecting a folder applies the move to the current selection or the open conversation.

## Shortcut map

This is the target binding set. Kyle should confirm it against his own Fastmail muscle memory before Phase 2, since some of these come from Fastmail's documentation rather than from watching him use it. The mechanism column is filled in during the Phase 1 spike.

| Key                           | Action                                      | Context       | Mechanism                                     |
| ----------------------------- | ------------------------------------------- | ------------- | --------------------------------------------- |
| `c`                           | Compose                                     | Anywhere      | Pass through (Outlook `C`)                    |
| `/`                           | Focus search                                | Anywhere      | Pass through                                  |
| `?`                           | Show OBB's shortcut help                    | Anywhere      | Extension overlay                             |
| `Esc`                         | Close, cancel, or return to list            | Anywhere      | Pass through                                  |
| `j` / `k`                     | Next / previous conversation                | List, reading | Pass through                                  |
| `o` or `Enter`                | Open conversation                           | List          | Pass through                                  |
| `x`                           | Select or deselect                          | List          | Pass through                                  |
| `Shift+X`                     | Select range since last selection           | List          | Likely extension logic                        |
| `* A` / `* N`                 | Select all / none                           | List          | Not yet                                       |
| `y`                           | Archive                                     | List, reading | Drive Archive button, else `E`                |
| `]` / `[`                     | Archive, then open previous / next          | Reading       | Composite                                     |
| `d`                           | Delete                                      | List, reading | Drive Delete button, else `#`                 |
| `!`                           | Report spam / junk                          | List, reading | Pass through (unverified)                     |
| `m`                           | Move to folder (palette)                    | List, reading | Drive Outlook's Move to picker, else `V`      |
| `g`                           | Go to folder (palette)                      | List, reading | Palette (folders plus Mail, Calendar, People) |
| `b`                           | Snooze                                      | List, reading | Pass through (Outlook `B`)                    |
| `s`                           | Toggle pin / flag                           | List, reading | Dropped (not wanted)                          |
| `. R` / `. U`                 | Mark read / unread                          | List, reading | Drive Read / Unread toggle by row state       |
| `r` / `a` / `f`               | Reply / reply all / forward                 | Reading       | Pass through                                  |
| `z`                           | Undo                                        | Anywhere      | Drive Undo button                             |
| `u`                           | Refresh                                     | Anywhere      | Drive Sync button                             |
| `Shift+G M` / `C` / `A` / `S` | Go to Mail / Calendar / Contacts / Settings | Anywhere      | Drive app rail / Settings                     |

## Settings-first checklist (do before writing code)

Some of the problems can be fixed with Outlook's own settings, and settings never break. Apply these first, then build only for what remains. Record the final settings in `docs/outlook-settings.md` so they can be reapplied if the account is ever reset.
- Keyboard shortcuts set to Gmail (Settings → General → Accessibility).
- Layout density set to the most compact option.
- Reading pane position and conversation view chosen deliberately.
- Focused Inbox turned off, unless Kyle wants it.
- Simplified ribbon, and any options to hide the ribbon or app rail.
- Notification and sound settings.
- Any per-feature toggles for Copilot or promotional surfaces that the tenant allows users to change.

## Phases

Each phase ends with something Kyle can use. Exit criteria are listed so the Claude Code session knows when a phase is done.

### Phase 0: Feasibility and DOM survey

- Check `chrome://policy` and `chrome://extensions` on the work Mac to confirm whether unpacked extensions and developer mode are allowed. If they are not, stop and switch to the Stylus and userscript fallback before doing anything else.
- Apply the settings-first checklist.
- Using Chrome access, survey the live Outlook page. For each element Kyle wants hidden or controlled, record its best available stable anchor (ID, role, `aria-label`, `data-*`) in `docs/dom-survey.md`. Include screenshots, with mail content redacted, of every banner and piece of clutter.
- Check whether Outlook responds to synthetic keyboard events. Dispatch Gmail-mode keys from the console and see whether the actions fire.

**Exit criteria:** policy status is known, the settings are applied, `docs/dom-survey.md` covers every declutter target and every control the keymap needs, and the synthetic-event question has a yes or no answer.

### Phase 1: Skeleton and spikes

- Scaffold the repo: Manifest V3, TypeScript, build script, content script, options page, and the selector registry with health check.
- Spike the keymap: implement `j`, `k`, `y`, and `x` using whichever strategies Phase 0 showed to be viable, and fill in the mechanism column of the shortcut map.
- Spike the palette: prove one of the two folder-list approaches against Kyle's real folder structure, including nested folders.

**Exit criteria:** the extension loads unpacked, the health check reports correctly, four keys work, and a decision on the palette approach is written down in `docs/decisions.md` with the reasoning.

### Phase 2: Keyboard parity

- Implement the full shortcut map, including sequences and the pending-prefix hint.
- Ship the folder palette for `g` and `m`.
- Add a `?` help overlay listing OBB's bindings.
- Unit tests for the sequence engine and action table.

**Exit criteria:** Kyle can triage a normal day of mail without touching the mouse.

### Phase 3: Declutter and density

- Implement the declutter modules, one toggle per target from the DOM survey.
- Implement the density module.
- Options page with all toggles, stored in `chrome.storage.sync`.

**Exit criteria:** every item Kyle flagged as clutter in Phase 0 can be hidden, and none of the hiding breaks layout (no leftover empty gaps where something was removed).

### Phase 4: Look and feel

- Implement the theme module with custom properties, working in both light and dark mode.
- Iterate with Kyle on screenshots until it feels right.

**Exit criteria:** Kyle prefers it to stock Outlook without reservation.

### Later, if wanted

- Calendar improvements, once Kyle knows what he needs.
- A command palette (`⌘K`) for actions beyond folders.
- Distributing to colleagues, which would raise the bar on robustness and probably require going through IT.

## Testing

There is no way to run Outlook on the web in CI, so testing is split by layer.
- **Pure logic** (sequence engine, action table, fuzzy matcher) gets ordinary unit tests.
- **Selectors** are protected mainly by the runtime health check. Optionally, sanitised DOM fixtures can be committed under `test/fixtures/` and tested with jsdom or Playwright. Fixtures must be scrubbed of all names, addresses, subjects, and message bodies before commit. A small script that replaces text nodes with placeholder text would make this safe and repeatable.
- **Manual smoke test**: a short checklist in `docs/smoke-test.md` that Kyle runs after any Outlook UI change or extension update. It should take under two minutes.

## Suggested repo layout

```
outlook-but-better/
  manifest.json
  src/
    content.ts          entry point: runs health check, tags elements, mounts layers
    selectors.ts        the only file that knows Outlook's DOM
    keymap/
      engine.ts         pure sequence engine
      actions.ts        action table and execution strategies
      bindings.ts       the Fastmail-style binding set
    palette/
      palette.ts
      fuzzy.ts
    styles/
      declutter.css
      density.css
      theme.css
    options/
      options.html
      options.ts
  test/
  docs/
    dom-survey.md
    decisions.md
    outlook-settings.md
    smoke-test.md
  PLAN.md
```

## Notes for the Claude Code session

- Start with Phase 0. Do not scaffold code until the policy check and the DOM survey are done, because both can change the plan.
- Use Chrome access to inspect the live page, but do not read, quote, or store mail content beyond what is needed to identify elements. Redact screenshots.
- When choosing a selector, record in `docs/dom-survey.md` why it was chosen and what the fallback is.
- When a decision is made between alternatives, write it in `docs/decisions.md` with the reasoning, so a future session does not relitigate it.
- Ask Kyle before adding any new permission to the manifest.

## Open questions for Kyle

1. Does the work Mac allow unpacked Chrome extensions? (Phase 0 answers this.)
2. ~~Should Fastmail's "pin" map to Outlook's Pin or to Flag?~~ Neither: `s` is dropped.
3. Is there anything in Outlook's look that Kyle likes and wants preserved?
4. Should the extension ever be shareable with colleagues, or is it strictly personal? This affects how much robustness and polish is worth investing.
