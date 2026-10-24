# DOM survey

Phase 0 output. For every element Overlook hides or drives: the chosen anchor, why, and the
fallback. `src/selectors.ts` must agree with this file.

Status: **not started**. Everything in `selectors.ts` is a guess until this table is filled in.

| Registry name              | Anchor (chosen) | Fallback                                     | Why / notes                                                           | Verified |
| -------------------------- | --------------- | -------------------------------------------- | --------------------------------------------------------------------- | -------- |
| messageList                | `#MailList`     | `[role=listbox][aria-label*="Message list"]` |                                                                       | ☐        |
| readingPane                |                 |                                              |                                                                       | ☐        |
| folderPane                 |                 |                                              |                                                                       | ☐        |
| folderItem                 |                 |                                              | Are collapsed subfolders in the DOM? Does text include unread counts? | ☐        |
| searchBox                  |                 |                                              |                                                                       | ☐        |
| appRail                    |                 |                                              |                                                                       | ☐        |
| ribbon                     |                 |                                              |                                                                       | ☐        |
| copilot                    |                 |                                              |                                                                       | ☐        |
| promoBanner                |                 |                                              |                                                                       | ☐        |
| externalTag                |                 |                                              |                                                                       | ☐        |
| composeButton … undoButton |                 |                                              | Toolbar labels; check with and without a selection                    | ☐        |

## Synthetic keyboard events

Does Outlook act on `isTrusted: false` key events? From the console with Gmail shortcuts on:

```js
document.activeElement.dispatchEvent(new KeyboardEvent("keydown", { key: "j", bubbles: true }));
```

Result: _unknown_.

## Clutter screenshots

Redacted screenshots of each banner and clutter item go here.
