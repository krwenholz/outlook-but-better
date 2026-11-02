# DOM survey

Anchors for everything Overlook hides or drives. `src/selectors.ts` must agree with this file.
Source: a saved inbox page from Oct 2026 (`test/fixtures/inbox.html`), English UI, light mode,
Gmail keyboard shortcuts on.

## Anchors

| Registry name                   | Anchor                                                                                                                            | Fallback                                         | Notes                                                                                                                               |
| ------------------------------- | --------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------ | ----------------------------------------------------------------------------------------------------------------------------------- |
| titleBar                        | `#OwaTitleBar`                                                                                                                    | –                                                | Structural ID.                                                                                                                      |
| searchBox                       | `#topSearchInput`                                                                                                                 | `[role=search] input`                            |                                                                                                                                     |
| appRail                         | `#LeftRail`                                                                                                                       | `[role=navigation][aria-label=left-rail-appbar]` |                                                                                                                                     |
| ribbon                          | `[data-app-section=Ribbon]`                                                                                                       | `[role=region][aria-label=Ribbon]`               | `data-app-section` names every major region.                                                                                        |
| ribbonTabs                      | `[role=tablist][aria-label="Ribbon tabs"]`                                                                                        | –                                                |                                                                                                                                     |
| folderPane                      | `[data-app-section=NavigationPane]`                                                                                               | `[role=region][aria-label="Navigation pane"]`    |                                                                                                                                     |
| folderItem                      | `[role=treeitem][data-folder-name]`                                                                                               | `[role=treeitem]`                                | `data-folder-name` is the lowercased name; `title` adds counts ("Inbox - 9 items (1 unread)"). Favorites duplicate mailbox folders. |
| listPane                        | `[data-app-section=MessageList]`                                                                                                  | –                                                | Includes Focused/Other tabs and filter buttons.                                                                                     |
| messageList                     | `#MailList`                                                                                                                       | `[role=listbox][aria-label="Message list"]`      | Virtualised (react-virtuoso).                                                                                                       |
| messageRow                      | `[role=option][data-convid]`                                                                                                      | `[role=option]`                                  | Row internals have only generated classes; sender/subject/date are not addressable.                                                 |
| groupHeader                     | `[role=button][id^=groupHeader]`                                                                                                  | –                                                | "This Month", "Last Month", month names.                                                                                            |
| readingPane                     | `[data-app-section=MailReadCompose]`                                                                                              | `#ReadingPaneContainerId`                        |                                                                                                                                     |
| messageBody                     | `[role=document][aria-label="Message body"]`                                                                                      | –                                                |                                                                                                                                     |
| infoBar                         | `[role=group]:has(> * > [id^=infoBar])`                                                                                           | –                                                | Inline-coloured notices above a message.                                                                                            |
| upsellBanner                    | `[data-app-section=UpsellBannerSection]`                                                                                          | –                                                | Present but empty in this capture.                                                                                                  |
| appBanner                       | `#mainApp > div > [role=group].fui-MessageBar`                                                                                    | –                                                | Uses Fluent's public static class.                                                                                                  |
| headerExtras                    | `#owaTeamsBtn_container`, `#owaTimePanelBtn_container`                                                                            | –                                                | Teams, My Day.                                                                                                                      |
| appRailExtras                   | `[role=none][aria-label="To Do"]` etc.                                                                                            | –                                                | Wrappers carry the app name.                                                                                                        |
| ribbonExtras                    | `[aria-label="Share to Teams"]`, Sweep, Assign policy, Viva Insights, More apps, Quick steps                                      | –                                                |                                                                                                                                     |
| rowAvatar                       | `[role=img][id^=avatar-]`                                                                                                         | –                                                |                                                                                                                                     |
| rowHoverActions                 | `button[aria-label="Flag this message"]` etc.                                                                                     | –                                                |                                                                                                                                     |
| composeButton … undoButton      | `button[aria-label="New mail"]`, `"Archive"`, `"Delete"`, `"Move to"`, `"Snooze"`, `"Flag / Unflag"`, `"Read / Unread"`, `"Undo"` | –                                                | Ribbon buttons.                                                                                                                     |
| syncButton, settingsButton      | `#owaSyncButton_container`, `#owaSettingsBtn_container`                                                                           | –                                                |                                                                                                                                     |
| mailApp, calendarApp, peopleApp | `button[aria-label=Mail]` etc. inside the app rail                                                                                | –                                                |                                                                                                                                     |

Not seen in this capture (still guesses): Copilot surfaces, External tags, a Report/Junk button
(this tenant has a "Phish Alert" add-in instead).

## Nov 2026 captures

`list.html` and `move-menu.html` use the simplified (single-row) ribbon; every anchor above still
resolves. The ribbon's "Move to" opens a menu with its own `Search for a folder` box, recent
folders, and "Select a different folder", so `m` (which clicks it) gets type-ahead for free.

## Outlook's Gmail-mode keys

From the ribbon's own tooltips: `C` new mail, `#` delete, `E` archive, `V` move to folder,
`⇧I`/`⇧U` read/unread, `B` snooze, `L` categorize, `R`/`A`/`F` reply/reply all/forward,
`⌘Z` undo.

## Dark mode

Outlook writes `--native-scrollbar-color-scheme: light|dark` among the palette variables on
`<html>`. Overlook mirrors it to `data-overlook-outlook-scheme` for "Follow Outlook" palettes.
Unverified in dark mode.

## Open questions

- Does Outlook act on synthetic (`isTrusted: false`) key events? Matters for `. r`/`. u`.
- Are collapsed subfolders present in the DOM? Matters for the `g` palette.
