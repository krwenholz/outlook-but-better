# Fixtures

Scrubbed snapshots of real Outlook pages. Every `.html` file here is loaded by the e2e suite and
must pass the selector health check. See `docs/capturing-pages.md` for how to add one.

| Fixture | Captured | State |
|---|---|---|
| `inbox.html` | Oct 2026, "Save page" | Inbox, a meeting invite open in the reading pane, classic ribbon |
| `list.html` | Nov 2026, DevTools snippet | Inbox, nothing selected, simplified ribbon |
| `move-menu.html` | Nov 2026, DevTools snippet | A message open with the ribbon's "Move to" menu showing |

- `vendor/`: Outlook's runtime styles from snippet captures. Microsoft's, not ours; see
  `vendor/NOTICE.md`.
- `assets/` (gitignored): Outlook's static stylesheets, fetched by `mise run fixtures:assets`.
- `shim.css` / `shim-saved.css`: test-only patches for capture artifacts; the second only for
  "Save page" captures.
