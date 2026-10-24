# Smoke test (under two minutes)

Run after any Outlook UI change or extension update.
1. Reload Outlook. No "selectors broken" badge in the top right. Console shows `[overlook] health: ok`.
2. Clutter toggles that are on are hidden, with no empty gaps left behind.
3. `j` / `k` move between conversations; `x` selects.
4. `y` archives the current conversation; `z` undoes it.
5. `g`, type part of a nested folder name, Enter: lands in that folder.
6. `.` shows the pending hint; `r` marks read.
7. `?` shows the shortcut list; any key closes it.
8. Typing `y` in the search box types a "y" (single-key shortcuts are off while editing).
9. Options: switch Colors to Solarized and Layout to Calm. Outlook recolours without reloading,
   including menus and dialogs; an open email stays readable.
