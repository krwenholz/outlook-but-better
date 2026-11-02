# Capturing an Outlook page for fixtures

Fixtures are scrubbed snapshots of real Outlook pages (`test/fixtures/`). Tests and screenshots
run against them, so a fresh capture is the way to pick up Outlook UI changes without a live
session.

## Best: the DevTools snippet (keeps runtime styles)

Chrome's "Save page as…" loses the styles Fluent injects at runtime, so saved pages render
half-styled. Instead, with Outlook open in the state you want (a message selected, a menu open,
dark mode…), open DevTools → Console, paste this, and press Enter. It downloads one HTML file.

```js
(() => {
  const sheets = [...document.styleSheets, ...document.adoptedStyleSheets].filter((s) => !s.href);
  const rules = sheets.flatMap((s) => {
    try {
      return [...s.cssRules].map((r) => r.cssText);
    } catch {
      return [];
    }
  });
  const style = `<style data-overlook-runtime>${rules.join("\n")}</style>`;
  const html = `<!doctype html>${document.documentElement.outerHTML.replace("</head>", `${style}</head>`)}`;
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([html], { type: "text/html" }));
  a.download = `outlook-${new Date().toISOString().slice(0, 16).replace(/\D/g, "")}.html`;
  a.click();
})();
```

"Save page as… → Webpage, Complete" also works; you get the `_files` folder too.

## Scrub and commit

1. Put the file in `fixtures-raw/` (gitignored; never commit it).
2. Optional: list extra words that must never appear (employer, colleagues, clients) in
   `fixtures-raw/denylist.txt`, one per line. The scrub fails if any survive.
3. `mise run scrub fixtures-raw/<file>.html test/fixtures --name <state>`
   (e.g. `inbox`, `reading-dark`, `compose`).
4. `mise run health test/fixtures/<state>.html` to see which selectors resolve.
5. `mise run screenshots` to look at it, skim the HTML, then commit.

The scrubber keeps structure, roles, semantic `data-*` values and Outlook's own UI words, and
replaces everything else (names, subjects, bodies, addresses, phone numbers, ids, links) with
deterministic fake words. A snippet capture's runtime styles go to
`test/fixtures/vendor/<state>.runtime.css` (see the NOTICE there). Outlook's static stylesheets
are not committed; `mise run fixtures:assets` fetches them from Microsoft's CDN. Screenshots also
load fonts and icons from that CDN, so they need network access; tests don't.
