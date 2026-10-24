# Fixtures

Scrubbed snapshots of real Outlook pages. Every `.html` file here is loaded by the e2e suite and
must pass the selector health check.

To add one:

1. In Chrome on Outlook, save the page (or copy `document.documentElement.outerHTML` from the
   console) into `fixtures-raw/` at the repo root. That directory is gitignored. Never commit it.
2. `mise run scrub fixtures-raw/inbox.html test/fixtures/inbox.html`
3. Skim the output for anything personal before committing. The scrubber masks all text and
   labels except the ones the selector registry depends on, but treat it as a safety net.
