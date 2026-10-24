# Overlook (outlook-but-better)

A Chrome extension that makes Outlook on the web tolerable: Fastmail-style keyboard shortcuts with
a fuzzy folder picker, less clutter, and a denser, calmer look. Everything happens in the page; it
makes no network requests and needs no Microsoft Graph access.

See [PLAN.md](PLAN.md) for the design and roadmap.

## Try it

```sh
mise install          # Node, pinned in mise.toml
mise run install      # npm deps
mise run build        # writes dist/
```

Then open `chrome://extensions`, turn on Developer mode, choose **Load unpacked**, and pick `dist/`.
Open Outlook and press `?` for the shortcut list. Options live on the extension's details page.

## Develop

`mise run dev` rebuilds on change (reload the extension afterwards). `mise run check` runs
everything. The first time on a new machine, `mise run e2e:setup` downloads Chromium for the
end-to-end tests. See [AGENTS.md](AGENTS.md) for layout and rules.
