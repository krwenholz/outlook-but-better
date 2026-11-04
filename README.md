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
Open Outlook and press `?` for the shortcut list. Click the toolbar icon for options (pin it from
Chrome's puzzle-piece menu). `g` jumps to a folder or to Mail, Calendar or People.

## Looks

Pick colours, light/dark and a layout on the options page. Palettes recolour Outlook through
its own colour variables, so everything matches, and "Follow Outlook" tracks its dark mode.

|                                                                                    |                                                                              |
| ---------------------------------------------------------------------------------- | ---------------------------------------------------------------------------- |
| ![Solarized light, calm](docs/screenshots/solarized-light-calm.png)                | ![Solarized dark, calm](docs/screenshots/solarized-dark-calm.png)            |
| ![Fastmail palette, Fastmail layout](docs/screenshots/fastmail-light-fastmail.png) | ![Fastmail dark, calm](docs/screenshots/fastmail-dark-calm.png)              |
| ![Base16 dark, Fastmail layout](docs/screenshots/base16-dark-fastmail.png)         | ![Base16 light, Fastmail layout](docs/screenshots/base16-light-fastmail.png) |
| ![Paper, focus](docs/screenshots/paper-light-focus.png)                            | ![Solarized dark, focus](docs/screenshots/solarized-dark-focus.png)          |

Screenshots come from scrubbed captures (`mise run screenshots`), so names and text are
placeholder words. The Fastmail and Focus layouts assume Outlook's "Reading pane:
Hide" setting.

## Develop

`mise run dev` rebuilds on change (reload the extension afterwards). `mise run check` runs
everything. The first time on a new machine, `mise run e2e:setup` downloads Chromium for the
end-to-end tests. See [AGENTS.md](AGENTS.md) for layout and rules.
