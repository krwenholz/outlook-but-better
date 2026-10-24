#!/bin/bash
# Prepares a Claude Code cloud session: installs mise, the pinned toolchain, and npm deps.
# Local machines: run `mise install && mise run install` yourself instead.
set -euo pipefail

if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

cd "$CLAUDE_PROJECT_DIR"

export PATH="$HOME/.local/bin:$HOME/.local/share/mise/shims:$PATH"
if ! command -v mise >/dev/null 2>&1; then
  curl -fsSL https://mise.run | sh
fi

mise trust --yes "$CLAUDE_PROJECT_DIR/mise.toml"
mise install
mise run install

# Playwright is pinned to the version whose Chromium is preinstalled in the cloud image; this is
# a no-op when that build is already present.
mise exec -- npx playwright install chromium >/dev/null 2>&1 || echo "playwright chromium install failed; e2e may not run" >&2

if [ -n "${CLAUDE_ENV_FILE:-}" ]; then
  echo "export PATH=\"$HOME/.local/bin:$HOME/.local/share/mise/shims:\$PATH\"" >> "$CLAUDE_ENV_FILE"
fi
