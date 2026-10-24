// Entry point: applies settings, tags elements for CSS, runs the health check, mounts the keymap.

import { ACTIONS } from "./keymap/actions.ts";
import { BINDINGS } from "./keymap/bindings.ts";
import { DEFAULT_TIMEOUT_MS, keyFromEvent, SequenceEngine } from "./keymap/engine.ts";
import { currentContext, execute } from "./keymap/execute.ts";
import { brokenEntries, healthCheck, resolve, tagElements } from "./selectors.ts";
import { featureTokens, loadSettings, onSettingsChanged, type Settings } from "./settings.ts";
import { isFromOverlay, isModalOpen, showHealthBadge, showHint } from "./ui/overlay.ts";

const LOG = "[overlook]";
const APP_READY_TIMEOUT_MS = 20_000;

let settings: Settings;
let healthDismissed = false;

function applySettings(next: Settings): void {
  settings = next;
  document.documentElement.setAttribute("data-overlook-features", featureTokens(settings));
  if (!settings.healthIndicator) showHealthBadge([], () => {});
}

function runHealthCheck(): void {
  const results = healthCheck();
  const broken = brokenEntries(results);
  const fallbacks = results.filter((r) => r.candidate > 0).map((r) => r.name);
  console.groupCollapsed(`${LOG} health: ${broken.length ? `${broken.length} broken` : "ok"}`);
  console.table(results);
  console.groupEnd();
  if (broken.length) console.warn(`${LOG} selectors not found:`, broken.join(", "));
  if (fallbacks.length) console.info(`${LOG} using fallback selectors for:`, fallbacks.join(", "));
  if (settings.healthIndicator && !healthDismissed) {
    showHealthBadge(broken, () => {
      healthDismissed = true;
    });
  }
}

/** Resolves once Outlook's mail UI has rendered, or after a timeout so the check still reports. */
function whenAppReady(): Promise<void> {
  return new Promise((done) => {
    if (resolve("messageList")) return done();
    const observer = new MutationObserver(() => {
      if (resolve("messageList")) finish();
    });
    const timer = setTimeout(finish, APP_READY_TIMEOUT_MS);
    function finish() {
      observer.disconnect();
      clearTimeout(timer);
      done();
    }
    observer.observe(document.body, { childList: true, subtree: true });
  });
}

/** Re-tag on DOM churn (React re-renders), and re-check health when the SPA navigates. */
function watchPage(): void {
  let scheduled = false;
  let lastUrl = location.href;
  new MutationObserver(() => {
    if (scheduled) return;
    scheduled = true;
    requestAnimationFrame(() => {
      scheduled = false;
      tagElements();
      if (location.href !== lastUrl) {
        lastUrl = location.href;
        runHealthCheck();
      }
    });
  }).observe(document.body, { childList: true, subtree: true });
}

function isEditable(target: EventTarget | undefined): boolean {
  if (!(target instanceof HTMLElement)) return false;
  return target.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(target.tagName);
}

function mountKeymap(): void {
  const engine = new SequenceEngine(BINDINGS);
  let hintTimer: ReturnType<typeof setTimeout> | undefined;
  window.addEventListener(
    "keydown",
    (event) => {
      if (!settings.keymap || !event.isTrusted || isFromOverlay(event) || isModalOpen()) return;
      if (isEditable(event.composedPath()[0])) return;
      const key = keyFromEvent(event);
      if (key === null) return;

      const result = engine.feed(key, currentContext(), event.timeStamp);
      if (result.kind === "none") return;
      showHint(
        result.kind === "pending"
          ? `${result.prefix.join(" ")} … ${result.next
              .map((b) => `${b.keys.at(-1)} ${ACTIONS[b.action].description}`)
              .join(" · ")}`
          : null,
      );
      clearTimeout(hintTimer);
      if (result.kind === "pending")
        hintTimer = setTimeout(() => showHint(null), DEFAULT_TIMEOUT_MS);
      if (result.kind === "match" && !execute(result.action)) return;
      event.preventDefault();
      event.stopImmediatePropagation();
    },
    true,
  );
}

async function main(): Promise<void> {
  applySettings(await loadSettings());
  onSettingsChanged(applySettings);
  mountKeymap();
  await whenAppReady();
  tagElements();
  runHealthCheck();
  watchPage();
}

void main();
