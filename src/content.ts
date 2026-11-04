// Entry point: applies settings, tags elements for CSS, runs the health check, mounts the keymap.

import { ACTIONS } from "./keymap/actions.ts";
import { BINDINGS } from "./keymap/bindings.ts";
import { DEFAULT_TIMEOUT_MS, keyFromEvent, SequenceEngine } from "./keymap/engine.ts";
import { currentContext, execute } from "./keymap/execute.ts";
import {
  brokenEntries,
  type HealthResult,
  healthCheck,
  outlookColorScheme,
  resolve,
  tagElements,
} from "./selectors.ts";
import { applyToDocument, loadSettings, onSettingsChanged, type Settings } from "./settings.ts";
import { isFromOverlay, isModalOpen, showHealthBadge, showHint } from "./ui/overlay.ts";

const LOG = "[obb]";
const APP_READY_TIMEOUT_MS = 20_000;
const HEALTH_SETTLE_MS = 1_000;
/** Waits between re-checks of a failed health check before it is reported. */
const HEALTH_RETRY_MS = [3_000, 10_000];

let settings: Settings;
let healthDismissed = false;

function applySettings(next: Settings): void {
  settings = next;
  applyToDocument(settings, document.documentElement);
  syncColorScheme();
  if (!settings.healthIndicator) showHealthBadge([], () => {});
}

/** Records whether Outlook shows a reading pane, so layouts only reshape panes when it doesn't. */
function syncReadingPane(): void {
  const state = resolve("readingPane") ? "shown" : "none";
  if (document.documentElement.dataset.overlookReadingPane !== state) {
    document.documentElement.dataset.overlookReadingPane = state;
  }
}

/** Mirrors Outlook's own light/dark mode onto <html> so "Follow Outlook" palettes can match it. */
function syncColorScheme(): void {
  const scheme = outlookColorScheme(document.documentElement);
  if (document.documentElement.getAttribute("data-overlook-outlook-scheme") !== scheme) {
    document.documentElement.setAttribute("data-overlook-outlook-scheme", scheme);
  }
}

function reportHealth(results: HealthResult[]): void {
  const broken = brokenEntries(results);
  const fallbacks = results.filter((r) => r.candidate > 0).map((r) => r.name);
  document.documentElement.dataset.overlookHealth = broken.length ? "broken" : "ok";
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

let healthTimer: ReturnType<typeof setTimeout> | undefined;

/**
 * Checks the registry once Outlook has settled. Parts of the page (the app rail, for one) render
 * after the message list, so a miss is re-checked a few times before it is reported.
 */
function scheduleHealthCheck(): void {
  clearTimeout(healthTimer);
  // Only the mail views are surveyed; Calendar and People legitimately lack mail elements.
  if (!location.pathname.includes("/mail")) return;
  const attempt = (n: number) => {
    const results = healthCheck();
    const retryIn = HEALTH_RETRY_MS[n];
    if (brokenEntries(results).length === 0 || retryIn === undefined) reportHealth(results);
    else healthTimer = setTimeout(() => attempt(n + 1), retryIn);
  };
  healthTimer = setTimeout(() => attempt(0), HEALTH_SETTLE_MS);
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
      syncColorScheme();
      syncReadingPane();
      if (location.href !== lastUrl) {
        lastUrl = location.href;
        scheduleHealthCheck();
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
  syncReadingPane();
  scheduleHealthCheck();
  watchPage();
}

void main();
