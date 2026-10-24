// Pure key-sequence engine. No DOM access, so it is fully unit-testable.
//
// Keys are KeyboardEvent.key values, which already reflect Shift ("G", "*", "!", "?").
// A binding like Fastmail's "Shift+G then M" is written ["G", "m"].

export type Context = "list" | "reading" | "palette";

export interface Binding<A extends string = string> {
  keys: readonly string[];
  action: A;
  contexts: readonly Context[];
}

export type FeedResult<A extends string> =
  | { kind: "match"; action: A }
  | { kind: "pending"; prefix: readonly string[]; next: readonly Binding<A>[] }
  /** A prefix was pending and this key completed nothing. The key should be swallowed. */
  | { kind: "cancelled" }
  | { kind: "none" };

export const DEFAULT_TIMEOUT_MS = 1500;

const startsWith = (keys: readonly string[], prefix: readonly string[]) =>
  prefix.length <= keys.length && prefix.every((k, i) => keys[i] === k);

export class SequenceEngine<A extends string> {
  readonly #bindings: readonly Binding<A>[];
  readonly #timeoutMs: number;
  #prefix: string[] = [];
  #lastAt = 0;

  constructor(bindings: readonly Binding<A>[], timeoutMs = DEFAULT_TIMEOUT_MS) {
    this.#bindings = bindings;
    this.#timeoutMs = timeoutMs;
  }

  get pending(): readonly string[] {
    return this.#prefix;
  }

  reset(): void {
    this.#prefix = [];
  }

  feed(key: string, context: Context, now: number): FeedResult<A> {
    if (this.#prefix.length > 0 && now - this.#lastAt > this.#timeoutMs) this.reset();
    const wasPending = this.#prefix.length > 0;
    const seq = [...this.#prefix, key];
    const active = this.#bindings.filter((b) => b.contexts.includes(context));
    const longer = active.filter((b) => b.keys.length > seq.length && startsWith(b.keys, seq));

    if (longer.length > 0) {
      this.#prefix = seq;
      this.#lastAt = now;
      return { kind: "pending", prefix: seq, next: longer };
    }

    this.reset();
    const exact = active.find((b) => b.keys.length === seq.length && startsWith(b.keys, seq));
    if (exact) return { kind: "match", action: exact.action };
    return wasPending ? { kind: "cancelled" } : { kind: "none" };
  }
}

/** Bindings that shadow each other (same keys, or one a prefix of another) in a shared context. */
export function findConflicts(bindings: readonly Binding[]): string[] {
  const conflicts: string[] = [];
  bindings.forEach((a, i) => {
    for (const b of bindings.slice(i + 1)) {
      const shared = a.contexts.some((c) => b.contexts.includes(c));
      const [short, long] = a.keys.length <= b.keys.length ? [a, b] : [b, a];
      if (shared && startsWith(long.keys, short.keys)) {
        conflicts.push(
          `${short.keys.join(" ")} (${short.action}) vs ${long.keys.join(" ")} (${long.action})`,
        );
      }
    }
  });
  return conflicts;
}

interface KeyLike {
  key: string;
  ctrlKey: boolean;
  metaKey: boolean;
  altKey: boolean;
}

const MODIFIER_KEYS = new Set(["Shift", "Control", "Alt", "Meta", "CapsLock", "Dead"]);

/** The engine key for an event, or null when it belongs to the browser/OS (Ctrl, Cmd, Alt chords). */
export function keyFromEvent(event: KeyLike): string | null {
  if (event.ctrlKey || event.metaKey || event.altKey) return null;
  if (MODIFIER_KEYS.has(event.key)) return null;
  return event.key;
}
