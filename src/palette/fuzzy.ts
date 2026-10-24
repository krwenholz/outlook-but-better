// Pure fuzzy matcher for the folder palette: case-insensitive subsequence matching, scored to
// favour word starts and consecutive runs, the way Fastmail's and most editors' pickers feel.

export interface FuzzyMatch<T> {
  item: T;
  score: number;
  /** Indices into the item's text that matched, for highlighting. */
  indices: number[];
}

const isBoundary = (text: string, i: number) => i === 0 || /[\s/\\._-]/.test(text[i - 1] ?? "");

function isSubsequence(query: string, text: string, from: number): boolean {
  let at = from;
  for (const ch of query) {
    at = text.indexOf(ch, at) + 1;
    if (at === 0) return false;
  }
  return true;
}

/** Score `query` against `text`, or null if it is not a subsequence. Higher is better. */
export function score(query: string, text: string): { score: number; indices: number[] } | null {
  const q = query.toLowerCase();
  const t = text.toLowerCase();
  if (q.length === 0) return { score: 0, indices: [] };

  const indices: number[] = [];
  let total = 0;
  let from = 0;
  for (const [n, ch] of [...q].entries()) {
    // Prefer the next occurrence at a word boundary, as long as the rest of the query still
    // fits after it; otherwise take the nearest one.
    const rest = q.slice(n + 1);
    let at = -1;
    for (let i = t.indexOf(ch, from); i >= 0; i = t.indexOf(ch, i + 1)) {
      if (!isSubsequence(rest, t, i + 1)) break;
      if (at < 0) at = i;
      if (isBoundary(text, i)) {
        at = i;
        break;
      }
    }
    if (at < 0) return null;
    const prev = indices.at(-1);
    total += 1;
    if (isBoundary(text, at)) total += 8;
    if (prev !== undefined && at === prev + 1) total += 5;
    if (prev !== undefined) total -= Math.min(at - prev - 1, 3);
    indices.push(at);
    from = at + 1;
  }
  // Shorter targets win ties, so "Inbox" beats "Inbox/Receipts/2019" for "inb".
  total -= text.length * 0.01;
  return { score: total, indices };
}

export function fuzzyFilter<T>(
  query: string,
  items: readonly T[],
  text: (item: T) => string,
): FuzzyMatch<T>[] {
  const matches: FuzzyMatch<T>[] = [];
  for (const item of items) {
    const result = score(query, text(item));
    if (result) matches.push({ item, ...result });
  }
  return matches.sort((a, b) => b.score - a.score);
}
