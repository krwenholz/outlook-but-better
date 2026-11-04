// Folder palette: Fastmail's `g` (go to folder) type-ahead picker, plus Outlook's apps.
// Moving (`m`) uses Outlook's own "Move to" picker for now; see docs/decisions.md.

import { type FolderNode, readFolderTree, resolve, type SelectorName } from "../selectors.ts";
import { el, overlayRoot, toast } from "../ui/overlay.ts";
import { type FuzzyMatch, fuzzyFilter } from "./fuzzy.ts";

const MAX_RESULTS = 12;

function highlight(text: string, indices: readonly number[]): DocumentFragment {
  const frag = document.createDocumentFragment();
  const hit = new Set(indices);
  for (const [i, ch] of [...text].entries()) {
    frag.append(hit.has(i) ? el("mark", "", ch) : document.createTextNode(ch));
  }
  return frag;
}

const APPS: [string, SelectorName][] = [
  ["Mail", "mailApp"],
  ["Calendar", "calendarApp"],
  ["People", "peopleApp"],
];

/** Outlook's apps, listed alongside folders so `g` then "cal" switches to Calendar. */
function appTargets(): FolderNode[] {
  return APPS.flatMap(([name, target]) => {
    const element = resolve(target);
    return element ? [{ name, path: name, element }] : [];
  });
}

export function openFolderPalette(): void {
  const folders = [...appTargets(), ...readFolderTree()];
  if (folders.length === 0) {
    toast("Overlook couldn't find the folder pane");
    return;
  }

  const root = overlayRoot();
  const backdrop = el("div", "backdrop");
  const panel = el("div", "panel palette");
  const input = el("input", "");
  input.placeholder = "Go to folder, or Mail, Calendar, People…";
  input.setAttribute("aria-label", input.placeholder);
  const list = el("ul", "");
  list.setAttribute("role", "listbox");
  panel.append(input, list);
  backdrop.append(panel);
  root.append(backdrop);

  let matches: FuzzyMatch<FolderNode>[] = [];
  let selected = 0;

  const close = () => backdrop.remove();
  const commit = () => {
    const match = matches[selected];
    close();
    if (match) (match.item.element as HTMLElement).click();
  };

  const render = () => {
    matches = fuzzyFilter(input.value, folders, (f) => f.path).slice(0, MAX_RESULTS);
    selected = Math.min(selected, Math.max(matches.length - 1, 0));
    list.replaceChildren(
      ...matches.map((m, i) => {
        const li = el("li", "");
        li.setAttribute("role", "option");
        li.setAttribute("aria-selected", String(i === selected));
        li.append(highlight(m.item.path, m.indices));
        li.addEventListener("mousedown", (e) => {
          e.preventDefault();
          selected = i;
          commit();
        });
        return li;
      }),
    );
    if (matches.length === 0) list.append(el("li", "empty", "No matching folders"));
  };

  input.addEventListener("input", () => {
    selected = 0;
    render();
  });
  input.addEventListener("keydown", (e) => {
    e.stopPropagation();
    if (e.key === "Escape") close();
    else if (e.key === "Enter") commit();
    else if (e.key === "ArrowDown" || (e.ctrlKey && e.key === "n")) selected++;
    else if (e.key === "ArrowUp" || (e.ctrlKey && e.key === "p"))
      selected = Math.max(selected - 1, 0);
    else return;
    e.preventDefault();
    render();
  });
  backdrop.addEventListener("mousedown", (e) => {
    if (e.target === backdrop) close();
  });

  render();
  input.focus();
}
