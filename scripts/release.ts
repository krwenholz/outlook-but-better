// Publishes the built extension (dist/) so installs and updates need no build tools:
// - commits it to the `dist` branch, which people clone once and `git pull` to update;
// - writes .cache/obb-<version>.zip for a GitHub Release.
// Usage: node scripts/release.ts [--push]   (mise run release builds first)
// Runs the same locally and in CI (.github/workflows/ci.yml).

import { execFileSync } from "node:child_process";
import { cp, mkdir, readdir, readFile, rm, writeFile } from "node:fs/promises";
import path from "node:path";

const push = process.argv.includes("--push");
const git = (args: string[], cwd = ".", quiet = false) =>
  execFileSync("git", args, {
    cwd,
    encoding: "utf8",
    stdio: ["ignore", "pipe", quiet ? "ignore" : "inherit"],
  }).trim();

const { version } = JSON.parse(await readFile("dist/manifest.json", "utf8")) as { version: string };
const source = git(["rev-parse", "--short", "HEAD"]);
const tree = path.resolve(".cache/dist-branch");

// A worktree on `dist`: the existing branch if there is one, otherwise a fresh orphan.
await rm(tree, { recursive: true, force: true });
git(["worktree", "prune"]);
let exists = true;
try {
  git(["fetch", "--quiet", "origin", "dist"], ".", true);
} catch {
  exists = false;
}
if (exists) {
  git(["worktree", "add", "--quiet", "-B", "dist", tree, "origin/dist"]);
} else {
  git(["worktree", "add", "--quiet", "--detach", tree]);
  git(["checkout", "--quiet", "--orphan", "dist"], tree);
  git(["rm", "-rf", "--quiet", "."], tree);
}

// Replace everything but .git with the build and a short README.
for (const entry of await readdir(tree)) {
  if (entry !== ".git") await rm(path.join(tree, entry), { recursive: true, force: true });
}
await cp("dist", tree, { recursive: true, filter: (f) => !f.endsWith(".map") });
await writeFile(
  path.join(tree, "README.md"),
  `# OBB ${version}

Built from \`${source}\` on main. Load this folder in Chrome (chrome://extensions → Developer mode
→ Load unpacked). To update: \`git pull\` here, then click **Reload OBB** in the OBB popup and
refresh Outlook. Source and docs are on the main branch.
`,
);

git(["add", "--all"], tree);
const changed = git(["status", "--porcelain"], tree) !== "";
if (changed) {
  git(
    [
      "-c",
      "user.name=OBB release",
      "-c",
      "user.email=obb-release@users.noreply.github.com",
      "commit",
      "--quiet",
      "-m",
      `OBB ${version} (built from ${source})`,
    ],
    tree,
  );
  console.log(`dist branch: committed OBB ${version}`);
} else {
  console.log("dist branch: already up to date");
}
if (push && (changed || !exists)) {
  git(["push", "--quiet", "origin", "dist"], tree);
  console.log("dist branch: pushed");
}

// The zip for a GitHub Release.
await mkdir(".cache", { recursive: true });
const zip = path.resolve(`.cache/obb-${version}.zip`);
await rm(zip, { force: true });
execFileSync("zip", ["-qr", zip, ".", "-x", ".git*", "-x", "*.map"], { cwd: tree });
console.log(`wrote ${path.relative(".", zip)}`);

git(["worktree", "remove", "--force", tree]);
