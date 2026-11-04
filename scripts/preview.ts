// Stitches the captioned frames from `mise run screenshots` into docs/preview.webp, the
// crossfading preview at the top of the README. Needs ffmpeg with libwebp (a dev-time tool only).

import { execFileSync } from "node:child_process";
import { readdir } from "node:fs/promises";

const dir = ".cache/preview-frames";
const HOLD = 2.6; // seconds each look is shown
const FADE = 0.5;
const WIDTH = 1200;

const frames = (await readdir(dir)).filter((f) => f.endsWith(".png")).sort();
if (frames.length < 2) throw new Error(`run \`mise run screenshots\` first (${dir} is empty)`);

const inputs = frames.flatMap((f) => [
  "-loop",
  "1",
  "-t",
  String(HOLD + FADE),
  "-i",
  `${dir}/${f}`,
]);
// Chain crossfades: [0][1] -> v1, [v1][2] -> v2, …
const steps: string[] = [];
let last = "0";
for (let i = 1; i < frames.length; i++) {
  const next = `v${i}`;
  steps.push(
    `[${last}][${i}]xfade=transition=fade:duration=${FADE}:offset=${(HOLD * i).toFixed(2)}[${next}]`,
  );
  last = next;
}
steps.push(`[${last}]scale=${WIDTH}:-2:flags=lanczos,fps=12[out]`);

execFileSync(
  "ffmpeg",
  [
    "-y",
    "-loglevel",
    "error",
    ...inputs,
    "-filter_complex",
    steps.join(";"),
    "-map",
    "[out]",
    "-c:v",
    "libwebp_anim",
    "-lossless",
    "0",
    "-q:v",
    "72",
    "-loop",
    "0",
    "docs/preview.webp",
  ],
  { stdio: "inherit" },
);
console.log(`wrote docs/preview.webp from ${frames.length} frames`);
