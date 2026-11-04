// Bundles the extension into dist/. Usage: node scripts/build.ts [--watch]
import { watch as watchFile } from "node:fs";
import { cp, mkdir, readFile, rm, writeFile } from "node:fs/promises";
import * as esbuild from "esbuild";

const watch = process.argv.includes("--watch");
const outdir = "dist";
const styles = ["palettes", "layouts", "declutter", "density"].map(
  (name) => `src/styles/${name}.css`,
);
const staticFiles = [
  "manifest.json",
  "src/options/options.html",
  "src/options/options.css",
  ...styles,
];

/** Copies static files and concatenates the page stylesheets into content.css. */
const statics: esbuild.Plugin = {
  name: "statics",
  setup(build) {
    build.onEnd(async () => {
      await cp("manifest.json", `${outdir}/manifest.json`);
      await cp("src/options/options.html", `${outdir}/options.html`);
      await cp("icons", `${outdir}/icons`, { recursive: true, filter: (f) => !f.endsWith(".svg") });
      const css = await Promise.all(styles.map((f) => readFile(f, "utf8")));
      await writeFile(`${outdir}/content.css`, css.join("\n"));
      // The options page wears the user's palette, so it gets the palettes too.
      const optionsCss = await Promise.all(
        ["src/styles/palettes.css", "src/options/options.css"].map((f) => readFile(f, "utf8")),
      );
      await writeFile(`${outdir}/options.css`, optionsCss.join("\n"));
      if (watch) console.log(`[build] ${new Date().toLocaleTimeString()} rebuilt ${outdir}/`);
    });
  },
};

await rm(outdir, { recursive: true, force: true });
await mkdir(outdir, { recursive: true });

const ctx = await esbuild.context({
  entryPoints: { content: "src/content.ts", options: "src/options/options.ts" },
  outdir,
  bundle: true,
  format: "iife",
  target: "chrome120",
  loader: { ".css": "text" },
  sourcemap: "inline",
  logLevel: "warning",
  plugins: [statics],
});

if (watch) {
  await ctx.watch();
  // esbuild only watches the import graph; static files need their own watcher.
  for (const file of staticFiles) watchFile(file, () => void ctx.rebuild());
  console.log("[build] watching; reload the extension in chrome://extensions after changes");
} else {
  await ctx.rebuild();
  await ctx.dispose();
  console.log(`[build] wrote ${outdir}/ (load it unpacked from chrome://extensions)`);
}
