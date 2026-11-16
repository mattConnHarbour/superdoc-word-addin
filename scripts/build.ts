import { cp, mkdir } from "node:fs/promises";
import { build } from "esbuild";

await mkdir("dist/ui", { recursive: true });
await build({
  entryPoints: ["src/taskpane/taskpane.ts"],
  outdir: "dist/ui",
  bundle: true,
  format: "esm",
  platform: "browser",
  target: "es2022",
  minify: true,
  legalComments: "none",
});
await cp("src/taskpane/taskpane.html", "dist/ui/taskpane.html");
await cp("assets", "dist/ui/assets", { recursive: true });
await cp("manifest.xml", "dist/ui/manifest.xml");
