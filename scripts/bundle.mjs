import { build } from "esbuild";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDirectory = path.resolve(__dirname, "..");

await build({
  entryPoints: [path.join(rootDirectory, "src", "main.ts")],
  bundle: true,
  format: "iife",
  outfile: path.join(rootDirectory, "dist", "content-script.js"),
  target: "es2022",
  platform: "browser",
  minify: false,
  sourcemap: false
});

process.stdout.write("Bundled content script to dist/content-script.js\n");
