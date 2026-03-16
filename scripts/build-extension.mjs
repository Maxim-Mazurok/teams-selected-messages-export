import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const extensionSrcDir = path.join(rootDir, "extension-src");
const extensionDistDir = path.join(rootDir, "extension-dist");
const sourceScriptPath = path.join(rootDir, "dist", "content-script.js");

export async function build() {
  await fs.rm(extensionDistDir, { recursive: true, force: true });
  await fs.mkdir(extensionDistDir, { recursive: true });
  await fs.mkdir(path.join(extensionDistDir, "icons"), { recursive: true });

  await fs.copyFile(path.join(extensionSrcDir, "manifest.json"), path.join(extensionDistDir, "manifest.json"));
  await fs.copyFile(path.join(extensionSrcDir, "background.js"), path.join(extensionDistDir, "background.js"));
  await fs.copyFile(path.join(extensionSrcDir, "worker-hook.js"), path.join(extensionDistDir, "worker-hook.js"));
  await fs.copyFile(sourceScriptPath, path.join(extensionDistDir, "content-script.js"));
  await fs.cp(path.join(extensionSrcDir, "icons"), path.join(extensionDistDir, "icons"), { recursive: true });

  process.stdout.write(`Extension built at ${extensionDistDir}\n`);
}

if (process.argv[1] && path.resolve(process.argv[1]) === __filename) {
  build().catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
}
