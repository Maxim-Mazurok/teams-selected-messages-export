import fs from "node:fs/promises";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";
import { build } from "./build-extension.mjs";

const execFileAsync = promisify(execFile);
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const extensionDistDir = path.join(rootDir, "extension-dist");
const artifactsDir = path.join(rootDir, "artifacts");
const outputPath = path.join(artifactsDir, "teams-message-export-extension.zip");

async function zipWith(command, args, options) {
  try {
    await execFileAsync(command, args, options);
    return true;
  } catch {
    return false;
  }
}

async function main() {
  await build();
  await fs.mkdir(artifactsDir, { recursive: true });
  await fs.rm(outputPath, { force: true });

  const zippedWithZip = await zipWith("zip", ["-qr", outputPath, "."], {
    cwd: extensionDistDir
  });

  if (!zippedWithZip) {
    const zippedWithDitto = await zipWith("ditto", ["-c", "-k", "--keepParent", extensionDistDir, outputPath], {
      cwd: rootDir
    });

    if (!zippedWithDitto) {
      throw new Error("Unable to create extension zip with either `zip` or `ditto`.");
    }
  }

  process.stdout.write(`Packaged extension at ${outputPath}\n`);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
