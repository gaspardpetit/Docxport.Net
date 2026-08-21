import { copyFile, mkdir } from "node:fs/promises";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const sample = path.resolve(here, "../../samples/sample-no-sectPr.docx");

for (const app of ["react", "vue"]) {
  const publicDirectory = path.join(here, app, "public");
  await mkdir(publicDirectory, { recursive: true });
  await copyFile(sample, path.join(publicDirectory, "sample.docx"));
  run(path.join(here, "node_modules", "docxport", "tools", "copy-assets.mjs"), [path.join(publicDirectory, "docxport")]);
}

function run(command, args) {
  const result = spawnSync(process.execPath, [command, ...args], { stdio: "inherit" });
  if (result.status !== 0) process.exit(result.status ?? 1);
}
