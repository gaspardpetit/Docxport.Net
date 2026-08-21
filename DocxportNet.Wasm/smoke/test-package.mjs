import assert from "node:assert/strict";
import { access, mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";

const packageRoot = path.resolve("node_modules/docxport");
const copier = path.join(packageRoot, "tools", "copy-assets.mjs");
const installedCommand = path.resolve("node_modules/.bin", process.platform === "win32" ? "docxport-copy-assets.cmd" : "docxport-copy-assets");
for (const required of ["index.js", "index.d.ts", "README.md", "LICENSE", "package.json", "_framework/dotnet.js", "tools/copy-assets.mjs"]) {
  await access(path.join(packageRoot, required));
}

const missingArgument = spawnSync(process.execPath, [copier], { encoding: "utf8" });
assert.notEqual(missingArgument.status, 0);
assert.match(missingArgument.stderr, /Usage:/);
await access(installedCommand);

const temporaryRoot = await mkdtemp(path.join(os.tmpdir(), "docxport smoke spaces "));
const destination = path.join(temporaryRoot, "public assets", "docxport");
await mkdir(path.join(destination, "_framework"), { recursive: true });
await writeFile(path.join(destination, "keep.txt"), "keep");
await writeFile(path.join(destination, "_framework", "stale.txt"), "stale");

const copied = spawnSync(process.execPath, [copier, destination], { encoding: "utf8" });
assert.equal(copied.status, 0, copied.stderr);
await access(path.join(destination, "_framework", "dotnet.js"));
assert.equal(await readFile(path.join(destination, "keep.txt"), "utf8"), "keep");
assert.equal(spawnSync(process.execPath, ["-e", `require('fs').accessSync(${JSON.stringify(path.join(destination, "_framework", "stale.txt"))})`]).status, 1);
