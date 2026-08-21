#!/usr/bin/env node

import { cp, mkdir, rm, stat } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const packageRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const sourceFramework = path.join(packageRoot, "_framework");
const destinationArgument = process.argv[2];

if (!destinationArgument || process.argv.length !== 3) {
  fail("Usage: docxport-copy-assets <public-directory>");
}

const destinationRoot = path.resolve(process.cwd(), destinationArgument);
const destinationFramework = path.join(destinationRoot, "_framework");

try {
  const sourceInfo = await stat(sourceFramework);
  if (!sourceInfo.isDirectory()) {
    fail(`Docxport runtime assets were not found at ${sourceFramework}`);
  }

  await mkdir(destinationRoot, { recursive: true });
  await rm(destinationFramework, { recursive: true, force: true });
  await cp(sourceFramework, destinationFramework, { recursive: true, force: true });
  process.stdout.write(`Copied Docxport runtime assets to ${destinationFramework}\n`);
} catch (error) {
  fail(error instanceof Error ? error.message : String(error));
}

function fail(message) {
  process.stderr.write(`${message}\n`);
  process.exit(1);
}
