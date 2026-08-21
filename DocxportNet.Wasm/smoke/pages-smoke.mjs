import { chromium } from "@playwright/test";
import { createServer } from "node:http";
import { readFile, stat } from "node:fs/promises";
import path from "node:path";

const root = path.resolve("../bin/Release/net10.0/publish/wwwroot");
const prefix = "/Docxport.Net/";
const sample = path.resolve("../../samples/sample-no-sectPr.docx");
const mime = new Map([[".html", "text/html"], [".js", "text/javascript"], [".css", "text/css"], [".wasm", "application/wasm"], [".json", "application/json"], [".dat", "application/octet-stream"]]);
const server = createServer(async (request, response) => {
  try {
    const url = new URL(request.url ?? "/", "http://127.0.0.1");
    if (!url.pathname.startsWith(prefix)) throw new Error("Invalid prefix");
    const relative = decodeURIComponent(url.pathname.slice(prefix.length));
    let target = path.resolve(root, relative || ".");
    if (!target.startsWith(root)) throw new Error("Invalid path");
    if ((await stat(target)).isDirectory()) target = path.join(target, "index.html");
    response.setHeader("Content-Type", mime.get(path.extname(target)) ?? "application/octet-stream");
    response.end(await readFile(target));
  } catch {
    response.statusCode = 404;
    response.end("Not found");
  }
});
await new Promise(resolve => server.listen(0, "127.0.0.1", resolve));
const { port } = server.address();
const browser = await chromium.launch({ headless: true });
try {
  const page = await browser.newPage();
  const failed = [];
  page.on("response", response => { if (response.status() >= 400) failed.push(`${response.status()} ${response.url()}`); });
  page.on("pageerror", error => failed.push(`page error: ${error.message}`));
  await page.goto(`http://127.0.0.1:${port}${prefix}`);
  await page.waitForURL(`**${prefix}demo/`);
  await page.locator("#status").filter({ hasText: "Ready" }).waitFor({ timeout: 60_000 });
  await page.locator("#fileInput").setInputFiles(sample);
  await page.waitForFunction(() => {
    const status = document.querySelector("#status")?.textContent;
    return status === "Converted locally" || status === "Conversion failed" || status === "WASM failed to load";
  }, null, { timeout: 60_000 });
  const status = await page.locator("#status").textContent();
  if (status !== "Converted locally") {
    throw new Error(`Demo failed with status '${status}': ${await page.locator("#errorMessage").textContent()}\n${failed.join("\n")}`);
  }
  await page.locator('[data-view="raw"]').click();
  const html = await page.locator("#rawOutput code").textContent();
  await page.locator('input[name="format"][value="markdown"]').check();
  await page.waitForFunction(previous => document.querySelector("#rawOutput code")?.textContent !== previous, html, { timeout: 60_000 });
  const markdown = await page.locator("#rawOutput code").textContent();
  if (!html?.trim() || !markdown?.trim()) throw new Error("Demo produced empty output");
  if (failed.length) throw new Error(`Pages asset failures:\n${failed.join("\n")}`);
} finally {
  await browser.close();
  server.close();
}
