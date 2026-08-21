import { chromium } from "@playwright/test";
import { createServer } from "node:http";
import { readFile, stat } from "node:fs/promises";
import path from "node:path";

const root = path.resolve("dist");
const mime = new Map([[".html", "text/html"], [".js", "text/javascript"], [".wasm", "application/wasm"], [".json", "application/json"], [".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"]]);
const server = createServer(async (request, response) => {
  try {
    const url = new URL(request.url ?? "/", "http://127.0.0.1");
    let target = path.resolve(root, `.${decodeURIComponent(url.pathname)}`);
    if (!target.startsWith(root + path.sep)) throw new Error("Invalid path");
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
  for (const app of ["react", "vue"]) {
    const page = await browser.newPage();
    const failed = [];
    page.on("response", response => { if (response.status() >= 400) failed.push(`${response.status()} ${response.url()}`); });
    page.on("pageerror", error => failed.push(`page error: ${error.message}`));
    await page.goto(`http://127.0.0.1:${port}/nested/${app}/`);
    try {
      await page.waitForFunction(() => {
        const main = document.querySelector("main");
        return main && main.dataset.status !== "loading";
      }, null, { timeout: 60_000 });
    } catch (error) {
      throw new Error(`${app} timed out: ${await page.locator("body").innerText()}\n${failed.join("\n")}`, { cause: error });
    }
    const status = await page.locator("main").getAttribute("data-status");
    const output = await page.locator("pre").textContent();
    if (status !== "ready") throw new Error(`${app} failed: ${output}\n${failed.join("\n")}`);
    if (!output?.trim()) throw new Error(`${app} produced empty output`);
    if (failed.length) throw new Error(`${app} asset failures:\n${failed.join("\n")}`);
    await page.close();
  }
} finally {
  await browser.close();
  server.close();
}
