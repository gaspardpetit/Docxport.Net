import { chromium, firefox, webkit } from "@playwright/test";
import katex from "katex";
import { createServer } from "node:http";
import { readFile, stat } from "node:fs/promises";
import path from "node:path";

const root = path.resolve("../bin/Release/net10.0/publish/wwwroot");
const mime = new Map([[".html", "text/html"], [".js", "text/javascript"], [".wasm", "application/wasm"], [".json", "application/json"], [".dat", "application/octet-stream"]]);
const server = createServer(async (request, response) => {
  try {
    const relative = decodeURIComponent(new URL(request.url ?? "/", "http://127.0.0.1").pathname.slice(1));
    let target = path.resolve(root, relative || "index.html");
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
const omml = `<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:rad><m:deg/><m:e><m:r><m:t>b</m:t></m:r></m:e></m:rad></m:den></m:f></m:oMath>`;

try {
  for (const engine of [chromium, firefox, webkit]) {
    const browser = await engine.launch({ headless: true });
    try {
      const page = await browser.newPage();
      await page.goto(`http://127.0.0.1:${port}/index.html`);
      const result = await page.evaluate(async input => {
        const { createDocxport } = await import("/index.js");
        const client = await createDocxport();
        const mathml = await client.convertOmml(input, "mathml");
        const latex = await client.convertOmml(input, "latex");
        document.body.innerHTML = mathml;
        const math = document.querySelector("math");
        return { namespace: math?.namespaceURI, fraction: Boolean(math?.querySelector("mfrac mroot")), display: getComputedStyle(math).display, latex };
      }, omml);
      if (result.namespace !== "http://www.w3.org/1998/Math/MathML" || !result.fraction || result.display === "none" || result.latex !== "\\frac{a}{\\sqrt[]{b}}")
        throw new Error(`${engine.name()} rejected representative OMML output: ${JSON.stringify(result)}`);
      katex.renderToString(result.latex, { throwOnError: true, strict: "error" });
    } finally {
      await browser.close();
    }
  }
} finally {
  server.close();
}
