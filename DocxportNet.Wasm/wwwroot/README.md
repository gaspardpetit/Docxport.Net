# docxport

Convert DOCX files to HTML, Markdown, or plain text entirely in the browser. The package also exposes field resolution and resolved-DOCX output through a .NET WebAssembly runtime.

## Install

```bash
npm install docxport
```

The package is browser-only. Do not initialize it during server-side rendering.

## Vite, React, and Vue

.NET uses a directory of runtime assets rather than a single WASM file. Copy those assets into the application's public directory before building:

```json
{
  "scripts": {
    "copy:docxport": "docxport-copy-assets public/docxport",
    "build": "npm run copy:docxport && vite build",
    "dev": "npm run copy:docxport && vite"
  }
}
```

Initialize the package from browser code using the matching public URL:

```ts
import { createDocxport } from "docxport";

const assetBaseUrl = new URL(
  `${import.meta.env.BASE_URL}docxport/`,
  window.location.origin
);
const docxport = await createDocxport({ assetBaseUrl });

const html = await docxport.export(docxBytes, {
  format: "html",
  preset: "rich",
  fields: { mode: "cache" },
  onProgress(progress) {
    console.log(progress.phase, progress.percentage);
  }
});
```

The `import.meta.env.BASE_URL` form works when a Vite application is deployed below a nested URL. In React, call this from a client component or an effect or event handler. In Vue, call it from browser-side setup or an event handler.

## Other build systems

Run `docxport-copy-assets <static-directory>/docxport` after installing dependencies and before the application build. Configure the application to serve that directory at `/docxport/`, then pass that URL as `assetBaseUrl`.

The server must:

- Serve all files under `_framework` without renaming them.
- Serve `.wasm` files as `application/wasm`.
- Use HTTP or HTTPS; browser runtimes cannot load from `file://` URLs.

The default asset location is relative to the installed ESM loader and is useful for direct unbundled imports. Bundled production applications should use the explicit copy command and `assetBaseUrl`.

## API

```ts
const docxport = await createDocxport({ assetBaseUrl: "/docxport/" });
const info = await docxport.inspect(docxBytes);
const markdown = await docxport.export(docxBytes, {
  format: "markdown",
  preset: "plain",
  markdown: { mathOutputFormat: "latex", emitMathDelimiters: true, mathDelimiterStyle: "auto" }
});
const resolvedBytes = await docxport.resolveDocx(docxBytes, {
  fields: { mode: "evaluate", variables: { Customer: "Ada" } }
});

const mathml = await docxport.convertOmml(ommlXml, "mathml");
const latex = await docxport.convertOmml(ommlXml, "latex");
```

Standalone OMML conversion also supports `html`, `unicodemath`, and `text`.
DOCX export accepts `mathOutputFormat` (`mathml`, `latex`, `unicodemath`,
`text`, or `none`) in the HTML, Markdown, and text option groups. Defaults are
MathML for HTML, LaTeX for Markdown, and readable text for text export.
Markdown `mathDelimiterStyle` accepts `auto` (the default), `dollar`, or
`backslash`.
See `index.d.ts` for format-specific options.
The optional `onProgress` callback receives the current phase, completed and
total paragraph units, and a nullable percentage. Supplying it enables the
lightweight paragraph-counting pre-pass.

## Build locally

```powershell
./DocxportNet.Wasm/build-package.ps1
cd ./DocxportNet.Wasm/bin/Release/net10.0/publish/wwwroot
npm pack
```

## Publishing setup

The first `docxport` version must be published manually from the validated tarball with npm 2FA and `npm publish <tarball> --access public`. Then configure the package's npm trusted publisher for `gaspardpetit/Docxport.Net`, workflow `publish-npm.yml`, and the `npm publish` action. Published GitHub Releases subsequently publish through OIDC without an npm token.
