import { createDocxport } from "../index.js";
import { renderMarkdown } from "./markdown.js";

const $ = selector => document.querySelector(selector);
const elements = {
  drop: $("#dropZone"), input: $("#fileInput"), fileCard: $("#fileCard"),
  fileName: $("#fileName"), fileSize: $("#fileSize"), clear: $("#clearButton"),
  presetField: $("#presetField"), presetLabel: $("#presetLabel"), preset: $("#presetSelect"),
  settings: $("#advancedSettings"), status: $("#status"),
  copy: $("#copyButton"), download: $("#downloadButton"), viewToggle: $("#viewToggle"),
  empty: $("#emptyState"), error: $("#errorState"),
  errorMessage: $("#errorMessage"), rendered: $("#renderedOutput"),
  raw: $("#rawOutput"), pdf: $("#pdfOutput"),
  pdfFontField: $("#pdfFontField"), pdfFontSelect: $("#pdfFontSelect"),
  localFontsButton: $("#localFontsButton"), localFontsMessage: $("#localFontsMessage")
};

const common = [
  ["emitImages", "Include images"], ["emitStyleFont", "Font styles"],
  ["emitRunColor", "Text colors"], ["emitRunBackground", "Text backgrounds"],
  ["emitTableBorders", "Table borders"], ["emitDocumentColors", "Document colors"],
  ["emitParagraphAlignment", "Paragraph alignment"], ["preserveListSymbols", "List symbols"],
  ["richTables", "Rich tables"], ["emitSectionHeadersFooters", "Headers and footers"],
  ["emitUnreferencedBookmarks", "Unreferenced bookmarks"], ["emitPageNumbers", "Page numbers"],
  ["emitFieldInstructions", "Field instructions"], ["usePlainComments", "Emit comments as hidden source comments"],
  ["emitCustomProperties", "Custom properties"], ["emitTimeline", "Change timeline"]
];
const settingSchemas = {
  html: [
    ...common, ["emitParagraphMetadata", "Paragraph metadata"],
    ["embedDefaultStylesheet", "Embed default stylesheet"],
    ["rootCssClass", "Root CSS class", "text"], ["stylesheetHref", "Stylesheet URL", "text"],
    ["trackedChangeMode", "Tracked changes", "select", ["accept", "reject", "inline", "split"]],
    ["headerSelection", "Header selection", "select", ["none", "first", "last"]],
    ["footerSelection", "Footer selection", "select", ["none", "first", "last"]]
  ],
  markdown: [
    ...common, ["emitRichLayoutHtml", "Rich layout HTML"], ["usePlainCodeBlocks", "Plain code blocks"],
    ["useMarkdownInlineStyles", "Markdown inline styles"],
    ["trackedChangeMode", "Tracked changes", "select", ["accept", "reject", "inline", "split"]]
  ],
  text: [
    ["emitDocumentProperties", "Document properties"], ["emitCustomProperties", "Custom properties"],
    ["imagePlaceholder", "Image placeholder", "text"]
  ],
  pdf: []
};

let api;
let documentBytes;
let rawOutput = "";
let pdfUrl;
let pdfFontReady;
const localFontFaces = new Map();
const registeredLocalFonts = new Set();
let currentView = "rendered";
let generation = 0;
let renderedSource = "";
let hasTrackedChanges = false;

function currentFormat() { return $("input[name=format]:checked").value; }
function formatSize(bytes) { return bytes < 1024 * 1024 ? `${Math.ceil(bytes / 1024)} KB` : `${(bytes / 1024 / 1024).toFixed(1)} MB`; }
function titleCase(value) { return value.replace(/([A-Z])/g, " $1").replace(/^./, c => c.toUpperCase()); }

function renderSettings() {
  const format = currentFormat();
  const textTrackedChanges = format === "text";
  elements.presetField.hidden = format === "pdf" || textTrackedChanges && !hasTrackedChanges;
  elements.settings.parentElement.hidden = format === "pdf";
  elements.viewToggle.hidden = format === "pdf";
  elements.copy.hidden = format === "pdf";
  elements.download.hidden = format !== "pdf";
  elements.pdfFontField.hidden = format !== "pdf";
  elements.presetLabel.textContent = textTrackedChanges ? "Tracked changes" : "Preset";
  elements.preset.replaceChildren();
  const presets = format === "text"
    ? [["accept", "Accept changes"], ["reject", "Reject changes"]]
    : [["rich", "Rich"], ["plain", "Plain"]];
  for (const [value, label] of presets) {
    const option = document.createElement("option");
    option.value = value; option.textContent = label; elements.preset.append(option);
  }
  elements.settings.replaceChildren();
  for (const [key, label, type = "checkbox", values] of settingSchemas[format]) {
    if (key === "trackedChangeMode" && !hasTrackedChanges) continue;
    const wrapper = document.createElement("label");
    wrapper.className = type === "checkbox" ? "setting setting-check" : "setting";
    const caption = document.createElement("span"); caption.textContent = label;
    let input;
    if (type === "select") {
      input = document.createElement("select");
      input.innerHTML = `<option value="">Preset default</option>${values.map(v => `<option value="${v}">${titleCase(v)}</option>`).join("")}`;
      if (key === "trackedChangeMode") input.value = "accept";
    } else if (type === "text") {
      input = document.createElement("input"); input.type = "text"; input.placeholder = "Preset default";
    } else {
      input = document.createElement("input"); input.type = "checkbox"; input.indeterminate = true;
      input.addEventListener("click", () => { if (input.indeterminate) { input.indeterminate = false; input.checked = true; } });
    }
    input.dataset.key = key;
    wrapper.append(type === "checkbox" ? input : caption, type === "checkbox" ? caption : input);
    elements.settings.append(wrapper);
  }
}

function collectOptions() {
  const options = {};
  for (const input of elements.settings.querySelectorAll("[data-key]")) {
    if (input.type === "checkbox") {
      if (!input.indeterminate) options[input.dataset.key] = input.checked;
    } else if (input.value !== "") options[input.dataset.key] = input.value;
  }
  return options;
}

function request() {
  const format = currentFormat();
  const options = collectOptions();
  if (format === "text") options.trackedChangeMode = elements.preset.value;
  return {
    format,
    ...(format === "text" ? {} : { preset: elements.preset.value }),
    fields: { mode: "cache" },
    [format]: options
  };
}

function renderedDocument() {
  const format = currentFormat();
  if (format === "html") return rawOutput;
  const body = format === "markdown" ? renderMarkdown(rawOutput) : `<pre>${escapeHtml(rawOutput)}</pre>`;
  return `<!doctype html><html><head><meta charset="utf-8"><style>body{max-width:900px;margin:36px auto;padding:0 28px;color:#20232a;font:16px/1.6 system-ui,sans-serif}img{max-width:100%}table{border-collapse:collapse;max-width:100%}td,th{padding:.45rem;border:1px solid #ccd1d8}pre{white-space:pre-wrap;word-break:break-word}code{background:#f1f3f5;padding:.12rem .3rem;border-radius:4px}blockquote{border-left:3px solid #5e6ad2;margin-left:0;padding-left:1rem;color:#59616d}</style></head><body>${body}</body></html>`;
}

function escapeHtml(value) { const div = document.createElement("div"); div.textContent = value; return div.innerHTML; }

function showOutput() {
  if (currentFormat() === "pdf") {
    elements.raw.hidden = elements.rendered.hidden = true;
    elements.pdf.hidden = !pdfUrl;
    if (pdfUrl) elements.pdf.src = pdfUrl;
    return;
  }
  elements.pdf.hidden = true;
  const raw = currentView === "raw";
  elements.raw.hidden = !raw;
  elements.rendered.hidden = raw;
  if (raw) elements.raw.querySelector("code").textContent = rawOutput;
  else {
    const source = renderedDocument();
    if (source !== renderedSource) {
      renderedSource = source;
      elements.rendered.srcdoc = source;
    }
  }
}

async function convert() {
  if (!documentBytes || !api) return;
  const token = ++generation;
  elements.status.textContent = "Converting…";
  elements.status.className = "status busy";
  elements.copy.disabled = true;
  elements.download.disabled = true;
  elements.error.hidden = true;
  try {
    if (currentFormat() === "pdf") {
      pdfFontReady ??= fetch(new URL("./fonts/NotoSans-Regular.ttf", import.meta.url))
        .then(response => {
          if (!response.ok) throw new Error(`Demo font could not be loaded (${response.status}).`);
          return response.arrayBuffer();
        }).then(bytes => api.registerPdfFont("Noto Sans", bytes));
      await pdfFontReady;
      const selectedFamily = elements.pdfFontSelect.value;
      if (selectedFamily && !registeredLocalFonts.has(selectedFamily)) {
        const face = localFontFaces.get(selectedFamily);
        if (!face) throw new Error(`Local font '${selectedFamily}' is no longer available.`);
        api.registerPdfFont(selectedFamily, await (await face.blob()).arrayBuffer());
        registeredLocalFonts.add(selectedFamily);
      }
      if (token !== generation) return;
      const bytes = await api.exportPdf(documentBytes,
        { fields: { mode: "cache" }, pdfFontFamily: selectedFamily || "Noto Sans" });
      if (token !== generation) return;
      if (pdfUrl) URL.revokeObjectURL(pdfUrl);
      pdfUrl = URL.createObjectURL(new Blob([bytes], { type: "application/pdf" }));
      showOutput();
      elements.empty.hidden = true;
      elements.download.disabled = false;
      elements.status.textContent = "Converted locally";
      elements.status.className = "status ready";
      return;
    }
    const output = await api.export(documentBytes, request());
    if (token !== generation) return;
    if (pdfUrl) { URL.revokeObjectURL(pdfUrl); pdfUrl = undefined; elements.pdf.removeAttribute("src"); }
    rawOutput = output;
    showOutput();
    elements.empty.hidden = true;
    elements.copy.disabled = false;
    elements.status.textContent = "Converted locally";
    elements.status.className = "status ready";
  } catch (error) {
    if (token !== generation) return;
    elements.rendered.hidden = elements.raw.hidden = elements.pdf.hidden = true;
    elements.empty.hidden = true;
    elements.error.hidden = false;
    elements.errorMessage.textContent = error?.message ?? String(error);
    elements.status.textContent = "Conversion failed";
    elements.status.className = "status failed";
  }
}

async function loadFile(file) {
  if (!file || !/\.docx?$/i.test(file.name)) {
    elements.error.hidden = false; elements.empty.hidden = true;
    elements.errorMessage.textContent = "Choose a .doc or .docx file."; return;
  }
  documentBytes = new Uint8Array(await file.arrayBuffer());
  elements.fileName.textContent = file.name; elements.fileSize.textContent = formatSize(file.size);
  elements.drop.hidden = true; elements.fileCard.hidden = false;
  try {
    const info = await api.inspect(documentBytes);
    hasTrackedChanges = info.hasTrackedChanges;
  } catch {
    hasTrackedChanges = false;
  }
  renderSettings();
  await convert();
}

function clearDocument() {
  generation++; documentBytes = undefined; rawOutput = ""; renderedSource = ""; hasTrackedChanges = false; elements.input.value = "";
  if (pdfUrl) { URL.revokeObjectURL(pdfUrl); pdfUrl = undefined; elements.pdf.removeAttribute("src"); }
  elements.rendered.removeAttribute("srcdoc");
  elements.fileCard.hidden = true; elements.drop.hidden = false; elements.empty.hidden = false;
  elements.error.hidden = elements.rendered.hidden = elements.raw.hidden = elements.pdf.hidden = true;
  elements.copy.disabled = true; elements.status.textContent = api ? "Ready" : "Loading WebAssembly…";
  elements.download.disabled = true;
  renderSettings();
}

elements.drop.addEventListener("click", () => elements.input.click());
elements.drop.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") elements.input.click(); });
elements.input.addEventListener("change", () => loadFile(elements.input.files[0]));
for (const name of ["dragenter", "dragover"]) elements.drop.addEventListener(name, e => { e.preventDefault(); elements.drop.classList.add("dragging"); });
for (const name of ["dragleave", "drop"]) elements.drop.addEventListener(name, e => { e.preventDefault(); elements.drop.classList.remove("dragging"); });
elements.drop.addEventListener("drop", e => loadFile(e.dataTransfer.files[0]));
elements.clear.addEventListener("click", clearDocument);
$("#formatControl").addEventListener("change", () => { renderSettings(); convert(); });
elements.preset.addEventListener("change", convert);
elements.settings.addEventListener("change", convert);
elements.settings.addEventListener("input", e => { if (e.target.type === "text") { clearTimeout(e.target.timer); e.target.timer = setTimeout(convert, 250); } });
elements.localFontsButton.disabled = typeof window.queryLocalFonts !== "function";
if (elements.localFontsButton.disabled)
  elements.localFontsMessage.textContent = "This browser does not offer local font access. The bundled font remains available.";
elements.localFontsButton.addEventListener("click", async () => {
  elements.localFontsButton.disabled = true;
  try {
    const faces = await window.queryLocalFonts();
    localFontFaces.clear();
    for (const face of faces) {
      if (!face.family || !face.blob) continue;
      const previous = localFontFaces.get(face.family);
      if (!previous || (face.style === "Regular" && previous.style !== "Regular"))
        localFontFaces.set(face.family, face);
    }
    elements.pdfFontSelect.replaceChildren(new Option("Bundled Noto Sans", ""));
    for (const family of [...localFontFaces.keys()].sort((a, b) => a.localeCompare(b)))
      elements.pdfFontSelect.add(new Option(family, family));
    elements.localFontsMessage.textContent = localFontFaces.size
      ? `Choose one of ${localFontFaces.size} available font families. The PDF may embed the selected font.`
      : "No installed fonts were made available. The bundled font remains available.";
  } catch (error) {
    elements.localFontsMessage.textContent = error?.name === "NotAllowedError"
      ? "Local font permission was declined. The bundled font remains available."
      : `Local fonts could not be listed: ${error?.message ?? String(error)}`;
  } finally {
    elements.localFontsButton.disabled = false;
  }
});
elements.pdfFontSelect.addEventListener("change", convert);
document.querySelectorAll("[data-view]").forEach(button => button.addEventListener("click", () => {
  currentView = button.dataset.view;
  document.querySelectorAll("[data-view]").forEach(item => item.classList.toggle("active", item === button));
  if (rawOutput) showOutput();
}));
elements.copy.addEventListener("click", async () => {
  await navigator.clipboard.writeText(rawOutput);
  const original = elements.copy.textContent; elements.copy.textContent = "Copied";
  setTimeout(() => elements.copy.textContent = original, 1200);
});
elements.download.addEventListener("click", () => {
  if (!pdfUrl) return;
  const link = document.createElement("a");
  link.href = pdfUrl;
  link.download = elements.fileName.textContent.replace(/\.docx?$/i, "") + ".pdf";
  link.click();
});

renderSettings();
try {
  api = await createDocxport({ assetBaseUrl: new URL("../", import.meta.url) });
  elements.status.textContent = "Ready"; elements.status.className = "status ready";
  const sampleUrl = new URLSearchParams(location.search).get("sample");
  if (sampleUrl) {
    const response = await fetch(sampleUrl);
    if (!response.ok) throw new Error(`Sample download failed (${response.status}).`);
    const name = new URL(sampleUrl, location.href).pathname.split("/").pop() || "sample.docx";
    await loadFile(new File([await response.arrayBuffer()], name, {
      type: /\.doc$/i.test(name) ? "application/msword" : "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    }));
  }
} catch (error) {
  elements.status.textContent = "WASM failed to load"; elements.status.className = "status failed";
  elements.error.hidden = false; elements.empty.hidden = true; elements.errorMessage.textContent = error?.message ?? String(error);
}
