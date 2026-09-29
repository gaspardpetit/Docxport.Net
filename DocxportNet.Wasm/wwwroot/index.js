let initialization;

function normalizeBaseUrl(value) {
  if (!value) return new URL("./", import.meta.url);
  return value instanceof URL ? value : new URL(value, globalThis.location?.href ?? import.meta.url);
}

async function initialize(options = {}) {
  const baseUrl = normalizeBaseUrl(options.assetBaseUrl);
  const runtimeUrl = new URL("_framework/dotnet.js", baseUrl);
  const { dotnet } = await import(runtimeUrl.href);
  const runtime = await dotnet
    .withDiagnosticTracing(Boolean(options.diagnosticTracing))
    .withApplicationEnvironment(options.environment ?? "Production")
    .create();
  const config = runtime.getConfig();
  const exports = await runtime.getAssemblyExports(config.mainAssemblyName);
  const api = exports.DocxportNet.Wasm.BrowserExports;
  if (!api) throw new Error("DocxportNet WASM exports could not be loaded.");
  return api;
}

function requireBytes(input) {
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  throw new TypeError("DOC or DOCX input must be a Uint8Array or ArrayBuffer.");
}

function base64Bytes(input) {
  const bytes = requireBytes(input);
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  return btoa(binary);
}

export async function createDocxport(options = {}) {
  initialization ??= initialize(options);
  const api = await initialization;

  return Object.freeze({
    async convertOmml(omml, format = "mathml") {
      if (typeof omml !== "string" || !omml.trim()) throw new TypeError("OMML input must be a non-empty string.");
      return api.ConvertOmml(omml, format);
    },
    async inspect(input) {
      return JSON.parse(api.Inspect(requireBytes(input)));
    },
    async export(input, request = {}) {
      const { onProgress, ...serializableRequest } = request;
      if (onProgress !== undefined && typeof onProgress !== "function") {
        throw new TypeError("onProgress must be a function.");
      }
      const progressCallback = onProgress
        ? value => onProgress(JSON.parse(value))
        : null;
      return api.Export(requireBytes(input), JSON.stringify(serializableRequest), progressCallback);
    },
    async projectDocx(input) {
      const result = api.ProjectDocx(requireBytes(input));
      return result instanceof Uint8Array ? result : new Uint8Array(result);
    },
    async exportDoc(input, request = {}) {
      const result = api.ExportDoc(requireBytes(input), JSON.stringify(request));
      return result instanceof Uint8Array ? result : new Uint8Array(result);
    },
    async editDocEnvelope(input, edits) {
      if (!Array.isArray(edits) || edits.length === 0) {
        throw new TypeError("At least one envelope edit is required.");
      }
      const serialized = edits.map(edit => edit.operation === "set"
        ? { ...edit, envelope: {
            ...edit.envelope,
            attachments: edit.envelope?.attachments?.map(attachment => ({
              fileName: attachment.fileName,
              content: base64Bytes(attachment.content)
            }))
          } }
        : edit);
      const result = api.EditDocEnvelope(requireBytes(input), JSON.stringify(serialized));
      return result instanceof Uint8Array ? result : new Uint8Array(result);
    },
    async resolveDocx(input, request = {}) {
      const result = api.ResolveDocx(requireBytes(input), JSON.stringify(request));
      return result instanceof Uint8Array ? result : new Uint8Array(result);
    }
  });
}
