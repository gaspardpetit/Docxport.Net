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
  throw new TypeError("DOCX input must be a Uint8Array or ArrayBuffer.");
}

export async function createDocxport(options = {}) {
  initialization ??= initialize(options);
  const api = await initialization;

  return Object.freeze({
    async inspect(input) {
      return JSON.parse(api.Inspect(requireBytes(input)));
    },
    async export(input, request = {}) {
      return api.Export(requireBytes(input), JSON.stringify(request));
    },
    async resolveDocx(input, request = {}) {
      const result = api.ResolveDocx(requireBytes(input), JSON.stringify(request));
      return result instanceof Uint8Array ? result : new Uint8Array(result);
    }
  });
}
