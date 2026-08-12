using DocxportNet;
using DocxportNet.API;
using DocxportNet.Fields;
using DocxportNet.Fields.Resolution;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;
using Microsoft.Extensions.Logging;
using System.Text.Json;

if (args.Length == 0 || args.Contains("--help") || args.Contains("-h"))
{
    PrintHelp();
    return;
}

if (args.Contains("--version") || args.Contains("-v"))
{
    Console.WriteLine(GetVersion());
    return;
}

string? inputPath = null;
string? outputPath = null;
string format = "markdown";
string tracked = "accept";
bool plainOutput = false;
bool formatExplicit = false;
var cliVariables = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
var includePaths = new List<string>();
string? varsPath = null;
LogLevel logLevel = LogLevel.Warning;
DxpFieldEvalExportMode fieldMode = DxpFieldEvalExportMode.Cache;

for (int i = 0; i < args.Length; i++)
{
    string arg = args[i];

    if (arg.StartsWith("--format=", StringComparison.OrdinalIgnoreCase))
    {
        format = arg[(arg.IndexOf('=') + 1)..];
        formatExplicit = true;
    }
    else if (arg.Equals("--format", StringComparison.OrdinalIgnoreCase))
    {
        if (i + 1 >= args.Length)
        {
            Console.Error.WriteLine("--format requires a value.");
            return;
        }
        format = args[++i];
        formatExplicit = true;
    }
    else if (arg.StartsWith("--tracked=", StringComparison.OrdinalIgnoreCase))
        tracked = arg[(arg.IndexOf('=') + 1)..];
    else if (arg.Equals("--tracked", StringComparison.OrdinalIgnoreCase))
    {
        if (i + 1 >= args.Length)
        {
            Console.Error.WriteLine("--tracked requires a value.");
            return;
        }
        tracked = args[++i];
    }
    else if (arg.Equals("--plain", StringComparison.OrdinalIgnoreCase))
        plainOutput = true;
    else if (arg.StartsWith("--fields=", StringComparison.OrdinalIgnoreCase))
        fieldMode = ParseFieldMode(arg[(arg.IndexOf('=') + 1)..]);
    else if (arg.Equals("--fields", StringComparison.OrdinalIgnoreCase))
    {
        if (i + 1 >= args.Length)
        {
            Console.Error.WriteLine("--fields requires a value.");
            return;
        }
        fieldMode = ParseFieldMode(args[++i]);
    }
    else if (arg.StartsWith("--vars=", StringComparison.OrdinalIgnoreCase))
        varsPath = arg[(arg.IndexOf('=') + 1)..];
    else if (arg.StartsWith("--include-path=", StringComparison.OrdinalIgnoreCase))
        includePaths.Add(arg[(arg.IndexOf('=') + 1)..]);
    else if (arg.Equals("--include-path", StringComparison.OrdinalIgnoreCase))
    {
        if (i + 1 >= args.Length)
        {
            Console.Error.WriteLine("--include-path requires a directory path.");
            return;
        }
        includePaths.Add(args[++i]);
    }
    else if (arg.Equals("--vars", StringComparison.OrdinalIgnoreCase))
    {
        if (i + 1 >= args.Length)
        {
            Console.Error.WriteLine("--vars requires a file path.");
            return;
        }
        varsPath = args[++i];
    }
    else if (arg.StartsWith("-D", StringComparison.Ordinal))
    {
        var spec = arg.Length > 2 && arg[2] == '=' ? arg[3..] : arg[2..];
        if (string.IsNullOrWhiteSpace(spec))
        {
            if (i + 1 >= args.Length)
            {
                Console.Error.WriteLine("-D requires name=value.");
                return;
            }
            spec = args[++i];
        }

        var equals = spec.IndexOf('=');
        if (equals <= 0)
        {
            Console.Error.WriteLine("-D requires name=value.");
            return;
        }
        var name = spec[..equals].Trim();
        var value = spec[(equals + 1)..];
        if (string.IsNullOrWhiteSpace(name))
        {
            Console.Error.WriteLine("-D requires name=value.");
            return;
        }
        cliVariables[name] = value;
    }
    else if (arg.StartsWith("--output=", StringComparison.OrdinalIgnoreCase))
        outputPath = arg[(arg.IndexOf('=') + 1)..];
    else if (arg.Equals("--output", StringComparison.OrdinalIgnoreCase))
    {
        if (i + 1 >= args.Length)
        {
            Console.Error.WriteLine("--output requires a value.");
            return;
        }
        outputPath = args[++i];
    }
    else if (arg.StartsWith("-o=", StringComparison.OrdinalIgnoreCase))
        outputPath = arg[(arg.IndexOf('=') + 1)..];
    else if (arg.Equals("-o", StringComparison.OrdinalIgnoreCase))
    {
        if (i + 1 >= args.Length)
        {
            Console.Error.WriteLine("-o requires a value.");
            return;
        }
        outputPath = args[++i];
    }
    else if (arg.StartsWith("--log-level=", StringComparison.OrdinalIgnoreCase))
        logLevel = ParseLogLevel(arg[(arg.IndexOf('=') + 1)..]);
    else if (arg.Equals("--log-level", StringComparison.OrdinalIgnoreCase))
    {
        if (i + 1 >= args.Length)
        {
            Console.Error.WriteLine("--log-level requires a value.");
            return;
        }
        logLevel = ParseLogLevel(args[++i]);
    }
    else if (arg.Equals("--log", StringComparison.OrdinalIgnoreCase))
        logLevel = LogLevel.Information;
    else if (inputPath is null)
        inputPath = arg;
}

if (string.IsNullOrWhiteSpace(inputPath))
{
    Console.Error.WriteLine("Input DOCX path is required.");
    PrintHelp();
    return;
}

if (!File.Exists(inputPath))
{
    Console.Error.WriteLine($"Input file not found: {inputPath}");
    return;
}

DxpTrackedChangeMode trackedMode = ParseTrackedChangeMode(tracked);

if (!formatExplicit && !string.IsNullOrWhiteSpace(outputPath))
{
    string ext = Path.GetExtension(outputPath).ToLowerInvariant();
    if (ext is ".html" or ".htm")
        format = "html";
    else if (ext is ".md" or ".markdown")
        format = "markdown";
    else if (ext is ".txt")
        format = "text";
}

switch (format.ToLowerInvariant())
{
    case "markdown":
    case "md":
        ExportMarkdown(inputPath, outputPath, trackedMode, plainOutput, fieldMode, varsPath, cliVariables, includePaths, logLevel);
        break;
    case "html":
        ExportHtml(inputPath, outputPath, trackedMode, plainOutput, fieldMode, varsPath, cliVariables, includePaths, logLevel);
        break;
    case "text":
    case "txt":
        if (plainOutput)
            Console.Error.WriteLine("Warning: --plain is only supported for markdown/html; ignoring.");
        ExportPlainText(inputPath, outputPath, trackedMode, fieldMode, varsPath, cliVariables, includePaths, logLevel);
        break;
    default:
        Console.Error.WriteLine($"Unknown format '{format}'. Expected markdown|html|text.");
        PrintHelp();
        break;
}

static void ExportMarkdown(
    string inputPath,
    string? outputPath,
    DxpTrackedChangeMode trackedMode,
    bool plainOutput,
    DxpFieldEvalExportMode fieldMode,
    string? varsPath,
    IReadOnlyDictionary<string, string> cliVariables,
    IReadOnlyList<string> includePaths,
    LogLevel logLevel)
{
    var config = plainOutput ? DxpMarkdownVisitorConfig.CreatePlainConfig() : DxpMarkdownVisitorConfig.CreateRichConfig();
    config = config with { TrackedChangeMode = trackedMode };

    string output = outputPath ?? Path.ChangeExtension(inputPath, plainOutput ? ".plain.md" : ".md");
    using var loggerFactory = CreateLoggerFactory(logLevel);
    var logger = loggerFactory.CreateLogger("docxport");
    var visitor = new DxpMarkdownVisitor(config, logger);
    ApplyFieldContext(visitor, varsPath, cliVariables, includePaths);
    var exportOptions = new DxpExportOptions { FieldEvalMode = fieldMode };
    DxpExport.ExportToFile(inputPath, visitor, output, exportOptions, logger);
    Console.WriteLine($"Wrote Markdown to {output}");
}

static void ExportHtml(
    string inputPath,
    string? outputPath,
    DxpTrackedChangeMode trackedMode,
    bool plainOutput,
    DxpFieldEvalExportMode fieldMode,
    string? varsPath,
    IReadOnlyDictionary<string, string> cliVariables,
    IReadOnlyList<string> includePaths,
    LogLevel logLevel)
{
    var config = (plainOutput ? DxpHtmlVisitorConfig.CreatePlainConfig() : DxpHtmlVisitorConfig.CreateRichConfig()) with { TrackedChangeMode = trackedMode };
    string output = outputPath ?? Path.ChangeExtension(inputPath, trackedMode == DxpTrackedChangeMode.RejectChanges ? ".reject.html" : ".html");
    using var loggerFactory = CreateLoggerFactory(logLevel);
    var logger = loggerFactory.CreateLogger("docxport");
    var visitor = new DxpHtmlVisitor(config, logger);
    ApplyFieldContext(visitor, varsPath, cliVariables, includePaths);
    var exportOptions = new DxpExportOptions { FieldEvalMode = fieldMode };
    DxpExport.ExportToFile(inputPath, visitor, output, exportOptions, logger);
    Console.WriteLine($"Wrote HTML to {output}");
}

static void ExportPlainText(
    string inputPath,
    string? outputPath,
    DxpTrackedChangeMode trackedMode,
    DxpFieldEvalExportMode fieldMode,
    string? varsPath,
    IReadOnlyDictionary<string, string> cliVariables,
    IReadOnlyList<string> includePaths,
    LogLevel logLevel)
{
    var textMode = trackedMode == DxpTrackedChangeMode.RejectChanges
        ? DxpPlainTextTrackedChangeMode.RejectChanges
        : DxpPlainTextTrackedChangeMode.AcceptChanges;
    var config = new DxpPlainTextVisitorConfig { TrackedChangeMode = textMode };
    string output = outputPath ?? Path.ChangeExtension(inputPath, textMode == DxpPlainTextTrackedChangeMode.RejectChanges ? ".reject.txt" : ".txt");
    using var loggerFactory = CreateLoggerFactory(logLevel);
    var logger = loggerFactory.CreateLogger("docxport");
    var visitor = new DxpPlainTextVisitor(config, logger);
    ApplyFieldContext(visitor, varsPath, cliVariables, includePaths);
    var exportOptions = new DxpExportOptions { FieldEvalMode = fieldMode };
    DxpExport.ExportToFile(inputPath, visitor, output, exportOptions, logger);
    Console.WriteLine($"Wrote text to {output}");
}

static DxpTrackedChangeMode ParseTrackedChangeMode(string value)
{
    return value.ToLowerInvariant() switch {
        "accept" => DxpTrackedChangeMode.AcceptChanges,
        "reject" => DxpTrackedChangeMode.RejectChanges,
        "inline" => DxpTrackedChangeMode.InlineChanges,
        "split" => DxpTrackedChangeMode.SplitChanges,
        _ => DxpTrackedChangeMode.AcceptChanges
    };
}

static LogLevel ParseLogLevel(string value)
{
    return value.ToLowerInvariant() switch {
        "trace" => LogLevel.Trace,
        "debug" => LogLevel.Debug,
        "info" => LogLevel.Information,
        "information" => LogLevel.Information,
        "warn" => LogLevel.Warning,
        "warning" => LogLevel.Warning,
        "error" => LogLevel.Error,
        "critical" => LogLevel.Critical,
        "none" => LogLevel.None,
        _ => LogLevel.Warning
    };
}

static DxpFieldEvalExportMode ParseFieldMode(string value)
{
    return value.ToLowerInvariant() switch {
        "none" => DxpFieldEvalExportMode.None,
        "evaluate" => DxpFieldEvalExportMode.Evaluate,
        "cache" => DxpFieldEvalExportMode.Cache,
        _ => DxpFieldEvalExportMode.Cache
    };
}

static ILoggerFactory CreateLoggerFactory(LogLevel minimumLevel)
{
    return LoggerFactory.Create(builder => {
        builder
            .SetMinimumLevel(minimumLevel)
            .AddSimpleConsole(options => {
                options.SingleLine = true;
                options.TimestampFormat = "HH:mm:ss ";
                options.UseUtcTimestamp = false;
            });
    });
}

static void PrintHelp()
{
    Console.WriteLine($"""
docxport ({GetVersion()})
Usage: docxport <input.docx> [--format=markdown|html|text] [--tracked=accept|reject|inline|split] [--plain] [--fields=evaluate|cache|none] [-o|--output=path] [--vars=path] [-D name=value] [--include-path=directory]

Options:
  --format=...   Output format (default: markdown)
                If --format is omitted, the format is inferred from -o/--output extension (.md/.html/.txt).
  --tracked=...  Tracked change mode (accept, reject, inline, split). Plain text supports accept/reject.
    --plain        Plain output for markdown/html (minimal styling/features)
    --fields=...   Field result mode (evaluate, cache, none). Default: cache.
  -o, --output=...  Output file path (default: swaps extension)
  --vars=...     Load DOCVARIABLE values from a JSON or INI file.
  -D name=value  Define a DOCVARIABLE (repeatable). CLI values override --vars.
  --include-path=...  Allow and search this directory for INCLUDETEXT DOCX files (repeatable).
  --log          Enable info-level logging to stderr.
  --log-level=...  Set log level (trace|debug|info|warn|error|critical|none).
  -v, --version  Show CLI version
  -h, --help     Show this help
""");
}

static string GetVersion()
{
    var attr = typeof(Program).Assembly
        .GetCustomAttributes(typeof(System.Reflection.AssemblyInformationalVersionAttribute), false)
        .OfType<System.Reflection.AssemblyInformationalVersionAttribute>()
        .FirstOrDefault();
    return attr?.InformationalVersion ?? "unknown";
}

static void ApplyFieldContext(
    DxpIVisitor visitor,
    string? varsPath,
    IReadOnlyDictionary<string, string> cliVariables,
    IReadOnlyList<string> includePaths)
{
    if (visitor is not DxpIFieldEvalProvider provider)
        return;

    var context = provider.FieldEval.Context;
    var fileVars = LoadVariables(varsPath);
    foreach (var kvp in fileVars)
        context.SetDocVariable(kvp.Key, kvp.Value);
    foreach (var kvp in cliVariables)
        context.SetDocVariable(kvp.Key, kvp.Value);
    if (includePaths.Count > 0)
        context.IncludeTextResolver = new DxpFileSystemIncludeTextResolver(includePaths);
}

static Dictionary<string, string> LoadVariables(string? varsPath)
{
    var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    if (string.IsNullOrWhiteSpace(varsPath))
        return result;
    if (!File.Exists(varsPath))
    {
        Console.Error.WriteLine($"Variables file not found: {varsPath}");
        return result;
    }

    var ext = Path.GetExtension(varsPath).ToLowerInvariant();
    try
    {
        if (ext == ".json")
            return LoadVariablesFromJson(varsPath);
        if (ext == ".ini")
            return LoadVariablesFromIni(varsPath);
        Console.Error.WriteLine($"Unsupported vars file extension '{ext}'. Use .json or .ini.");
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"Failed to load vars file: {ex.Message}");
    }

    return result;
}

static Dictionary<string, string> LoadVariablesFromJson(string path)
{
    var json = File.ReadAllText(path);
    var data = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
    return data ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
}

static Dictionary<string, string> LoadVariablesFromIni(string path)
{
    var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    foreach (var rawLine in File.ReadAllLines(path))
    {
        var line = rawLine.Trim();
        if (line.Length == 0 || line.StartsWith("#", StringComparison.Ordinal) || line.StartsWith(";", StringComparison.Ordinal))
            continue;
        var idx = line.IndexOf('=');
        if (idx <= 0)
            continue;
        var key = line[..idx].Trim();
        var value = line[(idx + 1)..].Trim();
        if (key.Length == 0)
            continue;
        result[key] = value;
    }
    return result;
}
