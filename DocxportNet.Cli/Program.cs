using DocxportNet;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;
using DocxportNet.API;

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
bool plainMarkdown = false;
bool formatExplicit = false;

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
		plainMarkdown = true;
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
		ExportMarkdown(inputPath, outputPath, trackedMode, plainMarkdown);
		break;
	case "html":
		if (plainMarkdown)
			Console.Error.WriteLine("Warning: --plain is only supported for markdown; ignoring.");
		ExportHtml(inputPath, outputPath, trackedMode);
		break;
	case "text":
	case "txt":
		if (plainMarkdown)
			Console.Error.WriteLine("Warning: --plain is only supported for markdown; ignoring.");
		ExportPlainText(inputPath, outputPath, trackedMode);
		break;
	default:
		Console.Error.WriteLine($"Unknown format '{format}'. Expected markdown|html|text.");
		PrintHelp();
		break;
}

static void ExportMarkdown(string inputPath, string? outputPath, DxpTrackedChangeMode trackedMode, bool plainMarkdown)
{
	var config = plainMarkdown ? DxpMarkdownVisitorConfig.CreatePlainConfig(): DxpMarkdownVisitorConfig.CreateRichConfig();
	config = config with { TrackedChangeMode = trackedMode };

	string output = outputPath ?? Path.ChangeExtension(inputPath, plainMarkdown ? ".plain.md" : ".md");
	var visitor = new DxpMarkdownVisitor(config);
	DxpExport.ExportToFile(inputPath, visitor, output);
	Console.WriteLine($"Wrote Markdown to {output}");
}

static void ExportHtml(string inputPath, string? outputPath, DxpTrackedChangeMode trackedMode)
{
	var config = DxpHtmlVisitorConfig.CreateRichConfig() with { TrackedChangeMode = trackedMode };
	string output = outputPath ?? Path.ChangeExtension(inputPath, trackedMode == DxpTrackedChangeMode.RejectChanges ? ".reject.html" : ".html");
	var visitor = new DxpHtmlVisitor(config);
	DxpExport.ExportToFile(inputPath, visitor, output);
	Console.WriteLine($"Wrote HTML to {output}");
}

static void ExportPlainText(string inputPath, string? outputPath, DxpTrackedChangeMode trackedMode)
{
	var textMode = trackedMode == DxpTrackedChangeMode.RejectChanges
		? DxpPlainTextTrackedChangeMode.RejectChanges
		: DxpPlainTextTrackedChangeMode.AcceptChanges;
	var config = new DxpPlainTextVisitorConfig { TrackedChangeMode = textMode };
	string output = outputPath ?? Path.ChangeExtension(inputPath, textMode == DxpPlainTextTrackedChangeMode.RejectChanges ? ".reject.txt" : ".txt");
	var visitor = new DxpPlainTextVisitor(config);
	DxpExport.ExportToFile(inputPath, visitor, output);
	Console.WriteLine($"Wrote text to {output}");
}

static DxpTrackedChangeMode ParseTrackedChangeMode(string value)
{
	return value.ToLowerInvariant() switch
	{
		"accept" => DxpTrackedChangeMode.AcceptChanges,
		"reject" => DxpTrackedChangeMode.RejectChanges,
		"inline" => DxpTrackedChangeMode.InlineChanges,
		"split" => DxpTrackedChangeMode.SplitChanges,
		_ => DxpTrackedChangeMode.AcceptChanges
	};
}

static void PrintHelp()
{
	Console.WriteLine($"""
docxport ({GetVersion()})
Usage: docxport <input.docx> [--format=markdown|html|text] [--tracked=accept|reject|inline|split] [--plain] [-o|--output=path]

Options:
  --format=...   Output format (default: markdown)
                If --format is omitted, the format is inferred from -o/--output extension (.md/.html/.txt).
  --tracked=...  Tracked change mode (accept, reject, inline, split). Plain text supports accept/reject.
	--plain        Plain Markdown (only for markdown format)
  -o, --output=...  Output file path (default: swaps extension)
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
