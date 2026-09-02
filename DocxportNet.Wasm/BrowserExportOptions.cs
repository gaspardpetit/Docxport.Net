using System.Text.Json.Serialization;

namespace DocxportNet.Wasm;

[JsonConverter(typeof(JsonStringEnumConverter<BrowserExportFormat>))]
public enum BrowserExportFormat { Html, Markdown, Text }

[JsonConverter(typeof(JsonStringEnumConverter<BrowserPreset>))]
public enum BrowserPreset { Rich, Plain }

[JsonConverter(typeof(JsonStringEnumConverter<BrowserFieldMode>))]
public enum BrowserFieldMode { None, Evaluate, Cache }

[JsonConverter(typeof(JsonStringEnumConverter<BrowserTrackedChangeMode>))]
public enum BrowserTrackedChangeMode { Accept, Reject, Inline, Split }

[JsonConverter(typeof(JsonStringEnumConverter<BrowserHeaderFooterSelection>))]
public enum BrowserHeaderFooterSelection { None, First, Last }

[JsonConverter(typeof(JsonStringEnumConverter<BrowserMathOutputFormat>))]
public enum BrowserMathOutputFormat { None, MathMl, Latex, UnicodeMath, Text }

[JsonConverter(typeof(JsonStringEnumConverter<BrowserMathDelimiterStyle>))]
public enum BrowserMathDelimiterStyle { Dollar, Backslash, Auto }

public sealed class BrowserExportRequest
{
    public BrowserExportFormat Format { get; set; } = BrowserExportFormat.Html;
    public BrowserPreset Preset { get; set; } = BrowserPreset.Rich;
    public BrowserFieldOptions? Fields { get; set; }
    public BrowserHtmlOptions? Html { get; set; }
    public BrowserMarkdownOptions? Markdown { get; set; }
    public BrowserTextOptions? Text { get; set; }
}

public sealed class BrowserResolveRequest
{
    public BrowserFieldOptions? Fields { get; set; }
}

public sealed class BrowserDocumentInfo
{
    public bool HasTrackedChanges { get; set; }
}

public sealed class BrowserFieldOptions
{
    public BrowserFieldMode? Mode { get; set; }
    public Dictionary<string, string?>? Variables { get; set; }
}

public sealed class BrowserHtmlOptions
{
    public BrowserMathOutputFormat? MathOutputFormat { get; set; }
    public bool? EmitImages { get; set; }
    public bool? EmitParagraphMetadata { get; set; }
    public bool? EmitStyleFont { get; set; }
    public bool? EmitRunColor { get; set; }
    public bool? EmitRunBackground { get; set; }
    public bool? EmitTableBorders { get; set; }
    public bool? EmitDocumentColors { get; set; }
    public bool? EmitParagraphAlignment { get; set; }
    public bool? PreserveListSymbols { get; set; }
    public bool? RichTables { get; set; }
    public bool? EmitSectionHeadersFooters { get; set; }
    public bool? EmitUnreferencedBookmarks { get; set; }
    public bool? EmitPageNumbers { get; set; }
    public bool? EmitFieldInstructions { get; set; }
    public bool? UsePlainComments { get; set; }
    public bool? EmitCustomProperties { get; set; }
    public bool? EmitTimeline { get; set; }
    public string? StylesheetHref { get; set; }
    public bool? EmbedDefaultStylesheet { get; set; }
    public string? RootCssClass { get; set; }
    public BrowserTrackedChangeMode? TrackedChangeMode { get; set; }
    public BrowserHeaderFooterSelection? HeaderSelection { get; set; }
    public BrowserHeaderFooterSelection? FooterSelection { get; set; }
}

public sealed class BrowserMarkdownOptions
{
    public BrowserMathOutputFormat? MathOutputFormat { get; set; }
    public bool? EmitMathDelimiters { get; set; }
    public BrowserMathDelimiterStyle? MathDelimiterStyle { get; set; }
    public bool? EmitImages { get; set; }
    public bool? EmitStyleFont { get; set; }
    public bool? EmitRunColor { get; set; }
    public bool? EmitRunBackground { get; set; }
    public bool? EmitTableBorders { get; set; }
    public bool? EmitDocumentColors { get; set; }
    public bool? EmitParagraphAlignment { get; set; }
    public bool? EmitRichLayoutHtml { get; set; }
    public bool? PreserveListSymbols { get; set; }
    public bool? RichTables { get; set; }
    public bool? UsePlainCodeBlocks { get; set; }
    public bool? UseMarkdownInlineStyles { get; set; }
    public bool? EmitSectionHeadersFooters { get; set; }
    public bool? EmitUnreferencedBookmarks { get; set; }
    public bool? EmitPageNumbers { get; set; }
    public bool? EmitFieldInstructions { get; set; }
    public bool? UsePlainComments { get; set; }
    public bool? EmitCustomProperties { get; set; }
    public bool? EmitTimeline { get; set; }
    public BrowserTrackedChangeMode? TrackedChangeMode { get; set; }
}

public sealed class BrowserTextOptions
{
    public BrowserMathOutputFormat? MathOutputFormat { get; set; }
    public BrowserTrackedChangeMode? TrackedChangeMode { get; set; }
    public string? ImagePlaceholder { get; set; }
    public bool? EmitDocumentProperties { get; set; }
    public bool? EmitCustomProperties { get; set; }
}

[JsonSourceGenerationOptions(
    PropertyNameCaseInsensitive = true,
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    UseStringEnumConverter = true)]
[JsonSerializable(typeof(BrowserExportRequest))]
[JsonSerializable(typeof(BrowserResolveRequest))]
[JsonSerializable(typeof(BrowserDocumentInfo))]
internal partial class BrowserJsonContext : JsonSerializerContext
{
}
