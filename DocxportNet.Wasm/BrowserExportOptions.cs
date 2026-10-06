using System.Text.Json.Serialization;
using DocxportNet.Doc;
using DocxportNet;

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
    public DxpCoreMetadata CoreProperties { get; set; } = new();
    public DxpExtendedMetadata? ExtendedProperties { get; set; }
    public IReadOnlyList<DxpLanguageRatio>? Language { get; set; }
    public bool HasTrackedChanges { get; set; }
    public bool HasComments { get; set; }
}

public sealed class BrowserExportProgress
{
    public string Phase { get; set; } = string.Empty;
    public long CompletedUnits { get; set; }
    public long TotalUnits { get; set; }
    public double? Percentage { get; set; }
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

public sealed class BrowserDocEnvelopeEdit
{
    public string Operation { get; set; } = "";
    public BrowserDocEmailEnvelope? Envelope { get; set; }
    public bool? Visible { get; set; }
}

public sealed class BrowserDocEmailEnvelope
{
    public string? Subject { get; set; }
    public string? Introduction { get; set; }
    public IReadOnlyList<DocEmailAddress>? To { get; set; }
    public IReadOnlyList<DocEmailAddress>? Cc { get; set; }
    public IReadOnlyList<DocEmailAddress>? Bcc { get; set; }
    public IReadOnlyList<DocEmailAddress>? ReplyTo { get; set; }
    public IReadOnlyList<DocEmailAttachment>? Attachments { get; set; }
    public DocEmailImportance? Importance { get; set; }
    public DocEmailSensitivity? Sensitivity { get; set; }
    public bool? RequestDeliveryReceipt { get; set; }
    public bool? RequestReadReceipt { get; set; }
    public bool? Visible { get; set; }
    public string? Categories { get; set; }
    public DateTimeOffset? DeliverAfter { get; set; }
    public DateTimeOffset? ExpiresAt { get; set; }

    public DocEmailEnvelope ToModel() => new()
    {
        Subject = Subject ?? "", Introduction = Introduction ?? "",
        To = To ?? Array.Empty<DocEmailAddress>(),
        Cc = Cc ?? Array.Empty<DocEmailAddress>(),
        Bcc = Bcc ?? Array.Empty<DocEmailAddress>(),
        ReplyTo = ReplyTo ?? Array.Empty<DocEmailAddress>(),
        Attachments = Attachments ?? Array.Empty<DocEmailAttachment>(),
        Importance = Importance ?? DocEmailImportance.Normal,
        Sensitivity = Sensitivity ?? DocEmailSensitivity.Normal,
        RequestDeliveryReceipt = RequestDeliveryReceipt ?? false,
        RequestReadReceipt = RequestReadReceipt ?? false,
        Visible = Visible ?? true,
        Categories = Categories ?? "", DeliverAfter = DeliverAfter, ExpiresAt = ExpiresAt
    };
}

[JsonSourceGenerationOptions(
    PropertyNameCaseInsensitive = true,
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    UseStringEnumConverter = true)]
[JsonSerializable(typeof(BrowserExportRequest))]
[JsonSerializable(typeof(BrowserResolveRequest))]
[JsonSerializable(typeof(BrowserDocumentInfo))]
[JsonSerializable(typeof(BrowserExportProgress))]
[JsonSerializable(typeof(BrowserDocEnvelopeEdit[]))]
internal partial class BrowserJsonContext : JsonSerializerContext
{
}
