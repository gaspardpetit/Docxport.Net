using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;
using System.Text.Json;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocxportNet.API;
using DocxportNet.Fields;
using DocxportNet.Omml;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Wasm;

public static partial class BrowserExports
{
    [JSExport]
    [SupportedOSPlatform("browser")]
    public static string ConvertOmml(string omml, string format) => format.ToLowerInvariant() switch
    {
        "mathml" or "html" => DxpOmmlConverter.ToMathMl(omml),
        "latex" => DxpOmmlConverter.ToLatex(omml),
        "unicodemath" => DxpOmmlConverter.ToUnicodeMath(omml),
        "text" => DxpOmmlConverter.ToText(omml),
        _ => throw new ArgumentException("OMML format must be mathml, html, latex, unicodemath, or text.", nameof(format))
    };

    [JSExport]
    [SupportedOSPlatform("browser")]
    public static string Export(byte[] docxBytes, string requestJson)
        => ExportCore(docxBytes, DeserializeExportRequest(requestJson));

    private static string ExportCore(byte[] docxBytes, BrowserExportRequest request)
    {
        ValidateBytes(docxBytes);
        var eval = CreateFieldEval(request.Fields);
        var exportOptions = CreateExportOptionsOrDefault(request.Fields);

        return request.Format switch
        {
            BrowserExportFormat.Html => DxpExport.ExportToString(docxBytes,
                new DxpHtmlVisitor(CreateHtmlConfig(request), fieldEval: eval), exportOptions),
            BrowserExportFormat.Markdown => DxpExport.ExportToString(docxBytes,
                new DxpMarkdownVisitor(CreateMarkdownConfig(request), fieldEval: eval), exportOptions),
            BrowserExportFormat.Text => DxpExport.ExportToString(docxBytes,
                new DxpPlainTextVisitor(CreateTextConfig(request), fieldEval: eval), exportOptions),
            _ => throw new ArgumentOutOfRangeException(nameof(request.Format))
        };
    }

    [JSExport]
    [SupportedOSPlatform("browser")]
    public static byte[] ResolveDocx(byte[] docxBytes, string requestJson)
        => ResolveDocxCore(docxBytes, DeserializeResolveRequest(requestJson));

    [JSExport]
    [SupportedOSPlatform("browser")]
    public static string Inspect(byte[] docxBytes)
        => JsonSerializer.Serialize(InspectCore(docxBytes), BrowserJsonContext.Default.BrowserDocumentInfo);

    public static BrowserDocumentInfo InspectForTests(byte[] docxBytes) => InspectCore(docxBytes);

    private static BrowserDocumentInfo InspectCore(byte[] docxBytes)
    {
        ValidateBytes(docxBytes);
        using var stream = new MemoryStream(docxBytes, writable: false);
        using var document = WordprocessingDocument.Open(stream, false);
        return new BrowserDocumentInfo { HasTrackedChanges = EnumerateStoryRoots(document).Any(HasTrackedChanges) };
    }

    private static byte[] ResolveDocxCore(byte[] docxBytes, BrowserResolveRequest request)
    {
        ValidateBytes(docxBytes);
        var fields = request.Fields ?? new BrowserFieldOptions();
        return DxpDocxExport.Export(docxBytes,
            CreateExportOptions(fields, BrowserFieldMode.Evaluate),
            fieldEval: CreateFieldEval(fields));
    }

    public static string ExportForTests(byte[] docxBytes, BrowserExportRequest request) =>
        ExportCore(docxBytes, request);

    public static byte[] ResolveDocxForTests(byte[] docxBytes, BrowserResolveRequest request) =>
        ResolveDocxCore(docxBytes, request);

    private static BrowserExportRequest DeserializeExportRequest(string json)
    {
        if (string.IsNullOrWhiteSpace(json))
            throw new ArgumentException("A JSON request is required.", nameof(json));
        return JsonSerializer.Deserialize(json, BrowserJsonContext.Default.BrowserExportRequest)
            ?? throw new ArgumentException("The JSON request was empty.", nameof(json));
    }

    private static BrowserResolveRequest DeserializeResolveRequest(string json)
    {
        if (string.IsNullOrWhiteSpace(json))
            throw new ArgumentException("A JSON request is required.", nameof(json));
        return JsonSerializer.Deserialize(json, BrowserJsonContext.Default.BrowserResolveRequest)
            ?? throw new ArgumentException("The JSON request was empty.", nameof(json));
    }

    private static void ValidateBytes(byte[] bytes)
    {
        if (bytes == null || bytes.Length == 0)
            throw new ArgumentException("A non-empty DOCX byte array is required.", nameof(bytes));
    }

    private static DxpFieldEval CreateFieldEval(BrowserFieldOptions? fields)
    {
        var eval = new DxpFieldEval();
        if (fields?.Variables != null)
            foreach (var item in fields.Variables)
                eval.Context.SetDocVariable(item.Key, item.Value);
        return eval;
    }

    private static DxpExportOptions CreateExportOptionsOrDefault(BrowserFieldOptions? fields) =>
        CreateExportOptions(fields ?? new BrowserFieldOptions(), BrowserFieldMode.Cache);

    private static DxpExportOptions CreateExportOptions(
        BrowserFieldOptions fields,
        BrowserFieldMode defaultMode = BrowserFieldMode.Cache) => new()
    {
        FieldEvalMode = (fields.Mode ?? defaultMode) switch
        {
            BrowserFieldMode.None => DxpFieldEvalExportMode.None,
            BrowserFieldMode.Evaluate => DxpFieldEvalExportMode.Evaluate,
            _ => DxpFieldEvalExportMode.Cache
        }
    };

    private static DxpHtmlVisitorConfig CreateHtmlConfig(BrowserExportRequest request)
    {
        var config = request.Preset == BrowserPreset.Plain
            ? DxpHtmlVisitorConfig.CreatePlainConfig()
            : DxpHtmlVisitorConfig.CreateRichConfig();
        config.TrackedChangeMode = DxpTrackedChangeMode.AcceptChanges;
        var o = request.Html;
        if (o == null) return config;
        if (o.MathOutputFormat.HasValue) config.MathOutputFormat = ToMathOutputFormat(o.MathOutputFormat.Value);
        if (o.EmitImages.HasValue) config.EmitImages = o.EmitImages.Value;
        if (o.EmitParagraphMetadata.HasValue) config.EmitParagraphMetadata = o.EmitParagraphMetadata.Value;
        if (o.EmitStyleFont.HasValue) config.EmitStyleFont = o.EmitStyleFont.Value;
        if (o.EmitRunColor.HasValue) config.EmitRunColor = o.EmitRunColor.Value;
        if (o.EmitRunBackground.HasValue) config.EmitRunBackground = o.EmitRunBackground.Value;
        if (o.EmitTableBorders.HasValue) config.EmitTableBorders = o.EmitTableBorders.Value;
        if (o.EmitDocumentColors.HasValue) config.EmitDocumentColors = o.EmitDocumentColors.Value;
        if (o.EmitParagraphAlignment.HasValue) config.EmitParagraphAlignment = o.EmitParagraphAlignment.Value;
        if (o.PreserveListSymbols.HasValue) config.PreserveListSymbols = o.PreserveListSymbols.Value;
        if (o.RichTables.HasValue) config.RichTables = o.RichTables.Value;
        if (o.EmitSectionHeadersFooters.HasValue) config.EmitSectionHeadersFooters = o.EmitSectionHeadersFooters.Value;
        if (o.EmitUnreferencedBookmarks.HasValue) config.EmitUnreferencedBookmarks = o.EmitUnreferencedBookmarks.Value;
        if (o.EmitPageNumbers.HasValue) config.EmitPageNumbers = o.EmitPageNumbers.Value;
        if (o.EmitFieldInstructions.HasValue) config.EmitFieldInstructions = o.EmitFieldInstructions.Value;
        if (o.UsePlainComments.HasValue) config.UsePlainComments = o.UsePlainComments.Value;
        if (o.EmitCustomProperties.HasValue) config.EmitCustomProperties = o.EmitCustomProperties.Value;
        if (o.EmitTimeline.HasValue) config.EmitTimeline = o.EmitTimeline.Value;
        if (o.StylesheetHref != null) config.StylesheetHref = o.StylesheetHref;
        if (o.EmbedDefaultStylesheet.HasValue) config.EmbedDefaultStylesheet = o.EmbedDefaultStylesheet.Value;
        if (o.RootCssClass != null) config.RootCssClass = o.RootCssClass;
        if (o.TrackedChangeMode.HasValue) config.TrackedChangeMode = ToTrackedMode(o.TrackedChangeMode.Value);
        if (o.HeaderSelection.HasValue) config.HeaderSelection = ToHeaderFooter(o.HeaderSelection.Value);
        if (o.FooterSelection.HasValue) config.FooterSelection = ToHeaderFooter(o.FooterSelection.Value);
        return config;
    }

    private static DxpMarkdownVisitorConfig CreateMarkdownConfig(BrowserExportRequest request)
    {
        var config = request.Preset == BrowserPreset.Plain
            ? DxpMarkdownVisitorConfig.CreatePlainConfig()
            : DxpMarkdownVisitorConfig.CreateRichConfig();
        config.TrackedChangeMode = DxpTrackedChangeMode.AcceptChanges;
        var o = request.Markdown;
        if (o == null) return config;
        if (o.MathOutputFormat.HasValue) config.MathOutputFormat = ToMathOutputFormat(o.MathOutputFormat.Value);
        if (o.EmitMathDelimiters.HasValue) config.EmitMathDelimiters = o.EmitMathDelimiters.Value;
        if (o.EmitImages.HasValue) config.EmitImages = o.EmitImages.Value;
        if (o.EmitStyleFont.HasValue) config.EmitStyleFont = o.EmitStyleFont.Value;
        if (o.EmitRunColor.HasValue) config.EmitRunColor = o.EmitRunColor.Value;
        if (o.EmitRunBackground.HasValue) config.EmitRunBackground = o.EmitRunBackground.Value;
        if (o.EmitTableBorders.HasValue) config.EmitTableBorders = o.EmitTableBorders.Value;
        if (o.EmitDocumentColors.HasValue) config.EmitDocumentColors = o.EmitDocumentColors.Value;
        if (o.EmitParagraphAlignment.HasValue) config.EmitParagraphAlignment = o.EmitParagraphAlignment.Value;
        if (o.EmitRichLayoutHtml.HasValue) config.EmitRichLayoutHtml = o.EmitRichLayoutHtml.Value;
        if (o.PreserveListSymbols.HasValue) config.PreserveListSymbols = o.PreserveListSymbols.Value;
        if (o.RichTables.HasValue) config.RichTables = o.RichTables.Value;
        if (o.UsePlainCodeBlocks.HasValue) config.UsePlainCodeBlocks = o.UsePlainCodeBlocks.Value;
        if (o.UseMarkdownInlineStyles.HasValue) config.UseMarkdownInlineStyles = o.UseMarkdownInlineStyles.Value;
        if (o.EmitSectionHeadersFooters.HasValue) config.EmitSectionHeadersFooters = o.EmitSectionHeadersFooters.Value;
        if (o.EmitUnreferencedBookmarks.HasValue) config.EmitUnreferencedBookmarks = o.EmitUnreferencedBookmarks.Value;
        if (o.EmitPageNumbers.HasValue) config.EmitPageNumbers = o.EmitPageNumbers.Value;
        if (o.EmitFieldInstructions.HasValue) config.EmitFieldInstructions = o.EmitFieldInstructions.Value;
        if (o.UsePlainComments.HasValue) config.UsePlainComments = o.UsePlainComments.Value;
        if (o.EmitCustomProperties.HasValue) config.EmitCustomProperties = o.EmitCustomProperties.Value;
        if (o.EmitTimeline.HasValue) config.EmitTimeline = o.EmitTimeline.Value;
        if (o.TrackedChangeMode.HasValue) config.TrackedChangeMode = ToTrackedMode(o.TrackedChangeMode.Value);
        return config;
    }

    private static DxpPlainTextVisitorConfig CreateTextConfig(BrowserExportRequest request)
    {
        var config = DxpPlainTextVisitorConfig.CreateAcceptConfig();
        var o = request.Text;
        if (o == null) return config;
        if (o.MathOutputFormat.HasValue) config.MathOutputFormat = ToMathOutputFormat(o.MathOutputFormat.Value);
        if (o.TrackedChangeMode.HasValue)
            config.TrackedChangeMode = o.TrackedChangeMode == BrowserTrackedChangeMode.Reject
                ? DxpPlainTextTrackedChangeMode.RejectChanges
                : DxpPlainTextTrackedChangeMode.AcceptChanges;
        if (o.ImagePlaceholder != null) config.ImagePlaceholder = o.ImagePlaceholder;
        if (o.EmitDocumentProperties.HasValue) config.EmitDocumentProperties = o.EmitDocumentProperties.Value;
        if (o.EmitCustomProperties.HasValue) config.EmitCustomProperties = o.EmitCustomProperties.Value;
        return config;
    }

    private static DxpTrackedChangeMode ToTrackedMode(BrowserTrackedChangeMode mode) => mode switch
    {
        BrowserTrackedChangeMode.Accept => DxpTrackedChangeMode.AcceptChanges,
        BrowserTrackedChangeMode.Reject => DxpTrackedChangeMode.RejectChanges,
        BrowserTrackedChangeMode.Split => DxpTrackedChangeMode.SplitChanges,
        _ => DxpTrackedChangeMode.InlineChanges
    };

    private static DxpHeaderFooterSelection ToHeaderFooter(BrowserHeaderFooterSelection value) => value switch
    {
        BrowserHeaderFooterSelection.None => DxpHeaderFooterSelection.None,
        BrowserHeaderFooterSelection.Last => DxpHeaderFooterSelection.Last,
        _ => DxpHeaderFooterSelection.First
    };

    private static DxpOmmlOutputFormat? ToMathOutputFormat(BrowserMathOutputFormat value) => value switch
    {
        BrowserMathOutputFormat.None => null,
        BrowserMathOutputFormat.MathMl => DxpOmmlOutputFormat.MathMl,
        BrowserMathOutputFormat.Latex => DxpOmmlOutputFormat.Latex,
        BrowserMathOutputFormat.UnicodeMath => DxpOmmlOutputFormat.UnicodeMath,
        _ => DxpOmmlOutputFormat.Text,
    };

    private static IEnumerable<OpenXmlElement> EnumerateStoryRoots(WordprocessingDocument document)
    {
        var main = document.MainDocumentPart;
        if (main?.Document != null) yield return main.Document;
        if (main == null) yield break;
        foreach (var part in main.HeaderParts)
            if (part.Header != null) yield return part.Header;
        foreach (var part in main.FooterParts)
            if (part.Footer != null) yield return part.Footer;
        if (main.FootnotesPart?.Footnotes != null) yield return main.FootnotesPart.Footnotes;
        if (main.EndnotesPart?.Endnotes != null) yield return main.EndnotesPart.Endnotes;
    }

    private static bool HasTrackedChanges(OpenXmlElement root) =>
        root.Descendants().Any(element => TrackedChangeNames.Contains(element.LocalName));

    private static readonly HashSet<string> TrackedChangeNames = new(StringComparer.Ordinal)
    {
        "ins", "del", "moveFrom", "moveTo", "moveFromRangeStart", "moveFromRangeEnd",
        "moveToRangeStart", "moveToRangeEnd", "conflictIns", "conflictDel"
    };
}
