using DocumentFormat.OpenXml.Packaging;
using DocxportNet.Fields;
using DocxportNet.Fields.Resolution;
using DocxportNet.Visitors.PlainText;
using Microsoft.Extensions.Logging;

namespace DocxportNet;

public static class DxpMergeLoop
{
    public static IReadOnlyList<string> MergePlainText(
        string docxPath,
        IDxpMergeRecordCursor cursor,
        DxpPlainTextVisitorConfig? config = null,
        DxpFieldEval? fieldEval = null,
        DxpExportOptions? options = null,
        ILogger? logger = null)
    {
        using var document = WordprocessingDocument.Open(docxPath, false);
        return MergePlainText(document, cursor, config, fieldEval, options, logger);
    }

    public static IReadOnlyList<string> MergePlainText(
        WordprocessingDocument document,
        IDxpMergeRecordCursor cursor,
        DxpPlainTextVisitorConfig? config = null,
        DxpFieldEval? fieldEval = null,
        DxpExportOptions? options = null,
        ILogger? logger = null)
    {
        var visitorConfig = config ?? DxpPlainTextVisitorConfig.CreateAcceptConfig();
        return DxpExport.ExportToStrings(
            document,
            cursor,
            eval => new DxpPlainTextVisitor(visitorConfig, logger, eval),
            fieldEval,
            options,
            logger);
    }

    public static IReadOnlyList<string> MergeHtml(
        string docxPath,
        IDxpMergeRecordCursor cursor,
        DocxportNet.Visitors.Html.DxpHtmlVisitorConfig? config = null,
        DxpFieldEval? fieldEval = null,
        DxpExportOptions? options = null,
        ILogger? logger = null)
    {
        using var document = WordprocessingDocument.Open(docxPath, false);
        return MergeHtml(document, cursor, config, fieldEval, options, logger);
    }

    public static IReadOnlyList<string> MergeHtml(
        WordprocessingDocument document,
        IDxpMergeRecordCursor cursor,
        DocxportNet.Visitors.Html.DxpHtmlVisitorConfig? config = null,
        DxpFieldEval? fieldEval = null,
        DxpExportOptions? options = null,
        ILogger? logger = null)
    {
        var visitorConfig = config ?? DocxportNet.Visitors.Html.DxpHtmlVisitorConfig.CreateRichConfig();
        return DxpExport.ExportToStrings(
            document,
            cursor,
            eval => new DocxportNet.Visitors.Html.DxpHtmlVisitor(visitorConfig, logger, eval),
            fieldEval,
            options,
            logger);
    }

    public static IReadOnlyList<string> MergeMarkdown(
        string docxPath,
        IDxpMergeRecordCursor cursor,
        DocxportNet.Visitors.Markdown.DxpMarkdownVisitorConfig? config = null,
        DxpFieldEval? fieldEval = null,
        DxpExportOptions? options = null,
        ILogger? logger = null)
    {
        using var document = WordprocessingDocument.Open(docxPath, false);
        return MergeMarkdown(document, cursor, config, fieldEval, options, logger);
    }

    public static IReadOnlyList<string> MergeMarkdown(
        WordprocessingDocument document,
        IDxpMergeRecordCursor cursor,
        DocxportNet.Visitors.Markdown.DxpMarkdownVisitorConfig? config = null,
        DxpFieldEval? fieldEval = null,
        DxpExportOptions? options = null,
        ILogger? logger = null)
    {
        var visitorConfig = config ?? DocxportNet.Visitors.Markdown.DxpMarkdownVisitorConfig.CreateRichConfig();
        return DxpExport.ExportToStrings(
            document,
            cursor,
            eval => new DocxportNet.Visitors.Markdown.DxpMarkdownVisitor(visitorConfig, logger, eval),
            fieldEval,
            options,
            logger);
    }
}
