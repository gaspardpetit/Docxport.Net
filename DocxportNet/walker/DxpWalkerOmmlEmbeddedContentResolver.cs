using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Omml;
using DocxportNet.Visitors.PlainText;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Walker;

/// <summary>
/// Resolves embedded WordprocessingML through Docxport's normal walker and plain-text visitor.
/// This adapter is intended for document-pipeline integrations; standalone callers can omit it.
/// </summary>
public sealed class DxpWalkerOmmlEmbeddedContentResolver : IDxpOmmlEmbeddedContentResolver
{
    private readonly ILogger? _logger;

    public DxpWalkerOmmlEmbeddedContentResolver(ILogger? logger = null)
    {
        _logger = logger;
    }

    public string? Resolve(DxpOmmlEmbeddedContentRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        if (request.OutputFormat != DxpOmmlOutputFormat.Latex)
            return null;
        if (request.RevisionMode == DxpOmmlRevisionMode.Preserve || request.FieldMode == DxpOmmlFieldMode.Omit)
            return null;

        using MemoryStream stream = new();
        using (WordprocessingDocument document = WordprocessingDocument.Create(
                   stream, WordprocessingDocumentType.Document, true))
        {
            MainDocumentPart main = document.AddMainDocumentPart();
            Paragraph paragraph = new();
            if (request.OpenXmlElements.Count != 0)
                paragraph.Append(request.OpenXmlElements.Select(element => element.CloneNode(true)));
            else if (request.XmlElements.Count != 0)
                paragraph.InnerXml = string.Concat(request.XmlElements.Select(element =>
                    element.ToString(System.Xml.Linq.SaveOptions.DisableFormatting)));
            else
                return null;
            main.Document = new Document(new Body(paragraph));
            main.Document.Save();
        }

        stream.Position = 0;
        using WordprocessingDocument readDocument = WordprocessingDocument.Open(stream, false);
        DxpPlainTextVisitorConfig config = new()
        {
            TrackedChangeMode = request.RevisionMode == DxpOmmlRevisionMode.Reject
                ? DxpPlainTextTrackedChangeMode.RejectChanges
                : DxpPlainTextTrackedChangeMode.AcceptChanges,
            EmitDocumentProperties = false,
            EmitCustomProperties = false,
        };
        string text = DxpExport.ExportToString(
            readDocument,
            new DxpPlainTextVisitor(config, _logger),
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.None },
            _logger);
        return text.TrimEnd('\r', '\n');
    }
}
