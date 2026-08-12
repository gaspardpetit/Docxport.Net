using DocxportNet.API;
using DocxportNet.Walker;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields;

internal interface IDxpIncludeTextSpliceCollector
{
    bool Record(DxpIncludeTextExpansion expansion);
    void Complete();
}

internal sealed record DxpFieldNodeBufferSplicePart(
    DxpFieldNodeBuffer? Inline,
    DxpIncludeTextExpansion? Expansion);

internal sealed record DxpIncludeTextExpansion(
    string Path,
    string Identity,
    byte[] Content,
    string? Bookmark,
    DxpFieldNodeBuffer? CachedResult,
    DxpFieldEval Eval,
    Microsoft.Extensions.Logging.ILogger? Logger)
{
    internal void Emit(
        DxpIVisitor visitor,
        DxpIDocumentContext parentContext,
        DocumentFormat.OpenXml.Wordprocessing.Paragraph parentParagraph,
        DxpFieldNodeBuffer? before,
        DxpFieldNodeBuffer? after)
    {
        if (!Eval.Context.TryEnterIncludeText(Identity, out var recursionError))
        {
            Logger?.LogWarning("{Error} Using cached INCLUDETEXT result.", recursionError);
            EmitCache(visitor, parentContext, parentParagraph, before, after);
            return;
        }

        try
        {
            MemoryStream? stream = null;
            DocumentFormat.OpenXml.Packaging.WordprocessingDocument? document = null;
            try
            {
                stream = new MemoryStream(Content, writable: false);
                document = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, false);
                if (document.MainDocumentPart?.Document?.Body == null)
                    throw new InvalidOperationException("DOCX has no main document body.");
            }
            catch (Exception ex) when (ex is DocumentFormat.OpenXml.Packaging.OpenXmlPackageException
                or FileFormatException
                or InvalidOperationException)
            {
                document?.Dispose();
                stream?.Dispose();
                Logger?.LogWarning(ex, "INCLUDETEXT source '{Path}' is not a valid DOCX; using cached result.", Path);
                EmitCache(visitor, parentContext, parentParagraph, before, after);
                return;
            }

            using (stream)
            using (document)
            {
                IReadOnlyList<DocumentFormat.OpenXml.OpenXmlElement>? blocks = null;
                if (!string.IsNullOrWhiteSpace(Bookmark))
                {
                    var body = document.MainDocumentPart!.Document.Body!;
                    if (!Resolution.DxpBookmarkRangeProjector.TryProject(body, Bookmark!, out var projected, out var error))
                    {
                        Logger?.LogWarning("{Error} Using cached INCLUDETEXT result.", error);
                        EmitCache(visitor, parentContext, parentParagraph, before, after);
                        return;
                    }
                    blocks = projected;
                }

                var pipeline = DocxportNet.Middleware.DxpVisitorMiddleware.Chain(
                    visitor,
                    next => DxpFieldEvalMiddleware.CreateEvaluatedFieldMiddleware(next, Eval, logger: Logger),
                    next => new DocxportNet.Middleware.DxpContextMiddleware(next, Logger));
                new DxpWalker(Logger).AcceptEmbeddedBodySpliced(document, pipeline, parentContext, parentParagraph, before, after, blocks);
            }
        }
        finally
        {
            Eval.Context.ExitIncludeText(Identity);
        }
    }

    private void EmitCache(DxpIVisitor visitor, DxpIDocumentContext context,
        DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph,
        DxpFieldNodeBuffer? before, DxpFieldNodeBuffer? after)
    {
        var merged = new DxpFieldNodeBuffer();
        merged.Append(before);
        merged.Append(CachedResult);
        merged.Append(after);
        context.Walker.ReplayBufferedParentParagraph(paragraph, context, visitor, merged);
    }
}
