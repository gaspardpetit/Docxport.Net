using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocxportNet.API;
using DocxportNet.Fields.Eval;
using DocxportNet.Fields.Resolution;
using DocxportNet.Middleware;
using DocxportNet.Walker;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields;

internal static class DxpEmbeddedIncludeTextRunner
{
    internal static bool TryRun(
        string path,
        byte[] content,
        string? bookmark,
        DxpIVisitor terminal,
        DxpFieldEval eval,
        ILogger? logger,
        Func<DxpWalker, WordprocessingDocument, DxpIVisitor, IReadOnlyList<OpenXmlElement>?, DxpIDocumentContext> walk)
    {
        MemoryStream? stream = null;
        WordprocessingDocument? document = null;
        try
        {
            stream = new MemoryStream(content, writable: false);
            document = WordprocessingDocument.Open(stream, false);
            if (document.MainDocumentPart?.Document?.Body == null)
                throw new InvalidOperationException("DOCX has no main document body.");
        }
        catch (Exception ex) when (ex is OpenXmlPackageException
            or FileFormatException
            or InvalidOperationException)
        {
            document?.Dispose();
            stream?.Dispose();
            logger?.LogWarning(ex, "INCLUDETEXT source '{Path}' is not a valid DOCX; using cached result.", path);
            return false;
        }

        using (stream)
        using (document)
        {
            IReadOnlyList<OpenXmlElement>? blocks = null;
            if (!string.IsNullOrWhiteSpace(bookmark))
            {
                var body = document.MainDocumentPart!.Document.Body!;
                if (!DxpBookmarkRangeProjector.TryProject(body, bookmark!, out var projected, out var error))
                {
                    logger?.LogWarning("{Error} Using cached INCLUDETEXT result.", error);
                    return false;
                }
                blocks = projected;
            }

            var pipeline = CreateChildPipeline(terminal, eval, logger);
            var walker = new DxpWalker(logger);
            DxpIDocumentContext documentContext = walk(walker, document, pipeline, blocks);

            DxpVisitorMiddleware.CompleteEmbeddedWalk(pipeline, documentContext);
            return true;
        }
    }

    private static DxpIVisitor CreateChildPipeline(
        DxpIVisitor terminal,
        DxpFieldEval eval,
        ILogger? logger)
        => DxpVisitorMiddleware.Chain(
            terminal,
            next => DxpFieldEvalMiddleware.CreateEvaluatedFieldMiddleware(
                next,
                eval,
                logger: logger,
                options: new DxpEvaluateFieldMiddlewareOptions
                {
                    PreserveLayoutDependentFields = eval.Context.PreserveLayoutDependentFields,
                    EmitStructuredDatabaseResults = eval.Context.EmitStructuredDatabaseResults
                }),
            next => new DxpContextMiddleware(next, logger));
}
