using DocxportNet.API;
using DocxportNet.Walker;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields;

internal interface IDxpStructuredFieldSpliceCollector
{
    bool Record(DxpFieldNodeBuffer buffer);
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
    internal void EmitStandalone(DxpIVisitor visitor, DxpIDocumentContext context)
    {
        if (!Eval.Context.TryEnterIncludeText(Identity, out var recursionError))
        {
            Logger?.LogWarning("{Error} Using cached INCLUDETEXT result.", recursionError);
            CachedResult?.Replay(visitor, context);
            return;
        }

        try
        {
            bool emitted = DxpEmbeddedIncludeTextRunner.TryRun(
                Path,
                Content,
                Bookmark,
                visitor,
                Eval,
                Logger,
                static (walker, document, pipeline, blocks) =>
                    walker.AcceptEmbeddedBody(document, pipeline, blocks));
            if (!emitted)
                CachedResult?.Replay(visitor, context);
        }
        finally
        {
            Eval.Context.ExitIncludeText(Identity);
        }
    }

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
            bool emitted = DxpEmbeddedIncludeTextRunner.TryRun(
                Path,
                Content,
                Bookmark,
                visitor,
                Eval,
                Logger,
                (walker, document, pipeline, blocks) =>
                {
                    DxpIDocumentContext childContext = walker.AcceptEmbeddedBodySpliced(
                        document, pipeline, parentContext, parentParagraph, before, after, blocks);
                    while (Eval.Context.TryTakeDeferredStructuredFieldResult(childContext, out var deferred) && deferred != null)
                        EmitDeferred(deferred, visitor, parentContext, parentParagraph);
                    return childContext;
                });
            if (!emitted)
            {
                EmitCache(visitor, parentContext, parentParagraph, before, after);
                return;
            }
        }
        finally
        {
            Eval.Context.ExitIncludeText(Identity);
        }
    }

    private static void EmitDeferred(
        DxpFieldNodeBuffer buffer,
        DxpIVisitor visitor,
        DxpIDocumentContext context,
        DocumentFormat.OpenXml.Wordprocessing.Paragraph parentParagraph)
    {
        var parts = buffer.SplitIncludeTextExpansions();
        if (parts.Count == 3 && parts[1].Expansion != null)
        {
            parts[1].Expansion!.Emit(visitor, context, parentParagraph,
                parts[0].Inline, parts[2].Inline);
            return;
        }

        foreach (var part in parts)
        {
            if (part.Inline != null)
                part.Inline.Replay(visitor, context);
            else
                part.Expansion?.Emit(visitor, context, parentParagraph, null, null);
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
