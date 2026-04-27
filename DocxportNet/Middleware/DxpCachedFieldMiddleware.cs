using DocxportNet.API;
using DocxportNet.Fields;
using DocxportNet.Fields.Eval;
using DocxportNet.Fields.Frames;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Walker;

internal sealed class DxpCachedFieldMiddleware : DxpFieldMiddlewareBase
{
    public DxpCachedFieldMiddleware(
        DxpIVisitor next,
        DxpFieldEval eval,
        bool includeDocumentProperties = true,
        bool includeCustomProperties = true,
        Func<DateTimeOffset>? nowProvider = null,
        ILogger? logger = null,
        DxpCachedFieldMiddlewareOptions? options = null)
        : base(
            next,
            eval,
            includeDocumentProperties,
            includeCustomProperties,
            nowProvider,
            logger,
            "DxpCachedFieldMiddleware")
    {
        _ = options;
    }

    internal override DxpIFieldEvalFrame CreateComplexFieldFrame()
        => new DxpCachedFieldRouterFrame(GetChainedNext(), Context, Logger);

    internal override DxpIFieldEvalFrame CreateSimpleFieldFrame(string? instructionText)
        => new DxpCachedFieldRouterFrame(
            GetChainedNext(),
            Context,
            Logger,
            initialInResult: true,
            initialSeenSeparate: true,
            initialInstructionText: instructionText);
}
