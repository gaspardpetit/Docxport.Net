using DocxportNet.API;
using DocxportNet.Fields;
using DocxportNet.Fields.Eval;
using DocxportNet.Fields.Frames;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Walker;

internal sealed class DxpCachedFieldMiddleware : DxpFieldMiddlewareBase
{
    private readonly bool _evaluateDateTimeFields;

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
        _evaluateDateTimeFields = options?.EvaluateDateTimeFields == true;
    }

    internal override DxpIFieldEvalFrame CreateComplexFieldFrame()
        => new DxpCachedFieldRouterFrame(GetChainedNext(), Eval, Context, Logger,
            _evaluateDateTimeFields);

    internal override DxpIFieldEvalFrame CreateSimpleFieldFrame(string? instructionText)
        => new DxpCachedFieldRouterFrame(
            GetChainedNext(),
            Eval,
            Context,
            Logger,
            _evaluateDateTimeFields,
            initialInResult: true,
            initialSeenSeparate: true,
            initialInstructionText: instructionText);
}
