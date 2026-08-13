using DocxportNet.API;
using DocxportNet.Fields;
using DocxportNet.Fields.Eval;
using DocxportNet.Fields.Frames;
using DocxportNet.Fields.Resolution;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Walker;

internal sealed class DxpEvaluateFieldMiddleware : DxpFieldMiddlewareBase
{
    private readonly DxpEvaluateFieldMiddlewareOptions _options;

    public DxpEvaluateFieldMiddleware(
        DxpIVisitor next,
        DxpFieldEval eval,
        bool includeDocumentProperties = true,
        bool includeCustomProperties = true,
        Func<DateTimeOffset>? nowProvider = null,
        ILogger? logger = null,
        DxpEvaluateFieldMiddlewareOptions? options = null)
        : base(
            next,
            eval,
            includeDocumentProperties,
            includeCustomProperties,
            nowProvider,
            logger,
            "DxpEvaluateFieldMiddleware")
    {
        _options = options ?? new DxpEvaluateFieldMiddlewareOptions();
    }

    internal override DxpIFieldEvalFrame CreateComplexFieldFrame()
        => new DxpEvaluateFieldRouterFrame(GetChainedNext(), Eval, Context, Logger);

    internal override DxpIFieldEvalFrame CreateSimpleFieldFrame(string? instructionText)
        => new DxpEvaluateFieldRouterFrame(
            GetChainedNext(),
            Eval,
            Context,
            Logger,
            initialInResult: true,
            initialSeenSeparate: true,
            initialInstructionText: instructionText);

    protected override void InitializeModeSpecificContext(DxpIDocumentContext documentContext)
    {
        _ = documentContext;
        Context.RefResolver ??= _options.RefResolver ?? new DxpRefIndexResolver();
        Context.EmitStructuredDatabaseResults = _options.EmitStructuredDatabaseResults;
        Context.UseSemanticFieldResults = _options.UseSemanticFieldResults;
        if (_options.PreserveLayoutDependentFields)
            Context.PreserveLayoutDependentFields = true;
    }
}
