using DocxportNet.API;
using DocxportNet.Fields;
using DocxportNet.Fields.Eval;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Walker;

public static class DxpFieldEvalMiddleware
{
    public static DxpIVisitor CreateCachedFieldMiddleware(
        DxpIVisitor next,
        DxpFieldEval eval,
        bool includeDocumentProperties = true,
        bool includeCustomProperties = true,
        Func<DateTimeOffset>? nowProvider = null,
        ILogger? logger = null,
        DxpCachedFieldMiddlewareOptions? options = null)
        => new DxpCachedFieldMiddleware(
            next,
            eval,
            includeDocumentProperties,
            includeCustomProperties,
            nowProvider,
            logger,
            options);

    public static DxpIVisitor CreateEvaluatedFieldMiddleware(
        DxpIVisitor next,
        DxpFieldEval eval,
        bool includeDocumentProperties = true,
        bool includeCustomProperties = true,
        Func<DateTimeOffset>? nowProvider = null,
        ILogger? logger = null,
        DxpEvaluateFieldMiddlewareOptions? options = null)
        => new DxpEvaluateFieldMiddleware(
            next,
            eval,
            includeDocumentProperties,
            includeCustomProperties,
            nowProvider,
            logger,
            options);
}
