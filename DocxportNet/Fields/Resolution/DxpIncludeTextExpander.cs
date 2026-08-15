using DocxportNet.Visitors.Docx;
using DocxportNet.Fields.Eval;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Resolution;

/// <summary>
/// Expands INCLUDETEXT fields into a rebuilt DOCX while preserving every other
/// field and its cached result unchanged.
/// </summary>
public static class DxpIncludeTextExpander
{
    public static byte[] Expand(
        byte[] document,
        IDxpIncludeTextResolver resolver,
        CancellationToken cancellationToken = default,
        ILogger? logger = null)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (resolver == null) throw new ArgumentNullException(nameof(resolver));
        cancellationToken.ThrowIfCancellationRequested();

        var eval = new DxpFieldEval(logger: logger);
        eval.Context.IncludeTextResolver = resolver;
        eval.Context.CancellationToken = cancellationToken;
        using var visitor = new DxpDocxVisitor(logger, eval);
        byte[] result = DxpExport.ExportToBytes(document, visitor, new DxpExportOptions
        {
            FieldEvalMode = DxpFieldEvalExportMode.Evaluate,
            FieldEvaluationFilter = DxpFieldInstructionClassifier.IsIncludeTextInstruction
        }, logger);
        cancellationToken.ThrowIfCancellationRequested();
        return result;
    }
}
