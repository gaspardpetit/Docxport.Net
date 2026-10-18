using DocxportNet.API;
using DocxportNet.Fields.Frames;
using DocxportNet.Middleware;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Eval;

internal sealed class DxpCachedFieldFrameFactory
{
    public DxpIFieldEvalFrame Create(
        string? instruction,
        DxpIVisitor next,
        DxpFieldEval eval,
        DxpFieldEvalContext context,
        ILogger? logger,
        bool evaluateDateTimeFields)
    {
        if (evaluateDateTimeFields && DxpFieldInstructionClassifier.IsDateTimeInstruction(instruction))
            return new DxpDateTimeFieldEvalFrame(next, eval, logger, instruction);

        if (next is DxpMiddleware { Next: IDxpCachedHyperlinkFieldVisitor } &&
            DxpFieldInstructionClassifier.IsHyperlinkInstruction(instruction))
            return new DxpHyperlinkFieldCachedFrame(next, instruction!);

        if (DxpFieldInstructionClassifier.IsSetInstruction(instruction))
            return new DxpSetFieldCachedFrame(context, logger);

        if (DxpFieldInstructionClassifier.IsNextInstruction(instruction))
            return new DxpNextFieldCachedFrame();

        return new DxpSimpleFieldCachedFrame(next, instruction);
    }
}
