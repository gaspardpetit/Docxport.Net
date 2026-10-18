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

        if (DxpFieldInstructionClassifier.IsHyperlinkInstruction(instruction))
        {
            DxpIVisitor terminal = next;
            while (terminal is DxpMiddleware middleware && middleware.Next != null)
                terminal = middleware.Next;
            if (terminal is IDxpLayoutAwareFieldVisitor)
                return new DxpHyperlinkFieldCachedFrame(next, terminal, instruction!);
        }

        if (DxpFieldInstructionClassifier.IsSetInstruction(instruction))
            return new DxpSetFieldCachedFrame(context, logger);

        if (DxpFieldInstructionClassifier.IsNextInstruction(instruction))
            return new DxpNextFieldCachedFrame();

        return new DxpSimpleFieldCachedFrame(next, instruction);
    }
}
