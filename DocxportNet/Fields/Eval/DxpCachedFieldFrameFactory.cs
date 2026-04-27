using DocxportNet.API;
using DocxportNet.Fields.Frames;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Eval;

internal sealed class DxpCachedFieldFrameFactory
{
    public DxpIFieldEvalFrame Create(
        string? instruction,
        DxpIVisitor next,
        DxpFieldEvalContext context,
        ILogger? logger)
    {
        if (DxpFieldInstructionClassifier.IsSetInstruction(instruction))
            return new DxpSetFieldCachedFrame(context, logger);

        if (DxpFieldInstructionClassifier.IsNextInstruction(instruction))
            return new DxpNextFieldCachedFrame();

        return new DxpSimpleFieldCachedFrame(next, instruction);
    }
}
