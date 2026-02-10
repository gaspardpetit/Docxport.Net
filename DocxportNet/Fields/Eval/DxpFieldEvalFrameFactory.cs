using DocxportNet.API;
using DocxportNet.Fields.Frames;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Eval;

internal sealed class DxpFieldEvalFrameFactory
{
    public DxpIFieldEvalFrame Create(
        DxpIFieldEvalFrame frame,
        DxpIVisitor next,
        DxpFieldEval eval,
        DxpFieldEvalContext context,
        ILogger? logger,
        DxpEvalFieldMode mode)
    {
        if (IsRefInstruction(frame.InstructionText))
        {
            return mode == DxpEvalFieldMode.Cache
                ? new DxpRefFieldCachedFrame(next, context, logger)
                : new DxpRefFieldEvalFrame(next, eval, context, logger);
        }
        if (IsDocVariableInstruction(frame.InstructionText))
            return mode == DxpEvalFieldMode.Cache
                ? new DxpDocVariableFieldCachedFrame(next, context, logger)
                : new DxpDocVariableFieldEvalFrame(next, eval, context, logger);
		if (IsIfInstruction(frame.InstructionText))
            return new DxpIFFieldEvalFrame(next, eval, context, logger, mode);
        if (IsSetInstruction(frame.InstructionText))
        {
            return mode == DxpEvalFieldMode.Cache
                ? new DxpSetFieldCachedFrame(eval, context, logger)
                : new DxpSetFieldEvalFrame(eval, context, logger);
        }
        if (logger?.IsEnabled(LogLevel.Debug) == true)
            logger.LogDebug("FieldFrameFactory: falling back to GenericFieldEvalFrame for {FrameType}.", frame.GetType().Name);
        return new DxpEvalGenericFieldFrame(next, eval, context, logger, mode);
    }

    internal static bool IsSetInstruction(string? instruction)
    {
        if (string.IsNullOrWhiteSpace(instruction))
            return false;
        var trimmed = instruction!.TrimStart();
        if (!trimmed.StartsWith("SET", StringComparison.OrdinalIgnoreCase))
            return false;
        return trimmed.Length == 3 || char.IsWhiteSpace(trimmed[3]);
    }

    internal static bool IsRefInstruction(string? instruction)
    {
        if (string.IsNullOrWhiteSpace(instruction))
            return false;
        var trimmed = instruction!.TrimStart();
        if (!trimmed.StartsWith("REF", StringComparison.OrdinalIgnoreCase))
            return false;
        return trimmed.Length == 3 || char.IsWhiteSpace(trimmed[3]);
    }

    internal static bool IsDocVariableInstruction(string? instruction)
    {
        if (string.IsNullOrWhiteSpace(instruction))
            return false;
        var trimmed = instruction!.TrimStart();
        if (!trimmed.StartsWith("DOCVARIABLE", StringComparison.OrdinalIgnoreCase))
            return false;
        return trimmed.Length == 11 || char.IsWhiteSpace(trimmed[11]);
    }

    internal static bool IsIfInstruction(string? instruction)
    {
        if (string.IsNullOrWhiteSpace(instruction))
            return false;
        var trimmed = instruction!.TrimStart();
        if (!trimmed.StartsWith("IF", StringComparison.OrdinalIgnoreCase))
            return false;
        return trimmed.Length == 2 || char.IsWhiteSpace(trimmed[2]);
    }
}
