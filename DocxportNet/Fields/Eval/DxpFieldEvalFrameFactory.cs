using DocxportNet.API;
using DocxportNet.Fields.Frames;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Eval;

internal sealed class DxpFieldEvalFrameFactory
{
    public DxpIFieldEvalFrame Create(
        string? instruction,
        DxpIVisitor next,
        DxpFieldEval eval,
        DxpFieldEvalContext context,
        ILogger? logger,
        DxpEvalFieldMode mode)
    {
        if (IsRefInstruction(instruction))
        {
            return mode == DxpEvalFieldMode.Cache
                ? new DxpRefFieldCachedFrame(next)
                : new DxpRefFieldEvalFrame(next, eval, logger, codeRun: null, instructionText: instruction);
        }
        if (IsDocVariableInstruction(instruction))
            return mode == DxpEvalFieldMode.Cache
                ? new DxpDocVariableFieldCachedFrame(next)
                : new DxpDocVariableFieldEvalFrame(next, eval, logger, codeRun: null, instructionText: instruction);
        if (IsIfInstruction(instruction))
            return mode == DxpEvalFieldMode.Cache
                ? new DxpIFFieldCachedFrame(next)
                : new DxpIFFieldEvalFrame(next, eval, logger);
		if (IsSetInstruction(instruction))
        {
            return mode == DxpEvalFieldMode.Cache
                ? new DxpSetFieldCachedFrame(context, logger)
                : new DxpSetFieldEvalFrame(eval, context, logger, instruction);
        }
        if (logger?.IsEnabled(LogLevel.Debug) == true)
            logger.LogDebug("FieldFrameFactory: falling back to GenericFieldEvalFrame for instruction '{Instruction}'.", instruction ?? string.Empty);
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
