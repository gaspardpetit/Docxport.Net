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
                : new DxpRefFieldEvalFrame(next, eval, logger, instruction);
        }
        if (IsDocVariableInstruction(instruction))
            return mode == DxpEvalFieldMode.Cache
                ? new DxpDocVariableFieldCachedFrame(next)
                : new DxpDocVariableFieldEvalFrame(next, eval, logger, instruction);
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
        if (IsAskInstruction(instruction))
        {
            return mode == DxpEvalFieldMode.Cache
                ? new DxpAskFieldCachedFrame(next)
                : new DxpAskFieldEvalFrame(next, eval, logger, instruction);
        }
        if (IsSkipIfInstruction(instruction))
        {
            return mode == DxpEvalFieldMode.Cache
                ? new DxpSkipIfFieldCachedFrame(next)
                : new DxpSkipIfFieldEvalFrame(next, eval, logger, instruction);
        }
        if (IsDocPropertyInstruction(instruction) ||
            IsMergeFieldInstruction(instruction) ||
            IsSeqInstruction(instruction) ||
            IsDateTimeInstruction(instruction) ||
            IsCompareInstruction(instruction))
        {
            if (IsDocPropertyInstruction(instruction))
            {
                return mode == DxpEvalFieldMode.Cache
                    ? new DxpDocPropertyFieldCachedFrame(next)
                    : new DxpDocPropertyFieldEvalFrame(next, eval, logger, instruction);
            }
            if (IsMergeFieldInstruction(instruction))
            {
                return mode == DxpEvalFieldMode.Cache
                    ? new DxpMergeFieldCachedFrame(next)
                    : new DxpMergeFieldEvalFrame(next, eval, logger, instruction);
            }
            if (IsSeqInstruction(instruction))
            {
                return mode == DxpEvalFieldMode.Cache
                    ? new DxpSeqFieldCachedFrame(next)
                    : new DxpSeqFieldEvalFrame(next, eval, logger, instruction);
            }
            if (IsDateTimeInstruction(instruction))
            {
                return mode == DxpEvalFieldMode.Cache
                    ? new DxpDateTimeFieldCachedFrame(next)
                    : new DxpDateTimeFieldEvalFrame(next, eval, logger, instruction);
            }
            if (IsCompareInstruction(instruction))
            {
                return mode == DxpEvalFieldMode.Cache
                    ? new DxpCompareFieldCachedFrame(next)
                    : new DxpCompareFieldEvalFrame(next, eval, logger, instruction);
            }
        }
        if (IsFormulaInstruction(instruction))
        {
            return mode == DxpEvalFieldMode.Cache
                ? new DxpFormulaFieldCachedFrame(next)
                : new DxpFormulaFieldEvalFrame(next, eval, logger, instruction);
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

    internal static bool IsDocPropertyInstruction(string? instruction)
        => StartsWithField(instruction, "DOCPROPERTY");

    internal static bool IsMergeFieldInstruction(string? instruction)
        => StartsWithField(instruction, "MERGEFIELD");

    internal static bool IsSeqInstruction(string? instruction)
        => StartsWithField(instruction, "SEQ");

    internal static bool IsCompareInstruction(string? instruction)
        => StartsWithField(instruction, "COMPARE");

    internal static bool IsAskInstruction(string? instruction)
        => StartsWithField(instruction, "ASK");

    internal static bool IsSkipIfInstruction(string? instruction)
    {
        if (StartsWithField(instruction, "SKIPIF"))
            return true;
        return StartsWithField(instruction, "NEXTIF");
    }

    internal static bool IsDateTimeInstruction(string? instruction)
    {
        if (StartsWithField(instruction, "DATE"))
            return true;
        if (StartsWithField(instruction, "TIME"))
            return true;
        if (StartsWithField(instruction, "CREATEDATE"))
            return true;
        if (StartsWithField(instruction, "SAVEDATE"))
            return true;
        return StartsWithField(instruction, "PRINTDATE");
    }

    internal static bool IsFormulaInstruction(string? instruction)
    {
        if (string.IsNullOrWhiteSpace(instruction))
            return false;
        var trimmed = instruction!.TrimStart();
        return trimmed.Length > 0 && trimmed[0] == '=';
    }

    private static bool StartsWithField(string? instruction, string fieldType)
    {
        if (string.IsNullOrWhiteSpace(instruction))
            return false;
        var trimmed = instruction!.TrimStart();
        if (!trimmed.StartsWith(fieldType, StringComparison.OrdinalIgnoreCase))
            return false;
        return trimmed.Length == fieldType.Length || char.IsWhiteSpace(trimmed[fieldType.Length]);
    }
}
