using DocxportNet.API;
using DocxportNet.Fields.Frames;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Eval;

internal sealed class DxpFieldEvalFrameFactory
{
    public DxpIFieldEvalFrame? Create(
        string? instruction,
        DxpIVisitor next,
        DxpFieldEval eval,
        DxpFieldEvalContext context,
        ILogger? logger,
        DxpEvalFieldMode mode)
    {
        if (mode == DxpEvalFieldMode.Cache)
        {
            if (IsSetInstruction(instruction))
                return new DxpSetFieldCachedFrame(context, logger);

            if (IsNextInstruction(instruction))
                return new DxpNextFieldCachedFrame();

            return new DxpSimpleFieldCachedFrame(next, instruction);
        }

        if (IsRefInstruction(instruction))
            return new DxpRefFieldEvalFrame(next, eval, logger, instruction);

        if (IsDocVariableInstruction(instruction))
            return new DxpDocVariableFieldEvalFrame(next, eval, logger, instruction);

        if (IsIfInstruction(instruction))
            return new DxpIFFieldEvalFrame(next, eval, logger);

        if (IsSetInstruction(instruction))
            return new DxpSetFieldEvalFrame(eval, context, logger, instruction);

        if (IsAskInstruction(instruction))
            return new DxpAskFieldEvalFrame(next, eval, logger, instruction);

        if (IsFillInInstruction(instruction))
            return new DxpValueFieldEvalFrame(next, eval, logger, instruction);

        if (IsNextInstruction(instruction))
            return new DxpValueFieldEvalFrame(next, eval, logger, instruction);

        if (IsSkipIfInstruction(instruction))
            return new DxpSkipIfFieldEvalFrame(next, eval, logger, instruction);

        if (IsDocPropertyInstruction(instruction) ||
            IsMergeFieldInstruction(instruction) ||
            IsMergeRecInstruction(instruction) ||
            IsMergeSeqInstruction(instruction) ||
            IsGreetingLineInstruction(instruction) ||
            IsAddressBlockInstruction(instruction) ||
            IsDatabaseInstruction(instruction) ||
            IsSeqInstruction(instruction) ||
            IsDateTimeInstruction(instruction) ||
            IsCompareInstruction(instruction) ||
            IsDocumentMetricInstruction(instruction))
        {
            if (IsDocPropertyInstruction(instruction))
                return new DxpDocPropertyFieldEvalFrame(next, eval, logger, instruction);

            if (IsMergeFieldInstruction(instruction))
                return new DxpMergeFieldEvalFrame(next, eval, logger, instruction);

            if (IsMergeRecInstruction(instruction) ||
                IsMergeSeqInstruction(instruction) ||
                IsGreetingLineInstruction(instruction) ||
                IsAddressBlockInstruction(instruction) ||
                IsDatabaseInstruction(instruction) ||
                IsDocumentMetricInstruction(instruction))
            {
                return new DxpValueFieldEvalFrame(next, eval, logger, instruction);
            }

            if (IsSeqInstruction(instruction))
                return new DxpSeqFieldEvalFrame(next, eval, logger, instruction);

            if (IsDateTimeInstruction(instruction))
                return new DxpDateTimeFieldEvalFrame(next, eval, logger, instruction);

            if (IsCompareInstruction(instruction))
                return new DxpCompareFieldEvalFrame(next, eval, logger, instruction);
        }
        if (IsFormulaInstruction(instruction))
            return new DxpFormulaFieldEvalFrame(next, eval, logger, instruction);

        if (logger?.IsEnabled(LogLevel.Debug) == true)
            logger.LogDebug("FieldFrameFactory: no evaluate frame for instruction '{Instruction}'.", instruction ?? string.Empty);
        return null;
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

    internal static bool IsMergeRecInstruction(string? instruction)
        => StartsWithField(instruction, "MERGEREC");

    internal static bool IsMergeSeqInstruction(string? instruction)
        => StartsWithField(instruction, "MERGESEQ");

    internal static bool IsGreetingLineInstruction(string? instruction)
        => StartsWithField(instruction, "GREETINGLINE");

    internal static bool IsAddressBlockInstruction(string? instruction)
        => StartsWithField(instruction, "ADDRESSBLOCK");

    internal static bool IsDatabaseInstruction(string? instruction)
        => StartsWithField(instruction, "DATABASE");

    internal static bool IsNextInstruction(string? instruction)
        => StartsWithField(instruction, "NEXT");

    internal static bool IsSeqInstruction(string? instruction)
        => StartsWithField(instruction, "SEQ");

    internal static bool IsCompareInstruction(string? instruction)
        => StartsWithField(instruction, "COMPARE");

    internal static bool IsAskInstruction(string? instruction)
        => StartsWithField(instruction, "ASK");

    internal static bool IsFillInInstruction(string? instruction)
        => StartsWithField(instruction, "FILLIN");

    internal static bool IsDocumentMetricInstruction(string? instruction)
    {
        if (StartsWithField(instruction, "NUMPAGES"))
            return true;
        if (StartsWithField(instruction, "NUMWORDS"))
            return true;
        return StartsWithField(instruction, "NUMCHARS");
    }

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
