namespace DocxportNet.Fields.Eval;

internal static class DxpFieldInstructionClassifier
{
    internal static bool TryGetImplicitRefName(
        string? instruction,
        DxpFieldEvalContext context,
        out string bookmark)
    {
        bookmark = string.Empty;
        if (string.IsNullOrWhiteSpace(instruction))
            return false;

        var parse = new DxpFieldParser().Parse(instruction!);
        string candidate = parse.Ast.FieldType ?? string.Empty;
        if (!string.IsNullOrWhiteSpace(parse.Ast.ArgumentsText))
            return false;
        if (!IsBookmarkIdentifier(candidate) || !context.TryGetBookmarkNodes(candidate, out _))
            return false;

        bookmark = candidate;
        return true;
    }

    internal static string RewriteImplicitRefInstruction(string instruction, string bookmark)
    {
        string trimmed = instruction.Trim();
        string suffix = trimmed.Length > bookmark.Length
            ? trimmed.Substring(bookmark.Length)
            : string.Empty;
        return $"REF {bookmark}{suffix}";
    }

    internal static bool IsSetInstruction(string? instruction)
    {
        if (string.IsNullOrWhiteSpace(instruction))
            return false;
        var trimmed = instruction.TrimStart();
        if (!trimmed.StartsWith("SET", StringComparison.OrdinalIgnoreCase))
            return false;
        return trimmed.Length == 3 || char.IsWhiteSpace(trimmed[3]);
    }

    internal static bool IsRefInstruction(string? instruction)
    {
        if (string.IsNullOrWhiteSpace(instruction))
            return false;
        var trimmed = instruction.TrimStart();
        if (!trimmed.StartsWith("REF", StringComparison.OrdinalIgnoreCase))
            return false;
        return trimmed.Length == 3 || char.IsWhiteSpace(trimmed[3]);
    }

    internal static bool IsDocVariableInstruction(string? instruction)
    {
        if (string.IsNullOrWhiteSpace(instruction))
            return false;
        var trimmed = instruction.TrimStart();
        if (!trimmed.StartsWith("DOCVARIABLE", StringComparison.OrdinalIgnoreCase))
            return false;
        return trimmed.Length == 11 || char.IsWhiteSpace(trimmed[11]);
    }

    internal static bool IsIfInstruction(string? instruction)
    {
        if (string.IsNullOrWhiteSpace(instruction))
            return false;
        var trimmed = instruction.TrimStart();
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

    internal static bool IsIncludeTextInstruction(string? instruction)
        => StartsWithField(instruction, "INCLUDETEXT");

    internal static bool IsAutoNumberInstruction(string? instruction)
        => StartsWithField(instruction, "AUTONUM")
            || StartsWithField(instruction, "AUTONUMLGL")
            || StartsWithField(instruction, "AUTONUMOUT");

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

    internal static bool IsLayoutDependentInstruction(string? instruction)
        => StartsWithField(instruction, "PAGE")
            || StartsWithField(instruction, "SECTION")
            || StartsWithField(instruction, "SECTIONPAGES")
            || StartsWithField(instruction, "PAGEREF");

    internal static bool IsPaginationDependentInstruction(string? instruction)
        => IsLayoutDependentInstruction(instruction)
            || StartsWithField(instruction, "NUMPAGES");

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
        var trimmed = instruction.TrimStart();
        return trimmed.Length > 0 && trimmed[0] == '=';
    }

    private static bool StartsWithField(string? instruction, string fieldType)
    {
        if (string.IsNullOrWhiteSpace(instruction))
            return false;
        var trimmed = instruction.TrimStart();
        if (!trimmed.StartsWith(fieldType, StringComparison.OrdinalIgnoreCase))
            return false;
        return trimmed.Length == fieldType.Length || char.IsWhiteSpace(trimmed[fieldType.Length]);
    }

    private static bool IsBookmarkIdentifier(string value)
    {
        if (value.Length == 0 || !char.IsLetter(value[0]))
            return false;
        for (int i = 1; i < value.Length; i++)
        {
            if (!char.IsLetterOrDigit(value[i]) && value[i] != '_')
                return false;
        }
        return true;
    }
}
