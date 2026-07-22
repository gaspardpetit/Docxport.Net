using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Walker;
using System.Globalization;

namespace DocxportNet.Fields.Eval;

internal sealed record DxpAutoNumberResult(bool Handled, string Text, bool Suppressed, int Value);

internal sealed class DxpAutoNumberResolver
{
    private static readonly NumberFormatValues[] OutlineFormats = {
        NumberFormatValues.UpperRoman,
        NumberFormatValues.UpperLetter,
        NumberFormatValues.Decimal,
        NumberFormatValues.LowerLetter,
        NumberFormatValues.LowerRoman,
        NumberFormatValues.Decimal,
        NumberFormatValues.LowerLetter,
        NumberFormatValues.LowerRoman,
        NumberFormatValues.Decimal
    };

    private readonly DxpFieldEvalContext _context;
    private readonly DxpFieldParser _parser = new();
    private readonly DxpFieldFormatter _formatter = new();

    internal DxpAutoNumberResolver(DxpFieldEvalContext context) => _context = context;

    internal DxpAutoNumberResult Resolve(string instruction, DxpIDocumentContext? documentContext)
    {
        var parse = _parser.Parse(instruction);
        var fieldType = parse.Ast.FieldType?.ToUpperInvariant();
        if (!parse.Success || fieldType is not ("AUTONUM" or "AUTONUMLGL" or "AUTONUMOUT"))
            return new DxpAutoNumberResult(false, string.Empty, false, 0);

        string storyKey = _context.CurrentStoryKeyProvider?.Invoke()
            ?? documentContext?.CurrentPart?.Uri.ToString()
            ?? "main";
        int headingLevel = _context.CurrentBuiltInHeadingLevelProvider?.Invoke() ?? 0;
        var story = _context.AutoNumbers.GetStory(storyKey);
        var family = fieldType switch {
            "AUTONUMLGL" => story.Legal,
            "AUTONUMOUT" => story.Outline,
            _ => story.AutoNum
        };

        int value;
        IReadOnlyList<int> components;
        if (headingLevel > 0)
        {
            value = family.AdvanceHeading(headingLevel);
            components = fieldType == "AUTONUM"
                ? new[] { value }
                : family.CurrentPath(headingLevel);
        }
        else
        {
            value = family.AdvanceBody();
            var activeLevel = Array.FindLastIndex(family.HeadingCounters, n => n > 0) + 1;
            components = fieldType == "AUTONUM" || activeLevel == 0
                ? new[] { value }
                : family.CurrentPath(activeLevel).Concat(new[] { value }).ToArray();
        }

        string separator = ResolveSeparator(instruction);
        string text = fieldType switch {
            "AUTONUM" => _formatter.Format(new DxpFieldValue(value), parse.Ast.FormatSpecs, _context) + separator,
            "AUTONUMLGL" => string.Join(separator, components.Select(n => n.ToString(CultureInfo.InvariantCulture)))
                + (HasSwitch(instruction, 'e') ? string.Empty : separator),
            _ => FormatOutline(components, separator)
        };

        _context.SetNumberedItem(fieldType, text.TrimEnd(separator.ToCharArray()));
        bool suppressed = IsNestedInIf(documentContext);
        return new DxpAutoNumberResult(true, suppressed ? string.Empty : text, suppressed, value);
    }

    private string FormatOutline(IReadOnlyList<int> components, string separator)
    {
        var labels = components.Select((number, index) =>
            DxpLists.FormatNumber(number, OutlineFormats[Math.Min(index, OutlineFormats.Length - 1)], _context.Culture));
        return string.Join(separator, labels) + separator;
    }

    internal static string ResolveSeparator(string instruction)
    {
        for (int i = 0; i < instruction.Length - 1; i++)
        {
            if (instruction[i] != '\\' || char.ToLowerInvariant(instruction[i + 1]) != 's')
                continue;
            int index = i + 2;
            while (index < instruction.Length && char.IsWhiteSpace(instruction[index]))
                index++;
            return index < instruction.Length && instruction[index] != '\\'
                ? instruction[index].ToString()
                : string.Empty;
        }
        return ".";
    }

    private static bool HasSwitch(string instruction, char name)
    {
        for (int i = 0; i < instruction.Length - 1; i++)
            if (instruction[i] == '\\' && char.ToLowerInvariant(instruction[i + 1]) == char.ToLowerInvariant(name))
                return true;
        return false;
    }

    private static bool IsNestedInIf(DxpIDocumentContext? documentContext)
    {
        if (documentContext == null)
            return false;
        bool skippedCurrent = false;
        foreach (var frame in documentContext.CurrentFields.FieldStack)
        {
            if (!skippedCurrent)
            {
                skippedCurrent = true;
                continue;
            }
            var instruction = frame.InstructionText?.TrimStart();
            if (instruction != null && instruction.StartsWith("IF", StringComparison.OrdinalIgnoreCase)
                && (instruction.Length == 2 || char.IsWhiteSpace(instruction[2])))
                return true;
        }
        return false;
    }
}
