using DocxportNet.Fields.Formatting;

namespace DocxportNet.Fields;

public sealed class DxpFieldAst
{
    public string RawText { get; }
    public string? FieldType { get; }
    public string? ArgumentsText { get; }
    public IReadOnlyList<IDxpFieldFormatSpec> FormatSpecs { get; }

    public DxpFieldAst(string rawText, string? fieldType, string? argumentsText, IReadOnlyList<IDxpFieldFormatSpec> formatSpecs)
    {
        RawText = rawText ?? string.Empty;
        FieldType = fieldType;
        ArgumentsText = argumentsText;
        FormatSpecs = formatSpecs;
    }
}

public sealed record DxpFieldParseResult(DxpFieldAst Ast, IReadOnlyList<string> Errors)
{
    public bool Success => Errors.Count == 0;
}

public sealed class DxpFieldParser
{
    public DxpFieldParseResult Parse(string instructionText)
    {
        var raw = instructionText ?? string.Empty;
        ParseSegments(raw, out var main, out var switchSegments);
        var fieldType = ParseFieldType(main, out var argsText);
        var specs = ParseFormatSpecs(switchSegments);
        var ast = new DxpFieldAst(raw, fieldType, argsText, specs);
        return new DxpFieldParseResult(ast, []);
    }

    private static void ParseSegments(string text, out string main, out List<string> switches)
    {
        switches = new List<string>();
        if (text.Length == 0)
        {
            main = string.Empty;
            return;
        }

        bool inQuote = false;
        int braceDepth = 0;
        var switchStarts = new List<int>();
        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            if (inQuote && ch == '\\' && i + 1 < text.Length && text[i + 1] == '"')
            {
                i++;
                continue;
            }
            if (ch == '"')
            {
                inQuote = !inQuote;
                continue;
            }

            if (!inQuote)
            {
                if (ch == '{')
                {
                    braceDepth++;
                    continue;
                }
                if (ch == '}' && braceDepth > 0)
                {
                    braceDepth--;
                    continue;
                }
            }

            if (!inQuote && braceDepth == 0 && ch == '\\')
                switchStarts.Add(i);
        }

        if (switchStarts.Count == 0)
        {
            main = text;
            return;
        }

        main = text.Substring(0, switchStarts[0]);
        for (int i = 0; i < switchStarts.Count; i++)
        {
            int start = switchStarts[i];
            int end = i + 1 < switchStarts.Count ? switchStarts[i + 1] : text.Length;
            switches.Add(text.Substring(start, end - start));
        }
    }

    private static string? ParseFieldType(string main, out string? argsText)
    {
        argsText = null;
        if (string.IsNullOrWhiteSpace(main))
            return null;

        int i = 0;
        while (i < main.Length && char.IsWhiteSpace(main[i]))
            i++;
        int start = i;
        while (i < main.Length && !char.IsWhiteSpace(main[i]))
            i++;
        if (start == i)
            return null;
        string fieldType = main.Substring(start, i - start);
        argsText = i < main.Length ? main.Substring(i).Trim() : null;
        return fieldType;
    }

    private List<IDxpFieldFormatSpec> ParseFormatSpecs(List<string> switchSegments)
    {
        var specs = new List<IDxpFieldFormatSpec>(switchSegments.Count);
        foreach (var seg in switchSegments)
        {
            if (string.IsNullOrWhiteSpace(seg))
                continue;

            string raw = seg.Trim();
            int i = 0;
            while (i < raw.Length && raw[i] == '\\')
                i++;
            while (i < raw.Length && char.IsWhiteSpace(raw[i]))
                i++;
            if (i >= raw.Length)
                continue;

            char kindChar = raw[i];
            string? arg = i + 1 < raw.Length ? raw.Substring(i + 1).Trim() : null;
            string? unquoted = Unquote(arg);
            if (string.IsNullOrWhiteSpace(unquoted))
                continue;
            string argValue = unquoted!;
            switch (kindChar)
            {
                case '*':
                    specs.Add(new DxpTextTransformFormatSpec(ParseTextTransform(argValue), raw, argValue));
                    break;
                case '#':
                    specs.Add(ParseNumericFormat(argValue, raw));
                    break;
                case '@':
                    specs.Add(ParseDateTimeFormat(argValue, raw));
                    break;
            }
        }
        return specs;
    }

    private string? Unquote(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return value;
        string? trimmed = value?.Trim();
        if (trimmed?.Length >= 2 && trimmed[0] == '"' && trimmed[trimmed.Length - 1] == '"')
            return trimmed.Substring(1, trimmed.Length - 2);
        return trimmed;
    }

    private static DxpTextTransformKind ParseTextTransform(string arg)
    {
        return arg.Trim().ToLowerInvariant() switch {
            "caps" => DxpTextTransformKind.Caps,
            "firstcap" => DxpTextTransformKind.FirstCap,
            "upper" => DxpTextTransformKind.Upper,
            "lower" => DxpTextTransformKind.Lower,
            "alphabetic" => DxpTextTransformKind.Alphabetic,
            "arabic" => DxpTextTransformKind.Arabic,
            "arabicdash" => DxpTextTransformKind.ArabicDash,
            "cardtext" => DxpTextTransformKind.CardText,
            "dollartext" => DxpTextTransformKind.DollarText,
            "hex" => DxpTextTransformKind.Hex,
            "ordtext" => DxpTextTransformKind.OrdText,
            "ordinal" => DxpTextTransformKind.Ordinal,
            "roman" => DxpTextTransformKind.Roman,
            "charformat" => DxpTextTransformKind.Charformat,
            "mergeformat" => DxpTextTransformKind.MergeFormat,
            _ => DxpTextTransformKind.Caps
        };
    }

    private static DxpDateTimeFormatSpec ParseDateTimeFormat(string format, string raw)
    {
        var tokens = new List<DxpDateTimeToken>();
        int i = 0;
        while (i < format.Length)
        {
            char ch = format[i];
            if (ch == '`')
            {
                int start = ++i;
                while (i < format.Length && format[i] != '`')
                    i++;
                string label = format.Substring(start, i - start);
                tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.NumberedItem, label));
                if (i < format.Length && format[i] == '`')
                    i++;
                continue;
            }
            if (ch == '\'')
            {
                int start = ++i;
                while (i < format.Length && format[i] != '\'')
                    i++;
                string literal = format.Substring(start, i - start);
                tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.Literal, literal));
                if (i < format.Length && format[i] == '\'')
                    i++;
                continue;
            }

            if (i + 4 < format.Length)
            {
                string ampm = format.Substring(i, 5);
                if (ampm == "AM/PM" || ampm == "am/pm")
                {
                    tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.AmPm, ampm));
                    i += 5;
                    continue;
                }
            }

            int run = CountRun(format, i);
            string seq = format.Substring(i, run);
            switch (seq)
            {
                case "M":
                    tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.MonthNumeric, seq));
                    break;
                case "MM":
                    tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.MonthNumeric2, seq));
                    break;
                case "MMM":
                    tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.MonthShortName, seq));
                    break;
                case "MMMM":
                    tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.MonthFullName, seq));
                    break;
                case "d":
                    tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.DayNumeric, seq));
                    break;
                case "dd":
                    tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.DayNumeric2, seq));
                    break;
                case "ddd":
                    tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.WeekdayShortName, seq));
                    break;
                case "dddd":
                    tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.WeekdayFullName, seq));
                    break;
                case "yy":
                    tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.Year2, seq));
                    break;
                case "yyyy":
                    tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.Year4, seq));
                    break;
                case "h":
                    tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.Hour12, seq));
                    break;
                case "hh":
                    tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.Hour12_2, seq));
                    break;
                case "H":
                    tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.Hour24, seq));
                    break;
                case "HH":
                    tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.Hour24_2, seq));
                    break;
                case "m":
                    tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.Minute, seq));
                    break;
                case "mm":
                    tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.Minute2, seq));
                    break;
                case "s":
                    tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.Second, seq));
                    break;
                case "ss":
                    tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.Second2, seq));
                    break;
                case "AM/PM":
                case "am/pm":
                    tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.AmPm, seq));
                    break;
                default:
                    tokens.Add(new DxpDateTimeToken(DxpDateTimeTokenKind.Literal, seq));
                    break;
            }
            i += run;
        }

        return new DxpDateTimeFormatSpec(raw, tokens);
    }

    private static int CountRun(string format, int start)
    {
        char ch = format[start];
        int i = start;
        while (i < format.Length && format[i] == ch)
            i++;
        return i - start;
    }

    private static DxpNumericFormatSpec ParseNumericFormat(string format, string raw)
    {
        var sections = SplitNumericSections(format);
        var positive = ParseNumericSection(sections[0]);
        DxpNumericFormatSection? negative = sections.Count > 1 ? ParseNumericSection(sections[1]) : null;
        DxpNumericFormatSection? zero = sections.Count > 2 ? ParseNumericSection(sections[2]) : null;
        return new DxpNumericFormatSpec(raw, positive, negative, zero);
    }

    private static List<string> SplitNumericSections(string format)
    {
        var sections = new List<string>();
        char? quote = null;
        int last = 0;
        for (int i = 0; i < format.Length; i++)
        {
            char ch = format[i];
            if (ch == '"' || ch == '\'')
            {
                if (quote == null)
                    quote = ch;
                else if (quote == ch)
                    quote = null;
                continue;
            }

            if (quote == null && ch == ';')
            {
                sections.Add(format.Substring(last, i - last));
                last = i + 1;
            }
        }
        sections.Add(format.Substring(last));
        return sections;
    }

    private static DxpNumericFormatSection ParseNumericSection(string section)
    {
        var tokens = new List<DxpNumericToken>();
        int i = 0;
        while (i < section.Length)
        {
            char ch = section[i];
            if (ch == '"' || ch == '\'')
            {
                char quote = ch;
                int start = ++i;
                while (i < section.Length && section[i] != quote)
                    i++;
                string literal = section.Substring(start, i - start);
                tokens.Add(new DxpNumericToken(DxpNumericTokenKind.Literal, literal));
                if (i < section.Length && section[i] == quote)
                    i++;
                continue;
            }
            if (ch == '`')
            {
                int start = ++i;
                while (i < section.Length && section[i] != '`')
                    i++;
                string label = section.Substring(start, i - start);
                tokens.Add(new DxpNumericToken(DxpNumericTokenKind.NumberedItem, label));
                if (i < section.Length && section[i] == '`')
                    i++;
                continue;
            }

            switch (ch)
            {
                case '0':
                    tokens.Add(new DxpNumericToken(DxpNumericTokenKind.DigitZero, "0"));
                    break;
                case '#':
                    tokens.Add(new DxpNumericToken(DxpNumericTokenKind.DigitOptional, "#"));
                    break;
                case 'x':
                case 'X':
                    tokens.Add(new DxpNumericToken(DxpNumericTokenKind.DropDigit, ch.ToString()));
                    break;
                case '.':
                    tokens.Add(new DxpNumericToken(DxpNumericTokenKind.DecimalPoint, "."));
                    break;
                case ',':
                    tokens.Add(new DxpNumericToken(DxpNumericTokenKind.GroupingSeparator, ","));
                    break;
                case '-':
                    tokens.Add(new DxpNumericToken(DxpNumericTokenKind.MinusSign, "-"));
                    break;
                case '+':
                    tokens.Add(new DxpNumericToken(DxpNumericTokenKind.PlusSign, "+"));
                    break;
                case '%':
                    tokens.Add(new DxpNumericToken(DxpNumericTokenKind.Percent, "%"));
                    break;
                default:
                    tokens.Add(new DxpNumericToken(DxpNumericTokenKind.Literal, ch.ToString()));
                    break;
            }
            i++;
        }
        return new DxpNumericFormatSection(tokens);
    }
}
