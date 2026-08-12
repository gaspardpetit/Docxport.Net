namespace DocxportNet.Fields.Resolution;

internal sealed record DxpDatabaseFieldSpec(
    string? QueryText,
    string? DataSource,
    string? ConnectionInfo,
    bool IncludeColumnHeadings,
    int? FirstRecord,
    int? LastRecord,
    int? TableFormat,
    int? TableFormatAttributes,
    bool InsertAtMergeStart,
    IReadOnlyDictionary<string, string?> Options,
    IReadOnlyList<string> PositionalArguments);

internal static class DxpDatabaseFieldParser
{
    public static DxpDatabaseFieldSpec Parse(string instructionText)
    {
        var parsed = new DxpFieldParser().Parse(instructionText);
        var positional = string.IsNullOrWhiteSpace(parsed.Ast.ArgumentsText)
            ? []
            : DxpFieldTokenization.TokenizeArgs(parsed.Ast.ArgumentsText!);
        var switches = ParseSwitches(parsed.Ast.RawText);
        switches.TryGetValue("s", out string? query);
        bool legacyPositionalQuery = string.IsNullOrWhiteSpace(query) && positional.Count > 0;
        if (legacyPositionalQuery)
            query = positional[0];
        switches.TryGetValue("d", out string? dataSource);
        switches.TryGetValue("c", out string? connectionInfo);

        return new DxpDatabaseFieldSpec(
            query,
            dataSource,
            connectionInfo,
            switches.ContainsKey("h") || legacyPositionalQuery,
            ParsePositiveInt(switches, "f"),
            ParsePositiveInt(switches, "t"),
            ParseInt(switches, "l"),
            ParseInt(switches, "b"),
            switches.ContainsKey("o"),
            switches,
            positional);
    }

    private static Dictionary<string, string?> ParseSwitches(string rawText)
    {
        var switches = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        bool inQuote = false;
        int braceDepth = 0;
        var starts = new List<int>();
        for (int i = 0; i < rawText.Length; i++)
        {
            char ch = rawText[i];
            if (inQuote && ch == '\\' && i + 1 < rawText.Length && rawText[i + 1] == '"')
            {
                i++;
                continue;
            }
            if (ch == '"')
            {
                inQuote = !inQuote;
                continue;
            }
            if (!inQuote && ch == '{') { braceDepth++; continue; }
            if (!inQuote && ch == '}' && braceDepth > 0) { braceDepth--; continue; }
            if (!inQuote && braceDepth == 0 && ch == '\\')
                starts.Add(i);
        }

        for (int i = 0; i < starts.Count; i++)
        {
            int start = starts[i] + 1;
            int end = i + 1 < starts.Count ? starts[i + 1] : rawText.Length;
            while (start < end && char.IsWhiteSpace(rawText[start])) start++;
            if (start >= end) continue;
            string key = char.ToLowerInvariant(rawText[start]).ToString();
            string argument = rawText.Substring(start + 1, end - start - 1).Trim();
            switches[key] = Unquote(argument);
        }
        return switches;
    }

    private static string? Unquote(string value)
    {
        if (value.Length == 0)
            return null;
        // Word may split an instruction so that its field-value tokenizer has
        // already consumed one delimiter. Strip each delimiter independently.
        if (value.Length > 0 && value[0] == '"')
            value = value.Substring(1);
        if (value.Length > 0 && value[value.Length - 1] == '"')
            value = value.Substring(0, value.Length - 1);
        return value.Replace("\\\"", "\"");
    }

    private static int? ParsePositiveInt(IReadOnlyDictionary<string, string?> switches, string key)
    {
        int? value = ParseInt(switches, key);
        return value > 0 ? value : null;
    }

    private static int? ParseInt(IReadOnlyDictionary<string, string?> switches, string key)
        => switches.TryGetValue(key, out string? value) &&
           int.TryParse(value, System.Globalization.NumberStyles.Integer,
               System.Globalization.CultureInfo.InvariantCulture, out int parsed)
            ? parsed
            : null;
}
