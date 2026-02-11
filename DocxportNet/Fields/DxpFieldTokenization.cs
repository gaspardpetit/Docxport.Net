namespace DocxportNet.Fields;

internal static class DxpFieldTokenization
{
    public static bool TryGetFirstToken(string? argsText, out string token)
    {
        token = string.Empty;
        if (string.IsNullOrWhiteSpace(argsText))
            return false;

        var tokens = TokenizeArgs(argsText!);
        if (tokens.Count == 0)
            return false;
        token = tokens[0];
        return true;
    }

    public static List<string> TokenizeArgs(string text)
    {
        var tokens = new List<string>();
        bool inQuote = false;
        bool justClosedQuote = false;
        var current = new System.Text.StringBuilder();
        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            if (ch == '"')
            {
                if (inQuote && i > 0 && text[i - 1] == '\\')
                {
                    if (current.Length > 0)
                        current.Length -= 1; // remove escape backslash
                    current.Append('"');
                    continue;
                }
                inQuote = !inQuote;
                if (!inQuote)
                    justClosedQuote = true;
                continue;
            }

            if (!inQuote && ch == '{')
            {
                int start = i;
                int depth = 0;
                for (; i < text.Length; i++)
                {
                    if (text[i] == '{')
                        depth++;
                    else if (text[i] == '}')
                    {
                        depth--;
                        if (depth == 0)
                        {
                            i++;
                            break;
                        }
                    }
                }
                string token = text.Substring(start, i - start).Trim();
                if (current.Length > 0)
                {
                    tokens.Add(current.ToString());
                    current.Clear();
                }
                if (token.Length > 0)
                    tokens.Add(token);
                justClosedQuote = false;
                i--;
                continue;
            }

            if (!inQuote && char.IsWhiteSpace(ch))
            {
                if (current.Length > 0)
                {
                    tokens.Add(current.ToString());
                    current.Clear();
                    justClosedQuote = false;
                }
                else if (justClosedQuote)
                {
                    tokens.Add(string.Empty);
                    justClosedQuote = false;
                }
                continue;
            }

            current.Append(ch);
            justClosedQuote = false;
        }

        if (current.Length > 0)
            tokens.Add(current.ToString());
        else if (justClosedQuote)
            tokens.Add(string.Empty);
        return tokens;
    }
}
