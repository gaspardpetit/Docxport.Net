using System.Text;

namespace DocxportNet.Fields.Semantic;

internal static class DxpFieldExpressionTokenizer
{
    public static IReadOnlyList<DxpFieldExpressionToken> Tokenize(DxpFieldExpression expression)
    {
        var tokens = new List<DxpFieldExpressionToken>();
        var current = new List<DxpFieldTemplatePart>();
        var literal = new StringBuilder();
        bool inQuote = false;
        bool justClosedQuote = false;

        void FlushLiteral()
        {
            if (literal.Length == 0)
                return;
            current.Add(new DxpFieldTemplateText(literal.ToString()));
            literal.Clear();
        }

        void FlushToken(bool allowEmpty = false)
        {
            FlushLiteral();
            if (current.Count == 0 && !allowEmpty)
                return;
            tokens.Add(new DxpFieldExpressionToken(current.ToArray()));
            current.Clear();
        }

        foreach (DxpFieldExpressionPart part in expression.Parts)
        {
            if (part is DxpFieldExpressionChild child)
            {
                FlushLiteral();
                current.Add(new DxpFieldTemplateChild(child.Expression));
                justClosedQuote = false;
                continue;
            }
            if (part is DxpFieldExpressionParagraph)
            {
                FlushLiteral();
                current.Add(new DxpFieldTemplateParagraph());
                continue;
            }

            string text = ((DxpFieldExpressionText)part).Text;
            for (int index = 0; index < text.Length; index++)
            {
                char ch = text[index];
                if (ch == '"')
                {
                    if (inQuote && index > 0 && text[index - 1] == '\\')
                    {
                        if (literal.Length > 0)
                            literal.Length--;
                        literal.Append('"');
                        continue;
                    }
                    inQuote = !inQuote;
                    if (!inQuote)
                        justClosedQuote = true;
                    continue;
                }

                if (!inQuote && char.IsWhiteSpace(ch))
                {
                    FlushToken(justClosedQuote);
                    justClosedQuote = false;
                    continue;
                }

                literal.Append(ch);
                justClosedQuote = false;
            }
        }

        FlushToken(justClosedQuote);
        return tokens;
    }
}
