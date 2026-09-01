using System.Globalization;
using System.Text;

namespace DocxportNet.Omml;

internal static class OmmlTokenClassifier
{
    public static IReadOnlyList<OmmlToken> Classify(string value, bool textMode)
    {
        if (value.Length == 0) return new[] { new OmmlToken(OmmlTokenKind.Text, string.Empty) };
        if (textMode) return new[] { new OmmlToken(OmmlTokenKind.Text, value) };
        List<OmmlToken> result = new();
        StringBuilder current = new();
        OmmlTokenKind? kind = null;
        for (int i = 0; i < value.Length;)
        {
            int length = char.IsHighSurrogate(value[i]) && i + 1 < value.Length && char.IsLowSurrogate(value[i + 1]) ? 2 : 1;
            string scalar = value.Substring(i, length);
            UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(value, i);
            OmmlTokenKind next = GetKind(category, scalar);
            if ((scalar == "." || scalar == ",") && i > 0 && i + length < value.Length &&
                char.IsDigit(value, i - 1) && char.IsDigit(value, i + length))
                next = OmmlTokenKind.Number;
            if ((category == UnicodeCategory.NonSpacingMark || category == UnicodeCategory.SpacingCombiningMark) && kind.HasValue)
                next = kind.Value;
            if (kind.HasValue && (next != kind.Value || scalar == "\u200B" || current.ToString() == "\u200B")) { result.Add(new OmmlToken(kind.Value, current.ToString())); current.Clear(); }
            kind = next; current.Append(scalar); i += length;
        }
        if (kind.HasValue) result.Add(new OmmlToken(kind.Value, current.ToString()));
        return result;
    }

    private static OmmlTokenKind GetKind(UnicodeCategory category, string scalar)
    {
        if (char.IsWhiteSpace(scalar, 0) || scalar == "\u200B" || category == UnicodeCategory.Control || category == UnicodeCategory.Format) return OmmlTokenKind.Text;
        if (category == UnicodeCategory.DecimalDigitNumber) return OmmlTokenKind.Number;
        if (category == UnicodeCategory.UppercaseLetter || category == UnicodeCategory.LowercaseLetter ||
            category == UnicodeCategory.TitlecaseLetter || category == UnicodeCategory.ModifierLetter ||
            category == UnicodeCategory.OtherLetter || category == UnicodeCategory.LetterNumber)
            return OmmlTokenKind.Identifier;
        return OmmlTokenKind.Operator;
    }
}
