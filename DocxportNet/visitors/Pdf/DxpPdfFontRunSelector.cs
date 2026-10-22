using System.Text;

namespace DocxportNet.Visitors.Pdf;

internal static class DxpPdfFontRunSelector
{
    internal sealed record Run(string Text, string Family);

    public static IReadOnlyList<Run> Split(string text, string requestedFamily,
        bool bold, bool italic, IDxpPdfFontProvider provider)
    {
        var runs = new List<Run>();
        string? previous = null;
        int start = 0;
        int offset = 0;
        foreach (Rune rune in text.EnumerateRunes())
        {
            string family = provider.ResolveForCharacter(requestedFamily, bold, italic, rune)?.FamilyName
                ?? requestedFamily;
            if (previous != null && !string.Equals(previous, family, StringComparison.OrdinalIgnoreCase))
            {
                runs.Add(new Run(text[start..offset], previous));
                start = offset;
            }
            previous = family;
            offset += rune.Utf16SequenceLength;
        }
        if (start < text.Length)
            runs.Add(new Run(text[start..], previous ?? requestedFamily));
        return runs;
    }
}
