using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet.Walker;

// Exact custom aliases recognized by Word; these are not free-form masks.
internal static class DxpCustomListFormats
{
    private static readonly Dictionary<string, NumberFormatValues> Aliases = new(StringComparer.Ordinal)
    {
        ["A"] = NumberFormatValues.UpperLetter,
        ["a"] = NumberFormatValues.LowerLetter,
        ["I"] = NumberFormatValues.UpperRoman,
        ["i"] = NumberFormatValues.LowerRoman,
    };

    internal static bool TryResolve(NumberingFormat? format, out NumberFormatValues value)
    {
        value = default;
        return format?.Val?.Value == NumberFormatValues.Custom &&
            format.Format?.Value is string pattern && Aliases.TryGetValue(pattern, out value);
    }
}
