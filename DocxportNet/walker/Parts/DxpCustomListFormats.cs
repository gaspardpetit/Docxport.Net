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
        ["\u090f\u0915, \u0926\u094b, \u0924\u0940\u0928, ..."] = NumberFormatValues.HindiCounting,
        ["\u0e2b\u0e19\u0e36\u0e48\u0e07, \u0e2a\u0e2d\u0e07, \u0e2a\u0e32\u0e21, ..."] = NumberFormatValues.ThaiCounting,
        ["m\u00f4\u0323t, hai, ba, ..."] = NumberFormatValues.VietnameseCounting,
    };

    internal static bool TryResolve(NumberingFormat? format, out NumberFormatValues value)
    {
        value = default;
        return format?.Val?.Value == NumberFormatValues.Custom &&
            format.Format?.Value is string pattern && Aliases.TryGetValue(pattern, out value);
    }
}
