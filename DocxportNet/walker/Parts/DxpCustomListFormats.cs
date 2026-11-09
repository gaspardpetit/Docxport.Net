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
        ["\u0623, \u0628, \u062a, ..."] = NumberFormatValues.ArabicAlpha,
        ["\u0623, \u0628, \u062c, ..."] = NumberFormatValues.ArabicAbjad,
        ["\u05d0, \u05d9, \u05e7, ..."] = NumberFormatValues.Hebrew1,
        ["\u05d0, \u05d1, \u05d2, ..."] = NumberFormatValues.Hebrew2,
        ["\u0905, \u0906, \u0907, ..."] = NumberFormatValues.HindiConsonants,
        ["\u0915, \u0916, \u0917, ..."] = NumberFormatValues.HindiVowels,
        ["\u0e01, \u0e02, \u0e04, ..."] = NumberFormatValues.ThaiLetters,
        ["\u0430, \u0431, \u0432, ..."] = NumberFormatValues.RussianLower,
        ["\u0410, \u0411, \u0412, ..."] = NumberFormatValues.RussianUpper,
        ["\uff71, \uff72, \uff73, ..."] = NumberFormatValues.Aiueo,
        ["\u30a2, \u30a4, \u30a6, ..."] = NumberFormatValues.AiueoFullWidth,
        ["\uff72, \uff9b, \uff8a, ..."] = NumberFormatValues.Iroha,
        ["\u30a4, \u30ed, \u30cf, ..."] = NumberFormatValues.IrohaFullWidth,
        ["\uac00, \ub098, \ub2e4, ..."] = NumberFormatValues.Ganada,
        ["\u3131, \u3134, \u3137, ..."] = NumberFormatValues.Chosung,
    };

    internal static bool TryResolve(NumberingFormat? format, out NumberFormatValues value)
    {
        value = default;
        return format?.Val?.Value == NumberFormatValues.Custom &&
            format.Format?.Value is string pattern && Aliases.TryGetValue(pattern, out value);
    }
}
