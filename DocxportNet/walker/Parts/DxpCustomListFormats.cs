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
        ["\uc77c, \uc774, \uc0bc, ..."] = NumberFormatValues.KoreanCounting,
        ["\ud558\ub098, \ub458, \uc14b, ..."] = NumberFormatValues.KoreanLegal,
        ["\uc77c, \uc77c\uc601, \uc77c\uc601\uc601, ..."] = NumberFormatValues.KoreanDigital,
        ["\u4e00, \u4e00\u96f6, \u4e00\u96f6\u96f6, ..."] = NumberFormatValues.KoreanDigital2,
        ["\u0967, \u0968, \u0969, ..."] = NumberFormatValues.HindiNumbers,
        ["\u0e51, \u0e52, \u0e53, ..."] = NumberFormatValues.ThaiNumbers,
        ["\u4e00, \u4e00\u25cb, \u4e00\u25cb\u25cb(\u7e41), ..."] = NumberFormatValues.TaiwaneseDigital,
        ["\u7532, \u4e59, \u4e19, ..."] = NumberFormatValues.IdeographTraditional,
        ["\u5b50, \u4e11, \u5bc5, ..."] = NumberFormatValues.IdeographZodiac,
        ["\u7532\u5b50, \u4e59\u4e11, \u4e19\u5bc5, ..."] = NumberFormatValues.IdeographZodiacTraditional,
        ["\u3220, \u3221, \u3222, ..."] = NumberFormatValues.IdeographEnclosedCircle,
        ["1, A, B, ..."] = NumberFormatValues.Hex,
        ["*, \u2020, \u2021, ..."] = NumberFormatValues.Chicago,
        ["\u2460, \u2461, \u2462, ..."] = NumberFormatValues.DecimalEnclosedCircle,
        ["\u2488, \u2489, \u248a, ..."] = NumberFormatValues.DecimalEnclosedFullstop,
        ["\u2474, \u2475, \u2476, ..."] = NumberFormatValues.DecimalEnclosedParen,
    };

    internal static bool TryResolve(NumberingFormat? format, out NumberFormatValues value)
    {
        value = default;
        return format?.Val?.Value == NumberFormatValues.Custom &&
            format.Format?.Value is string pattern && Aliases.TryGetValue(pattern, out value);
    }
}
