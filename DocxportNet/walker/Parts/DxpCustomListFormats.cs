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
        ["\uff11, \uff12, \uff13, ..."] = NumberFormatValues.DecimalFullWidth,
        ["\u4e00, \u4e8c, \u4e09, ..."] = NumberFormatValues.JapaneseCounting,
        ["\u4e00, \u4e8c, \u4e09, \u4e07, ..."] = NumberFormatValues.JapaneseDigitalTenThousand,
        ["\u58f1, \u5f10, \u53c2, ..."] = NumberFormatValues.JapaneseLegal,
        ["\u4e00, \u5341, \u4e00\u25cb\u25cb(\u7b80), ..."] = NumberFormatValues.ChineseCounting,
        ["\u4e00, \u5341, \u4e00\u767e(\u7b80), ..."] = NumberFormatValues.ChineseCountingThousand,
        ["\u4e00, \u5341, \u4e00\u25cb\u25cb(\u7e41), ..."] = NumberFormatValues.TaiwaneseCounting,
        ["\u4e00, \u5341, \u4e00\u767e(\u7e41), ..."] = NumberFormatValues.TaiwaneseCountingThousand,
        ["\u58f9, \u8d30, \u53c1, ..."] = NumberFormatValues.ChineseLegalSimplified,
        ["\u58f9, \u8cb3, \u53c3, ..."] = NumberFormatValues.IdeographLegalTraditional,
        ["\u4e00, \u4e00\u3007, \u4e00\u3007\u3007, ..."] = NumberFormatValues.IdeographDigital,
        ["1st, 2nd, 3rd, ..."] = NumberFormatValues.Ordinal,
        ["- 1 -, - 2 -, - 3 -, ..."] = NumberFormatValues.NumberInDash,
        ["One, Two, Three, ..."] = NumberFormatValues.CardinalText,
        ["First, Second, Third, ..."] = NumberFormatValues.OrdinalText,
        [""] = NumberFormatValues.None,
    };

    internal static bool TryResolve(NumberingFormat? format, out NumberFormatValues value)
    {
        value = default;
        return format?.Val?.Value == NumberFormatValues.Custom &&
            format.Format?.Value is string pattern && Aliases.TryGetValue(pattern, out value);
    }
    private static readonly Dictionary<string, string> ExtensionAlphabets = new(StringComparer.Ordinal)
    {
        ["A, \u00c7, \u011c, ..."] = "ABC\u00c7DEFG\u011eHI\u0130JKLMNO\u00d6PRS\u015eTU\u00dcVYZ",
        ["a, \u00e7, \u011d, ..."] = "abc\u00e7defg\u011fh\u0131ijklmno\u00f6prs\u015ftu\u00fcvyz",
        ["\u0410, \u0419, \u041a, ..."] = "\u0410\u0411\u0412\u0413\u0414\u0415\u0416\u0417\u0418\u0419\u041a\u041b\u041c\u041d\u041e\u041f\u0420\u0421\u0422\u0423\u0424\u0425\u0426\u0427\u0428\u0429\u042a\u042e\u042f",
        ["\u0430, \u0439, \u043a, ..."] = "\u0430\u0431\u0432\u0433\u0434\u0435\u0436\u0437\u0438\u0439\u043a\u043b\u043c\u043d\u043e\u043f\u0440\u0441\u0442\u0443\u0444\u0445\u0446\u0447\u0448\u0449\u044a\u044e\u044f",
    };

    private const string GreekUpperPattern = "\u0391, \u0392, \u0393, ...";
    private const string GreekLowerPattern = "\u03b1, \u03b2, \u03b3, ...";
    private static readonly string[] GreekUnits = { "\u0391", "\u0392", "\u0393", "\u0394", "\u0395", "\u03a3\u03a4", "\u0396", "\u0397", "\u0398" };
    private const string GreekTens = "\u0399\u039a\u039b\u039c\u039d\u039e\u039f\u03a0\u03de";
    private const string GreekHundreds = "\u03a1\u03a3\u03a4\u03a5\u03a6\u03a7\u03a8\u03a9\u03e0";

    internal static string? FormatExtension(NumberingFormat? format, int number)
    {
        if (format?.Val?.Value != NumberFormatValues.Custom || format.Format?.Value is not string pattern)
            return null;
        if (ExtensionAlphabets.TryGetValue(pattern, out var alphabet))
        {
            if (number < 0) return "";
            if (number == 0) return " ";
            // Word repeats each alphabet letter, wrapping after thirty repetitions.
            return new string(alphabet[(number - 1) % alphabet.Length], (number - 1) / alphabet.Length % 30 + 1);
        }
        if (pattern != GreekUpperPattern && pattern != GreekLowerPattern) return null;
        if (number < 0) return "";
        if (number == 0) return " ";
        // Greek numeric lists cycle through 1..9999, using additive symbols.
        var value = (number - 1) % 9999 + 1;
        var text = value >= 1000 ? "," + GreekUnits[value / 1000 - 1] : "";
        value %= 1000;
        if (value >= 100) text += GreekHundreds[value / 100 - 1];
        value %= 100;
        if (value >= 10) text += GreekTens[value / 10 - 1];
        if (value % 10 != 0) text += GreekUnits[value % 10 - 1];
        return pattern == GreekLowerPattern ? text.ToLowerInvariant() : text;
    }

}
