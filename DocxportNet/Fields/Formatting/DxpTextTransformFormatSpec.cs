using DocxportNet.Formatting;

namespace DocxportNet.Fields.Formatting;

public enum DxpTextTransformKind
{
    Caps,
    FirstCap,
    Upper,
    Lower,
    Alphabetic,
    Arabic,
    ArabicDash,
    CardText,
    DollarText,
    Hex,
    OrdText,
    Ordinal,
    Roman,
    Charformat,
    MergeFormat
}

public sealed class DxpTextTransformFormatSpec : IDxpFieldFormatSpec
{
    public DxpTextTransformKind Kind { get; }
    public string RawText { get; }
    public string Argument { get; }

    public DxpTextTransformFormatSpec(DxpTextTransformKind kind, string rawText, string argument)
    {
        Kind = kind;
        RawText = rawText;
        Argument = argument;
    }

    public string Apply(string text, DxpFieldValue value, DxpFieldEvalContext context)
    {
        return Kind switch {
            DxpTextTransformKind.Upper => ToUpper(text, context),
            DxpTextTransformKind.Lower => ToLower(text, context),
            DxpTextTransformKind.FirstCap => FirstCap(text, context),
            DxpTextTransformKind.Caps => Caps(text, context),
            DxpTextTransformKind.Alphabetic => ApplyCaseFromArgument(FormatAlphabetic(value, text, context), Argument),
            DxpTextTransformKind.Arabic => FormatArabic(value, text, context),
            DxpTextTransformKind.ArabicDash => "-" + FormatArabic(value, text, context) + "-",
            DxpTextTransformKind.CardText => ApplyCaseFromArgument(FormatCardText(value, text, context), Argument),
            DxpTextTransformKind.DollarText => ApplyCaseFromArgument(FormatDollarText(value, text, context), Argument),
            DxpTextTransformKind.Hex => ApplyCaseFromArgument(FormatHex(value, text, context), Argument),
            DxpTextTransformKind.OrdText => ApplyCaseFromArgument(FormatOrdText(value, text, context), Argument),
            DxpTextTransformKind.Ordinal => ApplyCaseFromArgument(FormatOrdinal(value, text, context), Argument),
            DxpTextTransformKind.Roman => ApplyCaseFromArgument(FormatRoman(value, text, context), Argument),
            _ => text
        };
    }

    private static string FirstCap(string text, DxpFieldEvalContext context)
    {
        if (string.IsNullOrEmpty(text))
            return text;
        int i = 0;
        while (i < text.Length && char.IsWhiteSpace(text[i]))
            i++;
        if (i >= text.Length)
            return text;
        var ti = (context.Culture ?? System.Globalization.CultureInfo.CurrentCulture).TextInfo;
        return text.Substring(0, i) + ti.ToUpper(text[i].ToString()) + text.Substring(i + 1);
    }

    private static string Caps(string text, DxpFieldEvalContext context)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        var ti = (context.Culture ?? System.Globalization.CultureInfo.CurrentCulture).TextInfo;
        char[] chars = text.ToCharArray();
        bool startWord = true;
        for (int i = 0; i < chars.Length; i++)
        {
            char ch = chars[i];
            if (char.IsLetter(ch))
            {
                if (startWord)
                    chars[i] = ti.ToUpper(ch.ToString())[0];
                startWord = false;
            }
            else
            {
                startWord = char.IsWhiteSpace(ch);
            }
        }
        return new string(chars);
    }

    private static string ToUpper(string text, DxpFieldEvalContext context)
    {
        var ti = (context.Culture ?? System.Globalization.CultureInfo.CurrentCulture).TextInfo;
        return ti.ToUpper(text);
    }

    private static string ToLower(string text, DxpFieldEvalContext context)
    {
        var ti = (context.Culture ?? System.Globalization.CultureInfo.CurrentCulture).TextInfo;
        return ti.ToLower(text);
    }

    private static string ApplyCaseFromArgument(string text, string argument)
    {
        if (string.IsNullOrEmpty(text))
            return text;
        if (string.IsNullOrWhiteSpace(argument))
            return text;

        string arg = argument.Trim();
        if (IsAllUpper(arg))
            return text.ToUpperInvariant();
        if (IsAllLower(arg))
            return text.ToLowerInvariant();
        return text;
    }

    private static bool IsAllUpper(string text)
    {
        bool any = false;
        foreach (char ch in text)
        {
            if (char.IsLetter(ch))
            {
                any = true;
                if (!char.IsUpper(ch))
                    return false;
            }
        }
        return any;
    }

    private static bool IsAllLower(string text)
    {
        bool any = false;
        foreach (char ch in text)
        {
            if (char.IsLetter(ch))
            {
                any = true;
                if (!char.IsLower(ch))
                    return false;
            }
        }
        return any;
    }

    private static bool TryGetNumber(DxpFieldValue value, string text, DxpFieldEvalContext context, out double number)
    {
        if (value.Kind == DxpFieldValueKind.Number && value.NumberValue.HasValue)
        {
            number = value.NumberValue.Value;
            return true;
        }

        if (!string.IsNullOrWhiteSpace(text))
        {
            var culture = context.Culture ?? System.Globalization.CultureInfo.CurrentCulture;
            if (double.TryParse(text, System.Globalization.NumberStyles.Any, culture, out number))
                return true;
            if (context.AllowInvariantNumericFallback &&
                double.TryParse(text, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out number))
                return true;
        }

        number = 0;
        return false;
    }

    private static string FormatArabic(DxpFieldValue value, string text, DxpFieldEvalContext context)
    {
        if (TryGetNumber(value, text, context, out var number))
            return number.ToString(context.Culture ?? System.Globalization.CultureInfo.CurrentCulture);
        return text;
    }

    private static string FormatHex(DxpFieldValue value, string text, DxpFieldEvalContext context)
    {
        if (!TryGetNumber(value, text, context, out var number))
            return text;
        long n = (long)Math.Round(number);
        return n.ToString("X", System.Globalization.CultureInfo.InvariantCulture);
    }

    private static string FormatAlphabetic(DxpFieldValue value, string text, DxpFieldEvalContext context)
    {
        if (!TryGetNumber(value, text, context, out var number))
            return text;
        int n = (int)Math.Round(number);
        if (n <= 0)
            return text;
        return ToAlphabetic(n);
    }

    private static string ToAlphabetic(int value)
    {
        var chars = new Stack<char>();
        int n = value;
        while (n > 0)
        {
            n--;
            chars.Push((char)('A' + (n % 26)));
            n /= 26;
        }
        return new string(chars.ToArray());
    }

    private static string FormatRoman(DxpFieldValue value, string text, DxpFieldEvalContext context)
    {
        if (!TryGetNumber(value, text, context, out var number))
            return text;
        int n = (int)Math.Round(number);
        if (n <= 0 || n > 3999)
            return text;
        return ToRoman(n);
    }

    private static string ToRoman(int value)
    {
        var map = new (int Value, string Symbol)[]
        {
            (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
            (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
            (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
        };
        var sb = new System.Text.StringBuilder();
        int n = value;
        foreach (var (val, sym) in map)
        {
            while (n >= val)
            {
                sb.Append(sym);
                n -= val;
            }
        }
        return sb.ToString();
    }

    private static string FormatOrdinal(DxpFieldValue value, string text, DxpFieldEvalContext context)
    {
        if (!TryGetNumber(value, text, context, out var number))
            return text;
        long n = (long)Math.Round(number);
        return n.ToString(System.Globalization.CultureInfo.InvariantCulture) + OrdinalSuffix(n);
    }

    private static string OrdinalSuffix(long n)
    {
        long abs = Math.Abs(n);
        long mod100 = abs % 100;
        if (mod100 >= 11 && mod100 <= 13)
            return "th";
        return (abs % 10) switch {
            1 => "st",
            2 => "nd",
            3 => "rd",
            _ => "th"
        };
    }

    private static string FormatOrdText(DxpFieldValue value, string text, DxpFieldEvalContext context)
    {
        if (!TryGetNumber(value, text, context, out var number))
            return text;
        int n = (int)Math.Round(number);
        var provider = ResolveNumberProvider(context);
        return provider.ToOrdinalWords(n);
    }

    private static string FormatCardText(DxpFieldValue value, string text, DxpFieldEvalContext context)
    {
        if (!TryGetNumber(value, text, context, out var number))
            return text;
        int n = (int)Math.Round(number);
        var provider = ResolveNumberProvider(context);
        return provider.ToCardinal(n);
    }

    private static string FormatDollarText(DxpFieldValue value, string text, DxpFieldEvalContext context)
    {
        if (!TryGetNumber(value, text, context, out var number))
            return text;
        var provider = ResolveNumberProvider(context);
        return provider.ToDollarText(number);
    }

    private static IDxpNumberToWordsProvider ResolveNumberProvider(DxpFieldEvalContext context)
    {
        if (context.NumberToWordsProvider != null)
            return context.NumberToWordsProvider;
        var culture = context.Culture ?? System.Globalization.CultureInfo.CurrentCulture;
        return context.NumberToWordsRegistry.Resolve(culture);
    }
}
