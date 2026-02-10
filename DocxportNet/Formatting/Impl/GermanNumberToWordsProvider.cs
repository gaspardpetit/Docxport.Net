using System.Globalization;

namespace DocxportNet.Formatting.Impl;

public sealed class GermanNumberToWordsProvider : DxpINumberToWordsProvider
{
    public bool CanHandle(CultureInfo culture) => culture.TwoLetterISOLanguageName.Equals("de", StringComparison.OrdinalIgnoreCase);

    public string ToCardinal(int number)
    {
        if (number == 0)
            return "null";
        if (number < 0)
            return "minus " + ToCardinal(Math.Abs(number));
        if (number > 32767)
            return number.ToString(CultureInfo.InvariantCulture);

        if (number < 20)
            return Small[number];
        if (number < 100)
            return TwoDigit(number);
        if (number < 1000)
            return ThreeDigit(number);

        int thousands = number / 1000;
        int rest = number % 1000;
        string thousandWord = thousands == 1 ? "tausend" : ToCardinal(thousands) + "tausend";
        return rest == 0 ? thousandWord : thousandWord + ThreeDigit(rest);
    }

    public string ToOrdinalWords(int number)
    {
        if (number == 0)
            return "nullte";
        if (number < 0)
            return "minus " + ToOrdinalWords(Math.Abs(number));
        if (number > 32767)
            return number.ToString(CultureInfo.InvariantCulture);

        if (number == 1)
            return "erste";
        if (number == 2)
            return "zweite";
        if (number == 3)
            return "dritte";
        if (number == 7)
            return "siebte";
        if (number == 8)
            return "achte";

        string baseCard = ToCardinal(number);
        if (number < 20)
            return baseCard + "te";
        return baseCard + "ste";
    }

    public string ToDollarText(double number)
    {
        double abs = Math.Abs(number);
        long dollars = (long)Math.Floor(abs);
        int cents = (int)Math.Round((abs - dollars) * 100, MidpointRounding.AwayFromZero);
        if (cents == 100)
        {
            dollars += 1;
            cents = 0;
        }

        string words = ToCardinal((int)dollars);
        string centsText = cents.ToString("00", CultureInfo.InvariantCulture) + "/100";
        string result = words + " und " + centsText;
        if (number < 0)
            result = "minus " + result;
        return result;
    }

    private static string TwoDigit(int number)
    {
        int tens = number / 10;
        int ones = number % 10;
        if (ones == 0)
            return Tens[tens];
        string onesWord = ones == 1 ? "ein" : Small[ones];
        return onesWord + "und" + Tens[tens];
    }

    private static string ThreeDigit(int number)
    {
        int hundreds = number / 100;
        int rest = number % 100;
        string hundredWord = hundreds == 1 ? "einhundert" : Small[hundreds] + "hundert";
        if (rest == 0)
            return hundredWord;
        return hundredWord + (rest < 20 ? Small[rest] : TwoDigit(rest));
    }

    private static readonly string[] Small =
    {
        "null","eins","zwei","drei","vier","fünf","sechs","sieben","acht","neun","zehn",
        "elf","zwölf","dreizehn","vierzehn","fünfzehn","sechzehn","siebzehn","achtzehn","neunzehn"
    };

    private static readonly string[] Tens =
    {
        "", "", "zwanzig", "dreißig", "vierzig", "fünfzig", "sechzig", "siebzig", "achtzig", "neunzig"
    };
}
