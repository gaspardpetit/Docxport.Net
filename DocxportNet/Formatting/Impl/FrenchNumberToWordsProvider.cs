using System.Globalization;

namespace DocxportNet.Formatting.Impl;

public sealed class FrenchNumberToWordsProvider : DxpINumberToWordsProvider
{
    public bool CanHandle(CultureInfo culture) => culture.TwoLetterISOLanguageName.Equals("fr", StringComparison.OrdinalIgnoreCase);

    public string ToCardinal(int number)
    {
        if (number == 0)
            return "zéro";
        if (number < 0)
            return "moins " + ToCardinal(Math.Abs(number));
        if (number > 32767)
            return number.ToString(CultureInfo.InvariantCulture);

        if (number < 17)
            return Small[number];
        if (number < 20)
            return "dix-" + Small[number - 10];
        if (number < 70)
            return TwoDigit(number);
        if (number < 80)
            return "soixante-" + ToCardinal(number - 60);
        if (number < 100)
            return number == 80 ? "quatre-vingts" : "quatre-vingt-" + ToCardinal(number - 80);
        if (number < 1000)
            return ThreeDigit(number);
        int thousands = number / 1000;
        int rest = number % 1000;
        string mille = thousands == 1 ? "mille" : ToCardinal(thousands) + " mille";
        return rest == 0 ? mille : mille + " " + ToCardinal(rest);
    }

    public string ToOrdinalWords(int number)
    {
        if (number == 0)
            return "zéroième";
        if (number < 0)
            return "moins " + ToOrdinalWords(Math.Abs(number));
        if (number > 32767)
            return number.ToString(CultureInfo.InvariantCulture);
        if (number == 1)
            return "premier";

        string baseCard = ToCardinal(number);
        if (baseCard.EndsWith("cinq", StringComparison.Ordinal))
            return baseCard.Substring(0, baseCard.Length - 4) + "cinquième";
        if (baseCard.EndsWith("neuf", StringComparison.Ordinal))
            return baseCard.Substring(0, baseCard.Length - 3) + "neuvième";
        if (baseCard.EndsWith("e", StringComparison.Ordinal))
            baseCard = baseCard.Substring(0, baseCard.Length - 1);
        return baseCard + "ième";
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
        string result = words + " et " + centsText;
        if (number < 0)
            result = "moins " + result;
        return result;
    }

    private static string TwoDigit(int number)
    {
        int tens = number / 10;
        int ones = number % 10;
        string tensWord = Tens[tens];
        if (ones == 0)
            return tensWord;
        if (ones == 1 && (tens == 2 || tens == 3 || tens == 4 || tens == 5 || tens == 6))
            return tensWord + " et un";
        return tensWord + "-" + Small[ones];
    }

    private string ThreeDigit(int number)
    {
        int hundreds = number / 100;
        int rest = number % 100;
        string hundredWord;
        if (hundreds == 1)
            hundredWord = "cent";
        else
            hundredWord = Small[hundreds] + " cent";
        if (rest == 0 && hundreds > 1)
            return hundredWord + "s";
        if (rest == 0)
            return hundredWord;
        return hundredWord + " " + ToCardinal(rest);
    }

    private static readonly string[] Small =
    {
        "zéro","un","deux","trois","quatre","cinq","six","sept","huit","neuf","dix",
        "onze","douze","treize","quatorze","quinze","seize"
    };

    private static readonly string[] Tens =
    {
        "", "", "vingt", "trente", "quarante", "cinquante", "soixante"
    };
}
