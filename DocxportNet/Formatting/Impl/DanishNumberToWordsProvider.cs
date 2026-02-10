using System.Globalization;

namespace DocxportNet.Formatting.Impl;

public sealed class DanishNumberToWordsProvider : DxpINumberToWordsProvider
{
    public bool CanHandle(CultureInfo culture) => culture.TwoLetterISOLanguageName.Equals("da", StringComparison.OrdinalIgnoreCase);

    public string ToCardinal(int number)
    {
        if (number == 0)
            return "nul";
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
        string thousandWord = thousands == 1 ? "tusind" : ToCardinal(thousands) + "tusind";
        return rest == 0 ? thousandWord : thousandWord + ToCardinal(rest);
    }

    public string ToOrdinalWords(int number)
    {
        if (number == 0)
            return "nulte";
        if (number < 0)
            return "minus " + ToOrdinalWords(Math.Abs(number));
        if (number > 32767)
            return number.ToString(CultureInfo.InvariantCulture);
        if (number == 1)
            return "første";
        if (number == 2)
            return "anden";
        if (number == 3)
            return "tredje";
        if (number == 4)
            return "fjerde";
        if (number == 5)
            return "femte";
        if (number == 6)
            return "sjette";
        if (number == 7)
            return "syvende";
        if (number == 8)
            return "ottende";
        if (number == 9)
            return "niende";
        if (number == 10)
            return "tiende";

        return ToCardinal(number) + "ende";
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
        string result = words + " og " + centsText;
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
        if (tens == 2)
            return Small[ones] + "ogtyve";
        if (tens == 3)
            return Small[ones] + "ogtredive";
        if (tens == 4)
            return Small[ones] + "ogfyrre";
        if (tens == 5)
            return Small[ones] + "oghalvtreds";
        if (tens == 6)
            return Small[ones] + "ogtres";
        if (tens == 7)
            return Small[ones] + "oghalvfjerds";
        if (tens == 8)
            return Small[ones] + "ogfirs";
        if (tens == 9)
            return Small[ones] + "oghalvfems";
        return Small[ones] + "og" + Tens[tens];
    }

    private string ThreeDigit(int number)
    {
        int hundreds = number / 100;
        int rest = number % 100;
        string hundredWord = hundreds == 1 ? "hundrede" : Small[hundreds] + "hundrede";
        if (rest == 0)
            return hundredWord;
        return hundredWord + ToCardinal(rest);
    }

    private static readonly string[] Small =
    {
        "nul","en","to","tre","fire","fem","seks","syv","otte","ni","ti",
        "elleve","tolv","tretten","fjorten","femten","seksten","sytten","atten","nitten"
    };

    private static readonly string[] Tens =
    {
        "", "", "tyve", "tredive", "fyrre", "halvtreds", "tres", "halvfjerds", "firs", "halvfems"
    };
}
