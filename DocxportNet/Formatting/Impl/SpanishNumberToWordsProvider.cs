using System.Globalization;

namespace DocxportNet.Formatting.Impl;

public sealed class SpanishNumberToWordsProvider : IDxpNumberToWordsProvider
{
    public bool CanHandle(CultureInfo culture) => culture.TwoLetterISOLanguageName.Equals("es", StringComparison.OrdinalIgnoreCase);

    public string ToCardinal(int number)
    {
        if (number == 0)
            return "cero";
        if (number < 0)
            return "menos " + ToCardinal(Math.Abs(number));
        if (number > 32767)
            return number.ToString(CultureInfo.InvariantCulture);

        if (number < 30)
            return Small[number];
        if (number < 100)
        {
            int tens = number / 10;
            int ones = number % 10;
            if (ones == 0)
                return Tens[tens];
            return Tens[tens] + " y " + Small[ones];
        }
        if (number < 1000)
            return ThreeDigit(number);

        int thousands = number / 1000;
        int rest = number % 1000;
        string thousandWord = thousands == 1 ? "mil" : ToCardinal(thousands) + " mil";
        return rest == 0 ? thousandWord : thousandWord + " " + ToCardinal(rest);
    }

    public string ToOrdinalWords(int number)
    {
        if (number == 0)
            return "cero";
        if (number < 0)
            return "menos " + ToOrdinalWords(Math.Abs(number));
        if (number > 32767)
            return number.ToString(CultureInfo.InvariantCulture);
        if (number == 1)
            return "primero";
        if (number == 2)
            return "segundo";
        if (number == 3)
            return "tercero";
        if (number == 4)
            return "cuarto";
        if (number == 5)
            return "quinto";
        if (number == 6)
            return "sexto";
        if (number == 7)
            return "séptimo";
        if (number == 8)
            return "octavo";
        if (number == 9)
            return "noveno";
        if (number == 10)
            return "décimo";

        return ToCardinal(number) + "avo";
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
        string result = words + " y " + centsText;
        if (number < 0)
            result = "menos " + result;
        return result;
    }

    private string ThreeDigit(int number)
    {
        int hundreds = number / 100;
        int rest = number % 100;
        string hundredWord = hundreds switch {
            1 => rest == 0 ? "cien" : "ciento",
            5 => "quinientos",
            7 => "setecientos",
            9 => "novecientos",
            _ => Small[hundreds] + "cientos"
        };
        if (rest == 0)
            return hundredWord;
        return hundredWord + " " + ToCardinal(rest);
    }

    private static readonly string[] Small =
    {
        "cero","uno","dos","tres","cuatro","cinco","seis","siete","ocho","nueve","diez",
        "once","doce","trece","catorce","quince","dieciséis","diecisiete","dieciocho","diecinueve","veinte",
        "veintiuno","veintidós","veintitrés","veinticuatro","veinticinco","veintiséis","veintisiete","veintiocho","veintinueve"
    };

    private static readonly string[] Tens =
    {
        "", "", "veinte", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa"
    };
}
