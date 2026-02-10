using System.Globalization;

namespace DocxportNet.Formatting.Impl;

public sealed class PortugueseNumberToWordsProvider : DxpINumberToWordsProvider
{
    public bool CanHandle(CultureInfo culture) => culture.TwoLetterISOLanguageName.Equals("pt", StringComparison.OrdinalIgnoreCase);

    public string ToCardinal(int number)
    {
        if (number == 0)
            return "zero";
        if (number < 0)
            return "menos " + ToCardinal(Math.Abs(number));
        if (number > 32767)
            return number.ToString(CultureInfo.InvariantCulture);

        if (number < 20)
            return Small[number];
        if (number < 100)
        {
            int tens = number / 10;
            int ones = number % 10;
            if (ones == 0)
                return Tens[tens];
            return Tens[tens] + " e " + Small[ones];
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
            return "zero";
        if (number < 0)
            return "menos " + ToOrdinalWords(Math.Abs(number));
        if (number > 32767)
            return number.ToString(CultureInfo.InvariantCulture);

        if (number == 1)
            return "primeiro";
        if (number == 2)
            return "segundo";
        if (number == 3)
            return "terceiro";
        if (number == 4)
            return "quarto";
        if (number == 5)
            return "quinto";
        if (number == 6)
            return "sexto";
        if (number == 7)
            return "sétimo";
        if (number == 8)
            return "oitavo";
        if (number == 9)
            return "nono";
        if (number == 10)
            return "décimo";

        return ToCardinal(number) + "º";
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
        string result = words + " e " + centsText;
        if (number < 0)
            result = "menos " + result;
        return result;
    }

    private string ThreeDigit(int number)
    {
        int hundreds = number / 100;
        int rest = number % 100;
        string hundredWord = hundreds switch {
            1 => rest == 0 ? "cem" : "cento",
            2 => "duzentos",
            3 => "trezentos",
            4 => "quatrocentos",
            5 => "quinhentos",
            6 => "seiscentos",
            7 => "setecentos",
            8 => "oitocentos",
            9 => "novecentos",
            _ => Small[hundreds] + "centos"
        };
        if (rest == 0)
            return hundredWord;
        return hundredWord + " e " + ToCardinal(rest);
    }

    private static readonly string[] Small =
    {
        "zero","um","dois","três","quatro","cinco","seis","sete","oito","nove","dez",
        "onze","doze","treze","catorze","quinze","dezesseis","dezessete","dezoito","dezenove"
    };

    private static readonly string[] Tens =
    {
        "", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa"
    };
}
