using System.Globalization;

namespace DocxportNet.Formatting.Impl;

public sealed class GreekNumberToWordsProvider : DxpINumberToWordsProvider
{
    public bool CanHandle(CultureInfo culture) => culture.TwoLetterISOLanguageName.Equals("el", StringComparison.OrdinalIgnoreCase);

    public string ToCardinal(int number)
    {
        if (number == 0)
            return "μηδέν";
        if (number < 0)
            return "μείον " + ToCardinal(Math.Abs(number));
        if (number > 32767)
            return number.ToString(CultureInfo.InvariantCulture);

        if (number < 20)
            return Small[number];
        if (number < 100)
        {
            int tens = number / 10;
            int ones = number % 10;
            return ones == 0 ? Tens[tens] : Tens[tens] + " " + Small[ones];
        }
        if (number < 1000)
        {
            int hundreds = number / 100;
            int rest = number % 100;
            string hundredWord = hundreds == 1 && rest > 0 ? "εκατόν" : Hundreds[hundreds];
            return rest == 0 ? hundredWord : hundredWord + " " + ToCardinal(rest);
        }
        if (number < 1000000)
        {
            int thousands = number / 1000;
            int rest = number % 1000;
            string thousandWord = thousands == 1 ? "χίλια" : ToCardinal(thousands) + " χιλιάδες";
            return rest == 0 ? thousandWord : thousandWord + " " + ToCardinal(rest);
        }
        int millions = number / 1000000;
        int remainder = number % 1000000;
        string millionWord = millions == 1 ? "ένα εκατομμύριο" : ToCardinal(millions) + " εκατομμύρια";
        return remainder == 0 ? millionWord : millionWord + " " + ToCardinal(remainder);
    }

    public string ToOrdinalWords(int number)
    {
        if (number == 0)
            return "μηδενικός";
        if (number < 0)
            return "μείον " + ToOrdinalWords(Math.Abs(number));
        if (number > 32767)
            return number.ToString(CultureInfo.InvariantCulture);

        if (number < 20)
            return OrdinalSmall[number];
        if (number < 100)
        {
            int tens = number / 10;
            int ones = number % 10;
            return ones == 0 ? OrdinalTens[tens] : OrdinalTens[tens] + " " + OrdinalSmall[ones];
        }
        if (number < 1000)
        {
            int hundreds = number / 100;
            int rest = number % 100;
            if (rest == 0)
                return OrdinalHundreds[hundreds];
            string hundredWord = hundreds == 1 ? "εκατόν" : Hundreds[hundreds];
            return hundredWord + " " + ToOrdinalWords(rest);
        }
        if (number < 1000000)
        {
            int thousands = number / 1000;
            int rest = number % 1000;
            if (rest == 0)
                return thousands == 1 ? "χιλιοστός" : ToCardinal(thousands) + " χιλιοστός";
            string thousandWord = thousands == 1 ? "χίλια" : ToCardinal(thousands) + " χιλιάδες";
            return thousandWord + " " + ToOrdinalWords(rest);
        }
        int millions = number / 1000000;
        int remainder = number % 1000000;
        if (remainder == 0)
            return millions == 1 ? "εκατομμυριοστός" : ToCardinal(millions) + " εκατομμυριοστός";
        string millionWord = millions == 1 ? "ένα εκατομμύριο" : ToCardinal(millions) + " εκατομμύρια";
        return millionWord + " " + ToOrdinalWords(remainder);
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
        string result = words + " και " + centsText;
        if (number < 0)
            result = "μείον " + result;
        return result;
    }

    private static readonly string[] Small =
    {
        "μηδέν","ένα","δύο","τρία","τέσσερα","πέντε","έξι","επτά","οκτώ","εννέα","δέκα",
        "έντεκα","δώδεκα","δεκατρία","δεκατέσσερα","δεκαπέντε","δεκαέξι","δεκαεπτά","δεκαοκτώ","δεκαεννέα"
    };

    private static readonly string[] Tens =
    {
        "", "", "είκοσι", "τριάντα", "σαράντα", "πενήντα", "εξήντα", "εβδομήντα", "ογδόντα", "ενενήντα"
    };

    private static readonly string[] Hundreds =
    {
        "", "εκατό", "διακόσια", "τριακόσια", "τετρακόσια", "πεντακόσια", "εξακόσια", "επτακόσια", "οκτακόσια", "εννιακόσια"
    };

    private static readonly string[] OrdinalSmall =
    {
        "μηδενικός","πρώτος","δεύτερος","τρίτος","τέταρτος","πέμπτος","έκτος","έβδομος","όγδοος","ένατος","δέκατος",
        "ενδέκατος","δωδέκατος","δέκατος τρίτος","δέκατος τέταρτος","δέκατος πέμπτος","δέκατος έκτος","δέκατος έβδομος","δέκατος όγδοος","δέκατος ένατος"
    };

    private static readonly string[] OrdinalTens =
    {
        "", "", "εικοστός", "τριακοστός", "τεσσαρακοστός", "πεντηκοστός", "εξηκοστός", "εβδομηκοστός", "ογδοηκοστός", "ενενηκοστός"
    };

    private static readonly string[] OrdinalHundreds =
    {
        "", "εκατοστός", "διακοσιοστός", "τριακοσιοστός", "τετρακοσιοστός", "πεντακοσιοστός", "εξακοσιοστός", "επτακοσιοστός", "οκτακοσιοστός", "εννιακοσιοστός"
    };
}
