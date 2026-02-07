using System.Globalization;

namespace DocxportNet.Formatting.Impl;

public sealed class FinnishNumberToWordsProvider : IDxpNumberToWordsProvider
{
	public bool CanHandle(CultureInfo culture) => culture.TwoLetterISOLanguageName.Equals("fi", StringComparison.OrdinalIgnoreCase);

	public string ToCardinal(int number)
	{
		if (number == 0)
			return "nolla";
		if (number < 0)
			return "miinus " + ToCardinal(Math.Abs(number));
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
		string thousandWord = thousands == 1 ? "tuhat" : ToCardinal(thousands) + "tuhatta";
		return rest == 0 ? thousandWord : thousandWord + ToCardinal(rest);
	}

	public string ToOrdinalWords(int number)
	{
		if (number == 0)
			return "nollas";
		if (number < 0)
			return "miinus " + ToOrdinalWords(Math.Abs(number));
		if (number > 32767)
			return number.ToString(CultureInfo.InvariantCulture);

		if (number < 20)
			return OrdinalSmall[number];
		if (number < 100)
			return TwoDigitOrdinal(number);
		if (number < 1000)
			return ThreeDigitOrdinal(number);

		int thousands = number / 1000;
		int rest = number % 1000;
		string thousandWord = thousands == 1 ? "tuhannes" : ToCardinal(thousands) + "tuhannes";
		return rest == 0 ? thousandWord : thousandWord + ToOrdinalWords(rest);
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
		string result = words + " ja " + centsText;
		if (number < 0)
			result = "miinus " + result;
		return result;
	}

	private static string TwoDigit(int number)
	{
		int tens = number / 10;
		int ones = number % 10;
		if (ones == 0)
			return Tens[tens];
		return Tens[tens] + Small[ones];
	}

	private string ThreeDigit(int number)
	{
		int hundreds = number / 100;
		int rest = number % 100;
		string hundredWord = hundreds == 1 ? "sata" : Small[hundreds] + "sataa";
		if (rest == 0)
			return hundredWord;
		return hundredWord + ToCardinal(rest);
	}

	private static string TwoDigitOrdinal(int number)
	{
		int tens = number / 10;
		int ones = number % 10;
		if (ones == 0)
			return OrdinalTens[tens];
		return OrdinalTens[tens] + OrdinalSmall[ones];
	}

	private string ThreeDigitOrdinal(int number)
	{
		int hundreds = number / 100;
		int rest = number % 100;
		string hundredWord = OrdinalHundreds[hundreds];
		if (rest == 0)
			return hundredWord;
		return hundredWord + ToOrdinalWords(rest);
	}

	private static readonly string[] Small =
	{
		"nolla","yksi","kaksi","kolme","neljä","viisi","kuusi","seitsemän","kahdeksan","yhdeksän","kymmenen",
		"yksitoista","kaksitoista","kolmetoista","neljätoista","viisitoista","kuusitoista","seitsemäntoista","kahdeksantoista","yhdeksäntoista"
	};

	private static readonly string[] Tens =
	{
		"", "", "kaksikymmentä", "kolmekymmentä", "neljäkymmentä", "viisikymmentä", "kuusikymmentä", "seitsemänkymmentä", "kahdeksankymmentä", "yhdeksänkymmentä"
	};

	private static readonly string[] OrdinalSmall =
	{
		"nollas","ensimmäinen","toinen","kolmas","neljäs","viides","kuudes","seitsemäs","kahdeksas","yhdeksäs","kymmenes",
		"yhdestoista","kahdestoista","kolmastoista","neljästoista","viidestoista","kuudestoista","seitsemästoista","kahdeksastoista","yhdeksästoista"
	};

	private static readonly string[] OrdinalTens =
	{
		"", "", "kahdeskymmenes", "kolmaskymmenes", "neljäskymmenes", "viideskymmenes", "kuudeskymmenes", "seitsemäskymmenes", "kahdeksaskymmenes", "yhdeksäskymmenes"
	};

	private static readonly string[] OrdinalHundreds =
	{
		"", "sadas", "kahdessadas", "kolmassadas", "neljässadas", "viidessadas", "kuudessadas", "seitsemässadas", "kahdeksassadas", "yhdeksässadas"
	};
}
