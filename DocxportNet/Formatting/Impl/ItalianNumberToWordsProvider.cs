using System.Globalization;

namespace DocxportNet.Formatting.Impl;

public sealed class ItalianNumberToWordsProvider : IDxpNumberToWordsProvider
{
	public bool CanHandle(CultureInfo culture) => culture.TwoLetterISOLanguageName.Equals("it", StringComparison.OrdinalIgnoreCase);

	public string ToCardinal(int number)
	{
		if (number == 0)
			return "zero";
		if (number < 0)
			return "meno " + ToCardinal(Math.Abs(number));
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
		string thousandWord = thousands == 1 ? "mille" : ToCardinal(thousands) + "mila";
		return rest == 0 ? thousandWord : thousandWord + ToCardinal(rest);
	}

	public string ToOrdinalWords(int number)
	{
		if (number == 0)
			return "zero";
		if (number < 0)
			return "meno " + ToOrdinalWords(Math.Abs(number));
		if (number > 32767)
			return number.ToString(CultureInfo.InvariantCulture);

		if (number == 1)
			return "primo";
		if (number == 2)
			return "secondo";
		if (number == 3)
			return "terzo";
		if (number == 4)
			return "quarto";
		if (number == 5)
			return "quinto";
		if (number == 6)
			return "sesto";
		if (number == 7)
			return "settimo";
		if (number == 8)
			return "ottavo";
		if (number == 9)
			return "nono";
		if (number == 10)
			return "decimo";

		string baseCard = ToCardinal(number);
		if (baseCard.EndsWith("e", StringComparison.Ordinal))
			baseCard = baseCard.Substring(0, baseCard.Length - 1);
		if (baseCard.EndsWith("a", StringComparison.Ordinal))
			baseCard = baseCard.Substring(0, baseCard.Length - 1);
		return baseCard + "esimo";
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
			result = "meno " + result;
		return result;
	}

	private static string TwoDigit(int number)
	{
		int tens = number / 10;
		int ones = number % 10;
		if (ones == 0)
			return Tens[tens];
		string tensWord = Tens[tens];
		if (ones == 1 || ones == 8)
			tensWord = tensWord.Substring(0, tensWord.Length - 1);
		return tensWord + Small[ones];
	}

	private string ThreeDigit(int number)
	{
		int hundreds = number / 100;
		int rest = number % 100;
		string hundredWord = hundreds == 1 ? "cento" : Small[hundreds] + "cento";
		if (rest == 0)
			return hundredWord;
		return hundredWord + ToCardinal(rest);
	}

	private static readonly string[] Small =
	{
		"zero","uno","due","tre","quattro","cinque","sei","sette","otto","nove","dieci",
		"undici","dodici","tredici","quattordici","quindici","sedici","diciassette","diciotto","diciannove"
	};

	private static readonly string[] Tens =
	{
		"", "", "venti", "trenta", "quaranta", "cinquanta", "sessanta", "settanta", "ottanta", "novanta"
	};
}
