using System.Globalization;
namespace DocxportNet.Formatting.Impl;

public sealed class EnglishNumberToWordsProvider : IDxpNumberToWordsProvider
{
	public bool CanHandle(CultureInfo culture) => culture.TwoLetterISOLanguageName.Equals("en", StringComparison.OrdinalIgnoreCase);

	public string ToCardinal(int number)
	{
		if (number == 0)
			return "zero";
		if (number < 0)
			return "minus " + ToCardinal(Math.Abs(number));
		if (number > 32767)
			return number.ToString(CultureInfo.InvariantCulture);

		if (number < 20)
			return SmallNumbers[number];
		if (number < 100)
		{
			int tens = number / 10;
			int ones = number % 10;
			return ones == 0 ? TensNumbers[tens] : TensNumbers[tens] + "-" + SmallNumbers[ones];
		}
		if (number < 1000)
		{
			int hundreds = number / 100;
			int rest = number % 100;
			return rest == 0 ? SmallNumbers[hundreds] + " hundred" : SmallNumbers[hundreds] + " hundred " + ToCardinal(rest);
		}
		if (number < 1000000)
		{
			int thousands = number / 1000;
			int rest = number % 1000;
			return rest == 0 ? ToCardinal(thousands) + " thousand" : ToCardinal(thousands) + " thousand " + ToCardinal(rest);
		}
		int millions = number / 1000000;
		int remainder = number % 1000000;
		return remainder == 0 ? ToCardinal(millions) + " million" : ToCardinal(millions) + " million " + ToCardinal(remainder);
	}

	public string ToOrdinalWords(int number)
	{
		if (number == 0)
			return "zeroth";
		if (number < 0)
			return "minus " + ToOrdinalWords(Math.Abs(number));
		if (number > 32767)
			return number.ToString(CultureInfo.InvariantCulture);

		if (number < 20)
			return OrdinalSmall[number];
		if (number < 100)
		{
			int tens = number / 10;
			int ones = number % 10;
			return ones == 0 ? OrdinalTens[tens] : TensNumbers[tens] + "-" + OrdinalSmall[ones];
		}
		if (number < 1000)
		{
			int hundreds = number / 100;
			int rest = number % 100;
			return rest == 0 ? SmallNumbers[hundreds] + " hundredth" : SmallNumbers[hundreds] + " hundred " + ToOrdinalWords(rest);
		}
		if (number < 1000000)
		{
			int thousands = number / 1000;
			int rest = number % 1000;
			return rest == 0 ? ToCardinal(thousands) + " thousandth" : ToCardinal(thousands) + " thousand " + ToOrdinalWords(rest);
		}
		int millions = number / 1000000;
		int remainder = number % 1000000;
		return remainder == 0 ? ToCardinal(millions) + " millionth" : ToCardinal(millions) + " million " + ToOrdinalWords(remainder);
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
		string result = words + " and " + centsText;
		if (number < 0)
			result = "minus " + result;
		return result;
	}

	private static readonly string[] SmallNumbers =
	{
		"zero","one","two","three","four","five","six","seven","eight","nine","ten",
		"eleven","twelve","thirteen","fourteen","fifteen","sixteen","seventeen","eighteen","nineteen"
	};

	private static readonly string[] TensNumbers =
	{
		"", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"
	};

	private static readonly string[] OrdinalSmall =
	{
		"zeroth","first","second","third","fourth","fifth","sixth","seventh","eighth","ninth","tenth",
		"eleventh","twelfth","thirteenth","fourteenth","fifteenth","sixteenth","seventeenth","eighteenth","nineteenth"
	};

	private static readonly string[] OrdinalTens =
	{
		"", "", "twentieth", "thirtieth", "fortieth", "fiftieth", "sixtieth", "seventieth", "eightieth", "ninetieth"
	};
}
