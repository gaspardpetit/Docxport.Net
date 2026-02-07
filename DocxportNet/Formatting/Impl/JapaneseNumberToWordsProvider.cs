using System.Globalization;
using System.Text;

namespace DocxportNet.Formatting.Impl;

public sealed class JapaneseNumberToWordsProvider : IDxpNumberToWordsProvider
{
	public bool CanHandle(CultureInfo culture) => culture.TwoLetterISOLanguageName.Equals("ja", StringComparison.OrdinalIgnoreCase);

	public string ToCardinal(int number)
	{
		if (number == 0)
			return "〇";
		if (number < 0)
			return "マイナス " + ToCardinal(Math.Abs(number));
		if (number > 32767)
			return number.ToString(CultureInfo.InvariantCulture);

		return ToJapaneseNumber(number);
	}

	public string ToOrdinalWords(int number)
	{
		if (number == 0)
			return "第〇";
		if (number < 0)
			return "第マイナス " + ToCardinal(Math.Abs(number));
		if (number > 32767)
			return number.ToString(CultureInfo.InvariantCulture);

		return "第" + ToCardinal(number);
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
		string result = words + " と " + centsText;
		if (number < 0)
			result = "マイナス " + result;
		return result;
	}

	private static string ToJapaneseNumber(int number)
	{
		var sb = new StringBuilder();
		int man = number / 10000;
		int rest = number % 10000;
		if (man > 0)
		{
			if (man > 1)
				sb.Append(ToJapaneseBelow10000(man));
			sb.Append("万");
		}
		if (rest > 0)
			sb.Append(ToJapaneseBelow10000(rest));
		return sb.ToString();
	}

	private static string ToJapaneseBelow10000(int number)
	{
		var sb = new StringBuilder();
		int thousands = number / 1000;
		int rest = number % 1000;
		int hundreds = rest / 100;
		rest %= 100;
		int tens = rest / 10;
		int ones = rest % 10;

		if (thousands > 0)
		{
			if (thousands > 1)
				sb.Append(Digits[thousands]);
			sb.Append("千");
		}
		if (hundreds > 0)
		{
			if (hundreds > 1)
				sb.Append(Digits[hundreds]);
			sb.Append("百");
		}
		if (tens > 0)
		{
			if (tens > 1)
				sb.Append(Digits[tens]);
			sb.Append("十");
		}
		if (ones > 0)
			sb.Append(Digits[ones]);

		return sb.ToString();
	}

	private static readonly string[] Digits =
	{
		"〇","一","二","三","四","五","六","七","八","九"
	};
}
