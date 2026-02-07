using System.Globalization;
using System.Text;

namespace DocxportNet.Formatting.Impl;

public sealed class ChineseSimplifiedNumberToWordsProvider : IDxpNumberToWordsProvider
{
	public bool CanHandle(CultureInfo culture) => culture.Name.StartsWith("zh", StringComparison.OrdinalIgnoreCase);

	public string ToCardinal(int number)
	{
		if (number == 0)
			return "零";
		if (number < 0)
			return "负" + ToCardinal(Math.Abs(number));
		if (number > 32767)
			return number.ToString(CultureInfo.InvariantCulture);

		return ToChinese(number);
	}

	public string ToOrdinalWords(int number)
	{
		if (number == 0)
			return "第零";
		if (number < 0)
			return "第负" + ToCardinal(Math.Abs(number));
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
		string result = words + "和" + centsText;
		if (number < 0)
			result = "负" + result;
		return result;
	}

	private static string ToChinese(int number)
	{
		int wan = number / 10000;
		int rest = number % 10000;
		var sb = new StringBuilder();
		if (wan > 0)
		{
			sb.Append(ToChineseBelow10000(wan));
			sb.Append("万");
			if (rest > 0 && rest < 1000)
				sb.Append("零");
		}
		if (rest > 0)
			sb.Append(ToChineseBelow10000(rest));
		return sb.ToString();
	}

	private static string ToChineseBelow10000(int number)
	{
		int thousands = number / 1000;
		int hundreds = number / 100 % 10;
		int tens = number / 10 % 10;
		int ones = number % 10;
		var sb = new StringBuilder();

		if (thousands > 0)
		{
			sb.Append(Digits[thousands]).Append("千");
		}
		if (hundreds > 0)
		{
			sb.Append(Digits[hundreds]).Append("百");
		}
		else if (thousands > 0 && (tens > 0 || ones > 0))
		{
			sb.Append("零");
		}

		if (tens > 0)
		{
			if (tens == 1 && thousands == 0 && hundreds == 0)
				sb.Append("十");
			else
				sb.Append(Digits[tens]).Append("十");
		}
		else if ((hundreds > 0 || thousands > 0) && ones > 0)
		{
			sb.Append("零");
		}

		if (ones > 0)
			sb.Append(Digits[ones]);

		return sb.ToString();
	}

	private static readonly string[] Digits =
	{
		"零","一","二","三","四","五","六","七","八","九"
	};
}
