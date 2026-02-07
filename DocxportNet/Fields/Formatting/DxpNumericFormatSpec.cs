using System.Globalization;
using System.Text;

namespace DocxportNet.Fields.Formatting;

public enum DxpNumericTokenKind
{
	DigitZero,
	DigitOptional,
	DropDigit,
	DecimalPoint,
	GroupingSeparator,
	MinusSign,
	PlusSign,
	Percent,
	Literal,
	NumberedItem
}

public sealed record DxpNumericToken(DxpNumericTokenKind Kind, string Text);

public sealed class DxpNumericFormatSection
{
	public IReadOnlyList<DxpNumericToken> Tokens { get; }

	public DxpNumericFormatSection(IReadOnlyList<DxpNumericToken> tokens)
	{
		Tokens = tokens;
	}
}

public sealed class DxpNumericFormatSpec : IDxpFieldFormatSpec
{
	public string RawText { get; }
	public DxpNumericFormatSection Positive { get; }
	public DxpNumericFormatSection? Negative { get; }
	public DxpNumericFormatSection? Zero { get; }

	public DxpNumericFormatSpec(string rawText, DxpNumericFormatSection positive, DxpNumericFormatSection? negative, DxpNumericFormatSection? zero)
	{
		RawText = rawText;
		Positive = positive;
		Negative = negative;
		Zero = zero;
	}

	public string Apply(string text, DxpFieldValue value, DxpFieldEvalContext context)
	{
		if (value.Kind != DxpFieldValueKind.Number || value.NumberValue == null)
			return text;

		double number = value.NumberValue.Value;
		if (number > 0)
			return RenderSection(Positive, number, context);
		if (number < 0 && Negative != null)
			return RenderSection(Negative, number, context);
		if (number == 0 && Zero != null)
			return RenderSection(Zero, number, context);
		if (number < 0)
			return RenderSection(Positive, number, context);
		return RenderSection(Positive, number, context);
	}

	private static string RenderSection(DxpNumericFormatSection section, double number, DxpFieldEvalContext context)
	{
		var culture = context.Culture ?? CultureInfo.CurrentCulture;
		if (section.Tokens.Count == 0)
			return number.ToString(culture);

		SplitTokens(section.Tokens, out var leftTokens, out var rightTokens, out bool hasDecimal);
		bool isNegative = number < 0;
		double abs = Math.Abs(number);

		int leftRequired = CountRequired(leftTokens);
		int leftTotal = CountTotalPlaceholders(leftTokens);
		int rightRequired = CountRequired(rightTokens);
		int rightTotal = CountTotalPlaceholders(rightTokens);

		double rounded = rightTotal > 0 ? Math.Round(abs, rightTotal, MidpointRounding.AwayFromZero) : Math.Round(abs, 0, MidpointRounding.AwayFromZero);
		long intPart = (long)Math.Floor(rounded);
		double fracPart = rounded - intPart;

		string intDigits = intPart.ToString(CultureInfo.InvariantCulture);
		int dropCount = CountDropDigits(leftTokens, out int dropIndex);
		bool hasDrop = dropCount > 0;
		string extraPrefix = string.Empty;
		string placeholderDigits = string.Empty;
		if (leftTotal > 0)
		{
			if (intDigits.Length > leftTotal)
			{
				if (!hasDrop)
					extraPrefix = intDigits.Substring(0, intDigits.Length - leftTotal);
				placeholderDigits = intDigits.Substring(intDigits.Length - leftTotal);
			}
			else
			{
				placeholderDigits = intDigits;
			}

			if (leftRequired > placeholderDigits.Length)
				placeholderDigits = placeholderDigits.PadLeft(leftRequired, '0');

			if (hasDrop && dropIndex > 0)
			{
				placeholderDigits = dropIndex >= placeholderDigits.Length
					? string.Empty
					: placeholderDigits.Substring(dropIndex);
			}
		}

		string fracDigits = rightTotal > 0 ? ((long)Math.Round(fracPart * Math.Pow(10, rightTotal), 0, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture).PadLeft(rightTotal, '0') : string.Empty;
		if (rightTotal > 0 && fracDigits.Length > rightTotal)
			fracDigits = fracDigits.Substring(fracDigits.Length - rightTotal);

		int trimIndex = fracDigits.Length;
		while (trimIndex > rightRequired && trimIndex > 0 && fracDigits[trimIndex - 1] == '0')
			trimIndex--;
		if (trimIndex < fracDigits.Length)
			fracDigits = fracDigits.Substring(0, trimIndex);

		var leftBuilder = new StringBuilder();
		int? groupSize = GetGroupingSize(leftTokens);
		AppendLeft(leftBuilder, leftTokens, extraPrefix, placeholderDigits, isNegative, number, culture, context);
		string leftText = groupSize.HasValue
			? ApplyGroupingToLeftText(leftBuilder.ToString(), groupSize.Value, culture)
			: leftBuilder.ToString();

		bool hasRightContent = rightTokens.Count > 0 && (fracDigits.Length > 0 || HasRequiredFraction(rightTokens) || HasRightLiterals(rightTokens));
		if (hasDecimal && hasRightContent)
		{
			var sb = new StringBuilder();
			sb.Append(leftText);
			sb.Append(culture.NumberFormat.NumberDecimalSeparator);
			AppendRight(sb, rightTokens, fracDigits, culture, context);
			return sb.ToString();
		}

		return leftText;
	}

	private static void SplitTokens(IReadOnlyList<DxpNumericToken> tokens, out List<DxpNumericToken> left, out List<DxpNumericToken> right, out bool hasDecimal)
	{
		left = new List<DxpNumericToken>();
		right = new List<DxpNumericToken>();
		hasDecimal = false;
		bool seenDecimal = false;
		foreach (var token in tokens)
		{
			if (!seenDecimal && token.Kind == DxpNumericTokenKind.DecimalPoint)
			{
				seenDecimal = true;
				hasDecimal = true;
				continue;
			}
			if (seenDecimal)
				right.Add(token);
			else
				left.Add(token);
		}
	}

	private static int CountRequired(IReadOnlyList<DxpNumericToken> tokens)
	{
		int count = 0;
		foreach (var token in tokens)
		{
			if (token.Kind == DxpNumericTokenKind.DigitZero || token.Kind == DxpNumericTokenKind.DropDigit)
				count++;
		}
		return count;
	}

	private static int CountTotalPlaceholders(IReadOnlyList<DxpNumericToken> tokens)
	{
		int count = 0;
		foreach (var token in tokens)
		{
			if (token.Kind == DxpNumericTokenKind.DigitZero || token.Kind == DxpNumericTokenKind.DigitOptional || token.Kind == DxpNumericTokenKind.DropDigit)
				count++;
		}
		return count;
	}

	private static bool HasRequiredFraction(IReadOnlyList<DxpNumericToken> tokens)
	{
		foreach (var token in tokens)
		{
			if (token.Kind == DxpNumericTokenKind.DigitZero || token.Kind == DxpNumericTokenKind.DropDigit)
				return true;
		}
		return false;
	}

	private static bool HasRightLiterals(IReadOnlyList<DxpNumericToken> tokens)
	{
		foreach (var token in tokens)
		{
			if (token.Kind == DxpNumericTokenKind.Literal || token.Kind == DxpNumericTokenKind.NumberedItem ||
				token.Kind == DxpNumericTokenKind.Percent || token.Kind == DxpNumericTokenKind.MinusSign || token.Kind == DxpNumericTokenKind.PlusSign)
				return true;
		}
		return false;
	}

	private static int CountDropDigits(IReadOnlyList<DxpNumericToken> tokens, out int firstDropIndex)
	{
		int count = 0;
		firstDropIndex = -1;
		int placeholderIndex = 0;
		foreach (var token in tokens)
		{
			if (token.Kind == DxpNumericTokenKind.DigitZero ||
				token.Kind == DxpNumericTokenKind.DigitOptional ||
				token.Kind == DxpNumericTokenKind.DropDigit)
			{
				if (token.Kind == DxpNumericTokenKind.DropDigit)
				{
					count++;
					if (firstDropIndex < 0)
						firstDropIndex = placeholderIndex;
				}
				placeholderIndex++;
			}
		}
		if (firstDropIndex < 0)
			firstDropIndex = 0;
		return count;
	}

	private static bool HasGroupingSeparator(IReadOnlyList<DxpNumericToken> tokens)
	{
		foreach (var token in tokens)
		{
			if (token.Kind == DxpNumericTokenKind.GroupingSeparator)
				return true;
		}
		return false;
	}

	private static int? GetGroupingSize(IReadOnlyList<DxpNumericToken> tokens)
	{
		int lastSeparator = -1;
		int index = 0;
		for (int i = 0; i < tokens.Count; i++)
		{
			if (tokens[i].Kind == DxpNumericTokenKind.GroupingSeparator)
				lastSeparator = index;
			if (tokens[i].Kind == DxpNumericTokenKind.DigitZero ||
				tokens[i].Kind == DxpNumericTokenKind.DigitOptional ||
				tokens[i].Kind == DxpNumericTokenKind.DropDigit)
				index++;
		}
		if (lastSeparator < 0)
			return null;
		int total = CountTotalPlaceholders(tokens);
		int size = total - lastSeparator;
		return size > 0 ? size : (int?)null;
	}

	private static void AppendLeft(StringBuilder sb, IReadOnlyList<DxpNumericToken> tokens, string extraPrefix, string digits, bool isNegative, double number, CultureInfo culture, DxpFieldEvalContext context)
	{
		int totalPlaceholders = CountTotalPlaceholders(tokens);
		int leadingBlanks = Math.Max(0, totalPlaceholders - digits.Length);
		int placeholderIndex = 0;
		bool extraEmitted = string.IsNullOrEmpty(extraPrefix);
		char space = ' ';
		for (int i = 0; i < tokens.Count; i++)
		{
			var token = tokens[i];
			switch (token.Kind)
			{
				case DxpNumericTokenKind.DigitZero:
				case DxpNumericTokenKind.DigitOptional:
				case DxpNumericTokenKind.DropDigit:
				{
					if (!extraEmitted)
					{
						sb.Append(extraPrefix);
						extraEmitted = true;
					}

					int digitIndex = placeholderIndex - leadingBlanks;
					placeholderIndex++;
					if (digitIndex >= 0 && digitIndex < digits.Length)
						sb.Append(digits[digitIndex]);
					else if (token.Kind == DxpNumericTokenKind.DigitZero || token.Kind == DxpNumericTokenKind.DropDigit)
						sb.Append('0');
					else
						sb.Append(space);
					break;
				}
				case DxpNumericTokenKind.GroupingSeparator:
					break;
				case DxpNumericTokenKind.MinusSign:
					sb.Append(isNegative ? '-' : space);
					break;
				case DxpNumericTokenKind.PlusSign:
					sb.Append(number > 0 ? '+' : number < 0 ? '-' : space);
					break;
				case DxpNumericTokenKind.Percent:
					sb.Append('%');
					break;
				case DxpNumericTokenKind.Literal:
					sb.Append(token.Text);
					break;
				case DxpNumericTokenKind.NumberedItem:
					sb.Append(ResolveNumberedItem(context, token.Text));
					break;
				case DxpNumericTokenKind.DecimalPoint:
					break;
			}
		}

		if (!extraEmitted)
			sb.Append(extraPrefix);
	}

	private static string ApplyGroupingToLeftText(string text, int groupSize, CultureInfo culture)
	{
		if (groupSize <= 0)
			return text;

		var sb = new StringBuilder(text);
		int digitCount = 0;
		for (int i = sb.Length - 1; i >= 0; i--)
		{
			char ch = sb[i];
			if (char.IsDigit(ch))
			{
				digitCount++;
				if (digitCount % groupSize == 0 && i > 0)
				{
					int j = i - 1;
					while (j >= 0 && !char.IsDigit(sb[j]))
						j--;
					if (j >= 0)
						sb.Insert(i, culture.NumberFormat.NumberGroupSeparator);
				}
			}
			else
			{
				// non-digit: do not count toward grouping
			}
		}
		return sb.ToString();
	}

	private static void AppendRight(StringBuilder sb, IReadOnlyList<DxpNumericToken> tokens, string digits, CultureInfo culture, DxpFieldEvalContext context)
	{
		int digitIndex = 0;
		foreach (var token in tokens)
		{
			switch (token.Kind)
			{
				case DxpNumericTokenKind.DigitZero:
				case DxpNumericTokenKind.DigitOptional:
				case DxpNumericTokenKind.DropDigit:
				{
					if (digitIndex < digits.Length)
						sb.Append(digits[digitIndex++]);
					else if (token.Kind == DxpNumericTokenKind.DigitZero || token.Kind == DxpNumericTokenKind.DropDigit)
						sb.Append('0');
					break;
				}
				case DxpNumericTokenKind.GroupingSeparator:
					sb.Append(culture.NumberFormat.NumberGroupSeparator);
					break;
				case DxpNumericTokenKind.MinusSign:
					sb.Append('-');
					break;
				case DxpNumericTokenKind.PlusSign:
					sb.Append('+');
					break;
				case DxpNumericTokenKind.Percent:
					sb.Append('%');
					break;
				case DxpNumericTokenKind.Literal:
					sb.Append(token.Text);
					break;
				case DxpNumericTokenKind.NumberedItem:
					sb.Append(ResolveNumberedItem(context, token.Text));
					break;
				case DxpNumericTokenKind.DecimalPoint:
					break;
			}
		}
	}

	private static string ResolveNumberedItem(DxpFieldEvalContext context, string name)
	{
		if (context.TryGetNumberedItem(name, out var value))
			return value ?? string.Empty;
		return string.Empty;
	}
}
