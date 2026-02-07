namespace DocxportNet.Fields;

public enum DxpFieldValueKind
{
	String,
	Number,
	DateTime,
}

public readonly struct DxpFieldValue
{
	public DxpFieldValueKind Kind { get; }
	public string? StringValue { get; }
	public double? NumberValue { get; }
	public DateTimeOffset? DateTimeValue { get; }

	public DxpFieldValue(string value)
	{
		Kind = DxpFieldValueKind.String;
		StringValue = value;
		NumberValue = null;
		DateTimeValue = null;
	}

	public DxpFieldValue(double value)
	{
		Kind = DxpFieldValueKind.Number;
		StringValue = null;
		NumberValue = value;
		DateTimeValue = null;
	}

	public DxpFieldValue(DateTimeOffset value)
	{
		Kind = DxpFieldValueKind.DateTime;
		StringValue = null;
		NumberValue = null;
		DateTimeValue = value;
	}

	public bool TryConvertToKind(DxpFieldValueKind targetKind, DxpFieldEvalContext context, out DxpFieldValue converted)
	{
		converted = this;
		if (Kind == targetKind)
			return true;

		switch (targetKind)
		{
			case DxpFieldValueKind.Number:
			{
				if (NumberValue.HasValue)
				{
					converted = new DxpFieldValue(NumberValue.Value);
					return true;
				}
				if (!string.IsNullOrWhiteSpace(StringValue) &&
					TryParseNumber(StringValue!, context, out var number))
				{
					converted = new DxpFieldValue(number);
					return true;
				}
				return false;
			}
			case DxpFieldValueKind.DateTime:
			{
				if (DateTimeValue.HasValue)
				{
					converted = new DxpFieldValue(DateTimeValue.Value);
					return true;
				}
				if (!string.IsNullOrWhiteSpace(StringValue) &&
					TryParseDateTime(StringValue!, context, out var dateTime))
				{
					converted = new DxpFieldValue(dateTime);
					return true;
				}
				return false;
			}
			case DxpFieldValueKind.String:
			default:
				return false;
		}
	}

	private static bool TryParseNumber(string text, DxpFieldEvalContext context, out double number)
	{
		var culture = context.Culture ?? System.Globalization.CultureInfo.CurrentCulture;
		if (double.TryParse(text, System.Globalization.NumberStyles.Any, culture, out number))
			return true;

		if (context.AllowInvariantNumericFallback &&
			double.TryParse(text, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out number))
			return true;

		number = default;
		return false;
	}

	private static bool TryParseDateTime(string text, DxpFieldEvalContext context, out DateTimeOffset dateTime)
	{
		var culture = context.Culture ?? System.Globalization.CultureInfo.CurrentCulture;
		var styles = System.Globalization.DateTimeStyles.AllowWhiteSpaces | System.Globalization.DateTimeStyles.AssumeLocal;
		if (DateTimeOffset.TryParse(text, culture, styles, out dateTime))
			return true;
		if (DateTimeOffset.TryParse(text, System.Globalization.CultureInfo.InvariantCulture, styles, out dateTime))
			return true;
		return false;
	}
}
