namespace DocxportNet.Fields;

public sealed class DxpFieldFormatter
{
	public string Format(DxpFieldValue value, IReadOnlyList<Formatting.IDxpFieldFormatSpec> specs, DxpFieldEvalContext context)
	{
		string text = FormatBaseValue(value, context);
		foreach (var spec in specs)
			text = spec.Apply(text, value, context);
		return text;
	}

	private static string FormatBaseValue(DxpFieldValue value, DxpFieldEvalContext context)
	{
		var culture = context.Culture ?? System.Globalization.CultureInfo.CurrentCulture;
		return value.Kind switch
		{
			DxpFieldValueKind.String => value.StringValue ?? string.Empty,
			DxpFieldValueKind.Number => value.NumberValue?.ToString(culture) ?? string.Empty,
			DxpFieldValueKind.DateTime => value.DateTimeValue?.ToString(culture) ?? string.Empty,
			_ => string.Empty
		};
	}
}
