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
}
