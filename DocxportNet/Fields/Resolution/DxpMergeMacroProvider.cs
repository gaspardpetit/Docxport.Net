using System.Globalization;
using DocxportNet.Fields;

namespace DocxportNet.Fields.Resolution;

public abstract class DxpMergeMacroProvider : DxpIMergeMacroProvider
{
    public abstract bool CanHandle(CultureInfo culture);
    public abstract string? Resolve(string macroName, IDxpMergeRecordCursor cursor, CultureInfo culture);

    protected static string? GetFieldString(IDxpMergeRecordCursor cursor, string fieldName, CultureInfo culture)
    {
        var value = cursor.GetValue(fieldName);
        if (value == null)
            return null;
        return ToStringValue(value.Value, culture);
    }

    protected static string? ToStringValue(DxpFieldValue value, CultureInfo culture)
    {
        switch (value.Kind)
        {
            case DxpFieldValueKind.String:
                return value.StringValue;
            case DxpFieldValueKind.Number:
                return value.NumberValue.HasValue ? value.NumberValue.Value.ToString(culture) : null;
            case DxpFieldValueKind.DateTime:
                return value.DateTimeValue.HasValue ? value.DateTimeValue.Value.ToString(culture) : null;
            default:
                return null;
        }
    }
}
