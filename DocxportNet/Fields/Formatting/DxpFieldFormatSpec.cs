namespace DocxportNet.Fields.Formatting;

public interface IDxpFieldFormatSpec
{
    string Apply(string text, DxpFieldValue value, DxpFieldEvalContext context);
}
