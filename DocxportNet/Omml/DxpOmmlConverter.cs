using DocumentFormat.OpenXml;
using OfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;
using OfficeMathParagraph = DocumentFormat.OpenXml.Math.Paragraph;

namespace DocxportNet.Omml;

/// <summary>
/// Converts standalone Office Math Markup Language (OMML) without requiring a
/// document, walker, or export visitor.
/// </summary>
public static class DxpOmmlConverter
{
    public static DxpOmmlConversionResult Convert(
        string omml,
        DxpOmmlOutputFormat format,
        DxpOmmlConversionOptions? options = null)
    {
        options ??= new DxpOmmlConversionOptions();
        OmmlDocument document = OmmlParser.Parse(omml, options);
        return EnforceOutputLimit(OmmlWriter.Write(document, format, options), options);
    }

    public static DxpOmmlConversionResult Convert(
        OfficeMath omml,
        DxpOmmlOutputFormat format,
        DxpOmmlConversionOptions? options = null) =>
        ConvertOpenXml(omml, format, options);

    public static DxpOmmlConversionResult Convert(
        OfficeMathParagraph omml,
        DxpOmmlOutputFormat format,
        DxpOmmlConversionOptions? options = null) =>
        ConvertOpenXml(omml, format, options);

    public static bool TryConvert(
        string? omml,
        DxpOmmlOutputFormat format,
        out DxpOmmlConversionResult? result,
        out DxpOmmlException? error,
        DxpOmmlConversionOptions? options = null)
    {
        try
        {
            result = Convert(omml!, format, options);
            error = null;
            return true;
        }
        catch (DxpOmmlException exception)
        {
            result = null;
            error = exception;
            return false;
        }
    }

    public static string ToMathMl(string omml, DxpOmmlConversionOptions? options = null) =>
        Convert(omml, DxpOmmlOutputFormat.MathMl, options).Output;

    public static string ToHtml(string omml, DxpOmmlConversionOptions? options = null) =>
        ToMathMl(omml, options);

    public static string ToLatex(string omml, DxpOmmlConversionOptions? options = null) =>
        Convert(omml, DxpOmmlOutputFormat.Latex, options).Output;

    public static string ToUnicodeMath(string omml, DxpOmmlConversionOptions? options = null) =>
        Convert(omml, DxpOmmlOutputFormat.UnicodeMath, options).Output;

    public static string ToText(string omml, DxpOmmlConversionOptions? options = null) =>
        Convert(omml, DxpOmmlOutputFormat.Text, options).Output;

    private static DxpOmmlConversionResult ConvertOpenXml(
        OpenXmlElement? omml,
        DxpOmmlOutputFormat format,
        DxpOmmlConversionOptions? options)
    {
        if (omml is null)
            throw new ArgumentNullException(nameof(omml));

        options ??= new DxpOmmlConversionOptions();
        OmmlDocument document = OmmlParser.Parse(omml, options);
        return EnforceOutputLimit(OmmlWriter.Write(document, format, options), options);
    }

    private static DxpOmmlConversionResult EnforceOutputLimit(
        DxpOmmlConversionResult result,
        DxpOmmlConversionOptions options)
    {
        if (options.MaxOutputCharacters <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxOutputCharacters), "The output limit must be positive.");
        if (result.Output.Length > options.MaxOutputCharacters)
            throw new DxpOmmlResourceLimitException($"OMML output exceeds the {options.MaxOutputCharacters} character limit.");
        return result;
    }
}
