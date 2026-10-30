using DocumentFormat.OpenXml;

namespace DocxportNet.Walker.Context;

internal static class DxpOpenXmlDate
{
    internal static DateTime? Read(DateTimeValue? value)
    {
        if (value == null)
            return null;

        try { return value.Value; }
        catch (FormatException) { return null; }
    }
}
