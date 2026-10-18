using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Walker;

namespace DocxportNet.Tests;

public sealed class ListMarkerFormatterTests
{
    [Fact]
    public void FormatsNativeListNumbers()
    {
        Assert.Equal("㉑", DxpListMarkerFormatter.FormatNumber(21, NumberFormatValues.DecimalEnclosedCircle));
        Assert.Equal("- 3 -", DxpListMarkerFormatter.FormatNumber(3, NumberFormatValues.NumberInDash));
        Assert.Equal("１２", DxpListMarkerFormatter.FormatNumber(12, NumberFormatValues.DecimalFullWidth2));
        Assert.Equal("一二", DxpListMarkerFormatter.FormatNumber(12, NumberFormatValues.IdeographDigital));
        Assert.Equal("**", DxpListMarkerFormatter.FormatNumber(5, NumberFormatValues.Chicago));
        Assert.Equal("а", DxpListMarkerFormatter.FormatNumber(1, NumberFormatValues.RussianLower));
        Assert.Equal("А", DxpListMarkerFormatter.FormatNumber(1, NumberFormatValues.RussianUpper));
    }
}
