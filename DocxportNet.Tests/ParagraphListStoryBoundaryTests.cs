using DocxportNet.API;
using DocxportNet.Visitors.Pdf;
using PdfSharp.Fonts;

namespace DocxportNet.Tests;

public sealed class ParagraphListStoryBoundaryTests
{
    [Theory]
    [InlineData("WordMixedSectionStories.docx")]
    [InlineData("WordMixedSectionStoryContent.docx")]
    [InlineData("WordMultisectionPositionalTabsAllStories.docx")]
    public void EmptySectionBreakDoesNotAdvanceListNumber(string fixture)
    {
        using var pdf = Export(fixture);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("9.Secondpagenote", pdf.GetPage(2).Text);
        Assert.DoesNotContain("10.Secondpagenote", pdf.GetPage(2).Text);
    }

    [Fact]
    public void HeaderFooterAndBodyUseSeparateListCounters()
    {
        using var pdf = Export("WordMixedSectionStyledStories.docx");
        var text = System.Text.RegularExpressions.Regex.Replace(pdf.GetPage(2).Text, @"\s+", "");
        Assert.Contains("1.Secondsectionheader", text);
        Assert.Contains("2.Secondsectionfooter", text);
        Assert.Contains("9.Secondpagenote", text);
    }

    [Theory]
    [InlineData("WordContextualSpacingAllStories.docx", 102.4)]
    [InlineData("WordContextualSpacingInheritedAllStories.docx", 102.4)]
    [InlineData("WordStyledEmptySpacingAllStories.docx", 116.4)]
    public void ContextualSpacingReservesHeaderSpace(string fixture, double firstBodyTop)
    {
        using var pdf = Export(fixture);
        var page = pdf.GetPage(1);
        var bodyTop = page.Letters.Where(letter => letter.Value == "B" &&
                page.Height - letter.GlyphRectangle.Top is > 80 and < 250)
            .Min(letter => page.Height - letter.GlyphRectangle.Top);
        Assert.InRange(bodyTop, firstBodyTop - 2, firstBodyTop + 2);
    }

    [Fact]
    public void ContextualSpacingSurvivesSectionBoundary()
    {
        using var pdf = Export("WordContextualSpacingSections.docx");
        var page = pdf.GetPage(2);
        Assert.Contains(page.Letters, letter => letter.Value == "B" &&
            page.Height - letter.GlyphRectangle.Top is > 89 and < 94);
    }

    private static UglyToad.PdfPig.PdfDocument Export(string fixture)
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "legacy", fixture);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        return UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
    }
}
