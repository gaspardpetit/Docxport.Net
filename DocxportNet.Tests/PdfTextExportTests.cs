using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Visitors.Pdf;
using PdfSharp.Fonts;

namespace DocxportNet.Tests;

public sealed class PdfTextExportTests
{
    private static readonly IFontResolver FontResolver = new TestFontResolver();

    [Fact]
    public void ExportsParagraphsAndExplicitPageBreak()
    {
        GlobalFontSettings.FontResolver = FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("First paragraph"))),
                new Paragraph(new Run(new Break { Type = BreakValues.Page }, new Text("Second page")))));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpPdfVisitor());

        Assert.StartsWith("%PDF", System.Text.Encoding.ASCII.GetString(result, 0, 4));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("First paragraph", pdf.GetPage(1).Text);
        Assert.Contains("Second page", pdf.GetPage(2).Text);
    }

    [Fact]
    public void WrapsLongParagraphAcrossPages()
    {
        GlobalFontSettings.FontResolver = FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text(
                string.Join(" ", Enumerable.Repeat("wrapped text", 3000)))))));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.True(pdf.NumberOfPages > 1);
        Assert.Contains("wrapped", pdf.GetPage(1).Text);
        Assert.Contains("wrapped", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void UsesEffectiveRunFontSizeWithoutBreakingStyledWords()
    {
        GlobalFontSettings.FontResolver = FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("small ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new RunProperties(new Bold(), new Italic(), new FontSize { Val = "40" }),
                    new Text("inter")),
                new Run(new Text("national")))));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains("international", page.Text.Replace(" ", string.Empty, StringComparison.Ordinal));
        var small = Assert.Single(page.Letters.Where(letter => letter.Value == "s"));
        var large = page.Letters.First(letter => letter.Value == "i");
        Assert.True(large.FontSize > small.FontSize * 1.5);
        Assert.Equal(large.StartBaseLine.Y, small.StartBaseLine.Y);
    }

    [Fact]
    public void InheritsFontSizeFromParagraphStyle()
    {
        GlobalFontSettings.FontResolver = FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new Style
            {
                StyleId = "Large", Type = StyleValues.Paragraph,
                StyleRunProperties = new StyleRunProperties(new FontSize { Val = "36" })
            });
            main.Document = new Document(new Body(
                new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Large" }),
                    new Run(new Text("Big"))),
                new Paragraph(new Run(new Text("small")))));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        var letters = Assert.Single(pdf.GetPages()).Letters.ToArray();
        Assert.True(letters.First(letter => letter.Value == "B").FontSize >
            letters.First(letter => letter.Value == "s").FontSize * 1.5);
    }

    private sealed class TestFontResolver : IFontResolver
    {
        private readonly byte[] _fontData = File.ReadAllBytes(OperatingSystem.IsWindows()
            ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf")
            : "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf");

        public FontResolverInfo ResolveTypeface(string familyName, bool bold, bool italic) => new("test-font");

        public byte[] GetFont(string faceName) => _fontData;
    }
}
