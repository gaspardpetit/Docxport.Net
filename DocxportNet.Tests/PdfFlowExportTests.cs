using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Visitors.Pdf;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace DocxportNet.Tests;

public sealed class PdfFlowExportTests
{
    [Fact]
    public void RendersTableCellsAndParagraphs()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Introduction"))),
                new Table(new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("Left cell")))),
                    new TableCell(new Paragraph(new Run(new Text("Right cell")))))),
                new Paragraph(new Run(new Text("After table")))));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains("Introduction", page.Text);
        Assert.Contains("Leftcell", page.Text);
        Assert.Contains("Rightcell", page.Text);
        Assert.Contains("Aftertable", page.Text);
    }

    [Fact]
    public void PreservesPageGeometryAndExplicitBreak()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("First page"))),
                new Paragraph(new Run(new Break { Type = BreakValues.Page }, new Text("Second page"))),
                new SectionProperties(new PageSize { Width = 12240, Height = 15840 })));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.InRange(pdf.GetPage(1).Width, 611.9, 612.1);
        Assert.InRange(pdf.GetPage(1).Height, 791.9, 792.1);
        Assert.Contains("Firstpage", pdf.GetPage(1).Text);
        Assert.Contains("Secondpage", pdf.GetPage(2).Text);
    }

    [Fact]
    public void StartsOddSectionOnAnOddPage()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("First section")),
                    new ParagraphProperties(new SectionProperties())),
                new Paragraph(new Run(new Text("Odd section"))),
                new SectionProperties(new SectionType { Val = SectionMarkValues.OddPage })));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Contains("Firstsection", pdf.GetPage(1).Text);
        Assert.Contains("Oddsection", pdf.GetPage(3).Text);
    }

    [Fact]
    public void KeepsSimpleContinuousSectionOnTheSamePage()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Before")),
                    new ParagraphProperties(new SectionProperties())),
                new Paragraph(new Run(new Text("After"))),
                new SectionProperties(new SectionType { Val = SectionMarkValues.Continuous })));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains("Before", page.Text);
        Assert.Contains("After", page.Text);
    }

    [Fact]
    public void ExplicitEvenSectionBreakWinsOverOddPageNumberRestart()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings(new EvenAndOddHeaders());
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("First section")),
                    new ParagraphProperties(new SectionProperties())),
                new Paragraph(new Run(new Text("Even section"))),
                new SectionProperties(
                    new SectionType { Val = SectionMarkValues.EvenPage },
                    new PageNumberType { Start = 7 })));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("Evensection", pdf.GetPage(2).Text);
    }

    [Fact]
    public void PreservesSignedTopMargin()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Signed margin"))),
                new SectionProperties(new PageMargin { Top = -1100, Bottom = -1200 })));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Single(pdf.GetPages());
        Assert.True(pdf.GetPage(1).Letters.Min(letter => letter.GlyphRectangle.Bottom) > 720);
    }

    [Fact]
    public void FormatsFootnoteReferenceAndNoteFromSectionProperties()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var notes = main.AddNewPart<FootnotesPart>();
            notes.Footnotes = new Footnotes(new Footnote(
                new Paragraph(new Run(new FootnoteReferenceMark()), new Run(new Text(" Note body"))))
                { Id = 1 });
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Body")), new Run(new FootnoteReference { Id = 1 })),
                new SectionProperties(new FootnoteProperties(
                    new NumberingFormat { Val = NumberFormatValues.UpperRoman },
                    new NumberingStart { Val = 3 }))));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Contains("III", string.Concat(pdf.GetPages().Select(page => page.Text)));
        Assert.DoesNotContain("Body1", string.Concat(pdf.GetPages().Select(page => page.Text)));
    }

    [Fact]
    public void DoesNotPaintAmbiguousHighlightMatches()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new PdfDocument();
        var page = source.AddPage();
        using (var graphics = XGraphics.FromPdfPage(page))
        {
            var font = new XFont("Arial", 12);
            graphics.DrawString("Important", font, XBrushes.Black, 40, 80);
            graphics.DrawString("Important", font, XBrushes.Black, 40, 120);
        }
        var originalStreamCount = page.Contents.Elements.Count;
        using var output = new MemoryStream();

        DxpPdfHighlightOverlay.Save(source, output,
            [new DxpPdfHighlight("Important", HighlightColorValues.Yellow)]);

        using var result = PdfReader.Open(new MemoryStream(output.ToArray()), PdfDocumentOpenMode.Import);
        Assert.Equal(originalStreamCount, result.Pages[0].Contents.Elements.Count);
    }

    [Fact]
    public void RendersCarriageReturnAsLineBreak()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("First"), new CarriageReturn(), new Text("Second")))));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        var words = pdf.GetPage(1).GetWords().ToArray();
        Assert.True(words[0].BoundingBox.Bottom > words[1].BoundingBox.Bottom);
    }

    [Fact]
    public void BottomAlignsMultilineSectionWithoutSplittingIt()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("First line"), new Break(),
                    new Text("Second line"), new Break(), new Text("Third line"))),
                new SectionProperties(new VerticalTextAlignmentOnPage { Val = VerticalJustificationValues.Bottom })));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains("Thirdline", page.Text);
        Assert.InRange(page.Letters.Min(letter => letter.GlyphRectangle.Bottom), 0, 100);
    }

    [Fact]
    public void PlacesFootnoteWithMirroredParagraphIndent()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var notes = main.AddNewPart<FootnotesPart>();
            notes.Footnotes = new Footnotes(new Footnote(new Paragraph(
                new Run(new FootnoteReferenceMark()), new Run(new Text(" Mirrored note")))) { Id = 1 });
            main.Document = new Document(new Body(new Paragraph(
                new ParagraphProperties(new MirrorIndents()),
                new Run(new Text("Mirrored body")), new Run(new FootnoteReference { Id = 1 }))));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains("Mirrorednote", page.Text);
    }

    [Fact]
    public void EmbedsInlineImage()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "probes", "inline-image-paragraph-probe__24.docx");

        byte[] result = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        var page = Assert.Single(pdf.GetPages());
        var image = Assert.Single(page.GetImages());
        Assert.InRange(image.Bounds.Width, 23.9, 24.1);
        Assert.InRange(image.Bounds.Height, 23.9, 24.1);
    }

    [Theory]
    [InlineData("EndnoteBasic.docx")]
    [InlineData("EndnoteTable.docx")]
    public void RendersEndnoteFixtures(string fileName)
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "features", fileName);

        byte[] result = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.True(pdf.NumberOfPages > 0);
        Assert.Contains(pdf.GetPages(), page => page.Letters.Any());
    }

    [Fact]
    public void RendersTextBoxFixture()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "features", "text-boxes",
            "TextBoxVerticalTextSingleLine.docx");

        byte[] result = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Contains(pdf.GetPages(), page => page.Letters.Any());
    }
}
