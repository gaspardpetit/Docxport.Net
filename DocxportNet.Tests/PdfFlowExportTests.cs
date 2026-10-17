using System.Buffers.Binary;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
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
    public void PaintsRepeatedHighlightsAtTheirSourceOffsets()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new PdfDocument();
        var page = source.AddPage();
        using (var graphics = XGraphics.FromPdfPage(page))
            graphics.DrawString("Important Important", new XFont("Arial", 12), XBrushes.Black, 40, 80);
        var originalStreamCount = page.Contents.Elements.Count;
        var paragraphText = new System.Text.StringBuilder("Important Important");
        using var output = new MemoryStream();

        DxpPdfHighlightOverlay.Save(source, output,
        [
            new DxpPdfHighlight("Important", HighlightColorValues.Yellow,
                ParagraphText: paragraphText, SourceOffset: 0),
            new DxpPdfHighlight("Important", HighlightColorValues.Green,
                ParagraphText: paragraphText, SourceOffset: 9)
        ]);

        using var result = PdfReader.Open(new MemoryStream(output.ToArray()), PdfDocumentOpenMode.Import);
        Assert.True(result.Pages[0].Contents.Elements.Count > originalStreamCount);
    }

    [Fact]
    public void PaintsShortAndSplitParagraphHighlights()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new PdfDocument();
        var firstPage = source.AddPage();
        using (var graphics = XGraphics.FromPdfPage(firstPage))
            graphics.DrawString("Opening A", new XFont("Arial", 12), XBrushes.Black, 40, 80);
        var originalStreamCount = firstPage.Contents.Elements.Count;
        var secondPage = source.AddPage();
        using (var graphics = XGraphics.FromPdfPage(secondPage))
            graphics.DrawString("following text", new XFont("Arial", 12), XBrushes.Black, 40, 80);
        using var output = new MemoryStream();

        DxpPdfHighlightOverlay.Save(source, output,
            [new DxpPdfHighlight("A", HighlightColorValues.Yellow,
                ParagraphText: new System.Text.StringBuilder("Opening A following text"), SourceOffset: 7)]);

        using var result = PdfReader.Open(new MemoryStream(output.ToArray()), PdfDocumentOpenMode.Import);
        Assert.True(result.Pages[0].Contents.Elements.Count > originalStreamCount);
    }

    [Fact]
    public void PaintsRepeatedHeaderHighlightOnEachPage()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new PdfDocument();
        for (var pageIndex = 0; pageIndex < 2; pageIndex++)
        {
            var page = source.AddPage();
            using var graphics = XGraphics.FromPdfPage(page);
            graphics.DrawString("Header", new XFont("Arial", 12), XBrushes.Black, 40, 40);
        }
        var originalStreamCount = source.Pages[0].Contents.Elements.Count;
        using var output = new MemoryStream();

        DxpPdfHighlightOverlay.Save(source, output,
            [new DxpPdfHighlight("Header", HighlightColorValues.Yellow,
                ParagraphText: new System.Text.StringBuilder("Header"), RepeatedStory: true)]);

        using var result = PdfReader.Open(new MemoryStream(output.ToArray()), PdfDocumentOpenMode.Import);
        Assert.All(result.Pages.Cast<PdfPage>(), page =>
            Assert.True(page.Contents.Elements.Count > originalStreamCount));
    }

    [Fact]
    public void PaintsWrappedHighlightOnBothLines()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new PdfDocument();
        var page = source.AddPage();
        using (var graphics = XGraphics.FromPdfPage(page))
        {
            var font = new XFont("Arial", 12);
            graphics.DrawString("First", font, XBrushes.Black, 40, 80);
            graphics.DrawString("Second", font, XBrushes.Black, 40, 100);
        }
        using var output = new MemoryStream();

        DxpPdfHighlightOverlay.Save(source, output,
            [new DxpPdfHighlight("FirstSecond", HighlightColorValues.Yellow,
                ParagraphText: new System.Text.StringBuilder("FirstSecond"))]);

        using var result = PdfReader.Open(new MemoryStream(output.ToArray()), PdfDocumentOpenMode.Import);
        var content = System.Text.Encoding.Latin1.GetString(
            result.Pages[0].Contents.CreateSingleContent().Stream!.UnfilteredValue);
        Assert.True(System.Text.RegularExpressions.Regex.Matches(content, @"\bre\b").Count >= 2);
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
    public void KeepsCarriageReturnInsideHyperlinkInTextOrder()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Hyperlink(new Run(new Text("First"), new CarriageReturn(),
                    new Text("Second"))) { Anchor = "destination" },
                new BookmarkStart { Name = "destination", Id = "1" },
                new BookmarkEnd { Id = "1" })));
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
    public void AlignsSectionBodyWhenFootnotesExist()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var notes = main.AddNewPart<FootnotesPart>();
            notes.Footnotes = new Footnotes(new Footnote(new Paragraph(
                new Run(new FootnoteReferenceMark()), new Run(new Text(" Note")))) { Id = 1 });
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Aligned body")),
                    new Run(new FootnoteReference { Id = 1 })),
                new SectionProperties(new VerticalTextAlignmentOnPage
                    { Val = VerticalJustificationValues.Bottom })));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        var page = Assert.Single(pdf.GetPages());
        var body = page.GetWords().First(word => word.Text == "Aligned");
        Assert.InRange(body.BoundingBox.Bottom, 0, 150);
    }

    [Fact]
    public void ConvertsRedRgbTiffToRedBmp()
    {
        var tiff = new byte[158];
        tiff[0] = (byte)'I';
        tiff[1] = (byte)'I';
        BinaryPrimitives.WriteUInt16LittleEndian(tiff.AsSpan(2), 42);
        BinaryPrimitives.WriteInt32LittleEndian(tiff.AsSpan(4), 8);
        BinaryPrimitives.WriteUInt16LittleEndian(tiff.AsSpan(8), 11);
        WriteEntry(0, 256, 4, 1);
        WriteEntry(1, 257, 4, 1);
        WriteEntry(2, 258, 3, 146, 3);
        WriteEntry(3, 259, 3, 5);
        WriteEntry(4, 262, 3, 2);
        WriteEntry(5, 273, 4, 152);
        WriteEntry(6, 274, 3, 1);
        WriteEntry(7, 277, 3, 3);
        WriteEntry(8, 278, 4, 1);
        WriteEntry(9, 279, 4, 6);
        WriteEntry(10, 284, 3, 1);
        for (var sample = 0; sample < 3; sample++)
            BinaryPrimitives.WriteUInt16LittleEndian(tiff.AsSpan(146 + sample * 2), 8);
        var codes = new[] { 256, 255, 0, 0, 257 };
        for (var codeIndex = 0; codeIndex < codes.Length; codeIndex++)
            for (var bit = 0; bit < 9; bit++)
            {
                var bitIndex = codeIndex * 9 + bit;
                tiff[152 + bitIndex / 8] |= (byte)(((codes[codeIndex] >> (8 - bit)) & 1) <<
                    (7 - bitIndex % 8));
            }

        var bitmap = DxpTiffConverter.ToBmp(tiff);

        Assert.Equal(new byte[] { 0, 0, 255 }, bitmap.AsSpan(54, 3).ToArray());

        void WriteEntry(int index, ushort tag, ushort type, int value, int count = 1)
        {
            var offset = 10 + index * 12;
            BinaryPrimitives.WriteUInt16LittleEndian(tiff.AsSpan(offset), tag);
            BinaryPrimitives.WriteUInt16LittleEndian(tiff.AsSpan(offset + 2), type);
            BinaryPrimitives.WriteInt32LittleEndian(tiff.AsSpan(offset + 4), count);
            BinaryPrimitives.WriteInt32LittleEndian(tiff.AsSpan(offset + 8), value);
        }
    }

    [Fact]
    public void TransformsInlineCropsAndQuarterTurns()
    {
        Assert.True(DxpFlowPdfVisitor.RequiresRasterTransform(false,
            new DxpImagePresentation { Crop = new DxpImageCrop(0.1, 0, 0, 0) }));
        Assert.True(DxpFlowPdfVisitor.RequiresRasterTransform(false,
            new DxpImagePresentation { RotationDegrees = 90 }));
        Assert.False(DxpFlowPdfVisitor.RequiresRasterTransform(false, new DxpImagePresentation()));
    }

    [Fact]
    public void ConvertsSkippedRle4PixelsThroughPaletteZero()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "legacy", "WordBmpRle4DeltaAllStories.docx");
        using var document = WordprocessingDocument.Open(source, false);
        using var image = Assert.Single(document.MainDocumentPart!.ImageParts).GetStream();
        using var data = new MemoryStream();
        image.CopyTo(data);

        var bitmap = DxpBmp24Converter.Normalize(data.ToArray());

        Assert.Equal(new byte[] { 0, 200, 0 }, bitmap.AsSpan(54 + 4 * 3, 3).ToArray());
    }

    [Fact]
    public void LetsImportedHtmlTableRowsGrowWithWrappedText()
    {
        var document = new MigraDoc.DocumentObjectModel.Document();
        var section = document.AddSection();
        var html = "<table><tr><td>Long text that wraps into several lines in a narrow table cell</td></tr>" +
            "<tr><td>Following row</td></tr></table>";
        using var source = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(html));

        DxpHtmlChunkImporter.Append(section, source, "Arial");

        var table = section.Elements.OfType<MigraDoc.DocumentObjectModel.Tables.Table>().Single();
        Assert.All(table.Rows.Cast<MigraDoc.DocumentObjectModel.Tables.Row>(), row =>
            Assert.Equal(MigraDoc.DocumentObjectModel.Tables.RowHeightRule.AtLeast, row.HeightRule));
    }

    [Fact]
    public void DoesNotImportNestedHtmlRowsAsOuterRows()
    {
        var document = new MigraDoc.DocumentObjectModel.Document();
        var section = document.AddSection();
        using var source = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(
            "<table><tr><td>Outer<table><tr><td>Inner</td></tr></table></td></tr></table>"));

        DxpHtmlChunkImporter.Append(section, source, "Arial");

        var table = section.Elements.OfType<MigraDoc.DocumentObjectModel.Tables.Table>().Single();
        Assert.Single(table.Rows);
    }

    [Theory]
    [InlineData("", 0.75)]
    [InlineData(" width='20'", 15)]
    public void ImportsHtmlImageUsingIntrinsicDimensions(string attributes, double expectedWidth)
    {
        var document = new MigraDoc.DocumentObjectModel.Document();
        var section = document.AddSection();
        var html = "<p><img src='data:image/png;base64," +
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4//8/AwAI/AL+p5qgoAAAAABJRU5ErkJggg==" +
            "'" + attributes + "></p>";
        using var source = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(html));

        DxpHtmlChunkImporter.Append(section, source, "Arial");

        var paragraph = section.Elements.OfType<MigraDoc.DocumentObjectModel.Paragraph>().Single();
        var image = paragraph.Elements.OfType<MigraDoc.DocumentObjectModel.Shapes.Image>().Single();
        Assert.Equal(expectedWidth, image.Width.Point, 2);
        Assert.Equal(expectedWidth, image.Height.Point, 2);
    }

    [Fact]
    public void ImportsHtmlUsingDeclaredWindows1252Encoding()
    {
        var document = new MigraDoc.DocumentObjectModel.Document();
        var section = document.AddSection();
        var html = "<meta charset='windows-1252'><p>Caf\u00e9</p>";
        var bytes = System.Text.Encoding.Latin1.GetBytes(html);
        using var source = new MemoryStream(bytes);

        DxpHtmlChunkImporter.Append(section, source, "Arial");

        var paragraph = section.Elements.OfType<MigraDoc.DocumentObjectModel.Paragraph>().Single();
        Assert.Contains("Caf\u00e9", paragraph.Elements.OfType<MigraDoc.DocumentObjectModel.FormattedText>()
            .SelectMany(span => span.Elements.OfType<MigraDoc.DocumentObjectModel.Text>())
            .Select(text => text.Content));
    }

    [Fact]
    public void ImportsHtmlUsingPartContentTypeCharset()
    {
        var document = new MigraDoc.DocumentObjectModel.Document();
        var section = document.AddSection();
        using var source = new MemoryStream(System.Text.Encoding.Latin1.GetBytes("<p>Caf\u00e9</p>"));

        DxpHtmlChunkImporter.Append(section, source, "Arial", contentType: "text/html; charset=windows-1252");

        var paragraph = section.Elements.OfType<MigraDoc.DocumentObjectModel.Paragraph>().Single();
        Assert.Contains("Caf\u00e9", paragraph.Elements.OfType<MigraDoc.DocumentObjectModel.FormattedText>()
            .SelectMany(span => span.Elements.OfType<MigraDoc.DocumentObjectModel.Text>())
            .Select(text => text.Content));
    }

    [Fact]
    public void ReadsPlainTextChunkUsingPartContentTypeCharset()
    {
        var encoding = DxpHtmlChunkImporter.DeclaredEncoding("text/plain; charset=windows-1252");
        Assert.NotNull(encoding);
        Assert.Equal("Caf\u00e9", encoding.GetString(System.Text.Encoding.Latin1.GetBytes("Caf\u00e9")));
    }

    [Fact]
    public void ImportsHtmlTableCellSpans()
    {
        var document = new MigraDoc.DocumentObjectModel.Document();
        var section = document.AddSection();
        var html = "<table><tr><td rowspan='2'>A</td><td>B</td></tr>" +
            "<tr><td>C</td></tr><tr><td colspan='2'>D</td></tr></table>";
        using var source = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(html));

        DxpHtmlChunkImporter.Append(section, source, "Arial");

        var table = section.Elements.OfType<MigraDoc.DocumentObjectModel.Tables.Table>().Single();
        Assert.Equal(2, table.Columns.Count);
        Assert.Equal(3, table.Rows.Count);
        Assert.Equal(1, table.Rows[0].Cells[0].MergeDown);
        Assert.Empty(table.Rows[1].Cells[0].Elements);
        Assert.NotEmpty(table.Rows[1].Cells[1].Elements);
        Assert.Equal(1, table.Rows[2].Cells[0].MergeRight);
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
