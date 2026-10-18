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
    public void TracksSourceAndOffsetsAcrossMultipleParagraphFragments()
    {
        var fragments = new DxpPdfParagraphFragments();
        var source = new Paragraph(new Run(new Text("FirstSecondThird")));
        var document = new MigraDoc.DocumentObjectModel.Document();
        var section = document.AddSection();
        var first = section.AddParagraph();
        var second = section.AddParagraph();
        var third = section.AddParagraph();

        fragments.RegisterSource(first, source, styleRefCandidate: true);
        fragments.RegisterContinuation(first, second, 5);
        fragments.RegisterContinuation(second, third, 6);

        Assert.True(fragments.TryGet(third, out var last));
        Assert.Same(source, last.Source);
        Assert.Same(first, last.Root);
        Assert.Same(first, fragments.RootOf(third));
        Assert.Equal(11, last.RenderedTextStart);
        Assert.Equal(2, last.Index);
        Assert.True(last.StyleRefCandidate);
        Assert.True(fragments.HasInheritedFlag(new HashSet<MigraDoc.DocumentObjectModel.Paragraph> { first }, third));
        Assert.True(fragments.HasInheritedFlag(new HashSet<MigraDoc.DocumentObjectModel.Paragraph> { third }, third));
        var widths = new Dictionary<MigraDoc.DocumentObjectModel.Paragraph, double> { [first] = 144 };
        Assert.True(fragments.TryGetInheritedValue(widths, third, out var inheritedWidth));
        Assert.Equal(144, inheritedWidth);
        widths[third] = 90;
        Assert.True(fragments.TryGetInheritedValue(widths, third, out var localWidth));
        Assert.Equal(90, localWidth);
        fragments.Clear();
        Assert.False(fragments.TryGet(third, out _));
        Assert.Same(third, fragments.RootOf(third));
    }

    [Fact]
    public void IgnoresContinuationsOfSkippedSourceParagraphs()
    {
        var fragments = new DxpPdfParagraphFragments();
        var section = new MigraDoc.DocumentObjectModel.Document().AddSection();
        var first = section.AddParagraph();
        var second = section.AddParagraph();
        var third = section.AddParagraph();

        fragments.RegisterContinuation(first, second, 0);
        fragments.RegisterContinuation(second, third, 0);

        Assert.False(fragments.TryGet(second, out _));
        Assert.False(fragments.TryGet(third, out _));
    }

    [Fact]
    public void ExportsMultipleBreaksInSkippedParagraphWithoutThrowing()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new ParagraphProperties(new ParagraphMarkRunProperties(new Deleted())),
                new Run(new Break { Type = BreakValues.Page },
                    new Break { Type = BreakValues.Page }))));
        }

        var result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        Assert.NotEmpty(result);
    }

    [Fact]
    public void InheritsVerticalAlignmentMetricsAcrossFragments()
    {
        var fragments = new DxpPdfParagraphFragments();
        var section = new MigraDoc.DocumentObjectModel.Document().AddSection();
        var first = section.AddParagraph();
        var second = section.AddParagraph();
        fragments.RegisterSource(first, new Paragraph(), styleRefCandidate: false);
        fragments.RegisterContinuation(first, second, 5);

        var topSizes = new Dictionary<MigraDoc.DocumentObjectModel.Paragraph, double> { [first] = 24 };
        var alignedFonts = new Dictionary<MigraDoc.DocumentObjectModel.Paragraph, (double Size, double Factor)>
        {
            [first] = (24, 0.35)
        };
        var headerFlags = new HashSet<MigraDoc.DocumentObjectModel.Paragraph> { first };

        Assert.True(fragments.TryGetInheritedValue(topSizes, second, out var topSize));
        Assert.Equal(24, topSize);
        Assert.True(fragments.TryGetInheritedValue(alignedFonts, second, out var alignedFont));
        Assert.Equal((24, 0.35), alignedFont);
        Assert.True(fragments.HasInheritedFlag(headerFlags, second));
    }

    [Fact]
    public void ResolvesPageRelativeImageAlignment()
    {
        var horizontal = new DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalPosition(
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignment("center"));
        horizontal.SetAttribute(new OpenXmlAttribute("relativeFrom", null, "page"));
        var vertical = new DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalPosition(
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalAlignment("bottom"));
        vertical.SetAttribute(new OpenXmlAttribute("relativeFrom", null, "page"));

        Assert.Equal(246, DxpFlowPdfVisitor.PageImageOffset(horizontal, true, 612, 120));
        Assert.Equal(692, DxpFlowPdfVisitor.PageImageOffset(vertical, false, 792, 100));

        horizontal.RemoveAllChildren();
        horizontal.AppendChild(new DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignment("inside"));
        Assert.Equal(0, DxpFlowPdfVisitor.PageImageOffset(horizontal, true, 612, 120));
        Assert.Equal(492, DxpFlowPdfVisitor.MirroredPageImageLeft(0, 120, 612, 2, true));
        horizontal.RemoveAllChildren();
        horizontal.AppendChild(new DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignment("outside"));
        Assert.Equal(492, DxpFlowPdfVisitor.PageImageOffset(horizontal, true, 612, 120));
        Assert.Equal(0, DxpFlowPdfVisitor.MirroredPageImageLeft(492, 120, 612, 2, true));
    }

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
    public void PlacesThreeColumnBreaksInDistinctColumns()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("First"), new Break { Type = BreakValues.Column })),
                new Paragraph(new Run(new Text("Second"), new Break { Type = BreakValues.Column })),
                new Paragraph(new Run(new Text("Third"))),
                new SectionProperties(new Columns { ColumnCount = 3, Space = "360" })));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        var words = Assert.Single(pdf.GetPages()).GetWords().ToDictionary(word => word.Text);
        Assert.Equal(3, words.Count);
        Assert.True(words["First"].BoundingBox.Left < words["Second"].BoundingBox.Left);
        Assert.True(words["Second"].BoundingBox.Left < words["Third"].BoundingBox.Left);
    }

    [Fact]
    public void KeepsLongParagraphsInExplicitMultiColumnFlow()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("A longer paragraph with enough words to exceed the old character cutoff"),
                    new Break { Type = BreakValues.Column })),
                new Paragraph(new Run(new Text("Second column"))),
                new SectionProperties(new Columns { ColumnCount = 2, Space = "360" })));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        var words = Assert.Single(pdf.GetPages()).GetWords().ToArray();
        Assert.True(words.Single(word => word.Text == "Second").BoundingBox.Left >
            words.First(word => word.Text.Contains("longer", StringComparison.OrdinalIgnoreCase)).BoundingBox.Left);
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
    public void KeepsImplicitNextPageBreakWithOddPageNumberRestart()
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
                new Paragraph(new Run(new Text("Second section"))),
                new SectionProperties(new PageNumberType { Start = 7 })));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("Secondsection", pdf.GetPage(2).Text);
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
    public void UsesCustomEndnoteMarkWithoutGeneratedOrdinal()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var notes = main.AddNewPart<EndnotesPart>();
            notes.Endnotes = new Endnotes(new Endnote(new Paragraph(
                new Run(new EndnoteReferenceMark(), new Text("* N")))) { Id = 1 });
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("B")),
                new Run(new EndnoteReference { Id = 1, CustomMarkFollows = true }),
                new Run(new Text("*")))));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        var text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("*", text);
        Assert.DoesNotContain("i", text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RestartsEndnoteLabelsInNextSection()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<EndnotesPart>().Endnotes = new Endnotes(
                new Endnote(new Paragraph(new Run(new EndnoteReferenceMark(), new Text(" A")))) { Id = 1 },
                new Endnote(new Paragraph(new Run(new EndnoteReferenceMark(), new Text(" B")))) { Id = 2 });
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("First"), new EndnoteReference { Id = 1 }),
                    new ParagraphProperties(new SectionProperties(new EndnoteProperties(
                        new NumberingRestart { Val = RestartNumberValues.EachSection })))),
                new Paragraph(new Run(new Text("Second"), new EndnoteReference { Id = 2 })),
                new SectionProperties(new EndnoteProperties(
                    new NumberingRestart { Val = RestartNumberValues.EachSection }))));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        var text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Equal("FirstiSecondiiAiB", text);
    }

    [Fact]
    public void ClearsCustomEndnoteMarksBeforeReusingVisitor()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        byte[] CreateSource(bool custom)
        {
            using var source = new MemoryStream();
            using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
            {
                var main = document.AddMainDocumentPart();
                main.AddNewPart<EndnotesPart>().Endnotes = new Endnotes(new Endnote(new Paragraph(
                    new Run(new EndnoteReferenceMark(), new Text(" N")))) { Id = 1 });
                main.Document = new Document(new Body(new Paragraph(new Run(new Text("B"),
                    new EndnoteReference { Id = 1, CustomMarkFollows = custom }))));
            }
            return source.ToArray();
        }

        var visitor = new DxpFlowPdfVisitor();
        _ = DxpExport.ExportToBytes(CreateSource(true), visitor);
        byte[] result = DxpExport.ExportToBytes(CreateSource(false), visitor);

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Contains("i", string.Concat(pdf.GetPages().Select(page => page.Text)));
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
    public void KeepsHyperlinkTextAfterPageBreakOnNextPage()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Hyperlink(new Run(new Text("Before"),
                    new Break { Type = BreakValues.Page }, new Text("After"))) { Anchor = "destination" },
                new BookmarkStart { Name = "destination", Id = "1" },
                new BookmarkEnd { Id = "1" })));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("Before", pdf.GetPage(1).Text);
        Assert.DoesNotContain("After", pdf.GetPage(1).Text);
        Assert.Contains("After", pdf.GetPage(2).Text);
    }

    [Fact]
    public void KeepsParagraphAlignmentAfterInParagraphPageBreak()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                new Run(new Text("Aligned"), new Break { Type = BreakValues.Page },
                    new Text("Aligned")))));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Equal(2, pdf.NumberOfPages);
        var first = Assert.Single(pdf.GetPage(1).GetWords());
        var second = Assert.Single(pdf.GetPage(2).GetWords());
        Assert.InRange(Math.Abs(first.BoundingBox.Left - second.BoundingBox.Left), 0, 0.5);
    }

    [Fact]
    public void KeepsNarrowJustificationOnLastContinuationLine()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new ParagraphProperties(new Justification { Val = JustificationValues.Both }),
                new InsertedRun(new Run(new Text("Opening text"))),
                new Run(new Break { Type = BreakValues.Page }, new Text("Alpha Beta Gamma")))));
        }

        var result = DxpExport.ExportToBytes(source.ToArray(),
            new DxpFlowPdfVisitor(availableFontFamilies: ["Aptos"]));

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Equal(2, pdf.NumberOfPages);
        var words = pdf.GetPage(2).GetWords().ToArray();
        Assert.Equal(["Alpha", "Beta", "Gamma"], words.Select(word => word.Text));
        Assert.InRange(words[1].BoundingBox.Left - words[0].BoundingBox.Right, 2.9, 3.6);
        Assert.InRange(words[2].BoundingBox.Left - words[1].BoundingBox.Right, 2.9, 3.6);
    }

    [Fact]
    public void RestoresDenseTrackedPaintWidthForManualContinuation()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var paragraph = new Paragraph(new ParagraphProperties(
                new Justification { Val = JustificationValues.Both }));
            for (var index = 0; index < 120; index++)
            {
                paragraph.Append(new Run(new Text("lorem ipsum ")
                    { Space = SpaceProcessingModeValues.Preserve }));
                if (index == 59)
                    paragraph.Append(new Run(new Break { Type = BreakValues.Page }));
            }
            main.Document = new Document(new Body(paragraph,
                new Paragraph(new InsertedRun(new Run(new Text("tracked"))))));
        }

        var visitor = new DxpFlowPdfVisitor(availableFontFamilies: ["Aptos"]);
        var result = DxpExport.ExportToBytes(source.ToArray(), visitor);
        var field = typeof(DxpFlowPdfVisitor).GetField("_document",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!;
        var rendered = (MigraDoc.DocumentObjectModel.Document)field.GetValue(visitor)!;
        var paragraphs = rendered.Sections[0].Elements
            .OfType<MigraDoc.DocumentObjectModel.Paragraph>().ToArray();
        Assert.Equal(3, paragraphs.Length);
        Assert.InRange(Math.Abs(paragraphs[0].Format.RightIndent.Point -
            paragraphs[1].Format.RightIndent.Point), 0, 0.01);
        Assert.Equal(0, paragraphs[1].Format.RightIndent.Point);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Equal(2, pdf.NumberOfPages);
    }

    [Fact]
    public void CarriesBodySpacingStateAcrossRepeatedFootnoteSplits()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var source = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "../../../../fixtures/docx/samples/Right to Repair.docx"));
        using var docx = new MemoryStream();
        docx.Write(File.ReadAllBytes(source));
        docx.Position = 0;
        using (var word = WordprocessingDocument.Open(docx, true))
        {
            foreach (var section in word.MainDocumentPart!.Document.Descendants<SectionProperties>())
            {
                var size = section.GetFirstChild<PageSize>() ?? section.AppendChild(new PageSize());
                size.Height = 14000;
            }
            word.MainDocumentPart.Document.Save();
        }
        var visitor = new DxpFlowPdfVisitor(availableFontFamilies: ["Aptos"]);
        var result = DxpExport.ExportToBytes(docx.ToArray(), visitor);
        var flags = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
        var compressed = (HashSet<MigraDoc.DocumentObjectModel.Paragraph>)
            typeof(DxpFlowPdfVisitor).GetField("_compressedBodyParagraphs", flags)!.GetValue(visitor)!;
        var widths = (HashSet<MigraDoc.DocumentObjectModel.Paragraph>)
            typeof(DxpFlowPdfVisitor).GetField("_layoutWidthParagraphs", flags)!.GetValue(visitor)!;
        var splits = (Dictionary<MigraDoc.DocumentObjectModel.Paragraph,
            MigraDoc.DocumentObjectModel.Paragraph>)
            typeof(DxpFlowPdfVisitor).GetField("_footnoteSplitContinuations", flags)!.GetValue(visitor)!;
        var repeated = splits.Where(entry =>
            DxpFootnoteParagraphSplitter.TryGetText(entry.Key, out var text) &&
            text.StartsWith("a detailed report21", StringComparison.Ordinal)).ToArray();
        foreach (var entry in repeated)
        {
            Assert.Contains(entry.Key, compressed);
            Assert.Contains(entry.Value, compressed);
            Assert.Contains(entry.Value.Elements.OfType<MigraDoc.DocumentObjectModel.FormattedText>(),
                span => Math.Abs(span.Size.Point - 9.5) < 0.01 &&
                    span.Elements.OfType<MigraDoc.DocumentObjectModel.Text>()
                        .Any(text => text.Content.Length > 0 &&
                            text.Content.All(character => character == ' ')));
        }
        var layoutSplits = splits.Where(entry =>
            DxpFootnoteParagraphSplitter.TryGetText(entry.Key, out var text) &&
            text.StartsWith("The World Intellectual Property", StringComparison.Ordinal)).ToArray();
        foreach (var entry in layoutSplits)
            Assert.Contains(entry.Value, widths);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.True(pdf.NumberOfPages > 1);
    }

    [Fact]
    public void TransfersBodySpacingStateAcrossSyntheticFootnoteSplits()
    {
        var visitor = new DxpFlowPdfVisitor();
        var flags = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
        var compressed = (HashSet<MigraDoc.DocumentObjectModel.Paragraph>)
            typeof(DxpFlowPdfVisitor).GetField("_compressedBodyParagraphs", flags)!.GetValue(visitor)!;
        var widths = (HashSet<MigraDoc.DocumentObjectModel.Paragraph>)
            typeof(DxpFlowPdfVisitor).GetField("_layoutWidthParagraphs", flags)!.GetValue(visitor)!;
        var section = new MigraDoc.DocumentObjectModel.Document().AddSection();
        var original = section.AddParagraph();
        var first = section.AddParagraph();
        var second = section.AddParagraph();
        compressed.Add(original);
        var firstSpace = first.AddFormattedText(" ");
        firstSpace.Size = MigraDoc.DocumentObjectModel.Unit.FromPoint(8.5);
        var secondSpace = second.AddFormattedText(" ");
        secondSpace.Size = MigraDoc.DocumentObjectModel.Unit.FromPoint(8.5);

        visitor.CarryFootnoteSplitState(original, first, 1, 30);
        visitor.CarryFootnoteSplitState(first, second, 1, 20);

        Assert.Contains(first, compressed);
        Assert.Contains(second, compressed);
        Assert.Equal(9.5, firstSpace.Size.Point);
        Assert.Equal(9.5, secondSpace.Size.Point);

        var layoutOriginal = section.AddParagraph();
        var layoutContinuation = section.AddParagraph();
        compressed.Add(layoutOriginal);
        widths.Add(layoutOriginal);
        visitor.CarryFootnoteSplitState(layoutOriginal, layoutContinuation, 1, 30);
        Assert.Contains(layoutContinuation, compressed);
        Assert.Contains(layoutContinuation, widths);
    }

    [Fact]
    public void DoesNotRepeatFirstLineIndentAfterPageBreak()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new ParagraphProperties(new Indentation { FirstLine = "600" }),
                new Run(new Text("First"), new Break { Type = BreakValues.Page },
                    new Text("Second")))));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Equal(2, pdf.NumberOfPages);
        var first = Assert.Single(pdf.GetPage(1).GetWords());
        var second = Assert.Single(pdf.GetPage(2).GetWords());
        Assert.InRange(first.BoundingBox.Left - second.BoundingBox.Left, 29.5, 30.5);
    }

    [Fact]
    public void AppliesParagraphSpaceBeforeOnlyBeforeFirstFragment()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new ParagraphProperties(new SpacingBetweenLines { Before = "400", After = "400" }),
                new Run(new Text("First"), new Break { Type = BreakValues.Page },
                    new Text("Second")))));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Equal(2, pdf.NumberOfPages);
        var first = Assert.Single(pdf.GetPage(1).GetWords());
        var second = Assert.Single(pdf.GetPage(2).GetWords());
        Assert.InRange(second.BoundingBox.Bottom - first.BoundingBox.Bottom, 19, 21);
    }

    [Fact]
    public void NumbersTextAfterInParagraphPageBreak()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("First"), new Break { Type = BreakValues.Page },
                    new Text("Second"))),
                new SectionProperties(new LineNumberType { CountBy = 1 })));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("1", pdf.GetPage(2).Text);
        Assert.Contains("Second", pdf.GetPage(2).Text);
    }

    [Fact]
    public void ContinuesLineNumbersAcrossSections()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new ParagraphProperties(new SectionProperties(
                        new LineNumberType { CountBy = 1 })),
                    new Run(new Text("First section"))),
                new Paragraph(new Run(new Text("Second section"))),
                new SectionProperties(new LineNumberType
                    { CountBy = 1, Restart = LineNumberRestartValues.Continuous })));
        }

        var result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains(pdf.GetPage(1).GetWords(), word => word.Text == "1");
        Assert.Contains(pdf.GetPage(2).GetWords(), word => word.Text == "2");
        Assert.DoesNotContain(pdf.GetPage(2).GetWords(), word => word.Text == "1");
    }

    [Fact]
    public void KeepsRunCharacterSpacingAfterPageBreak()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(
                new RunProperties(new Spacing { Val = 40 }),
                new Text("AAA"), new Break { Type = BreakValues.Page }, new Text("AAA")))));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Equal(2, pdf.NumberOfPages);
        var first = Assert.Single(pdf.GetPage(1).GetWords());
        var second = Assert.Single(pdf.GetPage(2).GetWords());
        Assert.InRange(Math.Abs(first.BoundingBox.Width - second.BoundingBox.Width), 0, 0.5);
    }

    [Fact]
    public void KeepsFitTextWidthAfterPageBreak()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(
                new RunProperties(new FitText { Val = 400 }),
                new Text("WWWW"), new Break { Type = BreakValues.Page }, new Text("WWWW"))),
                new Paragraph(new Run(new Text("WWWW")))));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Equal(2, pdf.NumberOfPages);
        var first = Assert.Single(pdf.GetPage(1).GetWords());
        var secondPageWords = pdf.GetPage(2).GetWords().ToArray();
        Assert.Equal(2, secondPageWords.Length);
        var second = secondPageWords[0];
        Assert.InRange(Math.Abs(first.BoundingBox.Width - second.BoundingBox.Width), 0, 0.5);
        Assert.True(second.BoundingBox.Width < secondPageWords[1].BoundingBox.Width);
    }

    [Fact]
    public void KeepsHangulBaselineCorrectionAfterPageBreak()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new RunProperties(new RunFonts
                    { Ascii = "Calibri", EastAsia = "Calibri" }),
                    new Text("ABC한"), new Break { Type = BreakValues.Page },
                    new Text("ABC한"))),
                new Paragraph(new Run(new RunProperties(new FitText { Val = 400 }), new Text("Fit")))));
        }

        var result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Equal(2, pdf.NumberOfPages);
        var first = pdf.GetPage(1).Letters.First(letter => letter.Value == "A");
        var second = pdf.GetPage(2).Letters.First(letter => letter.Value == "A");
        Assert.InRange(Math.Abs(first.StartBaseLine.Y - second.StartBaseLine.Y), 0, 0.1);
    }

    [Fact]
    public void KeepsSourceBaselineWithoutLeakingStrikeAcrossMultiplePageBreaks()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new RunProperties(new RunFonts { Ascii = "Calibri", EastAsia = "Calibri" }),
                    new Text("ABC한"), new Break { Type = BreakValues.Page }),
                new Run(new RunProperties(new RunFonts { Ascii = "Calibri", EastAsia = "Calibri" },
                    new Strike()), new Text("ABC한"), new Break { Type = BreakValues.Page }),
                new Run(new RunProperties(new RunFonts { Ascii = "Calibri", EastAsia = "Calibri" }),
                    new Text("ABC한")))));
        }

        var visitor = new DxpFlowPdfVisitor();
        var result = DxpExport.ExportToBytes(source.ToArray(), visitor);

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Equal(3, pdf.NumberOfPages);
        var baselines = Enumerable.Range(1, 3)
            .Select(page => pdf.GetPage(page).Letters.First(letter => letter.Value == "A").StartBaseLine.Y)
            .ToArray();
        Assert.All(baselines.Skip(1), baseline =>
            Assert.InRange(Math.Abs(baseline - baselines[0]), 0, 0.1));
        var flags = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
        var ranges = (System.Collections.IDictionary)
            typeof(DxpFlowPdfVisitor).GetField("_strikeRanges", flags)!.GetValue(visitor)!;
        Assert.Single(ranges.Keys.Cast<MigraDoc.DocumentObjectModel.Paragraph>());
    }

    [Fact]
    public void KeepsAdjustedGridGlyphPositionAfterPageBreak()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var grid = new DocGrid { Type = DocGridValues.LinesAndChars, LinePitch = 480 };
            grid.SetAttribute(new OpenXmlAttribute("w", "charSpace",
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "40960"));
            main.Document = new Document(new Body(
                new Paragraph(new ParagraphProperties(new AdjustRightIndent { Val = true }),
                    new Run(new RunProperties(new RunFonts
                        { Ascii = "Yu Mincho", EastAsia = "Yu Mincho" }),
                        new Text("日本語"), new Break { Type = BreakValues.Page },
                        new Text("日本語"))),
                new SectionProperties(grid)));
        }

        var result = DxpExport.ExportToBytes(source.ToArray(),
            new DxpFlowPdfVisitor(availableFontFamilies: ["Yu Mincho"]));

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Equal(2, pdf.NumberOfPages);
        var first = pdf.GetPage(1).Letters.ToArray();
        var second = pdf.GetPage(2).Letters.ToArray();
        Assert.Equal(first.Length, second.Length);
        Assert.NotEmpty(first);
        for (var index = 0; index < first.Length; index++)
        {
            Assert.InRange(Math.Abs(first[index].StartBaseLine.X - second[index].StartBaseLine.X),
                0, 0.1);
            Assert.InRange(Math.Abs(first[index].StartBaseLine.Y - second[index].StartBaseLine.Y),
                0, 0.1);
        }
    }

    [Fact]
    public void KeepsCompactArialUnderlineGeometryAfterPageBreak()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(
                new RunProperties(new RunFonts { Ascii = "Arial" }, new FontSize { Val = "18" },
                    new Underline { Val = UnderlineValues.Single, Color = "FF0000" }),
                new Text("Compact"), new Break { Type = BreakValues.Page },
                new Text("Compact")))));
        }

        var result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(result), PdfDocumentOpenMode.Import);
        Assert.Equal(2, pdf.PageCount);
        var rectangles = Enumerable.Range(0, pdf.PageCount).Select(index => pdf.Pages[index].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Where(reference => reference.Value is PdfSharp.Pdf.Advanced.PdfContent)
            .Select(reference => (PdfSharp.Pdf.Advanced.PdfContent)reference.Value)
            .Where(content => content.Stream != null)
            .SelectMany(content => System.Text.RegularExpressions.Regex.Matches(
                System.Text.Encoding.Latin1.GetString(content.Stream!.UnfilteredValue),
                @"(-?[\d.]+) (-?[\d.]+) (-?[\d.]+) (-?[\d.]+) re")
                .Select(match => match.Value)).ToArray()).ToArray();
        var first = Assert.Single(rectangles[0]).Split(' ');
        var second = Assert.Single(rectangles[1]).Split(' ');
        Assert.Equal(first[3], second[3]);
        Assert.Equal(first[1], second[1]);
    }

    [Fact]
    public void KeepsDirectParagraphSideBordersAcrossPageBreak()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        ParagraphProperties Bordered() => new(new ParagraphBorders(
            new LeftBorder { Val = BorderValues.Single, Color = "FF0000", Size = 8 },
            new RightBorder { Val = BorderValues.Single, Color = "FF0000", Size = 8 }),
            new SpacingBetweenLines { After = "120" });
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(Bordered(), new Run(new Text("Before"),
                    new Break { Type = BreakValues.Page }, new Text("After"))),
                new Paragraph(Bordered(), new Run(new Text("Next")))));
        }

        var result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(result), PdfDocumentOpenMode.Import);
        Assert.Equal(2, pdf.PageCount);
        var streams = pdf.Pages[1].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Where(reference => reference.Value is PdfSharp.Pdf.Advanced.PdfContent)
            .Select(reference => (PdfSharp.Pdf.Advanced.PdfContent)reference.Value)
            .Where(content => content.Stream != null)
            .Select(content => System.Text.Encoding.Latin1.GetString(content.Stream!.UnfilteredValue));
        var strokes = streams.SelectMany(stream => System.Text.RegularExpressions.Regex.Matches(stream,
            @"(?m)(-?[\d.]+) (-?[\d.]+) m\s+(-?[\d.]+) (-?[\d.]+) l\s+S")
            .Select(match => match.Value)).ToArray();
        var bridges = strokes.Select(stroke => stroke.Split([' ', '\r', '\n'],
                StringSplitOptions.RemoveEmptyEntries))
            .Where(parts => parts[0] == parts[3] &&
                Math.Abs(double.Parse(parts[4], System.Globalization.CultureInfo.InvariantCulture) -
                    double.Parse(parts[1], System.Globalization.CultureInfo.InvariantCulture)) is > 5.9 and < 6.1)
            .ToArray();
        Assert.Equal(2, bridges.Length);
    }

    [Fact]
    public void KeepsFooterShadingClearanceAfterPageBreak()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            styles.Styles = new Styles(new Style(new StyleRunProperties(new Border
                { Val = BorderValues.Single, Color = "FF0000", Size = 8, Space = 4 }))
            {
                Type = StyleValues.Paragraph,
                StyleId = "ShadedFooter"
            });
            var footer = main.AddNewPart<FooterPart>();
            footer.Footer = new Footer(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "ShadedFooter" },
                    new Shading { Fill = "00FF00" }),
                new Run(new Text("Footer A"), new Break { Type = BreakValues.Page },
                    new Text("Footer B"))));
            main.Document = new Document(new Body(new Paragraph(new Run(new Text("Body"))),
                new SectionProperties(new FooterReference
                    { Id = main.GetIdOfPart(footer), Type = HeaderFooterValues.Default })));
        }

        var result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(result), PdfDocumentOpenMode.Import);
        var rectangles = Enumerable.Range(0, pdf.PageCount).Select(index => pdf.Pages[index].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Where(reference => reference.Value is PdfSharp.Pdf.Advanced.PdfContent)
            .Select(reference => (PdfSharp.Pdf.Advanced.PdfContent)reference.Value)
            .Where(content => content.Stream != null)
            .SelectMany(content => System.Text.RegularExpressions.Regex.Matches(
                System.Text.Encoding.Latin1.GetString(content.Stream!.UnfilteredValue),
                @"(-?[\d.]+) (-?[\d.]+) (-?[\d.]+) (-?[\d.]+) re")
                .Select(match => match.Value)).ToArray()).ToArray();
        var shaded = rectangles.SelectMany(page => page)
            .Count(rectangle => double.Parse(rectangle.Split(' ')[2],
                System.Globalization.CultureInfo.InvariantCulture) > 400);
        Assert.Equal(2, shaded);
    }

    [Fact]
    public void KeepsFooterDoubleRunBorderOffsetAfterPageBreak()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var footer = main.AddNewPart<FooterPart>();
            footer.Footer = new Footer(new Paragraph(new Run(
                new RunProperties(new Border { Val = BorderValues.Double, Color = "FF0000",
                    Size = 8, Space = 4 }),
                new Text("Footer A"), new Break { Type = BreakValues.Page },
                new Text("Footer B"))));
            main.Document = new Document(new Body(new Paragraph(new Run(new Text("Body"))),
                new SectionProperties(new FooterReference
                    { Id = main.GetIdOfPart(footer), Type = HeaderFooterValues.Default })));
        }

        var result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(result), PdfDocumentOpenMode.Import);
        var rectangles = Enumerable.Range(0, pdf.PageCount).Select(index => pdf.Pages[index].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Where(reference => reference.Value is PdfSharp.Pdf.Advanced.PdfContent)
            .Select(reference => (PdfSharp.Pdf.Advanced.PdfContent)reference.Value)
            .Where(content => content.Stream != null)
            .SelectMany(content => System.Text.RegularExpressions.Regex.Matches(
                System.Text.Encoding.Latin1.GetString(content.Stream!.UnfilteredValue),
                @"(-?[\d.]+) (-?[\d.]+) (-?[\d.]+) (-?[\d.]+) re")
                .Select(match => match.Value)).ToArray()).ToArray();
        var outerBorders = Assert.Single(rectangles).Where((_, index) => index % 2 == 0).ToArray();
        Assert.Equal(2, outerBorders.Length);
        Assert.Equal(outerBorders[0].Split(' ')[1], outerBorders[1].Split(' ')[1]);
    }

    [Fact]
    public void KeepsBorderedTabInsetAfterPageBreak()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new ParagraphProperties(new Tabs(new TabStop
                    { Position = 3000, Val = TabStopValues.Left })),
                new Run(new RunProperties(new Border { Val = BorderValues.Single, Color = "FF0000",
                    Size = 8, Space = 4 }),
                    new Text("Alpha"), new TabChar(), new Text("Beta"),
                    new Break { Type = BreakValues.Page },
                    new Text("Gamma"), new TabChar(), new Text("Delta")))));
        }

        var result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(result), PdfDocumentOpenMode.Import);
        var rectangles = Enumerable.Range(0, pdf.PageCount).Select(index => pdf.Pages[index].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Where(reference => reference.Value is PdfSharp.Pdf.Advanced.PdfContent)
            .Select(reference => (PdfSharp.Pdf.Advanced.PdfContent)reference.Value)
            .Where(content => content.Stream != null)
            .SelectMany(content => System.Text.RegularExpressions.Regex.Matches(
                System.Text.Encoding.Latin1.GetString(content.Stream!.UnfilteredValue),
                @"(-?[\d.]+) (-?[\d.]+) (-?[\d.]+) (-?[\d.]+) re")
                .Select(match => match.Value)).ToArray()).ToArray();
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        var gaps = rectangles.Select((page, index) =>
        {
            var rectangle = page.Single().Split(' ');
            var right = double.Parse(rectangle[0], System.Globalization.CultureInfo.InvariantCulture) +
                double.Parse(rectangle[2], System.Globalization.CultureInfo.InvariantCulture);
            var last = parsed.GetPage(index + 1).Letters.Last();
            return right - last.EndBaseLine.X;
        }).ToArray();
        Assert.Equal(2, gaps.Length);
        Assert.InRange(gaps[0], 14, 15);
        Assert.InRange(Math.Abs(gaps[0] - gaps[1]), 0, 0.1);
    }

    [Fact]
    public void DoesNotRepeatFirstBodyDoubleBorderInsetAfterPageBreak()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(
                new RunProperties(new Border { Val = BorderValues.Double, Color = "FF0000",
                    Size = 8, Space = 4 }),
                new Text("Same text"), new Break { Type = BreakValues.Page },
                new Text("Same text")))));
        }

        var result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(result), PdfDocumentOpenMode.Import);
        Assert.Equal(2, pdf.PageCount);
        var rectangles = Enumerable.Range(0, pdf.PageCount).Select(index => pdf.Pages[index].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Where(reference => reference.Value is PdfSharp.Pdf.Advanced.PdfContent)
            .Select(reference => (PdfSharp.Pdf.Advanced.PdfContent)reference.Value)
            .Where(content => content.Stream != null)
            .SelectMany(content => System.Text.RegularExpressions.Regex.Matches(
                System.Text.Encoding.Latin1.GetString(content.Stream!.UnfilteredValue),
                @"(-?[\d.]+) (-?[\d.]+) (-?[\d.]+) (-?[\d.]+) re")
                .Select(match => match.Value)).ToArray()).ToArray();
        Assert.Equal(2, rectangles[0].Length);
        Assert.Equal(rectangles[0], rectangles[1]);
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
    public void KeepsNamedDestinationsWhenCombiningMirroredPages()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new ParagraphProperties(new MirrorIndents()),
                    new BookmarkStart { Name = "oddDestination", Id = "1" },
                    new Run(new Text("First page")), new BookmarkEnd { Id = "1" },
                    new Hyperlink(new Run(new Text("Next"))) { Anchor = "evenDestination" }),
                new Paragraph(new Run(new Break { Type = BreakValues.Page }),
                    new BookmarkStart { Name = "evenDestination", Id = "2" },
                    new Run(new Text("Second page")), new BookmarkEnd { Id = "2" },
                    new Hyperlink(new Run(new Text("Return"))) { Anchor = "oddDestination" })));
        }

        var result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(result), PdfDocumentOpenMode.Import);
        Assert.Equal(2, pdf.PageCount);
        var names = pdf.Internals.Catalog.Names!.NameTree!.GetNames(true);
        Assert.Contains("oddDestination", names);
        Assert.Contains("evenDestination", names);
    }

    [Fact]
    public void ClipsExactLineImageWhenCombiningMirroredPages()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "legacy", "WordImageStories.docx");
        using var docx = new MemoryStream();
        docx.Write(File.ReadAllBytes(source));
        docx.Position = 0;
        using (var word = WordprocessingDocument.Open(docx, true))
        {
            var body = word.MainDocumentPart!.Document.Body!;
            var imageParagraph = body.Descendants<Paragraph>()
                .First(paragraph => paragraph.Descendants<Drawing>().Any());
            var properties = imageParagraph.ParagraphProperties ??
                imageParagraph.PrependChild(new ParagraphProperties());
            properties.Append(new MirrorIndents(),
                new SpacingBetweenLines { Line = "200", LineRule = LineSpacingRuleValues.Exact });
            var nextPage = new Paragraph(
                new ParagraphProperties(new PageBreakBefore(), new MirrorIndents(),
                    new SpacingBetweenLines { Line = "200", LineRule = LineSpacingRuleValues.Exact }),
                new Run(imageParagraph.Descendants<Drawing>().First().CloneNode(true)),
                new Run(new Text("Following page")));
            body.InsertBefore(nextPage, body.GetFirstChild<SectionProperties>());
            word.MainDocumentPart.Document.Save();
        }

        var result = DxpExport.ExportToBytes(docx.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(result), PdfDocumentOpenMode.Import);
        Assert.Equal(2, pdf.PageCount);
        foreach (PdfPage page in pdf.Pages)
            Assert.Contains(page.Contents.Elements.OfType<PdfSharp.Pdf.Advanced.PdfReference>()
                .Where(reference => reference.Value is PdfSharp.Pdf.Advanced.PdfContent)
                .Select(reference => (PdfSharp.Pdf.Advanced.PdfContent)reference.Value)
                .Where(content => content.Stream != null)
                .Select(content => System.Text.Encoding.Latin1.GetString(content.Stream!.UnfilteredValue)),
                commands => commands.Contains(" re W n", StringComparison.Ordinal));
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
    public void PreservesLargeImageCropInsteadOfDroppingTransform()
    {
        var pixels = new byte[4200 * 4];
        for (var column = 0; column < 4200; column++)
        {
            pixels[column * 4] = column < 2100 ? (byte)255 : (byte)0;
            pixels[column * 4 + 2] = column < 2100 ? (byte)0 : (byte)255;
            pixels[column * 4 + 3] = 255;
        }
        using var stream = new MemoryStream();
        new StbImageWriteSharp.ImageWriter().WritePng(pixels, 4200, 1,
            StbImageWriteSharp.ColorComponents.RedGreenBlueAlpha, stream);

        var transformed = DxpRasterImageTransform.Transform(stream.ToArray(), 420, 1, 0, false, false,
            new DxpImageCrop(0.01, 0, 0, 0));

        Assert.NotNull(transformed);
        var image = StbImageSharp.ImageResult.FromMemory(transformed.Value.Image,
            StbImageSharp.ColorComponents.RedGreenBlueAlpha);
        Assert.Equal(4096, image.Width);
        Assert.Equal(1, image.Height);
        Assert.True(image.Data[0] > image.Data[2]);
        Assert.True(image.Data[^2] > image.Data[^4]);
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
    public void RotatesBmpFallbackInDrawingMlDirection()
    {
        var bitmap = new byte[62];
        bitmap[0] = (byte)'B';
        bitmap[1] = (byte)'M';
        BinaryPrimitives.WriteInt32LittleEndian(bitmap.AsSpan(2), bitmap.Length);
        BinaryPrimitives.WriteInt32LittleEndian(bitmap.AsSpan(10), 54);
        BinaryPrimitives.WriteInt32LittleEndian(bitmap.AsSpan(14), 40);
        BinaryPrimitives.WriteInt32LittleEndian(bitmap.AsSpan(18), 2);
        BinaryPrimitives.WriteInt32LittleEndian(bitmap.AsSpan(22), -1);
        BinaryPrimitives.WriteUInt16LittleEndian(bitmap.AsSpan(26), 1);
        BinaryPrimitives.WriteUInt16LittleEndian(bitmap.AsSpan(28), 24);
        bitmap[54] = 1;
        bitmap[57] = 2;

        var clockwise = DxpBmp24Converter.Transform(bitmap, 90, false, false);
        var counterclockwise = DxpBmp24Converter.Transform(bitmap, -90, false, false);

        Assert.Equal(1, BinaryPrimitives.ReadInt32LittleEndian(clockwise.AsSpan(18)));
        Assert.Equal(-2, BinaryPrimitives.ReadInt32LittleEndian(clockwise.AsSpan(22)));
        Assert.Equal((byte)1, clockwise[54]);
        Assert.Equal((byte)2, clockwise[58]);
        Assert.Equal((byte)2, counterclockwise[54]);
        Assert.Equal((byte)1, counterclockwise[58]);
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

    [Fact]
    public void RendersInlineDrawingTextBoxInAFrame()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "features", "text-boxes",
            "TextBoxSimple.docx");
        using var docx = new MemoryStream();
        docx.Write(File.ReadAllBytes(source));
        docx.Position = 0;
        using (var word = WordprocessingDocument.Open(docx, true))
        {
            var drawing = word.MainDocumentPart!.Document.Descendants<Drawing>().First();
            var anchor = drawing.ChildElements.First(element => element.LocalName == "anchor");
            var inline = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline();
            foreach (var child in anchor.ChildElements)
                if (child.LocalName is "extent" or "effectExtent" or "docPr" or
                    "cNvGraphicFramePr" or "graphic")
                    inline.Append(child.CloneNode(true));
            drawing.ReplaceChild(inline, anchor);
            word.MainDocumentPart.Document.Save();
        }

        var visitor = new DxpFlowPdfVisitor();
        var result = DxpExport.ExportToBytes(docx.ToArray(), visitor);
        var field = typeof(DxpFlowPdfVisitor).GetField("_document",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!;
        var rendered = (MigraDoc.DocumentObjectModel.Document)field.GetValue(visitor)!;

        Assert.Contains(rendered.Sections.Cast<MigraDoc.DocumentObjectModel.Section>()
            .SelectMany(section => section.Elements.OfType<MigraDoc.DocumentObjectModel.Shapes.TextFrame>()),
            frame => frame.Elements.OfType<MigraDoc.DocumentObjectModel.Paragraph>().Any());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Contains(pdf.GetPages(), page => page.Letters.Any());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void RendersAltChunkInsideTableCell(bool html)
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var chunkPart = main.AddAlternativeFormatImportPart(html
                ? AlternativeFormatImportPartType.Html
                : AlternativeFormatImportPartType.TextPlain, "rIdCellChunk");
            using var content = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(
                html ? "<html><body><p>Imported cell</p></body></html>" : "Imported cell"));
            chunkPart.FeedData(content);
            main.Document = new Document(new Body(new Table(new TableRow(new TableCell(
                new Paragraph(new Run(new Text("Before"))),
                new AltChunk { Id = "rIdCellChunk" },
                new Paragraph(new Run(new Text("After"))))))));
        }

        var result = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Contains("Imported", string.Concat(pdf.GetPages().Select(page => page.Text)));
        Assert.Contains("After", string.Concat(pdf.GetPages().Select(page => page.Text)));
    }

    [Theory]
    [InlineData("shared-hyperlink-carriage-return.docx")]
    [InlineData("shared-repeated-highlight-stories.docx")]
    [InlineData("shared-altchunk-charset.docx")]
    [InlineData("shared-altchunk-table-cell-html.docx")]
    [InlineData("shared-altchunk-table-cell-text.docx")]
    [InlineData("shared-line-number-continuous-sections.docx")]
    public void RendersSharedExportProbes(string fileName)
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "probes", fileName);
        using (var document = WordprocessingDocument.Open(source, false))
            Assert.NotNull(document.MainDocumentPart?.Document?.Body);

        byte[] result = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Contains(pdf.GetPages(), page => page.Letters.Any());
    }
}
