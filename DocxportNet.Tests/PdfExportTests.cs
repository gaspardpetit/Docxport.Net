using DocumentFormat.OpenXml;
using System.Buffers.Binary;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Fields;
using DocxportNet.Fields.Resolution;
using DocxportNet.Visitors.Pdf;
using DocxportNet.Walker.Context;
using PdfSharp.Fonts;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;
using MathXml = DocumentFormat.OpenXml.Math;

namespace DocxportNet.Tests;

public class PdfExportTests
{
    [Fact]
    public void ImplicitStylelessAptosLineRuleRequiresPlainDirectText()
    {
        var paragraph = new Paragraph(new Run(
            new RunProperties(new RunFonts { Ascii = "Aptos" }, new FontSize { Val = "24" }),
            new Text("Aptos body text")));

        Assert.True(DxpFlowPdfVisitor.HasImplicitPlainAptos12Runs(paragraph));
        Assert.False(DxpFlowPdfVisitor.HasImplicitPlainAptos12Runs(new Paragraph()));

        var withDrawing = (Paragraph)paragraph.CloneNode(true);
        withDrawing.Append(new Run(new Drawing()));
        Assert.False(DxpFlowPdfVisitor.HasImplicitPlainAptos12Runs(withDrawing));

        var withAfterSpacing = (Paragraph)paragraph.CloneNode(true);
        withAfterSpacing.PrependChild(new ParagraphProperties(
            new SpacingBetweenLines { After = "0" }));
        Assert.False(DxpFlowPdfVisitor.HasImplicitPlainAptos12Runs(withAfterSpacing));

        var wrongFont = (Paragraph)paragraph.CloneNode(true);
        wrongFont.GetFirstChild<Run>()!.RunProperties!.RunFonts!.Ascii = "Arial";
        Assert.False(DxpFlowPdfVisitor.HasImplicitPlainAptos12Runs(wrongFont));
    }

    [Fact]
    public void FlowVisitorUsesWordPitchForImplicitStylelessAptosText()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            "repeated-space-probes__docx__ln-space-1-width-85.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var baselines = pdf.GetPage(1).Letters
            .Select(letter => Math.Round(letter.StartBaseLine.Y, 2))
            .Distinct().OrderByDescending(baseline => baseline).ToArray();
        Assert.True(baselines.Length >= 2);
        Assert.InRange(baselines[0] - baselines[1], 16.9, 17.05);
    }

    [Fact]
    public void FlowVisitorTrimsRepeatedSpaceAtWrappedWord()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            "repeated-space-alpha-probes__docx__ln-space-2-width-45.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        var singleSpaceBytes = DxpExport.ExportToBytes(
            source.Replace("ln-space-2", "ln-space-1", StringComparison.Ordinal),
            new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        using var singleSpacePdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(singleSpaceBytes));
        var alpha = Assert.Single(pdf.GetPage(1).GetWords().Where(word => word.Text == "Alpha"));
        var singleSpaceAlpha = Assert.Single(singleSpacePdf.GetPage(1).GetWords()
            .Where(word => word.Text == "Alpha"));
        var extraSpaceAdvance = alpha.BoundingBox.Left - singleSpaceAlpha.BoundingBox.Left;
        Assert.InRange(extraSpaceAdvance, -0.1, 0.1);
    }

    [Theory]
    [InlineData("repeated-space-probes", 1)]
    [InlineData("repeated-space-probes", 2)]
    [InlineData("repeated-space-styleless-probes", 1)]
    [InlineData("repeated-space-styleless-probes", 2)]
    public void NarrowEmergencyBreakStartsOversizedWordAfterLeadingMark(string family, int spaces)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            $"{family}__docx__ln-space-{spaces}-width-45.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var lines = Assert.Single(pdf.GetPages()).Letters
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .OrderByDescending(group => group.Key)
            .Select(group => string.Concat(group.OrderBy(letter => letter.GlyphRectangle.Left)
                .Select(letter => letter.Value)).Replace(" ", "", StringComparison.Ordinal))
            .ToArray();
        Assert.True(lines.Any(line => line.Contains("LN(10)!", StringComparison.Ordinal)),
            string.Join(" | ", lines));
        Assert.Contains(lines, line => line == "Syntax");
        Assert.DoesNotContain(lines, line => line.Contains("!S", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(35, false)]
    [InlineData(36, true)]
    public void FlowVisitorBreaksOversizeClosingPunctuationAtWordBoundary(int width,
        bool staysWithWord)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            $"punctuation-arial9-probes__closing-{width}pt.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var lines = pdf.GetPage(1).Letters
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Select(group => string.Concat(group.OrderBy(letter => letter.GlyphRectangle.Left)
                .Select(letter => letter.Value)))
            .ToArray();
        var quoteLine = Assert.Single(lines.Where(line => line.Contains('"')));
        Assert.Equal(staysWithWord, quoteLine.Contains("authors.", StringComparison.Ordinal));
        foreach (var mark in new[] { ')', ']', '}', '!' })
        {
            var markLine = Assert.Single(lines.Where(line => line.Contains(mark)));
            Assert.Equal(staysWithWord, markLine.Contains("authors.", StringComparison.Ordinal));
        }
    }

    [Fact]
    public void FlowVisitorBreaksOversizeFinalQuoteFromThePrecedingWord()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            "punctuation-arial18-probes__closing-75pt.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var lines = pdf.GetPage(1).Letters
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Select(group => string.Concat(group.OrderBy(letter => letter.GlyphRectangle.Left)
                .Select(letter => letter.Value)))
            .ToArray();
        var markLine = Assert.Single(lines.Where(line => line.Contains('\u00bb')));
        Assert.DoesNotContain("authors.", markLine);
    }

    [Fact]
    public void FlowVisitorUsesWordLinePitchForStylelessArialParagraphs()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            "punctuation-arial18-probes__closing-80pt.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var baselines = pdf.GetPage(1).Letters
            .Select(letter => Math.Round(letter.StartBaseLine.Y, 2))
            .Distinct()
            .OrderByDescending(value => value)
            .Take(2)
            .ToArray();
        Assert.Equal(2, baselines.Length);
        Assert.InRange(baselines[0] - baselines[1], 23.94, 24.02);
    }

    [Theory]
    [InlineData("around")]
    [InlineData("notBeside")]
    public void FlowVisitorPositionsAndWrapsParagraphFrame(string wrap)
    {
        if (!Enum.TryParse<MigraDoc.DocumentObjectModel.Shapes.WrapStyle>("BothSides", out _)) return;
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var frameProperties = new FrameProperties();
            foreach (var (name, value) in new[]
            {
                ("w", "2880"), ("h", "1440"), ("x", "2880"), ("y", "720"),
                ("hAnchor", "margin"), ("vAnchor", "margin"), ("wrap", wrap)
            })
                frameProperties.SetAttribute(new OpenXmlAttribute("w", name,
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main", value));
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Before positioned frame"))),
                new Paragraph(new ParagraphProperties(frameProperties),
                    new Run(new Text("Positioned frame text"))),
                new Paragraph(new Run(new Text(string.Concat(Enumerable.Repeat(
                    "This paragraph must flow beside the positioned frame. ", 15))))),
                new SectionProperties(new PageSize { Width = 12240, Height = 15840 },
                    new PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 })));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var words = page.GetWords().ToArray();
        Assert.Contains(words, word => word.Text == "Positioned" &&
            word.BoundingBox.Left is > 215 and < 217);
        var firstBeside = words.First(word => word.Text == "beside");
        if (wrap == "around")
        {
            var paragraphTops = words.Where(word => word.Text == "paragraph")
                .Select(word => Math.Round(word.BoundingBox.Top, 2)).Distinct()
                .OrderByDescending(top => top).ToArray();
            Assert.InRange(paragraphTops[0] - paragraphTops[1], 14.64, 14.67);
        }
        Assert.InRange(firstBeside.BoundingBox.Left,
            wrap == "around" ? 360.8 : 197, wrap == "around" ? 361.2 : 200);
        if (wrap == "notBeside")
        {
            var frameWord = words.Single(word => word.Text == "Positioned");
            var firstBodyWord = words.First(word => word.Text == "This");
            Assert.InRange(frameWord.BoundingBox.Top - firstBodyWord.BoundingBox.Top, 65, 80);
        }
    }

    [Fact]
    public void FlowVisitorContinuesPositionedTableAcrossPagesWhenSupported()
    {
        if (!Enum.TryParse<MigraDoc.DocumentObjectModel.Shapes.WrapStyle>("BothSides", out _) ||
            typeof(MigraDoc.Rendering.DocumentRenderer).GetProperty("PaginateTextFrame") == null)
            return;

        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var position = new TablePositionProperties();
            foreach (var (name, value) in new[]
            {
                ("leftFromText", "120"), ("rightFromText", "120"),
                ("vertAnchor", "margin"), ("horzAnchor", "margin"),
                ("tblpX", "2880"), ("tblpY", "720")
            })
                position.SetAttribute(new OpenXmlAttribute("w", name,
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main", value));
            var borders = new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 8 },
                new LeftBorder { Val = BorderValues.Single, Size = 8 },
                new BottomBorder { Val = BorderValues.Single, Size = 8 },
                new RightBorder { Val = BorderValues.Single, Size = 8 });
            var table = new Table(
                new TableProperties(
                    new TableWidth { Width = "2880", Type = TableWidthUnitValues.Dxa },
                    position, borders),
                new TableGrid(new GridColumn { Width = "2880" }));
            for (var rowNumber = 1; rowNumber <= 35; rowNumber++)
                table.Append(new TableRow(new TableCell(
                    new TableCellProperties(new TableCellWidth
                        { Width = "2880", Type = TableWidthUnitValues.Dxa }),
                    new Paragraph(new Run(new Text($"Floating cell {rowNumber}"))))));
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Before floating table"))),
                table,
                new Paragraph(new Run(new Text("After floating table " +
                    string.Concat(Enumerable.Repeat(
                        "This paragraph must flow beside the positioned table. ", 16))))),
                new SectionProperties(
                    new PageSize { Width = 12240, Height = 15840 },
                    new PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440,
                        Header = 720, Footer = 720 })));
            main.Document.Save();
        }

        var previousSetting = Environment.GetEnvironmentVariable("DOCXPORT_PAGINATE_FLOATING_TABLES");
        byte[] bytes;
        try
        {
            Environment.SetEnvironmentVariable("DOCXPORT_PAGINATE_FLOATING_TABLES", null);
            bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        }
        finally
        {
            Environment.SetEnvironmentVariable("DOCXPORT_PAGINATE_FLOATING_TABLES", previousSetting);
        }

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        var firstPageWords = pdf.GetPage(1).GetWords().ToArray();
        var secondPageWords = pdf.GetPage(2).GetWords().ToArray();
        Assert.Contains(firstPageWords, word => word.Text == "1");
        Assert.DoesNotContain(firstPageWords, word => word.Text == "35");
        Assert.Contains(secondPageWords, word => word.Text == "28");
        Assert.Contains(secondPageWords, word => word.Text == "35");
        Assert.Contains(secondPageWords, word => word.Text == "After");
        var firstContinuation = secondPageWords.Where(word => int.TryParse(word.Text, out _))
            .OrderBy(word => int.Parse(word.Text)).First();
        Assert.InRange(pdf.GetPage(2).Height - firstContinuation.BoundingBox.Top, 75.3, 75.6);
        Assert.Contains(pdf.GetPage(2).Paths, path =>
            path.IsStroked && path.GetBoundingRectangle() is { } bounds &&
            Math.Abs(bounds.Left - 215) < 0.1 && Math.Abs(bounds.Right - 360) < 0.1 &&
            Math.Abs(pdf.GetPage(2).Height - bounds.Top - 72.5) < 0.1);
        var rowNumbers = firstPageWords.Concat(secondPageWords)
            .Select(word => word.Text)
            .Where(text => int.TryParse(text, out _))
            .Select(int.Parse)
            .OrderBy(number => number);
        Assert.Equal(Enumerable.Range(1, 35), rowNumbers);
    }

    [Fact]
    public void FlowClosingPunctuationOverhangDependsOnParagraphAlignment()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var font = new XFont("Arial", 12);
        var paragraph = new MigraDoc.DocumentObjectModel.Paragraph();

        paragraph.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Justify;
        var justified = DxpFlowPdfVisitor.GetLineEndOverhang(paragraph, "authors.\u201d", font);
        Assert.NotNull(justified);
        Assert.True(justified > 0);

        paragraph.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Left;
        var left = DxpFlowPdfVisitor.GetLineEndOverhang(paragraph, "authors.\u201d", font);
        Assert.NotNull(left);
        Assert.InRange(left.Value, 0, font.Size / 3.0);
        Assert.True(left <= justified);
        var twoWord = DxpFlowPdfVisitor.GetLineEndOverhang(paragraph,
            "authors.\u201d", font, twoWordParagraph: true);
        Assert.NotNull(twoWord);
        Assert.True(twoWord >= left);
        Assert.InRange(twoWord.Value, 0, font.Size * 0.6);
        Assert.Null(DxpFlowPdfVisitor.GetLineEndOverhang(paragraph,
            "authors.\"", font, twoWordParagraph: true));
        paragraph.Format.RightIndent = MigraDoc.DocumentObjectModel.Unit.FromPoint(10);
        Assert.Null(DxpFlowPdfVisitor.GetLineEndOverhang(paragraph,
            "authors.\u2019", font, twoWordParagraph: true));
        paragraph.Format.RightIndent = MigraDoc.DocumentObjectModel.Unit.Zero;
        Assert.Equal(font.Size * 1.05,
            DxpFlowPdfVisitor.GetLineEndOverhang(paragraph, "\u300d", font));

        paragraph.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Center;
        Assert.Null(DxpFlowPdfVisitor.GetLineEndOverhang(paragraph, "authors.\u201d", font));
        Assert.Null(DxpFlowPdfVisitor.GetLineEndOverhang(paragraph, "\u300d", font));
        paragraph.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Justify;
        Assert.Null(DxpFlowPdfVisitor.GetLineEndOverhang(paragraph, "authors.)", font));
    }

    [Fact]
    public void PairedClosingMarkOverhangUsesMeasuredPunctuationWidths()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var font = new XFont("Arial", 12);
        using var graphics = XGraphics.CreateMeasureContext(new XSize(100, 100),
            XGraphicsUnit.Point, XPageDirection.Downwards);
        var halfHyphen = graphics.MeasureString("-", font).Width / 2;

        Assert.Equal(halfHyphen,
            DxpFlowPdfVisitor.GetPairedClosingMarkOverhang("authors-", font));
        Assert.Equal(graphics.MeasureString("\u201D", font).Width + halfHyphen,
            DxpFlowPdfVisitor.GetPairedClosingMarkOverhang("\u201D", font));
        foreach (var mark in new[] { ';', ':', '!', '/' })
            Assert.Equal(graphics.MeasureString($"{mark}\u201D", font).Width,
                DxpFlowPdfVisitor.GetPairedClosingMarkOverhang($"authors{mark}\u201D", font));
        Assert.Null(DxpFlowPdfVisitor.GetPairedClosingMarkOverhang("authors", font));
    }

    [Fact]
    public void FlowCalibriClosingPunctuationToleranceCoversCloseCategory()
    {
        const MigraDoc.DocumentObjectModel.ParagraphAlignment left =
            MigraDoc.DocumentObjectModel.ParagraphAlignment.Left;

        foreach (var mark in new[] { ')', ']', '}' })
            Assert.Equal(0.2, DxpFlowPdfVisitor.GetClosePunctuationTolerance(left,
                $"authors.{mark}", "Calibri", 12, true));

        Assert.Null(DxpFlowPdfVisitor.GetClosePunctuationTolerance(left,
            "authors.}", "Calibri", 12, false));
        Assert.Null(DxpFlowPdfVisitor.GetClosePunctuationTolerance(left,
            "authors.}", "Arial", 12, true));
        Assert.Null(DxpFlowPdfVisitor.GetClosePunctuationTolerance(left,
            "authors.\u300d", "Calibri", 12, true));
        Assert.Null(DxpFlowPdfVisitor.GetClosePunctuationTolerance(
            MigraDoc.DocumentObjectModel.ParagraphAlignment.Justify,
            "authors.}", "Calibri", 12, true));
    }

    [Fact]
    public void FlowArialHangingTerminalSpaceRequiresTwoWordClosingMark()
    {
        const MigraDoc.DocumentObjectModel.ParagraphAlignment left =
            MigraDoc.DocumentObjectModel.ParagraphAlignment.Left;

        Assert.Equal(0.25, DxpFlowPdfVisitor.GetHangingTerminalLineSpaceThreshold(
            left, true, "authors.\u201d", "Arial"));
        Assert.Null(DxpFlowPdfVisitor.GetHangingTerminalLineSpaceThreshold(
            left, false, "authors.\u201d", "Arial"));
        Assert.Null(DxpFlowPdfVisitor.GetHangingTerminalLineSpaceThreshold(
            left, true, "authors.)", "Arial"));
        Assert.Null(DxpFlowPdfVisitor.GetHangingTerminalLineSpaceThreshold(
            MigraDoc.DocumentObjectModel.ParagraphAlignment.Justify,
            true, "authors.\u201d", "Arial"));
    }

    [Fact]
    public void FlowNearEdgeHangingSpaceIsLimitedToArialClosingMarks()
    {
        const MigraDoc.DocumentObjectModel.ParagraphAlignment left =
            MigraDoc.DocumentObjectModel.ParagraphAlignment.Left;

        Assert.Equal(0.25, DxpFlowPdfVisitor.GetHangingTerminalLineNearEdgeThreshold(
            left, true, "authors.\u201d", "Arial"));
        Assert.Null(DxpFlowPdfVisitor.GetHangingTerminalLineNearEdgeThreshold(
            left, true, "authors.\u201d", "Aptos"));
        Assert.Null(DxpFlowPdfVisitor.GetHangingTerminalLineNearEdgeThreshold(
            left, true, "authors.\u201d", "Calibri"));
        Assert.Null(DxpFlowPdfVisitor.GetHangingTerminalLineNearEdgeThreshold(
            left, false, "authors.\u201d", "Arial"));
        Assert.Null(DxpFlowPdfVisitor.GetHangingTerminalLineNearEdgeThreshold(
            left, true, "authors.)", "Arial"));
    }

    [Fact]
    public void FlowVisitorRecognizesAsciiWordsOnlyAfterSourceRunBoundaries()
    {
        var split = new[]
        {
            new Run(new Text("Reference number ") { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new Text("A")),
            new Run(new Text("b")),
        };
        var single = new[] { new Run(new Text("Header A")) };

        Assert.Equal(new[] { "A" }, DxpFlowPdfVisitor.GetAsciiWordsAfterRunBoundary(split));
        Assert.Empty(DxpFlowPdfVisitor.GetAsciiWordsAfterRunBoundary(single));
    }

    [Fact]
    public void FlowVisitorAppliesLeadingKerningOnlyAcrossSourceRunsWhenSupported()
    {
        if (typeof(MigraDoc.Rendering.DocumentRenderer).GetProperty("LeadingKerningAfterBlank") == null)
            return;
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var styledSource = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            "styled-word-probes__styled-word-93pt.docx");
        var controlSource = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleBoldOnlyDefaultsAllStories.docx");

        using var styledPdf = UglyToad.PdfPig.PdfDocument.Open(
            DxpExport.ExportToBytes(styledSource, new DxpFlowPdfVisitor()));
        var lines = styledPdf.GetPage(1).Letters
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Select(group => string.Concat(group.OrderBy(letter => letter.GlyphRectangle.Left)
                .Select(letter => letter.Value))).ToArray();
        Assert.Contains(lines, line => line.Contains("ReferencenumberAbc", StringComparison.Ordinal));

        using var controlPdf = UglyToad.PdfPig.PdfDocument.Open(
            DxpExport.ExportToBytes(controlSource, new DxpFlowPdfVisitor()));
        var control = controlPdf.GetPage(1);
        var headerLetter = control.Letters.Single(letter => letter.Value == "A" &&
            letter.GlyphRectangle.Left is > 100 and < 120 &&
            control.Height - letter.StartBaseLine.Y is > 40 and < 70);
        Assert.InRange(headerLetter.StartBaseLine.X, 111.5, 111.6);
    }

    [Fact]
    public void FlowStylelessFullStopOverhangUsesGlyphAdvance()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var font = new XFont("Arial", 12);
        var allowance = DxpFlowPdfVisitor.GetStylelessFullStopOverhang("et.", font);

        Assert.NotNull(allowance);
        Assert.InRange(allowance.Value, 0.01, font.Size / 3.0);
        Assert.Null(DxpFlowPdfVisitor.GetStylelessFullStopOverhang("et,", font));
        Assert.Null(DxpFlowPdfVisitor.GetStylelessFullStopOverhang("word", font));
    }

    [Theory]
    [InlineData("日本語", "Word123", true, true)]
    [InlineData("日本語", "Word123", false, false)]
    [InlineData("日本語 ", "Word123", true, false)]
    [InlineData("日本語", "456", true, false)]
    [InlineData("Word123", "日本語", true, false)]
    [InlineData("ABCé", "東京", true, true)]
    [InlineData("ABC", "東京", true, true)]
    [InlineData("日本語", "élève", true, true)]
    [InlineData("שלום", "東京", true, false)]
    [InlineData("مرحبا", "東京", true, false)]
    [InlineData("ABCé", "東京", false, false)]
    [InlineData("", "Word123", true, false)]
    public void FlowAutoSpaceDeRecognizesCrossRunLetterBoundaries(
        string previous, string next, bool autoSpaceDE, bool expected)
    {
        Assert.Equal(expected, DxpFlowPdfVisitor.NeedsAutoSpaceAcrossRuns(previous, next, autoSpaceDE));
    }

    [Fact]
    public void FlowVisitorSpacesFittedLatinToCjkHeaderBoundary()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextCjkPunctuationFontsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var header = page.Letters.Where(letter => page.Height - letter.StartBaseLine.Y < 65).ToArray();
        var accented = header.Single(letter => letter.Value == "é");
        var japanese = header.SkipWhile(letter => !ReferenceEquals(letter, accented)).Skip(1).First();
        Assert.InRange(japanese.StartBaseLine.X - accented.GlyphRectangle.Right, 0.5, 2);
    }

    [Fact]
    public void FlowVisitorPaintsFittedMixedCjkBodyAtWordNominalSize()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextCjkPunctuationFontsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var body = page.Letters.Where(letter =>
            letter.StartBaseLine.X is > 71.9 and < 72.2 &&
            page.Height - letter.StartBaseLine.Y is > 75 and < 105).ToArray();
        Assert.NotEmpty(body);
        Assert.All(body, letter => Assert.InRange(letter.PointSize, 11.99, 12.01));
        Assert.InRange(body[0].StartBaseLine.X, 71.9, 72.2);
    }

    [Fact]
    public void FlowVisitorQuantizesFittedMixedCjkBodyAdvancesWithoutMovingFollowingParagraph()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextCjkPunctuationFontsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var following = page.Letters.First(letter => letter.Value == "S" &&
            page.Height - letter.StartBaseLine.Y is > 105 and < 130);
        Assert.InRange(following.StartBaseLine.X, 71.9, 72.1);

    }

    [Theory]
    [InlineData(6.75, 6.72)]
    [InlineData(8.1, 8.04)]
    [InlineData(10.125, 10.08)]
    public void FittedCjkAdvanceUsesWordPointTwelveGrid(double advance, double expected)
    {
        Assert.Equal(expected, DxpFlowPdfVisitor.QuantizeFittedCjkAdvance(advance), 4);
    }

    [Theory]
    [InlineData(61.3635, 56)]
    [InlineData(73.6362, 67)]
    [InlineData(92.0452, 84)]
    public void FittedCjkPaintScalePreservesQuantizedAdvance(double layoutScale, double expected)
    {
        Assert.Equal(expected, DxpFlowPdfVisitor.FittedCjkPaintScale(layoutScale), 4);
    }

    [Fact]
    public void SourceSpaceRestorerAlignsSplitFirstWordWithoutSplittingIt()
    {
        const string content = "(U) Tj\n8 0 Td\n(nder) Tj\n27 0 Td\n(the) Tj\n20 0 Td\n(new) Tj\n25 0 Td\n(section.) Tj";

        Assert.True(DxpPdfSourceSpaceRestorer.TryRestore(content, "Under the new section.",
            "Under", 0, out var revised, out var count, out var cursor));

        Assert.Equal(3, count);
        Assert.Equal("Under the new section.".Length, cursor);
        Assert.Contains("(U) Tj\n8 0 Td\n(nder) Tj\n( ) Tj\n27 0 Td\n(the) Tj", revised);
    }

    [Fact]
    public void SourceSpaceRestorerRejectsAmbiguousAnchors()
    {
        const string repeated = "(The) Tj\n12 0 Td\n(owners) Tj\n30 0 Td\n(repair) Tj\n25 0 Td\n(devices.) Tj";
        var content = repeated + "\n" + repeated;

        Assert.False(DxpPdfSourceSpaceRestorer.TryRestore(content, "The owners repair devices.",
            "The", 0, out var revised, out var count, out _));

        Assert.Equal(content, revised);
        Assert.Equal(0, count);
    }

    [Fact]
    public void SourceSpaceRestorerPreservesSpaceAcrossColorChange()
    {
        const string content = "(For) Tj\n20 0 Td\n(instance,) Tj\n.129 .145 .161 rg\n49.918 0 Td\n" +
            "(the) Tj\n20 0 Td\n(legislator) Tj";

        Assert.True(DxpPdfSourceSpaceRestorer.TryRestore(content, "For instance, the legislator",
            "For", 0, out var revised, out var count, out var cursor));

        Assert.Equal(3, count);
        Assert.Equal("For instance, the legislator".Length, cursor);
        Assert.Contains("(instance,) Tj\n( ) Tj\n.129 .145 .161 rg\n49.918 0 Td\n(the) Tj", revised);
    }

    [Fact]
    public void SourceSpaceRestorerPreservesSpacesAroundRaisedFootnoteMarker()
    {
        const string content = "(Quebec) Tj\n20 0 Td\n(became) Tj\n20 0 Td\n(the) Tj\n" +
            "20 0 Td\n(first) Tj\n20 0 Td\n(last) Tj\n20 0 Td\n(year.) Tj\n" +
            "/F0 6.769 Tf\n26.3145 3.6911 Td\n(29) Tj\n/F0 12 Tf\n" +
            "11.4687 -3.6911 Td\n(In) Tj\n20 0 Td\n(the) Tj";

        Assert.True(DxpPdfSourceSpaceRestorer.TryRestore(content,
            "Quebec became the first last year.  In the", "Quebec", 0,
            out var revised, out var count, out var cursor));

        Assert.Equal(8, count);
        Assert.Equal("Quebec became the first last year.  In the".Length, cursor);
        Assert.Contains("(year.) Tj\n( ) Tj\n/F0 6.769 Tf\n26.3145 3.6911 Td\n" +
            "(29) Tj\n( ) Tj\n/F0 12 Tf", revised);
    }

    [Fact]
    public void SourceSpaceRestorerPreservesSpaceAcrossUnderlinedTextBlock()
    {
        const string content = "(In) Tj\n20 0 Td\n(the) Tj\n20 0 Td\n(first) Tj\n" +
            "20 0 Td\n(instance,) Tj\nET\n72 519.3614 m\n115.2832 519.3614 l\n" +
            "S\nBT\n125.47 520.477 Td\n(which) Tj\n20 0 Td\n(suggests) Tj";

        Assert.True(DxpPdfSourceSpaceRestorer.TryRestore(content,
            "In the first instance, which suggests", "In", 0,
            out var revised, out var count, out var cursor));

        Assert.Equal(5, count);
        Assert.Equal("In the first instance, which suggests".Length, cursor);
        Assert.Contains("(instance,) Tj\n( ) Tj\nET\n72 519.3614 m", revised);
    }

    [Fact]
    public void FlowVisitorKeepsSuperscriptHighlightAtFullLineHeight()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "with_breaking_style.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var highlight = Assert.Single(page.Paths.Where(path => path.IsFilled &&
            path.FillColor != null && path.FillColor.ToRGBValues() is var color && color.r > 0.99 &&
            color.g > 0.99 && color.b < 0.01 &&
            path.GetBoundingRectangle() is { } bounds && bounds.Left is > 71 and < 73));
        var rectangle = highlight.GetBoundingRectangle()!.Value;
        Assert.InRange(rectangle.Right, 185, 195);
        Assert.InRange(page.Height - rectangle.Top, 89, 92);
        Assert.InRange(page.Height - rectangle.Bottom, 112, 116);
    }

    [Theory]
    [InlineData("WordHighlightPaletteAllStories")]
    [InlineData("WordShadeResetWithHighlightAllStories")]
    public void FlowVisitorAlignsSmallFooterHighlightToWordLineBox(string fixture)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixture + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var highlight = Assert.Single(page.Paths.Where(path => path.IsFilled &&
            path.FillColor?.ToRGBValues() is { } color && color.r < 0.01 &&
            color.g > 0.99 && color.b > 0.99 &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 71 and < 73 && page.Height - bounds.Top > 650));
        var rectangle = highlight.GetBoundingRectangle()!.Value;
        var footer = page.Letters.First(letter => letter.Value == "F" &&
            page.Height - letter.StartBaseLine.Y > 650);
        var baseline = page.Height - footer.StartBaseLine.Y;
        Assert.InRange(baseline - (page.Height - rectangle.Top), 12.1, 12.4);
        Assert.InRange(rectangle.Height, 14.5, 14.8);
    }

    [Fact]
    public void FlowVisitorPreservesIndentedBlankLineBeforeNumberedList()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Tracked.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var firstMarker = page.Letters.First(letter => letter.Value == "1" &&
            letter.GlyphRectangle.Left is > 90 and < 92 &&
            page.Height - letter.GlyphRectangle.Top is > 110 and < 140);
        Assert.InRange(page.Height - firstMarker.GlyphRectangle.Top, 123, 127);
    }

    [Fact]
    public void FlowVisitorCollapsesOnlyIdenticalThinStrokes()
    {
        var source = string.Join('\n',
            "q", "0.5 w", "0 0 m", "100 0 l", "S",
            "100 0 m", "0 0 l", "S",
            "1 0 0 RG", "0 0 m", "100 0 l", "S",
            "1 w", "0 0 m", "100 0 l", "S", "Q");
        var rewritten = DxpFlowPdfVisitor.RewriteDuplicateThinStrokes(source);

        Assert.NotNull(rewritten);
        Assert.Equal(3, System.Text.RegularExpressions.Regex.Matches(rewritten, @"(?m)^S$").Count);
        Assert.Contains("1 0 0 RG\n0 0 m\n100 0 l\nS", rewritten);
        Assert.Contains("1 w\n0 0 m\n100 0 l\nS", rewritten);
        Assert.Null(DxpFlowPdfVisitor.RewriteDuplicateThinStrokes(rewritten));
    }

    [Fact]
    public void FlowVisitorCollapsesQuarterPointStrokesWithoutMergingDifferentWidths()
    {
        var source = string.Join('\n',
            "0.25 w", "0 0 m", "100 0 l", "S", "100 0 m", "0 0 l", "S",
            "0.5 w", "0 0 m", "100 0 l", "S");
        var rewritten = DxpFlowPdfVisitor.RewriteDuplicateThinStrokes(source);

        Assert.NotNull(rewritten);
        Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(rewritten, @"(?m)^S$").Count);
        Assert.Contains("0.5 w\n0 0 m\n100 0 l\nS", rewritten);
    }

    [Theory]
    [InlineData(".", "”")]
    [InlineData(".", "’")]
    [InlineData(".", "\"")]
    [InlineData(".", ")")]
    [InlineData(".", "]")]
    [InlineData(".", "}")]
    [InlineData("?", "”")]
    [InlineData(";", "”")]
    [InlineData("/", "”")]
    [InlineData("-", "”")]
    [InlineData("+", "”")]
    [InlineData("&", "”")]
    [InlineData("…", "”")]
    [InlineData("—", "”")]
    [InlineData("…", "’")]
    public void FlowVisitorKeepsClosingPunctuationWithPreviousWordAcrossRuns(
        string punctuation, string closingMark)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var body = new Body();
            for (var rightIndent = 0; rightIndent <= 600; rightIndent += 10)
                body.Append(new Paragraph(
                    new ParagraphProperties(
                        new SpacingBetweenLines { After = "0" },
                        new Indentation { Right = rightIndent.ToString() }),
                    new Run(new Text("A short sequence xxxxxxxxx authors")),
                    new Run(new Text(punctuation)),
                    new Run(new RunProperties(new Languages { Val = "en-CA" }), new Text(closingMark)),
                    new Run(new Text(" next words here.") { Space = SpaceProcessingModeValues.Preserve })));
            body.Append(new SectionProperties(
                new PageSize { Width = 5760, Height = 15840 },
                new PageMargin { Top = 720, Right = 720, Bottom = 720, Left = 720 }));
            main.Document = new Document(body);
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var lines = pdf.GetPages().SelectMany(page => page.Letters
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 0))
            .Select(group => string.Concat(group.OrderBy(letter => letter.GlyphRectangle.Left)
                .Select(letter => letter.Value)).TrimEnd())).ToArray();
        Assert.True(lines.Length > 61);
        Assert.Contains(lines, line => line.Contains("authors" + punctuation + closingMark,
            StringComparison.Ordinal));
        Assert.Contains(lines, line => line.EndsWith("authors" + punctuation + closingMark,
            StringComparison.Ordinal));
        Assert.DoesNotContain(lines, line => line.EndsWith("authors" + punctuation,
            StringComparison.Ordinal));
        Assert.DoesNotContain(lines, line => line.StartsWith(closingMark, StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("!Syntax", 1, false)]
    [InlineData("!Syntax", 2, false)]
    [InlineData("?Syntax", 2, false)]
    [InlineData(".Syntax", 2, false)]
    [InlineData(",Syntax", 2, false)]
    [InlineData(";Syntax", 2, false)]
    [InlineData(":Syntax", 2, false)]
    [InlineData("%Syntax", 2, false)]
    [InlineData(")Syntax", 2, false)]
    [InlineData("]Syntax", 2, false)]
    [InlineData("”Syntax", 2, false)]
    [InlineData("-Syntax", 2, true)]
    [InlineData("/Syntax", 2, true)]
    [InlineData("+Syntax", 2, true)]
    [InlineData("&Syntax", 2, true)]
    [InlineData("(Syntax", 2, true)]
    [InlineData("Alpha", 2, true)]
    public void FlowVisitorKeepsLeadingClosingPunctuationWithPreviousWordAcrossSpaces(
        string followingWord, int spaces, bool firstLineHasExpression)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("= LN(10)" + new string(' ', spaces) +
                    followingWord + " Error, 10") { Space = SpaceProcessingModeValues.Preserve })),
                new SectionProperties(
                    new PageSize { Width = 4000, Height = 15840 },
                    new PageMargin { Top = 720, Right = 960, Bottom = 720, Left = 1440 })));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var firstLine = pdf.GetPage(1).Letters
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 0))
            .OrderByDescending(group => group.Key)
            .First();
        var text = string.Concat(firstLine.OrderBy(letter => letter.GlyphRectangle.Left)
            .Select(letter => letter.Value));
        Assert.Equal(firstLineHasExpression, text.Contains("LN(10)", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void FlowVisitorAllowsBreakAfterEmDashWithinRun(int spaces)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("= LN(10)" + new string(' ', spaces) +
                    "—Syntax Error, 10") { Space = SpaceProcessingModeValues.Preserve })),
                new SectionProperties(
                    new PageSize { Width = 4000, Height = 15840 },
                    new PageMargin { Top = 720, Right = 960, Bottom = 720, Left = 1440 })));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var lines = pdf.GetPage(1).Letters
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 0))
            .OrderByDescending(group => group.Key)
            .Select(group => string.Concat(group.OrderBy(letter => letter.GlyphRectangle.Left)
                .Select(letter => letter.Value)))
            .ToArray();
        Assert.Equal(2, lines.Length);
        Assert.Contains("LN(10)", lines[0]);
        Assert.StartsWith("Syntax", lines[1]);
    }

    [Theory]
    [InlineData(".", "”")]
    [InlineData("!", "”")]
    [InlineData("?", ")")]
    [InlineData(".", "]")]
    [InlineData(";", "”")]
    public void FlowVisitorKeepsDifferentlyStyledClosingMarkWithPreviousWord(
        string punctuation, string closingMark)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var body = new Body();
            for (var rightIndent = 0; rightIndent <= 1600; rightIndent += 20)
                body.Append(new Paragraph(
                    new ParagraphProperties(new SpacingBetweenLines { After = "0" },
                        new Indentation { Right = rightIndent.ToString() }),
                    new Run(new Text("A short sequence xxxxxxxxx authors" + punctuation)),
                    new Run(new RunProperties(new Color { Val = "FF0000" }), new Text(closingMark)),
                    new Run(new Text(" next words here.") { Space = SpaceProcessingModeValues.Preserve })));
            body.Append(new SectionProperties(new PageSize { Width = 5760, Height = 15840 },
                new PageMargin { Top = 720, Right = 720, Bottom = 720, Left = 720 }));
            main.Document = new Document(body);
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var lines = pdf.GetPages().SelectMany(page => page.Letters
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 0))
            .Select(group => string.Concat(group.OrderBy(letter => letter.GlyphRectangle.Left)
                .Select(letter => letter.Value)).TrimEnd())).ToArray();
        Assert.Contains(lines, line => line.EndsWith("authors" + punctuation + closingMark,
            StringComparison.Ordinal));
        Assert.DoesNotContain(lines, line => line.StartsWith(closingMark, StringComparison.Ordinal));
        Assert.DoesNotContain(lines, line => line.EndsWith("authors" + punctuation,
            StringComparison.Ordinal));
        Assert.Equal(81, pdf.GetPages().SelectMany(page => page.Letters)
            .Count(letter => letter.Value == closingMark &&
                letter.Color.ToRGBValues() is var color &&
                color.r > 0.9 && color.g < 0.1 && color.b < 0.1));
        Assert.All(pdf.GetPages().SelectMany(page => page.GetWords())
            .Where(word => word.Text == "next"), word =>
        {
            var color = word.Letters[0].Color.ToRGBValues();
            Assert.True(color.r < 0.1 && color.g < 0.1 && color.b < 0.1);
        });
    }

    [Fact]
    public void FlowVisitorDoesNotRecolorAnIdenticalUnstyledClosingMark()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("First authors.")),
                    new Run(new RunProperties(new Color { Val = "FF0000" }), new Text("”"))),
                new Paragraph(new Run(new Text("Second authors.”"))),
                new SectionProperties(new PageSize { Width = 5760, Height = 15840 },
                    new PageMargin { Top = 720, Right = 720, Bottom = 720, Left = 720 })));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var marks = pdf.GetPages().SelectMany(page => page.Letters)
            .Where(letter => letter.Value == "”").ToArray();
        Assert.Equal(2, marks.Length);
        Assert.True(marks[0].Color.ToRGBValues().r > 0.9);
        Assert.True(marks[1].Color.ToRGBValues().r < 0.1);
    }

    [Theory]
    [InlineData("“")]
    [InlineData("‘")]
    [InlineData("(")]
    [InlineData("[")]
    [InlineData("{")]
    public void FlowVisitorKeepsOpeningPunctuationWithNextWordAcrossCompatibleRuns(string openingMark)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var body = new Body();
            for (var rightIndent = 0; rightIndent <= 600; rightIndent += 10)
                body.Append(new Paragraph(
                    new ParagraphProperties(
                        new SpacingBetweenLines { After = "0" },
                        new Indentation { Right = rightIndent.ToString() }),
                    new Run(new Text("A short sequence xxxxxxxxx ") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new Text(openingMark)),
                    new Run(new Text("authors")),
                    new Run(new Text(" next words here.") { Space = SpaceProcessingModeValues.Preserve })));
            body.Append(new SectionProperties(
                new PageSize { Width = 5760, Height = 15840 },
                new PageMargin { Top = 720, Right = 720, Bottom = 720, Left = 720 }));
            main.Document = new Document(body);
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var lines = pdf.GetPages().SelectMany(page => page.Letters
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 0))
            .Select(group => string.Concat(group.OrderBy(letter => letter.GlyphRectangle.Left)
                .Select(letter => letter.Value)).TrimEnd())).ToArray();
        Assert.True(lines.Length > 61);
        Assert.Contains(lines, line => line.Contains(openingMark + "authors", StringComparison.Ordinal));
        Assert.DoesNotContain(lines, line => line.EndsWith(openingMark, StringComparison.Ordinal));
    }

    [Fact]
    public void FlowVisitorBlendsLegacyRunShadingPatterns()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var fixture = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc");

        var solid = DxpExport.ExportToBytes(Path.Combine(fixture, "WordLegacyShadedRun.docx"),
            new DxpFlowPdfVisitor());
        using var solidPdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(solid));
        var solidFills = solidPdf.GetPage(1).Paths.Where(path => path.IsFilled && path.FillColor != null)
            .Select(path => path.FillColor!.ToRGBValues()).ToArray();
        Assert.Contains(solidFills, color => color.r > 0.99 && color.g < 0.01 && color.b < 0.01);
        var leftEdges = solidPdf.GetPage(1).Paths.Where(path => path.IsFilled)
            .Select(path => path.GetBoundingRectangle()?.Left ?? double.NaN).ToArray();
        Assert.InRange(leftEdges.Min(), 71, 73);

        var percentages = DxpExport.ExportToBytes(Path.Combine(fixture,
            "WordLegacyShadingPercentages.docx"), new DxpFlowPdfVisitor());
        using var percentagePdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(percentages));
        var fills = percentagePdf.GetPage(1).Paths.Where(path => path.IsFilled && path.FillColor != null)
            .Select(path => path.FillColor!.ToRGBValues()).ToArray();
        foreach (var green in new[] { 191 / 255.0, 127 / 255.0, 64 / 255.0 })
            Assert.Contains(fills, color => color.r > 0.99 && Math.Abs(color.g - green) < 0.001 &&
                color.b < 0.01);
        var bodyShades = percentagePdf.GetPage(1).Paths.Where(path => path.IsFilled &&
            path.FillColor?.ToRGBValues() is { } color && color.r > 0.99 && color.b < 0.01 &&
            path.GetBoundingRectangle() is { } bounds && bounds.Left is > 71 and < 118 &&
            percentagePdf.GetPage(1).Height - bounds.Top is > 71 and < 74)
            .OrderBy(path => path.GetBoundingRectangle()!.Value.Left).ToArray();
        Assert.Equal(2, bodyShades.Length);
        Assert.InRange(bodyShades[1].GetBoundingRectangle()!.Value.Left -
            bodyShades[0].GetBoundingRectangle()!.Value.Right, 2, 4);
    }

    [Fact]
    public void FlowVisitorAppliesInheritedCharacterSpacingWithoutBorders()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordThreeLevelStyleOverrides.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var body = pdf.GetPage(1).Letters.Where(letter => letter.PointSize is > 9 and < 11 &&
            letter.StartBaseLine.X is >= 89 and < 140).ToArray();
        var derived = body.Take(7).ToArray();
        Assert.Equal("Derived", string.Concat(derived.Select(letter => letter.Value)));
        Assert.InRange(derived[^1].GlyphRectangle.Right, 129, 132);
        var secondLineDerived = pdf.GetPage(1).GetWords().Single(word => word.Text == "derived");
        Assert.InRange(secondLineDerived.BoundingBox.Left, 147.5, 149.5);
        var bodyWord = pdf.GetPage(1).GetWords().Single(word => word.Text == "body");
        Assert.InRange(bodyWord.BoundingBox.Left, 135.5, 137.5);
    }

    [Fact]
    public void MigraDocRendersFootnoteObjectIntoAnotherPdfPage()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var noteDocument = new MigraDoc.DocumentObjectModel.Document();
        var note = noteDocument.AddSection().AddParagraph("Footnote body appears here");
        note.Format.Font.Name = "Arial";
        note.Format.Font.Size = MigraDoc.DocumentObjectModel.Unit.FromPoint(8);
        var renderer = new MigraDoc.Rendering.PdfDocumentRenderer { Document = noteDocument };
        renderer.RenderDocument();
        using var output = new PdfSharp.Pdf.PdfDocument();
        var page = output.AddPage();
        using (var graphics = PdfSharp.Drawing.XGraphics.FromPdfPage(page))
            renderer.DocumentRenderer.RenderObject(graphics,
                new PdfSharp.Drawing.XUnitPt(72), new PdfSharp.Drawing.XUnitPt(700),
                new PdfSharp.Drawing.XUnitPt(468), note);
        using var rendered = new MemoryStream();
        output.Save(rendered, false);
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(rendered.ToArray()));
        Assert.Contains("Footnotebodyappearshere", parsed.GetPage(1).Text);
        Assert.Contains(parsed.GetPage(1).Letters, letter => letter.Value == "F" &&
            parsed.GetPage(1).Height - letter.GlyphRectangle.Top is > 690 and < 725);
    }

    [Fact]
    public void MigraDocUsesSameSpaceAfterStyledAndPlainUrl()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var document = new MigraDoc.DocumentObjectModel.Document();
        var section = document.AddSection();
        section.PageSetup.PageWidth = MigraDoc.DocumentObjectModel.Unit.FromPoint(612);
        foreach (var linked in new[] { false, true })
        {
            var paragraph = section.AddParagraph();
            paragraph.Format.Font.Name = "Arial";
            paragraph.Format.Font.Size = MigraDoc.DocumentObjectModel.Unit.FromPoint(9);
            if (linked)
            {
                var url = paragraph.AddFormattedText(" https://www.revisor.mn.gov/laws/2023/0/Session+Law/Chapter/57/");
                url.Color = MigraDoc.DocumentObjectModel.Colors.Blue;
                url.Underline = MigraDoc.DocumentObjectModel.Underline.Single;
                url.Size = MigraDoc.DocumentObjectModel.Unit.FromPoint(9);
                ((MigraDoc.DocumentObjectModel.Text)url.Elements[url.Elements.Count - 1]).Content += ";";
                paragraph.AddFormattedText(" Sidley,").Size = MigraDoc.DocumentObjectModel.Unit.FromPoint(9);
                paragraph.AddLineBreak();
                paragraph.AddText("Next line");
            }
            else
                paragraph.AddText(" https://www.revisor.mn.gov/laws/2023/0/Session+Law/Chapter/57/; Sidley,");
        }
        var control = section.AddParagraph("Hello; Sidley,");
        control.Format.Font.Name = "Arial";
        control.Format.Font.Size = MigraDoc.DocumentObjectModel.Unit.FromPoint(9);

        var renderer = new MigraDoc.Rendering.PdfDocumentRenderer { Document = document };
        renderer.RenderDocument();
        using var output = new MemoryStream();
        renderer.PdfDocument.Save(output, false);
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(output.ToArray()));
        var punctuation = parsed.GetPage(1).Letters.Where(letter => letter.Value == ";")
            .OrderByDescending(letter => letter.GlyphRectangle.Bottom).ToArray();
        var initials = parsed.GetPage(1).GetWords().Where(word => word.Text == "Sidley,")
            .OrderByDescending(word => word.BoundingBox.Bottom).ToArray();

        Assert.Equal(3, punctuation.Length);
        Assert.Equal(3, initials.Length);
        var plainGap = initials[0].BoundingBox.Left - punctuation[0].GlyphRectangle.Right;
        var linkedGap = initials[1].BoundingBox.Left - punctuation[1].GlyphRectangle.Right;
        var controlGap = initials[2].BoundingBox.Left - punctuation[2].GlyphRectangle.Right;
        Assert.InRange(Math.Abs(plainGap - linkedGap), 0, 0.1);
        using var rawDocument = new PdfSharp.Pdf.PdfDocument();
        var rawPage = rawDocument.AddPage();
        using (var graphics = PdfSharp.Drawing.XGraphics.FromPdfPage(rawPage))
            graphics.DrawString("Hello; Sidley,", new PdfSharp.Drawing.XFont("Arial", 9),
                PdfSharp.Drawing.XBrushes.Black, new PdfSharp.Drawing.XPoint(72, 72));
        using var rawStream = new MemoryStream();
        rawDocument.Save(rawStream, false);
        using var rawPdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(rawStream.ToArray()));
        var rawSemicolon = rawPdf.GetPage(1).Letters.Single(letter => letter.Value == ";");
        var rawSidley = rawPdf.GetPage(1).GetWords().Single(word => word.Text == "Sidley,");
        var rawGap = rawSidley.BoundingBox.Left - rawSemicolon.GlyphRectangle.Right;
        Assert.True(Math.Abs(controlGap - rawGap) < 0.1,
            $"MigraDoc gap {controlGap}; XGraphics gap {rawGap}");
        using var modified = PdfSharp.Pdf.IO.PdfReader.Open(new MemoryStream(output.ToArray()),
            PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify);
        typeof(DxpFlowPdfVisitor).GetMethod("ApplyFontKerning",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
            .Invoke(new DxpFlowPdfVisitor(), [modified]);
        using var adjustedStream = new MemoryStream();
        modified.Save(adjustedStream, false);
        using var adjustedPdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(adjustedStream.ToArray()));
        var adjustedPunctuation = adjustedPdf.GetPage(1).Letters.Where(letter => letter.Value == ";")
            .OrderByDescending(letter => letter.GlyphRectangle.Bottom).ToArray();
        var adjustedInitials = adjustedPdf.GetPage(1).GetWords().Where(word => word.Text == "Sidley,")
            .OrderByDescending(word => word.BoundingBox.Bottom).ToArray();
        for (var index = 0; index < 2; index++)
        {
            var punctuationShift = adjustedPunctuation[index].StartBaseLine.X -
                punctuation[index].StartBaseLine.X;
            var followingShift = adjustedInitials[index].BoundingBox.Left - initials[index].BoundingBox.Left;
            Assert.InRange(-punctuationShift, 0.5, 2);
            Assert.InRange(Math.Abs(punctuationShift - followingShift), 0, 0.05);
        }
        Assert.InRange(Math.Abs(adjustedInitials[2].BoundingBox.Left - initials[2].BoundingBox.Left), 0, 0.05);
        var nextBefore = parsed.GetPage(1).GetWords().Single(word => word.Text == "Next");
        var nextAfter = adjustedPdf.GetPage(1).GetWords().Single(word => word.Text == "Next");
        Assert.InRange(Math.Abs(nextBefore.BoundingBox.Left - nextAfter.BoundingBox.Left), 0, 0.05);
    }

    [Fact]
    public void FlowKerningKeepsTabAlignedTextFixed()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var document = new MigraDoc.DocumentObjectModel.Document();
        var section = document.AddSection();
        section.PageSetup.PageWidth = MigraDoc.DocumentObjectModel.Unit.FromPoint(612);
        var paragraph = section.AddParagraph();
        paragraph.Format.Font.Name = "Arial";
        paragraph.Format.Font.Size = MigraDoc.DocumentObjectModel.Unit.FromPoint(9);
        paragraph.Format.AddTabStop(MigraDoc.DocumentObjectModel.Unit.FromPoint(400));
        paragraph.AddText("https://www.revisor.mn.gov/laws/2023/0/Session+Law/Chapter/57/;");
        paragraph.AddTab();
        paragraph.AddText("Tabbed");

        var renderer = new MigraDoc.Rendering.PdfDocumentRenderer { Document = document };
        renderer.RenderDocument();
        using var initialStream = new MemoryStream();
        renderer.PdfDocument.Save(initialStream, false);
        using var initialPdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(initialStream.ToArray()));
        using var modified = PdfSharp.Pdf.IO.PdfReader.Open(new MemoryStream(initialStream.ToArray()),
            PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify);
        typeof(DxpFlowPdfVisitor).GetMethod("ApplyFontKerning",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
            .Invoke(new DxpFlowPdfVisitor(), [modified]);
        using var adjustedStream = new MemoryStream();
        modified.Save(adjustedStream, false);
        using var adjustedPdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(adjustedStream.ToArray()));

        var initialPage = initialPdf.GetPage(1);
        var adjustedPage = adjustedPdf.GetPage(1);
        var initialSemicolon = initialPage.Letters.Single(letter => letter.Value == ";");
        var adjustedSemicolon = adjustedPage.Letters.Single(letter => letter.Value == ";");
        var initialTab = initialPage.GetWords().Single(word => word.Text == "Tabbed");
        var adjustedTab = adjustedPage.GetWords().Single(word => word.Text == "Tabbed");
        Assert.InRange(initialSemicolon.StartBaseLine.X - adjustedSemicolon.StartBaseLine.X, 0.5, 2);
        Assert.InRange(Math.Abs(initialTab.BoundingBox.Left - adjustedTab.BoundingBox.Left), 0, 0.05);
    }

    [Fact]
    public void FlowKerningPositionsAccentedLatinWord()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var document = new MigraDoc.DocumentObjectModel.Document();
        var paragraph = document.AddSection().AddParagraph("Trépanier next");
        paragraph.Format.Font.Name = "Arial";
        paragraph.Format.Font.Size = MigraDoc.DocumentObjectModel.Unit.FromPoint(12);
        var renderer = new MigraDoc.Rendering.PdfDocumentRenderer { Document = document };
        renderer.RenderDocument();
        using var original = new MemoryStream();
        renderer.PdfDocument.Save(original, false);
        using var before = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(original.ToArray()));
        var first = before.GetPage(1).GetWords().Single(word => word.Text == "Trépanier");
        var next = before.GetPage(1).GetWords().Single(word => word.Text == "next");
        using var modified = PdfSharp.Pdf.IO.PdfReader.Open(new MemoryStream(original.ToArray()),
            PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify);
        typeof(DxpFlowPdfVisitor).GetMethod("ApplyFontKerning",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
            .Invoke(new DxpFlowPdfVisitor(), [modified]);
        using var adjusted = new MemoryStream();
        modified.Save(adjusted, false);
        using var after = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(adjusted.ToArray()));
        var adjustedFirst = after.GetPage(1).GetWords().Single(word => word.Text == "Trépanier");
        var adjustedNext = after.GetPage(1).GetWords().Single(word => word.Text == "next");
        Assert.InRange(Math.Abs(adjustedFirst.BoundingBox.Left - first.BoundingBox.Left), 0, 0.05);
        Assert.InRange(next.BoundingBox.Left - adjustedNext.BoundingBox.Left, 0.3, 1);
    }

    [Fact]
    public void FlowKerningProcessesCompressedContentStream()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var document = new MigraDoc.DocumentObjectModel.Document();
        var paragraph = document.AddSection().AddParagraph("Trépanier next");
        paragraph.Format.Font.Name = "Arial";
        paragraph.Format.Font.Size = MigraDoc.DocumentObjectModel.Unit.FromPoint(12);
        var renderer = new MigraDoc.Rendering.PdfDocumentRenderer { Document = document };
        renderer.RenderDocument();
        var content = Assert.IsType<PdfSharp.Pdf.Advanced.PdfContent>(
            Assert.IsType<PdfSharp.Pdf.Advanced.PdfReference>(
                renderer.PdfDocument.Pages[0].Contents.Elements[0]).Value);
        content.Compressed = true;
        Assert.True(content.Elements.ContainsKey("/Filter"));

        typeof(DxpFlowPdfVisitor).GetMethod("ApplyFontKerning",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
            .Invoke(new DxpFlowPdfVisitor(), [renderer.PdfDocument]);

        Assert.False(content.Elements.ContainsKey("/Filter"));
        Assert.Contains("] TJ", System.Text.Encoding.Latin1.GetString(content.Stream!.Value));
        using var stream = new MemoryStream();
        renderer.PdfDocument.Save(stream, false);
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(stream.ToArray()));
        Assert.Contains(parsed.GetPage(1).GetWords(), word => word.Text == "Trépanier");
    }

    [Fact]
    public void FlowCachedFieldsCanRefreshOnlyDateAndTime()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var dateSource = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordDateTimeFields.docx");
        var cachedOptions = new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Cache };
        var refreshOptions = new DxpExportOptions {
            FieldEvalMode = DxpFieldEvalExportMode.Cache,
            EvaluateDateTimeFieldsInCacheMode = true
        };

        string Extract(string source, DxpExportOptions options)
        {
            var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(), options: options);
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
            return string.Join(" ", pdf.GetPages().SelectMany(page => page.GetWords())
                .Select(word => word.Text));
        }

        var before = DateTime.Now.ToString("yyyy-MM-dd");
        var cached = Extract(dateSource, cachedOptions);
        var refreshed = Extract(dateSource, refreshOptions);
        var after = DateTime.Now.ToString("yyyy-MM-dd");
        Assert.Contains("2026-10-01", cached, StringComparison.Ordinal);
        Assert.True(refreshed.Contains(before, StringComparison.Ordinal) ||
            refreshed.Contains(after, StringComparison.Ordinal));
        Assert.Contains("Page 1 of 2", refreshed, StringComparison.Ordinal);

        var tocSource = Path.Combine(root, "samples", "TOC.docx");
        Assert.Equal(Extract(tocSource, cachedOptions), Extract(tocSource, refreshOptions));
    }

    [Fact]
    public void FlowVisitorBuildsSingleSpaceAfterFootnoteUrl()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var visitor = new DxpFlowPdfVisitor();
        var bytes = DxpExport.ExportToBytes(Path.Combine(root, "samples", "Right to Repair.docx"), visitor);
        var field = typeof(DxpFlowPdfVisitor).GetField("_footnoteParagraphs",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!;
        var entries = (System.Collections.IEnumerable)field.GetValue(visitor)!;
        var paragraphs = entries.Cast<object>()
            .Select(entry => (MigraDoc.DocumentObjectModel.Paragraph)entry.GetType()
                .GetField("Item2")!.GetValue(entry)!).ToArray();
        var note = paragraphs.Single(paragraph => FlattenText(paragraph).Contains("Sidley"));
        var source = FlattenText(note);

        Assert.Contains("/Chapter/57/; Sidley", source);
        Assert.Equal(MigraDoc.DocumentObjectModel.ParagraphAlignment.Left, note.Format.Alignment);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPages().Single(candidate => candidate.GetWords().Any(word => word.Text == "Sidley,"));
        var sidley = page.GetWords().Single(word => word.Text == "Sidley,");
        var preceding = page.Letters.Where(letter => letter.Value == ";" &&
                Math.Abs(letter.StartBaseLine.Y - sidley.Letters[0].StartBaseLine.Y) < 1 &&
                letter.GlyphRectangle.Right < sidley.BoundingBox.Left)
            .OrderByDescending(letter => letter.GlyphRectangle.Right).First();
        var gap = sidley.BoundingBox.Left - preceding.GlyphRectangle.Right;
        var isolated = new MigraDoc.DocumentObjectModel.Document();
        isolated.AddSection().Elements.Add((MigraDoc.DocumentObjectModel.DocumentObject)note.Clone());
        var renderer = new MigraDoc.Rendering.PdfDocumentRenderer { Document = isolated };
        renderer.RenderDocument();
        using var isolatedStream = new MemoryStream();
        renderer.PdfDocument.Save(isolatedStream, false);
        using var isolatedPdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(isolatedStream.ToArray()));
        var isolatedPage = isolatedPdf.GetPage(1);
        var isolatedSidley = isolatedPage.GetWords().Single(word => word.Text == "Sidley,");
        var isolatedPreceding = isolatedPage.Letters.Where(letter => letter.Value == ";" &&
                Math.Abs(letter.StartBaseLine.Y - isolatedSidley.Letters[0].StartBaseLine.Y) < 1 &&
                letter.GlyphRectangle.Right < isolatedSidley.BoundingBox.Left)
            .OrderByDescending(letter => letter.GlyphRectangle.Right).First();
        var isolatedGap = isolatedSidley.BoundingBox.Left - isolatedPreceding.GlyphRectangle.Right;
        Assert.InRange(Math.Abs(gap - isolatedGap), 0, 0.1);

        static string FlattenText(MigraDoc.DocumentObjectModel.Paragraph paragraph)
        {
            var result = new System.Text.StringBuilder();
            Append(paragraph.Elements);
            return result.ToString();

            void Append(MigraDoc.DocumentObjectModel.DocumentObjectCollection elements)
            {
                foreach (var element in elements)
                {
                    if (element is MigraDoc.DocumentObjectModel.Text text)
                        result.Append(text.Content);
                    else if (element is MigraDoc.DocumentObjectModel.FormattedText span)
                        Append(span.Elements);
                    else if (element is MigraDoc.DocumentObjectModel.Hyperlink hyperlink)
                        Append(hyperlink.Elements);
                }
            }
        }
    }


    [Fact]
    public void FlowVisitorPlacesFittingFootnoteWithoutAppendixPage()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("Body with a note")),
                new Run(new FootnoteReference { Id = 1 }))));
            main.Document.Save();
            var notes = main.AddNewPart<FootnotesPart>();
            notes.Footnotes = new Footnotes(new Footnote(new Paragraph(
                new Run(new FootnoteReferenceMark(), new Text("Unique note body")))) { Id = 1 });
            notes.Footnotes.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Single(pdf.GetPages());
        Assert.Contains("Uniquenotebody", pdf.GetPage(1).Text);
        Assert.Equal(1, System.Text.RegularExpressions.Regex.Matches(pdf.GetPage(1).Text,
            "Uniquenotebody").Count);
        Assert.Contains(pdf.GetPage(1).Letters, letter => letter.Value == "U" &&
            pdf.GetPage(1).Height - letter.GlyphRectangle.Top > 680);
        using var pdfDocument = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.True(pdfDocument.Internals.Catalog.Names.NameTree.ContainsName("dxp-note-1", true));
        Assert.Contains(pdfDocument.Pages[0].Contents.Elements, element =>
            (element as PdfSharp.Pdf.Advanced.PdfReference)?.Value is PdfSharp.Pdf.PdfDictionary content &&
            content.Stream != null &&
            System.Text.Encoding.Latin1.GetString(content.Stream.UnfilteredValue).Contains(".6 w"));
    }

    [Fact]
    public void FlowVisitorSizesImplicitTablesInBodyAndEndnote()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Body with an endnote")),
                    new Run(new EndnoteReference { Id = 1 })),
                CreateImplicitTable("Body")));
            main.Document.Save();
            var notes = main.AddNewPart<EndnotesPart>();
            notes.Endnotes = new Endnotes(new Endnote(
                new Paragraph(new Run(new EndnoteReferenceMark(), new Text("Unique endnote body"))),
                CreateImplicitTable("Note")) { Id = 1 });
            notes.Endnotes.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var bodyMarker = page.Letters.Where(letter => letter.Value == "i" &&
                page.Height - letter.GlyphRectangle.Top < 90)
            .OrderByDescending(letter => letter.GlyphRectangle.Left).First();
        var bodyText = page.Letters.First(letter => letter.Value == "B" &&
            page.Height - letter.GlyphRectangle.Top < 90);
        Assert.Equal(bodyText.FontSize, bodyMarker.FontSize);
        foreach (var (initial, firstRow, secondColumn) in new[]
                 {
                     ("B", 97.4, 139.34),
                     ("N", 205.9, 138.50)
                 })
        {
            var matching = page.Letters.Where(letter => letter.Value == initial &&
                page.Height - letter.GlyphRectangle.Top >= firstRow - 1)
                .OrderBy(letter => page.Height - letter.GlyphRectangle.Top)
                .ThenBy(letter => letter.GlyphRectangle.Left).ToArray();
            Assert.Equal(4, matching.Length);
            var firstTop = page.Height - matching[0].GlyphRectangle.Top;
            var secondTop = page.Height - matching[2].GlyphRectangle.Top;
            Assert.InRange(firstTop, firstRow - (initial == "B" ? 1 : 3),
                firstRow + (initial == "B" ? 1 : 3));
            Assert.InRange(secondTop - firstTop, 22, 26);
            Assert.InRange(matching[0].GlyphRectangle.Left, 71.5, 73.5);
            Assert.InRange(matching[1].GlyphRectangle.Left, secondColumn - 5, secondColumn + 5);
        }

        static Table CreateImplicitTable(string label) => new(
            new TableRow(new TableCell(new Paragraph(new Run(new Text($"{label} cell 1-1")))),
                new TableCell(new Paragraph(new Run(new Text($"{label} cell 1-2"))))),
            new TableRow(new TableCell(new Paragraph(new Run(new Text($"{label} cell 2-1")))),
                new TableCell(new Paragraph(new Run(new Text($"{label} cell 2-2"))))));
    }

    [Fact]
    public void FlowVisitorKeepsStylelessEndnoteTableAtWordSeparatorSpacing()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfFeatures",
            "EndnoteTable.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var body = page.Letters.First(letter => letter.Value == "B");
        var note = page.Letters.First(letter => letter.Value == "U");
        var noteOffset = body.GlyphRectangle.Top - note.GlyphRectangle.Top;
        Assert.InRange(noteOffset, 36.5, 38.5);
    }

    [Fact]
    public void FlowVisitorRendersInlineOmmlRunBetweenTextRuns()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("Before ") { Space = SpaceProcessingModeValues.Preserve }),
                new MathXml.OfficeMath(new MathXml.Run(new MathXml.Text("x"))),
                new Run(new Text(" after") { Space = SpaceProcessingModeValues.Preserve }))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains(page.GetWords(), word => word.Text.Contains("𝑥", StringComparison.Ordinal) ||
            word.Text.Contains("x", StringComparison.Ordinal));
        Assert.Contains("after", page.Text);
    }

    [Fact]
    public void FlowVisitorRendersMultipleInlineOmmlRuns()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("Before ") { Space = SpaceProcessingModeValues.Preserve }),
                new MathXml.OfficeMath(
                    new MathXml.Run(new MathXml.Text("x")),
                    new MathXml.Run(new MathXml.Text("+")),
                    new MathXml.Run(new MathXml.Text("y"))),
                new Run(new Text(" after") { Space = SpaceProcessingModeValues.Preserve }))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains("x+y", string.Concat(page.Text.Where(character => !char.IsWhiteSpace(character))));
        Assert.Contains("after", page.Text);
        var plus = Assert.Single(page.Letters.Where(letter => letter.Value == "+"));
        var preceding = page.Letters.Last(letter => letter.StartBaseLine.X < plus.StartBaseLine.X &&
            (letter.Value == "𝑥" || letter.Value == "x"));
        var following = page.Letters.First(letter => letter.StartBaseLine.X > plus.StartBaseLine.X &&
            (letter.Value == "𝑦" || letter.Value == "y"));
        Assert.InRange(plus.StartBaseLine.X - preceding.StartBaseLine.X, 8, 11);
        Assert.InRange(following.StartBaseLine.X - plus.StartBaseLine.X, 8, 9.6);
    }

    [Fact]
    public void FlowVisitorHonorsRunsOnlyOmmlPlainStyle()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new MathXml.OfficeMath(
                    new MathXml.Run(new MathXml.RunProperties(
                        new MathXml.Style { Val = MathXml.StyleValues.Plain }),
                        new MathXml.Text("F")),
                    new MathXml.Run(new MathXml.Text("x"))))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var letters = page.Letters.Where(letter => letter.Value is "F" or "x").ToArray();
        Assert.Equal(2, letters.Length);
        Assert.False(letters[0].Font.Name.Contains("Italic", StringComparison.OrdinalIgnoreCase));
        Assert.True(letters[1].Font.Name.Contains("Italic", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void FlowVisitorHonorsRunsOnlyOmmlFontSize()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new MathXml.OfficeMath(new MathXml.Run(
                    new RunProperties(new FontSize { Val = "40" }),
                    new MathXml.Text("x"))))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var letter = Assert.Single(page.Letters.Where(candidate => candidate.Value == "x"));
        Assert.InRange(letter.PointSize, 19.5, 20.5);
    }

    [Fact]
    public void FlowVisitorRendersRunsOnlyDisplayOmml()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Before equation"))),
                new Paragraph(new MathXml.Paragraph(new MathXml.OfficeMath(
                    new MathXml.Run(new MathXml.Text("x")),
                    new MathXml.Run(new MathXml.Text("+")),
                    new MathXml.Run(new MathXml.Text("y"))))),
                new Paragraph(new Run(new Text("After equation")))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains("x+y", page.Text);
        Assert.Contains("Afterequation", page.Text);
        var equation = Assert.Single(page.GetWords().Where(word => word.Text == "x+y"));
        Assert.InRange((equation.BoundingBox.Left + equation.BoundingBox.Right) / 2,
            page.Width / 2 - 5, page.Width / 2 + 5);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(40)]
    [InlineData(-40)]
    public void FlowVisitorRendersStackedOmmlFraction(int explicitHalfPoints)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Before equation"))),
                new Paragraph(new MathXml.Paragraph(new MathXml.OfficeMath(
                    new MathXml.Fraction(
                        new MathXml.Numerator(CreateMathRun("1")),
                        new MathXml.Denominator(CreateMathRun("2")))))),
                new Paragraph(new Run(new Text("After equation")))));
            if (explicitHalfPoints < 0)
                main.Document.Descendants<Paragraph>().ElementAt(1).PrependChild(
                    new ParagraphProperties(new ParagraphMarkRunProperties(
                        new FontSize { Val = (-explicitHalfPoints).ToString() })));
            if (explicitHalfPoints > 0)
                Assert.Equal(2, main.Document.Descendants<FontSize>().Count());
            main.Document.Save();
        }

        if (explicitHalfPoints > 0)
        {
            using var reopened = WordprocessingDocument.Open(new MemoryStream(source.ToArray()), false);
            Assert.Equal(2, reopened.MainDocumentPart!.Document.Descendants<FontSize>().Count());
            Assert.Equal("40", reopened.MainDocumentPart.Document.Descendants<FontSize>().First().Val?.Value);
        }
        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var numerator = Assert.Single(page.Letters.Where(letter => letter.Value == "1"));
        var denominator = Assert.Single(page.Letters.Where(letter => letter.Value == "2"));
        Assert.InRange(numerator.GlyphRectangle.Left, explicitHalfPoints <= 0 ? 301 : 299,
            explicitHalfPoints <= 0 ? 305 : 303);
        Assert.InRange(denominator.GlyphRectangle.Left, explicitHalfPoints <= 0 ? 301 : 299,
            explicitHalfPoints <= 0 ? 305 : 303);
        Assert.InRange(numerator.GlyphRectangle.Bottom - denominator.GlyphRectangle.Top, 0, 15);
        Assert.InRange(numerator.PointSize, explicitHalfPoints <= 0 ? 11.5 : 19.5,
            explicitHalfPoints <= 0 ? 12.5 : 20.5);
        Assert.Contains("Afterequation", page.Text);

        MathXml.Run CreateMathRun(string value) => explicitHalfPoints == 0
            ? new MathXml.Run(new MathXml.Text(value))
            : new MathXml.Run(new RunProperties(new FontSize { Val = explicitHalfPoints.ToString() }),
                new MathXml.Text(value));
    }

    [Theory]
    [InlineData(0, 72.51, false)]
    [InlineData(40, 71.65, false)]
    [InlineData(0, 85.34, true)]
    public void FlowVisitorPositionsStandaloneDisplayFractionAgainstWord(int explicitHalfPoints,
        double expectedNumeratorTop, bool surrounded)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            MathXml.Run MathRun(string value) => explicitHalfPoints == 0
                ? new MathXml.Run(new MathXml.Text(value))
                : new MathXml.Run(new RunProperties(new FontSize { Val = explicitHalfPoints.ToString() }),
                    new MathXml.Text(value));
            var body = new Body();
            if (surrounded) body.Append(new Paragraph(new Run(new Text("Before equation"))));
            body.Append(new Paragraph(
                new Run(new Text("")),
                new MathXml.Paragraph(new MathXml.OfficeMath(new MathXml.Fraction(
                    new MathXml.Numerator(MathRun("1")),
                    new MathXml.Denominator(MathRun("2"))))),
                new Run(new Text(""))));
            if (surrounded) body.Append(new Paragraph(new Run(new Text("After equation"))));
            main.Document = new Document(body);
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var numerator = Assert.Single(page.Letters.Where(letter => letter.Value == "1"));
        var denominator = Assert.Single(page.Letters.Where(letter => letter.Value == "2"));
        Assert.InRange(page.Height - numerator.GlyphRectangle.Top,
            expectedNumeratorTop - 0.6, expectedNumeratorTop + 0.6);
        Assert.InRange(page.Height - denominator.GlyphRectangle.Top - expectedNumeratorTop,
            16.5, 21.5);
    }

    [Fact]
    public void FlowVisitorCentersUnequalWidthFractionTerms()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Before equation"))),
                new Paragraph(new MathXml.Paragraph(new MathXml.OfficeMath(
                    new MathXml.Fraction(
                        new MathXml.Numerator(new MathXml.Run(new MathXml.Text("12"))),
                        new MathXml.Denominator(new MathXml.Run(new MathXml.Text("345"))))))),
                new Paragraph(new Run(new Text("After equation")))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var numerator = Assert.Single(page.GetWords().Where(word => word.Text == "12"));
        var denominator = Assert.Single(page.GetWords().Where(word => word.Text == "345"));
        Assert.InRange(numerator.BoundingBox.Left, 298, 301);
        Assert.InRange(denominator.BoundingBox.Left, 294, 299);
        Assert.InRange(Math.Abs((numerator.BoundingBox.Left + numerator.BoundingBox.Right) / 2 -
            (denominator.BoundingBox.Left + denominator.BoundingBox.Right) / 2), 0, 0.5);
    }

    [Fact]
    public void FlowVisitorRendersInlineOmmlFractionBetweenTextRuns()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("Before ") { Space = SpaceProcessingModeValues.Preserve }),
                new MathXml.OfficeMath(new MathXml.Fraction(
                    new MathXml.Numerator(new MathXml.Run(new MathXml.Text("1"))),
                    new MathXml.Denominator(new MathXml.Run(new MathXml.Text("2"))))),
                new Run(new Text(" after") { Space = SpaceProcessingModeValues.Preserve }))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var numerator = Assert.Single(page.Letters.Where(letter => letter.Value == "1"));
        var denominator = Assert.Single(page.Letters.Where(letter => letter.Value == "2"));
        var following = Assert.Single(page.GetWords().Where(word => word.Text == "after"));
        Assert.InRange(numerator.GlyphRectangle.Left, 107, 110);
        Assert.InRange(denominator.GlyphRectangle.Left, 107, 110);
        Assert.InRange(following.BoundingBox.Left, 112, 116);
        Assert.InRange(numerator.StartBaseLine.Y - denominator.StartBaseLine.Y, 12.5, 13.5);
        Assert.Contains(page.Paths, path => path.IsFilled &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 107 and < 110 && bounds.Width is > 4 and < 6 &&
            page.Height - bounds.Top is > 81.5 and < 82.5);
    }

    [Fact]
    public void FlowVisitorRendersNestedInlineOmmlFractionBetweenTextRuns()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("Before ") { Space = SpaceProcessingModeValues.Preserve }),
                new MathXml.OfficeMath(new MathXml.Fraction(
                    new MathXml.Numerator(new MathXml.Fraction(
                        new MathXml.Numerator(new MathXml.Run(new MathXml.Text("1"))),
                        new MathXml.Denominator(new MathXml.Run(new MathXml.Text("3"))))),
                    new MathXml.Denominator(new MathXml.Run(new MathXml.Text("2"))))),
                new Run(new Text(" after") { Space = SpaceProcessingModeValues.Preserve }))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var numerator = Assert.Single(page.Letters.Where(letter => letter.Value == "1"));
        var middle = Assert.Single(page.Letters.Where(letter => letter.Value == "3"));
        var denominator = Assert.Single(page.Letters.Where(letter => letter.Value == "2"));
        var following = Assert.Single(page.GetWords().Where(word => word.Text == "after"));
        Assert.InRange(numerator.GlyphRectangle.Left, 108, 112);
        Assert.InRange(middle.GlyphRectangle.Left, 108, 112);
        Assert.InRange(denominator.GlyphRectangle.Left, 108, 112);
        Assert.InRange(following.BoundingBox.Left, 114, 119);
        Assert.InRange(numerator.StartBaseLine.Y - middle.StartBaseLine.Y, 8, 11);
        Assert.InRange(middle.StartBaseLine.Y - denominator.StartBaseLine.Y, 8, 11);
        Assert.InRange(numerator.PointSize, 6.9, 7.1);
        Assert.InRange(middle.PointSize, 6.9, 7.1);
        Assert.InRange(denominator.PointSize, 8.4, 8.7);
        Assert.Contains(page.Paths, path => path.IsFilled &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Width is > 3.8 and < 4.5 &&
            page.Height - bounds.Top is > 79.9 and < 80.5);
        Assert.Contains(page.Paths, path => path.IsFilled &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Width is > 7 and < 7.8 &&
            page.Height - bounds.Top is > 88.1 and < 88.7);
    }

    [Fact]
    public void FlowVisitorReservesWidthForWideDenominatorInNestedInlineFraction()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("Before ") { Space = SpaceProcessingModeValues.Preserve }),
                new MathXml.OfficeMath(new MathXml.Fraction(
                    new MathXml.Numerator(new MathXml.Fraction(
                        new MathXml.Numerator(new MathXml.Run(new MathXml.Text("1"))),
                        new MathXml.Denominator(new MathXml.Run(new MathXml.Text("3"))))),
                    new MathXml.Denominator(new MathXml.Run(new MathXml.Text("456"))))),
                new Run(new Text(" after") { Space = SpaceProcessingModeValues.Preserve }))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var numerator = Assert.Single(page.Letters.Where(letter => letter.Value == "1"));
        var middle = Assert.Single(page.Letters.Where(letter => letter.Value == "3"));
        var denominator = page.Letters.Where(letter => letter.Value is "4" or "5" or "6")
            .OrderBy(letter => letter.GlyphRectangle.Left).ToArray();
        var following = Assert.Single(page.GetWords().Where(word => word.Text == "after"));
        Assert.Equal(3, denominator.Length);
        Assert.InRange(numerator.GlyphRectangle.Left, 112, 117);
        Assert.InRange(middle.GlyphRectangle.Left, 112, 117);
        Assert.InRange(denominator[0].GlyphRectangle.Left, 107, 111);
        Assert.InRange(following.BoundingBox.Left, 127.5, 128.1);
        Assert.InRange(numerator.PointSize, 6.9, 7.1);
        Assert.All(denominator, letter => Assert.InRange(letter.PointSize, 8.4, 8.7));
        Assert.Contains(page.Paths, path => path.IsFilled &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Width is > 16.8 and < 17.5 &&
            page.Height - bounds.Top is > 88.1 and < 88.7);
    }

    [Fact]
    public void FlowVisitorRendersNestedInlineDenominatorWithoutLoweringSurroundingText()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("Before ") { Space = SpaceProcessingModeValues.Preserve }),
                new MathXml.OfficeMath(new MathXml.Fraction(
                    new MathXml.Numerator(new MathXml.Run(new MathXml.Text("1"))),
                    new MathXml.Denominator(new MathXml.Fraction(
                        new MathXml.Numerator(new MathXml.Run(new MathXml.Text("2"))),
                        new MathXml.Denominator(new MathXml.Run(new MathXml.Text("2"))))))),
                new Run(new Text(" after") { Space = SpaceProcessingModeValues.Preserve }))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var numerator = Assert.Single(page.Letters.Where(letter => letter.Value == "1"));
        var denominator = page.Letters.Where(letter => letter.Value == "2")
            .OrderByDescending(letter => letter.StartBaseLine.Y).ToArray();
        var before = Assert.Single(page.GetWords().Where(word => word.Text == "Before"));
        var after = Assert.Single(page.GetWords().Where(word => word.Text == "after"));
        Assert.Equal(2, denominator.Length);
        Assert.InRange(before.BoundingBox.Bottom - after.BoundingBox.Bottom, -0.1, 0.1);
        Assert.InRange(numerator.GlyphRectangle.Top - denominator[0].GlyphRectangle.Top, 9.5, 11);
        Assert.InRange(denominator[0].GlyphRectangle.Top - denominator[1].GlyphRectangle.Top, 8.5, 10.5);
        Assert.InRange(numerator.PointSize, 8.4, 8.7);
        Assert.All(denominator, letter => Assert.InRange(letter.PointSize, 6.9, 7.1));
        Assert.Contains(page.Paths, path => path.IsFilled &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Width is > 7 and < 7.8 &&
            page.Height - bounds.Top is > 81.9 and < 82.7);
        Assert.Contains(page.Paths, path => path.IsFilled &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Width is > 3.8 and < 4.5 &&
            page.Height - bounds.Top is > 90.1 and < 90.9);
    }

    [Theory]
    [InlineData(true, "+")]
    [InlineData(false, "+")]
    [InlineData(true, "=")]
    [InlineData(false, "=")]
    public void FlowVisitorRendersMixedOmmlRunAndFractionInSourceOrder(bool runFirst,
        string mathOperator)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var math = new MathXml.OfficeMath();
            if (runFirst)
                math.Append(new MathXml.Run(new MathXml.Text("x" + mathOperator)));
            math.Append(new MathXml.Fraction(
                new MathXml.Numerator(new MathXml.Run(new MathXml.Text("1"))),
                new MathXml.Denominator(new MathXml.Run(new MathXml.Text("2")))));
            if (!runFirst)
                math.Append(new MathXml.Run(new MathXml.Text(mathOperator + "x")));
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("Before ") { Space = SpaceProcessingModeValues.Preserve }),
                math,
                new Run(new Text(" after") { Space = SpaceProcessingModeValues.Preserve }))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var variable = Assert.Single(page.Letters.Where(letter => letter.Value is "x" or "𝑥"));
        var operation = Assert.Single(page.Letters.Where(letter => letter.Value == mathOperator));
        var numerator = Assert.Single(page.Letters.Where(letter => letter.Value == "1"));
        var denominator = Assert.Single(page.Letters.Where(letter => letter.Value == "2"));
        var after = Assert.Single(page.GetWords().Where(word => word.Text == "after"));
        if (runFirst)
        {
            Assert.True(operation.GlyphRectangle.Left > variable.GlyphRectangle.Right + 2);
            Assert.True(numerator.GlyphRectangle.Left > operation.GlyphRectangle.Right + 2);
            Assert.True(after.BoundingBox.Left > numerator.GlyphRectangle.Right);
        }
        else
        {
            Assert.True(operation.GlyphRectangle.Left > numerator.GlyphRectangle.Right + 2);
            Assert.True(variable.GlyphRectangle.Left > operation.GlyphRectangle.Right + 2);
            Assert.True(after.BoundingBox.Left > variable.GlyphRectangle.Right);
        }
        Assert.InRange(numerator.GlyphRectangle.Left - denominator.GlyphRectangle.Left, -1, 1);
    }

    [Fact]
    public void FlowVisitorRendersMixedDisplayOmmlRunAndFraction()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var math = new MathXml.OfficeMath(
                new MathXml.Run(new MathXml.Text("x+")),
                new MathXml.Fraction(
                    new MathXml.Numerator(new MathXml.Run(new MathXml.Text("1"))),
                    new MathXml.Denominator(new MathXml.Run(new MathXml.Text("2")))));
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("")), new MathXml.Paragraph(math), new Run(new Text("")))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains(page.Letters, letter => letter.Value is "x" or "𝑥");
        Assert.Contains(page.Letters, letter => letter.Value == "+");
        var numerator = Assert.Single(page.Letters.Where(letter => letter.Value == "1"));
        var denominator = Assert.Single(page.Letters.Where(letter => letter.Value == "2"));
        Assert.InRange(numerator.PointSize, 11.9, 12.1);
        Assert.InRange(denominator.PointSize, 11.9, 12.1);
        Assert.InRange(numerator.GlyphRectangle.Left - denominator.GlyphRectangle.Left, -1, 1);
    }

    [Fact]
    public void FlowVisitorRendersMixedDisplayOmmlRunAndRadical()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var math = new MathXml.OfficeMath(
                new MathXml.Run(new MathXml.Text("1")),
                new MathXml.Radical(new MathXml.Degree(),
                    new MathXml.Base(new MathXml.Run(new MathXml.Text("x")))),
                new MathXml.Run(new MathXml.Text("2")));
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("")), new MathXml.Paragraph(math), new Run(new Text("")))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = Assert.Single(pdf.GetPages()).Letters.ToArray();
        var first = Assert.Single(letters.Where(letter => letter.Value == "1"));
        var root = Assert.Single(letters.Where(letter => letter.Value == "√"));
        var last = Assert.Single(letters.Where(letter => letter.Value == "2"));
        Assert.True(first.GlyphRectangle.Left < root.GlyphRectangle.Left);
        Assert.True(root.GlyphRectangle.Left < last.GlyphRectangle.Left);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FlowVisitorRendersMixedDisplayOmmlRunAndComplexElement(bool matrix)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            OpenXmlElement element = matrix
                ? new MathXml.Matrix(new MathXml.MatrixRow(
                    new MathXml.Base(new MathXml.Run(new MathXml.Text("1"))),
                    new MathXml.Base(new MathXml.Run(new MathXml.Text("2")))))
                : new MathXml.Superscript(
                    new MathXml.Base(new MathXml.Run(new MathXml.Text("x"))),
                    new MathXml.SuperArgument(new MathXml.Run(new MathXml.Text("2"))));
            var math = new MathXml.OfficeMath(
                new MathXml.Run(new MathXml.Text("7")), element,
                new MathXml.Run(new MathXml.Text("8")));
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("")), new MathXml.Paragraph(math), new Run(new Text("")))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = Assert.Single(pdf.GetPages()).Letters.ToArray();
        var before = Assert.Single(letters.Where(letter => letter.Value == "7"));
        var after = Assert.Single(letters.Where(letter => letter.Value == "8"));
        var middle = Assert.Single(letters.Where(letter => letter.Value == "1" ||
            !matrix && letter.Value is "x" or "𝑥"));
        Assert.True(before.GlyphRectangle.Left < middle.GlyphRectangle.Left);
        Assert.True(middle.GlyphRectangle.Left < after.GlyphRectangle.Left);
    }

    [Fact]
    public void FlowVisitorCentersSurroundingTextAgainstInlineMatrixInk()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            "omml-matrix-probe__Matrix2x3.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var before = page.Letters.Single(letter => letter.Value == "B");
        var matrix = page.Letters.Single(letter => letter.Value == "a" &&
            page.Height - letter.StartBaseLine.Y < 85);
        Assert.InRange(matrix.StartBaseLine.Y - before.StartBaseLine.Y, 6.9, 7.2);
    }

    [Fact]
    public void FlowVisitorAlignsNumericMatrixColumnsAndFollowingText()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            "omml-matrix-probe__Matrix2x3Digits.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var after = page.GetWords().Single(word => word.Text == "after");
        var lastColumn = page.Letters.Single(letter => letter.Value == "3");
        Assert.InRange(after.BoundingBox.Left, 154.4, 154.9);
        Assert.InRange(lastColumn.GlyphRectangle.Left, 145.4, 145.9);
    }

    [Fact]
    public void FlowVisitorLowersOnlyTheMixedInlineRadicalRoot()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            "mixed-radical-probe__MixedRadicalDisplay.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var radical = page.Letters.Single(letter => letter.Value == "√");
        var radicand = page.Letters.Single(letter => letter.Value == "x");
        Assert.InRange(radicand.StartBaseLine.Y - radical.StartBaseLine.Y, 0.8, 0.9);
    }

    [Theory]
    [InlineData(false, 83.35)]
    [InlineData(true, 77.58)]
    public void FlowVisitorKeepsArialNestedFractionStoryBaseline(bool denominatorNested,
        double expectedTop)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var nested = new MathXml.Fraction(
                new MathXml.Numerator(new MathXml.Run(new MathXml.Text("1"))),
                new MathXml.Denominator(new MathXml.Run(new MathXml.Text("2"))));
            var fraction = denominatorNested
                ? new MathXml.Fraction(
                    new MathXml.Numerator(new MathXml.Run(new MathXml.Text("3"))),
                    new MathXml.Denominator(nested))
                : new MathXml.Fraction(
                    new MathXml.Numerator(nested),
                    new MathXml.Denominator(new MathXml.Run(new MathXml.Text("3"))));
            Run ArialRun(string value) => new(
                new RunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                    new FontSize { Val = "24" }),
                new Text(value) { Space = SpaceProcessingModeValues.Preserve });
            main.Document = new Document(new Body(new Paragraph(
                ArialRun("Before "), new MathXml.OfficeMath(fraction), ArialRun(" after"))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var first = Assert.Single(page.Letters.Where(letter => letter.Value == "B"));
        var following = Assert.Single(page.GetWords().Where(word => word.Text == "after"));
        Assert.InRange(page.Height - first.GlyphRectangle.Top, expectedTop - 0.3, expectedTop + 0.3);
        Assert.InRange(page.Height - following.BoundingBox.Top, expectedTop - 0.3, expectedTop + 0.3);
    }

    [Fact]
    public void FlowVisitorRendersMultipleInlineOmmlFractionsInOneParagraph()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("Before ") { Space = SpaceProcessingModeValues.Preserve }),
                Fraction("1", "2"),
                new Run(new Text(" between ") { Space = SpaceProcessingModeValues.Preserve }),
                Fraction("3", "4"),
                new Run(new Text(" after") { Space = SpaceProcessingModeValues.Preserve }))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var numerators = new[] { "1", "3" }.Select(value =>
            Assert.Single(page.Letters.Where(letter => letter.Value == value))).ToArray();
        var denominators = new[] { "2", "4" }.Select(value =>
            Assert.Single(page.Letters.Where(letter => letter.Value == value))).ToArray();
        Assert.True(numerators[1].GlyphRectangle.Left > numerators[0].GlyphRectangle.Left + 25);
        for (var index = 0; index < numerators.Length; index++)
        {
            Assert.InRange(Math.Abs((numerators[index].GlyphRectangle.Left +
                numerators[index].GlyphRectangle.Right) / 2 -
                (denominators[index].GlyphRectangle.Left +
                 denominators[index].GlyphRectangle.Right) / 2), 0, 0.5);
            Assert.InRange(numerators[index].StartBaseLine.Y -
                denominators[index].StartBaseLine.Y, 12.5, 13.5);
        }

        static MathXml.OfficeMath Fraction(string numerator, string denominator) => new(
            new MathXml.Fraction(
                new MathXml.Numerator(new MathXml.Run(new MathXml.Text(numerator))),
                new MathXml.Denominator(new MathXml.Run(new MathXml.Text(denominator)))));
    }

    [Theory]
    [InlineData("12", "345", 14, 16)]
    [InlineData("123", "4567", 19, 21)]
    public void FlowVisitorReservesWidthForUnequalInlineOmmlFractionTerms(
        string numeratorText, string denominatorText, double minimumRuleWidth, double maximumRuleWidth)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("Before ") { Space = SpaceProcessingModeValues.Preserve }),
                new MathXml.OfficeMath(new MathXml.Fraction(
                    new MathXml.Numerator(new MathXml.Run(new MathXml.Text(numeratorText))),
                    new MathXml.Denominator(new MathXml.Run(new MathXml.Text(denominatorText))))),
                new Run(new Text(" after") { Space = SpaceProcessingModeValues.Preserve }))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var numerator = Assert.Single(page.GetWords().Where(word => word.Text == numeratorText));
        var denominator = Assert.Single(page.GetWords().Where(word => word.Text == denominatorText));
        var following = Assert.Single(page.GetWords().Where(word => word.Text == "after"));
        Assert.InRange(following.BoundingBox.Left, 120, 140);
        Assert.InRange(Math.Abs((numerator.BoundingBox.Left + numerator.BoundingBox.Right) / 2 -
            (denominator.BoundingBox.Left + denominator.BoundingBox.Right) / 2), 0, 0.5);
        Assert.Contains(page.Paths, path => path.IsFilled &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 107 and < 110 &&
            bounds.Width > minimumRuleWidth && bounds.Width < maximumRuleWidth);
    }

    [Fact]
    public void FlowVisitorReservesInlineOmmlFractionWidthWithExplicitArialRuns()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                ArialRun("Before "),
                new MathXml.OfficeMath(new MathXml.Fraction(
                    new MathXml.Numerator(new MathXml.Run(new MathXml.Text("12"))),
                    new MathXml.Denominator(new MathXml.Run(new MathXml.Text("345"))))),
                ArialRun(" after"))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var following = Assert.Single(page.GetWords().Where(word => word.Text == "after"));
        var bar = Assert.Single(page.Paths.Where(path => path.IsFilled &&
            path.GetBoundingRectangle() is { } bounds && bounds.Width is > 14 and < 16));
        var rectangle = Assert.IsType<UglyToad.PdfPig.Core.PdfRectangle>(bar.GetBoundingRectangle());
        Assert.InRange(following.BoundingBox.Left - rectangle.Right, 2, 5);
        Assert.InRange(page.Height - rectangle.Top, 81.5, 82.5);

        static Run ArialRun(string value) => new(
            new RunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                new FontSize { Val = "24" }),
            new Text(value) { Space = SpaceProcessingModeValues.Preserve });
    }

    [Fact]
    public void FlowVisitorKeepsWrappedInlineOmmlFractionWithFollowingText()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(
                    new Run(new Text("Before words before words before words before words ")
                        { Space = SpaceProcessingModeValues.Preserve }),
                    new MathXml.OfficeMath(new MathXml.Fraction(
                        new MathXml.Numerator(new MathXml.Run(new MathXml.Text("12"))),
                        new MathXml.Denominator(new MathXml.Run(new MathXml.Text("345"))))),
                    new Run(new Text(" after words") { Space = SpaceProcessingModeValues.Preserve })),
                new SectionProperties(
                    new PageSize { Width = 5760, Height = 15840 },
                    new PageMargin { Top = 720, Right = 720, Bottom = 720, Left = 720 })));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var numerator = Assert.Single(page.GetWords().Where(word => word.Text == "12"));
        var denominator = Assert.Single(page.GetWords().Where(word => word.Text == "345"));
        var following = Assert.Single(page.GetWords().Where(word => word.Text == "after"));
        var firstLetter = Assert.Single(page.Letters.Where(letter => letter.Value == "B"));
        Assert.InRange(page.Height - firstLetter.StartBaseLine.Y, 46.0, 46.7);
        Assert.InRange(following.BoundingBox.Left - denominator.BoundingBox.Right, 0, 8);
        Assert.InRange(Math.Abs(following.BoundingBox.Top - numerator.BoundingBox.Top), 0, 12);
        Assert.InRange(numerator.BoundingBox.Left, 90, 130);
    }

    [Fact]
    public void FlowVisitorRendersDisplayOmmlRadical()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Before equation"))),
                new Paragraph(new MathXml.Paragraph(new MathXml.OfficeMath(
                    new MathXml.Radical(new MathXml.Degree(),
                        new MathXml.Base(new MathXml.Run(new MathXml.Text("x"))))))),
                new Paragraph(new Run(new Text("After equation")))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var radical = Assert.Single(page.Letters.Where(letter => letter.Value == "√"));
        var variable = Assert.Single(page.Letters.Where(letter => letter.Value == "x"));
        var following = Assert.Single(page.GetWords().Where(word => word.Text == "After"));
        Assert.InRange(radical.GlyphRectangle.Left, 299, 303);
        Assert.InRange(variable.GlyphRectangle.Left, 306, 311);
        Assert.InRange(page.Height - radical.GlyphRectangle.Top, 87, 90);
        Assert.InRange(page.Height - following.BoundingBox.Top, 109, 114);
        Assert.Contains("Afterequation", page.Text);
    }

    [Fact]
    public void FlowVisitorRendersInlineOmmlRadical()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("Before ") { Space = SpaceProcessingModeValues.Preserve }),
                new MathXml.OfficeMath(new MathXml.Radical(new MathXml.Degree(),
                    new MathXml.Base(new MathXml.Run(new MathXml.Text("x"))))),
                new Run(new Text(" after") { Space = SpaceProcessingModeValues.Preserve }))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains(page.Letters, letter => letter.Value == "√");
        Assert.Contains(page.Letters, letter => letter.Value == "𝑥" || letter.Value == "x");
        Assert.Contains(page.Paths, path => path.IsFilled &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 110 and < 120 &&
            bounds.Width is > 4 and < 12 && bounds.Height is > 0.7 and < 1);
        Assert.Contains("Before", page.Text);
        Assert.Contains("after", page.Text);
    }

    [Fact]
    public void FlowVisitorKeepsExplicitInlineRadicandSize()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new MathXml.OfficeMath(new MathXml.Radical(new MathXml.Degree(),
                    new MathXml.Base(new MathXml.Run(
                        new RunProperties(new FontSize { Val = "40" }),
                        new MathXml.Text("x"))))))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = Assert.Single(pdf.GetPages()).Letters;
        var radical = Assert.Single(letters.Where(
            letter => letter.Value == "√"));
        var radicand = Assert.Single(letters.Where(
            letter => letter.Value == "𝑥" || letter.Value == "x"));
        Assert.InRange(radicand.PointSize, 19.9, 20.1);
        if (radical.Font.Name.Contains("Cambria Math", StringComparison.OrdinalIgnoreCase))
            Assert.InRange(radical.PointSize, 11.9, 12.1);
        else Assert.InRange(radical.PointSize, 19.9, 20.1);
    }

    [Fact]
    public void FlowVisitorKeepsImplicitParagraphGapWithExplicitRunFont()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;

        double SecondBaseline(bool explicitZeroAfter)
        {
            using var source = new MemoryStream();
            using (var document = WordprocessingDocument.Create(source,
                WordprocessingDocumentType.Document, true))
            {
                var first = new Paragraph(new Run(new Text("Alpha")));
                if (explicitZeroAfter)
                    first.ParagraphProperties = new ParagraphProperties(
                        new SpacingBetweenLines { After = "0" });
                var second = new Paragraph(new Run(
                    new RunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }),
                    new Text("Beta")));
                var main = document.AddMainDocumentPart();
                main.Document = new Document(new Body(first, second));
                main.Document.Save();
            }

            var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
            var letters = Assert.Single(pdf.GetPages()).Letters.ToArray();
            return Assert.Single(letters.Where(letter => letter.Value == "B")).StartBaseLine.Y;
        }

        Assert.InRange(SecondBaseline(false) - SecondBaseline(true), -8.2, -7.9);
    }

    [Fact]
    public void FlowVisitorRendersMixedOmmlRunsAndRadicalInSourceOrder()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("Before ") { Space = SpaceProcessingModeValues.Preserve }),
                new MathXml.OfficeMath(
                    new MathXml.Run(new MathXml.Text("1")),
                    new MathXml.Radical(new MathXml.Degree(),
                        new MathXml.Base(new MathXml.Run(new MathXml.Text("x")))),
                    new MathXml.Run(new MathXml.Text("2"))),
                new Run(new Text(" after") { Space = SpaceProcessingModeValues.Preserve }))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = Assert.Single(pdf.GetPages()).Letters.ToArray();
        var first = Assert.Single(letters.Where(letter => letter.Value == "1"));
        var root = Assert.Single(letters.Where(letter => letter.Value == "√"));
        var last = Assert.Single(letters.Where(letter => letter.Value == "2"));
        Assert.True(first.GlyphRectangle.Left < root.GlyphRectangle.Left);
        Assert.True(root.GlyphRectangle.Left < last.GlyphRectangle.Left);
    }

    [Fact]
    public void FlowVisitorAvoidsDuplicateSpacesAtMixedMathBoundaries()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            "mixed-radical-probe__MixedRadicalOperators.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = Assert.Single(pdf.GetPages()).Letters.ToArray();
        var plus = Assert.Single(letters.Where(letter => letter.Value == "+"));
        var radical = Assert.Single(letters.Where(letter => letter.Value == "√"));
        var equals = Assert.Single(letters.Where(letter => letter.Value == "="));
        var following = Assert.Single(letters.Where(letter => letter.Value == "a"));
        var preceding = letters.Where(letter => letter.Value == "e" &&
            letter.StartBaseLine.X < plus.StartBaseLine.X)
            .MaxBy(letter => letter.StartBaseLine.X)!;
        Assert.InRange(plus.StartBaseLine.X - preceding.StartBaseLine.X, 7, 11.5);
        Assert.InRange(radical.StartBaseLine.X - plus.StartBaseLine.X, 9.5, 10.7);
        Assert.InRange(following.StartBaseLine.X - equals.StartBaseLine.X, 8, 13);
    }

    [Fact]
    public void FlowVisitorShiftsMixedDisplayMathRunWithoutMovingFraction()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            "display-mixed-omml-probe__source__display-mixed.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var plus = Assert.Single(page.Letters.Where(letter => letter.Value == "+"));
        var numerator = Assert.Single(page.Letters.Where(letter => letter.Value == "1"));
        var baselineDifference = numerator.StartBaseLine.Y - plus.StartBaseLine.Y;
        Assert.InRange(baselineDifference, 4.5, 5.7);
    }

    [Fact]
    public void FlowVisitorSpacesAlphanumericRunBeforeMixedRadical()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            "mixed-radical-probe__MixedRadical.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = Assert.Single(pdf.GetPages()).Letters.ToArray();
        var leading = Assert.Single(letters.Where(letter => letter.Value == "1"));
        var radical = Assert.Single(letters.Where(letter => letter.Value == "√"));
        Assert.InRange(radical.StartBaseLine.X - leading.StartBaseLine.X, 9.8, 10.6);
    }

    [Fact]
    public void FlowVisitorSelectsMathRadicalVariantByRadicandDescent()
    {
        var path = Path.Combine(Environment.GetFolderPath(
            Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc");
        if (!File.Exists(path)) return;
        var metrics = new DocxportNet.Doc.DocOpenTypeAdvances(File.ReadAllBytes(path), 1);
        Assert.True(DxpFlowPdfVisitor.ShouldUseTallerMathRadical(metrics, 0x1D466));
        Assert.True(DxpFlowPdfVisitor.ShouldUseTallerMathRadical(metrics, 0x1D454));
        Assert.True(DxpFlowPdfVisitor.ShouldUseTallerMathRadical(metrics, 0x1D457));
        Assert.False(DxpFlowPdfVisitor.ShouldUseTallerMathRadical(metrics, 0x1D465));
        Assert.False(DxpFlowPdfVisitor.ShouldUseTallerMathRadical(metrics, 0x1D461));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void FlowVisitorRendersMixedOmmlRunsAndScriptInSourceOrder(bool superscript)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var script = superscript
                ? (OpenXmlElement)new MathXml.Superscript(
                    new MathXml.Base(new MathXml.Run(new MathXml.Text("x"))),
                    new MathXml.SuperArgument(new MathXml.Run(new MathXml.Text("2"))))
                : new MathXml.Subscript(
                    new MathXml.Base(new MathXml.Run(new MathXml.Text("x"))),
                    new MathXml.SubArgument(new MathXml.Run(new MathXml.Text("2"))));
            main.Document = new Document(new Body(new Paragraph(new MathXml.OfficeMath(
                new MathXml.Run(new MathXml.Text("1")), script,
                new MathXml.Run(new MathXml.Text("3"))))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = Assert.Single(pdf.GetPages()).Letters.ToArray();
        var first = Assert.Single(letters.Where(letter => letter.Value == "1"));
        var baseLetter = Assert.Single(letters.Where(letter => letter.Value is "x" or "𝑥"));
        var exponent = Assert.Single(letters.Where(letter => letter.Value == "2"));
        var last = Assert.Single(letters.Where(letter => letter.Value == "3"));
        Assert.True(first.GlyphRectangle.Left < baseLetter.GlyphRectangle.Left);
        Assert.True(baseLetter.GlyphRectangle.Left < last.GlyphRectangle.Left);
        Assert.Equal(superscript,
            exponent.GlyphRectangle.Bottom > baseLetter.GlyphRectangle.Bottom);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void FlowVisitorRendersInlineOmmlScript(bool isSuperscript)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("Before ") { Space = SpaceProcessingModeValues.Preserve }),
                new MathXml.OfficeMath(isSuperscript
                    ? new MathXml.Superscript(
                        new MathXml.Base(new MathXml.Run(new MathXml.Text("x"))),
                        new MathXml.SuperArgument(new MathXml.Run(new MathXml.Text("2"))))
                    : new MathXml.Subscript(
                        new MathXml.Base(new MathXml.Run(new MathXml.Text("x"))),
                        new MathXml.SubArgument(new MathXml.Run(new MathXml.Text("2"))))),
                new Run(new Text(" after") { Space = SpaceProcessingModeValues.Preserve }))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var baseLetter = Assert.Single(page.Letters.Where(letter => letter.Value == "x"));
        var exponent = Assert.Single(page.Letters.Where(letter => letter.Value == "2"));
        Assert.InRange(baseLetter.GlyphRectangle.Left, 104, 108);
        Assert.InRange(exponent.GlyphRectangle.Left, 110, 114);
        Assert.Equal(isSuperscript,
            exponent.GlyphRectangle.Bottom > baseLetter.GlyphRectangle.Bottom);
        Assert.Contains("after", page.Text);
    }

    [Fact]
    public void FlowVisitorAdvancesPastInlineOmmlSuperscript()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            "mixed-script-probe__MixedSuperscript.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var lastMathRun = Assert.Single(page.Letters.Where(letter => letter.Value == "3"));
        Assert.InRange(lastMathRun.GlyphRectangle.Left, 125.2, 125.6);
        Assert.DoesNotContain("2 3", page.Text);
    }

    [Fact]
    public void FlowVisitorPreservesAdjacentRunSpaces()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            "nested-omml-probe__arial-source__ArialNoMath.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var words = Assert.Single(pdf.GetPages()).GetWords().ToArray();
        var before = Assert.Single(words.Where(word => word.Text == "Before"));
        var after = Assert.Single(words.Where(word => word.Text == "after"));
        Assert.InRange(after.BoundingBox.Left - before.BoundingBox.Right, 6.3, 7.0);
    }

    [Fact]
    public void FlowVisitorRendersInlineOmmlCombinedScript()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("Before ") { Space = SpaceProcessingModeValues.Preserve }),
                new MathXml.OfficeMath(new MathXml.SubSuperscript(
                    new MathXml.Base(new MathXml.Run(new MathXml.Text("x"))),
                    new MathXml.SubArgument(new MathXml.Run(new MathXml.Text("2"))),
                    new MathXml.SuperArgument(new MathXml.Run(new MathXml.Text("3"))))),
                new Run(new Text(" after") { Space = SpaceProcessingModeValues.Preserve }))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var subscript = Assert.Single(page.Letters.Where(letter => letter.Value == "2"));
        var superscript = Assert.Single(page.Letters.Where(letter => letter.Value == "3"));
        Assert.InRange(subscript.GlyphRectangle.Left - superscript.GlyphRectangle.Left, -2, 2);
        Assert.True(subscript.StartBaseLine.Y < superscript.StartBaseLine.Y);
        Assert.Contains("after", page.Text);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void FlowVisitorRendersInlineOmmlMatrixBetweenText(int dimension)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var rows = Enumerable.Range(0, dimension)
                .Select(row => new MathXml.MatrixRow(Enumerable.Range(0, dimension)
                    .Select(column => (DocumentFormat.OpenXml.OpenXmlElement)new MathXml.Base(
                        new MathXml.Run(new MathXml.Text(
                            ((char)('a' + row * dimension + column)).ToString())))).ToArray()))
                .Cast<DocumentFormat.OpenXml.OpenXmlElement>().ToArray();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("Before ") { Space = SpaceProcessingModeValues.Preserve }),
                new MathXml.OfficeMath(new MathXml.Matrix(rows)),
                new Run(new Text(" after") { Space = SpaceProcessingModeValues.Preserve }))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains("Before", page.Text);
        Assert.Contains("after", page.Text);
        var cells = page.Letters.Where(letter => letter.GlyphRectangle.Left is > 100 and < 170 &&
            Math.Abs(letter.PointSize - 12) < 0.1).ToArray();
        Assert.Equal(dimension * dimension, cells.Length);
        var columnCenters = cells.Select(letter =>
            (letter.GlyphRectangle.Left + letter.GlyphRectangle.Right) / 2)
            .OrderBy(center => center).ToArray();
        Assert.Equal(dimension, 1 + columnCenters.Zip(columnCenters.Skip(1))
            .Count(pair => pair.Second - pair.First > 5));
        Assert.Equal(dimension, cells.Select(letter => Math.Round(letter.StartBaseLine.Y)).Distinct().Count());
        if (dimension == 3)
        {
            var before = Assert.Single(page.Letters.Where(letter => letter.Value == "B"));
            Assert.InRange(cells.Max(letter => letter.StartBaseLine.Y) - before.StartBaseLine.Y,
                14.4, 15.6);
            var firstRow = cells.OrderByDescending(letter => letter.StartBaseLine.Y)
                .Take(3).OrderBy(letter => letter.GlyphRectangle.Left).ToArray();
            Assert.InRange(firstRow[1].GlyphRectangle.Left - firstRow[0].GlyphRectangle.Left,
                18.9, 19.5);
            Assert.InRange(firstRow[2].GlyphRectangle.Left - firstRow[0].GlyphRectangle.Left,
                37.7, 39.3);
        }
    }

    [Fact]
    public void FlowVisitorRendersMixedOmmlRunAndMatrixInSourceOrder()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var matrix = new MathXml.Matrix(
                new MathXml.MatrixRow(
                    new MathXml.Base(new MathXml.Run(new MathXml.Text("1"))),
                    new MathXml.Base(new MathXml.Run(new MathXml.Text("2")))),
                new MathXml.MatrixRow(
                    new MathXml.Base(new MathXml.Run(new MathXml.Text("3"))),
                    new MathXml.Base(new MathXml.Run(new MathXml.Text("4")))));
            main.Document = new Document(new Body(new Paragraph(new MathXml.OfficeMath(
                new MathXml.Run(new MathXml.Text("7")), matrix,
                new MathXml.Run(new MathXml.Text("8"))))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = Assert.Single(pdf.GetPages()).Letters.ToArray();
        var before = Assert.Single(letters.Where(letter => letter.Value == "7"));
        var after = Assert.Single(letters.Where(letter => letter.Value == "8"));
        foreach (var digit in new[] { "1", "2", "3", "4" })
        {
            var cell = Assert.Single(letters.Where(letter => letter.Value == digit));
            Assert.InRange(cell.GlyphRectangle.Left, before.GlyphRectangle.Right, after.GlyphRectangle.Left);
        }
    }

    [Theory]
    [InlineData("default", "i")]
    [InlineData("section", "5")]
    [InlineData("document", "i")]
    public void FlowVisitorRendersReferencedEndnoteAtDocumentEnd(string numberingScope, string expectedMarker)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var body = new Body(new Paragraph(
                new Run(new Text("Body with an endnote")),
                new Run(new EndnoteReference { Id = 2 })));
            if (numberingScope == "section")
                body.Append(new SectionProperties(new EndnoteProperties(
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new NumberingStart { Val = 5 })));
            else if (numberingScope == "document")
                body.Append(new SectionProperties());
            main.Document = new Document(body);
            main.Document.Save();
            var notes = main.AddNewPart<EndnotesPart>();
            notes.Endnotes = new Endnotes(
                new Endnote(new Paragraph(new Run(new Text("Orphaned endnote body")))) { Id = 1 },
                new Endnote(new Paragraph(
                    new Run(new EndnoteReferenceMark(), new Text("Unique endnote body")))) { Id = 2 });
            notes.Endnotes.Save();
            if (numberingScope == "document")
            {
                var settings = main.AddNewPart<DocumentSettingsPart>();
                settings.Settings = new Settings(new EndnoteDocumentWideProperties(
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new NumberingStart { Val = 5 }));
                settings.Settings.Save();
            }
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Single(pdf.GetPages());
        var page = pdf.GetPage(1);
        var bodyText = string.Concat(page.Letters.Where(letter =>
            page.Height - letter.GlyphRectangle.Top < 100).Select(letter => letter.Value));
        var noteText = string.Concat(page.Letters.Where(letter =>
            page.Height - letter.GlyphRectangle.Top > 100).Select(letter => letter.Value));
        Assert.EndsWith("Bodywithanendnote" + expectedMarker, bodyText);
        Assert.StartsWith(expectedMarker + "Uniqueendnotebody", noteText);
        Assert.DoesNotContain("Orphanedendnotebody", string.Concat(pdf.GetPages().Select(page => page.Text)));
        var noteLetters = page.Letters.Where(letter =>
            page.Height - letter.GlyphRectangle.Top > 100).ToArray();
        Assert.Equal(noteLetters.First(letter => letter.Value == "U").PointSize,
            noteLetters.First(letter => letter.Value == expectedMarker).PointSize);
    }

    [Theory]
    [InlineData("PdfFeatures", "EndnoteBasic", 1)]
    [InlineData("PdfFeatures", "EndnoteSectionEnd", 1)]
    [InlineData("PdfProbeCorpus", "endnote-probe__endnote-reordered", 2)]
    [InlineData("PdfProbeCorpus", "endnote-probe__endnote-two-sections", 2)]
    public void FlowVisitorMakesEndnoteReferencesActive(
        string fixtureGroup, string fixtureName, int expectedLinks)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", fixtureGroup,
            fixtureName + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var destinations = pdf.Internals.Catalog.Names!.NameTree!;
        var links = pdf.Pages.Cast<PdfSharp.Pdf.PdfPage>()
            .SelectMany(page => page.Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>())
            .Select(annotation => annotation.Elements.GetDictionary("/A")?.Elements.GetString("/D"))
            .Where(destination => destination?.StartsWith("dxp-endnote-", StringComparison.Ordinal) == true)
            .ToArray();
        Assert.Equal(expectedLinks, links.Length);
        for (var index = 1; index <= expectedLinks; index++)
        {
            var destination = $"dxp-endnote-{index}";
            Assert.Contains(destination, links);
            Assert.True(destinations.ContainsName(destination, true));
        }
    }

    [Fact]
    public void FlowVisitorOrdersEndnotesByReferenceRatherThanPartOrder()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("First reference"), new EndnoteReference { Id = 2 })),
                new Paragraph(new Run(new Text("Second reference"), new EndnoteReference { Id = 1 }))));
            main.Document.Save();
            var notes = main.AddNewPart<EndnotesPart>();
            notes.Endnotes = new Endnotes(
                new Endnote(new Paragraph(new Run(new EndnoteReferenceMark(),
                    new Text("First note body")))) { Id = 1 },
                new Endnote(new Paragraph(new Run(new EndnoteReferenceMark(),
                    new Text("Second note body")))) { Id = 2 });
            notes.Endnotes.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var text = string.Concat(pdf.GetPages().SelectMany(page => page.Letters)
            .Select(letter => letter.Value));
        Assert.Contains("FirstreferenceiSecondreferenceii", text);
        Assert.Contains("iSecondnotebodyiiFirstnotebody", text);
    }
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void FlowVisitorPlacesEndnotesAtDocumentWideSectionEnd(bool documentWide)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var firstSection = new SectionProperties();
            if (!documentWide)
                firstSection.Append(new EndnoteProperties(new EndnotePosition
                    { Val = EndnotePositionValues.SectionEnd }));
            main.Document = new Document(new Body(
                new Paragraph(new ParagraphProperties(firstSection),
                    new Run(new Text("First section reference"), new EndnoteReference { Id = 1 })),
                new Paragraph(new Run(new Text("Second section text"))),
                new SectionProperties()));
            main.Document.Save();
            var notes = main.AddNewPart<EndnotesPart>();
            notes.Endnotes = new Endnotes(new Endnote(new Paragraph(new Run(
                new EndnoteReferenceMark(), new Text("Section end note body")))) { Id = 1 });
            notes.Endnotes.Save();
            if (documentWide)
            {
                var settings = main.AddNewPart<DocumentSettingsPart>();
                settings.Settings = new Settings(new EndnoteDocumentWideProperties(
                    new EndnotePosition { Val = EndnotePositionValues.SectionEnd }));
                settings.Settings.Save();
            }
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("Sectionendnotebody", string.Concat(pdf.GetPage(documentWide ? 1 : 2)
            .Letters.Select(letter => letter.Value)));
        Assert.DoesNotContain("Sectionendnotebody", string.Concat(pdf.GetPage(documentWide ? 2 : 1)
            .Letters.Select(letter => letter.Value)));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void FlowVisitorPlacesMultipleEndnotesAtTheirSectionEnds(
        bool firstSectionDecimal, bool secondSectionDecimal)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var firstSection = new SectionProperties();
            if (firstSectionDecimal)
                firstSection.Append(new EndnoteProperties(
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new NumberingStart { Val = 5 }));
            var secondSection = new SectionProperties();
            if (secondSectionDecimal)
                secondSection.Append(new EndnoteProperties(
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new NumberingStart { Val = 5 }));
            main.Document = new Document(new Body(
                new Paragraph(new ParagraphProperties(firstSection),
                    new Run(new Text("First section"), new EndnoteReference { Id = 1 })),
                new Paragraph(new Run(new Text("Second section"), new EndnoteReference { Id = 2 })),
                secondSection));
            main.Document.Save();
            var notes = main.AddNewPart<EndnotesPart>();
            notes.Endnotes = new Endnotes(
                new Endnote(new Paragraph(new Run(new EndnoteReferenceMark(),
                    new Text("First note body")))) { Id = 1 },
                new Endnote(new Paragraph(new Run(new EndnoteReferenceMark(),
                    new Text("Second note body")))) { Id = 2 });
            notes.Endnotes.Save();
            var settings = main.AddNewPart<DocumentSettingsPart>();
            settings.Settings = new Settings(new EndnoteDocumentWideProperties(
                new EndnotePosition { Val = EndnotePositionValues.SectionEnd },
                new NumberingRestart { Val = RestartNumberValues.EachSection }));
            settings.Settings.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        var firstPage = string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value));
        var secondPage = string.Concat(pdf.GetPage(2).Letters.Select(letter => letter.Value));
        Assert.Contains("Firstsection" + (firstSectionDecimal ? "5" : "i"), firstPage);
        Assert.Contains((firstSectionDecimal ? "5" : "i") + "Firstnotebody", firstPage);
        Assert.DoesNotContain("Secondnotebody", firstPage);
        Assert.Contains("Secondsection" + (secondSectionDecimal ? "6" : "ii"), secondPage);
        Assert.Contains((secondSectionDecimal ? "6" : "ii") + "Secondnotebody", secondPage);
        Assert.DoesNotContain("Firstnotebody", secondPage);
    }
    [Theory]
    [InlineData("<span style=\"font-family:Symbol\">&#183;</span>", "•")]
    [InlineData("1.", "1.")]
    public void ListMarkersArePlainText(string marker, string expected)
    {
        Assert.Equal(expected, DxpPdfVisitor.NormalizeMarker(marker));
    }

    [Fact]
    public void ExportsWalkerTextToReadablePdf()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("First paragraph"))),
                new Paragraph(new Run(new Text("Second paragraph")))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpPdfVisitor());

        Assert.Equal("%PDF-", System.Text.Encoding.ASCII.GetString(bytes, 0, 5));
        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Single(pdf.Pages);
    }

    [Fact]
    public void FlowVisitorExportsMultiPageDocument()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("First page"))),
                new Paragraph(new Run(new Break { Type = BreakValues.Page }, new Text("Second page")))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Equal(2, pdf.PageCount);
        Assert.Equal(612, pdf.Pages[0].Width.Point);
        Assert.Equal(792, pdf.Pages[0].Height.Point);
    }

    [Fact]
    public void FlowVisitorKeepsSampleTableOnOnePage()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "TestTableSpan.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Single(pdf.Pages);
    }

    [Fact]
    public void FlowVisitorInsetsTextInFullyDoubleBorderedTable()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "TestTableSpan.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var firstCells = pdf.GetPage(1).GetWords().Where(word => word.Text == "1,1").ToArray();
        Assert.InRange(firstCells[0].BoundingBox.Left, 78.0, 78.3);
        Assert.InRange(firstCells[1].BoundingBox.Left, 205.3, 205.6);
    }

    [Fact]
    public void FlowVisitorCentersTextAcrossVerticallyMergedCells()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "TestTableSpan.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var tops = page.GetWords().Where(word => word.Text == "+")
            .Select(word => page.Height - word.BoundingBox.Top)
            .OrderBy(top => top).ToArray();
        Assert.Contains(tops, top => top is > 310 and < 312);
        Assert.Contains(tops, top => top is > 410 and < 412);
    }

    [Fact]
    public void FlowVisitorLetsDirectCellParagraphAlignmentOverrideConditionalStyle()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordLayeredCornerCenterParagraphAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var words = pdf.GetPage(1).GetWords().ToArray();
        Assert.Contains(words, word => word.Text == "B" && word.BoundingBox.Left is > 350 and < 355);
        Assert.Contains(words, word => word.Text == "D" && word.BoundingBox.Left is > 300 and < 370);
        Assert.Contains(words, word => word.Text == "B" && word.BoundingBox.Left is > 526 and < 530);
    }

    [Fact]
    public void FlowVisitorKeepsFirstTableCellSpacingInsideItsBorder()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            "spacing-probe__no-style-auto.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var textPdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        using var drawingPdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var candidates = new List<double>();
        foreach (var page in textPdf.GetPages())
        {
            var beforeWords = page.GetWords().Where(word => word.Text == "Before").ToArray();
            var afterWords = page.GetWords().Where(word => word.Text == "After").ToArray();
            if (beforeWords.Length == 0 || afterWords.Length == 0)
                continue;

            var drawingPage = drawingPdf.Pages[page.Number - 1];
            var content = string.Concat(drawingPage.Contents.Elements
                .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
                .Select(reference => reference.Value)
                .OfType<PdfSharp.Pdf.Advanced.PdfContent>()
                .Select(stream => System.Text.Encoding.Latin1.GetString(stream.Stream!.UnfilteredValue)));
            var borderTops = System.Text.RegularExpressions.Regex.Matches(content,
                    @"(?m)(?<left>-?[\d.]+) (?<y>-?[\d.]+) m\s+(?<right>-?[\d.]+) \k<y> l\s+S")
                .Select(match => drawingPage.Height.Point - double.Parse(match.Groups["y"].Value,
                    System.Globalization.CultureInfo.InvariantCulture))
                .ToArray();
            foreach (var before in beforeWords)
            {
                var headingTop = page.Height - before.BoundingBox.Top;
                if (!afterWords.Any(after => Math.Abs(after.BoundingBox.Top - before.BoundingBox.Top) < 1))
                    continue;
                candidates.AddRange(borderTops.Select(borderTop => headingTop - borderTop)
                    .Where(clearance => clearance is > 0 and < 12));
            }
        }

        Assert.Contains(candidates, clearance => clearance is > 5.5 and < 11);
    }

    [Fact]
    public void FlowVisitorKeepsStyledFixedTableBorderInsetAcrossAlignments()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordLayeredCornerParagraphAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var words = page.GetWords().Where(word =>
            page.Height - word.BoundingBox.Top is > 120 and < 170).ToArray();
        var rightAligned = words.Single(word => word.Text == "B" && word.BoundingBox.Left > 500);
        var leftAligned = words.Single(word => word.Text == "Body" &&
            word.BoundingBox.Left is > 300 and < 320);
        Assert.InRange(rightAligned.BoundingBox.Right, 535.0, 535.4);
        Assert.InRange(leftAligned.BoundingBox.Left, 311.8, 312.1);
    }

    [Fact]
    public void FlowVisitorAlignsNumberedTextAtListIndent()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "TestLists.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = pdf.GetPage(1).Letters;
        Assert.Contains(letters, letter => letter.Value == "d" &&
            letter.GlyphRectangle.Left is > 107 and < 109);
        Assert.Contains(letters, letter => letter.Value == "d" &&
            letter.GlyphRectangle.Left is > 395 and < 397);
        var romanMarkers = pdf.GetPage(2).Letters;
        Assert.Contains(romanMarkers, letter => letter.Value == "I" &&
            letter.GlyphRectangle.Left is > 80 and < 90);
        Assert.Contains(romanMarkers, letter => letter.Value == "I" &&
            letter.GlyphRectangle.Left is > 70 and < 80);
    }

    [Theory]
    [InlineData("WordRunFitTextBoldAllStories.docx")]
    [InlineData("WordRunFitTextItalicAllStories.docx")]
    public void FlowVisitorPositionsMixedScriptAfterEmphasizedFit(string filename)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", filename);

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var firstLine = page.Letters.First(letter => letter.Value == "S" &&
            page.Height - letter.GlyphRectangle.Top is > 90 and < 120);
        Assert.InRange(page.Height - firstLine.GlyphRectangle.Top, 101.2, 102.7);
    }

    [Fact]
    public void FlowVisitorUsesFullListIndentWhenMarkerIsSuppressed()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleNoMarkerListsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var starts = page.Letters.Where(letter => letter.Value == "F" &&
            letter.GlyphRectangle.Left is > 107 and < 109).ToArray();
        Assert.True(starts.Length >= 3);
        Assert.Contains(starts, letter => page.Height - letter.GlyphRectangle.Top is > 70 and < 77);
    }

    [Fact]
    public void FlowVisitorPreservesFirstLineIndentForNoSuffixHeading()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "file-sample_1MB.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.Letters, letter => letter.Value == "L" &&
            letter.GlyphRectangle.Left is > 34 and < 37 &&
            page.Height - letter.GlyphRectangle.Top is > 128 and < 145);
    }

    [Fact]
    public void FlowVisitorKeepsCenteredParagraphAnchoredImageFlow()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "file-sample_1MB.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(5, pdf.NumberOfPages);
        Assert.Contains(pdf.GetPage(5).GetWords(), word => word.Text == "In");
    }

    [Fact]
    public void FlowVisitorRetainsFileSampleTableInteriorBorders()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "file-sample_1MB.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(2);
        var rowBorders = page.Paths.Where(path => path.IsStroked &&
                path.GetBoundingRectangle() is { } bounds &&
                bounds.Left is > 92 and < 94 && bounds.Width is > 282 and < 285 &&
                bounds.Height < 0.5)
            .ToArray();
        Assert.True(rowBorders.Select(path => Math.Round(path.GetBoundingRectangle()!.Value.Top, 1))
            .Distinct().Count() >= 6);
    }

    [Fact]
    public void FlowVisitorDoesNotDisplayCachedAskResults()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "TestDocVariables.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(),
            options: new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Cache });

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var text = pdf.GetPage(2).Text;
        Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(text, "Bob").Count);
        Assert.Equal(1, System.Text.RegularExpressions.Regex.Matches(text, "Montreal").Count);
    }

    [Fact]
    public void FlowVisitorStartsAfterEmptyParagraphPageBreakAtTopMargin()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "TestLists.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(5, pdf.NumberOfPages);
        foreach (var pageNumber in new[] { 2, 3, 4, 5 })
        {
            var page = pdf.GetPage(pageNumber);
            var firstLine = page.Letters.Where(letter => letter.GlyphRectangle.Left > 100)
                .Max(letter => letter.GlyphRectangle.Top);
            Assert.InRange(page.Height - firstLine, 69, 78);
        }
    }

    [Fact]
    public void FlowVisitorKeepsContextualBulletListPitch()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "TestLists.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var lines = pdf.GetPage(1).GetWords().Where(word => word.Text == "default")
            .Select(word => pdf.GetPage(1).Height - word.BoundingBox.Top).Order().ToArray();
        Assert.Equal(17, lines.Length);
        Assert.InRange(lines[^1] - lines[0], 403, 405);
        Assert.Equal(5, pdf.NumberOfPages);
    }

    [Fact]
    public void FlowVisitorSpacesConsecutiveContextualBullets()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordMixedListsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var baselines = page.GetWords().Where(word => word.Text == "body")
            .Select(word => page.Height - word.Letters[0].StartBaseLine.Y)
            .Order().ToArray();
        Assert.Equal(3, baselines.Length);
        Assert.All(baselines.Zip(baselines.Skip(1), (first, second) => second - first),
            spacing => Assert.InRange(spacing, 16.6, 16.9));
    }

    [Fact]
    public void FlowVisitorKeepsEmptyListGroupSeparatorsAtWordPositions()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "TestLists.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(2);
        var words = page.GetWords().ToArray();
        var firstNumbered = words.First(word => word.Text == "numbered");
        var firstRoman = words.First(word => word.Text == "Roman");
        var firstItem = words.First(word => word.Text == "Item");
        Assert.InRange(firstNumbered.BoundingBox.Top - firstRoman.BoundingBox.Top, 175, 180);
        Assert.InRange(firstNumbered.BoundingBox.Top - firstItem.BoundingBox.Top, 320, 325);
    }

    [Fact]
    public void FlowVisitorKeepsLanguageOnlyEmptyListSeparatorsAtWordPitch()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "TestLists.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var words = pdf.GetPage(3).GetWords().ToArray();
        var lastItem = words.Last(word => word.Text == "Item");
        var firstFancy = words.First(word => word.Text == "Fancy");
        Assert.InRange(lastItem.BoundingBox.Top - firstFancy.BoundingBox.Top, 49, 51);
    }

    [Fact]
    public void FlowVisitorKeepsPreservedLeadingSpaceAfterNumberingTabOnSameLine()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "TestLists.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var words = pdf.GetPage(4).GetWords().ToArray();
        var marker = words.Single(word => word.Text == "a." &&
            word.BoundingBox.Left is > 195 and < 201 &&
            pdf.GetPage(4).Height - word.BoundingBox.Top is > 400 and < 410);
        var item = words.Single(word => word.Text == "Multi" &&
            word.BoundingBox.Left is > 218 and < 221 &&
            Math.Abs(word.BoundingBox.Top - marker.BoundingBox.Top) < 3);
        Assert.InRange(item.BoundingBox.Left, 218.5, 220);
    }

    [Fact]
    public void FlowVisitorPrintsWebHiddenRunsButNotVanishRuns()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new RunProperties(new WebHidden()), new Text("PRINTED")),
                new Run(new RunProperties(new Vanish()), new Text("HIDDEN")))));
            main.Document.Save();
        }
        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains("PRINTED", page.Text);
        Assert.DoesNotContain("HIDDEN", page.Text);
    }

    [Fact]
    public void FlowVisitorRetainsFirstSdtParagraphSpaceBefore()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "TOC.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(),
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Cache });

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var heading = page.Letters.First(letter => letter.Value == "C");
        Assert.InRange(page.Height - heading.GlyphRectangle.Top, 81, 89);
    }

    [Fact]
    public void FlowVisitorUsesTextRunSizeForDotTabLeaders()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new ParagraphProperties(new Tabs(new TabStop
                {
                    Val = TabStopValues.Right,
                    Leader = TabStopLeaderCharValues.Dot,
                    Position = 6000
                })),
                new Run(new RunProperties(new RunFonts { Ascii = "Arial" }, new FontSize { Val = "24" }),
                    new Text("Item"), new TabChar(), new Text("1")))));
            main.Document.Save();
        }
        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Contains(pdf.GetPage(1).Letters, letter => letter.Value == "." &&
            letter.PointSize is > 11.5 and < 12.5);
    }

    [Fact]
    public void FlowVisitorPrintsTocLinksWithoutWebHyperlinkColor()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                new Style(new StyleRunProperties(new DocumentFormat.OpenXml.Wordprocessing.Color { Val = "0563C1" },
                    new Underline { Val = UnderlineValues.Single }))
                { StyleId = "Hyperlink", Type = StyleValues.Character });
            main.Document = new Document(new Body(
                new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "TOC1" }),
                    new DocumentFormat.OpenXml.Wordprocessing.Hyperlink(
                        new Run(new RunProperties(new RunStyle { Val = "Hyperlink" }), new Text("TOC")))),
                new Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Hyperlink(
                    new Run(new RunProperties(new RunStyle { Val = "Hyperlink" }), new Text("WEB"))))));
            main.Document.Save();
        }
        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var tocColor = page.Letters.First(letter => letter.Value == "T").Color.ToRGBValues();
        var webColor = page.Letters.First(letter => letter.Value == "W").Color.ToRGBValues();
        Assert.True(tocColor.r < 0.1 && tocColor.g < 0.1 && tocColor.b < 0.1);
        Assert.True(webColor.b > 0.6 && webColor.r < 0.1);
    }

    [Fact]
    public void FlowVisitorUpdatesArabicPageFieldWithoutChangingRomanCachedResult()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            static Paragraph PageField(string instruction, string cached, bool hidden = false) => new(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(instruction)),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(hidden ? new RunProperties(new Vanish()) : new RunProperties(), new Text(cached)),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            main.Document = new Document(new Body(
                PageField(" PAGE \\* MERGEFORMAT ", "9"), PageField(" PAGE \\* ROMAN ", "IX"),
                PageField(" PAGE ", "7", hidden: true)));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Contains("1", pdf.GetPage(1).Text);
        Assert.Contains("IX", pdf.GetPage(1).Text);
        Assert.DoesNotContain("9", pdf.GetPage(1).Text);
        Assert.DoesNotContain("7", pdf.GetPage(1).Text);
    }

    [Fact]
    public void FlowVisitorReservesSpacingAfterFittedText()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var starts = page.Letters.Where(letter => letter.Value == "S" &&
            letter.GlyphRectangle.Left is > 71 and < 74)
            .Select(letter => page.Height - letter.GlyphRectangle.Top).ToArray();
        Assert.True(starts.Any(top => top is > 100 and < 125), string.Join(", ", starts));
        Assert.InRange(page.Letters.First(letter => letter.Value == "S" &&
            letter.GlyphRectangle.Left is > 71 and < 74 &&
            page.Height - letter.GlyphRectangle.Top is > 100 and < 125).PointSize, 11.9, 12.1);
    }

    [Fact]
    public void FlowVisitorPlacesCjkFollowupAfterDirectFittedRun()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextHighAnsiFontsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var secondTop = page.Letters.Where(letter => letter.Value == "S" &&
            letter.GlyphRectangle.Left is > 71 and < 73 &&
            page.Height - letter.GlyphRectangle.Top is > 90 and < 120)
            .Select(letter => page.Height - letter.GlyphRectangle.Top).Min();
        Assert.InRange(secondTop, 104.7, 106.0);
    }

    [Fact]
    public void FlowVisitorUsesHighAnsiFontForAccentedLatinGlyphs()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextHighAnsiFontsAllStories.docx");
        using var documentBytes = new MemoryStream();
        using (var input = File.OpenRead(source)) input.CopyTo(documentBytes);
        documentBytes.Position = 0;
        using (var document = WordprocessingDocument.Open(documentBytes, true))
        {
            document.MainDocumentPart!.HeaderParts.Single().Header!
                .Descendants<Text>().First().Text = "A¡£§©®±²³µ¶·ßÿé.—“”‘’€Ω漢B";
            document.MainDocumentPart.HeaderParts.Single().Header!.Save();
        }
        var bytes = DxpExport.ExportToBytes(documentBytes.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = pdf.GetPage(1).Letters;
        foreach (var character in "¡£§©®±²³µ¶·ßÿé—“”‘’€Ω")
            Assert.Contains(letters, letter => letter.Value == character.ToString() &&
                letter.FontName.Contains("Black", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(letters, letter => letter.Value == "A" &&
            !letter.FontName.Contains("Black", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(letters, letter => letter.Value == "B" &&
            !letter.FontName.Contains("Black", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(letters, letter => letter.Value == "." &&
            pdf.GetPage(1).Height - letter.GlyphRectangle.Top < 70 &&
            !letter.FontName.Contains("Black", StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData("PdfProbeCorpus", "highansi-script-probe__HighAnsiLatin1Probe.docx", "¡", 49.2)]
    [InlineData("PdfProbeCorpus", "highansi-script-probe__HighAnsiScriptProbe.docx", "é", 49.2)]
    [InlineData("Doc", "WordRunFitTextHighAnsiFontsAllStories.docx", "é", 49.21)]
    public void FlowVisitorUsesPaintFontForMixedHighAnsiHeaderLayout(
        string fixtureDirectory, string fileName, string glyph, double expectedBaseline)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", fixtureDirectory, fileName);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var firstHeaderGlyph = page.Letters.First(letter => letter.Value == glyph &&
            letter.FontName.Contains("Black", StringComparison.OrdinalIgnoreCase) &&
            page.Height - letter.StartBaseLine.Y < 70);
        Assert.InRange(page.Height - firstHeaderGlyph.StartBaseLine.Y,
            expectedBaseline - 0.05, expectedBaseline + 0.05);
    }

    [Fact]
    public void FlowVisitorKeepsExpandedTimesFitTextSearchable()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextInheritedParagraphStyleAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source,
            new DxpFlowPdfVisitor(availableFontFamilies: ["Times New Roman"]));

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.Letters, letter => letter.Value == "F" &&
            letter.FontName.Contains("Times New Roman", StringComparison.OrdinalIgnoreCase));
        using var drawnPdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Contains(drawnPdf.Pages[0].Contents.Elements, element =>
            (element as PdfSharp.Pdf.Advanced.PdfReference)?.Value is PdfSharp.Pdf.Advanced.PdfContent content &&
            content.Stream != null &&
            System.Text.Encoding.Latin1.GetString(content.Stream.UnfilteredValue) is string commands &&
            System.Text.RegularExpressions.Regex.IsMatch(commands,
                @"100 Tz\s+1\.\d+ Tc\s+\(Fitted\) Tj") &&
            commands.Contains("(body.) Tj"));
    }

    [Fact]
    public void FlowVisitorClearsMultiParagraphFittedHeaderBeforeBody()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextHeaderCjkFallback.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var bodyTop = page.Letters.Where(letter => letter.Value == "F" &&
            letter.GlyphRectangle.Left is > 71 and < 73)
            .Select(letter => page.Height - letter.GlyphRectangle.Top).Min();
        Assert.InRange(bodyTop, 85, 100);
        var cjkHeaderTop = page.Letters.Where(letter => letter.Value == "H" &&
            page.Height - letter.GlyphRectangle.Top is > 55 and < 85)
            .Select(letter => page.Height - letter.GlyphRectangle.Top).Min();
        Assert.InRange(cjkHeaderTop, 58.8, 59.5);
    }

    [Fact]
    public void FlowVisitorSpacesFittedFooterCjkContinuation()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextFooterCjkFallback.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var baselines = page.Letters.Where(letter => letter.Value == "F" &&
                page.Height - letter.StartBaseLine.Y > 700)
            .Select(letter => page.Height - letter.StartBaseLine.Y)
            .OrderBy(value => value).ToArray();
        Assert.Equal(2, baselines.Length);
        Assert.InRange(baselines[1] - baselines[0], 20.3, 20.7);
    }

    [Theory]
    [InlineData("WordRunFitTextHeaderCjkFallback", 112, 119, 122.5, 12)]
    [InlineData("WordRunFitTextFooterCjkFallback", 97, 102, 107.5, 12)]
    [InlineData("WordRunFitTextSupplementaryHanFallback", 97, 102, 109, 11)]
    public void FlowVisitorSizesLatinBodyAfterFittedParagraph(string name, double topMin, double topMax,
        double expectedBaseline, double expectedPointSize)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var initial = page.Letters.First(letter => letter.Value == "S" &&
            letter.GlyphRectangle.Left is > 71 and < 73 &&
            page.Height - letter.GlyphRectangle.Top is > 90 and < 140);
        Assert.InRange(initial.PointSize, expectedPointSize - 0.1, expectedPointSize + 0.1);
        Assert.InRange(page.Height - initial.GlyphRectangle.Top, topMin, topMax);
        Assert.InRange(page.Height - initial.StartBaseLine.Y, expectedBaseline - 0.5, expectedBaseline + 0.5);
    }


    [Theory]
    [InlineData("WordRunFitTextBoldAllStories", 110.6)]
    [InlineData("WordRunFitTextItalicAllStories", 110.6)]
    [InlineData("WordRunFitTextHighAnsiFontsAllStories", 114.1)]
    [InlineData("WordRunFitTextParagraphOverrideAllStories", 110.5)]
    [InlineData("WordRunFitTextCjkPunctuationFontsAllStories", 115.4)]
    public void FlowVisitorSpacesYuMinchoAfterUniformLatinFit(string name, double expectedBaseline)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var second = page.GetWords().First(word => word.Text == "Second");
        Assert.InRange(page.Height - second.Letters[0].StartBaseLine.Y,
            expectedBaseline - 0.7, expectedBaseline + 0.7);
    }

    [Fact]
    public void FlowVisitorSpacesUnfittedKoreanAndLatinText()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextKoreanFallback.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var line = page.Letters.Where(letter => page.Height - letter.GlyphRectangle.Top is > 90 and < 120 &&
            letter.GlyphRectangle.Left >= 72).ToArray();
        Assert.Contains(line, letter => letter.Value == "S" && letter.PointSize is > 11.9 and < 12.1);
        Assert.InRange(line.Max(letter => letter.EndBaseLine.X), 520, 540);
        var secondParagraph = line.First(letter => letter.Value == "S");
        Assert.InRange(page.Height - secondParagraph.StartBaseLine.Y, 109.9, 110.5);
    }

    [Theory]
    [InlineData("WordRunFitTextHeaderCjkFallback", 84, 105, 142, 149)]
    [InlineData("WordRunFitTextStyleAllStories", 70, 95, 132, 139)]
    public void FlowVisitorFitsUniformRunGroupsToDeclaredWidths(string name, double bodyTopMin,
        double bodyTopMax, double bodyWidthMin, double bodyWidthMax)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var letters = page.Letters;
        var headerLine = letters.Where(letter => page.Height - letter.GlyphRectangle.Top is > 30 and < 55 &&
            letter.StartBaseLine.X >= 72).ToArray();
        var bodyLine = letters.Where(letter => page.Height - letter.GlyphRectangle.Top > bodyTopMin &&
            page.Height - letter.GlyphRectangle.Top < bodyTopMax &&
            letter.StartBaseLine.X is >= 72 and < 230).ToArray();
        Assert.InRange(headerLine.Max(letter => letter.EndBaseLine.X) -
            headerLine.Min(letter => letter.StartBaseLine.X), 88, 92);
        Assert.InRange(bodyLine.Max(letter => letter.EndBaseLine.X) -
            bodyLine.Min(letter => letter.StartBaseLine.X), bodyWidthMin, bodyWidthMax);
    }

    [Theory]
    [InlineData("WordRunFitTextHeaderCjkFallback")]
    [InlineData("WordRunFitTextFooterCjkFallback")]
    public void FlowVisitorSizesCalibriFittedStoryRunsToTwelvePoints(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var header = page.Letters.First(letter => letter.Value == "H" &&
            page.Height - letter.GlyphRectangle.Top < 55);
        var footer = page.Letters.First(letter => letter.Value == "F" &&
            page.Height - letter.GlyphRectangle.Top > 700);
        Assert.InRange(header.PointSize, 12.0, 12.2);
        Assert.InRange(footer.PointSize, 12.0, 12.5);
        if (name == "WordRunFitTextFooterCjkFallback")
            Assert.InRange(page.Height - footer.GlyphRectangle.Top, 713.6, 714.5);
    }

    [Fact]
    public void FlowVisitorPositionsLatinFittedFooterNearWord()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextLayeredBoldItalicAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var footer = page.Letters.First(letter => letter.Value == "F" &&
            page.Height - letter.StartBaseLine.Y > 700);
        Assert.InRange(page.Height - footer.StartBaseLine.Y, 743.0, 743.5);
    }

    [Fact]
    public void FlowVisitorReservesSpacingAfterParagraphStyleFitText()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextParagraphStyleAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var firstTop = page.Letters.Where(letter => letter.Value == "F" &&
            letter.GlyphRectangle.Left is > 71 and < 73 &&
            page.Height - letter.GlyphRectangle.Top < 200)
            .Select(letter => page.Height - letter.GlyphRectangle.Top).Min();
        var secondTop = page.Letters.Where(letter => letter.Value == "S" &&
            letter.GlyphRectangle.Left is > 71 and < 73 &&
            page.Height - letter.GlyphRectangle.Top < 200)
            .Select(letter => page.Height - letter.GlyphRectangle.Top).Min();
        Assert.InRange(firstTop, 74.5, 76);
        Assert.InRange(secondTop, 102.2, 103);
        Assert.InRange(secondTop - firstTop, 25, 32);
        Assert.InRange(page.Letters.First(letter => letter.Value == "F" &&
            letter.GlyphRectangle.Left is > 71 and < 73 &&
            page.Height - letter.GlyphRectangle.Top < 200).PointSize, 11.9, 12.1);
    }

    [Fact]
    public void FlowVisitorKeepsParagraphStyleFitTextNaturalInHeaderAndFooter()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextParagraphStyleAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source,
            new DxpFlowPdfVisitor(availableFontFamilies: ["Times New Roman"]));

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var headerRight = page.Letters.Where(letter => letter.Value == "t" &&
            page.Height - letter.GlyphRectangle.Top < 60).Max(letter => letter.GlyphRectangle.Right);
        var footerRight = page.Letters.Where(letter => letter.Value == "t" &&
            page.Height - letter.GlyphRectangle.Top > 700).Max(letter => letter.GlyphRectangle.Right);
        var bodyRight = page.Letters.Where(letter => letter.Value == "." &&
            page.Height - letter.GlyphRectangle.Top is > 70 and < 95)
            .Max(letter => letter.GlyphRectangle.Right);

        Assert.InRange(headerRight, 155.5, 156.3);
        Assert.InRange(footerRight, 152.2, 153.0);
        Assert.InRange(bodyRight, 214.5, 217.5);
    }

    [Fact]
    public void FittedVisualWidthDistinguishesCalibriBoldFromBoldItalic()
    {
        Assert.Equal(88.75, DxpFlowPdfVisitor.FittedVisualTargetWidth(
            "Calibri,Bold", 90, tracking: false, story: true));
        Assert.Equal(90, DxpFlowPdfVisitor.FittedVisualTargetWidth(
            "Calibri,Bold", 90, tracking: false, story: false));
        Assert.Equal(90, DxpFlowPdfVisitor.FittedVisualTargetWidth(
            "Calibri,BoldItalic", 90, tracking: false, story: true));
        Assert.Equal(88.5, DxpFlowPdfVisitor.FittedVisualTargetWidth(
            "Calibri,Italic", 90, tracking: false, story: true));
        Assert.Equal(90, DxpFlowPdfVisitor.FittedVisualTargetWidth(
            "Calibri,Italic", 90, tracking: true, story: true));
    }


    [Fact]
    public void FlowVisitorLimitsKerningCarryToMatchingPlainLatinLines()
    {
        var eligible = new HashSet<string>(StringComparer.Ordinal)
        {
            "SecondparagraphhasonlyLatinlettersanddigits"
        };
        var plain = new[]
        {
            "-109.9767 -25.0299 Td", "100 Tz", "0 Tc",
            "[(Sec) 4.8828 (ond)] TJ", "38.7773 0 Td",
            "[(par) 21.9727 (agr) 21.9727 (aph)] TJ",
            "54.0234 0 Td", "(has) Tj", "19.8633 0 Td", "(only) Tj",
            "24.2344 0 Td", "(Latin) Tj", "27.2578 0 Td",
            "[(le) 6.8359 (t) 11.7188 (t) 12.207 (er) 12.207 (s)] TJ",
            "35.1152 0 Td", "(and) Tj", "21.5156 0 Td", "(digits.) Tj",
            "-109.9767 -20 Td", "(next paragraph) Tj"
        };

        Assert.True(DxpFlowPdfVisitor.IsPlainLatinAfterFitLine(plain, 1, eligible));
        Assert.False(DxpFlowPdfVisitor.IsPlainLatinAfterFitLine(
            plain, 1, new HashSet<string> { "Anotherparagraph" }));
        var mixed = plain.ToArray();
        mixed[8] = "<0B1B02EF07EA> Tj";
        Assert.False(DxpFlowPdfVisitor.IsPlainLatinAfterFitLine(mixed, 1, eligible));
    }

    [Fact]
    public void FlowVisitorTracksSeparateFitTextGroupsWithinParagraph()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextParagraphOverrideAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.Letters, letter => letter.Value == "t" &&
            letter.StartBaseLine.X is > 210 and < 222 &&
            letter.StartBaseLine.Y is > 680 and < 730);
        Assert.Contains(page.Letters, letter => letter.Value == "F" &&
            letter.GlyphRectangle.Left is > 71 and < 73 &&
            page.Height - letter.GlyphRectangle.Top < 200 &&
            letter.PointSize is > 11.9 and < 12.1);
        Assert.Contains(page.Letters, letter => letter.Value == "i" &&
            letter.StartBaseLine.X is > 84 and < 85 &&
            page.Height - letter.GlyphRectangle.Top is > 70 and < 80);
        Assert.Contains(page.Letters, letter => letter.Value == "S" &&
            letter.GlyphRectangle.Left is > 71 and < 73 &&
            page.Height - letter.GlyphRectangle.Top is > 90 and < 110);
        var footerTop = page.Letters.Where(letter => letter.Value == "F" &&
            letter.GlyphRectangle.Left is > 71 and < 73 &&
            page.Height - letter.GlyphRectangle.Top > 700)
            .Select(letter => page.Height - letter.GlyphRectangle.Top).Single();
        Assert.InRange(footerTop, 731, 736);
        var finalPeriod = page.Letters.Single(letter => letter.Value == "." &&
            page.Height - letter.GlyphRectangle.Top is > 70 and < 90);
        Assert.InRange(finalPeriod.GlyphRectangle.Left, 272.5, 274.5);
    }

    [Fact]
    public void FlowVisitorAccountsForHalfLeftBorderInCellText()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordConditionalFirstRowRunStyleAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = pdf.GetPage(1).Letters;
        var positions = letters.Where(letter => letter.Value == "B")
            .Select(letter => letter.GlyphRectangle.Left).ToArray();
        Assert.True(positions.Any(left => left is > 78.5 and < 79.5), string.Join(", ", positions));
        Assert.True(positions.Any(left => left is > 312.5 and < 313.5), string.Join(", ", positions));
    }

    [Fact]
    public void FlowVisitorKeepsDocumentVariablesSampleOnTwoPages()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "TestDocVariables.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(),
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Cache });

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Equal(2, pdf.PageCount);
    }

    [Fact]
    public void FlowVisitorUsesNormalStyleSpacingForWidowFixture()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", "WordWidowVisibleOn.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Equal(2, pdf.PageCount);
    }

    [Fact]
    public void FlowVisitorExportsMirroredIndentStoriesOnTwoPages()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyledParagraphShadingResetAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Equal(2, pdf.PageCount);
    }

    [Fact]
    public void FlowVisitorTopAlignsMixedSizeRunsWithoutOverridingBaselineReset()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordHyphenationStyleAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var words = page.GetWords().ToArray();
        var header = words.First(word => word.Text == "Mirrored" &&
            page.Height - word.BoundingBox.Top < 60);
        var headerTall = words.First(word => word.Text == "Tall" &&
            page.Height - word.BoundingBox.Top < 60);
        Assert.InRange(Math.Abs(header.BoundingBox.Top - headerTall.BoundingBox.Top), 0, 3);
        var reset = Assert.Single(words.Where(word => word.Text == "Reset"));
        var resetTall = words.First(word => word.Text == "Tall" &&
            page.Height - word.BoundingBox.Top is > 95 and < 125);
        Assert.InRange(Math.Abs(reset.BoundingBox.Top - resetTall.BoundingBox.Top), 8, 13);
        Assert.InRange(page.Height - reset.BoundingBox.Top, 123, 135);
        var footer = words.First(word => word.Text == "Mirrored" &&
            page.Height - word.BoundingBox.Top > 690);
        Assert.InRange(page.Height - footer.BoundingBox.Top, 711, 724);
    }

    [Fact]
    public void FlowVisitorDrawsRunHighlightPalette()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordHighlightPaletteAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var colors = pdf.GetPage(1).Paths.Where(path => path.IsFilled && path.FillColor != null)
            .Select(path => path.FillColor!.ToRGBValues()).Distinct().ToList();
        Assert.True(colors.Count >= 12);
        Assert.Contains(colors, color => color.r < 0.01 && color.g < 0.01 &&
            Math.Abs(color.b - 139.0 / 255) < 0.005);
        Assert.Contains(colors, color => color.r < 0.01 &&
            Math.Abs(color.g - 139.0 / 255) < 0.005 &&
            Math.Abs(color.b - 139.0 / 255) < 0.005);
        Assert.Contains(colors, color => color.r < 0.01 &&
            Math.Abs(color.g - 100.0 / 255) < 0.005 && color.b < 0.01);
        Assert.Contains(colors, color =>
            Math.Abs(color.r - 139.0 / 255) < 0.005 && color.g < 0.01 && color.b < 0.01);
        var page = pdf.GetPage(1);
        Assert.Contains(page.Paths, path => path.IsFilled &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 71 and < 73 &&
            page.Height - bounds.Top is > 64.5 and < 65.5 &&
            bounds.Height is > 13.7 and < 14);
        var paletteLine = page.Paths.Where(path => path.IsFilled && path.FillColor != null &&
            path.GetBoundingRectangle() is { } bounds &&
            page.Height - bounds.Top is > 64.5 and < 66 && bounds.Left is > 70 and < 190).ToArray();
        var yellow = paletteLine.Single(path => path.FillColor!.ToRGBValues() is var rgb &&
            rgb.r > 0.99 && rgb.g > 0.99 && rgb.b < 0.01);
        var black = paletteLine.Single(path => path.FillColor!.ToRGBValues() is var rgb &&
            rgb.r < 0.01 && rgb.g < 0.01 && rgb.b < 0.01);
        Assert.InRange(black.GetBoundingRectangle()!.Value.Left -
            yellow.GetBoundingRectangle()!.Value.Right,
            -0.2, 0.2);
        var cyanFills = page.Paths.Where(path => path.IsFilled && path.FillColor != null &&
            path.FillColor.ToRGBValues() is var rgb && rgb.r < 0.01 && rgb.g > 0.99 &&
            rgb.b > 0.99).ToArray();
        Assert.Contains(cyanFills, path => path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 71 and < 73 && bounds.Height is > 14 and < 15 &&
            page.Height - bounds.Top is > 35.5 and < 36.5);
        Assert.Contains(cyanFills, path => path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 71 and < 73 && bounds.Height is > 13.7 and < 14 &&
            page.Height - bounds.Top is > 79.5 and < 80.3);
        Assert.Contains(page.Letters, letter => letter.Value == "F" &&
            letter.PointSize is > 6.8 and < 7.25 &&
            page.Height - letter.GlyphRectangle.Top > 700);
    }

    [Fact]
    public void FlowVisitorKeepsCommentAnchorsButOmitsCommentStories()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "TestComments.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var text = pdf.GetPage(1).Text;
        Assert.Contains("BeforeComment", text);
        Assert.Contains("Comment", text);
        Assert.Contains("AnotherComment", text);
        Assert.DoesNotContain("FirstReply", text);
        Assert.DoesNotContain("Commentwithimage", text);
        using var clean = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.DoesNotContain(clean.Pages.Cast<PdfSharp.Pdf.PdfPage>()
            .SelectMany(page => page.Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>()),
            annotation => annotation.Elements.GetName("/Subtype") == "/Text");
    }

    [Fact]
    public void FlowVisitorExportsCommentThreadsAsPdfAnnotations()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "TestComments.docx");
        var bytes = DxpExport.ExportToBytes(source,
            new DxpFlowPdfVisitor(includePdfComments: true));

        using var annotated = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var comments = annotated.Pages.Cast<PdfSharp.Pdf.PdfPage>()
            .SelectMany(page => page.Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>())
            .ToArray();
        Assert.Equal(6, comments.Length);
        Assert.Contains(comments, comment => comment.Contents.Contains("Comment 1"));
        Assert.Contains(comments, comment => comment.Contents.Contains("First Reply"));
        Assert.Contains(comments, comment => comment.Contents.Contains("Another Comment"));
        var roots = comments.Where(comment => comment.Elements.GetReference("/IRT") == null).ToArray();
        Assert.Equal(2, roots.Length);
        var firstThread = Assert.Single(roots.Where(comment => comment.Contents == "Comment 1"));
        Assert.All(roots, rootAnnotation =>
        {
            Assert.Equal("/Highlight", rootAnnotation.Elements.GetName("/Subtype"));
            Assert.Equal(8, rootAnnotation.Elements.GetArray("/QuadPoints")!.Elements.Count);
            Assert.InRange(rootAnnotation.Rectangle.X1, 70, 73);
        });
        Assert.InRange(firstThread.Rectangle.X2, 122, 125);
        var secondThread = Assert.Single(roots.Where(comment => comment.Contents == "Another Comment"));
        Assert.InRange(secondThread.Rectangle.X2, 166, 170);
        Assert.True(secondThread.Rectangle.Y2 < firstThread.Rectangle.Y1);
        var replies = comments.Where(comment => comment.Elements.GetReference("/IRT") != null).ToArray();
        Assert.Equal(4, replies.Length);
        Assert.All(replies, reply =>
        {
            Assert.Equal("/Text", reply.Elements.GetName("/Subtype"));
            Assert.Equal("/R", reply.Elements.GetName("/RT"));
            Assert.Equal(firstThread.Reference!.ObjectID,
                reply.Elements.GetReference("/IRT")!.ObjectID);
        });
        Assert.All(comments, comment =>
        {
            Assert.Equal("Gaspard Petit", comment.Elements.GetString("/T"));
        });
        Assert.All(replies, reply => Assert.InRange(reply.Rectangle.X1,
            annotated.Pages[0].Width.Point - 30, annotated.Pages[0].Width.Point - 8));
    }

    [Fact]
    public void FlowVisitorAnchorsPdfCommentsOnTheirSourcePages()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfFeatures",
            "CommentMultiPage.docx");
        var bytes = DxpExport.ExportToBytes(source,
            new DxpFlowPdfVisitor(includePdfComments: true));

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Equal(2, pdf.PageCount);
        var first = Assert.Single(pdf.Pages[0].Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>()
            .Where(annotation => annotation.Elements.GetName("/Subtype") is "/Text" or "/Highlight"));
        var second = Assert.Single(pdf.Pages[1].Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>()
            .Where(annotation => annotation.Elements.GetName("/Subtype") is "/Text" or "/Highlight"));
        Assert.Contains("First page note", first.Contents);
        Assert.Contains("Second page note", second.Contents);
    }

    [Fact]
    public void FlowVisitorHighlightsCommentRangeAcrossPages()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfFeatures",
            "CommentCrossPageRange.docx");
        var bytes = DxpExport.ExportToBytes(source,
            new DxpFlowPdfVisitor(includePdfComments: true));

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Equal(2, pdf.PageCount);
        var highlights = new List<PdfSharp.Pdf.Annotations.PdfAnnotation>();
        for (var pageIndex = 0; pageIndex < pdf.PageCount; pageIndex++)
        {
            var highlight = Assert.Single(pdf.Pages[pageIndex].Annotations
                .Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>()
                .Where(annotation => annotation.Elements.GetName("/Subtype") == "/Highlight"));
            Assert.Equal("Cross-page note", highlight.Contents);
            Assert.Equal(8, highlight.Elements.GetArray("/QuadPoints")!.Elements.Count);
            highlights.Add(highlight);
        }
        Assert.True(highlights[0].Rectangle.X2 > highlights[1].Rectangle.X2 + 40);
    }

    [Fact]
    public void FlowVisitorExcludesRepeatedHeadersAndFootersFromCrossPageComment()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfFeatures",
            "CommentThreePageRange.docx");
        var bytes = DxpExport.ExportToBytes(source,
            new DxpFlowPdfVisitor(includePdfComments: true));

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Equal(3, pdf.PageCount);
        foreach (PdfSharp.Pdf.PdfPage page in pdf.Pages)
        {
            var highlight = Assert.Single(page.Annotations
                .Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>()
                .Where(annotation => annotation.Elements.GetName("/Subtype") == "/Highlight"));
            Assert.Equal("Three-page note", highlight.Contents);
            Assert.Equal(8, highlight.Elements.GetArray("/QuadPoints")!.Elements.Count);
        }
    }

    [Theory]
    [InlineData("240")]
    [InlineData("160")]
    public void FlowVisitorHighlightsMultilineCommentRange(string lineSpacing)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfFeatures",
            "CommentMultiPage.docx");
        using var stream = new MemoryStream();
        stream.Write(File.ReadAllBytes(source));
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var rangeEnd = document.MainDocumentPart!.Document.Descendants<CommentRangeEnd>()
                .Single(end => end.Id?.Value == "0");
            rangeEnd.Parent!.InsertBefore(new Run(
                new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
                new Text(" raised")), rangeEnd);
            rangeEnd.Parent!.InsertBefore(new Run(new Break(), new Text("Continuation")), rangeEnd);
            var paragraph = rangeEnd.Ancestors<Paragraph>().First();
            paragraph.ParagraphProperties ??= new ParagraphProperties();
            paragraph.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines
            {
                Line = lineSpacing,
                LineRule = LineSpacingRuleValues.Auto
            };
            document.MainDocumentPart.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(stream.ToArray(),
            new DxpFlowPdfVisitor(includePdfComments: true));
        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var first = Assert.Single(pdf.Pages[0].Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>());
        var second = Assert.Single(pdf.Pages[1].Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>());
        Assert.Equal("/Highlight", first.Elements.GetName("/Subtype"));
        Assert.Equal(lineSpacing == "240" ? 24 : 16,
            first.Elements.GetArray("/QuadPoints")!.Elements.Count);
        Assert.Equal("/Highlight", second.Elements.GetName("/Subtype"));
        using var text = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var raised = Assert.Single(text.GetPage(1).GetWords()
            .Where(word => word.Text == "raised"));
        Assert.InRange(raised.BoundingBox.Left, first.Rectangle.X1, first.Rectangle.X2);
        Assert.InRange(raised.BoundingBox.Right, first.Rectangle.X1, first.Rectangle.X2 + 1);
        Assert.InRange(raised.BoundingBox.Top, first.Rectangle.Y1, first.Rectangle.Y2);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FlowVisitorLimitsCommentHighlightToItsTableCell(bool multiline)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfFeatures",
            "CommentMultiPage.docx");
        using var stream = new MemoryStream();
        stream.Write(File.ReadAllBytes(source));
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var body = document.MainDocumentPart!.Document.Body!;
            var paragraph = body.Elements<Paragraph>().First();
            var left = (Paragraph)paragraph.CloneNode(true);
            var rangeEnd = left.Descendants<CommentRangeEnd>().Single();
            rangeEnd.Parent!.InsertBefore(new Run(new Text(" nearby extra phrase")), rangeEnd);
            if (multiline)
                rangeEnd.Parent!.InsertBefore(new Run(new Break(), new Text("Continuation")), rangeEnd);
            var right = new Paragraph(new Run(new Text("Other column first line")),
                new Run(new Break(), new Text("Other column second line")));
            var table = new Table(
                new TableProperties(new TableWidth { Width = "8000", Type = TableWidthUnitValues.Dxa }),
                new TableGrid(new GridColumn { Width = "4000" }, new GridColumn { Width = "4000" }),
                new TableRow(new TableCell(left), new TableCell(right)));
            body.ReplaceChild(table, paragraph);
            document.MainDocumentPart.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(stream.ToArray(),
            new DxpFlowPdfVisitor(includePdfComments: true));
        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var annotation = Assert.Single(pdf.Pages[0].Annotations
            .Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>());
        Assert.Equal("/Highlight", annotation.Elements.GetName("/Subtype"));
        Assert.Equal(multiline ? 16 : 8,
            annotation.Elements.GetArray("/QuadPoints")!.Elements.Count);
        Assert.True(annotation.Rectangle.X2 < 300);
        using var text = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var words = text.GetPage(1).GetWords().ToArray();
        var nearby = Assert.Single(words.Where(word => word.Text == "phrase"));
        var other = words.First(word => word.Text == "Other");
        Assert.InRange(other.BoundingBox.Left - nearby.BoundingBox.Right, 0, 60);
    }

    [Fact]
    public void FlowVisitorHighlightsPromotedNestedTableCommentRange()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            "table-nested-side-by-side-comment__source.docx");
        var bytes = DxpExport.ExportToBytes(source,
            new DxpFlowPdfVisitor(includePdfComments: true));

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Equal(2, pdf.PageCount);
        var first = Assert.Single(pdf.Pages[0].Annotations
            .Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>());
        var second = Assert.Single(pdf.Pages[1].Annotations
            .Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>());
        Assert.Equal("/Highlight", first.Elements.GetName("/Subtype"));
        Assert.Equal("/Highlight", second.Elements.GetName("/Subtype"));
        Assert.Equal(16, first.Elements.GetArray("/QuadPoints")!.Elements.Count);
        Assert.Equal(8, second.Elements.GetArray("/QuadPoints")!.Elements.Count);
        Assert.InRange(first.Rectangle.X2, 244, 249);
    }

    [Theory]
    [InlineData("table-nested-side-by-side-comment__source.docx", "Other", 300.53)]
    [InlineData("table-nested-side-by-side-short__source.docx", "Short", 409.51)]
    [InlineData("table-nested-side-by-side-long__source.docx", "Other", 264.17)]
    [InlineData("table-nested-side-by-side-two-row__source.docx", "Other", 300.53)]
    public void FlowVisitorKeepsSimpleNestedTableCellsSideBySide(
        string fileName, string rightWord, double wordLeft)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            fileName);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        var words = pdf.GetPage(1).GetWords().ToArray();
        var phrase = words.First(word => word.Text == "phrase");
        var other = words.First(word => word.Text == rightWord);
        Assert.InRange(other.BoundingBox.Top - phrase.BoundingBox.Top, -1, 1);
        Assert.InRange(other.BoundingBox.Left, wordLeft - 1.5, wordLeft + 1.5);
    }

    [Fact]
    public void FlowVisitorKeepsSecondNestedTableRowSideBySide()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            "table-nested-side-by-side-two-row__source.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        var words = pdf.GetPage(1).GetWords().ToArray();
        var left = words.First(word => word.Text == "Left");
        var right = words.First(word => word.Text == "Right");
        Assert.InRange(right.BoundingBox.Top - left.BoundingBox.Top, -1, 1);
        Assert.InRange(left.BoundingBox.Left, 82, 84);
        Assert.InRange(right.BoundingBox.Left, 299, 302);
        Assert.Contains("Second-pagecommentanchor.", pdf.GetPage(2).Text);
    }

    [Theory]
    [InlineData("PdfFeatures", "LegacyVmlRunImage.docx")]
    [InlineData("PdfFeatures", "LegacyVmlObjectPreview.docx")]
    [InlineData("PdfProbeCorpus", "vml-run-probe__vml-run-inch.docx")]
    [InlineData("PdfProbeCorpus", "vml-run-probe__vml-run-pixel.docx")]
    public void FlowVisitorTrimsUnpreservedLeadingSpaceAfterVmlImage(
        string fixtureDirectory, string fileName)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", fixtureDirectory, fileName);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var words = pdf.GetPage(1).GetWords().ToArray();
        var before = words.First(word => word.Text == "Before");
        var after = words.First(word => word.Text == "After");
        Assert.InRange(after.BoundingBox.Left - before.BoundingBox.Left, 103.5, 105.5);
    }

    [Theory]
    [InlineData("TextBoxTableNextPageRow.docx")]
    [InlineData("TextBoxTableNextPageRowBehind.docx")]
    [InlineData("TextBoxTableVerticalMergePage.docx")]
    [InlineData("TextBoxTableVerticalMergePageBehind.docx")]
    [InlineData("TextBoxTableDashedPage.docx")]
    [InlineData("TextBoxTableDashedPageBehind.docx")]
    [InlineData("TextBoxTableDottedPage.docx")]
    [InlineData("TextBoxTableDottedPageBehind.docx")]
    [InlineData("TextBoxTableRepeatHeader.docx")]
    public void FlowVisitorReservesTableFragmentTopBorderForText(string fileName)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes", fileName);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        var firstPage = pdf.GetPage(1);
        var secondPage = pdf.GetPage(2);
        var firstTop = firstPage.Height - firstPage.GetWords()
            .First(word => word.Text == "Previous").BoundingBox.Top;
        var secondWord = fileName.Contains("NextPageRow", StringComparison.Ordinal) ||
            fileName.Contains("RepeatHeader", StringComparison.Ordinal) ? "Previous" : "Text";
        var secondTop = secondPage.Height - secondPage.GetWords()
            .First(word => word.Text == secondWord).BoundingBox.Top;
        Assert.InRange(secondTop - firstTop, -0.15, 0.15);
    }

    [Theory]
    [InlineData("TextBoxTableDashedPage.docx")]
    [InlineData("TextBoxTableDashedPageBehind.docx")]
    [InlineData("TextBoxTableDottedPage.docx")]
    [InlineData("TextBoxTableDottedPageBehind.docx")]
    public void FlowVisitorReservesTableFragmentTopBorderForFigure(string fileName)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes", fileName);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var words = pdf.GetPage(2).GetWords().ToArray();
        var body = words.First(word => word.Text == "Text");
        var figure = words.First(word => word.Text == "Figure:");
        Assert.InRange(body.BoundingBox.Top - figure.BoundingBox.Top, 46.4, 47.4);
    }

    [Fact]
    public void FlowVisitorPreservesResolvedCommentState()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "TestComments.docx");
        using var stream = new MemoryStream();
        stream.Write(File.ReadAllBytes(source));
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var comment = document.MainDocumentPart!.WordprocessingCommentsExPart!.CommentsEx!
                .Elements<DocumentFormat.OpenXml.Office2013.Word.CommentEx>().First();
            comment.Done = true;
            document.MainDocumentPart.WordprocessingCommentsExPart.CommentsEx.Save();
        }

        var bytes = DxpExport.ExportToBytes(stream.ToArray(),
            new DxpFlowPdfVisitor(includePdfComments: true));
        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var comments = pdf.Pages.Cast<PdfSharp.Pdf.PdfPage>()
            .SelectMany(page => page.Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>())
            .ToArray();
        var resolved = Assert.Single(comments.Where(annotation => annotation.Contents == "Comment 1"));
        Assert.Equal("Review", resolved.Elements.GetString("/StateModel"));
        Assert.Equal("Completed", resolved.Elements.GetString("/State"));
        Assert.All(comments.Where(annotation => annotation.Contents != "Comment 1"),
            annotation => Assert.Equal(string.Empty, annotation.Elements.GetString("/State")));
    }

    [Fact]
    public void FlowVisitorLinksNestedRepliesEvenWhenChildSortsBeforeParent()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "TestComments.docx");
        using var stream = new MemoryStream();
        stream.Write(File.ReadAllBytes(source));
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var commentsEx = document.MainDocumentPart!.WordprocessingCommentsExPart!.CommentsEx!;
            var entries = commentsEx.Elements<DocumentFormat.OpenXml.Office2013.Word.CommentEx>().ToArray();
            entries[2].ParaIdParent = entries[1].ParaId!.Value;
            commentsEx.Save();
            var comments = document.MainDocumentPart.WordprocessingCommentsPart!.Comments!
                .Elements<Comment>().ToArray();
            comments[1].Date = new DateTime(2020, 1, 2, 0, 0, 0, DateTimeKind.Utc);
            comments[2].Date = new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            document.MainDocumentPart.WordprocessingCommentsPart.Comments.Save();
        }

        var bytes = DxpExport.ExportToBytes(stream.ToArray(),
            new DxpFlowPdfVisitor(includePdfComments: true));
        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var annotations = pdf.Pages.Cast<PdfSharp.Pdf.PdfPage>()
            .SelectMany(page => page.Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>())
            .ToArray();
        var parent = Assert.Single(annotations.Where(annotation => annotation.Contents == "First Reply"));
        var child = Assert.Single(annotations.Where(annotation => annotation.Contents.StartsWith("Comment with image")));
        Assert.Equal(parent.Reference!.ObjectID, child.Elements.GetReference("/IRT")!.ObjectID);
    }

    [Fact]
    public void FlowVisitorRendersParagraphLevelRubyBaseAndGuide()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfFeatures", "RubyBasic.docx");
        using var stream = new MemoryStream(File.ReadAllBytes(source));
        using var document = WordprocessingDocument.Open(stream, true);
        var ruby = document.MainDocumentPart!.Document.Descendants<OpenXmlUnknownElement>()
            .Single(element => element.LocalName == "ruby");
        ruby.InnerXml = ruby.InnerXml.Replace("日本", "base").Replace("にほん", "guide");
        document.MainDocumentPart.Document.Save();
        var bytes = DxpExport.ExportToBytes(document, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var text = pdf.GetPage(1).Text;
        Assert.Contains("base", text);
        Assert.Contains("guide", text);
    }


    [Theory]
    [InlineData("」", "Yu Mincho")]
    [InlineData("「", "MS Mincho")]
    [InlineData("，", "MS Mincho")]
    [InlineData("。", "MS Mincho")]
    public void FlowVisitorUsesWordFallbackForFullWidthPunctuation(string punctuation, string expectedFont)
    {
        Assert.Equal(expectedFont, DxpFlowPdfVisitor.SelectFullWidthPunctuationFont(
            punctuation, "MS Mincho", yuMinchoAvailable: true));
        Assert.Equal("MS Mincho", DxpFlowPdfVisitor.SelectFullWidthPunctuationFont(
            punctuation, "MS Mincho", yuMinchoAvailable: false));
    }

    [Theory]
    [InlineData("TextBoxSimple", 117.5, 121.5)]
    [InlineData("TextBoxSquareWrap", 151.5, 233.5)]
    [InlineData("TextBoxInsets", 128.3, 123.9)]
    [InlineData("TextBoxCentered", 117.5, 158.7)]
    [InlineData("TextBoxCenteredInsets", 128.3, 155.7)]
    public void FlowVisitorPositionsTextBoxAndPreservesAnchorParagraph(string name,
        double expectedLeft, double expectedTop)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes", name + ".docx");
        using (var document = WordprocessingDocument.Open(source, false))
        {
            Assert.Contains(document.MainDocumentPart!.Document.Descendants<TextBoxContent>(),
                box => box.Ancestors<Drawing>().Any(drawing =>
                    drawing.Descendants().Any(element => element.LocalName == "anchor")));
            var choice = Assert.Single(document.MainDocumentPart.Document.Descendants<AlternateContentChoice>());
            Assert.Equal(name != "TextBoxSquareWrap", new DxpFlowPdfVisitor().AcceptAlternateContentChoice(
                choice, new[] { "wps" }, null!));
        }
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.GetWords(), word => word.Text == "probe:");
        var figure = Assert.Single(page.GetWords().Where(word => word.Text == "Figure:"));
        Assert.InRange(figure.BoundingBox.Left, expectedLeft - 1, expectedLeft + 1);
        Assert.InRange(page.Height - figure.BoundingBox.Top, expectedTop - 1, expectedTop + 1);
        if (name == "TextBoxSimple")
        {
            using var painted = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
            var commands = string.Concat(painted.Pages[0].Contents.Elements
                .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
                .Select(reference => System.Text.Encoding.Latin1.GetString(
                    ((PdfSharp.Pdf.Advanced.PdfContent)reference.Value!).Stream!.UnfilteredValue)));
            Assert.True(commands.IndexOf("(The) Tj", StringComparison.Ordinal) <
                commands.IndexOf("(Figure:) Tj", StringComparison.Ordinal));
        }
    }

    [Fact]
    public void FlowVisitorPositionsTextBoxAfterPageBreakBefore()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes",
            "TextBoxLaterPage.docx");
        using (var document = WordprocessingDocument.Open(source, false))
            Assert.Contains(document.MainDocumentPart!.Document.Descendants<Paragraph>(),
                paragraph => paragraph.ParagraphProperties?.PageBreakBefore != null &&
                             paragraph.Descendants<TextBoxContent>().Any());

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain("Figure:", pdf.GetPage(1).Text);
        var secondPage = pdf.GetPage(2);
        var figure = Assert.Single(secondPage.GetWords().Where(word => word.Text == "Figure:"));
        Assert.InRange(figure.BoundingBox.Left, 116.5, 118.5);
        Assert.InRange(secondPage.Height - figure.BoundingBox.Top, 120.5, 122.5);
    }

    [Theory]
    [InlineData("vml-gif-probe__drawing-gif-line-exact12.docx", "9.6")]
    [InlineData("vml-gif-probe__drawing-gif-line-exact16.docx", "12.8")]
    [InlineData("vml-gif-probe__drawing-gif-line-exact20.docx", "16")]
    [InlineData("vml-gif-probe__drawing-png-line-exact12.docx", "9.6")]
    public void FlowVisitorClipsTallInlineImageToExactLineBox(string fileName, string clipHeight)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            fileName);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var commands = string.Concat(pdf.Pages[0].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Select(reference => System.Text.Encoding.Latin1.GetString(
                ((PdfSharp.Pdf.Advanced.PdfContent)reference.Value!).Stream!.UnfilteredValue)));
        Assert.Contains($"72 {clipHeight} re W n", commands);
    }




    [Fact]
    public void FlowVisitorPreservesRotatedShapeAnchorParagraph()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes", "TextBoxRotated.docx");
        using (var document = WordprocessingDocument.Open(source, false))
            Assert.Contains(document.MainDocumentPart!.Document.Descendants(),
                element => element.LocalName == "xfrm" && element.GetAttribute("rot", "").Value == "1800000");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var text = page.Text;
        Assert.Contains("Textboxprobe:Rotated", text);
        Assert.Contains("Figure:Rotated", text);
        Assert.Contains(page.Letters, letter => letter.Value == "T" &&
            page.Height - letter.GlyphRectangle.Top is > 240 and < 250);
    }

    [Fact]
    public void FlowVisitorAppliesDrawingShadeInLinearLight()
    {
        var baseColor = MigraDoc.DocumentObjectModel.Color.FromRgb(0x15, 0x60, 0x82);
        var shaded = DxpFlowPdfVisitor.ApplyDrawingShade(baseColor, 0.15);
        Assert.Equal(4u, shaded.R);
        Assert.Equal(36u, shaded.G);
        Assert.Equal(51u, shaded.B);
    }

    [Theory]
    [InlineData("TextBoxRotated", null)]
    [InlineData("TextBoxRotatedForeground", "1 w")]
    [InlineData("TextBoxRotatedForegroundLine1", ".5 w")]
    [InlineData("TextBoxRotatedForegroundLine3", "1.5 w")]
    public void FlowVisitorUsesDrawingThemeColorsForTextBoxTextAndStroke(string name,
        string? expectedStrokeWidth)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var commands = string.Concat(pdf.Pages[0].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Select(reference => System.Text.Encoding.Latin1.GetString(
                ((PdfSharp.Pdf.Advanced.PdfContent)reference.Value!).Stream!.UnfilteredValue)));
        var textPosition = commands.IndexOf("(Figure:) Tj", StringComparison.Ordinal);
        Assert.True(textPosition > 0);
        var colors = System.Text.RegularExpressions.Regex.Matches(commands[..textPosition],
            @"(?m)^(?:\d+(?:\.\d+)? ){2}\d+(?:\.\d+)? rg$");
        Assert.NotEmpty(colors);
        Assert.Equal("1 1 1 rg", colors[^1].Value);
        Assert.Contains(".016 .141 .2 RG", commands);
        if (expectedStrokeWidth != null)
            Assert.Contains(expectedStrokeWidth, commands);
    }

    [Theory]
    [InlineData(1, 0.5)]
    [InlineData(2, 1.0)]
    [InlineData(3, 1.5)]
    public void FlowVisitorReadsThemeLineStyleWidth(int styleIndex, double expectedWidth)
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes",
            "TextBoxRotatedForeground.docx");
        using var document = WordprocessingDocument.Open(source, false);

        Assert.Equal(expectedWidth, DxpFlowPdfVisitor.ReadDrawingThemeLineWidth(document, styleIndex));
    }

    [Fact]
    public void FlowVisitorStacksOverlappingTextBoxesByRelativeHeight()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes",
            "TextBoxOverlapping.docx");
        using (var document = WordprocessingDocument.Open(source, false))
        {
            var heights = document.MainDocumentPart!.Document.Descendants()
                .Where(element => element.LocalName == "anchor")
                .Select(element => long.Parse(element.GetAttribute("relativeHeight", "").Value))
                .ToArray();
            Assert.Equal(2, heights.Length);
            Assert.True(heights[0] > heights[1]);
        }

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Single(pdf.GetPages());
        var text = pdf.GetPage(1).Text;
        Assert.Contains("Figure:Overlapping", text);
        Assert.Contains("Frontboxcontent", text);
        Assert.True(text.IndexOf("Figure:Overlapping", StringComparison.Ordinal) <
            text.IndexOf("Frontboxcontent", StringComparison.Ordinal));
    }

    [Fact]
    public void FlowVisitorPaintsBehindDocumentTextBoxBeforeBodyText()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes",
            "TextBoxBehindBody.docx");
        using (var document = WordprocessingDocument.Open(source, false))
            Assert.Contains(document.MainDocumentPart!.Document.Descendants(),
                element => element.LocalName == "anchor" &&
                    element.GetAttribute("behindDoc", "").Value == "1");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Single(pdf.GetPages());
        var text = pdf.GetPage(1).Text;
        Assert.Contains("Figure:Simple", text);
        Assert.Contains("surroundingparagraph", text);
        Assert.True(text.IndexOf("Figure:Simple", StringComparison.Ordinal) <
            text.IndexOf("surroundingparagraph", StringComparison.Ordinal));
    }

    [Fact]
    public void FlowVisitorPositionsPageRelativeTextBoxOverHeader()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes",
            "TextBoxBehindHeader.docx");
        using (var document = WordprocessingDocument.Open(source, false))
            Assert.Contains(document.MainDocumentPart!.Document.Descendants(),
                element => element.LocalName == "positionV" &&
                    element.GetAttribute("relativeFrom", "").Value == "page");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains("HEADERMARKER", page.Text);
        var figure = Assert.Single(page.GetWords().Where(word => word.Text == "Figure:"));
        Assert.InRange(figure.BoundingBox.Left, 116.5, 118.5);
        Assert.InRange(page.Height - figure.BoundingBox.Top, 30, 45);
    }

    [Fact]
    public void FlowVisitorPositionsPageRelativeTextBoxHorizontally()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes",
            "TextBoxPageHorizontal.docx");
        using (var document = WordprocessingDocument.Open(source, false))
            Assert.Contains(document.MainDocumentPart!.Document.Descendants(),
                element => element.LocalName == "positionH" &&
                    element.GetAttribute("relativeFrom", "").Value == "page");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var figure = Assert.Single(page.GetWords().Where(word => word.Text == "Figure:"));
        Assert.InRange(figure.BoundingBox.Left, 44.5, 46.5);
        Assert.InRange(page.Height - figure.BoundingBox.Top, 30, 45);
    }

    [Fact]
    public void FlowVisitorPositionsMarginRelativeTextBoxAfterPrecedingParagraph()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes",
            "TextBoxMarginVertical.docx");
        using (var document = WordprocessingDocument.Open(source, false))
            Assert.Contains(document.MainDocumentPart!.Document.Descendants(),
                element => element.LocalName == "positionV" &&
                    element.GetAttribute("relativeFrom", "").Value == "margin");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains(page.GetWords(), word => word.Text == "Preceding");
        var figure = Assert.Single(page.GetWords().Where(word => word.Text == "Figure:"));
        Assert.InRange(page.Height - figure.BoundingBox.Top, 120.5, 122.5);
    }

    [Fact]
    public void FlowVisitorPositionsTextBoxInSecondColumn()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes",
            "TextBoxSecondColumn.docx");
        using (var document = WordprocessingDocument.Open(source, false))
        {
            Assert.True(document.MainDocumentPart!.Document.Body!.GetFirstChild<SectionProperties>()?
                .GetFirstChild<Columns>()?.ColumnCount?.Value == 2);
            Assert.Contains(document.MainDocumentPart.Document.Descendants<TextBoxContent>(),
                box => box.Ancestors<Drawing>().Any());
        }

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var words = page.GetWords().ToArray();
        Assert.InRange(Assert.Single(words.Where(word => word.Text == "probe:")).BoundingBox.Left,
            365, 375);
        Assert.InRange(Assert.Single(words.Where(word => word.Text == "Figure:")).BoundingBox.Left,
            368, 371);
        Assert.Contains(words, word => word.Text == "Second" && word.BoundingBox.Left > 320);
    }

    [Theory]
    [InlineData("TextBoxSecondColumnContinuation", 196.5)]
    [InlineData("TextBoxSecondColumnContinuationBehind", 153.5)]
    public void FlowVisitorKeepsColumnAnchoredTextBoxWithLongFlow(string name, double expectedFigureTop)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes",
            name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var pages = pdf.GetPages().ToArray();
        Assert.Equal(3, pages.Length);
        Assert.DoesNotContain("Figure:", pages[0].Text);
        Assert.DoesNotContain("Figure:", pages[1].Text);
        var figure = Assert.Single(pages[2].GetWords().Where(word => word.Text == "Figure:"));
        Assert.InRange(figure.BoundingBox.Left, 117, 118);
        Assert.InRange(pages[2].Height - figure.BoundingBox.Top,
            expectedFigureTop - 1, expectedFigureTop + 1);
        for (var index = 1; index <= 80; index++)
            Assert.Contains($"Columnflowline{index:000}",
                string.Concat(pages.Select(page => page.Text)));
    }

    [Fact]
    public void FlowVisitorPaintsSecondColumnBehindTextBoxBeforeBodyText()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes",
            "TextBoxSecondColumnBehind.docx");
        using (var document = WordprocessingDocument.Open(source, false))
            Assert.Contains(document.MainDocumentPart!.Document.Descendants(),
                element => element.LocalName == "anchor" &&
                    element.GetAttribute("behindDoc", "").Value == "1");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var figure = Assert.Single(page.GetWords().Where(word => word.Text == "Figure:"));
        Assert.InRange(figure.BoundingBox.Left, 368, 371);
        Assert.True(page.Text.IndexOf("Figure:Simple", StringComparison.Ordinal) <
            page.Text.IndexOf("Secondcolumnbody", StringComparison.Ordinal));
    }

    [Fact]
    public void FlowVisitorPaintsBehindDocumentTextBoxOverFooter()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes",
            "TextBoxBehindFooter.docx");
        using (var document = WordprocessingDocument.Open(source, false))
        {
            Assert.NotNull(document.MainDocumentPart!.FooterParts.SingleOrDefault());
            Assert.Contains(document.MainDocumentPart.Document.Descendants(),
                element => element.LocalName == "anchor" &&
                    element.GetAttribute("behindDoc", "").Value == "1");
        }

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var figure = Assert.Single(page.GetWords().Where(word => word.Text == "Figure:"));
        Assert.InRange(page.Height - figure.BoundingBox.Top, 670, 680);
        Assert.True(page.Text.IndexOf("FOOTERMARKER", StringComparison.Ordinal) <
            page.Text.IndexOf("Figure:Simple", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("TextBoxTableCell", false, 123.1, 122.0, 1)]
    [InlineData("TextBoxTableCellBehind", true, 123.1, 122.0, 1)]
    [InlineData("TextBoxTableSecondCell", false, 195.1, 122.0, 1)]
    [InlineData("TextBoxTableSecondCellBehind", true, 195.1, 122.0, 1)]
    [InlineData("TextBoxTableSecondRow", false, 123.1, 146.0, 1)]
    [InlineData("TextBoxTableSecondRowBehind", true, 123.1, 146.0, 1)]
    [InlineData("TextBoxTableNextPageRow", false, 123.1, 217.45, 2)]
    [InlineData("TextBoxTableNextPageRowBehind", true, 123.1, 217.45, 2)]
    [InlineData("TextBoxTableMergedCell", false, 195.1, 122.0, 1)]
    [InlineData("TextBoxTableMergedCellBehind", true, 195.1, 122.0, 1)]
    [InlineData("TextBoxTableRepeatHeader", false, 123.1, 241.93, 2)]
    [InlineData("TextBoxTableRepeatHeaderBehind", true, 123.1, 241.93, 2)]
    [InlineData("TextBoxTableVerticalMerge", false, 123.1, 122.0, 1)]
    [InlineData("TextBoxTableVerticalMergeBehind", true, 123.1, 122.0, 1)]
    [InlineData("TextBoxTableVerticalMergeBorder", false, 123.1, 122.0, 1)]
    [InlineData("TextBoxTableVerticalMergeBorderBehind", true, 123.1, 122.0, 1)]
    [InlineData("TextBoxTableVerticalMergePage", false, 123.1, 121.5, 2)]
    [InlineData("TextBoxTableVerticalMergePageBehind", true, 123.1, 121.5, 2)]
    [InlineData("TextBoxTableDashedPage", false, 123.1, 122.5, 2)]
    [InlineData("TextBoxTableDashedPageBehind", true, 123.1, 122.5, 2)]
    [InlineData("TextBoxTableDottedPage", false, 123.1, 122.5, 2)]
    [InlineData("TextBoxTableDottedPageBehind", true, 123.1, 122.5, 2)]
    [InlineData("TextBoxTableDoublePage", false, 123.1, 121.5, 2)]
    [InlineData("TextBoxTableDoublePageBehind", true, 123.1, 121.5, 2)]
    public void FlowVisitorPositionsTextBoxFromTableCellParagraph(string name, bool behindDocument,
        double expectedLeft, double expectedTop, int expectedPage)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes", name + ".docx");
        using (var document = WordprocessingDocument.Open(source, false))
            Assert.Contains(document.MainDocumentPart!.Document.Descendants<TextBoxContent>(),
                box => box.Ancestors<TableCell>().Any());

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var pages = pdf.GetPages().ToArray();
        Assert.Equal(expectedPage, pages.Length);
        var page = pages[expectedPage - 1];
        var figure = Assert.Single(page.GetWords().Where(word => word.Text == "Figure:"));
        Assert.InRange(figure.BoundingBox.Left, expectedLeft - 0.3, expectedLeft + 0.3);
        Assert.InRange(page.Height - figure.BoundingBox.Top, expectedTop - 1.0, expectedTop + 1.0);
        var figureIndex = page.Text.IndexOf("Figure:Simple", StringComparison.Ordinal);
        var bodyIndex = page.Text.IndexOf("Textboxprobe:Simple", StringComparison.Ordinal);
        Assert.True(behindDocument ? figureIndex < bodyIndex : figureIndex > bodyIndex);
        if (name.StartsWith("TextBoxTableVerticalMerge", StringComparison.Ordinal) &&
            !name.Contains("Page", StringComparison.Ordinal))
            Assert.Equal(2, page.Paths.Count(path => path.IsStroked &&
                path.GetBoundingRectangle() is { } bounds && bounds.Width > 430 &&
                bounds.Height < 1));
        if (name.Contains("MergePage", StringComparison.Ordinal) ||
            name.Contains("DashedPage", StringComparison.Ordinal) ||
            name.Contains("DottedPage", StringComparison.Ordinal) ||
            name.StartsWith("TextBoxTableRepeatHeader", StringComparison.Ordinal) ||
            name.StartsWith("TextBoxTableNextPageRow", StringComparison.Ordinal))
            Assert.All(pages, fragment => Assert.Equal(2, fragment.Paths.Count(path =>
                path.IsStroked && path.GetBoundingRectangle() is { } bounds &&
                bounds.Width > 430 && bounds.Height < 1)));
        if (name.Contains("DoublePage", StringComparison.Ordinal))
            Assert.All(pages, fragment => Assert.True(fragment.Paths.Count(path =>
                path.IsStroked && path.GetBoundingRectangle() is { } bounds &&
                bounds.Width > 430 && bounds.Height < 1) >= 3));
        if (name.Contains("Border", StringComparison.Ordinal))
        {
            using var raw = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
            var commands = string.Concat(raw.Pages[0].Contents.Elements
                .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
                .Select(reference => System.Text.Encoding.Latin1.GetString(
                    ((PdfSharp.Pdf.Advanced.PdfContent)reference.Value!).Stream!.UnfilteredValue)));
            Assert.Contains("1 0 0 RG", commands);
            Assert.Contains("2 w", commands);
        }
    }

    [Theory]
    [InlineData("TextBoxTableDoublePage")]
    [InlineData("TextBoxTableDoublePageBehind")]
    public void FlowVisitorAlignsDoubleBorderAtTableFragmentEnd(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var pages = pdf.GetPages().ToArray();
        Assert.Equal(2, pages.Length);
        var horizontal = pages[0].Paths
            .Where(path => path.IsStroked && path.GetBoundingRectangle() is { } bounds &&
                bounds.Width > 430 && bounds.Height < 1)
            .Select(path => path.GetBoundingRectangle()!.Value).OrderBy(bounds => bounds.Bottom).ToArray();
        var sideEnd = pages[0].Paths
            .Where(path => path.IsStroked && path.GetBoundingRectangle() is { } bounds &&
                bounds.Width < 1 && bounds.Height > 10 && bounds.Left is > 72 and < 73)
            .Select(path => path.GetBoundingRectangle()!.Value.Bottom).Min();
        Assert.True(horizontal.Length >= 3);
        Assert.InRange(sideEnd - horizontal[1].Bottom, 0, 1);
    }

    [Theory]
    [InlineData("TextBoxTableSpanningMerge")]
    [InlineData("TextBoxTableSpanningMergeBehind")]
    public void FlowVisitorKeepsSpanningMergedCellTextBoxOnAnchorPage(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var pages = pdf.GetPages().ToArray();
        Assert.Equal(3, pages.Length);
        Assert.Single(pages[0].GetWords().Where(word => word.Text == "Figure:"));
        Assert.DoesNotContain("Figure:", pages[1].Text);
        Assert.DoesNotContain("Figure:", pages[2].Text);
        for (var index = 1; index <= 65; index++)
            Assert.Contains($"Spanningmergedcellline{index}.",
                string.Concat(pages.Select(page => page.Text)));
    }

    [Theory]
    [InlineData("OverpageMergedCell")]
    [InlineData("OverpageTableCell")]
    [InlineData("OverpageVerticalMerge")]
    [InlineData("OverpageTwoColumnCell")]
    [InlineData("OverpageThreeColumnCell")]
    [InlineData("OverpageMultiParagraphSideCell")]
    [InlineData("OverpageVariableParagraphs")]
    public void FlowVisitorPaginatesOverpageTableCellWithoutLosingLines(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "DocKnownGaps",
            name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.InRange(pdf.NumberOfPages, 6, name == "OverpageVariableParagraphs" ? 7 : 6);
        var allNumbers = new List<int>();
        double? firstLineTop = null;
        var continuationOffset = typeof(MigraDoc.Rendering.DocumentRenderer)
            .GetProperty("TableFragmentTopBorderWidth") is null ? 0.5 : 0;
        for (var pageNumber = 1; pageNumber <= 6; pageNumber++)
        {
            var page = pdf.GetPage(pageNumber);
            var currentFirstLineTop = page.GetWords()
                .Where(word => word.Text is "Cell" or "Merged" or "Variable")
                .Max(word => word.BoundingBox.Top);
            if (firstLineTop is null)
                firstLineTop = currentFirstLineTop;
            else
                Assert.InRange(currentFirstLineTop - firstLineTop.Value - continuationOffset,
                    -0.1, 0.1);
            var numbers = page.GetWords()
                .Where(word => word.BoundingBox.Bottom >= 0 && word.BoundingBox.Top <= page.Height)
                .Select(word => int.TryParse(word.Text, out var number) ? number : -1)
                .Where(number => number > 0).ToArray();
            allNumbers.AddRange(numbers);
        }
        Assert.Equal(Enumerable.Range(1, 150), allNumbers.Distinct().Order());
        if (pdf.NumberOfPages == 7)
            Assert.Empty(pdf.GetPage(7).GetWords());
        if (name is "OverpageTwoColumnCell" or "OverpageThreeColumnCell" or
            "OverpageMultiParagraphSideCell")
        {
            var pagesWithSideCell = pdf.GetPages()
                .Select(page => page.GetWords().Count(word => word.Text == "Side"))
                .ToArray();
            Assert.Equal(new[] { name == "OverpageMultiParagraphSideCell" ? 4 :
                name == "OverpageThreeColumnCell" ? 2 : 1, 0, 0, 0, 0, 0 },
                pagesWithSideCell);
        }
    }

    [Theory]
    [InlineData("OverpageInlineImageCell", 6, 3, 24)]
    [InlineData("OverpageTallInlineImageCell", 7, 4, 48)]
    public void FlowVisitorKeepsInlineImageInPaginatedTableCell(
        string name, int pageCount, int imagePageNumber, int imageSize)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "DocKnownGaps",
            name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.InRange(pdf.NumberOfPages, imageSize == 48 ? 6 : pageCount, pageCount);
        var allNumbers = new HashSet<int>();
        var imageCount = 0;
        for (var pageNumber = 1; pageNumber <= pdf.NumberOfPages; pageNumber++)
        {
            var page = pdf.GetPage(pageNumber);
            var numbers = page.GetWords()
                .Select(word => int.TryParse(word.Text, out var number) ? number : -1)
                .Where(number => number > 0)
                .ToArray();
            Assert.NotEmpty(numbers);
            foreach (var number in numbers)
                allNumbers.Add(number);
            foreach (var image in page.GetImages())
            {
                imageCount++;
                Assert.InRange(pageNumber, imageSize == 48 ? imagePageNumber - 1 : imagePageNumber,
                    imagePageNumber);
                Assert.InRange(image.Bounds.Width, imageSize - 0.1, imageSize + 0.1);
                Assert.InRange(image.Bounds.Height, imageSize - 0.1, imageSize + 0.1);
            }
        }
        Assert.Equal(1, imageCount);
        Assert.Equal(Enumerable.Range(1, 150), allNumbers.Order());
    }

    [Theory]
    [InlineData(12)]
    [InlineData(16)]
    [InlineData(24)]
    [InlineData(48)]
    public void FlowVisitorKeepsInlineImageInOrdinaryParagraph(int imageSize)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            $"inline-image-paragraph-probe__{imageSize}.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(1, pdf.NumberOfPages);
        var page = pdf.GetPage(1);
        var image = Assert.Single(page.GetImages());
        Assert.InRange(image.Bounds.Width, imageSize - 0.1, imageSize + 0.1);
        Assert.InRange(image.Bounds.Height, imageSize - 0.1, imageSize + 0.1);
        foreach (var number in Enumerable.Range(1, 6))
            Assert.Contains($"Line{number}", page.Text);
    }

    [Theory]
    [InlineData("OverpageParallelCells", 150)]
    [InlineData("OverpageUnequalCells", 100)]
    public void FlowVisitorPaginatesParallelOverpageTableCells(string name, int expectedRightTotal)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "DocKnownGaps",
            name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(6, pdf.NumberOfPages);
        var leftTotal = 0;
        var rightTotal = 0;
        foreach (var page in pdf.GetPages())
        {
            var words = page.GetWords().ToArray();
            var leftCount = words.Count(word => word.Text == "Left");
            var rightCount = words.Count(word => word.Text == "Right");
            Assert.InRange(rightCount, 0, leftCount);
            leftTotal += leftCount;
            rightTotal += rightCount;
        }
        Assert.Equal(150, leftTotal);
        Assert.Equal(expectedRightTotal, rightTotal);
    }

    [Theory]
    [InlineData("OverpageLongRightInlineImageCell", 0, 150)]
    [InlineData("OverpageLongRightNestedCell", 0, 150)]
    [InlineData("OverpageLongMiddleInlineImageCell", 150, 40)]
    public void FlowVisitorPaginatesLongNonfirstCellWithInlineImage(
        string name, int middleTotal, int rightTotal)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "DocKnownGaps",
            name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.InRange(pdf.NumberOfPages, 6, 7);
        var pages = pdf.GetPages().ToArray();
        Assert.Equal(80, pages.Sum(page => page.GetWords().Count(word => word.Text == "Left")));
        Assert.Equal(middleTotal, pages.Sum(page => page.GetWords().Count(word => word.Text == "Middle")));
        Assert.Equal(rightTotal, pages.Sum(page => page.GetWords().Count(word => word.Text == "Right")));
        var images = pages.SelectMany(page => page.GetImages()).ToArray();
        var image = Assert.Single(images);
        Assert.InRange(image.Bounds.Width, 47.9, 48.1);
        Assert.InRange(image.Bounds.Height, 47.9, 48.1);
        if (name == "OverpageLongRightNestedCell")
        {
            Assert.Equal(6, pages.Sum(page => page.GetWords().Count(word => word.Text == "Nested")));
        }
    }

    [Fact]
    public void FlowVisitorKeepsNestedTableInsideOverpageCell()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "DocKnownGaps",
            "OverpageNestedTableCell.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.InRange(pdf.NumberOfPages, 6, 7);
        var words = pdf.GetPages().SelectMany(page => page.GetWords()).ToArray();
        Assert.Equal(150, words.Count(word => word.Text == "Cell"));
        Assert.Equal(2, words.Count(word => word.Text == "Nested"));
    }

    [Fact]
    public void FlowVisitorRetainsOverpageNestedTwoColumnCellContent()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "DocKnownGaps",
            "OverpageNestedTwoColumnCell.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.InRange(pdf.NumberOfPages, 6, 7);
        var words = pdf.GetPages().SelectMany(page => page.GetWords()).ToArray();
        Assert.Equal(150, words.Count(word => word.Text == "Cell"));
        var nestedWords = words.Where(word => word.Text == "Nested").ToArray();
        Assert.Equal(4, nestedWords.Length);
        var left = nestedWords.Where(word => word.BoundingBox.Left < 200)
            .OrderByDescending(word => word.BoundingBox.Top).ToArray();
        var right = nestedWords.Where(word => word.BoundingBox.Left > 200)
            .OrderByDescending(word => word.BoundingBox.Top).ToArray();
        Assert.Equal(2, left.Length);
        Assert.Equal(2, right.Length);
        Assert.All(right, word => Assert.InRange(word.BoundingBox.Left, 281, 285));
        for (var index = 0; index < 2; index++)
            Assert.InRange(right[index].BoundingBox.Top - left[index].BoundingBox.Top, -1, 1);
    }

    [Theory]
    [InlineData("OverpageNestedTallTwoColumnCell")]
    [InlineData("OverpageNestedTallBorderedTwoColumnCell")]
    public void FlowVisitorFragmentsTallNestedTwoColumnCell(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "DocKnownGaps",
            name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.InRange(pdf.NumberOfPages, 7, 8);
        var words = pdf.GetPages().SelectMany(page => page.GetWords()).ToArray();
        Assert.Equal(150, words.Count(word => word.Text == "Cell"));
        Assert.Equal(35, words.Count(word => word.Text == "left"));
        Assert.Equal(2, words.Count(word => word.Text == "right"));
        Assert.Equal(37, words.Count(word => word.Text == "Nested"));
    }

    [Fact]
    public void FlowVisitorDoesNotDuplicateTallNestedCellBorders()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "DocKnownGaps",
            "OverpageNestedTallBorderedTwoColumnCell.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var nestedPages = pdf.GetPages().Where(page =>
            page.GetWords().Any(word => word.Text == "Nested")).ToArray();
        Assert.InRange(nestedPages.Length, 2, 3);
        foreach (var page in nestedPages)
        {
            var horizontalEdges = page.Paths.Where(path => path.IsStroked &&
                    path.GetBoundingRectangle() is { } bounds &&
                    bounds.Left is > 76 and < 80 && bounds.Width is > 195 and < 205 &&
                    bounds.Height < 0.5)
                .Select(path => path.GetBoundingRectangle()!.Value.Top)
                .OrderBy(position => position).ToArray();
            var edgeGroups = horizontalEdges.Aggregate(new List<double>(), (groups, position) =>
            {
                if (groups.Count == 0 || position - groups[^1] > 1)
                    groups.Add(position);
                return groups;
            });
            Assert.True(edgeGroups.Count is >= 1 and <= 2,
                $"Page {page.Number}: {string.Join(", ", horizontalEdges)}");
        }
    }

    [Fact]
    public void FlowVisitorUsesExplicitDrawingTextBoxStroke()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes",
            "TextBoxExplicitStroke.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var commands = string.Concat(pdf.Pages[0].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Select(reference => System.Text.Encoding.Latin1.GetString(
                ((PdfSharp.Pdf.Advanced.PdfContent)reference.Value!).Stream!.UnfilteredValue)));
        Assert.Contains("1 0 0 RG", commands);
        Assert.Contains("2 w", commands);
    }

    [Theory]
    [InlineData("TextBoxVerticalText", "vert", "0 -1 1 0", 16.4, 16.8)]
    [InlineData("TextBoxVerticalUp", "vert270", "0 1 -1 0", 16.4, 16.8)]
    [InlineData("TextBoxVerticalTextSingleLine", "vert", "0 -1 1 0", 14.0, 14.4)]
    public void FlowVisitorOrientsVerticalDrawingTextBoxContent(string name, string orientation,
        string transform, double minimumPitch, double maximumPitch)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes",
            name + ".docx");
        using (var document = WordprocessingDocument.Open(source, false))
            Assert.Contains(document.MainDocumentPart!.Document.Descendants(),
                element => element.LocalName == "bodyPr" && element.GetAttribute("vert", "").Value == orientation);

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var commands = string.Concat(pdf.Pages[0].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Select(reference => System.Text.Encoding.Latin1.GetString(
                ((PdfSharp.Pdf.Advanced.PdfContent)reference.Value!).Stream!.UnfilteredValue)));
        Assert.Contains(transform, commands);
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = parsed.GetPage(1);
        var frameLetters = page.Letters.Where(letter =>
            page.Height - letter.GlyphRectangle.Top is > 115 and < 205 &&
            letter.GlyphRectangle.Left is > 100 and < 300).ToArray();
        var figure = Assert.Single(frameLetters.Where(letter => letter.Value == "F"));
        var following = frameLetters.Where(letter => letter.Value == "V")
            .MinBy(letter => Math.Abs(letter.GlyphRectangle.Left - figure.GlyphRectangle.Left));
        Assert.NotNull(following);
        Assert.InRange(Math.Abs(following.GlyphRectangle.Left - figure.GlyphRectangle.Left),
            minimumPitch, maximumPitch);
    }

    [Theory]
    [InlineData("TextBoxOverflow.docx")]
    [InlineData("TextBoxRotatedOverflow.docx")]
    public void FlowVisitorClipsOverflowingTextBoxContent(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfTextBoxes", name);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Equal(1, pdf.PageCount);
        var commands = string.Concat(pdf.Pages[0].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Select(reference => System.Text.Encoding.Latin1.GetString(
                ((PdfSharp.Pdf.Advanced.PdfContent)reference.Value!).Stream!.UnfilteredValue)));
        Assert.Contains("W* n", commands);
        using var extracted = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Contains("Visiblefirstline", extracted.GetPage(1).Text);
    }

    [Theory]
    [InlineData("RevisionMovedText.docx", 1)]
    [InlineData("RevisionMovedTextSplitRuns.docx", 2)]
    public void FlowVisitorKeepsMovedInTextInCleanView(string name, int runCount)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfFeatures",
            name);
        using (var document = WordprocessingDocument.Open(source, false))
        {
            Assert.Single(document.MainDocumentPart!.Document.Descendants<MoveFromRun>());
            Assert.Single(document.MainDocumentPart.Document.Descendants<MoveToRun>());
            Assert.Equal(runCount, document.MainDocumentPart.Document.Descendants<MoveToRun>()
                .Single().Elements<Run>().Count());
        }

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Single(pdf.GetPages());
        Assert.Contains("new", pdf.GetPage(1).Text);
        Assert.DoesNotContain("old", pdf.GetPage(1).Text);
    }

    [Theory]
    [InlineData("RevisionMovedText.docx")]
    [InlineData("RevisionMovedTextSplitRuns.docx")]
    public void FlowVisitorRendersMovedTextInMarkupView(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfFeatures",
            name);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(
            includeRevisionMarkup: true, revisionMarkupColor: XColor.FromArgb(44, 98, 52)));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Single(pdf.GetPages());
        Assert.Contains("new", pdf.GetPage(1).Text);
        Assert.Contains("old", pdf.GetPage(1).Text);
        var decorations = pdf.GetPage(1).Paths.Select(path => path.GetBoundingRectangle())
            .Where(bounds => bounds != null && bounds.Value.Left > 90 && bounds.Value.Right < 135)
            .ToArray();
        Assert.Equal(4, decorations.Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FlowVisitorKeepsMovedTextOnCorrectPagesAcrossPageBreak(bool markup)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfFeatures",
            "RevisionMovedTextAcrossPages.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(
            includeRevisionMarkup: markup, revisionMarkupColor: XColor.FromArgb(44, 98, 52)));

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("new", pdf.GetPage(1).Text);
        Assert.DoesNotContain("old", pdf.GetPage(1).Text);
        Assert.Equal(markup, pdf.GetPage(2).Text.Contains("old", StringComparison.Ordinal));
        Assert.Contains("Afterthemove.", pdf.GetPage(2).Text);
        if (markup)
        {
            for (var pageNumber = 1; pageNumber <= 2; pageNumber++)
            {
                var decorations = pdf.GetPage(pageNumber).Paths
                    .Select(path => path.GetBoundingRectangle())
                    .Where(bounds => bounds != null && bounds.Value.Left is > 70 and < 120 &&
                        bounds.Value.Right < 130).ToArray();
                Assert.Equal(2, decorations.Length);
            }
        }
    }

    [Theory]
    [InlineData("RevisionMovedTextAcrossCells.docx", false)]
    [InlineData("RevisionMovedTextAcrossCells.docx", true)]
    [InlineData("RevisionMovedTextCenteredCell.docx", false)]
    [InlineData("RevisionMovedTextCenteredCell.docx", true)]
    [InlineData("RevisionMovedTextNestedTable.docx", false)]
    [InlineData("RevisionMovedTextNestedTable.docx", true)]
    [InlineData("RevisionTableWidth-short.docx", false)]
    [InlineData("RevisionTableWidth-short.docx", true)]
    [InlineData("RevisionTableWidth-medium.docx", false)]
    [InlineData("RevisionTableWidth-medium.docx", true)]
    [InlineData("RevisionTableWidth-long.docx", false)]
    [InlineData("RevisionTableWidth-long.docx", true)]
    public void FlowVisitorKeepsMovedTextInSeparateTableCells(string name, bool markup)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfFeatures",
            name);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(
            includeRevisionMarkup: markup, revisionMarkupColor: XColor.FromArgb(44, 98, 52)));

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains("new", page.Text);
        var movedText = name switch
        {
            "RevisionTableWidth-short.docx" => "x",
            "RevisionTableWidth-medium.docx" => "former",
            "RevisionTableWidth-long.docx" => "previously",
            _ => "old",
        };
        Assert.Equal(markup, page.Text.Contains(movedText, StringComparison.Ordinal));
        Assert.Contains("firstcell.", page.Text);
        Assert.Contains("secondcell.", page.Text);
        if (!markup && (name == "RevisionMovedTextAcrossCells.docx" ||
                        name.StartsWith("RevisionTableWidth-", StringComparison.Ordinal)))
        {
            var secondRemains = page.GetWords().Where(word => word.Text == "remains")
                .OrderBy(word => word.BoundingBox.Left).Last();
            Assert.InRange(secondRemains.BoundingBox.Left, 314.3, 315.2);
        }
        if (markup)
        {
            var decorations = page.Paths.Select(path => path.GetBoundingRectangle())
                .Where(bounds => bounds != null).ToArray();
            var bars = decorations.Count(bounds => bounds!.Value.Left is > 35 and < 37);
            Assert.Equal(name == "RevisionMovedTextCenteredCell.docx" ? 2 : 1, bars);
            Assert.Equal(4, decorations.Length - bars);
        }
        else Assert.Empty(page.Paths);
    }

    [Theory]
    [InlineData("RevisionDeletedTextAcrossCells.docx")]
    [InlineData("RevisionDeletedHeaderTable.docx")]
    [InlineData("RevisionDeletedFooterTable.docx")]
    [InlineData("RevisionDeletedFooterTrailingTable.docx")]
    public void FlowVisitorPreservesSpaceAndRevisionBarForDeletedTableText(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfFeatures",
            name);
        var cleanBytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        var markupBytes = DxpExport.ExportToBytes(source,
            new DxpFlowPdfVisitor(includeRevisionMarkup: true));

        using var clean = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(cleanBytes));
        using var markup = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(markupBytes));
        var cleanPage = Assert.Single(clean.GetPages());
        var markupPage = Assert.Single(markup.GetPages());
        Assert.DoesNotContain("old", cleanPage.Text);
        Assert.Contains("old", markupPage.Text);
        var secondRemains = cleanPage.GetWords().Where(word => word.Text == "remains")
            .OrderBy(word => word.BoundingBox.Left).Last();
        Assert.InRange(secondRemains.BoundingBox.Left, 314.3, 315.2);
        Assert.Empty(cleanPage.Paths);
        Assert.Contains(markupPage.Paths, path => path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 35 and < 37 && bounds.Height > 20);
    }

    [Fact]
    public void FlowVisitorRetainsFooterTableTrailingParagraphPosition()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var fixtures = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfFeatures");
        double TableTextTop(string name)
        {
            var bytes = DxpExport.ExportToBytes(Path.Combine(fixtures, name), new DxpFlowPdfVisitor());
            using var document = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
            var page = Assert.Single(document.GetPages());
            var word = page.GetWords().First(item => item.Text == "The");
            return page.Height - word.BoundingBox.Top;
        }

        var tableOnly = TableTextTop("RevisionDeletedFooterTable.docx");
        var withTrailingParagraph = TableTextTop("RevisionDeletedFooterTrailingTable.docx");
        var size10 = TableTextTop("RevisionDeletedFooterTableSize10.docx");
        var size14 = TableTextTop("RevisionDeletedFooterTableSize14.docx");
        Assert.InRange(tableOnly - withTrailingParagraph, 24, 25.5);
        Assert.InRange(size10 - tableOnly, 1.8, 3.4);
        Assert.InRange(size14 - tableOnly, -3.4, -1.8);
    }

    [Fact]
    public void FlowVisitorRendersInlineRevisionMarkupOnlyWhenRequested()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfFeatures",
            "RevisionInline.docx");
        using (var document = WordprocessingDocument.Open(source, false))
        {
            Assert.Single(document.MainDocumentPart!.Document.Descendants<InsertedRun>());
            Assert.Single(document.MainDocumentPart.Document.Descendants<DeletedRun>());
        }

        var cleanBytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        var markupBytes = DxpExport.ExportToBytes(source,
            new DxpFlowPdfVisitor(includeRevisionMarkup: true));
        using var clean = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(cleanBytes));
        using var markup = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(markupBytes));
        Assert.DoesNotContain("old", clean.GetPage(1).Text);
        Assert.Contains("new", clean.GetPage(1).Text);
        Assert.Contains("old", markup.GetPage(1).Text);
        Assert.Contains("new", markup.GetPage(1).Text);
        var decorations = markup.GetPage(1).Paths.Select(path => path.GetBoundingRectangle())
            .Where(bounds => bounds != null && bounds.Value.Left > 80)
            .Select(bounds => bounds!.Value)
            .OrderBy(bounds => bounds.Left).ToArray();
        Assert.Equal(2, decorations.Length);
        Assert.Contains(markup.GetPage(1).Paths, path =>
            path.GetBoundingRectangle() is { } bounds &&
            path.IsFilled && bounds.Left is > 35 and < 37 &&
            bounds.Width is > 0.7 and < 0.75 && bounds.Height is > 23.5 and < 26);
        Assert.InRange(decorations[1].Left - decorations[0].Right, -0.5, 0.5);
        var deletedLetter = markup.GetPage(1).Letters.First(letter =>
            Math.Abs(letter.StartBaseLine.X - decorations[1].Left) < 0.5);
        var strikeFromBaseline = markup.GetPage(1).Height - deletedLetter.StartBaseLine.Y -
            (markup.GetPage(1).Height - decorations[1].Top);
        Assert.InRange(strikeFromBaseline, deletedLetter.PointSize / 3 - 0.5,
            deletedLetter.PointSize / 3 + 0.5);
        using var painted = PdfReader.Open(new MemoryStream(markupBytes), PdfDocumentOpenMode.Import);
        var decorationStream = (PdfSharp.Pdf.Advanced.PdfContent)
            ((PdfSharp.Pdf.Advanced.PdfReference)painted.Pages[0].Contents.Elements[^1]).Value!;
        var decorationCommands = System.Text.Encoding.Latin1.GetString(
            decorationStream.Stream!.UnfilteredValue);
        Assert.Contains(".82 .204 .22 RG", decorationCommands);
        var textStream = (PdfSharp.Pdf.Advanced.PdfContent)
            ((PdfSharp.Pdf.Advanced.PdfReference)painted.Pages[0].Contents.Elements[0]).Value!;
        Assert.Contains(".82 .204 .22 rg", System.Text.Encoding.Latin1.GetString(
            textStream.Stream!.UnfilteredValue));
    }

    [Fact]
    public void FlowVisitorIncludesDeletedParagraphTextOnlyInMarkupView()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Tracked.docx");
        var cleanBytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        var markupBytes = DxpExport.ExportToBytes(source,
            new DxpFlowPdfVisitor(includeRevisionMarkup: true));
        using var clean = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(cleanBytes));
        using var markup = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(markupBytes));
        Assert.DoesNotContain("Deleted4", clean.GetPage(1).Text);
        Assert.DoesNotContain("Delete8", clean.GetPage(1).Text);
        Assert.Contains("Deleted4", markup.GetPage(1).Text);
        Assert.Contains("Delete8", markup.GetPage(1).Text);
        var markupText = System.Text.RegularExpressions.Regex.Replace(markup.GetPage(1).Text,
            @"\s+", string.Empty);
        Assert.Contains("4:Deleted4", markupText);
        Assert.Contains("6:Delete8", markupText);
    }

    [Fact]
    public void FlowVisitorColorsRevisedListMarkersInMarkupView()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfFeatures",
            "RevisionNoFormattingBalloon.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(
            includeRevisionMarkup: true, revisionMarkupColor: XColor.FromArgb(0, 120, 212)));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var markers = pdf.GetPage(1).GetWords()
            .Where(word => word.Text is "4:" or "5:" or "6:")
            .ToArray();
        Assert.Equal(5, markers.Length);
        foreach (var marker in markers)
        {
            var color = marker.Letters[0].Color.ToRGBValues();
            if (marker.Text == "5:")
                Assert.True(color.r < 0.05 && color.g < 0.05 && color.b < 0.05);
            else
                Assert.True(color.r < 0.05 && color.g > 0.45 && color.b > 0.8);
        }
        var revisionLines = pdf.GetPage(1).Paths
            .Select(path => path.GetBoundingRectangle())
            .Where(bounds => bounds is { Width: > 60 and < 75, Height: < 0.1 } &&
                pdf.GetPage(1).Height - bounds.Value.Top is > 175 and < 255)
            .ToArray();
        Assert.Contains(revisionLines, bounds => bounds!.Value.Width is > 72 and < 74);
    }

    [Fact]
    public void FlowVisitorKeepsRevisedListDecorationAcrossSplitRuns()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfFeatures",
            "RevisionNoFormattingBalloon.docx");
        var splitSource = Path.Combine(Path.GetTempPath(), $"revision-split-{Guid.NewGuid():N}.docx");
        File.Copy(source, splitSource);
        try
        {
            using (var document = WordprocessingDocument.Open(splitSource, true))
            {
                var body = document.MainDocumentPart!.Document;
                var deleted = body.Descendants<DeletedRun>().First(revision =>
                    revision.Descendants<DeletedText>().Any(text => text.Text.Contains("Deleted 4")));
                var deletedRun = deleted.GetFirstChild<Run>()!;
                var deletedText = deletedRun.GetFirstChild<DeletedText>()!;
                var deletedContent = deletedText.Text;
                deletedText.Text = deletedContent[..5];
                var deletedRemainder = (Run)deletedRun.CloneNode(true);
                deletedRemainder.GetFirstChild<DeletedText>()!.Text = deletedContent[5..];
                deletedRun.InsertAfterSelf(deletedRemainder);
                var inserted = body.Descendants<InsertedRun>().First(revision =>
                    revision.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                        .Any(text => text.Text.Contains("Added 5")));
                var insertedRun = inserted.GetFirstChild<Run>()!;
                var insertedText = insertedRun.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Text>()!;
                var insertedContent = insertedText.Text;
                insertedText.Text = insertedContent[..5];
                var insertedRemainder = (Run)insertedRun.CloneNode(true);
                insertedRemainder.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Text>()!.Text =
                    insertedContent[5..];
                insertedRun.InsertAfterSelf(insertedRemainder);
                body.Save();
            }
            var bytes = DxpExport.ExportToBytes(splitSource, new DxpFlowPdfVisitor(
                includeRevisionMarkup: true, revisionMarkupColor: XColor.FromArgb(0, 120, 212)));
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
            var page = Assert.Single(pdf.GetPages());
            var normalizedText = System.Text.RegularExpressions.Regex.Replace(page.Text, @"\s+", "");
            Assert.Contains("Deleted4", normalizedText);
            Assert.Contains("Added5", normalizedText);
            var revisionLines = page.Paths.Select(path => path.GetBoundingRectangle())
                .Where(bounds => bounds is { Height: < 0.1, Right: > 150 and < 165 } &&
                    page.Height - bounds.Value.Top is > 175 and < 210)
                .ToArray();
            Assert.NotEmpty(revisionLines);
        }
        finally
        {
            File.Delete(splitSource);
        }
    }

    [Fact]
    public void RevisionBalloonControlPreservesTextRevisions()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        using var original = WordprocessingDocument.Open(Path.Combine(root, "samples", "Tracked.docx"), false);
        using var control = WordprocessingDocument.Open(Path.Combine(root, "DocxportNet.Tests", "Fixtures",
            "PdfFeatures", "RevisionNoFormattingBalloon.docx"), false);
        using var a4Control = WordprocessingDocument.Open(Path.Combine(root, "DocxportNet.Tests", "Fixtures",
            "PdfFeatures", "RevisionFormattingBalloonA4.docx"), false);
        var originalBody = original.MainDocumentPart!.Document;
        var controlBody = control.MainDocumentPart!.Document;
        var a4Body = a4Control.MainDocumentPart!.Document;
        Assert.Single(originalBody.Descendants<RunPropertiesChange>());
        Assert.Empty(controlBody.Descendants<RunPropertiesChange>());
        Assert.Single(a4Body.Descendants<RunPropertiesChange>());
        Assert.Equal(originalBody.Descendants<InsertedRun>().Count(),
            controlBody.Descendants<InsertedRun>().Count());
        Assert.Equal(originalBody.Descendants<DeletedRun>().Count(),
            controlBody.Descendants<DeletedRun>().Count());
        Assert.Equal(originalBody.InnerText, controlBody.InnerText);
        Assert.Equal(originalBody.InnerText, a4Body.InnerText);
        Assert.Equal(11906U, a4Body.Descendants<PageSize>().Single().Width!.Value);
        Assert.Equal(16838U, a4Body.Descendants<PageSize>().Single().Height!.Value);
    }

    [Fact]
    public void FlowVisitorStrikesEachDeletedLineInMarkupView()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfFeatures",
            "RevisionMultilineDeletion.docx");
        var bytes = DxpExport.ExportToBytes(source,
            new DxpFlowPdfVisitor(includeRevisionMarkup: true));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var strikes = page.Paths.Select(path => path.GetBoundingRectangle())
            .Where(bounds => bounds is { Width: > 100, Height: < 0.1 })
            .Select(bounds => bounds!.Value).OrderByDescending(bounds => bounds.Top).ToArray();
        Assert.Equal(3, strikes.Length);
        Assert.InRange(strikes[1].Top - strikes[0].Top, -18, -15.5);
        Assert.InRange(strikes[2].Top - strikes[1].Top, -18, -15.5);
    }

    [Fact]
    public void FlowVisitorDrawsPageBorderWithoutMarkingBlankOddPage()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordMixedGeometryPageBorderStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Equal(4, pdf.PageCount);
        Assert.True(pdf.Pages[0].Contents.Elements.Count > pdf.Pages[1].Contents.Elements.Count);
    }

    [Fact]
    public void FlowVisitorStopsPageBorderSidesAtHorizontalBorders()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordMixedGeometryPageBorderStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var sides = page.Paths.Where(path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds && bounds.Height > 500 && bounds.Width < 1).ToArray();
        Assert.Equal(2, sides.Length);
        Assert.All(sides, side =>
        {
            var bounds = side.GetBoundingRectangle()!.Value;
            Assert.InRange(page.Height - bounds.Top, 9.3, 9.6);
            Assert.InRange(page.Height - bounds.Bottom, 599.7, 600.0);
        });
    }

    [Fact]
    public void FlowVisitorPaginatesRepeatingHeaderTableByRowHeight()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", "WordVisibleRepeatingHeaderTable.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Equal(2, pdf.PageCount);
        using var extracted = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Contains(extracted.GetPage(1).GetWords(), word => word.Text == "15");
        Assert.DoesNotContain(extracted.GetPage(1).GetWords(), word => word.Text == "16");
        Assert.Contains(extracted.GetPage(2).GetWords(), word => word.Text == "16");
    }

    [Fact]
    public void FlowVisitorEmbedsWindowsBmpDrawing()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", "WordBmpAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.NotNull(pdf.Pages[0].Resources.Elements.GetDictionary("/XObject"));
    }

    [Fact]
    public void Rle4DeltaSkippedPixelsMatchWordBlackBackground()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordBmpRle4DeltaAllStories.docx");
        using var document = WordprocessingDocument.Open(source, false);
        using var image = Assert.Single(document.MainDocumentPart!.ImageParts).GetStream();
        using var copy = new MemoryStream();
        image.CopyTo(copy);

        var bitmap = DxpBmp24Converter.Normalize(copy.ToArray());
        Assert.Equal(16, BinaryPrimitives.ReadInt32LittleEndian(bitmap.AsSpan(18)));
        Assert.Equal(16, Math.Abs(BinaryPrimitives.ReadInt32LittleEndian(bitmap.AsSpan(22))));
        var black = 0;
        var cyan = 0;
        for (var offset = 54; offset < bitmap.Length; offset += 3)
        {
            var pixel = bitmap.AsSpan(offset, 3);
            if (pixel.SequenceEqual(new byte[] { 0, 0, 0 })) black++;
            else if (pixel.SequenceEqual(new byte[] { 255, 200, 0 })) cyan++;
            else Assert.Fail($"Unexpected RLE4 delta pixel at byte {offset}.");
        }
        Assert.Equal(192, black);
        Assert.Equal(64, cyan);
    }

    [Fact]
    public void FlowVisitorEmbedsAllBmpVariants()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var fixtures = Directory.GetFiles(Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc"),
            "WordBmp*AllStories.docx");

        Assert.Equal(14, fixtures.Length);
        foreach (var source in fixtures)
        {
            var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
            using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
            Assert.NotNull(pdf.Pages[0].Resources.Elements.GetDictionary("/XObject"));
        }
    }

    [Theory]
    [InlineData("WordTiffAllStories.docx")]
    [InlineData("WordFloatingTiffAllStories.docx")]
    public void FlowVisitorEmbedsTiffDrawings(string fixtureName)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixtureName);

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.NotNull(pdf.Pages[0].Resources.Elements.GetDictionary("/XObject"));
    }

    [Fact]
    public void FlowVisitorPlacesParagraphAnchoredTiffBesideBodyText()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordFloatingTiffAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.GetImages(), image =>
            image.Bounds.Left is > 71 and < 73 &&
            page.Height - image.Bounds.Top is > 71 and < 73);
        Assert.Contains(page.Letters, letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left is > 128 and < 131 &&
            page.Height - letter.GlyphRectangle.Top is > 71 and < 76);
    }

    [Theory]
    [InlineData("WordFloatingTiffAllStories.docx")]
    [InlineData("WordFloatingBmpAllStories.docx")]
    public void FlowVisitorPlacesParagraphAnchoredHeaderAndFooterImagesBesideText(string fixtureName)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixtureName);

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var images = page.GetImages().OrderByDescending(image => image.Bounds.Bottom).ToArray();
        Assert.Equal(3, images.Length);
        Assert.InRange(page.Height - images[0].Bounds.Top, 35.5, 37.5);
        Assert.InRange(page.Height - images[2].Bounds.Top, 720.5, 723.5);
        var header = page.GetWords().First(word => word.Text == "Header");
        Assert.InRange(header.BoundingBox.Left, 128.5, 129.5);
        var footer = page.GetWords().First(word => word.Text == "Footer");
        Assert.InRange(footer.BoundingBox.Left, 128.5, 129.5);
        Assert.InRange(page.Height - footer.BoundingBox.Top, 742, 746);
    }

    [Theory]
    [InlineData("WordRightAngleImageStories.docx")]
    [InlineData("WordFlippedImageStories.docx")]
    [InlineData("WordArbitraryRotationStories.docx")]
    public void FlowVisitorEmbedsSixteenBitPngDrawings(string fixtureName)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixtureName);

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.NotNull(pdf.Pages[0].Resources.Elements.GetDictionary("/XObject"));
    }

    [Fact]
    public void FlowVisitorAppliesInlinePngFlipsAcrossStories()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordFlippedImageStories.docx");
        using var document = WordprocessingDocument.Open(source, false);
        using var imageStream = Assert.Single(document.MainDocumentPart!.ImageParts).GetStream();
        using var imageBytes = new MemoryStream();
        imageStream.CopyTo(imageBytes);
        var bitmap = DxpPng16Converter.ToBmpOrOriginal(imageBytes.ToArray());
        var expected = new[] { (false, true), (true, false), (true, true) }
            .Select(flips => DxpRasterImageTransform.Transform(bitmap, 60, 60, 0,
                flips.Item1, flips.Item2))
            .Select(transformed =>
            {
                Assert.NotNull(transformed);
                var image = StbImageSharp.ImageResult.FromMemory(transformed.Value.Image,
                    StbImageSharp.ColorComponents.RedGreenBlueAlpha);
                var rgb = new byte[image.Width * image.Height * 3];
                for (var pixel = 0; pixel < image.Width * image.Height; pixel++)
                    Buffer.BlockCopy(image.Data, pixel * 4, rgb, pixel * 3, 3);
                return rgb;
            }).ToArray();
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var objects = pdf.Pages[0].Resources.Elements.GetDictionary("/XObject");
        Assert.NotNull(objects);
        var images = objects.Elements.Values
            .Select(value => (value as PdfSharp.Pdf.Advanced.PdfReference)?.Value as PdfSharp.Pdf.PdfDictionary ??
                value as PdfSharp.Pdf.PdfDictionary)
            .Where(image => image != null && image.Elements.GetName("/Subtype") == "/Image" &&
                image.Elements.GetInteger("/Width") == 80 && image.Elements.GetInteger("/Height") == 80)
            .Select(image => image!.Stream!.UnfilteredValue).ToArray();
        Assert.Equal(3, images.Length);
        foreach (var pixels in expected)
            Assert.Contains(images, image => image.SequenceEqual(pixels));
    }

    [Fact]
    public void FlowVisitorCentersArbitrarilyRotatedHeaderImageOverInlineFrame()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordArbitraryRotationStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.GetImages(), image =>
            image.Bounds.Left is > 59 and < 61 &&
            page.Height - image.Bounds.Top is > 23 and < 25 &&
            image.Bounds.Width is > 84 and < 86);
        Assert.Contains(page.Letters, letter => letter.Value == "H" &&
            letter.GlyphRectangle.Left is > 131 and < 134);
    }

    [Fact]
    public void RotatedRasterKeepsSourceEdgeColorUnderTransparentPixels()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordArbitraryRotationStories.docx");
        using var document = WordprocessingDocument.Open(source, false);
        var imagePart = document.MainDocumentPart!.ImageParts.Single();
        using var stream = imagePart.GetStream();
        using var bytes = new MemoryStream();
        stream.CopyTo(bytes);

        var rotated = DxpRasterImageTransform.Transform(bytes.ToArray(), 80, 80, 45, false, false);

        Assert.NotNull(rotated);
        var image = StbImageSharp.ImageResult.FromMemory(rotated.Value.Image,
            StbImageSharp.ColorComponents.RedGreenBlueAlpha);
        Assert.InRange(image.Width, 226, 229);
        Assert.Equal(image.Width, image.Height);
        var transparentPixels = Enumerable.Range(0, image.Width * image.Height)
            .Where(index => image.Data[index * 4 + 3] == 0).ToArray();
        Assert.NotEmpty(transparentPixels);
        Assert.All(transparentPixels, index =>
            Assert.Equal(new byte[] { 255, 255, 255 }, image.Data[(index * 4)..(index * 4 + 3)]));
    }

    [Theory]
    [InlineData("WordCroppedFloatingPngBody", 0.13331, 0.09999, 0, 0, 104, 108)]
    [InlineData("WordFourSidedCroppedFloatingPngBody", 0.13331, 0.09999, 0.06666, 0.08888, 96, 97)]
    public void FloatingRasterAppliesSourceRectangleBeforeEmbedding(string fixtureName,
        double left, double top, double right, double bottom, int expectedWidth, int expectedHeight)
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixtureName + ".docx");
        using var document = WordprocessingDocument.Open(source, false);
        using var stream = document.MainDocumentPart!.ImageParts.Single().GetStream();
        using var bytes = new MemoryStream();
        stream.CopyTo(bytes);

        var transformed = DxpRasterImageTransform.Transform(bytes.ToArray(), expectedWidth, 81,
            0, false, false, new DxpImageCrop(left, top, right, bottom));

        Assert.NotNull(transformed);
        var image = StbImageSharp.ImageResult.FromMemory(transformed.Value.Image,
            StbImageSharp.ColorComponents.RedGreenBlueAlpha);
        Assert.Equal(expectedWidth, image.Width);
        Assert.Equal(expectedHeight, image.Height);
        Assert.DoesNotContain(Enumerable.Range(0, image.Width * image.Height), index =>
            image.Data[index * 4] < 200 && image.Data[index * 4 + 1] < 200 &&
            image.Data[index * 4 + 2] < 200);
    }

    [Theory]
    [InlineData("WordCroppedFloatingPngBody", 104, 81)]
    [InlineData("WordFourSidedCroppedFloatingPngBody", 96, 73)]
    public void FloatingRasterUsesSimplePagePosition(string fixtureName, int width, int height)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixtureName + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var image = Assert.Single(pdf.GetPage(1).GetImages());
        Assert.InRange(image.Bounds.Left, 187.8, 188.1);
        Assert.InRange(pdf.GetPage(1).Height - image.Bounds.Top, 180.8, 181.1);
        Assert.InRange(image.Bounds.Width, width - 0.1, width + 0.1);
        Assert.InRange(image.Bounds.Height, height - 0.1, height + 0.1);
    }

    [Theory]
    [InlineData("WordTiffAllStories", 72.84)]
    [InlineData("WordImageStories", 99.37)]
    public void InlineImageFirstParagraphKeepsHeaderAndBodyPositions(string fixtureName, double bodyTop)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixtureName + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var images = page.GetImages().OrderBy(image => image.Bounds.Left).ToArray();
        Assert.True(images.Any(image => Math.Abs(page.Height - image.Bounds.Top - 36.84) < 0.2),
            string.Join(", ", images.Select(image => (page.Height - image.Bounds.Top).ToString("F2"))));
        Assert.True(images.Any(image => Math.Abs(page.Height - image.Bounds.Top - bodyTop) < 0.2),
            string.Join(", ", images.Select(image => (page.Height - image.Bounds.Top).ToString("F2"))));
    }

    [Theory]
    [InlineData("PdfFeatures", "AltChunkInlineImage")]
    [InlineData("PdfProbeCorpus", "altchunk-probe__html-inline-image")]
    public void HtmlAltChunkDoesNotAddSpaceAfterInlineImage(string fixtureGroup, string fixtureName)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", fixtureGroup,
            fixtureName + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var image = Assert.Single(page.GetImages());
        var following = page.Letters.Where(letter => letter.Value == "i" &&
            letter.GlyphRectangle.Left > image.Bounds.Right - 1)
            .MinBy(letter => letter.GlyphRectangle.Left)!;
        Assert.InRange(following.GlyphRectangle.Left - image.Bounds.Right, -0.5, 0.8);
    }

    [Fact]
    public void StandaloneDrawingImageSpacingMovesImageAndFollowingTextTogether()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var fixtures = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus");
        var previousSetting = Environment.GetEnvironmentVariable("DOCXPORT_DIAG_STANDALONE_IMAGE_DESCENT");

        byte[] Export(string fixtureName, string? setting, byte[]? sourceBytes = null)
        {
            try
            {
                Environment.SetEnvironmentVariable("DOCXPORT_DIAG_STANDALONE_IMAGE_DESCENT", setting);
                return sourceBytes == null
                    ? DxpExport.ExportToBytes(Path.Combine(fixtures, fixtureName + ".docx"),
                        new DxpFlowPdfVisitor())
                    : DxpExport.ExportToBytes(sourceBytes, new DxpFlowPdfVisitor());
            }
            finally
            {
                Environment.SetEnvironmentVariable("DOCXPORT_DIAG_STANDALONE_IMAGE_DESCENT", previousSetting);
            }
        }

        const string drawingFixture = "vml-gif-probe__drawing-gif-line-single";
        using var baseline = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(Export(drawingFixture, null)));
        using var adjusted = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(Export(drawingFixture, "3")));
        Assert.Equal(baseline.NumberOfPages, adjusted.NumberOfPages);

        var baselinePage = baseline.GetPage(1);
        var adjustedPage = adjusted.GetPage(1);
        double ImageTop(UglyToad.PdfPig.Content.Page page) =>
            page.Height - Assert.Single(page.GetImages()).Bounds.Top;
        double Baseline(UglyToad.PdfPig.Content.Page page, string glyph) =>
            page.Height - page.Letters.First(letter => letter.Value == glyph).StartBaseLine.Y;

        Assert.InRange(ImageTop(baselinePage) - ImageTop(adjustedPage), 0.34, 0.38);
        Assert.InRange(Baseline(baselinePage, "A") - Baseline(adjustedPage, "A"), 0.34, 0.38);
        Assert.InRange(Math.Abs(Baseline(baselinePage, "B") - Baseline(adjustedPage, "B")), 0, 0.01);

        const string vmlFixture = "vml-gif-probe__vml-gif";
        using var vmlBaseline = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(Export(vmlFixture, null)));
        using var vmlAdjusted = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(Export(vmlFixture, "3")));
        Assert.InRange(Math.Abs(ImageTop(vmlBaseline.GetPage(1)) - ImageTop(vmlAdjusted.GetPage(1))),
            0, 0.01);

        using var pngSource = new MemoryStream();
        pngSource.Write(File.ReadAllBytes(Path.Combine(fixtures, drawingFixture + ".docx")));
        using (var document = WordprocessingDocument.Open(pngSource, true))
        {
            var main = document.MainDocumentPart!;
            var oldImage = Assert.Single(main.ImageParts);
            var replacement = main.AddImagePart(ImagePartType.Png, "rIdPng");
            const string png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC";
            using (var data = new MemoryStream(Convert.FromBase64String(png)))
                replacement.FeedData(data);
            Assert.Single(main.Document.Descendants<DocumentFormat.OpenXml.Drawing.Blip>()).Embed =
                main.GetIdOfPart(replacement);
            main.DeletePart(oldImage);
            main.Document.Save();
        }
        var pngBytes = pngSource.ToArray();
        using var pngBaseline = UglyToad.PdfPig.PdfDocument.Open(
            new MemoryStream(Export(drawingFixture, null, pngBytes)));
        using var pngAdjusted = UglyToad.PdfPig.PdfDocument.Open(
            new MemoryStream(Export(drawingFixture, "3", pngBytes)));
        var pngBaselinePage = pngBaseline.GetPage(1);
        var pngAdjustedPage = pngAdjusted.GetPage(1);
        Assert.InRange(ImageTop(pngBaselinePage) - ImageTop(pngAdjustedPage),
            0.34, 0.38);
        Assert.InRange(Baseline(pngBaselinePage, "A") - Baseline(pngAdjustedPage, "A"),
            0.34, 0.38);
        Assert.InRange(Math.Abs(Baseline(pngBaselinePage, "B") - Baseline(pngAdjustedPage, "B")),
            0, 0.01);
    }

    [Theory]
    [InlineData("vml-gif-probe__drawing-gif-line-double")]
    [InlineData("vml-gif-probe__drawing-png-line-double")]
    [InlineData("vml-gif-probe__vml-gif-after-240")]
    public void StandaloneImageUsesSourceRunSizeForLineMetrics(string fixtureName)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            fixtureName + ".docx");
        var previousDescent = Environment.GetEnvironmentVariable("DOCXPORT_DIAG_STANDALONE_IMAGE_DESCENT");
        var previousFont = Environment.GetEnvironmentVariable("DOCXPORT_DIAG_STANDALONE_IMAGE_SOURCE_FONT");
        var previousStyleless = Environment.GetEnvironmentVariable("DOCXPORT_DIAG_STYLELESS_BASELINE");

        byte[] Export(bool sourceFont, string descentMode = "1")
        {
            Environment.SetEnvironmentVariable("DOCXPORT_DIAG_STANDALONE_IMAGE_DESCENT", descentMode);
            Environment.SetEnvironmentVariable("DOCXPORT_DIAG_STANDALONE_IMAGE_SOURCE_FONT",
                sourceFont ? "1" : null);
            Environment.SetEnvironmentVariable("DOCXPORT_DIAG_STYLELESS_BASELINE", "1");
            return DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        }

        byte[] normalBytes;
        byte[] sourceFontBytes;
        byte[] combinedBytes;
        try
        {
            normalBytes = Export(false);
            sourceFontBytes = Export(true);
            combinedBytes = Export(true, "3");
        }
        finally
        {
            Environment.SetEnvironmentVariable("DOCXPORT_DIAG_STANDALONE_IMAGE_DESCENT", previousDescent);
            Environment.SetEnvironmentVariable("DOCXPORT_DIAG_STANDALONE_IMAGE_SOURCE_FONT", previousFont);
            Environment.SetEnvironmentVariable("DOCXPORT_DIAG_STYLELESS_BASELINE", previousStyleless);
        }

        using var normal = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(normalBytes));
        using var sourceFont = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(sourceFontBytes));
        using var combined = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(combinedBytes));
        var normalPage = Assert.Single(normal.GetPages());
        var sourceFontPage = Assert.Single(sourceFont.GetPages());
        var combinedPage = Assert.Single(combined.GetPages());
        double ImageTop(UglyToad.PdfPig.Content.Page page) =>
            page.Height - Assert.Single(page.GetImages()).Bounds.Top;
        double AfterBaseline(UglyToad.PdfPig.Content.Page page) =>
            page.Height - page.Letters.First(letter => letter.Value == "A").StartBaseLine.Y;
        double BeforeBaseline(UglyToad.PdfPig.Content.Page page) =>
            page.Height - page.Letters.First(letter => letter.Value == "B").StartBaseLine.Y;

        Assert.InRange(ImageTop(normalPage) - ImageTop(sourceFontPage), -0.05, -0.02);
        Assert.InRange(AfterBaseline(normalPage) - AfterBaseline(sourceFontPage), -0.05, -0.02);
        Assert.InRange(Math.Abs(BeforeBaseline(normalPage) - BeforeBaseline(sourceFontPage)), 0, 0.01);
        Assert.InRange(Math.Abs(ImageTop(sourceFontPage) - ImageTop(combinedPage)), 0, 0.01);
        Assert.InRange(Math.Abs(AfterBaseline(sourceFontPage) - AfterBaseline(combinedPage)), 0, 0.01);
    }

    [Theory]
    [InlineData(10)]
    [InlineData(12)]
    [InlineData(14)]
    public void StandaloneAtLeastImageUsesDescentPlacementByDefault(int fontSize)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            $"vml-gif-probe__drawing-gif-atleast-font-{fontSize}-min-60.docx");
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(
            new MemoryStream(DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor())));
        var page = Assert.Single(pdf.GetPages());
        var imageTop = page.Height - Assert.Single(page.GetImages()).Bounds.Top;
        Assert.InRange(imageTop, 151.9, 152.1);
    }

    [Theory]
    [InlineData("WordInternalBookmarkLinkStories")]
    [InlineData("WordSixSlotStyledTableHeaderLinks")]
    [InlineData("WordSixSlotStyledTableLinks")]
    [InlineData("WordSixSlotStyledTablePageFields")]
    public void FirstSectionFooterTableTracksPageHeight(string fixtureName)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixtureName + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var landscape = pdf.GetPage(2);
        var portrait = pdf.GetPage(3);
        var landscapeFooter = landscape.GetWords().Where(word => word.Text == "Footer")
            .MinBy(word => landscape.Height - word.BoundingBox.Top)!;
        var portraitFooter = portrait.GetWords().Where(word => word.Text == "Footer")
            .MinBy(word => portrait.Height - word.BoundingBox.Top)!;
        var topDifference = portrait.Height - portraitFooter.BoundingBox.Top -
            (landscape.Height - landscapeFooter.BoundingBox.Top);
        Assert.InRange(topDifference, 179.8, 180.2);
    }

    [Fact]
    public void ParagraphAnchoredImageDoesNotShiftFollowingText()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "file-sample_1MB.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(4);
        var images = page.GetImages().ToArray();
        Assert.Equal(2, images.Length);
        Assert.All(images, image => Assert.InRange(image.Bounds.Width, 481.8, 482));
        var imageTops = images.Select(image => page.Height - image.Bounds.Top).Order().ToArray();
        Assert.InRange(imageTops[1] - imageTops[0], 1.2, 1.6);
        var heading = page.GetWords().First(word => word.Text == "Maecenas");
        Assert.InRange(page.Height - heading.BoundingBox.Top, 456.4, 457.0);
    }

    [Fact]
    public void FlowVisitorRendersNumberingPictureBullet()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "TestLists.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(2);
        Assert.Contains(page.GetImages(), image =>
            image.Bounds.Left is > 89.9 and < 90.1 &&
            image.Bounds.Width is > 9.9 and < 10.1);
        Assert.DoesNotContain(page.Letters, letter => letter.Value == "•" &&
            letter.GlyphRectangle.Left is > 89 and < 101 &&
            page.Height - letter.GlyphRectangle.Top is > 154 and < 171);
    }

    [Theory]
    [InlineData(false, "72pt", "36pt")]
    [InlineData(false, "1in", "0.5in")]
    [InlineData(false, "96px", "48px")]
    [InlineData(true, "72pt", "36pt")]
    [InlineData(true, "1in", "0.5in")]
    [InlineData(true, "96px", "48px")]
    public void FlowVisitorRendersInlineVmlPicture(bool embeddedObject, string width, string height)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var imagePart = main.AddImagePart(ImagePartType.Png, "rIdVML");
            var imageBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAGAAAAAwCAIAAABhdOiYAAAAZElEQVR4nO3QQQ0AIBDAMLScfymIwgE82aPJBCxde0aX1veDeIAAAQIEKBwgQIAAAQoHCBAgQIDCAQIECBCgcIAAAQIEKBwgQIAAAQoHCBAgQIDCAQIECBCgcIAAAQIEKBygRweJKyEeCj8YAgAAAABJRU5ErkJggg==");
            using (var imageStream = imagePart.GetStream(FileMode.Create, FileAccess.Write))
                imageStream.Write(imageBytes);
            OpenXmlCompositeElement carrier = embeddedObject ? new EmbeddedObject() : new Picture();
            carrier.InnerXml =
                "<v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                $"id=\"_x0000_i1025\" type=\"#_x0000_t75\" style=\"width:{width};height:{height}\">" +
                "<v:imagedata r:id=\"rIdVML\"/></v:shape>";
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("Before ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(carrier),
                new Run(new Text(" After") { Space = SpaceProcessingModeValues.Preserve }))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Single(page.GetImages());
        var image = page.GetImages().Single();
        Assert.InRange(image.Bounds.Width, 71.9, 72.1);
        Assert.InRange(image.Bounds.Height, 35.9, 36.1);
        Assert.Contains("Before", page.Text);
        Assert.Contains("After", page.Text);
    }

    [Theory]
    [InlineData(false, false, false, 160, 278, "auto", 12, 1)]
    [InlineData(false, false, true, 160, 278, "auto", 12, 1)]
    [InlineData(true, false, false, 160, 278, "auto", 12, 1)]
    [InlineData(true, true, false, 160, 278, "auto", 12, 1)]
    [InlineData(true, true, true, 160, 278, "auto", 12, 1)]
    [InlineData(true, true, false, 0, 278, "auto", 12, 1)]
    [InlineData(true, true, false, 240, 278, "auto", 12, 1)]
    [InlineData(true, true, true, 0, 278, "auto", 12, 1)]
    [InlineData(true, true, true, 240, 278, "auto", 12, 1)]
    [InlineData(true, true, true, 160, 240, "exact", 12, 1)]
    [InlineData(true, true, true, 160, 240, "auto", 12, 1)]
    [InlineData(true, true, true, 160, 360, "auto", 12, 1)]
    [InlineData(true, true, true, 160, 480, "auto", 12, 1)]
    [InlineData(true, true, true, 160, 360, "auto", 10, 1)]
    [InlineData(true, true, true, 160, 480, "auto", 10, 1)]
    [InlineData(true, true, true, 160, 360, "auto", 14, 1)]
    [InlineData(true, true, true, 160, 480, "auto", 14, 1)]
    [InlineData(true, true, true, 160, 240, "atLeast", 12, 1)]
    [InlineData(true, true, true, 160, 360, "atLeast", 12, 1)]
    [InlineData(true, true, true, 160, 1200, "atLeast", 12, 1)]
    [InlineData(true, true, false, 160, 278, "auto", 12, 2)]
    [InlineData(true, true, true, 160, 278, "auto", 12, 2)]
    public void FlowVisitorRendersGifPicture(bool cachedFieldResult, bool evaluateFields,
        bool drawingResult, int paragraphAfterTwips, int lineTwips, string lineRule,
        int fontSizePoints, int imageCount)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        using var fixture = System.IO.Compression.ZipFile.OpenRead(
            Path.Combine(root, "samples", "TestLists.docx"));
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles
            {
                InnerXml = $"<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val=\"{fontSizePoints * 2}\"/></w:rPr>" +
                    "</w:rPrDefault><w:pPrDefault><w:pPr>" +
                    $"<w:spacing w:after=\"{paragraphAfterTwips}\" w:line=\"{lineTwips}\" w:lineRule=\"{lineRule}\"/>" +
                    "</w:pPr></w:pPrDefault></w:docDefaults>"
            };
            var imagePart = main.AddImagePart(ImagePartType.Gif, "rIdVML");
            using (var imageStream = imagePart.GetStream(FileMode.Create, FileAccess.Write))
            using (var gifStream = fixture.GetEntry("word/media/image1.gif")!.Open())
                gifStream.CopyTo(imageStream);
            var picture = new Picture();
            picture.InnerXml =
                "<v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                "style=\"width:72pt;height:48pt\"><v:imagedata r:id=\"rIdVML\"/></v:shape>";
            OpenXmlElement graphicElement = picture;
            if (drawingResult)
            {
                var drawing = new Drawing();
                drawing.InnerXml =
                    "<wp:inline xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                    "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                    "<wp:extent cx=\"914400\" cy=\"609600\"/><wp:docPr id=\"1\" name=\"Picture\"/>" +
                    "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                    "<pic:pic><pic:nvPicPr><pic:cNvPr id=\"1\" name=\"Picture\"/><pic:cNvPicPr/></pic:nvPicPr>" +
                    "<pic:blipFill><a:blip r:embed=\"rIdVML\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
                    "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"609600\"/></a:xfrm>" +
                    "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>" +
                    "</a:graphicData></a:graphic></wp:inline>";
                graphicElement = drawing;
            }
            var imageParagraph = cachedFieldResult
                ? new Paragraph(
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                    new Run(new FieldCode(" ") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new FieldCode(" INCLUDEPICTURE image1.gif ")),
                    new Run(new FieldCode(" ") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                    new Run(graphicElement),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.End }))
                : new Paragraph(new Run(graphicElement));
            if (imageCount == 2)
            {
                var secondImage = new Run(graphicElement.CloneNode(true));
                if (cachedFieldResult)
                    imageParagraph.InsertBefore(secondImage, imageParagraph.LastChild!);
                else
                    imageParagraph.Append(secondImage);
            }
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Before"))),
                imageParagraph,
                new Paragraph(new Run(new Text("After")))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor(),
            new DxpExportOptions { FieldEvalMode = evaluateFields
                ? DxpFieldEvalExportMode.Evaluate : DxpFieldEvalExportMode.Cache });
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var images = pdf.GetPage(1).GetImages().ToArray();
        Assert.Equal(imageCount, images.Length);
        var image = images[0];
        Assert.InRange(image.Bounds.Width, 71.9, 72.1);
        Assert.InRange(image.Bounds.Height, 47.9, 48.1);
        if (drawingResult && fontSizePoints == 12 && lineTwips == 278)
            Assert.InRange(pdf.GetPage(1).Height - image.Bounds.Top,
                88 + paragraphAfterTwips / 20.0, 89 + paragraphAfterTwips / 20.0);
        if (drawingResult && lineRule == "exact")
            Assert.InRange(pdf.GetPage(1).Height - image.Bounds.Top, 53, 54);
        var following = Assert.Single(pdf.GetPage(1).GetWords().Where(word => word.Text == "After"));
        var expectedGap = paragraphAfterTwips / 20.0 + 4.9 +
            (lineTwips / 240.0 - 278 / 240.0) * fontSizePoints * 1.22;
        Assert.InRange(image.Bounds.Bottom - following.BoundingBox.Top,
            lineRule == "exact" ? 7 : lineRule == "atLeast" ?
                (lineTwips >= 960 ? 56.5 : 9) : expectedGap - 2,
            lineRule == "exact" ? 12 : lineRule == "atLeast" ?
                (lineTwips >= 960 ? 63 : 17) : expectedGap + 2);
    }

    [Fact]
    public void FlowVisitorJoinsFirstRowDoubleBorderAcrossCells()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRowBorderExceptionsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var commands = string.Join('\n', pdf.Pages[0].Contents.Elements
            .Select(element => (element as PdfSharp.Pdf.Advanced.PdfReference)?.Value ?? element)
            .OfType<PdfSharp.Pdf.Advanced.PdfContent>()
            .Select(content => System.Text.Encoding.Latin1.GetString(content.Stream!.UnfilteredValue)));
        Assert.True(System.Text.RegularExpressions.Regex.Matches(commands,
            @"73(?:\.0*)? [\d.]+ m\s+539(?:\.0*)? [\d.]+ l").Count >= 2);
    }

    [Fact]
    public void FlowVisitorAllocatesStyledSpacedTableBorderWithoutShiftingSecondColumn()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyledCellSpacingResetAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var words = pdf.GetPage(1).GetWords().ToArray();
        Assert.Contains(words, word => word.Text == "Header" &&
            word.BoundingBox.Left is > 329.5 and < 330.5);
        Assert.Contains(words, word => word.Text == "Footer" &&
            word.BoundingBox.Left is > 329.5 and < 330.5);
    }

    [Fact]
    public void FlowVisitorPaintsDoubleLogicalCellBorder()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleCustomCellBorderColors.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var strokes = pdf.GetPage(1).Paths.Where(path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 304 and < 309 && bounds.Height > 12 &&
            pdf.GetPage(1).Height - bounds.Top is > 49 and < 52).ToArray();
        Assert.Contains(strokes, path => path.GetBoundingRectangle()!.Value.Left is > 304 and < 306);
        Assert.Contains(strokes, path => path.GetBoundingRectangle()!.Value.Left is > 306 and < 308);
    }

    [Fact]
    public void FlowVisitorDrawsInheritedCharacterStyleStrikeThrough()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordBaselineOffsetLayeredStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.Paths, path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 71 and < 73 && bounds.Width is > 63 and < 66 &&
            page.Height - bounds.Top is > 79 and < 81);
    }

    [Fact]
    public void FlowVisitorExtendsStrikeThroughAcrossStyledTab()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordTextControlsNbspAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.Paths, path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 71 and < 73 && bounds.Right is > 143 and < 145 &&
            page.Height - bounds.Top is > 79 and < 83);
    }

    [Theory]
    [InlineData("WordBmpRle4AbsoluteAllStories.docx")]
    [InlineData("WordBmpRle4DeltaAllStories.docx")]
    public void FlowVisitorDisablesInterpolationForEmbeddedBitmap(string fixtureName)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixtureName);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var objects = pdf.Pages[0].Resources.Elements.GetDictionary("/XObject");
        Assert.NotNull(objects);
        Assert.Contains(objects.Elements.Values, value =>
            (value as PdfSharp.Pdf.Advanced.PdfReference)?.Value is PdfSharp.Pdf.PdfDictionary image &&
            image.Elements.GetName("/Subtype") == "/Image" &&
            image.Elements.GetInteger("/Width") == 16 &&
            !image.Elements.GetBoolean("/Interpolate"));
    }

    [Theory]
    [InlineData("WordFlippedImageStories")]
    [InlineData("WordRightAngleImageStories")]
    [InlineData("WordArbitraryRotationStories")]
    public void FlowVisitorDisablesInterpolationForPngImages(string fixture)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            fixture + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var objects = pdf.Pages[0].Resources.Elements.GetDictionary("/XObject");
        Assert.NotNull(objects);
        Assert.Contains(objects.Elements.Values, value =>
            (value as PdfSharp.Pdf.Advanced.PdfReference)?.Value is PdfSharp.Pdf.PdfDictionary image &&
            image.Elements.GetName("/Subtype") == "/Image" &&
            image.Elements.GetName("/Filter") == "/FlateDecode" &&
            !image.Elements.GetBoolean("/Interpolate"));
    }

    [Fact]
    public void FlowVisitorKeepsInterpolationForJpegImage()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordTwoFloatingPicturesBody.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var objects = pdf.Pages[0].Resources.Elements.GetDictionary("/XObject");
        Assert.NotNull(objects);
        Assert.Contains(objects.Elements.Values, value =>
            (value as PdfSharp.Pdf.Advanced.PdfReference)?.Value is PdfSharp.Pdf.PdfDictionary image &&
            image.Elements.GetName("/Subtype") == "/Image" &&
            image.Elements.GetName("/Filter") == "/DCTDecode" &&
            image.Elements.GetBoolean("/Interpolate"));
    }

    [Fact]
    public void FlowVisitorRendersSuperscriptCapsHeaderAtWordSize()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordCapsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.Letters, letter => letter.Value == "H" &&
            letter.GlyphRectangle.Left is > 71 and < 73 &&
            page.Height - letter.GlyphRectangle.Top < 50 &&
            letter.FontSize is > 6.9 and < 7.1);
    }

    [Fact]
    public void FlowVisitorDrawsLayeredCharacterOutlines()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordCharacterEffectsLayeredStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Contains("/ca 0", System.Text.Encoding.Latin1.GetString(bytes));
        var streams = pdf.Pages[0].Contents.Elements.Select(element =>
            (element as PdfSharp.Pdf.Advanced.PdfReference)?.Value is PdfSharp.Pdf.Advanced.PdfContent content &&
            content.Stream != null
                ? System.Text.Encoding.Latin1.GetString(content.Stream.UnfilteredValue)
                : "").ToArray();
        Assert.Contains(streams, stream => stream.Contains("0.65 0.65 0.65 RG\n0.5 w\nBT\n2 Tr"));
        Assert.Contains(streams, stream => stream.Contains("0 0 0 RG\n0.14 w\nBT\n2 Tr"));
    }

    [Fact]
    public void FlowVisitorKeepsInheritedOutlineWithConflictingLayeredEffects()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyledCharacterEffectsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var streams = pdf.Pages[0].Contents.Elements.Select(element =>
            (element as PdfSharp.Pdf.Advanced.PdfReference)?.Value is PdfSharp.Pdf.Advanced.PdfContent content &&
            content.Stream != null
                ? System.Text.Encoding.Latin1.GetString(content.Stream.UnfilteredValue)
                : "").ToArray();
        Assert.Contains(streams, stream => stream.Contains("0.65 0.65 0.65 RG\n0.5 w\nBT\n2 Tr"));
        Assert.Contains(streams, stream => stream.Contains("0 0 0 RG\n0.14 w\nBT\n2 Tr"));
        using var extracted = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = extracted.GetPage(1).Letters;
        Assert.True(letters.Count(letter => letter.Value == "H" && letter.StartBaseLine.Y > 700) >= 2);
        Assert.True(letters.Count(letter => letter.Value == "F" && letter.StartBaseLine.Y < 100) >= 2);
    }

    [Fact]
    public void FlowVisitorAppliesInheritedAndDirectCharacterScaling()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordCharacterScaleLayeredStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.Letters, letter => letter.Value == "d" &&
            page.Height - letter.GlyphRectangle.Top is > 35 and < 50 &&
            letter.GlyphRectangle.Right is > 116 and < 118);
        Assert.Contains(page.Letters, letter => letter.Value == "d" &&
            page.Height - letter.GlyphRectangle.Top is > 70 and < 90 &&
            letter.GlyphRectangle.Right is > 148 and < 150);
        Assert.Contains(page.Letters, letter => letter.Value == "d" &&
            page.Height - letter.GlyphRectangle.Top is > 730 and < 750 &&
            letter.GlyphRectangle.Right is > 131 and < 134);
        foreach (var expectedRight in new[] { 117.12, 149.04, 132.62 })
            Assert.Contains(page.Paths, path => path.IsStroked &&
                path.GetBoundingRectangle() is { } bounds &&
                bounds.Left is > 71 and < 73 && bounds.Height < 1 &&
                Math.Abs(bounds.Right - expectedRight) < 0.5);
    }

    [Fact]
    public void FlowVisitorDrawsInheritedTextBordersWithCharacterSpacing()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordCharacterSpacingStyleOverridesAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.True(page.Paths.Count(path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Width is > 65 and < 110 && bounds.Height is > 19 and < 23) >= 3);
        var secondPage = pdf.GetPage(2);
        var letters = secondPage.Letters.Where(letter =>
            secondPage.Height - letter.GlyphRectangle.Top is > 70 and < 90).ToArray();
        var mirroredStart = letters.First(letter => letter.Value == "m");
        var precedingPageEnd = letters.Where(letter => letter.Value == "e" &&
                letter.GlyphRectangle.Right < mirroredStart.GlyphRectangle.Left)
            .Max(letter => letter.GlyphRectangle.Right);
        var bodyStart = letters.First(letter => letter.Value == "b");
        var precedingMirroredEnd = letters.Where(letter => letter.Value == "d" &&
                letter.GlyphRectangle.Right < bodyStart.GlyphRectangle.Left)
            .Max(letter => letter.GlyphRectangle.Right);
        Assert.InRange(mirroredStart.GlyphRectangle.Left - precedingPageEnd, 7.5, 10);
        Assert.InRange(bodyStart.GlyphRectangle.Left - precedingMirroredEnd, 7.5, 10);
    }

    [Fact]
    public void FlowVisitorKeepsMirroredStoryRunShadingWithinTextWidth()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordCombinedStyleVisualAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(2);
        var shades = page.Paths.Where(path => path.IsFilled &&
            path.FillColor?.ToRGBValues() is { } color &&
            color.r is > 0.71 and < 0.74 && color.g is > 0.87 and < 0.89 &&
            color.b is > 0.95 and < 0.97).ToArray();
        Assert.Contains(shades, path => path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 111 and < 112 &&
            bounds.Width is > 76 and < 80 &&
            page.Height - bounds.Top is > 727 and < 729);
        Assert.Contains(shades, path => path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 107.5 and < 108.5 &&
            bounds.Width is > 70 and < 71 &&
            page.Height - bounds.Top is > 35 and < 37);
    }

    [Fact]
    public void FlowVisitorDrawsRunBorderShadows()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRgbShadowRunBordersAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Contains(pdf.GetPage(1).Paths, path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 233 and < 239 && bounds.Width < 2 &&
            bounds.Height is > 17 and < 19);
        var page = pdf.GetPage(1);
        var mirrored = page.GetWords().Where(word => word.Text == "Mirrored")
            .OrderBy(word => page.Height - word.BoundingBox.Top).ToArray();
        Assert.Equal(3, mirrored.Length);
        Assert.All(mirrored, word => Assert.InRange(word.BoundingBox.Left, 148.5, 149.5));
        Assert.Contains(pdf.GetPage(2).GetWords(), word => word.Text == "Mirrored" &&
            word.BoundingBox.Left is > 112.5 and < 113.5);
        var bodyTop = page.Height - mirrored[1].BoundingBox.Top;
        var footerTop = page.Height - mirrored[2].BoundingBox.Top;
        Assert.InRange(footerTop - bodyTop, 645, 649);
    }

    [Theory]
    [InlineData("WordCharacterStyleRunBordersAllStories")]
    [InlineData("WordDerivedCharacterStyleRunBordersAllStories")]
    [InlineData("WordRgbShadowRunBordersAllStories")]
    public void FlowVisitorReservesDirectRunBorderClearance(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.Letters, letter => letter.Value == "M" &&
            letter.GlyphRectangle.Left is > 145 and < 153 &&
            page.Height - letter.GlyphRectangle.Top is > 35 and < 49);
        Assert.Contains(page.Letters, letter => letter.Value == "R" &&
            letter.GlyphRectangle.Left is > 107 and < 113 &&
            page.Height - letter.GlyphRectangle.Top is > 101 and < 113);
        if (name == "WordRgbShadowRunBordersAllStories") return;
        var header = page.GetWords().First(word => word.Text == "Mirrored" &&
            word.BoundingBox.Left > 140);
        Assert.InRange(page.Height - header.Letters[0].StartBaseLine.Y, 50.4, 51.3);
    }

    [Fact]
    public void FlowVisitorStopsInheritedRunBorderAtReset()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedRunBorderResetAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Contains(pdf.GetPage(1).Paths, path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 144 and < 146 && bounds.Right is > 221 and < 225 &&
            bounds.Width is > 75 and < 81);
    }

    [Fact]
    public void FlowVisitorReservesInlineClearanceAfterInheritedRunBorder()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedRunBorderResetAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var words = pdf.GetPage(1).GetWords().ToArray();
        var plainWords = words.Where(word => word.Text == "Plain")
            .OrderByDescending(word => word.BoundingBox.Bottom).ToArray();
        Assert.Equal(3, plainWords.Length);
        foreach (var plain in plainWords)
        {
            var preceding = words.Where(word =>
                    Math.Abs(word.BoundingBox.Bottom - plain.BoundingBox.Bottom) < 1 &&
                    word.BoundingBox.Right < plain.BoundingBox.Left)
                .MaxBy(word => word.BoundingBox.Right);
            Assert.NotNull(preceding);
            Assert.InRange(plain.BoundingBox.Left - preceding.BoundingBox.Right, 4.7, 5.3);
        }
    }

    [Theory]
    [InlineData("WordClearedInheritedTabsAllStories", 728.75)]
    [InlineData("WordInheritedTabLeadersAllStories", 728.75)]
    [InlineData("WordStyledTabLeadersAllStories", 728.75)]
    [InlineData("WordTabLeadersAllStories", 738.94)]
    public void FlowVisitorAlignsTabbedFooterBeforeEmptyParagraph(string name, double expectedBaseline)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var footer = page.Letters.First(letter => letter.Value == "F" && letter.StartBaseLine.Y < 100);
        Assert.InRange(page.Height - footer.StartBaseLine.Y, expectedBaseline - 0.4, expectedBaseline + 0.4);
    }

    [Fact]
    public void FlowVisitorDrawsDoubleRunBorders()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordLayeredRunBordersAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.GetPage(1).Paths.Count(path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds && bounds.Left is > 140 and < 150 &&
            bounds.Width is > 75 and < 90 && bounds.Height is > 18 and < 30) >= 2);
        var firstPage = pdf.GetPage(1);
        var body = firstPage.GetWords().First(word => word.Text == "Mirrored" &&
            firstPage.Height - word.BoundingBox.Top is > 60 and < 200);
        var footer = firstPage.GetWords().Last(word => word.Text == "Mirrored");
        Assert.InRange(firstPage.Height - body.BoundingBox.Top, 78, 87);
        Assert.InRange(firstPage.Height - body.Letters[0].StartBaseLine.Y, 89.3, 90.6);
        Assert.InRange(firstPage.Height - footer.BoundingBox.Top, 720, 735);
        Assert.InRange(body.BoundingBox.Left, 150, 151);
        Assert.InRange(footer.BoundingBox.Left, 150, 151);
        Assert.Contains(firstPage.GetWords(), word => word.Text == "Reset" &&
            word.BoundingBox.Left is > 111 and < 112);
        var secondPage = pdf.GetPage(2);
        var footerBorder = secondPage.Paths.Where(path => path.IsStroked &&
                path.GetBoundingRectangle() is { } bounds && bounds.Left is > 107 and < 110 &&
                bounds.Width is > 85 and < 91 && bounds.Height is > 23 and < 28)
            .Select(path => path.GetBoundingRectangle()!.Value).First();
        Assert.InRange(secondPage.Height - footerBorder.Top, 717.0, 717.6);
    }

    [Fact]
    public void DoubleRunBorderTextPaddingRestoresTheNextParagraphOrigin()
    {
        var stream = string.Join('\n', "BT", "144 744 Td", "(Mirrored) Tj",
            "46.5 0 Td", "(header) Tj", "-43 -686 Td", "(Mirrored) Tj",
            "46.5 0 Td", "(body) Tj", "-82.5 -36 Td", "(Reset) Tj", "ET");
        var revised = DxpFlowPdfVisitor.RewriteDoubleBorderTextPadding(stream,
            new Dictionary<string, double> { ["Mirrored\0body"] = 3 });

        Assert.Contains("-40 -686 Td\n(Mirrored) Tj", revised);
        Assert.Contains("-85.5 -36 Td\n(Reset) Tj", revised);
        Assert.Contains("144 744 Td\n(Mirrored) Tj", revised);
    }

    [Fact]
    public void FlowVisitorKeepsHeaderFooterGridGapRowsCompact()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleRowGridGapsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var headers = page.GetWords().Where(word => word.Text == "Header" &&
            word.BoundingBox.Left is > 76 and < 80).OrderByDescending(word => word.BoundingBox.Top).ToArray();
        var footers = page.GetWords().Where(word => word.Text == "Footer" &&
            word.BoundingBox.Left is > 76 and < 80).OrderByDescending(word => word.BoundingBox.Top).ToArray();
        Assert.Equal(2, headers.Length);
        Assert.Equal(2, footers.Length);
        Assert.InRange(headers[0].BoundingBox.Top - headers[1].BoundingBox.Top, 27, 30);
        Assert.InRange(footers[0].BoundingBox.Top - footers[1].BoundingBox.Top, 27, 30);
        var headerIntro = page.GetWords().Single(word => word.Text == "Header" &&
            word.BoundingBox.Left is > 71 and < 73);
        var bodyIntro = page.GetWords().Single(word => word.Text == "Body" &&
            word.BoundingBox.Left is > 71 and < 73);
        Assert.InRange(headerIntro.BoundingBox.Top - bodyIntro.BoundingBox.Top, 72, 75);
    }

    [Fact]
    public void FlowVisitorClearsSingleRowHeaderBeforeGridGapBody()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleRowGridGaps.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var header = page.GetWords().Single(word => word.Text == "Header" &&
            word.BoundingBox.Left is > 71 and < 73);
        var body = page.GetWords().Single(word => word.Text == "Body" &&
            word.BoundingBox.Left is > 71 and < 73);
        Assert.InRange(header.BoundingBox.Top - body.BoundingBox.Top, 43, 45);
    }

    [Fact]
    public void FlowVisitorPlacesColoredBodyUnderlineBelowBaseline()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordUnderlineColorStyleOverridesAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var underline = Assert.Single(page.Paths.Where(path => path.IsFilled &&
            path.FillColor != null && path.FillColor.ToRGBValues() is var rgb &&
            rgb.r is > 0.78 and < 0.8 && rgb.g is > 0.2 and < 0.23 &&
            path.GetBoundingRectangle() is { } bounds && bounds.Width is > 70 and < 80));
        var box = underline.GetBoundingRectangle()!.Value;
        Assert.InRange(page.Height - box.Top, 87.4, 87.7);
        Assert.InRange(box.Height, 0.2, 0.3);
    }

    [Fact]
    public void FlowVisitorUsesWordSizedNativeUnderline()
    {
        if (typeof(MigraDoc.Rendering.DocumentRenderer).GetProperty("UnderlineStrokeWidthFactor") == null)
            return;
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Right to Repair.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var streams = pdf.Pages.Cast<PdfSharp.Pdf.PdfPage>().SelectMany(page => page.Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Select(reference => reference.Value)
            .OfType<PdfSharp.Pdf.Advanced.PdfContent>()
            .Where(content => content.Stream != null)
            .Select(content => System.Text.Encoding.Latin1.GetString(content.Stream!.UnfilteredValue)));
        var hasFontAwareGeometry = typeof(MigraDoc.Rendering.DocumentRenderer)
            .GetProperty("UnderlineGeometry") != null;
        var expectedWidth = hasFontAwareGeometry ? @"1\.14" : @"(?:0)?\.3";
        var underline = streams.Select(stream => System.Text.RegularExpressions.Regex.Match(stream,
                $@"(?<width>{expectedWidth}) w\s+(?<x>[\d.]+) (?<y>[\d.]+) m\s+[\d.]+ [\d.]+ l\s+S"))
            .First(match => match.Success);
        var width = double.Parse(underline.Groups["width"].Value,
            System.Globalization.CultureInfo.InvariantCulture);
        Assert.InRange(width, hasFontAwareGeometry ? 1.13 : 0.29,
            hasFontAwareGeometry ? 1.15 : 0.31);
    }

    [Fact]
    public void FlowVisitorConfiguresNinePointAptosDisplayUnderlineDepth()
    {
        if (typeof(MigraDoc.Rendering.DocumentRenderer).GetProperty("UnderlineGeometry") == null)
            return;
        var visitor = new DxpFlowPdfVisitor();
        var create = typeof(DxpFlowPdfVisitor).GetMethod("CreatePdfRenderer",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!;
        var renderer = (MigraDoc.Rendering.PdfDocumentRenderer)create.Invoke(visitor,
            [new MigraDoc.DocumentObjectModel.Document()])!;
        var callback = (Func<MigraDoc.DocumentObjectModel.Font, (double StrokeWidthFactor,
            double VerticalOffsetFactor)?>)typeof(MigraDoc.Rendering.DocumentRenderer)
            .GetProperty("UnderlineGeometry")!.GetValue(renderer.DocumentRenderer)!;
        var font = new MigraDoc.DocumentObjectModel.Font
        {
            Name = "Aptos Display",
            Size = MigraDoc.DocumentObjectModel.Unit.FromPoint(9)
        };

        Assert.Equal((0.025, 0.29), callback(font));
    }

    [Fact]
    public void FlowVisitorConfiguresDejaVuSansUnderlineGeometry()
    {
        if (typeof(MigraDoc.Rendering.DocumentRenderer).GetProperty("UnderlineGeometry") == null)
            return;
        var visitor = new DxpFlowPdfVisitor();
        var create = typeof(DxpFlowPdfVisitor).GetMethod("CreatePdfRenderer",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!;
        var renderer = (MigraDoc.Rendering.PdfDocumentRenderer)create.Invoke(visitor,
            [new MigraDoc.DocumentObjectModel.Document()])!;
        var callback = (Func<MigraDoc.DocumentObjectModel.Font, (double StrokeWidthFactor,
            double VerticalOffsetFactor)?>)typeof(MigraDoc.Rendering.DocumentRenderer)
            .GetProperty("UnderlineGeometry")!.GetValue(renderer.DocumentRenderer)!;
        var font = new MigraDoc.DocumentObjectModel.Font
        {
            Name = "DejaVu Sans",
            Size = MigraDoc.DocumentObjectModel.Unit.FromPoint(10.5)
        };

        Assert.Equal((0.045, -0.15), callback(font));
    }

    [Fact]
    public void FlowVisitorPreservesExternalHyperlinkAnnotationsAcrossStories()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordLinksAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Equal(2, pdf.PageCount);
        Assert.Equal(3, pdf.Pages[0].Annotations.Count);
        Assert.Equal(2, pdf.Pages[1].Annotations.Count);
    }

    [Theory]
    [InlineData(DxpFieldEvalExportMode.Cache)]
    [InlineData(DxpFieldEvalExportMode.Evaluate)]
    public void FlowVisitorPreservesComplexFieldHyperlinksInFootnotes(DxpFieldEvalExportMode fieldMode)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Right to Repair.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(),
            new DxpExportOptions { FieldEvalMode = fieldMode });

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var links = pdf.Pages.Cast<PdfSharp.Pdf.PdfPage>()
            .SelectMany(page => page.Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>())
            .Select(annotation => annotation.Elements.GetDictionary("/A")?.Elements.GetString("/URI"))
            .Where(uri => !string.IsNullOrEmpty(uri)).ToArray();
        Assert.True(links.Length >= 45, $"Expected at least 45 links; found {links.Length}.");
        Assert.Equal(30, links.Distinct().Count());
        Assert.Contains("https://www.parl.ca/legisinfo/en/bill/44-1/c-244", links);
    }

    [Fact]
    public void FlowVisitorPreservesInternalBookmarkLinkAnnotations()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInternalBookmarkLinkStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Equal(3, pdf.PageCount);
        Assert.Equal(1, pdf.Pages[1].Annotations.Count);
        Assert.Equal(1, pdf.Pages[2].Annotations.Count);
    }

    [Fact]
    public void FlowVisitorKeepsTocTabsAndPageNumbersInsideInternalLinks()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "TOC.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(),
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Cache });

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Equal(6, pdf.Pages[0].Annotations.Count);
        using var extracted = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var text = extracted.GetPage(1).Text;
        Assert.Contains("A.1", text);
        Assert.DoesNotContain("A.11", text);
    }

    [Fact]
    public void FlowVisitorCoalescesRunBorderAcrossTab()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordSixSlotCombinedStyles.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Single(pdf.GetPage(1).Paths.Where(path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 71.7 and < 72.3 && bounds.Right is > 337 and < 338 &&
            bounds.Height is > 13 and < 16));
        var storyBorders = pdf.GetPage(1).Paths.Where(path => path.IsStroked &&
                path.GetBoundingRectangle() is { } bounds &&
                bounds.Left is > 72.3 and < 72.7 && bounds.Width is > 140 and < 150 &&
                bounds.Height is > 13 and < 14)
            .Select(path => pdf.GetPage(1).Height - path.GetBoundingRectangle()!.Value.Top)
            .OrderBy(top => top).ToArray();
        Assert.Equal(2, storyBorders.Length);
        Assert.InRange(storyBorders[0], 36.4, 36.8);
        Assert.InRange(storyBorders[1], 737.9, 738.3);
        var blueShades = pdf.GetPage(1).Paths.Where(path => path.IsFilled &&
            path.FillColor != null && path.FillColor.ToRGBValues() is var rgb &&
            rgb.r is > 0.71 and < 0.74 && rgb.g is > 0.87 and < 0.89 &&
            rgb.b is > 0.95 and < 0.97).ToArray();
        Assert.Contains(blueShades, path => path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 73 and < 75 && bounds.Right is > 217 and < 219 &&
            pdf.GetPage(1).Height - bounds.Top is > 37 and < 40);
        Assert.Contains(blueShades, path => path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 179.5 and < 180.5 && bounds.Right is > 336 and < 337 &&
            pdf.GetPage(1).Height - bounds.Top is > 73 and < 75);
        Assert.Contains(blueShades, path => path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 73.5 and < 74.5 && bounds.Right is > 213.5 and < 214.5 &&
            pdf.GetPage(1).Height - bounds.Top is > 738 and < 741);
        var redUnderlines = pdf.GetPage(1).Paths.Where(path => path.IsFilled &&
            path.FillColor != null && path.FillColor.ToRGBValues() is var rgb &&
            rgb.r is > 0.78 and < 0.8 &&
            rgb.g is > 0.2 and < 0.23 && rgb.b is > 0.14 and < 0.17 &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Width > 130 && bounds.Height is > 0.9 and < 1.1).ToArray();
        Assert.True(redUnderlines.Length >= 2);
        var bodyWords = pdf.GetPage(1).GetWords()
            .Where(word => word.Text is "First" or "paragraph.")
            .ToArray();
        Assert.InRange(bodyWords.Single(word => word.Text == "paragraph.").BoundingBox.Right,
            332, 335);
        Assert.InRange(bodyWords.Single(word => word.Text == "First").BoundingBox.Left,
            179.7, 180.4);
        var bodyTab = pdf.GetPage(1).GetWords().Single(word => word.Text == "Tab");
        Assert.InRange(pdf.GetPage(1).Height - bodyTab.BoundingBox.Top, 75.5, 76.2);
        Assert.Contains(pdf.GetPage(1).Paths, path => path.IsFilled &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 87 and < 89 && bounds.Right is > 179.5 and < 180.5 &&
            pdf.GetPage(1).Height - bounds.Top is > 73 and < 75);
        var header = pdf.GetPage(1).Letters.First(letter => letter.Value == "F" &&
            letter.GlyphRectangle.Left is > 73 and < 75);
        Assert.InRange(pdf.GetPage(1).Height - header.GlyphRectangle.Top, 39, 41);
        var linkedHeader = pdf.GetPage(3).Letters.First(letter => letter.Value == "F" &&
            letter.GlyphRectangle.Left is > 73 and < 75);
        Assert.InRange(pdf.GetPage(3).Height - linkedHeader.GlyphRectangle.Top, 39, 41);
    }

    [Fact]
    public void FlowVisitorKeepsUnshadedRunBorderHeaderInSecondSection()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordSixSlotRunBorderInheritance.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var first = pdf.GetPage(1).Letters.First(letter => letter.Value == "F" &&
            letter.GlyphRectangle.Left is > 75 and < 77);
        var second = pdf.GetPage(2).Letters.First(letter => letter.Value == "F" &&
            letter.GlyphRectangle.Left is > 90 and < 94 &&
            pdf.GetPage(2).Height - letter.GlyphRectangle.Top < 100);
        Assert.InRange(pdf.GetPage(1).Height - first.GlyphRectangle.Top, 38, 42);
        Assert.InRange(pdf.GetPage(2).Height - second.GlyphRectangle.Top, 35, 38);
        Assert.Contains(pdf.GetPage(1).Paths, path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 71.7 and < 72.3 && bounds.Right is > 309.7 and < 310.3 &&
            pdf.GetPage(1).Height - bounds.Top is > 73 and < 74);
        Assert.Contains(pdf.GetPage(1).Paths, path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 72.7 and < 72.8 && bounds.Height is > 15 and < 16 &&
            pdf.GetPage(1).Height - bounds.Top is > 36.7 and < 37.2);
        var bodyFirst = pdf.GetPage(1).Letters.First(letter => letter.Value == "F" &&
            letter.GlyphRectangle.Left is > 175 and < 185);
        Assert.InRange(bodyFirst.GlyphRectangle.Left, 179.7, 180.6);
    }

    [Fact]
    public void FlowVisitorPositionsUnstyledFootersAcrossPageGeometries()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordMixedGeometrySectionStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(4, pdf.NumberOfPages);
        foreach (var (pageNumber, top) in new[] { (1, 561.6), (3, 751.6), (4, 751.6) })
        {
            var page = pdf.GetPage(pageNumber);
            var footer = page.Letters.First(letter =>
                page.Height - letter.GlyphRectangle.Top > page.Height - 70);
            Assert.InRange(page.Height - footer.GlyphRectangle.Top, top - 1, top + 1);
        }
    }

    [Theory]
    [InlineData("WordMixedGeometrySectionStories.docx")]
    [InlineData("WordMixedGeometryPageBorderStories.docx")]
    public void FlowVisitorUsesSectionPageNumberFormatInFooter(string fileName)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fileName);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(4, pdf.NumberOfPages);
        Assert.Contains("PortraitevenfooterVIII", pdf.GetPage(4).Text);
        Assert.DoesNotContain("Portraitevenfooter8", pdf.GetPage(4).Text);
    }

    [Theory]
    [InlineData("lowerRoman", "viii")]
    [InlineData("upperLetter", "H")]
    [InlineData("lowerLetter", "h")]
    public void FlowVisitorFormatsPageFieldWithSectionNumbering(string format, string expected)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordMixedGeometrySectionStories.docx");
        using var stream = new MemoryStream();
        stream.Write(File.ReadAllBytes(source));
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            foreach (var pageNumbers in document.MainDocumentPart!.Document.Descendants<PageNumberType>())
                pageNumbers.Format = format switch
                {
                    "lowerRoman" => NumberFormatValues.LowerRoman,
                    "upperLetter" => NumberFormatValues.UpperLetter,
                    _ => NumberFormatValues.LowerLetter
                };
            document.MainDocumentPart.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(stream.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(4, pdf.NumberOfPages);
        Assert.Contains("Portraitevenfooter" + expected, pdf.GetPage(4).Text);
    }


    [Fact]
    public void FlowVisitorTogglesInheritedBoldAndItalicInCharacterStyle()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordCharacterStyleBaseToggleStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var firstLetters = page.Letters.Where(letter => letter.Value is "H" or "B").ToArray();
        Assert.Equal(2, firstLetters.Length);
        Assert.All(firstLetters, letter =>
        {
            Assert.DoesNotContain("Bold", letter.Font.Name, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Italic", letter.Font.Name, StringComparison.OrdinalIgnoreCase);
        });
    }

    [Fact]
    public void FlowVisitorTogglesInheritedStrikeInCharacterStyle()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordCharacterStyleInheritedStrikeStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.DoesNotContain(pdf.GetPage(1).Paths, path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds && bounds.Width > 20);
    }

    [Fact]
    public void BmpTransformRotatesAndFlipsPixels()
    {
        var bitmap = new byte[70];
        bitmap[0] = (byte)'B';
        bitmap[1] = (byte)'M';
        BinaryPrimitives.WriteInt32LittleEndian(bitmap.AsSpan(18), 2);
        BinaryPrimitives.WriteInt32LittleEndian(bitmap.AsSpan(22), -2);
        BinaryPrimitives.WriteUInt16LittleEndian(bitmap.AsSpan(28), 24);
        bitmap[54] = 1;
        bitmap[57] = 2;
        bitmap[62] = 3;
        bitmap[65] = 4;

        var rotated = DxpBmp24Converter.Transform(bitmap, 90, false, false);
        var flipped = DxpBmp24Converter.Transform(bitmap, 0, true, false);

        Assert.Equal(new byte[] { 2, 4, 1, 3 }, new[] { rotated[54], rotated[57], rotated[62], rotated[65] });
        Assert.Equal(new byte[] { 2, 1, 4, 3 }, new[] { flipped[54], flipped[57], flipped[62], flipped[65] });
    }

    [Fact]
    public void FlowVisitorRendersSampleChart()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "file-sample_1MB.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Equal(5, pdf.PageCount);
        Assert.True(bytes.Length > 100_000);

        using var textPdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.DoesNotContain("12.0", textPdf.GetPage(1).Text);
        var chartPage = textPdf.GetPage(1);
        var chartWords = chartPage.GetWords().ToArray();
        var lastBullet = chartWords.Single(word => word.Text == "tincidunt");
        var firstCategory = chartWords.Where(word => word.Text == "Row")
            .OrderByDescending(word => word.BoundingBox.Top).First();
        Assert.InRange(lastBullet.BoundingBox.Top - firstCategory.BoundingBox.Top, 173.5, 174.5);
        var gridlines = chartPage.Paths.Where(path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds && bounds.Width is > 215 and < 225)
            .Select(path => path.GetBoundingRectangle()!.Value).ToArray();
        Assert.Contains(gridlines, bounds => bounds.Left is > 82 and < 84 &&
            bounds.Width is > 219 and < 221 &&
            chartPage.Height - bounds.Top is > 528 and < 532);
        if (typeof(MigraDoc.DocumentObjectModel.Shapes.Charts.Axis)
            .GetProperty("MajorTickMarkLength") != null)
        {
            var categoryTicks = chartPage.Paths.Where(path => path.IsStroked &&
                    path.GetBoundingRectangle() is { } bounds &&
                    bounds.Left is > 82 and < 305 && bounds.Width < 0.2 &&
                    bounds.Height is > 3.0 and < 3.3)
                .Select(path => path.GetBoundingRectangle()!.Value.Left)
                .OrderBy(left => left).ToArray();
            Assert.Equal(5, categoryTicks.Length);
            foreach (var (actual, expected) in categoryTicks.Zip(
                new[] { 83.55, 138.35, 193.14, 247.94, 302.74 }))
                Assert.InRange(actual, expected - 0.1, expected + 0.1);
            var zeroLabel = chartWords.First(word => word.Text == "0" &&
                word.BoundingBox.Left < 80);
            Assert.InRange(zeroLabel.BoundingBox.Top, 189.15, 189.30);
        }
        var categoryLabel = chartPage.Letters.First(letter => letter.Value == "R" &&
            letter.GlyphRectangle.Left is > 90 and < 110);
        Assert.InRange(categoryLabel.PointSize, 9.9, 10.1);
        var legendLabel = chartPage.GetWords().First(word => word.Text == "Column");
        Assert.InRange(legendLabel.BoundingBox.Left, 326.5, 328.5);
        var legendSwatches = chartPage.Paths.Where(path => path.IsFilled &&
                path.GetBoundingRectangle() is { } bounds &&
                bounds.Left is > 318 and < 321 && bounds.Width is > 5 and < 6)
            .Select(path => path.GetBoundingRectangle()!.Value).ToArray();
        Assert.Equal(3, legendSwatches.Length);
        Assert.All(legendSwatches, bounds => Assert.InRange(bounds.Height, 5, 6));
    }

    [Fact]
    public void FlowVisitorShiftsOnlyMatchingRightChartLegendStream()
    {
        var stream = "q\n1 -0 -0 1 63.7 555.97 cm\n0.15 w\n" +
            "245 -330 7 7 re\nB\nBT\n256 -330 Td\n(Column 1) Tj\nET\n" +
            "245 -348 7 7 re\nB\nBT\n256 -348 Td\n(Column 2) Tj\nET\n" +
            "1 w\n18 -402 m\nQ\nBT\n33 -415 Td\n(Row 1) Tj\n55 0 Td\n(Row 2) Tj\nET\n" +
            "BT\n/F6 10 Tf\n6 -405 Td\n(0) Tj\nET\nBT\n6 -384 Td\n(2) Tj\nET\nQ";
        var labels = new[] { new[] { "Column 1", "Column 2" } };

        var rewritten = DxpFlowPdfVisitor.RewriteRightLegendStream(stream, labels, labels);

        Assert.NotNull(rewritten);
        Assert.Contains("cm\n1 0 0 1 7 0 cm\n0.15 w", rewritten);
        Assert.Contains("248.3 -330 5.5 5.5 re", rewritten);
        Assert.Contains("248.3 -348.87 5.5 5.5 re", rewritten);
        Assert.Contains("256 -330.3 Td", rewritten);
        Assert.Contains("256 -349.17 Td", rewritten);
        Assert.Contains("0.9982 0 0 1 0.76 0 cm\n1 w\n18 -402 m\nQ\n" +
            "1.0018032 0 0 1 -0.761371 0 cm", rewritten);
        Assert.Contains("33.7 -418.17 Td\n(Row 1) Tj", rewritten);
        Assert.Contains("54.901 0 Td\n(Row 2) Tj", rewritten);
        Assert.Contains("3.95 -405.1 Td\n(0) Tj", rewritten);
        Assert.Contains("3.95 -384.175 Td\n(2) Tj", rewritten);
        Assert.Contains("(Column 2) Tj\nET\n1 0 0 1 -7 0 cm\n" +
            "0.9982 0 0 1 0.76 0 cm\n1 w", rewritten);
        Assert.Contains("33 -415 Td\n(Row 1) Tj",
            DxpFlowPdfVisitor.RewriteRightLegendStream(stream, labels));
        Assert.Contains("6 -405 Td\n(0) Tj",
            DxpFlowPdfVisitor.RewriteRightLegendStream(stream, labels));
        var tickStream = stream.Replace("BT\n/F6 10 Tf\n6 -405 Td",
            "1 w\n17.1336 -402.3936 m\n17.1336 -405.5236 l\nS\n" +
            "72.3262 -402.3936 m\n72.3262 -405.5236 l\nS\n" +
            "BT\n/F6 10 Tf\n6 -405 Td");
        var tickRewritten = DxpFlowPdfVisitor.RewriteRightLegendStream(tickStream, labels, labels);
        Assert.Contains("17.8628 -402.3936 m\n17.8628 -405.5236 l", tickRewritten);
        Assert.Contains("72.956 -402.3936 m\n72.956 -405.5236 l", tickRewritten);
        Assert.Null(DxpFlowPdfVisitor.RewriteRightLegendStream(stream,
            [new[] { "Another Series" }]));

        var previewStream = stream.Replace("1 -0 -0 1 63.7 555.97 cm", "1 0 0 1 63.7 555.97 cm")
            .Replace("0.15 w", ".15 w");
        var previewRewritten = DxpFlowPdfVisitor.RewriteRightLegendStream(previewStream, labels, labels);
        Assert.NotNull(previewRewritten);
        Assert.Contains("cm\n1 0 0 1 7 0 cm\n.15 w", previewRewritten);
        Assert.Contains("248.3 -330 5.5 5.5 re", previewRewritten);
        Assert.Contains("33.7 -418.17 Td\n(Row 1) Tj", previewRewritten);
    }

    [Fact]
    public void FlowVisitorPreservesExplicitCellBorderAfterNilTableBorder()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "file-sample_1MB.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(2);
        Assert.Contains(page.Paths, path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 91 and < 94 && bounds.Width < 1 &&
            bounds.Height > 25 && page.Height - bounds.Top is > 450 and < 700);
        var horizontalBorders = page.Paths.Where(path => path.IsStroked &&
                path.GetBoundingRectangle() is { } bounds && bounds.Width > 250 && bounds.Height < 1)
            .Select(path => page.Height - path.GetBoundingRectangle()!.Value.Top)
            .OrderBy(position => position).ToArray();
        Assert.Equal(7, horizontalBorders.Length);
        Assert.InRange(horizontalBorders[2] - horizontalBorders[1], 31.15, 31.35);
        var following = page.GetWords().First(word => word.Text == "Etiam" &&
            page.Height - word.BoundingBox.Top > horizontalBorders[^1]);
        Assert.InRange(page.Height - following.BoundingBox.Top - horizontalBorders[^1], 28, 28.4);
        using var drawnPdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Contains(drawnPdf.Pages[1].Contents.Elements, element =>
            element is PdfSharp.Pdf.Advanced.PdfReference reference &&
            reference.Value is PdfSharp.Pdf.Advanced.PdfContent content &&
            System.Text.RegularExpressions.Regex.IsMatch(
                System.Text.Encoding.Latin1.GetString(content.Stream.UnfilteredValue),
                @"(?<![\d.])(?:0)?\.25 w"));
    }

    [Fact]
    public void FlowVisitorKeepsItalicDejaVuBulletPitch()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "file-sample_1MB.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var words = pdf.GetPage(1).GetWords().ToArray();
        var italic = words.Single(word => word.Text == "facilisi.");
        var following = words.Single(word => word.Text == "Aenean");
        Assert.InRange(Math.Abs(italic.BoundingBox.Bottom - following.BoundingBox.Bottom), 25.5, 26.0);
    }

    [Fact]
    public void FlowVisitorUsesRegularDejaVuBodyRise()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "file-sample_1MB.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(2);
        var firstLetter = page.Letters.First(letter => letter.Value == "I" &&
            letter.StartBaseLine.X is > 56 and < 57 &&
            page.Height - letter.StartBaseLine.Y is > 190 and < 200);
        Assert.InRange(page.Height - firstLetter.StartBaseLine.Y, 195.23, 195.25);
    }

    [Fact]
    public void FlowVisitorAlignsRaisedDejaVuListMarkersWithText()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "file-sample_1MB.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        foreach (var text in new[] { "Aenean", "Mauris", "Morbi", "Maecenas" })
        {
            var word = page.GetWords().First(word => word.Text == text &&
                page.Height - word.Letters[0].StartBaseLine.Y > 420);
            var marker = page.Letters.First(letter => letter.Value == "•" &&
                letter.StartBaseLine.X < word.Letters[0].StartBaseLine.X &&
                Math.Abs(letter.StartBaseLine.Y - word.Letters[0].StartBaseLine.Y) < 3);
            Assert.InRange(Math.Abs(marker.StartBaseLine.Y - word.Letters[0].StartBaseLine.Y), 0, 0.2);
        }
    }

    [Fact]
    public void FlowVisitorPlacesDirectFitTextFooterAboveBottomMargin()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextLayeredArialCjkAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var footerTop = page.Letters.Where(letter => letter.Value == "F" &&
            letter.GlyphRectangle.Bottom < 100).Max(letter => page.Height - letter.GlyphRectangle.Top);
        Assert.InRange(footerTop, 729, 740);
    }

    [Fact]
    public void FlowVisitorNarrowsAutomaticLatinCjkSpacingAfterFitText()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextLayeredArialCjkAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var right = page.Letters.Where(letter => page.Height - letter.GlyphRectangle.Top is > 90 and < 120)
            .Max(letter => letter.GlyphRectangle.Right);
        Assert.InRange(right, 538, 540);
    }

    [Theory]
    [InlineData("WordRunFitTextLayeredArialCjkAllStories")]
    [InlineData("WordRunFitTextLayeredCalibriCjkAllStories")]
    public void FlowVisitorPositionsMixedFitTextBodyLatinBaseline(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var latin = page.Letters.First(letter => letter.Value == "S" &&
            letter.StartBaseLine.X is > 70 and < 80 &&
            page.Height - letter.StartBaseLine.Y is > 95 and < 120);
        Assert.InRange(page.Height - latin.StartBaseLine.Y, 108.0, 108.3);
    }

    [Fact]
    public void FlowVisitorUsesWordPaintSizeForPureCalibriFittedLines()
    {
        var source = "BT\n/F0 11 Tf\n72 690 Td\n138 Tz\n(Fitted) Tj\n/F0 12 Tf\n0 -24 Td\n100 Tz\n(Next) Tj\nET";
        var revised = DxpFlowPdfVisitor.RewriteFittedCalibriPaintSize(source,
            new HashSet<string> { "F0" });

        Assert.Contains("/F0 12 Tf\n72 690 Td\n126.5 Tz\n(Fitted) Tj", revised);
        Assert.Contains("/F0 12 Tf\n0 -24 Td\n100 Tz\n(Next) Tj", revised);
        Assert.Equal(source, DxpFlowPdfVisitor.RewriteFittedCalibriPaintSize(source,
            new HashSet<string>()));
        Assert.Equal(source.Replace("/F0 12 Tf\n0 -24 Td", "/F1 11 Tf\n0 -24 Td"),
            DxpFlowPdfVisitor.RewriteFittedCalibriPaintSize(
                source.Replace("/F0 12 Tf\n0 -24 Td", "/F1 11 Tf\n0 -24 Td"),
                new HashSet<string> { "F0" }));
    }

    [Theory]
    [InlineData("WordRunFitTextBoldAllStories", 83.42)]
    [InlineData("WordRunFitTextHeaderCjkFallback", 98.42)]
    [InlineData("WordRunFitTextHighAnsiFontsAllStories", 85.06)]
    [InlineData("WordRunFitTextMixedFontsAllStories", 82.32)]
    public void FlowVisitorPositionsFittedBodyForHeaderContext(string name, double expectedBaseline)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var fittedBody = page.GetWords().First(word =>
            word.BoundingBox.Left is > 70 and < 80 &&
            page.Height - word.Letters[0].StartBaseLine.Y is > 75 and < 110);
        Assert.InRange(page.Height - fittedBody.Letters[0].StartBaseLine.Y,
            expectedBaseline - 0.25, expectedBaseline + 0.25);
    }


    [Theory]
    [InlineData("WordRunFitTextMixedCjkAllStories")]
    [InlineData("WordRunFitTextMixedFontsAllStories")]
    [InlineData("WordRunFitTextMixedScriptFontsAllStories")]
    public void FlowVisitorPositionsFittedYuMinchoFooter(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var footer = page.Letters.First(letter => letter.Value == "W" &&
            page.Height - letter.StartBaseLine.Y > 700);
        Assert.InRange(page.Height - footer.StartBaseLine.Y, 739.7, 740.2);
    }

    [Fact]
    public void FlowVisitorKeepsConditionalBorderCellTextAtSourceMargin()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordConditionalColumnBorders.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var labels = page.GetWords().Where(word =>
            word.Text is "A1" or "B1" or "C1" or "A2" or "B2" or "C2" or
                "A3" or "B3" or "C3").ToArray();
        Assert.Equal(9, labels.Length);
        foreach (var label in labels)
        {
            var column = label.Text[0] - 'A';
            Assert.InRange(label.Letters[0].StartBaseLine.X - column * 90, 77.3, 77.5);
        }
    }

    [Theory]
    [InlineData("WordRunFitTextMixedCjkAllStories")]
    [InlineData("WordRunFitTextMixedFontsAllStories")]
    public void FlowVisitorIncludesTerminalAdvanceInSingleYuMinchoStoryFit(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var headerEnd = page.Letters.Where(letter => page.Height - letter.StartBaseLine.Y < 60)
            .MaxBy(letter => letter.StartBaseLine.X)!;
        var footerEnd = page.Letters.Where(letter => page.Height - letter.StartBaseLine.Y > 700)
            .MaxBy(letter => letter.StartBaseLine.X)!;
        Assert.InRange(headerEnd.StartBaseLine.X, 155.2, 156.0);
        Assert.InRange(footerEnd.StartBaseLine.X, 155.1, 155.9);
    }

    [Fact]
    public void FlowVisitorIncludesTerminalAdvanceInHighAnsiStoryFit()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextHighAnsiFontsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var headerEnd = page.Letters.Where(letter => page.Height - letter.StartBaseLine.Y < 60)
            .MaxBy(letter => letter.StartBaseLine.X)!;
        var footerEnd = page.Letters.Where(letter => page.Height - letter.StartBaseLine.Y > 700)
            .MaxBy(letter => letter.StartBaseLine.X)!;
        Assert.InRange(headerEnd.StartBaseLine.X, 158.7, 159.4);
        Assert.InRange(footerEnd.StartBaseLine.X, 158.7, 159.4);
    }

    [Fact]
    public void FlowVisitorSpacesMixedInheritedAndDirectFitTextBeforeCjkBody()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextStyleAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var second = page.Letters.First(letter => letter.Value == "S" &&
            page.Height - letter.StartBaseLine.Y is > 100 and < 130);
        Assert.InRange(page.Height - second.StartBaseLine.Y, 112.4, 112.9);
    }

    [Fact]
    public void FlowVisitorPositionsDirectYuMinchoFittedStories()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextCjkPunctuationFontsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var header = page.Letters.First(letter => letter.Value == "A" &&
            page.Height - letter.StartBaseLine.Y < 70);
        var footer = page.Letters.First(letter => letter.Value == "A" &&
            page.Height - letter.StartBaseLine.Y > 700);
        Assert.InRange(page.Height - header.StartBaseLine.Y, 47.8, 48.3);
        Assert.InRange(page.Height - footer.StartBaseLine.Y, 738.8, 739.7);
    }

    [Theory]
    [InlineData("WordEastAsianAutomaticSpacingAllStories")]
    [InlineData("WordEastAsianLineBreakingAllStories")]
    public void FlowVisitorAlignsUnfittedYuMinchoStoryMargins(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var header = page.Letters.First(letter => page.Height - letter.StartBaseLine.Y < 70);
        var body = page.Letters.First(letter =>
            page.Height - letter.StartBaseLine.Y is > 75 and < 110);
        var footer = page.Letters.First(letter => page.Height - letter.StartBaseLine.Y > 700);
        Assert.InRange(page.Height - header.StartBaseLine.Y, 51.0, 51.5);
        Assert.InRange(page.Height - body.StartBaseLine.Y, 87.2, 87.6);
        Assert.InRange(page.Height - footer.StartBaseLine.Y, 740.0, 740.3);
    }

    [Fact]
    public void FlowVisitorPositionsCenteredAndBottomAlignedSections()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordPageBorderSectionSetup.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        var firstPage = pdf.GetPage(1);
        var first = Assert.Single(firstPage.GetWords().Where(word => word.Text == "First"));
        Assert.InRange(first.BoundingBox.Left, 540, 555);
        Assert.InRange(firstPage.Height - first.BoundingBox.Top, 283, 294);
        var secondPage = pdf.GetPage(2);
        var second = Assert.Single(secondPage.GetWords().Where(word => word.Text == "Second"));
        Assert.InRange(secondPage.Height - second.BoundingBox.Top, 702, 712);
    }

    [Fact]
    public void FlowVisitorPreservesCompoundFontBodyLinePitch()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "file-sample_1MB.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var bodyLineStarts = pdf.GetPage(3).Letters
            .Where(letter => letter.Value is "N" or "P" &&
                letter.GlyphRectangle.Left is > 55 and < 60)
            .OrderByDescending(letter => letter.GlyphRectangle.Bottom).ToArray();
        Assert.True(bodyLineStarts.Length >= 2);
        Assert.Contains(Enumerable.Range(0, bodyLineStarts.Length - 1), index =>
            bodyLineStarts[index].GlyphRectangle.Bottom - bodyLineStarts[index + 1].GlyphRectangle.Bottom
                is >= 16.5 and <= 17.0);
    }

    [Theory]
    [InlineData("WordCharacterGridAdjustRightAllStories")]
    [InlineData("WordInheritedTextAlignmentAllStories")]
    public void FlowVisitorPreservesGridParagraphSpacing(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var secondParagraph = pdf.GetPage(1).Letters.Single(letter => letter.Value == "W" &&
            letter.GlyphRectangle.Left is > 200 and < 215 &&
            letter.GlyphRectangle.Bottom is > 580 and < 630);
        var thirdParagraph = pdf.GetPage(1).Letters.Single(letter => letter.Value == "W" &&
            letter.GlyphRectangle.Left is > 220 and < 235 &&
            letter.GlyphRectangle.Bottom is > 480 and < 540);
        Assert.InRange(secondParagraph.GlyphRectangle.Bottom - thirdParagraph.GlyphRectangle.Bottom,
            91.0, 93.0);
    }

    [Theory]
    [InlineData("WordCharacterGridAdjustRightAllStories", 27.75)]
    [InlineData("WordInheritedTextAlignmentAllStories", 27.84)]
    public void FlowVisitorUsesGridPitchForTextAlignment(string name, double expectedPitch)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var baselines = page.Letters.Where(letter => letter.Value == "W" &&
                page.Height - letter.StartBaseLine.Y is > 75 and < 165)
            .Select(letter => page.Height - letter.StartBaseLine.Y).OrderBy(top => top).ToArray();
        Assert.True(baselines.Length >= 2, string.Join(", ", baselines));
        Assert.InRange(baselines[1] - baselines[0], expectedPitch - 0.15, expectedPitch + 0.15);
    }

    [Theory]
    [InlineData("WordSnapToCharacterGridAllStories")]
    [InlineData("WordInheritedRunGridSnapAllStories")]
    public void FlowVisitorPlacesSnappedGridHeaderAtWordBaseline(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var header = page.Letters.First(letter => letter.StartBaseLine.X is > 71 and < 73 &&
            page.Height - letter.StartBaseLine.Y is > 45 and < 55);
        Assert.InRange(page.Height - header.StartBaseLine.Y, 51.2, 51.5);
    }

    [Theory]
    [InlineData("WordSnapToCharacterGridAllStories")]
    [InlineData("WordInheritedRunGridSnapAllStories")]
    public void FlowVisitorAlignsSnappedParagraphAfterUnsnappedGridParagraph(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var snappedBaseline = page.Letters.Where(letter => letter.StartBaseLine.X is > 76 and < 78 &&
                page.Height - letter.StartBaseLine.Y is > 175 and < 190)
            .Select(letter => page.Height - letter.StartBaseLine.Y).Distinct().ToArray();
        Assert.Single(snappedBaseline);
        Assert.InRange(snappedBaseline[0], 181.2, 181.6);
    }

    [Fact]
    public void FlowVisitorAlignsInheritedEmphasisAfterSnappedGridParagraph()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedCharacterEmphasisGridAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var commands = string.Concat(pdf.Pages[0].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Select(reference => reference.Value)
            .OfType<PdfSharp.Pdf.Advanced.PdfContent>()
            .Where(stream => stream.Stream != null)
            .Select(stream => System.Text.Encoding.Latin1.GetString(stream.Stream!.UnfilteredValue)));
        var paragraphTransitions = System.Text.RegularExpressions.Regex.Matches(commands,
                @"(?m)^-(?<horizontal>\d+(?:\.\d+)?) -(?<vertical>\d+(?:\.\d+)?) Td\r?$")
            .Select(match => (
                Horizontal: double.Parse(match.Groups["horizontal"].Value,
                    System.Globalization.CultureInfo.InvariantCulture),
                Vertical: double.Parse(match.Groups["vertical"].Value,
                    System.Globalization.CultureInfo.InvariantCulture)))
            .Where(position => position.Horizontal is > 350 and < 385 &&
                position.Vertical is > 34 and < 36).ToArray();
        Assert.Single(paragraphTransitions);
        Assert.InRange(paragraphTransitions[0].Vertical, 34.9, 35.0);
    }

    [Fact]
    public void FlowVisitorUsesCharacterGridHeaderDistance()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordCharacterGridAdjustRightAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var headerWord = page.Letters.First(letter => letter.Value == "W" &&
            page.Height - letter.StartBaseLine.Y < 70);
        Assert.InRange(page.Height - headerWord.StartBaseLine.Y, 51.0, 52.0);
    }

    [Theory]
    [InlineData("WordCharacterGridAdjustRightAllStories")]
    [InlineData("WordInheritedAdjustRightGridAllStories")]
    public void FlowVisitorOmitsGridGlyphRiseWithoutTextAlignment(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var commands = string.Join('\n', pdf.Pages[0].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Select(reference => reference.Value)
            .OfType<PdfSharp.Pdf.Advanced.PdfContent>()
            .Select(content => System.Text.Encoding.Latin1.GetString(content.Stream!.UnfilteredValue)));
        Assert.DoesNotContain("0.5 Ts", commands);
    }

    [Fact]
    public void FlowVisitorKeepsStyledBlankMarginLinesAtTheirInheritedHeight()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordDistributedStyleInheritanceAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var headerLast = page.Letters.Single(letter => letter.Value == "H" &&
            letter.GlyphRectangle.Left is > 71 and < 74 &&
            page.Height - letter.GlyphRectangle.Top is > 55 and < 85);
        var footerFirst = page.Letters.Single(letter => letter.Value == "F" &&
            letter.GlyphRectangle.Left is > 71 and < 74 &&
            page.Height - letter.GlyphRectangle.Top is > 690 and < 720);
        var footerLast = page.Letters.Single(letter => letter.Value == "F" &&
            letter.GlyphRectangle.Left is > 71 and < 74 &&
            page.Height - letter.GlyphRectangle.Top is > 720 and < 750);
        Assert.InRange(page.Height - headerLast.GlyphRectangle.Top, 75, 79);
        Assert.InRange(page.Height - footerLast.GlyphRectangle.Top, 741, 745);
        Assert.InRange(page.Height - footerFirst.StartBaseLine.Y, 714.7, 715.3);
        Assert.InRange(page.Height - footerLast.StartBaseLine.Y, 753.3, 753.8);
    }

    [Fact]
    public void FlowVisitorUsesModernDefaultForDocumentWithoutSectionProperties()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        using var modern = WordprocessingDocument.Open(
            Path.Combine(root, "samples", "sample-no-sectPr.docx"), false);
        using var legacy = WordprocessingDocument.Open(Path.Combine(root, "DocxportNet.Tests",
            "Fixtures", "Doc", "WordCharacterGridAdjustRightAllStories.docx"), false);
        using var styled = WordprocessingDocument.Open(Path.Combine(root, "DocxportNet.Tests",
            "Fixtures", "PdfProbeCorpus", "vml-gif-probe__vml-gif-with-styles.docx"), false);

        Assert.True(DxpFlowPdfVisitor.HasModernUnthemedDefaults(modern));
        Assert.True(DxpFlowPdfVisitor.HasModernUnthemedDefaults(styled));
        Assert.False(DxpFlowPdfVisitor.HasModernUnthemedDefaults(legacy));
    }

    [Fact]
    public void GridSpacingDoesNotTreatClosingCjkPunctuationAsLatin()
    {
        var snapGrid = DxpFlowPdfVisitor.AddGridBreakOpportunities("日。本", false, true);
        var linesAndChars = DxpFlowPdfVisitor.AddGridBreakOpportunities("日。本", false, false);

        Assert.Equal("日\u2003\u200b。\u2003\u200b本", snapGrid);
        Assert.Equal("日\u2002\u2002\u200b。\u2003\u200b本", linesAndChars);
    }

    [Fact]
    public void FootnoteUrlBreaksAtHyphensWithoutSplittingPathSegments()
    {
        var url = "https://example.org/documents/2024/10/28/2024-24563/repair-restrictions/report.pdf";

        var result = DxpFlowPdfVisitor.AddFootnoteUrlBreakOpportunities(url);

        Assert.Equal("https://example.org/documents/2024/10/28/2024-\u200b24563/repair-\u200brestrictions/report.pdf",
            result);
    }

    [Fact]
    public void FlowVisitorCarriesKerningIntoFollowingWordPosition()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordWidowVisibleOn.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(2);
        var firstLine = page.GetWords().Where(word =>
            page.Height - word.BoundingBox.Top is > 48 and < 55).ToArray();
        Assert.InRange(firstLine.Single(word => word.Text == "line").BoundingBox.Left, 76, 77.5);
        Assert.InRange(firstLine.Single(word => word.Text == "one").BoundingBox.Left, 98, 99.5);
    }

    [Fact]
    public void FlowVisitorCarriesSmallArialKerningIntoFollowingWord()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            "line-spacing-probe__isolated__LineSpacing-exact-20.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var glyphs = pdf.GetPage(1).GetWords().Single(word => word.Text == "glyphs");
        Assert.InRange(glyphs.BoundingBox.Left, 130, 131);
    }

    [Fact]
    public void FlowVisitorAdjustsOnlyTheMatchingKernedUnderline()
    {
        var lines = new[] {
            "ET", "0.563 w", "0.275 0.471 0.525 RG",
            "72 205.5348 m", "352.1909 205.5348 l", "S",
            "72 220 m", "352.1909 220 l", "S", "BT"
        };

        Assert.True(DxpFlowPdfVisitor.AdjustKernedUnderline(lines, 1, 72, 206.37, -2.36));
        Assert.Equal("349.8309 205.5348 l", lines[4]);
        Assert.Equal("352.1909 220 l", lines[7]);
        Assert.False(DxpFlowPdfVisitor.AdjustKernedUnderline(lines, 1, 90, 206.37, -2.36));
    }

    [Fact]
    public void FlowVisitorKernsOnlyContiguousUrlFragments()
    {
        var lines = new[] { "188.9605 0 Td", "192.1 0 Td", "bad" };

        Assert.True(DxpFlowPdfVisitor.TryAdjustKernedFragmentMove(lines, 0, 188.96, -0.68));
        Assert.Equal("188.2805 0 Td", lines[0]);
        Assert.False(DxpFlowPdfVisitor.TryAdjustKernedFragmentMove(lines, 1, 188.96, -0.68));
        Assert.False(DxpFlowPdfVisitor.TryAdjustKernedFragmentMove(lines, 2, 188.96, -0.68));
    }

    [Fact]
    public void FlowVisitorKernsWrappedDomainPathUrl()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Right to Repair.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var canada = pdf.GetPage(3).GetWords().Single(word =>
            word.Text == "Canada" && word.BoundingBox.Bottom > 105 &&
            word.BoundingBox.Bottom < 120);
        Assert.InRange(canada.BoundingBox.Left, 314.0, 314.3);
    }

    [Fact]
    public void FlowVisitorKernsDottedInitialsInCitation()
    {
        Assert.True(DxpFlowPdfVisitor.IsKerningCandidate("P.C."));
        Assert.True(DxpFlowPdfVisitor.IsKerningCandidate("S.C."));
        Assert.True(DxpFlowPdfVisitor.IsKerningCandidate("Word123"));
        Assert.True(DxpFlowPdfVisitor.IsKerningCandidate("FOOTER-"));
        Assert.False(DxpFlowPdfVisitor.IsKerningCandidate("Second-"));
        Assert.True(DxpFlowPdfVisitor.IsKerningCandidate("This\u00a0table\u00a0cell"));
        Assert.True(DxpFlowPdfVisitor.IsKerningCandidate("\"BAD\""));
        Assert.True(DxpFlowPdfVisitor.IsKerningCandidate("\"App*\""));
        Assert.True(DxpFlowPdfVisitor.IsKerningCandidate("\"Value"));
        Assert.True(DxpFlowPdfVisitor.IsKerningCandidate("yyyy-"));
        Assert.False(DxpFlowPdfVisitor.IsKerningCandidate("This table cell"));
        Assert.False(DxpFlowPdfVisitor.IsKerningCandidate("P."));
        Assert.False(DxpFlowPdfVisitor.IsKerningCandidate("2012-1392"));
        Assert.True(DxpFlowPdfVisitor.CarriesScaledKerning("Calibri"));
        Assert.True(DxpFlowPdfVisitor.CarriesScaledKerning("Arial Black"));
        Assert.False(DxpFlowPdfVisitor.CarriesScaledKerning("Aptos"));
        Assert.False(DxpFlowPdfVisitor.CarriesScaledKerning("Calibri", bold: true, italic: true));
        Assert.True(DxpFlowPdfVisitor.IsWinAnsiClosingKerningCandidate("authors.\u0094"));
        Assert.True(DxpFlowPdfVisitor.IsWinAnsiClosingKerningCandidate("authors?\u0092"));
        Assert.False(DxpFlowPdfVisitor.IsWinAnsiClosingKerningCandidate("authors.\u0093"));
        Assert.Equal('”', DxpFlowPdfVisitor.NormalizeWinAnsiQuote('\u0094'));
        Assert.Equal('’', DxpFlowPdfVisitor.NormalizeWinAnsiQuote('\u0092'));
        Assert.True(DxpFlowPdfVisitor.ShouldSuppressRepeatedPairExpansion("xxxxxxxxx", 0.1));
        Assert.False(DxpFlowPdfVisitor.ShouldSuppressRepeatedPairExpansion("xxxy", 0.1));
        Assert.False(DxpFlowPdfVisitor.ShouldSuppressRepeatedPairExpansion("xxxx", -0.1));
        Assert.False(DxpFlowPdfVisitor.ShouldSuppressRepeatedPairExpansion("1111", 0.1));
        Assert.False(DxpFlowPdfVisitor.ShouldSuppressRepeatedPairExpansion("漢漢漢漢", 0.1));
    }

    [Fact]
    public void FlowVisitorKernsNoBreakRunWithoutMovingFollowingText()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var document = new PdfSharp.Pdf.PdfDocument();
        var page = document.AddPage();
        using (var graphics = PdfSharp.Drawing.XGraphics.FromPdfPage(page))
        {
            var font = new PdfSharp.Drawing.XFont("Arial", 12);
            graphics.DrawString("To\u00a0AV\u00a0World", font, PdfSharp.Drawing.XBrushes.Black,
                new PdfSharp.Drawing.XPoint(72, 72));
            graphics.DrawString("Next", font, PdfSharp.Drawing.XBrushes.Black,
                new PdfSharp.Drawing.XPoint(300, 72));
        }
        using var stream = new MemoryStream();
        document.Save(stream, false);
        using var modified = PdfReader.Open(new MemoryStream(stream.ToArray()), PdfDocumentOpenMode.Modify);
        typeof(DxpFlowPdfVisitor).GetMethod("ApplyFontKerning",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
            .Invoke(new DxpFlowPdfVisitor(), [modified]);
        using var revised = new MemoryStream();
        modified.Save(revised, false);
        using var before = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(stream.ToArray()));
        using var after = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(revised.ToArray()));
        var originalNext = before.GetPage(1).GetWords().Single(word => word.Text == "Next");
        var revisedNext = after.GetPage(1).GetWords().Single(word => word.Text == "Next");
        Assert.InRange(Math.Abs(originalNext.BoundingBox.Left - revisedNext.BoundingBox.Left), 0, 0.01);
        var commands = string.Concat(modified.Pages[0].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Select(reference => reference.Value)
            .OfType<PdfSharp.Pdf.Advanced.PdfContent>()
            .Select(content => System.Text.Encoding.Latin1.GetString(content.Stream!.UnfilteredValue)));
        Assert.Contains(" TJ", commands);
    }

    [Fact]
    public void FlowKerningPositionsLettersBeforeDigits()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var document = new MigraDoc.DocumentObjectModel.Document();
        var paragraph = document.AddSection().AddParagraph("Word123 next");
        paragraph.Format.Font.Name = "Arial";
        paragraph.Format.Font.Size = MigraDoc.DocumentObjectModel.Unit.FromPoint(12);
        var renderer = new MigraDoc.Rendering.PdfDocumentRenderer { Document = document };
        renderer.RenderDocument();
        using var original = new MemoryStream();
        renderer.PdfDocument.Save(original, false);
        using var before = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(original.ToArray()));
        using var modified = PdfSharp.Pdf.IO.PdfReader.Open(new MemoryStream(original.ToArray()),
            PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify);
        typeof(DxpFlowPdfVisitor).GetMethod("ApplyFontKerning",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
            .Invoke(new DxpFlowPdfVisitor(), [modified]);
        using var adjusted = new MemoryStream();
        modified.Save(adjusted, false);
        using var after = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(adjusted.ToArray()));
        var firstBefore = before.GetPage(1).GetWords().Single(word => word.Text == "Word123");
        var firstAfter = after.GetPage(1).GetWords().Single(word => word.Text == "Word123");
        var nextBefore = before.GetPage(1).GetWords().Single(word => word.Text == "next");
        var nextAfter = after.GetPage(1).GetWords().Single(word => word.Text == "next");
        Assert.InRange(Math.Abs(firstAfter.BoundingBox.Left - firstBefore.BoundingBox.Left), 0, 0.05);
        Assert.InRange(firstBefore.BoundingBox.Right - firstAfter.BoundingBox.Right, 0.1, 1);
        Assert.InRange(nextBefore.BoundingBox.Left - nextAfter.BoundingBox.Left, 0.1, 1);
    }

    [Fact]
    public void LinesAndCharsGridKeepsLatinWordsTogether()
    {
        var content = DxpFlowPdfVisitor.AddGridBreakOpportunities("漢字Word456漢字", false, false);

        Assert.True(content.Contains("W\u2003o\u2003r\u2003d\u20034\u20035\u20036",
            StringComparison.Ordinal));
        Assert.False(content.Contains("W\u2003\u200bo", StringComparison.Ordinal));
        Assert.True(content.Contains("漢\u2003\u200b字", StringComparison.Ordinal));
        Assert.Contains("字\u2003\u2002\u200bW", content, StringComparison.Ordinal);
    }

    [Fact]
    public void OversizedGridTextUsesHalfEmCharacterSpacing()
    {
        var content = DxpFlowPdfVisitor.AddGridBreakOpportunities("LARGE", false, false, true);
        var leading = DxpFlowPdfVisitor.AddGridBreakOpportunities(" LARGE", false, false, true);

        Assert.Equal("L\u2002A\u2002R\u2002G\u2002E", content);
        Assert.Equal(" \u2003\u200bL\u2002A\u2002R\u2002G\u2002E", leading);
    }

    [Fact]
    public void FlowVisitorTopAlignsInheritedSizeWithLargeRun()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedTextAlignmentAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Contains(pdf.Pages[0].Contents.Elements, element =>
            (element as PdfSharp.Pdf.Advanced.PdfReference)?.Value is PdfSharp.Pdf.Advanced.PdfContent content &&
            content.Stream != null &&
            System.Text.Encoding.Latin1.GetString(content.Stream.UnfilteredValue).Contains("6.3 Ts"));
        Assert.Contains(pdf.Pages[0].Contents.Elements, element =>
            (element as PdfSharp.Pdf.Advanced.PdfReference)?.Value is PdfSharp.Pdf.Advanced.PdfContent content &&
            content.Stream != null &&
            System.Text.RegularExpressions.Regex.IsMatch(
                System.Text.Encoding.Latin1.GetString(content.Stream.UnfilteredValue),
                @"21 0 Td\s+/\S+ 14 Tf\s+128\.5714 Tz\s+0\.3889 Tc\s+0 Ts\s+<[0-9A-F]+> Tj"));
        Assert.Equal(2, pdf.Pages[0].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Select(reference => reference.Value)
            .OfType<PdfSharp.Pdf.Advanced.PdfContent>()
            .Where(content => content.Stream != null)
            .Sum(content => System.Text.RegularExpressions.Regex.Matches(
                System.Text.Encoding.Latin1.GetString(content.Stream!.UnfilteredValue),
                @"/\S+ 18 Tf\s+100 Tz\s+0 Tc\s+0 Ts\s+\(LARGE\) Tj").Count));
    }

    [Fact]
    public void FlowVisitorRaisesMixedSizeTopAlignedHeaderWithoutMovingFooter()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedTextAlignmentAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var commands = string.Concat(pdf.Pages[0].Contents.Elements.Select(element =>
            (element as PdfSharp.Pdf.Advanced.PdfReference)?.Value is
                PdfSharp.Pdf.Advanced.PdfContent content && content.Stream != null
                ? System.Text.Encoding.Latin1.GetString(content.Stream.UnfilteredValue) : ""));

        Assert.Matches(@"6\.86 Ts[\s\S]*6\.3 Ts", commands);
    }



    [Fact]
    public void FlowVisitorAppliesLinesOnlyGridPitchToBodyParagraphs()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordSectionLineGridSnapAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var bodyStarts = page.Letters.Where(letter => letter.Value is "F" or "S" or "T" &&
            page.Height - letter.GlyphRectangle.Top is > 60 and < 170)
            .Select(letter => page.Height - letter.GlyphRectangle.Top).OrderBy(top => top).ToArray();
        Assert.Contains(bodyStarts, top => top is > 70 and < 80);
        Assert.Contains(bodyStarts, top => top is > 102 and < 112);
        Assert.Contains(bodyStarts, top => top is > 138 and < 148);
        Assert.InRange(bodyStarts[1] - bodyStarts[0], 31.3, 31.8);
        var footer = page.GetWords().Single(word => word.Text == "Grid" &&
            page.Height - word.BoundingBox.Top > 700);
        Assert.InRange(page.Height - footer.BoundingBox.Top, 734, 736);
    }

    [Fact]
    public void FlowVisitorPlacesLinesAndCharactersGridFooter()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordLineCharacterGridAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var footerBaseline = page.Letters.Where(letter =>
                page.Height - letter.StartBaseLine.Y > 700)
            .Max(letter => page.Height - letter.StartBaseLine.Y);
        Assert.InRange(footerBaseline, 739.9, 740.5);
        var gridText = DxpFlowPdfVisitor.AddGridBreakOpportunities("す。漢", true, false);
        Assert.Equal("す\u00a0\u00a0\u00a0\u200b。\u00a0\u00a0\u00a0\u200b漢", gridText);
    }

    [Fact]
    public void FlowVisitorDistributesSingleLineParagraphWords()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordDistributedParagraphAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = pdf.GetPage(1).Letters
            .Where(letter => letter.GlyphRectangle.Bottom is > 680 and < 710).ToArray();
        Assert.Contains(letters, letter => letter.Value == "F" &&
            letter.GlyphRectangle.Left is > 71 and < 73);
        Assert.Contains(letters, letter => letter.Value == "b" &&
            letter.GlyphRectangle.Left is > 275 and < 285);
        Assert.Contains(letters, letter => letter.Value == "p" &&
            letter.GlyphRectangle.Left is > 485 and < 495);
        var storyWords = pdf.GetPage(1).GetWords().Where(word => word.Text == "first").ToArray();
        Assert.Contains(storyWords, word => word.BoundingBox.Left is > 515 and < 525 &&
            word.BoundingBox.Bottom > 700);
        Assert.Contains(storyWords, word => word.BoundingBox.Left is > 515 and < 525 &&
            word.BoundingBox.Bottom < 100);
    }

    [Fact]
    public void FlowVisitorDrawsSpacedTableOuterBorderOutsideCellGap()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyledCellSpacingResetAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var borders = pdf.GetPage(1).Paths.Where(path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 72 and < 73 && bounds.Height > 20).ToArray();
        Assert.Equal(3, borders.Length);
        Assert.Equal(2, borders.Count(path => path.GetBoundingRectangle() is
            { Height: > 54 and < 56 }));
        Assert.Single(borders, path => path.GetBoundingRectangle() is
            { Height: > 65 and < 67 });
        Assert.True(pdf.GetPage(1).Paths.Count(path => path.IsFilled &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Width is > 230 and < 240 && bounds.Height is > 23 and < 25) >= 6);
        var headerWords = pdf.GetPage(1).GetWords().Where(word => word.Text == "Header")
            .ToArray();
        Assert.Contains(headerWords, word => word.BoundingBox.Left is > 90.8 and < 91.0 &&
            pdf.GetPage(1).Height - word.BoundingBox.Top is > 60 and < 80);
        Assert.Contains(headerWords, word => word.BoundingBox.Left is > 78.3 and < 78.5 &&
            pdf.GetPage(1).Height - word.BoundingBox.Top is > 90 and < 110);
    }

    [Fact]
    public void FlowVisitorKeepsSeparateSpacedCellBorders()
    {
        if (typeof(MigraDoc.Rendering.DocumentRenderer).GetProperty("WrapContour") == null) return;
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordCellSpacingAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var strokes = page.Paths.Where(path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Bottom is > 600 and < 650 &&
            bounds.Width is > 257 and < 271 && bounds.Height < 0.1).ToArray();
        Assert.Equal(4, strokes.Length);
    }

    [Fact]
    public void FlowVisitorAppliesRowTablePropertyExceptionTopBorders()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRowBorderExceptionsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.GetPage(1).Paths.Count(path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Width > 200 && bounds.Height < 2) >= 8);
        Assert.True(pdf.GetPage(1).Paths.Count(path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Width > 460 && bounds.Height < 2) >= 2);
        var letters = pdf.GetPage(1).Letters;
        var intro = letters.Single(letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left is > 71 and < 73);
        var firstCell = letters.Where(letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left is > 77 and < 79)
            .MaxBy(letter => letter.GlyphRectangle.Bottom)!;
        Assert.InRange(intro.GlyphRectangle.Bottom - firstCell.GlyphRectangle.Bottom, 26.5, 27.5);
    }

    [Theory]
    [InlineData("WordCellAlignmentAllStories")]
    [InlineData("WordCellDecorationsAllStories")]
    [InlineData("WordNestedTableBody")]
    [InlineData("WordRowCellMarginsAllStories")]
    public void FlowVisitorClearsPlainHeaderTableBeforeBody(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = pdf.GetPage(1).Letters;
        var header = letters.Single(letter => letter.Value == "H" &&
            letter.GlyphRectangle.Left is > 71 and < 73);
        var body = letters.Single(letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left is > 71 and < 73);
        Assert.InRange(header.GlyphRectangle.Bottom - body.GlyphRectangle.Bottom, 44, 45);
    }

    [Theory]
    [InlineData("WordConditionalBandBordersThreeRowsAllStories", 75.1, 75.4)]
    [InlineData("WordConditionalBandInheritedAllStories", 75.1, 75.4)]
    [InlineData("WordConditionalBandInheritedDirectAllStories", 75.1, 75.4)]
    [InlineData("WordVisibleAutoRowGridGapsAllStories", 73.1, 73.4)]
    public void FlowVisitorKeepsMultirowHeaderClearanceWithGridGapBody(
        string name, double minimum, double maximum)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var words = pdf.GetPage(1).GetWords();
        var header = words.First(word => word.Text == "Header" &&
            word.BoundingBox.Left is > 71 and < 73);
        var body = words.First(word => word.Text == "Body" &&
            word.BoundingBox.Left is > 71 and < 73);
        Assert.InRange(header.BoundingBox.Top - body.BoundingBox.Top, minimum, maximum);
    }

    [Theory]
    [InlineData("WordTablesAllStories")]
    [InlineData("WordCellFitNoWrapAllStories")]
    public void FlowVisitorClearsZeroMarginHeaderTableBeforeBody(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var body = page.Letters.First(letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left is > 71 and < 73);
        Assert.InRange(page.Height - body.GlyphRectangle.Top, 82, 84);
    }

    [Theory]
    [InlineData("WordNestedTablesAllStories")]
    [InlineData("WordNestedTableListsAllStories")]
    public void FlowVisitorClearsSingleNestedHeaderRowBeforeBody(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = pdf.GetPage(1).Letters;
        var body = letters.Single(letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left is > 71 and < 73);
        var bodyTop = pdf.GetPage(1).Height - body.GlyphRectangle.Top;
        Assert.InRange(bodyTop, 96.5, 98);
    }

    [Theory]
    [InlineData("WordTableBackgroundAutoDirectAllStories")]
    [InlineData("WordTableBackgroundAutoStyleAllStories")]
    public void FlowVisitorClearsFilledSpacedHeaderTable(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "DocNativeOracle",
            name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = pdf.GetPage(1).Letters;
        var header = letters.Single(letter => letter.Value == "H" &&
            letter.GlyphRectangle.Left is > 83 and < 85);
        var body = letters.Single(letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left is > 83 and < 85);
        Assert.InRange(header.GlyphRectangle.Bottom - body.GlyphRectangle.Bottom, 48, 50);
        var secondHeader = letters.Single(letter => letter.Value == "H" &&
            letter.GlyphRectangle.Left > 130);
        var secondBody = letters.Single(letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left is > 130 and < 140 &&
            pdf.GetPage(1).Height - letter.GlyphRectangle.Top is > 90 and < 200);
        Assert.InRange(secondHeader.GlyphRectangle.Left, 144.6, 145.1);
        Assert.InRange(secondBody.GlyphRectangle.Left, 133.5, 134);
    }

    [Fact]
    public void FlowVisitorKeepsColoredSpacedTableRowBottomNearWord()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "DocNativeOracle",
            "WordTableBackgroundAutoStyleAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.Paths, path => path.IsFilled && path.FillColor != null &&
            path.FillColor.ToRGBValues() is var rgb &&
            rgb.r > 0.99 && rgb.g > 0.99 && rgb.b > 0.99 &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 83 and < 85 && bounds.Right is > 131 and < 133 &&
            page.Height - bounds.Top is > 47 and < 49 && bounds.Height is > 22.5 and < 24);
    }

    [Theory]
    [InlineData("WordWidowVisibleOn", false)]
    [InlineData("WordWidowVisibleOff", true)]
    public void FlowVisitorRespectsWidowControlForExplicitLineBreaks(string name, bool firstLineOnFirstPage)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Equal(firstLineOnFirstPage, pdf.GetPage(1).Text.Contains("Targetlineone"));
        Assert.Contains("Targetlinethree", pdf.GetPage(2).Text);
    }

    [Theory]
    [InlineData("WordGrowAutofitNoWrapAllStories", "Doc")]
    [InlineData("WordVisibleNoWrapAllStories", "Doc")]
    [InlineData("WordVisibleNoWrapOffAllStories", "Doc")]
    [InlineData("WordVisibleNoWrapOffArialAllStories", "Doc")]
    [InlineData("WordVisibleNoWrapArialAllStories", "DocKnownGaps")]
    public void FlowVisitorClearsImbalancedAutoFitHeaderTable(string name, string folder)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", folder, name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = pdf.GetPage(1).Letters;
        var header = letters.Where(letter => letter.Value == "H" &&
            letter.GlyphRectangle.Left is > 71 and < 73)
            .OrderByDescending(letter => letter.GlyphRectangle.Bottom).First();
        var body = letters.Where(letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left is > 71 and < 73)
            .OrderByDescending(letter => letter.GlyphRectangle.Bottom).First();
        Assert.InRange(header.GlyphRectangle.Bottom - body.GlyphRectangle.Bottom, 53, 64);
        if (name is "WordGrowAutofitNoWrapAllStories" or "WordVisibleNoWrapAllStories" or
            "WordVisibleNoWrapOffAllStories")
            Assert.InRange(pdf.GetPage(1).Height - body.GlyphRectangle.Top, 97.0, 97.5);
        if (name == "WordGrowAutofitNoWrapAllStories")
            Assert.True(letters.Count(letter => letter.Value == "T" &&
                letter.GlyphRectangle.Left is > 71 and < 73) >= 3);
        if (name.EndsWith("ArialAllStories", StringComparison.Ordinal))
            Assert.InRange(header.GlyphRectangle.Bottom - body.GlyphRectangle.Bottom, 54, 57);
    }

    [Theory]
    [InlineData("WordVisibleNoWrapOffArialAllStories", "Doc", 664.5, 665.3)]
    [InlineData("WordVisibleNoWrapArialAllStories", "DocKnownGaps", 664.5, 665.3)]
    [InlineData("WordVisibleNoWrapOffAllStories", "Doc", 663.7, 664.2)]
    public void FlowVisitorPositionsOneRowFooterTableForThemeFont(
        string name, string folder, double minimum, double maximum)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", folder, name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var words = pdf.GetPage(1).GetWords();
        var header = words.First(word => word.Text == "Header" && word.BoundingBox.Left is > 71 and < 73);
        var footer = words.First(word => word.Text == "Footer" && word.BoundingBox.Left is > 71 and < 73);
        Assert.InRange(header.BoundingBox.Top - footer.BoundingBox.Top, minimum, maximum);
    }

    [Fact]
    public void FlowVisitorBreaksOversizedWordInNarrowNoWrapCompanionCell()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleNoWrapAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var headerWords = pdf.GetPage(1).GetWords().Where(word =>
            word.BoundingBox.Left > 490 &&
            pdf.GetPage(1).Height - word.BoundingBox.Top is > 45 and < 85).ToArray();
        Assert.Contains(headerWords, word => word.Text is "Head" or "Heade");
        Assert.Contains(headerWords, word => word.Text is "er" or "r");
        var headerStart = headerWords.First(word => word.Text is "Head" or "Heade").BoundingBox.Left;
        var footerStart = pdf.GetPage(1).GetWords().First(word => word.Text is "Foot" or "Foote" &&
            word.BoundingBox.Left > 490).BoundingBox.Left;
        Assert.InRange(footerStart - headerStart, 3, 6);
    }

    [Fact]
    public void FlowVisitorDoesNotBalanceGrowAutoFitCompanionCell()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordGrowAutofitNoWrapAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var headerWords = pdf.GetPage(1).GetWords().Where(word =>
            word.BoundingBox.Left > 520 &&
            pdf.GetPage(1).Height - word.BoundingBox.Top is > 45 and < 85).ToArray();
        Assert.Contains(headerWords, word => word.Text == "Header");
        Assert.Contains(headerWords, word => word.Text == "B");
    }

    [Fact]
    public void FlowVisitorReservesNestedTableStoryHeightAndIndent()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordDeepNestedTablesAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = pdf.GetPage(1).Letters;
        var header = letters.First(letter => letter.Value == "H" &&
            letter.GlyphRectangle.Left is > 71 and < 73);
        var body = letters.First(letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left is > 71 and < 73);
        var nested = letters.First(letter => letter.Value == "N" &&
            letter.GlyphRectangle.Left is > 82 and < 84);
        Assert.InRange(header.GlyphRectangle.Bottom - body.GlyphRectangle.Bottom, 70, 76);
        Assert.InRange(nested.GlyphRectangle.Left - header.GlyphRectangle.Left, 10, 12);
    }

    [Fact]
    public void FlowVisitorClipsNestedCellsOutsideParentWidth()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordNestedMergedCell.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var text = pdf.GetPage(1).Text.Replace(" ", "");
        Assert.Contains("Body12", text);
        Assert.DoesNotContain("Body22", text);
        Assert.DoesNotContain("Header12", text);
        Assert.DoesNotContain("Footer22", text);
    }

    [Fact]
    public void FlowVisitorMeasuresImbalancedAutoFitColumnWidths()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleNoWrapOffArialAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = pdf.GetPage(1).Letters;
        Assert.InRange(letters.First(letter => letter.Value == "H" &&
            letter.GlyphRectangle.Left > 400).GlyphRectangle.Left, 486.5, 488);
        Assert.InRange(letters.Where(letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left > 400).OrderByDescending(letter =>
            letter.GlyphRectangle.Bottom).Skip(1).First().GlyphRectangle.Left, 497.5, 498.3);
        Assert.InRange(letters.First(letter => letter.Value == "F" &&
            letter.GlyphRectangle.Left > 400).GlyphRectangle.Left, 490.5, 491.5);
    }

    [Fact]
    public void FlowVisitorWrapsNarrowAutoFitFooterCellWithoutNoWrapSetting()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleNoWrapOffAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = pdf.GetPage(1).Letters;
        var footer = letters.Single(letter => letter.Value == "F" &&
            letter.GlyphRectangle.Left > 400);
        var nextLine = letters.Where(letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left > 400).MinBy(letter => letter.GlyphRectangle.Bottom)!;
        Assert.InRange(nextLine.GlyphRectangle.Left, 492, 494);
        Assert.InRange(footer.GlyphRectangle.Bottom - nextLine.GlyphRectangle.Bottom, 13.5, 14.2);
        var page = pdf.GetPage(1);
        var body = page.GetWords().Single(word => word.Text == "Body" &&
            word.BoundingBox.Left > 490 &&
            page.Height - word.BoundingBox.Top is > 155 and < 175);
        var lastLetter = page.GetWords().Single(word => word.Text == "D" &&
            word.BoundingBox.Left > 490 &&
            page.Height - word.BoundingBox.Top is > 175 and < 195);
        Assert.InRange(body.BoundingBox.Top - lastLetter.BoundingBox.Top, 15.8, 17.5);
    }

    [Theory]
    [InlineData("WordLeftPositionalTabAllStories", 42)]
    [InlineData("WordLeftPositionalTabBoundariesAllStories", 109)]
    public void FlowVisitorPlacesLeftMarginPositionalTabs(string name, int bodyOffset)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = pdf.GetPage(1).Letters;
        var header = letters.Where(letter => letter.Value == "H" &&
            letter.GlyphRectangle.Left is > 71 and < 73 &&
            letter.GlyphRectangle.Bottom > 700)
            .OrderByDescending(letter => letter.GlyphRectangle.Bottom).First();
        var body = letters.Where(letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left is > 71 and < 73 &&
            letter.GlyphRectangle.Bottom is > 550 and < 720)
            .OrderByDescending(letter => letter.GlyphRectangle.Bottom).First();
        Assert.InRange(header.GlyphRectangle.Bottom - body.GlyphRectangle.Bottom,
            bodyOffset - 2, bodyOffset + 2);
    }

    [Fact]
    public void FlowVisitorKeepsHeaderLeftTabParagraphBetweenNeighboringParagraphs()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordLeftPositionalTabBoundariesAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var headers = page.GetWords().Where(word => word.Text == "Header" &&
            page.Height - word.Letters[0].StartBaseLine.Y < 150)
            .OrderByDescending(word => word.Letters[0].StartBaseLine.Y).ToArray();
        Assert.Equal(3, headers.Length);
        Assert.InRange(headers[0].Letters[0].StartBaseLine.Y -
            headers[1].Letters[0].StartBaseLine.Y, 27.0, 27.5);
        Assert.InRange(headers[1].Letters[0].StartBaseLine.Y -
            headers[2].Letters[0].StartBaseLine.Y, 41.9, 42.3);
    }

    [Theory]
    [InlineData("WordLeftPositionalTabAllStories", 58.5, 99.5, 733.7)]
    [InlineData("WordLeftPositionalTabBoundariesAllStories", 82.7, 191.8, 692.4)]
    public void FlowVisitorKeepsLeftMarginTabLinePitchAcrossStories(string name,
        double headerTop, double bodyTop, double footerTop)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var following = page.GetWords().Where(word => word.Text == "Following")
            .OrderByDescending(word => word.BoundingBox.Top).ToArray();
        Assert.Equal(3, following.Length);
        Assert.InRange(page.Height - following[0].BoundingBox.Top, headerTop - 1, headerTop + 1);
        Assert.InRange(page.Height - following[1].BoundingBox.Top, bodyTop - 1, bodyTop + 1);
        Assert.InRange(page.Height - following[2].BoundingBox.Top, footerTop - 1, footerTop + 1);
    }

    [Theory]
    [InlineData("WordPositionalTabCellMarginsAllStories", 42)]
    [InlineData("WordPositionalTabTenTwipCellMarginsAllStories", 25)]
    public void FlowVisitorSpacesPositionalTabsInTableCells(string name, int paragraphPitch)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "DocKnownGaps",
            name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var bodyStarts = pdf.GetPage(1).Letters.Where(letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left is > 70 and < 80 &&
            letter.GlyphRectangle.Bottom is > 600 and < 710)
            .OrderByDescending(letter => letter.GlyphRectangle.Bottom).ToArray();
        Assert.Equal(2, bodyStarts.Length);
        Assert.InRange(bodyStarts[0].GlyphRectangle.Bottom - bodyStarts[1].GlyphRectangle.Bottom,
            paragraphPitch - 1, paragraphPitch + 1);
    }

    [Fact]
    public void FlowVisitorDoesNotExpandMultirowTableForTrailingPositionalTab()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "DocKnownGaps",
            "WordTableCellPositionalTabProbe.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var rowStarts = page.Letters.Where(letter => letter.Value == "B" &&
                letter.GlyphRectangle.Left is > 76 and < 79 &&
                page.Height - letter.GlyphRectangle.Top is > 100 and < 150)
            .Select(letter => page.Height - letter.GlyphRectangle.Top).OrderBy(top => top).ToArray();
        Assert.Equal(2, rowStarts.Length);
        Assert.InRange(rowStarts[1] - rowStarts[0], 24, 26);
        Assert.Contains(page.Letters, letter => letter.Value == "F" &&
            page.Height - letter.GlyphRectangle.Top is > 708 and < 718);
    }

    [Fact]
    public void FlowVisitorKeepsIdeographDigitalListAtWordLinePitch()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordIdeographDigitalListsAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var bodyTops = page.GetWords().Where(word => word.Text is "First" or "Second" or "Third" &&
                word.BoundingBox.Left is > 107 and < 110 &&
                page.Height - word.BoundingBox.Top is > 70 and < 140)
            .Select(word => page.Height - word.BoundingBox.Top).OrderBy(top => top).ToArray();
        Assert.Equal(3, bodyTops.Length);
        Assert.InRange(bodyTops[1] - bodyTops[0], 19, 21);
        Assert.InRange(bodyTops[2] - bodyTops[1], 19, 21);
        var marginTops = page.GetWords().Where(word => word.Text is "First" or "Second" &&
                word.BoundingBox.Left is > 107 and < 110 &&
                (page.Height - word.BoundingBox.Top < 70 ||
                 page.Height - word.BoundingBox.Top > 700))
            .Select(word => page.Height - word.BoundingBox.Top).OrderBy(top => top).ToArray();
        Assert.Equal(4, marginTops.Length);
        Assert.InRange(marginTops[1] - marginTops[0], 17, 19);
        Assert.InRange(marginTops[3] - marginTops[2], 17, 19);
    }

    [Fact]
    public void FlowVisitorKeepsIdeographicZeroInDigitalMarkerFontSpan()
    {
        Assert.Equal(2, DxpFlowPdfVisitor.GetIdeographMarkerPrefixLength("一〇."));
        Assert.Equal(2, DxpFlowPdfVisitor.GetIdeographMarkerPrefixLength("一二."));
    }

    [Fact]
    public void FlowVisitorContinuesShadingThroughAdjacentParagraphSpacing()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyledParagraphShadingResetAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.Paths, path => path.IsFilled &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 142 and < 144 && bounds.Width is > 362 and < 364 &&
            bounds.Height is > 31.5 and < 34 &&
            page.Height - bounds.Top is > 71 and < 73);
        Assert.Contains(page.Paths, path => path.IsFilled &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 106 and < 108 && bounds.Width is > 362 and < 364 &&
            bounds.Height is > 23.5 and < 26 &&
            page.Height - bounds.Top is > 103 and < 106);
        Assert.Equal(2, pdf.NumberOfPages);
        foreach (var (pageNumber, expectedLeft) in new[] { (1, 142.5), (2, 106.5) })
        {
            var footerPage = pdf.GetPage(pageNumber);
            Assert.Contains(footerPage.Paths, path => path.IsFilled &&
                path.FillColor != null && path.FillColor.ToRGBValues() is var color &&
                color.r is > 0.76 and < 0.78 &&
                color.g is > 0.89 and < 0.91 && color.b is > 0.75 and < 0.77 &&
                path.GetBoundingRectangle() is { } bounds &&
                Math.Abs(bounds.Left - expectedLeft) < 2 &&
                bounds.Height is > 23.5 and < 26);
        }
    }

    [Fact]
    public void FlowVisitorPreservesCachedFieldResultRunFormatting()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "TestDocVariables.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(),
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Cache });

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Contains(pdf.GetPage(1).Letters, letter => letter.Value == "E" &&
            letter.GlyphRectangle.Left is > 310 and < 325 &&
            letter.FontName?.Contains("Bold", StringComparison.OrdinalIgnoreCase) == true);
    }

    [Fact]
    public void FlowVisitorDrawsDiagonalBordersAcrossStories()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordDiagonalBordersAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var diagonals = pdf.GetPage(1).Paths.Where(path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Width > 200 && bounds.Height > 10).ToArray();
        Assert.Equal(4, diagonals.Length);
    }

    [Fact]
    public void FlowVisitorAlignsColoredDiagonalTableEdge()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordLegacyTableBordersAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Contains(pdf.GetPage(1).Paths, path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 66 and < 67 && bounds.Width < 1 &&
            bounds.Height is > 14 and < 16);
        Assert.Contains(pdf.GetPage(1).Paths, path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 66.7 and < 67 && bounds.Right is > 300.2 and < 300.5 &&
            bounds.Height is > 23 and < 24);
    }

    [Theory]
    [InlineData("WordCellSpacingAllStories.docx")]
    [InlineData("WordCellSpacingResetAllStories.docx")]
    public void FlowVisitorPositionsSpacedSingleRowFooter(string fixture)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixture);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var footer = page.GetWords().First(word => word.Text == "Footer" &&
            page.Height - word.BoundingBox.Top > 600);
        Assert.InRange(page.Height - footer.BoundingBox.Top, 691, 694);
    }

    [Theory]
    [InlineData("WordCellSpacingAllStories.docx", 41.6, 42.1)]
    [InlineData("WordCellSpacingResetAllStories.docx", 29.6, 30.1)]
    public void FlowVisitorKeepsBodyRowPitchWhenCellSpacingIncreases(
        string fixture, double minimum, double maximum)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixture);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var rows = page.GetWords().Where(word => word.Text is "A" or "C" &&
                word.BoundingBox.Left is > 100 and < 150 &&
                page.Height - word.BoundingBox.Top is > 125 and < 220)
            .OrderBy(word => word.BoundingBox.Top).ToArray();
        Assert.Equal(2, rows.Length);
        Assert.InRange(rows[1].BoundingBox.Top - rows[0].BoundingBox.Top, minimum, maximum);
    }

    [Theory]
    [InlineData("WordCellSpacingResetAllStories.docx")]
    [InlineData("WordStyledCellSpacingResetAllStories.docx")]
    public void FlowVisitorKeepsRowPitchWhenCellSpacingResets(string fixture)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixture);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var rowTops = page.GetWords().Where(word => word.Text == "Body" &&
                word.BoundingBox.Left is > 75 and < 100 &&
                page.Height - word.BoundingBox.Top is > 125 and < 220)
            .Select(word => page.Height - word.BoundingBox.Top).OrderBy(top => top).ToArray();
        Assert.Equal(2, rowTops.Length);
        Assert.InRange(rowTops[1] - rowTops[0], 29.5, 30.5);
        if (fixture.StartsWith("WordStyled", StringComparison.Ordinal))
        {
            var footerTops = page.GetWords().Where(word => word.Text == "Footer" &&
                    word.BoundingBox.Left is > 75 and < 100 &&
                    page.Height - word.BoundingBox.Top > 680)
                .Select(word => page.Height - word.BoundingBox.Top).OrderBy(top => top).ToArray();
            Assert.Equal(2, footerTops.Length);
            Assert.InRange(footerTops[1] - footerTops[0], 29.5, 30.5);
        }
    }

    [Theory]
    [InlineData("WordMixedSectionStories.docx")]
    [InlineData("WordMixedSectionStoryContent.docx")]
    [InlineData("WordMultisectionPositionalTabsAllStories.docx")]
    public void FlowVisitorDoesNotNumberEmptySectionBreak(string fixture)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixture);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("9.Secondpagenote", pdf.GetPage(2).Text);
        Assert.DoesNotContain("10.Secondpagenote", pdf.GetPage(2).Text);
    }

    [Theory]
    [InlineData("WordMixedSectionStoryContent.docx")]
    [InlineData("WordMultisectionPositionalTabsAllStories.docx")]
    public void FlowVisitorKeepsSecondSectionBodyBelowMixedHeader(string fixture)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixture);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(2);
        var note = page.GetWords().First(word => word.Text == "note" &&
            page.Height - word.BoundingBox.Top < 100 && word.BoundingBox.Left >= 90);
        Assert.InRange(page.Height - note.BoundingBox.Top, 74, 76);
    }

    [Fact]
    public void FlowVisitorPreservesStyledMixedHeaderAfterSpacing()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordMixedSectionStyledStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(2);
        var note = page.GetWords().First(word => word.Text == "note" &&
            page.Height - word.BoundingBox.Top < 100);
        var secondHeader = page.GetWords().First(word => word.Text == "Second" &&
            page.Height - word.Letters[0].StartBaseLine.Y is > 74 and < 78);
        Assert.InRange(page.Height - secondHeader.Letters[0].StartBaseLine.Y, 75.5, 76.5);
        Assert.InRange(page.Height - note.Letters[0].StartBaseLine.Y, 88.8, 89.7);
        Assert.InRange(page.Height - note.BoundingBox.Top, 78, 82);
    }

    [Fact]
    public void FlowVisitorCentersInlineImageAfterSectionHeaderPositionalTab()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordMultisectionPositionalTabsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        var image = Assert.Single(pdf.GetPage(2).GetImages());
        Assert.InRange(image.Bounds.Left, 371.5, 372.5);
        Assert.InRange(image.Bounds.Width, 17.5, 18.5);
        var firstPage = pdf.GetPage(1);
        Assert.Contains(firstPage.Letters, letter => letter.Value == "F" &&
            firstPage.Height - letter.GlyphRectangle.Top is > 68 and < 79);
    }

    [Fact]
    public void FlowVisitorUsesNumberingTabInZeroIndentTableCell()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordMultisectionPositionalTabsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var image = Assert.Single(page.GetImages());
        Assert.InRange(image.Bounds.Left, 374.5, 376);
        var delta = page.GetWords().First(word => word.Text == "Delta");
        Assert.InRange(delta.BoundingBox.Left, 346.5, 348.5);
    }

    [Theory]
    [InlineData("WordMixedSectionStories")]
    [InlineData("WordMultisectionPositionalTabsAllStories")]
    public void FlowVisitorBalancesNumberedCellInlineImageLine(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var words = pdf.GetPage(1).GetWords();
        var gamma = words.First(word => word.Text == "Gamma");
        var delta = words.First(word => word.Text == "Delta");
        var following = words.First(word => word.Text == "8.");
        Assert.InRange(gamma.BoundingBox.Top - delta.BoundingBox.Top, 7.5, 7.9);
        Assert.InRange(delta.BoundingBox.Top - following.BoundingBox.Top, 23.4, 23.7);
    }

    [Fact]
    public void FlowVisitorSeparatesBodyAndHeaderFooterListCounters()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordMixedSectionStyledStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = System.Text.RegularExpressions.Regex.Replace(pdf.GetPage(2).Text, @"\s+", "");
        Assert.Contains("1.Secondsectionheader", page);
        Assert.Contains("2.Secondsectionfooter", page);
        Assert.Contains("9.Secondpagenote", page);
    }

    [Fact]
    public void FlowVisitorUsesDocumentDefaultTabInterval()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleDefaultTabAndTextAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.GetWords(), word => word.Text == "right" &&
            word.BoundingBox.Left is > 179.5 and < 180.5);
        Assert.Contains(page.GetWords(), word => word.Text == "right" &&
            word.BoundingBox.Left is > 125.5 and < 126.5);
        Assert.Contains(page.GetWords(), word => word.Text == "First" &&
            word.BoundingBox.Left is > 125.5 and < 126.5 &&
            page.Height - word.BoundingBox.Top is > 92.8 and < 93.4);
    }

    [Theory]
    [InlineData("WordVisibleCustomDefaultTabsAllStories.docx")]
    [InlineData("WordVisibleDefaultTabAndTextAllStories.docx")]
    public void FlowVisitorKeepsStorySpecificListTabStops(string fixture)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixture);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.GetWords(), word => word.Text == "First" &&
            word.BoundingBox.Left is > 305.5 and < 306.5 &&
            page.Height - word.BoundingBox.Top < 80);
        Assert.Contains(page.GetWords(), word => word.Text == "First" &&
            word.BoundingBox.Left is > 125.5 and < 126.5 &&
            page.Height - word.BoundingBox.Top is > 70 and < 120);
    }

    [Fact]
    public void FlowVisitorInheritsFirstPageFooterAcrossSections()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordSixSlotDiagonalCellBorders.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Contains("First-pagesection1footer", pdf.GetPage(1).Text);
        Assert.Contains("First-pagesection2footer", pdf.GetPage(2).Text);
        Assert.Contains("First-pagesection2footer", pdf.GetPage(3).Text);
        Assert.Contains("FooterD", pdf.GetPage(3).Text);
        var secondPage = pdf.GetPage(2);
        var footerLabel = secondPage.GetWords().First(word => word.Text == "footer2");
        Assert.InRange(secondPage.Height - footerLabel.BoundingBox.Top, 520.0, 521.1);
        var thirdPage = pdf.GetPage(3);
        var inheritedFooterLabel = thirdPage.GetWords().First(word => word.Text == "footer3");
        Assert.InRange(thirdPage.Height - inheritedFooterLabel.BoundingBox.Top, 700.0, 701.2);
        var footerCell = secondPage.GetWords().First(word => word.Text == "Footer");
        Assert.InRange(footerCell.BoundingBox.Left, 95.0, 97.0);
    }

    [Fact]
    public void FlowVisitorRendersPositionalTabLeadersAcrossStories()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordPositionalTabLeadersAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.True(page.Letters.Count(letter => letter.Value == ".") > 100);
        Assert.True(page.Letters.Count(letter => letter.Value == "-") > 80);
        Assert.Contains(page.Letters, letter => letter.Value == "d" &&
            letter.GlyphRectangle.Left is > 71 and < 74 &&
            page.Height - letter.GlyphRectangle.Top is > 134 and < 140);
        var headerLines = page.GetWords().Where(word =>
                word.Text is "dot" or "hyphen" or "underscore" or "middleDot" &&
                page.Height - word.BoundingBox.Top < 130)
            .OrderByDescending(word => word.BoundingBox.Top).ToArray();
        Assert.Equal(4, headerLines.Length);
        for (var index = 1; index < headerLines.Length; index++)
            Assert.InRange(headerLines[index - 1].BoundingBox.Top -
                headerLines[index].BoundingBox.Top, 24.8, 25.1);
        var footerLines = page.GetWords().Where(word =>
                word.Text is "dot" or "hyphen" or "underscore" or "middleDot" &&
                page.Height - word.BoundingBox.Top > 600)
            .OrderByDescending(word => word.BoundingBox.Top).ToArray();
        Assert.Equal(4, footerLines.Length);
        for (var index = 1; index < footerLines.Length; index++)
            Assert.InRange(footerLines[index - 1].BoundingBox.Top -
                footerLines[index].BoundingBox.Top, 24.8, 25.1);
    }

    [Theory]
    [InlineData("WordPositionalTabLeadersAllStories", "dot", 138.32)]
    [InlineData("WordLeftPositionalTabBoundariesAllStories", "Body", 146.54)]
    public void FlowVisitorPreservesHeaderPositionalTabBodyClearance(
        string name, string firstBodyWord, double expectedTop)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var body = page.GetWords().Single(word => word.Text == firstBodyWord &&
            page.Height - word.BoundingBox.Top is > 130 and < 155);
        Assert.InRange(page.Height - body.BoundingBox.Top, expectedTop - 0.25, expectedTop + 0.25);
    }

    [Theory]
    [InlineData("WordStyledTabLeadersAllStories", "S", ".")]
    [InlineData("WordInheritedTabLeadersAllStories", "D", "-")]
    [InlineData("WordClearedInheritedTabsAllStories", "C", "-")]
    public void FlowVisitorAlignsInheritedLeaderBaselineAndFont(string name,
        string firstLetter, string leader)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var first = page.Letters.First(letter => letter.Value == firstLetter &&
            letter.GlyphRectangle.Left is > 71 and < 74 &&
            page.Height - letter.GlyphRectangle.Top is > 65 and < 100);
        Assert.InRange(page.Height - first.GlyphRectangle.Top, 77, 79);
        var leaderLetter = page.Letters.First(letter => letter.Value == leader &&
            letter.GlyphRectangle.Left > 120 &&
            page.Height - letter.GlyphRectangle.Top is > 65 and < 100);
        Assert.InRange(leaderLetter.PointSize, 11.5, 12.5);
    }

    [Fact]
    public void FlowVisitorRendersInheritedSmallCapsAcrossStories()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedEmptyMarkAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains("HEADERANCHOR", page.Text);
        Assert.Contains("BODYANCHOR", page.Text);
        Assert.Contains("FOOTERANCHOR", page.Text);
        Assert.Contains(page.Letters, letter => letter.Value == "H" && letter.PointSize is > 13 and < 15);
        Assert.Contains(page.Letters, letter => letter.Value == "E" && letter.PointSize is > 10 and < 12);
        Assert.Contains(page.Letters, letter => letter.Value == "H" &&
            page.Height - letter.GlyphRectangle.Top is > 70 and < 74);
        Assert.Contains(page.Letters, letter => letter.Value == "B" &&
            page.Height - letter.GlyphRectangle.Top is > 141 and < 145);
        Assert.Contains(page.Letters, letter => letter.Value == "F" &&
            page.Height - letter.GlyphRectangle.Top is > 725 and < 730);
    }

    [Fact]
    public void FlowVisitorUsesInheritedFontSizeForEmptyBodyParagraphMark()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordCharacterStyleCapsResetAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var lastBody = page.GetWords().First(word => word.Text == "LAST" &&
            page.Height - word.BoundingBox.Top > 100);
        Assert.InRange(page.Height - lastBody.BoundingBox.Top, 128, 133);
    }

    [Theory]
    [InlineData("WordVisibleBoldOnlyDefaultsAllStories.docx", "Times New Roman")]
    [InlineData("WordVisibleFontOnlyDefaultsAllStories.docx", "Aptos")]
    public void FlowVisitorUsesTenPointClassicPartialRunDefaults(string fixture, string font)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixture);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = pdf.GetPage(1).Letters.Where(letter => letter.Value == "B").ToArray();
        Assert.NotEmpty(letters);
        Assert.All(letters, letter =>
        {
            Assert.InRange(letter.PointSize, 9.9, 10.1);
            if (font == "Times New Roman")
                Assert.Contains(font, letter.FontName, StringComparison.OrdinalIgnoreCase);
        });
        if (font == "Times New Roman")
        {
            var page = pdf.GetPage(1);
            var footer = page.GetWords().First(word => word.Text == "Footer");
            Assert.InRange(page.Height - footer.BoundingBox.Top, 723.5, 725);
        }
        else
        {
            var page = pdf.GetPage(1);
            var body = page.Letters.Single(letter => letter.Value == "B" &&
                letter.GlyphRectangle.Left is > 71 and < 73 &&
                page.Height - letter.StartBaseLine.Y is > 75 and < 90);
            Assert.InRange(page.Height - body.StartBaseLine.Y, 81.7, 82.1);
            var footer = page.Letters.Single(letter => letter.Value == "F" &&
                letter.GlyphRectangle.Left is > 71 and < 73 &&
                page.Height - letter.StartBaseLine.Y is > 720 and < 735);
            Assert.InRange(page.Height - footer.StartBaseLine.Y, 730.1, 730.6);
        }
    }

    [Fact]
    public void FlowVisitorPaintsEastAsianEmphasisWithoutChangingText()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordEastAsianEmphasisAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains("Secondparagraph", page.Text);
        var headerWord = page.Letters.First(letter => letter.Value == "W" &&
            page.Height - letter.StartBaseLine.Y < 70);
        Assert.InRange(page.Height - headerWord.StartBaseLine.Y, 55.8, 56.8);
        var glyphUnderDots = page.Letters.Where(letter => letter.Value == "．" &&
            page.Height - letter.StartBaseLine.Y < 70).ToArray();
        if (glyphUnderDots.Length > 0)
        {
            Assert.Equal(43, glyphUnderDots.Length);
            Assert.InRange(headerWord.StartBaseLine.Y - glyphUnderDots[0].StartBaseLine.Y,
                4.3, 4.7);
        }
        else
        {
            Assert.True(page.Paths.Count(path => path.IsFilled || path.IsStroked) > 100);
            var headerUnderDots = page.Paths.Where(path => path.IsFilled &&
                path.GetBoundingRectangle() is { } bounds &&
                bounds.Width is > 1.3 and < 1.5 && bounds.Height is > 1.3 and < 1.5 &&
                page.Height - bounds.Top is > 58.9 and < 59.3).ToArray();
            Assert.Equal(43, headerUnderDots.Length);
            var underDotCenters = headerUnderDots.Select(path => path.GetBoundingRectangle())
                .Select(bounds => (bounds!.Value.Left + bounds.Value.Right) / 2)
                .OrderBy(center => center).ToArray();
            Assert.InRange(underDotCenters[25] - underDotCenters[24], 8.5, 9.5);
            Assert.InRange(underDotCenters[26] - underDotCenters[25], 8.5, 9.5);
        }
        var secondParagraph = page.Letters.First(letter => letter.Value == "S" &&
            page.Height - letter.GlyphRectangle.Top > 120);
        var glyphTop = page.Height - secondParagraph.GlyphRectangle.Top;
        var commaMark = page.Letters.FirstOrDefault(letter => letter.Value == "、" &&
            letter.StartBaseLine.X >= secondParagraph.StartBaseLine.X &&
            letter.StartBaseLine.X <= secondParagraph.StartBaseLine.X + 2);
        if (commaMark != null)
            Assert.InRange(Math.Abs(secondParagraph.StartBaseLine.Y - commaMark.StartBaseLine.Y),
                11.5, 12.7);
        else
            Assert.Contains(page.Paths, path => path.IsStroked &&
                path.GetBoundingRectangle() is { } bounds &&
                bounds.Left >= secondParagraph.GlyphRectangle.Left &&
                bounds.Right <= secondParagraph.GlyphRectangle.Right &&
                bounds.Width is > 1.5 and < 3 &&
                glyphTop - (page.Height - bounds.Top) is > 3 and < 7);
    }

    [Fact]
    public void FlowVisitorPlacesRubyGuideAboveBaseWithoutAdvancingFollowingText()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var ruby = new Ruby(
                new RubyProperties { InnerXml = "<w:rubyAlign w:val=\"center\"/><w:hps w:val=\"10\"/>" +
                    "<w:hpsRaise w:val=\"16\"/><w:hpsBaseText w:val=\"24\"/>" },
                new RubyContent(new Run(new Text("ni"))),
                new RubyBase(new Run(new Text("JP"))));
            main.Document = new Document(new Body(
                new Paragraph(
                    new Run(new Text("JP Before ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }),
                    new Run(ruby),
                    new Run(new Text(" after.") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains("JP", page.Text);
        Assert.Contains("ni", page.Text);
        var baseLetter = page.Letters.Last(letter => letter.Value == "J");
        var guideLetter = page.Letters.First(letter => letter.Value == "n");
        var following = page.Letters.First(letter => letter.Value == "a" &&
            letter.StartBaseLine.X > baseLetter.StartBaseLine.X);
        Assert.InRange(following.StartBaseLine.X - baseLetter.StartBaseLine.X, 11, 23);
        Assert.InRange(guideLetter.StartBaseLine.X - baseLetter.StartBaseLine.X, -5, 12);
        Assert.True(guideLetter.PointSize < baseLetter.PointSize / 2);
        Assert.True(guideLetter.StartBaseLine.Y > baseLetter.StartBaseLine.Y);
    }

    [Fact]
    public void FlowVisitorKeepsUnpreservedSpaceBeforeDirectCjkButNotRuby()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));

        double FollowingTextOrigin(string relativePath)
        {
            var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", relativePath);
            var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
            return pdf.GetPage(1).Letters.First(letter => letter.Value == "a").StartBaseLine.X;
        }

        var directOrigin = FollowingTextOrigin(Path.Combine("PdfProbeCorpus", "ruby-raise-probe__ruby-control.docx"));
        var rubyOrigin = FollowingTextOrigin(Path.Combine("PdfFeatures", "RubyYuMincho.docx"));
        Assert.InRange(directOrigin - rubyOrigin, 2, 3.5);
    }

    [Theory]
    [InlineData("ruby-exact-raise-0")]
    [InlineData("ruby-exact-raise-32")]
    public void FlowVisitorAppliesExactLineRubyBaselineOffset(string fixture)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            $"ruby-raise-probe__{fixture}.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var first = page.Letters.First(letter => letter.Value == "B");
        var offset = Environment.GetEnvironmentVariable("DOCXPORT_DIAG_ARIAL_EXACT_BASELINE") == "1" ? 0.2 : 0;
        Assert.InRange(page.Height - first.StartBaseLine.Y, 82.35 + offset, 82.55 + offset);
    }

    [Theory]
    [InlineData(10, 6, 13)]
    [InlineData(16, 9, 14)]
    public void FlowVisitorExpandsRubyLineForRaisedGuide(int guideHalfPoints,
        double minimumBaseShift, double maximumBaseShift)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;

        (double BaseTop, double GuideRise) Render(int raiseHalfPoints, int? secondRaiseHalfPoints = null)
        {
            using var source = new MemoryStream();
            using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
            {
                var main = document.AddMainDocumentPart();
                Ruby CreateRuby(int halfPoints) => new(
                    new RubyProperties { InnerXml =
                        $"<w:rubyAlign w:val=\"center\"/><w:hps w:val=\"{guideHalfPoints}\"/>" +
                        $"<w:hpsRaise w:val=\"{halfPoints}\"/><w:hpsBaseText w:val=\"24\"/>" },
                    new RubyContent(new Run(new Text("ni"))),
                    new RubyBase(new Run(new Text("JP"))));
                var paragraph = new Paragraph(new Run(new Text("Before")),
                    new Run(CreateRuby(raiseHalfPoints)), new Run(new Text("after")));
                if (secondRaiseHalfPoints is int secondRaise)
                    paragraph.Append(new Run(CreateRuby(secondRaise)));
                main.Document = new Document(new Body(paragraph));
                main.Document.Save();
            }

            var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
            var page = pdf.GetPage(1);
            var baseLetter = page.Letters.First(letter => letter.Value == "B");
            var guideLetter = page.Letters.First(letter => letter.Value == "n");
            return (page.Height - baseLetter.GlyphRectangle.Top,
                guideLetter.StartBaseLine.Y - baseLetter.StartBaseLine.Y);
        }

        var normal = Render(0);
        var raised = Render(32);
        Assert.InRange(raised.BaseTop - normal.BaseTop, minimumBaseShift, maximumBaseShift);
        Assert.InRange(normal.GuideRise, -0.5, 0.5);
        Assert.InRange(raised.GuideRise, 15.5, 16.5);
        var lowThenHigh = Render(0, 32);
        var highThenLow = Render(32, 0);
        Assert.InRange(Math.Abs(lowThenHigh.BaseTop - highThenLow.BaseTop), 0, 0.5);
    }

    [Fact]
    public void FlowVisitorKeepsAtLeastLineSpacingDistinctFromExact()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;

        double FollowingTop(LineSpacingRuleValues lineRule)
        {
            using var source = new MemoryStream();
            using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
            {
                var main = document.AddMainDocumentPart();
                main.Document = new Document(new Body(
                    new Paragraph(
                        new ParagraphProperties(new SpacingBetweenLines
                        {
                            Line = "240", LineRule = lineRule
                        }),
                        new Run(new RunProperties(new FontSize { Val = "48" }), new Text("Tall"))),
                    new Paragraph(new Run(new Text("After")))));
                main.Document.Save();
            }

            var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
            var page = pdf.GetPage(1);
            var following = page.GetWords().Single(word => word.Text == "After");
            return page.Height - following.BoundingBox.Top;
        }

        var exact = FollowingTop(LineSpacingRuleValues.Exact);
        var atLeast = FollowingTop(LineSpacingRuleValues.AtLeast);
        Assert.InRange(atLeast - exact, 12, 28);
    }

    [Fact]
    public void FlowVisitorUsesNaturalBaselineForExplicitAutoSingleSpacing()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;

        double FirstWordTop(LineSpacingRuleValues lineRule)
        {
            using var source = new MemoryStream();
            using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
            {
                var main = document.AddMainDocumentPart();
                main.Document = new Document(new Body(new Paragraph(
                    new ParagraphProperties(new SpacingBetweenLines
                    {
                        Before = "0", After = "0", Line = "240", LineRule = lineRule
                    }),
                    new Run(new RunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                        new FontSize { Val = "48" }), new Text("Tall")))));
                main.Document.Save();
            }

            var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
            var page = pdf.GetPage(1);
            return page.Height - page.GetWords().Single(word => word.Text == "Tall").BoundingBox.Top;
        }

        Assert.InRange(Math.Abs(FirstWordTop(LineSpacingRuleValues.Auto) -
            FirstWordTop(LineSpacingRuleValues.AtLeast)), 0, 0.15);
    }

    [Fact]
    public void FlowVisitorUsesImplicitLineSpacingWithoutStyles()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;

        double ThirdParagraphTop(bool explicitAuto)
        {
            using var source = new MemoryStream();
            using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
            {
                var main = document.AddMainDocumentPart();
                var followingSpacing = new SpacingBetweenLines { Before = "0", After = "0" };
                if (explicitAuto)
                {
                    followingSpacing.Line = "240";
                    followingSpacing.LineRule = LineSpacingRuleValues.Auto;
                }
                main.Document = new Document(new Body(
                    new Paragraph(
                        new ParagraphProperties(new SpacingBetweenLines
                        {
                            Before = "0", After = "0", Line = "240", LineRule = LineSpacingRuleValues.Exact
                        }),
                        new Run(new RunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                            new FontSize { Val = "20" }), new Text("Exact line"))),
                    new Paragraph(new ParagraphProperties(followingSpacing),
                        new Run(new RunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                            new FontSize { Val = "22" }), new Text("Following line"))),
                    new Paragraph(
                        new ParagraphProperties(new SpacingBetweenLines
                        {
                            Before = "0", After = "0", Line = "240", LineRule = LineSpacingRuleValues.AtLeast
                        }),
                        new Run(new RunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                            new FontSize { Val = "20" }), new Text("Final line")))));
                main.Document.Save();
            }

            var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
            var page = pdf.GetPage(1);
            var final = page.GetWords().Single(word => word.Text == "Final");
            return page.Height - final.BoundingBox.Top;
        }

        Assert.InRange(ThirdParagraphTop(false) - ThirdParagraphTop(true), 1.5, 2.5);
    }


    [Theory]
    [InlineData("left", true, false)]
    [InlineData("right", false, true)]
    [InlineData("distributeLetter", true, true)]
    [InlineData("distributeSpace", false, false)]
    public void FlowVisitorAlignsRubyGuideToBase(string alignment, bool matchesLeft, bool matchesRight)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var ruby = new Ruby(
                new RubyProperties { InnerXml = $"<w:rubyAlign w:val=\"{alignment}\"/><w:hps w:val=\"10\"/>" },
                new RubyContent(new Run(new Text("ni"))),
                new RubyBase(new Run(new Text("JP"))));
            main.Document = new Document(new Body(new Paragraph(new Run(ruby))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = pdf.GetPage(1).Letters;
        var baseStart = Assert.Single(letters.Where(letter => letter.Value == "J"));
        var baseEnd = Assert.Single(letters.Where(letter => letter.Value == "P"));
        var guideStart = Assert.Single(letters.Where(letter => letter.Value == "n"));
        var guideEnd = Assert.Single(letters.Where(letter => letter.Value == "i"));
        Assert.Equal(matchesLeft,
            Math.Abs(guideStart.StartBaseLine.X - baseStart.StartBaseLine.X) < 1.5);
        Assert.Equal(matchesRight,
            Math.Abs(guideEnd.EndBaseLine.X - baseEnd.EndBaseLine.X) < 1.5);
        if (alignment == "distributeSpace")
        {
            Assert.InRange(guideStart.StartBaseLine.X - baseStart.StartBaseLine.X, 1.5, 4);
            Assert.InRange(baseEnd.EndBaseLine.X - guideEnd.EndBaseLine.X, 1.5, 4);
        }
    }

    [Fact]
    public void FlowVisitorStacksRightVerticalRubyGuideBesideBase()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var ruby = new Ruby(
                new RubyProperties { InnerXml = "<w:rubyAlign w:val=\"rightVertical\"/>" +
                    "<w:hps w:val=\"10\"/><w:hpsRaise w:val=\"16\"/>" },
                new RubyContent(new Run(new Text("nic"))),
                new RubyBase(new Run(new Text("JP"))));
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("Before ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }),
                new Run(ruby),
                new Run(new Text(" after.") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var letters = pdf.GetPage(1).Letters;
        Assert.DoesNotContain(letters, letter => letter.Value == "\0");
        var baseRight = Assert.Single(letters.Where(letter => letter.Value == "P")).EndBaseLine.X;
        var guide = letters.Where(letter => letter.Value is "n" or "i" or "c").ToArray();
        Assert.Equal(3, guide.Length);
        Assert.All(guide, letter => Assert.InRange(letter.StartBaseLine.X - baseRight, -2, 2));
        Assert.InRange(guide[0].StartBaseLine.Y - guide[1].StartBaseLine.Y, 4.5, 5.5);
        Assert.InRange(guide[1].StartBaseLine.Y - guide[2].StartBaseLine.Y, 4.5, 5.5);
        var following = Assert.Single(letters.Where(letter => letter.Value == "a"));
        Assert.True(following.StartBaseLine.X > guide[0].EndBaseLine.X + 1);
    }

    [Fact]
    public void FlowVisitorImportsHtmlAltChunkWithHeadingAndBoldRun()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var chunkPart = main.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html,
                "rIdHtmlChunk");
            using (var chunkStream = chunkPart.GetStream(FileMode.Create, FileAccess.Write))
            {
                var html = System.Text.Encoding.UTF8.GetBytes(
                    "<html><body><h1>Imported heading</h1><p>Imported <b>bold</b> paragraph.</p></body></html>");
                chunkStream.Write(html, 0, html.Length);
            }
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Before chunk"))),
                new AltChunk { Id = "rIdHtmlChunk" },
                new Paragraph(new Run(new Text("After chunk")))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains("Importedheading", page.Text);
        Assert.Contains("Importedboldparagraph.", page.Text);
        var heading = page.Letters.First(letter => letter.Value == "I" &&
            page.Height - letter.StartBaseLine.Y is > 110 and < 140);
        var bold = page.Letters.First(letter => letter.Value == "b" &&
            page.Height - letter.StartBaseLine.Y is > 140 and < 170);
        Assert.InRange(heading.PointSize, 23.9, 24.1);
        Assert.Contains("Bold", bold.FontName);
        var after = page.Letters.First(letter => letter.Value == "A" &&
            page.Height - letter.StartBaseLine.Y > 165);
        Assert.InRange(page.Height - after.StartBaseLine.Y, 175, 195);
    }

    [Fact]
    public void FlowVisitorImportsPlainTextAltChunkBetweenParagraphs()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var chunkPart = main.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.TextPlain,
                "rIdPlainChunk");
            using (var chunkStream = chunkPart.GetStream(FileMode.Create, FileAccess.Write))
            {
                var content = System.Text.Encoding.UTF8.GetBytes(
                    "Imported first line\r\nImported second line\r\n");
                chunkStream.Write(content, 0, content.Length);
            }
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Before chunk"))),
                new AltChunk { Id = "rIdPlainChunk" },
                new Paragraph(new Run(new Text("After chunk")))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains("ImportedfirstlineImportedsecondline", page.Text);
        var imported = page.GetWords().Where(word => word.Text == "Imported")
            .OrderByDescending(word => word.BoundingBox.Top).ToArray();
        Assert.Equal(2, imported.Length);
        Assert.InRange(imported[0].BoundingBox.Top - imported[1].BoundingBox.Top, 11.5, 12.5);
        Assert.Contains("Afterchunk", page.Text);
    }

    [Fact]
    public void FlowVisitorPreservesTrailingBlankLineInPlainTextAltChunk()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var chunkPart = main.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.TextPlain,
                "rIdPlainChunk");
            using (var chunkStream = chunkPart.GetStream(FileMode.Create, FileAccess.Write))
            {
                var content = System.Text.Encoding.UTF8.GetBytes("First line\r\nSecond line\r\n\r\n");
                chunkStream.Write(content);
            }
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Before chunk"))),
                new AltChunk { Id = "rIdPlainChunk" },
                new Paragraph(new Run(new Text("After chunk")))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var second = Assert.Single(page.GetWords().Where(word => word.Text == "Second"));
        var after = Assert.Single(page.GetWords().Where(word => word.Text == "After"));
        Assert.InRange(second.BoundingBox.Top - after.BoundingBox.Top, 23, 25);
    }

    [Theory]
    [InlineData("\f", 2)]
    [InlineData("\v", 1)]
    public void FlowVisitorImportsPlainTextControlBreaks(string control, int expectedPages)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var chunkPart = main.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.TextPlain,
                "rIdPlainChunk");
            using (var chunkStream = chunkPart.GetStream(FileMode.Create, FileAccess.Write))
            {
                var content = System.Text.Encoding.UTF8.GetBytes($"First line{control}Second line\r\n");
                chunkStream.Write(content);
            }
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Before chunk"))),
                new AltChunk { Id = "rIdPlainChunk" },
                new Paragraph(new Run(new Text("After chunk")))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var pages = pdf.GetPages().ToArray();
        Assert.Equal(expectedPages, pages.Length);
        Assert.Contains("Firstline", pages[0].Text);
        Assert.Contains("Secondline", pages[^1].Text);
        Assert.Contains("Afterchunk", pages[^1].Text);
    }

    [Theory]
    [InlineData(false, 24.36)]
    [InlineData(true, 52.16)]
    public void FlowVisitorImportsHtmlAltChunkDataImage(bool precedingHtmlParagraph,
        double expectedImageOffset)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var chunkPart = main.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html,
                "rIdHtmlImage");
            using (var chunkStream = chunkPart.GetStream(FileMode.Create, FileAccess.Write))
            {
                const string image = "iVBORw0KGgoAAAANSUhEUgAAAGAAAAAwCAIAAABhdOiYAAAAZElEQVR4nO3QQQ0AIBDAMLScfymIwgE82aPJBCxde0aX1veDeIAAAQIEKBwgQIAAAQoHCBAgQIDCAQIECBCgcIAAAQIEKBwgQIAAAQoHCBAgQIDCAQIECBCgcIAAAQIEKBygRweJKyEeCj8YAgAAAABJRU5ErkJggg==";
                var html = System.Text.Encoding.UTF8.GetBytes(
                    $"<html><body>{(precedingHtmlParagraph ? "<p>Prelude</p>" : "")}<p>Image before <img src=\"data:image/png;base64,{image}\" " +
                    "width=\"96\" height=\"48\"/> image after</p></body></html>");
                chunkStream.Write(html);
            }
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Before chunk"))),
                new AltChunk { Id = "rIdHtmlImage" },
                new Paragraph(new Run(new Text("After chunk"))),
                new SectionProperties(
                    new PageSize { Width = 12240, Height = 15840 },
                    new PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 })));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var imageResult = Assert.Single(page.GetImages());
        Assert.InRange(imageResult.Bounds.Width, 71.9, 72.1);
        Assert.InRange(imageResult.Bounds.Height, 35.9, 36.1);
        var first = page.GetWords().First(word => word.Text == "Before");
        var firstTop = page.Height - first.BoundingBox.Top;
        Assert.InRange(page.Height - imageResult.Bounds.Top - firstTop,
            expectedImageOffset - 0.15, expectedImageOffset + 0.15);
        Assert.Contains("Imagebefore", page.Text);
        Assert.Contains("imageafter", page.Text);
        var following = page.GetWords().First(word => word.Text == "After");
        Assert.True(following.BoundingBox.Top < imageResult.Bounds.Bottom);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FlowVisitorReservesMissingHtmlImageBox(bool fromBytes)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            "altchunk-probe__html-relative-image.docx");
        var bytes = fromBytes
            ? DxpExport.ExportToBytes(File.ReadAllBytes(source), new DxpFlowPdfVisitor())
            : DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var image = Assert.Single(page.GetImages());
        Assert.InRange(image.Bounds.Width, 71.9, 72.1);
        Assert.InRange(image.Bounds.Height, 35.9, 36.1);
        Assert.Contains(page.Paths, path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Width is > 71.9 and < 72.1 && bounds.Height is > 35.9 and < 36.1);
        var before = page.GetWords().First(word => word.Text == "Image");
        var after = page.GetWords().First(word => word.Text == "After");
        Assert.InRange(page.Height - before.BoundingBox.Top, 125, 135);
        Assert.InRange(page.Height - after.BoundingBox.Top, 152, 162);
    }

    [Fact]
    public void FlowVisitorImportsHtmlAltChunkListAsSeparateItems()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var chunkPart = main.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html,
                "rIdHtmlList");
            using (var chunkStream = chunkPart.GetStream(FileMode.Create, FileAccess.Write))
            {
                var html = System.Text.Encoding.UTF8.GetBytes(
                    "<html><body><ul><li>First item</li><li>Second item</li></ul></body></html>");
                chunkStream.Write(html, 0, html.Length);
            }
            main.Document = new Document(new Body(
                new AltChunk { Id = "rIdHtmlList" },
                new Paragraph(new Run(new Text("After chunk")))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains("Firstitem", page.Text);
        Assert.Contains("Seconditem", page.Text);
        var markers = page.Letters.Where(letter => letter.Value == "●").ToArray();
        Assert.Equal(2, markers.Length);
        Assert.All(markers, marker => Assert.InRange(marker.StartBaseLine.X, 89.5, 90.5));
        var first = page.Letters.First(letter => letter.Value == "F");
        var second = page.Letters.First(letter => letter.Value == "S");
        Assert.InRange(first.StartBaseLine.X, 107.5, 108.5);
        Assert.InRange(second.StartBaseLine.X, 107.5, 108.5);
        Assert.InRange(first.StartBaseLine.Y - second.StartBaseLine.Y, 13.5, 15.5);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FlowVisitorImportsHtmlAltChunkTableAsCells(bool followingParagraph)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var chunkPart = main.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html,
                "rIdHtmlTable");
            using (var chunkStream = chunkPart.GetStream(FileMode.Create, FileAccess.Write))
            {
                var html = System.Text.Encoding.UTF8.GetBytes(
                    "<html><body><table border=\"1\"><tr><th>Header A</th><th>Header B</th></tr>" +
                    "<tr><td>Body A</td><td>Body B</td></tr></table>" +
                    (followingParagraph ? "<p>Inside after table.</p>" : "") + "</body></html>");
                chunkStream.Write(html, 0, html.Length);
            }
            main.Document = new Document(new Body(
                new AltChunk { Id = "rIdHtmlTable" },
                new Paragraph(new Run(new Text("After chunk")))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains("HeaderA", page.Text);
        Assert.Contains("HeaderB", page.Text);
        Assert.Contains("BodyA", page.Text);
        Assert.Contains("BodyB", page.Text);
        if (followingParagraph) Assert.Contains("Insideaftertable.", page.Text);
        var headers = page.Letters.Where(letter => letter.Value == "H")
            .OrderBy(letter => letter.StartBaseLine.X).ToArray();
        Assert.Equal(2, headers.Length);
        Assert.InRange(headers[0].StartBaseLine.X, 71.5, 72.5);
        Assert.InRange(headers[1].StartBaseLine.X, 126.4, 126.9);
        Assert.Contains(page.Paths, path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds && bounds.Width > 50);
        using var painted = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var commands = string.Concat(painted.Pages[0].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Select(reference => System.Text.Encoding.Latin1.GetString(
                ((PdfSharp.Pdf.Advanced.PdfContent)reference.Value!).Stream!.UnfilteredValue)));
        Assert.Contains(".937 .937 .937 RG", commands);
        Assert.Contains(".624 .624 .624 RG", commands);
    }

    [Theory]
    [InlineData("padding0")]
    [InlineData("padding-default")]
    public void FlowVisitorKeepsNarrowPaddingHtmlTableHeadersOnOneLine(string variant)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            $"altchunk-bevel-probe__{variant}.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = Assert.Single(pdf.GetPages());
        var firstHeader = Assert.Single(page.Letters.Where(letter => letter.Value == "H" &&
            letter.StartBaseLine.X is > 71.5 and < 72.5));
        var firstHeaderSuffix = Assert.Single(page.Letters.Where(letter => letter.Value == "A" &&
            letter.StartBaseLine.X is > 112 and < 124));
        Assert.InRange(Math.Abs(firstHeader.StartBaseLine.Y - firstHeaderSuffix.StartBaseLine.Y), 0, 0.1);
    }

    [Theory]
    [InlineData("WordInheritedEmphasisAllStories", 43.7, 44.1)]
    [InlineData("WordInheritedCharacterEmphasisGridAllStories", 45.0, 45.4)]
    public void FlowVisitorPositionsInheritedHeaderDotsAtWordHeight(string name,
        double minimum, double maximum)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.Paths, path => path.IsFilled &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 74 and < 85 && bounds.Width is > 1.3 and < 1.5 &&
            page.Height - bounds.Top >= minimum && page.Height - bounds.Top <= maximum);
    }

    [Theory]
    [InlineData("WordEastAsianAutomaticSpacingAllStories.docx", 127, 135)]
    [InlineData("WordEastAsianEmphasisAllStories.docx", 142, 151)]
    public void FlowVisitorUsesCjkParagraphPitch(string fixture, double minimumTop, double maximumTop)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixture);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var positions = page.Letters.Where(letter => letter.Value == "S")
            .Select(letter => (letter.GlyphRectangle.Left, Top: page.Height - letter.GlyphRectangle.Top))
            .ToArray();
        Assert.True(positions.Any(position => position.Left is > 71 and < 74 &&
            position.Top > minimumTop && position.Top < maximumTop),
            string.Join(", ", positions));
    }

    [Fact]
    public void DirectParagraphAfterSpacingRetainsInheritedLineSpacing()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        using var document = WordprocessingDocument.Open(Path.Combine(root, "samples", "file-sample_1MB.docx"), false);
        var paragraph = document.MainDocumentPart!.Document.Body!.Elements<Paragraph>()
            .First(candidate => candidate.InnerText.StartsWith("Vestibulum neque massa", StringComparison.Ordinal));

        var spacing = new DxpStyleResolver(document).GetParagraphSpacing(paragraph);

        Assert.Equal("225", spacing?.After?.Value);
        Assert.Equal("288", spacing?.Line?.Value);
        Assert.Equal(LineSpacingRuleValues.Auto, spacing?.LineRule?.Value);
    }

    [Fact]
    public void ParagraphSpacingExplicitValuesOverrideInheritedAutoSpacingFlags()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        using var document = WordprocessingDocument.Open(Path.Combine(root, "samples", "Right to Repair.docx"), false);
        var paragraph = document.MainDocumentPart!.Document.Body!.Elements<Paragraph>()
            .First(candidate => candidate.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "NormalWeb");

        var spacing = new DxpStyleResolver(document).GetParagraphSpacing(paragraph);

        Assert.Equal("120", spacing?.Before?.Value);
        Assert.Equal("120", spacing?.After?.Value);
        Assert.Null(spacing?.BeforeAutoSpacing);
        Assert.Null(spacing?.AfterAutoSpacing);

        var heading = document.MainDocumentPart.Document.Body.Elements<Paragraph>()
            .First(candidate => candidate.InnerText == "What it means");
        var headingSpacing = new DxpStyleResolver(document).GetParagraphSpacing(heading);
        Assert.True(headingSpacing?.BeforeAutoSpacing?.Value);
        Assert.True(headingSpacing?.AfterAutoSpacing?.Value);

        var explicitAutoOff = (Paragraph)heading.CloneNode(true);
        explicitAutoOff.ParagraphProperties!.SpacingBetweenLines = new SpacingBetweenLines
        {
            BeforeAutoSpacing = false,
            AfterAutoSpacing = false
        };
        var explicitSpacing = new DxpStyleResolver(document).GetParagraphSpacing(explicitAutoOff);
        Assert.Equal("100", explicitSpacing?.Before?.Value);
        Assert.Equal("100", explicitSpacing?.After?.Value);
        Assert.False(explicitSpacing?.BeforeAutoSpacing?.Value);
        Assert.False(explicitSpacing?.AfterAutoSpacing?.Value);
    }

    [Fact]
    public void FlowVisitorAlignsDejaVuBodyBaselineWithoutChangingPagination()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "file-sample_1MB.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(5, pdf.NumberOfPages);
        var page = pdf.GetPage(3);
        var body = page.Letters.First(letter => letter.Value == "p" &&
            letter.StartBaseLine.X is > 67 and < 69 &&
            page.Height - letter.StartBaseLine.Y is > 240 and < 270);
        Assert.InRange(page.Height - body.StartBaseLine.Y, 254, 256);

        var firstPage = pdf.GetPage(1);
        var regular = firstPage.Letters.First(letter => letter.Value == "c" &&
            firstPage.Height - letter.StartBaseLine.Y is > 230 and < 260);
        var bold = firstPage.Letters.First(letter => letter.Value == "V" &&
            firstPage.Height - letter.StartBaseLine.Y is > 230 and < 260);
        Assert.InRange(Math.Abs(regular.StartBaseLine.Y - bold.StartBaseLine.Y), 0, 0.2);

    }

    [Fact]
    public void FlowVisitorRendersFootnoteReferenceAsSuperscript()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Right to Repair.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var commands = string.Concat(pdf.Pages[0].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Select(reference => System.Text.Encoding.Latin1.GetString(
                ((PdfSharp.Pdf.Advanced.PdfContent)reference.Value!).Stream!.UnfilteredValue)));
        var marker = System.Text.RegularExpressions.Regex.Match(commands,
            @"/F\d+ (?<size>\d+(?:\.\d+)?) Tf\s+\d+(?:\.\d+)? (?<rise>\d+(?:\.\d+)?) Td\s+100 Tz\s+0 Tc\s+\(1\) Tj");
        Assert.True(marker.Success);
        Assert.InRange(double.Parse(marker.Groups["size"].Value,
            System.Globalization.CultureInfo.InvariantCulture), 6.8, 7.2);
        Assert.InRange(double.Parse(marker.Groups["rise"].Value,
            System.Globalization.CultureInfo.InvariantCulture), 3, 4.5);
        Assert.Contains(pdf.Pages[0].Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>(),
            annotation => annotation.Elements.GetDictionary("/A")?.Elements.GetString("/D") ==
                "dxp-note-1");
        var secondNote = Assert.IsType<PdfSharp.Pdf.PdfArray>(
            pdf.Internals.Catalog.Names!.NameTree!.GetValue("dxp-note-2", true));
        var targetPage = Assert.IsType<PdfSharp.Pdf.Advanced.PdfReference>(secondNote.Elements[0]);
        Assert.Equal(pdf.Pages[1].Reference!.ObjectID, targetPage.ObjectID);
    }

    [Fact]
    public void RightToRepairFootnoteOrdinalRetainsSuperscriptStyle()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        using var document = WordprocessingDocument.Open(
            Path.Combine(root, "samples", "Right to Repair.docx"), false);
        var paragraph = document.MainDocumentPart!.FootnotesPart!.Footnotes!
            .Descendants<Paragraph>().First(candidate => candidate.InnerText.Contains("Minnesota Session Laws"));
        var ordinal = paragraph.Descendants<Run>().Single(run => run.InnerText == "rd");

        Assert.True(new DxpStyleResolver(document).ResolveRunStyle(paragraph, ordinal).Superscript);
    }

    [Fact]
    public void FlowVisitorRaisesFootnoteOrdinal()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var bytes = DxpExport.ExportToBytes(Path.Combine(root, "samples", "Right to Repair.docx"),
            new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var page = pdf.GetPage(7);
        var ordinal = page.Letters.Single(letter => letter.Value == "r" &&
            letter.PointSize is > 5.45 and < 5.65);
        var numeral = page.Letters.Where(letter => letter.Value == "3" &&
                letter.GlyphRectangle.Left < ordinal.GlyphRectangle.Left &&
                Math.Abs(letter.StartBaseLine.Y - ordinal.StartBaseLine.Y) < 4)
            .OrderByDescending(letter => letter.GlyphRectangle.Left).First();

        Assert.InRange(ordinal.PointSize, 5.45, 5.65);
        Assert.InRange(ordinal.StartBaseLine.Y - numeral.StartBaseLine.Y, 2.8, 3.2);
    }



    [Fact]
    public void MigraDocRaisesExplicitlySizedInlineSuperscript()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var document = new MigraDoc.DocumentObjectModel.Document();
        var paragraph = document.AddSection().AddParagraph();
        paragraph.Format.Font.Name = "Arial";
        paragraph.Format.Font.Size = MigraDoc.DocumentObjectModel.Unit.FromPoint(9);
        paragraph.AddText("93");
        var ordinal = paragraph.AddFormattedText("rd");
        ordinal.Superscript = true;
        ordinal.Size = MigraDoc.DocumentObjectModel.Unit.FromPoint(9);
        paragraph.AddText(" Legislature");

        var renderer = new MigraDoc.Rendering.PdfDocumentRenderer { Document = document };
        renderer.RenderDocument();
        using var output = new MemoryStream();
        renderer.PdfDocument.Save(output, false);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(output.ToArray());
        var numeral = pdf.GetPage(1).Letters.First(letter => letter.Value == "3");
        var superscript = pdf.GetPage(1).Letters.First(letter => letter.Value == "r");

        Assert.InRange(superscript.StartBaseLine.Y - numeral.StartBaseLine.Y, 2.5, 4);
    }

    [Fact]
    public void FlowVisitorPreservesNonbreakingSpaceWithoutXmlSpacePreserve()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Right to Repair.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var streams = pdf.Pages[4].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Select(reference => reference.Value)
            .OfType<PdfSharp.Pdf.PdfDictionary>()
            .Select(content => System.Text.Encoding.Latin1.GetString(content.Stream!.UnfilteredValue));
        Assert.Contains(streams, stream => stream.Contains("(:\u00a0) Tj"));
    }

    [Fact]
    public void FlowVisitorRestoresNarrowSpaceOnSplitTableFinalLine()
    {
        var line = "(its) Tj\n14.1992 0 Td\n(operation,) Tj\n";
        var commands = "BT\n77.65 694.084 Td\n" + line +
            "ET\nBT\n311.4 694.084 Td\n" + line + "ET";

        var revised = DxpFlowPdfVisitor.JustifyFinalTextLine(commands,
            ["its", "operation,"], 0.82, 311.4, 694.084);

        Assert.Contains("77.65 694.084 Td\n" + line, revised);
        Assert.Contains("311.4 694.084 Td\n(its) Tj\n15.0192 0 Td\n(operation,) Tj", revised);
    }


    [Fact]
    public void FlowVisitorPlacesOpeningTableBorderBesideHeaderText()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Right to Repair.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPages().First(candidate => candidate.GetWords().Any(word => word.Text == "Before") &&
            candidate.GetWords().Any(word => word.Text == "After"));
        var before = page.GetWords().First(word => word.Text == "Before");
        var border = page.Paths.Where(path => path.IsStroked &&
                path.GetBoundingRectangle() is { } bounds && bounds.Width is > 230 and < 235 &&
                Math.Abs(bounds.Left - 72) < 1)
            .ToArray();

        var topBorder = Assert.Single(border).GetBoundingRectangle()!.Value;
        Assert.InRange(topBorder.Top - before.BoundingBox.Top, 7, 11);
    }

    [Fact]
    public void FlowVisitorConnectsMatchingParagraphSideBorders()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Right to Repair.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var streams = pdf.Pages[0].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Select(reference => reference.Value)
            .OfType<PdfSharp.Pdf.Advanced.PdfContent>()
            .Where(content => content.Stream != null)
            .Select(content => System.Text.Encoding.Latin1.GetString(content.Stream!.UnfilteredValue))
            .ToArray();
        foreach (var x in new[] { "66.25", "545.75", "102.25", "539.75" })
            Assert.True(streams.Sum(stream => System.Text.RegularExpressions.Regex.Matches(stream,
                $@"{x} [\d.]+ m\s+{x} [\d.]+ l\s+S").Count) >= 3);
        var joinedBorders = streams.SelectMany(stream =>
            System.Text.RegularExpressions.Regex.Matches(stream,
                    @"66\.25 ([\d.]+) m\s+66\.25 ([\d.]+) l\s+S")
                .Select(match => Math.Abs(double.Parse(match.Groups[1].Value,
                    System.Globalization.CultureInfo.InvariantCulture) -
                    double.Parse(match.Groups[2].Value,
                        System.Globalization.CultureInfo.InvariantCulture)))).ToArray();
        Assert.Equal(2, joinedBorders.Count(length => Math.Abs(length - 6) < 0.01));
    }

    [Fact]
    public void FlowVisitorPreservesFootnoteBodies()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Right to Repair.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var text = string.Join(" ", pdf.GetPages().Select(page => page.Text));
        var normalized = System.Text.RegularExpressions.Regex.Replace(text, @"\s+", "");
        Assert.Contains("CopyrightModernizationAct", normalized, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("NintendoofAmerica", normalized, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("LEGISinfo:C-59", normalized, StringComparison.OrdinalIgnoreCase);

        using var document = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var destinations = document.Internals.Catalog.Names.NameTree;
        Assert.True(destinations.ContainsName("dxp-note-13", true));
        Assert.False(destinations.ContainsName("dxp-note-14", true));
        Assert.False(destinations.ContainsName("dxp-note-15", true));
        Assert.True(destinations.ContainsName("dxp-note-16", true));
        var noteDestination = Assert.IsType<PdfSharp.Pdf.PdfArray>(destinations.GetValue("dxp-note-16", true));
        var noteTop = Assert.IsType<PdfSharp.Pdf.PdfReal>(noteDestination.Elements[3]).Value;
        Assert.InRange(noteTop, 640, 660);
        var secondNote = Assert.IsType<PdfSharp.Pdf.PdfArray>(destinations.GetValue("dxp-note-2", true));
        var thirdNote = Assert.IsType<PdfSharp.Pdf.PdfArray>(destinations.GetValue("dxp-note-3", true));
        var secondY = Assert.IsType<PdfSharp.Pdf.PdfReal>(secondNote.Elements[3]).Value;
        var thirdY = Assert.IsType<PdfSharp.Pdf.PdfReal>(thirdNote.Elements[3]).Value;
        Assert.InRange(thirdY - secondY, 20, 22.5);
        Assert.False(destinations.ContainsName("dxp-note-28", true));
        Assert.True(destinations.ContainsName("dxp-note-29", true));
    }

    [Fact]
    public void FlowVisitorLinksEveryVisibleRightToRepairFootnote()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Right to Repair.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var links = pdf.Pages.Cast<PdfSharp.Pdf.PdfPage>()
            .SelectMany(page => page.Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>())
            .Select(annotation => annotation.Elements.GetDictionary("/A")?.Elements.GetString("/D"))
            .Where(name => name?.StartsWith("dxp-note-", StringComparison.Ordinal) == true)
            .ToArray();
        Assert.Equal(26, links.Length);
        foreach (var index in Enumerable.Range(1, 13).Concat(Enumerable.Range(16, 12)).Append(29))
        {
            Assert.Contains($"dxp-note-{index}", links);
            Assert.True(pdf.Internals.Catalog.Names.NameTree.ContainsName($"dxp-note-{index}", true));
        }
    }

    [Fact]
    public void FlowVisitorTrimsTrailingWhitespaceWithoutXmlPreserve()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "with_breaking_style.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var superscript = pdf.GetPage(1).Letters.Single(letter => letter.Value == "S");
        Assert.InRange(superscript.GlyphRectangle.Left, 129, 131);
    }

    [Fact]
    public void LargeSuperscriptRiseRestoresTextStateForFollowingRuns()
    {
        const string stream = "BT\n/F0 20 Tf\n(This) Tj\n/F0 11.6 Tf\n(Super) Tj\n(script) Tj\n(plain) Tj\nET";
        var rewritten = DxpFlowPdfVisitor.RewriteLargeSuperscriptTextStream(stream,
            new HashSet<string> { "Super", "script" });

        Assert.Contains("/F0 12 Tf\n7 Ts\n(Super) Tj\n(script) Tj", rewritten);
        Assert.Contains("0 Ts\n/F0 11.6 Tf\n(plain) Tj", rewritten);
        Assert.Null(DxpFlowPdfVisitor.RewriteLargeSuperscriptTextStream(stream,
            new HashSet<string> { "different" }));
    }

    [Fact]
    public void FlowVisitorKeepsRepairFootnotesNearTheirReferences()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Right to Repair.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(8, pdf.NumberOfPages);
        Assert.Contains("WIPOCopyrightTreaty", pdf.GetPage(3).Text);
        Assert.Contains("technologicalprotectionmeasure", pdf.GetPage(4).Text);
        Assert.Contains("CharlesBernard", pdf.GetPage(5).Text);
        Assert.Contains("CatherineLovrics", pdf.GetPage(5).Text);
        Assert.Contains("pp.142andfollowing", pdf.GetPage(6).Text);
        Assert.Contains("CopyrightOffice", pdf.GetPage(7).Text);
    }

    [Fact]
    public void FlowVisitorDoesNotApplyInheritedBeforeSpacingToSingleFooterParagraph()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordThreeLevelStyleOverrides.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var footer = page.GetWords().First(word => word.Text == "Derived" &&
            page.Height - word.BoundingBox.Top > 700);
        Assert.InRange(page.Height - footer.BoundingBox.Top, 727, 735);
    }

    [Fact]
    public void FlowVisitorPositionsMarginAnchoredHeaderImageWithBody()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "DocKnownGaps",
            "WordFloatingHeaderWrapAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var body = page.GetWords().First(word => word.Text == "Floating");
        Assert.InRange(page.Height - body.BoundingBox.Top, 79, 83);
        var continuation = page.GetWords().First(word => word.Text == "paragraph.");
        Assert.InRange(continuation.BoundingBox.Left, 230, 233);
        var headerContinuation = page.GetWords().First(word => word.Text == "continue");
        Assert.InRange(headerContinuation.BoundingBox.Left, 230, 233);
        var images = page.GetImages()
            .Where(image => image.WidthInSamples > 1 && image.HeightInSamples > 1)
            .OrderByDescending(image => image.Bounds.Bottom).ToArray();
        Assert.Equal(3, images.Length);
        Assert.InRange(page.Height - images[0].Bounds.Top, 53, 56);
        Assert.InRange(page.Height - images[1].Bounds.Top, 179, 182);
        Assert.InRange(page.Height - images[2].Bounds.Top, 763, 766);
    }

    [Fact]
    public void FlowVisitorAppliesInheritedTitleCharacterSpacing()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Right to Repair.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var titleLetters = pdf.GetPage(1).Letters.Where(letter => letter.PointSize > 20).ToArray();
        var titleWidth = titleLetters.Max(letter => letter.EndBaseLine.X) -
            titleLetters.Min(letter => letter.StartBaseLine.X);
        Assert.InRange(titleWidth, 399, 405);
    }

    [Fact]
    public void FlowVisitorMirrorsInheritedIndentsLikeWord()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordHyphenationCenterBottomAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains(pdf.GetPage(1).GetWords(), word => word.Text == "Mirrored" &&
            word.BoundingBox.Left is > 143 and < 145);
        Assert.Contains(pdf.GetPage(2).GetWords(), word => word.Text == "Second-page" &&
            word.BoundingBox.Left is > 107 and < 109);
    }

    [Theory]
    [InlineData("WordDerivedFromUnbasedDefaultsAllStories.docx")]
    [InlineData("WordUnbasedParagraphDefaultsAgainstNormalAllStories.docx")]
    public void FlowVisitorUsesDocumentSpacingForUnbasedParagraphStyle(string fixture)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixture);

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var reset = page.GetWords().First(word => word.Text == "Reset");
        var footer = page.GetWords().First(word => word.Text == "footer");
        Assert.InRange(page.Height - reset.BoundingBox.Top, 95, 100);
        Assert.InRange(page.Height - footer.BoundingBox.Top, 732, 737);
    }

    [Fact]
    public void FlowVisitorPositionsMixedSizeCenterAndBottomParagraphsLikeWord()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordHyphenationCenterBottomAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var first = pdf.GetPage(1);
        var tallTops = first.GetWords().Where(word => word.Text == "Tall")
            .Select(word => first.Height - word.BoundingBox.Top).OrderBy(top => top).ToArray();
        Assert.Equal(4, tallTops.Length);
        Assert.InRange(tallTops[1] - tallTops[0], 39, 41);
        Assert.InRange(tallTops[2] - tallTops[1], 39, 41);
        Assert.InRange(tallTops[3], 720, 723);
        var tall = first.GetWords().First(word => word.Text == "Tall");
        Assert.InRange(tall.BoundingBox.Width, 35.5, 36.5);
        var mirrored = first.GetWords().First(word => word.Text == "Mirrored" &&
            first.Height - word.BoundingBox.Top is > 70 and < 100);
        Assert.InRange(first.Height - mirrored.BoundingBox.Top, 85.0, 86.2);
        var reset = first.GetWords().First(word => word.Text == "Reset");
        Assert.InRange(first.Height - reset.BoundingBox.Top, 130.0, 132.0);
    }


    [Fact]
    public void FlowVisitorOmitsDeletedOnlyListParagraphs()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "TestLists.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(5);
        Assert.DoesNotContain("Delete8", page.Text);
        Assert.Contains("Original9", page.Text);
        var lastOriginalTop = page.Letters.Where(letter => letter.Value == "O")
            .Max(letter => page.Height - letter.GlyphRectangle.Top);
        Assert.InRange(lastOriginalTop, 168, 173);
    }

    [Fact]
    public void FlowVisitorHidesOccupiedHorizontalMergeContinuations()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordOccupiedThreeCellHMergeArialAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var text = pdf.GetPage(1).Text;
        Assert.Contains("HEADER-START", text);
        Assert.Contains("HEADER-LEFT", text);
        Assert.DoesNotContain("CONTINUATION", text);
        Assert.DoesNotContain("MORE", text);
        var headerLetters = pdf.GetPage(1).Letters.Where(letter => letter.Value == "H" &&
            pdf.GetPage(1).Height - letter.GlyphRectangle.Top is > 55 and < 85)
            .Select(letter => letter.GlyphRectangle.Left).ToArray();
        Assert.Contains(headerLetters, left => left is > 155 and < 158);
        Assert.Contains(headerLetters, left => left is > 248 and < 251);
    }

    [Theory]
    [InlineData("WordOccupiedThreeCellHMergeAllStories", 109, 113)]
    [InlineData("WordOccupiedThreeCellHMergeArialAllStories", 110, 113)]
    [InlineData("WordOccupiedHMergeDocDefaultsAllStories", 110, 113)]
    public void FlowVisitorReservesOccupiedHeaderMergeClearance(string name,
        double minimum, double maximum)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var body = page.GetWords().Single(word => word.Text == "START");
        Assert.InRange(page.Height - body.BoundingBox.Top, minimum, maximum);
    }

    [Theory]
    [InlineData("WordOccupiedHMergeAllStories")]
    [InlineData("WordOccupiedHMergeArialAllStories")]
    [InlineData("WordOccupiedHMergeNoWrapOffAllStories")]
    [InlineData("WordOccupiedHMergeDocDefaultsAllStories")]
    public void FlowVisitorHonorsImplicitHorizontalMergeCellInset(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var first = pdf.GetPage(1).GetWords().First(word => word.Text.Contains("START"));
        Assert.InRange(first.BoundingBox.Left, 72.4, 72.6);
    }

    [Theory]
    [InlineData("WordOccupiedHMergeAllStories", 110, 113)]
    [InlineData("WordOccupiedHMergeArialAllStories", 113, 116)]
    public void FlowVisitorLeavesCellClearanceInNarrowHorizontalMergeRows(string name,
        double minimum, double maximum)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var rightCell = page.Letters.Single(letter => letter.Value == "R" &&
            page.Height - letter.GlyphRectangle.Top is > 130 and < 155);
        Assert.InRange(rightCell.GlyphRectangle.Left, minimum, maximum);
    }

    [Theory]
    [InlineData("WordOccupiedHMergeAllStories")]
    [InlineData("WordOccupiedHMergeExplicitMarginsAllStories")]
    [InlineData("WordOccupiedThreeCellHMergeAllStories")]
    public void FlowVisitorPositionsAptosHorizontalMergeFooter(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var footer = page.GetWords().Single(word => word.Text == "FOOTER-START");
        Assert.InRange(page.Height - footer.BoundingBox.Top, 686.5, 688);
    }

    [Theory]
    [InlineData("WordOccupiedHMergeDocDefaultsAllStories", 695.38, 719.38)]
    [InlineData("WordOccupiedHMergeThemeDefaultsAllStories", 695.38, 719.38)]
    [InlineData("WordOccupiedHMergeArialAllStories", 694.42, 718.42)]
    public void FlowVisitorAlignsHorizontalMergeFooterWithDocumentDefaults(
        string name, double firstBaseline, double secondBaseline)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var footerBaselines = page.Letters.Where(letter => letter.Value == "F" &&
                page.Height - letter.StartBaseLine.Y > 680)
            .Select(letter => page.Height - letter.StartBaseLine.Y).Distinct().OrderBy(value => value).ToArray();
        Assert.Equal(2, footerBaselines.Length);
        Assert.InRange(footerBaselines[0], firstBaseline - 0.1, firstBaseline + 0.1);
        Assert.InRange(footerBaselines[1], secondBaseline - 0.1, secondBaseline + 0.1);
    }

    [Fact]
    public void FlowVisitorPositionsFirstArialHeadingAtSectionTop()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordNextStyleStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var flow = page.GetWords().Where(word => word.Text is "Flow" &&
                word.BoundingBox.Left is > 70 and < 75)
            .OrderByDescending(word => word.BoundingBox.Bottom).ToArray();
        Assert.Equal(2, flow.Length);
        Assert.InRange(page.Height - flow[0].Letters[0].StartBaseLine.Y, 93.8, 94.4);
        Assert.InRange(page.Height - flow[1].Letters[0].StartBaseLine.Y, 116.7, 117.3);
    }

    [Theory]
    [InlineData("WordVisibleRepeatingHeaderTable", 87, 93)]
    [InlineData("WordConditionalFirstRowParagraphAllStories", 82, 84)]
    public void FlowVisitorReservesHeaderTableClearance(string name, double minimum, double maximum)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var body = page.Letters.First(letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left is > 70 and < 75);
        Assert.InRange(page.Height - body.GlyphRectangle.Top, minimum, maximum);
    }

    [Theory]
    [InlineData("WordVisibleRepeatingHeaderTable")]
    [InlineData("WordVisibleRowHeightsAllStories")]
    public void FlowVisitorReservesExactHeightHeaderRowClearance(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var body = page.Letters.First(letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left is > 70 and < 75);
        Assert.InRange(page.Height - body.GlyphRectangle.Top, 92.7, 93.1);
    }

    [Fact]
    public void FlowVisitorReservesMinimumHeightHeaderRow()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRowLayoutAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var body = page.Letters.First(letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left is > 70 and < 75);
        Assert.InRange(page.Height - body.GlyphRectangle.Top, 93, 95);
    }

    [Fact]
    public void FlowVisitorUsesCompactHeaderRowWithParagraphOnlyDefaults()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleParagraphOnlyDefaultsAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var body = page.Letters.First(letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left is > 70 and < 75);
        Assert.InRange(page.Height - body.GlyphRectangle.Top, 82, 84);
    }

    [Fact]
    public void FlowVisitorPlacesEmptyModernDefaultsBodyAtWordOrigin()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleEmptyDefaultsAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var body = page.Letters.First(letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left is > 71 and < 73);
        Assert.InRange(page.Height - body.StartBaseLine.Y, 82.1, 82.4);
    }

    [Theory]
    [InlineData("WordMirrorIndentsAllStories", "Mirrored footer")]
    [InlineData("WordLinkedStylesAllStories", "Footer linked")]
    public void FlowVisitorPreservesImplicitSpacingOfSimpleUnbasedStyle(string name, string footerText)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var footer = page.Letters.First(letter => letter.Value == footerText[..1] &&
            letter.GlyphRectangle.Top < 65);
        Assert.InRange(page.Height - footer.GlyphRectangle.Top, 730, 737);
    }

    [Fact]
    public void FlowVisitorUsesWordMirrorIndentPageParity()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordMirrorIndentsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var odd = pdf.GetPage(1);
        var even = pdf.GetPage(2);
        Assert.Contains(odd.Letters, letter => letter.Value == "M" &&
            letter.GlyphRectangle.Left is > 143 and < 145 &&
            odd.Height - letter.GlyphRectangle.Top is > 70 and < 75);
        Assert.Contains(even.Letters, letter => letter.Value == "S" &&
            letter.GlyphRectangle.Left is > 107 and < 109 &&
            even.Height - letter.GlyphRectangle.Top is > 70 and < 75);
    }

    [Fact]
    public void FlowVisitorKeepsShadedBorderedTextAtNaturalBaseline()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyledParagraphShadingResetAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.Letters, letter => letter.Value == "R" &&
            letter.GlyphRectangle.Left is > 109 and < 115 &&
            page.Height - letter.GlyphRectangle.Top is > 104 and < 115);
    }

    [Theory]
    [InlineData("WordSymbolGlyphAllStories", "F")]
    [InlineData("WordLayeredAutomaticColorAllStories", "F")]
    [InlineData("WordAdjacentHighlightNoneAllStories", "M")]
    [InlineData("WordCenteredPositionalTabAllStories", "F")]
    public void FlowVisitorPositionsFooterWithEmptyNormalStyle(string name, string firstLetter)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var footer = page.Letters.First(letter => letter.Value == firstLetter &&
            page.Height - letter.GlyphRectangle.Top > 700);
        Assert.InRange(page.Height - footer.GlyphRectangle.Top, 730, 737);
    }

    [Theory]
    [InlineData("WordMixedRotatedCellsAllStories")]
    [InlineData("WordRotatedTableCellsAllStories")]
    public void FlowVisitorRotatesBorderlessTableCellText(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.Letters, letter => letter.Value == "B" &&
            Math.Abs(letter.StartBaseLine.Y - letter.EndBaseLine.Y) >
            2 * Math.Abs(letter.StartBaseLine.X - letter.EndBaseLine.X));
        Assert.Contains(page.Letters, letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left is > 70 and < 75 &&
            page.Height - letter.GlyphRectangle.Top is > 82 and < 86);
        Assert.DoesNotContain(page.Letters, letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left is > 75 and < 100 &&
            page.Height - letter.GlyphRectangle.Top is > 45 and < 70 &&
            Math.Abs(letter.StartBaseLine.X - letter.EndBaseLine.X) >
            2 * Math.Abs(letter.StartBaseLine.Y - letter.EndBaseLine.Y));
    }

    [Theory]
    [InlineData("WordRotatedTableCellsAllStories", "(H) Tj", "(r A) Tj")]
    [InlineData("WordMixedRotatedCellsAllStories", "(He) Tj", "(er ) Tj")]
    public void FlowVisitorWrapsRotatedCellTextLikeWord(string name, string firstChunk, string lastChunk)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var commands = string.Concat(pdf.Pages[0].Contents.Elements.Select(element =>
            (element as PdfSharp.Pdf.Advanced.PdfReference)?.Value is
                PdfSharp.Pdf.Advanced.PdfContent content && content.Stream != null
                ? System.Text.Encoding.Latin1.GetString(content.Stream.UnfilteredValue) : ""));
        Assert.Contains(firstChunk, commands);
        Assert.Contains(lastChunk, commands);
    }

    [Fact]
    public void FlowVisitorPreservesOppositeTableCellTextDirections()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordMixedRotatedCellsAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var bodyLetters = page.Letters.Where(letter => letter.Value == "B" &&
            page.Height - letter.GlyphRectangle.Top is > 100 and < 155).ToArray();
        Assert.Contains(bodyLetters, letter => letter.GlyphRectangle.Left is > 280 and < 305 &&
            letter.StartBaseLine.Y > letter.EndBaseLine.Y);
        Assert.Contains(bodyLetters, letter => letter.GlyphRectangle.Left is > 305 and < 345 &&
            letter.StartBaseLine.Y < letter.EndBaseLine.Y);
        Assert.Contains(bodyLetters, letter => letter.GlyphRectangle.Left is > 515 and < 540 &&
            letter.StartBaseLine.Y > letter.EndBaseLine.Y);
    }

    [Fact]
    public void FlowVisitorPositionsGridGapFooterTable()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleAutoRowGridGapsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var footer = page.Letters.First(letter => letter.Value == "F" &&
            letter.GlyphRectangle.Left is > 71 and < 73 &&
            page.Height - letter.GlyphRectangle.Top > 600);
        Assert.InRange(page.Height - footer.GlyphRectangle.Top, 683, 689);
    }


    [Fact]
    public void FlowVisitorFirstParagraphSpacingDoesNotShiftLaterPages()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "AutoNumberCompatibility.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        var secondPage = pdf.GetPage(2);
        var firstGlyphTop = secondPage.Height - secondPage.Letters.Max(letter => letter.GlyphRectangle.Top);
        Assert.InRange(firstGlyphTop, 68, 78);
    }

    [Fact]
    public void FlowVisitorUsesWordCachedThemeShadeWhenItMatchesTheTheme()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "AutoNumberCompatibility.docx");
        using var document = WordprocessingDocument.Open(source, false);
        var paragraph = document.MainDocumentPart!.Document.Body!.Descendants<Paragraph>()
            .First(candidate => candidate.Descendants<Run>().Any());
        var run = paragraph.Descendants<Run>().First();

        Assert.Equal("#0F4761", new DxpStyleResolver(document).ResolveRunStyle(paragraph, run).ColorCss);
    }

    [Theory]
    [InlineData("WordTableBackgroundAutoStyleAllStories", "#FFD700")]
    [InlineData("WordTableBackgroundAutoDirectAllStories", "#00B0F0")]
    public void FlowVisitorResolvesTableShadingInAllStories(string name, string expectedColor)
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "DocNativeOracle", name + ".docx");
        using var document = WordprocessingDocument.Open(source, false);
        var resolver = new DxpTableStyleResolver(document);
        var body = document.MainDocumentPart!.Document.Body!;
        var header = document.MainDocumentPart.HeaderParts.Single().Header;
        var footer = document.MainDocumentPart.FooterParts.Single().Footer;
        foreach (var table in new[] { body.Descendants<Table>().Single(),
            header.Descendants<Table>().Single(), footer.Descendants<Table>().Single() })
        {
            var style = resolver.ResolveTableStyle(table.TableProperties, 1, 2);
            Assert.Equal(expectedColor, DxpTableStyleResolver.ComputeCellStyle(
                style, table.Descendants<TableCell>().First().TableCellProperties, 0, 0, 1, 1).BackgroundColorCss);
        }

        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var secondHeaderCell = pdf.GetPage(1).Letters.Where(letter => letter.Value == "H" &&
            letter.GlyphRectangle.Left > 130).Min(letter => letter.GlyphRectangle.Left);
        Assert.InRange(secondHeaderCell, 143, 146);
    }

    [Theory]
    [InlineData("WordLogicalTableBordersAllStories", false)]
    [InlineData("WordStyledLogicalTableBordersAllStories", false)]
    [InlineData("WordStyleDirectLogicalTableBordersAllStories", false)]
    [InlineData("WordStyleDirectLogicalTableBordersAllStories", true)]
    public void LogicalTableBordersResolveByVisualDirection(string name, bool rightToLeft)
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        using var document = WordprocessingDocument.Open(source, false);
        var table = document.MainDocumentPart!.Document.Body!.Descendants<Table>().First();
        if (rightToLeft) table.TableProperties!.PrependChild(new BiDiVisual());

        var resolved = new DxpTableStyleResolver(document).ResolveTableStyle(table.TableProperties, 1, 2);
        var borders = resolved.ComputedTableStyle.TableBorders!;
        Assert.Equal("#FF0000", (rightToLeft ? borders.Right : borders.Left)?.ColorCss);
        if (name == "WordStyleDirectLogicalTableBordersAllStories")
            Assert.Equal("#008000", (rightToLeft ? borders.Left : borders.Right)?.ColorCss);
    }

    [Fact]
    public void LogicalCellBordersAndVisualOrderRespectRightToLeftTable()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordLogicalCellBordersAllStories.docx");
        using var document = WordprocessingDocument.Open(source, false);
        var headerTable = document.MainDocumentPart!.HeaderParts
            .SelectMany(part => part.Header.Descendants<Table>()).Single();
        var resolver = new DxpTableStyleResolver(document);
        var resolved = resolver.ResolveTableStyle(headerTable.TableProperties, 1, 2);
        Assert.True(resolved.RightToLeft);
        var firstCell = headerTable.Elements<TableRow>().First().Elements<TableCell>().First();
        var cellStyle = DxpTableStyleResolver.ComputeCellStyle(resolved, firstCell.TableCellProperties,
            0, 0, 1, 1);
        Assert.Equal("#008000", cellStyle.Borders?.Left?.ColorCss);

        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var headerLetters = pdf.GetPage(1).Letters.Where(letter => letter.Value is "A" or "B" &&
            letter.GlyphRectangle.Bottom > 690).ToArray();
        Assert.True(headerLetters.Any(letter => letter.Value == "B" && letter.GlyphRectangle.Left < 150));
        Assert.True(headerLetters.Any(letter => letter.Value == "A" && letter.GlyphRectangle.Left > 300));
        var bodyWords = pdf.GetPage(1).GetWords().Where(word => word.Text == "Body").ToArray();
        Assert.Contains(bodyWords, word => word.BoundingBox.Left is > 77.7 and < 78.6);
        Assert.Contains(bodyWords, word => word.BoundingBox.Left is > 311.8 and < 312.6);
    }

    [Fact]
    public void FlowVisitorDrawsDoubleBorderAtVisualLeftOfRightToLeftHeaderTable()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleCustomTableBorderColors.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var headerBorders = page.Paths.Where(path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds &&
            page.Height - bounds.Top is > 48 and < 68 && bounds.Height > 12).ToArray();
        Assert.Equal(2, headerBorders.Length);
        Assert.All(headerBorders, path =>
            Assert.InRange(path.GetBoundingRectangle()!.Value.Left, 69, 74));
    }

    [Fact]
    public void FlowVisitorDrawsDoubleBorderAtVisualRightOfTableStories()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyleDirectLogicalTableBordersAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var borders = page.Paths.Where(path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 538 and < 543 && bounds.Height > 12).ToArray();
        Assert.Equal(6, borders.Length);
    }

    [Fact]
    public void FlowVisitorHonorsParagraphBorderSpaceAroundShading()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordParagraphBordersShadingStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var wideShades = pdf.GetPage(1).Paths.Where(path => path.IsFilled &&
            path.GetBoundingRectangle() is { } bounds && bounds.Width > 465).ToArray();
        Assert.True(wideShades.Length >= 3);
        Assert.All(wideShades, path =>
        {
            var bounds = path.GetBoundingRectangle()!.Value;
            Assert.InRange(bounds.Left, 70.3, 70.8);
            Assert.InRange(bounds.Width, 470.8, 471.2);
        });
        var page = pdf.GetPage(1);
        var headerLast = page.GetWords().Single(word => word.Text == "last" &&
            page.Height - word.BoundingBox.Top < 100);
        var footerFirst = page.GetWords().Single(word => word.Text == "first" &&
            page.Height - word.BoundingBox.Top > 700);
        Assert.InRange(page.Height - headerLast.BoundingBox.Top, 68.2, 69);
        Assert.InRange(page.Height - footerFirst.BoundingBox.Top, 715.8, 716.7);
        var bodyFirst = page.GetWords().Single(word => word.Text == "First" &&
            page.Height - word.BoundingBox.Top is > 80 and < 120);
        var bodyLast = page.GetWords().Single(word => word.Text == "Last" &&
            page.Height - word.BoundingBox.Top is > 120 and < 180);
        Assert.InRange(page.Height - bodyFirst.BoundingBox.Top, 85.4, 86.2);
        Assert.InRange(page.Height - bodyLast.BoundingBox.Top, 133.2, 133.9);
    }

    [Theory]
    [InlineData("WordStyledParagraphDecorations.docx")]
    [InlineData("WordStyledDistributedParagraphAllStories.docx")]
    [InlineData("WordDistributedStyleInheritanceAllStories.docx")]
    public void FlowVisitorAlignsFirstStyleInheritedBorderedBodyParagraph(string fixture)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixture);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var bodyFirst = page.GetWords().First(word => word.Text == "First" &&
            page.Height - word.BoundingBox.Top is > 85 and < 120);
        Assert.InRange(page.Height - bodyFirst.BoundingBox.Top, 95.8, 96.3);
    }

    [Theory]
    [InlineData("WordStyledParagraphDecorations.docx")]
    [InlineData("WordStyledDistributedParagraphAllStories.docx")]
    [InlineData("WordDistributedStyleInheritanceAllStories.docx")]
    public void FlowVisitorExtendsInheritedShadingToBorderDistance(string fixture)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixture);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var commands = string.Join('\n', pdf.Pages[0].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Select(reference => reference.Value)
            .OfType<PdfSharp.Pdf.Advanced.PdfContent>()
            .Select(content => System.Text.Encoding.Latin1.GetString(content.Stream!.UnfilteredValue)));
        Assert.Equal(3, System.Text.RegularExpressions.Regex.Matches(commands,
            @"70\.5 \d+(?:\.\d+)? 471 \d+(?:\.\d+)? re").Count);
    }

    [Theory]
    [InlineData("WordInheritedStyledRotatedTables.docx")]
    [InlineData("WordInheritedTableMixedContent.docx")]
    [InlineData("WordInheritedTablePageFields.docx")]
    [InlineData("WordSixSlotInheritedStyledTables.docx")]
    public void FlowVisitorUsesFirstPageHeaderClearanceForSectionBody(string fixture)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixture);
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(2);
        var firstWord = page.GetWords().First(word => word.Text == "Second");
        Assert.InRange(page.Height - firstWord.BoundingBox.Top, 74, 75);
        var footer = page.GetWords().First(word => word.Text == "First-page" &&
            page.Height - word.BoundingBox.Top > 500);
        Assert.InRange(page.Height - footer.Letters[0].StartBaseLine.Y, 570, 571);
    }

    [Fact]
    public void FlowVisitorJoinsRightToRepairParagraphSideBordersAcrossParagraphSpacing()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Right to Repair.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var edges = page.Paths.Where(path => path.IsStroked &&
                path.GetBoundingRectangle() is { } bounds &&
                bounds.Left is > 66 and < 67 && bounds.Height > 2 &&
                page.Height - bounds.Top is > 230 and < 530)
            .Select(path => path.GetBoundingRectangle()!.Value)
            .OrderByDescending(bounds => bounds.Top).ToArray();
        Assert.True(edges.Length >= 3);
        for (var index = 1; index < edges.Length; index++)
            Assert.True(edges[index - 1].Bottom - edges[index].Top < 1,
                "The paragraph side border has a visible gap between adjacent paragraphs.");
        var nextBody = pdf.GetPage(2).GetWords().First(word => word.Text == "Under");
        Assert.InRange(pdf.GetPage(2).Height - nextBody.BoundingBox.Top, 240, 245);
    }


    [Fact]
    public void FlowVisitorKeepsAdjacentBorderedBodyParagraphSpacing()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Right to Repair.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var questions = page.GetWords().Single(word => word.Text == "questions");
        var adoption = page.GetWords().First(word => word.Text == "adoption");
        Assert.InRange(page.Height - questions.BoundingBox.Top, 376, 378);
        Assert.InRange(questions.BoundingBox.Top - adoption.BoundingBox.Top, 103, 105);
    }

    [Theory]
    [InlineData("WordRussianLowerListsAllStories", "а.", "б.", "в.")]
    [InlineData("WordRussianUpperListsAllStories", "А.", "Б.", "В.")]
    public void FlowVisitorRendersRussianListMarkersInAllStories(
        string name, string first, string second, string third)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var words = pdf.GetPage(1).GetWords().Select(word => word.Text).ToArray();
        Assert.Equal(3, words.Count(word => word == first));
        Assert.Equal(3, words.Count(word => word == second));
        Assert.Contains(third, words);
    }

    [Theory]
    [InlineData(DxpFieldEvalExportMode.Cache)]
    [InlineData(DxpFieldEvalExportMode.Evaluate)]
    public void FlowVisitorResolvesHeaderStyleReferenceOnEachPage(DxpFieldEvalExportMode mode)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyleReferenceAcrossPages.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(),
            new DxpExportOptions { FieldEvalMode = mode });

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("Heading:Openexamplesite", pdf.GetPage(1).Text);
        Assert.Contains("Heading:Secondheading", pdf.GetPage(2).Text);
        Assert.DoesNotContain("Heading:Secondheading", pdf.GetPage(1).Text);
    }

    [Fact]
    public void FlowVisitorRendersNoBreakHyphenInFooter()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleTopGutterAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Contains("Footer-joinsoft", pdf.GetPage(1).Text);
    }

    [Fact]
    public void FlowVisitorKeepsNoBreakHyphenWithBothSidesAtLineBoundary()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleTopGutterAllStories.docx");
        using var stream = new MemoryStream();
        stream.Write(File.ReadAllBytes(source));
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var footer = document.MainDocumentPart!.FooterParts
                .Single(part => part.Footer.Descendants<NoBreakHyphen>().Any()).Footer;
            var prefix = footer.Descendants<Text>().Single(text => text.Text == "Footer");
            prefix.Space = SpaceProcessingModeValues.Preserve;
            prefix.Text = string.Concat(Enumerable.Repeat("The quick brown fox ", 3)) + "a Footer";
            footer.Save();
        }
        var bytes = DxpExport.ExportToBytes(stream.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var words = pdf.GetPage(1).GetWords().ToArray();
        Assert.Contains(words, word => word.Text.StartsWith("Footer-join", StringComparison.Ordinal));
        Assert.DoesNotContain(words, word => word.Text == "Footer-");
    }

    [Fact]
    public void FlowVisitorShowsSoftHyphenAtFooterLineBreak()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleTopGutterAllStories.docx");
        using var stream = new MemoryStream();
        stream.Write(File.ReadAllBytes(source));
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var footer = document.MainDocumentPart!.FooterParts
                .Single(part => part.Footer.Descendants<SoftHyphen>().Any()).Footer;
            var prefix = footer.Descendants<Text>().Single(text => text.Text == "Footer");
            prefix.Space = SpaceProcessingModeValues.Preserve;
            prefix.Text = string.Concat(Enumerable.Repeat("The quick brown fox ", 3)) + "Footer";
            footer.Save();
        }
        var bytes = DxpExport.ExportToBytes(stream.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var words = page.GetWords().ToArray();
        var first = Assert.Single(words.Where(word => word.Text == "Footer-join-"));
        var second = Assert.Single(words.Where(word => word.Text == "soft"));
        Assert.True(first.BoundingBox.Bottom > second.BoundingBox.Bottom);
    }

    [Fact]
    public void FlowVisitorKeepsRightToRepairBodyBeforePageSevenFootnotes()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Right to Repair.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(8, pdf.NumberOfPages);
        Assert.Contains(pdf.GetPage(6).GetWords(), word => word.Text == "United" &&
            word.Letters[0].PointSize > 10);
        Assert.Contains(pdf.GetPage(7).GetWords(), word => word.Text == "Here");
        Assert.DoesNotContain(pdf.GetPage(8).GetWords(), word => word.Text == "Here");
        Assert.Contains(pdf.GetPage(8).GetWords(), word => word.Text == "diagnosis");
    }

    [Fact]
    public void FlowVisitorKeepsFirstDirectBorderClearanceInRightToRepair()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Right to Repair.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var words = pdf.GetPage(1).GetWords().ToArray();
        var preceding = Assert.Single(words.Where(word => word.Text == "paragraphs:"));
        var bordered = Assert.Single(words.Where(word => word.Text == "Diagnosis,"));
        var gap = preceding.BoundingBox.Top - bordered.BoundingBox.Top;
        Assert.InRange(gap, 31.2, 31.8);
    }

    [Fact]
    public void FlowVisitorKeepsBorderedParagraphExitClearanceAfterPageSplit()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Right to Repair.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(2);
        var words = page.GetWords().ToArray();
        var heading = Assert.Single(words.Where(word => word.Text == "Non-application"));
        var following = Assert.Single(words.Where(word => word.Text == "Under"));
        Assert.InRange(page.Height - heading.BoundingBox.Top, 158.5, 159.5);
        Assert.InRange(page.Height - following.BoundingBox.Top, 243, 244);
    }

    [Fact]
    public void FlowVisitorKeepsTrackedWordFragmentsTogether()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Right to Repair.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Contains(Enumerable.Range(1, pdf.NumberOfPages)
            .SelectMany(page => pdf.GetPage(page).GetWords()), word => word.Text == "excluding");
    }

    [Fact]
    public void FlowVisitorDoesNotInheritNarrowSpaceSizeForTrackedWord()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Right to Repair.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var word = Assert.Single(pdf.GetPage(1).GetWords().Where(item => item.Text == "software"));
        Assert.All(word.Letters, letter => Assert.InRange(letter.PointSize, 11.9, 12.1));
    }

    [Fact]
    public void FlowVisitorOmitsDeletedFootnoteReferenceAndKeepsVisibleNumber()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Right to Repair.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var words = pdf.GetPage(6).GetWords().Select(word => word.Text).ToArray();
        var apple = Array.FindIndex(words, word => word == "Apple");
        Assert.True(apple >= 0);
        var repair = Array.FindIndex(words, apple, word => word.StartsWith("repair", StringComparison.Ordinal));
        Assert.True(repair > apple);
        Assert.Equal("repair", words[repair]);
        Assert.Equal("and,", words[repair + 1]);
        Assert.Contains(words.Skip(repair), word => word == "safety.16" || word == "16");
    }

    [Fact]
    public void PdfUnicodeMapEncodesSupplementaryHanAsUtf16()
    {
        var mapping = "1 beginbfrange\n<01C5><01C5><20000>\nendbfrange";
        var corrected = DocxportNet.Visitors.Pdf.DxpPdfHighlightOverlay.RepairUnicodeMapText(mapping);

        Assert.Equal("1 beginbfrange\n<01C5><01C5><D840DC00>\nendbfrange", corrected);
        Assert.Equal("<0754><0754><D835DC65>",
            DocxportNet.Visitors.Pdf.DxpPdfHighlightOverlay.RepairUnicodeMapText(
                "<0754><0754><1D465>"));
        Assert.Equal("<01C5><0041>",
            DocxportNet.Visitors.Pdf.DxpPdfHighlightOverlay.RepairUnicodeMapText("<01C5><0041>"));
    }

    [Fact]
    public void PdfUnicodeMapUsesStandardBulletForSymbolFont()
    {
        var mapping = "1 beginbfrange\n<0078><0078><F0B7>\nendbfrange";

        Assert.Equal("1 beginbfrange\n<0078><0078><2022>\nendbfrange",
            DocxportNet.Visitors.Pdf.DxpPdfHighlightOverlay.RepairUnicodeMapText(mapping, true));
        Assert.Equal(mapping,
            DocxportNet.Visitors.Pdf.DxpPdfHighlightOverlay.RepairUnicodeMapText(mapping));
    }

    [Fact]
    public void PdfUnicodeMapUsesWingdingsSymbolsForSelectableListMarkers()
    {
        var mapping = "4 beginbfrange\n<0001><0001><F0A7>\n<0002><0002><F0D8>\n<0003><0003><F076>\n<0004><0004><F0FC>\nendbfrange";

        Assert.Equal("4 beginbfrange\n<0001><0001><25AA>\n<0002><0002><27A2>\n<0003><0003><2756>\n<0004><0004><2713>\nendbfrange",
            DocxportNet.Visitors.Pdf.DxpPdfHighlightOverlay.RepairUnicodeMapText(
                mapping, wingdingsFont: true));
    }

    [Theory]
    [InlineData("RevisionFormattingBalloonA4", 1)]
    [InlineData("RevisionFormattingBalloonThirdItem", 1)]
    [InlineData("RevisionFormattingBalloonTwoItems", 2)]
    public void FlowVisitorPrintsFormattingChangeBalloons(string name, int expectedCount)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfFeatures",
            name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(
            includeRevisionMarkup: true, revisionMarkupColor: XColor.FromArgb(0, 120, 212)));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(expectedCount, System.Text.RegularExpressions.Regex.Matches(
            pdf.GetPage(1).Text, "Formatted: English \\(United States\\)").Count);
    }

    [Fact]
    public void MarkupPrintLayoutMovesLinksDestinationsAndCommentQuads()
    {
        using var document = new PdfSharp.Pdf.PdfDocument();
        var page = document.AddPage();
        page.Width = XUnit.FromPoint(612);
        page.Height = XUnit.FromPoint(792);
        using (var graphics = XGraphics.FromPdfPage(page))
            graphics.DrawLine(XPens.Black, 72, 192, 172, 192);
        page.AddWebLink(new PdfSharp.Pdf.PdfRectangle(new XPoint(72, 600),
            new XPoint(172, 620)), "https://example.com");
        document.AddNamedDestination("target", 1,
            PdfSharp.Pdf.Advanced.PdfNamedDestinationParameters.CreatePosition(72, 600));
        var highlight = new PdfSharp.Pdf.Annotations.PdfTextAnnotation(document)
        {
            Rectangle = new PdfSharp.Pdf.PdfRectangle(new XPoint(72, 600),
                new XPoint(172, 620)),
            Contents = "selected text"
        };
        highlight.Elements.SetName("/Subtype", "/Highlight");
        var quads = new PdfSharp.Pdf.PdfArray(document);
        foreach (var coordinate in new double[] { 72, 620, 172, 620, 72, 600, 172, 600 })
            quads.Elements.Add(new PdfSharp.Pdf.PdfReal(coordinate));
        highlight.Elements.SetValue("/QuadPoints", quads);
        page.Annotations.Add(highlight);

        DxpPdfMarkupPrintLayout.Apply(document,
            [Array.Empty<DxpPdfMarkupPrintLayout.Balloon>()]);

        using var stream = new MemoryStream();
        document.Save(stream, false);
        using var result = PdfReader.Open(new MemoryStream(stream.ToArray()), PdfDocumentOpenMode.Import);
        var scale = 612.0 / (612 + 194.4);
        var offset = (1 - scale) * 792 / 2;
        var annotations = result.Pages[0].Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>()
            .ToArray();
        Assert.Equal(2, annotations.Length);
        Assert.All(annotations, annotation =>
        {
            Assert.InRange(annotation.Rectangle.X1, 72 * scale - 0.01, 72 * scale + 0.01);
            Assert.InRange(annotation.Rectangle.Y1, 600 * scale + offset - 0.01,
                600 * scale + offset + 0.01);
        });
        var selected = Assert.Single(annotations.Where(annotation =>
            annotation.Elements.GetName("/Subtype") == "/Highlight"));
        var quad = selected.Elements.GetArray("/QuadPoints")!;
        static double Number(PdfSharp.Pdf.PdfItem item) => item switch
        {
            PdfSharp.Pdf.PdfReal real => real.Value,
            PdfSharp.Pdf.PdfInteger integer => integer.Value,
            _ => throw new InvalidOperationException("Expected PDF number")
        };
        Assert.InRange(Number(quad.Elements[0]),
            72 * scale - 0.01, 72 * scale + 0.01);
        Assert.InRange(Number(quad.Elements[1]),
            620 * scale + offset - 0.01, 620 * scale + offset + 0.01);
        var destination = (PdfSharp.Pdf.PdfArray)result.Internals.Catalog.Names!.NameTree!
            .GetValue("target", true)!;
        var position = destination.Elements[1] is PdfSharp.Pdf.PdfLiteral literal
            ? literal.ToString().Split(' ', StringSplitOptions.RemoveEmptyEntries)
            : null;
        var destinationLeft = position is { Length: >= 3 }
            ? double.Parse(position[1], System.Globalization.CultureInfo.InvariantCulture)
            : Number(destination.Elements[2]);
        var destinationTop = position is { Length: >= 3 }
            ? double.Parse(position[2], System.Globalization.CultureInfo.InvariantCulture)
            : Number(destination.Elements[3]);
        Assert.InRange(destinationLeft,
            72 * scale - 0.01, 72 * scale + 0.01);
        Assert.InRange(destinationTop,
            600 * scale + offset - 0.01, 600 * scale + offset + 0.01);
    }

    [Fact]
    public void FlowVisitorExportsSearchableSymbolBullets()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordMixedListsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var text = pdf.GetPage(1).Text;
        Assert.Contains('•', text);
        Assert.DoesNotContain('\uf0b7', text);
    }

    [Fact]
    public void ContinuedJustifiedPageLineSpreadsWordsWithoutSplittingGlyphRuns()
    {
        var commands = "(Charles) Tj\n40 0 Td\n(Bernard,) Tj\n42 0 Td\n(Association) Tj\n12 0 Td\n(,) Tj\nET";
        var revised = DxpFlowPdfVisitor.JustifyFinalTextLine(commands,
            ["Charles", "Bernard,", "Association,"], 20);

        Assert.Equal("(Charles) Tj\n50 0 Td\n(Bernard,) Tj\n52 0 Td\n(Association) Tj\n12 0 Td\n(,) Tj\nET", revised);
    }

    [Fact]
    public void FlowVisitorJustifiesRightToRepairPageFourContinuation()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Right to Repair.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var words = pdf.GetPage(4).GetWords().ToArray();
        var anchor = words.First(word => word.Text == "means" && word.BoundingBox.Left > 300);
        var continuation = words.Where(word =>
                Math.Abs(word.Letters[0].StartBaseLine.Y - anchor.Letters[0].StartBaseLine.Y) < 1 &&
                word.BoundingBox.Left >= anchor.BoundingBox.Left - 1)
            .OrderBy(word => word.BoundingBox.Left).ToArray();

        Assert.Contains(continuation, word => word.Text.Contains("technology"));
        Assert.InRange(continuation.Last().BoundingBox.Right, 527.5, 529.0);
    }

    [Fact]
    public void FlowVisitorRightAnchorsTrackedAptosKerning()
    {
        var metrics = DocxportNet.Doc.DocSystemFontAdvances.Find("Aptos", includeCloudFonts: true);
        Assert.NotNull(metrics);
        Assert.InRange(metrics.LegacyKerningAdjustment('a', 't', 12), -0.12, -0.09);
        var commands = "BT\n/F1 12 Tf\n100 100 Td\n100 Tz\n0 Tc\n(Gaspard) Tj\n" +
            "46.9922 0 Td\n(Petit,) Tj\n29.7363 0 Td\n(Laura) Tj\n" +
            "31.8984 0 Td\n(Trépanier-) Tj\n53.0859 0 Td\n(Champagne) Tj\n" +
            "-404.6074 -24.9678 Td\n(On) Tj\nET";

        var revised = DxpFlowPdfVisitor.RewriteRightAlignedKernedLine(commands,
            "Gaspard Petit, Laura Trépanier-Champagne", metrics, new HashSet<string> { "F1" });

        Assert.NotNull(revised);
        Assert.Contains("102.2266 100 Td", revised);
        Assert.Contains("[(G)", revised);
        Assert.Contains("-404.666 -24.9678 Td", revised);
        Assert.Contains("(On) Tj", revised);
        Assert.Null(DxpFlowPdfVisitor.RewriteRightAlignedKernedLine(commands,
            "Gaspard Petit, Laura Trépanier-Champagne", metrics, new HashSet<string> { "F2" }));
        Assert.Null(DxpFlowPdfVisitor.RewriteRightAlignedKernedLine(
            commands + "\n" + commands,
            "Gaspard Petit, Laura Trépanier-Champagne", metrics, new HashSet<string> { "F1" }));
    }

    [Fact]
    public void FlowVisitorRebalancesOnlyIntraWordKerningOnLongJustifiedLine()
    {
        var metrics = DocxportNet.Doc.DocSystemFontAdvances.Find("Aptos", includeCloudFonts: true);
        Assert.NotNull(metrics);
        var commands = "BT\n/F0 12 Tf\n72 500 Td\n" +
            "(ubiquitous,) Tj\n( ) Tj\n60.743 0 Td\n(appearing) Tj\n( ) Tj\n" +
            "53.5301 0 Td\n(in) Tj\n( ) Tj\n11.1668 0 Td\n(cars,) Tj\n" +
            "-130 -17 Td\n(Next) Tj\nET";
        var revised = DxpFlowPdfVisitor.RewriteLongJustifiedKernedLines(commands,
            ["Before ubiquitous, appearing in cars, after"], metrics,
            new HashSet<string> { "F0" });

        Assert.NotNull(revised);
        Assert.Contains(" TJ", revised);
        Assert.Contains("( ) Tj", revised);
        Assert.Contains("(Next) Tj", revised);
        Assert.Null(DxpFlowPdfVisitor.RewriteLongJustifiedKernedLines(commands,
            [], metrics, new HashSet<string> { "F0" }));

        var numericLine = commands.Replace("(ubiquitous,)", "(protecting)")
            .Replace("(appearing)", "(TPMs)")
            .Replace("(in)", "(1996,)");
        Assert.NotNull(DxpFlowPdfVisitor.RewriteLongJustifiedKernedLines(numericLine,
            ["Before protecting TPMs 1996, cars, after"], metrics,
            new HashSet<string> { "F0" }));

        var escapedLine = commands.Replace("(appearing)", @"(\(CA\),)");
        var escapedRevision = DxpFlowPdfVisitor.RewriteLongJustifiedKernedLines(escapedLine,
            ["Before ubiquitous, (CA), in cars, after"], metrics,
            new HashSet<string> { "F0" });
        Assert.NotNull(escapedRevision);
        Assert.Contains(@"(\(CA\),) Tj", escapedRevision);

        var splitWordLine = "BT\n/F0 12 Tf\n72 500 Td\n(for) Tj\n( ) Tj\n" +
            "16.7412 0 Td\n(diagnosi) Tj\n43.7109 0 Td\n(s,) Tj\n( ) Tj\n" +
            "11.749 0 Td\n(maintenance,) Tj\n( ) Tj\n74.4385 0 Td\n" +
            "(and) Tj\n( ) Tj\n22.208 0 Td\n(repair) Tj\nET";
        Assert.NotNull(DxpFlowPdfVisitor.RewriteLongJustifiedKernedLines(splitWordLine,
            ["Before for diagnosis, maintenance, and repair after"], metrics,
            new HashSet<string> { "F0" }));
    }

    [Fact]
    public void FlowVisitorCarriesOrganizationKerningAcrossFootnoteMarker()
    {
        var metrics = DocxportNet.Doc.DocSystemFontAdvances.Find("Aptos", includeCloudFonts: true);
        Assert.NotNull(metrics);
        var commands = "BT\n/F0 12 Tf\n72 500 Td\n" +
            "(Federal) Tj\n( ) Tj\n44.6281 0 Td\n(Trade) Tj\n( ) Tj\n" +
            "34.632 0 Td\n(Commission) Tj\n( ) Tj\n71.4816 0 Td\n" +
            "(\\(FTC\\)) Tj\n( ) Tj\n32.8273 0 Td\n(filed) Tj\n" +
            "31.582 0 Td\n(report) Tj\n/F0 6.769 Tf\n31.582 3.6911 Td\n" +
            "(21) Tj\n/F0 12 Tf\n12.684 -3.6911 Td\n(on) Tj\nET";
        var revised = DxpFlowPdfVisitor.RewriteCapitalizedAcronymMarkerKerning(commands,
            ["The Federal Trade Commission (FTC) filed a report on repair"], metrics,
            new HashSet<string> { "F0" });

        Assert.NotNull(revised);
        Assert.Contains("[(Commis)", revised);
        Assert.DoesNotContain("44.6281 0 Td", revised);
        Assert.DoesNotContain("12.684 -3.6911 Td", revised);
        Assert.Contains("(on) Tj", revised);
        Assert.Null(DxpFlowPdfVisitor.RewriteCapitalizedAcronymMarkerKerning(commands,
            [], metrics, new HashSet<string> { "F0" }));
        Assert.Null(DxpFlowPdfVisitor.RewriteCapitalizedAcronymMarkerKerning(
            commands.Replace("/F0 6.769 Tf", "/F1 6.769 Tf"),
            ["The Federal Trade Commission (FTC) filed a report on repair"], metrics,
            new HashSet<string> { "F0" }));
        Assert.Null(DxpFlowPdfVisitor.RewriteCapitalizedAcronymMarkerKerning(
            commands.Replace("/F0 6.769 Tf", "ET\nBT\n/F0 6.769 Tf"),
            ["The Federal Trade Commission (FTC) filed a report on repair"], metrics,
            new HashSet<string> { "F0" }));
        Assert.NotNull(DxpFlowPdfVisitor.RewriteCapitalizedAcronymMarkerKerning(
            commands.Replace("Commission", "Council").Replace("FTC", "FRC"),
            ["The Federal Trade Council (FRC) filed a report on repair"], metrics,
            new HashSet<string> { "F0" }));
    }

    [Fact]
    public void FlowVisitorKeepsNaturalSpacingOnFinalJustifiedBodyLines()
    {
        var commands = "(maintenance,) Tj\n73.7754 0 Td\n(repair) Tj\n32.1504 0 Td\n" +
            "(and) Tj\n21.5449 0 Td\n(diagnosis.) Tj\nET";

        var revised = DxpFlowPdfVisitor.JustifyFinalTextLine(commands,
            ["maintenance,", "repair", "and", "diagnosis."], 1.65);

        Assert.Contains("74.3254 0 Td\n(repair) Tj", revised);
        Assert.Contains("32.7004 0 Td\n(and) Tj", revised);
        Assert.Contains("22.0949 0 Td\n(diagnosis.) Tj", revised);
    }

    [Fact]
    public void FinalLineAdjustmentUsesBaselineWhenWordsRepeatOnPreviousLine()
    {
        var commands = "BT\n72 700 Td\n(same) Tj\n20 0 Td\n(words) Tj\n" +
            "-20 -17 Td\n(same) Tj\n20 0 Td\n(words) Tj\nET";

        var revised = DxpFlowPdfVisitor.JustifyFinalTextLine(commands,
            ["same", "words"], 4, 72, 683);

        Assert.Contains("72 700 Td\n(same) Tj\n20 0 Td\n(words) Tj", revised);
        Assert.Contains("-20 -17 Td\n(same) Tj\n24 0 Td\n(words) Tj", revised);
    }

    [Fact]
    public void ContinuedTableLineSpreadsOnlyTheMatchingColumn()
    {
        var line = "(component) Tj\n61 0 Td\n(that,) Tj\n26 0 Td\n(in) Tj\n11 0 Td\n(the) Tj\n" +
            "18 0 Td\n(ordinary) Tj\n44 0 Td\n(course) Tj\n37 0 Td\n(of) Tj\n";
        var commands = "BT\n77.65 708 Td\n" + line +
            "ET\nBT\n311.4 708 Td\n" + line + "ET";
        var revised = DxpFlowPdfVisitor.JustifyFinalTextLine(commands,
            ["component", "that,", "in", "the", "ordinary", "course", "of"], 9, 311.4);

        Assert.Contains("77.65 708 Td\n" + line, revised);
        Assert.Contains("311.4 708 Td\n(component) Tj\n62.5 0 Td\n(that,) Tj", revised);
    }

    [Fact]
    public void SplitTableLastLineSpreadsOnlyTheSelectedColumn()
    {
        var line = "(means) Tj\n37 0 Td\n(any) Tj\n20 0 Td\n(effective) Tj\n" +
            "46 0 Td\n(technology,) Tj\n62 0 Td\n(device) Tj\n35 0 Td\n(or) Tj\n";
        var commands = "BT\n77.65 82.35 Td\n" + line +
            "ET\nBT\n311.4 82.35 Td\n" + line + "ET";

        var revised = DxpFlowPdfVisitor.JustifyFinalTextLine(commands,
            ["means", "any", "effective", "technology,", "device", "or"],
            5, 311.4, 82.35);

        Assert.Contains("77.65 82.35 Td\n" + line, revised);
        Assert.Contains("311.4 82.35 Td\n(means) Tj\n38 0 Td\n(any) Tj", revised);
    }

    [Fact]
    public void FinalLineMatchesEncodedLeadingPunctuationWithoutMovingFollowingTable()
    {
        var commands = "BT\n77.65 460 Td\n(\u0093technological) Tj\n( ) Tj\n" +
            "77 0 Td\n(protection) Tj\n( ) Tj\n55 0 Td\n(necessary.) Tj\n" +
            "-130 -25 Td\n(Before) Tj\nET";

        var revised = DxpFlowPdfVisitor.JustifyFinalTextLine(commands,
            ["“technological", "protection", "necessary."], 2, 77.65, 460,
            allowEncodedPunctuation: true);

        Assert.Contains("78 0 Td\n(protection) Tj", revised);
        Assert.Contains("56 0 Td\n(necessary.) Tj", revised);
        Assert.Contains("-132 -25 Td\n(Before) Tj", revised);
    }

    [Fact]
    public void TrackedUnderlineMovesBothEndsAcrossSplitRuns()
    {
        var lines = new[] {
            "BT", "90 700 Td", "(Derived) Tj", "ET", "90 698.1 m", "124.46 698.1 l", "S",
            "BT", "128 700 Td", "(body) Tj", "(paragraph) Tj", "ET",
            "124.46 698.1 m", "226.42 698.1 l", "S",
            "BT", "90 650 Td", "(unrelated) Tj", "ET", "90 648 m", "130 648 l", "S"
        };

        DxpFlowPdfVisitor.AdjustTrackedUnderlines(lines,
            new Dictionary<int, (double Baseline, double EndShift)> {
                [3] = (700, 7), [11] = (700, 22)
            });

        Assert.Equal("90 698.1 m", lines[4]);
        Assert.Equal("131.46 698.1 l", lines[5]);
        Assert.Equal("131.46 698.1 m", lines[12]);
        Assert.Equal("248.42 698.1 l", lines[13]);
        Assert.Equal("130 648 l", lines[20]);
    }

    [Fact]
    public void FlowVisitorKeepsInheritedTrackingUnderlinesAlignedWithText()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordThreeLevelStyleOverrides.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var page = pdf.Pages[0];
        var content = string.Concat(page.Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Select(reference => reference.Value)
            .OfType<PdfSharp.Pdf.Advanced.PdfContent>()
            .Where(stream => stream.Stream != null)
            .Select(stream => System.Text.Encoding.Latin1.GetString(stream.Stream!.UnfilteredValue)));
        var rules = System.Text.RegularExpressions.Regex.Matches(content,
            @"(?m)^(?<left>\d+(?:\.\d+)?) (?<y>\d+(?:\.\d+)?) m\r?\n(?<right>\d+(?:\.\d+)?) \k<y> l\r?\nS$")
            .Select(match => (
                Left: double.Parse(match.Groups["left"].Value, System.Globalization.CultureInfo.InvariantCulture),
                Right: double.Parse(match.Groups["right"].Value, System.Globalization.CultureInfo.InvariantCulture),
                Top: page.Height.Point - double.Parse(match.Groups["y"].Value,
                    System.Globalization.CultureInfo.InvariantCulture))).ToArray();

        Assert.Contains(rules, rule => rule.Top is > 54 and < 59 &&
            rule.Left is > 89.9 and < 90.1 && rule.Right is > 202.6 and < 203.1);
        Assert.Contains(rules, rule => rule.Top is > 90 and < 95 &&
            rule.Left is > 131.2 and < 131.7 && rule.Right is > 248.2 and < 248.7);
    }

    [Fact]
    public void ContinuedJustifiedPageLineIncludesInlineFootnoteMarker()
    {
        var commands = "(last) Tj\n30 0 Td\n(year) Tj\n4 0 Td\n(.) Tj\n" +
            "/F1 7 Tf\n6 3.6911 Td\n(29) Tj\n/F0 12 Tf\n8 -3.6911 Td\n(next) Tj\nET";
        var revised = DxpFlowPdfVisitor.JustifyFinalTextLine(commands,
            ["last", "year.", "29", "next"], 9);

        Assert.Contains("33 0 Td\n(year) Tj", revised);
        Assert.Contains("9 3.6911 Td\n(29) Tj", revised);
        Assert.Contains("11 -3.6911 Td\n(next) Tj", revised);
        Assert.Contains("4 0 Td\n(.) Tj", revised);

        var baselineWords = DxpFlowPdfVisitor.JustifyFinalTextLine(commands,
            ["last", "year.", "next"], 6);
        Assert.Contains("33 0 Td\n(year) Tj", baselineWords);
        Assert.Contains("6 3.6911 Td\n(29) Tj", baselineWords);
        Assert.Contains("11 -3.6911 Td\n(next) Tj", baselineWords);
    }

    [Fact]
    public void ParagraphFinalLineSpreadsBeforeUnextractedRaisedFootnoteMarker()
    {
        var commands = "BT\n72 435.64 Td\n(the) Tj\n18 0 Td\n(numerous) Tj\n" +
            "30 0 Td\n(reviews.) Tj\n/F1 7 Tf\n42 3.6911 Td\n(7) Tj\nET";

        var revised = DxpFlowPdfVisitor.JustifyFinalTextLine(commands,
            ["the", "numerous", "reviews."], 1.1, 72, 435.64);

        Assert.Contains("18.55 0 Td\n(numerous) Tj", revised);
        Assert.Contains("30.55 0 Td\n(reviews.) Tj", revised);
        Assert.Contains("42 3.6911 Td\n(7) Tj", revised);
    }

    [Fact]
    public void ContinuedJustifiedFragmentOnlySpreadsItsOwnLineInSharedTextBlock()
    {
        var commands = "(first) Tj\n20 0 Td\n(line) Tj\n/F1 7 Tf\n5 3.6911 Td\n(10) Tj\n" +
            "/F0 12 Tf\n8 -3.6911 Td\n(after) Tj\n-100 -14 Td\n" +
            "(second) Tj\n20 0 Td\n(line) Tj\n20 0 Td\n(end) Tj\nET";

        var first = DxpFlowPdfVisitor.JustifyFinalTextLine(commands,
            ["first", "line", "after"], 12);
        var second = DxpFlowPdfVisitor.JustifyFinalTextLine(first,
            ["second", "line", "end"], 10);

        Assert.Contains("26 0 Td\n(line) Tj", first);
        Assert.Contains("14 -3.6911 Td\n(after) Tj", first);
        Assert.Contains("-112 -14 Td\n(second) Tj\n20 0 Td\n(line) Tj", first);
        Assert.Contains("(second) Tj\n25 0 Td\n(line) Tj\n25 0 Td\n(end) Tj", second);
    }

    [Fact]
    public void ContinuedLayoutLineCanNarrowWithoutShiftingTheNextLine()
    {
        var commands = "BT\n72 708 Td\n(under) Tj\n40 0 Td\n(sections) Tj\n" +
            "50 0 Td\n(following) Tj\n-90 -17 Td\n(next) Tj\nET";

        var revised = DxpFlowPdfVisitor.JustifyFinalTextLine(commands,
            ["under", "sections", "following"], -1, 72, 708);

        Assert.Contains("39.5 0 Td\n(sections) Tj\n49.5 0 Td\n(following) Tj", revised);
        Assert.Contains("-89 -17 Td\n(next) Tj", revised);
    }

    [Fact]
    public void ShortContinuationWidensOnlyCompressedSpaceSpans()
    {
        var paragraph = new MigraDoc.DocumentObjectModel.Paragraph();
        var space = paragraph.AddFormattedText(" ");
        space.Size = MigraDoc.DocumentObjectModel.Unit.FromPoint(8.5);
        var word = paragraph.AddFormattedText("repair");
        word.Size = MigraDoc.DocumentObjectModel.Unit.FromPoint(8.5);

        DxpFlowPdfVisitor.WidenShortContinuationSpaces(paragraph);

        Assert.Equal(9.5, space.Size.Point);
        Assert.Equal(8.5, word.Size.Point);
    }

    [Theory]
    [InlineData(42, 171, true)]
    [InlineData(143, 70, true)]
    [InlineData(94, 22, false)]
    [InlineData(85, 9, false)]
    [InlineData(73, 12, false)]
    public void CompressedContinuationSpaceWideningTracksRemainingText(
        int fittingSourceWords, int remainingSourceWords, bool expected)
    {
        Assert.Equal(expected, DxpFlowPdfVisitor.ShouldWidenCompressedContinuation(
            fittingSourceWords, remainingSourceWords));
    }



    [Fact]
    public void FlowVisitorFallsBackToDecimalWhenCircledMarkerFontIsUnavailable()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordSymbolicCircleOnly.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.DoesNotContain(page.Letters, letter => letter.Value == "\0");
        Assert.Contains(page.GetWords(), word => word.Text == "1." &&
            word.BoundingBox.Bottom > page.Height / 2);
    }

    [Fact]
    public void FlowVisitorRendersChicagoSymbolsInBodyList()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "DocKnownGaps",
            "WordSymbolicChicagoOnly.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var text = pdf.GetPage(1).Text;
        Assert.Contains("*.", text);
        Assert.Contains("†.", text);
        Assert.Contains("‡.", text);
    }

    [Fact]
    public void FlowVisitorUsesWordMergePlaceholdersInCachedMarginStories()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordMergeFieldsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(),
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Cache });

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Contains("«CustomerName»", pdf.GetPage(1).Text);
        Assert.Contains("«CustomerName»", pdf.GetPage(2).Text);
        Assert.Contains("Alice", pdf.GetPage(2).Text);
    }

    [Fact]
    public void FlowVisitorUsesWordMergePlaceholdersInEvaluatedMarginStories()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordMergeFieldsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(),
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate });

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Contains("«CustomerName»", pdf.GetPage(1).Text);
        Assert.DoesNotContain("Alice", pdf.GetPage(2).Text);
    }

    [Fact]
    public void FlowVisitorKeepsResolvedMergeValuesInMarginStories()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordMergeFieldsAllStories.docx");
        var visitor = new DxpFlowPdfVisitor();
        visitor.FieldEval.Context.ValueResolver = new CustomerNameResolver();
        var bytes = DxpExport.ExportToBytes(source, visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate });

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Contains("Taylor", pdf.GetPage(1).Text);
        Assert.DoesNotContain("«CustomerName»", pdf.GetPage(1).Text);
    }

    [Fact]
    public void FlowVisitorNumbersSourceBodyLinesWithoutCountingLayoutSpacers()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "TestNumberedParagraphs.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var marginLines = pdf.GetPage(1).Letters.Where(letter =>
            letter.GlyphRectangle.Left < 65 && letter.GlyphRectangle.Bottom > 280).ToArray();
        Assert.Equal("51015", string.Concat(marginLines.Select(letter => letter.Value)));
        Assert.InRange(marginLines[0].GlyphRectangle.Left, 50, 53);
        Assert.InRange(marginLines[1].GlyphRectangle.Left, 44, 47);
        Assert.InRange(marginLines[3].GlyphRectangle.Left, 44, 47);
        Assert.InRange(pdf.GetPage(1).Height - marginLines[0].GlyphRectangle.Top, 165, 165.6);
        var specificationLines = pdf.GetPage(1).GetWords()
            .Where(word => word.Text == "Line").ToArray();
        Assert.Equal(2, specificationLines.Length);
        Assert.All(specificationLines, word => Assert.InRange(word.BoundingBox.Left, 131, 136));
        var patentLabels = pdf.GetPage(1).GetWords()
            .Where(word => word.Text.StartsWith("[000", StringComparison.Ordinal)).ToArray();
        Assert.Equal(3, patentLabels.Length);
        Assert.All(patentLabels, word =>
            Assert.Contains("Bold", word.Letters[0].FontName, StringComparison.OrdinalIgnoreCase));
        Assert.InRange(specificationLines[1].BoundingBox.Top - patentLabels[2].BoundingBox.Top,
            46.6, 47.1);
        var itemWords = pdf.GetPage(1).GetWords()
            .Where(word => word.Text is "Item").ToArray();
        Assert.Equal(4, itemWords.Length);
        Assert.InRange(itemWords[3].BoundingBox.Top - itemWords[2].BoundingBox.Top,
            -65.6, -65.2);
    }

    [Fact]
    public void DirectAutomaticParagraphShadingClearsInheritedFill()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyledParagraphShadingResetAllStories.docx");
        using var document = WordprocessingDocument.Open(source, false);
        var resolver = new DxpStyleResolver(document);
        var body = document.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();
        var header = document.MainDocumentPart.HeaderParts.Single().Header.Descendants<Paragraph>().Single();
        var footer = document.MainDocumentPart.FooterParts.Single().Footer.Descendants<Paragraph>().Single();

        Assert.Equal("B8E0F4", resolver.GetParagraphShading(body)?.Fill?.Value);
        Assert.Equal("auto", resolver.GetParagraphShading(header)?.Fill?.Value);
        Assert.Equal("C5E5C2", resolver.GetParagraphShading(footer)?.Fill?.Value);
    }

    [Fact]
    public void FlowVisitorAutoFitsImbalancedTableColumns()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "DocKnownGaps",
            "WordVisibleNoWrapArialAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var secondColumnStart = pdf.GetPage(1).Letters.Where(letter => letter.Value == "H" &&
            letter.GlyphRectangle.Left > 400).Min(letter => letter.GlyphRectangle.Left);
        Assert.InRange(secondColumnStart, 490, 510);
        var bodyColumnStart = pdf.GetPage(1).Letters.Where(letter => letter.Value == "B" &&
            letter.GlyphRectangle.Left > 400 &&
            pdf.GetPage(1).Height - letter.GlyphRectangle.Top is > 90 and < 200)
            .Min(letter => letter.GlyphRectangle.Left);
        Assert.InRange(bodyColumnStart - secondColumnStart, 7, 12);
        var words = pdf.GetPage(1).GetWords();
        var headerWord = words.First(word => word.Text == "Head" && word.BoundingBox.Left > 400);
        var footerWord = words.First(word => word.Text == "Foote" && word.BoundingBox.Left > 400);
        Assert.InRange(footerWord.BoundingBox.Left - headerWord.BoundingBox.Left, 2.5, 5.5);
        var bodyWrap = Assert.Single(words.Where(word => word.Text == "wrap" &&
            word.BoundingBox.Left is > 75 and < 120 &&
            pdf.GetPage(1).Height - word.BoundingBox.Top is > 90 and < 200));
        var bodySetting = words.First(word => word.Text == "setting" &&
            word.BoundingBox.Left is > 100 and < 120 &&
            pdf.GetPage(1).Height - word.BoundingBox.Top is > 90 and < 200);
        Assert.InRange(Math.Abs(bodyWrap.Letters[0].StartBaseLine.Y -
            bodySetting.Letters[0].StartBaseLine.Y), 0, 0.3);
    }

    [Theory]
    [InlineData("WordTwoColumnsStories", 315, "Column paragraph 53")]
    [InlineData("WordUnequalColumnsStories", 260, "Column paragraph 53")]
    [InlineData("WordExplicitColumnBreakStories", 315, "Column paragraph 36")]
    public void FlowVisitorFlowsSimpleParagraphsAcrossColumns(string name, double secondColumnX,
        string nextPageParagraph)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains(new string(nextPageParagraph.Where(character => !char.IsWhiteSpace(character)).ToArray()),
            new string(pdf.GetPage(2).Text.Where(character => !char.IsWhiteSpace(character)).ToArray()));
        var columnStarts = pdf.GetPage(1).Letters.Where(letter => letter.Value == "C" &&
            letter.GlyphRectangle.Left > 230).Select(letter => letter.GlyphRectangle.Left);
        Assert.Contains(columnStarts, left => Math.Abs(left - secondColumnX) < 2);
    }

    [Fact]
    public void FlowVisitorDrawsInheritedRunShadingAndDirectColorOverride()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyledRunShadingResetAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var colors = pdf.GetPage(1).Paths.Where(path => path.IsFilled && path.FillColor != null)
            .Select(path => path.FillColor!.ToRGBValues()).Distinct().ToArray();
        Assert.True(colors.Length >= 2);
    }

    [Fact]
    public void FlowVisitorPositionsMarginAnchoredImagesInAllStories()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordTransformedFloatingImageStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var images = pdf.GetPage(1).GetImages()
            .Where(image => image.WidthInSamples > 1 && image.HeightInSamples > 1)
            .OrderByDescending(image => image.Bounds.Bottom).ToArray();
        Assert.Equal(3, images.Length);
        Assert.InRange(images[0].Bounds.Left, 175, 183);
        Assert.InRange(images[1].Bounds.Left, 153, 163);
        Assert.InRange(images[1].Bounds.Width, 145, 151);
        Assert.InRange(images[2].Bounds.Left, 100, 115);
        Assert.InRange(pdf.GetPage(1).Height - images[2].Bounds.Top, 752, 762);
    }

    [Fact]
    public void FlowVisitorPositionsColumnAnchoredImageRelativeToColumn()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordColumnPositionedFloatingPngBody.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var image = Assert.Single(pdf.GetPage(1).GetImages());
        Assert.InRange(image.Bounds.Left, 141.5, 142.5);
        Assert.InRange(pdf.GetPage(1).Height - image.Bounds.Top, 161.5, 162.5);
    }

    [Fact]
    public void FlowVisitorPositionsPageAnchoredImageWithoutReflowingBody()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordPagePositionedFloatingPngBody.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var image = Assert.Single(page.GetImages());
        Assert.InRange(image.Bounds.Left, 119.5, 120.5);
        Assert.InRange(page.Height - image.Bounds.Top, 159.5, 160.5);
        Assert.Contains(page.GetWords(), word => word.Text == "Following" &&
            page.Height - word.BoundingBox.Top is > 96 and < 99);
    }

    [Fact]
    public void FlowVisitorUsesInsetRectangularWrapPolygon()
    {
        if (typeof(MigraDoc.Rendering.DocumentRenderer).GetProperty("WrapContour") == null) return;
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "PdfProbeCorpus",
            "square-wrap-probe__overlap-sources__WordThroughLongFloatingPngBody.docx");
        var insetSource = File.ReadAllBytes(source);
        using var fullBoundsSource = new MemoryStream();
        fullBoundsSource.Write(insetSource);
        using (var document = WordprocessingDocument.Open(fullBoundsSource, true))
        {
            var polygon = Assert.Single(document.MainDocumentPart!.Document!.Descendants<
                DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapPolygon>());
            foreach (var vertex in polygon.ChildElements.Where(vertex =>
                         vertex.GetAttribute("x", "").Value == "21330"))
                vertex.SetAttribute(new OpenXmlAttribute("", "x", "", "21600"));
            document.MainDocumentPart.Document.Save();
        }

        double WrappedWordX(byte[] sourceBytes)
        {
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(
                DxpExport.ExportToBytes(sourceBytes, new DxpFlowPdfVisitor()));
            var page = pdf.GetPage(1);
            return Assert.Single(page.GetWords().Where(word => word.Text == "Words" &&
                page.Height - word.BoundingBox.Top is > 160 and < 170)).BoundingBox.Left;
        }

        var insetX = WrappedWordX(insetSource);
        var fullBoundsX = WrappedWordX(fullBoundsSource.ToArray());
        Assert.InRange(fullBoundsX - insetX, 1.0, 2.0);
    }

    [Fact]
    public void FlowVisitorWrapsTextAroundParagraphAnchoredImage()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordParagraphPositionedFloatingPngBody.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var image = Assert.Single(page.GetImages());
        Assert.InRange(image.Bounds.Left, 151.5, 152.5);
        Assert.InRange(page.Height - image.Bounds.Top, 101.5, 102.5);
        Assert.Contains(page.GetWords(), word => word.Text == "Following" &&
            word.BoundingBox.Left is > 71 and < 73);
        Assert.Contains(page.GetWords(), word => word.Text == "paragraph." &&
            word.BoundingBox.Left is > 280 and < 282 &&
            page.Height - word.BoundingBox.Top is > 96 and < 99);
    }

    [Fact]
    public void FlowVisitorWrapsShortHeaderAndBodyAroundMarginAnchoredImage()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordFloatingPngHeaderFooterOnly.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.GetWords(), word => word.Text == "picture" &&
            word.BoundingBox.Left is > 229 and < 233 &&
            page.Height - word.BoundingBox.Top is > 35 and < 52);
        Assert.Contains(page.GetWords(), word => word.Text == "paragraph." &&
            word.BoundingBox.Left is > 229 and < 233 &&
            page.Height - word.BoundingBox.Top is > 70 and < 88);
    }

    [Fact]
    public void FlowVisitorWrapsHeaderAroundTransformedImageBounds()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordTransformedFloatingImageStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.GetWords(), word => word.Text == "picture" &&
            word.BoundingBox.Left is > 222 and < 225 &&
            page.Height - word.BoundingBox.Top is > 35 and < 52);
    }

    [Theory]
    [InlineData("WordBalancedSpaceWithStories.docx", 78.0)]
    [InlineData("WordReferenceWithStories.docx", 75.3)]
    public void FlowVisitorPreservesLeadingSpaceWithWordBalancing(string fixtureName, double expectedLeft)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixtureName);

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(2);
        var reference = page.GetWords().First(word => word.Text == "Ref:" &&
            page.Height - word.BoundingBox.Top is > 70 and < 75);
        Assert.InRange(reference.BoundingBox.Left, expectedLeft - 0.5, expectedLeft + 0.5);
    }

    [Fact]
    public void FlowVisitorWidensBalancedFooterFieldSpacesWithoutMovingFollowingText()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordBalancedSpaceWithStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var footer = page.GetWords().Where(word => word.BoundingBox.Bottom < 65)
            .OrderBy(word => word.BoundingBox.Left).ToArray();
        Assert.Equal(["Page", "1", "of", "2"], footer.Select(word => word.Text));
        Assert.InRange(footer[1].BoundingBox.Left, 102.8, 103.8);
        Assert.Contains(page.GetWords(), word => word.Text == "Open" &&
            word.BoundingBox.Left is > 71.5 and < 72.5);
    }

    [Fact]
    public void FlowVisitorPreservesTheNextTextOriginAfterWideningBalancedSpaces()
    {
        var source = string.Join('\n', "BT", "/F0 12 Tf", "72 39 Td", "(Page) Tj",
            "27.873 0 Td", "(1) Tj", "8.8477 0 Td", "(of) Tj", "12.6738 0 Td", "(2) Tj",
            "-49.3945 669.3516 Td", "(Open) Tj", "ET");
        var revised = DxpFlowPdfVisitor.RewriteBalancedFooterFieldSpaces(source,
            new Dictionary<string, int> { ["Page"] = 3 }, ["F0"], 792);

        Assert.Contains("31.233 0 Td\n(1) Tj", revised);
        Assert.Contains("12.2077 0 Td\n(of) Tj", revised);
        Assert.Contains("16.0338 0 Td\n(2) Tj", revised);
        Assert.Contains("-59.4745 669.3516 Td\n(Open) Tj", revised);
    }

    [Fact]
    public void FlowVisitorLayersMarginAnchoredImagesByRelativeHeight()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordTwoFloatingPicturesBody.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var images = pdf.GetPage(1).GetImages()
            .Where(image => image.WidthInSamples > 1 && image.HeightInSamples > 1).ToArray();
        Assert.Equal(2, images.Length);
        Assert.InRange(images[0].Bounds.Left, 171.5, 172.5);
        Assert.InRange(images[1].Bounds.Left, 241, 242);
    }

    [Fact]
    public void FlowVisitorRendersFullWidthDecimalListMarkersInAllStories()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "DocKnownGaps",
            "WordDecimalFullWidth2ListsAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.True(page.Letters.Count(letter => letter.GlyphRectangle.Left is > 91 and < 92 &&
            letter.GlyphRectangle.Width is > 5 and < 7) >= 7);
        Assert.DoesNotContain(page.Letters, letter => letter.Value is "1" or "2" or "3");
    }

    [Theory]
    [InlineData("Doc", "WordSymbolicCircleOnly")]
    [InlineData("DocKnownGaps", "WordDecimalFullWidth2ListsAllStories")]
    public void FlowVisitorNarrowsConsecutiveNumberedHeaderSpacing(string folder, string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", folder, name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfSharp.Pdf.IO.PdfReader.Open(new MemoryStream(bytes),
            PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import);
        var commands = string.Join('\n', pdf.Pages[0].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Select(reference => reference.Value)
            .OfType<PdfSharp.Pdf.Advanced.PdfContent>()
            .Select(content => System.Text.Encoding.Latin1.GetString(content.Stream!.UnfilteredValue)));
        Assert.Matches(@"-\d+(?:\.\d+)? -12\.[0-4]\d* Td", commands);
    }

    [Theory]
    [InlineData("Doc", "WordVisibleWidthDecimalListsAllStories")]
    [InlineData("DocKnownGaps", "WordDecimalFullWidth2ListsAllStories")]
    [InlineData("Doc", "WordIdeographDigitalListsAllStories")]
    public void FlowVisitorUsesWordPitchForNumberedFooters(string folder, string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", folder, name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var footerWords = page.GetWords().Where(word =>
            page.Height - word.BoundingBox.Top > 700).ToArray();
        var first = footerWords.First(word => word.Text == "First");
        var second = footerWords.First(word => word.Text == "Second");
        Assert.InRange(first.BoundingBox.Top - second.BoundingBox.Top, 17.4, 17.8);
    }

    [Fact]
    public void FlowVisitorPositionsFullWidth2NumberedHeaderAtWordBaseline()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "DocKnownGaps",
            "WordDecimalFullWidth2ListsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var first = page.Letters.First(letter => letter.Value == "F" &&
            page.Height - letter.StartBaseLine.Y < 75);
        Assert.InRange(page.Height - first.StartBaseLine.Y, 45.6, 46.0);
    }

    [Fact]
    public void FlowVisitorPreservesNonconsecutiveRightToRepairFootnoteReferences()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "samples", "Right to Repair.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(6);
        var references = page.GetWords().Where(word =>
                page.Height - word.BoundingBox.Top is > 110 and < 190 &&
                word.Text is "12" or "13" or "14" or "15" or "16")
            .Select(word => word.Text).ToArray();
        Assert.Contains("12", references);
        Assert.Contains("13", references);
        Assert.Contains("16", references);
        Assert.DoesNotContain("14", references);
        Assert.DoesNotContain("15", references);
    }

    [Fact]
    public void CjkListRiseAlignsBodyMarkersAndPreservesOtherStories()
    {
        const string stream = "BT\n/F0 12 Tf\n90 704.24 Td\n<04B2> Tj\n/F1 12 Tf\n" +
            "12 0 Td\n(First body item) Tj\n/F0 12 Tf\n-80 -20 Td\n<0CCD> Tj\n" +
            "/F1 12 Tf\n12 0 Td\n(Second body item) Tj\n/F0 12 Tf\n" +
            "-80 -680 Td\n<04B2> Tj\n/F1 12 Tf\n12 0 Td\n(Footer item) Tj\nET";

        var revised = DxpFlowPdfVisitor.RaiseCjkListTextRuns(stream,
            new HashSet<string> { "F0" }, new HashSet<string> { "F1" },
            [(70, 140)], 792);

        Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(revised, "1.38 Ts").Count);
        Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(revised, "1.46 Ts").Count);
        Assert.Contains("1.46 Ts\n<04B2> Tj\n0 Ts\n/F1 12 Tf\n1.38 Ts", revised);
        Assert.Contains("0 Ts\n/F0 12 Tf\n-80 -20 Td\n1.46 Ts\n<0CCD> Tj", revised);
        Assert.Contains("<04B2> Tj\n/F1 12 Tf\n12 0 Td\n(Footer item) Tj", revised);
    }

    [Fact]
    public void BulletBaselineRiseResetsBeforeUnmarkedContinuation()
    {
        const string stream = "BT\n/F0 12 Tf\n90 708.68 Td\n<0078> Tj\n" +
            "/F1 12 Tf\n18 0 Td\n(First) Tj\n-40 -17.23 Td\n(next) Tj\n" +
            "/F2 12 Tf\n-20 -17 Td\n(o) Tj\n/F1 12 Tf\n18 0 Td\n(Second) Tj\nET";

        var revised = DxpFlowPdfVisitor.RewriteBulletBaselineStream(stream,
            new Dictionary<string, string> { ["F0"] = "Symbol", ["F2"] = "Courier New" },
            [(70.0, 90.0, "Symbol", 0.82), (105.0, 125.0, "Courier New", 0.68)], 792);

        Assert.Contains("90 708.68 Td\n-0.82 Ts\n<0078> Tj", revised);
        Assert.Contains("(First) Tj\n0 Ts\n-40 -17.23 Td\n(next) Tj", revised);
        Assert.Contains("-20 -17 Td\n-0.68 Ts\n(o) Tj", revised);
        Assert.Contains("(Second) Tj\n0 Ts\nET", revised);
    }

    [Fact]
    public void BulletBaselineRiseCanAlignTaggedContinuationIndependently()
    {
        const string stream = "BT\n/F0 12 Tf\n90 708.68 Td\n<0078> Tj\n" +
            "/F1 12 Tf\n18 0 Td\n(First) Tj\n-40 -17.23 Td\n(next) Tj\nET";

        var revised = DxpFlowPdfVisitor.RewriteBulletBaselineStream(stream,
            new Dictionary<string, string> { ["F0"] = "Symbol", ["F1"] = "Aptos" },
            [(70.0, 90.0, "Symbol", 0.82), (95.0, 110.0, "Aptos", 0.51)], 792);

        Assert.Contains("(First) Tj\n0 Ts\n-40 -17.23 Td\n-0.51 Ts\n(next) Tj", revised);
        Assert.EndsWith("(next) Tj\n0 Ts\nET", revised);
    }

    [Fact]
    public void MixedDisplayMathRunShiftLeavesFractionAnchorUnchanged()
    {
        const string stream = "BT\n/F0 12 Tf\n292 705.83 Td\n<0754> Tj\n" +
            "/F1 12 Tf\n6 0 Td\n( ) Tj\n/F2 12 Tf\n3 0 Td\n(+) Tj\n" +
            "/F3 14.01 Tf\n2 0 Td\n( ) Tj\nET\n" +
            "BT\n/F0 12 Tf\n100 680 Td\n<0754> Tj\nET";

        var revised = DxpFlowPdfVisitor.RewriteMixedDisplayMathRunStream(stream,
            new HashSet<string> { "F0", "F2" }, [(80.0, 95.0)], 792);

        Assert.Contains("292 705.83 Td\n-4.81 Ts\n<0754> Tj\n0 Ts", revised);
        Assert.Contains("( ) Tj\n/F2 12 Tf\n3 0 Td\n-4.81 Ts\n(+) Tj\n0 Ts", revised);
        Assert.Contains("/F3 14.01 Tf\n2 0 Td\n( ) Tj", revised);
        Assert.EndsWith("100 680 Td\n<0754> Tj\nET", revised);
    }

    [Theory]
    [InlineData("WordVisibleListLabelCapsUnderlineAllStories", "A.", 2, 0.8, 0.2)]
    [InlineData("WordVisibleListLabelStyleAllStories", "1.", 1, 0.2, 0.4)]
    public void FlowVisitorUsesNumberingLevelRunFormattingForListLabels(string name,
        string expectedMarker, int underlineCount, double red, double green)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var markers = pdf.GetPage(1).GetWords().Where(word => word.Text == expectedMarker).ToArray();
        Assert.True(markers.Length >= 3);
        Assert.All(markers, marker =>
        {
            Assert.Contains("BoldItalic", marker.Letters[0].FontName,
                StringComparison.OrdinalIgnoreCase);
            var color = marker.Letters[0].Color.ToRGBValues();
            Assert.InRange(color.r, 0.19, 0.21);
            Assert.InRange(color.g, 0.39, 0.41);
            Assert.InRange(color.b, 0.79, 0.81);
        });
        var page = pdf.GetPage(1);
        var underlines = page.Paths.Where(path => path.IsFilled &&
                path.GetBoundingRectangle() is { } bounds &&
                bounds.Left is > 89 and < 91 && bounds.Width is > 9 and < 13 &&
                page.Height - bounds.Top is > 47 and < 51)
            .ToArray();
        Assert.Equal(underlineCount, underlines.Length);
        Assert.All(underlines, path =>
        {
            var color = path.FillColor!.ToRGBValues();
            Assert.InRange(color.r, red - 0.01, red + 0.01);
            Assert.InRange(color.g, green - 0.01, green + 0.01);
        });
    }

    [Fact]
    public void FlowVisitorStrikesThroughNumberingLevelLabels()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleListLabelEffectsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var strikes = page.Paths.Where(path => path.IsFilled &&
                path.GetBoundingRectangle() is { } bounds &&
                bounds.Left is > 89 and < 91 && bounds.Width is > 9 and < 11 &&
                page.Height - bounds.Top is > 42 and < 44)
            .ToArray();
        Assert.Single(strikes);
        Assert.InRange(strikes[0].GetBoundingRectangle()!.Value.Height, 0.59, 0.61);
    }

    [Fact]
    public void FlowVisitorKeepsFullWidthListBodyAtWordPitch()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleWidthDecimalListsAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var bodyLines = page.GetWords().Where(word => word.Text is "First" or "Second" or "Third" &&
                page.Height - word.BoundingBox.Top is > 70 and < 150)
            .Select(word => page.Height - word.BoundingBox.Top).OrderBy(top => top).ToArray();
        Assert.Equal(3, bodyLines.Length);
        Assert.InRange(bodyLines[1] - bodyLines[0], 19, 21);
        Assert.InRange(bodyLines[2] - bodyLines[1], 19, 21);
    }

    [Theory]
    [InlineData("WordTextControls")]
    [InlineData("WordTextControlsNbspAllStories")]
    [InlineData("WordVisibleMirrorMarginsAllStories")]
    [InlineData("WordVisibleMirrorGutterAllStories")]
    public void FlowVisitorAlignsBodyAfterInlineTextControls(string name)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", name + ".docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        var body = pdf.GetPage(1).Letters.First(letter => letter.Value == "B" &&
            pdf.GetPage(1).Height - letter.StartBaseLine.Y is > 75 and < 95);
        var continuation = pdf.GetPage(2).Letters.First(letter => letter.Value == "p" &&
            pdf.GetPage(2).Height - letter.StartBaseLine.Y is > 75 and < 95);
        Assert.InRange(pdf.GetPage(1).Height - body.StartBaseLine.Y, 83.2, 84.8);
        Assert.InRange(pdf.GetPage(2).Height - continuation.StartBaseLine.Y, 83.2, 84.8);
    }

    [Fact]
    public void FlowVisitorPlacesTopGutterAboveBodyOnEveryPage()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleTopGutterAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains(pdf.GetPage(1).Letters, letter => letter.Value == "B" &&
            pdf.GetPage(1).Height - letter.GlyphRectangle.Top is > 107 and < 112);
        Assert.Contains(pdf.GetPage(2).Letters, letter => letter.Value == "p" &&
            pdf.GetPage(2).Height - letter.GlyphRectangle.Top is > 107 and < 114);
        Assert.Contains(pdf.GetPage(2).Letters, letter => letter.Value == "H" &&
            letter.GlyphRectangle.Left is > 143 and < 145 &&
            pdf.GetPage(2).Height - letter.GlyphRectangle.Top < 80);
    }

    [Theory]
    [InlineData("WordVisibleMirrorMarginsAllStories.docx", 144)]
    [InlineData("WordVisibleMirrorGutterAllStories.docx", 180)]
    public void FlowVisitorMirrorsInsideMarginAndGutter(string fixture, double firstPageLeft)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixture);

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        foreach (var (pageNumber, expectedLeft) in new[] { (1, firstPageLeft), (2, 72.0) })
        {
            var page = pdf.GetPage(pageNumber);
            Assert.Contains(page.Letters, letter => letter.Value == "H" &&
                letter.GlyphRectangle.Left > expectedLeft - 1 &&
                letter.GlyphRectangle.Left < expectedLeft + 1 &&
                page.Height - letter.GlyphRectangle.Top < 80);
        }
    }

    [Theory]
    [InlineData("WordStyledEmptySpacingAllStories.docx", 28, 40, 26)]
    [InlineData("WordFinalEmptyMarkAllStories.docx", 14, 14, 26)]
    public void FlowVisitorRetainsSpacingAtContextualReset(string fixture,
        double firstGap, double secondGap, double thirdGap)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixture);

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var bodyStarts = page.Letters.Where(letter => letter.Value == "B" &&
                page.Height - letter.GlyphRectangle.Top is > 80 and < 250)
            .Select(letter => page.Height - letter.GlyphRectangle.Top)
            .OrderBy(top => top).ToArray();
        Assert.Equal(4, bodyStarts.Length);
        Assert.InRange(bodyStarts[1] - bodyStarts[0], firstGap - 0.5, firstGap + 0.5);
        Assert.InRange(bodyStarts[2] - bodyStarts[1], secondGap - 0.5, secondGap + 0.5);
        Assert.InRange(bodyStarts[3] - bodyStarts[2], thirdGap - 0.5, thirdGap + 0.5);
    }

    [Fact]
    public void FlowVisitorReservesTrailingEmptyMarkStories()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordFinalEmptyMarkAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.Letters, letter => letter.Value == "B" &&
            page.Height - letter.GlyphRectangle.Top is > 108 and < 113);
        Assert.Contains(page.Letters, letter => letter.Value == "F" &&
            page.Height - letter.GlyphRectangle.Top is > 707 and < 714);
    }

    [Fact]
    public void FlowVisitorKeepsContextualSpacingAcrossSectionBoundary()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordContextualSpacingSections.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(2);
        Assert.Contains(page.Letters, letter => letter.Value == "B" &&
            page.Height - letter.GlyphRectangle.Top is > 89 and < 94);
        Assert.Contains(page.Letters, letter => letter.Value == "F" &&
            page.Height - letter.GlyphRectangle.Top is > 715 and < 723);
    }

    [Theory]
    [InlineData("WordContextualSpacingAllStories.docx", 102.4)]
    [InlineData("WordContextualSpacingInheritedAllStories.docx", 102.4)]
    [InlineData("WordStyledEmptySpacingAllStories.docx", 116.4)]
    public void FlowVisitorReservesContextualHeaderSpacing(string fixture, double firstBodyTop)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", fixture);

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var bodyTop = page.Letters.Where(letter => letter.Value == "B" &&
                page.Height - letter.GlyphRectangle.Top is > 80 and < 250)
            .Min(letter => page.Height - letter.GlyphRectangle.Top);
        Assert.InRange(bodyTop, firstBodyTop - 2, firstBodyTop + 2);
    }

    [Fact]
    public void FlowVisitorReservesHeightOfSizedEmptyHeaderParagraph()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleSizedEmptyParagraphsAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var bodyTop = page.Letters.Where(letter => letter.Value == "F" &&
                page.Height - letter.GlyphRectangle.Top is > 80 and < 150)
            .Min(letter => page.Height - letter.GlyphRectangle.Top);
        Assert.InRange(bodyTop, 85, 89);
    }

    [Fact]
    public void FlowVisitorReservesExplicitSpacingOfEmptyHeaderParagraph()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordEmptyParagraphSpacingAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var bodyTop = page.Letters.Where(letter => letter.Value == "F" &&
                page.Height - letter.GlyphRectangle.Top is > 90 and < 140)
            .Min(letter => page.Height - letter.GlyphRectangle.Top);
        Assert.InRange(bodyTop, 102, 107);
    }

    [Fact]
    public void FlowVisitorPreservesSpacingAfterBodyStyledHeaderParagraph()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedCapsResetAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var headerLastTop = page.Letters.Where(letter => letter.Value == "H" &&
                page.Height - letter.GlyphRectangle.Top is > 60 and < 90)
            .Min(letter => page.Height - letter.GlyphRectangle.Top);
        var bodyTop = page.Letters.Where(letter => letter.Value == "F" &&
                page.Height - letter.GlyphRectangle.Top is > 90 and < 130)
            .Min(letter => page.Height - letter.GlyphRectangle.Top);
        Assert.InRange(headerLastTop, 73, 78);
        Assert.InRange(bodyTop, 102.8, 103.8);
    }

    [Fact]
    public void FlowVisitorReservesMixedNumberedHeaderWithTrailingTab()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleDefaultTabAndTextAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var bodyTop = page.Letters.Where(letter => letter.Value == "F" &&
                page.Height - letter.GlyphRectangle.Top is > 80 and < 140)
            .Min(letter => page.Height - letter.GlyphRectangle.Top);
        Assert.InRange(bodyTop, 92, 94);
    }


    [Fact]
    public void FlowVisitorResolvesThemeTintAndShadeRunColors()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleThemeTintShadeAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.Letters, letter => letter.Value == "H" &&
            letter.Color.ToRGBValues().r > 0.7 && letter.Color.ToRGBValues().g < 0.4);
        Assert.Contains(page.Letters, letter => letter.Value == "F" &&
            letter.Color.ToRGBValues().g < 0.3 && letter.Color.ToRGBValues().b < 0.2);
        Assert.Contains(page.Letters, letter => letter.Value == "B" &&
            letter.Color.ToRGBValues().b > 0.8 && letter.Color.ToRGBValues().g > 0.6);
    }

    [Fact]
    public void FlowVisitorAdjustsElevenPointAptosPaintWithoutChangingOtherFonts()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var document = new PdfSharp.Pdf.PdfDocument();
        var page = document.AddPage();
        using (var graphics = XGraphics.FromPdfPage(page))
        {
            graphics.DrawString("Aptos", new XFont("Arial", 11), XBrushes.Black, 72, 72);
            graphics.DrawString("Other", new XFont("Times New Roman", 11), XBrushes.Black, 72, 96);
        }
        var fonts = page.Resources.Elements.GetDictionary("/Font")!;
        var aptosKey = fonts.Elements.Keys.First(key =>
        {
            var entry = fonts.Elements[key];
            var font = (entry as PdfSharp.Pdf.Advanced.PdfReference)?.Value as PdfSharp.Pdf.PdfDictionary ??
                entry as PdfSharp.Pdf.PdfDictionary;
            return font?.Elements.GetName("/BaseFont").Contains("Arial", StringComparison.Ordinal) == true;
        });
        var aptosEntry = fonts.Elements[aptosKey];
        var aptosFont = (aptosEntry as PdfSharp.Pdf.Advanced.PdfReference)?.Value as PdfSharp.Pdf.PdfDictionary ??
            (PdfSharp.Pdf.PdfDictionary)aptosEntry!;
        aptosFont.Elements.SetName("/BaseFont", "/Aptos");

        DxpFlowPdfVisitor.MatchElevenPointAptosPaint(document);

        var content = string.Join('\n', page.Contents.Elements.Select(item =>
            System.Text.Encoding.Latin1.GetString(
                (((PdfSharp.Pdf.Advanced.PdfReference)item).Value as
                    PdfSharp.Pdf.Advanced.PdfContent)!.Stream.Value)));
        Assert.Contains($"{aptosKey} 11.04 Tf", content);
        Assert.Contains(" 11 Tf", content);
    }


    [Fact]
    public void FlowVisitorAppliesDocumentThemeColorMapping()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleThemeColorMappingAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains(page.Letters, letter => letter.Value == "H" &&
            letter.Color.ToRGBValues() is var color && color.r > 0.6 && color.b > 0.5);
        Assert.Contains(page.Letters, letter => letter.Value == "B" &&
            letter.Color.ToRGBValues() is var color && color.g > 0.5 && color.b > 0.7);
    }

    [Fact]
    public void FlowVisitorCapitalizesTextualListLabelsAndAlignsHeaderFooterItems()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleTextualListFormatsAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains("One.", page.Text);
        Assert.Contains("Two.", page.Text);
        Assert.Contains("First.", page.Text);
        Assert.Contains("Second.", page.Text);
        Assert.True(page.Letters.Count(letter => letter.Value == "F" &&
            letter.GlyphRectangle.Left is > 305 and < 307) >= 2);
        Assert.True(page.Letters.Count(letter => letter.Value == "S" &&
            letter.GlyphRectangle.Left is > 305 and < 307) >= 2);
        Assert.Contains(page.Letters, letter => letter.Value == "F" &&
            letter.GlyphRectangle.Left is > 143.5 and < 145);
        Assert.Contains(page.Letters, letter => letter.Value == "S" &&
            letter.GlyphRectangle.Left is > 143.5 and < 145);
        Assert.Contains(page.Letters, letter => letter.Value == "T" &&
            letter.GlyphRectangle.Left is > 143.5 and < 145);
    }

    [Fact]
    public void FlowVisitorResolvesStartAndEndAlignmentFromStyledBidi()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleStyledBidiStyledAlignmentAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.All(page.Letters.Where(letter => letter.Value is "H" or "B" or "F"),
            letter => Assert.InRange(letter.GlyphRectangle.Left, 420, 545));
        var bodyTops = page.Letters.Where(letter => letter.Value == "B" &&
            page.Height - letter.GlyphRectangle.Top is > 70 and < 140)
            .Select(letter => page.Height - letter.GlyphRectangle.Top)
            .DistinctBy(top => Math.Round(top, 1)).OrderBy(top => top).ToArray();
        Assert.Equal(2, bodyTops.Length);
        Assert.InRange(bodyTops[0], 82, 85);
        Assert.InRange(bodyTops[1], 105, 108);
        var footerDigits = page.Letters.Where(letter => letter.Value == "7" &&
            letter.GlyphRectangle.Bottom < 100).Min(letter => letter.GlyphRectangle.Left);
        var footerHebrew = page.Letters.Single(letter => letter.Value == "ם" &&
            letter.GlyphRectangle.Bottom < 100);
        Assert.True(footerDigits < footerHebrew.GlyphRectangle.Left);
        Assert.Contains(page.Paths, path => path.IsStroked &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left is > 430 and < 435 && bounds.Width > 104 &&
            page.Height - bounds.Top is > 738 and < 745);
    }

    [Fact]
    public void StyledBidiFooterInheritsCharacterStrike()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleStyledBidiStyledAlignmentAllStories.docx");
        using var document = WordprocessingDocument.Open(source, false);
        var footer = document.MainDocumentPart!.FooterParts.SelectMany(part =>
            part.Footer.Descendants<Paragraph>()).First(paragraph =>
            paragraph.InnerText.Contains("Footer reset", StringComparison.Ordinal));

        var style = new DxpStyleResolver(document).ResolveRunStyle(footer,
            footer.Elements<Run>().First());

        Assert.True(style.Strike);
        Assert.False(style.DoubleStrike);
    }

    [Theory]
    [InlineData("Footer שלום 789", "Footer 789 םולש")]
    [InlineData("Header مرحبا 456", "Header 456 ابحرم")]
    public void FlowVisitorOrdersMixedRtlGroupsForPdf(string logical, string visual)
    {
        Assert.Equal(visual, DxpFlowPdfVisitor.ReorderRtlForPdf(logical));
    }

    [Theory]
    [InlineData("Body مرحبا 123", "123 مرحبا Body")]
    [InlineData("Footer שלום 789", "789 םולש Footer")]
    public void FlowVisitorOrdersExplicitRtlRunsForPdf(string logical, string visual)
    {
        Assert.Equal(visual, DxpFlowPdfVisitor.ReorderExplicitRtlRunForPdf(logical));
    }

    [Theory]
    [InlineData("مرحبا", "ﻣﺮﺣﺒﺎ")]
    [InlineData("Body مرحبا 123", "Body ﻣﺮﺣﺒﺎ 123")]
    [InlineData("باب", "ﺑﺎﺏ")]
    [InlineData("سمك", "ﺳﻤﻚ")]
    public void FlowVisitorShapesArabicJoiningForms(string logical, string visual)
    {
        Assert.Equal(visual, DxpFlowPdfVisitor.ShapeExplicitArabicForPdf(logical));
    }

    [Theory]
    [InlineData("Body ﻣﺮﺣﺒﺎ 123", "Body ﺎﺒﺣﺮﻣ 123")]
    [InlineData("Header 456 ﻣﺮﺣﺒﺎ", "Header 456 ﺎﺒﺣﺮﻣ")]
    public void FlowVisitorOrdersShapedArabicForPdf(string shaped, string painted)
    {
        Assert.Equal(painted, DxpFlowPdfVisitor.OrderShapedArabicForPdf(shaped));
    }

    [Theory]
    [InlineData("Header مرحبا 456", "Header 456 ﻣﺮﺣﺒﺎ")]
    [InlineData("Body مرحبا 123", "Body 123 ﻣﺮﺣﺒﺎ")]
    public void FlowVisitorKeepsNumericGroupsBeforeImplicitShapedArabic(string logical, string visual)
    {
        Assert.Equal(visual, DxpFlowPdfVisitor.ShapeImplicitArabicForPdf(logical));
    }

    [Fact]
    public void FlowVisitorPlacesExplicitRtlRunsInWordVisualOrder()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunDirectionOverridesAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var body = page.Letters.First(letter => letter.Value == "B");
        var numeral = page.Letters.First(letter => letter.Value == "1");
        Assert.True(numeral.GlyphRectangle.Left < body.GlyphRectangle.Left);
        var bodyArabic = page.Letters.Where(letter =>
            letter.GlyphRectangle.Left > numeral.GlyphRectangle.Left + 10 &&
            letter.GlyphRectangle.Left < body.GlyphRectangle.Left - 5 &&
            Math.Abs(letter.StartBaseLine.Y - body.StartBaseLine.Y) < 3).ToArray();
        Assert.True(bodyArabic.Length >= 4);
        Assert.All(bodyArabic, letter => Assert.Contains("Arial", letter.FontName));
        Assert.InRange(body.PointSize, 11.9, 12.1);
        var footer = page.Letters.First(letter => letter.Value == "F");
        var footerNumber = page.Letters.First(letter => letter.Value == "7");
        Assert.True(footerNumber.GlyphRectangle.Left < footer.GlyphRectangle.Left);
        Assert.InRange(footer.PointSize, 11.9, 12.1);
        Assert.DoesNotContain("Bold", footer.FontName, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Italic", footer.FontName, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void FlowVisitorUsesExplicitComplexScriptWeightAndSlantForRtlRuns()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        using var source = new MemoryStream(File.ReadAllBytes(Path.Combine(root,
            "DocxportNet.Tests", "Fixtures", "Doc", "WordRunDirectionOverridesAllStories.docx")));
        using (var document = WordprocessingDocument.Open(source, true))
        {
            var run = document.MainDocumentPart!.FooterParts
                .SelectMany(part => part.Footer!.Descendants<Run>())
                .First(candidate => candidate.InnerText.Contains("Footer", StringComparison.Ordinal));
            run.RunProperties!.Append(new BoldComplexScript(), new ItalicComplexScript());
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var footer = pdf.GetPage(1).Letters.First(letter => letter.Value == "F");
        Assert.Contains("Bold", footer.FontName, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Italic", footer.FontName, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void FlowVisitorAlignsUnjustifiedParagraphsWithEffectiveBidi()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleStyledBidiPrecedenceAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var initialLetters = pdf.GetPage(1).Letters.Where(letter =>
            letter.Value is "H" or "B" or "F").ToArray();
        Assert.True(initialLetters.Count(letter => letter.GlyphRectangle.Left is > 70 and < 90) >= 3);
        Assert.True(initialLetters.Count(letter => letter.GlyphRectangle.Left is > 420 and < 550) >= 3);
    }


    [Fact]
    public void EffectiveRunColorHonorsDirectAutomaticReset()
    {
        var style = new DxpEffectiveRunStyleBuilder();
        DxpEffectiveRunStyleBuilder.ApplyStyleRunProperties(
            new StyleRunProperties(new Color { Val = "CC00CC" }), null, ref style);
        Assert.Equal("#CC00CC", style.ToImmutable().ColorCss);

        DxpEffectiveRunStyleBuilder.ApplyRunProperties(
            new RunProperties(new Color { Val = "auto" }), null, ref style);
        Assert.Null(style.ToImmutable().ColorCss);
    }

    [Fact]
    public void ConditionalTableTextStyleResolvesCornersAndEdges()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", "WordStyledCellSpacingResetAllStories.docx");
        using var document = WordprocessingDocument.Open(source, false);
        var table = document.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        var resolver = new DxpTableStyleResolver(document);
        var resolved = resolver.ResolveTableStyle(table.TableProperties, 2, 2);
        var firstRow = table.Elements<TableRow>().First();
        var firstCell = DxpTableStyleResolver.ComputeCellStyle(resolved,
            firstRow.Elements<TableCell>().First().TableCellProperties, 0, 0, 1, 1);
        var lastCell = DxpTableStyleResolver.ComputeCellStyle(resolved,
            firstRow.Elements<TableCell>().Last().TableCellProperties, 0, 1, 1, 1);

        Assert.Equal("#CC00CC", firstCell.ConditionalTextColorCss);
        Assert.True(firstCell.ConditionalBold);
        Assert.True(firstCell.ConditionalItalic);
        Assert.Equal("#FF8800", lastCell.ConditionalTextColorCss);
        Assert.True(lastCell.ConditionalUnderline);
        Assert.Equal(DxpComputedTextAlign.Right, lastCell.ConditionalTextAlign);
    }

    [Fact]
    public void DirectClearShadingResetsInheritedTableFillButNilDoesNot()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
        var source = Path.Combine(root, "DocxportNet.Tests", "Fixtures", "Doc", "WordDefaultShadingResetsAllStories.docx");
        using var document = WordprocessingDocument.Open(source, false);
        var table = document.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        var resolver = new DxpTableStyleResolver(document);
        var resolved = resolver.ResolveTableStyle(table.TableProperties, 2, 2);
        var cells = table.Elements<TableRow>().First().Elements<TableCell>().ToArray();

        var inherited = DxpTableStyleResolver.ComputeCellStyle(resolved, cells[0].TableCellProperties, 0, 0, 1, 1);
        var cleared = DxpTableStyleResolver.ComputeCellStyle(resolved, cells[1].TableCellProperties, 0, 1, 1, 1);

        Assert.Equal("#00FF00", inherited.BackgroundColorCss);
        Assert.Null(cleared.BackgroundColorCss);
    }

    private sealed class CustomerNameResolver : IDxpFieldValueResolver
    {
        public Task<DxpFieldValue?> ResolveAsync(string name, DxpFieldValueKindHint kind,
            DxpFieldEvalContext context) => Task.FromResult<DxpFieldValue?>(
                kind == DxpFieldValueKindHint.MergeField && name == "CustomerName"
                    ? new DxpFieldValue("Taylor") : null);
    }
}
