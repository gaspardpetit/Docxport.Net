using System.Buffers.Binary;
using System.Text;
using System.Text.RegularExpressions;
using System.Text;
using System.Text.RegularExpressions;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Fields;
using DocxportNet.Visitors.Pdf;
using DocxportNet.Walker.Context;
using DocxportNet.Walker.Context;
using DocxportNet.Walker.Context;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace DocxportNet.Tests;

public sealed class PdfFlowExportTests
{
    [Fact]
    public void SplitsFootnoteParagraphWithStyledRepeatedSpaces()
    {
        var paragraph = new MigraDoc.DocumentObjectModel.Document().AddSection().AddParagraph();
        var span = paragraph.AddFormattedText("First ");
        span.AddSpace(2);
        span.AddText("second third");

        Assert.True(DxpFootnoteParagraphSplitter.TryGetText(paragraph, out var original));
        Assert.Equal("First   second third", original);
        Assert.True(DxpFootnoteParagraphSplitter.TrySplit(paragraph, 2, out var continuation));
        Assert.True(DxpFootnoteParagraphSplitter.TryGetText(paragraph, out var first));
        Assert.True(DxpFootnoteParagraphSplitter.TryGetText(continuation!, out var second));
        Assert.Equal("First   second ", first);
        Assert.Equal("third", second);
    }

    [Fact]
    public void ReplacesStyleRefCacheContainingStyledRepeatedSpaces()
    {
        var paragraph = new MigraDoc.DocumentObjectModel.Document().AddSection().AddParagraph();
        var span = paragraph.AddFormattedText("Heading:First ");
        span.AddSpace(2);
        span.AddText("title suffix");
        var method = typeof(DxpFlowPdfVisitor).GetMethod("ReplaceStyleRefText",
            System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic)!;

        Assert.True((bool)method.Invoke(null, [paragraph, "First   title", "Second heading"])!);
        Assert.Equal("Heading:Second heading suffix", string.Concat(span.Elements.Select(element => element switch
        {
            MigraDoc.DocumentObjectModel.Text text => text.Content,
            MigraDoc.DocumentObjectModel.Character { SymbolName: MigraDoc.DocumentObjectModel.SymbolName.Blank } spaces =>
                new string(' ', spaces.Count),
            _ => ""
        })));
    }

    [Fact]
    public void PreservesRepeatedSpacesInPageSpecificStyleRefReplacement()
    {
        var paragraph = new MigraDoc.DocumentObjectModel.Document().AddSection().AddParagraph();
        var span = paragraph.AddFormattedText("Heading:First title suffix");
        var method = typeof(DxpFlowPdfVisitor).GetMethod("ReplaceStyleRefText",
            System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic)!;

        Assert.True((bool)method.Invoke(null, [paragraph, "First title", "Second   heading"])!);
        Assert.Equal("Heading:Second   heading suffix", string.Concat(span.Elements.Select(element => element switch
        {
            MigraDoc.DocumentObjectModel.Text text => text.Content,
            MigraDoc.DocumentObjectModel.Character { SymbolName: MigraDoc.DocumentObjectModel.SymbolName.Blank } spaces =>
                new string(' ', spaces.Count),
            _ => ""
        })));
        Assert.Contains(span.Elements, element => element is MigraDoc.DocumentObjectModel.Character
            { SymbolName: MigraDoc.DocumentObjectModel.SymbolName.Blank, Count: 2, Tag: not null });
    }

    [Fact]
    public void CollapsesStyledRepeatedSpacesOnMirroredPageClone()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            static Paragraph Paragraph(bool nextPage) => new(
                new ParagraphProperties(new MirrorIndents(), new Justification { Val = JustificationValues.Center },
                    nextPage ? new PageBreakBefore() : new KeepNext()),
                new Run(new RunProperties(new Bold()), new Text("Alpha   beta gamma delta") { Space = SpaceProcessingModeValues.Preserve }));
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(Paragraph(false), Paragraph(true),
                new SectionProperties(new PageSize { Width = 2100, Height = 5000 },
                    new PageMargin { Left = 600, Right = 600, Top = 600, Bottom = 600 })));
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        var first = pdf.GetPage(1).Letters.First(letter => letter.Value == "A");
        var second = pdf.GetPage(2).Letters.First(letter => letter.Value == "A");
        var wrapped = pdf.GetPage(1).Letters.First(letter => letter.Value == "b");
        Assert.True(wrapped.GlyphRectangle.Bottom < first.GlyphRectangle.Bottom);
        Assert.InRange(Math.Abs(first.GlyphRectangle.Left - second.GlyphRectangle.Left), 0, 0.2);
    }

    [Fact]
    public void StyledParagraphKeepsEachRepeatedSpaceAdvance()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "features", "StyledRepeatedSpaces.docx");
        var reference = Path.Combine(root, "truth", "pdf-word", "features", "StyledRepeatedSpaces.pdf");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        static double[] BetaOrigins(byte[] pdfBytes)
        {
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(pdfBytes));
            return pdf.GetPage(1).GetWords().Where(word => word.Text == "Beta")
                .OrderByDescending(word => word.BoundingBox.Top)
                .Select(word => word.BoundingBox.Left).ToArray();
        }

        var word = BetaOrigins(File.ReadAllBytes(reference));
        var flow = BetaOrigins(bytes);
        Assert.Equal(3, word.Length);
        Assert.Equal(3, flow.Length);
        var wordExtraBefore = word[1] - word[0];
        var wordExtraAfter = word[2] - word[0];
        Assert.InRange(wordExtraBefore, 3.0, 4.0);
        Assert.InRange(wordExtraAfter, 3.0, 4.0);
        Assert.InRange(flow[1] - flow[0], wordExtraBefore - 1.0, wordExtraBefore + 1.0);
        Assert.InRange(flow[2] - flow[0], wordExtraAfter - 1.0, wordExtraAfter + 1.0);
    }

    [Fact]
    public void AllCapsStyledParagraphKeepsRepeatedSpaceAdvance()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "features", "StyledRepeatedSpaces.docx");
        var reference = Path.Combine(root, "truth", "pdf-word", "features", "StyledRepeatedSpaces.pdf");
        var flow = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        static double[] Origins(byte[] bytes)
        {
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
            return pdf.GetPage(1).GetWords().Where(word => word.Text == "SIGMA")
                .OrderByDescending(word => word.BoundingBox.Top)
                .Select(word => word.BoundingBox.Left).ToArray();
        }

        var word = Origins(File.ReadAllBytes(reference));
        var actual = Origins(flow);
        Assert.Equal(2, word.Length);
        Assert.Equal(2, actual.Length);
        Assert.InRange(word[1] - word[0], 2, 6);
        Assert.InRange(actual[1] - actual[0], word[1] - word[0] - 1, word[1] - word[0] + 1);
    }

    [Theory]
    [InlineData("Delta")]
    [InlineData("Omega")]
    [InlineData("Theta")]
    public void RepeatedSpacesBeforeWrappedWordsDoNotIndentTheNextLine(string word)
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "features", "StyledRepeatedSpaces.docx");
        var reference = Path.Combine(root, "truth", "pdf-word", "features", "StyledRepeatedSpaces.pdf");
        var flow = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        static double[] Origins(byte[] pdfBytes, string text)
        {
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(pdfBytes));
            return pdf.GetPage(1).GetWords().Where(candidate => candidate.Text == text)
                .OrderByDescending(candidate => candidate.BoundingBox.Top)
                .Select(candidate => candidate.BoundingBox.Left).ToArray();
        }

        var wordOrigins = Origins(File.ReadAllBytes(reference), word);
        var flowOrigins = Origins(flow, word);
        Assert.Equal(2, wordOrigins.Length);
        Assert.Equal(2, flowOrigins.Length);
        Assert.InRange(Math.Abs(wordOrigins[1] - wordOrigins[0]), 0, 0.5);
        Assert.InRange(Math.Abs(flowOrigins[1] - flowOrigins[0]), 0, 0.5);
    }

    [Fact]
    public void BehindTextGradientRectangleIsRenderedUnderShadedTable()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "features", "mixed-shaded-behind.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Single(pdf.Pages);
        Assert.NotNull(pdf.Pages[0].Resources.Elements.GetDictionary("/Pattern"));
        using var text = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Contains("Mixedrails", text.GetPage(1).Text);
    }

    [Theory]
    [InlineData("foreground")]
    [InlineData("interiorStop")]
    [InlineData("shade")]
    public void UnsupportedBehindTextGradientIsNotRendered(string variation)
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "features", "mixed-shaded-behind.docx");
        using var docx = new MemoryStream();
        using (var input = File.OpenRead(source)) input.CopyTo(docx);
        docx.Position = 0;
        using (var document = WordprocessingDocument.Open(docx, true))
        {
            var choice = document.MainDocumentPart!.Document.Descendants<AlternateContentChoice>().Single();
            var element = variation switch
            {
                "foreground" => choice.Descendants().First(child => child.LocalName == "anchor"),
                "interiorStop" => choice.Descendants().First(child => child.LocalName == "gs"),
                _ => choice.Descendants().First(child => child.LocalName == "shade")
            };
            var attribute = variation switch
            {
                "foreground" => "behindDoc",
                "interiorStop" => "pos",
                _ => "val"
            };
            element.SetAttribute(new OpenXmlAttribute(attribute, "", variation == "foreground" ? "0" : "25000"));
            document.MainDocumentPart.Document.Save();
        }
        var bytes = DxpExport.ExportToBytes(docx.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Null(pdf.Pages[0].Resources.Elements.GetDictionary("/Pattern"));
    }

    [Fact]
    public void MixedBorderBandsKeepCellTextOnTheWordGrid()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "features", "mixed-shaded-behind.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var words = page.GetWords().ToArray();
        var expectedX = new[] { 57.6, 208.85, 362.23, 57.6, 211.01, 360.07, 59.76, 208.85, 360.07 };
        var expectedY = new[] { 149.25, 254.28, 359.28 };
        for (var row = 1; row <= 3; row++)
            for (var column = 1; column <= 3; column++)
            {
                var box = Assert.Single(words, word => word.Text == $"{row},{column}").BoundingBox;
                Assert.InRange(box.Left, expectedX[(row - 1) * 3 + column - 1] - 0.5,
                    expectedX[(row - 1) * 3 + column - 1] + 0.5);
                Assert.InRange(page.Height - box.Top, expectedY[row - 1] - 0.75,
                    expectedY[row - 1] + 0.75);
            }
    }

    [Fact]
    public void ModernCompatibilityModeAlignsTableBorderWithTextMargin()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "features", "mixed-shaded-behind.docx");
        using var stream = new MemoryStream();
        using (var input = File.OpenRead(source)) input.CopyTo(stream);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var setting = Assert.Single(document.MainDocumentPart!.DocumentSettingsPart!.Settings!
                .Descendants<CompatibilitySetting>(), candidate =>
                    candidate.Name?.Value == CompatSettingNameValues.CompatibilityMode);
            setting.Val = "15";
        }

        var bytes = DxpExport.ExportToBytes(stream.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var cellText = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "1,1");
        Assert.InRange(cellText.BoundingBox.Left, 63.5, 64.25);
    }

    [Theory]
    [InlineData(6, "dotDash")]
    [InlineData(8, "triple")]
    [InlineData(19, "doubleWave")]
    [InlineData(26, "double")]
    public void TableBorderStyleRetainsWordValueForPdfRendering(int tableIndex, string expected)
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "features", "WordBorderStyles3x3.docx");
        using var document = WordprocessingDocument.Open(source, false);
        var table = document.MainDocumentPart!.Document.Body!.Elements<Table>().ElementAt(tableIndex);
        var resolved = new DxpTableStyleResolver(document).ResolveTableStyle(table.TableProperties, 3, 3);
        var firstCell = table.Elements<TableRow>().First().Elements<TableCell>().First();
        var computed = DxpTableStyleResolver.ComputeCellStyle(resolved, firstCell.TableCellProperties,
            0, 0, 3, 3);
        Assert.Equal(expected, computed.Borders?.Top?.WordBorderStyle);
    }

    [Fact]
    public void CenteredFixedTablesUseCellWidthsAndOuterBorderExtent()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "features", "WordBorderStyles3x3.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        double? referenceMiddle = null;
        foreach (var pageNumber in Enumerable.Range(1, 8).Concat(Enumerable.Range(21, 6)))
        {
            var words = pdf.GetPage(pageNumber).GetWords().ToArray();
            double Center(string label)
            {
                var box = Assert.Single(words, word => word.Text == label).BoundingBox;
                return (box.Left + box.Right) / 2;
            }

            var first = Center("1,1");
            var middle = Center("1,2");
            var last = Center("1,3");
            Assert.InRange(middle - first, 148.5, 151.0);
            Assert.InRange(last - middle, 148.5, 151.0);
            Assert.InRange(middle, 305.5, 306.5);
            referenceMiddle ??= middle;
            Assert.InRange(middle - referenceMiddle.Value, -0.1, 0.1);
        }
    }

    [Fact]
    public void CellLevelDoubleBordersDrawBothHorizontalRails()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "features", "WordBorderStyles3x3.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());
        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var content = string.Concat(pdf.Pages[26].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Select(reference => reference.Value)
            .OfType<PdfSharp.Pdf.Advanced.PdfContent>()
            .Where(stream => stream.Stream != null)
            .Select(stream => Encoding.Latin1.GetString(stream.Stream!.UnfilteredValue)));
        Assert.True(Regex.Matches(content, @"(?m)[\d.]+ [\d.]+ [\d.]+ [\d.]+ re\s+S$").Count >= 10);
    }

    [Theory]
    [InlineData(null, null, 148.5, 151.0)]
    [InlineData("dxa", null, 160.5, 162.5)]
    [InlineData("dxa", "dxa", 148.5, 151.0)]
    [InlineData("pct", null, 148.5, 151.0)]
    [InlineData("dxa", "pct", 148.5, 151.0)]
    public void CenteredTableWidthsRespectInheritedCellSpacingAndDirectOverride(
        string? inheritedSpacingType, string? directSpacingType, double minimumColumnDistance,
        double maximumColumnDistance)
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "features", "WordBorderStyles3x3.docx");
        var copy = Path.Combine(Path.GetTempPath(), $"styled-table-width-{Guid.NewGuid():N}.docx");
        File.Copy(source, copy);
        try
        {
            using (var document = WordprocessingDocument.Open(copy, true))
            {
                var styles = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                var baseStyle = new Style { Type = StyleValues.Table, StyleId = "WidthBase" };
                if (inheritedSpacingType != null)
                    baseStyle.AppendChild(new StyleTableProperties(new TableCellSpacing
                    {
                        Width = "120", Type = inheritedSpacingType == "pct"
                            ? TableWidthUnitValues.Pct : TableWidthUnitValues.Dxa
                    }));
                styles.AppendChild(baseStyle);
                styles.AppendChild(new Style(new BasedOn { Val = "WidthBase" })
                {
                    Type = StyleValues.Table, StyleId = "WidthDerived"
                });
                var table = document.MainDocumentPart.Document.Body!.Descendants<Table>().First();
                table.TableProperties!.PrependChild(new TableStyle { Val = "WidthDerived" });
                if (directSpacingType != null)
                    table.TableProperties.AppendChild(new TableCellSpacing
                    {
                        Width = directSpacingType == "pct" ? "120" : "0",
                        Type = directSpacingType == "pct"
                            ? TableWidthUnitValues.Pct : TableWidthUnitValues.Dxa
                    });
            }

            var bytes = DxpExport.ExportToBytes(copy, new DxpFlowPdfVisitor());
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
            var words = pdf.GetPage(1).GetWords().ToArray();
            double Center(string label)
            {
                var box = Assert.Single(words, word => word.Text == label).BoundingBox;
                return (box.Left + box.Right) / 2;
            }
            Assert.InRange(Center("1,2") - Center("1,1"), minimumColumnDistance,
                maximumColumnDistance);
        }
        finally
        {
            File.Delete(copy);
        }
    }

    [Fact]
    public void DerivedTableStyleUsesInheritedFillForSpacedCellBorders()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "features", "WordBorderStyles3x3.docx");
        var copy = Path.Combine(Path.GetTempPath(), $"styled-table-fill-{Guid.NewGuid():N}.docx");
        File.Copy(source, copy);
        try
        {
            using (var document = WordprocessingDocument.Open(copy, true))
            {
                var styles = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.AppendChild(new Style(new StyleTableProperties(
                    new DocumentFormat.OpenXml.Wordprocessing.Shading { Fill = "21C57D" },
                    new TableCellSpacing { Width = "120", Type = TableWidthUnitValues.Dxa }))
                {
                    Type = StyleValues.Table, StyleId = "FillBase"
                });
                styles.AppendChild(new Style(new BasedOn { Val = "FillBase" })
                {
                    Type = StyleValues.Table, StyleId = "FillDerived"
                });
                var table = document.MainDocumentPart.Document.Body!.Descendants<Table>().First();
                table.TableProperties!.PrependChild(new TableStyle { Val = "FillDerived" });
            }

            var bytes = DxpExport.ExportToBytes(copy, new DxpFlowPdfVisitor());
            using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
            var content = string.Concat(pdf.Pages[0].Contents.Elements
                .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
                .Select(reference => reference.Value)
                .OfType<PdfSharp.Pdf.Advanced.PdfContent>()
                .Where(stream => stream.Stream != null)
                .Select(stream => Encoding.Latin1.GetString(stream.Stream!.UnfilteredValue)));
            Assert.Contains(".129 .773 .49 RG", content);
        }
        finally
        {
            File.Delete(copy);
        }
    }

    [Fact]
    public void EvaluatedMergeFieldsUseWordPlaceholdersInMarginStories()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "legacy", "WordMergeFieldsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(),
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate });

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Contains("«CustomerName»", pdf.GetPage(1).Text);
        Assert.DoesNotContain("Alice", pdf.GetPage(2).Text);
    }

    [Fact]
    public void ParagraphAnchoredTiffStaysBesideBodyText()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "legacy", "WordFloatingTiffAllStories.docx");
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

    [Fact]
    public void CachedAskResultsAreNotDisplayedInPdf()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "samples", "TestDocVariables.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(),
            options: new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Cache });

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var text = pdf.GetPage(2).Text;
        Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(text, "Bob").Count);
        Assert.Equal(1, System.Text.RegularExpressions.Regex.Matches(text, "Montreal").Count);
    }

    [Fact]
    public void OmitsCommentStoriesAndAnnotationsByDefault()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "samples", "TestComments.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var text = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Contains("BeforeComment", text.GetPage(1).Text);
        Assert.DoesNotContain("FirstReply", text.GetPage(1).Text);
        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.DoesNotContain(pdf.Pages.Cast<PdfPage>()
            .SelectMany(page => page.Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>()),
            annotation => annotation.Elements.GetName("/Subtype") is "/Text" or "/Highlight");
    }

    [Fact]
    public void ExportsCommentThreadsAsLinkedPdfAnnotations()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "samples", "TestComments.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(includePdfComments: true));

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var annotations = pdf.Pages.Cast<PdfPage>()
            .SelectMany(page => page.Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>())
            .ToArray();
        Assert.Equal(6, annotations.Length);
        var roots = annotations.Where(annotation => annotation.Elements.GetReference("/IRT") == null).ToArray();
        Assert.Equal(2, roots.Length);
        Assert.All(roots, annotation =>
        {
            Assert.Equal("/Highlight", annotation.Elements.GetName("/Subtype"));
            Assert.Equal(8, annotation.Elements.GetArray("/QuadPoints")!.Elements.Count);
        });
        var first = Assert.Single(roots.Where(annotation => annotation.Contents == "Comment 1"));
        var replies = annotations.Where(annotation => annotation.Elements.GetReference("/IRT") != null).ToArray();
        Assert.Equal(4, replies.Length);
        Assert.All(replies, reply =>
        {
            Assert.Equal("/Text", reply.Elements.GetName("/Subtype"));
            Assert.Equal("/R", reply.Elements.GetName("/RT"));
            Assert.Equal(first.Reference!.ObjectID, reply.Elements.GetReference("/IRT")!.ObjectID);
        });
    }

    [Fact]
    public void PlacesCommentAnnotationsOnTheirSourcePages()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "features", "CommentMultiPage.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(includePdfComments: true));

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Equal(2, pdf.PageCount);
        var first = Assert.Single(pdf.Pages[0].Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>());
        var second = Assert.Single(pdf.Pages[1].Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>());
        Assert.Contains("First page note", first.Contents);
        Assert.Contains("Second page note", second.Contents);
    }

    [Fact]
    public void ContinuesCommentHighlightAcrossPages()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "features", "CommentCrossPageRange.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(includePdfComments: true));

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Equal(2, pdf.PageCount);
        foreach (PdfPage page in pdf.Pages)
        {
            var highlight = Assert.Single(page.Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>());
            Assert.Equal("/Highlight", highlight.Elements.GetName("/Subtype"));
            Assert.Equal("Cross-page note", highlight.Contents);
            Assert.Equal(8, highlight.Elements.GetArray("/QuadPoints")!.Elements.Count);
        }
    }

    [Fact]
    public void ExcludesRepeatedHeadersAndFootersFromCommentHighlight()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "features", "CommentThreePageRange.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(includePdfComments: true));

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Equal(3, pdf.PageCount);
        foreach (PdfPage page in pdf.Pages)
        {
            var highlight = Assert.Single(page.Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>());
            Assert.Equal("/Highlight", highlight.Elements.GetName("/Subtype"));
            Assert.Equal("Three-page note", highlight.Contents);
            Assert.Equal(8, highlight.Elements.GetArray("/QuadPoints")!.Elements.Count);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LimitsCommentHighlightToItsTableCell(bool multiline)
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "features", "CommentMultiPage.docx");
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
        var highlight = Assert.Single(pdf.Pages[0].Annotations
            .Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>());
        Assert.Equal("/Highlight", highlight.Elements.GetName("/Subtype"));
        Assert.Equal(multiline ? 16 : 8, highlight.Elements.GetArray("/QuadPoints")!.Elements.Count);
        Assert.True(highlight.Rectangle.X2 < 300);
    }

    [Fact]
    public void CapitalizesTextualListLabelsInPdf()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "legacy",
            "WordVisibleTextualListFormatsAllStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        Assert.Contains("One.", page.Text);
        Assert.Contains("Two.", page.Text);
        Assert.Contains("First.", page.Text);
        Assert.Contains("Second.", page.Text);
    }

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
    public void ExplicitNextPageOddRestartInsertsBlankPageForOddEvenHeaders()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "legacy", "WordMixedGeometrySectionStories.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(4, pdf.NumberOfPages);
        Assert.Contains("Landscapebody", pdf.GetPage(1).Text);
        Assert.Empty(pdf.GetPage(2).Text);
        Assert.Contains("Portraitbody", pdf.GetPage(3).Text);
        Assert.Contains("Portraitlaterbody", pdf.GetPage(4).Text);
    }

    [Fact]
    public void DisabledOddEvenHeadersDoNotInsertBlankPageForOddRestart()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings(
                new EvenAndOddHeaders { Val = false });
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("First section")),
                    new ParagraphProperties(new SectionProperties())),
                new Paragraph(new Run(new Text("Second section"))),
                new SectionProperties(new SectionType { Val = SectionMarkValues.NextPage },
                    new PageNumberType { Start = 7 })));
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
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

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NegativeSectionMarginsMatchTheirPositiveDistances(bool bottomAligned)
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var negative = FirstLetterTop(-1100, -1200);
        var positive = FirstLetterTop(1100, 1200);

        Assert.InRange(Math.Abs(negative - positive), 0, 0.1);
        if (!bottomAligned) Assert.InRange(negative, 52, 60);

        double FirstLetterTop(int top, int bottom)
        {
            using var source = new MemoryStream();
            using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
            {
                var main = document.AddMainDocumentPart();
                var section = new SectionProperties(new PageMargin { Top = top, Bottom = bottom });
                if (bottomAligned)
                    section.Append(new VerticalTextAlignmentOnPage { Val = VerticalJustificationValues.Bottom });
                main.Document = new Document(new Body(
                    new Paragraph(new Run(new Text("Signed margin"))), section));
            }

            var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
            var page = Assert.Single(pdf.GetPages());
            return page.Height - page.Letters.First().GlyphRectangle.Top;
        }
    }

    [Fact]
    public void BottomAlignedNegativeMarginSectionKeepsItsBodyText()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "legacy", "WordPageBorderSectionSetup.docx");

        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("Second", pdf.GetPage(2).Text);
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
    public void ConvertsSkippedRle4PixelsToBlackBackground()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "legacy", "WordBmpRle4DeltaAllStories.docx");
        using var document = WordprocessingDocument.Open(source, false);
        using var image = Assert.Single(document.MainDocumentPart!.ImageParts).GetStream();
        using var data = new MemoryStream();
        image.CopyTo(data);

        var bitmap = DxpBmp24Converter.Normalize(data.ToArray());

        Assert.Equal(new byte[] { 0, 0, 0 }, bitmap.AsSpan(54 + 4 * 3, 3).ToArray());
        var black = 0;
        var cyan = 0;
        for (var offset = 54; offset < bitmap.Length; offset += 3)
        {
            var pixel = bitmap.AsSpan(offset, 3);
            if (pixel.SequenceEqual(new byte[] { 0, 0, 0 })) black++;
            else if (pixel.SequenceEqual(new byte[] { 255, 200, 0 })) cyan++;
            else Assert.Fail($"Unexpected RLE4 pixel at byte {offset}.");
        }
        Assert.Equal(192, black);
        Assert.Equal(64, cyan);
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

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OrdersEndnotesByReferenceRatherThanPartOrder(bool alternateContent)
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var first = new Paragraph(new Run(new Text("First reference")));
            if (alternateContent)
            {
                var choice = new AlternateContentChoice(new Run(new EndnoteReference { Id = 1 })) {
                    Requires = "w14"
                };
                choice.AddNamespaceDeclaration("w14",
                    "http://schemas.microsoft.com/office/word/2010/wordml");
                first.Append(new AlternateContent(choice,
                    new AlternateContentFallback(new Run(new EndnoteReference { Id = 2 }))));
            }
            else
                first.Append(new Run(new EndnoteReference { Id = 2 }));
            main.Document = new Document(new Body(first,
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

    [Fact]
    public void ConfiguresFragmentBorderReservationForSplitTables()
    {
        var visitor = new DxpFlowPdfVisitor();
        var table = new MigraDoc.DocumentObjectModel.Tables.Table();
        var tablesField = typeof(DxpFlowPdfVisitor).GetField("_splitOverpageTables",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!;
        var tables = (HashSet<MigraDoc.DocumentObjectModel.Tables.Table>)tablesField.GetValue(visitor)!;
        tables.Add(table);
        var createRenderer = typeof(DxpFlowPdfVisitor).GetMethod("CreatePdfRenderer",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!;
        var renderer = (MigraDoc.Rendering.PdfDocumentRenderer)createRenderer.Invoke(visitor,
            new object[] { new MigraDoc.DocumentObjectModel.Document() })!;

        var reservation = Assert.IsType<Func<MigraDoc.DocumentObjectModel.Tables.Table, int, double?>>(
            renderer.DocumentRenderer.TableFragmentTopBorderReservation);
        Assert.Equal(0.5, reservation(table, 1));
        Assert.NotNull(renderer.DocumentRenderer.WrapContour);
        Assert.NotNull(renderer.DocumentRenderer.WrapContourLineFontBandHeight);
        Assert.NotNull(renderer.DocumentRenderer.WrapLineFitHeight);
    }

    [Fact]
    public void PreservesCachedMergeFieldResultsInBodyAndUsesPlaceholdersInMarginStories()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "legacy", "WordMergeFieldsAllStories.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(),
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Cache });

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Contains("«CustomerName»", pdf.GetPage(1).Text);
        Assert.Contains("Alice", pdf.GetPage(2).Text);
    }

    [Fact]
    public void UpdatesArabicPageFieldWithoutChangingRomanCachedResult()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
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

    [Theory]
    [InlineData(DxpFieldEvalExportMode.Cache)]
    [InlineData(DxpFieldEvalExportMode.Evaluate)]
    public void ResolvesHeaderStyleReferenceOnEachPage(DxpFieldEvalExportMode mode)
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "legacy", "WordStyleReferenceAcrossPages.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(),
            new DxpExportOptions { FieldEvalMode = mode });

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("Heading:Openexamplesite", pdf.GetPage(1).Text);
        Assert.Contains("Heading:Secondheading", pdf.GetPage(2).Text);
        Assert.DoesNotContain("Heading:Secondheading", pdf.GetPage(1).Text);
    }

    [Theory]
    [InlineData(DxpFieldEvalExportMode.Cache)]
    [InlineData(DxpFieldEvalExportMode.Evaluate)]
    public void PreservesComplexFieldHyperlinksInFootnotes(DxpFieldEvalExportMode mode)
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var source = Path.Combine(root, "fixtures", "docx", "samples", "Right to Repair.docx");
        var bytes = DxpExport.ExportToBytes(source, new DxpFlowPdfVisitor(),
            new DxpExportOptions { FieldEvalMode = mode });

        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var links = pdf.Pages.Cast<PdfPage>()
            .SelectMany(page => page.Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>())
            .Select(annotation => annotation.Elements.GetDictionary("/A")?.Elements.GetString("/URI"))
            .Where(uri => !string.IsNullOrEmpty(uri)).ToArray();
        Assert.True(links.Length >= 45, $"Expected at least 45 links; found {links.Length}.");
        Assert.Equal(30, links.Distinct().Count());
        Assert.Contains("https://www.parl.ca/legisinfo/en/bill/44-1/c-244", links);
    }

    [Theory]
    [InlineData(false, false, DxpFieldEvalExportMode.Cache)]
    [InlineData(false, true, DxpFieldEvalExportMode.Cache)]
    [InlineData(true, false, DxpFieldEvalExportMode.Cache)]
    [InlineData(true, true, DxpFieldEvalExportMode.Cache)]
    [InlineData(false, false, DxpFieldEvalExportMode.Evaluate)]
    [InlineData(false, true, DxpFieldEvalExportMode.Evaluate)]
    public void PreservesFieldHyperlinks(bool nested, bool simple, DxpFieldEvalExportMode mode)
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var paragraph = new Paragraph();
            if (nested)
            {
                paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                    new Run(new FieldCode(" IF 1 = 1 \"yes\" \"no\" ")),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
            }
            if (simple)
                paragraph.Append(new SimpleField(new Run(new Text("Open link"))) {
                    Instruction = " HYPERLINK \"https://example.test/nested\" "
                });
            else
                paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                    new Run(new FieldCode(" HYPERLINK \"https://example.test/nested\" ")),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                    new Run(new Text("Open link")),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            if (nested)
                paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            main.Document = new Document(new Body(paragraph));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor(),
            new DxpExportOptions { FieldEvalMode = mode });
        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var links = pdf.Pages.Cast<PdfPage>()
            .SelectMany(page => page.Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>())
            .Select(annotation => annotation.Elements.GetDictionary("/A")?.Elements.GetString("/URI"))
            .Where(uri => !string.IsNullOrEmpty(uri)).ToArray();
        Assert.Contains("https://example.test/nested", links);
    }

    [Fact]
    public void CachedFieldsCanRefreshOnlyDateAndTime()
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        var dateSource = Path.Combine(root, "fixtures", "docx", "legacy", "WordDateTimeFields.docx");
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

        var tocSource = Path.Combine(root, "fixtures", "docx", "samples", "TOC.docx");
        Assert.Equal(Extract(tocSource, cachedOptions), Extract(tocSource, refreshOptions));
    }

    [Theory]
    [InlineData(DxpFieldEvalExportMode.Cache, false)]
    [InlineData(DxpFieldEvalExportMode.Cache, true)]
    [InlineData(DxpFieldEvalExportMode.Evaluate, false)]
    [InlineData(DxpFieldEvalExportMode.Evaluate, true)]
    public void IncludePictureFieldPreservesCachedImage(DxpFieldEvalExportMode mode, bool drawingResult)
    {
        GlobalFontSettings.FontResolver = PdfTextExportTests.FontResolver;
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../.."));
        using var fixture = System.IO.Compression.ZipFile.OpenRead(
            Path.Combine(root, "fixtures", "docx", "samples", "TestLists.docx"));
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var imagePart = main.AddImagePart(ImagePartType.Gif, "rIdImage");
            using (var imageStream = imagePart.GetStream(FileMode.Create, FileAccess.Write))
            using (var gifStream = fixture.GetEntry("word/media/image1.gif")!.Open())
                gifStream.CopyTo(imageStream);

            OpenXmlElement image;
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
                    "<pic:blipFill><a:blip r:embed=\"rIdImage\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
                    "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"609600\"/></a:xfrm>" +
                    "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>" +
                    "</a:graphicData></a:graphic></wp:inline>";
                image = drawing;
            }
            else
            {
                var picture = new Picture();
                picture.InnerXml =
                    "<v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                    "style=\"width:72pt;height:48pt\"><v:imagedata r:id=\"rIdImage\"/></v:shape>";
                image = picture;
            }

            main.Document = new Document(new Body(new Paragraph(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" INCLUDEPICTURE image1.gif ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(image),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }))));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor(),
            new DxpExportOptions { FieldEvalMode = mode });
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
        Assert.Single(pdf.GetPage(1).GetImages());
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

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void RendersDrawingTextBoxWithoutVmlFallback(bool inlineDrawing)
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
            var alternate = word.MainDocumentPart!.Document.Descendants<AlternateContent>().First();
            alternate.Elements<AlternateContentFallback>().First().Remove();
            var drawing = alternate.Descendants<Drawing>().First();
            var anchor = drawing.ChildElements.First(element => element.LocalName == "anchor");
            if (inlineDrawing)
            {
                var inline = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline();
                foreach (var child in anchor.ChildElements)
                    if (child.LocalName is "extent" or "effectExtent" or "docPr" or
                        "cNvGraphicFramePr" or "graphic")
                        inline.Append(child.CloneNode(true));
                drawing.ReplaceChild(inline, anchor);
            }
            else
            {
                var wrap = anchor.ChildElements.First(element => element.LocalName == "wrapNone");
                anchor.ReplaceChild(new DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapSquare
                    { WrapText = DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapTextValues.BothSides }, wrap);
            }
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
