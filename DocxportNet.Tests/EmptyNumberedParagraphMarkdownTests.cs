using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.PlainText;
using AngleSharp.Html.Parser;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxportNet.Tests;

public sealed class EmptyNumberedParagraphMarkdownTests
{
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void InlineChangesClosesRevisionTagsAfterTrailingEmptyNumberedMarker(bool inserted)
    {
        var path = Path.Combine(Path.GetTempPath(), $"numbered-inline-{Guid.NewGuid():N}.docx");
        try
        {
            using (var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                var numbering = main.AddNewPart<NumberingDefinitionsPart>();
                numbering.Numbering = new Numbering(
                    new AbstractNum(new Level(
                        new StartNumberingValue { Val = 1 },
                        new NumberingFormat { Val = NumberFormatValues.Decimal },
                        new LevelText { Val = "%1." }) { LevelIndex = 0 }) { AbstractNumberId = 1 },
                    new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
                numbering.Numbering.Save();
                var revision = inserted ? (OpenXmlElement)new Inserted { Id = "1", Author = "Test" }
                    : new Deleted { Id = "1", Author = "Test" };
                main.Document = new Document(new Body(new Paragraph(
                    new ParagraphProperties(
                        new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 }),
                        new ParagraphMarkRunProperties(revision)))));
                main.Document.Save();
            }

            var config = DxpMarkdownVisitorConfig.CreatePlainConfig();
            config.TrackedChangeMode = DxpTrackedChangeMode.InlineChanges;
            var markdown = DxpExport.ExportToString(path, new DxpMarkdownVisitor(config));
            Assert.Contains(inserted ? "<u>" : "<del>", markdown);
            Assert.Equal(Count(markdown, "<u>"), Count(markdown, "</u>"));
            Assert.Equal(Count(markdown, "<del>"), Count(markdown, "</del>"));
            Assert.Contains(@"1\.", markdown);
        }
        finally
        {
            File.Delete(path);
        }

        static int Count(string text, string pattern) => text.Split(pattern).Length - 1;
    }

    [Fact]
    public void SplitChangesKeepsTrailingEmptyNumberedMarker()
    {
        var path = Path.Combine(Path.GetTempPath(), $"numbered-split-{Guid.NewGuid():N}.docx");
        try
        {
            using (var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                var numbering = main.AddNewPart<NumberingDefinitionsPart>();
                numbering.Numbering = new Numbering(
                    new AbstractNum(new Level(
                        new StartNumberingValue { Val = 1 },
                        new NumberingFormat { Val = NumberFormatValues.Decimal },
                        new LevelText { Val = "%1." }) { LevelIndex = 0 }) { AbstractNumberId = 1 },
                    new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
                numbering.Numbering.Save();
                main.Document = new Document(new Body(new Paragraph(
                    new ParagraphProperties(new NumberingProperties(
                        new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 })))));
                main.Document.Save();
            }

            var config = DxpMarkdownVisitorConfig.CreatePlainConfig();
            config.TrackedChangeMode = DxpTrackedChangeMode.SplitChanges;
            var markdown = DxpExport.ExportToString(path, new DxpMarkdownVisitor(config));
            Assert.Contains(@"1\.", markdown);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void HtmlKeepsInlineDrawingInsideItsNumberedParagraph()
    {
        var path = Path.Combine(Path.GetTempPath(), $"numbered-drawing-{Guid.NewGuid():N}.docx");
        try
        {
            using (var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                var numbering = main.AddNewPart<NumberingDefinitionsPart>();
                numbering.Numbering = new Numbering(
                    new AbstractNum(new Level(
                        new StartNumberingValue { Val = 1 },
                        new NumberingFormat { Val = NumberFormatValues.Decimal },
                        new LevelText { Val = "%1." }) { LevelIndex = 0 }) { AbstractNumberId = 1 },
                    new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
                numbering.Numbering.Save();
                main.Document = new Document(new Body(new Paragraph(
                    new ParagraphProperties(new NumberingProperties(
                        new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 })),
                    new Run(new Drawing(new DW.Inline(new DW.DocProperties { Id = 1U, Name = "probe" }))))));
                main.Document.Save();
            }

            var config = DxpHtmlVisitorConfig.CreateRichConfig();
            config.EmitImages = false;
            var html = DxpExport.ExportToString(path, new DxpHtmlVisitor(config));
            var parsed = new HtmlParser().ParseDocument(html);
            var numbered = Assert.Single(parsed.QuerySelectorAll("p").Where(p =>
                p.QuerySelector(".dxp-marker")?.TextContent.Trim() == "1."));
            Assert.NotNull(numbered.QuerySelector(".dxp-image"));

            var markdown = DxpExport.ExportToString(path,
                new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()));
            Assert.Contains(@"1\. [IMAGE]", markdown);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void StandaloneOrderedMarkersAreEscapedAcrossListRestarts()
    {
        var path = Path.Combine(Path.GetTempPath(), $"numbered-restart-{Guid.NewGuid():N}.docx");
        try
        {
            using (var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                var numbering = main.AddNewPart<NumberingDefinitionsPart>();
                numbering.Numbering = new Numbering(
                    new AbstractNum(new Level(
                        new StartNumberingValue { Val = 1 },
                        new NumberingFormat { Val = NumberFormatValues.Decimal },
                        new LevelText { Val = "%1." }) { LevelIndex = 0 }) { AbstractNumberId = 1 },
                    new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 },
                    new NumberingInstance(new AbstractNumId { Val = 1 },
                        new LevelOverride(new StartOverrideNumberingValue { Val = 1 }) { LevelIndex = 0 })
                    { NumberID = 2 });
                numbering.Numbering.Save();
                Paragraph Item(int id, bool lineBreak) => new(
                    new ParagraphProperties(new NumberingProperties(
                        new NumberingLevelReference { Val = 0 }, new NumberingId { Val = id })),
                    lineBreak ? new Run(new Break()) : null);
                main.Document = new Document(new Body(Item(1, false), Item(2, true)));
                main.Document.Save();
            }

            var markdown = DxpExport.ExportToString(path,
                new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()));
            Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(markdown, @"(?m)^1\\\.(?:$| <br/>)").Count);
            Assert.DoesNotContain("1.\n", markdown);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void EmptyNumberedMarkerStaysInItsTableCell()
    {
        var path = Path.Combine(Path.GetTempPath(), $"numbered-cell-{Guid.NewGuid():N}.docx");
        try
        {
            using (var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                var numbering = main.AddNewPart<NumberingDefinitionsPart>();
                numbering.Numbering = new Numbering(
                    new AbstractNum(new Level(
                        new StartNumberingValue { Val = 1 },
                        new NumberingFormat { Val = NumberFormatValues.Decimal },
                        new LevelText { Val = "%1" }) { LevelIndex = 0 }) { AbstractNumberId = 1 },
                    new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
                numbering.Numbering.Save();
                var emptyNumbered = new Paragraph(new ParagraphProperties(new NumberingProperties(
                    new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 }),
                    new Shading { Fill = "FFFF00" },
                    new ParagraphMarkRunProperties(new FontSize { Val = "32" })));
                main.Document = new Document(new Body(new Table(new TableRow(
                    new TableCell(emptyNumbered),
                    new TableCell(new Paragraph(new Run(new Text("Next cell"))))))));
                main.Document.Save();
            }

            var markdown = DxpExport.ExportToString(path,
                new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()));
            Assert.Matches(@"1\s*</td>", markdown);
            Assert.DoesNotContain("1 Next cell", markdown);

            var html = DxpExport.ExportToString(path,
                new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig()));
            Assert.Matches(@">1</span>\s*</p>\s*</td>", html);
            Assert.DoesNotContain(">1</span> Next cell", html);
            var parsed = new HtmlParser().ParseDocument(html);
            var numbered = Assert.Single(parsed.QuerySelectorAll("td p").Where(p =>
                p.QuerySelector(".dxp-marker")?.TextContent.Trim() == "1"));
            Assert.Contains("font-size:16pt", numbered.GetAttribute("style"));
            Assert.Contains("background-color", numbered.GetAttribute("style"));

            var plainText = DxpExport.ExportToString(path,
                new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig()));
            Assert.Contains("1", plainText);
            Assert.Contains("Next cell", plainText);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void PlainMarkdownKeepsNumberWhenEmptyNumberedParagraphPrecedesItsText()
    {
        var path = Path.Combine(Path.GetTempPath(), $"numbered-continuation-{Guid.NewGuid():N}.docx");
        try
        {
            using (var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                var numbering = main.AddNewPart<NumberingDefinitionsPart>();
                numbering.Numbering = new Numbering(
                    new AbstractNum(new Level(
                        new StartNumberingValue { Val = 1 },
                        new NumberingFormat { Val = NumberFormatValues.Decimal },
                        new LevelText { Val = "%1" }) { LevelIndex = 0 }) { AbstractNumberId = 1 },
                    new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
                numbering.Numbering.Save();

                Paragraph Numbered(string? text, bool tabOnly = false, bool lineBreak = false)
                {
                    var p = new Paragraph(new ParagraphProperties(new NumberingProperties(
                        new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 })));
                    if (tabOnly) p.Append(new Run(new TabChar()));
                    if (lineBreak) p.Append(new Run(new Break()));
                    if (text != null) p.Append(new Run(new Text(text)));
                    return p;
                }

                main.Document = new Document(new Body(
                    Numbered("First paragraph"),
                    Numbered(null),
                    new Paragraph(new Run(new Text("Second paragraph"))),
                    Numbered(null, tabOnly: true),
                    new Paragraph(new Run(new Text("Third paragraph"))),
                    Numbered("Fourth paragraph"),
                    Numbered(null),
                    Numbered("Sixth paragraph"),
                    Numbered(null),
                    new Paragraph(),
                    new Paragraph(new Run(new Text("Separated paragraph"))),
                    Numbered(null, lineBreak: true),
                    new Paragraph(new Run(new Text("After line break"))),
                    Numbered("Text after break", lineBreak: true)));
                main.Document.Save();
            }

            var markdown = DxpExport.ExportToString(path,
                new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()));
            Assert.Contains("2\n", markdown);
            Assert.Contains("1 First paragraph", markdown);
            Assert.DoesNotContain("2 Second paragraph", markdown);
            Assert.Matches(@"(?m)^3(?:\s|$)", markdown);
            Assert.DoesNotContain("3 Third paragraph", markdown);
            Assert.Contains("4 Fourth paragraph", markdown);
            Assert.Contains("5\n", markdown);
            Assert.Contains("6 Sixth paragraph", markdown);
            Assert.Contains("7\n", markdown);
            Assert.DoesNotContain("7 Separated paragraph", markdown);
            Assert.Contains("8 <br/>", markdown);
            Assert.DoesNotContain("8 After line break", markdown);
            Assert.Contains("9", markdown);
            Assert.Contains("Text after break", markdown);

            var html = DxpExport.ExportToString(path,
                new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig()));
            Assert.Contains(">2</span> </p>", html);
            Assert.Contains(">3</span>", html);
            Assert.DoesNotContain(">2</span> Second paragraph", html);
            Assert.Contains(">5</span>", html);
            Assert.Contains(">7</span>", html);
            Assert.DoesNotContain(">7</span> Separated paragraph", html);
            var parsedHtml = new HtmlParser().ParseDocument(html);
            var breakParagraph = Assert.Single(parsedHtml.QuerySelectorAll("p").Where(p =>
                p.QuerySelector(".dxp-marker")?.TextContent.Trim() == "8"));
            Assert.NotNull(breakParagraph.QuerySelector("br"));
            Assert.DoesNotContain(">8</span> After line break", html);

            var plainText = DxpExport.ExportToString(path,
                new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig()));
            Assert.Contains("2\n", plainText);
            Assert.DoesNotContain("2 Second paragraph", plainText);
            Assert.Matches(@"(?m)^3(?:\s|$)", plainText);
            Assert.DoesNotContain("3 Third paragraph", plainText);
            Assert.Contains("5\n", plainText);
            Assert.Contains("7\n", plainText);
            Assert.DoesNotContain("7 Separated paragraph", plainText);
            Assert.Contains("8\n", plainText);
            Assert.DoesNotContain("8 After line break", plainText);
            Assert.Contains("9", plainText);
            Assert.Contains("Text after break", plainText);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
