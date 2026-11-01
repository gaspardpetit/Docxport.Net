using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Walker;

namespace DocxportNet.Tests;

public sealed class ParagraphNumberingInstanceTests
{
    [Theory]
    [InlineData(true, "10.2 Shared child")]
    [InlineData(false, "2.1 Shared child")]
    public void PlainMarkdownKeepsParentAndChildCountersAcrossInstances(bool sharedParentFirst, string expected)
    {
        var path = Path.Combine(Path.GetTempPath(), $"numbering-instance-child-{Guid.NewGuid():N}.docx");
        try
        {
            using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                var numbering = main.AddNewPart<NumberingDefinitionsPart>();
                numbering.Numbering = new Numbering(
                    new AbstractNum(
                        new Level(new StartNumberingValue { Val = 1 },
                            new NumberingFormat { Val = NumberFormatValues.Decimal },
                            new LevelText { Val = "%1." }) { LevelIndex = 0 },
                        new Level(new StartNumberingValue { Val = 1 },
                            new NumberingFormat { Val = NumberFormatValues.Decimal },
                            new LevelText { Val = "%1.%2" }) { LevelIndex = 1 })
                    { AbstractNumberId = 7 },
                    new NumberingInstance(new AbstractNumId { Val = 7 }) { NumberID = 1 },
                    new NumberingInstance(new AbstractNumId { Val = 7 },
                        new LevelOverride(new StartOverrideNumberingValue
                            { Val = sharedParentFirst ? 10 : 1 }) { LevelIndex = 0 })
                    { NumberID = 8 });
                numbering.Numbering.Save();

                Paragraph Item(int numId, int level, string text) => new(
                    new ParagraphProperties(new NumberingProperties(
                        new NumberingLevelReference { Val = level }, new NumberingId { Val = numId })),
                    new Run(new Text(text)));
                main.Document = new Document(new Body(sharedParentFirst
                    ? new[] { Item(1, 0, "Shared parent"), Item(8, 0, "Restarted parent"),
                        Item(8, 1, "Restarted child"), Item(1, 1, "Shared child") }
                    : new[] { Item(8, 0, "Restarted parent"), Item(8, 1, "Restarted child"),
                        Item(1, 0, "Shared parent"), Item(1, 1, "Shared child") }));
                main.Document.Save();
            }

            var markdown = DxpExport.ExportToString(path,
                new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()));
            Assert.Contains(expected, markdown);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ChildUsesMostRecentParentAcrossInstances()
    {
        var tracker = new DxpListTracker();

        Assert.Equal(1, tracker.NextIndex(1, 7, 0, 1, false));
        Assert.Equal(10, tracker.NextIndex(8, 7, 0, 10, true));
        Assert.Equal(1, tracker.NextIndex(8, 7, 1, 1, false));
        Assert.Equal(2, tracker.NextIndex(1, 7, 1, 1, false));
        Assert.Equal(10, tracker.GetActiveCurrent(7, 0));
    }

    [Fact]
    public void ParentAdvanceClearsPendingChildFromAnotherInstance()
    {
        var tracker = new DxpListTracker();

        Assert.Equal(1, tracker.NextIndex(8, 7, 0, 1, true));
        Assert.Equal(1, tracker.NextIndex(8, 7, 1, 1, false));
        Assert.Equal(2, tracker.NextIndex(1, 7, 0, 1, false));
        Assert.Equal(1, tracker.NextIndex(1, 7, 1, 1, false));
    }

    [Fact]
    public void PlainMarkdownHonorsLevelRestartRules()
    {
        var path = Path.Combine(Path.GetTempPath(), $"numbering-level-restart-{Guid.NewGuid():N}.docx");
        try
        {
            using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                var numbering = main.AddNewPart<NumberingDefinitionsPart>();
                numbering.Numbering = new Numbering(
                    new AbstractNum(
                        new Level(new StartNumberingValue { Val = 1 },
                            new NumberingFormat { Val = NumberFormatValues.Decimal },
                            new LevelText { Val = "[%1]" }) { LevelIndex = 0 },
                        new Level(new StartNumberingValue { Val = 1 },
                            new NumberingFormat { Val = NumberFormatValues.Decimal },
                            new LevelText { Val = "(%2)" }, new LevelRestart { Val = 0 }) { LevelIndex = 1 },
                        new Level(new StartNumberingValue { Val = 1 },
                            new NumberingFormat { Val = NumberFormatValues.Decimal },
                            new LevelText { Val = "{%3}" }, new LevelRestart { Val = 1 }) { LevelIndex = 2 })
                    { AbstractNumberId = 7 },
                    new NumberingInstance(new AbstractNumId { Val = 7 }) { NumberID = 1 });
                numbering.Numbering.Save();

                Paragraph Item(int level, string text) => new(
                    new ParagraphProperties(new NumberingProperties(
                        new NumberingLevelReference { Val = level }, new NumberingId { Val = 1 })),
                    new Run(new Text(text)));
                main.Document = new Document(new Body(
                    Item(0, "Parent one"), Item(1, "Child one"),
                    Item(2, "Grandchild one"), Item(2, "Grandchild two"),
                    Item(1, "Child two"), Item(2, "Grandchild three"),
                    Item(0, "Parent two"), Item(1, "Child three"),
                    Item(2, "Grandchild after parent")));
                main.Document.Save();
            }

            var markdown = DxpExport.ExportToString(path,
                new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()));
            Assert.Contains("(3) Child three", markdown);
            Assert.Contains("{3} Grandchild three", markdown);
            Assert.Contains("{1} Grandchild after parent", markdown);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void TwoRestartsContinueFromTheMostRecentInstance()
    {
        var tracker = new DxpListTracker();

        Assert.Equal(1, tracker.NextIndex(2, 7, 0, 1, true));
        Assert.Equal(2, tracker.NextIndex(2, 7, 0, 1, true));
        Assert.Equal(10, tracker.NextIndex(4, 7, 0, 10, true));
        Assert.Equal(11, tracker.NextIndex(4, 7, 0, 10, true));
        Assert.Equal(12, tracker.NextIndex(3, 7, 0, 1, false));
    }

    [Fact]
    public void NewInstanceContinuesRestartedChildAndParentLevels()
    {
        var tracker = new DxpListTracker();

        Assert.Equal(1, tracker.NextIndex(1, 7, 0, 1, false));
        Assert.Equal(1, tracker.NextIndex(1, 7, 1, 1, false));
        Assert.Equal(2, tracker.NextIndex(1, 7, 1, 1, false));
        Assert.Equal(1, tracker.NextIndex(2, 7, 0, 1, true));
        Assert.Equal(1, tracker.NextIndex(2, 7, 1, 1, false));
        Assert.Equal(2, tracker.NextIndex(3, 7, 1, 1, false));
        Assert.Equal(2, tracker.NextIndex(3, 7, 0, 1, false));
    }

    [Fact]
    public void NewSharedInstanceContinuesTheMostRecentRestart()
    {
        var tracker = new DxpListTracker();

        Assert.Equal(1, tracker.NextIndex(7, 5, 0, 1, false));
        Assert.Equal(2, tracker.NextIndex(7, 5, 0, 1, false));
        Assert.Equal(1, tracker.NextIndex(15, 5, 0, 1, true));
        Assert.Equal(2, tracker.NextIndex(8, 5, 0, 1, false));
        Assert.Equal(3, tracker.NextIndex(8, 5, 0, 1, false));
    }

    [Fact]
    public void PlainMarkdownRestartsHeadingsThenContinuesAcrossInstances()
    {
        var path = Path.Combine(Path.GetTempPath(), $"numbering-headings-{Guid.NewGuid():N}.docx");
        try
        {
            using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                var numbering = main.AddNewPart<NumberingDefinitionsPart>();
                numbering.Numbering = new Numbering(
                    new AbstractNum(new Level(
                        new StartNumberingValue { Val = 1 },
                        new NumberingFormat { Val = NumberFormatValues.UpperRoman },
                        new LevelText { Val = "%1." }) { LevelIndex = 0 }) { AbstractNumberId = 5 },
                    new NumberingInstance(new AbstractNumId { Val = 5 }) { NumberID = 7 },
                    new NumberingInstance(new AbstractNumId { Val = 5 }) { NumberID = 8 },
                    new NumberingInstance(new AbstractNumId { Val = 5 },
                        new LevelOverride(new StartOverrideNumberingValue { Val = 1 }) { LevelIndex = 0 })
                    { NumberID = 15 });
                numbering.Numbering.Save();

                var styles = main.AddNewPart<StyleDefinitionsPart>();
                styles.Styles = new Styles(new Style(
                    new StyleParagraphProperties(new NumberingProperties(new NumberingId { Val = 7 })))
                { Type = StyleValues.Paragraph, StyleId = "Heading" });
                styles.Styles.Save();

                Paragraph Heading(string text, int? numId = null) => new(
                    new ParagraphProperties(new ParagraphStyleId { Val = "Heading" },
                        numId is int id ? new NumberingProperties(new NumberingLevelReference { Val = 0 },
                            new NumberingId { Val = id }) : null),
                    new Run(new Text(text)));
                main.Document = new Document(new Body(
                    Heading("First section"), Heading("First facts"),
                    Heading("Second introduction", 15), Heading("Second facts", 8),
                    Heading("Second analysis", 8)));
                main.Document.Save();
            }

            var markdown = DxpExport.ExportToString(path,
                new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()));
            Assert.Contains("II. First facts", markdown);
            Assert.Contains("I. Second introduction", markdown);
            Assert.Contains("II. Second facts", markdown);
            Assert.Contains("III. Second analysis", markdown);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void FirstRestartedInstanceContinuesIntoStyleInstance()
    {
        var tracker = new DxpListTracker();

        Assert.Equal(1, tracker.NextIndex(2, 7, 0, 1, true));
        Assert.Equal(2, tracker.NextIndex(2, 7, 0, 1, true));
        Assert.Equal(3, tracker.NextIndex(2, 7, 0, 1, true));
        Assert.Equal(4, tracker.NextIndex(1, 7, 0, 1, false));
        Assert.Equal(5, tracker.NextIndex(1, 7, 0, 1, false));
    }

    [Fact]
    public void PlainMarkdownContinuesFromRestartedInstanceIntoStyleInstance()
    {
        var path = Path.Combine(Path.GetTempPath(), $"numbering-continuation-{Guid.NewGuid():N}.docx");
        try
        {
            using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                var numbering = main.AddNewPart<NumberingDefinitionsPart>();
                numbering.Numbering = new Numbering(
                    new AbstractNum(new Level(
                        new StartNumberingValue { Val = 1 },
                        new NumberingFormat { Val = NumberFormatValues.Decimal },
                        new LevelText { Val = "[%1]" }) { LevelIndex = 0 }) { AbstractNumberId = 7 },
                    new NumberingInstance(new AbstractNumId { Val = 7 }) { NumberID = 1 },
                    new NumberingInstance(new AbstractNumId { Val = 7 },
                        new LevelOverride(new StartOverrideNumberingValue { Val = 1 }) { LevelIndex = 0 })
                    { NumberID = 2 });
                numbering.Numbering.Save();

                var styles = main.AddNewPart<StyleDefinitionsPart>();
                styles.Styles = new Styles(new Style(
                    new StyleParagraphProperties(new NumberingProperties(new NumberingId { Val = 1 })))
                { Type = StyleValues.Paragraph, StyleId = "NumberedBody" });
                styles.Styles.Save();

                Paragraph Explicit(string text) => new(
                    new ParagraphProperties(new ParagraphStyleId { Val = "NumberedBody" },
                        new NumberingProperties(new NumberingLevelReference { Val = 0 },
                            new NumberingId { Val = 2 })),
                    new Run(new Text(text)));
                Paragraph Inherited(string text) => new(
                    new ParagraphProperties(new ParagraphStyleId { Val = "NumberedBody" }),
                    new Run(new Text(text)));
                main.Document = new Document(new Body(
                    Explicit("Majority one"), Explicit("Majority two"), Explicit("Majority three"),
                    Inherited("Dissent one"), Inherited("Dissent two")));
                main.Document.Save();
            }

            var markdown = DxpExport.ExportToString(path,
                new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()));
            Assert.Contains("[3] Majority three", markdown);
            Assert.Contains("[4] Dissent one", markdown);
            Assert.Contains("[5] Dissent two", markdown);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ExplicitRestartDoesNotOverwriteAnotherInstancesProgress()
    {
        var tracker = new DxpListTracker();

        Assert.Equal(1, tracker.NextIndex(1, 7, 0, 1, false));
        Assert.Equal(2, tracker.NextIndex(1, 7, 0, 1, false));
        Assert.Equal(10, tracker.NextIndex(8, 7, 0, 10, true));
        Assert.Equal(11, tracker.NextIndex(8, 7, 0, 10, true));
        Assert.Equal(3, tracker.NextIndex(1, 7, 0, 1, false));
        Assert.Equal(3, tracker.GetCurrent(1, 7, 0));
        Assert.Equal(11, tracker.GetCurrent(8, 7, 0));
    }

    [Fact]
    public void ExplicitRestartDoesNotResetAnotherInstancesDescendants()
    {
        var tracker = new DxpListTracker();

        Assert.Equal(1, tracker.NextIndex(1, 7, 0, 1, false));
        Assert.Equal(1, tracker.NextIndex(1, 7, 1, 1, false));
        Assert.Equal(2, tracker.NextIndex(1, 7, 1, 1, false));
        Assert.Equal(10, tracker.NextIndex(8, 7, 0, 10, true));
        Assert.Equal(1, tracker.NextIndex(8, 7, 1, 1, false));
        Assert.Equal(3, tracker.NextIndex(1, 7, 1, 1, false));
        Assert.Equal(2, tracker.NextIndex(8, 7, 1, 1, false));
        Assert.Equal(11, tracker.NextIndex(8, 7, 0, 10, true));
        Assert.Equal(1, tracker.NextIndex(8, 7, 1, 1, false));
        Assert.Equal(4, tracker.NextIndex(1, 7, 1, 1, false));
        Assert.Equal(11, tracker.GetCurrent(8, 7, 0));
        Assert.Equal(4, tracker.GetCurrent(1, 7, 1));
    }

    [Fact]
    public void PlainMarkdownContinuesParagraphNumberAfterHeadingUsesSameInstance()
    {
        var path = Path.Combine(Path.GetTempPath(), $"numbering-{Guid.NewGuid():N}.docx");
        try
        {
            using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                var numbering = main.AddNewPart<NumberingDefinitionsPart>();
                numbering.Numbering = new Numbering(
                    new AbstractNum(
                        new Level(
                            new StartNumberingValue { Val = 88 },
                            new NumberingFormat { Val = NumberFormatValues.Decimal },
                            new LevelText { Val = "[%1]" }) { LevelIndex = 0 },
                        new Level(
                            new StartNumberingValue { Val = 1 },
                            new NumberingFormat { Val = NumberFormatValues.Decimal },
                            new LevelText { Val = "(%4)" }) { LevelIndex = 3 })
                    { AbstractNumberId = 7 },
                    new NumberingInstance(new AbstractNumId { Val = 7 }) { NumberID = 1 },
                    new NumberingInstance(
                        new AbstractNumId { Val = 7 },
                        new LevelOverride(new StartOverrideNumberingValue { Val = 1 }) { LevelIndex = 0 },
                        new LevelOverride(new StartOverrideNumberingValue { Val = 2 }) { LevelIndex = 3 })
                    { NumberID = 8 });
                numbering.Numbering.Save();

                var styles = main.AddNewPart<StyleDefinitionsPart>();
                styles.Styles = new Styles(new Style(
                    new StyleParagraphProperties(new NumberingProperties(new NumberingId { Val = 1 })))
                { Type = StyleValues.Paragraph, StyleId = "NumberedBody" });
                styles.Styles.Save();

                main.Document = new Document(new Body(
                    new Paragraph(
                        new ParagraphProperties(
                            new NumberingProperties(new NumberingLevelReference { Val = 3 },
                                new NumberingId { Val = 8 })),
                        new Run(new Text("Analysis"))),
                    new Paragraph(
                        new ParagraphProperties(new ParagraphStyleId { Val = "NumberedBody" }),
                        new Run(new Text("First paragraph"))),
                    new Paragraph(
                        new ParagraphProperties(
                            new ParagraphStyleId { Val = "NumberedBody" },
                            new NumberingProperties(new NumberingLevelReference { Val = 0 },
                                new NumberingId { Val = 8 })),
                        new Run(new Text("Second paragraph")))));
                main.Document.Save();
            }

            var markdown = DxpExport.ExportToString(path,
                new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()));
            Assert.Contains("[88] First paragraph", markdown);
            Assert.Contains("[89] Second paragraph", markdown);
            Assert.DoesNotContain("[1] Second paragraph", markdown);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
