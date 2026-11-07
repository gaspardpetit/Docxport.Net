using System.Text.RegularExpressions;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class RestartLevelImportTests
{
    // Installed Word drops this entire replacement on import, even its startOverride.
    [Theory]
    [InlineData(2, false)]
    [InlineData(2, true)]
    [InlineData(3, false)]
    [InlineData(3, true)]
    [InlineData(7, false)]
    [InlineData(7, true)]
    public void InvalidReplacementRestartDiscardsItsStartOverride(int restart, bool overrideStart)
    {
        using var source = new MemoryStream();
        using var doc = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var abs = new AbstractNum { AbstractNumberId = 1 };
        for (var i = 0; i < 3; i++)
            abs.Append(new Level(new StartNumberingValue { Val = new[] { 3, 2, 4 }[i] },
                new NumberingFormat { Val = NumberFormatValues.Decimal },
                new LevelText { Val = "L" + i + "=%" + (i + 1) + "." }) { LevelIndex = i });
        var replacement = new LevelOverride { LevelIndex = 2 };
        if (overrideStart) replacement.Append(new StartOverrideNumberingValue { Val = 7 });
        replacement.Append(new Level(new StartNumberingValue { Val = 4 },
            new NumberingFormat { Val = NumberFormatValues.Decimal },
            new LevelRestart { Val = restart }, new LevelText { Val = "V=%3." }) { LevelIndex = 2 });
        var numbering = new Numbering(abs,
            new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 },
            new NumberingInstance(new AbstractNumId { Val = 1 }, replacement) { NumberID = 2 });
        main.AddNewPart<NumberingDefinitionsPart>().Numbering = numbering;
        var body = new Body();
        var levels = new[] { 0, 1, 2, 2, 2 };
        for (var i = 0; i < levels.Length; i++)
            body.Append(new Paragraph(new ParagraphProperties(new NumberingProperties(
                new NumberingLevelReference { Val = levels[i] }, new NumberingId { Val = i == 4 ? 2 : 1 })),
                new Run(new Text("P" + i))));
        main.Document = new Document(body);
        var original = numbering.OuterXml;
        DxpITextVisitor[] visitors = [new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreatePlainConfig()),
            new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig())];
        foreach (var visitor in visitors)
        {
            var output = DxpExport.ExportToString(doc, visitor);
            if (visitor is DxpHtmlVisitor) output = new HtmlParser().ParseDocument(output).Body!.TextContent;
            var markers = string.Join(" ", Regex.Matches(output, @"(?:L[012]|V)=[0-9]+\.").Cast<Match>().Select(m => m.Value));
            Assert.Equal("L0=3. L1=2. L2=4. L2=5. L2=6.", markers);
            Assert.Equal(original, numbering.OuterXml);
        }
    }

    // Visible sequences captured from Word, including numeric restoration on a new instance.
    [Theory]
    [InlineData(false, false, false, "L0=3. L1=2. Section Section V=6.")]
    [InlineData(false, true, false, "L0=3. L1=2. V=6.")]
    [InlineData(false, false, true, "Section L0=4. L1=2. Section V=5.")]
    [InlineData(true, false, false, "L0=3. L1=2. V=2.")]
    public void HiddenLevelsConsumeCountersAndSeedAncestors(bool invalidRestart, bool empty, bool firstDeep,
        string expected)
    {
        using var source = new MemoryStream();
        using var doc = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var abs = new AbstractNum { AbstractNumberId = 1 };
        for (var i = 0; i < 3; i++)
            abs.Append(new Level(new StartNumberingValue { Val = new[] { 3, 2, 4 }[i] },
                new NumberingFormat { Val = NumberFormatValues.Decimal },
                new LevelText { Val = "L" + i + "=%" + (i + 1) + "." }) { LevelIndex = i });
        var child = abs.Elements<Level>().Last();
        if (invalidRestart) child.InsertBefore(new LevelRestart { Val = 2 }, child.LevelText!);
        else
        {
            child.NumberingFormat!.Val = NumberFormatValues.None;
            child.LevelText!.Val = empty ? "" : "Section";
        }
        var replacement = new LevelOverride(new Level(new StartNumberingValue { Val = 4 },
            new NumberingFormat { Val = NumberFormatValues.Decimal },
            new LevelText { Val = "V=%3." }) { LevelIndex = 2 }) { LevelIndex = 2 };
        var numbering = new Numbering(abs,
            new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 },
            new NumberingInstance(new AbstractNumId { Val = 1 }, replacement) { NumberID = 2 });
        main.AddNewPart<NumberingDefinitionsPart>().Numbering = numbering;
        var body = new Body();
        var levels = firstDeep ? new[] { 2, 0, 1, 2, 2 } : new[] { 0, 1, 2, 2, 2 };
        for (var i = 0; i < levels.Length; i++)
            body.Append(new Paragraph(new ParagraphProperties(new NumberingProperties(
                new NumberingLevelReference { Val = levels[i] }, new NumberingId { Val = i == 4 ? 2 : 1 })),
                new Run(new Text("P" + i))));
        main.Document = new Document(body);
        var original = numbering.OuterXml;
        DxpITextVisitor[] visitors = [new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreatePlainConfig()),
            new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig())];
        foreach (var visitor in visitors)
        {
            var output = DxpExport.ExportToString(doc, visitor);
            if (visitor is DxpHtmlVisitor) output = new HtmlParser().ParseDocument(output).Body!.TextContent;
            var markers = string.Join(" ", Regex.Matches(output, @"(?:L[012]|V)=[0-9]+\.|Section").Cast<Match>().Select(m => m.Value));
            Assert.Equal(expected, markers);
            Assert.Equal(original, numbering.OuterXml);
        }
    }
}
