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

public sealed class NegativeStartListMarkerTests
{
    // Full sequences from Word distinguish invalid level declarations from valid negative overrides.
    [Theory]
    [InlineData(-1, true, 9, null, "L0=3. L1=2. V=2.")]
    [InlineData(-2147483648, true, 9, null, "L0=3. L1=2. V=2.")]
    [InlineData(4, false, 9, -1, "L0=3. L1=2. L2=4. L2=5. L2=-1.")]
    [InlineData(4, false, 9, -2147483648, "L0=3. L1=2. L2=4. L2=5. L2=.")]
    [InlineData(4, true, -1, null, "L0=3. L1=2. L2=4. L2=5. L2=6.")]
    [InlineData(4, true, -1, 7, "L0=3. L1=2. L2=4. L2=5. L2=6.")]
    [InlineData(4, true, 9, -1, "L0=3. L1=2. L2=4. L2=5. V=9.")]
    [InlineData(4, true, 0, -1, "L0=3. L1=2. L2=4. L2=5. V=0.")]
    [InlineData(4, true, null, -1, "L0=3. L1=2. L2=4. L2=5. L2=6.")]
    public void ImportPreservesWordNegativeStartDistinctions(int abstractStart, bool hasReplacement,
        int? replacementStart, int? overrideStart, string expected)
    {
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var abs = new AbstractNum { AbstractNumberId = 1 };
        for (var i = 0; i < 3; i++)
            abs.Append(new Level(new StartNumberingValue { Val = i == 2 ? abstractStart : i == 0 ? 3 : 2 },
                new NumberingFormat { Val = NumberFormatValues.Decimal },
                new LevelText { Val = "L" + i + "=%" + (i + 1) + "." }) { LevelIndex = i });
        var ov = new LevelOverride { LevelIndex = 2 };
        if (overrideStart.HasValue) ov.Append(new StartOverrideNumberingValue { Val = overrideStart.Value });
        if (hasReplacement)
        {
            var level = new Level { LevelIndex = 2 };
            if (replacementStart.HasValue) level.Append(new StartNumberingValue { Val = replacementStart.Value });
            level.Append(new NumberingFormat { Val = NumberFormatValues.Decimal }, new LevelText { Val = "V=%3." });
            ov.Append(level);
        }
        var numbering = new Numbering(abs, new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 },
            new NumberingInstance(new AbstractNumId { Val = 1 }, ov) { NumberID = 2 });
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
            var markers = string.Join(" ", Regex.Matches(output, @"(?:L[012]|V)=[-0-9]*\.").Cast<Match>().Select(m => m.Value));
            Assert.Equal(expected, markers);
            Assert.Equal(original, numbering.OuterXml);
        }
    }
}
