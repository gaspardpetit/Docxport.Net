using System.Text.RegularExpressions;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Walker;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class RestartBoundaryMatrixTests
{
    // Full sequences captured from Word PDF, covering all levels and adjacent restart boundaries.
    [Theory]
    [InlineData(0, -1, false, "L0=1. L0=2. L0=3. L0=4. L0=5. L0=6.")]
    [InlineData(0, -1, true, "L0=1. L0=2. L0=3. L0=4. V=9. V=10.")]
    [InlineData(0, 0, false, "L0=1. L0=2. L0=3. L0=4. L0=5. L0=6.")]
    [InlineData(0, 0, true, "L0=1. L0=2. L0=3. L0=4. V=9. V=10.")]
    [InlineData(0, 1, false, "L0=1. L0=2. L0=3. L0=4. L0=5. L0=6.")]
    [InlineData(0, 1, true, "L0=1. L0=2. L0=3. L0=4. V=9. V=10.")]
    [InlineData(0, 2, false, "L0=1. L0=2. L0=3. L0=4. L0=5. L0=6.")]
    [InlineData(0, 2, true, "L0=1. L0=2. L0=3. L0=4. V=9. V=10.")]
    [InlineData(0, 7, false, "L0=1. L0=2. L0=3. L0=4. L0=5. L0=6.")]
    [InlineData(0, 7, true, "L0=1. L0=2. L0=3. L0=4. V=9. V=10.")]
    [InlineData(0, 8, false, "L0=1. L0=2. L0=3. L0=4. L0=5. L0=6.")]
    [InlineData(0, 8, true, "L0=1. L0=2. L0=3. L0=4. V=9. V=10.")]
    [InlineData(0, 9, false, "L0=1. L0=2. L0=3. L0=4. L0=5. L0=6.")]
    [InlineData(0, 9, true, "L0=1. L0=2. L0=3. L0=4. V=9. V=10.")]
    [InlineData(0, 2147483647, false, "L0=1. L0=2. L0=3. L0=4. L0=5. L0=6.")]
    [InlineData(0, 2147483647, true, "L0=1. L0=2. L0=3. L0=4. V=9. V=10.")]
    [InlineData(1, -1, false, "L0=1. L1=2. L1=3. L0=2. L1=4. L1=5.")]
    [InlineData(1, -1, true, "L0=1. L1=2. L1=3. L0=2. V=9. V=10.")]
    [InlineData(1, 0, false, "L0=1. L1=2. L1=3. L0=2. L1=4. L1=5.")]
    [InlineData(1, 0, true, "L0=1. L1=2. L1=3. L0=2. V=9. V=10.")]
    [InlineData(1, 1, false, "L0=1. L0=2.")]
    [InlineData(1, 1, true, "L0=1. L1=2. L1=3. L0=2. L1=2. L1=3.")]
    [InlineData(1, 2, false, "L0=1. L0=2.")]
    [InlineData(1, 2, true, "L0=1. L1=2. L1=3. L0=2. L1=2. L1=3.")]
    [InlineData(1, 7, false, "L0=1. L0=2.")]
    [InlineData(1, 7, true, "L0=1. L1=2. L1=3. L0=2. L1=2. L1=3.")]
    [InlineData(1, 8, false, "L0=1. L1=2. L1=3. L0=2. L1=4. L1=5.")]
    [InlineData(1, 8, true, "L0=1. L1=2. L1=3. L0=2. V=9. V=10.")]
    [InlineData(1, 9, false, "L0=1. L1=2. L1=3. L0=2. L1=4. L1=5.")]
    [InlineData(1, 9, true, "L0=1. L1=2. L1=3. L0=2. V=9. V=10.")]
    [InlineData(1, 2147483647, false, "L0=1. L1=2. L1=3. L0=2. L1=4. L1=5.")]
    [InlineData(1, 2147483647, true, "L0=1. L1=2. L1=3. L0=2. V=9. V=10.")]
    [InlineData(2, -1, false, "L0=1. L2=3. L2=4. L0=2. L2=5. L2=6.")]
    [InlineData(2, -1, true, "L0=1. L2=3. L2=4. L0=2. V=9. V=10.")]
    [InlineData(2, 0, false, "L0=1. L2=3. L2=4. L0=2. L2=5. L2=6.")]
    [InlineData(2, 0, true, "L0=1. L2=3. L2=4. L0=2. V=9. V=10.")]
    [InlineData(2, 1, false, "L0=1. L2=3. L2=4. L0=2. L2=3. L2=4.")]
    [InlineData(2, 1, true, "L0=1. L2=3. L2=4. L0=2. V=9. V=10.")]
    [InlineData(2, 2, false, "L0=1. L0=2.")]
    [InlineData(2, 2, true, "L0=1. L2=3. L2=4. L0=2. L2=3. L2=4.")]
    [InlineData(2, 7, false, "L0=1. L0=2.")]
    [InlineData(2, 7, true, "L0=1. L2=3. L2=4. L0=2. L2=3. L2=4.")]
    [InlineData(2, 8, false, "L0=1. L2=3. L2=4. L0=2. L2=5. L2=6.")]
    [InlineData(2, 8, true, "L0=1. L2=3. L2=4. L0=2. V=9. V=10.")]
    [InlineData(2, 9, false, "L0=1. L2=3. L2=4. L0=2. L2=5. L2=6.")]
    [InlineData(2, 9, true, "L0=1. L2=3. L2=4. L0=2. V=9. V=10.")]
    [InlineData(2, 2147483647, false, "L0=1. L2=3. L2=4. L0=2. L2=5. L2=6.")]
    [InlineData(2, 2147483647, true, "L0=1. L2=3. L2=4. L0=2. V=9. V=10.")]
    [InlineData(3, -1, false, "L0=1. L3=4. L3=5. L0=2. L3=6. L3=7.")]
    [InlineData(3, -1, true, "L0=1. L3=4. L3=5. L0=2. V=9. V=10.")]
    [InlineData(3, 0, false, "L0=1. L3=4. L3=5. L0=2. L3=6. L3=7.")]
    [InlineData(3, 0, true, "L0=1. L3=4. L3=5. L0=2. V=9. V=10.")]
    [InlineData(3, 1, false, "L0=1. L3=4. L3=5. L0=2. L3=4. L3=5.")]
    [InlineData(3, 1, true, "L0=1. L3=4. L3=5. L0=2. V=9. V=10.")]
    [InlineData(3, 2, false, "L0=1. L3=4. L3=5. L0=2. L3=4. L3=5.")]
    [InlineData(3, 2, true, "L0=1. L3=4. L3=5. L0=2. V=9. V=10.")]
    [InlineData(3, 7, false, "L0=1. L0=2.")]
    [InlineData(3, 7, true, "L0=1. L3=4. L3=5. L0=2. L3=4. L3=5.")]
    [InlineData(3, 8, false, "L0=1. L3=4. L3=5. L0=2. L3=6. L3=7.")]
    [InlineData(3, 8, true, "L0=1. L3=4. L3=5. L0=2. V=9. V=10.")]
    [InlineData(3, 9, false, "L0=1. L3=4. L3=5. L0=2. L3=6. L3=7.")]
    [InlineData(3, 9, true, "L0=1. L3=4. L3=5. L0=2. V=9. V=10.")]
    [InlineData(3, 2147483647, false, "L0=1. L3=4. L3=5. L0=2. L3=6. L3=7.")]
    [InlineData(3, 2147483647, true, "L0=1. L3=4. L3=5. L0=2. V=9. V=10.")]
    [InlineData(4, -1, false, "L0=1. L4=5. L4=6. L0=2. L4=7. L4=8.")]
    [InlineData(4, -1, true, "L0=1. L4=5. L4=6. L0=2. V=9. V=10.")]
    [InlineData(4, 0, false, "L0=1. L4=5. L4=6. L0=2. L4=7. L4=8.")]
    [InlineData(4, 0, true, "L0=1. L4=5. L4=6. L0=2. V=9. V=10.")]
    [InlineData(4, 1, false, "L0=1. L4=5. L4=6. L0=2. L4=5. L4=6.")]
    [InlineData(4, 1, true, "L0=1. L4=5. L4=6. L0=2. V=9. V=10.")]
    [InlineData(4, 2, false, "L0=1. L4=5. L4=6. L0=2. L4=5. L4=6.")]
    [InlineData(4, 2, true, "L0=1. L4=5. L4=6. L0=2. V=9. V=10.")]
    [InlineData(4, 7, false, "L0=1. L0=2.")]
    [InlineData(4, 7, true, "L0=1. L4=5. L4=6. L0=2. L4=5. L4=6.")]
    [InlineData(4, 8, false, "L0=1. L4=5. L4=6. L0=2. L4=7. L4=8.")]
    [InlineData(4, 8, true, "L0=1. L4=5. L4=6. L0=2. V=9. V=10.")]
    [InlineData(4, 9, false, "L0=1. L4=5. L4=6. L0=2. L4=7. L4=8.")]
    [InlineData(4, 9, true, "L0=1. L4=5. L4=6. L0=2. V=9. V=10.")]
    [InlineData(4, 2147483647, false, "L0=1. L4=5. L4=6. L0=2. L4=7. L4=8.")]
    [InlineData(4, 2147483647, true, "L0=1. L4=5. L4=6. L0=2. V=9. V=10.")]
    [InlineData(5, -1, false, "L0=1. L5=6. L5=7. L0=2. L5=8. L5=9.")]
    [InlineData(5, -1, true, "L0=1. L5=6. L5=7. L0=2. V=9. V=10.")]
    [InlineData(5, 0, false, "L0=1. L5=6. L5=7. L0=2. L5=8. L5=9.")]
    [InlineData(5, 0, true, "L0=1. L5=6. L5=7. L0=2. V=9. V=10.")]
    [InlineData(5, 1, false, "L0=1. L5=6. L5=7. L0=2. L5=6. L5=7.")]
    [InlineData(5, 1, true, "L0=1. L5=6. L5=7. L0=2. V=9. V=10.")]
    [InlineData(5, 2, false, "L0=1. L5=6. L5=7. L0=2. L5=6. L5=7.")]
    [InlineData(5, 2, true, "L0=1. L5=6. L5=7. L0=2. V=9. V=10.")]
    [InlineData(5, 7, false, "L0=1. L0=2.")]
    [InlineData(5, 7, true, "L0=1. L5=6. L5=7. L0=2. L5=6. L5=7.")]
    [InlineData(5, 8, false, "L0=1. L5=6. L5=7. L0=2. L5=8. L5=9.")]
    [InlineData(5, 8, true, "L0=1. L5=6. L5=7. L0=2. V=9. V=10.")]
    [InlineData(5, 9, false, "L0=1. L5=6. L5=7. L0=2. L5=8. L5=9.")]
    [InlineData(5, 9, true, "L0=1. L5=6. L5=7. L0=2. V=9. V=10.")]
    [InlineData(5, 2147483647, false, "L0=1. L5=6. L5=7. L0=2. L5=8. L5=9.")]
    [InlineData(5, 2147483647, true, "L0=1. L5=6. L5=7. L0=2. V=9. V=10.")]
    [InlineData(6, -1, false, "L0=1. L6=7. L6=8. L0=2. L6=9. L6=10.")]
    [InlineData(6, -1, true, "L0=1. L6=7. L6=8. L0=2. V=9. V=10.")]
    [InlineData(6, 0, false, "L0=1. L6=7. L6=8. L0=2. L6=9. L6=10.")]
    [InlineData(6, 0, true, "L0=1. L6=7. L6=8. L0=2. V=9. V=10.")]
    [InlineData(6, 1, false, "L0=1. L6=7. L6=8. L0=2. L6=7. L6=8.")]
    [InlineData(6, 1, true, "L0=1. L6=7. L6=8. L0=2. V=9. V=10.")]
    [InlineData(6, 2, false, "L0=1. L6=7. L6=8. L0=2. L6=7. L6=8.")]
    [InlineData(6, 2, true, "L0=1. L6=7. L6=8. L0=2. V=9. V=10.")]
    [InlineData(6, 7, false, "L0=1. L0=2.")]
    [InlineData(6, 7, true, "L0=1. L6=7. L6=8. L0=2. L6=7. L6=8.")]
    [InlineData(6, 8, false, "L0=1. L6=7. L6=8. L0=2. L6=9. L6=10.")]
    [InlineData(6, 8, true, "L0=1. L6=7. L6=8. L0=2. V=9. V=10.")]
    [InlineData(6, 9, false, "L0=1. L6=7. L6=8. L0=2. L6=9. L6=10.")]
    [InlineData(6, 9, true, "L0=1. L6=7. L6=8. L0=2. V=9. V=10.")]
    [InlineData(6, 2147483647, false, "L0=1. L6=7. L6=8. L0=2. L6=9. L6=10.")]
    [InlineData(6, 2147483647, true, "L0=1. L6=7. L6=8. L0=2. V=9. V=10.")]
    [InlineData(7, -1, false, "L0=1. L7=8. L7=9. L0=2. L7=10. L7=11.")]
    [InlineData(7, -1, true, "L0=1. L7=8. L7=9. L0=2. V=9. V=10.")]
    [InlineData(7, 0, false, "L0=1. L7=8. L7=9. L0=2. L7=10. L7=11.")]
    [InlineData(7, 0, true, "L0=1. L7=8. L7=9. L0=2. V=9. V=10.")]
    [InlineData(7, 1, false, "L0=1. L7=8. L7=9. L0=2. L7=8. L7=9.")]
    [InlineData(7, 1, true, "L0=1. L7=8. L7=9. L0=2. V=9. V=10.")]
    [InlineData(7, 2, false, "L0=1. L7=8. L7=9. L0=2. L7=8. L7=9.")]
    [InlineData(7, 2, true, "L0=1. L7=8. L7=9. L0=2. V=9. V=10.")]
    [InlineData(7, 7, false, "L0=1. L0=2.")]
    [InlineData(7, 7, true, "L0=1. L7=8. L7=9. L0=2. L7=8. L7=9.")]
    [InlineData(7, 8, false, "L0=1. L7=8. L7=9. L0=2. L7=10. L7=11.")]
    [InlineData(7, 8, true, "L0=1. L7=8. L7=9. L0=2. V=9. V=10.")]
    [InlineData(7, 9, false, "L0=1. L7=8. L7=9. L0=2. L7=10. L7=11.")]
    [InlineData(7, 9, true, "L0=1. L7=8. L7=9. L0=2. V=9. V=10.")]
    [InlineData(7, 2147483647, false, "L0=1. L7=8. L7=9. L0=2. L7=10. L7=11.")]
    [InlineData(7, 2147483647, true, "L0=1. L7=8. L7=9. L0=2. V=9. V=10.")]
    [InlineData(8, -1, false, "L0=1. L8=9. L8=10. L0=2. L8=11. L8=12.")]
    [InlineData(8, -1, true, "L0=1. L8=9. L8=10. L0=2. V=9. V=10.")]
    [InlineData(8, 0, false, "L0=1. L8=9. L8=10. L0=2. L8=11. L8=12.")]
    [InlineData(8, 0, true, "L0=1. L8=9. L8=10. L0=2. V=9. V=10.")]
    [InlineData(8, 1, false, "L0=1. L8=9. L8=10. L0=2. L8=9. L8=10.")]
    [InlineData(8, 1, true, "L0=1. L8=9. L8=10. L0=2. V=9. V=10.")]
    [InlineData(8, 2, false, "L0=1. L8=9. L8=10. L0=2. L8=9. L8=10.")]
    [InlineData(8, 2, true, "L0=1. L8=9. L8=10. L0=2. V=9. V=10.")]
    [InlineData(8, 7, false, "L0=1. L8=9. L8=10. L0=2. L8=9. L8=10.")]
    [InlineData(8, 7, true, "L0=1. L8=9. L8=10. L0=2. V=9. V=10.")]
    [InlineData(8, 8, false, "L0=1. L8=9. L8=10. L0=2. L8=11. L8=12.")]
    [InlineData(8, 8, true, "L0=1. L8=9. L8=10. L0=2. V=9. V=10.")]
    [InlineData(8, 9, false, "L0=1. L8=9. L8=10. L0=2. L8=11. L8=12.")]
    [InlineData(8, 9, true, "L0=1. L8=9. L8=10. L0=2. V=9. V=10.")]
    [InlineData(8, 2147483647, false, "L0=1. L8=9. L8=10. L0=2. L8=11. L8=12.")]
    [InlineData(8, 2147483647, true, "L0=1. L8=9. L8=10. L0=2. V=9. V=10.")]
    [InlineData(2, 3, false, "L0=1. L0=2.")]
    [InlineData(2, 3, true, "L0=1. L2=3. L2=4. L0=2. L2=3. L2=4.")]
    [InlineData(3, 3, false, "L0=1. L0=2.")]
    [InlineData(3, 3, true, "L0=1. L3=4. L3=5. L0=2. L3=4. L3=5.")]
    [InlineData(3, 4, false, "L0=1. L0=2.")]
    [InlineData(3, 4, true, "L0=1. L3=4. L3=5. L0=2. L3=4. L3=5.")]
    [InlineData(4, 3, false, "L0=1. L4=5. L4=6. L0=2. L4=5. L4=6.")]
    [InlineData(4, 3, true, "L0=1. L4=5. L4=6. L0=2. V=9. V=10.")]
    [InlineData(4, 4, false, "L0=1. L0=2.")]
    [InlineData(4, 4, true, "L0=1. L4=5. L4=6. L0=2. L4=5. L4=6.")]
    [InlineData(4, 5, false, "L0=1. L0=2.")]
    [InlineData(4, 5, true, "L0=1. L4=5. L4=6. L0=2. L4=5. L4=6.")]
    [InlineData(5, 4, false, "L0=1. L5=6. L5=7. L0=2. L5=6. L5=7.")]
    [InlineData(5, 4, true, "L0=1. L5=6. L5=7. L0=2. V=9. V=10.")]
    [InlineData(5, 5, false, "L0=1. L0=2.")]
    [InlineData(5, 5, true, "L0=1. L5=6. L5=7. L0=2. L5=6. L5=7.")]
    [InlineData(5, 6, false, "L0=1. L0=2.")]
    [InlineData(5, 6, true, "L0=1. L5=6. L5=7. L0=2. L5=6. L5=7.")]
    [InlineData(6, 5, false, "L0=1. L6=7. L6=8. L0=2. L6=7. L6=8.")]
    [InlineData(6, 5, true, "L0=1. L6=7. L6=8. L0=2. V=9. V=10.")]
    [InlineData(6, 6, false, "L0=1. L0=2.")]
    [InlineData(6, 6, true, "L0=1. L6=7. L6=8. L0=2. L6=7. L6=8.")]
    [InlineData(7, 6, false, "L0=1. L7=8. L7=9. L0=2. L7=8. L7=9.")]
    [InlineData(7, 6, true, "L0=1. L7=8. L7=9. L0=2. V=9. V=10.")]
    public void RestartImportPreservesWordsBoundaryBehavior(int current, int restart, bool replacement, string expected)
    {
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var abs = new AbstractNum(new MultiLevelType { Val = MultiLevelValues.Multilevel }) { AbstractNumberId = 1 };
        for (var i = 0; i < 9; i++)
        {
            var level = new Level(new StartNumberingValue { Val = i + 1 },
                new NumberingFormat { Val = NumberFormatValues.Decimal }) { LevelIndex = i };
            if (!replacement && i == current) level.Append(new LevelRestart { Val = restart });
            level.Append(new LevelText { Val = "L" + i + "=%" + (i + 1) + "." });
            abs.Append(level);
        }
        var numbering = new Numbering(abs,
            new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
        if (replacement)
            numbering.Append(new NumberingInstance(new AbstractNumId { Val = 1 },
                new LevelOverride(new StartOverrideNumberingValue { Val = 7 },
                    new Level(new StartNumberingValue { Val = 9 },
                        new NumberingFormat { Val = NumberFormatValues.Decimal },
                        new LevelRestart { Val = restart },
                        new LevelText { Val = "V=%" + (current + 1) + "." }) { LevelIndex = current })
                    { LevelIndex = current }) { NumberID = 2 });
        main.AddNewPart<NumberingDefinitionsPart>().Numbering = numbering;
        var body = new Body();
        var levels = new[] { 0, current, current, 0, current, current };
        for (var i = 0; i < levels.Length; i++)
            body.Append(new Paragraph(new ParagraphProperties(new NumberingProperties(
                new NumberingLevelReference { Val = levels[i] },
                new NumberingId { Val = replacement && i >= 4 ? 2 : 1 })), new Run(new Text("P" + i))));
        main.Document = new Document(body);
        var original = numbering.OuterXml;
        if (current == 0) Assert.Equal(-1, new DxpNumberingResolver(doc).GetRestartTrigger(1, 0));
        DxpITextVisitor[] visitors = [new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreatePlainConfig()),
            new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig())];
        foreach (var visitor in visitors)
        {
            var output = DxpExport.ExportToString(doc, visitor);
            if (visitor is DxpHtmlVisitor) output = new HtmlParser().ParseDocument(output).Body!.TextContent;
            var markers = string.Join(" ", Regex.Matches(output, @"(?:L[0-8]|V)=[-0-9]+\.").Cast<Match>().Select(m => m.Value));
            Assert.Equal(expected, markers);
            Assert.Equal(original, numbering.OuterXml);
        }
    }
}
