using System.Text.RegularExpressions;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class DocFormatOverrideMarkerTests
{
    [Theory]
    [MemberData(nameof(DocOverrideFlagTests.Cases), MemberType = typeof(DocOverrideFlagTests))]
    public void FormattingReplacementRetainsAbstractStartUntilActivated(int depth, string mode)
    {
        var bytes = CreateSource(depth, mode);
        var binary = DxpDocExport.Export(bytes);
        using var sourceStream = new MemoryStream(bytes);
        using var source = WordprocessingDocument.Open(sourceStream, false);
        using var projectedStream = new MemoryStream(DxpDocToDocx.Project(binary).DocxBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
        Assert.Empty(new OpenXmlValidator().Validate(projected));
        var expected = Enumerable.Range(mode is "start" or "both" ? 7 : 2, 6)
            .Select(n => "BEGIN" + (mode is "format" or "both" ? Roman(n) : n.ToString()) + "END").ToArray();
        foreach (var document in new[] { source, projected })
        {
            DxpITextVisitor[] visitors = [new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()),
                new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreateRichConfig()),
                new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreatePlainConfig()),
                new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig()),
                new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig())];
            foreach (var visitor in visitors)
            {
                var output = DxpExport.ExportToString(document, visitor);
                if (visitor is DxpHtmlVisitor or DxpMarkdownVisitor)
                    output = new HtmlParser().ParseDocument(output).Body!.TextContent;
                Assert.Equal(expected, Regex.Matches(output, "BEGIN.*?END").Select(m => m.Value));
                Assert.Equal(6, Regex.Matches(output, @"Item\d+").Count);
            }
        }
    }

    public static byte[] CreateSource(int depth, string mode)
    {
        using var stream = new MemoryStream();
        stream.Write(DocOverrideFlagTests.CreateSource(depth, mode));
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var replacement = document.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
                .Elements<NumberingInstance>().Single().Elements<LevelOverride>().SingleOrDefault()?.Level;
            if (replacement != null) replacement.NumberingFormat!.Val = NumberFormatValues.UpperRoman;
        }
        return stream.ToArray();
    }

    public static string Roman(int value) => new[] { "", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII" }[value];
}
