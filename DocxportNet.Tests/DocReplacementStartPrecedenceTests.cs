using System.Text.RegularExpressions;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Doc;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class DocReplacementStartPrecedenceTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 0, 1, 8 })
            foreach (var mode in new[] { "shared-high", "shared-low", "single-high", "single-low", "shared-omitted", "single-omitted" })
                yield return new object[] { depth, mode };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void ReplacementStartPrevailsThroughDoc(int depth, string mode)
    {
        var bytes = CreateSource(depth, mode);
        var binary = DxpDocExport.Export(bytes);
        using (var input = new MemoryStream(binary))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            Assert.Single(index.Lists.Definitions);
            var instance = index.Lists.Instances.Last();
            var expectedStart = mode.EndsWith("high") || mode.EndsWith("omitted") ? 7 : 2;
            Assert.Equal(expectedStart, instance.StartOverrides[depth]);
            Assert.Equal(expectedStart, instance.FormattingOverrides[depth].StartAt);
            Assert.Equal(mode.StartsWith("single") ? 1 : 2, index.Lists.Instances.Count);
        }
        using var sourceStream = new MemoryStream(bytes);
        using var source = WordprocessingDocument.Open(sourceStream, false);
        var numbering = source.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        var original = numbering.OuterXml;
        _ = DocOpenXmlListReader.Read(source.MainDocumentPart, _ => DocCharacterFormatting.Empty, _ => DocParagraphFormatting.Empty);
        Assert.Equal(original, numbering.OuterXml);
        using var projectedStream = new MemoryStream(DxpDocToDocx.Project(binary).DocxBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
        Assert.Empty(new OpenXmlValidator().Validate(projected));
        var projectedNumbering = projected.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        Assert.Equal(1, projectedNumbering.Elements<AbstractNum>().Count());
        Assert.Equal(mode.StartsWith("single") ? 1 : 2, projectedNumbering.Elements<NumberingInstance>().Count());
        Assert.Equal(1,
            projectedNumbering.Descendants<LevelOverride>().Count(x => x.Level != null));
        var expected = Labels(depth, mode);
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
                Assert.Equal(string.Concat(expected.Select((marker, i) => marker + "Item" + i)), Regex.Replace(output, @"\s+", ""));
            }
        }
    }

    public static string[] Labels(int depth, string mode)
    {
        var start = mode.EndsWith("high") || mode.EndsWith("omitted") ? 7 : 2;
        var values = mode.StartsWith("single")
            ? Enumerable.Range(start, 6).ToArray()
            : new[] { 1, 2, start, start + 1, start + 2, start + 3 };
        string Roman(int value) => new[] { "", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII" }[value];
        return values.Select((value, i) => "BEGIN" +
            (mode.StartsWith("single") || i is 2 or 3 or 5 ? Roman(value) : value.ToString()) + "END").ToArray();
    }

    public static byte[] CreateSource(int depth, string mode)
    {
        using var stream = new MemoryStream();
        stream.Write(DocSharedCounterOwnerTests.CreateSource(depth, "both"));
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            var replacement = main.NumberingDefinitionsPart!.Numbering!.Elements<NumberingInstance>()
                .Single(x => x.NumberID!.Value == 2).Elements<LevelOverride>().Single();
            replacement.StartOverrideNumberingValue!.Val = mode.EndsWith("high") ? 2 : 7;
            replacement.Level!.StartNumberingValue!.Val = mode.EndsWith("high") || mode.EndsWith("omitted") ? 7 : 2;
            if (mode.EndsWith("omitted")) replacement.Level.StartNumberingValue!.Remove();
            if (mode.StartsWith("single"))
                foreach (var paragraph in main.Document.Body!.Elements<Paragraph>())
                    paragraph.ParagraphProperties!.NumberingProperties!.NumberingId!.Val = 2;
        }
        return stream.ToArray();
    }
}
