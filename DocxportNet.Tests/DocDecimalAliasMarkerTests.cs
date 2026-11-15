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

public sealed class DocDecimalAliasMarkerTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 0, 8 })
            foreach (var mode in new[] { "definition-1", "definition-98", "override-1", "override-98" })
                yield return new object[] { depth, mode };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void DecimalAliasMarkersSurviveDoc(int depth, string mode)
    {
        var bytes = CreateSource(depth, mode);
        var binary = CreateBinary(depth, mode);
        using (var input = new MemoryStream(binary))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            Assert.Single(index.Lists.Definitions);
            var instance = Assert.Single(index.Lists.Instances);
            Assert.Empty(instance.StartOverrides);
            Assert.Equal(mode.StartsWith("override-"), instance.FormattingOverrides.ContainsKey(depth));
            Assert.Equal((byte)0x28, mode.StartsWith("override-") ? instance.FormattingOverrides[depth].NumberFormat
                : index.Lists.Definitions.Single().Levels[depth].NumberFormat);
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
        Assert.Single(projectedNumbering.Elements<NumberingInstance>());
        Assert.Equal(mode.StartsWith("override-") ? 1 : 0, projectedNumbering.Descendants<LevelOverride>().Count());
        Assert.All(projectedNumbering.Descendants<NumberingFormat>(),
            format => Assert.Equal(NumberFormatValues.Decimal, format.Val!.Value));
        using (var canonicalStream = new MemoryStream(DxpDocExport.Export(projectedStream.ToArray())))
        using (var canonical = new DocTextIndexWalker().Index(canonicalStream))
        {
            Assert.All(canonical.Lists.Definitions.SelectMany(x => x.Levels), level => Assert.Equal((byte)0, level.NumberFormat));
            Assert.All(canonical.Lists.Instances.SelectMany(x => x.FormattingOverrides.Values), level => Assert.Equal((byte)0, level.NumberFormat));
        }
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
                Assert.Equal(Regex.Replace(string.Concat(expected.Select((marker, i) => marker + "Item" + i)), @"\s+", ""), Regex.Replace(output, @"\s+", ""));
            }
        }
    }

    // Editable ListString values confirmed in Word on 2026-11-15. Formatting-only
    // overrides retain the abstract level's start (2), ignoring their own start.
    public static string[] Labels(int depth, string mode) => mode.StartsWith("override-")
        ? ["BEGIN2END", "BEGIN3END", "BEGIN4END", "BEGIN5END", "BEGIN6END", "BEGIN7END"]
        : mode.EndsWith("-98")
        ? ["BEGIN98END", "BEGIN99END", "BEGIN100END", "BEGIN101END", "BEGIN102END", "BEGIN103END"]
        : ["BEGIN1END", "BEGIN2END", "BEGIN3END", "BEGIN4END", "BEGIN5END", "BEGIN6END"];

    // Construct a generic DOC by changing only the selected LVLF.nfc byte of a
    // canonical decimal export. Use structure offsets rather than byte searching.
    public static byte[] CreateBinary(int depth, string mode)
    {
        var bytes = DxpDocExport.Export(CreateSource(depth, mode));
        string streamName;
        long offset;
        using (var input = new MemoryStream(bytes))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            var name = mode.StartsWith("override-") ? "ListOverrides" : "ListDefinitions";
            var location = index.Structure.FindLocation(name)!;
            var node = new DocStructureNode("FibLocation", name, location.StreamName, location.Offset, location.Length);
            DocListNavigator.Expand(index.Structure, node);
            static IEnumerable<DocStructureNode> Walk(DocStructureNode node) =>
                node.Children.SelectMany(child => new[] { child }.Concat(Walk(child)));
            var levels = Walk(node).Where(x => x.Kind == "LVL").ToArray();
            var level = levels[mode.StartsWith("override-") ? 0 : depth];
            streamName = level.StreamName!;
            offset = level.Offset!.Value + 4; // LVLF.iStartAt precedes nfc.
            Assert.Equal(0, index.Structure.ReadRange(streamName, offset, 1)[0]);
        }
        using var output = new MemoryStream();
        output.Write(bytes);
        output.Position = 0;
        using (var compound = OpenMcdf.RootStorage.Open(output, OpenMcdf.StorageModeFlags.LeaveOpen))
        using (var table = compound.OpenStream(streamName))
        {
            table.Position = offset;
            table.WriteByte(0x28);
            compound.Flush();
        }
        return output.ToArray();
    }

    public static byte[] CreateSource(int depth, string mode)
    {
        using var stream = new MemoryStream();
        stream.Write(DocOverrideFlagTests.CreateSource(depth, mode.StartsWith("override-") ? "format" : "none"));
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            var numbering = main.NumberingDefinitionsPart!.Numbering!;
            var level = mode.StartsWith("override-")
                ? numbering.Descendants<LevelOverride>().Single().GetFirstChild<Level>()!
                : numbering.Elements<AbstractNum>().Single().Elements<Level>().Single(x => x.LevelIndex!.Value == depth);
            level.StartNumberingValue!.Val = int.Parse(mode.Split('-').Last());
            level.NumberingFormat!.Val = NumberFormatValues.Decimal;
            level.AppendChild(new NumberingSymbolRunProperties(new RunFonts
                { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial" }));
        }
        return stream.ToArray();
    }
}
