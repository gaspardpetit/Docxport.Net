using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Visitors.Doc;

namespace DocxportNet.Tests;

public sealed class MarkerDocRtlNormalizationTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var label in new[] { "ab", "αβ", "אב", "بت" })
        foreach (var state in new[] { "both-on", "cs-on-rtl-off", "cs-off-rtl-on", "both-off" })
            yield return new object[] { label, state };
        foreach (var label in new[] { "AéαБ", "AאבZ", "AبتZ", "אב 12", "(אב).", "A漢あZ" })
        foreach (var state in new[] { "absent", "off", "on", "rtl-only" })
            yield return new object[] { label, state };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void DocNormalizesOnlyListRtlFormattingAndPreservesScriptSlots(string label, string state)
    {
        var source = CreateSource(label, state);
        var rtl = state is "both-on" or "cs-off-rtl-on" or "rtl-only";
        bool? cs = state is "absent" or "rtl-only" ? null : state is "both-on" or "cs-on-rtl-off" or "on";
        using (var original = WordprocessingDocument.Open(new MemoryStream(source), false))
        {
            var numbering = original.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
            var before = numbering.OuterXml;
            _ = DxpExport.ExportToBytes(original, new DxpDocVisitor(), options: null);
            Assert.Equal(before, numbering.OuterXml);
        }
        using var projectedStream = new MemoryStream(DxpDocToDocx.Project(DxpDocExport.Export(source)).DocxBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(projected));
        var actual = projected.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Descendants<NumberingSymbolRunProperties>().Single();
        var forced = cs == true && !rtl;
        Assert.Equal(forced ? "Courier New" : "Arial", actual.RunFonts?.Ascii?.Value);
        Assert.Equal(forced ? "Courier New" : rtl ? "Arial" : "Times New Roman", actual.RunFonts?.HighAnsi?.Value);
        Assert.Equal(forced ? "Courier New" : "MS Mincho", actual.RunFonts?.EastAsia?.Value);
        Assert.Equal("Courier New", actual.RunFonts?.ComplexScript?.Value);
        Assert.Equal("24", actual.FontSize?.Val?.Value);
        Assert.Equal("30", actual.FontSizeComplexScript?.Val?.Value);
        if (rtl) Assert.False(actual.GetFirstChild<RightToLeftText>()?.Val?.Value ?? true);
        foreach (var run in projected.MainDocumentPart.Document.Descendants<Run>())
        {
            var body = run.RunProperties!;
            Assert.NotNull(body.GetFirstChild<RightToLeftText>());
            Assert.NotNull(body.GetFirstChild<ComplexScript>());
            Assert.True(body.GetFirstChild<RightToLeftText>()!.Val?.Value ?? true);
            Assert.True(body.GetFirstChild<ComplexScript>()!.Val?.Value ?? true);
        }
    }
    [Theory]
    [InlineData("ascii", "Symbol")]
    [InlineData("hAnsi", "Symbol")]
    [InlineData("eastAsia", "Symbol")]
    [InlineData("cs", "Symbol")]
    [InlineData("ascii", "Wingdings")]
    [InlineData("hAnsi", "Wingdings")]
    [InlineData("eastAsia", "Wingdings")]
    [InlineData("cs", "Wingdings")]
    public void SymbolSlotKeepsRtlEncoding(string slot, string family)
    {
        using var stream = new MemoryStream();
        stream.Write(CreateSource("ab", "cs-off-rtl-on"));
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var fonts = doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
                .Descendants<NumberingSymbolRunProperties>().Single().RunFonts!;
            fonts.SetAttribute(new OpenXmlAttribute("w", slot,
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main", family));
        }
        using var projectedStream = new MemoryStream(DxpDocToDocx.Project(DxpDocExport.Export(stream.ToArray())).DocxBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
        var props = projected.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Descendants<NumberingSymbolRunProperties>().Single();
        using var original = WordprocessingDocument.Open(new MemoryStream(stream.ToArray()), false);
        var originalFonts = original.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Descendants<NumberingSymbolRunProperties>().Single().RunFonts!;
        Assert.Equal(originalFonts.Ascii?.Value, props.RunFonts?.Ascii?.Value);
        Assert.Equal(originalFonts.HighAnsi?.Value, props.RunFonts?.HighAnsi?.Value);
        Assert.Equal(originalFonts.EastAsia?.Value, props.RunFonts?.EastAsia?.Value);
        Assert.Equal(originalFonts.ComplexScript?.Value, props.RunFonts?.ComplexScript?.Value);
        Assert.NotNull(props.GetFirstChild<RightToLeftText>());
        Assert.True(props.GetFirstChild<RightToLeftText>()!.Val?.Value ?? true);
    }

    public static byte[] CreateSource(string label, string state)
    {
        var rtl = state is "both-on" or "cs-off-rtl-on" or "rtl-only";
        bool? cs = state is "absent" or "rtl-only" ? null : state is "both-on" or "cs-on-rtl-off" or "on";
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var props = new NumberingSymbolRunProperties(new RunFonts
            { Ascii = "Arial", HighAnsi = "Times New Roman", EastAsia = "MS Mincho", ComplexScript = "Courier New" },
                new FontSize { Val = "24" }, new FontSizeComplexScript { Val = "30" });
            if (state != "absent" && state is not ("off" or "on")) props.Append(new RightToLeftText { Val = rtl });
            if (cs.HasValue) props.Append(new ComplexScript { Val = cs.Value });
            main.AddNewPart<NumberingDefinitionsPart>().Numbering = new Numbering(
                new AbstractNum(new Level(new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.Bullet }, new LevelText { Val = label }, props)
                    { LevelIndex = 0 }) { AbstractNumberId = 1 },
                new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
            main.Document = new Document(new Body(Enumerable.Range(0, 3).Select(i => new Paragraph(new ParagraphProperties(new NumberingProperties(
                new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 })),
                new Run(new RunProperties(new RightToLeftText(), new ComplexScript()), new Text("Item" + i))))));
            Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator(FileFormatVersions.Office2013).Validate(doc));
        }
        return stream.ToArray();
    }
}
