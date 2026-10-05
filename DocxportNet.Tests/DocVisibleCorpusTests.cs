using DocxportNet.Doc;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Validation;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace DocxportNet.Tests;

public class DocVisibleCorpusTests
{
    [Theory]
    [InlineData("WordPositionalTabTenTwipCellMarginsAllStories")]
    [InlineData("WordPositionalTabCellMarginsAllStories")]
    [InlineData("WordTableCellPositionalTabProbe")]
    public void TableCellPositionalTabKnownGapsRetainIndexedStops(string name)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures",
            "DocKnownGaps"));
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var native = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
        var expected = ReadStories(source);
        var expectedStops = ReadEffectiveTabStops(source);
        foreach (var binary in new[] { native, DxpDocExport.Export(source) })
        {
            var projected = DxpDocToDocx.Project(binary);
            Assert.Empty(projected.Coverage.OmittedCharacters);
            Assert.Equal(expected, ReadStories(projected.DocxBytes));
            Assert.Equal(expectedStops, ReadEffectiveTabStops(projected.DocxBytes));
            Assert.Empty(Validate(projected.DocxBytes));
        }
    }

    [Theory]
    [InlineData("WordLineCharacterGridAllStories", (byte)1)]
    [InlineData("WordSnapToCharacterGridAllStories", (byte)3)]
    public void CharacterGridPairsRetainPitchModeAndRunOverrides(string name, byte mode)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var stories = ReadStories(source);
        foreach (var doc in new[]
        {
            File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
            DxpDocExport.Export(source)
        })
        {
            foreach (var binary in new[] { doc,
                DxpDocExport.Export(DxpDocToDocx.Project(doc).DocxBytes) })
            {
                using (var input = new MemoryStream(binary))
                using (var index = new DocTextIndexWalker().Index(input))
                {
                    var section = DocSectionFormatting.Read(index,
                        Assert.Single(index.Sections));
                    Assert.Equal(mode, section.GridMode);
                    Assert.Equal((ushort)480, section.GridLinePitch);
                    Assert.Equal(40960, section.GridCharacterSpace);
                    Assert.Contains(index.CharacterFormatting,
                        x => x.Formatting.SnapToGrid == false);
                }
                var projection = DxpDocToDocx.Project(binary);
                Assert.Empty(projection.Coverage.OmittedCharacters);
                using var package = WordprocessingDocument.Open(
                    new MemoryStream(projection.DocxBytes), false);
                var main = package.MainDocumentPart!;
                var grid = main.Document.Body!.Elements<SectionProperties>()
                    .Single().GetFirstChild<DocGrid>()!;
                Assert.Equal(mode == 1 ? DocGridValues.LinesAndChars
                    : DocGridValues.SnapToChars, grid.Type!.Value);
                Assert.Equal(40960, grid.CharacterSpace!.Value);
                Assert.False(main.HeaderParts.Single().Header!.Descendants<Run>()
                    .Single().RunProperties!.GetFirstChild<SnapToGrid>()!.Val!.Value);
                Assert.Equal(stories, ReadStories(projection.DocxBytes));
                Assert.Empty(Validate(projection.DocxBytes));
            }
        }
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InheritedCharacterEmphasisAndGridSnapRetainLayeredValues(bool native)
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedCharacterEmphasisGridAllStories"));
        var source = File.ReadAllBytes(stem + ".docx");
        var expected = ReadEffectiveRunFormatting(source);
        Assert.Equal("none", expected["body.paragraph0.char0.emphasis"]);
        Assert.Equal("true", expected["body.paragraph0.char0.snapToGrid"]);
        Assert.Equal("dot", expected["body.paragraph1.char0.emphasis"]);
        Assert.Equal("false", expected["body.paragraph1.char0.snapToGrid"]);
        Assert.Equal("dot", expected["section0.header.default.paragraph0.char0.emphasis"]);
        Assert.Equal("dot", expected["section0.footer.default.paragraph0.char0.emphasis"]);
        var bytes = native ? File.ReadAllBytes(stem + ".doc") :
            DxpDocExport.Export(source);
        foreach (var binary in new[] { bytes,
            DxpDocExport.Export(DxpDocToDocx.Project(bytes).DocxBytes) })
        {
            using var index = new DocTextIndexWalker().Index(new MemoryStream(binary));
            var baseStyle = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Grid run base");
            var derived = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Grid run derived");
            Assert.Equal(baseStyle.Index, derived.BasedOnIndex);
            Assert.Equal((byte)1, baseStyle.DirectCharacterFormatting?.EmphasisMarkCode);
            Assert.Null(derived.DirectCharacterFormatting?.EmphasisMarkCode);
            Assert.Equal((byte)1, derived.CharacterFormatting.EmphasisMarkCode);
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            AssertEffectiveRunFormatting(source, projected);
            Assert.Equal(ReadStories(source), ReadStories(projected));
            Assert.Empty(Validate(projected));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InheritedEmphasisRetainsStyleAndDirectRunOverrides(bool native)
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedEmphasisAllStories"));
        var source = File.ReadAllBytes(stem + ".docx");
        var expected = ReadEffectiveRunFormatting(source);
        Assert.Equal("none", expected["body.paragraph0.char0.emphasis"]);
        Assert.Equal("comma", expected["body.paragraph1.char0.emphasis"]);
        Assert.Equal("dot", expected["section0.header.default.paragraph0.char0.emphasis"]);
        Assert.Equal("dot", expected["section0.footer.default.paragraph0.char0.emphasis"]);
        var bytes = native ? File.ReadAllBytes(stem + ".doc") :
            DxpDocExport.Export(source);
        foreach (var binary in new[] { bytes,
            DxpDocExport.Export(DxpDocToDocx.Project(bytes).DocxBytes) })
        {
            using var index = new DocTextIndexWalker().Index(new MemoryStream(binary));
            var baseStyle = Assert.Single(index.StyleDefinitions,
                x => x.Name == "CJK spacing base");
            var derived = Assert.Single(index.StyleDefinitions,
                x => x.Name == "CJK spacing derived");
            Assert.Equal(baseStyle.Index, derived.BasedOnIndex);
            Assert.Equal((byte)1, baseStyle.DirectCharacterFormatting?.EmphasisMarkCode);
            Assert.Null(derived.DirectCharacterFormatting?.EmphasisMarkCode);
            Assert.Equal((byte)1, derived.CharacterFormatting.EmphasisMarkCode);
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            AssertEffectiveRunFormatting(source, projected);
            Assert.Equal(ReadStories(source), ReadStories(projected));
            Assert.Empty(Validate(projected));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InheritedRunGridSnapRetainsCharacterStylesAndDirectReset(bool native)
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedRunGridSnapAllStories"));
        var source = File.ReadAllBytes(stem + ".docx");
        var expected = ReadEffectiveRunFormatting(source);
        Assert.Equal("true", expected["body.paragraph0.char0.snapToGrid"]);
        Assert.Equal("false", expected["body.paragraph1.char0.snapToGrid"]);
        Assert.Equal("false", expected["section0.header.default.paragraph0.char0.snapToGrid"]);
        Assert.Equal("false", expected["section0.footer.default.paragraph0.char0.snapToGrid"]);
        var bytes = native ? File.ReadAllBytes(stem + ".doc") :
            DxpDocExport.Export(source);
        foreach (var binary in new[] { bytes,
            DxpDocExport.Export(DxpDocToDocx.Project(bytes).DocxBytes) })
        {
            using var index = new DocTextIndexWalker().Index(new MemoryStream(binary));
            var baseStyle = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Grid run base");
            var derived = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Grid run derived");
            Assert.Equal(baseStyle.Index, derived.BasedOnIndex);
            Assert.False(baseStyle.DirectCharacterFormatting?.SnapToGrid);
            Assert.Null(derived.DirectCharacterFormatting?.SnapToGrid);
            Assert.False(derived.CharacterFormatting.SnapToGrid);
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            AssertEffectiveRunFormatting(source, projected);
            Assert.Equal(ReadStories(source), ReadStories(projected));
            Assert.Empty(Validate(projected));
        }
    }

    [Fact]
    public void SectionLineGridPairRetainsSnapStylesAndDirectOverrides()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordSectionLineGridSnapAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var stories = ReadStories(source);
        foreach (var doc in new[]
        {
            File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
            DxpDocExport.Export(source)
        })
        {
            foreach (var binary in new[] { doc,
                DxpDocExport.Export(DxpDocToDocx.Project(doc).DocxBytes) })
            {
                using (var input = new MemoryStream(binary))
                using (var index = new DocTextIndexWalker().Index(input))
                {
                    Assert.Equal((ushort)480, DocSectionFormatting.Read(index,
                        Assert.Single(index.Sections)).GridLinePitch);
                    Assert.True(Assert.Single(index.StyleDefinitions,
                        x => x.Name == "Grid base").ParagraphFormatting?.SnapToGrid != false);
                    Assert.False(Assert.Single(index.StyleDefinitions,
                        x => x.Name == "Grid derived").ParagraphFormatting?.SnapToGrid);
                    Assert.Contains(index.ParagraphStyles,
                        x => x.Formatting?.SnapToGrid == false);
                    Assert.Contains(index.ParagraphStyles,
                        x => x.Formatting?.SnapToGrid == true);
                }
                var projection = DxpDocToDocx.Project(binary);
                Assert.Empty(projection.Coverage.OmittedCharacters);
                using var package = WordprocessingDocument.Open(
                    new MemoryStream(projection.DocxBytes), false);
                var main = package.MainDocumentPart!;
                Assert.Equal(480, main.Document.Body!.Elements<SectionProperties>()
                    .Single().GetFirstChild<DocGrid>()!.LinePitch!.Value);
                Assert.False(main.Document.Body!.Elements<Paragraph>().First()
                    .ParagraphProperties!.GetFirstChild<SnapToGrid>()!.Val!.Value);
                Assert.True(main.FooterParts.Single().Footer!.Elements<Paragraph>()
                    .Single().ParagraphProperties!.GetFirstChild<SnapToGrid>()!
                    .Val!.Value);
                Assert.Equal(stories, ReadStories(projection.DocxBytes));
                Assert.Empty(Validate(projection.DocxBytes));
            }
        }
    }
    [Fact]
    public void CharacterGridAdjustRightPairRetainsStyleAndDirectOverrides()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordCharacterGridAdjustRightAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var stories = ReadStories(source);
        foreach (var doc in new[]
        {
            File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
            DxpDocExport.Export(source)
        })
        {
            foreach (var binary in new[] { doc,
                DxpDocExport.Export(DxpDocToDocx.Project(doc).DocxBytes) })
            {
                using (var input = new MemoryStream(binary))
                using (var index = new DocTextIndexWalker().Index(input))
                {
                    var grid = DocSectionFormatting.Read(index, Assert.Single(index.Sections));
                    Assert.Equal((ushort)480, grid.GridLinePitch);
                    Assert.Equal((byte)1, grid.GridMode);
                    Assert.Equal(40960, grid.GridCharacterSpace);
                    Assert.True(Assert.Single(index.StyleDefinitions,
                        x => x.Name == "Right grid base").ParagraphFormatting?.AdjustRightIndent != false);
                    Assert.False(Assert.Single(index.StyleDefinitions,
                        x => x.Name == "Right grid derived").ParagraphFormatting?.AdjustRightIndent);
                    Assert.Contains(index.ParagraphStyles,
                        x => x.Formatting?.AdjustRightIndent == false);
                    Assert.Contains(index.ParagraphStyles,
                        x => x.Formatting?.AdjustRightIndent == true);
                }
                var projection = DxpDocToDocx.Project(binary);
                Assert.Empty(projection.Coverage.OmittedCharacters);
                using var package = WordprocessingDocument.Open(
                    new MemoryStream(projection.DocxBytes), false);
                var main = package.MainDocumentPart!;
                Assert.Equal(480, main.Document.Body!.Elements<SectionProperties>()
                    .Single().GetFirstChild<DocGrid>()!.LinePitch!.Value);
                Assert.False(main.Document.Body!.Elements<Paragraph>().First()
                    .ParagraphProperties!.GetFirstChild<AdjustRightIndent>()!.Val!.Value);
                Assert.True(main.FooterParts.Single().Footer!.Elements<Paragraph>()
                    .Single().ParagraphProperties!.GetFirstChild<AdjustRightIndent>()!
                    .Val!.Value);
                Assert.Equal(stories, ReadStories(projection.DocxBytes));
                Assert.Empty(Validate(projection.DocxBytes));
            }
        }
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InheritedTextAlignmentRetainsVisibleStyleAndDirectReset(bool native)
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedTextAlignmentAllStories"));
        var source = File.ReadAllBytes(stem + ".docx");
        var expected = ReadEffectiveParagraphLayout(source);
        Assert.Equal("baseline", expected["body.paragraph0.textAlignment"]);
        Assert.Equal("top", expected["body.paragraph2.textAlignment"]);
        Assert.Equal("top", expected["section0.header.default.paragraph0.textAlignment"]);
        Assert.Equal("top", expected["section0.footer.default.paragraph0.textAlignment"]);
        var bytes = native ? File.ReadAllBytes(stem + ".doc") :
            DxpDocExport.Export(source);
        foreach (var binary in new[] { bytes,
            DxpDocExport.Export(DxpDocToDocx.Project(bytes).DocxBytes) })
        {
            using var index = new DocTextIndexWalker().Index(new MemoryStream(binary));
            var baseStyle = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Right grid base");
            var derived = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Right grid derived");
            Assert.Equal((short)0, baseStyle.DirectParagraphFormatting?.TextAlignmentCode);
            Assert.Null(derived.DirectParagraphFormatting?.TextAlignmentCode);
            Assert.Equal((short)0, derived.ParagraphFormatting?.TextAlignmentCode);
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            AssertEffectiveParagraphLayout(source, projected);
            Assert.Equal(ReadStories(source), ReadStories(projected));
            Assert.Empty(Validate(projected));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InheritedGridAdjustedRightIndentRetainsVisibleDerivedStyle(bool native)
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedAdjustRightGridAllStories"));
        var source = File.ReadAllBytes(stem + ".docx");
        Assert.Equal("false", ReadEffectiveParagraphLayout(source)[
            "body.paragraph2.adjustRightIndent"]);
        var bytes = native ? File.ReadAllBytes(stem + ".doc") :
            DxpDocExport.Export(source);
        foreach (var binary in new[] { bytes,
            DxpDocExport.Export(DxpDocToDocx.Project(bytes).DocxBytes) })
        {
            using var index = new DocTextIndexWalker().Index(new MemoryStream(binary));
            var style = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Right grid derived");
            Assert.False(style.DirectParagraphFormatting?.AdjustRightIndent);
            Assert.False(style.ParagraphFormatting?.AdjustRightIndent);
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            AssertEffectiveParagraphLayout(source, projected);
            Assert.Equal(ReadStories(source), ReadStories(projected));
            Assert.Empty(Validate(projected));
        }
    }

    [Fact]
    public void EastAsianLineBreakingPairRetainsStyleAndDirectControlsInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordEastAsianLineBreakingAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var native = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
        var expectedStories = ReadStories(source);
        foreach (var doc in new[] { native, DxpDocExport.Export(source) })
        {
            foreach (var binary in new[] { doc,
                DxpDocExport.Export(DxpDocToDocx.Project(doc).DocxBytes) })
            {
                using (var input = new MemoryStream(binary))
                using (var index = new DocTextIndexWalker().Index(input))
                {
                    var baseStyle = Assert.Single(index.StyleDefinitions,
                        x => x.Name == "CJK break base");
                    Assert.True(baseStyle.ParagraphFormatting?.Kinsoku != false);
                    Assert.False(baseStyle.ParagraphFormatting?.WordWrap);
                    Assert.True(Assert.Single(index.StyleDefinitions,
                        x => x.Name == "CJK break derived")
                        .ParagraphFormatting?.WordWrap);
                    Assert.Contains(index.ParagraphStyles,
                        x => x.Formatting?.Kinsoku == false);
                    Assert.Contains(index.ParagraphStyles,
                        x => x.Formatting?.WordWrap == true);
                }
                var projected = DxpDocToDocx.Project(binary).DocxBytes;
                using var package = WordprocessingDocument.Open(
                    new MemoryStream(projected), false);
                var main = package.MainDocumentPart!;
                var styles = main.StyleDefinitionsPart!.Styles!;
                Assert.False(styles.Elements<Style>().Single(x =>
                    x.StyleName?.Val?.Value == "CJK break base")
                    .StyleParagraphProperties!.WordWrap!.Val!.Value);
                Assert.True(styles.Elements<Style>().Single(x =>
                    x.StyleName?.Val?.Value == "CJK break derived")
                    .StyleParagraphProperties!.WordWrap!.Val?.Value ?? true);
                Assert.False(main.Document.Body!.Elements<Paragraph>().First()
                    .ParagraphProperties!.Kinsoku!.Val!.Value);
                Assert.False(main.FooterParts.Single().Footer!.Elements<Paragraph>()
                    .Single().ParagraphProperties!.Kinsoku!.Val!.Value);
                Assert.Equal(expectedStories, ReadStories(projected));
                Assert.Empty(Validate(projected));
            }
        }
    }
    [Fact]
    public void EastAsianAutomaticSpacingPairRetainsStyleDirectAndDefaultFontInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordEastAsianAutomaticSpacingAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var native = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
        var expectedStories = ReadStories(source);
        foreach (var doc in new[] { native, DxpDocExport.Export(source) })
        {
            foreach (var binary in new[] { doc,
                DxpDocExport.Export(DxpDocToDocx.Project(doc).DocxBytes) })
            {
                using (var input = new MemoryStream(binary))
                using (var index = new DocTextIndexWalker().Index(input))
                {
                    Assert.Equal("Yu Mincho", index.DefaultCharacterFormatting.EastAsiaFontName);
                    var baseStyle = Assert.Single(index.StyleDefinitions,
                        x => x.Name == "CJK spacing base");
                    Assert.True(baseStyle.ParagraphFormatting?.AutoSpaceDE != false);
                    Assert.False(baseStyle.ParagraphFormatting?.AutoSpaceDN);
                    Assert.True(Assert.Single(index.StyleDefinitions,
                        x => x.Name == "CJK spacing derived")
                        .ParagraphFormatting?.AutoSpaceDN);
                    Assert.Contains(index.ParagraphStyles,
                        x => x.Formatting?.AutoSpaceDE == false);
                    Assert.Contains(index.ParagraphStyles,
                        x => x.Formatting?.AutoSpaceDN == true);
                }
                var projected = DxpDocToDocx.Project(binary).DocxBytes;
                using var package = WordprocessingDocument.Open(
                    new MemoryStream(projected), false);
                var main = package.MainDocumentPart!;
                var styles = main.StyleDefinitionsPart!.Styles!;
                Assert.False(styles.Elements<Style>().Single(x =>
                    x.StyleName?.Val?.Value == "CJK spacing base")
                    .StyleParagraphProperties!.AutoSpaceDN!.Val!.Value);
                Assert.True(styles.Elements<Style>().Single(x =>
                    x.StyleName?.Val?.Value == "CJK spacing derived")
                    .StyleParagraphProperties!.AutoSpaceDN!.Val?.Value ?? true);
                Assert.False(main.Document.Body!.Elements<Paragraph>().First()
                    .ParagraphProperties!.AutoSpaceDE!.Val!.Value);
                Assert.False(main.FooterParts.Single().Footer!.Elements<Paragraph>()
                    .Single().ParagraphProperties!.AutoSpaceDE!.Val!.Value);
                Assert.Equal(expectedStories, ReadStories(projected));
                Assert.Empty(Validate(projected));
            }
        }
    }
    [Fact]
    public void EastAsianEmphasisPairRetainsStyleAndDirectMarksInAllStories()
    {
        Assert.Equal((byte)4, DocCharacterFormattingReader.ParseSprms(
            new byte[] { 0x34, 0x2A, 4 }).EmphasisMarkCode);
        Assert.Throws<InvalidDataException>(() => DocCharacterFormattingReader.ParseSprms(
            new byte[] { 0x34, 0x2A, 5 }));
        Assert.Throws<InvalidDataException>(() =>
            DocPlainTextWriter.EncodeCharacterProperties(
                DocCharacterFormatting.Empty with { EmphasisMarkCode = 5 }));
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordEastAsianEmphasisAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var expectedStories = ReadStories(source);
        foreach (var doc in new[] { File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
            DxpDocExport.Export(source) })
        {
            foreach (var binary in new[] { doc,
                DxpDocExport.Export(DxpDocToDocx.Project(doc).DocxBytes) })
            {
                using (var input = new MemoryStream(binary))
                using (var index = new DocTextIndexWalker().Index(input))
                {
                    Assert.Equal((byte)1, Assert.Single(index.StyleDefinitions,
                        x => x.Name == "CJK spacing base")
                        .CharacterFormatting.EmphasisMarkCode);
                    Assert.Equal((byte)3, Assert.Single(index.StyleDefinitions,
                        x => x.Name == "CJK spacing derived")
                        .CharacterFormatting.EmphasisMarkCode);
                    foreach (var code in new byte[] { 0, 2, 4 })
                        Assert.Contains(index.CharacterFormatting,
                            x => x.Formatting.EmphasisMarkCode == code);
                }
                var projection = DxpDocToDocx.Project(binary);
                Assert.Empty(projection.Coverage.OmittedCharacters);
                Assert.Equal(expectedStories, ReadStories(projection.DocxBytes));
                using var package = WordprocessingDocument.Open(
                    new MemoryStream(projection.DocxBytes), false);
                var main = package.MainDocumentPart!;
                var styles = main.StyleDefinitionsPart!.Styles!;
                Assert.Equal(EmphasisMarkValues.Dot,
                    Assert.Single(styles.Elements<Style>(), x =>
                        x.StyleName?.Val?.Value == "CJK spacing base")
                        .StyleRunProperties!.GetFirstChild<Emphasis>()!.Val!.Value);
                Assert.Equal(EmphasisMarkValues.Circle,
                    Assert.Single(styles.Elements<Style>(), x =>
                        x.StyleName?.Val?.Value == "CJK spacing derived")
                        .StyleRunProperties!.GetFirstChild<Emphasis>()!.Val!.Value);
                Assert.Equal(EmphasisMarkValues.None,
                    main.Document.Body!.Elements<Paragraph>().First()
                        .Descendants<Run>().First().RunProperties!
                        .GetFirstChild<Emphasis>()!.Val!.Value);
                Assert.Equal(EmphasisMarkValues.Comma,
                    main.Document.Body!.Elements<Paragraph>().Skip(1).First()
                        .Descendants<Run>().First().RunProperties!
                        .GetFirstChild<Emphasis>()!.Val!.Value);
                Assert.Equal(EmphasisMarkValues.UnderDot,
                    main.HeaderParts.Single().Header!.Descendants<Run>().First()
                        .RunProperties!.GetFirstChild<Emphasis>()!.Val!.Value);
                Assert.Empty(Validate(projection.DocxBytes));
            }
        }
    }

    [Fact]
    public void OpenTypeAdvanceReaderMapsFormatFourAndRejectsMissingGlyphs()
    {
        var font = new byte[244];
        static void Word(byte[] bytes, int offset, ushort value) =>
            System.Buffers.Binary.BinaryPrimitives.WriteUInt16BigEndian(
                bytes.AsSpan(offset), value);
        static void Dword(byte[] bytes, int offset, uint value) =>
            System.Buffers.Binary.BinaryPrimitives.WriteUInt32BigEndian(
                bytes.AsSpan(offset), value);
        Dword(font, 0, 0x00010000); Word(font, 4, 5);
        var tables = new[] { ("head", 92, 54), ("hhea", 146, 36),
            ("maxp", 182, 6), ("hmtx", 188, 12), ("cmap", 200, 44) };
        for (var i = 0; i < tables.Length; i++)
        {
            var (name, offset, length) = tables[i];
            System.Text.Encoding.ASCII.GetBytes(name).CopyTo(font, 12 + i * 16);
            Dword(font, 12 + i * 16 + 8, (uint)offset);
            Dword(font, 12 + i * 16 + 12, (uint)length);
        }
        Word(font, 92 + 18, 1000);  // unitsPerEm
        Word(font, 146 + 34, 3);   // numberOfHMetrics
        Word(font, 182 + 4, 3);    // glyph count
        Word(font, 188 + 4, 500); // A advance
        Word(font, 188 + 8, 750); // B advance
        Word(font, 200 + 2, 1);   // one cmap encoding
        Word(font, 200 + 4, 3); Word(font, 200 + 6, 1);
        Dword(font, 200 + 8, 12);
        var subtable = 212;
        Word(font, subtable, 4); Word(font, subtable + 2, 32);
        Word(font, subtable + 6, 4); // two segments, including the sentinel
        Word(font, subtable + 14, 66); Word(font, subtable + 16, 0xFFFF);
        Word(font, subtable + 20, 65); Word(font, subtable + 22, 0xFFFF);
        Word(font, subtable + 24, unchecked((ushort)-64));
        Word(font, subtable + 26, 1);
        var metrics = new DocOpenTypeAdvances(font);
        Assert.True(metrics.TryMeasure("AB", 12, out var width));
        Assert.Equal(15, width);
        Assert.False(metrics.TryMeasure("AC", 12, out _));
        Assert.Throws<InvalidDataException>(() => new DocOpenTypeAdvances(font[..90]));

        // Add a horizontal format-0 kern table for A/B without changing the
        // cmap or advances. The table fit opts into this pair adjustment.
        var kerned = new byte[284];
        Array.Copy(font, 0, kerned, 0, 92);
        Array.Copy(font, 92, kerned, 108, 152);
        Word(kerned, 4, 6);
        for (var i = 0; i < tables.Length; i++)
            Dword(kerned, 12 + i * 16 + 8, (uint)(tables[i].Item2 + 16));
        System.Text.Encoding.ASCII.GetBytes("kern").CopyTo(kerned, 92);
        Dword(kerned, 92 + 8, 260); Dword(kerned, 92 + 12, 24);
        Word(kerned, 260 + 2, 1); // one subtable
        Word(kerned, 264 + 2, 20); Word(kerned, 264 + 4, 1);
        Word(kerned, 264 + 6, 1); // one A/B pair
        Word(kerned, 278, 1); Word(kerned, 280, 2);
        Word(kerned, 282, unchecked((ushort)-50));
        var kernedMetrics = new DocOpenTypeAdvances(kerned);
        Assert.True(kernedMetrics.TryMeasure("AB", 12, out var withoutKerning));
        Assert.True(kernedMetrics.TryMeasure("AB", 12, out var withKerning,
            useKerning: true));
        Assert.Equal(15, withoutKerning);
        Assert.Equal(14.4, withKerning, 3);
    }

    [Fact]
    public void OpenTypeAdvanceReaderMeasuresFittedWordFont()
    {
        var font = Path.Combine(Environment.GetFolderPath(
            Environment.SpecialFolder.Windows), "Fonts", "yumin.ttf");
        if (!File.Exists(font)) return;
        var metrics = Assert.IsType<DocOpenTypeAdvances>(
            DocSystemFontAdvances.Find("Yu Mincho"));
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        using var package = WordprocessingDocument.Open(Path.Combine(directory,
            "WordRunFitTextMixedCjkAllStories.docx"), false);
        var headerText = string.Concat(package.MainDocumentPart!.HeaderParts.Single()
            .Header!.Descendants<Text>().Select(x => x.Text));
        Assert.True(metrics.TryMeasure(headerText, 12, out var headerWidth));
        Assert.InRange(headerWidth, 405, 415);
        Assert.True(metrics.TryMeasure("Fitted text in the body.", 12,
            out var bodyWidth));
        Assert.InRange(bodyWidth, 119, 121);
        Assert.True(DocFitTextScaleEstimator.TryEstimate(1800, headerText,
            "Yu Mincho", 12, true, false, out var headerScale));
        Assert.Equal((ushort)20, headerScale);
        Assert.True(DocFitTextScaleEstimator.TryEstimate(2880,
            "Fitted text in the body.", "Yu Mincho", 12, false, false,
            out var bodyScale));
        Assert.Equal((ushort)120, bodyScale);
    }

    [Fact]
    public void NativeFitTextWidthBoundariesAreExplicit()
    {
        static byte[] Write(int width)
        {
            const string text = "Fit\r";
            using var output = new MemoryStream();
            DocPlainTextWriter.Write(output, new DocPlainTextDocument(
                new DocPlainTextStory(text,
                    [new DocStoryCharacterRun(0, 3,
                        DocCharacterFormatting.Empty with
                        { FitText = new DocFitText(width, 4) })], [])
                {
                    Paragraphs = [new DocStoryParagraphRange(0, text.Length,
                        DocParagraphEnd.ParagraphMark)]
                },
                [new DocPlainTextSection(text.Length, new DocPlainTextStory?[6])]));
            return output.ToArray();
        }
        var negative = Write(-1);
        using (var input = new MemoryStream(negative))
        using (var index = new DocTextIndexWalker().Index(input))
            Assert.Contains(index.CharacterFormatting,
                x => x.Formatting.FitText == new DocFitText(-1, 4));
        Assert.Throws<NotSupportedException>(() => DxpDocToDocx.Project(negative));
        var ignored = Write(0);
        var projected = DxpDocToDocx.Project(ignored);
        using var package = WordprocessingDocument.Open(
            new MemoryStream(projected.DocxBytes), false);
        Assert.Empty(package.MainDocumentPart!.Document.Body!.Descendants<FitText>());
        Assert.Empty(Validate(projected.DocxBytes));
    }

    [Fact]
    public void DocxFitTextWidthAboveSignedDocRangeIsRejected()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        using var source = new MemoryStream();
        source.Write(File.ReadAllBytes(Path.Combine(directory,
            "WordRunFitTextAllStories.docx")));
        source.Position = 0;
        using (var package = WordprocessingDocument.Open(source, true))
            package.MainDocumentPart!.Document.Body!.Descendants<FitText>()
                .First().Val = (uint)int.MaxValue + 1;
        Assert.Throws<NotSupportedException>(() => DxpDocExport.Export(source.ToArray()));
    }

    [Fact]
    public void RunFitTextParagraphStylesRetainInheritedWidthsInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordRunFitTextParagraphStyleAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var stories = ReadStories(source);
        foreach (var doc in new[] { File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
            DxpDocExport.Export(source) })
            foreach (var binary in new[] { doc,
                DxpDocExport.Export(DxpDocToDocx.Project(doc).DocxBytes) })
            {
                using (var input = new MemoryStream(binary))
                using (var index = new DocTextIndexWalker().Index(input))
                    foreach (var (styleName, fit) in new[] {
                        ("FitParagraphBody", new DocFitText(2880, 7)),
                        ("FitParagraphHeader", new DocFitText(1800, 8)),
                        ("FitParagraphFooter", new DocFitText(1800, 9)) })
                        Assert.Equal(fit, Assert.Single(index.StyleDefinitions,
                            x => x.Name == styleName).CharacterFormatting.FitText);
                var projection = DxpDocToDocx.Project(binary);
                Assert.Empty(projection.Coverage.OmittedCharacters);
                Assert.Equal(stories.Where(x => !x.Key.EndsWith(".customStyles")).ToArray(),
                    ReadStories(projection.DocxBytes).Where(x => !x.Key.EndsWith(".customStyles")).ToArray());
                Assert.Empty(Validate(projection.DocxBytes));
                using var package = WordprocessingDocument.Open(
                    new MemoryStream(projection.DocxBytes), false);
                var main = package.MainDocumentPart!;
                var stylePart = main.StyleDefinitionsPart!.Styles!;
                foreach (var (styleName, fit) in new[] {
                    ("FitParagraphBody", new DocFitText(2880, 7)),
                    ("FitParagraphHeader", new DocFitText(1800, 8)),
                    ("FitParagraphFooter", new DocFitText(1800, 9)) })
                {
                    var style = Assert.Single(stylePart.Elements<Style>(), x =>
                        x.StyleName?.Val?.Value == styleName);
                    Assert.Equal((uint)fit.WidthTwips,
                        style.StyleRunProperties!.GetFirstChild<FitText>()!.Val!.Value);
                    Assert.Equal(fit.Id,
                        style.StyleRunProperties.GetFirstChild<FitText>()!.Id!.Value);
                    var paragraph = styleName.EndsWith("Body")
                        ? main.Document.Body!.Elements<Paragraph>().First()
                        : styleName.EndsWith("Header")
                            ? main.HeaderParts.Single().Header!.Elements<Paragraph>().First()
                            : main.FooterParts.Single().Footer!.Elements<Paragraph>().First();
                    Assert.Equal(style.StyleId!.Value,
                        paragraph.ParagraphProperties!.ParagraphStyleId!.Val!.Value);
                }

            }
    }

    [Fact]
    public void RunFitTextBasedOnParagraphStylesRetainInheritedWidthsInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordRunFitTextInheritedParagraphStyleAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var stories = ReadStories(source);
        foreach (var doc in new[] { File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
            DxpDocExport.Export(source) })
            foreach (var binary in new[] { doc,
                DxpDocExport.Export(DxpDocToDocx.Project(doc).DocxBytes) })
            {
                using (var input = new MemoryStream(binary))
                using (var index = new DocTextIndexWalker().Index(input))
                    foreach (var (styleName, fit) in new[] {
                        ("FitParagraphBody", new DocFitText(2880, 7)),
                        ("FitParagraphHeader", new DocFitText(1800, 8)),
                        ("FitParagraphFooter", new DocFitText(1800, 9)) })
                        Assert.Equal(fit, Assert.Single(index.StyleDefinitions,
                            x => x.Name == styleName).CharacterFormatting.FitText);
                var projection = DxpDocToDocx.Project(binary);
                Assert.Empty(projection.Coverage.OmittedCharacters);
                Assert.Equal(stories.Where(x => !x.Key.EndsWith(".customStyles")).ToArray(),
                    ReadStories(projection.DocxBytes).Where(x => !x.Key.EndsWith(".customStyles")).ToArray());
                Assert.Empty(Validate(projection.DocxBytes));
                using var package = WordprocessingDocument.Open(
                    new MemoryStream(projection.DocxBytes), false);
                var main = package.MainDocumentPart!;
                var stylePart = main.StyleDefinitionsPart!.Styles!;
                foreach (var (styleName, fit) in new[] {
                    ("FitParagraphBody", new DocFitText(2880, 7)),
                    ("FitParagraphHeader", new DocFitText(1800, 8)),
                    ("FitParagraphFooter", new DocFitText(1800, 9)) })
                {
                    var style = Assert.Single(stylePart.Elements<Style>(), x =>
                        x.StyleName?.Val?.Value == styleName);
                    Assert.Equal((uint)fit.WidthTwips,
                        style.StyleRunProperties!.GetFirstChild<FitText>()!.Val!.Value);
                    Assert.Equal(fit.Id,
                        style.StyleRunProperties.GetFirstChild<FitText>()!.Id!.Value);
                    var derived = Assert.Single(stylePart.Elements<Style>(), x =>
                        x.StyleName?.Val?.Value == "Inherited" + styleName);
                    Assert.Equal(style.StyleId!.Value, derived.BasedOn!.Val!.Value);
                    var paragraph = styleName.EndsWith("Body")
                        ? main.Document.Body!.Elements<Paragraph>().First()
                        : styleName.EndsWith("Header")
                            ? main.HeaderParts.Single().Header!.Elements<Paragraph>().First()
                            : main.FooterParts.Single().Footer!.Elements<Paragraph>().First();
                    Assert.Equal(derived.StyleId!.Value,
                        paragraph.ParagraphProperties!.ParagraphStyleId!.Val!.Value);
                }

            }
    }

    [Fact]
    public void RunFitTextParagraphOverrideRetainsDirectAndInheritedWidthsInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordRunFitTextParagraphOverrideAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var stories = ReadStories(source);
        foreach (var doc in new[] { File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
            DxpDocExport.Export(source) })
            foreach (var binary in new[] { doc,
                DxpDocExport.Export(DxpDocToDocx.Project(doc).DocxBytes) })
            {
                using (var input = new MemoryStream(binary))
                using (var index = new DocTextIndexWalker().Index(input))
                    foreach (var (styleName, fit) in new[] {
                        ("FitParagraphBody", new DocFitText(2880, 7)),
                        ("FitParagraphHeader", new DocFitText(1800, 8)),
                        ("FitParagraphFooter", new DocFitText(1800, 9)) })
                        Assert.Equal(fit, Assert.Single(index.StyleDefinitions,
                            x => x.Name == styleName).CharacterFormatting.FitText);
                using (var input = new MemoryStream(binary))
                using (var index = new DocTextIndexWalker().Index(input))
                    Assert.Contains(index.CharacterFormatting, x =>
                        x.Formatting.FitText == new DocFitText(1200, 10));
                var projection = DxpDocToDocx.Project(binary);
                Assert.Empty(projection.Coverage.OmittedCharacters);
                Assert.Equal(stories.Where(x => !x.Key.EndsWith(".customStyles")).ToArray(),
                    ReadStories(projection.DocxBytes).Where(x => !x.Key.EndsWith(".customStyles")).ToArray());
                Assert.Empty(Validate(projection.DocxBytes));
                using var package = WordprocessingDocument.Open(
                    new MemoryStream(projection.DocxBytes), false);
                var main = package.MainDocumentPart!;
                var stylePart = main.StyleDefinitionsPart!.Styles!;
                Assert.Contains(main.Document.Body!.Elements<Paragraph>().First()
                    .Descendants<Run>(), x => x.RunProperties?.GetFirstChild<FitText>()
                    is { } direct && direct.Val?.Value == 1200 && direct.Id?.Value == 10);

                foreach (var (styleName, fit) in new[] {
                    ("FitParagraphBody", new DocFitText(2880, 7)),
                    ("FitParagraphHeader", new DocFitText(1800, 8)),
                    ("FitParagraphFooter", new DocFitText(1800, 9)) })
                {
                    var style = Assert.Single(stylePart.Elements<Style>(), x =>
                        x.StyleName?.Val?.Value == styleName);
                    Assert.Equal((uint)fit.WidthTwips,
                        style.StyleRunProperties!.GetFirstChild<FitText>()!.Val!.Value);
                    Assert.Equal(fit.Id,
                        style.StyleRunProperties.GetFirstChild<FitText>()!.Id!.Value);
                    var derived = Assert.Single(stylePart.Elements<Style>(), x =>
                        x.StyleName?.Val?.Value == "Inherited" + styleName);
                    Assert.Equal(style.StyleId!.Value, derived.BasedOn!.Val!.Value);
                    var paragraph = styleName.EndsWith("Body")
                        ? main.Document.Body!.Elements<Paragraph>().First()
                        : styleName.EndsWith("Header")
                            ? main.HeaderParts.Single().Header!.Elements<Paragraph>().First()
                            : main.FooterParts.Single().Footer!.Elements<Paragraph>().First();
                    Assert.Equal(derived.StyleId!.Value,
                        paragraph.ParagraphProperties!.ParagraphStyleId!.Val!.Value);
                }

            }
    }

    [Fact]
    public void RunFitTextHighAnsiFontsRetainScalesAndEditableFontsInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordRunFitTextHighAnsiFontsAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var native = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
        var generated = DxpDocExport.Export(source);
        static ushort?[] Scales(byte[] doc)
        {
            using var stream = new MemoryStream(doc);
            using var index = new DocTextIndexWalker().Index(stream);
            return new[] { 7, 8, 9 }.Select(id => index.CharacterFormatting
                .FirstOrDefault(x => x.Formatting.FitText?.Id == id)?
                .Formatting.CharacterScalePercent).ToArray();
        }
        var nativeScales = Scales(native);
        var generatedScales = Scales(generated);
        for (var i = 0; i < nativeScales.Length; i++)
            Assert.InRange(generatedScales[i]!.Value,
                nativeScales[i]!.Value - 1, nativeScales[i]!.Value + 1);
        foreach (var doc in new[] { native, generated })
            foreach (var binary in new[] { doc,
                DxpDocExport.Export(DxpDocToDocx.Project(doc).DocxBytes) })
            {
                var projection = DxpDocToDocx.Project(binary);
                Assert.Empty(projection.Coverage.OmittedCharacters);
                Assert.Equal(ReadStories(source), ReadStories(projection.DocxBytes));
                Assert.Empty(Validate(projection.DocxBytes));
                using var package = WordprocessingDocument.Open(
                    new MemoryStream(projection.DocxBytes), false);
                var main = package.MainDocumentPart!;
                foreach (var run in new[] {
                    main.Document.Body!.Elements<Paragraph>().First().Elements<Run>().First(),
                    main.HeaderParts.Single().Header!.Descendants<Run>().First(),
                    main.FooterParts.Single().Footer!.Descendants<Run>().First() })
                {
                    Assert.Equal("Calibri", run.RunProperties!.RunFonts!.Ascii!.Value);
                    Assert.Equal("Arial Black", run.RunProperties.RunFonts.HighAnsi!.Value);
                }

            }
    }

    [Fact]
    public void RunFitTextBoldUsesBoldFaceAdvancesInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordRunFitTextBoldAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var native = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
        var generated = DxpDocExport.Export(source);
        static ushort?[] Scales(byte[] doc)
        {
            using var stream = new MemoryStream(doc);
            using var index = new DocTextIndexWalker().Index(stream);
            return new[] { 7, 8, 9 }.Select(id => index.CharacterFormatting
                .FirstOrDefault(x => x.Formatting.FitText?.Id == id)?
                .Formatting.CharacterScalePercent).ToArray();
        }
        var nativeScales = Scales(native);
        var generatedScales = Scales(generated);
        for (var i = 0; i < nativeScales.Length; i++)
            Assert.InRange(generatedScales[i]!.Value,
                nativeScales[i]!.Value - 1, nativeScales[i]!.Value + 1);
        foreach (var doc in new[] { native, generated })
            foreach (var binary in new[] { doc,
                DxpDocExport.Export(DxpDocToDocx.Project(doc).DocxBytes) })
            {
                var projection = DxpDocToDocx.Project(binary);
                Assert.Empty(projection.Coverage.OmittedCharacters);
                Assert.Equal(ReadStories(source), ReadStories(projection.DocxBytes));
                Assert.Empty(Validate(projection.DocxBytes));
                using var package = WordprocessingDocument.Open(
                    new MemoryStream(projection.DocxBytes), false);
                var main = package.MainDocumentPart!;
                foreach (var run in new[] {
                    main.Document.Body!.Elements<Paragraph>().First().Elements<Run>().First(),
                    main.HeaderParts.Single().Header!.Descendants<Run>().First(),
                    main.FooterParts.Single().Footer!.Descendants<Run>().First() })
                {
                    Assert.Equal("Calibri", run.RunProperties!.RunFonts!.Ascii!.Value);
                    Assert.Equal("Calibri", run.RunProperties.RunFonts.HighAnsi!.Value);
                    Assert.NotNull(run.RunProperties.Bold);
                }

            }
    }

    [Theory]
    [InlineData("WordRunFitTextLayeredCalibriCjkAllStories", "Calibri")]
    [InlineData("WordRunFitTextLayeredArialCjkAllStories", "Arial")]
    [InlineData("WordRunFitTextHeaderCjkFallback", "Calibri")]
    [InlineData("WordRunFitTextFooterCjkFallback", "Calibri")]
    [InlineData("WordRunFitTextKoreanFallback", "Calibri")]
    [InlineData("WordRunFitTextChineseFallback", "Calibri")]
    [InlineData("WordRunFitTextSupplementaryHanFallback", "Calibri")]
    public void CjkFontSubstitutionRetainsStoriesAndThirdHop(string name,
        string authoredFont)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        static (bool? Balance, bool? Breaking) Flags(byte[] doc)
        {
            using var stream = new MemoryStream(doc);
            using var index = new DocTextIndexWalker().Index(stream);
            return (index.BalanceSingleByteDoubleByteWidth, index.ApplyBreakingRules);
        }
        var native = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
        var generated = DxpDocExport.Export(source);
        Assert.Equal(Flags(native), Flags(generated));
        var authoredFace = DocSystemFontAdvances.Find(authoredFont);
        var fallbackFace = DocSystemFontAdvances.Find("Yu Mincho");
        var sample = name.Contains("Korean", StringComparison.Ordinal) ? "한" :
            name.Contains("Supplementary", StringComparison.Ordinal) ? "𠀀" : "日";
        if (authoredFace != null && fallbackFace != null &&
            !authoredFace.TryMeasure(sample, 12, out _))
        {
            using var nativeStream = new MemoryStream(native);
            using var generatedStream = new MemoryStream(generated);
            using var nativeIndex = new DocTextIndexWalker().Index(nativeStream);
            using var generatedIndex = new DocTextIndexWalker().Index(generatedStream);
            Assert.Equal("Yu Mincho",
                nativeIndex.DefaultCharacterFormatting.EastAsiaFontName);
            Assert.Equal(nativeIndex.DefaultCharacterFormatting.EastAsiaFontName,
                generatedIndex.DefaultCharacterFormatting.EastAsiaFontName);
        }
        var stories = ReadStories(source);
        foreach (var doc in new[] { native, generated })
        {
            var projected = DxpDocToDocx.Project(doc);
            Assert.Empty(projected.Coverage.OmittedCharacters);
            Assert.Equal(stories, ReadStories(projected.DocxBytes));
            Assert.Empty(Validate(projected.DocxBytes));
            var thirdHop = DxpDocExport.Export(projected.DocxBytes);
            var thirdProjection = DxpDocToDocx.Project(thirdHop);
            Assert.Empty(thirdProjection.Coverage.OmittedCharacters);
            Assert.Equal(stories, ReadStories(thirdProjection.DocxBytes));
            Assert.Empty(Validate(thirdProjection.DocxBytes));
        }
    }

    [Fact]
    public void SoleAuthoredEastAsianFontStaysImplicitWhenNoCjkGlyphNeedsFallback()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextLayeredCalibriCjkAllStories.docx"));
        using var source = new MemoryStream();
        source.Write(File.ReadAllBytes(path));
        source.Position = 0;
        using (var document = WordprocessingDocument.Open(source, true))
        {
            var paragraph = document.MainDocumentPart!.Document.Body!
                .Elements<Paragraph>().Skip(1).First();
            Assert.Single(paragraph.Descendants<Text>()).Text =
                "Second paragraph has only Latin letters and digits.";
            document.MainDocumentPart.Document.Save();
        }
        var generated = DxpDocExport.Export(source.ToArray());
        using var stream = new MemoryStream(generated);
        using var index = new DocTextIndexWalker().Index(stream);
        Assert.Equal("Calibri", index.DefaultCharacterFormatting.EastAsiaFontName);
        var projected = DxpDocToDocx.Project(generated);
        Assert.Empty(projected.Coverage.OmittedCharacters);
        Assert.Empty(Validate(projected.DocxBytes));
        Assert.Equal(ReadStories(source.ToArray()), ReadStories(projected.DocxBytes));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EastAsianFallbackFindsCjkTextOnlyInHeaderOrFooter(bool inHeader)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextLayeredCalibriCjkAllStories.docx"));
        using var source = new MemoryStream();
        source.Write(File.ReadAllBytes(path));
        source.Position = 0;
        using (var document = WordprocessingDocument.Open(source, true))
        {
            Assert.Single(document.MainDocumentPart!.Document.Body!
                .Elements<Paragraph>().Skip(1).First().Descendants<Text>()).Text =
                "Second paragraph has only Latin letters and digits.";
            document.MainDocumentPart.Document.Save();
            if (inHeader)
            {
                var header = document.MainDocumentPart.HeaderParts.Single().Header!;
                header.Append(new Paragraph(new Run(new Text("日本語"))));
                header.Save();
            }
            else
            {
                var footer = document.MainDocumentPart.FooterParts.Single().Footer!;
                footer.Append(new Paragraph(new Run(new Text("日本語"))));
                footer.Save();
            }
        }
        var generated = DxpDocExport.Export(source.ToArray());
        using var stream = new MemoryStream(generated);
        using var index = new DocTextIndexWalker().Index(stream);
        var authoredFace = DocSystemFontAdvances.Find("Calibri");
        var fallbackFace = DocSystemFontAdvances.Find("Yu Mincho");
        var substitutes = authoredFace != null && fallbackFace != null &&
            !authoredFace.TryMeasure("日", 12, out _) &&
            fallbackFace.TryMeasure("日", 12, out _);
        Assert.Equal(substitutes ? "Yu Mincho" : "Calibri",
            index.DefaultCharacterFormatting.EastAsiaFontName);
        var projected = DxpDocToDocx.Project(generated);
        Assert.Empty(projected.Coverage.OmittedCharacters);
        Assert.Empty(Validate(projected.DocxBytes));
        Assert.Equal(ReadStories(source.ToArray()), ReadStories(projected.DocxBytes));
        var third = DxpDocToDocx.Project(DxpDocExport.Export(projected.DocxBytes));
        Assert.Empty(third.Coverage.OmittedCharacters);
        Assert.Equal(ReadStories(source.ToArray()), ReadStories(third.DocxBytes));
        Assert.Empty(Validate(third.DocxBytes));
    }

    [Fact]
    public void RunFitTextLayeredBoldItalicStylesRetainNativeSpacing()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordRunFitTextLayeredBoldItalicAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var native = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
        var generated = DxpDocExport.Export(source);
        static ushort?[] Scales(byte[] doc)
        {
            using var stream = new MemoryStream(doc);
            using var index = new DocTextIndexWalker().Index(stream);
            return new[] { 7, 8, 9 }.Select(id => index.CharacterFormatting
                .FirstOrDefault(x => x.Formatting.FitText?.Id == id)?
                .Formatting.CharacterScalePercent).ToArray();
        }
        var nativeScales = Scales(native);
        var generatedScales = Scales(generated);
        static bool HasFit(byte[] doc)
        {
            using var stream = new MemoryStream(doc);
            using var index = new DocTextIndexWalker().Index(stream);
            return index.CharacterFormatting.Any(x => x.Formatting.FitText != null);
        }
        Assert.True(HasFit(native));
        Assert.True(HasFit(generated));
        Assert.All(nativeScales, Assert.Null);
        Assert.All(generatedScales, Assert.Null);
        static short?[] Spacing(byte[] doc)
        {
            using var stream = new MemoryStream(doc);
            using var index = new DocTextIndexWalker().Index(stream);
            return new[] { 7, 8, 9 }.Select(id => index.CharacterFormatting
                .FirstOrDefault(x => x.Formatting.FitText?.Id == id)?
                .Formatting.CharacterSpacingTwips).ToArray();
        }
        var nativeSpacing = Spacing(native);
        var generatedSpacing = Spacing(generated);
        Assert.All(generatedSpacing, value => Assert.NotNull(value));
        for (var i = 0; i < nativeSpacing.Length; i++)
            Assert.InRange(generatedSpacing[i]!.Value,
                nativeSpacing[i]!.Value - 2, nativeSpacing[i]!.Value + 2);
        foreach (var doc in new[] { native, generated })
            foreach (var binary in new[] { doc,
                DxpDocExport.Export(DxpDocToDocx.Project(doc).DocxBytes) })
            {
                var projection = DxpDocToDocx.Project(binary);
                Assert.Empty(projection.Coverage.OmittedCharacters);
                Assert.Equal(ReadStories(source), ReadStories(projection.DocxBytes));
                Assert.Empty(Validate(projection.DocxBytes));

            }
    }

    [Fact]
    public void RunFitTextItalicUsesItalicFaceAdvancesInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordRunFitTextItalicAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var native = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
        var generated = DxpDocExport.Export(source);
        static ushort?[] Scales(byte[] doc)
        {
            using var stream = new MemoryStream(doc);
            using var index = new DocTextIndexWalker().Index(stream);
            return new[] { 7, 8, 9 }.Select(id => index.CharacterFormatting
                .FirstOrDefault(x => x.Formatting.FitText?.Id == id)?
                .Formatting.CharacterScalePercent).ToArray();
        }
        var nativeScales = Scales(native);
        var generatedScales = Scales(generated);
        for (var i = 0; i < nativeScales.Length; i++)
            Assert.InRange(generatedScales[i]!.Value,
                nativeScales[i]!.Value - 1, nativeScales[i]!.Value + 1);
        foreach (var doc in new[] { native, generated })
            foreach (var binary in new[] { doc,
                DxpDocExport.Export(DxpDocToDocx.Project(doc).DocxBytes) })
            {
                var projection = DxpDocToDocx.Project(binary);
                Assert.Empty(projection.Coverage.OmittedCharacters);
                Assert.Equal(ReadStories(source), ReadStories(projection.DocxBytes));
                Assert.Empty(Validate(projection.DocxBytes));
                using var package = WordprocessingDocument.Open(
                    new MemoryStream(projection.DocxBytes), false);
                var main = package.MainDocumentPart!;
                foreach (var run in new[] {
                    main.Document.Body!.Elements<Paragraph>().First().Elements<Run>().First(),
                    main.HeaderParts.Single().Header!.Descendants<Run>().First(),
                    main.FooterParts.Single().Footer!.Descendants<Run>().First() })
                {
                    Assert.Equal("Calibri", run.RunProperties!.RunFonts!.Ascii!.Value);
                    Assert.Equal("Calibri", run.RunProperties.RunFonts.HighAnsi!.Value);
                    Assert.NotNull(run.RunProperties.Italic);
                }

            }
    }

    [Fact]
    public void RunFitTextCjkPunctuationUsesEastAsianFaceInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordRunFitTextCjkPunctuationFontsAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var native = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
        var generated = DxpDocExport.Export(source);
        static ushort?[] Scales(byte[] doc)
        {
            using var stream = new MemoryStream(doc);
            using var index = new DocTextIndexWalker().Index(stream);
            return new[] { 7, 8, 9 }.Select(id => index.CharacterFormatting
                .FirstOrDefault(x => x.Formatting.FitText?.Id == id)?
                .Formatting.CharacterScalePercent).ToArray();
        }
        var nativeScales = Scales(native);
        var generatedScales = Scales(generated);
        Assert.True(nativeScales.All(x => x != null) && generatedScales.All(x => x != null),
            $"native={string.Join(",", nativeScales)} generated={string.Join(",", generatedScales)}");
        for (var i = 0; i < nativeScales.Length; i++)
            Assert.InRange(generatedScales[i]!.Value,
                nativeScales[i]!.Value - 1, nativeScales[i]!.Value + 1);
        foreach (var doc in new[] { native, generated })
            foreach (var binary in new[] { doc,
                DxpDocExport.Export(DxpDocToDocx.Project(doc).DocxBytes) })
            {
                var projection = DxpDocToDocx.Project(binary);
                Assert.Empty(projection.Coverage.OmittedCharacters);
                Assert.Equal(ReadStories(source), ReadStories(projection.DocxBytes));
                Assert.Empty(Validate(projection.DocxBytes));
                using var package = WordprocessingDocument.Open(
                    new MemoryStream(projection.DocxBytes), false);
                var main = package.MainDocumentPart!;
                foreach (var storyRuns in new[] {
                    main.Document.Body!.Elements<Paragraph>().First().Descendants<Run>(),
                    main.HeaderParts.Single().Header!.Descendants<Run>(),
                    main.FooterParts.Single().Footer!.Descendants<Run>() })
                {
                    Assert.True(storyRuns.Any(run =>
                        run.RunProperties?.RunFonts?.Ascii?.Value == "Calibri" &&
                        run.RunProperties.RunFonts.HighAnsi?.Value == "Arial Black"),
                        string.Join(";", storyRuns.Select(run =>
                            $"{run.RunProperties?.RunFonts?.Ascii?.Value}/{run.RunProperties?.RunFonts?.HighAnsi?.Value}/{run.RunProperties?.RunFonts?.EastAsia?.Value}")));
                    if (ReferenceEquals(doc, generated) && ReferenceEquals(binary, generated))
                        Assert.Contains(storyRuns, run =>
                            run.RunProperties?.RunFonts?.EastAsia?.Value == "Yu Mincho");
                }
            }
    }

    [Fact]
    public void RunFitTextPairRetainsGroupedWidthsInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordRunFitTextAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var stories = ReadStories(source);
        var expected = new[] { (2880, 7), (1800, 8), (1800, 9) };
        foreach (var doc in new[] { File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
            DxpDocExport.Export(source) })
        {
            foreach (var binary in new[] { doc,
                DxpDocExport.Export(DxpDocToDocx.Project(doc).DocxBytes) })
            {
                using (var input = new MemoryStream(binary))
                using (var index = new DocTextIndexWalker().Index(input))
                {
                    var actual = index.CharacterFormatting
                        .Where(x => x.Formatting.FitText != null)
                        .Select(x => (x.Formatting.FitText!.WidthTwips,
                            x.Formatting.FitText.Id)).Distinct().ToArray();
                    Assert.Equal(expected, actual);
                }
                var projection = DxpDocToDocx.Project(binary);
                Assert.Empty(projection.Coverage.OmittedCharacters);
                Assert.Equal(stories, ReadStories(projection.DocxBytes));
                using var package = WordprocessingDocument.Open(
                    new MemoryStream(projection.DocxBytes), false);
                var main = package.MainDocumentPart!;
                var body = main.Document.Body!.Elements<Paragraph>().First()
                    .Descendants<Run>().Select(x => x.RunProperties?.GetFirstChild<FitText>())
                    .Where(x => x != null).ToArray();
                Assert.Equal(2, body.Length);
                Assert.All(body, x => { Assert.Equal((uint)2880, x!.Val!.Value);
                    Assert.Equal(7, x.Id!.Value); });
                var header = main.HeaderParts.Single().Header!.Descendants<Run>()
                    .Select(x => x.RunProperties?.GetFirstChild<FitText>())
                    .First(x => x?.Id?.Value == 8)!;
                var footer = main.FooterParts.Single().Footer!.Descendants<Run>()
                    .Select(x => x.RunProperties?.GetFirstChild<FitText>())
                    .First(x => x?.Id?.Value == 9)!;
                Assert.Equal((uint)1800, header.Val!.Value);
                Assert.Equal(8, header.Id!.Value);
                Assert.Equal((uint)1800, footer.Val!.Value);
                Assert.Equal(9, footer.Id!.Value);
                Assert.Empty(Validate(projection.DocxBytes));
            }
        }
    }
    [Fact]
    public void RunFitTextCharacterStyleCombinesWithDirectRunInBody()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordRunFitTextStyleAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var stories = ReadStories(source);
        foreach (var doc in new[] { File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
            DxpDocExport.Export(source) })
        {
            foreach (var binary in new[] { doc,
                DxpDocExport.Export(DxpDocToDocx.Project(doc).DocxBytes) })
            {
                using (var input = new MemoryStream(binary))
                using (var index = new DocTextIndexWalker().Index(input))
                {
                    Assert.Equal(new DocFitText(2880, 7), Assert.Single(
                        index.StyleDefinitions, x => x.Name == "Fit text character")
                        .CharacterFormatting.FitText);
                    Assert.Contains(index.CharacterFormatting,
                        x => x.Formatting.FitText == new DocFitText(2880, 7));
                }
                var projected = DxpDocToDocx.Project(binary);
                Assert.Empty(projected.Coverage.OmittedCharacters);
                Assert.Equal(stories, ReadStories(projected.DocxBytes));
                using var package = WordprocessingDocument.Open(
                    new MemoryStream(projected.DocxBytes), false);
                var style = package.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                    .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                        "Fit text character");
                Assert.Equal((uint)2880, style.StyleRunProperties!
                    .GetFirstChild<FitText>()!.Val!.Value);
                Assert.Empty(Validate(projected.DocxBytes));
            }
        }
    }
    [Fact]
    public void MixedScriptFontFitTextRetainsGroupedRunsAcrossStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordRunFitTextMixedScriptFontsAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var generated = DxpDocExport.Export(source);
        var native = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
        var stories = ReadStories(source);
        static (ushort? Header, ushort? Footer) Scales(byte[] doc)
        {
            using var stream = new MemoryStream(doc);
            using var index = new DocTextIndexWalker().Index(stream);
            return (index.CharacterFormatting.FirstOrDefault(x =>
                x.Formatting.FitText?.Id == 8)?.Formatting.CharacterScalePercent,
                index.CharacterFormatting.FirstOrDefault(x =>
                x.Formatting.FitText?.Id == 9)?.Formatting.CharacterScalePercent);
        }
        var generatedScales = Scales(generated);
        var nativeScales = Scales(native);
        foreach (var doc in new[] { native, generated })
            foreach (var binary in new[] { doc,
                DxpDocExport.Export(DxpDocToDocx.Project(doc).DocxBytes) })
            {
                using (var input = new MemoryStream(binary))
                using (var index = new DocTextIndexWalker().Index(input))
                {
                    Assert.Contains(index.CharacterFormatting, x =>
                        x.Formatting.FitText?.Id == 8);
                    Assert.Contains(index.CharacterFormatting, x =>
                        x.Formatting.FitText?.Id == 9);
                    Assert.Contains(index.Fonts, x => x.Name == "Calibri");
                    Assert.Contains(index.Fonts, x => x.Name == "Yu Mincho");
                }
                var projection = DxpDocToDocx.Project(binary);
                Assert.Empty(projection.Coverage.OmittedCharacters);
                Assert.Equal(stories, ReadStories(projection.DocxBytes));
                Assert.Empty(Validate(projection.DocxBytes));
            }
        if (DocSystemFontAdvances.Find("Calibri") != null &&
            DocSystemFontAdvances.Find("Yu Mincho") != null)
        {
            Assert.Equal(nativeScales.Header, generatedScales.Header);
            Assert.Equal(nativeScales.Footer, generatedScales.Footer);
        }
        if (Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_WORD") != "1") return;
        var path = Path.Combine(Path.GetTempPath(),
            "docxport-mixed-script-font-fit-" + Guid.NewGuid().ToString("N") + ".doc");
        File.WriteAllBytes(path, generated);
        dynamic word = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application")!)!;
        dynamic? opened = null;
        try
        {
            word.Visible = false;
            word.DisplayAlerts = 0;
            (double Middle, double End) Positions(dynamic range)
            {
                dynamic chars = range.Characters;
                return ((double)chars.Item(20).Information(5),
                    (double)chars.Item(chars.Count).Information(5));
            }
            ((double Middle, double End) Header, (double Middle, double End) Footer)
                Read(string file)
            {
                opened = word.Documents.Open(file, false, true);
                var positions = (Positions(opened.Sections.Item(1).Headers.Item(1).Range),
                    Positions(opened.Sections.Item(1).Footers.Item(1).Range));
                opened.Close(false);
                opened = null;
                return positions;
            }
            var expected = Read(Path.Combine(directory, name + ".docx"));
            var actual = Read(path);
            Assert.True(Math.Abs(expected.Header.End - actual.Header.End) <= 2.5 &&
                Math.Abs(expected.Footer.End - actual.Footer.End) <= 2.5,
                $"expected={expected}, actual={actual}, nativeScale={nativeScales}, " +
                $"generatedScale={generatedScales}");
        }
        finally
        {
            if (opened != null) opened.Close(false);
            word.Quit();
            File.Delete(path);
        }
    }

    [Fact]
    public void MixedFontFitTextRegionUsesBothInstalledFaces()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordRunFitTextMixedFontsAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var stories = ReadStories(source);
        foreach (var doc in new[] { File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
            DxpDocExport.Export(source) })
        {
            foreach (var binary in new[] { doc,
                DxpDocExport.Export(DxpDocToDocx.Project(doc).DocxBytes) })
            {
                using (var input = new MemoryStream(binary))
                using (var index = new DocTextIndexWalker().Index(input))
                {
                    var body = index.CharacterFormatting.Where(x => x.CpStart < 24 &&
                        x.Formatting.FitText == new DocFitText(2880, 7)).ToArray();
                    Assert.True(body.Length >= 2);
                    Assert.Contains(body, x => x.Formatting.AsciiFontName == "Calibri");
                    if (DocSystemFontAdvances.Find("Calibri") != null &&
                        DocSystemFontAdvances.Find("Yu Mincho") != null)
                        Assert.Single(body.Select(x => x.Formatting.CharacterScalePercent)
                            .Distinct());
                }
                var projected = DxpDocToDocx.Project(binary);
                Assert.Empty(projected.Coverage.OmittedCharacters);
                Assert.Equal(stories, ReadStories(projected.DocxBytes));
                Assert.Empty(Validate(projected.DocxBytes));
            }
        }
        if (Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_WORD") != "1") return;
        var path = Path.Combine(Path.GetTempPath(),
            "docxport-mixed-font-fit-" + Guid.NewGuid().ToString("N") + ".doc");
        File.WriteAllBytes(path, DxpDocExport.Export(source));
        dynamic word = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application")!)!;
        dynamic? opened = null;
        try
        {
            word.Visible = false;
            word.DisplayAlerts = 0;
            (double Middle, double End) Positions(string file)
            {
                opened = word.Documents.Open(file, false, true);
                dynamic characters = opened.Paragraphs.Item(1).Range.Characters;
                double middle = characters.Item(20).Information(5);
                double final = characters.Item(characters.Count).Information(5);
                opened.Close(false);
                opened = null;
                return (middle, final);
            }
            var expected = Positions(Path.Combine(directory, name + ".docx"));
            var actual = Positions(path);
            Assert.True(Math.Abs(expected.Middle - actual.Middle) <= 4 &&
                Math.Abs(expected.End - actual.End) <= 2,
                $"body positions expected={expected}, actual={actual}");
        }
        finally
        {
            if (opened != null) opened.Close(false);
            word.Quit();
            File.Delete(path);
        }
    }

    [Fact]
    public void MixedScriptFitTextUsesMeasuredScaleInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordRunFitTextMixedCjkAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var native = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
        var generated = DxpDocExport.Export(source);
        static (int FitRuns, ushort? Scale) ReadHeader(byte[] doc)
        {
            using var input = new MemoryStream(doc);
            using var index = new DocTextIndexWalker().Index(input);
            var fitted = index.CharacterFormatting.Where(x =>
                x.Formatting.FitText == new DocFitText(1800, 8)).ToArray();
            return (fitted.Length, fitted.Select(x =>
                x.Formatting.CharacterScalePercent).FirstOrDefault(x => x != null));
        }
        var nativeHeader = ReadHeader(native);
        var generatedHeader = ReadHeader(generated);
        Assert.True(nativeHeader.FitRuns > 0);
        Assert.True(generatedHeader.FitRuns > 0);
        Assert.Equal((ushort)20, nativeHeader.Scale);
        if (DocSystemFontAdvances.Find("Yu Mincho") != null)
            Assert.Equal((ushort)20, generatedHeader.Scale);
        foreach (var doc in new[] { native, generated })
        {
            var projected = DxpDocToDocx.Project(doc);
            Assert.Empty(projected.Coverage.OmittedCharacters);
            Assert.Equal(ReadStories(source), ReadStories(projected.DocxBytes));
            Assert.Empty(Validate(projected.DocxBytes));
        }
        if (Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_WORD") != "1") return;
        var generatedPath = Path.Combine(Path.GetTempPath(),
            "docxport-mixed-fit-" + Guid.NewGuid().ToString("N") + ".doc");
        File.WriteAllBytes(generatedPath, generated);
        dynamic word = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application")!)!;
        dynamic? opened = null;
        try
        {
            word.Visible = false;
            word.DisplayAlerts = 0;
            double Position(string path)
            {
                opened = word.Documents.Open(path, false, true);
                double position = opened.Sections.Item(1).Headers.Item(1)
                    .Range.Characters.Item(20).Information(5);
                opened.Close(false);
                opened = null;
                return position;
            }
            var sourcePosition = Position(Path.Combine(directory, name + ".docx"));
            var nativePosition = Position(Path.Combine(directory, name + ".doc"));
            var generatedPosition = Position(generatedPath);
            Assert.InRange(Math.Abs(sourcePosition - nativePosition), 0, 0.5);
            Assert.InRange(Math.Abs(generatedPosition - sourcePosition), 0, 2);
        }
        finally
        {
            if (opened != null) opened.Close(false);
            word.Quit();
            File.Delete(generatedPath);
        }
    }
    [Fact]
    public void WordNativeDocDropsOmittedFitTextIdsInHeaderAndFooter()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "DocKnownGaps"));
        const string name = "WordRunFitTextOmittedIdAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var native = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
        var generated = DxpDocExport.Export(source);
        static DocFitText[] FitTexts(byte[] doc)
        {
            using var input = new MemoryStream(doc);
            using var index = new DocTextIndexWalker().Index(input);
            return index.CharacterFormatting.Select(x => x.Formatting.FitText)
                .OfType<DocFitText>().Distinct().ToArray();
        }
        Assert.Equal(new[] { new DocFitText(2880, 7) }, FitTexts(native));
        Assert.Equal(new[] { new DocFitText(2880, 7), new DocFitText(1800, 0) },
            FitTexts(generated));
        var projected = DxpDocToDocx.Project(generated);
        Assert.Equal(ReadStories(source), ReadStories(projected.DocxBytes));
        Assert.Empty(Validate(projected.DocxBytes));
    }
    [Fact]
    public void OutlineStylePairRetainsLevelsAndOverridesInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordOutlineStylesAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        foreach (var doc in new[]
        {
            File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
            DxpDocExport.Export(source)
        })
        {
            foreach (var binary in new[] { doc,
                DxpDocExport.Export(DxpDocToDocx.Project(doc).DocxBytes) })
            {
                using (var input = new MemoryStream(binary))
                using (var index = new DocTextIndexWalker().Index(input))
                {
                    Assert.Equal((byte)0, Assert.Single(index.StyleDefinitions,
                        x => x.Name == "Outline pair base").ParagraphFormatting?.OutlineLevel);
                    Assert.Equal((byte)2, Assert.Single(index.StyleDefinitions,
                        x => x.Name == "Outline pair derived").ParagraphFormatting?.OutlineLevel);
                    Assert.Contains(index.ParagraphStyles, x => x.Formatting?.OutlineLevel == 1);
                    Assert.Contains(index.ParagraphStyles, x => x.Formatting?.OutlineLevel == 9);
                }
                var projected = DxpDocToDocx.Project(binary).DocxBytes;
                using var package = WordprocessingDocument.Open(
                    new MemoryStream(projected), false);
                var main = package.MainDocumentPart!;
                var styles = main.StyleDefinitionsPart!.Styles!;
                Assert.Equal(0, styles.Elements<Style>().Single(x =>
                    x.StyleName?.Val?.Value == "Outline pair base")
                    .StyleParagraphProperties!.OutlineLevel!.Val!.Value);
                Assert.Equal(2, styles.Elements<Style>().Single(x =>
                    x.StyleName?.Val?.Value == "Outline pair derived")
                    .StyleParagraphProperties!.OutlineLevel!.Val!.Value);
                var body = main.Document.Body!.Elements<Paragraph>().ToArray();
                Assert.Equal(1, body[0].ParagraphProperties!.OutlineLevel!.Val!.Value);
                Assert.Null(body[1].ParagraphProperties!.OutlineLevel);
                Assert.Null(main.HeaderParts.Single(x => x.Header!.InnerText.Contains("Time"))
                    .Header!.Elements<Paragraph>().Single().ParagraphProperties!.OutlineLevel);
                Assert.Equal(9, main.FooterParts.Single(x => x.Footer!.InnerText.Contains("Page"))
                    .Footer!.Elements<Paragraph>().Single().ParagraphProperties!
                    .OutlineLevel!.Val!.Value);
                Assert.Empty(Validate(projected));
            }
        }
    }

    [Fact]
    public void RevisionPropertyPairPreservesEditableFieldsAndMetadata()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordRevisionPropertyAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var native = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
        var expectedFields = ReadEditableFieldSemantics(source);
        Assert.Equal(3, expectedFields.Count(x =>
            x == "FIELD:DOCPROPERTY:RevisionNumber"));
        foreach (var binary in new[] { native, DxpDocExport.Export(source) })
        {
            using var input = new MemoryStream(binary);
            using var index = new DocTextIndexWalker().Index(input);
            var nativeRevision = DocSummaryInformation.ReadRevisionNumber(index.Structure);
            Assert.Equal("2", nativeRevision);
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            using var package = WordprocessingDocument.Open(new MemoryStream(projected), false);
            Assert.Equal(nativeRevision, package.PackageProperties.Revision);
            Assert.Equal(expectedFields, ReadEditableFieldSemantics(projected));
            Assert.Empty(Validate(projected));
        }
    }

    [Fact]
    public void WordSymbolGlyphsKeepFontAndCodeInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordSymbolGlyphAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var native = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
        var generated = DxpDocExport.Export(source);
        foreach (var binary in new[] { native, generated,
            DxpDocExport.Export(DxpDocToDocx.Project(generated).DocxBytes) })
        {
            using var stream = new MemoryStream(binary);
            using var index = new DocTextIndexWalker().Index(stream);
            var symbols = index.CharacterFormatting.Where(x =>
                x.Formatting.SymbolCharacter == 0xF0FC &&
                x.Formatting.SymbolFontName == "Wingdings").ToArray();
            Assert.Equal(3, symbols.Length);
            Assert.All(symbols, x => Assert.Equal((uint)1, x.CpEnd - x.CpStart));
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            var stories = ReadStories(projected);
            Assert.Contains("[symbol:Wingdings:F0FC]", stories["body"]);
            Assert.Contains(stories, x => x.Key.Contains(".header.") &&
                x.Value.Contains("[symbol:Wingdings:F0FC]"));
            Assert.Contains(stories, x => x.Key.Contains(".footer.") &&
                x.Value.Contains("[symbol:Wingdings:F0FC]"));
            Assert.Empty(Validate(projected));
        }
    }

    [Fact]
    public void WordSavedGrowAutofitPairRetainsSettingsAndStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures",
            "Doc"));
        const string name = "WordGrowAutofitNoWrapAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        static bool Enabled(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes);
            using var document = WordprocessingDocument.Open(stream, false);
            var option = document.MainDocumentPart!.DocumentSettingsPart?.Settings?
                .GetFirstChild<Compatibility>()?.GetFirstChild<GrowAutofit>();
            return option != null && (option.Val?.Value ?? true);
        }
        Assert.True(Enabled(source));
        foreach (var binary in new[]
        {
            File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
            DxpDocExport.Export(source)
        })
        {
            using (var stream = new MemoryStream(binary))
            using (var index = new DocTextIndexWalker().Index(stream))
                Assert.True(index.GrowAutofit);
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            Assert.True(Enabled(projected));
            Assert.Equal(ReadStories(source), ReadStories(projected));
            AssertTableGrid(source, projected);
            var repeated = DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes;
            Assert.True(Enabled(repeated));
            Assert.Equal(ReadStories(source), ReadStories(repeated));
            AssertTableGrid(source, repeated);
            Assert.Empty(Validate(projected));
            Assert.Empty(Validate(repeated));
        }
    }

    [Fact]
    public void GrowAutofitCompatibilityOptionSurvivesDocRoundTrips()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleNoWrapAllStories.docx"));
        using var source = new MemoryStream();
        source.Write(File.ReadAllBytes(path));
        using (var document = WordprocessingDocument.Open(source, true))
        {
            var settings = document.MainDocumentPart!.DocumentSettingsPart!.Settings!;
            var compatibility = settings.GetFirstChild<Compatibility>();
            if (compatibility == null)
            {
                compatibility = new Compatibility();
                settings.AddChild(compatibility, true);
            }
            compatibility.AddChild(new GrowAutofit(), true);
            settings.Save();
        }
        var generated = DxpDocExport.Export(source.ToArray());
        using (var stream = new MemoryStream(generated))
        using (var index = new DocTextIndexWalker().Index(stream))
            Assert.True(index.GrowAutofit);
        if (Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_WORD_RENDER") == "1")
        {
            var tempPath = Path.Combine(Path.GetTempPath(),
                $"docxport-grow-autofit-{Guid.NewGuid():N}.doc");
            File.WriteAllBytes(tempPath, generated);
            dynamic word = Activator.CreateInstance(Type.GetTypeFromProgID(
                "Word.Application")!)!;
            word.Visible = false;
            try
            {
                dynamic opened = word.Documents.Open(tempPath, ReadOnly: true,
                    AddToRecentFiles: false);
                try { Assert.True((bool)opened.Compatibility(50)); }
                finally { opened.Close(false); }
            }
            finally
            {
                word.Quit(false);
                File.Delete(tempPath);
            }
        }
        for (var hop = 0; hop < 2; hop++)
        {
            var projected = DxpDocToDocx.Project(generated).DocxBytes;
            using var stream = new MemoryStream(projected);
            using var document = WordprocessingDocument.Open(stream, false);
            Assert.NotNull(document.MainDocumentPart!.DocumentSettingsPart!
                .Settings!.GetFirstChild<Compatibility>()?.GetFirstChild<GrowAutofit>());
            generated = DxpDocExport.Export(projected);
        }
    }

    [Fact]
    public void WordClearedInheritedTabsDoNotReviveTheParentStop()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var name = "WordClearedInheritedTabsAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var expected = ReadEffectiveTabStops(source);
        Assert.Contains(expected, x => x.Key.StartsWith("body.paragraph0.",
            StringComparison.Ordinal) && x.Value == "right|hyphen");
        Assert.DoesNotContain(expected, x => x.Key == "body.paragraph0.tab5600");
        Assert.Contains(expected, x => x.Key.StartsWith("body.paragraph1.",
            StringComparison.Ordinal) && x.Value == "right|dot");
        Assert.Contains(expected, x => x.Key.Contains(".header.",
            StringComparison.Ordinal) && x.Value == "right|hyphen");
        Assert.Contains(expected, x => x.Key.Contains(".footer.",
            StringComparison.Ordinal) && x.Value == "right|dot");
        foreach (var binary in new[]
        {
            File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
            DxpDocExport.Export(source)
        })
        {
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            Assert.Equal(expected, ReadEffectiveTabStops(projected));
            var repeated = DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes;
            Assert.Equal(expected, ReadEffectiveTabStops(repeated));
        }
    }

    [Fact]
    public void WordInheritedTabLeadersExerciseStyleChainAndDirectOverrides()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedTabLeadersAllStories.docx"));
        var bytes = File.ReadAllBytes(path);
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var styles = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        var root = Assert.Single(styles.Elements<Style>(), x =>
            x.StyleId?.Value == "LeaderRoot");
        var child = Assert.Single(styles.Elements<Style>(), x =>
            x.StyleId?.Value == "LeaderChild");
        Assert.Equal("Normal", root.BasedOn?.Val?.Value);
        Assert.Equal("LeaderRoot", child.BasedOn?.Val?.Value);
        var tabs = ReadEffectiveTabStops(bytes);
        Assert.Contains(tabs, x => x.Key.StartsWith("body.", StringComparison.Ordinal) &&
            x.Value == "right|hyphen");
        Assert.Contains(tabs, x => x.Key.StartsWith("body.", StringComparison.Ordinal) &&
            x.Value == "right|underscore");
        Assert.Contains(tabs, x => x.Key.Contains(".header.", StringComparison.Ordinal) &&
            x.Value == "right|dot");
        Assert.Contains(tabs, x => x.Key.Contains(".footer.", StringComparison.Ordinal) &&
            x.Value == "right|heavy");
    }

    [Theory]
    [InlineData("WordTabLeadersAllStories")]
    [InlineData("WordStyledTabLeadersAllStories")]
    public void WordTabLeaderFixtureExercisesAllThreeStories(string name)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            name + ".docx"));
        var tabs = ReadEffectiveTabStops(File.ReadAllBytes(path));
        Assert.Contains(tabs, x => x.Key.StartsWith("body.", StringComparison.Ordinal) &&
            x.Value == "right|dot");
        Assert.Contains(tabs, x => x.Key.StartsWith("body.", StringComparison.Ordinal) &&
            x.Value == "right|hyphen");
        Assert.Contains(tabs, x => x.Key.Contains(".header.", StringComparison.Ordinal) &&
            x.Value == "right|underscore");
        Assert.Contains(tabs, x => x.Key.Contains(".footer.", StringComparison.Ordinal) &&
            x.Value == "right|heavy");
    }

    [Theory]
    [InlineData("WordNumWordsBodyWithStories")]
    [InlineData("WordNumWordsCharsMixedStories")]
    public void WordStatisticPropertiesSurviveBothDocRoutesAndThirdHop(string name)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        static (string?, string?, string?) Stats(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes);
            using var document = WordprocessingDocument.Open(stream, false);
            var properties = document.ExtendedFilePropertiesPart?.Properties;
            return (properties?.Pages?.Text, properties?.Words?.Text,
                properties?.Characters?.Text);
        }
        var expected = Stats(source);
        Assert.False(string.IsNullOrWhiteSpace(expected.Item2));
        Assert.False(string.IsNullOrWhiteSpace(expected.Item3));
        foreach (var (binary, savedStats) in new[]
        {
            (File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
                name == "WordNumWordsCharsMixedStories"
                    ? ("1", "12", "67") : expected),
            (DxpDocExport.Export(source), expected)
        })
        {
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            Assert.Equal(savedStats, Stats(projected));
            var repeated = DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes;
            Assert.Equal(savedStats, Stats(repeated));
        }
    }

    [Fact]
    public void OmittedTableLookUsesWordDefaultsInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        using var source = new MemoryStream();
        source.Write(File.ReadAllBytes(Path.Combine(directory,
            "WordConditionalFirstRowParagraphAllStories.docx")));
        using (var document = WordprocessingDocument.Open(source, true))
        {
            var main = document.MainDocumentPart!;
            var style = Assert.Single(main.StyleDefinitionsPart!.Styles!
                .Elements<Style>(), x => x.StyleId?.Value == "LogicalStart");
            style.StyleTableProperties!.PrependChild(
                new TableStyleColumnBandSize { Val = 1 });
            style.StyleTableProperties.PrependChild(
                new TableStyleRowBandSize { Val = 1 });
            style.AppendChild(new TableStyleProperties(
                new TableStyleConditionalFormattingTableCellProperties(
                    new Shading { Val = ShadingPatternValues.Clear,
                        Fill = "FFFF00" }))
            { Type = TableStyleOverrideValues.Band1Horizontal });
            style.AppendChild(new TableStyleProperties(
                new TableStyleConditionalFormattingTableCellProperties(
                    new Shading { Val = ShadingPatternValues.Clear,
                        Fill = "0000FF" }))
            { Type = TableStyleOverrideValues.Band1Vertical });
            main.StyleDefinitionsPart.Styles.Save();
            foreach (var table in main.Document.Body!.Elements<Table>()
                .Concat(main.HeaderParts.SelectMany(x => x.Header!.Elements<Table>()))
                .Concat(main.FooterParts.SelectMany(x => x.Footer!.Elements<Table>())))
            {
                table.TableProperties!.GetFirstChild<TableLook>()?.Remove();
                if (table.Elements<TableRow>().Count() == 1)
                    table.AppendChild(new TableRow(
                        new TableCell(new Paragraph(new Run(new Text("Band A")))),
                        new TableCell(new Paragraph(new Run(new Text("Band B"))))));
            }
            main.Document.Save();
            foreach (var part in main.HeaderParts) part.Header!.Save();
            foreach (var part in main.FooterParts) part.Footer!.Save();
        }
        var generated = DxpDocExport.Export(source.ToArray());
        if (Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_WORD") == "1")
        {
            var sourcePath = Path.Combine(Path.GetTempPath(),
                $"docxport-implicit-look-{Guid.NewGuid():N}.docx");
            var docPath = Path.ChangeExtension(sourcePath, ".doc");
            File.WriteAllBytes(sourcePath, source.ToArray());
            File.WriteAllBytes(docPath, generated);
            dynamic word = Activator.CreateInstance(
                Type.GetTypeFromProgID("Word.Application")!)!;
            word.Visible = false;
            word.DisplayAlerts = 0;
            dynamic? opened = null;
            int[]? sourceBandColors = null;
            try
            {
                foreach (var path in new[] { sourcePath, docPath })
                {
                    opened = word.Documents.Open(path, ReadOnly: true,
                        AddToRecentFiles: false);
                    var storyTables = new dynamic[]
                    {
                        opened.Tables.Item(1),
                        opened.Sections.Item(1).Headers.Item(1).Range.Tables.Item(1),
                        opened.Sections.Item(1).Footers.Item(1).Range.Tables.Item(1)
                    };
                    foreach (dynamic table in storyTables)
                        Assert.Equal(1, (int)table.Cell(1, 1).Range
                            .ParagraphFormat.Alignment);
                    var bandColors = storyTables.Select(table => (int)table
                        .Cell(2, 1).Shading.BackgroundPatternColor).ToArray();
                    Assert.All(bandColors, color => Assert.Equal(65535, color));
                    if (sourceBandColors == null) sourceBandColors = bandColors;
                    else Assert.Equal(sourceBandColors, bandColors);
                    opened.Close(false);
                    opened = null;
                }
            }
            finally
            {
                if (opened != null) opened.Close(false);
                word.Quit();
                File.Delete(sourcePath);
                File.Delete(docPath);
            }
        }
        foreach (var candidate in new[]
        {
            DxpDocToDocx.Project(generated).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(generated).DocxBytes)).DocxBytes
        })
        {
            using var stream = new MemoryStream(candidate);
            using var document = WordprocessingDocument.Open(stream, false);
            var main = document.MainDocumentPart!;
            var tables = main.Document.Body!.Elements<Table>()
                .Concat(main.HeaderParts.SelectMany(x => x.Header!.Elements<Table>()))
                .Concat(main.FooterParts.SelectMany(x => x.Footer!.Elements<Table>()))
                .ToArray();
            Assert.Equal(3, tables.Length);
            Assert.All(tables, table =>
            {
                Assert.Null(table.TableProperties?.GetFirstChild<TableLook>());
                Assert.All(table.Elements<TableRow>().First().Descendants<Paragraph>(),
                    paragraph => Assert.Equal(JustificationValues.Center,
                        paragraph.ParagraphProperties?.Justification?.Val?.Value));
                Assert.Equal(2, table.Elements<TableRow>().Count());
                Assert.All(table.Elements<TableRow>().Last().Elements<TableCell>(),
                    cell => Assert.Equal("FFFF00",
                        cell.TableCellProperties?.Shading?.Fill?.Value));
            });
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
    }

    [Fact]
    public void WordAuthoredConditionalBordersKeepReusableGeneratedRules()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordConditionalColumnBorders";
        var generated = DxpDocExport.Export(File.ReadAllBytes(
            Path.Combine(directory, name + ".docx")));
        var rules = new (ushort Code, TableStyleOverrideValues Kind,
            string Edge, uint Color)[]
        {
            (0x0001, TableStyleOverrideValues.FirstRow, "top", 0x000000AA),
            (0x0002, TableStyleOverrideValues.LastRow, "bottom", 0x000088CC),
            (0x0004, TableStyleOverrideValues.FirstColumn, "left", 0x0000AA00),
            (0x0008, TableStyleOverrideValues.LastColumn, "right", 0x00AA0000),
            (0x0100, TableStyleOverrideValues.NorthEastCell, "top", 0x00442277),
            (0x0200, TableStyleOverrideValues.NorthWestCell, "top", 0x00774411),
            (0x0400, TableStyleOverrideValues.SouthEastCell, "bottom", 0x00114488),
            (0x0800, TableStyleOverrideValues.SouthWestCell, "bottom", 0x00447722)
        };
        using (var input = new MemoryStream(generated))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            var style = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Column Borders");
            foreach (var (code, _, edge, color) in rules)
            {
                var borders = style.ConditionalTableBorders?[code];
                var actual = edge switch
                {
                    "top" => borders?.Top,
                    "bottom" => borders?.Bottom,
                    "left" => borders?.Left,
                    _ => borders?.Right
                };
                Assert.Equal(color, actual?.ColorRgb);
            }
        }
        var projected = DxpDocToDocx.Project(generated).DocxBytes;
        foreach (var binary in new[]
        {
            generated,
            File.ReadAllBytes(Path.Combine(directory, name + ".doc"))
        })
        {
            using var lookStream = new MemoryStream(DxpDocToDocx.Project(binary).DocxBytes);
            using var lookDocument = WordprocessingDocument.Open(lookStream, false);
            var look = Assert.Single(lookDocument.MainDocumentPart!.Document.Body!
                .Elements<Table>()).TableProperties?.GetFirstChild<TableLook>();
            Assert.Equal("07E0", look?.Val?.Value);
        }
        foreach (var candidate in new[]
        {
            projected,
            DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes
        })
        {
            using var stream = new MemoryStream(candidate);
            using var document = WordprocessingDocument.Open(stream, false);
            var style = Assert.Single(document.MainDocumentPart!
                .StyleDefinitionsPart!.Styles!.Elements<Style>(), x =>
                x.StyleName?.Val?.Value == "Column Borders");
            foreach (var (_, kind, edge, color) in rules)
            {
                var rule = style.Elements<TableStyleProperties>().Single(x =>
                    x.Type?.Value == kind);
                var border = rule.Descendants<TableCellBorders>().Single()
                    .ChildElements.Single(x => x.LocalName == edge);
                Assert.Equal($"{((color & 0xFF) << 16) | (color & 0xFF00) | (color >> 16):X6}",
                    border.GetAttribute("color",
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value);
            }
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
    }

    [Fact]
    public void AllConditionalTableStyleRegionsSurviveGeneratedDocAndThirdHop()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        using var source = new MemoryStream();
        source.Write(File.ReadAllBytes(Path.Combine(directory,
            "WordConditionalFirstRowParagraphAllStories.docx")));
        var regions = new (ushort Code, TableStyleOverrideValues Kind)[]
        {
            (0x0001, TableStyleOverrideValues.FirstRow),
            (0x0002, TableStyleOverrideValues.LastRow),
            (0x0004, TableStyleOverrideValues.FirstColumn),
            (0x0008, TableStyleOverrideValues.LastColumn),
            (0x0010, TableStyleOverrideValues.Band1Vertical),
            (0x0020, TableStyleOverrideValues.Band2Vertical),
            (0x0040, TableStyleOverrideValues.Band1Horizontal),
            (0x0080, TableStyleOverrideValues.Band2Horizontal),
            (0x0100, TableStyleOverrideValues.NorthEastCell),
            (0x0200, TableStyleOverrideValues.NorthWestCell),
            (0x0400, TableStyleOverrideValues.SouthEastCell),
            (0x0800, TableStyleOverrideValues.SouthWestCell)
        };
        using (var document = WordprocessingDocument.Open(source, true))
        {
            var styles = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            var style = Assert.Single(styles.Elements<Style>(), x =>
                x.StyleId?.Value == "LogicalStart");
            style.StyleTableProperties!.PrependChild(
                new TableStyleColumnBandSize { Val = 3 });
            style.StyleTableProperties.PrependChild(
                new TableStyleRowBandSize { Val = 2 });
            var first = style.Elements<TableStyleProperties>().Single(x =>
                x.Type?.Value == TableStyleOverrideValues.FirstRow);
            var borders = new TableCellBorders(
                new TopBorder { Val = BorderValues.Single, Size = 8, Color = "CC0000" },
                new LeftBorder { Val = BorderValues.Single, Size = 8, Color = "CC0000" },
                new BottomBorder { Val = BorderValues.Single, Size = 8, Color = "CC0000" },
                new RightBorder { Val = BorderValues.Single, Size = 8, Color = "CC0000" },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 8, Color = "CC0000" },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 8, Color = "CC0000" },
                new TopLeftToBottomRightCellBorder { Val = BorderValues.Single, Size = 8, Color = "CC0000" },
                new TopRightToBottomLeftCellBorder { Val = BorderValues.Single, Size = 8, Color = "CC0000" });
            first.TableStyleConditionalFormattingTableCellProperties!
                .PrependChild(borders);
            foreach (var (_, kind) in regions.Skip(1))
            {
                var copy = (TableStyleProperties)first.CloneNode(true);
                copy.Type = kind;
                style.AppendChild(copy);
            }
            styles.Save();
        }
        var generated = DxpDocExport.Export(source.ToArray());
        using (var input = new MemoryStream(generated))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            var style = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Logical Start");
            Assert.Equal((byte)2, style.TableFormatting?.TableHorizontalBandSize);
            Assert.Equal((byte)3, style.TableFormatting?.TableVerticalBandSize);
            foreach (var (code, _) in regions)
            {
                Assert.Equal((byte)1,
                    style.ConditionalParagraphFormatting?[code].Justification);
                Assert.True(style.ConditionalCharacterFormatting?[code].Bold);
                Assert.Equal(0x0000FF00u,
                    style.ConditionalTableShading?[code].FillRgb);
                var edges = style.ConditionalTableBorders?[code];
                Assert.NotNull(edges?.Top);
                Assert.NotNull(edges?.Bottom);
                Assert.NotNull(edges?.Left);
                Assert.NotNull(edges?.Right);
                Assert.NotNull(edges?.InsideHorizontal);
                Assert.NotNull(edges?.InsideVertical);
                Assert.NotNull(edges?.TopLeftToBottomRight);
                Assert.NotNull(edges?.TopRightToBottomLeft);
            }
        }
        var projected = DxpDocToDocx.Project(generated).DocxBytes;
        if (Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_WORD") == "1")
        {
            var docPath = Path.Combine(Path.GetTempPath(),
                $"docxport-all-conditions-{Guid.NewGuid():N}.doc");
            var docxPath = Path.ChangeExtension(docPath, ".docx");
            File.WriteAllBytes(docPath, generated);
            dynamic word = Activator.CreateInstance(
                Type.GetTypeFromProgID("Word.Application")!)!;
            word.Visible = false;
            word.DisplayAlerts = 0;
            dynamic? opened = null;
            try
            {
                opened = word.Documents.Open(docPath, ReadOnly: true,
                    AddToRecentFiles: false);
                opened.SaveAs2(docxPath, 16);
                opened.Close(false);
                opened = null;
                using var saved = WordprocessingDocument.Open(docxPath, false);
                var style = Assert.Single(saved.MainDocumentPart!
                    .StyleDefinitionsPart!.Styles!.Elements<Style>(), x =>
                    x.StyleName?.Val?.Value == "Logical Start");
                Assert.Equal(2, style.StyleTableProperties?
                    .GetFirstChild<TableStyleRowBandSize>()?.Val?.Value);
                Assert.Equal(3, style.StyleTableProperties?
                    .GetFirstChild<TableStyleColumnBandSize>()?.Val?.Value);
                foreach (var (_, kind) in regions)
                {
                    var rule = style.Elements<TableStyleProperties>().Single(x =>
                        x.Type?.Value == kind);
                    Assert.Equal(JustificationValues.Center,
                        rule.Descendants<Justification>().Single().Val?.Value);
                    Assert.Single(rule.Descendants<Bold>());
                    Assert.Equal("00FF00", rule.Descendants<Shading>().Single().Fill?.Value);
                    Assert.Equal("CC0000", rule.Descendants<TopBorder>().Single().Color?.Value);
                    Assert.Equal(8U, rule.Descendants<InsideVerticalBorder>().Single().Size?.Value);
                }
            }
            finally
            {
                if (opened != null) opened.Close(false);
                word.Quit();
                File.Delete(docPath);
                File.Delete(docxPath);
            }
        }
        foreach (var candidate in new[]
        {
            projected,
            DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes
        })
        {
            using var stream = new MemoryStream(candidate);
            using var document = WordprocessingDocument.Open(stream, false);
            var style = Assert.Single(document.MainDocumentPart!
                .StyleDefinitionsPart!.Styles!.Elements<Style>(), x =>
                x.StyleName?.Val?.Value == "Logical Start");
            Assert.Equal(2, style.StyleTableProperties?
                .GetFirstChild<TableStyleRowBandSize>()?.Val?.Value);
            Assert.Equal(3, style.StyleTableProperties?
                .GetFirstChild<TableStyleColumnBandSize>()?.Val?.Value);
            foreach (var (_, kind) in regions)
            {
                var rule = style.Elements<TableStyleProperties>().Single(x =>
                    x.Type?.Value == kind);
                Assert.Equal(JustificationValues.Center,
                    rule.Descendants<Justification>().Single().Val?.Value);
                Assert.Single(rule.Descendants<Bold>());
                Assert.Equal("FF0000", rule.Descendants<Color>().Single().Val?.Value);
                Assert.Equal("00FF00", rule.Descendants<Shading>().Single().Fill?.Value);
                Assert.Equal("CC0000", rule.Descendants<TopBorder>().Single().Color?.Value);
                Assert.Equal(8U, rule.Descendants<InsideVerticalBorder>().Single().Size?.Value);
                Assert.Single(rule.Descendants<TopLeftToBottomRightCellBorder>());
                Assert.Single(rule.Descendants<TopRightToBottomLeftCellBorder>());
            }
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
    }

    [Theory]
    [InlineData("WordDocPropertyLastSavedByBodyWithMergeStories",
        "FIELD:DOCPROPERTY:LastSavedBy")]
    [InlineData("WordLastSavedByBodyWithMergeStories", "FIELD:LASTSAVEDBY")]
    [InlineData("WordLastSavedByBodyHeader", "FIELD:LASTSAVEDBY")]
    [InlineData("WordLastSavedByBodyFooter", "FIELD:LASTSAVEDBY")]
    public void LastSavedByFieldAndMetadataSurviveBothDocRoutesAndThirdHop(
        string name, string fieldSemantics)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, name);
        var source = File.ReadAllBytes(stem + ".docx");
        var generated = DxpDocExport.Export(source);
        var native = File.ReadAllBytes(stem + ".doc");
        static string? LastAuthor(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes);
            using var document = WordprocessingDocument.Open(stream, false);
            return document.PackageProperties.LastModifiedBy;
        }
        var expected = LastAuthor(source);
        Assert.False(string.IsNullOrWhiteSpace(expected));
        Assert.Contains(fieldSemantics, ReadEditableFieldSemantics(source));
        foreach (var doc in new[] { generated, native })
        {
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            Assert.Equal(expected, LastAuthor(projected));
            Assert.Equal(ReadEditableFieldSemantics(source),
                ReadEditableFieldSemantics(projected));
            Assert.Equal(ReadStories(source), ReadStories(projected));
            var repeated = DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes;
            Assert.Equal(expected, LastAuthor(repeated));
            Assert.Equal(ReadEditableFieldSemantics(source),
                ReadEditableFieldSemantics(repeated));
            Assert.Equal(ReadStories(source), ReadStories(repeated));
            Assert.Empty(Validate(projected));
            Assert.Empty(Validate(repeated));
        }
    }

    [Theory]
    [InlineData("WordAuthorBodyWithMergeStories", "AUTHOR", "Ada Lovelace")]
    [InlineData("WordTitleBodyWithMergeStories", "TITLE", "Quarterly Report")]
    [InlineData("WordSubjectBodyWithMergeStories", "SUBJECT", "Internal Review")]
    [InlineData("WordKeywordsBodyWithMergeStories", "KEYWORDS", "Project Alpha")]
    [InlineData("WordCommentsBodyWithMergeStories", "COMMENTS", "Approved draft")]
    public void StandaloneCorePropertyFieldAndMetadataSurviveBothDocRoutesAndThirdHop(
        string name, string fieldName, string expectedValue)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, name);
        var source = File.ReadAllBytes(stem + ".docx");
        Assert.Contains("FIELD:" + fieldName, ReadEditableFieldSemantics(source));
        foreach (var doc in new[]
        {
            DxpDocExport.Export(source), File.ReadAllBytes(stem + ".doc")
        })
        {
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            var repeated = DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes;
            foreach (var output in new[] { projected, repeated })
            {
                Assert.Equal(ReadStories(source), ReadStories(output));
                Assert.Equal(ReadEditableFieldSemantics(source),
                    ReadEditableFieldSemantics(output));
                using var stream = new MemoryStream(output);
                using var document = WordprocessingDocument.Open(stream, false);
                Assert.Equal(expectedValue, fieldName switch
                {
                    "AUTHOR" => document.PackageProperties.Creator,
                    "TITLE" => document.PackageProperties.Title,
                    "SUBJECT" => document.PackageProperties.Subject,
                    "KEYWORDS" => document.PackageProperties.Keywords,
                    "COMMENTS" => document.PackageProperties.Description,
                    _ => throw new InvalidOperationException()
                });
                Assert.Empty(Validate(output));
            }
        }
    }

    [Theory]
    [InlineData("WordDistributedParagraphAllStories", false)]
    [InlineData("WordStyledDistributedParagraphAllStories", true)]
    public void DistributedParagraphAlignmentSurvivesBothDocRoutes(
        string name, bool styled)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, name);
        var source = File.ReadAllBytes(stem + ".docx");
        foreach (var doc in new[]
        {
            File.ReadAllBytes(stem + ".doc"), DxpDocExport.Export(source)
        })
        {
            using var input = new MemoryStream(doc);
            using var index = new DocTextIndexWalker().Index(input);
            if (styled)
                Assert.Contains(index.StyleDefinitions, x =>
                    x.Name == "Decorated Note" &&
                    x.ParagraphFormatting?.Justification == 4);
            else
                Assert.True(index.ParagraphStyles.Count(x =>
                    x.Formatting?.Justification == 4) >= 3,
                    string.Join(",", index.ParagraphStyles.Select(x =>
                        x.Formatting?.Justification?.ToString() ?? "<none>")));
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            Assert.Equal(ReadStories(source), ReadStories(projected));
            AssertEffectiveParagraphLayout(source, projected);
            var repeated = DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes;
            AssertEffectiveParagraphLayout(source, repeated);
            Assert.Empty(Validate(projected));
            Assert.Empty(Validate(repeated));
        }
    }

    [Fact]
    public void DistributedStyleInheritanceAndHeaderOverrideSurviveBothDocRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, "WordDistributedStyleInheritanceAllStories");
        var source = File.ReadAllBytes(stem + ".docx");
        foreach (var doc in new[]
        {
            File.ReadAllBytes(stem + ".doc"), DxpDocExport.Export(source)
        })
        {
            using var input = new MemoryStream(doc);
            using var index = new DocTextIndexWalker().Index(input);
            var parent = Assert.Single(index.StyleDefinitions, x =>
                x.Name == "Decorated Note");
            var child = Assert.Single(index.StyleDefinitions, x =>
                x.Name == "Distributed Child");
            Assert.Equal(parent.Index, child.BasedOnIndex);
            Assert.Equal((byte)4, child.ParagraphFormatting?.Justification);
            Assert.Contains(index.ParagraphStyles, x =>
                x.Formatting?.Justification == 1);
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            var repeated = DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes;
            foreach (var output in new[] { projected, repeated })
            {
                Assert.Equal(ReadStories(source), ReadStories(output));
                AssertEffectiveParagraphLayout(source, output);
                Assert.Empty(Validate(output));
            }
        }
    }

    [Theory]
    [InlineData("WordConditionalFirstRowRunStyleAllStories", false, false)]
    [InlineData("WordConditionalFirstRowParagraphAllStories", true, false)]
    [InlineData("WordConditionalFirstLastRowParagraphAllStories", true, true)]
    [InlineData("WordMaskedTableLookParagraphAllStories", true, true)]
    public void ConditionalFirstRowRunStyleSurvivesBothDocRoutes(
        string name, bool centered, bool lastRowRight)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures",
            "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        using (var stream = new MemoryStream(source))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var style = Assert.Single(document.MainDocumentPart!
                .StyleDefinitionsPart!.Styles!.Elements<Style>(), x =>
                x.StyleId?.Value == "LogicalStart");
            var firstRow = style.Elements<TableStyleProperties>().Single(x =>
                x.Type?.Value == TableStyleOverrideValues.FirstRow);
            Assert.Single(firstRow.Descendants<Bold>());
            Assert.Equal("FF0000", firstRow.Descendants<Color>().Single().Val?.Value);
            if (centered)
                Assert.Equal(JustificationValues.Center,
                    firstRow.Descendants<Justification>().Single().Val?.Value);
            if (lastRowRight)
            {
                var lastRow = style.Elements<TableStyleProperties>().Single(x =>
                    x.Type?.Value == TableStyleOverrideValues.LastRow);
                Assert.Equal(JustificationValues.Right,
                    lastRow.Descendants<Justification>().Single().Val?.Value);
            }
        }
        var generated = DxpDocExport.Export(source);
        using (var input = new MemoryStream(generated))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            var tableStyle = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Logical Start");
            Assert.True(tableStyle.ConditionalCharacterFormatting?[1].Bold);
            Assert.Equal(0x000000FFu,
                tableStyle.ConditionalCharacterFormatting?[1].ColorRef);
            Assert.Equal(0x0000FF00u,
                tableStyle.ConditionalTableShading?[1].FillRgb);
        }
        if (centered)
        {
            using var input = new MemoryStream(generated);
            using var index = new DocTextIndexWalker().Index(input);
            var tableStyle = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Logical Start");
            Assert.Equal((byte)1, tableStyle.ConditionalParagraphFormatting?[1]
                .Justification);
            if (lastRowRight)
                Assert.Equal((byte)2, tableStyle.ConditionalParagraphFormatting?[2]
                    .Justification);
        }
        if (lastRowRight &&
            Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_WORD") == "1")
        {
            var nativePath = Path.Combine(Path.GetTempPath(),
                $"docxport-conditional-style-{Guid.NewGuid():N}.doc");
            var savedPath = Path.ChangeExtension(nativePath, ".docx");
            File.WriteAllBytes(nativePath, generated);
            dynamic word = Activator.CreateInstance(
                Type.GetTypeFromProgID("Word.Application")!)!;
            word.Visible = false;
            word.DisplayAlerts = 0;
            dynamic? opened = null;
            try
            {
                opened = word.Documents.Open(nativePath, ReadOnly: true,
                    AddToRecentFiles: false);
                opened.SaveAs2(savedPath, 16);
                opened.Close(false);
                opened = null;
                using var saved = WordprocessingDocument.Open(savedPath, false);
                var style = Assert.Single(saved.MainDocumentPart!
                    .StyleDefinitionsPart!.Styles!.Elements<Style>(), x =>
                    x.StyleName?.Val?.Value == "Logical Start");
                Assert.Equal(JustificationValues.Center,
                    style.Elements<TableStyleProperties>().Single(x =>
                        x.Type?.Value == TableStyleOverrideValues.FirstRow)
                        .Descendants<Justification>().Single().Val?.Value);
                Assert.Equal(JustificationValues.Right,
                    style.Elements<TableStyleProperties>().Single(x =>
                        x.Type?.Value == TableStyleOverrideValues.LastRow)
                        .Descendants<Justification>().Single().Val?.Value);
                var firstRow = style.Elements<TableStyleProperties>().Single(x =>
                    x.Type?.Value == TableStyleOverrideValues.FirstRow);
                Assert.Single(firstRow.Descendants<Bold>());
                Assert.Equal("FF0000", firstRow.Descendants<Color>().Single().Val?.Value);
                Assert.Equal("00FF00", firstRow.Descendants<Shading>().Single().Fill?.Value);
            }
            finally
            {
                if (opened != null) opened.Close(false);
                word.Quit();
                File.Delete(nativePath);
                File.Delete(savedPath);
            }
        }
        foreach (var binary in new[]
        {
            File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
            generated
        })
        {
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            foreach (var candidate in new[]
            {
                projected,
                DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes
            })
            {
                using var stream = new MemoryStream(candidate);
                using var document = WordprocessingDocument.Open(stream, false);
                var main = document.MainDocumentPart!;
                var tables = main.Document.Body!.Elements<Table>()
                    .Concat(main.HeaderParts.SelectMany(x => x.Header!.Elements<Table>()))
                    .Concat(main.FooterParts.SelectMany(x => x.Footer!.Elements<Table>()))
                    .ToArray();
                Assert.Equal(3, tables.Length);
                Assert.All(tables, table =>
                {
                    var runs = table.Elements<TableRow>().First()
                        .Descendants<Run>().ToArray();
                    Assert.NotEmpty(runs);
                    if (centered)
                        Assert.All(table.Elements<TableRow>().First()
                            .Descendants<Paragraph>(), paragraph =>
                            Assert.Equal(JustificationValues.Center,
                                paragraph.ParagraphProperties?
                                    .GetFirstChild<Justification>()?.Val?.Value));
                    if (lastRowRight)
                        Assert.All(table.Elements<TableRow>().Last()
                            .Descendants<Paragraph>(), paragraph =>
                            Assert.Equal(JustificationValues.Right,
                                paragraph.ParagraphProperties?
                                    .GetFirstChild<Justification>()?.Val?.Value));
                    Assert.All(runs, run =>
                    {
                        Assert.True(run.RunProperties?.Bold?.Val?.Value ??
                            run.RunProperties?.Bold != null);
                        Assert.Equal("FF0000", run.RunProperties?
                            .GetFirstChild<Color>()?.Val?.Value);
                    });
                });
                Assert.Empty(new OpenXmlValidator().Validate(document));
                if (centered && ReferenceEquals(binary, generated))
                {
                    var style = Assert.Single(main.StyleDefinitionsPart!.Styles!
                        .Elements<Style>(), x => x.StyleName?.Val?.Value ==
                            "Logical Start");
                    Assert.Equal(JustificationValues.Center,
                        style.Elements<TableStyleProperties>().Single(x =>
                            x.Type?.Value == TableStyleOverrideValues.FirstRow)
                            .Descendants<Justification>().Single().Val?.Value);
                    if (lastRowRight)
                        Assert.Equal(JustificationValues.Right,
                            style.Elements<TableStyleProperties>().Single(x =>
                                x.Type?.Value == TableStyleOverrideValues.LastRow)
                                .Descendants<Justification>().Single().Val?.Value);
                }
                if (ReferenceEquals(binary, generated))
                {
                    var style = Assert.Single(main.StyleDefinitionsPart!.Styles!
                        .Elements<Style>(), x => x.StyleName?.Val?.Value ==
                            "Logical Start");
                    var firstRow = style.Elements<TableStyleProperties>().Single(x =>
                        x.Type?.Value == TableStyleOverrideValues.FirstRow);
                    Assert.Single(firstRow.Descendants<Bold>());
                    Assert.Equal("FF0000", firstRow.Descendants<Color>().Single().Val?.Value);
                    Assert.Equal("00FF00", firstRow.Descendants<Shading>().Single().Fill?.Value);
                }
            }
        }
    }

    [Fact]
    public void WordNativeConditionalFirstRowFillSurvivesAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures",
            "Doc"));
        const string name = "WordConditionalFirstRowShadingAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        using (var stream = new MemoryStream(source))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var style = Assert.Single(document.MainDocumentPart!
                .StyleDefinitionsPart!.Styles!.Elements<Style>(), x =>
                x.StyleId?.Value == "LogicalStart");
            Assert.Equal("00FF00", style.Elements<TableStyleProperties>()
                .Single(x => x.Type?.Value == TableStyleOverrideValues.FirstRow)
                .Descendants<Shading>().Single().Fill?.Value);
        }
        foreach (var binary in new[]
        {
            File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
            DxpDocExport.Export(source)
        })
        {
            foreach (var projected in new[]
            {
                DxpDocToDocx.Project(binary).DocxBytes,
                DxpDocToDocx.Project(DxpDocExport.Export(
                    DxpDocToDocx.Project(binary).DocxBytes)).DocxBytes
            })
            {
                using var stream = new MemoryStream(projected);
                using var document = WordprocessingDocument.Open(stream, false);
                var main = document.MainDocumentPart!;
                var tables = main.Document.Body!.Elements<Table>()
                    .Concat(main.HeaderParts.SelectMany(x => x.Header!.Elements<Table>()))
                    .Concat(main.FooterParts.SelectMany(x => x.Footer!.Elements<Table>()))
                    .ToArray();
                Assert.Equal(3, tables.Length);
                Assert.All(tables, table =>
                {
                    var rows = table.Elements<TableRow>().ToArray();
                    Assert.NotEmpty(rows);
                    Assert.All(rows[0].Elements<TableCell>(), cell =>
                        Assert.Equal("00FF00", cell.TableCellProperties?
                            .GetFirstChild<Shading>()?.Fill?.Value));
                    foreach (var later in rows.Skip(1))
                        Assert.All(later.Elements<TableCell>(), cell =>
                            Assert.NotEqual("00FF00", cell.TableCellProperties?
                                .GetFirstChild<Shading>()?.Fill?.Value));
                });
                Assert.Empty(new OpenXmlValidator().Validate(document));
            }
        }
    }

    [Fact]
    public void StylelessTablesUseWordImplicitCellPadding()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures",
            "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordOccupiedHMergeAllStories.docx"));
        using (var stream = new MemoryStream(source))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            Assert.Null(document.MainDocumentPart!.StyleDefinitionsPart);
            Assert.Null(document.MainDocumentPart.Document.Body!
                .GetFirstChild<Table>()!.TableProperties?
                .GetFirstChild<TableCellMarginDefault>());
        }
        foreach (var binary in new[]
        {
            File.ReadAllBytes(Path.Combine(directory,
                "WordOccupiedHMergeAllStories.doc")),
            DxpDocExport.Export(source)
        })
        {
            using var stream = new MemoryStream(binary);
            using var index = new DocTextIndexWalker().Index(stream);
            var rows = index.ParagraphStyles.Where(x =>
                x.Formatting?.TableTerminator == true).ToArray();
            Assert.Equal(6, rows.Length);
            Assert.All(rows, row =>
            {
                Assert.Equal((ushort)10,
                    row.Formatting!.TableDefaultCellMargins?.Left);
                Assert.Equal((ushort)10,
                    row.Formatting.TableDefaultCellMargins?.Right);
            });
            AssertTableGrid(source, DxpDocToDocx.Project(binary).DocxBytes);
        }
    }

    [Fact]
    public void ThreeColumnOccupiedMergeFitsLongContinuationAcrossAllStories()
    {
        var metrics = DocSystemFontAdvances.Find("Arial");
        if (metrics == null) return;
        const string longText = "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW";
        Assert.True(metrics.TryMeasure(longText, 12, out var widthPoints));
        var minimumTwips = widthPoints * 20 + 20;
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordOccupiedThreeCellHMergeArialAllStories.docx"));
        using var source = new MemoryStream();
        source.Write(File.ReadAllBytes(path));
        source.Position = 0;
        using (var package = WordprocessingDocument.Open(source, true))
        {
            var main = package.MainDocumentPart!;
            var tables = main.Document.Body!.Elements<Table>()
                .Concat(main.HeaderParts.SelectMany(x => x.Header!.Elements<Table>()))
                .Concat(main.FooterParts.SelectMany(x => x.Footer!.Elements<Table>()));
            foreach (var table in tables)
                table.Elements<TableRow>().First().Elements<TableCell>()
                    .ElementAt(2).Descendants<Text>().First().Text = longText;
        }
        using var binary = new MemoryStream(DxpDocExport.Export(source.ToArray()));
        using var index = new DocTextIndexWalker().Index(binary);
        var rows = index.ParagraphStyles.Where(x =>
            x.Formatting?.TableTerminator == true).ToArray();
        Assert.Equal(6, rows.Length);
        foreach (var row in rows.Where((_, i) => i % 2 == 0))
        {
            var edges = row.Formatting!.TableCellEdges!;
            Assert.Equal(4, edges.Count);
            Assert.True(edges[3] - edges[0] >= minimumTwips,
                $"Merged text needs {minimumTwips:F1} twips, got {edges[3] - edges[0]}.");
        }
        using var projection = WordprocessingDocument.Open(new MemoryStream(
            DxpDocToDocx.Project(binary.ToArray()).DocxBytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(projection));
    }
    [Theory]
    [InlineData("WordOccupiedHMergeExplicitMarginsAllStories", 2)]
    [InlineData("WordOccupiedHMergeAllStories", 2)]
    [InlineData("WordOccupiedThreeCellHMergeAllStories", 3)]
    public void OccupiedHorizontalMergesKeepPhysicalCellsInAllStories(
        string name, int mergedCellCount)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var native = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
        var generated = DxpDocExport.Export(source);
        foreach (var binary in new[] { native, generated })
        {
            using (var input = new MemoryStream(binary))
            using (var index = new DocTextIndexWalker().Index(input))
            {
                var expected = new byte?[] { 2 }
                    .Concat(Enumerable.Repeat<byte?>(1, mergedCellCount - 1)).ToArray();
                Assert.Equal(3, index.ParagraphStyles.Count(range =>
                    range.Formatting?.TableCellHorizontalMerges is { } states &&
                    states.SequenceEqual(expected)));
            }
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            foreach (var docxBytes in new[]
            {
                projected,
                DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes
            })
            {
                using var stream = new MemoryStream(docxBytes);
                using var document = WordprocessingDocument.Open(stream, false);
                var main = document.MainDocumentPart!;
                var tables = main.Document.Body!.Elements<Table>()
                    .Concat(main.HeaderParts.SelectMany(part =>
                        part.Header!.Elements<Table>()))
                    .Concat(main.FooterParts.SelectMany(part =>
                        part.Footer!.Elements<Table>())).ToArray();
                Assert.Equal(3, tables.Length);
                foreach (var table in tables)
                {
                    var cells = table.Elements<TableRow>().First()
                        .Elements<TableCell>().ToArray();
                    Assert.Equal(mergedCellCount, cells.Length);
                    Assert.Equal(MergedCellValues.Restart,
                        cells[0].TableCellProperties!
                            .GetFirstChild<HorizontalMerge>()!.Val!.Value);
                    for (var i = 1; i < cells.Length; i++)
                    {
                        Assert.Equal(MergedCellValues.Continue,
                            cells[i].TableCellProperties!
                                .GetFirstChild<HorizontalMerge>()!.Val!.Value);
                        Assert.NotEmpty(cells[i].InnerText);
                    }
                }
                Assert.Empty(new OpenXmlValidator().Validate(document));
            }
        }
    }

    [Fact]
    public void AutoFitRowAndInlinePictureShareDataStream()
    {
        var picture = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Z7n8AAAAASUVORK5CYII=");
        const string text = "A\u0007\u0007\u0001\r";
        var inTable = DocParagraphFormatting.Empty with { InTable = true };
        var row = inTable with
        {
            TableTerminator = true,
            TableCellEdges = new short[] { 0, 2000 },
            TableAutoFit = true
        };
        var paragraphStyles = new DocStoryParagraphStyleRun[]
        {
            new(0, 2, 0, inTable), new(2, 3, 0, row),
            new(3, 5, 0, DocParagraphFormatting.Empty)
        };
        var document = new DocPlainTextDocument(
            new DocPlainTextStory(text, [], paragraphStyles)
            {
                Paragraphs = DocStoryParagraphRange.Capture(text,
                    paragraphStyles.Select(x => x.ToWriter(0)).ToArray()),
                Pictures = [new DocStoryInlinePicture(3)
                {
                    Payload = new DocInlinePicture(picture, "image/png", 914400, 914400)
                }]
            },
            [new DocPlainTextSection(text.Length, new DocPlainTextStory?[6])]);
        using var binary = new MemoryStream();
        DocPlainTextWriter.Write(binary, document);
        binary.Position = 0;
        using var structure = new DocStructureWalker().Accept(binary,
            new DocStructurePrintVisitor(TextWriter.Null));
        static IEnumerable<DocStructureNode> Descendants(DocStructureNode node) =>
            node.Children.SelectMany(child => new[] { child }.Concat(Descendants(child)));
        var rowBlock = Assert.Single(Descendants(structure.Root), node =>
            node.Kind == "PrcData" && node.Name == "TableProperties");
        Assert.Equal("Data", rowBlock.StreamName);
        Assert.True(rowBlock.Offset > 0);
        binary.Position = 0;
        using var index = new DocTextIndexWalker().Index(binary);
        using var projected = WordprocessingDocument.Open(new MemoryStream(
            new DocToDocxProjector().Project(index).DocxBytes), false);
        Assert.Single(projected.MainDocumentPart!.Document!.Body!
            .Descendants<Drawing>());
        using var projectedImage = Assert.Single(projected.MainDocumentPart!
            .ImageParts).GetStream();
        using var copied = new MemoryStream();
        projectedImage.CopyTo(copied);
        Assert.Equal(picture, copied.ToArray());
    }

    [Fact]
    public void ArialAutoWidthNoWrapKnownGapRetainsEditableSemanticsInBothDocRoutes()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures",
            "DocKnownGaps", "WordVisibleNoWrapArialAllStories.docx"));
        var source = File.ReadAllBytes(path);
        var expectedStories = ReadStories(source);
        foreach (var doc in new[]
        {
            File.ReadAllBytes(Path.ChangeExtension(path, ".doc")),
            DxpDocExport.Export(source)
        })
        {
            foreach (var binary in new[] { doc,
                DxpDocExport.Export(DxpDocToDocx.Project(doc).DocxBytes) })
            {
                using (var stream = new MemoryStream(binary))
                using (var index = new DocTextIndexWalker().Index(stream))
                {
                    var rows = index.ParagraphStyles.Where(row =>
                        row.Formatting?.TableTerminator == true).ToArray();
                    Assert.Equal(4, rows.Length);
                    Assert.All(rows, row =>
                    {
                        Assert.True(row.Formatting!.TableAutoFit);
                        Assert.Equal(new DocTablePreferredWidth(1, 0),
                            row.Formatting.TablePreferredWidth);
                        Assert.All(row.Formatting.TableCellPreferredWidths!, width =>
                            Assert.Equal(new DocTablePreferredWidth(1, 0), width));
                    });
                    Assert.True(rows[0].Formatting!.TableCellNoWraps![0]);
                }
                var projected = DxpDocToDocx.Project(binary);
                Assert.Empty(projected.Coverage.OmittedCharacters);
                Assert.Empty(projected.Coverage.ApproximateCharacters);
                Assert.Empty(projected.Coverage.DeferredParts);
                Assert.Equal(expectedStories, ReadStories(projected.DocxBytes));
                AssertTableGrid(source, projected.DocxBytes);
                using (var package = WordprocessingDocument.Open(
                    new MemoryStream(projected.DocxBytes), false))
                {
                    var main = package.MainDocumentPart!;
                    var tables = new DocumentFormat.OpenXml.OpenXmlElement[]
                    { main.Document.Body! }.Concat(main.HeaderParts.Select(x =>
                        (DocumentFormat.OpenXml.OpenXmlElement)x.Header!)).Concat(
                        main.FooterParts.Select(x =>
                            (DocumentFormat.OpenXml.OpenXmlElement)x.Footer!))
                        .SelectMany(x => x.Descendants<Table>()).ToArray();
                    Assert.Equal(3, tables.Length);
                    foreach (var table in tables)
                    {
                        Assert.Equal(TableWidthUnitValues.Auto,
                            table.TableProperties?.TableWidth?.Type?.Value);
                        Assert.All(table.Descendants<TableCell>(), cell =>
                            Assert.Equal(TableWidthUnitValues.Auto,
                                cell.TableCellProperties?.TableCellWidth?.Type?.Value));
                    }
                }
                Assert.Empty(Validate(projected.DocxBytes));
            }
        }
    }

    [Theory]
    [InlineData("WordVisibleNoWrapOffAllStories")]
    [InlineData("WordVisibleNoWrapAllStories")]
    public void GeneratedAutoWidthRowsCarryReferencedModernCellDefinitions(string name)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            name + ".docx"));
        var generated = DxpDocExport.Export(File.ReadAllBytes(path));
        using var input = new MemoryStream(generated);
        using var structure = new DocStructureWalker().Accept(input,
            new DocStructurePrintVisitor(TextWriter.Null));
        static IEnumerable<DocStructureNode> Descendants(DocStructureNode node) =>
            node.Children.SelectMany(child => new[] { child }.Concat(Descendants(child)));
        var rows = Descendants(structure.Root).Where(node =>
            node.Kind == "PrcData" && node.Name == "TableProperties").ToArray();
        Assert.Equal(4, rows.Length);
        foreach (var row in rows)
        {
            Assert.Equal("Data", row.StreamName);
            var codes = Descendants(row).Where(node => node.Kind == "Sprm")
                .Select(node => node.Attributes["code"]).ToArray();
            Assert.Contains("0x7621", codes);
            Assert.Equal(2, codes.Count(code => code == "0x7623"));
            // The modern gap-half operand makes Word reflow these auto-width
            // rows differently even when the compatibility PAPX retains it.
            Assert.DoesNotContain("0x9602", codes);
            Assert.Contains("0x3615", codes);
            Assert.Equal(4, codes.Count(code => code == "0xD634"));
            foreach (var padding in Descendants(row).Where(node =>
                node.Kind == "Prl" && node.Children.Any(child =>
                    child.Kind == "Sprm" && child.Attributes["code"] == "0xD634")))
            {
                var operand = padding.Children.Single(child => child.Kind == "CSSAOperand");
                var bytes = structure.ReadRange("Data", operand.Offset!.Value,
                    checked((int)operand.Length!.Value));
                // MS-DOC requires the row-wide default margin range 0..1.
                Assert.Equal(new byte[] { 6, 0, 1 }, bytes[..3]);
            }
            Assert.Contains("0x7479", codes);
        }
        input.Position = 0;
        using var index = new DocTextIndexWalker().Index(input);
        var generatedEdges = index.ParagraphStyles.Where(x =>
            x.Formatting?.TableTerminator == true)
            .Select(x => x.Formatting!.TableCellEdges![1]).ToArray();
        Assert.Equal(new short[] { 4680, 4680, 4680, 4680 }, generatedEdges);
        Assert.All(index.ParagraphStyles.Where(x =>
            x.Formatting?.TableTerminator == true), run =>
            {
                Assert.Equal((short)0, run.Formatting!.TableRowOriginTwips);
                Assert.Equal(new short[] { 0, 4680, 9360 },
                    run.Formatting.TableCellEdges);
                Assert.Equal((ushort)108, run.Formatting.TableDefaultCellMargins?.Left);
                Assert.Equal((ushort)108, run.Formatting.TableDefaultCellMargins?.Right);
                Assert.Equal(new DocTablePreferredWidth(1, 0),
                    run.Formatting.TablePreferredWidth);
                Assert.All(run.Formatting.TableCellPreferredWidths!,
                    width => Assert.Equal(new DocTablePreferredWidth(1, 0), width));
            });
    }

    [Fact]
    public void GeneratedThemeTintShadeMatchesSourceRgbAcrossStories()
    {
        var sourcePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleThemeTintShadeAllStories.docx"));
        var source = File.ReadAllBytes(sourcePath);
        var generatedDoc = DxpDocExport.Export(source);
        var generated = DxpDocToDocx.Project(generatedDoc).DocxBytes;
        var thirdHop = DxpDocToDocx.Project(DxpDocExport.Export(generated)).DocxBytes;
        var expected = ReadEffectiveRunFormatting(source)
            .Where(x => x.Key.EndsWith(".color", StringComparison.Ordinal))
            .ToDictionary(x => x.Key, x => x.Value);
        Assert.Contains(expected, x => x.Key.StartsWith("body.", StringComparison.Ordinal));
        Assert.Contains(expected, x => x.Key.Contains(".header.", StringComparison.Ordinal));
        Assert.Contains(expected, x => x.Key.Contains(".footer.", StringComparison.Ordinal));
        foreach (var result in new[] { generated, thirdHop })
        {
            var colors = ReadEffectiveRunFormatting(result)
                .Where(x => x.Key.EndsWith(".color", StringComparison.Ordinal))
                .ToDictionary(x => x.Key, x => x.Value);
            Assert.Equal(expected, colors);
        }
    }

    [Fact]
    public void WordNativeLegacyTableBorderOperandsSurviveBothRoutes()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordLegacyTableBordersAllStories.doc"));
        var native = File.ReadAllBytes(path);
        var generated = DxpDocExport.Export(File.ReadAllBytes(
            Path.ChangeExtension(path, ".docx")));
        static short?[] RowOrigins(byte[] bytes)
        {
            using var index = new DocTextIndexWalker().Index(new MemoryStream(bytes));
            return index.ParagraphStyles.Where(x =>
                x.Formatting?.TableTerminator == true).Select(x =>
                x.Formatting!.TableRowOriginTwips).ToArray();
        }
        Assert.Equal(new short?[] { 113, 113, 113, 113, 113, 113 },
            RowOrigins(native));
        Assert.Equal(RowOrigins(native), RowOrigins(generated));
        static IEnumerable<DocStructureNode> Descendants(DocStructureNode node) =>
            node.Children.SelectMany(child => new[] { child }.Concat(Descendants(child)));
        foreach (var (doc, route) in new[] { native, generated,
            DxpDocExport.Export(DxpDocToDocx.Project(native).DocxBytes),
            DxpDocExport.Export(DxpDocToDocx.Project(generated).DocxBytes) }
            .Select((bytes, route) => (bytes, route)))
        {
            using var input = new MemoryStream(doc);
            using var structure = new DocStructureWalker().Accept(input,
                new DocStructurePrintVisitor(TextWriter.Null));
            var legacyRows = Descendants(structure.Root).Count(node =>
                node.Kind == "Sprm" &&
                node.Attributes.TryGetValue("code", out var code) &&
                code == "0xD605");
            var legacyCells = Descendants(structure.Root).Count(node =>
                node.Kind == "Sprm" &&
                node.Attributes.TryGetValue("code", out var code) &&
                code == "0xD620");
            Assert.True(route == 0 ? legacyRows >= 6 : legacyRows + legacyCells >= 6,
                $"Route {route}: found {legacyRows} legacy rows and {legacyCells} legacy cells.");
        }
    }

    [Fact]
    public void WordNativeLegacyCellBorderOperandsSurviveBothRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        static IEnumerable<DocStructureNode> Descendants(DocStructureNode node) =>
            node.Children.SelectMany(child => new[] { child }.Concat(Descendants(child)));
        var path = Path.Combine(directory, "WordDiagonalBordersAllStories.doc");
        var native = File.ReadAllBytes(path);
        var generated = DxpDocExport.Export(File.ReadAllBytes(
            Path.ChangeExtension(path, ".docx")));
        foreach (var (doc, route) in new[] { native, generated,
            DxpDocExport.Export(DxpDocToDocx.Project(native).DocxBytes),
            DxpDocExport.Export(DxpDocToDocx.Project(generated).DocxBytes) }
            .Select((bytes, route) => (bytes, route)))
        {
            using var input = new MemoryStream(doc);
            using var structure = new DocStructureWalker().Accept(input,
                new DocStructurePrintVisitor(TextWriter.Null));
            var legacyCells = Descendants(structure.Root).Count(node =>
                node.Kind == "Sprm" &&
                node.Attributes.TryGetValue("code", out var code) &&
                code == "0xD620");
            Assert.True(legacyCells >= (route == 0 ? 8 : 4),
                $"Route {route}: found {legacyCells} legacy cell-border operands.");
        }
    }

    [Theory]
    [InlineData("nil", 0)]
    [InlineData("clear", 0)]
    [InlineData("nil", 1)]
    [InlineData("clear", 1)]
    public void GeneratedDocPreservesConditionalShadingResetsInAllStories(
        string pattern, int cellIndex)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        using var source = new MemoryStream();
        source.Write(File.ReadAllBytes(Path.Combine(directory,
            "WordConditionalFirstRowShadingAllStories.docx")));
        using (var package = WordprocessingDocument.Open(source, true))
        {
            var main = package.MainDocumentPart!;
            var tables = main.Document.Body!.Elements<Table>()
                .Concat(main.HeaderParts.SelectMany(x => x.Header!.Elements<Table>()))
                .Concat(main.FooterParts.SelectMany(x => x.Footer!.Elements<Table>()))
                .ToArray();
            Assert.Equal(3, tables.Length);
            foreach (var table in tables)
            {
                var properties = table.Elements<TableRow>().First()
                    .Elements<TableCell>().ElementAt(cellIndex).TableCellProperties!;
                properties.RemoveAllChildren<Shading>();
                properties.AddChild(new Shading
                {
                    Val = pattern == "nil" ? ShadingPatternValues.Nil :
                        ShadingPatternValues.Clear
                }, true);
            }
            main.Document.Save();
            foreach (var part in main.HeaderParts) part.Header!.Save();
            foreach (var part in main.FooterParts) part.Footer!.Save();
        }
        var generated = DxpDocExport.Export(source.ToArray());
        foreach (var projected in new[]
        {
            DxpDocToDocx.Project(generated).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(generated).DocxBytes)).DocxBytes
        })
        {
            using var input = new MemoryStream(projected);
            using var package = WordprocessingDocument.Open(input, false);
            var main = package.MainDocumentPart!;
            var tables = main.Document.Body!.Elements<Table>()
                .Concat(main.HeaderParts.SelectMany(x => x.Header!.Elements<Table>()))
                .Concat(main.FooterParts.SelectMany(x => x.Footer!.Elements<Table>()));
            Assert.All(tables, table =>
            {
                var shade = table.Elements<TableRow>().First()
                    .Elements<TableCell>().ElementAt(cellIndex).TableCellProperties?.Shading;
                if (pattern == "nil")
                    Assert.Equal("00FF00", shade?.Fill?.Value);
                else
                {
                    Assert.Equal(ShadingPatternValues.Clear, shade?.Val?.Value);
                    Assert.Null(shade?.Fill?.Value);
                }
            });
            Assert.Empty(new OpenXmlValidator().Validate(package));
        }
    }

    [Theory]
    [InlineData("WordConditionalShadingResetsAllStories", false)]
    [InlineData("WordDefaultShadingResetsAllStories", true)]
    public void PairedTableShadingResetsKeepVisibleFillsInAllStories(
        string name, bool shadedSecondRow)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var native = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
        var generated = DxpDocExport.Export(source);
        var routes = new[]
        {
            source,
            DxpDocToDocx.Project(native).DocxBytes,
            DxpDocToDocx.Project(generated).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(generated).DocxBytes)).DocxBytes
        };
        for (var route = 0; route < routes.Length; route++)
        {
            var bytes = routes[route];
            using var stream = new MemoryStream(bytes);
            using var package = WordprocessingDocument.Open(stream, false);
            var main = package.MainDocumentPart!;
            var tables = main.Document.Body!.Elements<Table>()
                .Concat(main.HeaderParts.SelectMany(x => x.Header!.Elements<Table>()))
                .Concat(main.FooterParts.SelectMany(x => x.Footer!.Elements<Table>()))
                .ToArray();
            Assert.Equal(3, tables.Length);
            foreach (var table in tables)
            {
                var cells = table.Elements<TableRow>().First()
                    .Elements<TableCell>().ToArray();
                Assert.Equal(2, cells.Length);
                string? EffectiveFill(TableCell cell, int rowIndex)
                {
                    var direct = cell.TableCellProperties?.Shading;
                    if (direct?.Val?.Value != ShadingPatternValues.Nil)
                    {
                        var fill = direct?.Fill?.Value;
                        if (fill != null && !fill.Equals("auto",
                            StringComparison.OrdinalIgnoreCase))
                            return fill.ToUpperInvariant();
                        if (direct?.Val?.Value == ShadingPatternValues.Clear)
                            return null;
                    }
                    var styleId = table.TableProperties?.TableStyle?.Val?.Value;
                    var style = main.StyleDefinitionsPart?.Styles?
                        .Elements<Style>().FirstOrDefault(x => x.StyleId?.Value == styleId);
                    var inherited = (rowIndex == 0 ? style?
                        .Elements<TableStyleProperties>()
                        .FirstOrDefault(x => x.Type?.Value ==
                            TableStyleOverrideValues.FirstRow)?
                        .Descendants<Shading>().FirstOrDefault()?.Fill?.Value
                        : null) ?? style?.GetFirstChild<StyleTableCellProperties>()?
                            .Shading?.Fill?.Value;
                    return inherited == null || inherited.Equals("auto",
                        StringComparison.OrdinalIgnoreCase)
                        ? null : inherited.ToUpperInvariant();
                }
                Assert.Equal("00FF00", EffectiveFill(cells[0], 0));
                Assert.Null(EffectiveFill(cells[1], 0));
                if (table.Elements<TableRow>().Skip(1).FirstOrDefault() is
                    { } secondRow)
                {
                    var secondRowCells = secondRow.Elements<TableCell>().ToArray();
                    Assert.Equal(2, secondRowCells.Length);
                    foreach (var cell in secondRowCells)
                        Assert.Equal(shadedSecondRow ? "00FF00" : null,
                            EffectiveFill(cell, 1));
                }
            }
            // Word's tblLook boolean attributes require the Office 2010 schema.
            var errors = new OpenXmlValidator(FileFormatVersions.Office2010)
                .Validate(package).ToArray();
            Assert.True(errors.Length == 0, string.Join(" | ", errors.Take(5)
                .Select(x => $"{x.Part?.Uri}: {x.Description}")));
        }
    }

    [Fact]
    public void GeneratedDocClearsExplicitNilParagraphShading()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        using var source = new MemoryStream();
        source.Write(File.ReadAllBytes(Path.Combine(directory,
            "WordStyledParagraphShadingResetAllStories.docx")));
        using (var package = WordprocessingDocument.Open(source, true))
        {
            var header = Assert.Single(package.MainDocumentPart!.HeaderParts).Header!;
            var shade = Assert.Single(header.Descendants<Paragraph>())
                .ParagraphProperties!.Shading!;
            shade.Val = ShadingPatternValues.Nil;
            shade.Fill = null;
            shade.Color = null;
            header.Save();
        }
        var generated = DxpDocExport.Export(source.ToArray());
        using (var input = new MemoryStream(generated))
        using (var index = new DocTextIndexWalker().Index(input))
            Assert.Contains(index.ParagraphStyles, x =>
                x.Formatting is { ShadingPattern: 0, FillRgb: null });
        foreach (var projected in new[]
        {
            DxpDocToDocx.Project(generated).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(generated).DocxBytes)).DocxBytes
        })
        {
            using var input = new MemoryStream(projected);
            using var package = WordprocessingDocument.Open(input, false);
            var header = Assert.Single(package.MainDocumentPart!.HeaderParts).Header!;
            var shade = Assert.Single(header.Descendants<Paragraph>())
                .ParagraphProperties?.Shading;
            Assert.Equal(ShadingPatternValues.Clear, shade?.Val?.Value);
            Assert.Null(shade?.Fill?.Value);
            Assert.Empty(new OpenXmlValidator().Validate(package));
        }
    }

    [Fact]
    public void WordSavedParagraphShadingClearIsIndexedFromBothDocRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordStyledParagraphShadingResetAllStories.docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            "WordStyledParagraphShadingResetAllStories.doc"));
        foreach (var (doc, route) in new[]
            { (native, "native"), (DxpDocExport.Export(source), "generated") })
        {
            using var input = new MemoryStream(doc);
            using var index = new DocTextIndexWalker().Index(input);
            var style = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Mirrored body");
            Assert.Equal(0xF4E0B8u, style.ParagraphFormatting?.FillRgb);
            Assert.True(index.ParagraphStyles.Any(x =>
                x.Formatting is { ShadingPattern: 0, FillRgb: null }),
                route + ": " + string.Join(", ", index.ParagraphStyles
                    .Where(x => x.Formatting?.ShadingPattern != null)
                    .Select(x => $"{x.CpStart}-{x.CpEnd}:" +
                        x.Formatting!.ShadingPattern + "/" +
                        x.Formatting.FillRgb)));
            Assert.Contains(index.ParagraphStyles, x =>
                x.Formatting?.FillRgb == 0xC2E5C5u);
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            using var projectedStream = new MemoryStream(projected);
            using var package = WordprocessingDocument.Open(projectedStream, false);
            var header = Assert.Single(package.MainDocumentPart!.HeaderParts).Header!;
            var clear = Assert.Single(header.Descendants<Paragraph>())
                .ParagraphProperties?.Shading;
            Assert.True(clear?.Val?.Value == ShadingPatternValues.Clear &&
                clear.Fill?.Value is null or "auto",
                route + " header: " + clear?.OuterXml);
            var footer = Assert.Single(package.MainDocumentPart.FooterParts).Footer!;
            Assert.Equal("C5E5C2", Assert.Single(footer.Descendants<Paragraph>())
                .ParagraphProperties?.Shading?.Fill?.Value);
        }
    }

    [Fact]
    public void WordSavedRunShadingClearIsIndexedFromBothDocRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordStyledRunShadingResetAllStories.docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            "WordStyledRunShadingResetAllStories.doc"));
        foreach (var (doc, route) in new[]
            { (DxpDocExport.Export(source), "generated"), (native, "native") })
        {
            using var input = new MemoryStream(doc);
            using var index = new DocTextIndexWalker().Index(input);
            Assert.True(index.CharacterFormatting.Any(x =>
                x.Formatting.Shading is
                    { Pattern: 0, FillRgb: null, ForegroundRgb: null }),
                route + ": " + string.Join(", ", index.CharacterFormatting
                    .Where(x => x.Formatting.Shading != null)
                    .Select(x => $"{x.CpStart}-{x.CpEnd}:" +
                        x.Formatting.Shading)));
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            using var projectedStream = new MemoryStream(projected);
            using var package = WordprocessingDocument.Open(projectedStream, false);
            var headerShade = Assert.Single(package.MainDocumentPart!.HeaderParts)
                .Header!.Descendants<Run>().First().RunProperties?.Shading;
            Assert.Equal(ShadingPatternValues.Clear, headerShade?.Val?.Value);
            Assert.True(headerShade?.Fill?.Value is null or "auto",
                route + " projected header has a fill");
        }
    }

    [Fact]
    public void GeneratedDocClearsExplicitNilRunShading()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var bytes = File.ReadAllBytes(Path.Combine(directory,
            "WordStyledRunShadingResetAllStories.docx"));
        using var source = new MemoryStream();
        source.Write(bytes);
        using (var package = WordprocessingDocument.Open(source, true))
        {
            var header = Assert.Single(package.MainDocumentPart!.HeaderParts).Header!;
            var shading = Assert.Single(header.Descendants<Run>())
                .RunProperties!.Shading!;
            shading.Val = ShadingPatternValues.Nil;
            shading.Fill = null;
            shading.Color = null;
            header.Save();
        }
        var generated = DxpDocExport.Export(source.ToArray());
        using (var input = new MemoryStream(generated))
        using (var index = new DocTextIndexWalker().Index(input))
            Assert.Contains(index.CharacterFormatting, x =>
                x.Formatting.Shading is
                    { Pattern: 0, FillRgb: null, ForegroundRgb: null });
        foreach (var projected in new[]
        {
            DxpDocToDocx.Project(generated).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(generated).DocxBytes)).DocxBytes
        })
        {
            using var input = new MemoryStream(projected);
            using var package = WordprocessingDocument.Open(input, false);
            var header = Assert.Single(package.MainDocumentPart!.HeaderParts).Header!;
            var clear = Assert.Single(header.Descendants<Run>())
                .RunProperties?.Shading;
            Assert.Equal(ShadingPatternValues.Clear, clear?.Val?.Value);
            Assert.Null(clear?.Fill?.Value);
            Assert.Empty(new OpenXmlValidator().Validate(package));
        }
    }

    [Fact]
    public void WordSavedStyledRgbRunShadingSurvivesBothRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordStyledRgbRunShadingAllStories.docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            "WordStyledRgbRunShadingAllStories.doc"));
        AssertStyledRgbRunShading(source);
        foreach (var doc in new[] { native, DxpDocExport.Export(source) })
        {
            using (var input = new MemoryStream(doc))
            using (var index = new DocTextIndexWalker().Index(input))
            {
                var style = Assert.Single(index.StyleDefinitions,
                    x => x.Name == "Mirrored body");
                Assert.Equal(0xF4E0B8u,
                    style.CharacterFormatting.Shading?.FillRgb);
                Assert.Contains(index.CharacterFormatting, x =>
                    x.Formatting.Shading?.FillRgb == 0xA8D6F4u);
                Assert.Contains(index.CharacterFormatting, x =>
                    x.Formatting.Shading?.FillRgb == 0xC2E5C5u);
            }
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            AssertStyledRgbRunShading(projected);
            AssertStyledRgbRunShading(DxpDocToDocx.Project(
                DxpDocExport.Export(projected)).DocxBytes);
        }
    }

    private static void AssertStyledRgbRunShading(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var style = Assert.Single(main.StyleDefinitionsPart!.Styles!
            .Elements<Style>(), x => x.StyleName?.Val?.Value == "Mirrored body");
        Assert.Equal("B8E0F4", style.StyleRunProperties?.Shading?.Fill?.Value);
        foreach (var (story, directFill) in new
            (OpenXmlElement Story, string? Fill)[]
        {
            (main.Document!.Body!, null),
            (Assert.Single(main.HeaderParts).Header!, "F4D6A8"),
            (Assert.Single(main.FooterParts).Footer!, "C5E5C2")
        })
        {
            var paragraph = Assert.Single(story.Descendants<Paragraph>(), x =>
                x.InnerText.StartsWith("Mirrored ", StringComparison.Ordinal));
            Assert.Equal(style.StyleId?.Value,
                paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
            Assert.Equal(directFill, Assert.Single(paragraph.Elements<Run>())
                .RunProperties?.Shading?.Fill?.Value);
        }
    }

    [Fact]
    public void WordSavedRgbRunShadingSurvivesBothRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordRgbRunShadingAllStories.docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            "WordRgbRunShadingAllStories.doc"));
        AssertRgbRunShading(source);
        foreach (var doc in new[] { native, DxpDocExport.Export(source) })
        {
            using (var input = new MemoryStream(doc))
            using (var index = new DocTextIndexWalker().Index(input))
                foreach (var color in new[] { 0xF4E0B8u, 0xA8D6F4u, 0xC2E5C5u })
                    Assert.Contains(index.CharacterFormatting, x =>
                        x.Formatting.Shading?.FillRgb == color);
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            AssertRgbRunShading(projected);
            AssertRgbRunShading(DxpDocToDocx.Project(
                DxpDocExport.Export(projected)).DocxBytes);
        }
    }

    private static void AssertRgbRunShading(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        foreach (var (story, expected) in new (OpenXmlElement Story, string Color)[]
        {
            (main.Document!.Body!, "B8E0F4"),
            (Assert.Single(main.HeaderParts).Header!, "F4D6A8"),
            (Assert.Single(main.FooterParts).Footer!, "C5E5C2")
        })
        {
            var paragraph = Assert.Single(story.Descendants<Paragraph>(), x =>
                x.InnerText.StartsWith("Mirrored ", StringComparison.Ordinal));
            var shading = Assert.IsType<Shading>(Assert.Single(
                paragraph.Elements<Run>()).RunProperties?.Shading);
            Assert.Equal(expected, shading.Fill?.Value);
            Assert.Equal(ShadingPatternValues.Clear, shading.Val?.Value);
        }
    }

    [Fact]
    public void WordSavedUnderlineColorStyleOverridesSurviveBothRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordUnderlineColorStyleOverridesAllStories.docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            "WordUnderlineColorStyleOverridesAllStories.doc"));
        AssertUnderlineColorStyleOverrides(source);
        foreach (var doc in new[] { native, DxpDocExport.Export(source) })
        {
            using (var input = new MemoryStream(doc))
            using (var index = new DocTextIndexWalker().Index(input))
            {
                var style = Assert.Single(index.StyleDefinitions,
                    x => x.Name == "Mirrored body");
                Assert.Equal(0x2836C9u,
                    style.CharacterFormatting.UnderlineColorRef);
                Assert.Contains(index.CharacterFormatting, x =>
                    x.Formatting.UnderlineColorRef == 0x6F9D2Bu);
                Assert.Contains(index.CharacterFormatting, x =>
                    x.Formatting.UnderlineColorRef == 0xCC6633u);
            }
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            AssertUnderlineColorStyleOverrides(projected);
            AssertUnderlineColorStyleOverrides(DxpDocToDocx.Project(
                DxpDocExport.Export(projected)).DocxBytes);
        }
    }

    private static void AssertUnderlineColorStyleOverrides(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var style = Assert.Single(main.StyleDefinitionsPart!.Styles!
            .Elements<Style>(), x => x.StyleName?.Val?.Value == "Mirrored body");
        var underline = Assert.IsType<Underline>(style.StyleRunProperties?
            .GetFirstChild<Underline>());
        Assert.Equal(UnderlineValues.Single, underline.Val?.Value);
        Assert.Equal("C93628", underline.Color?.Value);
        foreach (var (story, expectedValue, expectedColor) in new
            (OpenXmlElement Story, UnderlineValues? Value, string? Color)[]
        {
            (main.Document!.Body!, null, null),
            (Assert.Single(main.HeaderParts).Header!, UnderlineValues.Double,
                "2B9D6F"),
            (Assert.Single(main.FooterParts).Footer!, UnderlineValues.Single,
                "3366CC")
        })
        {
            var paragraph = Assert.Single(story.Descendants<Paragraph>(), x =>
                x.InnerText.StartsWith("Mirrored ", StringComparison.Ordinal));
            Assert.Equal(style.StyleId?.Value,
                paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
            var direct = Assert.Single(paragraph.Elements<Run>())
                .RunProperties?.GetFirstChild<Underline>();
            if (story is Footer)
                Assert.True(direct?.Val?.Value == null ||
                    direct.Val.Value == expectedValue);
            else
                Assert.Equal(expectedValue, direct?.Val?.Value);
            Assert.Equal(expectedColor, direct?.Color?.Value);
        }
    }

    [Fact]
    public void WordSavedCharacterSpacingStyleOverridesSurviveBothRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordCharacterSpacingStyleOverridesAllStories.docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            "WordCharacterSpacingStyleOverridesAllStories.doc"));
        AssertCharacterSpacingStyleOverrides(source);
        foreach (var doc in new[] { native, DxpDocExport.Export(source) })
        {
            using (var input = new MemoryStream(doc))
            using (var index = new DocTextIndexWalker().Index(input))
            {
                var style = Assert.Single(index.StyleDefinitions,
                    x => x.Name == "Mirrored body");
                Assert.Equal((short)40,
                    style.CharacterFormatting.CharacterSpacingTwips);
                Assert.Equal((ushort)1033, style.CharacterFormatting.LanguageId);
                Assert.Equal((ushort)1041,
                    style.CharacterFormatting.EastAsiaLanguageId);
                Assert.Equal((ushort)1025,
                    style.CharacterFormatting.ComplexScriptLanguageId);
                Assert.Contains(index.CharacterFormatting, x =>
                    x.Formatting.CharacterSpacingTwips == -20);
                Assert.Contains(index.CharacterFormatting, x =>
                    x.Formatting.CharacterSpacingTwips == 0);
                Assert.Contains(index.CharacterFormatting, x =>
                    x.Formatting.LanguageId == 3084 &&
                    x.Formatting.EastAsiaLanguageId == 2052 &&
                    x.Formatting.ComplexScriptLanguageId == 1037);
                Assert.Contains(index.CharacterFormatting, x =>
                    x.Formatting.LanguageId == 2057 &&
                    x.Formatting.EastAsiaLanguageId == 1042 &&
                    x.Formatting.ComplexScriptLanguageId == 1065);
            }
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            AssertCharacterSpacingStyleOverrides(projected);
            AssertCharacterSpacingStyleOverrides(DxpDocToDocx.Project(
                DxpDocExport.Export(projected)).DocxBytes);
        }
    }

    private static void AssertCharacterSpacingStyleOverrides(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var style = Assert.Single(main.StyleDefinitionsPart!.Styles!
            .Elements<Style>(), x => x.StyleName?.Val?.Value == "Mirrored body");
        Assert.Equal(40, style.StyleRunProperties?
            .GetFirstChild<Spacing>()?.Val?.Value);
        var styleLanguage = Assert.IsType<Languages>(style.StyleRunProperties?
            .GetFirstChild<Languages>());
        Assert.Equal("en-US", styleLanguage.Val?.Value);
        Assert.Equal("ja-JP", styleLanguage.EastAsia?.Value);
        Assert.Equal("ar-SA", styleLanguage.Bidi?.Value);
        foreach (var (story, expected, languages) in new
            (OpenXmlElement Story, int? Direct, string[]? Languages)[]
        {
            (main.Document!.Body!, null, null),
            (Assert.Single(main.HeaderParts).Header!, -20,
                ["fr-CA", "zh-CN", "he-IL"]),
            (Assert.Single(main.FooterParts).Footer!, 0,
                ["en-GB", "ko-KR", "fa-IR"])
        })
        {
            var paragraph = Assert.Single(story.Descendants<Paragraph>(), x =>
                x.InnerText.StartsWith("Mirrored ", StringComparison.Ordinal));
            Assert.Equal(style.StyleId?.Value,
                paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
            var run = Assert.Single(paragraph.Elements<Run>());
            Assert.Equal(expected,
                run.RunProperties?.GetFirstChild<Spacing>()?.Val?.Value);
            var direct = run.RunProperties?.GetFirstChild<Languages>();
            if (languages == null) Assert.Null(direct);
            else
            {
                Assert.Equal(languages[0], direct?.Val?.Value);
                Assert.Equal(languages[1], direct?.EastAsia?.Value);
                Assert.Equal(languages[2], direct?.Bidi?.Value);
            }
        }
    }

    [Fact]
    public void WordSavedSixSlotInheritedRunBordersSurviveBothRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordSixSlotRunBorderInheritance.docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            "WordSixSlotRunBorderInheritance.doc"));
        AssertSixSlotInheritedRunBorders(source);
        foreach (var doc in new[] { native, DxpDocExport.Export(source) })
        {
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            AssertSixSlotInheritedRunBorders(projected);
            AssertSixSlotInheritedRunBorders(DxpDocToDocx.Project(
                DxpDocExport.Export(projected)).DocxBytes);
        }
    }

    private static void AssertSixSlotInheritedRunBorders(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var styles = main.StyleDefinitionsPart!.Styles!;
        var baseStyle = Assert.Single(styles.Elements<Style>(), x =>
            x.StyleName?.Val?.Value == "Six Story Base");
        var derived = Assert.Single(styles.Elements<Style>(), x =>
            x.StyleName?.Val?.Value == "Six Story Derived");
        Assert.Equal(baseStyle.StyleId?.Value, derived.BasedOn?.Val?.Value);
        var border = Assert.IsType<Border>(baseStyle.StyleRunProperties?
            .GetFirstChild<Border>());
        Assert.Equal(BorderValues.Single, border.Val?.Value);
        Assert.Equal("2B9D6F", border.Color?.Value);
        var stories = main.HeaderParts.Select(x => (OpenXmlElement)x.Header!)
            .Concat(main.FooterParts.Select(x => (OpenXmlElement)x.Footer!))
            .ToArray();
        Assert.Equal(12, stories.Length);
        var resetText = new[]
        {
            "First section header", "First-page section 1 footer",
            "First-page section 2 header"
        };
        foreach (var story in stories)
        {
            var paragraph = Assert.Single(story.Descendants<Paragraph>());
            Assert.Equal(derived.StyleId?.Value,
                paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
            var run = Assert.Single(paragraph.Elements<Run>());
            var direct = run.RunProperties?.GetFirstChild<Border>();
            if (resetText.Contains(paragraph.InnerText))
                Assert.True(direct?.Val?.Value == BorderValues.Nil ||
                    direct?.Val?.Value == BorderValues.None);
            else
                Assert.Null(direct);
        }
    }

    [Fact]
    public void WordSavedLayeredRunBordersSurviveBothRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordLayeredRunBordersAllStories.docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            "WordLayeredRunBordersAllStories.doc"));
        AssertLayeredRunBorders(source);
        foreach (var doc in new[] { native, DxpDocExport.Export(source) })
        {
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            AssertLayeredRunBorders(projected);
            AssertLayeredRunBorders(DxpDocToDocx.Project(
                DxpDocExport.Export(projected)).DocxBytes);
        }
    }

    private static void AssertLayeredRunBorders(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var styles = main.StyleDefinitionsPart!.Styles!;
        var paragraphStyle = Assert.Single(styles.Elements<Style>(), x =>
            x.StyleName?.Val?.Value == "Mirrored body");
        var characterStyle = Assert.Single(styles.Elements<Style>(), x =>
            x.StyleName?.Val?.Value == "Layered border");
        Assert.Equal("2B9D6F", paragraphStyle.StyleRunProperties?
            .GetFirstChild<Border>()?.Color?.Value);
        var characterBorder = Assert.IsType<Border>(characterStyle
            .StyleRunProperties?.GetFirstChild<Border>());
        Assert.Equal(BorderValues.Double, characterBorder.Val?.Value);
        Assert.Equal("C93628", characterBorder.Color?.Value);
        foreach (var story in new OpenXmlElement[]
        {
            main.Document!.Body!, Assert.Single(main.HeaderParts).Header!,
            Assert.Single(main.FooterParts).Footer!
        })
        {
            var paragraph = Assert.Single(story.Descendants<Paragraph>(), x =>
                x.InnerText.StartsWith("Mirrored ", StringComparison.Ordinal));
            Assert.Equal(paragraphStyle.StyleId?.Value,
                paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
            var run = Assert.Single(paragraph.Elements<Run>());
            Assert.Equal(characterStyle.StyleId?.Value,
                run.RunProperties?.RunStyle?.Val?.Value);
            var direct = run.RunProperties?.GetFirstChild<Border>();
            if (story is Header)
                Assert.True(direct?.Val?.Value == BorderValues.Nil ||
                    direct?.Val?.Value == BorderValues.None);
            else
                Assert.Null(direct);
        }
    }

    [Fact]
    public void WordSavedParagraphStyleRunBorderResetSurvivesBothRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordParagraphStyleRunBorderResetAllStories.docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            "WordParagraphStyleRunBorderResetAllStories.doc"));
        AssertParagraphStyleRunBorderReset(source);
        foreach (var doc in new[] { native, DxpDocExport.Export(source) })
        {
            using (var input = new MemoryStream(doc))
            using (var index = new DocTextIndexWalker().Index(input))
            {
                var style = Assert.Single(index.StyleDefinitions,
                    x => x.Name == "Mirrored body");
                Assert.Equal(new DocParagraphBorder(1, 12, 2, 0x6F9D2B),
                    style.CharacterFormatting.Border);
                Assert.Contains(index.CharacterFormatting, x =>
                    x.Formatting.Border?.Type == 0);
            }
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            AssertParagraphStyleRunBorderReset(projected);
            AssertParagraphStyleRunBorderReset(DxpDocToDocx.Project(
                DxpDocExport.Export(projected)).DocxBytes);
        }
    }

    private static void AssertParagraphStyleRunBorderReset(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var style = Assert.Single(main.StyleDefinitionsPart!.Styles!
            .Elements<Style>(), x => x.StyleName?.Val?.Value == "Mirrored body");
        var border = Assert.IsType<Border>(style.StyleRunProperties?
            .GetFirstChild<Border>());
        Assert.Equal(BorderValues.Single, border.Val?.Value);
        Assert.Equal("2B9D6F", border.Color?.Value);
        foreach (var story in new OpenXmlElement[]
        {
            main.Document!.Body!, Assert.Single(main.HeaderParts).Header!,
            Assert.Single(main.FooterParts).Footer!
        })
        {
            var paragraph = Assert.Single(story.Descendants<Paragraph>(), x =>
                x.InnerText.StartsWith("Mirrored ", StringComparison.Ordinal));
            Assert.Equal(style.StyleId?.Value,
                paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
            var run = Assert.Single(paragraph.Elements<Run>());
            var direct = run.RunProperties?.GetFirstChild<Border>();
            if (story is Header)
                Assert.True(direct?.Val?.Value == BorderValues.Nil ||
                    direct?.Val?.Value == BorderValues.None);
            else
                Assert.Null(direct);
        }
    }

    [Fact]
    public void WordSavedRgbShadowRunBordersSurviveBothRoutesAndThirdHop()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordRgbShadowRunBordersAllStories.docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            "WordRgbShadowRunBordersAllStories.doc"));
        AssertRgbShadowRunBorders(source);
        foreach (var doc in new[] { native, DxpDocExport.Export(source) })
        {
            using (var input = new MemoryStream(doc))
            using (var index = new DocTextIndexWalker().Index(input))
                Assert.True(index.CharacterFormatting.Count(x =>
                    x.Formatting.Border is { Type: 1, WidthEighthPoints: 12,
                        SpacePoints: 2, ColorRgb: 0x6F9D2B,
                        Shadow: true }) >= 3);
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            AssertRgbShadowRunBorders(projected);
            AssertRgbShadowRunBorders(DxpDocToDocx.Project(
                DxpDocExport.Export(projected)).DocxBytes);
        }
    }

    private static void AssertRgbShadowRunBorders(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        void Check(OpenXmlElement? root)
        {
            var border = Assert.Single(root!.Descendants<Run>()
                .Select(x => x.RunProperties?.GetFirstChild<Border>())
                .Where(x => x != null));
            Assert.Equal(BorderValues.Single, border.Val?.Value);
            Assert.Equal((uint)12, border.Size?.Value);
            Assert.Equal((uint)2, border.Space?.Value);
            Assert.Equal("2B9D6F", border.Color?.Value);
            Assert.True(border.Shadow?.Value);
        }
        Check(main.Document?.Body);
        Check(Assert.Single(main.HeaderParts).Header);
        Check(Assert.Single(main.FooterParts).Footer);
    }

    [Fact]
    public void WordSavedInheritedRunBorderResetSurvivesBothRoutesAndThirdHop()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordInheritedRunBorderResetAllStories.docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            "WordInheritedRunBorderResetAllStories.doc"));
        AssertDirectRunBorderReset(source);
        foreach (var doc in new[] { native, DxpDocExport.Export(source) })
        {
            using (var input = new MemoryStream(doc))
            using (var index = new DocTextIndexWalker().Index(input))
                Assert.True(index.CharacterFormatting.Count(x =>
                    x.Formatting.Border?.Type == 0) >= 3);
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            AssertDirectRunBorderReset(projected);
            AssertDirectRunBorderReset(DxpDocToDocx.Project(
                DxpDocExport.Export(projected)).DocxBytes);
        }
    }

    private static void AssertDirectRunBorderReset(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        void Check(OpenXmlElement? root)
        {
            var run = Assert.Single(root!.Descendants<Run>(), x =>
                x.InnerText.EndsWith(" Plain", StringComparison.Ordinal));
            var border = Assert.IsType<Border>(run.RunProperties?
                .GetFirstChild<Border>());
            Assert.True(border.Val?.Value == BorderValues.Nil ||
                border.Val?.Value == BorderValues.None);
            var styleId = Assert.Single(main.StyleDefinitionsPart!.Styles!
                .Elements<Style>(), x => x.StyleName?.Val?.Value ==
                    "Text Frame Emphasis").StyleId!.Value;
            Assert.Equal(styleId, run.RunProperties?.RunStyle?.Val?.Value);
        }
        Check(main.Document?.Body);
        Check(Assert.Single(main.HeaderParts).Header);
        Check(Assert.Single(main.FooterParts).Footer);
    }

    [Fact]
    public void WordSavedDerivedCharacterStyleInheritsRunBordersThroughBothRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordDerivedCharacterStyleRunBordersAllStories.docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            "WordDerivedCharacterStyleRunBordersAllStories.doc"));
        AssertTextFrameStyleBorder(source, "Text Frame Emphasis");
        foreach (var (doc, expectBasedOn) in new[]
        {
            (Doc: native, ExpectBasedOn: false),
            (Doc: DxpDocExport.Export(source), ExpectBasedOn: true)
        })
        {
            using (var input = new MemoryStream(doc))
            using (var index = new DocTextIndexWalker().Index(input))
            {
                var parent = Assert.Single(index.StyleDefinitions,
                    x => x.Name == "Text Frame");
                var child = Assert.Single(index.StyleDefinitions,
                    x => x.Name == "Text Frame Emphasis");
                if (expectBasedOn)
                    Assert.Equal(parent.Index, child.BasedOnIndex);
                else
                    Assert.Equal(new DocParagraphBorder(1, 12, 2, 0x0000FF),
                        child.DirectCharacterFormatting?.Border);
                Assert.Equal(new DocParagraphBorder(1, 12, 2, 0x0000FF),
                    child.CharacterFormatting.Border);
                Assert.True(child.CharacterFormatting.Italic);
            }
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            AssertTextFrameStyleBorder(projected, "Text Frame Emphasis",
                expectBasedOn);
            var repeated = DxpDocToDocx.Project(DxpDocExport.Export(projected))
                .DocxBytes;
            AssertTextFrameStyleBorder(repeated, "Text Frame Emphasis",
                expectBasedOn);
        }
    }

    [Fact]
    public void WordSavedCharacterStyleRunBordersSurviveBothRoutesAndThirdHop()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordCharacterStyleRunBordersAllStories.docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            "WordCharacterStyleRunBordersAllStories.doc"));
        AssertTextFrameStyleBorder(source);
        foreach (var doc in new[] { native, DxpDocExport.Export(source) })
        {
            using (var input = new MemoryStream(doc))
            using (var index = new DocTextIndexWalker().Index(input))
            {
                var style = Assert.Single(index.StyleDefinitions,
                    x => x.Name == "Text Frame");
                Assert.Equal(new DocParagraphBorder(1, 12, 2, 0x0000FF),
                    style.CharacterFormatting.Border);
            }
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            AssertTextFrameStyleBorder(projected);
            var repeated = DxpDocToDocx.Project(DxpDocExport.Export(projected))
                .DocxBytes;
            AssertTextFrameStyleBorder(repeated);
        }
    }

    private static void AssertTextFrameStyleBorder(byte[] bytes,
        string appliedStyleName = "Text Frame", bool expectBasedOn = true)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var style = Assert.Single(main.StyleDefinitionsPart!.Styles!
            .Elements<Style>(), x => x.StyleName?.Val?.Value == "Text Frame");
        var styleId = style.StyleId!.Value;
        var border = Assert.IsType<Border>(style.StyleRunProperties?
            .GetFirstChild<Border>());
        Assert.Equal(BorderValues.Single, border.Val?.Value);
        Assert.Equal((uint)12, border.Size?.Value);
        Assert.Equal((uint)2, border.Space?.Value);
        Assert.Equal("FF0000", border.Color?.Value);
        var appliedStyle = Assert.Single(main.StyleDefinitionsPart.Styles
            .Elements<Style>(), x => x.StyleName?.Val?.Value == appliedStyleName);
        if (appliedStyleName != "Text Frame")
        {
            if (expectBasedOn)
                Assert.Equal(style.StyleId?.Value, appliedStyle.BasedOn?.Val?.Value);
            else
                Assert.Equal("FF0000", appliedStyle.StyleRunProperties?
                    .GetFirstChild<Border>()?.Color?.Value);
            Assert.True(appliedStyle.StyleRunProperties?.Italic?.Val?.Value ?? true);
        }
        styleId = appliedStyle.StyleId!.Value;
        Assert.Equal(1, main.Document!.Body!.Descendants<RunStyle>()
            .Count(x => x.Val?.Value == styleId));
        Assert.Equal(1, main.HeaderParts.SelectMany(x => x.Header!
            .Descendants<RunStyle>()).Count(x => x.Val?.Value == styleId));
        Assert.Equal(1, main.FooterParts.SelectMany(x => x.Footer!
            .Descendants<RunStyle>()).Count(x => x.Val?.Value == styleId));
    }

    [Fact]
    public void WordSavedRunBordersSurviveBothRoutesAndThirdHop()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordRunBordersAllStories.docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            "WordRunBordersAllStories.doc"));
        var expected = new[]
        {
            "body:single:12:2:FF0000",
            "header:single:12:2:FF0000",
            "footer:single:12:2:FF0000"
        };
        Assert.Equal(expected, ReadRunBorders(source));
        foreach (var doc in new[] { native, DxpDocExport.Export(source) })
        {
            using (var input = new MemoryStream(doc))
            using (var index = new DocTextIndexWalker().Index(input))
                Assert.True(index.CharacterFormatting.Count(x =>
                    x.Formatting.Border is { Type: 1, WidthEighthPoints: 12,
                        SpacePoints: 2 }) >= 3);
            var projected = DxpDocToDocx.Project(doc);
            Assert.Equal(expected, ReadRunBorders(projected.DocxBytes));
            var rewritten = DxpDocExport.Export(projected.DocxBytes);
            Assert.Equal(expected, ReadRunBorders(DxpDocToDocx.Project(rewritten)
                .DocxBytes));
        }
    }

    private static IReadOnlyList<string> ReadRunBorders(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var result = new List<string>();
        void Add(string story, OpenXmlElement? root)
        {
            if (root == null) return;
            foreach (var border in root.Descendants<Run>()
                .Select(run => run.RunProperties?.GetFirstChild<Border>())
                .Where(border => border != null))
                result.Add($"{story}:{border!.Val}:{border.Size}:{border.Space}:" +
                    border.Color?.Value);
        }
        Add("body", main.Document?.Body);
        foreach (var header in main.HeaderParts) Add("header", header.Header);
        foreach (var footer in main.FooterParts) Add("footer", footer.Footer);
        return result;
    }

    [Fact]
    public void WordSavedLegacyShadingPercentagesSurviveBothRoutes()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordLegacyShadingPercentages.doc"));
        var native = File.ReadAllBytes(path);
        var generated = DxpDocExport.Export(File.ReadAllBytes(
            Path.ChangeExtension(path, ".docx")));
        using (var source = new MemoryStream(native))
        using (var structure = new DocStructureWalker().Accept(source,
            new DocStructurePrintVisitor(TextWriter.Null)))
        {
            static IEnumerable<DocStructureNode> Descendants(DocStructureNode node) =>
                node.Children.SelectMany(child => new[] { child }.Concat(Descendants(child)));
            Assert.True(Descendants(structure.Root).Count(node =>
                node.Kind == "Chpx" && Descendants(node).Any(modifier =>
                    modifier.Kind == "Sprm" &&
                    modifier.Attributes.TryGetValue("code", out var code) &&
                    code == "0x4866")) >= 9);
        }
        foreach (var input in new[] { native, generated })
        foreach (var doc in new[] { input,
            DxpDocExport.Export(DxpDocToDocx.Project(input).DocxBytes) })
        {
            using var index = new DocTextIndexWalker().Index(new MemoryStream(doc));
            var patterns = index.CharacterFormatting.Where(x =>
                x.Formatting.Shading?.FillRgb == 0x00FFFFu &&
                x.Formatting.Shading.ForegroundRgb == 0x0000FFu)
                .Select(x => x.Formatting.Shading!.Pattern).ToArray();
            Assert.Equal(new ushort[] { 5, 8, 11, 5, 8, 11, 5, 8, 11 }, patterns);
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            using var stream = new MemoryStream(projected);
            using var package = WordprocessingDocument.Open(stream, false);
            Assert.Empty(new OpenXmlValidator().Validate(package));
        }
    }

    [Theory]
    [InlineData("WordLegacyShadedRun", 1)]
    [InlineData("WordPatternedLegacyRun", 8)]
    public void WordSavedDirectRunLegacyShadingSurvivesBothDocHops(
        string fixtureName, int pattern)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            fixtureName + ".doc"));
        var native = File.ReadAllBytes(path);
        var generated = DxpDocExport.Export(File.ReadAllBytes(
            Path.ChangeExtension(path, ".docx")));
        using (var input = new MemoryStream(native))
        using (var structure = new DocStructureWalker().Accept(input,
            new DocStructurePrintVisitor(TextWriter.Null)))
        {
            static IEnumerable<DocStructureNode> Descendants(DocStructureNode node) =>
                node.Children.SelectMany(child => new[] { child }.Concat(Descendants(child)));
            var legacyRunCount = Descendants(structure.Root).Count(node =>
                node.Kind == "Chpx" && Descendants(node).Any(modifier =>
                    modifier.Kind == "Sprm" &&
                    modifier.Attributes.TryGetValue("code", out var code) &&
                    code == "0x4866"));
            Assert.True(legacyRunCount >= 3, $"Found {legacyRunCount} legacy CHPX runs");
        }
        foreach (var doc in new[] { native, generated })
        {
            using var index = new DocTextIndexWalker().Index(new MemoryStream(doc));
            var matchingRuns = index.CharacterFormatting.Count(x =>
                x.Formatting.Shading?.FillRgb == 0x00FFFFu &&
                x.Formatting.Shading.ForegroundRgb == 0x0000FFu &&
                x.Formatting.Shading.Pattern == pattern);
            Assert.True(matchingRuns >= 3, $"Found {matchingRuns}: " +
                string.Join("; ", index.CharacterFormatting.Select(x =>
                    $"{x.CpStart}-{x.CpEnd} {x.Formatting.Shading}")));
            foreach (var projected in new[] { DxpDocToDocx.Project(doc).DocxBytes,
                DxpDocToDocx.Project(DxpDocExport.Export(
                    DxpDocToDocx.Project(doc).DocxBytes)).DocxBytes })
            {
                using var stream = new MemoryStream(projected);
                using var document = WordprocessingDocument.Open(stream, false);
                var main = document.MainDocumentPart!;
                foreach (var story in new[]
                {
                    new OpenXmlElement[] { main.Document.Body! },
                    main.HeaderParts.Select(part => (OpenXmlElement)part.Header!).ToArray(),
                    main.FooterParts.Select(part => (OpenXmlElement)part.Footer!).ToArray()
                })
                    Assert.Contains(story.SelectMany(root => root.Descendants<Run>()), run =>
                        run.RunProperties?.Shading?.Fill?.Value == "FFFF00" &&
                        run.RunProperties.Shading.Color?.Value == "FF0000" &&
                        run.RunProperties.Shading.Val?.Value ==
                            (pattern == 1 ? ShadingPatternValues.Solid :
                                ShadingPatternValues.Percent50));
                Assert.Empty(new OpenXmlValidator().Validate(document));
            }
        }
    }

    [Fact]
    public void WordSavedCharacterStyleLegacyShadingSurvivesDocProjectionAndThirdHop()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var nativePath = Path.Combine(directory, "WordStyledParagraphDecorations.doc");
        using (var input = File.OpenRead(nativePath))
        using (var structure = new DocStructureWalker().Accept(input,
            new DocStructurePrintVisitor(TextWriter.Null)))
        {
            static IEnumerable<DocStructureNode> Descendants(DocStructureNode node) =>
                node.Children.SelectMany(child => new[] { child }.Concat(Descendants(child)));
            Assert.Contains(Descendants(structure.Root), node =>
                node.Kind == "UpxChpx" && Descendants(node).Any(modifier =>
                    modifier.Kind == "Sprm" &&
                    modifier.Attributes.TryGetValue("code", out var code) &&
                    code == "0x4866"));
        }
        var native = File.ReadAllBytes(nativePath);
        var generated = DxpDocExport.Export(File.ReadAllBytes(
            Path.ChangeExtension(nativePath, ".docx")));
        foreach (var doc in new[] { native, generated })
        {
            using var index = new DocTextIndexWalker().Index(new MemoryStream(doc));
            var style = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Decorated Note Char");
            Assert.Equal(0x00FFFFu, style.CharacterFormatting.Shading?.FillRgb);
            Assert.Equal((ushort)0, style.CharacterFormatting.Shading?.Pattern);
        }
        foreach (var doc in new[] { native, generated })
        foreach (var bytes in new[] { DxpDocToDocx.Project(doc).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(doc).DocxBytes)).DocxBytes })
        {
            using var stream = new MemoryStream(bytes);
            using var document = WordprocessingDocument.Open(stream, false);
            var style = Assert.Single(document.MainDocumentPart!.StyleDefinitionsPart!
                .Styles!.Elements<Style>(), x => x.StyleName?.Val?.Value ==
                    "Decorated Note Char");
            Assert.Equal("FFFF00", style.StyleRunProperties?.Shading?.Fill?.Value);
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
    }

    [Theory]
    [InlineData("WordVisibleNoWrapOffAllStories")]
    [InlineData("WordVisibleNoWrapAllStories")]
    public void WordNativeAutoWidthGeometrySurvivesDocxAndDocHop(string name)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var native = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
        static short[] FirstCellEdges(byte[] bytes)
        {
            using var input = new MemoryStream(bytes);
            using var index = new DocTextIndexWalker().Index(input);
            return index.ParagraphStyles.Where(x => x.Formatting?.TableTerminator == true)
                .Select(x => x.Formatting!.TableCellEdges![1]).ToArray();
        }
        var projected = DxpDocToDocx.Project(native).DocxBytes;
        var rewritten = DxpDocExport.Export(projected);
        Assert.Equal(FirstCellEdges(native), FirstCellEdges(rewritten));
    }

    [Fact]
    public void DocxVisitorRebuildsEveryExplicitHeaderStoryFromBinaryDoc()
    {
        var inputPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordArbitraryRotationStories.doc"));
        var outputPath = Path.Combine(Path.GetTempPath(), $"docxport-headers-{Guid.NewGuid():N}.docx");
        try
        {
            DxpExport.ExportToFile(inputPath,
                new DocxportNet.Visitors.Docx.DxpDocxVisitor(), outputPath);
            using var output = WordprocessingDocument.Open(outputPath, false);
            var headers = output.MainDocumentPart!.HeaderParts
                .Select(part => part.Header!).ToArray();
            Assert.Single(headers.SelectMany(header => header.Descendants<Drawing>()));
            Assert.Empty(new OpenXmlValidator().Validate(output));
        }
        finally
        {
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void EvaluatedDocxExportKeepsHyperlinkFieldFormatting()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var inputPath = Path.Combine(directory, "WordFieldsAndLink.doc");
        var outputPath = Path.Combine(Path.GetTempPath(), $"docxport-links-{Guid.NewGuid():N}.docx");
        try
        {
            var projected = DxpDocToDocx.Project(File.ReadAllBytes(inputPath)).DocxBytes;
            DxpExport.ExportToFile(inputPath,
                new DocxportNet.Visitors.Docx.DxpDocxVisitor(), outputPath);
            var exported = File.ReadAllBytes(outputPath);
            AssertEffectiveRunFormatting(projected, exported);
            Assert.Equal(ReadEditableFieldSemantics(projected),
                ReadEditableFieldSemantics(exported));
            Assert.Empty(Validate(exported));
        }
        finally
        {
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Theory]
    [MemberData(nameof(PairedFixtures))]
    public void EvaluatedDocxExportRetainsEditableFieldsFromBinaryDoc(string name)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var inputPath = Path.Combine(directory, name + ".doc");
        var outputPath = Path.Combine(Path.GetTempPath(), $"docxport-fields-{Guid.NewGuid():N}.docx");
        try
        {
            var projected = DxpDocToDocx.Project(File.ReadAllBytes(inputPath)).DocxBytes;
            DxpExport.ExportToFile(inputPath,
                new DocxportNet.Visitors.Docx.DxpDocxVisitor(), outputPath);
            var exported = File.ReadAllBytes(outputPath);
            Assert.Equal(ReadStories(projected), ReadStories(exported));
            Assert.Equal(ReadSupportedCoreProperties(projected),
                ReadSupportedCoreProperties(exported));
            Assert.Equal(ReadDocumentStatistics(projected),
                ReadDocumentStatistics(exported));
            Assert.Equal(ReadEffectiveHeaderFooterStories(projected),
                ReadEffectiveHeaderFooterStories(exported));
            AssertExplicitSectionLayout(projected, exported);
            AssertEffectiveParagraphLayout(projected, exported);
            AssertParagraphDecorations(projected, exported);
            AssertEffectiveTabStops(projected, exported);
            AssertEffectiveRunFormatting(projected, exported);
            AssertDrawingGeometry(projected, exported);
            AssertTableGrid(projected, exported);
            Assert.Equal(ReadListSemantics(projected), ReadListSemantics(exported));
            Assert.Equal(ReadListSemantics(projected, true),
                ReadListSemantics(exported, true));
            Assert.Equal(ReadEditableFieldSemantics(projected),
                ReadEditableFieldSemantics(exported));
            var repeated = DxpDocToDocx.Project(DxpDocExport.Export(exported)).DocxBytes;
            Assert.Equal(ReadStories(projected), ReadStories(repeated));
            Assert.Equal(ReadSupportedCoreProperties(projected),
                ReadSupportedCoreProperties(repeated));
            Assert.Equal(ReadDocumentStatistics(projected),
                ReadDocumentStatistics(repeated));
            Assert.Equal(ReadEditableFieldSemantics(projected),
                ReadEditableFieldSemantics(repeated));
            Assert.Equal(ReadEffectiveHeaderFooterStories(projected),
                ReadEffectiveHeaderFooterStories(repeated));
            AssertExplicitSectionLayout(projected, repeated);
            AssertEffectiveParagraphLayout(projected, repeated);
            AssertParagraphDecorations(projected, repeated);
            AssertEffectiveTabStops(projected, repeated);
            AssertEffectiveRunFormatting(projected, repeated,
                name is "WordCharacterStyleToggleStories" or
                    "WordLinkedStylesAllStories" or
                    "WordCharacterStyleBaseToggleStories" or
                    "WordCharacterStyleStrikeOverrideStories" or
                    "WordCharacterStyleInheritedStrikeStories" or
                    "WordCharacterStyleRelativeStrikeStories",
                name is "WordCharacterStyleInheritedStrikeStories" or
                    "WordCharacterStyleRelativeStrikeStories",
                name == "WordVisibleThemeTintShadeAllStories");
            AssertTableGrid(projected, repeated);
            Assert.Equal(ReadListSemantics(projected), ReadListSemantics(repeated));
            Assert.Equal(ReadListSemantics(projected, true),
                ReadListSemantics(repeated, true));
            AssertDrawingGeometry(projected, repeated);
            Assert.Empty(Validate(exported));
            Assert.Empty(Validate(repeated));
        }
        finally
        {
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Theory]
    [MemberData(nameof(PairedFixtures))]
    public void DocxVisitorRetainsProjectedVisibleStoriesFromBinaryDoc(string name)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var inputPath = Path.Combine(directory, name + ".doc");
        var outputPath = Path.Combine(Path.GetTempPath(), $"docxport-visitor-{Guid.NewGuid():N}.docx");
        try
        {
            var projected = DxpDocToDocx.Project(File.ReadAllBytes(inputPath)).DocxBytes;
            DxpExport.ExportToFile(inputPath,
                new DocxportNet.Visitors.Docx.DxpDocxVisitor(), outputPath,
                new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.None });
            var exported = File.ReadAllBytes(outputPath);
            Assert.Equal(ReadStories(projected), ReadStories(exported));
            Assert.Equal(ReadSupportedCoreProperties(projected),
                ReadSupportedCoreProperties(exported));
            Assert.Equal(ReadDocumentStatistics(projected),
                ReadDocumentStatistics(exported));
            Assert.Equal(ReadEffectiveHeaderFooterStories(projected),
                ReadEffectiveHeaderFooterStories(exported));
            AssertExplicitSectionLayout(projected, exported);
            AssertEffectiveParagraphLayout(projected, exported);
            AssertParagraphDecorations(projected, exported);
            AssertEffectiveTabStops(projected, exported);
            AssertEffectiveRunFormatting(projected, exported);
            AssertDrawingGeometry(projected, exported);
            AssertTableGrid(projected, exported);
            Assert.Equal(ReadListSemantics(projected), ReadListSemantics(exported));
            Assert.Equal(ReadListSemantics(projected, true),
                ReadListSemantics(exported, true));
            Assert.Equal(ReadEditableFieldSemantics(projected),
                ReadEditableFieldSemantics(exported));
            Assert.Empty(Validate(exported));
        }
        finally
        {
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void WordSavedBaselineOffsetsSurviveBothDocRoutes()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordBaselineOffsetLayeredStories.docx"));
        var source = File.ReadAllBytes(path);
        foreach (var binary in new[] { File.ReadAllBytes(Path.ChangeExtension(path, ".doc")),
            DxpDocExport.Export(source) })
        {
            using var input = new MemoryStream(binary);
            using var index = new DocTextIndexWalker().Index(input);
            var derived = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Visible Derived");
            Assert.Equal((short)4, derived.CharacterFormatting.BaselineOffsetHalfPoints);
            Assert.Contains(index.CharacterFormatting, x =>
                x.Formatting.BaselineOffsetHalfPoints == -4);
            Assert.Contains(index.CharacterFormatting, x =>
                x.Formatting.BaselineOffsetHalfPoints == 6);
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            AssertEffectiveRunFormatting(source, projected);
            using (var stream = new MemoryStream(projected))
            using (var document = WordprocessingDocument.Open(stream, false))
                Assert.Empty(new OpenXmlValidator().Validate(document));
            var repeated = DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes;
            AssertEffectiveRunFormatting(source, repeated);
        }
    }

    [Fact]
    public void WordSavedCharacterScaleSurvivesBothDocRoutes()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordCharacterScaleLayeredStories.docx"));
        var source = File.ReadAllBytes(path);
        foreach (var binary in new[] { File.ReadAllBytes(Path.ChangeExtension(path, ".doc")),
            DxpDocExport.Export(source) })
        {
            using var input = new MemoryStream(binary);
            using var index = new DocTextIndexWalker().Index(input);
            var derived = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Visible Derived");
            Assert.Equal((ushort)80, derived.CharacterFormatting.CharacterScalePercent);
            Assert.Contains(index.CharacterFormatting, x =>
                x.Formatting.CharacterScalePercent == 120);
            Assert.Contains(index.CharacterFormatting, x =>
                x.Formatting.CharacterScalePercent == 60);
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            AssertEffectiveRunFormatting(source, projected);
            using (var stream = new MemoryStream(projected))
            using (var document = WordprocessingDocument.Open(stream, false))
                Assert.Empty(new OpenXmlValidator().Validate(document));
            var repeated = DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes;
            AssertEffectiveRunFormatting(source, repeated);
        }
    }

    [Fact]
    public void WordSavedSizedEmptyParagraphMarksSurviveBothDocRoutes()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleSizedEmptyParagraphsAllStories.docx"));
        var source = File.ReadAllBytes(path);
        var expected = ReadInteriorEmptyParagraphMarks(source, details: true);
        Assert.Equal("bold|72|C00000", expected["body.paragraph1"]);
        Assert.Equal("italic|36|0070C0",
            expected["section0.header.default.paragraph1"]);
        Assert.Equal("bold|48|008000",
            expected["section0.footer.default.paragraph1"]);
        foreach (var binary in new[] { File.ReadAllBytes(Path.ChangeExtension(path, ".doc")),
            DxpDocExport.Export(source) })
        {
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            Assert.Equal(expected, ReadInteriorEmptyParagraphMarks(projected,
                details: true));
            using (var stream = new MemoryStream(projected))
            using (var document = WordprocessingDocument.Open(stream, false))
                Assert.Empty(new OpenXmlValidator().Validate(document));
            var repeated = DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes;
            Assert.Equal(expected, ReadInteriorEmptyParagraphMarks(repeated,
                details: true));
        }
    }

    [Fact]
    public void DoubleStrikeStyleLayersSurviveBothDocRoutes()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordDoubleStrikeLayeredStories.docx"));
        var source = File.ReadAllBytes(path);
        var native = File.ReadAllBytes(Path.ChangeExtension(path, ".doc"));
        static void CheckDoc(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes);
            using var index = new DocTextIndexWalker().Index(stream);
            var inherited = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Visible Derived");
            Assert.True(inherited.CharacterFormatting.DoubleStrike);
            var accent = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Visible Accent");
            Assert.False(accent.CharacterFormatting.DoubleStrike);
            Assert.Contains(index.CharacterFormatting, x =>
                x.CpStart == 0 && x.Formatting.DoubleStrike == true);
        }
        foreach (var binary in new[] { native, DxpDocExport.Export(source) })
        {
            CheckDoc(binary);
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            using (var stream = new MemoryStream(projected))
            using (var document = WordprocessingDocument.Open(stream, false))
            {
                var baseStyle = Assert.Single(document.MainDocumentPart!
                    .StyleDefinitionsPart!.Styles!.Elements<Style>(),
                    x => x.StyleName?.Val?.Value == "Visible Base");
                Assert.True(baseStyle.StyleRunProperties?
                    .GetFirstChild<DoubleStrike>()?.Val?.Value);
                var bodyRun = Assert.Single(document.MainDocumentPart.Document.Body!
                    .Descendants<Run>(), x => x.InnerText == "Body layered");
                Assert.True(bodyRun.RunProperties?.GetFirstChild<DoubleStrike>()?
                    .Val?.Value);
                Assert.Empty(new OpenXmlValidator().Validate(document));
            }
            CheckDoc(DxpDocExport.Export(projected));
        }
    }

    [Fact]
    public void WordSavedCharacterEffectsSurviveBothDocRoutes()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordCharacterEffectsLayeredStories.docx"));
        var source = File.ReadAllBytes(path);
        var native = File.ReadAllBytes(Path.ChangeExtension(path, ".doc"));
        foreach (var binary in new[] { native, DxpDocExport.Export(source) })
        {
            using var input = new MemoryStream(binary);
            using var index = new DocTextIndexWalker().Index(input);
            var derived = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Visible Derived");
            Assert.True(derived.CharacterFormatting.Outline);
            Assert.Contains(index.CharacterFormatting, x => x.Formatting.Shadow == true);
            Assert.Contains(index.CharacterFormatting, x => x.Formatting.Emboss == true);
            Assert.Contains(index.CharacterFormatting, x => x.Formatting.Imprint == true);
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            AssertEffectiveRunFormatting(source, projected);
            using (var stream = new MemoryStream(projected))
            using (var document = WordprocessingDocument.Open(stream, false))
                Assert.Empty(new OpenXmlValidator().Validate(document));
            var repeated = DxpDocExport.Export(projected);
            AssertEffectiveRunFormatting(source,
                DxpDocToDocx.Project(repeated).DocxBytes);
        }
    }

    [Fact]
    public void WordSavedFontMetadataIsIndexed()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordMissingFontStyleStories.doc"));
        using var input = File.OpenRead(path);
        using var index = new DocTextIndexWalker().Index(input);
        var missing = Assert.Single(index.Fonts,
            x => x.Name == "Docxport Missing Font XYZ");
        Assert.Equal("Cambria", missing.AlternateName);
        Assert.Equal(400, missing.Weight);
        Assert.Equal((byte)0x10, missing.FamilyPitch);
        Assert.Equal(10, missing.Panose!.Length);
        Assert.Equal(24, missing.Signature!.Length);
        var symbol = Assert.Single(index.Fonts, x => x.Name == "Symbol");
        Assert.Equal((byte)2, symbol.Charset);
    }

    [Fact]
    public void DocxFontMetadataWritesRealDocFontRecord()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordMissingFontStyleStories.docx"));
        var source = File.ReadAllBytes(path);
        var native = File.ReadAllBytes(Path.ChangeExtension(path, ".doc"));
        var expectedStories = ReadStories(source);
        const string fontName = "Docxport Missing Font XYZ";
        var sourceStyles = ReadCustomStyleRunFormatting(source);
        foreach (var styleName in new[] { "Unavailable Font Accent.",
            "Unavailable Font Accent Char." })
        {
            Assert.Equal(fontName, sourceStyles[styleName + "fontAscii"]);
            Assert.Equal(fontName, sourceStyles[styleName + "fontHighAnsi"]);
        }
        static DocFontDefinition MissingFont(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes);
            using var index = new DocTextIndexWalker().Index(stream);
            return Assert.Single(index.Fonts,
                x => x.Name == "Docxport Missing Font XYZ");
        }
        var nativeFont = MissingFont(native);
        var generatedFont = MissingFont(DxpDocExport.Export(source));
        Assert.Equal(nativeFont.FamilyPitch, generatedFont.FamilyPitch);
        Assert.Equal(nativeFont.Weight, generatedFont.Weight);
        Assert.Equal(nativeFont.Charset, generatedFont.Charset);
        Assert.Equal(nativeFont.Panose, generatedFont.Panose);
        Assert.Equal(nativeFont.Signature, generatedFont.Signature);
        foreach (var binary in new[] { native, DxpDocExport.Export(source) })
        {
            using var input = new MemoryStream(binary);
            using var index = new DocTextIndexWalker().Index(input);
            var missing = Assert.Single(index.Fonts,
                x => x.Name == "Docxport Missing Font XYZ");
            Assert.Equal("Cambria", missing.AlternateName);
            Assert.Equal((byte)0x10, missing.FamilyPitch);
            Assert.Equal((byte)0, missing.Charset);
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            using (var stream = new MemoryStream(projected))
            using (var document = WordprocessingDocument.Open(stream, false))
            {
                var table = XElement.Parse(document.MainDocumentPart!
                    .FontTablePart!.Fonts!.OuterXml);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var font = Assert.Single(table.Elements(w + "font"),
                    x => (string?)x.Attribute(w + "name") == missing.Name);
                Assert.Equal("Cambria", (string?)font.Element(w + "altName")?
                    .Attribute(w + "val"));
                Assert.Equal("roman", (string?)font.Element(w + "family")?
                    .Attribute(w + "val"));
                Assert.NotNull(font.Element(w + "notTrueType"));
                Assert.Empty(new OpenXmlValidator().Validate(document));
            }
            var third = DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes;
            foreach (var docx in new[] { projected, third })
            {
                Assert.Equal(expectedStories, ReadStories(docx));
                var styles = ReadCustomStyleRunFormatting(docx);
                foreach (var styleName in new[] { "Unavailable Font Accent.",
                    "Unavailable Font Accent Char." })
                {
                    Assert.Equal(fontName, styles[styleName + "fontAscii"]);
                    Assert.Equal(fontName, styles[styleName + "fontHighAnsi"]);
                }
            }
            using var rewritten = new MemoryStream(DxpDocExport.Export(projected));
            using var repeated = new DocTextIndexWalker().Index(rewritten);
            var again = Assert.Single(repeated.Fonts, x => x.Name == missing.Name);
            Assert.Equal(missing.AlternateName, again.AlternateName);
            Assert.Equal(missing.FamilyPitch, again.FamilyPitch);
            Assert.Equal(missing.Charset, again.Charset);
        }
    }

    [Fact]
    public void TopGutterSurvivesBothDocRoutes()
    {
        var dir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(dir, "WordVisibleTopGutterAllStories.docx"));
        var native = File.ReadAllBytes(Path.Combine(dir, "WordVisibleTopGutterAllStories.doc"));
        static bool? Read(byte[] bytes)
        {
            using var input = new MemoryStream(bytes);
            using var index = new DocTextIndexWalker().Index(input);
            return index.GutterAtTop;
        }
        Assert.True(Read(native));
        Assert.True(Read(DxpDocExport.Export(source)));
        foreach (var binary in new[] { native, DxpDocExport.Export(source) })
        {
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            using var stream = new MemoryStream(projected);
            using var document = WordprocessingDocument.Open(stream, false);
            Assert.NotNull(document.MainDocumentPart!.DocumentSettingsPart!
                .Settings!.GetFirstChild<GutterAtTop>());
            Assert.Empty(new OpenXmlValidator().Validate(document));
            Assert.True(Read(DxpDocExport.Export(projected)));
        }
    }

    [Fact]
    public void MirrorMarginsSurviveBothDocRoutes()
    {
        var dir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(dir, "WordVisibleMirrorMarginsAllStories.docx"));
        var native = File.ReadAllBytes(Path.Combine(dir, "WordVisibleMirrorMarginsAllStories.doc"));
        static bool? Read(byte[] bytes)
        {
            using var input = new MemoryStream(bytes);
            using var index = new DocTextIndexWalker().Index(input);
            return index.MirrorMargins;
        }
        Assert.True(Read(native));
        Assert.True(Read(DxpDocExport.Export(source)));
        foreach (var binary in new[] { native, DxpDocExport.Export(source) })
        {
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            using var stream = new MemoryStream(projected);
            using var document = WordprocessingDocument.Open(stream, false);
            Assert.NotNull(document.MainDocumentPart!.DocumentSettingsPart!
                .Settings!.GetFirstChild<MirrorMargins>());
            Assert.Empty(new OpenXmlValidator().Validate(document));
            Assert.True(Read(DxpDocExport.Export(projected)));
        }
    }

    [Fact]
    public void DefaultTabIntervalSurvivesBothDocRoutes()
    {
        var dir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var native = File.ReadAllBytes(Path.Combine(dir, "WordVisibleTextualListFormatsAllStories.doc"));
        var source = File.ReadAllBytes(Path.Combine(dir, "WordVisibleTextualListFormatsAllStories.docx"));
        static short? TabInterval(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes);
            using var index = new DocTextIndexWalker().Index(stream);
            return index.DefaultTabStopTwips;
        }
        Assert.Equal((short?)720, TabInterval(native));
        Assert.Equal((short?)720, TabInterval(DxpDocExport.Export(source)));
        var nativeCustom = File.ReadAllBytes(Path.Combine(dir,
            "WordVisibleCustomDefaultTabsAllStories.doc"));
        Assert.Equal((short?)1080, TabInterval(nativeCustom));
        var fromNative = DxpDocToDocx.Project(nativeCustom).DocxBytes;
        using (var stream = new MemoryStream(fromNative))
        using (var document = WordprocessingDocument.Open(stream, false))
            Assert.Equal(1080, document.MainDocumentPart!.DocumentSettingsPart!
                .Settings!.GetFirstChild<DefaultTabStop>()!.Val!.Value);
        Assert.Equal((short?)1080, TabInterval(DxpDocExport.Export(fromNative)));
        using var customStream = new MemoryStream();
        customStream.Write(source);
        customStream.Position = 0;
        using (var document = WordprocessingDocument.Open(customStream, true))
        {
            var settingsPart = document.MainDocumentPart!.DocumentSettingsPart ??
                document.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings ??= new Settings();
            settingsPart.Settings.GetFirstChild<DefaultTabStop>()?.Remove();
            settingsPart.Settings.AppendChild(new DefaultTabStop { Val = 1080 });
            settingsPart.Settings.AppendChild(new EvenAndOddHeaders());
            settingsPart.Settings.Save();
        }
        var custom = customStream.ToArray();
        var generated = DxpDocExport.Export(custom);
        Assert.Equal((short?)1080, TabInterval(generated));
        var projected = DxpDocToDocx.Project(generated).DocxBytes;
        using (var stream = new MemoryStream(projected))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            Assert.Equal(1080, document.MainDocumentPart!.DocumentSettingsPart!
                .Settings!.GetFirstChild<DefaultTabStop>()!.Val!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        Assert.Equal((short?)1080, TabInterval(DxpDocExport.Export(projected)));
    }

    [Fact]
    public void ColorSchemeMappingChangesVisibleStyleAndStoryColors()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleThemeColorMappingAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using (var stream = new MemoryStream(source))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var mapping = document.MainDocumentPart!.DocumentSettingsPart!
                .Settings!.GetFirstChild<ColorSchemeMapping>()!;
            Assert.Equal(ColorSchemeIndexValues.Accent4, mapping.Accent1?.Value);
            Assert.Equal(ColorSchemeIndexValues.Accent5, mapping.Accent2?.Value);
            Assert.Equal(ColorSchemeIndexValues.Accent6, mapping.Accent3?.Value);
        }
        var effective = ReadEffectiveRunFormatting(source);
        Assert.Contains(effective, x => x.Key.StartsWith("body.", StringComparison.Ordinal) &&
            x.Key.EndsWith(".color", StringComparison.Ordinal) && x.Value == "0F9ED5");
        Assert.Contains(effective, x => x.Key.Contains(".header.", StringComparison.Ordinal) &&
            x.Key.EndsWith(".color", StringComparison.Ordinal) && x.Value == "A02B93");
        Assert.Contains(effective, x => x.Key.Contains(".footer.", StringComparison.Ordinal) &&
            x.Key.EndsWith(".color", StringComparison.Ordinal) && x.Value == "4EA72E");
    }

    [Fact]
    public void ThemeTextColorsOverrideDifferentFallbackValuesInAllStories()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleThemeTextColorsAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using (var stream = new MemoryStream(source))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var main = document.MainDocumentPart!;
            var baseStyle = main.StyleDefinitionsPart!.Styles!.Elements<Style>()
                .Single(x => x.StyleName?.Val?.Value == "Visible Base");
            Assert.Equal("FF0000", baseStyle.StyleRunProperties?
                .GetFirstChild<Color>()?.Val?.Value);
            Assert.Equal(ThemeColorValues.Accent1, baseStyle.StyleRunProperties?
                .GetFirstChild<Color>()?.ThemeColor?.Value);
            Assert.Contains(main.HeaderParts.SelectMany(x => x.Header!.Descendants<Color>()),
                x => x.Val?.Value == "00FF00" &&
                    x.ThemeColor?.Value == ThemeColorValues.Accent2);
            Assert.Contains(main.FooterParts.SelectMany(x => x.Footer!.Descendants<Color>()),
                x => x.Val?.Value == "0000FF" &&
                    x.ThemeColor?.Value == ThemeColorValues.Accent3);
        }
        var effective = ReadEffectiveRunFormatting(source);
        Assert.Contains(effective, x => x.Key.StartsWith("body.", StringComparison.Ordinal) &&
            x.Key.EndsWith(".color", StringComparison.Ordinal) && x.Value == "156082");
        Assert.Contains(effective, x => x.Key.Contains(".header.", StringComparison.Ordinal) &&
            x.Key.EndsWith(".color", StringComparison.Ordinal) && x.Value == "E97132");
        Assert.Contains(effective, x => x.Key.Contains(".footer.", StringComparison.Ordinal) &&
            x.Key.EndsWith(".color", StringComparison.Ordinal) && x.Value == "196B24");
    }

    [Theory]
    [InlineData("WordVisibleThemeTextColorsAllStories", "156082", "156082")]
    [InlineData("WordVisibleThemeColorMappingAllStories", "0F9ED5", "0F9ED5")]
    [InlineData("WordVisibleThemeTintShadeAllStories", "64BDE6", "63BDE6")]
    public void ThemedBaseStyleColorRemainsEditableAcrossBothDocRoutes(
        string name, string generatedRgb, string nativeRgb)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var native = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
        foreach (var (binary, expectedRgb, nativeRounding) in new[]
        {
            (native, nativeRgb, nativeRgb != generatedRgb),
            (DxpDocExport.Export(source), generatedRgb, false)
        })
        {
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            AssertBaseStyleColor(projected, expectedRgb);
            AssertEffectiveRunFormatting(source, projected,
                wordValidatedThemeLuminance: nativeRounding);
            var repeated = DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes;
            AssertBaseStyleColor(repeated, expectedRgb);
            AssertEffectiveRunFormatting(source, repeated,
                wordValidatedThemeLuminance: nativeRounding);
        }
    }

    private static void AssertBaseStyleColor(byte[] bytes, string expectedRgb)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var styles = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Style>().ToArray();
        var root = Assert.Single(styles, x => x.StyleName?.Val?.Value == "Visible Base");
        var derived = Assert.Single(styles, x => x.StyleName?.Val?.Value == "Visible Derived");
        Assert.Equal(root.StyleId?.Value, derived.BasedOn?.Val?.Value);
        Assert.Equal(expectedRgb, root.StyleRunProperties?
            .GetFirstChild<Color>()?.Val?.Value);
        Assert.Null(derived.StyleRunProperties?.GetFirstChild<Color>());
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2010).Validate(document));
    }

    [Fact]
    public void ThemeFontsResolveThroughStyledBodyAndDirectHeaderFooterRuns()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordVisibleThemeFontsAllStories.docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            "WordVisibleThemeFontsAllStories.doc"));
        static void AssertFonts(byte[] binary)
        {
            using var stream = new MemoryStream(binary);
            using var index = new DocTextIndexWalker().Index(stream);
            Assert.Contains(index.Fonts, x => x.Name == "Aptos Display");
            Assert.Contains(index.Fonts, x => x.Name == "Aptos");
            Assert.Contains(index.StyleDefinitions, x =>
                x.Name == "Visible Base" &&
                x.CharacterFormatting.AsciiFontName == "Aptos Display");
            Assert.Contains(index.CharacterFormatting, x =>
                x.Formatting.AsciiFontName == "Aptos");
        }
        AssertFonts(native);
        var generated = DxpDocExport.Export(source);
        AssertFonts(generated);
        AssertFonts(DxpDocExport.Export(DxpDocToDocx.Project(native).DocxBytes));
        AssertFonts(DxpDocExport.Export(DxpDocToDocx.Project(generated).DocxBytes));
        foreach (var binary in new[] { native, generated })
        {
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            AssertEffectiveRunFormatting(source, projected);
            Assert.Empty(Validate(projected));
            var repeated = DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes;
            AssertEffectiveRunFormatting(source, repeated);
            Assert.Empty(Validate(repeated));
        }
    }

    [Theory]
    [InlineData("WordThreeSectionEmptyInheritedHeader", false)]
    [InlineData("WordThreeSectionEmptyInheritedStories", true)]
    public void ExplicitEmptyStoriesStayDistinctFromInheritedStories(
        string fixtureName, bool emptyFooter)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            fixtureName + ".docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            fixtureName + ".doc"));
        void AssertEmptyThenInherited(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes);
            using var document = WordprocessingDocument.Open(stream, false);
            var main = document.MainDocumentPart!;
            var sections = main.Document!.Body!.Descendants<SectionProperties>().ToArray();
            Assert.Equal(3, sections.Length);
            var second = Assert.Single(sections[1].Elements<HeaderReference>(), x =>
                x.Type?.Value == HeaderFooterValues.Default);
            var header = (HeaderPart)main.GetPartById(second.Id!);
            Assert.Empty(header.Header!.Descendants<Text>().Where(x => x.Text.Length != 0));
            Assert.DoesNotContain(sections[2].Elements<HeaderReference>(), x =>
                x.Type?.Value == HeaderFooterValues.Default);
            if (emptyFooter)
            {
                var footerReference = Assert.Single(
                    sections[1].Elements<FooterReference>(), x =>
                        x.Type?.Value == HeaderFooterValues.Default);
                var footer = (FooterPart)main.GetPartById(footerReference.Id!);
                Assert.Empty(footer.Footer!.Descendants<Text>()
                    .Where(x => x.Text.Length != 0));
                Assert.DoesNotContain(sections[2].Elements<FooterReference>(), x =>
                    x.Type?.Value == HeaderFooterValues.Default);
            }
        }
        AssertEmptyThenInherited(source);
        foreach (var binary in new[] { native, DxpDocExport.Export(source) })
        {
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            AssertEmptyThenInherited(projected);
            AssertEmptyThenInherited(DxpDocToDocx.Project(
                DxpDocExport.Export(projected)).DocxBytes);
        }
    }

    [Fact]
    public void BaseDerivedAndDirectParagraphDirectionPrecedenceSurvivesAllDocRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordVisibleStyledBidiPrecedenceAllStories.docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            "WordVisibleStyledBidiPrecedenceAllStories.doc"));
        static void AssertDirectionPrecedence(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes);
            using var document = WordprocessingDocument.Open(stream, false);
            var main = document.MainDocumentPart!;
            var styles = main.StyleDefinitionsPart!.Styles!;
            var baseStyle = styles.Elements<Style>().Single(x =>
                x.StyleName?.Val?.Value == "Visible Base");
            var derived = styles.Elements<Style>().Single(x =>
                x.StyleName?.Val?.Value == "Visible Derived");
            var baseDirection = Assert.IsType<BiDi>(baseStyle.StyleParagraphProperties?
                .GetFirstChild<BiDi>());
            var derivedDirection = Assert.IsType<BiDi>(derived.StyleParagraphProperties?
                .GetFirstChild<BiDi>());
            Assert.NotEqual(false, baseDirection.Val?.Value);
            Assert.Equal(false, derivedDirection.Val?.Value);
            var stories = new DocumentFormat.OpenXml.OpenXmlElement[]
                { main.Document!.Body! }
                .Concat(main.HeaderParts.Select(x => (DocumentFormat.OpenXml.OpenXmlElement)x.Header!))
                .Concat(main.FooterParts.Select(x => (DocumentFormat.OpenXml.OpenXmlElement)x.Footer!));
            Assert.Equal(3, stories.SelectMany(x => x.Descendants<Paragraph>())
                .Count(x => x.InnerText.Contains("reset", StringComparison.Ordinal) &&
                    x.ParagraphProperties?.GetFirstChild<BiDi>()?.Val?.Value == true));
        }
        AssertDirectionPrecedence(source);
        foreach (var binary in new[] { native, DxpDocExport.Export(source) })
        {
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            AssertDirectionPrecedence(projected);
            AssertDirectionPrecedence(DxpDocToDocx.Project(
                DxpDocExport.Export(projected)).DocxBytes);
        }
    }

    [Theory]
    [InlineData("WordVisibleStyledBidiAllStories")]
    [InlineData("WordVisibleStyledBidiResetAllStories")]
    public void StyledParagraphDirectionSurvivesBodyHeaderFooterDocRoutes(string fixtureName)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            fixtureName + ".docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            fixtureName + ".doc"));
        static bool HasStyledBidi(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes);
            using var document = WordprocessingDocument.Open(stream, false);
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    "Visible Derived");
            return style.StyleParagraphProperties?.GetFirstChild<BiDi>()?
                .Val?.Value != false &&
                style.StyleParagraphProperties?.GetFirstChild<BiDi>() != null;
        }
        static bool HasDirectResets(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes);
            using var document = WordprocessingDocument.Open(stream, false);
            var main = document.MainDocumentPart!;
            var stories = new DocumentFormat.OpenXml.OpenXmlElement[]
                { main.Document!.Body! }
                .Concat(main.HeaderParts.Select(x => (DocumentFormat.OpenXml.OpenXmlElement)x.Header!))
                .Concat(main.FooterParts.Select(x => (DocumentFormat.OpenXml.OpenXmlElement)x.Footer!));
            return stories.SelectMany(x => x.Descendants<Paragraph>())
                .Count(x => x.InnerText.Contains("reset", StringComparison.Ordinal) &&
                    x.ParagraphProperties?.GetFirstChild<BiDi>()?.Val?.Value == false) == 3;
        }
        Assert.True(HasStyledBidi(source));
        if (fixtureName.Contains("Reset", StringComparison.Ordinal))
            Assert.True(HasDirectResets(source));
        foreach (var binary in new[] { native, DxpDocExport.Export(source) })
        {
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            Assert.True(HasStyledBidi(projected));
            if (fixtureName.Contains("Reset", StringComparison.Ordinal))
                Assert.True(HasDirectResets(projected));
            var thirdHop = DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes;
            Assert.True(HasStyledBidi(thirdHop));
            if (fixtureName.Contains("Reset", StringComparison.Ordinal))
                Assert.True(HasDirectResets(thirdHop));
        }
    }

    [Fact]
    public void WordNativeExplicitRtlHeaderCarriesModernAndCompatibilityRowProperties()
    {
        var fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordLogicalRightRtlHeader.doc"));
        using var input = File.OpenRead(fixture);
        using var structure = new DocStructureWalker().Accept(input,
            new DocStructurePrintVisitor(TextWriter.Null));
        static IEnumerable<DocStructureNode> Descendants(DocStructureNode node) =>
            node.Children.SelectMany(child => new[] { child }.Concat(Descendants(child)));
        var rows = Descendants(structure.Root).Where(node =>
            node.Kind == "PapxInFkp" && Descendants(node).Any(child =>
                child.Kind == "Sprm" && child.Attributes.TryGetValue("code", out var code) &&
                code == "0x560B"))
            .ToArray();
        var rtl = Assert.Single(rows, row => Descendants(row).Any(child =>
            child.Kind == "Sprm" && child.Attributes.TryGetValue("code", out var code) &&
            code == "0x548A"));
        var modern = Assert.Single(Descendants(rtl), node => node.Kind == "PrcData");
        var modernCodes = Descendants(modern).Where(node => node.Kind == "Sprm")
            .Select(node => node.Attributes["code"]).ToArray();
        Assert.Equal(new[] { "0x2416", "0x2417", "0x6649", "0x9601",
            "0x7621", "0x7623", "0x563A", "0x9602", "0x560B",
            "0xF614", "0xD635", "0xD62F", "0x7479", "0x548A" }, modernCodes);
        string Operand(string code)
        {
            var prl = Assert.Single(modern.Children, x => x.Kind == "Prl" &&
                x.Children.FirstOrDefault()?.Attributes.TryGetValue("code", out var found) == true &&
                found == code);
            return Convert.ToHexString(structure.ReadRange(prl.StreamName!,
                prl.Offset!.Value, checked((int)prl.Length!.Value)));
        }
        Assert.Equal("01966C00", Operand("0x9601"));
        Assert.Equal("217600026801", Operand("0x7621"));
        Assert.Equal("237600024812", Operand("0x7623"));
        Assert.Equal("02966C00", Operand("0x9602"));
        Assert.Equal("8A540200", Operand("0x548A"));
        var compatibilityCodes = Descendants(rtl).Where(node =>
                node.Kind == "Sprm" && node.StreamName == "WordDocument")
            .Select(node => node.Attributes["code"]).ToArray();
        Assert.Contains("0xD608", compatibilityCodes);
        Assert.Contains("0xF661", compatibilityCodes);
        Assert.Contains("0x548A", compatibilityCodes);
    }

    [Theory]
    [InlineData("WordVisibleAutoRowGridGapsAllStories", 1, 3120)]
    [InlineData("WordVisibleMultiRowGridGapsAllStories", 2, 5000)]
    public void WordSavedOmittedRowGapWidthUsesRowOriginInAllStories(
        string fixtureName, int skippedColumns, int leadingWidth)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            fixtureName + ".docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            fixtureName + ".doc"));
        using var stream = new MemoryStream(native);
        using var index = new DocTextIndexWalker().Index(stream);
        var rows = index.ParagraphStyles.Where(x =>
            x.Formatting.TableTerminator == true).Select(x => x.Formatting).ToArray();
        Assert.Equal((byte)3, rows[1].TableWidthBefore?.Unit);
        Assert.Equal((ushort)0, rows[1].TableWidthBefore?.Value);
        Assert.Equal(leadingWidth,
            rows[1].TableRowOriginTwips - rows[0].TableRowOriginTwips);
        static (int Before, int After)[] Gaps(byte[] bytes)
        {
            using var input = new MemoryStream(bytes);
            using var document = WordprocessingDocument.Open(input, false);
            var main = document.MainDocumentPart!;
            return main.Document.Body!.Elements<Table>()
                .Concat(main.HeaderParts.SelectMany(x => x.Header!.Descendants<Table>()))
                .Concat(main.FooterParts.SelectMany(x => x.Footer!.Descendants<Table>()))
                .SelectMany(table => table.Elements<TableRow>())
                .Select(row => ((int?)row.TableRowProperties?
                    .GetFirstChild<GridBefore>()?.Val?.Value ?? 0,
                    (int?)row.TableRowProperties?
                    .GetFirstChild<GridAfter>()?.Val?.Value ?? 0))
                .ToArray();
        }
        var expected = Enumerable.Range(0, 3).SelectMany(_ =>
            new[] { (0, 0), (skippedColumns, 0), (0, skippedColumns) }).ToArray();
        Assert.Equal(expected, Gaps(source));
        var nativeProjected = DxpDocToDocx.Project(native).DocxBytes;
        using (var input = new MemoryStream(nativeProjected))
        using (var document = WordprocessingDocument.Open(input, false))
        {
            var main = document.MainDocumentPart!;
            var tables = main.Document.Body!.Elements<Table>()
                .Concat(main.HeaderParts.SelectMany(x => x.Header!.Descendants<Table>()))
                .Concat(main.FooterParts.SelectMany(x => x.Footer!.Descendants<Table>()));
            Assert.All(tables, table => Assert.Null(table.Elements<TableRow>()
                .ElementAt(1).TableRowProperties?
                .GetFirstChild<WidthBeforeTableRow>()));
        }
        foreach (var (binary, fromGeneratedDoc) in new[]
        {
            (native, false), (DxpDocExport.Export(source), true)
        })
        {
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            Assert.Equal(expected, Gaps(projected));
            if (fromGeneratedDoc)
            {
                using var input = new MemoryStream(projected);
                using var projectedDocument = WordprocessingDocument.Open(input, false);
                var main = projectedDocument.MainDocumentPart!;
                var tables = main.Document.Body!.Elements<Table>()
                    .Concat(main.HeaderParts.SelectMany(x => x.Header!.Descendants<Table>()))
                    .Concat(main.FooterParts.SelectMany(x => x.Footer!.Descendants<Table>()));
                Assert.All(tables, table =>
                {
                    Assert.Null(table.Elements<TableRow>().ElementAt(1)
                        .TableRowProperties?.GetFirstChild<WidthBeforeTableRow>());
                    Assert.Null(table.Elements<TableRow>().ElementAt(2)
                        .TableRowProperties?.GetFirstChild<WidthAfterTableRow>());
                });
            }
            var thirdHop = DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes;
            Assert.Equal(expected, Gaps(thirdHop));
            if (fromGeneratedDoc)
            {
                using var input = new MemoryStream(thirdHop);
                using var roundTripDocument = WordprocessingDocument.Open(input, false);
                var main = roundTripDocument.MainDocumentPart!;
                var tables = main.Document.Body!.Elements<Table>()
                    .Concat(main.HeaderParts.SelectMany(x => x.Header!.Descendants<Table>()))
                    .Concat(main.FooterParts.SelectMany(x => x.Footer!.Descendants<Table>()));
                Assert.All(tables, table =>
                {
                    Assert.Null(table.Elements<TableRow>().ElementAt(1)
                        .TableRowProperties?.GetFirstChild<WidthBeforeTableRow>());
                    Assert.Null(table.Elements<TableRow>().ElementAt(2)
                        .TableRowProperties?.GetFirstChild<WidthAfterTableRow>());
                });
            }
            using var package = new MemoryStream(thirdHop);
            using var document = WordprocessingDocument.Open(package, false);
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
    }

    [Theory]
    [InlineData("WordVisibleRowGridGaps", 1)]
    [InlineData("WordVisibleRowGridGapsAllStories", 3)]
    public void WordSavedRowGridGapsKeepCellPositionsThroughBothDocRoutes(
        string fixtureName, int tableCount)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            fixtureName + ".docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            fixtureName + ".doc"));
        static IEnumerable<Table> Tables(WordprocessingDocument document)
        {
            var main = document.MainDocumentPart!;
            return main.Document.Body!.Elements<Table>()
                .Concat(main.HeaderParts.SelectMany(x => x.Header!.Descendants<Table>()))
                .Concat(main.FooterParts.SelectMany(x => x.Footer!.Descendants<Table>()))
                .Where(table => table.Elements<TableRow>().Count() == 3);
        }
        static (int Before, int After)[] Gaps(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes);
            using var document = WordprocessingDocument.Open(stream, false);
            return Tables(document).SelectMany(table => table.Elements<TableRow>())
                .Select(row => ((int?)row.TableRowProperties?
                    .GetFirstChild<GridBefore>()?.Val?.Value ?? 0,
                    (int?)row.TableRowProperties?
                    .GetFirstChild<GridAfter>()?.Val?.Value ?? 0))
                .ToArray();
        }
        static (string? Before, string? After)[] GapWidths(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes);
            using var document = WordprocessingDocument.Open(stream, false);
            return Tables(document).SelectMany(table => table.Elements<TableRow>())
                .Select(row => (row.TableRowProperties?
                    .GetFirstChild<WidthBeforeTableRow>()?.Width?.Value ?? "0",
                    row.TableRowProperties?
                    .GetFirstChild<WidthAfterTableRow>()?.Width?.Value ?? "0"))
                .ToArray();
        }
        var expected = Enumerable.Range(0, tableCount).SelectMany(_ =>
            new[] { (0, 0), (1, 0), (0, 1) }).ToArray();
        var expectedWidths = Enumerable.Range(0, tableCount).SelectMany(_ =>
            new[] { ("0", "0"), ("3120", "0"), ("0", "3120") }).ToArray();
        Assert.Equal(expected, Gaps(source));
        Assert.Equal(expectedWidths, GapWidths(source));
        using (var indexStream = new MemoryStream(native))
        using (var index = new DocTextIndexWalker().Index(indexStream))
            Assert.Equal(new[] { "0,3120,6240,9360", "0,3120,6240", "0,3120,6240" },
                index.ParagraphStyles.Where(x => x.Formatting.TableTerminator == true)
                    .Take(3).Select(x => string.Join(',', x.Formatting.TableCellEdges!)));
        foreach (var binary in new[] { native, DxpDocExport.Export(source) })
        {
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            Assert.Equal(expected, Gaps(projected));
            Assert.Equal(expectedWidths, GapWidths(projected));
            var thirdHop = DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes;
            Assert.Equal(expected, Gaps(thirdHop));
            Assert.Equal(expectedWidths, GapWidths(thirdHop));
            using var stream = new MemoryStream(thirdHop);
            using var document = WordprocessingDocument.Open(stream, false);
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
    }

    [Fact]
    public void ImplicitThemeFontReachesDocDefaultsInAllStoriesFixture()
    {
        var fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleImplicitThemeAllStories.docx"));
        using var stream = new MemoryStream(DxpDocExport.Export(File.ReadAllBytes(fixture)));
        using var index = new DocTextIndexWalker().Index(stream);
        Assert.Equal("Aptos", index.DefaultCharacterFormatting.AsciiFontName);
        Assert.Equal("Aptos", index.DefaultCharacterFormatting.HighAnsiFontName);
    }
    [Fact]
    public void WordNativeAutoWidthRowsExposeReferencedTableProperties()
    {
        var fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleNoWrapOffAllStories.doc"));
        using var input = File.OpenRead(fixture);
        using var structure = new DocStructureWalker().Accept(input,
            new DocStructurePrintVisitor(TextWriter.Null));
        static IEnumerable<DocStructureNode> Descendants(DocStructureNode node) =>
            node.Children.SelectMany(child => new[] { child }.Concat(Descendants(child)));
        var references = Descendants(structure.Root).Where(node =>
            node.Kind == "PrcData" && node.Name == "TableProperties").ToArray();
        Assert.Equal(4, references.Length);
        var expectedEdges = new[] { 8430, 8430, 8208, 8300 };
        for (var rowIndex = 0; rowIndex < references.Length; rowIndex++)
        {
            var reference = references[rowIndex];
            Assert.Equal("Data", reference.StreamName);
            var codes = Descendants(reference).Where(node => node.Kind == "Sprm")
                .Select(node => node.Attributes["code"]).ToArray();
            Assert.Equal(new[]
            {
                "0x2416", "0x2417", "0x6649", "0x9601", "0x7621", "0x7623",
                "0x7623", "0x563A", "0x9602", "0xF614", "0x3615", "0xD635",
                "0x7479"
            }, codes);
            var data = structure.ReadRange("Data", reference.Offset!.Value,
                checked((int)reference.Length!.Value));
            if (rowIndex == 0)
                // Keep the complete Word-authored row block as the writer's
                // byte-level reference, including its fixed modern modifiers.
                Assert.Equal(
                    "400016240117240149660100000001966C0021760002680123760001EE20" +
                    "23760102A2033A560B0002966C0014F601000015360135D605000201" +
                    "00007974A5382B00", Convert.ToHexString(data));
            Assert.Equal(66, data.Length);
            var formatting = DocParagraphFormatting.Parse(data.AsSpan(2));
            Assert.Equal(new short[] { 0, checked((short)expectedEdges[rowIndex]), 9360 },
                formatting.TableCellEdges);
            Assert.All(formatting.TableCellPreferredWidths!, cell =>
                Assert.Equal((byte)1, cell?.Unit));
        }
        input.Position = 0;
        using var index = new DocTextIndexWalker().Index(input);
        var indexedRows = index.ParagraphStyles.Where(x =>
            x.Formatting?.TableTerminator == true).Select(x => x.Formatting!).ToArray();
        Assert.Equal(4, indexedRows.Length);
        foreach (var row in indexedRows)
            Assert.Equal(new byte[] { 1, 1 }, row.TableCellPreferredWidths!
                .Select(cell => cell!.Unit));
    }
    [Fact]
    public void WordNativeAutoWidthRowsExposeContentFittedCellOrigins()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        int[] FirstCellEdges(string name)
        {
            using var stream = new MemoryStream(File.ReadAllBytes(Path.Combine(directory, name + ".doc")));
            using var index = new DocTextIndexWalker().Index(stream);
            return index.ParagraphStyles.Where(x =>
                x.Formatting?.TableTerminator == true).Select(x =>
                (int)x.Formatting!.TableCellEdges![1]).ToArray();
        }
        Assert.Equal(new[] { 8430, 8430, 8208, 8300 },
            FirstCellEdges("WordVisibleNoWrapOffAllStories"));
        Assert.Equal(new[] { 8653, 8653, 8440, 8521 },
            FirstCellEdges("WordVisibleNoWrapAllStories"));
        if (Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_WORD_RENDER") != "1")
            return;
        dynamic word = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application")!)!;
        word.Visible = false;
        dynamic? opened = null;
        try
        {
            foreach (var (name, expectedTwips) in new[]
            {
                ("WordVisibleNoWrapOffAllStories", 8430),
                ("WordVisibleNoWrapAllStories", 8653)
            })
            {
                opened = word.Documents.Open(Path.Combine(directory, name + ".docx"),
                    ReadOnly: true, AddToRecentFiles: false);
                var actualTwips = Convert.ToDouble(opened.Tables.Item(1)
                    .Cell(1, 1).Width) * 20;
                Assert.InRange(actualTwips, expectedTwips - 2, expectedTwips + 2);
                dynamic secondCell = opened.Tables.Item(1).Cell(1, 2);
                Assert.Equal(12d, Convert.ToDouble(secondCell.Range.Font.Size), 1);
                Assert.Equal("Aptos", Convert.ToString(secondCell.Range.Font.Name));
                Assert.Equal(5.4d, Convert.ToDouble(secondCell.LeftPadding), 1);
                Assert.Equal(5.4d, Convert.ToDouble(secondCell.RightPadding), 1);
                opened.Close(false);
                opened = null;
            }
        }
        finally
        {
            if (opened != null) opened.Close(false);
            word.Quit();
        }
    }
    [Fact]
    public void InheritedBandRulesAndDirectCellBordersKeepTheirPrecedence()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordConditionalBandInheritedDirectAllStories";
        foreach (var binary in new[]
        {
            File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
            DxpDocExport.Export(File.ReadAllBytes(Path.Combine(directory,
                name + ".docx")))
        })
        {
            using var stream = new MemoryStream(DxpDocToDocx.Project(binary).DocxBytes);
            using var document = WordprocessingDocument.Open(stream, false);
            var main = document.MainDocumentPart!;
            var body = Assert.Single(main.Document.Body!.Elements<Table>());
            var header = Assert.Single(main.HeaderParts.SelectMany(x =>
                x.Header!.Elements<Table>()));
            var footer = Assert.Single(main.FooterParts.SelectMany(x =>
                x.Footer!.Elements<Table>()));
            TableCell Cell(Table table, int row, int column) => table
                .Elements<TableRow>().ElementAt(row).Elements<TableCell>()
                .ElementAt(column);
            Assert.Equal("000000", Cell(body, 1, 1).TableCellProperties?
                .TableCellBorders?.LeftBorder?.Color?.Value);
            Assert.Equal("000000", Cell(header, 2, 0).TableCellProperties?
                .TableCellBorders?.TopBorder?.Color?.Value);
            Assert.Equal("000000", Cell(footer, 1, 1).TableCellProperties?
                .TableCellBorders?.LeftBorder?.Color?.Value);
            Assert.Equal("AA0000", Cell(header, 1, 0).TableCellProperties?
                .TableCellBorders?.TopBorder?.Color?.Value);
            Assert.Equal("FFFF00", Cell(body, 1, 0).TableCellProperties?
                .TableCellBorders?.TopBorder?.Color?.Value);
            Assert.Equal("00AA00", Cell(footer, 1, 0).TableCellProperties?
                .TableCellBorders?.LeftBorder?.Color?.Value);
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
    }

    [Fact]
    public void InheritedBandStyleIdentityFollowsEncodedRowReferences()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordConditionalBandInheritedAllStories";
        var native = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
        var generated = DxpDocExport.Export(File.ReadAllBytes(
            Path.Combine(directory, name + ".docx")));
        foreach (var (bytes, encodedReference) in new[]
        {
            (native, false), (generated, true)
        })
        {
            using var index = new DocTextIndexWalker().Index(new MemoryStream(bytes));
            var style = Assert.Single(index.StyleDefinitions.Where(definition =>
                definition.Type == 3 && definition.Name == "Band Derived"));
            var rowReferences = index.ParagraphStyles.Where(paragraph =>
                paragraph.Formatting?.TableTerminator == true)
                .Select(paragraph => paragraph.Formatting!.TableStyleIndex).ToArray();
            Assert.Equal(9, rowReferences.Length);
            if (encodedReference)
                Assert.All(rowReferences, reference =>
                    Assert.Equal((ushort)style.Index, reference));
            else
                Assert.All(rowReferences, reference => Assert.Null(reference));

            using var projectedStream = new MemoryStream(
                DxpDocToDocx.Project(bytes).DocxBytes);
            using var projected = WordprocessingDocument.Open(projectedStream, false);
            var tables = projected.MainDocumentPart!.Document.Body!
                .Elements<Table>()
                .Concat(projected.MainDocumentPart.HeaderParts.SelectMany(part =>
                    part.Header!.Elements<Table>()))
                .Concat(projected.MainDocumentPart.FooterParts.SelectMany(part =>
                    part.Footer!.Elements<Table>())).ToArray();
            Assert.Equal(3, tables.Length);
            Assert.All(tables, table =>
            {
                var reference = table.TableProperties?.GetFirstChild<TableStyle>();
                if (encodedReference)
                    Assert.Equal($"DocStyle{style.Index}", reference?.Val?.Value);
                else
                    Assert.Null(reference);
            });
        }
    }
    [Fact]
    public void DerivedBandStyleOverridesOnlyItsSecondHorizontalBand()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordConditionalBandInheritedAllStories";
        foreach (var binary in new[]
        {
            File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
            DxpDocExport.Export(File.ReadAllBytes(Path.Combine(directory,
                name + ".docx")))
        })
        {
            using var stream = new MemoryStream(DxpDocToDocx.Project(binary).DocxBytes);
            using var document = WordprocessingDocument.Open(stream, false);
            var main = document.MainDocumentPart!;
            var tables = new[]
            {
                Assert.Single(main.Document.Body!.Elements<Table>()),
                Assert.Single(main.HeaderParts.SelectMany(x => x.Header!.Elements<Table>())),
                Assert.Single(main.FooterParts.SelectMany(x => x.Footer!.Elements<Table>()))
            };
            for (var i = 0; i < tables.Length; i++)
            {
                var rows = tables[i].Elements<TableRow>().ToArray();
                Assert.Equal(3, rows.Length);
                var baseRow = i == 0 ? 0 : 1;
                var derivedRow = i == 0 ? 1 : 2;
                Assert.Equal("AA0000", rows[baseRow].Elements<TableCell>().First()
                    .TableCellProperties?.TableCellBorders?.TopBorder?.Color?.Value);
                Assert.Equal("FFFF00", rows[derivedRow].Elements<TableCell>().First()
                    .TableCellProperties?.TableCellBorders?.TopBorder?.Color?.Value);
                Assert.Equal("00AA00", rows[derivedRow].Elements<TableCell>().First()
                    .TableCellProperties?.TableCellBorders?.LeftBorder?.Color?.Value);
            }
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
    }

    [Fact]
    public void WordBandBordersAlternateAcrossBodyHeaderAndFooterRows()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordConditionalBandBordersThreeRowsAllStories";
        foreach (var binary in new[]
        {
            File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
            DxpDocExport.Export(File.ReadAllBytes(Path.Combine(directory,
                name + ".docx")))
        })
        {
            using var stream = new MemoryStream(DxpDocToDocx.Project(binary).DocxBytes);
            using var document = WordprocessingDocument.Open(stream, false);
            var main = document.MainDocumentPart!;
            var body = Assert.Single(main.Document.Body!.Elements<Table>());
            var header = Assert.Single(main.HeaderParts.SelectMany(x =>
                x.Header!.Elements<Table>()));
            var footer = Assert.Single(main.FooterParts.SelectMany(x =>
                x.Footer!.Elements<Table>()));
            static string?[] TopColors(Table table) => table.Elements<TableRow>()
                .Select(row => row.Elements<TableCell>().First()
                    .TableCellProperties?.TableCellBorders?.TopBorder?.Color?.Value)
                .ToArray();
            Assert.Equal(new[] { "AA0000", "0000AA", "AA0000" },
                TopColors(body));
            foreach (var table in new[] { header, footer })
            {
                var colors = TopColors(table);
                Assert.Equal(3, colors.Length);
                Assert.Equal("AA0000", colors[1]);
                Assert.Equal("0000AA", colors[2]);
            }
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
    }

    [Fact]
    public void WordBandBorderDirectOverridesWinInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordConditionalBandDirectOverrideAllStories";
        foreach (var binary in new[]
        {
            File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
            DxpDocExport.Export(File.ReadAllBytes(Path.Combine(directory,
                name + ".docx")))
        })
        {
            using var stream = new MemoryStream(DxpDocToDocx.Project(binary).DocxBytes);
            using var document = WordprocessingDocument.Open(stream, false);
            var main = document.MainDocumentPart!;
            var body = Assert.Single(main.Document.Body!.Elements<Table>());
            var header = Assert.Single(main.HeaderParts.SelectMany(x =>
                x.Header!.Elements<Table>()));
            var footer = Assert.Single(main.FooterParts.SelectMany(x =>
                x.Footer!.Elements<Table>()));
            Assert.Equal("000000", body.Elements<TableRow>().ElementAt(1)
                .Elements<TableCell>().ElementAt(1).TableCellProperties?
                .TableCellBorders?.LeftBorder?.Color?.Value);
            Assert.Equal("000000", header.Elements<TableRow>().Single()
                .Elements<TableCell>().First().TableCellProperties?
                .TableCellBorders?.LeftBorder?.Color?.Value);
            Assert.Equal("000000", footer.Elements<TableRow>().Single()
                .Elements<TableCell>().ElementAt(1).TableCellProperties?
                .TableCellBorders?.LeftBorder?.Color?.Value);
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
    }

    [Fact]
    public void HorizontalAndVerticalBandStyleBordersSurviveGeneratedDocHops()
    {
        var fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordConditionalFirstRowParagraphAllStories.docx"));
        using var source = new MemoryStream(File.ReadAllBytes(fixture));
        using (var document = WordprocessingDocument.Open(source, true))
        {
            var main = document.MainDocumentPart!;
            var style = Assert.Single(main.StyleDefinitionsPart!.Styles!
                .Elements<Style>(), x => x.StyleId?.Value == "LogicalStart");
            style.StyleTableProperties!.PrependChild(
                new TableStyleRowBandSize { Val = 1 });
            style.StyleTableProperties.PrependChild(
                new TableStyleColumnBandSize { Val = 1 });
            foreach (var (kind, color) in new[]
            {
                (TableStyleOverrideValues.Band1Horizontal, "AA0000"),
                (TableStyleOverrideValues.Band2Horizontal, "0000AA")
            })
            {
                var rule = new TableStyleProperties { Type = kind };
                rule.Append(new TableStyleConditionalFormattingTableCellProperties(
                    new TableCellBorders(new TopBorder
                    {
                        Val = BorderValues.Single, Size = 8, Color = color
                    })));
                style.Append(rule);
            }
            foreach (var (kind, color) in new[]
            {
                (TableStyleOverrideValues.Band1Vertical, "00AA00"),
                (TableStyleOverrideValues.Band2Vertical, "AA00AA")
            })
            {
                var rule = new TableStyleProperties { Type = kind };
                rule.Append(new TableStyleConditionalFormattingTableCellProperties(
                    new TableCellBorders(new LeftBorder
                    {
                        Val = BorderValues.Single, Size = 8, Color = color
                    })));
                style.Append(rule);
            }
            main.StyleDefinitionsPart.Styles.Save();
            foreach (var table in main.Document.Body!.Elements<Table>())
            {
                var look = table.TableProperties!.GetFirstChild<TableLook>()!;
                look.FirstRow = false;
                look.LastRow = false;
                look.FirstColumn = false;
                look.LastColumn = false;
                look.NoHorizontalBand = false;
                look.NoVerticalBand = false;
            }
            main.Document.Save();
        }
        var generated = DxpDocExport.Export(source.ToArray());
        var projected = DxpDocToDocx.Project(generated).DocxBytes;
        using (var projectedStream = new MemoryStream(projected))
        using (var document = WordprocessingDocument.Open(projectedStream, false))
        {
            var cells = document.MainDocumentPart!.Document.Body!
                .Elements<Table>().SelectMany(x => x.Elements<TableRow>())
                .SelectMany(x => x.Elements<TableCell>()).ToArray();
            Assert.Contains(cells, x => x.TableCellProperties?
                .TableCellBorders?.TopBorder?.Color?.Value == "AA0000");
            Assert.Contains(cells, x => x.TableCellProperties?
                .TableCellBorders?.TopBorder?.Color?.Value == "0000AA");
            Assert.Contains(cells, x => x.TableCellProperties?
                .TableCellBorders?.LeftBorder?.Color?.Value == "00AA00");
            Assert.Contains(cells, x => x.TableCellProperties?
                .TableCellBorders?.LeftBorder?.Color?.Value == "AA00AA");
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        foreach (var binary in new[] { generated, DxpDocExport.Export(projected) })
        {
            using var stream = new MemoryStream(binary);
            using var index = new DocTextIndexWalker().Index(stream);
            var rows = index.ParagraphStyles.Where(x =>
                x.Formatting?.TableTerminator == true).ToArray();
            Assert.Equal(4, rows.Length);
            var colors = rows.Select(x =>
                x.Formatting?.TableCellBorders?[0]?.Top?.ColorRgb).ToArray();
            Assert.Contains(0x000000AAu, colors);
            Assert.Contains(0x00AA0000u, colors);
            var leftColors = rows.SelectMany(x =>
                x.Formatting?.TableCellBorders?.Select(b => b?.Left?.ColorRgb) ??
                Array.Empty<uint?>()).ToArray();
            Assert.Contains(0x0000AA00u, leftColors);
            Assert.Contains(0x00AA00AAu, leftColors);
        }
        if (Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_WORD") == "1")
        {
            var docPath = Path.Combine(Path.GetTempPath(),
                $"docxport-band-borders-{Guid.NewGuid():N}.doc");
            var docxPath = Path.ChangeExtension(docPath, ".docx");
            File.WriteAllBytes(docPath, generated);
            dynamic word = Activator.CreateInstance(
                Type.GetTypeFromProgID("Word.Application")!)!;
            word.Visible = false;
            word.DisplayAlerts = 0;
            dynamic? opened = null;
            try
            {
                opened = word.Documents.Open(docPath, ReadOnly: true,
                    AddToRecentFiles: false);
                opened.SaveAs2(docxPath, 16);
                opened.Close(false);
                opened = null;
                using var saved = WordprocessingDocument.Open(docxPath, false);
                var style = Assert.Single(saved.MainDocumentPart!
                    .StyleDefinitionsPart!.Styles!.Elements<Style>(), x =>
                    x.StyleName?.Val?.Value == "Logical Start");
                foreach (var (kind, color, edge) in new[]
                {
                    (TableStyleOverrideValues.Band1Horizontal, "AA0000", "top"),
                    (TableStyleOverrideValues.Band2Horizontal, "0000AA", "top"),
                    (TableStyleOverrideValues.Band1Vertical, "00AA00", "left"),
                    (TableStyleOverrideValues.Band2Vertical, "AA00AA", "left")
                })
                {
                    var rule = style.Elements<TableStyleProperties>().Single(x =>
                        x.Type?.Value == kind);
                    var border = rule.Descendants<TableCellBorders>().Single()
                        .ChildElements.Single(x => x.LocalName == edge);
                    Assert.Equal(color, border.GetAttribute("color",
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value);
                }
            }
            finally
            {
                if (opened != null) opened.Close(false);
                word.Quit();
                File.Delete(docPath);
                File.Delete(docxPath);
            }
        }
    }

    [Fact]
    public void StyledStartAndDirectEndBordersAreBothFlattenedIntoDocCells()
    {
        var fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyleDirectLogicalTableBordersAllStories.docx"));
        var binary = DxpDocExport.Export(File.ReadAllBytes(fixture));
        using var stream = new MemoryStream(binary);
        using var index = new DocTextIndexWalker().Index(stream);
        var rows = index.ParagraphStyles.Where(x =>
            x.Formatting?.TableTerminator == true).ToArray();
        Assert.Contains(rows, row =>
            row.Formatting?.TableCellBorders?[0]?.Left?.ColorRgb == 0x0000FFu &&
            row.Formatting.TableCellBorders?[1]?.Right?.ColorRgb == 0x008000u);
    }

    [Fact]
    public void TablePropertyExceptionBorderAppliesOnlyToItsRowInDoc()
    {
        var fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRowBorderExceptionsAllStories.docx"));
        using var stream = new MemoryStream(DxpDocExport.Export(File.ReadAllBytes(fixture)));
        using var index = new DocTextIndexWalker().Index(stream);
        var rows = index.ParagraphStyles.Where(x =>
            x.Formatting?.TableTerminator == true).Select(x => x.Formatting!).ToArray();
        Assert.Equal(4, rows.Length);
        Assert.Equal(3, rows.Count(x =>
            x.TableCellBorders?[0]?.Top?.ColorRgb == 0x0088FFu));
        Assert.Contains(rows, x => x.TableCellBorders?[0]?.Top == null);
    }

    [Fact]
    public void TablePropertyExceptionCellMarginsApplyOnlyToTheirRows()
    {
        var fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRowCellMarginsAllStories.docx"));
        using var stream = new MemoryStream(DxpDocExport.Export(File.ReadAllBytes(fixture)));
        using var index = new DocTextIndexWalker().Index(stream);
        var rows = index.ParagraphStyles.Where(x =>
            x.Formatting?.TableTerminator == true).Select(x => x.Formatting!).ToArray();
        Assert.Equal(4, rows.Length);
        Assert.Equal(3, rows.Count(x =>
            x.TableDefaultCellMargins?.Left == 360 &&
            x.TableDefaultCellMargins?.Right == 108));
        Assert.Contains(rows, x => x.TableDefaultCellMargins?.Left != 360);
    }

    [Fact]
    public void RowCellMarginsRemainRowSpecificAfterBothDocRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordRowCellMarginsAllStories.docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            "WordRowCellMarginsAllStories.doc"));
        static int? LeftMargin(DocumentFormat.OpenXml.OpenXmlElement? margins)
        {
            var side = margins?.ChildElements.FirstOrDefault(x =>
                x.LocalName is "left" or "start");
            return int.TryParse(side?.GetAttribute("w",
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value,
                out var value) ? value : null;
        }
        static int[] EffectiveLeftMargins(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes);
            using var document = WordprocessingDocument.Open(stream, false);
            var main = document.MainDocumentPart!;
            var tables = main.Document!.Body!.Elements<Table>()
                .Concat(main.HeaderParts.SelectMany(x => x.Header!.Elements<Table>()))
                .Concat(main.FooterParts.SelectMany(x => x.Footer!.Elements<Table>()));
            return tables.SelectMany(table =>
            {
                var tableMargin = LeftMargin(table.TableProperties?
                    .GetFirstChild<TableCellMarginDefault>()) ?? 108;
                return table.Elements<TableRow>().Select(row =>
                    LeftMargin(row.Elements<TableCell>().First().TableCellProperties?
                        .GetFirstChild<TableCellMargin>()) ??
                    LeftMargin(row.TablePropertyExceptions?
                        .GetFirstChild<TableCellMarginDefault>()) ?? tableMargin);
            }).ToArray();
        }
        var expected = EffectiveLeftMargins(source);
        Assert.Equal(new[] { 360, 108, 360, 360 }, expected);
        foreach (var binary in new[] { native, DxpDocExport.Export(source) })
        {
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            Assert.Equal(expected, EffectiveLeftMargins(projected));
            using (var stream = new MemoryStream(projected))
            using (var document = WordprocessingDocument.Open(stream, false))
                Assert.Empty(new OpenXmlValidator().Validate(document));
            var thirdHop = DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes;
            Assert.Equal(expected, EffectiveLeftMargins(thirdHop));
        }
    }

    public static IEnumerable<object[]> PairedFixtures()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        return Directory.EnumerateFiles(directory, "*.docx")
            .Select(Path.GetFileNameWithoutExtension)
            .Where(name => name != null && File.Exists(Path.Combine(directory, name + ".doc")))
            .OrderBy(name => name, StringComparer.Ordinal)
            .Select(name => new object[] { name! });
    }

    [Fact]
    public void PairedPngPayloadsSurviveBothDocRoutesAndAnotherHop()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var checkedFixtures = 0;
        foreach (var fixture in PairedFixtures())
        {
            var name = (string)fixture[0];
            var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
            var expected = ReadDistinctPngPayloadHashes(source);
            if (expected.Length == 0) continue;
            checkedFixtures++;
            foreach (var (binary, nativeRoute) in new[]
            {
                (File.ReadAllBytes(Path.Combine(directory, name + ".doc")), true),
                (DxpDocExport.Export(source), false)
            })
            {
                var projected = DxpDocToDocx.Project(binary).DocxBytes;
                var projectedHashes = ReadDistinctPngPayloadHashes(projected);
                if (name == "WordTransformedFloatingImageStories" && nativeRoute)
                {
                    Assert.Equal(2, projectedHashes.Length);
                    Assert.Contains(expected[0], projectedHashes);
                }
                else Assert.Equal(expected, projectedHashes);
                var repeated = DxpDocToDocx.Project(
                    DxpDocExport.Export(projected)).DocxBytes;
                Assert.Equal(projectedHashes, ReadDistinctPngPayloadHashes(repeated));
            }
        }
        Assert.True(checkedFixtures >= 39,
            $"Only {checkedFixtures} paired PNG fixtures were checked.");
    }

    private static string[] ReadDistinctPngPayloadHashes(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var archive = new System.IO.Compression.ZipArchive(stream,
            System.IO.Compression.ZipArchiveMode.Read);
        return archive.Entries.Where(entry =>
                entry.FullName.Contains("/media/", StringComparison.Ordinal) ||
                entry.FullName.StartsWith("media/", StringComparison.Ordinal))
            .Where(entry => entry.FullName.EndsWith(".png",
                StringComparison.OrdinalIgnoreCase))
            .Select(entry =>
            {
                using var payload = entry.Open();
                using var copy = new MemoryStream();
                payload.CopyTo(copy);
                return Convert.ToHexString(
                    System.Security.Cryptography.SHA256.HashData(copy.ToArray()));
            })
            .Distinct(StringComparer.Ordinal).OrderBy(x => x, StringComparer.Ordinal)
            .ToArray();
    }

    [Fact]
    public void PairedDocCharacterRunsStayWithinOnePhysicalTextPiece()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var checkedRuns = 0;
        var multiPieceRuns = 0;
        foreach (var fixture in PairedFixtures())
        {
            var name = (string)fixture[0];
            using var stream = File.OpenRead(Path.Combine(directory, name + ".doc"));
            using var index = new DocTextIndexWalker().Index(stream);
            var stories = new[] { index.ReadStory("Main") }
                .Concat(index.HeaderStories.Where(x => !x.IsEmpty)
                    .Select(x => x.ReadIndexedContent()));
            foreach (var story in stories)
            foreach (var run in story.Runs)
            {
                var start = Assert.IsType<uint>(run.SourceCpStart);
                var end = Assert.IsType<uint>(run.SourceCpEnd);
                var piece = Assert.Single(index.Pieces, x =>
                    x.CpStart <= start && end <= x.CpEnd && start < end);
                Assert.Equal(piece.GetTextOffset(start), run.SourceTextOffsetStart);
                Assert.Equal(piece.GetTextOffset(end), run.SourceTextOffsetEnd);
                checkedRuns++;
                if (index.Pieces.Count > 1) multiPieceRuns++;
            }
        }
        Assert.True(checkedRuns > 100, $"Only {checkedRuns} indexed runs were checked.");
        Assert.True(multiPieceRuns > 0, "No multi-piece formatting was checked.");
    }
    [Fact]
    public void EveryPairedFixtureHasANamedWordRenderRoute()
    {
        static IEnumerable<string> FixtureNames(string methodName)
        {
            var method = typeof(DocStructureWalkerTests).GetMethod(methodName)!;
            return method.CustomAttributes
                .Where(attribute => attribute.AttributeType.FullName == "Xunit.InlineDataAttribute")
                .Select(attribute => ((IList<System.Reflection.CustomAttributeTypedArgument>)
                    attribute.ConstructorArguments[0].Value!).First().Value as string)
                .Where(name => name != null)
                .Select(name => Path.GetFileNameWithoutExtension(name!));
        }

        var covered = FixtureNames("WordAuthoredStoriesRenderCloseToReference")
            .Concat(FixtureNames("WordBmpImagesRoundTripThroughAllStories"))
            .ToHashSet(StringComparer.Ordinal);
        var paired = PairedFixtures().Select(fixture => (string)fixture[0]).ToArray();
        var missing = paired.Where(name => !covered.Contains(name)).ToArray();
        Assert.Equal(new[]
        {
            "WordVisibleNoWrapAllStories",
            "WordVisibleNoWrapOffAllStories"
        }, missing);
        var relaxed = FixtureNames(
            "AutoWidthNoWrapRenderStaysCloserToSourceThanWordNativeDoc")
            .ToHashSet(StringComparer.Ordinal);
        Assert.All(missing, name => Assert.Contains(name, relaxed));
        Assert.Equal(paired.Length, paired.Count(name => covered.Contains(name) ||
            relaxed.Contains(name)));
    }
    [Fact]
    public void InheritedHeaderFooterTablesKeepEditablePageFields()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordInheritedTablePageFields";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var binaries = new[]
        {
            File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
            DxpDocExport.Export(source)
        };
        foreach (var bytes in new[] { source }.Concat(binaries.Select(x =>
            DxpDocToDocx.Project(x).DocxBytes)).Concat(binaries.Select(x =>
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(x).DocxBytes)).DocxBytes)))
        {
            using var package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
            var main = package.MainDocumentPart!;
            var sections = main.Document.Body!.Descendants<SectionProperties>().ToArray();
            Assert.Equal(3, sections.Length);
            Assert.Empty(sections[2].Elements<HeaderReference>());
            Assert.Empty(sections[2].Elements<FooterReference>());
            foreach (var header in new[] { true, false })
            {
                var id = header
                    ? sections[1].Elements<HeaderReference>().Single(x =>
                        x.Type?.Value == HeaderFooterValues.Default).Id!.Value!
                    : sections[1].Elements<FooterReference>().Single(x =>
                        x.Type?.Value == HeaderFooterValues.Default).Id!.Value!;
                var root = header ? (OpenXmlElement)((HeaderPart)main.GetPartById(id))
                    .Header! : ((FooterPart)main.GetPartById(id)).Footer!;
                var cells = Assert.Single(root.Elements<Table>())
                    .Elements<TableRow>().Single().Elements<TableCell>().ToArray();
                Assert.Equal(2, cells.Length);
                var field = cells[1].Descendants<SimpleField>()
                    .Select(x => x.Instruction?.Value)
                    .Concat(cells[1].Descendants<FieldCode>().Select(x => x.Text))
                    .Where(x => x != null).ToArray();
                Assert.Single(field);
                Assert.Matches(@"(?i)\bPAGE\b", field[0]!);
                var styleId = cells[1].Descendants<Paragraph>().First()
                    .ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                Assert.Equal("Six Story Derived", main.StyleDefinitionsPart!.Styles!
                    .Elements<Style>().Single(x => x.StyleId?.Value == styleId)
                    .StyleName?.Val?.Value);
            }
            Assert.Empty(new OpenXmlValidator().Validate(package));
        }
        Assert.Equal(ReadEffectiveHeaderFooterStories(source),
            ReadEffectiveHeaderFooterStories(DxpDocToDocx.Project(
                DxpDocExport.Export(DxpDocToDocx.Project(binaries[1]).DocxBytes))
                .DocxBytes));
    }
    [Fact]
    public void InheritedTableStoriesKeepImagesAndEditablePageFields()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordInheritedTableMixedContent";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var binaries = new[]
        {
            File.ReadAllBytes(Path.Combine(directory, name + ".doc")),
            DxpDocExport.Export(source)
        };
        byte[]? expectedImage = null;
        foreach (var bytes in new[] { source }.Concat(binaries.Select(x =>
            DxpDocToDocx.Project(x).DocxBytes)).Concat(binaries.Select(x =>
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(x).DocxBytes)).DocxBytes)))
        {
            using var package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
            var main = package.MainDocumentPart!;
            var sections = main.Document.Body!.Descendants<SectionProperties>().ToArray();
            Assert.Equal(3, sections.Length);
            Assert.Empty(sections[2].Elements<HeaderReference>());
            Assert.Empty(sections[2].Elements<FooterReference>());
            foreach (var header in new[] { true, false })
            {
                var id = header
                    ? sections[1].Elements<HeaderReference>().Single(x =>
                        x.Type?.Value == HeaderFooterValues.Default).Id!.Value!
                    : sections[1].Elements<FooterReference>().Single(x =>
                        x.Type?.Value == HeaderFooterValues.Default).Id!.Value!;
                var root = header ? (OpenXmlElement)((HeaderPart)main.GetPartById(id))
                    .Header! : ((FooterPart)main.GetPartById(id)).Footer!;
                var cells = Assert.Single(root.Elements<Table>())
                    .Elements<TableRow>().Single().Elements<TableCell>().ToArray();
                Assert.Equal(2, cells.Length);
                var drawing = Assert.Single(cells[0].Descendants<Drawing>());
                var blip = Assert.Single(drawing.Descendants<
                    DocumentFormat.OpenXml.Drawing.Blip>());
                var part = header ? (OpenXmlPart)(HeaderPart)main.GetPartById(id)
                    : (FooterPart)main.GetPartById(id);
                var image = Assert.IsType<ImagePart>(part.GetPartById(
                    blip.Embed!.Value!));
                Assert.Equal("image/png", image.ContentType);
                using var payload = image.GetStream();
                using var copied = new MemoryStream();
                payload.CopyTo(copied);
                var imageBytes = copied.ToArray();
                Assert.True(imageBytes.Length > 100);
                if (expectedImage == null) expectedImage = imageBytes;
                else Assert.Equal(expectedImage, imageBytes);
                var field = cells[1].Descendants<SimpleField>()
                    .Select(x => x.Instruction?.Value)
                    .Concat(cells[1].Descendants<FieldCode>().Select(x => x.Text))
                    .Where(x => x != null).ToArray();
                Assert.Single(field);
                Assert.Matches(@"(?i)\bPAGE\b", field[0]!);
                var styleId = cells[1].Descendants<Paragraph>().First()
                    .ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                Assert.Equal("Six Story Derived", main.StyleDefinitionsPart!.Styles!
                    .Elements<Style>().Single(x => x.StyleId?.Value == styleId)
                    .StyleName?.Val?.Value);
            }
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2010)
                .Validate(package));
        }
        Assert.Equal(ReadEffectiveHeaderFooterStories(source),
            ReadEffectiveHeaderFooterStories(DxpDocToDocx.Project(
                DxpDocExport.Export(DxpDocToDocx.Project(binaries[1]).DocxBytes))
                .DocxBytes));
    }
    [Theory]
    [InlineData("WordSixSlotInheritedStyledTables", false)]
    [InlineData("WordInheritedStyledRotatedTables", true)]
    public void ThreeSectionStyledHeaderFooterTablesRemainInherited(
        string name, bool rotated)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        foreach (var bytes in new[]
        {
            source,
            DxpDocToDocx.Project(File.ReadAllBytes(Path.Combine(directory,
                name + ".doc"))).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(source)).DocxBytes
        })
        {
            using var package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
            var main = package.MainDocumentPart!;
            var sections = main.Document.Body!.Descendants<SectionProperties>().ToArray();
            Assert.Equal(3, sections.Length);
            Assert.Empty(sections[2].Elements<HeaderReference>());
            Assert.Empty(sections[2].Elements<FooterReference>());
            foreach (var (tag, header) in new[] { ("Header", true), ("Footer", false) })
            {
                var id = header
                    ? sections[1].Elements<HeaderReference>().Single(x =>
                        x.Type?.Value == HeaderFooterValues.Default).Id!.Value!
                    : sections[1].Elements<FooterReference>().Single(x =>
                        x.Type?.Value == HeaderFooterValues.Default).Id!.Value!;
                var root = header ? (OpenXmlElement)((HeaderPart)main.GetPartById(id))
                    .Header! : ((FooterPart)main.GetPartById(id)).Footer!;
                var table = Assert.Single(root.Elements<Table>());
                var cells = Assert.Single(table.Elements<TableRow>())
                    .Elements<TableCell>().ToArray();
                Assert.Equal(new[] { tag + " cell A", tag + " cell B" },
                    cells.Select(x => x.InnerText));
                var styles = main.StyleDefinitionsPart!.Styles!.Elements<Style>()
                    .ToDictionary(x => x.StyleId!.Value!, x => x.StyleName?.Val?.Value);
                Assert.All(cells, cell =>
                {
                    var styleId = Assert.Single(cell.Descendants<Paragraph>())
                        .ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                    Assert.Equal("Six Story Derived", styles[styleId!]);
                });
                Assert.Equal("C0504D", cells[1].Descendants<Color>()
                    .Single().Val?.Value);
                if (rotated)
                    Assert.Equal(header ? new[] { "tbRl", "lrTbV" } :
                        new[] { "btLr", "tbRlV" }, cells.Select(cell =>
                            cell.TableCellProperties?.GetFirstChild<TextDirection>()?
                                .Val?.InnerText));
            }
            Assert.Empty(new OpenXmlValidator().Validate(package));
        }
        var expected = ReadEffectiveHeaderFooterStories(source);
        Assert.Equal(expected, ReadEffectiveHeaderFooterStories(
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(DxpDocExport.Export(source)).DocxBytes))
                .DocxBytes));
    }
    [Fact]
    public void PairedCorpusExercisesLinkedParagraphAndCharacterStyles()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var name = "WordLinkedStylesAllStories";
        Assert.Contains(PairedFixtures(), x => (string)x[0] == name);
        var links = ReadLinkedCustomStyles(File.ReadAllBytes(Path.Combine(
            directory, name + ".docx")));
        Assert.Equal("Linked character", links["Linked paragraph"]);
        Assert.Equal("Linked paragraph", links["Linked character"]);
    }

    [Fact]
    public void CommonBuiltInParagraphStylesUseWordNativeInvariantIds()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordArbitraryRotationStories.docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            "WordArbitraryRotationStories.doc"));
        var expected = new Dictionary<string, int>
        {
            ["Title"] = 62, ["Subtitle"] = 74,
            ["Quote"] = 180, ["Intense Quote"] = 181,
            ["Intense Emphasis"] = 261, ["Intense Reference"] = 263
        };
        foreach (var bytes in new[] { native, DxpDocExport.Export(source) })
        {
            using var input = new MemoryStream(bytes);
            using var index = new DocTextIndexWalker().Index(input);
            foreach (var (name, invariant) in expected)
                Assert.Equal(invariant, Assert.Single(index.StyleDefinitions,
                    x => x.Name == name).InvariantStyleId);
            using var projection = new MemoryStream(DxpDocToDocx.Project(bytes).DocxBytes);
            using var docx = WordprocessingDocument.Open(projection, false);
            var styles = docx.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().ToArray();
            foreach (var name in expected.Keys)
                Assert.False(Assert.Single(styles,
                    x => x.StyleName?.Val?.Value == name).CustomStyle?.Value);
            Assert.Empty(new OpenXmlValidator().Validate(docx));
        }
    }

    [Fact]
    public void GeneratedDocMatchesWordNativeBuiltInStyleIdentifiersAcrossCorpus()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        foreach (var fixture in PairedFixtures())
        {
            var name = (string)fixture[0];
            using var native = new DocTextIndexWalker().Index(Path.Combine(directory,
                name + ".doc"));
            using var generatedStream = new MemoryStream(DxpDocExport.Export(
                File.ReadAllBytes(Path.Combine(directory, name + ".docx"))));
            using var generated = new DocTextIndexWalker().Index(generatedStream);
            var byName = generated.StyleDefinitions.GroupBy(x => x.Name,
                StringComparer.Ordinal).ToDictionary(x => x.Key,
                x => x.FirstOrDefault(style => style.InvariantStyleId != 0x0FFE) ??
                    x.First(), StringComparer.Ordinal);
            foreach (var style in native.StyleDefinitions.Where(x =>
                x.InvariantStyleId is not (null or 0x0FFE)))
                if (byName.TryGetValue(style.Name, out var written))
                    Assert.True(style.InvariantStyleId == written.InvariantStyleId,
                        $"{name}: {style.Name} has Word sti {style.InvariantStyleId}, " +
                        $"generated sti {written.InvariantStyleId}");
        }
    }

    [Fact]
    public void WordNativeDocFlattensCharacterStyleBaseWhileGeneratedDocRetainsIt()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordThreeLevelStyleOverrides.docx"));
        var source = File.ReadAllBytes(path);
        var native = File.ReadAllBytes(Path.ChangeExtension(path, ".doc"));
        var generatedProjection = DxpDocToDocx.Project(
            DxpDocExport.Export(source)).DocxBytes;
        var first = DxpDocToDocx.Project(native).DocxBytes;
        var third = DxpDocToDocx.Project(DxpDocExport.Export(first)).DocxBytes;
        var key = "GrandchildDisplay Char.basedOn";
        string Value(byte[] bytes) => ReadStyleRelationships(bytes, false)
            .TryGetValue(key, out var value) ? value : "<missing>";
        using var nativeIndex = new DocTextIndexWalker().Index(
            new MemoryStream(native));
        var nativeStyle = Assert.Single(nativeIndex.StyleDefinitions,
            x => x.Name == "GrandchildDisplay Char");
        Assert.Equal("DerivedDisplay Char", Value(source));
        Assert.Equal("DerivedDisplay Char", Value(generatedProjection));
        Assert.Null(nativeStyle.BasedOnIndex);
        Assert.Equal("<missing>", Value(first));
        Assert.Equal("<missing>", Value(third));
    }

    [Theory]
    [InlineData("WordLogicalCellBordersAllStories", null)]
    [InlineData("WordLogicalLeftRtlHeader", 0)]
    [InlineData("WordLogicalRightRtlHeader", 2)]
    public void WordRtlTablePhysicalAndLogicalJustificationStayDistinct(
        string name, int? logicalAlignment)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var binary = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
        using var input = new MemoryStream(binary);
        using var index = new DocTextIndexWalker().Index(input);
        var row = Assert.Single(index.ParagraphStyles, x =>
            x.Formatting.TableTerminator == true &&
            x.Formatting.TableRightToLeft == true);
        var nativeAlignment = name == "WordLogicalLeftRtlHeader"
            ? null : logicalAlignment;
        Assert.Equal(nativeAlignment, (int?)row.Formatting.TableJustification);
        var projected = DxpDocToDocx.Project(binary).DocxBytes;
        var generated = DxpDocToDocx.Project(DxpDocExport.Export(File.ReadAllBytes(
            Path.Combine(directory, name + ".docx")))).DocxBytes;
        var repeated = DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes;
        foreach (var (bytes, expectedAlignment) in new[]
            { (projected, nativeAlignment), (generated, logicalAlignment),
                (repeated, nativeAlignment) })
        {
            using var stream = new MemoryStream(bytes);
            using var document = WordprocessingDocument.Open(stream, false);
            var header = Assert.Single(document.MainDocumentPart!.HeaderParts,
                x => x.Header!.InnerText.Contains("Header A", StringComparison.Ordinal));
            var table = Assert.Single(header.Header!.Elements<Table>());
            Assert.NotNull(table.TableProperties?.GetFirstChild<BiDiVisual>());
            Assert.Equal(expectedAlignment switch
                {
                    0 => TableRowAlignmentValues.Left,
                    2 => TableRowAlignmentValues.Right,
                    _ => (TableRowAlignmentValues?)null
                },
                table.TableProperties?.GetFirstChild<TableJustification>()?.Val?.Value);
        }
    }

    [Theory]
    [InlineData("WordSixSlotStyledTableContent", false)]
    [InlineData("WordSixSlotStyledTablePageFields", false)]
    [InlineData("WordSixSlotCellBorderOverride", true)]
    [InlineData("WordSixSlotDiagonalCellBorders", false)]
    public void WordSavedFirstPageFooterTableRetainsBorders(string name,
        bool suppressFirstTop)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        using var input = File.OpenRead(Path.Combine(directory, name + ".doc"));
        using var index = new DocTextIndexWalker().Index(input);
        var rows = index.ParagraphStyles.Where(x =>
            x.Formatting?.TableTerminator == true).ToArray();
        Assert.Equal(2, rows.Length);
        Assert.All(rows, row =>
        {
            Assert.Equal((byte)1, row.Formatting!.TableBorders?.Top?.Type);
            if (suppressFirstTop && row == rows[0])
                Assert.Equal((byte)0, row.Formatting.TableCellBorders?[0]?.Top?.Type);
            else Assert.Null(row.Formatting.TableCellBorders?.FirstOrDefault()?.Top);
            if (name == "WordSixSlotDiagonalCellBorders" && row == rows[0])
            {
                Assert.Equal((byte)1, row.Formatting.TableCellBorders?[0]?
                    .TopLeftToBottomRight?.Type);
                Assert.Equal((byte)1, row.Formatting.TableCellBorders?[1]?
                    .TopRightToBottomLeft?.Type);
            }
        });
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var projected = DxpDocToDocx.Project(File.ReadAllBytes(Path.Combine(
            directory, name + ".doc"))).DocxBytes;
        var generated = DxpDocToDocx.Project(DxpDocExport.Export(source)).DocxBytes;
        var repeated = DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes;
        foreach (var (route, bytes) in new[]
            {
                ("native", projected), ("generated", generated),
                ("rewritten", repeated)
            })
        {
            using var stream = new MemoryStream(bytes);
            using var document = WordprocessingDocument.Open(stream, false);
            var footer = Assert.Single(document.MainDocumentPart!.FooterParts,
                x => x.Footer!.InnerText.Contains("Footer A", StringComparison.Ordinal));
            var table = Assert.Single(footer.Footer!.Elements<Table>());
            var cells = table.Descendants<TableCell>().ToArray();
            for (var i = 0; i < cells.Length; i++)
            {
                var borders = cells[i].TableCellProperties?.TableCellBorders;
                Assert.True(borders?.TopBorder?.Val?.Value ==
                    (suppressFirstTop && i == 0 ? BorderValues.Nil :
                        BorderValues.Single), $"{route} cell {i} top border");
                Assert.Equal(BorderValues.Single, borders?.LeftBorder?.Val?.Value);
                Assert.Equal(BorderValues.Single, borders?.BottomBorder?.Val?.Value);
                Assert.Equal(BorderValues.Single, borders?.RightBorder?.Val?.Value);
                if (name == "WordSixSlotDiagonalCellBorders" && i == 0)
                    Assert.Equal(BorderValues.Single, borders?
                        .TopLeftToBottomRightCellBorder?.Val?.Value);
                if (name == "WordSixSlotDiagonalCellBorders" && i == 1)
                    Assert.Equal(BorderValues.Single, borders?
                        .TopRightToBottomLeftCellBorder?.Val?.Value);
            }
        }
    }

    [Fact]
    public void WordPairedCorpusRetainsVisibleStoriesAfterThirdDocHop()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var failures = new List<string>();
        foreach (var fixture in PairedFixtures())
        {
            var name = (string)fixture[0];
            try
            {
                var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
                var doc = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
                var projected = DxpDocToDocx.Project(doc);
                Assert.Empty(projected.Coverage.OmittedCharacters);
                Assert.Empty(projected.Coverage.ApproximateCharacters);
                Assert.Empty(projected.Coverage.DeferredParts);
                var rewritten = DxpDocExport.Export(projected.DocxBytes);
                var repeated = DxpDocToDocx.Project(rewritten);
                Assert.Empty(repeated.Coverage.OmittedCharacters);
                Assert.Empty(repeated.Coverage.ApproximateCharacters);
                Assert.Empty(repeated.Coverage.DeferredParts);
                Assert.Equal(ReadStories(source), ReadStories(repeated.DocxBytes));
                Assert.Equal(ReadSupportedCoreProperties(source),
                    ReadSupportedCoreProperties(repeated.DocxBytes));
                AssertExplicitHyphenationSettings(source, repeated.DocxBytes);
                AssertDirectEmptyParagraphMarkFormatting(source, repeated.DocxBytes);
                AssertCustomStyleRunFormatting(source, repeated.DocxBytes);
                AssertCustomStyleParagraphFormatting(source, repeated.DocxBytes);
                AssertLinkedCustomStylesPreserved(source, repeated.DocxBytes);
                AssertCustomStyleRelationshipsPreserved(projected.DocxBytes,
                    repeated.DocxBytes);
                if (name is "WordReferenceWithStories" or "WordBalancedSpaceWithStories")
                    Assert.Equal(ReadBalanceSetting(source),
                        ReadBalanceSetting(repeated.DocxBytes));
                if (name == "WordExplicitColumnBreakStories")
                    Assert.Equal(ReadColumnBreaks(source),
                        ReadColumnBreaks(repeated.DocxBytes));
                Assert.Equal(ReadEditableFieldSemantics(source),
                    ReadEditableFieldSemantics(repeated.DocxBytes));
                Assert.Equal(ReadEffectiveFooterPageFieldCounts(source),
                    ReadEffectiveFooterPageFieldCounts(repeated.DocxBytes));
                Assert.Equal(ReadEffectiveFooterHyperlinkTargets(source),
                    ReadEffectiveFooterHyperlinkTargets(repeated.DocxBytes));
                Assert.Equal(ReadEffectiveHeaderHyperlinkTargets(source),
                    ReadEffectiveHeaderHyperlinkTargets(repeated.DocxBytes));
                Assert.Equal(ReadListSemantics(source),
                    ReadListSemantics(repeated.DocxBytes));
                Assert.Equal(ReadListSemantics(source, true),
                    ReadListSemantics(repeated.DocxBytes, true));
                Assert.Equal(ReadEffectiveHeaderFooterStories(source),
                    ReadEffectiveHeaderFooterStories(repeated.DocxBytes));
                AssertExplicitSectionLayout(source, repeated.DocxBytes);
                AssertEffectiveParagraphLayout(source, repeated.DocxBytes);
                AssertParagraphDecorations(source, repeated.DocxBytes);
                AssertEffectiveTabStops(source, repeated.DocxBytes);
                AssertDefaultTabInterval(source, repeated.DocxBytes);
                AssertMirrorMargins(source, repeated.DocxBytes);
                AssertGutterAtTop(source, repeated.DocxBytes);
                AssertEffectiveRunFormatting(source, repeated.DocxBytes,
                    name is "WordCharacterStyleToggleStories" or
                        "WordLinkedStylesAllStories" or
                        "WordCharacterStyleBaseToggleStories" or
                        "WordCharacterStyleStrikeOverrideStories" or
                        "WordCharacterStyleInheritedStrikeStories" or
                        "WordCharacterStyleRelativeStrikeStories",
                    name is "WordCharacterStyleInheritedStrikeStories" or
                        "WordCharacterStyleRelativeStrikeStories",
                    name == "WordVisibleThemeTintShadeAllStories");
                if (name == "WordTransformedFloatingImageStories")
                    AssertNativeFlattenedHeaderGeometry(source, repeated.DocxBytes);
                else AssertDrawingGeometry(source, repeated.DocxBytes);
                AssertTableGrid(source, repeated.DocxBytes);
                Assert.Equal(ReadSectionBodyText(source),
                    ReadSectionBodyText(repeated.DocxBytes));
                if (name is "WordEmptyParagraphsAllStories" or
                    "WordVisibleSizedEmptyParagraphsAllStories")
                    Assert.Equal(ReadInteriorEmptyParagraphMarks(source,
                            details: name == "WordVisibleSizedEmptyParagraphsAllStories"),
                        ReadInteriorEmptyParagraphMarks(repeated.DocxBytes,
                            details: name == "WordVisibleSizedEmptyParagraphsAllStories"));
                Assert.Empty(Validate(repeated.DocxBytes));
            }
            catch (Exception exception)
            {
                failures.Add(name + ": " + exception.Message.Split('\n')[0]);
            }
        }
        Assert.True(failures.Count == 0, string.Join(Environment.NewLine, failures));
    }

    [Fact]
    public void RenamedBuiltInHeadingRetainsAliasAndIdentityAcrossDocRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var docx = File.ReadAllBytes(Path.Combine(directory, "WordRenamedHeadingStories.docx"));
        var nativeDoc = File.ReadAllBytes(Path.Combine(directory, "WordRenamedHeadingStories.doc"));
        var generatedDoc = DxpDocExport.Export(docx);
        foreach (var doc in new[] { nativeDoc, generatedDoc })
        {
            using var input = new MemoryStream(doc);
            using var index = new DocTextIndexWalker().Index(input);
            Assert.Contains(index.StyleDefinitions, x => x.Name.Equals(
                "Heading 1,Chapter Title", StringComparison.OrdinalIgnoreCase) &&
                x.InvariantStyleId == 1);
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            using var stream = new MemoryStream(projected);
            using var opened = WordprocessingDocument.Open(stream, false);
            Assert.Contains(opened.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>(), x => string.Equals(x.StyleName?.Val?.Value,
                    "Heading 1", StringComparison.OrdinalIgnoreCase) &&
                x.Aliases?.Val?.Value == "Chapter Title" &&
                x.CustomStyle?.Value != true);
            Assert.Empty(Validate(projected));
        }
        var rewritten = DxpDocExport.Export(DxpDocToDocx.Project(nativeDoc).DocxBytes);
        using var repeatedInput = new MemoryStream(rewritten);
        using var repeatedIndex = new DocTextIndexWalker().Index(repeatedInput);
        Assert.Contains(repeatedIndex.StyleDefinitions, x => x.Name.Equals(
            "Heading 1,Chapter Title", StringComparison.OrdinalIgnoreCase) &&
            x.InvariantStyleId == 1);
    }

    [Fact]
    public void RenamedCustomStyleRetainsMultipleAliasesInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var docx = File.ReadAllBytes(Path.Combine(directory,
            "WordRenamedCustomStyleStories.docx"));
        var nativeDoc = File.ReadAllBytes(Path.Combine(directory,
            "WordRenamedCustomStyleStories.doc"));
        var generatedDoc = DxpDocExport.Export(docx);
        foreach (var doc in new[] { nativeDoc, generatedDoc,
            DxpDocExport.Export(DxpDocToDocx.Project(nativeDoc).DocxBytes) })
        {
            using var input = new MemoryStream(doc);
            using var index = new DocTextIndexWalker().Index(input);
            Assert.Contains(index.StyleDefinitions, x => x.Name ==
                "Final Accent,Previous Accent,Alternate Accent" &&
                x.InvariantStyleId == 0x0FFE);
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            using var stream = new MemoryStream(projected);
            using var opened = WordprocessingDocument.Open(stream, false);
            Assert.Contains(opened.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>(), x => x.StyleName?.Val?.Value == "Final Accent" &&
                x.Aliases?.Val?.Value == "Previous Accent,Alternate Accent" &&
                x.CustomStyle?.Value == true);
            Assert.Equal(ReadStories(docx), ReadStories(projected));
            Assert.Empty(Validate(projected));
        }
    }

    [Fact]
    public void CustomNextParagraphStyleRetainsLinkAcrossDocRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var docx = File.ReadAllBytes(Path.Combine(directory, "WordNextStyleStories.docx"));
        var nativeDoc = File.ReadAllBytes(Path.Combine(directory, "WordNextStyleStories.doc"));
        foreach (var doc in new[] { nativeDoc, DxpDocExport.Export(docx),
            DxpDocExport.Export(DxpDocToDocx.Project(nativeDoc).DocxBytes) })
        {
            using var input = new MemoryStream(doc);
            using var index = new DocTextIndexWalker().Index(input);
            var heading = Assert.Single(index.StyleDefinitions, x => x.Name == "Flow Heading");
            var body = Assert.Single(index.StyleDefinitions, x => x.Name == "Flow Text");
            Assert.Equal(body.Index, heading.NextIndex);
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            using var stream = new MemoryStream(projected);
            using var opened = WordprocessingDocument.Open(stream, false);
            var styles = opened.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().ToArray();
            var headingStyle = Assert.Single(styles,
                x => x.StyleName?.Val?.Value == "Flow Heading");
            var bodyStyle = Assert.Single(styles,
                x => x.StyleName?.Val?.Value == "Flow Text");
            Assert.Equal(bodyStyle.StyleId?.Value,
                headingStyle.NextParagraphStyle?.Val?.Value);
            Assert.Equal(ReadStories(docx), ReadStories(projected));
            Assert.Empty(Validate(projected));
        }
    }

    [Theory]
    [InlineData("WordMixedSectionStories")]
    [InlineData("WordMixedSectionStoryContent")]
    [InlineData("WordMixedSectionStyledStories")]
    public void WordSavedMixedSectionsCrossMainAndHeaderPieceBoundary(string name)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var sourceDocx = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var reference = ReadStories(sourceDocx);
        using var input = File.OpenRead(Path.Combine(directory, name + ".doc"));
        using var index = new DocTextIndexWalker().Index(input);
        var main = index.Parts.Single(x => x.Name == "Main");
        var headers = index.Parts.Single(x => x.Name == "Headers");
        Assert.Contains(index.Pieces, x => x.CpStart < main.CpEnd &&
            x.CpStart < headers.CpStart && x.CpEnd > headers.CpStart);
        var projected = DxpDocToDocx.Project(File.ReadAllBytes(Path.Combine(
            directory, name + ".doc")));
        Assert.Equal(reference, ReadStories(projected.DocxBytes));
        var rewritten = DxpDocExport.Export(projected.DocxBytes);
        var rewrittenProjection = DxpDocToDocx.Project(rewritten);
        Assert.Equal(reference, ReadStories(rewrittenProjection.DocxBytes));
        Assert.Equal(ReadEditableFieldSemantics(sourceDocx),
            ReadEditableFieldSemantics(rewrittenProjection.DocxBytes));
        Assert.Equal(ReadListSemantics(sourceDocx),
            ReadListSemantics(rewrittenProjection.DocxBytes));
        AssertEffectiveRunFormatting(sourceDocx, rewrittenProjection.DocxBytes);
        AssertEffectiveParagraphLayout(sourceDocx, rewrittenProjection.DocxBytes);
        AssertDrawingGeometry(sourceDocx, rewrittenProjection.DocxBytes);
        AssertTableGrid(sourceDocx, rewrittenProjection.DocxBytes);
        Assert.Equal(ReadEffectiveHeaderFooterStories(sourceDocx),
            ReadEffectiveHeaderFooterStories(rewrittenProjection.DocxBytes));
    }

    [Fact]
    public void PairedCorpusExercisesVisibleParagraphFlowProperties()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var properties = PairedFixtures()
            .SelectMany(x => ReadEffectiveParagraphLayout(File.ReadAllBytes(
                Path.Combine(directory, (string)x[0] + ".docx"))).Keys)
            .ToArray();
        Assert.Contains(properties, x => x.EndsWith(".line", StringComparison.Ordinal));
        Assert.Contains(properties, x => x.EndsWith(".keepNext", StringComparison.Ordinal));
        Assert.Contains(properties, x => x.EndsWith(".keepLines", StringComparison.Ordinal));
        var threeSections = File.ReadAllBytes(Path.Combine(directory,
            "WordThreeSectionStories.docx"));
        Assert.Equal("left|none", ReadEffectiveTabStops(threeSections)[
            "body.paragraph0.tab2160"]);
        Assert.StartsWith("Tab\t", ReadStories(threeSections)["body"]);
    }

    [Fact]
    public void PairedCorpusExercisesPageBreakBeforeAndWidowControl()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var layout = ReadEffectiveParagraphLayout(File.ReadAllBytes(Path.Combine(
            directory, "WordPaginationAllStories.docx")));
        Assert.Equal("true", layout["body.paragraph2.pageBreakBefore"]);
        Assert.Equal("false", layout["body.paragraph0.widowControl"]);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PairedCorpusExercisesInheritedCjkBreakingInAllStories(bool native)
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedCjkBreakAllStories"));
        var bytes = native ? File.ReadAllBytes(stem + ".doc") :
            DxpDocExport.Export(File.ReadAllBytes(stem + ".docx"));
        foreach (var binary in new[] { bytes,
            DxpDocExport.Export(DxpDocToDocx.Project(bytes).DocxBytes) })
        {
            using var index = new DocTextIndexWalker().Index(new MemoryStream(binary));
            var baseStyle = Assert.Single(index.StyleDefinitions,
                x => x.Name == "CJK break base");
            var derived = Assert.Single(index.StyleDefinitions,
                x => x.Name == "CJK break derived");
            Assert.Equal(baseStyle.Index, derived.BasedOnIndex);
            Assert.False(baseStyle.DirectParagraphFormatting?.Kinsoku);
            Assert.False(baseStyle.DirectParagraphFormatting?.WordWrap);
            Assert.Null(derived.DirectParagraphFormatting?.Kinsoku);
            Assert.Null(derived.DirectParagraphFormatting?.WordWrap);
            Assert.False(derived.ParagraphFormatting?.Kinsoku);
            Assert.False(derived.ParagraphFormatting?.WordWrap);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PairedCorpusExercisesInheritedCjkSpacingInAllStories(bool native)
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedCjkSpacingAllStories"));
        var bytes = native ? File.ReadAllBytes(stem + ".doc") :
            DxpDocExport.Export(File.ReadAllBytes(stem + ".docx"));
        foreach (var binary in new[] { bytes,
            DxpDocExport.Export(DxpDocToDocx.Project(bytes).DocxBytes) })
        {
            using var index = new DocTextIndexWalker().Index(new MemoryStream(binary));
            var baseStyle = Assert.Single(index.StyleDefinitions,
                x => x.Name == "CJK spacing base");
            var derived = Assert.Single(index.StyleDefinitions,
                x => x.Name == "CJK spacing derived");
            Assert.Equal(baseStyle.Index, derived.BasedOnIndex);
            Assert.False(baseStyle.DirectParagraphFormatting?.AutoSpaceDE);
            Assert.False(baseStyle.DirectParagraphFormatting?.AutoSpaceDN);
            Assert.Null(derived.DirectParagraphFormatting?.AutoSpaceDE);
            Assert.Null(derived.DirectParagraphFormatting?.AutoSpaceDN);
            Assert.False(derived.ParagraphFormatting?.AutoSpaceDE);
            Assert.False(derived.ParagraphFormatting?.AutoSpaceDN);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PairedCorpusExercisesInheritedOutlineInAllStories(bool native)
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedOutlineAllStories"));
        var bytes = native ? File.ReadAllBytes(stem + ".doc") :
            DxpDocExport.Export(File.ReadAllBytes(stem + ".docx"));
        foreach (var binary in new[] { bytes,
            DxpDocExport.Export(DxpDocToDocx.Project(bytes).DocxBytes) })
        {
            using var index = new DocTextIndexWalker().Index(new MemoryStream(binary));
            var baseStyle = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Outline pair base");
            var derived = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Outline pair derived");
            Assert.Equal(baseStyle.Index, derived.BasedOnIndex);
            Assert.Equal((byte)0, baseStyle.DirectParagraphFormatting?.OutlineLevel);
            Assert.Null(derived.DirectParagraphFormatting?.OutlineLevel);
            Assert.Equal((byte)0, derived.ParagraphFormatting?.OutlineLevel);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PairedCorpusExercisesInheritedBidiInAllStories(bool native)
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedBidiAllStories"));
        var bytes = native ? File.ReadAllBytes(stem + ".doc") :
            DxpDocExport.Export(File.ReadAllBytes(stem + ".docx"));
        foreach (var binary in new[] { bytes,
            DxpDocExport.Export(DxpDocToDocx.Project(bytes).DocxBytes) })
        {
            using var index = new DocTextIndexWalker().Index(new MemoryStream(binary));
            var baseStyle = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Visible Base");
            var derived = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Visible Derived");
            Assert.Equal(baseStyle.Index, derived.BasedOnIndex);
            Assert.True(baseStyle.DirectParagraphFormatting?.ParagraphRightToLeft);
            Assert.Null(derived.DirectParagraphFormatting?.ParagraphRightToLeft);
            Assert.True(derived.ParagraphFormatting?.ParagraphRightToLeft);
        }
    }

    [Fact]
    public void PairedCorpusExercisesInheritedPaginationStyle()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedPaginationStyleStories.docx"));
        var source = File.ReadAllBytes(path);
        var styles = ReadCustomStyleParagraphFormatting(source);
        Assert.Equal("False", styles["Pagination base.widowControl"]);
        Assert.Equal("False", styles["Pagination child.widowControl"]);
        Assert.Equal("True", styles["Pagination child.pageBreakBefore"]);
        var body = ReadEffectiveParagraphLayout(source);
        Assert.Equal("false", body["body.paragraph0.widowControl"]);
        Assert.Equal("true", body["body.paragraph2.pageBreakBefore"]);
        Assert.Equal("false", body["body.paragraph2.widowControl"]);
    }

    [Fact]
    public void PairedCorpusExercisesTextFlowAcrossTwoColumns()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordTwoColumnsStories.docx"));
        var layout = ReadSectionLayout(source);
        Assert.Equal("2", layout["section0.columns.count"]);
        Assert.Equal("360", layout["section0.columns.space"]);
        var stories = ReadStories(source);
        Assert.Contains("Column paragraph 01", stories["body"]);
        Assert.Contains("Column paragraph 65", stories["body"]);
        Assert.Equal("2", ReadEffectiveRunFormatting(source)[
            "body.paragraph0.char0.kern"]);
        Assert.Equal("Two-column header", stories["section0.header.default"]);
        Assert.Equal("Two-column footer", stories["section0.footer.default"]);
    }

    [Fact]
    public void PairedCorpusExercisesUnbasedStyleDefaultsAndKerningInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordKerningStyleAllStories.docx"));
        var formatting = ReadEffectiveRunFormatting(source);
        Assert.Equal("24", formatting["body.paragraph0.char0.size"]);
        Assert.Equal("4", formatting["body.paragraph0.char0.kern"]);
        Assert.Equal("0", formatting["body.paragraph1.char0.kern"]);
        Assert.Equal("4", formatting["section0.header.default.paragraph0.char0.kern"]);
        Assert.Equal("4", formatting["section0.footer.default.paragraph0.char0.kern"]);
    }

    [Fact]
    public void PairedCorpusKeepsUnbasedStyleDefaultsSeparateFromNormalOverride()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordUnbasedDefaultsAgainstNormalAllStories.docx"));
        using var document = WordprocessingDocument.Open(path, false);
        var normal = Assert.Single(document.MainDocumentPart!.StyleDefinitionsPart!
            .Styles!.Elements<Style>(), x => x.StyleId?.Value == "Normal");
        Assert.Equal("36", normal.StyleRunProperties!.FontSize!.Val!.Value);
        var formatting = ReadEffectiveRunFormatting(File.ReadAllBytes(path));
        Assert.Equal("24", formatting["body.paragraph0.char0.size"]);
        Assert.Equal("24", formatting["section0.header.default.paragraph0.char0.size"]);
        Assert.Equal("24", formatting["section0.footer.default.paragraph0.char0.size"]);
    }

    [Fact]
    public void PairedCorpusKeepsUnbasedParagraphSpacingSeparateFromNormalOverride()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordUnbasedParagraphDefaultsAgainstNormalAllStories.docx"));
        using var document = WordprocessingDocument.Open(path, false);
        var normal = Assert.Single(document.MainDocumentPart!.StyleDefinitionsPart!
            .Styles!.Elements<Style>(), x => x.StyleId?.Value == "Normal");
        Assert.Equal("400", normal.StyleParagraphProperties!
            .SpacingBetweenLines!.After!.Value);
        var layout = ReadEffectiveParagraphLayout(File.ReadAllBytes(path));
        Assert.Equal("160", layout["body.paragraph0.after"]);
        Assert.Equal("160", layout["section0.header.default.paragraph0.after"]);
        Assert.Equal("160", layout["section0.footer.default.paragraph0.after"]);
    }

    [Fact]
    public void PairedCorpusDerivesAllStoryStyleFromUnbasedDefaults()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordDerivedFromUnbasedDefaultsAllStories.docx"));
        var source = File.ReadAllBytes(path);
        Assert.Equal("Mirrored body", ReadStyleRelationships(source, false)
            ["Mirrored child.basedOn"]);
        var formatting = ReadEffectiveRunFormatting(source);
        foreach (var prefix in new[] { "body", "section0.header.default",
            "section0.footer.default" })
        {
            Assert.Equal("24", formatting[$"{prefix}.paragraph0.char0.size"]);
            Assert.Equal("4", formatting[$"{prefix}.paragraph0.char0.kern"]);
            Assert.Equal("true", formatting[$"{prefix}.paragraph0.char0.italic"]);
        }
        var layout = ReadEffectiveParagraphLayout(source);
        Assert.Equal("160", layout["body.paragraph0.after"]);
        Assert.Equal("160", layout["section0.header.default.paragraph0.after"]);
        Assert.Equal("160", layout["section0.footer.default.paragraph0.after"]);
    }

    [Fact]
    public void PairedCorpusExercisesUnequalColumnWidths()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var layout = ReadSectionLayout(File.ReadAllBytes(Path.Combine(directory,
            "WordUnequalColumnsStories.docx")));
        Assert.Equal("2", layout["section0.columns.count"]);
        Assert.Equal("False", layout["section0.columns.equalWidth"]);
        Assert.Equal("3400", layout["section0.columns.column0.width"]);
        Assert.Equal("360", layout["section0.columns.column0.space"]);
        Assert.Equal("5600", layout["section0.columns.column1.width"]);
    }

    [Fact]
    public void PairedCorpusExercisesSectionPageNumberRestart()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordSectionPageRestart.docx"));
        var layout = ReadSectionLayout(source);
        Assert.Equal("3", layout["sectionCount"]);
        Assert.Equal("6", layout["section1.pageNumber.start"]);
        Assert.Equal(new[] { "FIELD:PAGE" }, ReadEditableFieldSemantics(source));
    }

    [Fact]
    public void PairedCorpusExercisesFourSidedPageBorders()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var layout = ReadSectionLayout(File.ReadAllBytes(Path.Combine(directory,
            "WordPageBorderSectionSetup.docx")));
        Assert.Equal("single", layout["section0.pageBorders.top.val"]);
        Assert.Equal("AA3366", layout["section0.pageBorders.top.color"]);
        Assert.Equal("16", layout["section0.pageBorders.top.size"]);
        Assert.Equal("12", layout["section0.pageBorders.top.space"]);
        Assert.Equal("dashed", layout["section0.pageBorders.left.val"]);
        Assert.Equal("3366AA", layout["section0.pageBorders.left.color"]);
        Assert.Equal("single", layout["section0.pageBorders.bottom.val"]);
        Assert.Equal("dashed", layout["section0.pageBorders.right.val"]);
        Assert.Equal("firstPage", layout["section0.pageBorders.display"]);
        Assert.Equal("page", layout["section0.pageBorders.offsetFrom"]);
        Assert.Equal("back", layout["section0.pageBorders.zOrder"]);
    }

    [Fact]
    public void PairedCorpusExercisesPageBordersWithDistinctSectionStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordMixedGeometryPageBorderStories.docx"));
        var layout = ReadSectionLayout(source);
        Assert.Equal("2", layout["sectionCount"]);
        Assert.Equal("single", layout["section0.pageBorders.top.val"]);
        Assert.Equal("dashed", layout["section0.pageBorders.left.val"]);
        Assert.Equal("double", layout["section0.pageBorders.bottom.val"]);
        Assert.Equal("dotted", layout["section0.pageBorders.right.val"]);
        Assert.Equal("allPages", layout["section0.pageBorders.display"]);
        Assert.Equal("page", layout["section0.pageBorders.offsetFrom"]);
        Assert.Equal("front", layout["section0.pageBorders.zOrder"]);
        Assert.DoesNotContain(layout.Keys, key => key.StartsWith(
            "section1.pageBorders.", StringComparison.Ordinal));
        Assert.True(ReadEffectiveHeaderFooterStories(source).Count >= 6);
    }

    [Fact]
    public void PairedCorpusExercisesInlinePictureFlipsInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordFlippedImageStories.docx"));
        var geometry = ReadDrawingGeometry(source);
        Assert.Equal("True", geometry["body.drawing0.flipH"]);
        Assert.Equal("False", geometry["body.drawing0.flipV"]);
        Assert.Equal("False", geometry["section0.header.default.drawing0.flipH"]);
        Assert.Equal("True", geometry["section0.header.default.drawing0.flipV"]);
        Assert.Equal("True", geometry["section0.footer.default.drawing0.flipH"]);
        Assert.Equal("True", geometry["section0.footer.default.drawing0.flipV"]);
    }

    [Fact]
    public void PairedCorpusExercisesRightAngleInlinePictureRotationInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordRightAngleImageStories.docx"));
        var geometry = ReadDrawingGeometry(source);
        Assert.Equal("5400000", geometry["body.drawing0.rotation"]);
        Assert.Equal("16200000", geometry["section0.header.default.drawing0.rotation"]);
        Assert.Equal("10800000", geometry["section0.footer.default.drawing0.rotation"]);
    }

    [Fact]
    public void HeaderWrapKnownGapRetainsStoryAndFloatingPictureThroughBothDocRoutes()
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures",
            "DocKnownGaps", "WordFloatingHeaderWrapAllStories"));
        var source = File.ReadAllBytes(stem + ".docx");
        Assert.Contains("longer line of nearby words",
            ReadStories(source)["section0.header.default"]);
        foreach (var binary in new[] { File.ReadAllBytes(stem + ".doc"),
            DxpDocExport.Export(source) })
        {
            foreach (var doc in new[] { binary,
                DxpDocExport.Export(DxpDocToDocx.Project(binary).DocxBytes) })
            {
                var projection = DxpDocToDocx.Project(doc);
                Assert.Empty(projection.Coverage.OmittedCharacters);
                Assert.Equal(ReadStories(source), ReadStories(projection.DocxBytes));
                AssertDrawingGeometry(source, projection.DocxBytes);
                using (var stream = new MemoryStream(projection.DocxBytes))
                using (var package = WordprocessingDocument.Open(stream, false))
                    Assert.Contains(package.MainDocumentPart!.DocumentSettingsPart!
                        .Settings!.Descendants<CompatibilitySetting>(), setting =>
                        setting.Name?.Value == CompatSettingNameValues.CompatibilityMode &&
                        setting.Val?.Value == "15");
                Assert.Empty(Validate(projection.DocxBytes));
            }
        }
    }

    [Theory]
    [InlineData("WordTransformedFloatingImageStories")]
    [InlineData("WordRotatedFloatingImageStories")]
    public void FloatingPictureTransformsSurviveGeneratedDocAndThirdHop(string name)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures",
            "DocKnownGaps", name + ".docx"));
        var source = File.ReadAllBytes(path);
        var projected = DxpDocToDocx.Project(DxpDocExport.Export(source)).DocxBytes;
        AssertDrawingGeometry(source, projected);
        var repeated = DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes;
        AssertDrawingGeometry(source, repeated);
        using var stream = new MemoryStream(repeated);
        using var document = WordprocessingDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator().Validate(document));
    }

    [Theory]
    [InlineData("WordTransformedFloatingImageStories")]
    [InlineData("WordRotatedFloatingImageStories")]
    public void WordNativeFloatingPicturesSurviveAnotherDocHop(string name)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures",
            "DocKnownGaps", name + ".doc"));
        var native = File.ReadAllBytes(path);
        var first = DxpDocToDocx.Project(native).DocxBytes;
        var rewritten = DxpDocToDocx.Project(DxpDocExport.Export(first)).DocxBytes;
        Assert.Equal(ReadStories(first), ReadStories(rewritten));
        AssertDrawingGeometry(first, rewritten);
        Assert.Empty(Validate(rewritten));
    }

    [Fact]
    public void WordNativeHeaderTransformIsFlattenedToASeparateRaster()
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures",
            "Doc", "WordTransformedFloatingImageStories"));
        var source = File.ReadAllBytes(stem + ".docx");
        var native = DxpDocToDocx.Project(File.ReadAllBytes(stem + ".doc")).DocxBytes;
        AssertNativeFlattenedHeaderGeometry(source, native);
    }

    private static void AssertNativeFlattenedHeaderGeometry(byte[] source, byte[] native)
    {
        var generated = DxpDocToDocx.Project(DxpDocExport.Export(source)).DocxBytes;
        var original = ReadDrawingGeometry(source);
        var flattened = ReadDrawingGeometry(native);
        AssertDrawingGeometry(source, generated);
        const string header = "section0.header.default.drawing0.";
        Assert.Equal("635000", original[header + "cx"]);
        Assert.Equal("444500", original[header + "cy"]);
        Assert.Equal("5400000", original[header + "rotation"]);
        Assert.Equal("True", original[header + "flipV"]);
        Assert.Equal("444500", flattened[header + "cx"]);
        Assert.Equal("635000", flattened[header + "cy"]);
        Assert.Equal("0", flattened[header + "rotation"]);
        Assert.Equal("False", flattened[header + "flipV"]);
        Assert.Equal(95250, long.Parse(flattened[header + "horizontal.offset"]) -
            long.Parse(original[header + "horizontal.offset"]));
        Assert.Equal(-95250, long.Parse(flattened[header + "vertical.offset"]) -
            long.Parse(original[header + "vertical.offset"]));
        foreach (var key in new[] { "cx", "cy", "rotation", "flipV",
            "horizontal.offset", "vertical.offset" })
        {
            original.Remove(header + key);
            flattened.Remove(header + key);
        }
        Assert.Equal(original.Keys.OrderBy(key => key),
            flattened.Keys.OrderBy(key => key));
        foreach (var (key, expected) in original)
        {
            var actual = flattened[key];
            Assert.True(expected == actual ||
                ((key.EndsWith(".offset", StringComparison.Ordinal) ||
                  key.EndsWith(".cx", StringComparison.Ordinal) ||
                  key.EndsWith(".cy", StringComparison.Ordinal)) &&
                 long.TryParse(expected, out var expectedEmu) &&
                 long.TryParse(actual, out var actualEmu) &&
                 Math.Abs(expectedEmu - actualEmu) <= 635),
                $"{key}: expected {expected}, observed {actual}");
        }
        Assert.Single(ReadDistinctPngPayloadHashes(source));
        Assert.Equal(2, ReadDistinctPngPayloadHashes(native).Length);
        Assert.Empty(Validate(native));
    }
    [Fact]
    public void CombinedInheritedStyleAppliesInAllSixSlotsAcrossSections()
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures",
            "Doc", "WordSixSlotCombinedStyles"));
        var source = File.ReadAllBytes(stem + ".docx");
        var expected = ReadStories(source);
        for (var section = 0; section < 2; section++)
            foreach (var kind in new[] { "header", "footer" })
                foreach (var slot in new[] { "first", "even", "default" })
                    Assert.Equal("Six Story Derived", expected[
                        $"section{section}.{kind}.{slot}.customStyles"]);
        const string derived = "Six Story Derived.";
        var style = ReadCustomStyleRunFormatting(source);
        Assert.Equal("20", style[derived + "characterSpacing"]);
        Assert.Equal("C93628", style[derived + "underlineColor"]);
        Assert.Equal("2B9D6F", style[derived + "border.color"]);
        Assert.Equal("B8E0F4", style[derived + "shading.fill"]);
        foreach (var binary in new[] { File.ReadAllBytes(stem + ".doc"),
            DxpDocExport.Export(source) })
        {
            var first = DxpDocToDocx.Project(binary).DocxBytes;
            var third = DxpDocToDocx.Project(DxpDocExport.Export(first)).DocxBytes;
            foreach (var docx in new[] { first, third })
            {
                Assert.Equal(expected, ReadStories(docx));
                var observed = ReadCustomStyleRunFormatting(docx);
                foreach (var property in new[] { "characterSpacing", "underlineColor",
                    "border.color", "shading.fill" })
                    Assert.Equal(style[derived + property],
                        observed[derived + property]);
                Assert.Empty(Validate(docx));
            }
        }
    }

    [Fact]
    public void CombinedStyleEffectsRemainEditableAcrossBothDocRoutes()
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures",
            "Doc", "WordCombinedStyleVisualAllStories"));
        var source = File.ReadAllBytes(stem + ".docx");
        var runStyle = ReadCustomStyleRunFormatting(source);
        const string style = "Mirrored body.";
        Assert.Equal("40", runStyle[style + "characterSpacing"]);
        Assert.Equal("C93628", runStyle[style + "underlineColor"]);
        Assert.Equal("2B9D6F", runStyle[style + "border.color"]);
        Assert.Equal("B8E0F4", runStyle[style + "shading.fill"]);
        Assert.Equal("True", ReadCustomStyleParagraphFormatting(source)[
            style + "mirrorIndents"]);
        var stories = ReadStories(source);
        Assert.Contains(stories.Keys, key => key.Contains("header", StringComparison.Ordinal));
        Assert.Contains(stories.Keys, key => key.Contains("footer", StringComparison.Ordinal));
        foreach (var binary in new[] { File.ReadAllBytes(stem + ".doc"),
            DxpDocExport.Export(source) })
        {
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            var repeated = DxpDocToDocx.Project(DxpDocExport.Export(projected)).DocxBytes;
            foreach (var docx in new[] { projected, repeated })
            {
                var observed = ReadCustomStyleRunFormatting(docx);
                foreach (var property in new[] { "characterSpacing", "underlineColor",
                    "border.color", "shading.fill" })
                    Assert.Equal(runStyle[style + property], observed[style + property]);
                Assert.Equal("True", ReadCustomStyleParagraphFormatting(docx)[
                    style + "mirrorIndents"]);
                Assert.Equal(stories, ReadStories(docx));
                Assert.Empty(Validate(docx));
            }
        }
    }

    [Fact]
    public void WordNativeRotationReparameterizesAllThreeFloatingPictures()
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures",
            "DocKnownGaps", "WordRotatedFloatingImageStories"));
        var source = File.ReadAllBytes(stem + ".docx");
        var native = DxpDocToDocx.Project(File.ReadAllBytes(stem + ".doc")).DocxBytes;
        var original = ReadDrawingGeometry(source);
        var converted = ReadDrawingGeometry(native);
        AssertDrawingGeometry(source,
            DxpDocToDocx.Project(DxpDocExport.Export(source)).DocxBytes);
        const string body = "body.drawing0.";
        const string header = "section0.header.default.drawing0.";
        const string footer = "section0.footer.default.drawing0.";
        Assert.Equal(("1524000", "1143000", "2700000"),
            (original[body + "cx"], original[body + "cy"],
                original[body + "rotation"]));
        Assert.Equal(("1143000", "1524000", "2700000"),
            (converted[body + "cx"], converted[body + "cy"],
                converted[body + "rotation"]));
        Assert.Equal(("635000", "444500", "5400000"),
            (original[header + "cx"], original[header + "cy"],
                original[header + "rotation"]));
        Assert.Equal(("444500", "635000", "0"),
            (converted[header + "cx"], converted[header + "cy"],
                converted[header + "rotation"]));
        Assert.Equal(("571500", "381000", "10800000"),
            (original[footer + "cx"], original[footer + "cy"],
                original[footer + "rotation"]));
        Assert.Equal(("571500", "381000", "0"),
            (converted[footer + "cx"], converted[footer + "cy"],
                converted[footer + "rotation"]));
        Assert.Single(ReadDistinctPngPayloadHashes(source));
        Assert.Equal(3, ReadDistinctPngPayloadHashes(native).Length);
        Assert.Empty(Validate(native));
    }

    [Fact]
    public void WordNativeArbitraryRotationShapeFieldRestoresInlinePicture()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures",
            "Doc", "WordArbitraryRotationStories.doc"));
        var projection = DxpDocToDocx.Project(File.ReadAllBytes(path));
        Assert.Empty(projection.Coverage.OmittedCharacters);
        using var stream = new MemoryStream(projection.DocxBytes);
        using var document = WordprocessingDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator().Validate(document));
        Assert.Contains(document.MainDocumentPart!.HeaderParts,
            part => part.Header!.Descendants<Drawing>().Any(drawing =>
                drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline>()
                    .Any()));
        Assert.DoesNotContain(document.MainDocumentPart.HeaderParts,
            part => part.Header!.Descendants<FieldCode>().Any(code =>
                code.Text.Contains("SHAPE", StringComparison.OrdinalIgnoreCase)));
    }

    [Fact]
    public void PairedCorpusExercisesDistinctTextControls()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory, "WordTextControls.docx"));
        var stories = ReadStories(source);
        var body = stories["body"];
        Assert.Contains("Body\ttab\u2028line\fpage\u2011join\u00adsoft", body);
        Assert.Contains("Header\u2028line", stories["section0.header.default"]);
        Assert.Contains("Footer\u2011join\u00adsoft", stories["section0.footer.default"]);
    }

    [Fact]
    public void PairedCorpusKeepsNonbreakingSpacesInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        const string name = "WordTextControlsNbspAllStories";
        var source = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var native = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));
        var generated = DxpDocExport.Export(source);
        foreach (var bytes in new[]
        {
            source,
            DxpDocToDocx.Project(native).DocxBytes,
            DxpDocToDocx.Project(generated).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(generated).DocxBytes)).DocxBytes
        })
        {
            var stories = ReadStories(bytes);
            Assert.Contains("Body\u00A0space\ttab\u00A0styled\u2028line\fpage\u2011join\u00ADsoft",
                stories["body"]);
            Assert.Contains("Header\u00A0space\u2028line\u00A0styled",
                stories["section0.header.default"]);
            Assert.Contains("Footer\u00A0space\u2011join\u00A0styled\u00ADsoft",
                stories["section0.footer.default"]);
            using var stream = new MemoryStream(bytes);
            using var package = WordprocessingDocument.Open(stream, false);
            var main = package.MainDocumentPart!;
            static bool StyledNonbreakingSpace(Run run) =>
                run.InnerText.Contains('\u00A0') &&
                run.RunProperties?.RunStyle?.Val?.Value != null;
            Assert.Contains(main.Document!.Body!.Descendants<Run>(),
                StyledNonbreakingSpace);
            Assert.Contains(main.HeaderParts.SelectMany(x => x.Header!
                .Descendants<Run>()), StyledNonbreakingSpace);
            Assert.Contains(main.FooterParts.SelectMany(x => x.Footer!
                .Descendants<Run>()), StyledNonbreakingSpace);
        }
    }

    [Fact]
    public void WordHyperlinkBinaryPayloadIsNotAnInlinePicture()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        using var input = File.OpenRead(Path.Combine(directory, "WordFieldsAndLink.doc"));
        using var index = new DocTextIndexWalker().Index(input);
        var atoms = index.Parts.Where(x => x.Name is "Main" or "Headers")
            .SelectMany(x => DocStoryTextReader.Read(index, x.Name).Paragraphs)
            .SelectMany(x => x.Atoms)
            .ToArray();
        Assert.Contains(atoms, x => x.Kind == DocStoryAtomKind.FieldData);
        Assert.DoesNotContain(atoms, x => x.Kind == DocStoryAtomKind.InlinePicture);
        var sourceDocx = File.ReadAllBytes(Path.Combine(directory,
            "WordFieldsAndLink.docx"));
        Assert.Equal(new[] { "FIELD:NUMPAGES", "FIELD:PAGE",
            "LINK:https://example.com" },
            ReadEditableFieldSemantics(sourceDocx));
    }

    [Fact]
    public void PairedCorpusExercisesDateAndTimeFields()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordDateTimeFields.docx"));
        Assert.Equal(new[] { "FIELD:DATE", "FIELD:NUMPAGES", "FIELD:PAGE",
            "FIELD:TIME", "LINK:https://example.com" },
            ReadEditableFieldSemantics(source));
    }

    [Fact]
    public void PairedCorpusExercisesHyperlinksInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordLinksAllStories.docx"));
        Assert.Equal(new[] { "FIELD:DATE", "FIELD:NUMPAGES", "FIELD:PAGE",
            "FIELD:TIME", "LINK:https://example.com",
            "LINK:https://example.org/footer", "LINK:https://example.org/header" },
            ReadEditableFieldSemantics(source));
    }

    [Fact]
    public void PairedCorpusExercisesTablesInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var grid = ReadTableGrid(File.ReadAllBytes(Path.Combine(directory,
            "WordTablesAllStories.docx")));
        Assert.Equal("1", grid["body.tables"]);
        Assert.Equal("2", grid["body.table0.rows"]);
        Assert.Equal("Body D", grid["body.table0.row1.cell1.text"]);
        Assert.Equal("1", grid["section0.header.default.tables"]);
        Assert.Equal("Header B", grid["section0.header.default.table0.row0.cell1.text"]);
        Assert.Equal("1", grid["section0.footer.default.tables"]);
        Assert.Equal("Footer B", grid["section0.footer.default.table0.row0.cell1.text"]);
    }

    [Fact]
    public void PairedCorpusExercisesNumberedTableCellsInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordTableListsAllStories.docx"));
        var lists = ReadListSemantics(source);
        Assert.Equal(3, lists.Count);
        Assert.Contains(lists.Keys, x => x.StartsWith("body.",
            StringComparison.Ordinal));
        Assert.Contains(lists.Keys, x => x.StartsWith("section0.header.default.",
            StringComparison.Ordinal));
        Assert.Contains(lists.Keys, x => x.StartsWith("section0.footer.default.",
            StringComparison.Ordinal));
        Assert.All(lists.Values, x => Assert.StartsWith("decimal|%1.|1|", x));
        var grid = ReadTableGrid(source);
        Assert.Equal("1", grid["body.tables"]);
        Assert.Equal("1", grid["section0.header.default.tables"]);
        Assert.Equal("1", grid["section0.footer.default.tables"]);
    }

    [Fact]
    public void PairedCorpusExercisesCellDecorationsInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var grid = ReadTableGrid(File.ReadAllBytes(Path.Combine(directory,
            "WordCellDecorationsAllStories.docx")));
        Assert.Equal("FFFF00", grid["body.table0.row0.cell0.fill"]);
        Assert.Equal("0000FF", grid["body.table0.row0.cell0.bottom.color"]);
        Assert.Equal("99CCFF", grid[
            "section0.header.default.table0.row0.cell0.fill"]);
        Assert.Equal("FF0000", grid[
            "section0.header.default.table0.row0.cell0.bottom.color"]);
        Assert.Equal("CCFFCC", grid[
            "section0.footer.default.table0.row0.cell0.fill"]);
        Assert.Equal("008000", grid[
            "section0.footer.default.table0.row0.cell0.bottom.color"]);
    }

    [Fact]
    public void PairedCorpusExercisesCellVerticalAlignmentInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var grid = ReadTableGrid(File.ReadAllBytes(Path.Combine(directory,
            "WordCellAlignmentAllStories.docx")));
        Assert.Equal("center", grid["body.table0.row0.cell0.verticalAlign"]);
        Assert.Equal("bottom", grid[
            "section0.header.default.table0.row0.cell0.verticalAlign"]);
        Assert.Equal("center", grid[
            "section0.footer.default.table0.row0.cell0.verticalAlign"]);
    }

    [Fact]
    public void PairedCorpusExercisesRowLayoutInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var grid = ReadTableGrid(File.ReadAllBytes(Path.Combine(directory,
            "WordRowLayoutAllStories.docx")));
        Assert.Equal("exact:600", grid["body.table0.row0.height"]);
        Assert.Equal("atLeast:520", grid[
            "section0.header.default.table0.row0.height"]);
        Assert.Equal("exact:480", grid[
            "section0.footer.default.table0.row0.height"]);
        Assert.Equal("1", grid["body.table0.row0.repeatHeader"]);
        Assert.Equal("1", grid["body.table0.row0.cantSplit"]);
        Assert.Equal("1", grid["section0.header.default.table0.row0.cantSplit"]);
        Assert.Equal("1", grid["section0.footer.default.table0.row0.cantSplit"]);
    }

    [Fact]
    public void NestedTableFixtureRetainsDepthAndGridInBothRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures",
            "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordNestedTableBody.docx"));
        var grid = ReadTableGrid(source);
        Assert.Equal("2", grid["body.tables"]);
        Assert.Contains("Nested body", ReadStories(source)["body"]);
        using (var input = File.OpenRead(Path.Combine(directory,
            "WordNestedTableBody.doc")))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            Assert.Contains(index.ParagraphStyles,
                x => x.Formatting?.TableDepth == 2);
            Assert.Contains(index.ParagraphStyles,
                x => x.Formatting?.InnerTableCell == true);
            Assert.Contains(index.ParagraphStyles,
                x => x.Formatting?.InnerTableRow == true);
            var story = DocStoryTextReader.Read(index, "Main");
            Assert.Contains(story.Paragraphs, x => x.End == DocParagraphEnd.CellMark &&
                x.CpEnd > x.CpStart && index.ParagraphStyles.Any(style =>
                    style.CpStart <= x.CpEnd - 1 && x.CpEnd - 1 < style.CpEnd &&
                    style.Formatting?.TableDepth == 2));
            Assert.Contains(story.Paragraphs, x => x.End == DocParagraphEnd.RowMark &&
                x.CpEnd > x.CpStart && index.ParagraphStyles.Any(style =>
                    style.CpStart <= x.CpEnd - 1 && x.CpEnd - 1 < style.CpEnd &&
                    style.Formatting?.TableDepth == 2));
        }
        var projected = DxpDocToDocx.Project(File.ReadAllBytes(Path.Combine(
            directory, "WordNestedTableBody.doc")));
        Assert.Equal(ReadStories(source), ReadStories(projected.DocxBytes));
        AssertTableGrid(source, projected.DocxBytes);
        var written = DxpDocExport.Export(source);
        var roundTrip = DxpDocToDocx.Project(written);
        Assert.Equal(ReadStories(source), ReadStories(roundTrip.DocxBytes));
        AssertTableGrid(source, roundTrip.DocxBytes);
    }

    [Fact]
    public void PairedCorpusExercisesNestedTablesInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordNestedTablesAllStories.docx"));
        var grid = ReadTableGrid(source);
        Assert.Equal("2", grid["body.tables"]);
        Assert.Equal("2", grid["section0.header.default.tables"]);
        Assert.Equal("2", grid["section0.footer.default.tables"]);
        Assert.Equal("0", grid["body.table1.parentTable"]);
        Assert.Equal("0", grid["section0.header.default.table1.parentTable"]);
        Assert.Equal("0", grid["section0.footer.default.table1.parentTable"]);
        Assert.Contains("Nested body", grid["body.table1.row0.cell0.text"]);
        Assert.Contains("Nested header", grid[
            "section0.header.default.table1.row0.cell0.text"]);
        Assert.Contains("Nested footer", grid[
            "section0.footer.default.table1.row0.cell0.text"]);
    }

    [Fact]
    public void PairedCorpusExercisesThreeLevelTablesInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordDeepNestedTablesAllStories.docx"));
        var grid = ReadTableGrid(source);
        foreach (var story in new[] { "body", "section0.header.default",
            "section0.footer.default" })
        {
            Assert.Equal("3", grid[$"{story}.tables"]);
            Assert.Equal("0", grid[$"{story}.table1.parentTable"]);
            Assert.Equal("1", grid[$"{story}.table2.parentTable"]);
            Assert.Contains("Deep", grid[$"{story}.table2.row0.cell0.text"]);
        }
        using var input = File.OpenRead(Path.Combine(directory,
            "WordDeepNestedTablesAllStories.doc"));
        using var index = new DocTextIndexWalker().Index(input);
        Assert.Contains(index.ParagraphStyles, x => x.Formatting?.TableDepth == 3);
    }

    [Fact]
    public void PairedCorpusExercisesNestedTwoByTwoGridsInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var grid = ReadTableGrid(File.ReadAllBytes(Path.Combine(directory,
            "WordNestedGridsAllStories.docx")));
        foreach (var story in new[] { "body", "section0.header.default",
            "section0.footer.default" })
        {
            Assert.Equal("2", grid[$"{story}.tables"]);
            Assert.Equal("0", grid[$"{story}.table1.parentTable"]);
            Assert.Equal("2", grid[$"{story}.table1.rows"]);
            Assert.Equal("2", grid[$"{story}.table1.row0.cells"]);
            Assert.Equal("2", grid[$"{story}.table1.row1.cells"]);
        }
    }

    [Fact]
    public void PairedCorpusExercisesMergedCellInsideNestedTable()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var grid = ReadTableGrid(File.ReadAllBytes(Path.Combine(directory,
            "WordNestedMergedCell.docx")));
        Assert.Equal("0", grid["body.table1.parentTable"]);
        Assert.Equal("2", grid["body.table1.rows"]);
        Assert.Equal("1", grid["body.table1.row0.cells"]);
        Assert.Equal("2", grid["body.table1.row0.cell0.span"]);
        Assert.Equal("2", grid["body.table1.row1.cells"]);
    }

    [Fact]
    public void PairedCorpusExercisesNumberedParagraphsInsideNestedTables()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordNestedTableListsAllStories.docx"));
        var grid = ReadTableGrid(source);
        foreach (var story in new[] { "body", "section0.header.default",
            "section0.footer.default" })
            Assert.Equal("0", grid[$"{story}.table1.parentTable"]);
        var lists = ReadListSemantics(source);
        Assert.Equal(3, lists.Count);
        Assert.Contains(lists.Keys, x => x.StartsWith("body.", StringComparison.Ordinal));
        Assert.Contains(lists.Keys, x => x.StartsWith(
            "section0.header.default.", StringComparison.Ordinal));
        Assert.Contains(lists.Keys, x => x.StartsWith(
            "section0.footer.default.", StringComparison.Ordinal));
    }

    [Fact]
    public void PairedCorpusExercisesFormattedEmptyParagraphsInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordEmptyParagraphsAllStories.docx"));
        var stories = ReadStories(source);
        Assert.Contains("First body paragraph\n\nLast body paragraph", stories["body"]);
        Assert.Equal("Header first\n\nHeader last",
            stories["section0.header.default"]);
        Assert.Equal("Footer first\n\nFooter last",
            stories["section0.footer.default"]);
        var marks = ReadInteriorEmptyParagraphMarks(source);
        Assert.Equal("bold", marks["body.paragraph1"]);
        Assert.Equal("italic", marks["section0.header.default.paragraph1"]);
        Assert.Equal("bold", marks["section0.footer.default.paragraph1"]);
    }

    [Fact]
    public void PairedCorpusExercisesParagraphBordersAndShadingInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var decorations = ReadParagraphDecorations(File.ReadAllBytes(Path.Combine(
            directory, "WordParagraphBordersShadingStories.docx")));
        Assert.Equal("FFFF00", decorations["body.paragraph0.fill"]);
        Assert.Equal("0000FF", decorations["body.paragraph0.top.color"]);
        Assert.Equal("99CCFF", decorations[
            "section0.header.default.paragraph0.fill"]);
        Assert.Equal("FF0000", decorations[
            "section0.header.default.paragraph0.top.color"]);
        Assert.Equal("CCFFCC", decorations[
            "section0.footer.default.paragraph0.fill"]);
        Assert.Equal("008000", decorations[
            "section0.footer.default.paragraph0.top.color"]);
    }

    [Fact]
    public void PairedCorpusExercisesStyledDecorationsWithDirectOverrides()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordStyledParagraphDecorations.docx"));
        var decorations = ReadParagraphDecorations(source);
        Assert.Equal("FFFF00", decorations["body.paragraph0.fill"]);
        Assert.Equal("0000FF", decorations["body.paragraph0.top.color"]);
        Assert.Equal("99CCFF", decorations[
            "section0.header.default.paragraph0.fill"]);
        Assert.Equal("FF0000", decorations[
            "section0.header.default.paragraph0.top.color"]);
        Assert.Equal("CCFFCC", decorations[
            "section0.footer.default.paragraph0.fill"]);
        Assert.Equal("008000", decorations[
            "section0.footer.default.paragraph0.top.color"]);
        var stories = ReadStories(source);
        Assert.Contains("Decorated Note", stories["body.customStyles"]);
        Assert.Contains("Decorated Note", stories[
            "section0.header.default.customStyles"]);
        Assert.Contains("Decorated Note", stories[
            "section0.footer.default.customStyles"]);
    }

    [Fact]
    public void PairedCorpusExercisesCharacterEffectsInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var formatting = ReadEffectiveRunFormatting(File.ReadAllBytes(Path.Combine(
            directory, "WordCharacterEffectsStories.docx")));
        Assert.Equal("Arial", formatting["body.paragraph0.char0.font"]);
        Assert.Equal("yellow", formatting["body.paragraph0.char0.highlight"]);
        Assert.Equal("true", formatting["body.paragraph1.char0.hidden"]);
        Assert.Equal("superscript", formatting[
            "section0.header.default.paragraph0.char0.baseline"]);
        Assert.Equal("subscript", formatting[
            "section0.footer.default.paragraph0.char0.baseline"]);
    }

    [Fact]
    public void AdjacentHighlightNoneKeepsRunFormattingLocalInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, "WordAdjacentHighlightNoneAllStories");
        var source = File.ReadAllBytes(stem + ".docx");
        var generated = DxpDocExport.Export(source);
        var native = File.ReadAllBytes(stem + ".doc");
        var generatedProjection = DxpDocToDocx.Project(generated).DocxBytes;
        var nativeProjection = DxpDocToDocx.Project(native).DocxBytes;
        var routes = new[] { source, generatedProjection, nativeProjection,
            DxpDocToDocx.Project(DxpDocExport.Export(generatedProjection)).DocxBytes };
        for (var route = 0; route < routes.Length; route++)
        {
            var docx = routes[route];
            var formatting = ReadEffectiveRunFormatting(docx);
            foreach (var story in new[] { "body", "section0.header.default",
                "section0.footer.default" })
            {
                Assert.Equal("yellow", formatting[$"{story}.paragraph0.char0.highlight"]);
                Assert.DoesNotContain(
                    $"{story}.paragraph0.char7.highlight", formatting);
            }
            var errors = Validate(docx);
            Assert.True(errors.Count == 0, string.Join("; ",
                errors.Select(x => x.Description)));
        }
        foreach (var doc in new[] { generated, native })
        {
            using var stream = new MemoryStream(doc);
            using var index = new DocTextIndexWalker().Index(stream);
            Assert.Contains(index.StyleDefinitions, x =>
                x.CharacterFormatting?.ColorRef != null);
        }
    }

    [Fact]
    public void CharacterStyleAndDirectHighlightsSurviveAllStoriesAndBothDocRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory,
            "WordCharacterStyleDirectHighlightsAllStories");
        var source = File.ReadAllBytes(stem + ".docx");
        var generated = DxpDocExport.Export(source);
        var native = File.ReadAllBytes(stem + ".doc");
        var routes = new[]
        {
            source,
            DxpDocToDocx.Project(generated).DocxBytes,
            DxpDocToDocx.Project(native).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(generated).DocxBytes)).DocxBytes
        };
        for (var route = 0; route < routes.Length; route++)
        {
            var docx = routes[route];
            var formatting = ReadEffectiveRunFormatting(docx);
            foreach (var story in new[]
            {
                "body", "section0.header.default", "section0.footer.default"
            })
            {
                Assert.Equal("cyan", formatting[$"{story}.paragraph0.char0.highlight"]);
                Assert.Equal("yellow", formatting[
                    $"{story}.paragraph1.char0.highlight"]);
            }
            Assert.Empty(Validate(docx));
        }
    }

    [Fact]
    public void CharacterStyleShadingAndDirectHighlightSurviveAllStoriesAndBothDocRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory,
            "WordCharacterStyleShadeHighlightAllStories");
        var source = File.ReadAllBytes(stem + ".docx");
        var generated = DxpDocExport.Export(source);
        var native = File.ReadAllBytes(stem + ".doc");
        var routes = new[]
        {
            source,
            DxpDocToDocx.Project(generated).DocxBytes,
            DxpDocToDocx.Project(native).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(generated).DocxBytes)).DocxBytes
        };
        for (var route = 0; route < routes.Length; route++)
        {
            var docx = routes[route];
            var formatting = ReadEffectiveRunFormatting(docx);
            foreach (var story in new[]
            {
                "body", "section0.header.default", "section0.footer.default"
            })
            {
                Assert.Equal("cyan", formatting[$"{story}.paragraph0.char0.highlight"]);
                Assert.Equal("FFF2CC", formatting[
                    $"{story}.paragraph0.char0.shading.fill"]);
                Assert.Equal("FFF2CC", formatting[
                    $"{story}.paragraph1.char0.shading.fill"]);
                Assert.DoesNotContain($"{story}.paragraph1.char0.highlight",
                    formatting.Keys);
            }
            Assert.Empty(Validate(docx));
        }
    }

    [Fact]
    public void DirectRunDirectionOverridesSurviveStyledStoriesAndBothDocRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, "WordRunDirectionOverridesAllStories");
        var source = File.ReadAllBytes(stem + ".docx");
        var generated = DxpDocExport.Export(source);
        var native = File.ReadAllBytes(stem + ".doc");
        var routes = new[]
        {
            source,
            DxpDocToDocx.Project(generated).DocxBytes,
            DxpDocToDocx.Project(native).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(generated).DocxBytes)).DocxBytes
        };
        for (var route = 0; route < routes.Length; route++)
        {
            var docx = routes[route];
            var formatting = ReadEffectiveRunFormatting(docx);
            Assert.Equal("true", formatting["body.paragraph0.char0.rtl"]);
            var hasHeaderDirection = formatting.TryGetValue(
                "section0.header.default.paragraph0.char0.rtl", out var headerRtl);
            if (route != 2)
                Assert.True(hasHeaderDirection,
                    $"Route {route} lost the explicit header direction reset.");
            Assert.Equal("false", headerRtl ?? "false");
            Assert.Equal("true", formatting[
                "section0.footer.default.paragraph0.char0.rtl"]);
            Assert.Empty(Validate(docx));
        }
    }

    [Fact]
    public void HighlightPaletteSurvivesAllStoriesAndBothDocRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, "WordHighlightPaletteAllStories");
        var source = File.ReadAllBytes(stem + ".docx");
        var generated = DxpDocExport.Export(source);
        var expected = new[]
        {
            "black", "blue", "cyan", "green", "magenta", "red", "yellow",
            "white", "darkBlue", "darkCyan", "darkGreen", "darkMagenta",
            "darkRed", "darkYellow", "darkGray", "lightGray"
        };
        foreach (var docx in new[]
        {
            source,
            DxpDocToDocx.Project(generated).DocxBytes,
            DxpDocToDocx.Project(File.ReadAllBytes(stem + ".doc")).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(generated).DocxBytes)).DocxBytes
        })
        {
            using var stream = new MemoryStream(docx);
            using var document = WordprocessingDocument.Open(stream, false);
            var main = document.MainDocumentPart!;
            var stories = new DocumentFormat.OpenXml.OpenXmlElement[]
            {
                main.Document.Body!,
                main.HeaderParts.Single(x => x.Header!.InnerText.Contains("Header first"))
                    .Header!,
                main.FooterParts.Single(x => x.Footer!.InnerText.Contains("Footer first"))
                    .Footer!
            };
            foreach (var story in stories)
            {
                var palette = story.Descendants<Run>()
                    .Where(run => run.InnerText.Trim().StartsWith("P", StringComparison.Ordinal))
                    .OrderBy(run => run.InnerText.Trim(), StringComparer.Ordinal)
                    .Select(run => run.RunProperties?.Highlight?.Val?.InnerText)
                    .ToArray();
                Assert.Equal(expected, palette);
            }
            Assert.Empty(Validate(docx));
        }
    }

    [Fact]
    public void RunShadeClearKeepsHighlightOverCharacterStyleInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, "WordShadeResetWithHighlightAllStories");
        var source = File.ReadAllBytes(stem + ".docx");
        var generated = DxpDocExport.Export(source);
        foreach (var docx in new[]
        {
            source,
            DxpDocToDocx.Project(generated).DocxBytes,
            DxpDocToDocx.Project(File.ReadAllBytes(stem + ".doc")).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(generated).DocxBytes)).DocxBytes
        })
        {
            var formatting = ReadEffectiveRunFormatting(docx);
            foreach (var story in new[]
            {
                "body", "section0.header.default", "section0.footer.default"
            })
            {
                Assert.Equal("cyan", formatting[$"{story}.paragraph0.char0.highlight"]);
                Assert.Equal("true", formatting[
                    $"{story}.paragraph0.char0.shading.none"]);
                Assert.DoesNotContain($"{story}.paragraph0.char0.shading.fill",
                    formatting.Keys);
                Assert.Equal("FFF2CC", formatting[
                    $"{story}.paragraph1.char0.shading.fill"]);
            }
            Assert.Empty(Validate(docx));
        }
    }

    [Fact]
    public void MixedListTableImageFieldDocumentRetainsHighlightsInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory,
            "WordMixedHighlightedListTableFieldPages");
        var source = File.ReadAllBytes(stem + ".docx");
        var generated = DxpDocExport.Export(source);
        var routes = new[]
        {
            source,
            DxpDocToDocx.Project(generated).DocxBytes,
            DxpDocToDocx.Project(File.ReadAllBytes(stem + ".doc")).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(generated).DocxBytes)).DocxBytes
        };
        for (var route = 0; route < routes.Length; route++)
        {
            var docx = routes[route];
            using var stream = new MemoryStream(docx);
            using var document = WordprocessingDocument.Open(stream, false);
            var main = document.MainDocumentPart!;
            var bodyRun = main.Document.Body!.Descendants<Run>()
                .Single(x => x.InnerText == "Beta");
            var headerRun = main.HeaderParts.SelectMany(x => x.Header!
                .Descendants<Run>()).Single(x => x.InnerText == "Second header item");
            var footerRun = main.FooterParts.SelectMany(x => x.Footer!
                .Descendants<Run>()).Single(x => x.InnerText == "Second footer item");
            Assert.Equal(HighlightColorValues.Red,
                bodyRun.RunProperties?.Highlight?.Val?.Value);
            Assert.Equal(HighlightColorValues.Cyan,
                headerRun.RunProperties?.Highlight?.Val?.Value);
            Assert.Equal(HighlightColorValues.Yellow,
                footerRun.RunProperties?.Highlight?.Val?.Value);
            var errors = new OpenXmlValidator(FileFormatVersions.Office2010)
                .Validate(document).ToArray();
            Assert.True(errors.Length == 0,
                $"Route {route}: {string.Join(" | ", errors.Select(x => x.Description).Take(3))}");
        }
    }

    [Theory]
    [InlineData("WordDocPropertyTitleBodyHeader", true, "Title", "Quarterly Report")]
    [InlineData("WordDocPropertyTitleBodyFooter", false, "Title", "Quarterly Report")]
    [InlineData("WordDocPropertyKeywordsBodyHeader", true, "Keywords", "Project Alpha")]
    [InlineData("WordDocPropertyCommentsBodyFooter", false, "Comments", "Approved draft")]
    public void DocPropertyFieldRemainsEditableInHeaderOrFooter(
        string name, bool header, string propertyName, string value)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, name);
        var source = File.ReadAllBytes(stem + ".docx");
        var generated = DxpDocExport.Export(source);
        var native = File.ReadAllBytes(stem + ".doc");
        foreach (var docx in new[]
        {
            source,
            DxpDocToDocx.Project(generated).DocxBytes,
            DxpDocToDocx.Project(native).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(generated).DocxBytes)).DocxBytes
        })
        {
            using var stream = new MemoryStream(docx);
            using var document = WordprocessingDocument.Open(stream, false);
            Assert.Equal(value, propertyName switch
            {
                "Title" => document.PackageProperties.Title,
                "Keywords" => document.PackageProperties.Keywords,
                "Comments" => document.PackageProperties.Description,
                _ => throw new InvalidOperationException()
            });
            var main = document.MainDocumentPart!;
            var story = header
                ? (DocumentFormat.OpenXml.OpenXmlElement)main.HeaderParts.Single(x =>
                    x.Header!.InnerText.Contains("Reference header")).Header!
                : main.FooterParts.Single(x => x.Footer!.InnerText.Contains("of"))
                    .Footer!;
            Assert.True(story.Descendants<FieldCode>().Any(field =>
                    field.Text.Contains("DOCPROPERTY " + propertyName,
                        StringComparison.OrdinalIgnoreCase)) ||
                story.Descendants<SimpleField>().Any(field =>
                    field.Instruction?.Value?.Contains("DOCPROPERTY " + propertyName,
                        StringComparison.OrdinalIgnoreCase) == true));
            Assert.Contains(story.Descendants<Text>(), text =>
                text.Text == value);
            Assert.Equal(2, ReadEditableFieldSemantics(docx).Count(x =>
                x == "FIELD:DOCPROPERTY:" + propertyName));
            Assert.Equal(1, ReadEditableFieldSemantics(docx).Count(x =>
                x == "FIELD:MERGEFIELD:CustomerName"));
            Assert.Empty(Validate(docx));
        }
    }

    [Theory]
    [InlineData("WordDocPropertyTitleBodyWithMergeStories", "Title", "Quarterly Report")]
    [InlineData("WordUnicodeDocPropertyTitleBodyWithMergeStories", "Title", "Résumé – Über")]
    [InlineData("WordDocPropertySubjectBodyWithMergeStories", "Subject", "Internal Review")]
    [InlineData("WordDocPropertyAuthorBodyWithMergeStories", "Author", "Ada Lovelace")]
    [InlineData("WordDocPropertyKeywordsBodyWithMergeStories", "Keywords", "Project Alpha")]
    [InlineData("WordDocPropertyCommentsBodyWithMergeStories", "Comments", "Approved draft")]
    public void DocPropertyFieldRetainsMetadataAndCachedVisibleResultInBody(
        string name, string propertyName, string value)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, name);
        var source = File.ReadAllBytes(stem + ".docx");
        var generated = DxpDocExport.Export(source);
        var native = File.ReadAllBytes(stem + ".doc");
        foreach (var docx in new[]
        {
            source,
            DxpDocToDocx.Project(generated).DocxBytes,
            DxpDocToDocx.Project(native).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(generated).DocxBytes)).DocxBytes
        })
        {
            using var stream = new MemoryStream(docx);
            using var document = WordprocessingDocument.Open(stream, false);
            Assert.Equal(value, propertyName switch
            {
                "Title" => document.PackageProperties.Title,
                "Subject" => document.PackageProperties.Subject,
                "Author" => document.PackageProperties.Creator,
                "Keywords" => document.PackageProperties.Keywords,
                "Comments" => document.PackageProperties.Description,
                _ => throw new InvalidOperationException()
            });
            var body = document.MainDocumentPart!.Document.Body!;
            Assert.True(body.Descendants<FieldCode>().Any(field =>
                    field.Text.Contains("DOCPROPERTY " + propertyName,
                        StringComparison.OrdinalIgnoreCase)) ||
                body.Descendants<SimpleField>().Any(field =>
                    field.Instruction?.Value?.Contains("DOCPROPERTY " + propertyName,
                        StringComparison.OrdinalIgnoreCase) == true));
            Assert.Contains(body.Descendants<Text>(), text => text.Text == value);
            Assert.Equal(1, ReadEditableFieldSemantics(docx).Count(x =>
                x == "FIELD:DOCPROPERTY:" + propertyName));
            Assert.Equal(2, ReadEditableFieldSemantics(docx).Count(x =>
                x == "FIELD:MERGEFIELD:CustomerName"));
            Assert.Empty(Validate(docx));
        }
    }

    [Fact]
    public void MergeFieldsRemainEditableWithCachedResultsInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, "WordMergeFieldsAllStories");
        var source = File.ReadAllBytes(stem + ".docx");
        var generated = DxpDocExport.Export(source);
        var native = File.ReadAllBytes(stem + ".doc");
        foreach (var docx in new[]
        {
            source,
            DxpDocToDocx.Project(generated).DocxBytes,
            DxpDocToDocx.Project(native).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(generated).DocxBytes)).DocxBytes
        })
        {
            using var stream = new MemoryStream(docx);
            using var document = WordprocessingDocument.Open(stream, false);
            var main = document.MainDocumentPart!;
            foreach (var story in new DocumentFormat.OpenXml.OpenXmlElement[]
            {
                main.Document.Body!,
                main.HeaderParts.Single(x => x.Header!.InnerText.Contains("Reference header"))
                    .Header!,
                main.FooterParts.Single(x => x.Footer!.InnerText.Contains("of"))
                    .Footer!
            })
            {
                Assert.Contains(story.Descendants<FieldCode>(), field =>
                    field.Text.Contains("MERGEFIELD CustomerName",
                        StringComparison.OrdinalIgnoreCase));
                Assert.Contains(story.Descendants<Text>(), text => text.Text == "Alice");
            }
            Assert.Equal(3, ReadEditableFieldSemantics(docx).Count(x =>
                x == "FIELD:MERGEFIELD:CustomerName"));
            Assert.Empty(Validate(docx));
        }
    }

    [Fact]
    public void IfFieldInsideNumberedTableRetainsEditableResultAndHighlight()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, "WordIfInNumberedTableMixedPages");
        var source = File.ReadAllBytes(stem + ".docx");
        var generated = DxpDocExport.Export(source);
        var native = File.ReadAllBytes(stem + ".doc");
        foreach (var docx in new[]
        {
            source,
            DxpDocToDocx.Project(generated).DocxBytes,
            DxpDocToDocx.Project(native).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(generated).DocxBytes)).DocxBytes
        })
        {
            using var stream = new MemoryStream(docx);
            using var document = WordprocessingDocument.Open(stream, false);
            var main = document.MainDocumentPart!;
            var cell = main.Document.Body!.Descendants<Table>().First()
                .Elements<TableRow>().First().Elements<TableCell>().ElementAt(1);
            Assert.Contains(cell.Descendants<FieldCode>(), field =>
                field.Text.Contains("IF 1 = 1 \"Beta\" \"No\"",
                    StringComparison.OrdinalIgnoreCase));
            Assert.Contains(cell.Descendants<Run>(), run =>
                run.InnerText == "Beta" &&
                run.RunProperties?.GetFirstChild<Highlight>()?.Val?.Value ==
                    HighlightColorValues.Red);
            Assert.NotNull(cell.Descendants<NumberingProperties>().FirstOrDefault());
            Assert.Equal(1, ReadEditableFieldSemantics(docx).Count(x =>
                x == "FIELD:IF:1:=:1:Beta:No"));
            Assert.Contains("FIELD:PAGE", ReadEditableFieldSemantics(docx));
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2010)
                .Validate(document));
        }
    }

    [Fact]
    public void SectionNumberFieldsRemainEditableAcrossIndependentSectionFooters()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, "WordSectionNumberAcrossSections");
        var source = File.ReadAllBytes(stem + ".docx");
        var generated = DxpDocExport.Export(source);
        var native = File.ReadAllBytes(stem + ".doc");
        foreach (var docx in new[]
        {
            source,
            DxpDocToDocx.Project(generated).DocxBytes,
            DxpDocToDocx.Project(native).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(generated).DocxBytes)).DocxBytes
        })
        {
            using var stream = new MemoryStream(docx);
            using var document = WordprocessingDocument.Open(stream, false);
            var main = document.MainDocumentPart!;
            Assert.Equal(3, main.Document.Body!.Descendants<SectionProperties>().Count());
            var fieldCount = main.FooterParts.Sum(x => x.Footer!
                .Descendants<FieldCode>().Count(field =>
                    Regex.IsMatch(field.Text, @"\bSECTION\b", RegexOptions.IgnoreCase)) +
                x.Footer.Descendants<SimpleField>().Count(field =>
                    Regex.IsMatch(field.Instruction?.Value ?? string.Empty,
                        @"\bSECTION\b", RegexOptions.IgnoreCase)));
            Assert.Equal(2, fieldCount);
            Assert.Equal(2, ReadEditableFieldSemantics(docx).Count(x =>
                x == "FIELD:SECTION"));
            Assert.Contains(main.FooterParts.SelectMany(x => x.Footer!
                .Descendants<Text>()), text => text.Text == "2");
            Assert.Contains(main.FooterParts.SelectMany(x => x.Footer!
                .Descendants<Text>()), text => text.Text == "3");
            Assert.Empty(Validate(docx));
        }
    }

    [Fact]
    public void SectionPagesFieldUsesEffectiveFirstFooterAcrossSections()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, "WordSectionPagesAcrossSections");
        var source = File.ReadAllBytes(stem + ".docx");
        var generated = DxpDocExport.Export(source);
        var native = File.ReadAllBytes(stem + ".doc");
        foreach (var docx in new[]
        {
            source,
            DxpDocToDocx.Project(generated).DocxBytes,
            DxpDocToDocx.Project(native).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(generated).DocxBytes)).DocxBytes
        })
        {
            using var stream = new MemoryStream(docx);
            using var document = WordprocessingDocument.Open(stream, false);
            var main = document.MainDocumentPart!;
            Assert.Equal(3, main.Document.Body!.Descendants<SectionProperties>().Count());
            Assert.Single(main.FooterParts.SelectMany(x => x.Footer!
                .Descendants<FieldCode>()), field =>
                    field.Text.Contains("SECTIONPAGES", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(main.FooterParts.SelectMany(x => x.Footer!
                .Descendants<Text>()), text => text.Text == "1");
            Assert.Equal(1, ReadEditableFieldSemantics(docx).Count(x =>
                x == "FIELD:SECTIONPAGES"));
            Assert.Empty(Validate(docx));
        }
    }

    [Fact]
    public void SectionPagesFieldsRemainEditableInAllStoriesThroughBothDocRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, "WordSectionPagesAllStories");
        var source = File.ReadAllBytes(stem + ".docx");
        var generated = DxpDocExport.Export(source);
        var native = File.ReadAllBytes(stem + ".doc");
        foreach (var docx in new[]
        {
            source,
            DxpDocToDocx.Project(generated).DocxBytes,
            DxpDocToDocx.Project(native).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(generated).DocxBytes)).DocxBytes
        })
        {
            using var stream = new MemoryStream(docx);
            using var document = WordprocessingDocument.Open(stream, false);
            var main = document.MainDocumentPart!;
            foreach (var story in new DocumentFormat.OpenXml.OpenXmlElement[]
            {
                main.Document.Body!,
                main.HeaderParts.Single(x => x.Header!.InnerText.Contains("Reference header"))
                    .Header!,
                main.FooterParts.Single(x => x.Footer!.InnerText.Contains("of"))
                    .Footer!
            })
            {
                Assert.Contains(story.Descendants<FieldCode>(), field =>
                    field.Text.Contains("SECTIONPAGES", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(story.Descendants<Text>(), text => text.Text == "2");
            }
            Assert.Empty(Validate(docx));
            Assert.Equal(3, ReadEditableFieldSemantics(docx).Count(x =>
                x == "FIELD:SECTIONPAGES"));
        }
    }

    [Fact]
    public void IfFieldsRemainEditableInBodyHeaderAndFooterThroughBothDocRoutes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, "WordIfFieldsAllStories");
        var source = File.ReadAllBytes(stem + ".docx");
        var generated = DxpDocExport.Export(source);
        var native = File.ReadAllBytes(stem + ".doc");
        foreach (var docx in new[]
        {
            source,
            DxpDocToDocx.Project(generated).DocxBytes,
            DxpDocToDocx.Project(native).DocxBytes,
            DxpDocToDocx.Project(DxpDocExport.Export(
                DxpDocToDocx.Project(generated).DocxBytes)).DocxBytes
        })
        {
            using var stream = new MemoryStream(docx);
            using var document = WordprocessingDocument.Open(stream, false);
            var main = document.MainDocumentPart!;
            foreach (var story in new DocumentFormat.OpenXml.OpenXmlElement[]
            {
                main.Document.Body!,
                main.HeaderParts.Single(x => x.Header!.InnerText.Contains("Reference header"))
                    .Header!,
                main.FooterParts.Single(x => x.Footer!.InnerText.Contains("of"))
                    .Footer!
            })
            {
                Assert.Contains(story.Descendants<FieldCode>(), field =>
                    field.Text.Contains("IF 1 = 1 \"YES\" \"NO\"",
                        StringComparison.OrdinalIgnoreCase));
                Assert.Contains(story.Descendants<Text>(), text => text.Text == "YES");
            }
            Assert.Empty(Validate(docx));
            Assert.Equal(3, ReadEditableFieldSemantics(docx).Count(x =>
                x == "FIELD:IF:1:=:1:YES:NO"));
        }
        if (Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_WORD") != "1") return;
        var output = Path.Combine(Path.GetTempPath(),
            $"docxport-if-field-{Guid.NewGuid():N}.doc");
        File.WriteAllBytes(output, generated);
        dynamic word = Activator.CreateInstance(Type.GetTypeFromProgID(
            "Word.Application")!)!;
        word.Visible = false;
        dynamic? opened = null;
        try
        {
            opened = word.Documents.Open(output, ReadOnly: true,
                AddToRecentFiles: false);
            Assert.Equal(7, (int)opened.Fields.Item(2).Type);
            Assert.Equal(7, (int)opened.Sections.Item(1).Headers.Item(1)
                .Range.Fields.Item(1).Type);
            Assert.Equal(7, (int)opened.Sections.Item(1).Footers.Item(1)
                .Range.Fields.Item(1).Type);
        }
        finally
        {
            if (opened != null) opened.Close(false);
            word.Quit(false);
            File.Delete(output);
        }
    }

    [Fact]
    public void WordListParagraphRetainsBuiltInStyleIdentity()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        using var input = File.OpenRead(Path.Combine(directory, "WordListsAllStories.doc"));
        using var index = new DocTextIndexWalker().Index(input);
        Assert.Equal(179, Assert.Single(index.StyleDefinitions,
            x => x.Name == "List Paragraph").InvariantStyleId);
    }

    [Fact]
    public void PairedCorpusExercisesEditableListsInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var lists = ReadListSemantics(File.ReadAllBytes(Path.Combine(directory,
            "WordListsAllStories.docx")));
        Assert.Equal(7, lists.Count);
        Assert.Equal("decimal|%1.|1|tab|720|360|", lists["body.paragraph0"]);
        Assert.Equal(lists["body.paragraph0"],
            lists["section0.header.default.paragraph0"]);
        Assert.Equal(lists["body.paragraph0"],
            lists["section0.footer.default.paragraph0"]);
    }

    [Fact]
    public void PairedCorpusExercisesMixedBulletAndNumberedLists()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var lists = ReadListSemantics(File.ReadAllBytes(Path.Combine(directory,
            "WordMixedListsAllStories.docx")));
        Assert.Equal(7, lists.Count);
        Assert.StartsWith("bullet|", lists["body.paragraph0"]);
        Assert.EndsWith("|Symbol", lists["body.paragraph0"]);
        Assert.StartsWith("decimal|%1.|1|", lists[
            "section0.header.default.paragraph0"]);
        Assert.StartsWith("decimal|%1.|1|", lists[
            "section0.footer.default.paragraph0"]);
    }

    [Theory]
    [InlineData("WordVisibleDecimalZeroListsAllStories", 0x16, true)]
    [InlineData("WordVisibleNoMarkerListsAllStories", 0xFF, true)]
    [InlineData("WordSymbolicCircleOnly", 0x12, false)]
    [InlineData("WordSymbolicDashOnly", 0x39, false)]
    public void WordListFormatUsesNativeDocCodeInAllStories(string fixtureName,
        byte expectedCode, bool allStories)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, fixtureName);
        foreach (var bytes in new[]
        {
            File.ReadAllBytes(stem + ".doc"),
            DxpDocExport.Export(File.ReadAllBytes(stem + ".docx"))
        })
        {
            using var input = new MemoryStream(bytes);
            using var index = new DocTextIndexWalker().Index(input);
            Assert.Equal(3, index.Lists.Definitions.Count);
            if (allStories)
                Assert.All(index.Lists.Definitions,
                    definition => Assert.Equal(expectedCode,
                        Assert.Single(definition.Levels).NumberFormat));
            else
                Assert.Contains(index.Lists.Definitions,
                    definition => Assert.Single(definition.Levels).NumberFormat == expectedCode);
        }
    }

    [Fact]
    public void WordFullAndHalfWidthDecimalListsUseNativeDocCodes()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, "WordVisibleWidthDecimalListsAllStories");
        foreach (var bytes in new[]
        {
            File.ReadAllBytes(stem + ".doc"),
            DxpDocExport.Export(File.ReadAllBytes(stem + ".docx"))
        })
        {
            using var input = new MemoryStream(bytes);
            using var index = new DocTextIndexWalker().Index(input);
            Assert.Equal(new byte[] { 0x0E, 0x0E, 0x0F },
                index.Lists.Definitions.Select(x => Assert.Single(x.Levels).NumberFormat)
                    .OrderBy(x => x));
        }
    }

    [Theory]
    [InlineData("WordRussianLowerListsAllStories", (byte)0x3A, false)]
    [InlineData("WordRussianUpperListsAllStories", (byte)0x3B, true)]
    public void WordRussianListsRetainNativeCodeInAllStories(
        string name, byte code, bool upperCase)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, name);
        var source = File.ReadAllBytes(stem + ".docx");
        var generated = DxpDocExport.Export(source);
        foreach (var binary in new[]
        {
            File.ReadAllBytes(stem + ".doc"), generated,
            DxpDocExport.Export(DxpDocToDocx.Project(generated).DocxBytes)
        })
        {
            using var input = new MemoryStream(binary);
            using var index = new DocTextIndexWalker().Index(input);
            Assert.Equal(3, index.Lists.Definitions.Count);
            Assert.All(index.Lists.Definitions, definition =>
                Assert.Equal(code,
                    Assert.Single(definition.Levels).NumberFormat));
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            using var document = WordprocessingDocument.Open(
                new MemoryStream(projected), false);
            Assert.Equal(3, document.MainDocumentPart!.NumberingDefinitionsPart!
                .Numbering!.Descendants<NumberingFormat>().Count(x =>
                    x.Val?.Value == (upperCase ? NumberFormatValues.RussianUpper :
                        NumberFormatValues.RussianLower)));
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
    }

    [Fact]
    public void WordIdeographDigitalListsRetainNativeCodeInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, "WordIdeographDigitalListsAllStories");
        var source = File.ReadAllBytes(stem + ".docx");
        var generated = DxpDocExport.Export(source);
        foreach (var binary in new[]
        {
            File.ReadAllBytes(stem + ".doc"), generated,
            DxpDocExport.Export(DxpDocToDocx.Project(generated).DocxBytes)
        })
        {
            using var input = new MemoryStream(binary);
            using var index = new DocTextIndexWalker().Index(input);
            Assert.Equal(3, index.Lists.Definitions.Count);
            Assert.All(index.Lists.Definitions, definition =>
                Assert.Equal((byte)0x0A,
                    Assert.Single(definition.Levels).NumberFormat));
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            using var document = WordprocessingDocument.Open(
                new MemoryStream(projected), false);
            Assert.Equal(3, document.MainDocumentPart!.NumberingDefinitionsPart!
                .Numbering!.Descendants<NumberingFormat>().Count(x =>
                    x.Val?.Value == NumberFormatValues.IdeographDigital));
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
    }

    [Fact]
    public void WordDecimalFullWidth2ListsRetainNativeCodeInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures",
            "DocKnownGaps"));
        var stem = Path.Combine(directory, "WordDecimalFullWidth2ListsAllStories");
        var source = File.ReadAllBytes(stem + ".docx");
        Assert.Throws<NotSupportedException>(() => DxpDocExport.Export(source));
        var native = File.ReadAllBytes(stem + ".doc");
        using (var input = new MemoryStream(native))
        {
            using var index = new DocTextIndexWalker().Index(input);
            Assert.Equal(3, index.Lists.Definitions.Count);
            Assert.All(index.Lists.Definitions, definition =>
                Assert.Equal((byte)0x13,
                    Assert.Single(definition.Levels).NumberFormat));
            var projected = DxpDocToDocx.Project(native).DocxBytes;
            using var document = WordprocessingDocument.Open(
                new MemoryStream(projected), false);
            Assert.Equal(3, document.MainDocumentPart!.NumberingDefinitionsPart!
                .Numbering!.Descendants<NumberingFormat>().Count(x =>
                    x.Val?.Value == NumberFormatValues.DecimalFullWidth2));
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
    }

    [Fact]
    public void WordNativeChicagoListProjectsWithEditableNumberFormat()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures",
            "DocKnownGaps"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            "WordSymbolicChicagoOnly.doc"));
        using (var input = new MemoryStream(native))
        using (var index = new DocTextIndexWalker().Index(input))
            Assert.Contains(index.Lists.Definitions, definition =>
                definition.Levels.Any(level => level.NumberFormat == 0x09));

        using var projected = new MemoryStream(DxpDocToDocx.Project(native).DocxBytes);
        using var document = WordprocessingDocument.Open(projected, false);
        var numbering = document.MainDocumentPart!.NumberingDefinitionsPart!
            .Numbering!;
        Assert.Contains(numbering.Descendants<NumberingFormat>(), format =>
            format.Val?.Value == NumberFormatValues.Chicago);
        Assert.Empty(new OpenXmlValidator().Validate(document));

        if (Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_WORD") == "1")
        {
            var path = Path.Combine(Path.GetTempPath(),
                $"docxport-chicago-projection-{Guid.NewGuid():N}.docx");
            File.WriteAllBytes(path, projected.ToArray());
            dynamic word = Activator.CreateInstance(
                Type.GetTypeFromProgID("Word.Application")!)!;
            word.Visible = false;
            try
            {
                dynamic source = word.Documents.Open(Path.Combine(directory,
                    "WordSymbolicChicagoOnly.docx"), ReadOnly: true);
                dynamic result = word.Documents.Open(path, ReadOnly: true);
                try
                {
                    var sourceLabels = new List<string>();
                    var resultLabels = new List<string>();
                    for (var i = 1; i <= (int)source.Paragraphs.Count; i++)
                    {
                        var label = (string)source.Paragraphs.Item(i)
                            .Range.ListFormat.ListString;
                        if (!string.IsNullOrEmpty(label)) sourceLabels.Add(label);
                    }
                    for (var i = 1; i <= (int)result.Paragraphs.Count; i++)
                    {
                        var label = (string)result.Paragraphs.Item(i)
                            .Range.ListFormat.ListString;
                        if (!string.IsNullOrEmpty(label)) resultLabels.Add(label);
                    }
                    Assert.NotEmpty(sourceLabels);
                    Assert.Equal(sourceLabels, resultLabels);
                }
                finally
                {
                    result.Close(false);
                    source.Close(false);
                }
            }
            finally
            {
                word.Quit(false);
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void PairedCorpusExercisesNestedNumberingAndContinuation()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var lists = ReadListSemantics(File.ReadAllBytes(Path.Combine(directory,
            "WordNestedListsAllStories.docx")));
        Assert.Equal(7, lists.Count);
        Assert.StartsWith("decimal|%1.|1|tab|720|360|",
            lists["body.paragraph0"]);
        Assert.StartsWith("decimal|%2)|1|tab|1440|360|",
            lists["body.paragraph1"]);
        Assert.Equal(lists["body.paragraph0"], lists["body.paragraph2"]);
    }

    [Fact]
    public void PairedCorpusExercisesListStartOverride()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var lists = ReadListSemantics(File.ReadAllBytes(Path.Combine(directory,
            "WordRestartListsAllStories.docx")));
        Assert.Equal(7, lists.Count);
        Assert.StartsWith("decimal|%1.|5|tab|720|360|",
            lists["body.paragraph0"]);
        Assert.StartsWith("decimal|%2)|1|tab|1440|360|",
            lists["body.paragraph1"]);
        Assert.Equal(lists["body.paragraph0"], lists["body.paragraph2"]);
        Assert.StartsWith("decimal|%1.|1|",
            lists["section0.header.default.paragraph0"]);
        Assert.StartsWith("decimal|%1.|1|",
            lists["section0.footer.default.paragraph0"]);
    }

    [Fact]
    public void CharacterStyleTogglesPreserveStyleLinksAndVisibleRunResets()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var docx = File.ReadAllBytes(Path.Combine(directory,
            "WordCharacterStyleToggleStories.docx"));
        using var input = new MemoryStream(DxpDocExport.Export(docx));
        using var index = new DocTextIndexWalker().Index(input);
        var baseStyle = Assert.Single(index.StyleDefinitions,
            x => x.Name == "ToggleBoldBase");
        var derived = Assert.Single(index.StyleDefinitions,
            x => x.Name == "ToggleBoldDerived");
        Assert.Equal(baseStyle.Index, derived.BasedOnIndex);
        Assert.True(baseStyle.CharacterFormatting.Bold);
        Assert.True(baseStyle.CharacterFormatting.Italic);
        Assert.True(derived.CharacterFormatting.Bold);
        Assert.True(derived.CharacterFormatting.Italic);
        Assert.Contains(index.CharacterFormatting, x =>
            x.Formatting.CharacterStyleIndex == derived.Index &&
            x.Formatting.Bold == false && x.Formatting.Italic == false);
    }

    [Fact]
    public void NativeCharacterStyleRelativeOverridesResolveAgainstBothStyles()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        using var input = File.OpenRead(Path.Combine(directory,
            "WordCharacterStyleDirectOverrideStories.doc"));
        using var index = new DocTextIndexWalker().Index(input);
        var baseStyle = Assert.Single(index.StyleDefinitions,
            x => x.Name == "ToggleBoldBase");
        Assert.Contains(index.CharacterFormatting, x =>
            x.CpStart == 0 && x.CpEnd == 5 &&
            x.Formatting.CharacterStyleIndex == baseStyle.Index &&
            x.Formatting.Bold == true && x.Formatting.Italic == true);
    }

    [Fact]
    public void InheritedStrikeToggleWritesVisibleRunReset()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var docx = File.ReadAllBytes(Path.Combine(directory,
            "WordCharacterStyleInheritedStrikeStories.docx"));
        using var input = new MemoryStream(DxpDocExport.Export(docx));
        using var index = new DocTextIndexWalker().Index(input);
        var baseStyle = Assert.Single(index.StyleDefinitions,
            x => x.Name == "ToggleBoldBase");
        Assert.Contains(index.CharacterFormatting, x =>
            x.CpStart == 0 && x.CpEnd == 5 &&
            x.Formatting.CharacterStyleIndex == baseStyle.Index &&
            x.Formatting.Strike == false);
    }

    [Fact]
    public void NativeRelativeStrikeToggleUsesCombinedStyleState()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        using var input = File.OpenRead(Path.Combine(directory,
            "WordCharacterStyleRelativeStrikeStories.doc"));
        using var index = new DocTextIndexWalker().Index(input);
        var baseStyle = Assert.Single(index.StyleDefinitions,
            x => x.Name == "ToggleBoldBase");
        Assert.Contains(index.CharacterFormatting, x =>
            x.CpStart == 0 && x.CpEnd == 5 &&
            x.Formatting.CharacterStyleIndex == baseStyle.Index &&
            x.Formatting.Strike == true);
    }

    [Theory]
    [MemberData(nameof(PairedFixtures))]
    public void WordPairedCorpusRetainsSectionStoryContentInBothRoutes(string name)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var docxBytes = File.ReadAllBytes(Path.Combine(directory, name + ".docx"));
        var docBytes = File.ReadAllBytes(Path.Combine(directory, name + ".doc"));

        var expected = ReadStories(docxBytes);
        var expectedStyleLinks = ReadLinkedCustomStyles(docxBytes);
        var expectedStyleRelationships = ReadStyleRelationships(docxBytes,
            customOnly: true);
        var fromDoc = DxpDocToDocx.Project(docBytes);
        Assert.Empty(fromDoc.Coverage.OmittedCharacters);
        Assert.Empty(fromDoc.Coverage.ApproximateCharacters);
        Assert.Empty(fromDoc.Coverage.DeferredParts);
        Assert.Equal(expected, ReadStories(fromDoc.DocxBytes));
        Assert.Equal(ReadEditableFieldSemantics(docxBytes),
            ReadEditableFieldSemantics(fromDoc.DocxBytes));
        Assert.Equal(ReadSupportedCoreProperties(docxBytes),
            ReadSupportedCoreProperties(fromDoc.DocxBytes));
        AssertExplicitHyphenationSettings(docxBytes, fromDoc.DocxBytes);
        AssertDirectEmptyParagraphMarkFormatting(docxBytes, fromDoc.DocxBytes);
        AssertCustomStyleRunFormatting(docxBytes, fromDoc.DocxBytes);
        AssertCustomStyleParagraphFormatting(docxBytes, fromDoc.DocxBytes);
        AssertLinkedCustomStylesPreserved(expectedStyleLinks, fromDoc.DocxBytes);
        AssertDocStyleRelationshipsPreserved(docBytes, fromDoc.DocxBytes);
        if (name is "WordReferenceWithStories" or "WordBalancedSpaceWithStories")
            Assert.Equal(ReadBalanceSetting(docxBytes),
                ReadBalanceSetting(fromDoc.DocxBytes));
        if (name == "WordExplicitColumnBreakStories")
            Assert.Equal(ReadColumnBreaks(docxBytes),
                ReadColumnBreaks(fromDoc.DocxBytes));
        if (name == "WordSixSlotStyledTablePageFields")
        {
            var expectedFooterFields = ReadEffectiveFooterPageFieldCounts(docxBytes);
            Assert.Equal(new[] { 1, 1, 1, 1, 1, 2, 1, 1, 2 },
                expectedFooterFields);
        }
        if (name == "WordSixSlotStyledTableLinks")
        {
            var fields = ReadEditableFieldSemantics(docxBytes);
            Assert.Contains("LINK:https://example.org/default", fields);
            Assert.Contains("LINK:https://example.org/even", fields);
            Assert.Equal(new[]
            {
                "https://example.org/default", "", "", "",
                "https://example.org/even", "", "",
                "https://example.org/even", ""
            }, ReadEffectiveFooterHyperlinkTargets(docxBytes));
        }
        if (name == "WordSixSlotStyledTableHeaderLinks")
            Assert.Equal(new[]
            {
                "", "", "https://example.org/first-header",
                "https://example.org/default-header", "", "",
                "https://example.org/default-header", "", ""
            }, ReadEffectiveHeaderHyperlinkTargets(docxBytes));
        if (name == "WordInternalBookmarkLinkStories")
        {
            Assert.Contains("LINK:#BodyTarget", ReadEditableFieldSemantics(docxBytes));
            Assert.Equal(new[]
            {
                "", "", "https://example.org/first-header",
                "https://example.org/default-header", "", "#BodyTarget",
                "https://example.org/default-header", "", "#BodyTarget"
            }, ReadEffectiveHeaderHyperlinkTargets(docxBytes));
        }
        Assert.Equal(ReadEffectiveFooterPageFieldCounts(docxBytes),
            ReadEffectiveFooterPageFieldCounts(fromDoc.DocxBytes));
        Assert.Equal(ReadEffectiveFooterHyperlinkTargets(docxBytes),
            ReadEffectiveFooterHyperlinkTargets(fromDoc.DocxBytes));
        Assert.Equal(ReadEffectiveHeaderHyperlinkTargets(docxBytes),
            ReadEffectiveHeaderHyperlinkTargets(fromDoc.DocxBytes));
        if (name == "WordDateTimeFields")
            Assert.Equal(ReadDateTimeInstructions(docxBytes),
                ReadDateTimeInstructions(fromDoc.DocxBytes));
        Assert.Equal(ReadSectionBodyText(docxBytes), ReadSectionBodyText(fromDoc.DocxBytes));
        Assert.Equal(ReadEffectiveHeaderFooterStories(docxBytes),
            ReadEffectiveHeaderFooterStories(fromDoc.DocxBytes));
        AssertExplicitSectionLayout(docxBytes, fromDoc.DocxBytes);
        AssertEffectiveParagraphLayout(docxBytes, fromDoc.DocxBytes);
        AssertParagraphDecorations(docxBytes, fromDoc.DocxBytes);
        AssertEffectiveTabStops(docxBytes, fromDoc.DocxBytes);
        AssertDefaultTabInterval(docxBytes, fromDoc.DocxBytes);
        AssertMirrorMargins(docxBytes, fromDoc.DocxBytes);
        AssertGutterAtTop(docxBytes, fromDoc.DocxBytes);
        AssertEffectiveRunFormatting(docxBytes, fromDoc.DocxBytes,
            name is "WordCharacterStyleToggleStories" or
                "WordLinkedStylesAllStories" or
                "WordCharacterStyleBaseToggleStories" or
                "WordCharacterStyleStrikeOverrideStories" or
                "WordCharacterStyleInheritedStrikeStories" or
                "WordCharacterStyleRelativeStrikeStories",
            name is "WordCharacterStyleInheritedStrikeStories" or
                "WordCharacterStyleRelativeStrikeStories",
            name == "WordVisibleThemeTintShadeAllStories");
        if (name == "WordTransformedFloatingImageStories")
            AssertNativeFlattenedHeaderGeometry(docxBytes, fromDoc.DocxBytes);
        else AssertDrawingGeometry(docxBytes, fromDoc.DocxBytes);
        AssertTableGrid(docxBytes, fromDoc.DocxBytes);
        if (name is "WordEmptyParagraphsAllStories" or
            "WordVisibleSizedEmptyParagraphsAllStories")
            Assert.Equal(ReadInteriorEmptyParagraphMarks(docxBytes,
                    details: name == "WordVisibleSizedEmptyParagraphsAllStories"),
                ReadInteriorEmptyParagraphMarks(fromDoc.DocxBytes,
                    details: name == "WordVisibleSizedEmptyParagraphsAllStories"));
        Assert.Equal(ReadListSemantics(docxBytes), ReadListSemantics(fromDoc.DocxBytes));
        Assert.Equal(ReadListSemantics(docxBytes, true),
            ReadListSemantics(fromDoc.DocxBytes, true));
        var fromDocValidation = Validate(fromDoc.DocxBytes);
        Assert.True(fromDocValidation.Count == 0,
            string.Join(Environment.NewLine, fromDocValidation.Select(x => x.Description)));

        var writtenDoc = DxpDocExport.Export(docxBytes);
        using (var input = new MemoryStream(writtenDoc))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            Assert.NotEmpty(index.Pieces);
            Assert.NotEmpty(index.ParagraphStyles);
        }
        var fromWrittenDoc = DxpDocToDocx.Project(writtenDoc);
        Assert.Empty(fromWrittenDoc.Coverage.OmittedCharacters);
        Assert.Empty(fromWrittenDoc.Coverage.ApproximateCharacters);
        Assert.Empty(fromWrittenDoc.Coverage.DeferredParts);
        Assert.Equal(expected, ReadStories(fromWrittenDoc.DocxBytes));
        Assert.Equal(ReadEditableFieldSemantics(docxBytes),
            ReadEditableFieldSemantics(fromWrittenDoc.DocxBytes));
        Assert.Equal(ReadSupportedCoreProperties(docxBytes),
            ReadSupportedCoreProperties(fromWrittenDoc.DocxBytes));
        AssertExplicitHyphenationSettings(docxBytes, fromWrittenDoc.DocxBytes);
        AssertDirectEmptyParagraphMarkFormatting(docxBytes, fromWrittenDoc.DocxBytes);
        AssertCustomStyleRunFormatting(docxBytes, fromWrittenDoc.DocxBytes);
        AssertCustomStyleParagraphFormatting(docxBytes, fromWrittenDoc.DocxBytes);
        AssertLinkedCustomStylesPreserved(expectedStyleLinks, fromWrittenDoc.DocxBytes);
        AssertCustomStyleRelationshipsPreserved(expectedStyleRelationships,
            fromWrittenDoc.DocxBytes);
        if (name is "WordReferenceWithStories" or "WordBalancedSpaceWithStories")
            Assert.Equal(ReadBalanceSetting(docxBytes),
                ReadBalanceSetting(fromWrittenDoc.DocxBytes));
        if (name == "WordExplicitColumnBreakStories")
            Assert.Equal(ReadColumnBreaks(docxBytes),
                ReadColumnBreaks(fromWrittenDoc.DocxBytes));
        Assert.Equal(ReadEffectiveFooterPageFieldCounts(docxBytes),
            ReadEffectiveFooterPageFieldCounts(fromWrittenDoc.DocxBytes));
        Assert.Equal(ReadEffectiveFooterHyperlinkTargets(docxBytes),
            ReadEffectiveFooterHyperlinkTargets(fromWrittenDoc.DocxBytes));
        Assert.Equal(ReadEffectiveHeaderHyperlinkTargets(docxBytes),
            ReadEffectiveHeaderHyperlinkTargets(fromWrittenDoc.DocxBytes));
        if (name == "WordDateTimeFields")
            Assert.Equal(ReadDateTimeInstructions(docxBytes),
                ReadDateTimeInstructions(fromWrittenDoc.DocxBytes));
        Assert.Equal(ReadSectionBodyText(docxBytes), ReadSectionBodyText(fromWrittenDoc.DocxBytes));
        Assert.Equal(ReadEffectiveHeaderFooterStories(docxBytes),
            ReadEffectiveHeaderFooterStories(fromWrittenDoc.DocxBytes));
        AssertExplicitSectionLayout(docxBytes, fromWrittenDoc.DocxBytes);
        AssertEffectiveParagraphLayout(docxBytes, fromWrittenDoc.DocxBytes);
        AssertParagraphDecorations(docxBytes, fromWrittenDoc.DocxBytes);
        AssertEffectiveTabStops(docxBytes, fromWrittenDoc.DocxBytes);
        AssertDefaultTabInterval(docxBytes, fromWrittenDoc.DocxBytes);
        AssertMirrorMargins(docxBytes, fromWrittenDoc.DocxBytes);
        AssertGutterAtTop(docxBytes, fromWrittenDoc.DocxBytes);
        AssertEffectiveRunFormatting(docxBytes, fromWrittenDoc.DocxBytes,
            name is "WordCharacterStyleToggleStories" or
                "WordLinkedStylesAllStories" or
                "WordCharacterStyleBaseToggleStories" or
                "WordCharacterStyleStrikeOverrideStories" or
                "WordCharacterStyleInheritedStrikeStories" or
                "WordCharacterStyleRelativeStrikeStories",
            name is "WordCharacterStyleInheritedStrikeStories" or
                "WordCharacterStyleRelativeStrikeStories",
            name == "WordVisibleThemeTintShadeAllStories");
        AssertDrawingGeometry(docxBytes, fromWrittenDoc.DocxBytes);
        AssertTableGrid(docxBytes, fromWrittenDoc.DocxBytes);
        if (name is "WordEmptyParagraphsAllStories" or
            "WordVisibleSizedEmptyParagraphsAllStories")
            Assert.Equal(ReadInteriorEmptyParagraphMarks(docxBytes,
                    details: name == "WordVisibleSizedEmptyParagraphsAllStories"),
                ReadInteriorEmptyParagraphMarks(fromWrittenDoc.DocxBytes,
                    details: name == "WordVisibleSizedEmptyParagraphsAllStories"));
        Assert.Equal(ReadListSemantics(docxBytes),
            ReadListSemantics(fromWrittenDoc.DocxBytes));
        Assert.Equal(ReadListSemantics(docxBytes, true),
            ReadListSemantics(fromWrittenDoc.DocxBytes, true));
        Assert.Empty(Validate(fromWrittenDoc.DocxBytes));
    }

    private static Dictionary<string, string> ReadLinkedCustomStyles(byte[] bytes)
    {
        using var input = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(input, false);
        var styles = document.MainDocumentPart?.StyleDefinitionsPart?.Styles?
            .Elements<Style>().Where(x => x.StyleId?.Value != null)
            .ToDictionary(x => x.StyleId!.Value!, StringComparer.Ordinal)
            ?? new Dictionary<string, Style>(StringComparer.Ordinal);
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var style in styles.Values.Where(x => x.CustomStyle?.Value == true))
        {
            var linkedId = style.LinkedStyle?.Val?.Value;
            if (linkedId == null || !styles.TryGetValue(linkedId, out var linked) ||
                linked.CustomStyle?.Value != true) continue;
            result[style.StyleName?.Val?.Value ?? style.StyleId!.Value!] =
                linked.StyleName?.Val?.Value ?? linkedId;
        }
        return result;
    }

    private static void AssertLinkedCustomStylesPreserved(byte[] source, byte[] actual) =>
        AssertLinkedCustomStylesPreserved(ReadLinkedCustomStyles(source), actual);

    private static void AssertLinkedCustomStylesPreserved(
        IReadOnlyDictionary<string, string> expected, byte[] actual)
    {
        var observed = ReadLinkedCustomStyles(actual);
        foreach (var (name, partner) in expected)
            Assert.True(observed.TryGetValue(name, out var found) && found == partner,
                $"Linked style {name}: expected {partner}, observed " +
                (observed.TryGetValue(name, out var current) ? current : "<missing>"));
    }

    private static Dictionary<string, string> ReadStyleRelationships(byte[] bytes,
        bool customOnly)
    {
        using var input = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(input, false);
        var styles = document.MainDocumentPart?.StyleDefinitionsPart?.Styles?
            .Elements<Style>().Where(x => x.StyleId?.Value != null)
            .ToDictionary(x => x.StyleId!.Value!, StringComparer.Ordinal)
            ?? new Dictionary<string, Style>(StringComparer.Ordinal);
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var style in styles.Values.Where(x =>
            !customOnly || x.CustomStyle?.Value == true))
        {
            var name = style.StyleName?.Val?.Value ?? style.StyleId!.Value!;
            void Add(string relation, string? targetId)
            {
                if (targetId != null && styles.TryGetValue(targetId, out var target))
                    result[$"{name}.{relation}"] = target.StyleName?.Val?.Value ?? targetId;
            }
            var basedOnId = style.BasedOn?.Val?.Value;
            // The default character base is implicit in binary DOC.
            if (style.Type?.Value == StyleValues.Character &&
                basedOnId != null && styles.TryGetValue(basedOnId, out var baseStyle) &&
                baseStyle.StyleName?.Val?.Value == "Default Paragraph Font")
                basedOnId = null;
            Add("basedOn", basedOnId);
            Add("next", style.NextParagraphStyle?.Val?.Value);
            Add("linked", style.LinkedStyle?.Val?.Value);
        }
        return result;
    }

    private static void AssertCustomStyleRelationshipsPreserved(byte[] source,
        byte[] actual) => AssertCustomStyleRelationshipsPreserved(
            ReadStyleRelationships(source, customOnly: true), actual);

    private static void AssertCustomStyleRelationshipsPreserved(
        IReadOnlyDictionary<string, string> expected, byte[] actual)
    {
        var observed = ReadStyleRelationships(actual, customOnly: false);
        foreach (var (name, target) in expected)
            Assert.True(observed.TryGetValue(name, out var found) &&
                string.Equals(found, target, StringComparison.OrdinalIgnoreCase),
                $"Style relationship {name}: expected {target}, observed " +
                (observed.TryGetValue(name, out var current) ? current : "<missing>"));
    }

    private static void AssertDocStyleRelationshipsPreserved(byte[] sourceDoc,
        byte[] actualDocx)
    {
        using var input = new MemoryStream(sourceDoc);
        using var index = new DocTextIndexWalker().Index(input);
        var styles = index.StyleDefinitions.ToDictionary(x => x.Index);
        var expected = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        static string PrimaryName(string name) => name.Split(',', 2)[0];
        foreach (var style in index.StyleDefinitions.Where(x =>
            x.InvariantStyleId == 0x0FFE))
        {
            void Add(string relation, int? targetIndex)
            {
                if (targetIndex is int target && styles.TryGetValue(target, out var parent))
                {
                    if (relation == "basedOn" && style.Type == 2 &&
                        parent.Name == "Default Paragraph Font") return;
                    expected[$"{PrimaryName(style.Name)}.{relation}"] =
                        PrimaryName(parent.Name);
                }
            }
            Add("basedOn", style.BasedOnIndex);
            if (style.Type == 1) Add("next", style.NextIndex);
            Add("linked", style.LinkedStyleIndex);
        }
        AssertCustomStyleRelationshipsPreserved(expected, actualDocx);
    }

    private static Dictionary<int, int> ReadColumnBreaks(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return document.MainDocumentPart!.Document.Body!
            .Descendants<Paragraph>()
            .Select((paragraph, index) => (index, count: paragraph
                .Descendants<Break>()
                .Count(x => x.Type?.Value == BreakValues.Column)))
            .Where(x => x.count > 0)
            .ToDictionary(x => x.index, x => x.count);
    }

    [Fact]
    public void ExplicitColumnBreakRemainsAtItsBodyParagraphThroughDocHops()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var source = File.ReadAllBytes(Path.Combine(directory,
            "WordExplicitColumnBreakStories.docx"));
        var native = File.ReadAllBytes(Path.Combine(directory,
            "WordExplicitColumnBreakStories.doc"));
        var expected = new Dictionary<int, int> { [9] = 1 };
        Assert.Equal(expected, ReadColumnBreaks(source));
        foreach (var binary in new[] { native, DxpDocExport.Export(source) })
        {
            AssertIndexedColumnBreak(binary);
            var first = DxpDocToDocx.Project(binary).DocxBytes;
            Assert.Equal(expected, ReadColumnBreaks(first));
            Assert.Empty(Validate(first));
            var third = DxpDocToDocx.Project(DxpDocExport.Export(first)).DocxBytes;
            AssertIndexedColumnBreak(DxpDocExport.Export(first));
            Assert.Equal(expected, ReadColumnBreaks(third));
            Assert.Empty(Validate(third));
        }

        static void AssertIndexedColumnBreak(byte[] binary)
        {
            using var input = new MemoryStream(binary);
            using var index = new DocTextIndexWalker().Index(input);
            var paragraphs = DocStoryTextReader.Read(index, "Main").Paragraphs;
            Assert.Equal(9, Assert.Single(paragraphs
                .Select((paragraph, position) => (paragraph, position))
                .Where(x => x.paragraph.Atoms.Any(atom =>
                    atom.Kind == DocStoryAtomKind.ColumnBreak))).position);
            Assert.Single(paragraphs[9].Atoms,
                x => x.Kind == DocStoryAtomKind.ColumnBreak);
        }
    }

    [Theory]
    [InlineData("body", "WordVisibleSizedEmptyParagraphsAllStories")]
    [InlineData("header", "WordVisibleSizedEmptyParagraphsAllStories")]
    [InlineData("footer", "WordVisibleSizedEmptyParagraphsAllStories")]
    [InlineData("body", "WordFinalEmptyMarkAllStories")]
    [InlineData("header", "WordFinalEmptyMarkAllStories")]
    [InlineData("footer", "WordFinalEmptyMarkAllStories")]
    public void EmptyParagraphMarkFormattingIsPartOfThePairedGate(
        string story, string fixtureName)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            fixtureName + ".docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            var root = story switch
            {
                "body" => (DocumentFormat.OpenXml.OpenXmlElement)main.Document!.Body!,
                "header" => main.HeaderParts.Select(x => x.Header!)
                    .First(x => x.Descendants<Paragraph>().Any(p =>
                        p.ParagraphProperties?.GetFirstChild<ParagraphMarkRunProperties>()?
                            .GetFirstChild<FontSize>() != null)),
                _ => main.FooterParts.Select(x => x.Footer!)
                    .First(x => x.Descendants<Paragraph>().Any(p =>
                        p.ParagraphProperties?.GetFirstChild<ParagraphMarkRunProperties>()?
                            .GetFirstChild<FontSize>() != null))
            };
            var mark = root.Descendants<Paragraph>().First(p =>
                !p.Descendants<Text>().Any(t => t.Text.Length > 0) &&
                p.ParagraphProperties?.GetFirstChild<ParagraphMarkRunProperties>()?
                    .GetFirstChild<FontSize>() != null)
                .ParagraphProperties!.GetFirstChild<ParagraphMarkRunProperties>()!;
            mark.GetFirstChild<FontSize>()!.Val = "16";
            switch (root)
            {
                case Body: main.Document!.Save(); break;
                case Header header: header.Save(); break;
                case Footer footer: footer.Save(); break;
            }
        }
        var failure = Record.Exception(() =>
            AssertDirectEmptyParagraphMarkFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(story, failure.Message);
        Assert.Contains(".size", failure.Message);
    }

    [Theory]
    [InlineData("body", "underline")]
    [InlineData("footer", "strike")]
    public void EmptyParagraphMarkDecorationsArePartOfThePairedGate(
        string story, string property)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedEmptyMarkAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            var root = story == "body"
                ? (DocumentFormat.OpenXml.OpenXmlElement)main.Document!.Body!
                : main.FooterParts.Single().Footer!;
            var mark = root.Descendants<Paragraph>()
                .Select(p => p.ParagraphProperties?
                    .GetFirstChild<ParagraphMarkRunProperties>())
                .First(m => property == "underline"
                    ? m?.GetFirstChild<Underline>() != null
                    : m?.GetFirstChild<Strike>() != null)!;
            if (property == "underline") mark.GetFirstChild<Underline>()!.Remove();
            else mark.GetFirstChild<Strike>()!.Remove();
            if (story == "body") main.Document!.Save();
            else main.FooterParts.Single().Footer!.Save();
        }
        var failure = Record.Exception(() =>
            AssertDirectEmptyParagraphMarkFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(story, failure.Message);
        Assert.Contains(property, failure.Message);
    }

    [Fact]
    public void EmptyHeaderMarkFontIsPartOfThePairedGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedEmptyMarkAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var header = document.MainDocumentPart!.HeaderParts.Single().Header!;
            var mark = header.Descendants<Paragraph>()
                .Select(p => p.ParagraphProperties?
                    .GetFirstChild<ParagraphMarkRunProperties>())
                .First(m => m?.GetFirstChild<RunFonts>() != null)!;
            mark.GetFirstChild<RunFonts>()!.Remove();
            header.Save();
        }
        var failure = Record.Exception(() =>
            AssertDirectEmptyParagraphMarkFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("header", failure.Message);
        Assert.Contains("fontAscii", failure.Message);
    }

    [Fact]
    public void CustomStyleFontInheritanceIsPartOfThePairedGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedEmptyMarkAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var styles = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            var baseStyle = styles.Elements<Style>().Single(x =>
                x.StyleName?.Val?.Value == "Mark base");
            var fonts = baseStyle.StyleRunProperties!.GetFirstChild<RunFonts>()!;
            fonts.Ascii = "Times New Roman";
            fonts.HighAnsi = "Times New Roman";
            styles.Save();
        }
        var failure = Record.Exception(() =>
            AssertCustomStyleRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("Mark base.fontAscii", failure.Message);
    }

    [Theory]
    [InlineData("WordCharacterSpacingStyleOverridesAllStories", "characterSpacing")]
    [InlineData("WordCharacterSpacingStyleOverridesAllStories", "language.eastAsia")]
    [InlineData("WordKerningStyleAllStories", "kerning")]
    public void CustomStyleSpacingLanguageAndKerningArePartOfThePairedGate(
        string fixture, string property)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            fixture + ".docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    "Mirrored body");
            var runs = style.StyleRunProperties!;
            switch (property)
            {
                case "characterSpacing":
                    runs.GetFirstChild<Spacing>()!.Val = 0;
                    break;
                case "language.eastAsia":
                    runs.GetFirstChild<Languages>()!.EastAsia = "ko-KR";
                    break;
                default:
                    runs.GetFirstChild<Kern>()!.Val = 0;
                    break;
            }
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertCustomStyleRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("Mirrored body." + property, failure.Message);
    }
    [Theory]
    [InlineData("WordCharacterScaleLayeredStories", "scale")]
    [InlineData("WordBaselineOffsetLayeredStories", "position")]
    public void CustomStyleScaleAndPositionArePartOfThePairedGate(
        string fixture, string property)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            fixture + ".docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    "Visible Base");
            if (property == "scale")
                style.StyleRunProperties!.GetFirstChild<CharacterScale>()!.Val = 100;
            else
                style.StyleRunProperties!.GetFirstChild<Position>()!.Val = "0";
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertCustomStyleRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("Visible Base." + property, failure.Message);
    }
    [Fact]
    public void CustomStyleRunShadingIsPartOfThePairedGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyledRgbRunShadingAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    "Mirrored body");
            style.StyleRunProperties!.GetFirstChild<Shading>()!.Fill = "00FF00";
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertCustomStyleRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("Mirrored body.shading.fill", failure.Message);
    }
    [Theory]
    [InlineData("Mirrored body", "border.color")]
    [InlineData("Layered border", "border.style")]
    public void CustomStyleRunBorderIsPartOfThePairedGate(
        string styleName, string property)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordLayeredRunBordersAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    styleName);
            var border = style.StyleRunProperties!.GetFirstChild<Border>()!;
            if (property == "border.color") border.Color = "0000FF";
            else border.Val = BorderValues.Single;
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertCustomStyleRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(styleName + "." + property, failure.Message);
    }
    [Theory]
    [InlineData("Heading 4 Char", "complexItalic")]
    [InlineData("Heading 1 Char", "complexSize")]
    public void CustomStyleComplexScriptFormattingIsPartOfThePairedGate(
        string styleName, string property)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyleLayeredStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    styleName);
            var runs = style.StyleRunProperties!;
            if (property == "complexItalic")
                runs.GetFirstChild<ItalicComplexScript>()!.Val = false;
            else runs.GetFirstChild<FontSizeComplexScript>()!.Val = "24";
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertCustomStyleRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(styleName + "." + property, failure.Message);
    }
    [Fact]
    public void CustomStyleFitTextWidthIsPartOfThePairedGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRunFitTextParagraphStyleAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    "FitParagraphBody");
            style.StyleRunProperties!.GetFirstChild<FitText>()!.Val = 2500;
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertCustomStyleRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("FitParagraphBody.fitTextWidth", failure.Message);
    }
    [Fact]
    public void StyledCharacterEffectsReachBodyHeaderAndFooterThroughBothDocRoutes()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyledCharacterEffectsAllStories.docx"));
        var source = File.ReadAllBytes(path);
        var expected = ReadEffectiveRunFormatting(source);
        foreach (var story in new[] { "body.", ".header.", ".footer." })
        foreach (var effect in new[] { "shadow", "emboss", "imprint" })
            Assert.Contains(expected, x => x.Key.Contains(story,
                StringComparison.Ordinal) && x.Key.EndsWith("." + effect,
                StringComparison.Ordinal) && x.Value == "true");
        foreach (var binary in new[]
        {
            File.ReadAllBytes(Path.ChangeExtension(path, ".doc")),
            DxpDocExport.Export(source)
        })
        {
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            AssertEffectiveRunFormatting(source, projected);
            var repeated = DxpDocToDocx.Project(DxpDocExport.Export(projected))
                .DocxBytes;
            AssertEffectiveRunFormatting(source, repeated);
            Assert.Empty(Validate(repeated));
        }
    }
    [Theory]
    [InlineData("Visible Base", "shadow")]
    [InlineData("Visible Derived", "emboss")]
    [InlineData("Visible Accent", "imprint")]
    public void CustomStyleEffectsArePartOfThePairedGate(
        string styleName, string property)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyledCharacterEffectsAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var styles = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            var style = styles.Elements<Style>().Single(x =>
                x.StyleName?.Val?.Value == styleName);
            var properties = style.StyleRunProperties!;
            if (property == "shadow") properties.GetFirstChild<Shadow>()!.Val = false;
            else if (property == "emboss") properties.GetFirstChild<Emboss>()!.Val = false;
            else properties.GetFirstChild<Imprint>()!.Val = false;
            styles.Save();
        }
        var failure = Record.Exception(() =>
            AssertCustomStyleRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(styleName + "." + property, failure.Message);
    }
    [Fact]
    public void CustomStyleOutlineIsPartOfThePairedGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordCharacterEffectsLayeredStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var styles = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            var style = styles.Elements<Style>().Single(x =>
                x.StyleName?.Val?.Value == "Visible Base");
            style.StyleRunProperties!.GetFirstChild<Outline>()!.Val = false;
            styles.Save();
        }
        var failure = Record.Exception(() =>
            AssertCustomStyleRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("Visible Base.outline", failure.Message);
    }
    [Fact]
    public void UnexpectedEnabledRunEffectIsPartOfThePairedGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyleLayeredStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var body = document.MainDocumentPart!.Document!.Body!;
            var run = body.Descendants<Run>().First(x => x.InnerText == "Body layered");
            run.RunProperties ??= new RunProperties();
            run.RunProperties.AppendChild(new Shadow());
            document.MainDocumentPart.Document.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("Unexpected enabled run effect", failure.Message);
        Assert.Contains(".shadow", failure.Message);
    }
    [Fact]
    public void UnexpectedHiddenRunIsRejectedByPairedGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyleLayeredStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var run = document.MainDocumentPart!.Document!.Body!
                .Descendants<Run>().First(x => x.InnerText == "Body layered");
            run.RunProperties ??= new RunProperties();
            run.RunProperties.AppendChild(new Vanish());
            document.MainDocumentPart.Document.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("Unexpected enabled run effect", failure.Message);
        Assert.Contains(".hidden", failure.Message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnexpectedRunCapsAreRejectedByPairedGate(bool smallCaps)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyleLayeredStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var run = document.MainDocumentPart!.Document!.Body!
                .Descendants<Run>().First(x => x.InnerText == "Body layered");
            run.RunProperties ??= new RunProperties();
            if (smallCaps) run.RunProperties.AppendChild(new SmallCaps());
            else run.RunProperties.AppendChild(new Caps());
            document.MainDocumentPart.Document.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("Unexpected enabled run effect", failure.Message);
        Assert.Contains(smallCaps ? ".smallCaps" : ".caps", failure.Message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnexpectedRunDecorationIsRejectedByPairedGate(bool highlight)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyleLayeredStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var run = document.MainDocumentPart!.Document!.Body!
                .Descendants<Run>().First(x => x.InnerText == "Body layered");
            run.RunProperties ??= new RunProperties();
            if (highlight) run.RunProperties.AppendChild(new Highlight
                { Val = HighlightColorValues.Yellow });
            else run.RunProperties.AppendChild(new Underline
                { Val = UnderlineValues.Single });
            document.MainDocumentPart.Document.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("Unexpected run decoration", failure.Message);
        Assert.Contains(highlight ? ".highlight" : ".underline", failure.Message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnexpectedRunFillOrBorderIsRejectedByPairedGate(bool border)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyleLayeredStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var run = document.MainDocumentPart!.Document!.Body!
                .Descendants<Run>().First(x => x.InnerText == "Body layered");
            run.RunProperties ??= new RunProperties();
            if (border) run.RunProperties.AppendChild(new Border
                { Val = BorderValues.Single, Size = 8, Color = "FF0000" });
            else run.RunProperties.AppendChild(new Shading
                { Val = ShadingPatternValues.Clear, Fill = "FF0000" });
            document.MainDocumentPart.Document.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("Unexpected run decoration", failure.Message);
        Assert.Contains(border ? ".border.val" : ".shading.fill",
            failure.Message);
    }

    [Fact]
    public void CustomStyleDoubleStrikeIsPartOfThePairedGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordDoubleStrikeLayeredStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var styles = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            var style = styles.Elements<Style>().Single(x =>
                x.StyleName?.Val?.Value == "Visible Base");
            style.StyleRunProperties!.GetFirstChild<DoubleStrike>()!.Val = false;
            styles.Save();
        }
        var failure = Record.Exception(() =>
            AssertCustomStyleRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("Visible Base.doubleStrike", failure.Message);
    }
    [Fact]
    public void CustomStyleUnderlineColorIsPartOfThePairedGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordUnderlineColorStyleOverridesAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x =>
                    x.StyleName?.Val?.Value == "Mirrored body");
            style.StyleRunProperties!.GetFirstChild<Underline>()!.Color = "00FF00";
            document.MainDocumentPart.StyleDefinitionsPart.Styles.Save();
        }
        var failure = Record.Exception(() =>
            AssertCustomStyleRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("Mirrored body.underlineColor", failure.Message);
    }
    [Theory]
    [InlineData("WordUnderlineColorStyleOverridesAllStories", "Mirrored body",
        "underline")]
    [InlineData("WordCharacterStyleToggleStories", "ToggleBoldBase", "strike")]
    public void CustomStyleDecorationsArePartOfThePairedGate(
        string fixture, string styleName, string property)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            fixture + ".docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var styles = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            var style = styles.Elements<Style>().Single(x =>
                x.StyleName?.Val?.Value == styleName);
            if (property == "underline")
                style.StyleRunProperties!.GetFirstChild<Underline>()!.Remove();
            else style.StyleRunProperties!.GetFirstChild<Strike>()!.Val = false;
            styles.Save();
        }
        var failure = Record.Exception(() =>
            AssertCustomStyleRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(styleName + "." + property, failure.Message);
    }

    [Fact]
    public void CustomStyleSmallCapsIsPartOfThePairedGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedEmptyMarkAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var styles = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            var style = styles.Elements<Style>().Single(x =>
                x.StyleName?.Val?.Value == "Mark child");
            style.StyleRunProperties!.GetFirstChild<SmallCaps>()!.Val = false;
            styles.Save();
        }
        var failure = Record.Exception(() =>
            AssertCustomStyleRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("Mark child.smallCaps", failure.Message);
    }

    [Fact]
    public void CustomStyleCapsIsPartOfThePairedGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordCharacterStyleCapsResetAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var styles = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            var style = styles.Elements<Style>().Single(x =>
                x.StyleName?.Val?.Value == "Caps Character");
            style.StyleRunProperties!.GetFirstChild<Caps>()!.Val = false;
            styles.Save();
        }
        var failure = Record.Exception(() =>
            AssertCustomStyleRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("Caps Character.caps", failure.Message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnexpectedCustomStyleRunDecorationFailsThePairedGate(bool border)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyleLayeredStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    "Visible Base");
            if (border)
                style.StyleRunProperties!.AddChild(new Border
                {
                    Val = BorderValues.Single, Color = "FF0000", Size = 12,
                    Space = 2
                }, true);
            else style.StyleRunProperties!.AddChild(new Shading
            {
                Val = ShadingPatternValues.Clear, Fill = "00FF00"
            }, true);
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertCustomStyleRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(border ? "Visible Base.border." :
            "Visible Base.shading.", failure.Message);
    }
    [Fact]
    public void LiteralCustomStyleColorIsPartOfThePairedGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedDisplayStyleStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var styles = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            var baseStyle = styles.Elements<Style>().Single(x =>
                x.StyleName?.Val?.Value == "BaseDisplay");
            baseStyle.StyleRunProperties!.GetFirstChild<Color>()!.Val = "FF0000";
            styles.Save();
        }
        var failure = Record.Exception(() =>
            AssertCustomStyleRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("BaseDisplay.color", failure.Message);
    }

    private static void AssertCustomStyleRunFormatting(byte[] source,
        byte[] output)
    {
        var expected = ReadCustomStyleRunFormatting(source, literalColorsOnly: true);
        var observed = ReadCustomStyleRunFormatting(output);
        foreach (var (key, value) in expected)
            Assert.True(observed.TryGetValue(key, out var actual) &&
                (actual == value || key.EndsWith(".color", StringComparison.Ordinal) &&
                    NearLegacyRgb(value, actual)),
                $"Custom style run {key}: expected {value}, observed " +
                (observed.TryGetValue(key, out var found) ? found : "<missing>"));
        foreach (var (key, value) in observed)
        {
            if (expected.ContainsKey(key)) continue;
            var border = key.LastIndexOf(".border.", StringComparison.Ordinal);
            if (border >= 0 && observed.TryGetValue(
                key[..(border + ".border.".Length)] + "style", out var style) &&
                style is not ("nil" or "none"))
                Assert.Fail($"Unexpected custom style run {key}: {value}");
            if (key.EndsWith(".shading.fill", StringComparison.Ordinal) &&
                value != "AUTO" ||
                key.EndsWith(".shading.pattern", StringComparison.Ordinal) &&
                value is not ("clear" or "nil"))
                Assert.Fail($"Unexpected custom style run {key}: {value}");
        }
    }

    private static bool NearLegacyRgb(string expected, string observed)
    {
        if (expected.Length != 6 || observed.Length != 6 ||
            !int.TryParse(expected, System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture, out var left) ||
            !int.TryParse(observed, System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture, out var right))
            return false;
        return Enumerable.Range(0, 3).All(i =>
            Math.Abs(((left >> (i * 8)) & 0xff) -
                ((right >> (i * 8)) & 0xff)) <= 2);
    }

    private static Dictionary<string, string> ReadCustomStyleRunFormatting(
        byte[] bytes, bool literalColorsOnly = false)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var styles = document.MainDocumentPart?.StyleDefinitionsPart?.Styles?
            .Elements<Style>().Where(x => x.StyleId?.Value != null)
            .ToDictionary(x => x.StyleId!.Value!, StringComparer.Ordinal)
            ?? new Dictionary<string, Style>(StringComparer.Ordinal);
        var resolved = new Dictionary<string, Dictionary<string, string>>(
            StringComparer.Ordinal);
        var visiting = new HashSet<string>(StringComparer.Ordinal);
        Dictionary<string, string> Resolve(string id)
        {
            if (resolved.TryGetValue(id, out var cached)) return cached;
            if (!styles.TryGetValue(id, out var style))
                return new Dictionary<string, string>(StringComparer.Ordinal);
            Assert.True(visiting.Add(id), $"Cyclic based-on style {id}");
            var values = style.BasedOn?.Val?.Value is string parent
                ? new Dictionary<string, string>(Resolve(parent), StringComparer.Ordinal)
                : new Dictionary<string, string>(StringComparer.Ordinal);
            var properties = style.StyleRunProperties;
            if (properties?.GetFirstChild<RunFonts>() is { } fonts)
            {
                if (fonts.Ascii?.Value is string ascii) values["fontAscii"] = ascii;
                if (fonts.HighAnsi?.Value is string highAnsi)
                    values["fontHighAnsi"] = highAnsi;
            }
            if (properties?.GetFirstChild<Bold>() is { } bold)
                values["bold"] = (bold.Val?.Value ?? true).ToString();
            if (properties?.GetFirstChild<Italic>() is { } italic)
                values["italic"] = (italic.Val?.Value ?? true).ToString();
            if (properties?.GetFirstChild<ItalicComplexScript>() is { } complexItalic)
                values["complexItalic"] = (complexItalic.Val?.Value ?? true).ToString();
            if (properties?.GetFirstChild<FontSize>()?.Val?.Value is string size)
                values["size"] = size;
            if (properties?.GetFirstChild<FontSizeComplexScript>()?.Val?.Value is string complexSize)
                values["complexSize"] = complexSize;
            if (properties?.GetFirstChild<Spacing>()?.Val?.Value is int spacing)
                values["characterSpacing"] = spacing.ToString(
                    System.Globalization.CultureInfo.InvariantCulture);
            if (properties?.GetFirstChild<Kern>()?.Val?.Value is uint kern)
                values["kerning"] = kern.ToString(
                    System.Globalization.CultureInfo.InvariantCulture);
            if (properties?.GetFirstChild<CharacterScale>()?.Val?.Value is long scale)
                values["scale"] = scale.ToString(
                    System.Globalization.CultureInfo.InvariantCulture);
            if (properties?.GetFirstChild<Position>()?.Val?.Value is string position)
                values["position"] = position;
            if (properties?.GetFirstChild<FitText>()?.Val?.Value is uint fitWidth)
                values["fitTextWidth"] = fitWidth.ToString(
                    System.Globalization.CultureInfo.InvariantCulture);
            if (properties?.GetFirstChild<Languages>() is { } languages)
            {
                if (languages.Val?.Value is string language)
                    values["language"] = language;
                if (languages.EastAsia?.Value is string eastAsia)
                    values["language.eastAsia"] = eastAsia;
                if (languages.Bidi?.Value is string bidi)
                    values["language.bidi"] = bidi;
            }
            if (properties?.GetFirstChild<Strike>() is { } strike)
                values["strike"] = (strike.Val?.Value ?? true).ToString();
            if (properties?.GetFirstChild<DoubleStrike>() is { } doubleStrike)
                values["doubleStrike"] = (doubleStrike.Val?.Value ?? true).ToString();
            if (properties?.GetFirstChild<Outline>() is { } outline)
                values["outline"] = (outline.Val?.Value ?? true).ToString();
            if (properties?.GetFirstChild<Shadow>() is { } shadow)
                values["shadow"] = (shadow.Val?.Value ?? true).ToString();
            if (properties?.GetFirstChild<Emboss>() is { } emboss)
                values["emboss"] = (emboss.Val?.Value ?? true).ToString();
            if (properties?.GetFirstChild<Imprint>() is { } imprint)
                values["imprint"] = (imprint.Val?.Value ?? true).ToString();
            if (properties?.GetFirstChild<Caps>() is { } caps)
                values["caps"] = (caps.Val?.Value ?? true).ToString();
            if (properties?.GetFirstChild<SmallCaps>() is { } smallCaps)
                values["smallCaps"] = (smallCaps.Val?.Value ?? true).ToString();
            if (properties?.GetFirstChild<Underline>() is { } underline)
            {
                values["underline"] = underline.Val?.InnerText ?? "single";
                if (underline.Color?.Value is string underlineColor &&
                    !underlineColor.Equals("auto", StringComparison.OrdinalIgnoreCase))
                    values["underlineColor"] = underlineColor.ToUpperInvariant();
            }
            if (properties?.GetFirstChild<Border>() is { } border)
            {
                if (border.Val != null)
                    values["border.style"] = border.Val.InnerText;
                if (border.Color?.Value is string borderColor &&
                    !borderColor.Equals("auto", StringComparison.OrdinalIgnoreCase))
                    values["border.color"] = borderColor.ToUpperInvariant();
                if (border.Size?.Value is uint borderSize)
                    values["border.size"] = borderSize.ToString(
                        System.Globalization.CultureInfo.InvariantCulture);
                if (border.Space?.Value is uint borderSpace)
                    values["border.space"] = borderSpace.ToString(
                        System.Globalization.CultureInfo.InvariantCulture);
            }
            if (properties?.GetFirstChild<Shading>() is { } shading)
            {
                if (shading.Fill?.Value is string fill)
                    values["shading.fill"] = fill.ToUpperInvariant();
                if (shading.Val != null)
                    values["shading.pattern"] = shading.Val.InnerText;
            }
            if (properties?.GetFirstChild<Color>() is { } color)
            {
                if (literalColorsOnly && color.ThemeColor != null)
                    values.Remove("color");
                else if (color.Val?.Value is string rgb)
                    values["color"] = rgb.ToUpperInvariant();
            }
            visiting.Remove(id);
            resolved[id] = values;
            return values;
        }
        foreach (var style in styles.Values.Where(x =>
            x.CustomStyle?.Value == true))
        {
            var name = style.StyleName?.Val?.Value ?? style.StyleId!.Value!;
            foreach (var (property, value) in Resolve(style.StyleId!.Value!))
                result[$"{name}.{property}"] = value;
        }
        return result;
    }

    [Theory]
    [InlineData("WordInheritedBidiAllStories", "Visible Base", "bidi")]
    [InlineData("WordCharacterSpacingStyleOverridesAllStories", "Mirrored body", "mirrorIndents")]
    [InlineData("WordOutlineStylesAllStories", "Outline pair base", "outlineLevel")]
    public void CustomParagraphStyleDirectionAndOutlineArePartOfThePairedGate(
        string fixture, string styleName, string property)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            fixture + ".docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    styleName);
            var paragraph = style.StyleParagraphProperties!;
            switch (property)
            {
                case "bidi":
                    var bidi = paragraph.GetFirstChild<BiDi>()!;
                    bidi.Val = !(bidi.Val?.Value ?? true);
                    break;
                case "mirrorIndents":
                    var mirror = paragraph.GetFirstChild<MirrorIndents>()!;
                    mirror.Val = !(mirror.Val?.Value ?? true);
                    break;
                default:
                    paragraph.GetFirstChild<OutlineLevel>()!.Val = 2;
                    break;
            }
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertCustomStyleParagraphFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(styleName + "." + property, failure.Message);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CustomStyleTabLeaderAndInheritedClearArePartOfThePairedGate(
        bool removeClear)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordClearedInheritedTabsAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var styleName = removeClear ? "Clear Child" : "Clear Root";
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    styleName);
            var tabs = style.StyleParagraphProperties!.GetFirstChild<Tabs>()!;
            if (removeClear)
                tabs.Elements<TabStop>().Single(x => x.Val?.Value ==
                    TabStopValues.Clear).Remove();
            else tabs.Elements<TabStop>().Single().Leader =
                TabStopLeaderCharValues.Hyphen;
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertCustomStyleParagraphFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(removeClear ? "Clear Child.tab.5600" :
            "Clear Root.tab.5600", failure.Message);
    }
    [Fact]
    public void CustomParagraphStyleNumberingIsPartOfThePairedGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordMixedSectionStyledStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    "Section Story Accent");
            style.StyleParagraphProperties!.GetFirstChild<NumberingProperties>()!
                .Remove();
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertCustomStyleParagraphFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("Section Story Accent.styleNumbering", failure.Message);
    }
    [Theory]
    [InlineData("WordStyledEmptySpacingAllStories", "Contextual base", "before")]
    [InlineData("WordDistributedStyleInheritanceAllStories", "Decorated Note",
        "alignment")]
    [InlineData("WordSixSlotStyledContent", "Six Story Base", "keepNext")]
    [InlineData("WordInheritedEmptyMarkAllStories", "Mark base", "lineRule")]
    [InlineData("WordInheritedEmptyMarkAllStories", "Mark base", "hangingIndent")]
    [InlineData("WordContextualSpacingInheritedAllStories", "Contextual base", "contextualSpacing")]
    [InlineData("WordContextualSpacingInheritedAllStories", "Contextual body", "contextualSpacing")]
    [InlineData("WordInheritedPaginationStyleStories", "Pagination base", "widowControl")]
    [InlineData("WordInheritedPaginationStyleStories", "Pagination child", "pageBreakBefore")]
    [InlineData("WordInheritedDisplayStyleStories", "DerivedDisplay", "leftIndent")]
    public void CustomParagraphStyleLayoutIsPartOfThePairedGate(
        string fixture, string styleName, string property)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            fixture + ".docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var styles = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            var style = styles.Elements<Style>().Single(x =>
                x.StyleName?.Val?.Value == styleName);
            var paragraph = style.StyleParagraphProperties!;
            switch (property)
            {
                case "before":
                    paragraph.GetFirstChild<SpacingBetweenLines>()!.Before = "360";
                    break;
                case "alignment":
                    paragraph.Justification!.Val = JustificationValues.Center;
                    break;
                case "lineRule":
                    paragraph.GetFirstChild<SpacingBetweenLines>()!.LineRule =
                        LineSpacingRuleValues.Exact;
                    break;
                case "hangingIndent":
                    paragraph.GetFirstChild<Indentation>()!.Hanging = "240";
                    break;
                case "leftIndent":
                    paragraph.GetFirstChild<Indentation>()!.Left = "720";
                    break;
                case "contextualSpacing":
                    if (paragraph.ContextualSpacing is { } contextual)
                        contextual.Val = !(contextual.Val?.Value ?? true);
                    else paragraph.AppendChild(new ContextualSpacing { Val = false });
                    break;
                case "widowControl":
                    paragraph.WidowControl!.Val =
                        !(paragraph.WidowControl.Val?.Value ?? true);
                    break;
                case "pageBreakBefore":
                    paragraph.PageBreakBefore!.Val =
                        !(paragraph.PageBreakBefore.Val?.Value ?? true);
                    break;
                default:
                    paragraph.KeepNext!.Val = false;
                    break;
            }
            styles.Save();
        }
        var sourceValues = ReadCustomStyleParagraphFormatting(source);
        var changedValues = ReadCustomStyleParagraphFormatting(stream.ToArray());
        if (property == "alignment")
            Assert.NotEqual(sourceValues[styleName + ".alignment"],
                changedValues[styleName + ".alignment"]);
        var failure = Record.Exception(() =>
            AssertCustomStyleParagraphFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(styleName + "." + property, failure.Message);
    }

    [Theory]
    [InlineData("border.top.color")]
    [InlineData("shading.fill")]
    public void CustomParagraphStyleDecorationsArePartOfThePairedGate(
        string property)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyledParagraphDecorations.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var styles = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            var paragraph = styles.Elements<Style>().Single(x =>
                x.StyleName?.Val?.Value == "Decorated Note")
                .StyleParagraphProperties!;
            if (property == "border.top.color")
                paragraph.GetFirstChild<ParagraphBorders>()!.TopBorder!.Color =
                    "FF0000";
            else paragraph.GetFirstChild<Shading>()!.Fill = "00FF00";
            styles.Save();
        }
        var failure = Record.Exception(() =>
            AssertCustomStyleParagraphFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("Decorated Note." + property, failure.Message);
    }

    private static void AssertCustomStyleParagraphFormatting(byte[] source,
        byte[] output)
    {
        var expected = ReadCustomStyleParagraphFormatting(source);
        var observed = ReadCustomStyleParagraphFormatting(output);
        foreach (var (key, value) in expected)
            Assert.True(observed.TryGetValue(key, out var actual) && actual == value,
                $"Custom paragraph style {key}: expected {value}, observed " +
                (observed.TryGetValue(key, out var found) ? found : "<missing>"));
        foreach (var (key, value) in observed.Where(x =>
            x.Key.Contains(".tab.", StringComparison.Ordinal)))
            Assert.True(expected.TryGetValue(key, out var actual) && actual == value,
                $"Unexpected custom paragraph style {key}: {value}");
    }

    private static Dictionary<string, string> ReadCustomStyleParagraphFormatting(
        byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var styles = document.MainDocumentPart?.StyleDefinitionsPart?.Styles?
            .Elements<Style>().Where(x => x.StyleId?.Value != null)
            .ToDictionary(x => x.StyleId!.Value!, StringComparer.Ordinal)
            ?? new Dictionary<string, Style>(StringComparer.Ordinal);
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var resolved = new Dictionary<string, Dictionary<string, string>>(
            StringComparer.Ordinal);
        var visiting = new HashSet<string>(StringComparer.Ordinal);
        Dictionary<string, string> Resolve(string id)
        {
            if (resolved.TryGetValue(id, out var cached)) return cached;
            if (!styles.TryGetValue(id, out var style))
                return new Dictionary<string, string>(StringComparer.Ordinal);
            Assert.True(visiting.Add(id), $"Cyclic based-on style {id}");
            var values = style.BasedOn?.Val?.Value is string parent
                ? new Dictionary<string, string>(Resolve(parent), StringComparer.Ordinal)
                : new Dictionary<string, string>(StringComparer.Ordinal);
            var properties = style.StyleParagraphProperties;
            if (properties?.Justification?.Val?.InnerText is string alignment)
                values["alignment"] = alignment switch
                {
                    "start" => "left",
                    "end" => "right",
                    _ => alignment
                };
            var spacing = properties?.GetFirstChild<SpacingBetweenLines>();
            if (spacing?.Before?.Value is string before)
                values["before"] = before;
            if (spacing?.After?.Value is string after)
                values["after"] = after;
            if (spacing?.Line?.Value is string line)
                values["line"] = line;
            if (spacing?.LineRule?.InnerText is string rule)
                values["lineRule"] = rule;
            var indent = properties?.GetFirstChild<Indentation>();
            if ((indent?.Start?.Value ?? indent?.Left?.Value) is string left)
                values["leftIndent"] = left;
            if ((indent?.End?.Value ?? indent?.Right?.Value) is string right)
                values["rightIndent"] = right;
            if (indent?.FirstLine?.Value is string firstLine)
                values["firstLineIndent"] = firstLine;
            if (indent?.Hanging?.Value is string hanging)
                values["hangingIndent"] = hanging;
            if (properties?.KeepNext is { } keepNext)
                values["keepNext"] = (keepNext.Val?.Value ?? true).ToString();
            if (properties?.KeepLines is { } keepLines)
                values["keepLines"] = (keepLines.Val?.Value ?? true).ToString();
            if (properties?.ContextualSpacing is { } contextualSpacing)
                values["contextualSpacing"] =
                    (contextualSpacing.Val?.Value ?? true).ToString();
            if (properties?.WidowControl is { } widowControl)
                values["widowControl"] =
                    (widowControl.Val?.Value ?? true).ToString();
            if (properties?.PageBreakBefore is { } pageBreakBefore)
                values["pageBreakBefore"] =
                    (pageBreakBefore.Val?.Value ?? true).ToString();
            if (properties?.GetFirstChild<NumberingProperties>() is { } numbering &&
                numbering.NumberingId?.Val != null)
                values["styleNumbering"] = "true";            if (properties?.GetFirstChild<Tabs>() is { } tabs)
                foreach (var tab in tabs.Elements<TabStop>())
                {
                    var position = tab.Position?.Value;
                    if (position == null) continue;
                    var key = "tab." + position;
                    if (tab.Val?.Value == TabStopValues.Clear)
                        values.Remove(key);
                    else values[key] = (tab.Val?.InnerText ?? "left") + "|" +
                        (tab.Leader?.InnerText ?? "none");
                }            if (properties?.GetFirstChild<BiDi>() is { } bidi)
                values["bidi"] = (bidi.Val?.Value ?? true).ToString();
            if (properties?.GetFirstChild<MirrorIndents>() is { } mirror)
                values["mirrorIndents"] = (mirror.Val?.Value ?? true).ToString();
            if (properties?.GetFirstChild<OutlineLevel>()?.Val?.Value is int outline)
                values["outlineLevel"] = outline.ToString(
                    System.Globalization.CultureInfo.InvariantCulture);
            static string? Attribute(DocumentFormat.OpenXml.OpenXmlElement element,
                string name) => element.GetAttributes().FirstOrDefault(x =>
                    x.LocalName == name && x.NamespaceUri ==
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
                    .Value;
            void AddAttributes(DocumentFormat.OpenXml.OpenXmlElement? element,
                string prefix, params string[] names)
            {
                if (element == null) return;
                foreach (var name in names)
                    if (Attribute(element, name) is string value && value.Length != 0 &&
                        !(name == "color" && value.Equals("auto",
                            StringComparison.OrdinalIgnoreCase)))
                        values[$"{prefix}.{name}"] = name is "color" or "fill"
                            ? value.ToUpperInvariant() : value;
            }
            AddAttributes(properties?.GetFirstChild<Shading>(), "shading",
                "val", "color", "fill");
            var borders = properties?.GetFirstChild<ParagraphBorders>();
            if (borders != null)
                foreach (var edge in borders.ChildElements.Where(x =>
                    x.LocalName is "top" or "bottom" or "left" or "right"))
                    AddAttributes(edge, "border." + edge.LocalName,
                        "val", "sz", "space", "color");
            visiting.Remove(id);
            resolved[id] = values;
            return values;
        }
        foreach (var style in styles.Values.Where(x => x.Type?.Value ==
            StyleValues.Paragraph && x.CustomStyle?.Value == true))
        {
            var name = style.StyleName?.Val?.Value ?? style.StyleId!.Value!;
            foreach (var (property, value) in Resolve(style.StyleId!.Value!))
                result[$"{name}.{property}"] = value;
        }
        return result;
    }

    private static void AssertDirectEmptyParagraphMarkFormatting(byte[] source,
        byte[] output)
    {
        var expected = ReadDirectEmptyParagraphMarkFormatting(source);
        var observed = ReadDirectEmptyParagraphMarkFormatting(output);
        foreach (var (key, value) in expected)
            Assert.True(observed.TryGetValue(key, out var actual) && actual == value,
                $"Empty paragraph mark {key}: expected {value}, observed " +
                (observed.TryGetValue(key, out var found) ? found : "<missing>"));
    }

    private static Dictionary<string, string> ReadDirectEmptyParagraphMarkFormatting(
        byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        void AddStory(string key, DocumentFormat.OpenXml.OpenXmlElement story)
        {
            var paragraphs = story.Descendants<Paragraph>().ToArray();
            for (var i = 0; i < paragraphs.Length; i++)
            {
                if (paragraphs[i].Descendants<Text>().Any(x => x.Text.Length != 0))
                    continue;
                var mark = paragraphs[i].ParagraphProperties?
                    .GetFirstChild<ParagraphMarkRunProperties>();
                if (mark == null) continue;
                var prefix = $"{key}.paragraph{i}.";
                if (mark.GetFirstChild<Bold>() is { } bold)
                    result[prefix + "bold"] = (bold.Val?.Value ?? true).ToString();
                if (mark.GetFirstChild<Italic>() is { } italic)
                    result[prefix + "italic"] = (italic.Val?.Value ?? true).ToString();
                if (mark.GetFirstChild<FontSize>()?.Val?.Value is string size)
                    result[prefix + "size"] = size;
                if (mark.GetFirstChild<Color>()?.Val?.Value is string color)
                    result[prefix + "color"] = color.ToUpperInvariant();
                if (mark.GetFirstChild<Underline>() is { } underline)
                    result[prefix + "underline"] = underline.Val?.InnerText
                        ?? "single";
                if (mark.GetFirstChild<Strike>() is { } strike)
                    result[prefix + "strike"] = (strike.Val?.Value ?? true).ToString();
                if (mark.GetFirstChild<RunFonts>() is { } fonts)
                {
                    if (fonts.Ascii?.Value is string ascii)
                        result[prefix + "fontAscii"] = ascii;
                    if (fonts.HighAnsi?.Value is string highAnsi)
                        result[prefix + "fontHighAnsi"] = highAnsi;
                }
            }
        }
        AddStory("body", main.Document!.Body!);
        var sections = main.Document.Body.Descendants<SectionProperties>().ToArray();
        for (var i = 0; i < sections.Length; i++)
        {
            foreach (var reference in sections[i].Elements<HeaderReference>())
                AddStory($"section{i}.header.{Slot(reference.Type?.Value)}",
                    ((HeaderPart)main.GetPartById(reference.Id!)).Header!);
            foreach (var reference in sections[i].Elements<FooterReference>())
                AddStory($"section{i}.footer.{Slot(reference.Type?.Value)}",
                    ((FooterPart)main.GetPartById(reference.Id!)).Footer!);
        }
        return result;
    }

    private static Dictionary<string, string> ReadInteriorEmptyParagraphMarks(byte[] bytes,
        bool details = false)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        void AddStory(string key, DocumentFormat.OpenXml.OpenXmlElement story)
        {
            var paragraphs = story.Descendants<Paragraph>().ToArray();
            for (var i = 1; i + 1 < paragraphs.Length; i++)
            {
                static bool HasText(Paragraph paragraph) =>
                    paragraph.Descendants<Text>().Any(x => x.Text.Length != 0);
                if (HasText(paragraphs[i]) || !HasText(paragraphs[i - 1]) ||
                    !HasText(paragraphs[i + 1])) continue;
                var mark = paragraphs[i].ParagraphProperties?
                    .GetFirstChild<ParagraphMarkRunProperties>();
                var emphasis = mark?.GetFirstChild<Bold>() != null
                    ? "bold" : mark?.GetFirstChild<Italic>() != null
                    ? "italic" : "plain";
                result[$"{key}.paragraph{i}"] = details
                    ? $"{emphasis}|{mark?.GetFirstChild<FontSize>()?.Val?.Value}|" +
                        $"{mark?.GetFirstChild<Color>()?.Val?.Value}"
                    : emphasis;
            }
        }
        AddStory("body", main.Document!.Body!);
        var sections = main.Document.Body.Descendants<SectionProperties>().ToArray();
        for (var i = 0; i < sections.Length; i++)
        {
            foreach (var reference in sections[i].Elements<HeaderReference>())
                AddStory($"section{i}.header.{Slot(reference.Type?.Value)}",
                    ((HeaderPart)main.GetPartById(reference.Id!)).Header!);
            foreach (var reference in sections[i].Elements<FooterReference>())
                AddStory($"section{i}.footer.{Slot(reference.Type?.Value)}",
                    ((FooterPart)main.GetPartById(reference.Id!)).Footer!);
        }
        return result;
    }

    private static Dictionary<string, string> ReadListSemantics(byte[] bytes,
        bool includeLabelFormatting = false)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var numbering = main.NumberingDefinitionsPart?.Numbering;
        var instances = numbering?.Elements<NumberingInstance>()
            .Where(x => x.NumberID?.Value != null)
            .ToDictionary(x => x.NumberID!.Value) ??
            new Dictionary<int, NumberingInstance>();
        var definitions = numbering?.Elements<AbstractNum>()
            .Where(x => x.AbstractNumberId?.Value != null)
            .ToDictionary(x => x.AbstractNumberId!.Value) ??
            new Dictionary<int, AbstractNum>();
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        void AddStory(string key, DocumentFormat.OpenXml.OpenXmlElement story)
        {
            var paragraphs = story.Descendants<Paragraph>()
                .Where(x => x.Descendants<Text>().Any(t => t.Text.Length != 0))
                .ToArray();
            for (var i = 0; i < paragraphs.Length; i++)
            {
                var properties = paragraphs[i].ParagraphProperties?.NumberingProperties;
                if (properties?.NumberingId?.Val?.Value is not int numberId ||
                    numberId == 0) continue;
                Assert.True(instances.TryGetValue(numberId, out var instance),
                    $"Missing numbering instance {numberId} in {key}.");
                Assert.True(definitions.TryGetValue(instance!.AbstractNumId!.Val!.Value,
                    out var definition), $"Missing abstract list in {key}.");
                var levelIndex = properties.NumberingLevelReference?.Val?.Value ?? 0;
                var level = definition!.Elements<Level>().Single(x =>
                    x.LevelIndex?.Value == levelIndex);
                var overrideLevel = instance.Elements<LevelOverride>().FirstOrDefault(x =>
                    x.LevelIndex?.Value == levelIndex);
                var effective = overrideLevel?.Level ?? level;
                var indent = effective.PreviousParagraphProperties?
                    .GetFirstChild<Indentation>();
                var labelFont = effective.NumberingSymbolRunProperties?
                    .GetFirstChild<RunFonts>()?.Ascii?.Value ?? "";
                var labelProperties = effective.NumberingSymbolRunProperties;
                var labelFormatting = includeLabelFormatting
                    ? string.Join("|",
                        labelProperties?.GetFirstChild<Bold>() is { } bold &&
                            bold.Val?.Value != false,
                        labelProperties?.GetFirstChild<Italic>() is { } italic &&
                            italic.Val?.Value != false,
                        labelProperties?.GetFirstChild<Color>()?.Val?.Value ?? "",
                        labelProperties?.GetFirstChild<Underline>()?.Val?.Value.ToString() ?? "",
                        labelProperties?.GetFirstChild<FontSize>()?.Val?.Value ?? "",
                        labelProperties?.GetFirstChild<Strike>() is { } strike &&
                            strike.Val?.Value != false,
                        labelProperties?.GetFirstChild<SmallCaps>() is { } smallCaps &&
                            smallCaps.Val?.Value != false,
                        labelProperties?.GetFirstChild<Caps>() is { } caps &&
                            caps.Val?.Value != false,
                        labelProperties?.GetFirstChild<Underline>()?.Color?.Value ?? "")
                    : null;
                result[$"{key}.paragraph{i}"] = string.Join("|",
                    effective.NumberingFormat?.Val?.InnerText ?? "decimal",
                    effective.LevelText?.Val?.Value ?? "",
                    overrideLevel?.StartOverrideNumberingValue?.Val?.Value ??
                        effective.StartNumberingValue?.Val?.Value ?? 1,
                    effective.LevelSuffix?.Val?.InnerText ?? "tab",
                    indent?.Left?.Value?.ToString() ?? "",
                    indent?.Hanging?.Value?.ToString() ?? "",
                    labelFont) + (includeLabelFormatting ? "|" + labelFormatting : "");
            }
        }
        AddStory("body", main.Document!.Body!);
        var sections = main.Document.Body.Descendants<SectionProperties>().ToArray();
        for (var i = 0; i < sections.Length; i++)
        {
            foreach (var reference in sections[i].Elements<HeaderReference>())
                AddStory($"section{i}.header.{Slot(reference.Type?.Value)}",
                    ((HeaderPart)main.GetPartById(reference.Id!)).Header!);
            foreach (var reference in sections[i].Elements<FooterReference>())
                AddStory($"section{i}.footer.{Slot(reference.Type?.Value)}",
                    ((FooterPart)main.GetPartById(reference.Id!)).Footer!);
        }
        return result;
    }

    private static void AssertExplicitSectionLayout(byte[] reference, byte[] actual)
    {
        var expected = ReadSectionLayout(reference);
        var observed = ReadSectionLayout(actual);
        foreach (var (key, value) in expected)
            Assert.True(observed.TryGetValue(key, out var found) && found == value,
                $"Section layout {key}: expected {value}, observed " +
                (observed.TryGetValue(key, out var current) ? current : "<missing>"));
        Assert.Equal(expected.Where(x => x.Key.Contains(".pageBorders.",
                StringComparison.Ordinal)).OrderBy(x => x.Key),
            observed.Where(x => x.Key.Contains(".pageBorders.",
                StringComparison.Ordinal)).OrderBy(x => x.Key));
    }

    private static Dictionary<string, string> ReadSectionLayout(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var sections = main.Document!.Body!
            .Descendants<SectionProperties>().ToArray();
        var result = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["sectionCount"] = sections.Length.ToString(),
            ["evenAndOddHeaders"] = (main.DocumentSettingsPart?.Settings?
                .GetFirstChild<EvenAndOddHeaders>() is { } evenOdd &&
                (evenOdd.Val?.Value ?? true)).ToString()
        };
        for (var i = 0; i < sections.Length; i++)
        {
            var size = sections[i].GetFirstChild<PageSize>();
            var margin = sections[i].GetFirstChild<PageMargin>();
            var columns = sections[i].GetFirstChild<Columns>();
            var pageNumber = sections[i].GetFirstChild<PageNumberType>();
            void Add(string property, object? value)
            {
                if (value != null) result[$"section{i}.{property}"] = value.ToString()!;
            }
            Add("page.width", size?.Width?.Value);
            Add("page.height", size?.Height?.Value);
            Add("page.orientation", size?.Orient?.Value == PageOrientationValues.Landscape
                ? "landscape" : "portrait");
            Add("margin.top", margin?.Top?.Value);
            Add("margin.bottom", margin?.Bottom?.Value);
            Add("margin.left", margin?.Left?.Value);
            Add("margin.right", margin?.Right?.Value);
            Add("margin.header", margin?.Header?.Value);
            Add("margin.footer", margin?.Footer?.Value);
            Add("margin.gutter", margin?.Gutter?.Value);
            Add("columns.count", columns?.ColumnCount?.Value);
            Add("columns.space", columns?.Space?.Value);
            Add("columns.equalWidth", columns?.EqualWidth?.Value);
            Add("columns.separator", columns?.Separator?.Value);
            if (columns != null)
            {
                var definedColumns = columns.Elements<Column>().ToArray();
                foreach (var (column, columnIndex) in definedColumns
                    .Select((column, columnIndex) => (column, columnIndex)))
                {
                    Add($"columns.column{columnIndex}.width", column.Width?.Value);
                    if (columnIndex + 1 < definedColumns.Length)
                        Add($"columns.column{columnIndex}.space", column.Space?.Value);
                }
            }
            Add("pageNumber.start", pageNumber?.Start?.Value);
            Add("pageNumber.format", pageNumber?.Format?.Value ??
                NumberFormatValues.Decimal);
            Add("section.type", sections[i].GetFirstChild<SectionType>()?
                .Val?.Value ?? SectionMarkValues.NextPage);
            Add("verticalAlignment", sections[i]
                .GetFirstChild<VerticalTextAlignmentOnPage>()?.Val?.Value ??
                VerticalJustificationValues.Top);
            Add("bidirectional", (sections[i].GetFirstChild<BiDi>() is { } bidi &&
                (bidi.Val?.Value ?? true)));
            Add("gutterOnRight", (sections[i].GetFirstChild<GutterOnRight>() is
                { } rightGutter && (rightGutter.Val?.Value ?? true)));
            var borders = sections[i].GetFirstChild<PageBorders>();
            const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            if (borders != null)
            {
                Add("pageBorders.display", borders.GetAttribute("display", w).Value);
                Add("pageBorders.offsetFrom", borders.GetAttribute("offsetFrom", w).Value);
                Add("pageBorders.zOrder", borders.GetAttribute("zOrder", w).Value);
            }
            void AddBorder(string side, BorderType? border)
            {
                if (border != null)
                    Add($"pageBorders.{side}.val", border.GetAttribute("val", w).Value);
                Add($"pageBorders.{side}.color", border?.Color?.Value);
                Add($"pageBorders.{side}.size", border?.Size?.Value);
                Add($"pageBorders.{side}.space", border?.Space?.Value);
            }
            AddBorder("top", borders?.GetFirstChild<TopBorder>());
            AddBorder("left", borders?.GetFirstChild<LeftBorder>());
            AddBorder("bottom", borders?.GetFirstChild<BottomBorder>());
            AddBorder("right", borders?.GetFirstChild<RightBorder>());
            var titlePage = sections[i].GetFirstChild<TitlePage>();
            result[$"section{i}.titlePage"] =
                (titlePage != null && (titlePage.Val?.Value ?? true)).ToString();
        }
        return result;
    }

    private static void AssertParagraphDecorations(byte[] reference, byte[] actual)
    {
        var expected = ReadParagraphDecorations(reference);
        var observed = ReadParagraphDecorations(actual);
        foreach (var (key, value) in expected)
            Assert.True(observed.TryGetValue(key, out var found) && found == value,
                $"Paragraph decoration {key}: expected {value}, observed " +
                (observed.TryGetValue(key, out var current) ? current : "<missing>"));
    }

    private static Dictionary<string, string> ReadParagraphDecorations(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var styles = main.StyleDefinitionsPart?.Styles;
        var definitions = styles?.Elements<Style>()
            .Where(x => x.StyleId?.Value != null)
            .ToDictionary(x => x.StyleId!.Value!, StringComparer.Ordinal) ??
            new Dictionary<string, Style>(StringComparer.Ordinal);
        var defaultStyle = styles?.Elements<Style>()
            .FirstOrDefault(x => x.Type?.Value == StyleValues.Paragraph &&
                x.Default?.Value == true)?.StyleId?.Value;
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        void Apply(DocumentFormat.OpenXml.OpenXmlElement? properties, string prefix)
        {
            var shading = properties?.GetFirstChild<Shading>();
            var border = properties?.GetFirstChild<ParagraphBorders>()?
                .GetFirstChild<TopBorder>();
            if (shading != null)
            {
                result.Remove($"{prefix}.fill");
                result.Remove($"{prefix}.pattern");
                result.Remove($"{prefix}.none");
                var pattern = shading.Val?.InnerText;
                var fill = shading.Fill?.Value;
                if (pattern == "nil" ||
                    (pattern == "clear" &&
                        (fill == null || fill.Equals("auto",
                            StringComparison.OrdinalIgnoreCase))))
                    result[$"{prefix}.none"] = "true";
                else
                {
                    if (fill != null && !fill.Equals("auto",
                        StringComparison.OrdinalIgnoreCase))
                        result[$"{prefix}.fill"] = fill.ToUpperInvariant();
                    if (pattern != null) result[$"{prefix}.pattern"] = pattern;
                }
            }
            if (border?.Val?.Value is { })
                result[$"{prefix}.top.style"] = border.Val!.InnerText;
            if (border?.Size?.Value is uint size)
                result[$"{prefix}.top.size"] = size.ToString();
            if (border?.Space?.Value is uint space)
                result[$"{prefix}.top.space"] = space.ToString();
            if (border?.Color?.Value is string color)
                result[$"{prefix}.top.color"] = color.ToUpperInvariant();
        }
        void AddStory(string key, DocumentFormat.OpenXml.OpenXmlElement story)
        {
            var paragraphs = story.Descendants<Paragraph>()
                .Where(x => x.Descendants<Text>().Any(t => t.Text.Length != 0))
                .ToArray();
            for (var i = 0; i < paragraphs.Length; i++)
            {
                var prefix = $"{key}.paragraph{i}";
                var seen = new HashSet<string>(StringComparer.Ordinal);
                void ApplyStyle(string? id)
                {
                    if (id == null || !seen.Add(id) ||
                        !definitions.TryGetValue(id, out var style)) return;
                    ApplyStyle(style.BasedOn?.Val?.Value);
                    Apply(style.StyleParagraphProperties, prefix);
                }
                ApplyStyle(paragraphs[i].ParagraphProperties?
                    .ParagraphStyleId?.Val?.Value ?? defaultStyle);
                Apply(paragraphs[i].ParagraphProperties, prefix);
            }
        }
        AddStory("body", main.Document!.Body!);
        var sections = main.Document.Body.Descendants<SectionProperties>().ToArray();
        for (var i = 0; i < sections.Length; i++)
        {
            foreach (var reference in sections[i].Elements<HeaderReference>())
                AddStory($"section{i}.header.{Slot(reference.Type?.Value)}",
                    ((HeaderPart)main.GetPartById(reference.Id!)).Header!);
            foreach (var reference in sections[i].Elements<FooterReference>())
                AddStory($"section{i}.footer.{Slot(reference.Type?.Value)}",
                    ((FooterPart)main.GetPartById(reference.Id!)).Footer!);
        }
        return result;
    }

    [Theory]
    [InlineData("body")]
    [InlineData("header")]
    [InlineData("footer")]
    public void EmptyParagraphSpacingIsPartOfThePairedLayoutGate(string story)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordEmptyParagraphSpacingAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            var root = story switch
            {
                "body" => (DocumentFormat.OpenXml.OpenXmlElement)main.Document!.Body!,
                "header" => main.HeaderParts.Select(x => x.Header!)
                    .First(x => x.Descendants<Paragraph>().Any(p =>
                        !p.Descendants<Text>().Any(t => t.Text.Length > 0) &&
                        p.ParagraphProperties?.SpacingBetweenLines != null)),
                _ => main.FooterParts.Select(x => x.Footer!)
                    .First(x => x.Descendants<Paragraph>().Any(p =>
                        !p.Descendants<Text>().Any(t => t.Text.Length > 0) &&
                        p.ParagraphProperties?.SpacingBetweenLines != null))
            };
            var empty = root.Descendants<Paragraph>().First(p =>
                !p.Descendants<Text>().Any(t => t.Text.Length > 0) &&
                p.ParagraphProperties?.SpacingBetweenLines != null);
            empty.ParagraphProperties!.SpacingBetweenLines!.Before = "0";
            switch (root)
            {
                case Body: main.Document!.Save(); break;
                case Header header: header.Save(); break;
                case Footer footer: footer.Save(); break;
            }
        }
        var failure = Record.Exception(() =>
            AssertEffectiveParagraphLayout(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(story, failure.Message);
        Assert.Contains(".before", failure.Message);
    }

    [Fact]
    public void ImplicitCjkSpacingDefaultIsPartOfThePairedLayoutGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordThreeSectionStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var paragraph = document.MainDocumentPart!.Document!.Body!
                .Elements<Paragraph>().First();
            var properties = paragraph.ParagraphProperties ??
                paragraph.PrependChild(new ParagraphProperties());
            properties.AppendChild(new AutoSpaceDE { Val = false });
            document.MainDocumentPart.Document.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveParagraphLayout(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("body.paragraph0.autoSpaceDE", failure.Message);
    }

    [Theory]
    [InlineData("WordInheritedCjkBreakAllStories", "CJK break base", "kinsoku")]
    [InlineData("WordInheritedCjkBreakAllStories", "CJK break base", "wordWrap")]
    [InlineData("WordInheritedCjkSpacingAllStories", "CJK spacing base", "autoSpaceDE")]
    [InlineData("WordInheritedCjkSpacingAllStories", "CJK spacing base", "autoSpaceDN")]
    public void InheritedCjkParagraphFlagsArePartOfThePairedLayoutGate(
        string fixture, string styleName, string property)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            fixture + ".docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value == styleName);
            var paragraph = style.StyleParagraphProperties!;
            switch (property)
            {
                case "kinsoku": paragraph.Kinsoku!.Val = true; break;
                case "wordWrap": paragraph.WordWrap!.Val = true; break;
                case "autoSpaceDE": paragraph.AutoSpaceDE!.Val = true; break;
                default: paragraph.AutoSpaceDN!.Val = true; break;
            }
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveParagraphLayout(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("." + property, failure.Message);
    }

    [Fact]
    public void InheritedOutlineLevelIsPartOfThePairedLayoutGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedOutlineAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    "Outline pair base");
            style.StyleParagraphProperties!.GetFirstChild<OutlineLevel>()!.Val = 2;
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveParagraphLayout(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(".outlineLevel", failure.Message);
    }

    [Fact]
    public void DirectGridSnapIsPartOfThePairedLayoutGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordSectionLineGridSnapAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var paragraph = document.MainDocumentPart!.Document.Body!
                .Elements<Paragraph>().First(x => x.ParagraphProperties?
                    .GetFirstChild<SnapToGrid>()?.Val?.Value == false);
            paragraph.ParagraphProperties!.GetFirstChild<SnapToGrid>()!.Val = true;
            document.MainDocumentPart.Document.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveParagraphLayout(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(".snapToGrid", failure.Message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ConditionalCornerParagraphAlignmentSurvivesBothDocRoutes(bool native)
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordConditionalCornerParagraphAllStories"));
        var source = File.ReadAllBytes(stem + ".docx");
        var bytes = native ? File.ReadAllBytes(stem + ".doc") :
            DxpDocExport.Export(source);
        foreach (var binary in new[] { bytes,
            DxpDocExport.Export(DxpDocToDocx.Project(bytes).DocxBytes) })
        {
            using var index = new DocTextIndexWalker().Index(new MemoryStream(binary));
            if (!native)
            {
                var style = Assert.Single(index.StyleDefinitions,
                    x => x.Name == "Logical Start");
                var rules = style.ConditionalParagraphFormatting!;
                foreach (var condition in new ushort[]
                    { 0x0004, 0x0008, 0x0100, 0x0200, 0x0400, 0x0800 })
                    Assert.True(rules.ContainsKey(condition));
            }
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            AssertEffectiveParagraphLayout(source, projected);
            Assert.Equal(ReadStories(source), ReadStories(projected));
            Assert.Empty(Validate(projected));
        }
    }

    [Theory]
    [InlineData("NorthWestCell")]
    [InlineData("NorthEastCell")]
    [InlineData("SouthWestCell")]
    [InlineData("SouthEastCell")]
    public void CornerParagraphAlignmentIsPartOfPairedLayoutGate(
        string cornerName)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordConditionalCornerParagraphAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    "Logical Start");
            var condition = cornerName switch
            {
                "NorthWestCell" => TableStyleOverrideValues.NorthWestCell,
                "NorthEastCell" => TableStyleOverrideValues.NorthEastCell,
                "SouthWestCell" => TableStyleOverrideValues.SouthWestCell,
                "SouthEastCell" => TableStyleOverrideValues.SouthEastCell,
                _ => throw new ArgumentOutOfRangeException(nameof(cornerName))
            };
            var corner = style.Elements<TableStyleProperties>().Single(x =>
                x.Type?.Value == condition);
            corner.Descendants<Justification>().Single().Val =
                condition == TableStyleOverrideValues.SouthEastCell
                    ? JustificationValues.Left : JustificationValues.Both;
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveParagraphLayout(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(".alignment", failure.Message);
    }

    [Theory]
    [InlineData("0040")]
    [InlineData("0020")]
    public void MaskOnlyTableLookControlsEffectiveParagraphAlignment(string mask)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordMaskedTableLookParagraphAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            var tables = main.Document.Body!.Descendants<Table>()
                .Concat(main.HeaderParts.SelectMany(x => x.Header!.Descendants<Table>()))
                .Concat(main.FooterParts.SelectMany(x => x.Footer!.Descendants<Table>()))
                .ToArray();
            Assert.Equal(3, tables.Length);
            Assert.All(tables, table =>
            {
                var look = table.TableProperties!.GetFirstChild<TableLook>()!;
                Assert.Null(look.FirstRow);
                Assert.Null(look.LastRow);
                Assert.Equal("0060", look.Val?.Value);
                look.Val = mask;
            });
            main.Document.Save();
            foreach (var header in main.HeaderParts) header.Header!.Save();
            foreach (var footer in main.FooterParts) footer.Footer!.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveParagraphLayout(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(".alignment", failure.Message);
    }

    [Fact]
    public void MaskOnlyTableLookControlsEffectiveCornerRunFormatting()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordMaskedTableLookRunStyleAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            var tables = main.Document.Body!.Descendants<Table>()
                .Concat(main.HeaderParts.SelectMany(x => x.Header!.Descendants<Table>()))
                .Concat(main.FooterParts.SelectMany(x => x.Footer!.Descendants<Table>()))
                .ToArray();
            Assert.Equal(3, tables.Length);
            Assert.All(tables, table =>
            {
                var look = table.TableProperties!.GetFirstChild<TableLook>()!;
                Assert.Null(look.FirstRow);
                Assert.Null(look.LastRow);
                Assert.Equal("01E0", look.Val?.Value);
                look.Val = "01C0";
            });
            main.Document.Save();
            foreach (var header in main.HeaderParts) header.Header!.Save();
            foreach (var footer in main.FooterParts) footer.Footer!.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(".color", failure.Message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AllConditionalCornerRunStylesSurviveBothDocRoutes(bool native)
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordConditionalAllCornerRunStyleAllStories"));
        var source = File.ReadAllBytes(stem + ".docx");
        var expected = ReadEffectiveRunFormatting(source);
        foreach (var color in new[] { "CC00CC", "FF8800", "0000CC", "660099" })
            Assert.Contains(expected.Values, value => value == color);
        var bytes = native ? File.ReadAllBytes(stem + ".doc") :
            DxpDocExport.Export(source);
        foreach (var binary in new[] { bytes,
            DxpDocExport.Export(DxpDocToDocx.Project(bytes).DocxBytes) })
        {
            using var index = new DocTextIndexWalker().Index(new MemoryStream(binary));
            if (!native)
            {
                var style = Assert.Single(index.StyleDefinitions,
                    x => x.Name == "Logical Start");
                var colors = style.ConditionalCharacterFormatting!;
                Assert.Equal(0x00CC00CCu, colors[0x0200].ColorRef);
                Assert.Equal(0x000088FFu, colors[0x0100].ColorRef);
                Assert.Equal(0x00CC0000u, colors[0x0800].ColorRef);
                Assert.Equal(0x00990066u, colors[0x0400].ColorRef);
            }
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            AssertEffectiveRunFormatting(source, projected);
            Assert.Equal(ReadStories(source), ReadStories(projected));
            Assert.Empty(Validate(projected));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LowerCornerRunColorIsPartOfThePairedFormattingGate(bool southEast)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordConditionalAllCornerRunStyleAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    "Logical Start");
            var corner = style.Elements<TableStyleProperties>().Single(x =>
                x.Type?.Value == (southEast ? TableStyleOverrideValues.SouthEastCell :
                    TableStyleOverrideValues.SouthWestCell));
            corner.Descendants<Color>().Single().Val = "123456";
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(".color", failure.Message);
    }

    [Theory]
    [InlineData("WordLayeredCornerParagraphAllStories", false,
        "both")]
    [InlineData("WordLayeredCornerParagraphAllStories", true,
        "both")]
    [InlineData("WordLayeredCornerCenterParagraphAllStories", false,
        "center")]
    [InlineData("WordLayeredCornerCenterParagraphAllStories", true,
        "center")]
    public void CellConditionalAnnotationControlsDirectLeftAlignment(
        string fixtureName, bool removeAnnotations,
        string conditionalAlignment)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            fixtureName + ".docx"));
        var source = File.ReadAllBytes(path);
        if (removeAnnotations)
        {
            using var editable = new MemoryStream();
            editable.Write(source);
            using (var package = WordprocessingDocument.Open(editable, true))
            {
                var main = package.MainDocumentPart!;
                foreach (var cell in main.Document.Body!.Descendants<TableCell>()
                    .Concat(main.HeaderParts.SelectMany(x =>
                        x.Header!.Descendants<TableCell>()))
                    .Concat(main.FooterParts.SelectMany(x =>
                        x.Footer!.Descendants<TableCell>())))
                    foreach (var annotation in cell.TableCellProperties!
                        .ChildElements.Where(x => x.LocalName == "cnfStyle").ToArray())
                        annotation.Remove();
                main.Document.Save();
                foreach (var header in main.HeaderParts) header.Header!.Save();
                foreach (var footer in main.FooterParts) footer.Footer!.Save();
            }
            source = editable.ToArray();
            var original = File.ReadAllBytes(path);
            var failure = Record.Exception(() =>
                AssertEffectiveParagraphLayout(original, source));
            Assert.NotNull(failure);
            Assert.Contains(".alignment", failure.Message);
        }
        var projected = DxpDocToDocx.Project(DxpDocExport.Export(source)).DocxBytes;
        using var output = WordprocessingDocument.Open(
            new MemoryStream(projected), false);
        var mainOutput = output.MainDocumentPart!;
        var tables = mainOutput.Document.Body!.Descendants<Table>()
            .Concat(mainOutput.HeaderParts.SelectMany(x =>
                x.Header!.Descendants<Table>()))
            .Concat(mainOutput.FooterParts.SelectMany(x =>
                x.Footer!.Descendants<Table>())).ToArray();
        Assert.Equal(3, tables.Length);
        Assert.All(tables, table =>
        {
            var lastCell = table.Elements<TableRow>().Last()
                .Elements<TableCell>().Last();
            Assert.Equal(removeAnnotations
                    ? (conditionalAlignment == "center" ?
                        JustificationValues.Center : JustificationValues.Both) :
                        JustificationValues.Left,
                lastCell.Descendants<Paragraph>().Single().ParagraphProperties!
                    .GetFirstChild<Justification>()?.Val?.Value);
            Assert.Equal(removeAnnotations ? null : "000100000010",
                lastCell.TableCellProperties?
                    .GetFirstChild<ConditionalFormatStyle>()?.Val?.Value);
        });
        Assert.Empty(Validate(projected));
    }

    [Fact]
    public void CellSpacingZeroResetSurvivesBothDocRoutesAndThirdHop()
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordCellSpacingResetAllStories"));
        var source = File.ReadAllBytes(stem + ".docx");
        var sourceGrid = ReadTableGrid(source);
        Assert.Equal("120", sourceGrid["body.table0.row0.cellSpacing"]);
        Assert.Equal("0", sourceGrid["body.table0.row1.cellSpacing"]);
        foreach (var binary in new[] { File.ReadAllBytes(stem + ".doc"),
            DxpDocExport.Export(source) })
        {
            foreach (var doc in new[] { binary,
                DxpDocExport.Export(DxpDocToDocx.Project(binary).DocxBytes) })
            {
                var projected = DxpDocToDocx.Project(doc).DocxBytes;
                AssertTableGrid(source, projected);
                Assert.Equal("0", ReadTableGrid(projected)[
                    "body.table0.row1.cellSpacing"]);
                Assert.Empty(Validate(projected));
            }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StyledCellSpacingAndZeroRowOverrideStayEditable(bool native)
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyledCellSpacingResetAllStories"));
        var source = File.ReadAllBytes(stem + ".docx");
        var binary = native ? File.ReadAllBytes(stem + ".doc") :
            DxpDocExport.Export(source);
        foreach (var doc in new[] { binary,
            DxpDocExport.Export(DxpDocToDocx.Project(binary).DocxBytes) })
        {
            using var index = new DocTextIndexWalker().Index(new MemoryStream(doc));
            var style = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Logical Start");
            Assert.Equal((ushort)120, style.TableFormatting?.TableCellSpacingTwips);
            Assert.Contains(index.ParagraphStyles, x =>
                x.Formatting?.TableTerminator == true &&
                x.Formatting.TableCellSpacingTwips == 0);
            var projected = DxpDocToDocx.Project(doc).DocxBytes;
            AssertTableGrid(source, projected);
            Assert.Empty(Validate(projected));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StyledCellSpacingLayerMutationsFailTableGate(bool rowOverride)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordStyledCellSpacingResetAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var editable = new MemoryStream();
        editable.Write(source);
        using (var package = WordprocessingDocument.Open(editable, true))
        {
            var main = package.MainDocumentPart!;
            if (rowOverride)
            {
                var row = main.Document.Body!.Descendants<Table>().First()
                    .Elements<TableRow>().Last();
                row.TableRowProperties!.GetFirstChild<TableCellSpacing>()!
                    .Width = "120";
                row.TablePropertyExceptions!.GetFirstChild<TableCellSpacing>()!
                    .Width = "120";
                main.Document.Save();
            }
            else
            {
                var style = main.StyleDefinitionsPart!.Styles!.Elements<Style>()
                    .Single(x => x.StyleName?.Val?.Value == "Logical Start");
                style.StyleTableProperties!.GetFirstChild<TableCellSpacing>()!
                    .Width = "240";
                main.StyleDefinitionsPart.Styles.Save();
            }
        }
        var failure = Record.Exception(() => AssertTableGrid(source, editable.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("cellSpacing", failure.Message);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("header")]
    [InlineData("footer")]
    public void FixedLayoutMutationFailsAutoFitTableGate(string story)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures",
            "DocKnownGaps", "WordVisibleNoWrapArialAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var editable = new MemoryStream();
        editable.Write(source);
        using (var package = WordprocessingDocument.Open(editable, true))
        {
            var main = package.MainDocumentPart!;
            var section = main.Document.Body!.Descendants<SectionProperties>().First();
            var part = story switch
            {
                "header" => main.GetPartById(section.Elements<HeaderReference>()
                    .First(x => x.Type?.Value == HeaderFooterValues.Default).Id!),
                "footer" => main.GetPartById(section.Elements<FooterReference>()
                    .First(x => x.Type?.Value == HeaderFooterValues.Default).Id!),
                _ => (OpenXmlPart)main
            };
            var table = part.RootElement!.Descendants<Table>().First();
            var layout = table.TableProperties!.GetFirstChild<TableLayout>() ??
                table.TableProperties.AppendChild(new TableLayout());
            layout.Type = TableLayoutValues.Fixed;
            part.RootElement.Save();
        }
        var failure = Record.Exception(() => AssertTableGrid(source, editable.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(".fixedLayout", failure.Message);
    }

    [Theory]
    [InlineData("body", false)]
    [InlineData("header", false)]
    [InlineData("footer", false)]
    [InlineData("body", true)]
    [InlineData("header", true)]
    [InlineData("footer", true)]
    public void FixedPreferredWidthMutationFailsAutoWidthGate(string story, bool cellWidth)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures",
            "DocKnownGaps", "WordVisibleNoWrapArialAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var editable = new MemoryStream();
        editable.Write(source);
        using (var package = WordprocessingDocument.Open(editable, true))
        {
            var main = package.MainDocumentPart!;
            var section = main.Document.Body!.Descendants<SectionProperties>().First();
            var part = story switch
            {
                "header" => main.GetPartById(section.Elements<HeaderReference>()
                    .First(x => x.Type?.Value == HeaderFooterValues.Default).Id!),
                "footer" => main.GetPartById(section.Elements<FooterReference>()
                    .First(x => x.Type?.Value == HeaderFooterValues.Default).Id!),
                _ => (OpenXmlPart)main
            };
            var table = part.RootElement!.Descendants<Table>().First();
            TableWidthType width = cellWidth
                ? table.Descendants<TableCell>().First().TableCellProperties!
                    .GetFirstChild<TableCellWidth>()!
                : table.TableProperties!.GetFirstChild<TableWidth>()!;
            width.Type = TableWidthUnitValues.Dxa;
            width.Width = cellWidth ? "4680" : "9360";
            part.RootElement.Save();
        }
        var failure = Record.Exception(() => AssertTableGrid(source, editable.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(".preferredWidth", failure.Message);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("header")]
    [InlineData("footer")]
    public void UnexpectedNoWrapFailsTableGate(string story)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleNoWrapOffAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var editable = new MemoryStream();
        editable.Write(source);
        using (var package = WordprocessingDocument.Open(editable, true))
        {
            var main = package.MainDocumentPart!;
            var section = main.Document.Body!.Descendants<SectionProperties>().First();
            var part = story switch
            {
                "header" => main.GetPartById(section.Elements<HeaderReference>()
                    .First(x => x.Type?.Value == HeaderFooterValues.Default).Id!),
                "footer" => main.GetPartById(section.Elements<FooterReference>()
                    .First(x => x.Type?.Value == HeaderFooterValues.Default).Id!),
                _ => (OpenXmlPart)main
            };
            var cell = part.RootElement!.Descendants<Table>().First()
                .Descendants<TableCell>().First();
            (cell.TableCellProperties ?? cell.PrependChild(new TableCellProperties()))
                .Append(new NoWrap());
            part.RootElement.Save();
        }
        var failure = Record.Exception(() => AssertTableGrid(source, editable.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(".noWrap", failure.Message);
    }

    [Fact]
    public void UnexpectedCellFitTextFailsTableGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleNoWrapOffAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var editable = new MemoryStream();
        editable.Write(source);
        using (var package = WordprocessingDocument.Open(editable, true))
        {
            var main = package.MainDocumentPart!;
            var cell = main.Document.Body!.Descendants<TableCell>().First();
            (cell.TableCellProperties ?? cell.PrependChild(new TableCellProperties()))
                .Append(new TableCellFitText());
            main.Document.Save();
        }
        var failure = Record.Exception(() => AssertTableGrid(source, editable.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(".fitText", failure.Message);
    }

    [Fact]
    public void CellSpacingZeroResetMutationFailsTableGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordCellSpacingResetAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var editable = new MemoryStream();
        editable.Write(source);
        using (var package = WordprocessingDocument.Open(editable, true))
        {
            var row = package.MainDocumentPart!.Document.Body!
                .Descendants<Table>().First().Elements<TableRow>().Last();
            row.TableRowProperties!.GetFirstChild<TableCellSpacing>()!
                .Width = "120";
            row.TablePropertyExceptions!.GetFirstChild<TableCellSpacing>()!
                .Width = "120";
            package.MainDocumentPart.Document.Save();
        }
        var failure = Record.Exception(() => AssertTableGrid(source, editable.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("cellSpacing", failure.Message);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("header")]
    [InlineData("footer")]
    [InlineData("row")]
    public void CellSpacingChangesFailTableGate(string location)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordCellSpacingAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var editable = new MemoryStream();
        editable.Write(source);
        using (var package = WordprocessingDocument.Open(editable, true))
        {
            var main = package.MainDocumentPart!;
            var root = location switch
            {
                "body" or "row" => (OpenXmlElement)main.Document.Body!,
                "header" => main.HeaderParts.Single(x => x.Header!
                    .Descendants<Table>().Any()).Header!,
                _ => main.FooterParts.Single(x => x.Footer!
                    .Descendants<Table>().Any()).Footer!
            };
            var table = root.Descendants<Table>().First();
            if (location == "row")
            {
                var row = table.Elements<TableRow>().Last();
                row.TableRowProperties?.GetFirstChild<TableCellSpacing>()?.Remove();
                row.TablePropertyExceptions?.GetFirstChild<TableCellSpacing>()?.Remove();
            }
            else
            {
                table.TableProperties!.GetFirstChild<TableCellSpacing>()!.Remove();
                var firstRow = table.Elements<TableRow>().First();
                firstRow.TableRowProperties?.GetFirstChild<TableCellSpacing>()?.Remove();
                firstRow.TablePropertyExceptions?
                    .GetFirstChild<TableCellSpacing>()?.Remove();
            }
            if (location is "body" or "row") main.Document.Save();
            else if (location == "header") ((Header)root).Save();
            else ((Footer)root).Save();
        }
        var failure = Record.Exception(() => AssertTableGrid(source, editable.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("cellSpacing", failure.Message);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("header")]
    [InlineData("footer")]
    public void RotatedCellDirectionChangesFailTableGate(string story)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordRotatedTableCellsAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var editable = new MemoryStream();
        editable.Write(source);
        using (var package = WordprocessingDocument.Open(editable, true))
        {
            var main = package.MainDocumentPart!;
            var root = story switch
            {
                "body" => (OpenXmlElement)main.Document.Body!,
                "header" => main.HeaderParts.Single(x => x.Header!
                    .Descendants<Table>().Any()).Header!,
                _ => main.FooterParts.Single(x => x.Footer!
                    .Descendants<Table>().Any()).Footer!
            };
            root.Descendants<Table>().First().Descendants<TableCell>().First()
                .TableCellProperties!.GetFirstChild<TextDirection>()!.Remove();
            if (story == "body") main.Document.Save();
            else if (story == "header") ((Header)root).Save();
            else ((Footer)root).Save();
        }
        var failure = Record.Exception(() => AssertTableGrid(source, editable.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("textDirection", failure.Message);
    }

    [Fact]
    public void WordRotatedTableCellsRetainDirectionInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, "WordRotatedTableCellsAllStories");
        var source = File.ReadAllBytes(stem + ".docx");
        using (var sourcePackage = WordprocessingDocument.Open(
            new MemoryStream(source), false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2010).Validate(sourcePackage));
        var native = File.ReadAllBytes(stem + ".doc");
        var generated = DxpDocExport.Export(source);
        foreach (var binary in new[] { native, generated,
            DxpDocExport.Export(DxpDocToDocx.Project(native).DocxBytes) })
        {
            using var input = new MemoryStream(binary);
            using var index = new DocTextIndexWalker().Index(input);
            Assert.Equal(3, index.ParagraphStyles.Count(run =>
                run.Formatting?.TableTerminator == true &&
                run.Formatting.TableCellTextFlows?.FirstOrDefault() == 1));
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            using var document = WordprocessingDocument.Open(
                new MemoryStream(projected), false);
            var main = document.MainDocumentPart!;
            foreach (var root in new OpenXmlElement[] { main.Document.Body!,
                main.HeaderParts.Single(x => x.Header!.Descendants<Table>().Any()).Header!,
                main.FooterParts.Single(x => x.Footer!.Descendants<Table>().Any()).Footer! })
                Assert.Equal(TextDirectionValues.TopToBottomRightToLeft,
                    root.Descendants<Table>().First().Descendants<TableCell>().First()
                        .TableCellProperties!.GetFirstChild<TextDirection>()!.Val!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
    }

    [Fact]
    public void WordMixedRotatedCellsRetainFourDirectionsInAllStories()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc"));
        var stem = Path.Combine(directory, "WordMixedRotatedCellsAllStories");
        var source = File.ReadAllBytes(stem + ".docx");
        using (var sourcePackage = WordprocessingDocument.Open(
            new MemoryStream(source), false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2010).Validate(sourcePackage));
        var native = File.ReadAllBytes(stem + ".doc");
        var generated = DxpDocExport.Export(source);
        var expected = new[] { "1,3", "4,5", "3,5", "4,1" };
        foreach (var binary in new[] { native, generated,
            DxpDocExport.Export(DxpDocToDocx.Project(native).DocxBytes) })
        {
            using var input = new MemoryStream(binary);
            using var index = new DocTextIndexWalker().Index(input);
            var flows = index.ParagraphStyles.Where(run =>
                run.Formatting?.TableTerminator == true &&
                run.Formatting.TableCellTextFlows != null)
                .Select(run => string.Join(",", run.Formatting!.TableCellTextFlows!))
                .ToArray();
            Assert.Equal(expected.OrderBy(x => x), flows.OrderBy(x => x));
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            using var document = WordprocessingDocument.Open(
                new MemoryStream(projected), false);
            var main = document.MainDocumentPart!;
            static string Directions(OpenXmlElement root) => string.Join(",",
                root.Descendants<Table>().First().Descendants<TableCell>()
                    .Select(cell => cell.TableCellProperties!
                        .GetFirstChild<TextDirection>()!.Val!.InnerText));
            Assert.Equal("tbRl,btLr,lrTbV,tbRlV", Directions(main.Document.Body!));
            Assert.Equal("btLr,tbRlV", Directions(main.HeaderParts.Single(x =>
                x.Header!.Descendants<Table>().Any()).Header!));
            Assert.Equal("lrTbV,tbRl", Directions(main.FooterParts.Single(x =>
                x.Footer!.Descendants<Table>().Any()).Footer!));
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
    }

    [Theory]
    [InlineData("body")]
    [InlineData("header")]
    [InlineData("footer")]
    public void UnexpectedRotatedCellDirectionFailsTableGate(string story)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordTablesAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var editable = new MemoryStream();
        editable.Write(source);
        using (var package = WordprocessingDocument.Open(editable, true))
        {
            var main = package.MainDocumentPart!;
            var root = story switch
            {
                "body" => (OpenXmlElement)main.Document.Body!,
                "header" => main.HeaderParts.Single(x => x.Header!
                    .Descendants<Table>().Any()).Header!,
                _ => main.FooterParts.Single(x => x.Footer!
                    .Descendants<Table>().Any()).Footer!
            };
            root.Descendants<Table>().First().Descendants<TableCell>().First()
                .TableCellProperties!.AppendChild(new TextDirection
                { Val = TextDirectionValues.TopToBottomRightToLeft });
            if (story == "body") main.Document.Save();
            else if (story == "header") ((Header)root).Save();
            else ((Footer)root).Save();
        }
        var failure = Record.Exception(() => AssertTableGrid(source, editable.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("textDirection", failure.Message);
    }

    [Theory]
    [InlineData("WordLayeredCornerParagraphAllStories", false)]
    [InlineData("WordLayeredCornerParagraphAllStories", true)]
    [InlineData("WordLayeredCornerCenterParagraphAllStories", false)]
    [InlineData("WordLayeredCornerCenterParagraphAllStories", true)]
    public void LayeredCornerParagraphAlignmentSurvivesBothDocRoutes(
        string fixtureName, bool native)
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            fixtureName));
        var source = File.ReadAllBytes(stem + ".docx");
        var bytes = native ? File.ReadAllBytes(stem + ".doc") :
            DxpDocExport.Export(source);
        foreach (var binary in new[] { bytes,
            DxpDocExport.Export(DxpDocToDocx.Project(bytes).DocxBytes) })
        {
            using var index = new DocTextIndexWalker().Index(new MemoryStream(binary));
            Assert.Contains(index.StyleDefinitions, x => x.Name == "Corner Paragraph");
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            AssertEffectiveParagraphLayout(source, projected);
            Assert.Equal(ReadStories(source), ReadStories(projected));
            Assert.Empty(Validate(projected));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LayeredCornerParagraphAlignmentChangesFailLayoutGate(bool direct)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordLayeredCornerParagraphAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            if (direct)
            {
                main.Document.Body!.Descendants<Table>().First()
                    .Elements<TableRow>().Last().Elements<TableCell>().Last()
                    .Descendants<Paragraph>().Single().ParagraphProperties!
                    .GetFirstChild<Justification>()!.Val = JustificationValues.Right;
                main.Document.Save();
            }
            else
            {
                var style = main.StyleDefinitionsPart!.Styles!.Elements<Style>()
                    .Single(x => x.StyleName?.Val?.Value == "Corner Paragraph");
                style.StyleParagraphProperties!.GetFirstChild<Justification>()!
                    .Val = JustificationValues.Right;
                main.StyleDefinitionsPart.Styles.Save();
            }
        }
        var failure = Record.Exception(() =>
            AssertEffectiveParagraphLayout(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(".alignment", failure.Message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LayeredCornerRunColorsSurviveBothDocRoutes(bool native)
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordLayeredCornerRunStyleAllStories"));
        var source = File.ReadAllBytes(stem + ".docx");
        var values = ReadEffectiveRunFormatting(source);
        Assert.True(values.Values.Count(x => x == "0099AA") >= 3);
        Assert.True(values.Values.Count(x => x == "DDAA00") >= 3);
        var bytes = native ? File.ReadAllBytes(stem + ".doc") :
            DxpDocExport.Export(source);
        foreach (var binary in new[] { bytes,
            DxpDocExport.Export(DxpDocToDocx.Project(bytes).DocxBytes) })
        {
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            AssertEffectiveRunFormatting(source, projected);
            Assert.Equal(ReadStories(source), ReadStories(projected));
            Assert.Empty(Validate(projected));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LayeredCornerRunColorChangesFailFormattingGate(bool direct)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordLayeredCornerRunStyleAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            var color = direct
                ? main.Document.Body!.Descendants<Table>().First()
                    .Elements<TableRow>().Last().Elements<TableCell>().Last()
                    .Descendants<Run>().Single().RunProperties!
                    .GetFirstChild<Color>()!
                : main.StyleDefinitionsPart!.Styles!.Elements<Style>()
                    .Single(x => x.StyleName?.Val?.Value == "Corner Character")
                    .StyleRunProperties!.GetFirstChild<Color>()!;
            color.Val = "123456";
            if (direct) main.Document.Save();
            else main.StyleDefinitionsPart!.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(".color", failure.Message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ConditionalCornerRunStylesSurviveBothDocRoutes(bool native)
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordConditionalCornerRunStyleAllStories"));
        var source = File.ReadAllBytes(stem + ".docx");
        var expected = ReadEffectiveRunFormatting(source);
        Assert.Contains(expected.Values, value => value == "CC00CC");
        Assert.Contains(expected.Values, value => value == "FF8800");
        var bytes = native ? File.ReadAllBytes(stem + ".doc") :
            DxpDocExport.Export(source);
        foreach (var binary in new[] { bytes,
            DxpDocExport.Export(DxpDocToDocx.Project(bytes).DocxBytes) })
        {
            using var index = new DocTextIndexWalker().Index(new MemoryStream(binary));
            if (!native)
            {
                var style = Assert.Single(index.StyleDefinitions,
                    x => x.Name == "Logical Start");
                Assert.Equal(0x00CC00CCu,
                    style.ConditionalCharacterFormatting?[0x0200].ColorRef);
                Assert.Equal(0x000088FFu,
                    style.ConditionalCharacterFormatting?[0x0100].ColorRef);
            }
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            AssertEffectiveRunFormatting(source, projected);
            Assert.Equal(ReadStories(source), ReadStories(projected));
            Assert.Empty(Validate(projected));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ConditionalCornerRunColorIsPartOfThePairedFormattingGate(
        bool northEast)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordConditionalCornerRunStyleAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    "Logical Start");
            var corner = style.Elements<TableStyleProperties>().Single(x =>
                x.Type?.Value == (northEast ? TableStyleOverrideValues.NorthEastCell :
                    TableStyleOverrideValues.NorthWestCell));
            corner.Descendants<Color>().Single().Val = "123456";
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(".color", failure.Message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ConditionalColumnRunStylesSurviveBothDocRoutes(bool native)
    {
        var stem = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordConditionalColumnRunStyleAllStories"));
        var source = File.ReadAllBytes(stem + ".docx");
        var bytes = native ? File.ReadAllBytes(stem + ".doc") :
            DxpDocExport.Export(source);
        foreach (var binary in new[] { bytes,
            DxpDocExport.Export(DxpDocToDocx.Project(bytes).DocxBytes) })
        {
            using var index = new DocTextIndexWalker().Index(new MemoryStream(binary));
            var style = Assert.Single(index.StyleDefinitions,
                x => x.Name == "Logical Start");
            if (native)
            {
                Assert.Contains(index.CharacterFormatting,
                    x => x.Formatting.ColorRef == 0x00FF0000u);
                Assert.Contains(index.CharacterFormatting,
                    x => x.Formatting.ColorRef == 0x00008800u);
            }
            else
            {
                Assert.True(style.ConditionalCharacterFormatting?[0x0004].Italic);
                Assert.Equal(0x00FF0000u,
                    style.ConditionalCharacterFormatting?[0x0004].ColorRef);
                Assert.Equal(0x00008800u,
                    style.ConditionalCharacterFormatting?[0x0008].ColorRef);
            }
            var projected = DxpDocToDocx.Project(binary).DocxBytes;
            AssertEffectiveRunFormatting(source, projected);
            Assert.Equal(ReadStories(source), ReadStories(projected));
            Assert.Empty(Validate(projected));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ConditionalColumnRunColorIsPartOfThePairedFormattingGate(
        bool lastColumn)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordConditionalColumnRunStyleAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    "Logical Start");
            var column = style.Elements<TableStyleProperties>().Single(x =>
                x.Type?.Value == (lastColumn ? TableStyleOverrideValues.LastColumn :
                    TableStyleOverrideValues.FirstColumn));
            column.Descendants<Color>().Single().Val = "FF0000";
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(".color", failure.Message);
    }

    [Fact]
    public void FirstRowConditionalRunColorIsPartOfThePairedFormattingGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordConditionalFirstRowRunStyleAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    "Logical Start");
            var firstRow = style.Elements<TableStyleProperties>().Single(x =>
                x.Type?.Value == TableStyleOverrideValues.FirstRow);
            firstRow.Descendants<Color>().Single().Val = "0000FF";
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(".color", failure.Message);
    }

    [Fact]
    public void CharacterStyleEmphasisLayerIsPartOfThePairedFormattingGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedCharacterEmphasisGridAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    "Grid run base");
            style.StyleRunProperties!.GetFirstChild<Emphasis>()!.Val =
                EmphasisMarkValues.Circle;
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(".emphasis", failure.Message);
    }

    [Fact]
    public void InheritedEmphasisIsPartOfThePairedFormattingGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedEmphasisAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    "CJK spacing base");
            style.StyleRunProperties!.GetFirstChild<Emphasis>()!.Val =
                EmphasisMarkValues.Circle;
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(".emphasis", failure.Message);
    }

    [Fact]
    public void InheritedRunGridSnapIsPartOfThePairedFormattingGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedRunGridSnapAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    "Grid run base");
            style.StyleRunProperties!.GetFirstChild<SnapToGrid>()!.Val = true;
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(".snapToGrid", failure.Message);
    }

    [Fact]
    public void UnexpectedRunPositioningIsRejectedByPairedFormattingGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedCharacterEmphasisGridAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var run = document.MainDocumentPart!.Document.Body!
                .Descendants<Run>().First(x => x.GetFirstChild<Text>() != null);
            (run.RunProperties ?? run.PrependChild(new RunProperties()))
                .Append(new Position { Val = "6" });
            document.MainDocumentPart.Document.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveRunFormatting(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(".baselineOffset", failure.Message);
    }

    [Fact]
    public void InheritedTextAlignmentIsPartOfThePairedLayoutGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedTextAlignmentAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    "Right grid base");
            style.StyleParagraphProperties!.GetFirstChild<TextAlignment>()!
                .Val = VerticalTextAlignmentValues.Bottom;
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveParagraphLayout(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(".textAlignment", failure.Message);
    }

    [Fact]
    public void InheritedGridAdjustedRightIndentIsPartOfThePairedLayoutGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordInheritedAdjustRightGridAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().Single(x => x.StyleName?.Val?.Value ==
                    "Right grid derived");
            style.StyleParagraphProperties!.GetFirstChild<AdjustRightIndent>()!
                .Val = true;
            document.MainDocumentPart.StyleDefinitionsPart.Styles!.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveParagraphLayout(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("body.paragraph2.adjustRightIndent", failure.Message);
    }

    [Fact]
    public void GridAdjustedRightIndentIsPartOfThePairedLayoutGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordCharacterGridAdjustRightAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var paragraph = document.MainDocumentPart!.Document.Body!
                .Elements<Paragraph>().First(x => x.ParagraphProperties?
                    .GetFirstChild<AdjustRightIndent>()?.Val?.Value == false);
            paragraph.ParagraphProperties!.GetFirstChild<AdjustRightIndent>()!
                .Val = true;
            document.MainDocumentPart.Document.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveParagraphLayout(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains(".adjustRightIndent", failure.Message);
    }

    [Fact]
    public void UnexpectedAfterSpacingIsRejectedByPairedLayoutGate()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordVisibleEmptyDefaultsAllStories.docx"));
        var source = File.ReadAllBytes(path);
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var paragraph = document.MainDocumentPart!.Document.Body!
                .Elements<Paragraph>().First();
            paragraph.ParagraphProperties ??= new ParagraphProperties();
            paragraph.ParagraphProperties.RemoveAllChildren<SpacingBetweenLines>();
            paragraph.ParagraphProperties.Append(new SpacingBetweenLines
                { After = "320" });
            document.MainDocumentPart.Document.Save();
        }
        var failure = Record.Exception(() =>
            AssertEffectiveParagraphLayout(source, stream.ToArray()));
        Assert.NotNull(failure);
        Assert.Contains("body.paragraph0.after", failure.Message);
    }

    private static void AssertEffectiveParagraphLayout(byte[] reference, byte[] actual)
    {
        var expected = ReadEffectiveParagraphLayout(reference);
        var observed = ReadEffectiveParagraphLayout(actual);
        foreach (var (key, value) in expected)
            Assert.True((observed.TryGetValue(key, out var found) && found == value) ||
                (!observed.ContainsKey(key) && value == "0" &&
                    (key.EndsWith(".after", StringComparison.Ordinal) ||
                     key.EndsWith(".before", StringComparison.Ordinal))),
                $"Paragraph layout {key}: expected {value}, observed " +
                (observed.TryGetValue(key, out var current) ? current : "<missing>"));
        foreach (var (key, value) in observed)
        {
            if (expected.ContainsKey(key)) continue;
            var prefix = key[..(key.LastIndexOf('.') + 1)];
            // DOC table terminators can materialize a structural paragraph
            // absent from the DOCX source's visible paragraph inventory.
            if (!expected.Keys.Any(x => x.StartsWith(prefix,
                StringComparison.Ordinal))) continue;
            var property = key[(key.LastIndexOf('.') + 1)..];
            if ((property is "before" or "after" or "left" or "right" or
                "firstLine" or "hanging" or "line" or "leftChars" or
                "rightChars") && value == "0") continue;
            // Word's implicit spacing and line defaults may become explicit
            // in a DOC conversion; they do not add visible formatting.
            if ((property == "after" && value == "160") ||
                (property == "line" && value == "278") ||
                (property == "lineRule" && value == "auto") ||
                (property == "line" && value == "240") ||
                (property == "outlineLevel" && value == "9") ||
                ((property is "snapToGrid" or "adjustRightIndent" or "kinsoku" or
                    "wordWrap" or "autoSpaceDE" or "autoSpaceDN") &&
                    value == "true")) continue;
            Assert.Fail($"Unexpected paragraph layout {key}: {value}");
        }
    }

    private static void AssertDefaultTabInterval(byte[] reference, byte[] actual)
    {
        static int Interval(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes);
            using var document = WordprocessingDocument.Open(stream, false);
            return document.MainDocumentPart?.DocumentSettingsPart?.Settings?
                .GetFirstChild<DefaultTabStop>()?.Val?.Value ?? 720;
        }
        Assert.Equal(Interval(reference), Interval(actual));
    }

    private static void AssertMirrorMargins(byte[] reference, byte[] actual)
    {
        static bool Enabled(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes);
            using var document = WordprocessingDocument.Open(stream, false);
            var setting = document.MainDocumentPart?.DocumentSettingsPart?.Settings?
                .GetFirstChild<MirrorMargins>();
            return setting != null && (setting.Val?.Value ?? true);
        }
        Assert.Equal(Enabled(reference), Enabled(actual));
    }

    private static void AssertGutterAtTop(byte[] reference, byte[] actual)
    {
        static bool Enabled(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes);
            using var document = WordprocessingDocument.Open(stream, false);
            var setting = document.MainDocumentPart?.DocumentSettingsPart?.Settings?
                .GetFirstChild<GutterAtTop>();
            return setting != null && (setting.Val?.Value ?? true);
        }
        Assert.Equal(Enabled(reference), Enabled(actual));
    }

    private static void AssertEffectiveTabStops(byte[] reference, byte[] actual)
    {
        var expected = ReadEffectiveTabStops(reference);
        var observed = ReadEffectiveTabStops(actual);
        foreach (var (key, value) in expected)
            Assert.True(observed.TryGetValue(key, out var found) && found == value,
                $"Tab stop {key}: expected {value}, observed " +
                (observed.TryGetValue(key, out var current) ? current : "<missing>"));
        foreach (var (key, value) in observed)
            Assert.True(expected.TryGetValue(key, out var found) && found == value,
                $"Unexpected tab stop {key}: observed {value}, expected " +
                (expected.TryGetValue(key, out var current) ? current : "<missing>"));
    }

    private static Dictionary<string, string> ReadEffectiveTabStops(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var styles = main.StyleDefinitionsPart?.Styles;
        var definitions = styles?.Elements<Style>()
            .Where(x => x.StyleId?.Value != null)
            .ToDictionary(x => x.StyleId!.Value!, StringComparer.Ordinal) ??
            new Dictionary<string, Style>(StringComparer.Ordinal);
        var defaultStyle = styles?.Elements<Style>()
            .FirstOrDefault(x => x.Type?.Value == StyleValues.Paragraph &&
                x.Default?.Value == true)?.StyleId?.Value;
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        var sections = main.Document!.Body!.Descendants<SectionProperties>().ToArray();
        static string? Attribute(DocumentFormat.OpenXml.OpenXmlElement element,
            string name) => element.GetAttributes()
                .FirstOrDefault(x => x.LocalName == name).Value;
        static void Apply(DocumentFormat.OpenXml.OpenXmlElement? properties,
            IDictionary<string, string> values)
        {
            var tabs = properties?.ChildElements.FirstOrDefault(x => x.LocalName == "tabs");
            if (tabs == null) return;
            foreach (var tab in tabs.ChildElements.Where(x => x.LocalName == "tab"))
            {
                var position = Attribute(tab, "pos");
                if (position == null) continue;
                var alignment = Attribute(tab, "val") ?? "left";
                if (alignment == "clear") values.Remove(position);
                else values[position] = alignment + "|" + (Attribute(tab, "leader") ?? "none");
            }
        }
        void AddStory(string key, DocumentFormat.OpenXml.OpenXmlElement story,
            SectionProperties? section)
        {
            static int ContentWidth(SectionProperties? properties)
            {
                var page = properties?.GetFirstChild<PageSize>();
                var margins = properties?.GetFirstChild<PageMargin>();
                return (int)(page?.Width?.Value ?? 12240) -
                    (int)(margins?.Left?.Value ?? 1440) -
                    (int)(margins?.Right?.Value ?? 1440);
            }
            var contentWidth = ContentWidth(section);
            var paragraphs = story.Descendants<Paragraph>()
                .Where(x => x.Descendants<Text>().Any(t => t.Text.Length != 0))
                .ToArray();
            for (var i = 0; i < paragraphs.Length; i++)
            {
                var paragraphContentWidth = contentWidth;
                if (story is Body && sections.Length != 0)
                {
                    var ordinal = story.Descendants<Paragraph>()
                        .TakeWhile(x => !ReferenceEquals(x, paragraphs[i]))
                        .Count(x => x.ParagraphProperties?.SectionProperties != null);
                    paragraphContentWidth = ContentWidth(sections[Math.Min(ordinal,
                        sections.Length - 1)]);
                }
                if (paragraphs[i].Ancestors<TableCell>().FirstOrDefault() is { } cell)
                {
                    var table = cell.Ancestors<Table>().FirstOrDefault();
                    var tableMargins = table?.TableProperties?
                        .GetFirstChild<TableCellMarginDefault>();
                    var styleMargins = definitions.TryGetValue(
                        table?.TableProperties?.GetFirstChild<TableStyle>()?.Val?.Value ??
                        "TableNormal", out var tableStyle)
                        ? tableStyle.StyleTableProperties?
                            .GetFirstChild<TableCellMarginDefault>() : null;
                    var rowMargins = cell.Parent is TableRow row
                        ? row.TablePropertyExceptions?
                            .GetFirstChild<TableCellMarginDefault>() : null;
                    var cellMargins = cell.TableCellProperties?
                        .GetFirstChild<TableCellMargin>();
                    static int Side(DocumentFormat.OpenXml.OpenXmlElement? margins,
                        string name, int fallback)
                    {
                        var side = margins?.ChildElements.FirstOrDefault(x =>
                            x.LocalName == name);
                        return int.TryParse(side?.GetAttribute("w",
                            "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
                            .Value, out var width) ? width : fallback;
                    }
                    var cellLeft = Side(cellMargins, "left", Side(rowMargins, "left",
                        Side(tableMargins, "left", Side(styleMargins, "left", 10))));
                    var cellRight = Side(cellMargins, "right", Side(rowMargins, "right",
                        Side(tableMargins, "right", Side(styleMargins, "right", 10))));
                    if (int.TryParse(cell.TableCellProperties?
                        .GetFirstChild<TableCellWidth>()?.Width?.Value, out var cellWidth))
                        paragraphContentWidth = cellWidth - cellLeft - cellRight;
                }
                var values = new Dictionary<string, string>(StringComparer.Ordinal);
                var seen = new HashSet<string>(StringComparer.Ordinal);
                var left = 0; var right = 0; var firstLine = 0;
                void ApplyIndent(DocumentFormat.OpenXml.OpenXmlElement? properties)
                {
                    var indent = properties?.GetFirstChild<Indentation>();
                    if (indent == null) return;
                    if (int.TryParse(indent.Left?.Value, out var l)) left = l;
                    if (int.TryParse(indent.Right?.Value, out var r)) right = r;
                    if (int.TryParse(indent.FirstLine?.Value, out var first))
                        firstLine = first;
                    else if (int.TryParse(indent.Hanging?.Value, out var hanging))
                        firstLine = -hanging;
                }
                void ApplyStyle(string? id)
                {
                    if (id == null || !seen.Add(id) ||
                        !definitions.TryGetValue(id, out var style)) return;
                    ApplyStyle(style.BasedOn?.Val?.Value);
                    Apply(style.StyleParagraphProperties, values);
                    ApplyIndent(style.StyleParagraphProperties);
                }
                ApplyStyle(paragraphs[i].ParagraphProperties?.ParagraphStyleId?.Val?.Value ??
                    defaultStyle);
                Apply(paragraphs[i].ParagraphProperties, values);
                ApplyIndent(paragraphs[i].ParagraphProperties);
                if (paragraphs[i].Descendants<PositionalTab>().Any(x =>
                    x.Alignment?.Value != AbsolutePositionTabAlignmentValues.Left))
                    values.Clear();
                foreach (var tab in paragraphs[i].Descendants<PositionalTab>())
                    if (tab.RelativeTo?.Value == AbsolutePositionTabPositioningBaseValues.Margin ||
                        tab.RelativeTo?.Value == AbsolutePositionTabPositioningBaseValues.Indent)
                    {
                        var relativeIndent = tab.RelativeTo?.Value ==
                            AbsolutePositionTabPositioningBaseValues.Indent;
                        var leader = tab.GetAttribute("leader",
                            "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
                            .Value ?? "none";
                        if (tab.Alignment?.Value == AbsolutePositionTabAlignmentValues.Center)
                            values[(paragraphContentWidth / 2 + (relativeIndent ? (left + firstLine) / 2 : 0))
                                .ToString()] = "center|" + leader;
                        else if (tab.Alignment?.Value == AbsolutePositionTabAlignmentValues.Right)
                            values[(paragraphContentWidth - (relativeIndent ? right : 0)).ToString()] =
                                "right|" + leader;
                    }
                foreach (var (position, value) in values)
                    result[$"{key}.paragraph{i}.tab{position}"] = value;
            }
        }
        AddStory("body", main.Document.Body, sections.FirstOrDefault());
        for (var i = 0; i < sections.Length; i++)
        {
            foreach (var reference in sections[i].Elements<HeaderReference>())
                AddStory($"section{i}.header.{Slot(reference.Type?.Value)}",
                    ((HeaderPart)main.GetPartById(reference.Id!)).Header!, sections[i]);
            foreach (var reference in sections[i].Elements<FooterReference>())
                AddStory($"section{i}.footer.{Slot(reference.Type?.Value)}",
                    ((FooterPart)main.GetPartById(reference.Id!)).Footer!, sections[i]);
        }
        return result;
    }

    private static Dictionary<string, string> ReadEffectiveParagraphLayout(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var styles = main.StyleDefinitionsPart?.Styles;
        var definitions = styles?.Elements<Style>()
            .Where(x => x.StyleId?.Value != null)
            .ToDictionary(x => x.StyleId!.Value!, StringComparer.Ordinal) ??
            new Dictionary<string, Style>(StringComparer.Ordinal);
        var defaultParagraphStyle = styles?.Elements<Style>()
            .FirstOrDefault(x => x.Type?.Value == StyleValues.Paragraph &&
                x.Default?.Value == true)?.StyleId?.Value;
        var defaults = styles?.Elements<DocDefaults>().FirstOrDefault()?
            .GetFirstChild<ParagraphPropertiesDefault>()?
            .GetFirstChild<ParagraphPropertiesBaseStyle>();
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        const string ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        static string? Attribute(DocumentFormat.OpenXml.OpenXmlElement? element,
            string name)
        {
            var value = element?.GetAttributes()
                .FirstOrDefault(x => x.LocalName == name && x.NamespaceUri == ns).Value;
            return string.IsNullOrEmpty(value) ? null : value;
        }
        static void Apply(DocumentFormat.OpenXml.OpenXmlElement? properties,
            IDictionary<string, string> values)
        {
            if (properties == null) return;
            void Set(string key, string? value)
            {
                if (value != null) values[key] = value;
            }
            Set("alignment", Attribute(properties.GetFirstChild<Justification>(), "val"));
            var indentation = properties.GetFirstChild<Indentation>();
            Set("left", Attribute(indentation, "start") ?? Attribute(indentation, "left"));
            Set("right", Attribute(indentation, "end") ?? Attribute(indentation, "right"));
            Set("leftChars", Attribute(indentation, "startChars") ??
                Attribute(indentation, "leftChars"));
            Set("rightChars", Attribute(indentation, "endChars") ??
                Attribute(indentation, "rightChars"));
            Set("firstLine", Attribute(indentation, "firstLine"));
            Set("hanging", Attribute(indentation, "hanging"));
            Set("firstLineChars", Attribute(indentation, "firstLineChars"));
            Set("hangingChars", Attribute(indentation, "hangingChars"));
            var spacing = properties.GetFirstChild<SpacingBetweenLines>();
            Set("before", Attribute(spacing, "before"));
            Set("after", Attribute(spacing, "after"));
            Set("beforeAutoSpacing", Attribute(spacing, "beforeAutospacing"));
            Set("afterAutoSpacing", Attribute(spacing, "afterAutospacing"));
            Set("beforeLines", Attribute(spacing, "beforeLines"));
            Set("afterLines", Attribute(spacing, "afterLines"));
            Set("line", Attribute(spacing, "line"));
            Set("lineRule", Attribute(spacing, "lineRule"));
            Set("textAlignment", Attribute(properties.GetFirstChild<TextAlignment>(), "val"));
            Set("outlineLevel", Attribute(properties.GetFirstChild<OutlineLevel>(), "val"));
            foreach (var (tag, key) in new[] { ("keepNext", "keepNext"),
                ("keepLines", "keepLines"),
                ("contextualSpacing", "contextualSpacing"),
                ("mirrorIndents", "mirrorIndents"),
                ("suppressAutoHyphens", "suppressAutoHyphens"),
                ("suppressLineNumbers", "suppressLineNumbers"),
                ("snapToGrid", "snapToGrid"),
                ("adjustRightInd", "adjustRightIndent"),
                ("kinsoku", "kinsoku"),
                ("wordWrap", "wordWrap"),
                ("autoSpaceDE", "autoSpaceDE"),
                ("autoSpaceDN", "autoSpaceDN"),
                ("bidi", "bidi"),
                ("pageBreakBefore", "pageBreakBefore"),
                ("widowControl", "widowControl") })
            {
                var flag = properties.ChildElements.FirstOrDefault(x => x.LocalName == tag);
                if (flag != null)
                    values[key] = Attribute(flag, "val") is "0" or "false" or "off"
                        ? "false" : "true";
            }
        }
        void AddStory(string key, DocumentFormat.OpenXml.OpenXmlElement story)
        {
            var paragraphs = story.Descendants<Paragraph>().ToArray();
            for (var i = 0; i < paragraphs.Length; i++)
            {
                var values = new Dictionary<string, string>(StringComparer.Ordinal);
                Apply(defaults, values);
                var styleId = paragraphs[i].ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                var seen = new HashSet<string>(StringComparer.Ordinal);
                var paragraphStyleSetsAlignment = false;
                string? conditionalAlignment = null;
                var tableSetsAlignment = false;
                void ApplyStyle(string? id)
                {
                    if (id == null || !seen.Add(id) ||
                        !definitions.TryGetValue(id, out var style)) return;
                    ApplyStyle(style.BasedOn?.Val?.Value);
                    Apply(style.StyleParagraphProperties, values);
                    paragraphStyleSetsAlignment |= style.StyleParagraphProperties?
                        .Justification != null;
                }
                if (paragraphs[i].Ancestors<Table>().FirstOrDefault() is { } table &&
                    paragraphs[i].Ancestors<TableRow>().FirstOrDefault() is { } row)
                {
                    var look = table.TableProperties?.GetFirstChild<TableLook>();
                    ushort.TryParse(look?.Val?.Value,
                        System.Globalization.NumberStyles.HexNumber,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out var lookMask);
                    if (look == null) lookMask = 0x04A0;
                    var attributeOnlyLook = look != null && look.Val?.Value == null;
                    var firstColumnLook = look?.FirstColumn?.Value ??
                        (attributeOnlyLook || (lookMask & 0x0080) != 0);
                    var lastColumnLook = look?.LastColumn?.Value ??
                        (attributeOnlyLook || (lookMask & 0x0100) != 0);
                    var firstRowLook = look?.FirstRow?.Value ??
                        (lookMask & 0x0020) != 0;
                    var lastRowLook = look?.LastRow?.Value ??
                        (lookMask & 0x0040) != 0;
                    void ApplyRowStyle(TableStyleOverrideValues condition)
                    {
                        var seenTableStyles = new HashSet<string>(StringComparer.Ordinal);
                        void ApplyCondition(string? id)
                        {
                            if (id == null || !seenTableStyles.Add(id) ||
                                !definitions.TryGetValue(id, out var style)) return;
                            ApplyCondition(style.BasedOn?.Val?.Value);
                            var conditional = style.Elements<TableStyleProperties>()
                                .FirstOrDefault(x => x.Type?.Value == condition);
                            var paragraphProperties = conditional?
                                .GetFirstChild<StyleParagraphProperties>();
                            Apply(paragraphProperties, values);
                            tableSetsAlignment |= paragraphProperties?.Justification != null;
                        }
                        ApplyCondition(table.TableProperties?.TableStyle?.Val?.Value);
                    }
                    var rows = table.Elements<TableRow>();
                    if (firstRowLook &&
                        ReferenceEquals(row, rows.FirstOrDefault()))
                        ApplyRowStyle(TableStyleOverrideValues.FirstRow);
                    if (lastRowLook &&
                        ReferenceEquals(row, rows.LastOrDefault()))
                        ApplyRowStyle(TableStyleOverrideValues.LastRow);
                    var cells = row.Elements<TableCell>();
                    var cell = paragraphs[i].Ancestors<TableCell>().FirstOrDefault();
                    var firstRow = ReferenceEquals(row, rows.FirstOrDefault());
                    var lastRow = ReferenceEquals(row, rows.LastOrDefault());
                    var firstColumn = ReferenceEquals(cell, cells.FirstOrDefault());
                    var lastColumn = ReferenceEquals(cell, cells.LastOrDefault());
                    if (firstColumnLook && firstColumn)
                        ApplyRowStyle(TableStyleOverrideValues.FirstColumn);
                    if (lastColumnLook && lastColumn)
                        ApplyRowStyle(TableStyleOverrideValues.LastColumn);
                    if (firstRowLook && firstColumnLook && firstRow && firstColumn)
                        ApplyRowStyle(TableStyleOverrideValues.NorthWestCell);
                    if (firstRowLook && lastColumnLook && firstRow && lastColumn)
                        ApplyRowStyle(TableStyleOverrideValues.NorthEastCell);
                    if (lastRowLook && firstColumnLook && lastRow && firstColumn)
                        ApplyRowStyle(TableStyleOverrideValues.SouthWestCell);
                    if (lastRowLook && lastColumnLook && lastRow && lastColumn)
                        ApplyRowStyle(TableStyleOverrideValues.SouthEastCell);
                }
                if (tableSetsAlignment)
                    values.TryGetValue("alignment", out conditionalAlignment);
                ApplyStyle(styleId ?? defaultParagraphStyle);
                Apply(paragraphs[i].ParagraphProperties, values);
                if (conditionalAlignment != null &&
                    !paragraphStyleSetsAlignment &&
                    paragraphs[i].ParagraphProperties?.Justification?.Val?.Value ==
                        JustificationValues.Left &&
                    paragraphs[i].Ancestors<TableCell>().FirstOrDefault()?
                        .TableCellProperties?.ChildElements.Any(x =>
                            x.LocalName == "cnfStyle") != true)
                    values["alignment"] = conditionalAlignment;
                values.TryAdd("outlineLevel", "9");
                foreach (var property in new[] { "snapToGrid",
                    "adjustRightIndent", "kinsoku",
                    "wordWrap", "autoSpaceDE", "autoSpaceDN" })
                    values.TryAdd(property, "true");
                var rightToLeft = values.TryGetValue("bidi", out var direction) &&
                    direction == "true";
                if (values.TryGetValue("alignment", out var alignment))
                {
                    if (alignment is "start" or "end" or "left" or "right")
                        values["alignment"] =
                            (alignment is "start" or "left") == rightToLeft
                                ? "right" : "left";
                }
                else values["alignment"] = rightToLeft ? "right" : "left";
                foreach (var (property, value) in values)
                    result[$"{key}.paragraph{i}.{property}"] = value;
            }
        }
        AddStory("body", main.Document!.Body!);
        var sections = main.Document.Body.Descendants<SectionProperties>().ToArray();
        for (var i = 0; i < sections.Length; i++)
        {
            foreach (var reference in sections[i].Elements<HeaderReference>())
                AddStory($"section{i}.header.{Slot(reference.Type?.Value)}",
                    ((HeaderPart)main.GetPartById(reference.Id!)).Header!);
            foreach (var reference in sections[i].Elements<FooterReference>())
                AddStory($"section{i}.footer.{Slot(reference.Type?.Value)}",
                    ((FooterPart)main.GetPartById(reference.Id!)).Footer!);
        }
        return result;
    }

    private static string Slot(HeaderFooterValues? value) =>
        value == HeaderFooterValues.Even ? "even" :
        value == HeaderFooterValues.First ? "first" : "default";

    private static void AssertDrawingGeometry(byte[] reference, byte[] actual)
    {
        var expected = ReadDrawingGeometry(reference);
        var observed = ReadDrawingGeometry(actual);
        Assert.Equal(expected.Keys.OrderBy(x => x, StringComparer.Ordinal),
            observed.Keys.OrderBy(x => x, StringComparer.Ordinal));
        foreach (var (key, value) in expected)
            Assert.True(observed.TryGetValue(key, out var found) &&
                (found == value ||
                    ((key.EndsWith(".cx", StringComparison.Ordinal) ||
                      key.EndsWith(".cy", StringComparison.Ordinal) ||
                      key.EndsWith(".offset", StringComparison.Ordinal) ||
                      key.Contains(".wrapDistance.", StringComparison.Ordinal)) &&
                     long.TryParse(value, out var expectedEmu) &&
                     long.TryParse(found, out var foundEmu) &&
                     Math.Abs(expectedEmu - foundEmu) <= 635) ||
                    (key.Contains(".crop.", StringComparison.Ordinal) &&
                     long.TryParse(value, out var expectedCrop) &&
                     long.TryParse(found, out var foundCrop) &&
                     Math.Abs(expectedCrop - foundCrop) <= 2)),
                $"Drawing geometry {key}: expected {value}, observed " +
                (observed.TryGetValue(key, out var current) ? current : "<missing>"));
    }

    private static Dictionary<string, string> ReadDrawingGeometry(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        static string? Attribute(DocumentFormat.OpenXml.OpenXmlElement element,
            string name) => element.GetAttributes()
                .FirstOrDefault(x => x.LocalName == name).Value;
        void AddStory(string key, DocumentFormat.OpenXml.OpenXmlElement story)
        {
            var drawings = story.Descendants<Drawing>().ToArray();
            for (var i = 0; i < drawings.Length; i++)
            {
                var container = drawings[i].ChildElements.FirstOrDefault(x =>
                    x.LocalName is "anchor" or "inline");
                if (container == null) continue;
                var prefix = $"{key}.drawing{i}";
                result[$"{prefix}.kind"] = container.LocalName;
                var extent = container.ChildElements.FirstOrDefault(x => x.LocalName == "extent");
                if (extent != null)
                {
                    if (Attribute(extent, "cx") is string cx) result[$"{prefix}.cx"] = cx;
                    if (Attribute(extent, "cy") is string cy) result[$"{prefix}.cy"] = cy;
                }
                var transform = container.Descendants().FirstOrDefault(x =>
                    x.LocalName == "xfrm");
                result[$"{prefix}.flipH"] = (transform != null &&
                    Attribute(transform, "flipH") is "1" or "true").ToString();
                result[$"{prefix}.flipV"] = (transform != null &&
                    Attribute(transform, "flipV") is "1" or "true").ToString();
                var rawRotation = Attribute(transform ?? container, "rot");
                result[$"{prefix}.rotation"] = long.TryParse(rawRotation, out var angle)
                    ? ((angle % 21600000 + 21600000) % 21600000).ToString()
                    : "0";
                var crop = container.Descendants().FirstOrDefault(x => x.LocalName == "srcRect");
                if (crop != null)
                    foreach (var side in new[] { "l", "t", "r", "b" })
                        if (Attribute(crop, side) is string fraction &&
                            (!long.TryParse(fraction, out var value) || value != 0))
                            result[$"{prefix}.crop.{side}"] = fraction;
                if (container.LocalName != "anchor") continue;
                foreach (var (name, side, defaultValue) in new[]
                {
                    ("distT", "top", "0"), ("distB", "bottom", "0"),
                    ("distL", "left", "114300"),
                    ("distR", "right", "114300")
                })
                    result[$"{prefix}.wrapDistance.{side}"] =
                        Attribute(container, name) ?? defaultValue;
                foreach (var (name, axis) in new[] { ("positionH", "horizontal"),
                    ("positionV", "vertical") })
                {
                    var position = container.ChildElements.FirstOrDefault(x => x.LocalName == name);
                    if (position == null) continue;
                    if (Attribute(position, "relativeFrom") is string relativeFrom)
                        result[$"{prefix}.{axis}.relativeFrom"] = relativeFrom;
                    var offset = position.ChildElements.FirstOrDefault(x => x.LocalName == "posOffset");
                    if (offset?.InnerText is string offsetText && offsetText.Length != 0)
                        result[$"{prefix}.{axis}.offset"] = offsetText;
                    var align = position.ChildElements.FirstOrDefault(x => x.LocalName == "align");
                    if (align?.InnerText is string alignment && alignment.Length != 0)
                        result[$"{prefix}.{axis}.align"] = alignment;
                }
                var wrap = container.ChildElements.FirstOrDefault(x =>
                    x.LocalName.StartsWith("wrap", StringComparison.Ordinal));
                if (wrap != null) result[$"{prefix}.wrap"] = wrap.LocalName;
            }
        }
        AddStory("body", main.Document!.Body!);
        var sections = main.Document.Body.Descendants<SectionProperties>().ToArray();
        for (var i = 0; i < sections.Length; i++)
        {
            foreach (var reference in sections[i].Elements<HeaderReference>())
                AddStory($"section{i}.header.{Slot(reference.Type?.Value)}",
                    ((HeaderPart)main.GetPartById(reference.Id!)).Header!);
            foreach (var reference in sections[i].Elements<FooterReference>())
                AddStory($"section{i}.footer.{Slot(reference.Type?.Value)}",
                    ((FooterPart)main.GetPartById(reference.Id!)).Footer!);
        }
        return result;
    }

    private static void AssertEffectiveRunFormatting(byte[] reference, byte[] actual,
        bool wordValidatedToggles = false, bool wordValidatedStrike = false,
        bool wordValidatedThemeLuminance = false)
    {
        var expected = ReadEffectiveRunFormatting(reference);
        var observed = ReadEffectiveRunFormatting(actual);
        foreach (var (key, value) in expected)
        {
            // These inherited toggles are checked against Word's effective
            // per-character formatting in the optional render test. A simple
            // XML overlay reports a false positive for the source fixture.
            if (wordValidatedToggles &&
                (key.EndsWith(".bold", StringComparison.Ordinal) ||
                 key.EndsWith(".italic", StringComparison.Ordinal))) continue;
            if (wordValidatedStrike &&
                key.EndsWith(".strike", StringComparison.Ordinal)) continue;
            var closeThemeColor = false;
            if (wordValidatedThemeLuminance &&
                key.EndsWith(".color", StringComparison.Ordinal) &&
                value.Length == 6 && observed.TryGetValue(key, out var color) &&
                color.Length == 6)
            {
                closeThemeColor = Enumerable.Range(0, 3).All(i =>
                    Math.Abs(Convert.ToInt32(value.Substring(i * 2, 2), 16) -
                        Convert.ToInt32(color.Substring(i * 2, 2), 16)) <= 1);
            }
            Assert.True((observed.TryGetValue(key, out var found) && found == value) ||
                closeThemeColor ||
                (!observed.ContainsKey(key) && value == "false"),
                $"Run formatting {key}: expected {value}, observed " +
                (observed.TryGetValue(key, out var current) ? current : "<missing>") +
                "; related " + string.Join(", ", observed.Where(x =>
                    x.Key.StartsWith(key[..key.LastIndexOf('.')],
                        StringComparison.Ordinal)).Select(x =>
                    x.Key + "=" + x.Value)));
        }
        var visibleEffects = new HashSet<string>(StringComparer.Ordinal)
        {
            "doubleStrike", "shadow", "outline", "emboss", "imprint",
            "caps", "smallCaps", "hidden"
        };
        foreach (var (key, value) in observed.Where(x => x.Value == "true"))
        {
            var property = key[(key.LastIndexOf('.') + 1)..];
            if (visibleEffects.Contains(property) && !expected.ContainsKey(key))
                Assert.Fail($"Unexpected enabled run effect {key}.");
        }
        foreach (var (key, value) in observed.Where(x =>
            (x.Key.EndsWith(".underline", StringComparison.Ordinal) &&
                x.Value != "none") ||
            x.Key.EndsWith(".highlight", StringComparison.Ordinal) ||
            x.Key.EndsWith(".shading.fill", StringComparison.Ordinal) ||
            (x.Key.EndsWith(".border.val", StringComparison.Ordinal) &&
                x.Value != "none")))
            if (!expected.ContainsKey(key))
                Assert.Fail($"Unexpected run decoration {key}: {value}");
        foreach (var (key, value) in observed.Where(x =>
            (x.Key.EndsWith(".baseline", StringComparison.Ordinal) &&
                x.Value != "baseline") ||
            (x.Key.EndsWith(".baselineOffset", StringComparison.Ordinal) &&
                x.Value != "0")))
            if (!expected.ContainsKey(key))
                Assert.Fail($"Unexpected run positioning {key}: {value}");
        foreach (var (key, value) in observed.Where(x =>
            x.Key.EndsWith(".characterStyle", StringComparison.Ordinal)))
            Assert.True(expected.TryGetValue(key, out var original) && original == value,
                $"Unexpected character style at {key}: {value}");
    }

    private static Dictionary<string, string> ReadEffectiveRunFormatting(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var styles = main.StyleDefinitionsPart?.Styles;
        var definitions = styles?.Elements<Style>()
            .Where(x => x.StyleId?.Value != null)
            .ToDictionary(x => x.StyleId!.Value!, StringComparer.Ordinal) ??
            new Dictionary<string, Style>(StringComparer.Ordinal);
        var defaultParagraphStyle = styles?.Elements<Style>()
            .FirstOrDefault(x => x.Type?.Value == StyleValues.Paragraph &&
                x.Default?.Value == true)?.StyleId?.Value;
        var defaults = styles?.Elements<DocDefaults>().FirstOrDefault()?
            .GetFirstChild<RunPropertiesDefault>()?
            .GetFirstChild<RunPropertiesBaseStyle>();
        var themeColors = new Dictionary<string, string>(StringComparer.Ordinal);
        string? majorLatinFont = null;
        string? minorLatinFont = null;
        var themeMapping = main.DocumentSettingsPart?.Settings?
            .GetFirstChild<ColorSchemeMapping>()?.GetAttributes()
            .Where(x => x.LocalName.StartsWith("accent", StringComparison.Ordinal))
            .ToDictionary(x => x.LocalName, x => x.Value,
                StringComparer.Ordinal) ?? new Dictionary<string, string>(StringComparer.Ordinal);
        if (main.ThemePart is { } themePart)
        {
            using var themeStream = themePart.GetStream(FileMode.Open, FileAccess.Read);
            var theme = XDocument.Load(themeStream);
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            majorLatinFont = theme.Descendants(a + "majorFont").FirstOrDefault()?
                .Element(a + "latin")?.Attribute("typeface")?.Value;
            minorLatinFont = theme.Descendants(a + "minorFont").FirstOrDefault()?
                .Element(a + "latin")?.Attribute("typeface")?.Value;
            foreach (var accent in theme.Descendants(a + "clrScheme")
                .Elements().Where(x => x.Name.LocalName.StartsWith("accent",
                    StringComparison.Ordinal)))
            {
                var rgb = accent.Descendants(a + "srgbClr").FirstOrDefault()?
                    .Attribute("val")?.Value ?? accent.Descendants(a + "sysClr")
                    .FirstOrDefault()?.Attribute("lastClr")?.Value;
                if (rgb is { Length: 6 }) themeColors[accent.Name.LocalName] = rgb;
            }
        }
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        static string ApplyThemeLightness(string rgb, string modifier, bool tint)
        {
            var amount = Convert.ToByte(modifier, 16) / 255d;
            var channels = Enumerable.Range(0, 3)
                .Select(i => Convert.ToInt32(rgb.Substring(i * 2, 2), 16) / 255d)
                .ToArray();
            var maximum = channels.Max();
            var minimum = channels.Min();
            var delta = maximum - minimum;
            var lightness = (maximum + minimum) / 2;
            var saturation = delta == 0 ? 0 :
                delta / (1 - Math.Abs(2 * lightness - 1));
            var hue = delta == 0 ? 0 : maximum == channels[0]
                ? ((channels[1] - channels[2]) / delta + 6) % 6
                : maximum == channels[1]
                    ? (channels[2] - channels[0]) / delta + 2
                    : (channels[0] - channels[1]) / delta + 4;
            lightness = tint ? lightness * amount + 1 - amount :
                lightness * amount;
            var chroma = (1 - Math.Abs(2 * lightness - 1)) * saturation;
            var secondary = chroma * (1 - Math.Abs(hue % 2 - 1));
            var middle = lightness - chroma / 2;
            var components = hue switch
            {
                < 1 => new[] { chroma, secondary, 0d },
                < 2 => new[] { secondary, chroma, 0d },
                < 3 => new[] { 0d, chroma, secondary },
                < 4 => new[] { 0d, secondary, chroma },
                < 5 => new[] { secondary, 0d, chroma },
                _ => new[] { chroma, 0d, secondary }
            };
            return string.Concat(components.Select(x =>
                Math.Clamp((int)Math.Round((x + middle) * 255,
                    MidpointRounding.AwayFromZero), 0, 255).ToString("X2")));
        }
        void Apply(DocumentFormat.OpenXml.OpenXmlElement? properties,
            IDictionary<string, string> values)
        {
            if (properties == null) return;
            void SetToggle(string name, bool value)
            {
                values[name] = value ? "true" : "false";
            }
            if (properties.GetFirstChild<Bold>() is { } bold)
                SetToggle("bold", bold.Val?.Value ?? true);
            if (properties.GetFirstChild<Italic>() is { } italic)
                SetToggle("italic", italic.Val?.Value ?? true);
            if (properties.GetFirstChild<BoldComplexScript>() is { } complexBold)
                SetToggle("complexBold", complexBold.Val?.Value ?? true);
            if (properties.GetFirstChild<ItalicComplexScript>() is { } complexItalic)
                SetToggle("complexItalic", complexItalic.Val?.Value ?? true);
            if (properties.GetFirstChild<RightToLeftText>() is { } rightToLeft)
                SetToggle("rtl", rightToLeft.Val?.Value ?? true);
            if (properties.GetFirstChild<ComplexScript>() is { } complexScript)
                SetToggle("complexScript", complexScript.Val?.Value ?? true);
            if (properties.GetFirstChild<Caps>() is { } caps)
                SetToggle("caps", caps.Val?.Value ?? true);
            if (properties.GetFirstChild<SmallCaps>() is { } smallCaps)
                SetToggle("smallCaps", smallCaps.Val?.Value ?? true);
            if (properties.GetFirstChild<Strike>() is { } strike)
                SetToggle("strike", strike.Val?.Value ?? true);
            if (properties.GetFirstChild<DoubleStrike>() is { } doubleStrike)
                SetToggle("doubleStrike", doubleStrike.Val?.Value ?? true);
            if (properties.GetFirstChild<Shadow>() is { } shadow)
                SetToggle("shadow", shadow.Val?.Value ?? true);
            if (properties.GetFirstChild<Outline>() is { } outline)
                SetToggle("outline", outline.Val?.Value ?? true);
            if (properties.GetFirstChild<Emboss>() is { } emboss)
                SetToggle("emboss", emboss.Val?.Value ?? true);
            if (properties.GetFirstChild<Imprint>() is { } imprint)
                SetToggle("imprint", imprint.Val?.Value ?? true);
            if (properties.GetFirstChild<Vanish>() is { } vanish)
                SetToggle("hidden", vanish.Val?.Value ?? true);
            if (properties.GetFirstChild<SnapToGrid>() is { } snap)
                SetToggle("snapToGrid", snap.Val?.Value ?? true);
            if (properties.GetFirstChild<Emphasis>()?.Val is { } emphasis)
                values["emphasis"] = emphasis.InnerText;
            var fonts = properties.GetFirstChild<RunFonts>();
            var themeFont = fonts?.AsciiTheme?.Value;
            var themedFont = themeFont == ThemeFontValues.MajorAscii ||
                themeFont == ThemeFontValues.MajorHighAnsi ? majorLatinFont :
                themeFont == ThemeFontValues.MinorAscii ||
                themeFont == ThemeFontValues.MinorHighAnsi ? minorLatinFont : null;
            if ((themedFont ?? fonts?.Ascii?.Value) is string font)
                values["font"] = font;
            if (properties.GetFirstChild<Underline>() is { } underline)
            {
                if (underline.Val?.Value is { } underlineValue)
                    values["underline"] = underlineValue.ToString();
                if (underline.Color?.Value is string underlineColor)
                    values["underline.color"] = underlineColor.ToUpperInvariant();
            }
            if (properties.GetFirstChild<Border>() is { } border)
            {
                var borderValue = border.Val?.Value;
                if (borderValue == BorderValues.Nil || borderValue == BorderValues.None)
                    values["border.val"] = "none";
                else
                {
                    values["border.val"] = borderValue?.ToString() ?? "";
                    values["border.size"] = border.Size?.Value.ToString() ?? "";
                    values["border.space"] = border.Space?.Value.ToString() ?? "";
                    values["border.color"] = border.Color?.Value ?? "";
                    values["border.shadow"] =
                        (border.Shadow?.Value ?? false).ToString();
                    values["border.frame"] =
                        (border.Frame?.Value ?? false).ToString();
                }
            }
            if (properties.GetFirstChild<Shading>() is { } shading)
            {
                // A direct shd replaces the complete inherited shade. Word
                // may write an explicit nil or clear/auto for the same
                // unshaded appearance.
                foreach (var key in values.Keys.Where(x =>
                    x.StartsWith("shading.", StringComparison.Ordinal)).ToArray())
                    values.Remove(key);
                var pattern = shading.Val?.InnerText;
                var fill = shading.Fill?.Value;
                if (pattern == "nil" ||
                    (pattern == "clear" &&
                        (fill == null || fill.Equals("auto",
                            StringComparison.OrdinalIgnoreCase))))
                    values["shading.none"] = "true";
                else
                {
                    if (pattern != null) values["shading.pattern"] = pattern;
                    if (fill != null && !fill.Equals("auto",
                        StringComparison.OrdinalIgnoreCase))
                        values["shading.fill"] = fill.ToUpperInvariant();
                    if (shading.Color?.Value is string foreground &&
                        !foreground.Equals("auto", StringComparison.OrdinalIgnoreCase))
                        values["shading.color"] = foreground.ToUpperInvariant();
                }
            }
            if (properties.GetFirstChild<Highlight>()?.Val is { } highlight)
            {
                if (highlight.Value == HighlightColorValues.None)
                    values.Remove("highlight");
                else values["highlight"] = highlight.InnerText;
            }
            if (properties.GetFirstChild<VerticalTextAlignment>()?.Val != null)
                values["baseline"] = properties.GetFirstChild<VerticalTextAlignment>()!
                    .Val!.InnerText;
            if (properties.GetFirstChild<Languages>() is { } languages)
            {
                if (languages.Val?.Value is string language)
                    values["language"] = language;
                if (languages.EastAsia?.Value is string eastAsia)
                    values["language.eastAsia"] = eastAsia;
                if (languages.Bidi?.Value is string bidi)
                    values["language.bidi"] = bidi;
            }
            if (properties.GetFirstChild<Spacing>()?.Val?.Value is int spacing)
                values["spacing"] = spacing.ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (properties.GetFirstChild<Kern>()?.Val?.Value is uint kern)
                values["kern"] = kern.ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (properties.GetFirstChild<CharacterScale>()?.Val?.Value is long scale)
                values["characterScale"] = scale.ToString(
                    System.Globalization.CultureInfo.InvariantCulture);
            if (properties.GetFirstChild<Position>()?.Val?.Value is string position)
                values["baselineOffset"] = position;
            if (properties.GetFirstChild<FontSize>()?.Val?.Value is string size)
                values["size"] = size;
            if (properties.GetFirstChild<FontSizeComplexScript>()?.Val?.Value is
                string complexSize)
                values["complexSize"] = complexSize;
            if (properties.GetFirstChild<Color>() is { } color)
            {
                var key = color.ThemeColor?.Value switch
                {
                    var value when value == ThemeColorValues.Accent1 => "accent1",
                    var value when value == ThemeColorValues.Accent2 => "accent2",
                    var value when value == ThemeColorValues.Accent3 => "accent3",
                    var value when value == ThemeColorValues.Accent4 => "accent4",
                    var value when value == ThemeColorValues.Accent5 => "accent5",
                    var value when value == ThemeColorValues.Accent6 => "accent6",
                    _ => null
                };
                if (key != null)
                {
                    var visited = new HashSet<string>(StringComparer.Ordinal);
                    while (visited.Add(key) && themeMapping.TryGetValue(key, out var mapped) &&
                        mapped != key)
                        key = mapped;
                }
                var rgb = key != null && themeColors.TryGetValue(key, out var themed)
                    ? themed : color.Val?.Value;
                if (rgb is { Length: 6 } && key != null)
                {
                    if (color.ThemeTint?.Value is string tint)
                        rgb = ApplyThemeLightness(rgb, tint, true);
                    else if (color.ThemeShade?.Value is string shade)
                        rgb = ApplyThemeLightness(rgb, shade, false);
                }
                if (rgb is { Length: 6 }) values["color"] = rgb.ToUpperInvariant();
            }
        }
        void ApplyStyle(string? id, IDictionary<string, string> values,
            ISet<string> seen)
        {
            if (id == null || !seen.Add(id) ||
                !definitions.TryGetValue(id, out var style)) return;
            ApplyStyle(style.BasedOn?.Val?.Value, values, seen);
            Apply(style.StyleRunProperties, values);
        }
        void AddStory(string key, DocumentFormat.OpenXml.OpenXmlElement story)
        {
            var paragraphs = story.Descendants<Paragraph>()
                .Where(x => x.Descendants<Text>().Any(t => t.Text.Length != 0))
                .ToArray();
            for (var paragraphIndex = 0; paragraphIndex < paragraphs.Length;
                paragraphIndex++)
            {
                var paragraph = paragraphs[paragraphIndex];
                var position = 0;
                foreach (var run in paragraph.Descendants<Run>())
                {
                    var text = string.Concat(run.Descendants<Text>().Select(x => x.Text));
                    if (text.Length == 0) continue;
                    var values = new Dictionary<string, string>(StringComparer.Ordinal);
                    Apply(defaults, values);
                    var paragraphStyle = paragraph.ParagraphProperties?
                        .ParagraphStyleId?.Val?.Value ?? defaultParagraphStyle;
                    ApplyStyle(paragraphStyle, values,
                        new HashSet<string>(StringComparer.Ordinal));
                    if (paragraph.Ancestors<Table>().FirstOrDefault() is { } table &&
                        paragraph.Ancestors<TableRow>().FirstOrDefault() is { } row)
                    {
                        var look = table.TableProperties?.GetFirstChild<TableLook>();
                        ushort.TryParse(look?.Val?.Value,
                            System.Globalization.NumberStyles.HexNumber,
                            System.Globalization.CultureInfo.InvariantCulture,
                            out var lookMask);
                        if (look == null) lookMask = 0x04A0;
                        var attributeOnlyLook = look != null && look.Val?.Value == null;
                        var firstRowLook = look?.FirstRow?.Value ??
                            (lookMask & 0x0020) != 0;
                        var lastRowLook = look?.LastRow?.Value ??
                            (lookMask & 0x0040) != 0;
                        var firstColumnLook = look?.FirstColumn?.Value ??
                            (attributeOnlyLook || (lookMask & 0x0080) != 0);
                        var lastColumnLook = look?.LastColumn?.Value ??
                            (attributeOnlyLook || (lookMask & 0x0100) != 0);
                        void ApplyRowStyle(TableStyleOverrideValues condition)
                        {
                            var seenTableStyles = new HashSet<string>(StringComparer.Ordinal);
                            void ApplyCondition(string? id)
                            {
                                if (id == null || !seenTableStyles.Add(id) ||
                                    !definitions.TryGetValue(id, out var style)) return;
                                ApplyCondition(style.BasedOn?.Val?.Value);
                                var conditional = style.Elements<TableStyleProperties>()
                                    .FirstOrDefault(x => x.Type?.Value == condition);
                                Apply(conditional?.ChildElements.FirstOrDefault(x =>
                                    x.LocalName == "rPr"), values);
                            }
                            ApplyCondition(table.TableProperties?.TableStyle?.Val?.Value);
                        }
                        var rows = table.Elements<TableRow>();
                        if (firstRowLook &&
                            ReferenceEquals(row, rows.FirstOrDefault()))
                            ApplyRowStyle(TableStyleOverrideValues.FirstRow);
                        if (lastRowLook &&
                            ReferenceEquals(row, rows.LastOrDefault()))
                            ApplyRowStyle(TableStyleOverrideValues.LastRow);
                        var cells = row.Elements<TableCell>();
                        var cell = paragraph.Ancestors<TableCell>().FirstOrDefault();
                        if (firstColumnLook &&
                            ReferenceEquals(cell, cells.FirstOrDefault()))
                            ApplyRowStyle(TableStyleOverrideValues.FirstColumn);
                        if (lastColumnLook &&
                            ReferenceEquals(cell, cells.LastOrDefault()))
                            ApplyRowStyle(TableStyleOverrideValues.LastColumn);
                        // Corner rules are more specific than their intersecting
                        // row and column rules. Apply them before the run layers.
                        var firstRow = ReferenceEquals(row, rows.FirstOrDefault());
                        var lastRow = ReferenceEquals(row, rows.LastOrDefault());
                        var firstColumn = ReferenceEquals(cell, cells.FirstOrDefault());
                        var lastColumn = ReferenceEquals(cell, cells.LastOrDefault());
                        if (firstRow && firstColumn && firstRowLook &&
                            firstColumnLook)
                            ApplyRowStyle(TableStyleOverrideValues.NorthWestCell);
                        if (firstRow && lastColumn && firstRowLook &&
                            lastColumnLook)
                            ApplyRowStyle(TableStyleOverrideValues.NorthEastCell);
                        if (lastRow && firstColumn && lastRowLook &&
                            firstColumnLook)
                            ApplyRowStyle(TableStyleOverrideValues.SouthWestCell);
                        if (lastRow && lastColumn && lastRowLook &&
                            lastColumnLook)
                            ApplyRowStyle(TableStyleOverrideValues.SouthEastCell);
                    }
                    var characterStyle = run.RunProperties?.RunStyle?.Val?.Value;
                    ApplyStyle(characterStyle, values,
                        new HashSet<string>(StringComparer.Ordinal));
                    Apply(run.RunProperties, values);
                    values.TryAdd("snapToGrid", "true");
                    values.TryAdd("emphasis", "none");
                    if (characterStyle != null &&
                        definitions.TryGetValue(characterStyle, out var appliedStyle) &&
                        appliedStyle.CustomStyle?.Value == true)
                    {
                        var styleName = appliedStyle.StyleName?.Val?.Value ?? characterStyle;
                        if (styleName is not ("Hyperlink" or "Followed Hyperlink"))
                            values["characterStyle"] = styleName;
                    }
                    for (var offset = 0; offset < text.Length; offset++)
                        foreach (var (property, value) in values)
                            result[$"{key}.paragraph{paragraphIndex}.char{position + offset}.{property}"] =
                                value;
                    position += text.Length;
                }
            }
        }
        AddStory("body", main.Document!.Body!);
        var sections = main.Document.Body.Descendants<SectionProperties>().ToArray();
        for (var i = 0; i < sections.Length; i++)
        {
            foreach (var reference in sections[i].Elements<HeaderReference>())
                AddStory($"section{i}.header.{Slot(reference.Type?.Value)}",
                    ((HeaderPart)main.GetPartById(reference.Id!)).Header!);
            foreach (var reference in sections[i].Elements<FooterReference>())
                AddStory($"section{i}.footer.{Slot(reference.Type?.Value)}",
                    ((FooterPart)main.GetPartById(reference.Id!)).Footer!);
        }
        return result;
    }

    private static Dictionary<string, string> ReadStories(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var stories = new Dictionary<string, string>(StringComparer.Ordinal);
        var customStyleNames = main.StyleDefinitionsPart?.Styles?
            .Elements<Style>()
            .Where(x => x.CustomStyle?.Value == true && x.StyleId?.Value != null)
            .ToDictionary(x => x.StyleId!.Value!,
                x => x.StyleName?.Val?.Value ?? x.StyleId!.Value!,
                StringComparer.Ordinal) ?? new Dictionary<string, string>(StringComparer.Ordinal);
        void AddStory(string key, DocumentFormat.OpenXml.OpenXmlElement element)
        {
            stories[key] = TextOf(element);
            var drawings = element.Descendants<Drawing>().Count();
            if (drawings != 0) stories[key + ".drawings"] = drawings.ToString();
            var appliedStyles = element.Descendants<Paragraph>()
                .Select(x => x.ParagraphProperties?.ParagraphStyleId?.Val?.Value)
                .Select(id => id != null && customStyleNames.TryGetValue(id, out var style)
                    ? style : null)
                .Where(x => x != null).ToArray();
            if (appliedStyles.Length != 0)
                stories[key + ".customStyles"] = string.Join("\n", appliedStyles);
            var appliedCharacterStyles = element.Descendants<Run>()
                .Select(x => x.RunProperties?.RunStyle?.Val?.Value)
                .Where(id => id is not ("Hyperlink" or "FollowedHyperlink"))
                .Select(id => id != null && customStyleNames.TryGetValue(id, out var style)
                    ? style : null)
                .Where(x => x != null && x is not ("Hyperlink" or
                    "Followed Hyperlink"))
                .Distinct(StringComparer.Ordinal).ToArray();
            if (appliedCharacterStyles.Length != 0)
                stories[key + ".customCharacterStyles"] =
                    string.Join("\n", appliedCharacterStyles);
        }
        AddStory("body", main.Document!.Body!);
        var sections = main.Document.Body.Descendants<SectionProperties>().ToArray();
        for (var i = 0; i < sections.Length; i++)
        {
            foreach (var reference in sections[i].Elements<HeaderReference>())
            {
                var part = (HeaderPart)main.GetPartById(reference.Id!);
                if (TextOf(part.Header!).Length != 0 ||
                    part.Header.Descendants<Drawing>().Any())
                    AddStory($"section{i}.header.{Slot(reference.Type?.Value)}", part.Header);
            }
            foreach (var reference in sections[i].Elements<FooterReference>())
            {
                var part = (FooterPart)main.GetPartById(reference.Id!);
                if (TextOf(part.Footer!).Length != 0 ||
                    part.Footer.Descendants<Drawing>().Any())
                    AddStory($"section{i}.footer.{Slot(reference.Type?.Value)}", part.Footer);
            }
        }
        return stories;
    }

    private static string TextOf(DocumentFormat.OpenXml.OpenXmlElement story) =>
        string.Join("\n", story.Descendants<Paragraph>().Select(paragraph =>
        {
            var text = new System.Text.StringBuilder();
            var hasContentOnLine = false;
            foreach (var child in paragraph.Descendants())
            {
                string value = child switch
                {
                    Text runText => runText.Text,
                    SymbolChar symbol => $"[symbol:{symbol.Font?.Value}:{symbol.Char?.Value}]",
                    PositionalTab tab when tab.Alignment?.Value ==
                        AbsolutePositionTabAlignmentValues.Left =>
                        hasContentOnLine ? "\u2028" : "",
                    PositionalTab => "\t",
                    TabChar => "\t",
                    Break lineBreak => lineBreak.Type?.Value == BreakValues.Page
                        ? "\f" : lineBreak.Type?.Value == BreakValues.Column
                            ? "\u000e" : "\u2028",
                    _ when child.LocalName == "cr" => "\u2028",
                    _ when child.LocalName == "noBreakHyphen" => "\u2011",
                    _ when child.LocalName == "softHyphen" => "\u00ad",
                    _ => ""
                };
                text.Append(value);
                if (value.Length != 0)
                    hasContentOnLine = value.EndsWith("\u2028", StringComparison.Ordinal) ||
                        value.EndsWith("\f", StringComparison.Ordinal) ||
                        value.EndsWith("\u000e", StringComparison.Ordinal)
                        ? false : true;
            }
            return text.ToString();
        })).TrimEnd('\n');

    private static IReadOnlyList<string> ReadSupportedCoreProperties(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var properties = document.PackageProperties;
        return [properties.Title ?? "", properties.Subject ?? "",
            properties.Creator ?? "", properties.Keywords ?? "",
            properties.Description ?? ""];
    }

    private static void AssertExplicitHyphenationSettings(byte[] source, byte[] output)
    {
        using var stream = new MemoryStream(source);
        using var document = WordprocessingDocument.Open(stream, false);
        var settings = document.MainDocumentPart?.DocumentSettingsPart?.Settings;
        if (settings?.GetFirstChild<AutoHyphenation>() == null &&
            settings?.GetFirstChild<DoNotHyphenateCaps>() == null &&
            settings?.GetFirstChild<HyphenationZone>() == null &&
            settings?.GetFirstChild<ConsecutiveHyphenLimit>() == null)
            return;
        Assert.Equal(ReadHyphenationSettings(source), ReadHyphenationSettings(output));
    }
    private static (bool Automatic, bool HyphenateCaps, int ZoneTwips,
        ushort ConsecutiveLimit) ReadHyphenationSettings(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var settings = document.MainDocumentPart?.DocumentSettingsPart?.Settings;
        var automatic = settings?.GetFirstChild<AutoHyphenation>();
        var noCaps = settings?.GetFirstChild<DoNotHyphenateCaps>();
        var zoneText = settings?.GetFirstChild<HyphenationZone>()?.Val?.Value;
        return (automatic != null && (automatic.Val?.Value ?? true),
            noCaps == null || !(noCaps.Val?.Value ?? true),
            int.TryParse(zoneText, out var zone) ? zone : 0,
            settings?.GetFirstChild<ConsecutiveHyphenLimit>()?.Val?.Value ?? 0);
    }

    private static IReadOnlyList<string> ReadDocumentStatistics(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var properties = document.ExtendedFilePropertiesPart?.Properties;
        return [properties?.Pages?.Text ?? "", properties?.Words?.Text ?? "",
            properties?.Characters?.Text ?? ""];
    }

    private static IReadOnlyList<string> ReadEditableFieldSemantics(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var result = new List<string>();
        const string simpleIfPattern =
            @"\bIF\s+(\S+)\s*(=|<>|<=|>=|<|>)\s*(\S+)\s+""([^""]*)""\s+""([^""]*)""";
        foreach (var part in new OpenXmlPart[] { main }
            .Concat(main.HeaderParts).Concat(main.FooterParts))
        {
            var root = part.RootElement;
            if (root == null) continue;
            var relationships = part.HyperlinkRelationships
                .ToDictionary(x => x.Id, x => x.Uri.ToString(), StringComparer.Ordinal);
            foreach (var hyperlink in root.Descendants<Hyperlink>())
            {
                if (hyperlink.Anchor?.Value is string anchor)
                    result.Add("LINK:#" + anchor);
                if (hyperlink.Id?.Value is string id &&
                    relationships.TryGetValue(id, out var target))
                    result.Add("LINK:" + target.TrimEnd('/'));
            }
            foreach (var field in root.Descendants<SimpleField>())
                AddInstruction(field.Instruction?.Value);
            var instruction = string.Concat(root.Descendants<FieldCode>()
                .Select(x => x.Text));
            foreach (Match match in Regex.Matches(instruction,
                @"HYPERLINK\s+(?:\\l\s+)?""([^""]+)""|\bSTYLEREF\s+""([^""]+)""|\b(?:DOCPROPERTY|MERGEFIELD|SEQ|REF|PAGEREF)\s+[A-Za-z_][A-Za-z_0-9]*|\b(?:LASTSAVEDBY|AUTHOR|TITLE|SUBJECT|KEYWORDS|COMMENTS|SECTIONPAGES|SECTION|NUMPAGES|NUMWORDS|NUMCHARS|PAGE|DATE|TIME)\b|" + simpleIfPattern,
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                AddInstruction(match.Value);
        }
        return result.OrderBy(x => x, StringComparer.Ordinal).ToArray();

        void AddInstruction(string? instruction)
        {
            if (string.IsNullOrWhiteSpace(instruction)) return;
            var internalLink = Regex.Match(instruction,
                @"HYPERLINK\s+\\l\s+""([^""]+)""",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (internalLink.Success)
            {
                result.Add("LINK:#" + internalLink.Groups[1].Value);
                return;
            }
            var hyperlink = Regex.Match(instruction,
                @"HYPERLINK\s+""([^""]+)""",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (hyperlink.Success)
                result.Add("LINK:" + hyperlink.Groups[1].Value.TrimEnd('/'));
            else if (Regex.Match(instruction, simpleIfPattern,
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant) is
                { Success: true } conditional)
                result.Add("FIELD:IF:" + string.Join(":", Enumerable.Range(1, 5)
                    .Select(i => conditional.Groups[i].Value)));
            else if (Regex.Match(instruction,
                @"\bMERGEFIELD\s+([A-Za-z_][A-Za-z_0-9]*)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant) is
                { Success: true } mergeField)
                result.Add("FIELD:MERGEFIELD:" + mergeField.Groups[1].Value);
            else if (Regex.Match(instruction,
                @"\bDOCPROPERTY\s+([A-Za-z_][A-Za-z_0-9]*)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant) is
                { Success: true } docProperty)
                result.Add("FIELD:DOCPROPERTY:" + docProperty.Groups[1].Value);
            else if (Regex.Match(instruction,
                @"\bSTYLEREF\s+""([^""]+)""",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant) is { Success: true } styleReference)
                result.Add("FIELD:STYLEREF:" + styleReference.Groups[1].Value);
            else if (Regex.Match(instruction,
                @"\bSEQ\s+([A-Za-z_][A-Za-z_0-9]*)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant) is { Success: true } seq)
                result.Add("FIELD:SEQ:" + seq.Groups[1].Value);
            else if (Regex.Match(instruction,
                @"\bREF\s+([A-Za-z_][A-Za-z_0-9]*)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant) is { Success: true } reference)
                result.Add("FIELD:REF:" + reference.Groups[1].Value);
            else if (Regex.Match(instruction,
                @"\bPAGEREF\s+([A-Za-z_][A-Za-z_0-9]*)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant) is { Success: true } pageReference)
                result.Add("FIELD:PAGEREF:" + pageReference.Groups[1].Value);
            else if (Regex.IsMatch(instruction, @"\bSECTIONPAGES\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                result.Add("FIELD:SECTIONPAGES");
            else if (Regex.IsMatch(instruction, @"\bLASTSAVEDBY\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                result.Add("FIELD:LASTSAVEDBY");
            else if (Regex.IsMatch(instruction, @"\bAUTHOR\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                result.Add("FIELD:AUTHOR");
            else if (Regex.IsMatch(instruction, @"\bTITLE\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                result.Add("FIELD:TITLE");
            else if (Regex.IsMatch(instruction, @"\bSUBJECT\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                result.Add("FIELD:SUBJECT");
            else if (Regex.IsMatch(instruction, @"\bKEYWORDS\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                result.Add("FIELD:KEYWORDS");
            else if (Regex.IsMatch(instruction, @"\bCOMMENTS\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                result.Add("FIELD:COMMENTS");
            else if (Regex.IsMatch(instruction, @"\bNUMWORDS\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                result.Add("FIELD:NUMWORDS");
            else if (Regex.IsMatch(instruction, @"\bNUMCHARS\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                result.Add("FIELD:NUMCHARS");
            else if (Regex.IsMatch(instruction, @"\bSECTION\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                result.Add("FIELD:SECTION");
            else if (Regex.IsMatch(instruction, @"\bNUMPAGES\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                result.Add("FIELD:NUMPAGES");
            else if (Regex.IsMatch(instruction, @"\bPAGE\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                result.Add("FIELD:PAGE");
            else if (Regex.IsMatch(instruction, @"\bDATE\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                result.Add("FIELD:DATE");
            else if (Regex.IsMatch(instruction, @"\bTIME\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                result.Add("FIELD:TIME");
        }
    }

    private static bool ReadBalanceSetting(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var setting = document.MainDocumentPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<Compatibility>()?
            .GetFirstChild<BalanceSingleByteDoubleByteWidth>();
        return setting != null && (setting.Val?.Value ?? true);
    }

    private static IReadOnlyList<string> ReadDateTimeInstructions(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        return new DocumentFormat.OpenXml.Packaging.OpenXmlPart[] { main }
            .Concat(main.HeaderParts).Concat(main.FooterParts)
            .SelectMany(part => part.RootElement?.Descendants<FieldCode>() ??
                Enumerable.Empty<FieldCode>())
            .Select(field => Regex.Replace(field.Text.Trim(), @"\s+", " "))
            .Where(instruction => Regex.IsMatch(instruction,
                @"^(?:DATE|TIME)\b", RegexOptions.IgnoreCase |
                    RegexOptions.CultureInvariant))
            .OrderBy(instruction => instruction, StringComparer.Ordinal)
            .ToArray();
    }

    private static void AssertTableGrid(byte[] reference, byte[] actual)
    {
        var expected = ReadTableGrid(reference);
        var observed = ReadTableGrid(actual);
        foreach (var (key, value) in expected)
            Assert.True(observed.TryGetValue(key, out var found) && found == value,
                $"Table grid {key}: expected {value}, observed " +
                (observed.TryGetValue(key, out var current) ? current : "<missing>"));
        foreach (var (key, value) in observed.Where(x =>
            x.Key.EndsWith(".noWrap", StringComparison.Ordinal) ||
            x.Key.EndsWith(".fitText", StringComparison.Ordinal) ||
            x.Key.EndsWith(".preferredWidth", StringComparison.Ordinal) ||
            x.Key.EndsWith(".fixedLayout", StringComparison.Ordinal)))
            if (!expected.ContainsKey(key))
                Assert.Fail($"Unexpected table cell property {key}: {value}");
    }

    private static Dictionary<string, string> ReadTableGrid(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var implicitSideMargin = main.StyleDefinitionsPart == null ? 10 : 108;
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        static string? PreferredWidth(TableWidthType? width)
        {
            var value = width?.Width?.Value;
            if (value == null || value == "0" ||
                width?.Type?.Value == TableWidthUnitValues.Auto)
                return null;
            return $"{width.Type?.Value}:{value}";
        }
        void AddStory(string key, DocumentFormat.OpenXml.OpenXmlElement story)
        {
            var tables = story.Descendants<Table>().ToArray();
            result[$"{key}.tables"] = tables.Length.ToString();
            for (var tableIndex = 0; tableIndex < tables.Length; tableIndex++)
            {
                var rows = tables[tableIndex].Elements<TableRow>().ToArray();
                var prefix = $"{key}.table{tableIndex}";
                result[$"{prefix}.rows"] = rows.Length.ToString();
                var borderLayers = new List<TableBorders>();
                var tableStyleLayers = new List<Style>();
                var marginLayers = new List<DocumentFormat.OpenXml.OpenXmlElement>();
                if (tables[tableIndex].TableProperties?
                    .GetFirstChild<TableBorders>() is { } directBorders)
                    borderLayers.Add(directBorders);
                if (tables[tableIndex].TableProperties?
                    .GetFirstChild<TableCellMarginDefault>() is { } directMargins)
                    marginLayers.Add(directMargins);
                var styleId = tables[tableIndex].TableProperties?.TableStyle?.Val?.Value;
                var seenStyles = new HashSet<string>(StringComparer.Ordinal);
                while (styleId != null && seenStyles.Add(styleId))
                {
                    var style = main.StyleDefinitionsPart?.Styles?
                        .Elements<Style>().FirstOrDefault(x =>
                            x.Type?.Value == StyleValues.Table &&
                            x.StyleId?.Value == styleId);
                    if (style == null) break;
                    tableStyleLayers.Add(style);
                    if (style.StyleTableProperties?
                        .GetFirstChild<TableBorders>() is { } styleBorders)
                        borderLayers.Add(styleBorders);
                    if (style.StyleTableProperties?
                        .GetFirstChild<TableCellMarginDefault>() is { } styleMargins)
                        marginLayers.Add(styleMargins);
                    styleId = style.BasedOn?.Val?.Value;
                }
                if (PreferredWidth(tables[tableIndex].TableProperties?
                    .GetFirstChild<TableWidth>()) is { } tableWidth)
                    result[$"{prefix}.preferredWidth"] = tableWidth;
                if (tables[tableIndex].TableProperties?.GetFirstChild<TableLayout>()?
                    .Type?.Value == TableLayoutValues.Fixed)
                    result[$"{prefix}.fixedLayout"] = "1";
                var look = tables[tableIndex].TableProperties?
                    .GetFirstChild<TableLook>();
                var lookMask = ushort.TryParse(look?.Val?.Value,
                    System.Globalization.NumberStyles.HexNumber,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out var parsedLook) ? parsedLook : (ushort)0x04A0;
                var firstRowLook = look?.FirstRow?.Value ?? (lookMask & 0x0020) != 0;
                var firstColumnLook = look?.FirstColumn?.Value ??
                    (lookMask & 0x0080) != 0;
                var horizontalBandSize = look?.NoHorizontalBand?.Value == true ||
                    (look?.NoHorizontalBand == null && (lookMask & 0x0200) != 0)
                    ? 0 : tableStyleLayers.Select(x => x.StyleTableProperties?
                        .GetFirstChild<TableStyleRowBandSize>()?.Val?.Value)
                        .FirstOrDefault(x => x != null) ?? 0;
                var verticalBandSize = look?.NoVerticalBand?.Value == true ||
                    (look?.NoVerticalBand == null && (lookMask & 0x0400) != 0)
                    ? 0 : tableStyleLayers.Select(x => x.StyleTableProperties?
                        .GetFirstChild<TableStyleColumnBandSize>()?.Val?.Value)
                        .FirstOrDefault(x => x != null) ?? 0;
                var firstColumnBandOffset = firstColumnLook && tableStyleLayers.Any(x =>
                    x.Elements<TableStyleProperties>().Any(rule =>
                        rule.Type?.Value == TableStyleOverrideValues.FirstColumn)) ? 1 : 0;
                BorderType? ConditionalBandBorder(TableStyleOverrideValues kind,
                    Func<TableCellBorders, BorderType?> side) => tableStyleLayers
                    .SelectMany(x => x.Elements<TableStyleProperties>().Where(rule =>
                        rule.Type?.Value == kind))
                    .Select(rule => rule.Descendants<TableCellBorders>()
                        .Select(side).FirstOrDefault(edge => edge != null))
                    .FirstOrDefault(edge => edge != null);
                BorderType? TableBorder(Func<TableBorders, BorderType?> side) =>
                    borderLayers.Select(side).FirstOrDefault(x => x != null);
                static int? MarginSide(DocumentFormat.OpenXml.OpenXmlElement? margins,
                    string physical, string logical)
                {
                    var side = margins?.ChildElements.FirstOrDefault(x =>
                        x.LocalName == logical) ?? margins?.ChildElements.FirstOrDefault(x =>
                        x.LocalName == physical);
                    return int.TryParse(side?.GetAttribute("w",
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
                        .Value, out var value) ? value : null;
                }
                var parentTable = tables[tableIndex].Ancestors<Table>().FirstOrDefault();
                if (parentTable != null)
                    result[$"{prefix}.parentTable"] =
                        Array.IndexOf(tables, parentTable).ToString();
                if (tables[tableIndex].TableProperties?
                    .GetFirstChild<TableIndentation>()?.Width?.Value is int indentation &&
                    indentation != 0)
                    result[$"{prefix}.indent"] = indentation.ToString();
                for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
                {
                    var cells = rows[rowIndex].Elements<TableCell>().ToArray();
                    var rowBorderLayers = new List<TableBorders>();
                    var exceptionBorders = rows[rowIndex].TablePropertyExceptions?
                        .GetFirstChild<TableBorders>();
                    if (exceptionBorders != null)
                        rowBorderLayers.Add(exceptionBorders);
                    rowBorderLayers.AddRange(borderLayers);
                    var rowMarginLayers = new List<DocumentFormat.OpenXml.OpenXmlElement>();
                    if (rows[rowIndex].TablePropertyExceptions?
                        .GetFirstChild<TableCellMarginDefault>() is { } rowMargins)
                        rowMarginLayers.Add(rowMargins);
                    rowMarginLayers.AddRange(marginLayers);
                    BorderType? RowBorder(Func<TableBorders, BorderType?> side) =>
                        rowBorderLayers.Select(side).FirstOrDefault(x => x != null);
                    result[$"{prefix}.row{rowIndex}.cells"] = cells.Length.ToString();
                    var rowProperties = rows[rowIndex].TableRowProperties;
                    var rowSpacing = rowProperties?.GetFirstChild<TableCellSpacing>() ??
                        rows[rowIndex].TablePropertyExceptions?
                            .GetFirstChild<TableCellSpacing>() ??
                        tables[tableIndex].TableProperties?
                            .GetFirstChild<TableCellSpacing>() ??
                        tableStyleLayers.Select(x => x.StyleTableProperties?
                            .GetFirstChild<TableCellSpacing>())
                            .FirstOrDefault(x => x != null);
                    result[$"{prefix}.row{rowIndex}.cellSpacing"] =
                        rowSpacing?.Width?.Value ?? "0";
                    result[$"{prefix}.row{rowIndex}.gridBefore"] =
                        (rowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0)
                        .ToString();
                    result[$"{prefix}.row{rowIndex}.gridAfter"] =
                        (rowProperties?.GetFirstChild<GridAfter>()?.Val?.Value ?? 0)
                        .ToString();
                    if (rowProperties?.GetFirstChild<WidthBeforeTableRow>() is
                        { } widthBefore && widthBefore.Width?.Value is string beforeWidth &&
                        (beforeWidth != "0" || rowProperties?.GetFirstChild<GridBefore>()?.Val?.Value > 0))
                        result[$"{prefix}.row{rowIndex}.wBefore"] =
                            $"{widthBefore.Type?.Value}:{beforeWidth}";
                    if (rowProperties?.GetFirstChild<WidthAfterTableRow>() is
                        { } widthAfter && widthAfter.Width?.Value is string afterWidth &&
                        (afterWidth != "0" || rowProperties?.GetFirstChild<GridAfter>()?.Val?.Value > 0))
                        result[$"{prefix}.row{rowIndex}.wAfter"] =
                            $"{widthAfter.Type?.Value}:{afterWidth}";
                    if (rowProperties?.GetFirstChild<TableRowHeight>() is { } height)
                        result[$"{prefix}.row{rowIndex}.height"] =
                            $"{height.HeightType?.InnerText ?? "atLeast"}:{height.Val?.Value}";
                    if (rowProperties?.GetFirstChild<TableHeader>() != null)
                        result[$"{prefix}.row{rowIndex}.repeatHeader"] = "1";
                    if (rowProperties?.GetFirstChild<CantSplit>() != null)
                        result[$"{prefix}.row{rowIndex}.cantSplit"] = "1";
                    for (var cellIndex = 0; cellIndex < cells.Length; cellIndex++)
                    {
                        var cell = cells[cellIndex];
                        var cellPrefix = $"{prefix}.row{rowIndex}.cell{cellIndex}";
                        var cellMargins = cell.TableCellProperties?
                            .GetFirstChild<TableCellMargin>();
                        var topMargin = MarginSide(cellMargins, "top", "top") ??
                            rowMarginLayers.Select(x => MarginSide(x, "top", "top"))
                                .FirstOrDefault(x => x != null) ?? 0;
                        var leftMargin = MarginSide(cellMargins, "left", "start") ??
                            rowMarginLayers.Select(x => MarginSide(x, "left", "start"))
                                .FirstOrDefault(x => x != null) ??
                            (tables[tableIndex].TableProperties?.TableStyle != null
                                ? 0 : implicitSideMargin);
                        var rightMargin = MarginSide(cellMargins, "right", "end") ??
                            rowMarginLayers.Select(x => MarginSide(x, "right", "end"))
                                .FirstOrDefault(x => x != null) ??
                            (tables[tableIndex].TableProperties?.TableStyle != null
                                ? 0 : implicitSideMargin);
                        var bottomMargin = MarginSide(cellMargins, "bottom", "bottom") ??
                            rowMarginLayers.Select(x => MarginSide(x, "bottom", "bottom"))
                                .FirstOrDefault(x => x != null) ?? 0;
                        result[$"{cellPrefix}.margin.top"] = topMargin.ToString();
                        result[$"{cellPrefix}.margin.left"] = leftMargin.ToString();
                        result[$"{cellPrefix}.margin.right"] = rightMargin.ToString();
                        result[$"{cellPrefix}.margin.bottom"] = bottomMargin.ToString();
                        result[$"{cellPrefix}.text"] = TextOf(cell);
                        if (PreferredWidth(cell.TableCellProperties?
                            .GetFirstChild<TableCellWidth>()) is { } cellWidth)
                            result[$"{cellPrefix}.preferredWidth"] = cellWidth;
                        var direction = cell.TableCellProperties?
                            .GetFirstChild<TextDirection>()?.Val?.Value;
                        result[$"{cellPrefix}.textDirection"] = direction == null ||
                            direction == TextDirectionValues.LefToRightTopToBottom ||
                            direction == TextDirectionValues.LeftToRightTopToBottom2010
                                ? "lrTb" : direction.Value.ToString();
                        result[$"{cellPrefix}.span"] =
                            (cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1).ToString();
                        result[$"{cellPrefix}.merge"] =
                            cell.TableCellProperties?.VerticalMerge?.Val?.Value.ToString() ?? "none";
                        var fitText = cell.TableCellProperties?
                            .GetFirstChild<TableCellFitText>();
                        if (fitText != null && (fitText.Val == null ||
                            fitText.Val.Value == OnOffOnlyValues.On))
                            result[$"{cellPrefix}.fitText"] = "1";
                        var noWrap = cell.TableCellProperties?.GetFirstChild<NoWrap>();
                        if (noWrap != null && (noWrap.Val == null ||
                            noWrap.Val.Value == OnOffOnlyValues.On))
                            result[$"{cellPrefix}.noWrap"] = "1";
                        if (cell.TableCellProperties?.TableCellVerticalAlignment?.Val is
                            { } verticalAlignment)
                            result[$"{cellPrefix}.verticalAlign"] =
                                verticalAlignment.InnerText;
                        var shading = cell.TableCellProperties?.GetFirstChild<Shading>();
                        if (shading?.Fill?.Value is string fill &&
                            !fill.Equals("auto", StringComparison.OrdinalIgnoreCase))
                            result[$"{cellPrefix}.fill"] = fill.ToUpperInvariant();
                        if (shading?.Val != null &&
                            shading.Val.Value != ShadingPatternValues.Nil &&
                            !(shading.Val.Value == ShadingPatternValues.Clear &&
                                (shading.Fill?.Value == null ||
                                    shading.Fill.Value.Equals("auto",
                                        StringComparison.OrdinalIgnoreCase))))
                            result[$"{cellPrefix}.pattern"] = shading.Val.InnerText;
                        var borders = cell.TableCellProperties?
                            .GetFirstChild<TableCellBorders>();
                        var horizontalBand = horizontalBandSize > 0 &&
                            rowIndex >= (firstRowLook ? 1 : 0)
                            ? ((rowIndex - (firstRowLook ? 1 : 0)) /
                                horizontalBandSize) % 2 == 0
                                ? TableStyleOverrideValues.Band1Horizontal
                                : TableStyleOverrideValues.Band2Horizontal
                            : (TableStyleOverrideValues?)null;
                        var verticalBand = verticalBandSize > 0 &&
                            cellIndex >= firstColumnBandOffset
                            ? ((cellIndex - firstColumnBandOffset) /
                                verticalBandSize) % 2 == 0
                                ? TableStyleOverrideValues.Band1Vertical
                                : TableStyleOverrideValues.Band2Vertical
                            : (TableStyleOverrideValues?)null;
                        BorderType? BandBorder(Func<TableCellBorders, BorderType?> side) =>
                            (verticalBand is { } vertical
                                ? ConditionalBandBorder(vertical, side) : null) ??
                            (horizontalBand is { } horizontal
                                ? ConditionalBandBorder(horizontal, side) : null);
                        void AddBorder(string side, BorderType? border)
                        {
                            if (border?.Val != null)
                                result[$"{cellPrefix}.{side}.style"] =
                                    border.Val.Value == BorderValues.Nil ||
                                    border.Val.Value == BorderValues.None
                                        ? "none" : border.Val.InnerText;
                            if (border?.Size?.Value is uint size)
                                result[$"{cellPrefix}.{side}.size"] = size.ToString();
                            if (border?.Space?.Value is uint space)
                                result[$"{cellPrefix}.{side}.space"] = space.ToString();
                            if (border?.Color?.Value is string color)
                                result[$"{cellPrefix}.{side}.color"] = color.ToUpperInvariant();
                        }
                        AddBorder("top", (BorderType?)borders?.TopBorder ??
                            exceptionBorders?.TopBorder ??
                            BandBorder(x => x.TopBorder) ?? (rowIndex == 0
                                ? TableBorder(x => x.TopBorder) :
                                    RowBorder(x => x.InsideHorizontalBorder)));
                        AddBorder("left", (BorderType?)borders?.StartBorder ??
                            (BorderType?)borders?.LeftBorder ??
                            BandBorder(x => x.LeftBorder) ?? (cellIndex == 0
                            ? RowBorder(x => (BorderType?)x.StartBorder ?? x.LeftBorder) :
                                RowBorder(x => x.InsideVerticalBorder)));
                        AddBorder("bottom", (BorderType?)borders?.BottomBorder ??
                            exceptionBorders?.BottomBorder ??
                            BandBorder(x => x.BottomBorder) ??
                            (rowIndex == rows.Length - 1 ?
                                TableBorder(x => x.BottomBorder) :
                                RowBorder(x => x.InsideHorizontalBorder)));
                        AddBorder("right", (BorderType?)borders?.EndBorder ??
                            (BorderType?)borders?.RightBorder ??
                            BandBorder(x => x.RightBorder) ??
                            (cellIndex == cells.Length - 1 ?
                                RowBorder(x => (BorderType?)x.EndBorder ?? x.RightBorder) :
                                RowBorder(x => x.InsideVerticalBorder)));
                        AddBorder("tl2br", borders?.TopLeftToBottomRightCellBorder ??
                            BandBorder(x => x.TopLeftToBottomRightCellBorder));
                        AddBorder("tr2bl", borders?.TopRightToBottomLeftCellBorder ??
                            BandBorder(x => x.TopRightToBottomLeftCellBorder));
                    }
                }
            }
        }
        AddStory("body", main.Document!.Body!);
        var sections = main.Document.Body.Descendants<SectionProperties>().ToArray();
        for (var i = 0; i < sections.Length; i++)
        {
            foreach (var reference in sections[i].Elements<HeaderReference>())
                AddStory($"section{i}.header.{Slot(reference.Type?.Value)}",
                    ((HeaderPart)main.GetPartById(reference.Id!)).Header!);
            foreach (var reference in sections[i].Elements<FooterReference>())
                AddStory($"section{i}.footer.{Slot(reference.Type?.Value)}",
                    ((FooterPart)main.GetPartById(reference.Id!)).Footer!);
        }
        return result;
    }

    private static IReadOnlyList<string> ReadSectionBodyText(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var result = new List<string>();
        var paragraphs = new List<string>();
        foreach (var child in document.MainDocumentPart!.Document!.Body!.ChildElements)
        {
            if (child is Paragraph or Table)
                paragraphs.Add(TextOf(child));
            if (child is SectionProperties ||
                child is Paragraph paragraph &&
                paragraph.ParagraphProperties?.SectionProperties != null)
            {
                result.Add(string.Join("\n", paragraphs).TrimEnd('\n'));
                paragraphs.Clear();
            }
        }
        if (paragraphs.Count != 0) result.Add(string.Join("\n", paragraphs).TrimEnd('\n'));
        return result;
    }

    private static IReadOnlyList<string> ReadEffectiveHeaderFooterStories(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var result = new List<string>();
        var headers = new HeaderPart?[3];
        var footers = new FooterPart?[3];
        static int SlotIndex(HeaderFooterValues? value) =>
            value == HeaderFooterValues.Even ? 1 :
            value == HeaderFooterValues.First ? 2 : 0;
        foreach (var section in main.Document!.Body!.Descendants<SectionProperties>())
        {
            foreach (var reference in section.Elements<HeaderReference>())
                headers[SlotIndex(reference.Type?.Value)] =
                    (HeaderPart)main.GetPartById(reference.Id!);
            foreach (var reference in section.Elements<FooterReference>())
                footers[SlotIndex(reference.Type?.Value)] =
                    (FooterPart)main.GetPartById(reference.Id!);
            for (var slot = 0; slot < 3; slot++)
            {
                var header = headers[slot]?.Header;
                var footer = footers[slot]?.Footer;
                result.Add($"header{slot}:{(header == null ? "" : TextOf(header))}|" +
                    $"drawings:{header?.Descendants<Drawing>().Count() ?? 0}");
                result.Add($"footer{slot}:{(footer == null ? "" : TextOf(footer))}|" +
                    $"drawings:{footer?.Descendants<Drawing>().Count() ?? 0}");
            }
        }
        return result;
    }

    private static IReadOnlyList<int> ReadEffectiveFooterPageFieldCounts(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var result = new List<int>();
        var footers = new FooterPart?[3];
        foreach (var section in main.Document!.Body!.Descendants<SectionProperties>())
        {
            foreach (var reference in section.Elements<FooterReference>())
            {
                var slot = reference.Type?.Value == HeaderFooterValues.Even ? 1 :
                    reference.Type?.Value == HeaderFooterValues.First ? 2 : 0;
                footers[slot] = (FooterPart)main.GetPartById(reference.Id!);
            }
            foreach (var footer in footers)
            {
                var root = footer?.Footer;
                var simple = root?.Descendants<SimpleField>().Count(x =>
                    Regex.IsMatch(x.Instruction?.Value ?? "", @"\bPAGE\b",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)) ?? 0;
                var complex = root?.Descendants<FieldCode>().Count(x =>
                    Regex.IsMatch(x.Text ?? "", @"\bPAGE\b",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)) ?? 0;
                result.Add(simple + complex);
            }
        }
        return result;
    }

    private static IReadOnlyList<string> ReadEffectiveFooterHyperlinkTargets(byte[] bytes)
        => ReadEffectiveStoryHyperlinkTargets(bytes, header: false);

    private static IReadOnlyList<string> ReadEffectiveHeaderHyperlinkTargets(byte[] bytes)
        => ReadEffectiveStoryHyperlinkTargets(bytes, header: true);

    private static IReadOnlyList<string> ReadEffectiveStoryHyperlinkTargets(byte[] bytes,
        bool header)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var result = new List<string>();
        var parts = new OpenXmlPart?[3];
        foreach (var section in main.Document!.Body!.Descendants<SectionProperties>())
        {
            if (header)
            {
                foreach (var reference in section.Elements<HeaderReference>())
                {
                    var slot = reference.Type?.Value == HeaderFooterValues.Even ? 1 :
                        reference.Type?.Value == HeaderFooterValues.First ? 2 : 0;
                    parts[slot] = main.GetPartById(reference.Id!);
                }
            }
            else
            {
                foreach (var reference in section.Elements<FooterReference>())
                {
                    var slot = reference.Type?.Value == HeaderFooterValues.Even ? 1 :
                        reference.Type?.Value == HeaderFooterValues.First ? 2 : 0;
                    parts[slot] = main.GetPartById(reference.Id!);
                }
            }
            foreach (var part in parts)
            {
                var targets = new List<string>();
                if (part?.RootElement is { } root)
                {
                    var relationships = part.HyperlinkRelationships
                        .ToDictionary(x => x.Id, x => x.Uri.ToString(),
                            StringComparer.Ordinal);
                    foreach (var hyperlink in root.Descendants<Hyperlink>())
                    {
                        if (hyperlink.Anchor?.Value is string anchor)
                            targets.Add("#" + anchor);
                        if (hyperlink.Id?.Value is string id &&
                            relationships.TryGetValue(id, out var target))
                            targets.Add(target.TrimEnd('/'));
                    }
                    var instruction = string.Concat(root.Descendants<FieldCode>()
                        .Select(x => x.Text));
                    foreach (Match match in Regex.Matches(instruction,
                        @"HYPERLINK\s+\\l\s+""([^""]+)""",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                        targets.Add("#" + match.Groups[1].Value);
                    foreach (Match match in Regex.Matches(instruction,
                        "HYPERLINK\\s+\"([^\"]+)\"",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                        targets.Add(match.Groups[1].Value.TrimEnd('/'));
                }
                result.Add(string.Join("|", targets.Distinct(StringComparer.Ordinal)
                    .OrderBy(x => x, StringComparer.Ordinal)));
            }
        }
        return result;
    }

    [Theory]
    [InlineData("WordTableListsAllStories")]
    [InlineData("WordNestedTableListsAllStories")]
    public void IndexedListParagraphsRetainNumberedTableCellsInAllStories(string name)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            name + ".docx"));
        var source = File.ReadAllBytes(path);
        var generated = DxpDocExport.Export(source);
        var projected = DxpDocToDocx.Project(generated).DocxBytes;
        foreach (var bytes in new[] { File.ReadAllBytes(Path.ChangeExtension(path, ".doc")),
                     generated, DxpDocExport.Export(projected) })
        {
            using var stream = new MemoryStream(bytes);
            using var index = new DocTextIndexWalker().Index(stream);
            var stories = new[] { index.ReadStory("Main") }
                .Concat(index.HeaderStories.Select(x => x.ReadIndexedContent()))
                .ToArray();
            Assert.NotEmpty(stories[0].ListParagraphs);
            Assert.True(stories.Skip(1).Count(x => x.ListParagraphs.Count > 0) >= 2);
            foreach (var list in stories.SelectMany(x => x.ListParagraphs))
            {
                Assert.True(list.OverrideIndex > 0);
                Assert.NotNull(list.SourceCpStart);
                Assert.NotNull(list.SourceTextOffsetStart);
                Assert.Contains(index.ParagraphStyles, x =>
                    x.CpStart <= list.SourceCpStart &&
                    x.CpEnd >= list.SourceCpEnd &&
                    x.Formatting?.ListOverrideIndex == list.OverrideIndex &&
                    x.Formatting.InTable == true &&
                    x.Formatting.TableTerminator != true);
            }
        }
    }

    [Fact]
    public void IndexedListParagraphsBindToStoryPositionsAcrossAllStories()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordListsAllStories.docx"));
        foreach (var bytes in new[] { File.ReadAllBytes(Path.ChangeExtension(path, ".doc")),
                     DxpDocExport.Export(File.ReadAllBytes(path)) })
        {
            using var stream = new MemoryStream(bytes);
            using var index = new DocTextIndexWalker().Index(stream);
            var main = index.ReadStory("Main");
            Assert.NotEmpty(main.ListParagraphs);
            var headers = index.HeaderStories.Select(x => x.ReadIndexedContent()).ToArray();
            Assert.True(headers.Count(x => x.ListParagraphs.Count > 0) >= 2);
            foreach (var story in new[] { main }.Concat(headers))
                foreach (var list in story.ListParagraphs)
                {
                    Assert.InRange(list.Start, 0, checked((int)(story.Text.CpEnd -
                        story.Text.CpStart)) - 1);
                    Assert.True(list.End > list.Start);
                    Assert.Equal(story.Text.CpStart + (uint)list.Start,
                        list.SourceCpStart);
                    Assert.Equal(story.Text.CpStart + (uint)list.End,
                        list.SourceCpEnd);
                    Assert.Equal(index.GetTextOffset(list.SourceCpStart!.Value),
                        list.SourceTextOffsetStart);
                    Assert.Contains(story.ParagraphStyles, x =>
                        x.CpStart <= list.SourceCpStart &&
                        x.CpEnd >= list.SourceCpEnd &&
                        x.Formatting?.ListOverrideIndex == list.OverrideIndex &&
                        (x.Formatting.ListLevel ?? 0) == list.Level);
                }
        }
    }

    private static IReadOnlyList<DocumentFormat.OpenXml.Validation.ValidationErrorInfo>
        Validate(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return new OpenXmlValidator().Validate(document).ToArray();
    }
}
