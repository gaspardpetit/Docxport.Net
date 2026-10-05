using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Doc;
using DocxportNet.Walker.Context;
using System.Xml.Linq;
using Xunit.Abstractions;

namespace DocxportNet.Tests;

public class ThemeFontResolutionTests : TestBase<ThemeFontResolutionTests>
{
    private static readonly string ProjectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));

    public ThemeFontResolutionTests(ITestOutputHelper output) : base(output)
    {
    }

    [Fact]
    public void DefaultRunStyle_IncludesThemeLatinFont()
    {
        string path = Path.Combine(ProjectRoot, "samples", "with_breaking_style.docx");
        using var doc = WordprocessingDocument.Open(path, false);

        string? minorLatin = null;
        var themePart = doc.MainDocumentPart?.ThemePart;
        if (themePart != null)
        {
            using var s = themePart.GetStream(FileMode.Open, FileAccess.Read);
            var xdoc = XDocument.Load(s);
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            minorLatin = xdoc.Descendants(a + "fontScheme")
                .Elements(a + "minorFont")
                .Elements(a + "latin")
                .Attributes("typeface")
                .Select(a => a.Value)
                .FirstOrDefault();
        }

        var resolver = new DxpStyleResolver(doc);
        var style = resolver.GetDefaultRunStyle();

        Assert.False(string.IsNullOrWhiteSpace(minorLatin));
        Assert.False(string.IsNullOrWhiteSpace(style.FontName));
        Assert.Equal(minorLatin, style.FontName);
    }

    [Fact]
    public void DocRoundTripResolvesThemeLatinFontInDefaultsAndDirectRuns()
    {
        var path = Path.Combine(ProjectRoot, "samples", "with_breaking_style.docx");
        using var source = new MemoryStream(File.ReadAllBytes(path));
        string majorLatin;
        string minorLatin;
        using (var document = WordprocessingDocument.Open(source, true))
        {
            using var themeStream = document.MainDocumentPart!.ThemePart!
                .GetStream(FileMode.Open, FileAccess.Read);
            var theme = XDocument.Load(themeStream);
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            var scheme = theme.Descendants(a + "fontScheme").Single();
            majorLatin = scheme.Element(a + "majorFont")!.Element(a + "latin")!
                .Attribute("typeface")!.Value;
            minorLatin = scheme.Element(a + "minorFont")!.Element(a + "latin")!
                .Attribute("typeface")!.Value;
            var defaults = document.MainDocumentPart.StyleDefinitionsPart!.Styles!
                .GetFirstChild<DocDefaults>()!.RunPropertiesDefault!
                .RunPropertiesBaseStyle!;
            defaults.RemoveAllChildren<RunFonts>();
            defaults.AppendChild(new RunFonts { AsciiTheme = ThemeFontValues.MinorAscii,
                HighAnsiTheme = ThemeFontValues.MinorHighAnsi });
            var styled = document.MainDocumentPart.StyleDefinitionsPart.Styles
                .Elements<Style>().First(x => x.Type?.Value == StyleValues.Paragraph);
            styled.StyleRunProperties = new StyleRunProperties(new RunFonts
            {
                AsciiTheme = ThemeFontValues.MajorAscii,
                HighAnsiTheme = ThemeFontValues.MajorHighAnsi
            });
            var run = document.MainDocumentPart.Document!.Body!
                .Descendants<Run>().First(x => x.GetFirstChild<Text>() != null);
            run.RunProperties = new RunProperties(new RunFonts
            {
                Ascii = "Arial", AsciiTheme = ThemeFontValues.MajorAscii,
                HighAnsi = "Arial", HighAnsiTheme = ThemeFontValues.MajorHighAnsi
            });
            document.MainDocumentPart.StyleDefinitionsPart.Styles.Save();
            document.MainDocumentPart.Document.Save();
        }
        using (var reopened = WordprocessingDocument.Open(source, false))
        {
            var fonts = reopened.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .GetFirstChild<DocDefaults>()!.Descendants<RunFonts>().First();
            Assert.Equal(ThemeFontValues.MinorAscii, fonts.AsciiTheme!.Value);
        }
        var docBytes = DxpDocExport.Export(source.ToArray());
        using (var indexStream = new MemoryStream(docBytes))
        using (var index = new DocTextIndexWalker().Index(indexStream))
        {
            Assert.Contains(index.Fonts, x => x.Name == minorLatin);
            Assert.Contains(index.Fonts, x => x.Name == majorLatin);
            Assert.Equal(minorLatin, index.DefaultCharacterFormatting.AsciiFontName);
            Assert.Equal(minorLatin, index.DefaultCharacterFormatting.HighAnsiFontName);
            Assert.Contains(index.CharacterFormatting, x =>
                x.Formatting.AsciiFontName == majorLatin &&
                x.Formatting.HighAnsiFontName == majorLatin);
            Assert.Contains(index.StyleDefinitions, x =>
                x.Type == 1 && x.CharacterFormatting.AsciiFontName == majorLatin &&
                x.CharacterFormatting.HighAnsiFontName == majorLatin);
        }
        var projectedBytes = DxpDocToDocx.Project(docBytes).DocxBytes;
        using var projectedStream = new MemoryStream(projectedBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
        Assert.Equal(minorLatin, projected.MainDocumentPart!.StyleDefinitionsPart!
            .Styles!.GetFirstChild<DocDefaults>()!.Descendants<RunFonts>()
            .First().Ascii!.Value);
        Assert.Contains(projected.MainDocumentPart.Document!.Body!.Descendants<RunFonts>(),
            x => x.Ascii?.Value == majorLatin);
    }

    [Fact]
    public void DocRoundTripResolvesThemeEastAsianFontInDefaultsStylesAndRuns()
    {
        var path = Path.Combine(ProjectRoot, "samples", "with_breaking_style.docx");
        using var source = new MemoryStream();
        source.Write(File.ReadAllBytes(path));
        source.Position = 0;
        using (var document = WordprocessingDocument.Open(source, true))
        {
            var main = document.MainDocumentPart!;
            main.DocumentSettingsPart?.Settings?
                .GetFirstChild<ThemeFontLanguages>()?.Remove();
            main.DocumentSettingsPart?.Settings?.Save();
            var themePart = main.ThemePart!;
            XDocument theme;
            using (var stream = themePart.GetStream(FileMode.Open, FileAccess.Read))
                theme = XDocument.Load(stream);
            XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
            var scheme = theme.Descendants(drawing + "fontScheme").Single();
            scheme.Element(drawing + "majorFont")!.Element(drawing + "ea")!
                .SetAttributeValue("typeface", "Yu Gothic");
            scheme.Element(drawing + "minorFont")!.Element(drawing + "ea")!
                .SetAttributeValue("typeface", "MS Mincho");
            using (var stream = themePart.GetStream(FileMode.Create, FileAccess.Write))
                theme.Save(stream);

            var styles = main.StyleDefinitionsPart!.Styles!;
            var defaults = styles.GetFirstChild<DocDefaults>()!
                .RunPropertiesDefault!.RunPropertiesBaseStyle!;
            var defaultFonts = defaults.GetFirstChild<RunFonts>()!;
            defaultFonts.EastAsia = "Fallback Default";
            defaultFonts.EastAsiaTheme = ThemeFontValues.MinorEastAsia;
            var style = styles.Elements<Style>()
                .First(x => x.Type?.Value == StyleValues.Paragraph);
            style.StyleRunProperties = new StyleRunProperties(new RunFonts
            {
                EastAsia = "Fallback Style",
                EastAsiaTheme = ThemeFontValues.MajorEastAsia
            });
            var run = main.Document!.Body!.Descendants<Run>()
                .First(x => x.GetFirstChild<Text>() != null);
            run.RunProperties = new RunProperties(new RunFonts
            {
                EastAsia = "Fallback Run",
                EastAsiaTheme = ThemeFontValues.MajorEastAsia
            });
            styles.Save();
            main.Document.Save();
        }

        var docBytes = DxpDocExport.Export(source.ToArray());
        using (var stream = new MemoryStream(docBytes))
        using (var index = new DocTextIndexWalker().Index(stream))
        {
            Assert.Equal("MS Mincho",
                index.DefaultCharacterFormatting.EastAsiaFontName);
            Assert.Contains(index.StyleDefinitions, x =>
                x.Type == 1 && x.CharacterFormatting.EastAsiaFontName == "Yu Gothic");
            Assert.Contains(index.CharacterFormatting, x =>
                x.Formatting.EastAsiaFontName == "Yu Gothic");
        }
        using var projectedStream = new MemoryStream(DxpDocToDocx.Project(docBytes).DocxBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
        var outputStyles = projected.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Assert.Equal("MS Mincho", outputStyles.GetFirstChild<DocDefaults>()!
            .Descendants<RunFonts>().First().EastAsia!.Value);
        Assert.Contains(outputStyles.Elements<Style>(), x =>
            x.Type?.Value == StyleValues.Paragraph &&
            x.StyleRunProperties?.RunFonts?.EastAsia?.Value == "Yu Gothic");
        Assert.Contains(projected.MainDocumentPart.Document!.Body!.Descendants<RunFonts>(),
            x => x.EastAsia?.Value == "Yu Gothic");
    }

    [Theory]
    [InlineData("ja-JP", "Jpan")]
    [InlineData("ko-KR", "Hang")]
    [InlineData("zh-CN", "Hans")]
    [InlineData("zh-TW", "Hant")]
    public void DocRoundTripResolvesThemeEastAsianScriptFontFromSettings(
        string language, string script)
    {
        var path = Path.Combine(ProjectRoot, "samples", "with_breaking_style.docx");
        using var source = new MemoryStream();
        source.Write(File.ReadAllBytes(path));
        source.Position = 0;
        string majorFont;
        string minorFont;
        using (var document = WordprocessingDocument.Open(source, true))
        {
            var main = document.MainDocumentPart!;
            XDocument theme;
            using (var stream = main.ThemePart!.GetStream(FileMode.Open, FileAccess.Read))
                theme = XDocument.Load(stream);
            XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
            var scheme = theme.Descendants(drawing + "fontScheme").Single();
            string ScriptFont(string variant) => scheme.Element(drawing + variant)!
                .Elements(drawing + "font")
                .Single(x => (string?)x.Attribute("script") == script)
                .Attribute("typeface")!.Value;
            majorFont = ScriptFont("majorFont");
            minorFont = ScriptFont("minorFont");
            var settings = main.DocumentSettingsPart?.Settings;
            if (settings == null)
            {
                var part = main.AddNewPart<DocumentSettingsPart>();
                settings = part.Settings = new Settings();
            }
            settings.GetFirstChild<ThemeFontLanguages>()?.Remove();
            settings.AppendChild(new ThemeFontLanguages
            { Val = language, EastAsia = language });
            settings.Save();
            var styles = main.StyleDefinitionsPart!.Styles!;
            var defaults = styles.GetFirstChild<DocDefaults>()!
                .RunPropertiesDefault!.RunPropertiesBaseStyle!;
            var defaultFonts = defaults.GetFirstChild<RunFonts>()!;
            defaultFonts.EastAsia = "Fallback Default";
            defaultFonts.EastAsiaTheme = ThemeFontValues.MinorEastAsia;
            defaultFonts.Ascii = "Fallback Default";
            defaultFonts.AsciiTheme = ThemeFontValues.MinorAscii;
            var style = styles.Elements<Style>()
                .First(x => x.Type?.Value == StyleValues.Paragraph);
            style.StyleRunProperties = new StyleRunProperties(new RunFonts
            {
                EastAsia = "Fallback Style",
                EastAsiaTheme = ThemeFontValues.MajorEastAsia,
                Ascii = "Fallback Style",
                AsciiTheme = ThemeFontValues.MajorAscii
            });
            styles.Save();
            var run = main.Document!.Body!.Descendants<Run>()
                .First(x => x.GetFirstChild<Text>() != null);
            run.RunProperties = new RunProperties(new RunFonts
            {
                EastAsia = "Fallback Run",
                EastAsiaTheme = ThemeFontValues.MajorEastAsia,
                Ascii = "Fallback Run",
                AsciiTheme = ThemeFontValues.MajorAscii
            });
            var header = main.AddNewPart<HeaderPart>();
            header.Header = new Header(new Paragraph(new Run(
                new RunProperties(new RunFonts
                {
                    EastAsia = "Fallback Header",
                    EastAsiaTheme = ThemeFontValues.MajorEastAsia
                }), new Text("Header 漢字"))));
            header.Header.Save();
            var footer = main.AddNewPart<FooterPart>();
            footer.Footer = new Footer(new Paragraph(new Run(
                new RunProperties(new RunFonts
                {
                    EastAsia = "Fallback Footer",
                    EastAsiaTheme = ThemeFontValues.MinorEastAsia
                }), new Text("Footer 漢字"))));
            footer.Footer.Save();
            var section = main.Document.Body!.Descendants<SectionProperties>().Last();
            section.PrependChild(new FooterReference
            { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(footer) });
            section.PrependChild(new HeaderReference
            { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(header) });
            main.Document.Save();
        }
        var docBytes = DxpDocExport.Export(source.ToArray());
        using (var stream = new MemoryStream(docBytes))
        using (var index = new DocTextIndexWalker().Index(stream))
        {
            Assert.Equal(minorFont, index.DefaultCharacterFormatting.EastAsiaFontName);
            Assert.Equal(minorFont, index.DefaultCharacterFormatting.AsciiFontName);
            Assert.Contains(index.StyleDefinitions, x =>
                x.Type == 1 && x.CharacterFormatting.EastAsiaFontName == majorFont &&
                x.CharacterFormatting.AsciiFontName == majorFont);
            Assert.Contains(index.CharacterFormatting, x =>
                x.Formatting.EastAsiaFontName == majorFont &&
                x.Formatting.AsciiFontName == majorFont);
            var headerStory = index.HeaderStories[7];
            var footerStory = index.HeaderStories[9];
            Assert.False(headerStory.IsEmpty);
            Assert.False(footerStory.IsEmpty);
            Assert.Contains(index.CharacterFormatting, x =>
                x.CpStart >= headerStory.CpStart && x.CpStart < headerStory.CpEnd &&
                x.Formatting.EastAsiaFontName == majorFont);
            Assert.Contains(index.CharacterFormatting, x =>
                x.CpStart >= footerStory.CpStart && x.CpStart < footerStory.CpEnd &&
                x.Formatting.EastAsiaFontName == minorFont);
        }
        using var output = new MemoryStream(DxpDocToDocx.Project(docBytes).DocxBytes);
        using var projected = WordprocessingDocument.Open(output, false);
        Assert.Equal(minorFont, projected.MainDocumentPart!.StyleDefinitionsPart!
            .Styles!.GetFirstChild<DocDefaults>()!.Descendants<RunFonts>()
            .First().EastAsia!.Value);
        Assert.Equal(minorFont, projected.MainDocumentPart.StyleDefinitionsPart!
            .Styles!.GetFirstChild<DocDefaults>()!.Descendants<RunFonts>()
            .First().Ascii!.Value);
        Assert.Contains(projected.MainDocumentPart.Document!.Body!
            .Descendants<RunFonts>(), x => x.EastAsia?.Value == majorFont &&
                x.Ascii?.Value == majorFont);
        Assert.Contains(projected.MainDocumentPart.HeaderParts
            .SelectMany(x => x.Header!.Descendants<RunFonts>()),
            x => x.EastAsia?.Value == majorFont);
        Assert.Contains(projected.MainDocumentPart.FooterParts
            .SelectMany(x => x.Footer!.Descendants<RunFonts>()),
            x => x.EastAsia?.Value == minorFont);
        if (Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_WORD") != "1") return;
        var docPath = Path.Combine(Path.GetTempPath(),
            $"docxport-themed-script-{Guid.NewGuid():N}.doc");
        var docxPath = Path.ChangeExtension(docPath, ".docx");
        var nativePath = Path.Combine(Path.GetTempPath(),
            $"docxport-themed-script-native-{Guid.NewGuid():N}.doc");
        File.WriteAllBytes(docPath, docBytes);
        File.WriteAllBytes(docxPath, source.ToArray());
        dynamic word = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application")!)!;
        word.Visible = false;
        dynamic? opened = null;
        dynamic? original = null;
        dynamic? native = null;
        try
        {
            original = word.Documents.Open(docxPath, ReadOnly: true,
                AddToRecentFiles: false);
            original.SaveAs2(FileName: nativePath, FileFormat: 0);
            original.Close(false);
            original = null;
            native = word.Documents.Open(nativePath, ReadOnly: true,
                AddToRecentFiles: false);
            opened = word.Documents.Open(docPath, ReadOnly: true,
                AddToRecentFiles: false);
            Assert.Contains("Header 漢字",
                (string)opened.Sections.Item(1).Headers.Item(1).Range.Text);
            Assert.Contains("Footer 漢字",
                (string)opened.Sections.Item(1).Footers.Item(1).Range.Text);
            Assert.Equal((string)native.Sections.Item(1)
                .Headers.Item(1).Range.Characters.Item(8).Font.NameFarEast,
                (string)opened.Sections.Item(1)
                    .Headers.Item(1).Range.Characters.Item(8).Font.NameFarEast);
            Assert.Equal((string)native.Sections.Item(1)
                .Footers.Item(1).Range.Characters.Item(8).Font.NameFarEast,
                (string)opened.Sections.Item(1)
                    .Footers.Item(1).Range.Characters.Item(8).Font.NameFarEast);
        }
        finally
        {
            if (opened != null) opened.Close(false);
            if (native != null) native.Close(false);
            if (original != null) original.Close(false);
            word.Quit();
            File.Delete(docPath);
            File.Delete(docxPath);
            File.Delete(nativePath);
        }
    }

    [Theory]
    [InlineData("ar-SA", "Arab")]
    [InlineData("he-IL", "Hebr")]
    public void DocRoundTripPreservesComplexScriptThemeFonts(
        string language, string script)
    {
        var scriptText = script == "Hebr" ? "שלום" : "مرحبا";
        var path = Path.Combine(ProjectRoot, "samples", "with_breaking_style.docx");
        using var source = new MemoryStream();
        source.Write(File.ReadAllBytes(path));
        source.Position = 0;
        string majorFont;
        string minorFont;
        using (var document = WordprocessingDocument.Open(source, true))
        {
            var main = document.MainDocumentPart!;
            XDocument theme;
            using (var stream = main.ThemePart!.GetStream(FileMode.Open, FileAccess.Read))
                theme = XDocument.Load(stream);
            XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
            var scheme = theme.Descendants(drawing + "fontScheme").Single();
            string ScriptFont(string variant) => scheme.Element(drawing + variant)!
                .Elements(drawing + "font")
                .Single(x => (string?)x.Attribute("script") == script)
                .Attribute("typeface")!.Value;
            majorFont = ScriptFont("majorFont");
            minorFont = ScriptFont("minorFont");
            var settings = main.DocumentSettingsPart!.Settings!;
            settings.GetFirstChild<ThemeFontLanguages>()?.Remove();
            settings.AppendChild(new ThemeFontLanguages { Bidi = language });
            settings.Save();
            var styles = main.StyleDefinitionsPart!.Styles!;
            var defaults = styles.GetFirstChild<DocDefaults>()!
                .RunPropertiesDefault!.RunPropertiesBaseStyle!;
            defaults.GetFirstChild<RunFonts>()!.ComplexScriptTheme =
                ThemeFontValues.MinorBidi;
            defaults.AppendChild(new BoldComplexScript { Val = false });
            defaults.AppendChild(new ItalicComplexScript { Val = false });
            defaults.AppendChild(new FontSizeComplexScript { Val = "24" });
            var style = styles.Elements<Style>()
                .First(x => x.Type?.Value == StyleValues.Paragraph);
            style.StyleRunProperties = new StyleRunProperties(new RunFonts
            { ComplexScriptTheme = ThemeFontValues.MajorBidi },
                new BoldComplexScript(), new ItalicComplexScript { Val = false },
                new FontSizeComplexScript { Val = "30" });
            styles.AppendChild(new Style(new StyleName { Val = "OnlyComplex" },
                new StyleRunProperties(new RunFonts
                { ComplexScript = "Courier New" }))
            { Type = StyleValues.Character, StyleId = "OnlyComplex" });
            styles.Save();
            var run = main.Document!.Body!.Descendants<Run>()
                .First(x => x.GetFirstChild<Text>() != null);
            run.RunProperties = new RunProperties(new RunFonts
            { ComplexScriptTheme = ThemeFontValues.MajorBidi },
                new BoldComplexScript(), new ItalicComplexScript(),
                new FontSizeComplexScript { Val = "28" },
                new RightToLeftText(), new ComplexScript());
            var header = main.AddNewPart<HeaderPart>();
            header.Header = new Header(new Paragraph(
                new Run(new RunProperties(new RunFonts
                    { ComplexScriptTheme = ThemeFontValues.MajorBidi },
                    new BoldComplexScript(), new ItalicComplexScript(),
                    new FontSizeComplexScript { Val = "28" },
                    new RightToLeftText(), new ComplexScript()),
                    new Text("Header " + scriptText)),
                new Run(new RunProperties(new RunFonts
                    { ComplexScript = "Courier New" }),
                    new Text(" explicit"))));
            header.Header.Save();
            var footer = main.AddNewPart<FooterPart>();
            footer.Footer = new Footer(new Paragraph(new Run(new RunProperties(
                new RunFonts { ComplexScriptTheme = ThemeFontValues.MinorBidi },
                new BoldComplexScript { Val = false }, new ItalicComplexScript(),
                new FontSizeComplexScript { Val = "20" },
                new RightToLeftText { Val = false },
                new ComplexScript { Val = false }),
                new Text("Footer " + scriptText))));
            footer.Footer.Save();
            var section = main.Document.Body!.Descendants<SectionProperties>().Last();
            section.PrependChild(new FooterReference
            { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(footer) });
            section.PrependChild(new HeaderReference
            { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(header) });
            main.Document.Save();
        }

        var docBytes = DxpDocExport.Export(source.ToArray());
        using (var stream = new MemoryStream(docBytes))
        using (var index = new DocTextIndexWalker().Index(stream))
        {
            Assert.Equal(minorFont,
                index.DefaultCharacterFormatting.ComplexScriptFontName);
            var normal = Assert.Single(index.StyleDefinitions, x => x.Index == 0);
            Assert.False(normal.CharacterFormatting.ComplexScriptBold);
            Assert.False(normal.CharacterFormatting.ComplexScriptItalic);
            Assert.Equal((ushort)24,
                normal.CharacterFormatting.ComplexScriptSizeHalfPoints);
            Assert.Contains(index.StyleDefinitions, x => x.Type == 1 &&
                x.CharacterFormatting.ComplexScriptFontName == majorFont &&
                x.CharacterFormatting.ComplexScriptBold == true &&
                x.CharacterFormatting.ComplexScriptItalic == false &&
                x.CharacterFormatting.ComplexScriptSizeHalfPoints == 30);
            Assert.Contains(index.StyleDefinitions, x =>
                x.Name == "OnlyComplex" && x.Type == 2 &&
                x.CharacterFormatting.ComplexScriptFontName == "Courier New");
            Assert.Contains(index.CharacterFormatting, x =>
                x.Formatting.ComplexScriptFontName == majorFont &&
                x.Formatting.ComplexScriptBold == true &&
                x.Formatting.ComplexScriptItalic == true &&
                x.Formatting.ComplexScriptSizeHalfPoints == 28 &&
                x.Formatting.RightToLeftText == true &&
                x.Formatting.ForceComplexScript == true);
            var headerStory = index.HeaderStories[7];
            var footerStory = index.HeaderStories[9];
            Assert.Contains(index.CharacterFormatting, x =>
                x.CpStart >= headerStory.CpStart && x.CpStart < headerStory.CpEnd &&
                x.Formatting.ComplexScriptFontName == majorFont);
            Assert.Contains(index.CharacterFormatting, x =>
                x.CpStart >= headerStory.CpStart && x.CpStart < headerStory.CpEnd &&
                x.Formatting.ComplexScriptFontName == "Courier New");
            Assert.Contains(index.CharacterFormatting, x =>
                x.CpStart >= footerStory.CpStart && x.CpStart < footerStory.CpEnd &&
                x.Formatting.ComplexScriptFontName == minorFont &&
                x.Formatting.ComplexScriptBold == false &&
                x.Formatting.ComplexScriptItalic == true &&
                x.Formatting.ComplexScriptSizeHalfPoints == 20 &&
                x.Formatting.RightToLeftText == false &&
                x.Formatting.ForceComplexScript == false);
        }
        using var output = new MemoryStream(DxpDocToDocx.Project(docBytes).DocxBytes);
        using var projected = WordprocessingDocument.Open(output, false);
        Assert.Equal(minorFont, projected.MainDocumentPart!.StyleDefinitionsPart!
            .Styles!.GetFirstChild<DocDefaults>()!.Descendants<RunFonts>()
            .First().ComplexScript!.Value);
        var projectedDefaults = projected.MainDocumentPart.StyleDefinitionsPart.Styles
            .GetFirstChild<DocDefaults>()!.Descendants<RunPropertiesBaseStyle>().Single();
        Assert.False(projectedDefaults.GetFirstChild<BoldComplexScript>()!.Val!.Value);
        Assert.Equal("24", projectedDefaults.GetFirstChild<FontSizeComplexScript>()!
            .Val!.Value);
        Assert.Contains(projected.MainDocumentPart.StyleDefinitionsPart.Styles
            .Elements<Style>(), x => x.Type?.Value == StyleValues.Paragraph &&
                x.StyleRunProperties?.GetFirstChild<BoldComplexScript>()?.Val?.Value == true &&
                x.StyleRunProperties?.GetFirstChild<ItalicComplexScript>()?.Val?.Value == false &&
                x.StyleRunProperties?.GetFirstChild<FontSizeComplexScript>()?.Val?.Value == "30");
        Assert.Contains(projected.MainDocumentPart.StyleDefinitionsPart.Styles
            .Elements<Style>(), x => x.StyleName?.Val?.Value == "OnlyComplex" &&
                x.StyleRunProperties?.RunFonts?.ComplexScript?.Value == "Courier New");
        Assert.Contains(projected.MainDocumentPart.Document!.Body!
            .Descendants<RunProperties>(), x =>
                x.RunFonts?.ComplexScript?.Value == majorFont &&
                x.GetFirstChild<BoldComplexScript>()?.Val?.Value == true &&
                x.GetFirstChild<ItalicComplexScript>()?.Val?.Value == true &&
                x.GetFirstChild<FontSizeComplexScript>()?.Val?.Value == "28" &&
                x.GetFirstChild<RightToLeftText>()?.Val?.Value == true &&
                x.GetFirstChild<ComplexScript>()?.Val?.Value == true);
        Assert.Contains(projected.MainDocumentPart.HeaderParts
            .SelectMany(x => x.Header!.Descendants<RunFonts>()),
            x => x.ComplexScript?.Value == majorFont);
        Assert.Contains(projected.MainDocumentPart.HeaderParts
            .SelectMany(x => x.Header!.Descendants<RunFonts>()),
            x => x.ComplexScript?.Value == "Courier New");
        Assert.Contains(projected.MainDocumentPart.FooterParts
            .SelectMany(x => x.Footer!.Descendants<RunProperties>()),
            x => x.RunFonts?.ComplexScript?.Value == minorFont &&
                x.GetFirstChild<BoldComplexScript>()?.Val?.Value == false &&
                x.GetFirstChild<ItalicComplexScript>()?.Val?.Value == true &&
                x.GetFirstChild<FontSizeComplexScript>()?.Val?.Value == "20" &&
                x.GetFirstChild<RightToLeftText>()?.Val?.Value == false &&
                x.GetFirstChild<ComplexScript>()?.Val?.Value == false);
        var validationErrors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator()
            .Validate(projected).ToArray();
        Assert.True(validationErrors.Length == 0,
            string.Join(Environment.NewLine, validationErrors.Select(x =>
                x.Description + " " + x.Node?.OuterXml)));
        if (Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_WORD") != "1") return;
        var prefix = Path.Combine(Path.GetTempPath(),
            $"docxport-complex-font-{Guid.NewGuid():N}");
        var docxPath = prefix + ".docx";
        var docPath = prefix + ".doc";
        var nativePath = prefix + "-native.doc";
        File.WriteAllBytes(docxPath, source.ToArray());
        File.WriteAllBytes(docPath, docBytes);
        dynamic word = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application")!)!;
        word.Visible = false;
        dynamic? original = null, native = null, opened = null;
        try
        {
            original = word.Documents.Open(docxPath, ReadOnly: true,
                AddToRecentFiles: false);
            original.SaveAs2(FileName: nativePath, FileFormat: 0);
            original.Close(false);
            original = null;
            native = word.Documents.Open(nativePath, ReadOnly: true,
                AddToRecentFiles: false);
            opened = word.Documents.Open(docPath, ReadOnly: true,
                AddToRecentFiles: false);
            Assert.Contains(scriptText,
                (string)opened.Sections.Item(1).Headers.Item(1).Range.Text);
            Assert.Contains(scriptText,
                (string)opened.Sections.Item(1).Footers.Item(1).Range.Text);
            foreach (var headerStory in new[] { true, false })
            {
                dynamic expected = headerStory
                    ? native.Sections.Item(1).Headers.Item(1).Range.Characters.Item(8).Font
                    : native.Sections.Item(1).Footers.Item(1).Range.Characters.Item(8).Font;
                dynamic actual = headerStory
                    ? opened.Sections.Item(1).Headers.Item(1).Range.Characters.Item(8).Font
                    : opened.Sections.Item(1).Footers.Item(1).Range.Characters.Item(8).Font;
                Assert.Equal((string)expected.NameBi, (string)actual.NameBi);
                Assert.Equal((int)expected.BoldBi, (int)actual.BoldBi);
                Assert.Equal((int)expected.ItalicBi, (int)actual.ItalicBi);
                Assert.Equal((float)expected.SizeBi, (float)actual.SizeBi);
            }
        }
        finally
        {
            if (opened != null) opened.Close(false);
            if (native != null) native.Close(false);
            if (original != null) original.Close(false);
            word.Quit();
            File.Delete(docxPath);
            File.Delete(docPath);
            File.Delete(nativePath);
        }
    }
}
