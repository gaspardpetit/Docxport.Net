using System.Diagnostics;
using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Doc;

namespace DocxportNet.Tests;

public class DocLibreOfficeSmokeTests
{
    [Theory]
    [InlineData(false, 108, 0)]
    [InlineData(true, 108, 0)]
    [InlineData(false, 720, 0)]
    [InlineData(true, 720, 0)]
    [InlineData(false, 108, 1)]
    [InlineData(false, 108, 2)]
    public void GeneratedDocRetainsSimpleTableInLibreOffice(bool bordered,
        int indentTwips, byte justification)
    {
        if (Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_LIBREOFFICE") != "1") return;
        var office = Environment.GetEnvironmentVariable("DOCXPORT_LIBREOFFICE_PATH") ??
            @"C:\Program Files\LibreOffice\program\soffice.exe";
        if (!File.Exists(office)) throw new FileNotFoundException("LibreOffice is required.", office);
        static TableCell Cell(string value, bool border = false) => new(
            border ? new TableCellProperties(new TableCellBorders(new TopBorder
            { Val = BorderValues.Single, Size = 8, Color = "CC3300" })) :
                new TableCellProperties(),
            new Paragraph(new Run(new RunProperties(new RunFonts
            { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize { Val = "22" }),
                new Text(value))));
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Table(new TableProperties(
                    new TableJustification { Val = justification switch
                    {
                        0 => TableRowAlignmentValues.Left,
                        1 => TableRowAlignmentValues.Center,
                        _ => TableRowAlignmentValues.Right
                    } },
                    new TableIndentation { Width = indentTwips,
                        Type = TableWidthUnitValues.Dxa },
                    new TableCellMarginDefault(
                        new TableCellLeftMargin { Width = 108,
                            Type = TableWidthValues.Dxa },
                        new TableCellRightMargin { Width = 108,
                            Type = TableWidthValues.Dxa })),
                    new TableGrid(new GridColumn { Width = "2400" },
                        new GridColumn { Width = "3600" }),
                    new TableRow(new TableRowProperties(new CantSplit(),
                            new TableRowHeight { Val = 500,
                                HeightType = HeightRuleValues.Exact }, new TableHeader()),
                        Cell("A1", bordered), Cell("B1")),
                    new TableRow(new TableRowProperties(new TableRowHeight
                        { Val = 700, HeightType = HeightRuleValues.AtLeast }),
                        Cell("A2"), Cell("B2"))),
                new SectionProperties(new PageSize { Width = 12240, Height = 15840 },
                    new PageMargin { Left = 1440, Right = 1440, Top = 1440,
                        Bottom = 1440, Header = 720, Footer = 720 })));
            main.Document.Save();
        }
        var directory = Path.Combine(Path.GetTempPath(), "docxport-lo-table-" +
            Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var referencePath = Path.Combine(directory, "reference.docx");
            File.WriteAllBytes(referencePath, source.ToArray());
            var docPath = Path.Combine(directory, "table.doc");
            var binary = DxpDocExport.Export(source.ToArray());
            using (var indexedStream = new MemoryStream(binary))
            using (var index = new DocTextIndexWalker().Index(indexedStream))
            {
                Assert.Contains(index.ParagraphStyles, x =>
                    x.Formatting.TableIndentTwips == indentTwips &&
                    x.Formatting.TableJustification == justification);
                var projected = new DocToDocxProjector().Project(index);
                using var projectedStream = new MemoryStream(projected.DocxBytes);
                using var projectedDocx = WordprocessingDocument.Open(
                    projectedStream, false);
                Assert.Equal(indentTwips, Assert.Single(projectedDocx.MainDocumentPart!
                    .Document!.Body!.Elements<Table>()).TableProperties!
                    .GetFirstChild<TableIndentation>()!.Width!.Value);
                Assert.Equal(justification switch
                    {
                        0 => TableRowAlignmentValues.Left,
                        1 => TableRowAlignmentValues.Center,
                        _ => TableRowAlignmentValues.Right
                    }, Assert.Single(projectedDocx.MainDocumentPart.Document.Body
                        .Elements<Table>()).TableProperties!
                        .GetFirstChild<TableJustification>()!.Val!.Value);
                Assert.Empty(new OpenXmlValidator().Validate(projectedDocx));
            }
            File.WriteAllBytes(docPath, binary);
            if (justification != 0 &&
                Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_WORD") == "1")
            {
                var nativePath = Path.Combine(directory, "native.doc");
                dynamic word = Activator.CreateInstance(
                    Type.GetTypeFromProgID("Word.Application")!)!;
                word.Visible = false;
                dynamic? opened = null;
                try
                {
                    opened = word.Documents.Open(referencePath, ReadOnly: true,
                        AddToRecentFiles: false);
                    opened.SaveAs2(nativePath, 0);
                    opened.Close(false);
                    opened = word.Documents.Open(docPath, ReadOnly: true,
                        AddToRecentFiles: false);
                    Assert.Equal((int)justification,
                        (int)opened.Tables.Item(1).Rows.Alignment);
                    opened.Close(false);
                    opened = null;
                    using var nativeStream = new MemoryStream(File.ReadAllBytes(nativePath));
                    using var nativeIndex = new DocTextIndexWalker().Index(nativeStream);
                    Assert.Contains(nativeIndex.ParagraphStyles, x =>
                        x.Formatting.TableJustification == justification);
                    using var nativeProjectionStream = new MemoryStream(
                        new DocToDocxProjector().Project(nativeIndex).DocxBytes);
                    using var nativeProjection = WordprocessingDocument.Open(
                        nativeProjectionStream, false);
                    Assert.Equal(justification switch
                    {
                        1 => TableRowAlignmentValues.Center,
                        _ => TableRowAlignmentValues.Right
                    }, Assert.Single(nativeProjection.MainDocumentPart!.Document!
                        .Body!.Elements<Table>()).TableProperties!
                        .GetFirstChild<TableJustification>()!.Val!.Value);
                    Assert.Empty(new OpenXmlValidator().Validate(nativeProjection));
                }
                finally
                {
                    if (opened != null) opened.Close(false);
                    word.Quit();
                }
            }
            var convert = new ProcessStartInfo(office)
            {
                UseShellExecute = false, CreateNoWindow = true,
                RedirectStandardOutput = true, RedirectStandardError = true
            };
            foreach (var argument in new[] { "-env:UserInstallation=" +
                new Uri(Path.Combine(directory, "profile")).AbsoluteUri,
                "--headless", "--convert-to", "docx", "--outdir", directory, docPath })
                convert.ArgumentList.Add(argument);
            Run(convert);
            using var converted = WordprocessingDocument.Open(
                Path.Combine(directory, "table.docx"), false);
            var table = Assert.Single(converted.MainDocumentPart!.Document!.Body!
                .Elements<Table>());
            Assert.Equal(new[] { "A1", "B1", "A2", "B2" },
                table.Descendants<TableCell>().Select(x => x.InnerText));
            if (bordered)
                Assert.Equal(BorderValues.Single, table.Descendants<TableCell>()
                    .First().TableCellProperties?.TableCellBorders?.TopBorder?.Val?.Value);
            var rows = table.Elements<TableRow>().ToArray();
            Assert.NotNull(rows[0].TableRowProperties?.GetFirstChild<TableHeader>());
            Assert.NotNull(rows[0].TableRowProperties?.GetFirstChild<CantSplit>());
            Assert.Equal((uint)500, rows[0].TableRowProperties?
                .GetFirstChild<TableRowHeight>()?.Val?.Value);
            Assert.Equal((uint)700, rows[1].TableRowProperties?
                .GetFirstChild<TableRowHeight>()?.Val?.Value);
            var errors = new OpenXmlValidator().Validate(converted).ToArray();
            Assert.All(errors, error =>
            {
                Assert.Contains("invalid value 'true'", error.Description);
                Assert.True(error.Path?.XPath?.EndsWith("/w:cantSplit[1]") == true ||
                    error.Path?.XPath?.EndsWith("/w:tblHeader[1]") == true,
                    error.Path?.XPath + ": " + error.Description);
            });
            Assert.InRange(errors.Length, 0, 2); // Some LibreOffice versions emit w:val="true".
            if (Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_RENDER") == "1")
            {
                var ghostscript = Environment.GetEnvironmentVariable("DOCXPORT_GHOSTSCRIPT_PATH") ??
                    @"C:\Program Files\gs\gs10.06.0\bin\gswin64c.exe";
                var imageMagick = Environment.GetEnvironmentVariable("DOCXPORT_IMAGEMAGICK_PATH") ??
                    @"C:\Program Files\ImageMagick-7.1.2-Q16-HDRI\magick.exe";
                foreach (var input in new[] { referencePath, docPath })
                {
                    var render = new ProcessStartInfo(office)
                    {
                        UseShellExecute = false, CreateNoWindow = true,
                        RedirectStandardOutput = true, RedirectStandardError = true
                    };
                    foreach (var argument in new[] { "-env:UserInstallation=" +
                        new Uri(Path.Combine(directory, "render-profile")).AbsoluteUri,
                        "--headless", "--convert-to", "pdf", "--outdir", directory, input })
                        render.ArgumentList.Add(argument);
                    Run(render);
                    var name = Path.GetFileNameWithoutExtension(input);
                    var raster = new ProcessStartInfo(ghostscript)
                    {
                        UseShellExecute = false, CreateNoWindow = true,
                        RedirectStandardOutput = true, RedirectStandardError = true
                    };
                    foreach (var argument in new[] { "-q", "-dNOPAUSE", "-dBATCH",
                        "-sDEVICE=png16m", "-r96", "-sOutputFile=" +
                        Path.Combine(directory, name + "-page-%d.png"),
                        Path.Combine(directory, name + ".pdf") })
                        raster.ArgumentList.Add(argument);
                    Run(raster);
                }
                var referenceImage = Path.Combine(directory, "reference-page-1.png");
                var docImage = Path.Combine(directory, "table-page-1.png");
                Assert.True(File.Exists(referenceImage) && File.Exists(docImage));
                var compare = new ProcessStartInfo(imageMagick)
                {
                    UseShellExecute = false, CreateNoWindow = true,
                    RedirectStandardOutput = true, RedirectStandardError = true
                };
                foreach (var argument in new[] { "compare", "-metric", "RMSE",
                    referenceImage, docImage, "null:" })
                    compare.ArgumentList.Add(argument);
                using var process = Process.Start(compare)!;
                if (!process.WaitForExit(30000))
                { process.Kill(); throw new TimeoutException("Table image comparison timed out."); }
                var metric = process.StandardError.ReadToEnd();
                var match = Regex.Match(metric, @"\(([0-9.]+)\)");
                Assert.True(match.Success, "No table RMSE: " + metric);
                var rmse = double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
                // Explicit source padding matches Word's DOC default and gives
                // both LibreOffice imports the same table cell geometry.
                Assert.True(rmse <= 0.005, $"Table page differs by {rmse}; files: {directory}");
            }
        }
        finally
        {
            if (Environment.GetEnvironmentVariable("DOCXPORT_KEEP_LIBREOFFICE_FILES") != "1")
                Directory.Delete(directory, true);
        }
    }

    [Fact]
    public void GeneratedDocRetainsHeaderFooterVariantsInLibreOffice()
    {
        if (Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_LIBREOFFICE") != "1") return;
        var office = Environment.GetEnvironmentVariable("DOCXPORT_LIBREOFFICE_PATH") ??
            @"C:\Program Files\LibreOffice\program\soffice.exe";
        if (!File.Exists(office)) throw new FileNotFoundException("LibreOffice is required.", office);
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            styles.Styles = new Styles(new DocDefaults(new RunPropertiesDefault(
                new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                    new FontSize { Val = "22" }))));
            styles.Styles.Save();
            var settings = main.AddNewPart<DocumentSettingsPart>();
            settings.Settings = new Settings(new EvenAndOddHeaders());
            settings.Settings.Save();
            static Run StyledRun(string text) => new(
                new RunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                    new FontSize { Val = "22" }), new Text(text)
                    { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve });
            HeaderReference Header(string text, HeaderFooterValues kind)
            {
                var part = main.AddNewPart<HeaderPart>();
                part.Header = new Header(new Paragraph(StyledRun(text)));
                part.Header.Save();
                return new HeaderReference { Type = kind, Id = main.GetIdOfPart(part) };
            }
            FooterReference Footer(string text, HeaderFooterValues kind)
            {
                var part = main.AddNewPart<FooterPart>();
                var cachedPage = kind == HeaderFooterValues.First ? "1" :
                    kind == HeaderFooterValues.Even ? "2" : "3";
                part.Footer = new Footer(new Paragraph(StyledRun(text + " "),
                    new SimpleField(StyledRun(cachedPage)) { Instruction = " PAGE " }));
                part.Footer.Save();
                return new FooterReference { Type = kind, Id = main.GetIdOfPart(part) };
            }
            main.Document = new Document(new Body(
                new Paragraph(StyledRun("First page")),
                new Paragraph(new ParagraphProperties(new PageBreakBefore()),
                    StyledRun("Second page")),
                new Paragraph(new ParagraphProperties(new PageBreakBefore()),
                    StyledRun("Third page")),
                new SectionProperties(
                    Header("First header", HeaderFooterValues.First),
                    Header("Even header", HeaderFooterValues.Even),
                    Header("Odd header", HeaderFooterValues.Default),
                    Footer("First footer", HeaderFooterValues.First),
                    Footer("Even footer", HeaderFooterValues.Even),
                    Footer("Odd footer", HeaderFooterValues.Default),
                    new PageSize { Width = 12240, Height = 15840 },
                    new PageMargin { Left = 1440, Right = 1440, Top = 1440,
                        Bottom = 1440, Header = 720, Footer = 720 },
                    new TitlePage())));
            main.Document.Save();
        }
        var directory = Path.Combine(Path.GetTempPath(), "docxport-lo-variants-" +
            Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var referencePath = Path.Combine(directory, "reference.docx");
            File.WriteAllBytes(referencePath, source.ToArray());
            var docPath = Path.Combine(directory, "variants.doc");
            File.WriteAllBytes(docPath, DxpDocExport.Export(source.ToArray()));
            if (Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_WORD") == "1")
                VerifyVariantsInWord(docPath);
            var start = new ProcessStartInfo(office)
            {
                UseShellExecute = false, CreateNoWindow = true,
                RedirectStandardOutput = true, RedirectStandardError = true
            };
            start.ArgumentList.Add("-env:UserInstallation=" +
                new Uri(Path.Combine(directory, "profile")).AbsoluteUri);
            start.ArgumentList.Add("--headless");
            start.ArgumentList.Add("--convert-to");
            start.ArgumentList.Add("docx");
            start.ArgumentList.Add("--outdir");
            start.ArgumentList.Add(directory);
            start.ArgumentList.Add(docPath);
            Run(start);
            using var converted = WordprocessingDocument.Open(
                Path.Combine(directory, "variants.docx"), false);
            var main = converted.MainDocumentPart!;
            var headers = main.HeaderParts.Select(x => x.Header?.InnerText).ToArray();
            var footers = main.FooterParts.Select(x => x.Footer?.InnerText).ToArray();
            Assert.Contains("First header", headers);
            Assert.Contains("Even header", headers);
            Assert.Contains("Odd header", headers);
            Assert.Contains(footers, x => x?.StartsWith("First footer ") == true);
            Assert.Contains(footers, x => x?.StartsWith("Even footer ") == true);
            Assert.Contains(footers, x => x?.StartsWith("Odd footer ") == true);
            Assert.All(main.FooterParts, x =>
                Assert.Contains(x.Footer!.Descendants<FieldCode>(), code =>
                    code.Text.Contains("PAGE", StringComparison.OrdinalIgnoreCase)));
            Assert.NotNull(main.DocumentSettingsPart?.Settings?.GetFirstChild<EvenAndOddHeaders>());
            Assert.NotNull(main.Document?.Body?.GetFirstChild<SectionProperties>()?
                .GetFirstChild<TitlePage>());
            if (Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_RENDER") == "1")
            {
                var ghostscript = Environment.GetEnvironmentVariable("DOCXPORT_GHOSTSCRIPT_PATH") ??
                    @"C:\Program Files\gs\gs10.06.0\bin\gswin64c.exe";
                if (!File.Exists(ghostscript))
                    throw new FileNotFoundException("Ghostscript is required for rendering.", ghostscript);
                foreach (var input in new[] { referencePath, docPath })
                {
                    var render = new ProcessStartInfo(office)
                    {
                        UseShellExecute = false, CreateNoWindow = true,
                        RedirectStandardOutput = true, RedirectStandardError = true
                    };
                    render.ArgumentList.Add("-env:UserInstallation=" +
                        new Uri(Path.Combine(directory, "render-profile")).AbsoluteUri);
                    render.ArgumentList.Add("--headless");
                    render.ArgumentList.Add("--convert-to");
                    render.ArgumentList.Add("pdf");
                    render.ArgumentList.Add("--outdir");
                    render.ArgumentList.Add(directory);
                    render.ArgumentList.Add(input);
                    Run(render);
                    var name = Path.GetFileNameWithoutExtension(input);
                    var pdf = Path.Combine(directory, name + ".pdf");
                    Assert.True(File.Exists(pdf), "LibreOffice did not render " + input);
                    var raster = new ProcessStartInfo(ghostscript)
                    {
                        UseShellExecute = false, CreateNoWindow = true,
                        RedirectStandardOutput = true, RedirectStandardError = true
                    };
                    foreach (var argument in new[] { "-q", "-dNOPAUSE", "-dBATCH",
                        "-sDEVICE=png16m", "-r96", "-sOutputFile=" +
                        Path.Combine(directory, name + "-page-%d.png"), pdf })
                        raster.ArgumentList.Add(argument);
                    Run(raster);
                }
                Assert.Equal(3, Directory.GetFiles(directory, "variants-page-*.png").Length);
                Assert.Equal(3, Directory.GetFiles(directory, "reference-page-*.png").Length);
                var imageMagick = Environment.GetEnvironmentVariable("DOCXPORT_IMAGEMAGICK_PATH") ??
                    @"C:\Program Files\ImageMagick-7.1.2-Q16-HDRI\magick.exe";
                if (!File.Exists(imageMagick))
                    throw new FileNotFoundException("ImageMagick is required for comparison.", imageMagick);
                for (var page = 1; page <= 3; page++)
                {
                    var compare = new ProcessStartInfo(imageMagick)
                    {
                        UseShellExecute = false, CreateNoWindow = true,
                        RedirectStandardOutput = true, RedirectStandardError = true
                    };
                    compare.ArgumentList.Add("compare");
                    compare.ArgumentList.Add("-metric");
                    compare.ArgumentList.Add("RMSE");
                    compare.ArgumentList.Add(Path.Combine(directory, $"reference-page-{page}.png"));
                    compare.ArgumentList.Add(Path.Combine(directory, $"variants-page-{page}.png"));
                    compare.ArgumentList.Add("null:");
                    using var process = Process.Start(compare)!;
                    if (!process.WaitForExit(30000))
                    { process.Kill(); throw new TimeoutException("Page comparison timed out."); }
                    var metric = process.StandardError.ReadToEnd();
                    var match = Regex.Match(metric, @"\(([0-9.]+)\)");
                    Assert.True(match.Success, "No RMSE for page " + page + ": " + metric);
                    var normalizedRmse = double.Parse(match.Groups[1].Value,
                        CultureInfo.InvariantCulture);
                    Assert.True(normalizedRmse <= 0.005,
                        $"Page {page} differs by {normalizedRmse}; files: {directory}");
                }
            }
        }
        finally
        {
            if (Environment.GetEnvironmentVariable("DOCXPORT_KEEP_LIBREOFFICE_FILES") != "1")
                Directory.Delete(directory, true);
        }
    }

    [Fact]
    public void GeneratedDocOpensInIndependentWordProcessor()
    {
        if (Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_LIBREOFFICE") != "1") return;
        var office = Environment.GetEnvironmentVariable("DOCXPORT_LIBREOFFICE_PATH") ??
            @"C:\Program Files\LibreOffice\program\soffice.exe";
        if (!File.Exists(office)) throw new FileNotFoundException("LibreOffice is required for this check.", office);
        var renderPages = Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_RENDER") == "1";

        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(
                new Style(new StyleName { Val = "Visible Base" },
                    new StyleParagraphProperties(
                        new Justification { Val = JustificationValues.Center }),
                    new StyleRunProperties(
                    new RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                    new Bold(), new Spacing { Val = 20 },
                    new FontSize { Val = "22" }, new Languages { Val = "en-US" }))
                { Type = StyleValues.Paragraph, StyleId = "VisibleBase" },
                new Style(new StyleName { Val = "Visible Style" },
                    new BasedOn { Val = "VisibleBase" },
                    new StyleRunProperties(new Italic()))
                { Type = StyleValues.Paragraph, StyleId = "VisibleStyle" },
                new Style(new StyleName { Val = "Visual Accent" },
                    new StyleRunProperties(new Color { Val = "AA3366" },
                        new Underline { Val = UnderlineValues.Single }))
                { Type = StyleValues.Character, StyleId = "VisualAccent" });
            stylesPart.Styles.Save();
            var header = main.AddNewPart<HeaderPart>();
            header.Header = new Header(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "VisibleStyle" }),
                new Run(new Text("Reference header"))));
            header.Header.Save();
            var footer = main.AddNewPart<FooterPart>();
            footer.Footer = new Footer(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "VisibleStyle" }),
                new Run(new Text("Reference footer"))));
            footer.Footer.Save();
            var section = new SectionProperties(new HeaderReference
            {
                Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(header)
            }, new FooterReference
            {
                Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(footer)
            });
            if (renderPages)
            {
                section.AppendChild(new PageSize { Width = 12240, Height = 15840 });
                section.AppendChild(new PageMargin { Left = 1440, Right = 1440,
                    Top = 1440, Bottom = 1440, Header = 720, Footer = 720 });
            }
            main.Document = new Document(new Body(
                new Paragraph(new ParagraphProperties(
                    new ParagraphStyleId { Val = "VisibleStyle" },
                    new ParagraphBorders(new TopBorder { Val = BorderValues.Single,
                        Size = 8, Space = 2, Color = "112233" }),
                    new Shading
                {
                    Val = ShadingPatternValues.Percent50, Color = "123456", Fill = "E1F2A3"
                }), new Run(new Text("Reference body"))),
                new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId { Val = "VisibleStyle" }),
                    new Run(new RunProperties(new RunStyle { Val = "VisualAccent" },
                        new Bold { Val = false }), new Text("Accent override"))),
                section));
            main.Document.Save();
        }

        var directory = Path.Combine(Path.GetTempPath(), "docxport-lo-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var referencePath = Path.Combine(directory, "reference.docx");
            File.WriteAllBytes(referencePath, source.ToArray());
            var docPath = Path.Combine(directory, "generated.doc");
            File.WriteAllBytes(docPath, DxpDocExport.Export(source.ToArray()));
            if (Environment.GetEnvironmentVariable("DOCXPORT_VERIFY_WORD") == "1")
                VerifyStyledStoriesInWord(docPath);
            var projection = DxpDocToDocx.Project(File.ReadAllBytes(docPath));
            using (var projectedStream = new MemoryStream(projection.DocxBytes))
            using (var projected = WordprocessingDocument.Open(projectedStream, false))
            {
                var style = Assert.Single(projected.MainDocumentPart!
                    .StyleDefinitionsPart!.Styles!.Elements<Style>(),
                    x => x.StyleName?.Val?.Value == "Visible Style");
                var baseStyle = Assert.Single(projected.MainDocumentPart
                    .StyleDefinitionsPart.Styles.Elements<Style>(),
                    x => x.StyleName?.Val?.Value == "Visible Base");
                Assert.Equal(baseStyle.StyleId!.Value, style.BasedOn?.Val?.Value);
                Assert.Equal(JustificationValues.Center,
                    baseStyle.StyleParagraphProperties?.Justification?.Val?.Value);
                Assert.Equal(20, baseStyle.StyleRunProperties?.Spacing?.Val?.Value);
                Assert.Equal("en-US", baseStyle.StyleRunProperties?.Languages?.Val?.Value);
                Assert.Equal(style.StyleId!.Value, projected.MainDocumentPart.Document!
                    .Body!.Elements<Paragraph>().First().ParagraphProperties?
                    .ParagraphStyleId?.Val?.Value);
                Assert.Equal(style.StyleId.Value, projected.MainDocumentPart.HeaderParts
                    .Single().Header!.Elements<Paragraph>().Single().ParagraphProperties?
                    .ParagraphStyleId?.Val?.Value);
                Assert.Equal(style.StyleId.Value, projected.MainDocumentPart.FooterParts
                    .Single().Footer!.Elements<Paragraph>().Single().ParagraphProperties?
                    .ParagraphStyleId?.Val?.Value);
                var accent = Assert.Single(projected.MainDocumentPart.StyleDefinitionsPart
                    .Styles.Elements<Style>(),
                    x => x.StyleName?.Val?.Value == "Visual Accent");
                var accentRun = projected.MainDocumentPart.Document.Body
                    .Elements<Paragraph>().Skip(1).Single().Elements<Run>().Single();
                Assert.Equal(accent.StyleId!.Value,
                    accentRun.RunProperties?.RunStyle?.Val?.Value);
                Assert.False(accentRun.RunProperties?.Bold?.Val?.Value);
                Assert.Empty(new OpenXmlValidator().Validate(projected));
            }
            var profile = new Uri(Path.Combine(directory, "profile")).AbsoluteUri;
            var start = new ProcessStartInfo(office)
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            start.ArgumentList.Add("-env:UserInstallation=" + profile);
            start.ArgumentList.Add("--headless");
            start.ArgumentList.Add("--convert-to");
            start.ArgumentList.Add("docx");
            start.ArgumentList.Add("--outdir");
            start.ArgumentList.Add(directory);
            start.ArgumentList.Add(docPath);
            using var process = Process.Start(start)!;
            if (!process.WaitForExit(30000))
            {
                process.Kill();
                throw new TimeoutException("LibreOffice did not finish opening the generated DOC.");
            }
            var output = process.StandardOutput.ReadToEnd() + process.StandardError.ReadToEnd();
            Assert.True(process.ExitCode == 0, output);
            var convertedPath = Path.Combine(directory, "generated.docx");
            Assert.True(File.Exists(convertedPath), output);
            using var converted = WordprocessingDocument.Open(convertedPath, false);
            var body = converted.MainDocumentPart!.Document!.Body!;
            Assert.Contains("Reference body", body.InnerText);
            var visibleStyle = Assert.Single(converted.MainDocumentPart.StyleDefinitionsPart!
                .Styles!.Elements<Style>(), x => x.StyleName?.Val?.Value == "Visible Style");
            var visibleBase = Assert.Single(converted.MainDocumentPart.StyleDefinitionsPart
                .Styles.Elements<Style>(), x => x.StyleName?.Val?.Value == "Visible Base");
            Assert.Equal(visibleBase.StyleId!.Value, visibleStyle.BasedOn?.Val?.Value);
            Assert.Equal("Arial", visibleBase.StyleRunProperties?.RunFonts?.Ascii?.Value);
            Assert.Equal("22", visibleBase.StyleRunProperties?.FontSize?.Val?.Value);
            Assert.NotNull(visibleBase.StyleRunProperties?.Bold);
            Assert.Equal(JustificationValues.Center,
                visibleBase.StyleParagraphProperties?.Justification?.Val?.Value);
            Assert.NotNull(visibleStyle.StyleRunProperties?.Italic);
            var visualAccent = Assert.Single(converted.MainDocumentPart.StyleDefinitionsPart
                .Styles.Elements<Style>(), x => x.StyleName?.Val?.Value == "Visual Accent");
            Assert.Equal("AA3366", visualAccent.StyleRunProperties?.Color?.Val?.Value);
            Assert.Equal(visualAccent.StyleId!.Value, body.Elements<Paragraph>().Skip(1)
                .Single().Elements<Run>().Single().RunProperties?.RunStyle?.Val?.Value);
            Assert.Equal(visibleStyle.StyleId!.Value, body.Elements<Paragraph>().First()
                .ParagraphProperties?.ParagraphStyleId?.Val?.Value);
            Assert.Equal(visibleStyle.StyleId.Value,
                converted.MainDocumentPart.HeaderParts.Single().Header!
                    .Elements<Paragraph>().Single().ParagraphProperties?
                    .ParagraphStyleId?.Val?.Value);
            Assert.Equal(visibleStyle.StyleId.Value,
                converted.MainDocumentPart.FooterParts.Single().Footer!
                    .Elements<Paragraph>().Single().ParagraphProperties?
                    .ParagraphStyleId?.Val?.Value);
            var border = body.Elements<Paragraph>().First().ParagraphProperties?
                .ParagraphBorders?.TopBorder;
            Assert.True(border?.Val?.Value == BorderValues.Single,
                "LibreOffice did not retain the paragraph border; files: " + directory);
            Assert.True(converted.MainDocumentPart.HeaderParts.Any(part =>
                part.Header?.InnerText.Contains("Reference header") == true),
                "LibreOffice did not retain the header; files: " + directory);
            Assert.True(converted.MainDocumentPart.FooterParts.Any(part =>
                part.Footer?.InnerText.Contains("Reference footer") == true),
                "LibreOffice did not retain the footer; files: " + directory);

            if (renderPages)
            {
                var ghostscript = Environment.GetEnvironmentVariable("DOCXPORT_GHOSTSCRIPT_PATH") ??
                    @"C:\Program Files\gs\gs10.06.0\bin\gswin64c.exe";
                if (!File.Exists(ghostscript))
                    throw new FileNotFoundException("Ghostscript is required for rendering.", ghostscript);
                foreach (var input in new[] { referencePath, docPath })
                {
                    var render = new ProcessStartInfo(office)
                    {
                        UseShellExecute = false, CreateNoWindow = true,
                        RedirectStandardOutput = true, RedirectStandardError = true
                    };
                    render.ArgumentList.Add("-env:UserInstallation=" + profile);
                    render.ArgumentList.Add("--headless");
                    render.ArgumentList.Add("--convert-to");
                    render.ArgumentList.Add("pdf");
                    render.ArgumentList.Add("--outdir");
                    render.ArgumentList.Add(directory);
                    render.ArgumentList.Add(input);
                    Run(render);
                    var name = Path.GetFileNameWithoutExtension(input);
                    var pdf = Path.Combine(directory, name + ".pdf");
                    Assert.True(File.Exists(pdf), "LibreOffice did not render " + input);
                    var png = Path.Combine(directory, name + ".png");
                    var raster = new ProcessStartInfo(ghostscript)
                    {
                        UseShellExecute = false, CreateNoWindow = true,
                        RedirectStandardOutput = true, RedirectStandardError = true
                    };
                    foreach (var argument in new[] { "-q", "-dNOPAUSE", "-dBATCH",
                        "-sDEVICE=png16m", "-r96", "-dFirstPage=1", "-dLastPage=1",
                        "-sOutputFile=" + png, pdf }) raster.ArgumentList.Add(argument);
                    Run(raster);
                    Assert.True(File.Exists(png), "Ghostscript did not rasterize " + pdf);
                }
                var imageMagick = Environment.GetEnvironmentVariable("DOCXPORT_IMAGEMAGICK_PATH") ??
                    @"C:\Program Files\ImageMagick-7.1.2-Q16-HDRI\magick.exe";
                if (!File.Exists(imageMagick))
                    throw new FileNotFoundException("ImageMagick is required for image comparison.", imageMagick);
                var compare = new ProcessStartInfo(imageMagick)
                {
                    UseShellExecute = false, CreateNoWindow = true,
                    RedirectStandardOutput = true, RedirectStandardError = true
                };
                compare.ArgumentList.Add("compare");
                compare.ArgumentList.Add("-metric");
                compare.ArgumentList.Add("RMSE");
                compare.ArgumentList.Add(Path.Combine(directory, "reference.png"));
                compare.ArgumentList.Add(Path.Combine(directory, "generated.png"));
                compare.ArgumentList.Add("null:");
                using var comparison = Process.Start(compare)!;
                if (!comparison.WaitForExit(30000))
                {
                    comparison.Kill();
                    throw new TimeoutException("ImageMagick did not finish comparing pages.");
                }
                var metric = comparison.StandardError.ReadToEnd();
                var match = Regex.Match(metric, @"\(([0-9.]+)\)");
                Assert.True(match.Success, "ImageMagick returned no normalized RMSE: " + metric);
                var normalizedRmse = double.Parse(match.Groups[1].Value,
                    CultureInfo.InvariantCulture);
                Assert.InRange(normalizedRmse, 0, 0.001);
            }
        }
        finally
        {
            if (Environment.GetEnvironmentVariable("DOCXPORT_KEEP_LIBREOFFICE_FILES") != "1")
                Directory.Delete(directory, true);
        }
    }

    private static void Run(ProcessStartInfo start)
    {
        using var process = Process.Start(start)!;
        if (!process.WaitForExit(30000))
        {
            process.Kill();
            throw new TimeoutException(start.FileName + " did not finish.");
        }
        var output = process.StandardOutput.ReadToEnd() + process.StandardError.ReadToEnd();
        Assert.True(process.ExitCode == 0, output);
    }

    private static void VerifyVariantsInWord(string path)
    {
        if (!OperatingSystem.IsWindows())
            throw new PlatformNotSupportedException("Microsoft Word automation requires Windows.");
        var type = Type.GetTypeFromProgID("Word.Application") ??
            throw new InvalidOperationException("Microsoft Word is not installed.");
        dynamic word = Activator.CreateInstance(type)!;
        dynamic? document = null;
        try
        {
            word.Visible = false;
            word.DisplayAlerts = 0;
            document = word.Documents.Open(path, false, true);
            Assert.Contains("First page", (string)document.Content.Text);
            Assert.Contains("Third page", (string)document.Content.Text);
            dynamic section = document.Sections.Item(1);
            Assert.Contains("Odd header", (string)section.Headers.Item(1).Range.Text);
            Assert.Contains("First header", (string)section.Headers.Item(2).Range.Text);
            Assert.Contains("Even header", (string)section.Headers.Item(3).Range.Text);
            Assert.Contains("Odd footer", (string)section.Footers.Item(1).Range.Text);
            Assert.Equal(1, (int)section.Footers.Item(1).Range.Fields.Count);
            Assert.Contains("First footer", (string)section.Footers.Item(2).Range.Text);
            Assert.Equal(1, (int)section.Footers.Item(2).Range.Fields.Count);
            Assert.Contains("Even footer", (string)section.Footers.Item(3).Range.Text);
            Assert.Equal(1, (int)section.Footers.Item(3).Range.Fields.Count);
        }
        finally
        {
            if (document != null) document.Close(false);
            word.Quit();
        }
    }

    private static void VerifyStyledStoriesInWord(string path)
    {
        if (!OperatingSystem.IsWindows())
            throw new PlatformNotSupportedException("Microsoft Word automation requires Windows.");
        var type = Type.GetTypeFromProgID("Word.Application") ??
            throw new InvalidOperationException("Microsoft Word is not installed.");
        dynamic word = Activator.CreateInstance(type)!;
        dynamic? document = null;
        try
        {
            word.Visible = false;
            word.DisplayAlerts = 0;
            document = word.Documents.Open(path, false, true);
            dynamic bodyParagraph = document.Paragraphs.Item(1);
            Assert.Equal("Visible Style", (string)bodyParagraph.Style.NameLocal);
            Assert.Equal(-1, (int)bodyParagraph.Range.Font.Bold);
            Assert.Equal(-1, (int)bodyParagraph.Range.Font.Italic);
            Assert.Equal("Arial", (string)bodyParagraph.Range.Font.Name);
            Assert.Equal(11f, (float)bodyParagraph.Range.Font.Size);
            Assert.Equal(1, (int)bodyParagraph.Alignment);
            Assert.Equal(1, (int)bodyParagraph.Borders.Item(-1).LineStyle);
            Assert.Equal(10744545, (int)bodyParagraph.Shading.BackgroundPatternColor);
            dynamic accentParagraph = document.Paragraphs.Item(2);
            Assert.Equal("Visible Style", (string)accentParagraph.Style.NameLocal);
            Assert.Equal(0, (int)accentParagraph.Range.Font.Bold);
            Assert.Equal(-1, (int)accentParagraph.Range.Font.Italic);
            Assert.Equal(1, (int)accentParagraph.Range.Font.Underline);
            dynamic section = document.Sections.Item(1);
            Assert.Equal("Visible Style", (string)section.Headers.Item(1).Range
                .Paragraphs.Item(1).Style.NameLocal);
            Assert.Equal(-1, (int)section.Headers.Item(1).Range.Font.Bold);
            Assert.Equal(-1, (int)section.Headers.Item(1).Range.Font.Italic);
            Assert.Equal(1, (int)section.Headers.Item(1).Range.Paragraphs.Item(1).Alignment);
            Assert.Equal("Visible Style", (string)section.Footers.Item(1).Range
                .Paragraphs.Item(1).Style.NameLocal);
            Assert.Equal(-1, (int)section.Footers.Item(1).Range.Font.Bold);
            Assert.Equal(-1, (int)section.Footers.Item(1).Range.Font.Italic);
            Assert.Equal(1, (int)section.Footers.Item(1).Range.Paragraphs.Item(1).Alignment);
        }
        finally
        {
            if (document != null) document.Close(false);
            word.Quit();
        }
    }
}
