using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Tests.Utils;
using DocxportNet.Fields;
using DocxportNet.Fields.Resolution;
using Xunit.Abstractions;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxportNet.Tests;

public sealed class DocxExportTests : TestBase<DocxExportTests>
{
    private static readonly string ProjectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
    private static readonly string SamplesDirectory = Path.Combine(ProjectRoot, "samples");

    public DocxExportTests(ITestOutputHelper output) : base(output) { }

    [Theory]
    [InlineData("sample-no-sectPr.docx")]
    [InlineData("TestFields.docx")]
    [InlineData("TestTableSpan.docx")]
    public void Passthrough_RebuildsReadableMainDocumentWithSameText(string fileName)
    {
        string inputPath = Path.Combine(SamplesDirectory, fileName);
        byte[] sourceBytes = File.ReadAllBytes(inputPath);
        byte[] outputBytes = DxpDocxExport.Export(
            sourceBytes,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.None },
            Logger);

        using var source = Open(sourceBytes);
        using var output = Open(outputBytes);

        Assert.Equal(source.MainDocumentPart!.Document.Body!.InnerText, output.MainDocumentPart!.Document.Body!.InnerText);
        Assert.Equal(
            source.MainDocumentPart.Document.Body.Descendants<Table>().Count(),
            output.MainDocumentPart.Document.Body.Descendants<Table>().Count());
        Assert.Equal(
            source.MainDocumentPart.Document.Body.Descendants<FieldChar>().Count(),
            output.MainDocumentPart.Document.Body.Descendants<FieldChar>().Count());
        Assert.Equal(
            source.MainDocumentPart.Document.Body.Descendants<SimpleField>().Count(),
            output.MainDocumentPart.Document.Body.Descendants<SimpleField>().Count());
        Assert.Equal(source.MainDocumentPart.StyleDefinitionsPart?.Uri, output.MainDocumentPart.StyleDefinitionsPart?.Uri);
    }

    [Fact]
    public void Passthrough_PreservesPackagePartsAndRewritesReferencedHeaderFooter()
    {
        byte[] sourceBytes = File.ReadAllBytes(Path.Combine(SamplesDirectory, "Right to Repair.docx"));
        byte[] outputBytes = DxpDocxExport.Export(
            sourceBytes,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.None },
            Logger);

        using var source = Open(sourceBytes);
        using var output = Open(outputBytes);

        Assert.Equal(source.MainDocumentPart!.Parts.Count(), output.MainDocumentPart!.Parts.Count());
        Assert.Equal(
            source.MainDocumentPart.HeaderParts.Select(static p => p.Uri).OrderBy(static p => p.ToString()),
            output.MainDocumentPart.HeaderParts.Select(static p => p.Uri).OrderBy(static p => p.ToString()));
        Assert.Equal(
            source.MainDocumentPart.FooterParts.Select(static p => p.Uri).OrderBy(static p => p.ToString()),
            output.MainDocumentPart.FooterParts.Select(static p => p.Uri).OrderBy(static p => p.ToString()));
    }

    [Fact]
    public void Passthrough_DoesNotDuplicateTableCellProperties()
    {
        byte[] sourceBytes = File.ReadAllBytes(Path.Combine(SamplesDirectory, "file-sample_1MB.docx"));
        byte[] outputBytes = DxpDocxExport.Export(
            sourceBytes,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.None },
            Logger);

        using var source = Open(sourceBytes);
        using var output = Open(outputBytes);
        var sourceCells = source.MainDocumentPart!.Document.Body!.Descendants<TableCell>().ToList();
        var outputCells = output.MainDocumentPart!.Document.Body!.Descendants<TableCell>().ToList();

        Assert.Equal(sourceCells.Count, outputCells.Count);
        Assert.Equal(
            sourceCells.Select(static cell => cell.Elements<TableCellProperties>().Count()),
            outputCells.Select(static cell => cell.Elements<TableCellProperties>().Count()));
        Assert.All(outputCells, static cell => Assert.True(cell.Elements<TableCellProperties>().Count() <= 1));
    }

    [Fact]
    public void Passthrough_PreservesParagraphsInsideTableCells()
    {
        byte[] sourceBytes = File.ReadAllBytes(Path.Combine(SamplesDirectory, "file-sample_1MB.docx"));
        byte[] outputBytes = DxpDocxExport.Export(
            sourceBytes,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.None },
            Logger);

        using var source = Open(sourceBytes);
        using var output = Open(outputBytes);
        var sourceCells = source.MainDocumentPart!.Document.Body!.Descendants<TableCell>().ToList();
        var outputCells = output.MainDocumentPart!.Document.Body!.Descendants<TableCell>().ToList();

        Assert.Equal(sourceCells.Count, outputCells.Count);
        Assert.Equal(
            sourceCells.Select(static cell => cell.Elements<Paragraph>().Count()),
            outputCells.Select(static cell => cell.Elements<Paragraph>().Count()));
        Assert.All(outputCells, static cell => Assert.NotEmpty(cell.Elements<Paragraph>()));
    }

    [Fact]
    public void Passthrough_PreservesSdtCellWrapperAndSingleTablePropertyExceptions()
    {
        byte[] sourceBytes;
        using (var stream = new MemoryStream())
        {
            using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
            {
                var main = document.AddMainDocumentPart();
                main.Document = new Document(new Body(new Table(
                    new TableProperties(),
                    new TableGrid(new GridColumn()),
                    new TableRow(
                        new TablePropertyExceptions(),
                        new SdtCell(
                            new SdtProperties(),
                            new SdtContentCell(
                                new TableCell(new Paragraph(new Run(new Text("cell"))))))))));
                main.Document.Save();
            }
            sourceBytes = stream.ToArray();
        }

        byte[] outputBytes = DxpDocxExport.Export(
            sourceBytes,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.None },
            Logger);

        using var output = Open(outputBytes);
        var row = Assert.Single(output.MainDocumentPart!.Document.Body!.Descendants<TableRow>());
        Assert.Single(row.Elements<TablePropertyExceptions>());
        var sdt = Assert.Single(row.Elements<SdtCell>());
        Assert.NotNull(sdt.SdtContentCell);
        Assert.Single(sdt.SdtContentCell!.Elements<TableCell>());
        var errors = new OpenXmlValidator().Validate(output).ToArray();
        Assert.True(errors.Length == 0, string.Join(Environment.NewLine,
            errors.Select(error => $"{error.Description} ({error.Path?.XPath})")));
    }

    [Fact]
    public void Passthrough_PreservesDrawings()
    {
        byte[] sourceBytes = File.ReadAllBytes(Path.Combine(SamplesDirectory, "file-sample_1MB.docx"));
        byte[] outputBytes = DxpDocxExport.Export(
            sourceBytes,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.None },
            Logger);

        using var source = Open(sourceBytes);
        using var output = Open(outputBytes);
        var sourceDrawings = source.MainDocumentPart!.Document.Body!.Descendants<Drawing>().ToList();
        var outputDrawings = output.MainDocumentPart!.Document.Body!.Descendants<Drawing>().ToList();

        Assert.NotEmpty(sourceDrawings);
        Assert.Equal(sourceDrawings.Count, outputDrawings.Count);
        Assert.Equal(
            sourceDrawings.Select(static drawing => drawing.OuterXml),
            outputDrawings.Select(static drawing => drawing.OuterXml));
    }

    [Fact]
    public void Passthrough_PreservesCommentMarkersWithoutNestingCommentParagraphs()
    {
        byte[] sourceBytes = File.ReadAllBytes(Path.Combine(SamplesDirectory, "TestComments.docx"));
        byte[] outputBytes = DxpDocxExport.Export(
            sourceBytes,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.None },
            Logger);

        using var source = Open(sourceBytes);
        using var output = Open(outputBytes);
        var sourceBody = source.MainDocumentPart!.Document.Body!;
        var outputBody = output.MainDocumentPart!.Document.Body!;

        Assert.Equal(sourceBody.Descendants<CommentRangeStart>().Count(), outputBody.Descendants<CommentRangeStart>().Count());
        Assert.Equal(sourceBody.Descendants<CommentRangeEnd>().Count(), outputBody.Descendants<CommentRangeEnd>().Count());
        Assert.Equal(sourceBody.Descendants<CommentReference>().Count(), outputBody.Descendants<CommentReference>().Count());
        Assert.DoesNotContain(outputBody.Descendants<Paragraph>(), static paragraph => paragraph.Parent is Paragraph);
        Assert.Empty(new OpenXmlValidator().Validate(output));
    }

    [Fact]
    public void EvaluatedExport_ReplacesFieldsWithLiteralRuns()
    {
        byte[] sourceBytes = File.ReadAllBytes(Path.Combine(SamplesDirectory, "TestFields.docx"));
        byte[] outputBytes = DxpDocxExport.Export(
            sourceBytes,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate },
            Logger);

        using var output = Open(outputBytes);
        var body = output.MainDocumentPart!.Document.Body!;

        Assert.Empty(body.Descendants<FieldChar>());
        Assert.Empty(body.Descendants<FieldCode>());
        Assert.Empty(body.Descendants<SimpleField>());
        Assert.Contains("Expect No Error: Not Empty", body.InnerText, StringComparison.Ordinal);
    }

    [Fact]
    public void EvaluatedExport_PreservesNativePaginationFields()
    {
        byte[] sourceBytes;
        using (var stream = new MemoryStream())
        {
            using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
            {
                var main = document.AddMainDocumentPart();
                main.Document = new Document(new Body(
                    new Paragraph(
                        new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                        new Run(new FieldCode(" PAGE ") { Space = SpaceProcessingModeValues.Preserve }),
                        new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                        new Run(new Text("1")),
                        new Run(new FieldChar { FieldCharType = FieldCharValues.End })),
                    new Paragraph(new SimpleField(new Run(new Text("cached"))) { Instruction = " NUMPAGES " }),
                    new Paragraph(new SimpleField(new Run(new Text("missing"))) { Instruction = " DOCVARIABLE Name " })));
                main.Document.Save();
            }
            sourceBytes = stream.ToArray();
        }

        var eval = new DxpFieldEval();
        eval.Context.SetDocVariable("Name", "resolved");
        byte[] outputBytes = DxpDocxExport.Export(
            sourceBytes,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate },
            Logger,
            eval);

        using var output = Open(outputBytes);
        var fields = output.MainDocumentPart!.Document.Body!.Descendants<SimpleField>().ToList();
        Assert.Equal(new[] { "NUMPAGES" }, fields.Select(field => field.Instruction!.Value!.Trim()));
        Assert.Equal("PAGE", Assert.Single(output.MainDocumentPart.Document.Body.Descendants<FieldCode>()).Text.Trim());
        Assert.Equal(3, output.MainDocumentPart.Document.Body.Descendants<FieldChar>().Count());
        Assert.Contains("resolved", output.MainDocumentPart.Document.Body.InnerText, StringComparison.Ordinal);
        Assert.DoesNotContain("DOCVARIABLE", output.MainDocumentPart.Document.Body.OuterXml, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Passthrough_NormalizesSyntheticLegacyAndMisnestedMarkup()
    {
        byte[] sourceBytes = CreateDocument(body =>
        {
            var outer = new Paragraph(new Run(new Text("before")));
            outer.AppendChild(new Paragraph(new Run(new Text("nested"))));
            body.Append(outer, new Run(new Text("block run")));

            var legacy = new Paragraph();
            var smartTag = new OpenXmlUnknownElement(
                "w", "smartTag", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            smartTag.AppendChild(new Run(new Text("tagged")));
            legacy.AppendChild(smartTag);
            body.Append(legacy);

            body.Append(new Table(
                new TableProperties(),
                new TableGrid(new GridColumn()),
                new TableRow(
                    new TableRowProperties(new ConditionalFormatStyle
                    {
                        Val = "100000000000",
                        FirstRow = true,
                        LastRow = false
                    }),
                    new TableCell(new Paragraph(new Run(new Text("cell")))))));
        });

        byte[] outputBytes = DxpDocxExport.Export(sourceBytes,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.None }, Logger);

        using var output = Open(outputBytes);
        var body = output.MainDocumentPart!.Document.Body!;
        Assert.Contains("beforenestedblock runtaggedcell", body.InnerText, StringComparison.Ordinal);
        Assert.DoesNotContain(body.Descendants(), element => element.LocalName == "smartTag");
        Assert.DoesNotContain(body.Descendants<Paragraph>(), paragraph => paragraph.Parent is Paragraph or Run);
        Assert.Empty(body.Elements<Run>());
        Assert.All(body.Descendants<ConditionalFormatStyle>(), style =>
            Assert.All(style.GetAttributes(), attribute => Assert.Equal("val", attribute.LocalName)));
        Assert.Empty(new OpenXmlValidator().Validate(output));
    }

    [Fact]
    public void Passthrough_MakesSyntheticBookmarkAndDrawingIdsUnique()
    {
        byte[] sourceBytes = CreateDocument(body =>
        {
            var firstDrawing = new Drawing(new DW.Inline(new DW.DocProperties { Id = 4U, Name = "shape-a" }));
            var secondDrawing = new Drawing(new DW.Inline(new DW.DocProperties { Id = 4U, Name = "shape-b" }));
            body.Append(
                new Paragraph(new BookmarkEnd { Id = "7" }),
                new Paragraph(new BookmarkEnd { Id = "7" }),
                new Paragraph(new Run(firstDrawing), new Run(secondDrawing)));
        });

        byte[] outputBytes = DxpDocxExport.Export(sourceBytes,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.None }, Logger);

        using var output = Open(outputBytes);
        var body = output.MainDocumentPart!.Document.Body!;
        Assert.Equal(2, body.Descendants<BookmarkEnd>().Select(end => end.Id!.Value).Distinct().Count());
        Assert.Equal(2, body.Descendants<DW.DocProperties>().Select(properties => properties.Id!.Value).Distinct().Count());
    }

    [Fact]
    public void EvaluatedExport_DropsSectionRelationshipsFromSyntheticIncludedDocument()
    {
        byte[] childBytes = CreateDocument(body => body.Append(
            new Paragraph(
                new ParagraphProperties(
                    new SectionProperties(
                        new HeaderReference { Id = "rIdSynthetic", Type = HeaderFooterValues.Default },
                        new FooterReference { Id = "rIdSynthetic2", Type = HeaderFooterValues.Default })),
                new Run(new Text("included")))));
        byte[] parentBytes = CreateDocument(body => body.Append(
            new Paragraph(new SimpleField(new Run(new Text("cached"))) { Instruction = " INCLUDETEXT \"child.docx\" " })));

        var eval = new DxpFieldEval();
        eval.Context.IncludeTextResolver = new SyntheticIncludeResolver(childBytes);
        byte[] outputBytes = DxpDocxExport.Export(parentBytes,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate }, Logger, eval);

        using var output = Open(outputBytes);
        var body = output.MainDocumentPart!.Document.Body!;
        Assert.Contains("included", body.InnerText, StringComparison.Ordinal);
        Assert.Empty(body.Descendants<HeaderReference>());
        Assert.Empty(body.Descendants<FooterReference>());
        Assert.Empty(new OpenXmlValidator().Validate(output));
    }

    [Fact]
    public void EvaluatedExport_RendersSyntheticDatabaseResultAsWordTable()
    {
        byte[] sourceBytes = CreateDocument(body => body.Append(
            new Paragraph(new SimpleField(new Run(new Text("cached"))) {
                Instruction = " DATABASE \\s \"SELECT Value FROM Items\" \\h "
            }),
            new Paragraph(new Run(new Text("after")))));
        var eval = new DxpFieldEval();
        eval.Context.DatabaseProvider = new SyntheticDatabaseProvider();

        byte[] outputBytes = DxpDocxExport.Export(sourceBytes,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate }, Logger, eval);

        using var output = Open(outputBytes);
        var body = output.MainDocumentPart!.Document.Body!;
        var table = Assert.Single(body.Elements<Table>());
        var rows = table.Elements<TableRow>().ToArray();
        Assert.Equal(3, rows.Length);
        Assert.Equal("Value", rows[0].InnerText);
        Assert.Equal("alpha", rows[1].InnerText);
        Assert.Equal("beta", rows[2].InnerText);
        Assert.Contains("after", body.InnerText, StringComparison.Ordinal);
        Assert.Empty(body.Descendants<FieldChar>());
        Assert.Empty(body.Descendants<SimpleField>());
        Assert.Empty(new OpenXmlValidator().Validate(output));
    }

    [Fact]
    public void EvaluatedExport_RendersSixtyTwoDatabaseColumnsAsTabs()
    {
        byte[] sourceBytes = CreateDocument(body => body.Append(
            new Paragraph(new SimpleField(new Run(new Text("cached"))) {
                Instruction = " DATABASE \\s \"SELECT synthetic columns\" "
            }),
            new Paragraph(new Run(new Text("after")))));
        var columns = Enumerable.Range(1, 62)
            .Select(index => new DxpDatabaseColumn($"C{index}"))
            .ToArray();
        var values = Enumerable.Range(1, 62)
            .Select(index => (DxpFieldValue?)new DxpFieldValue(index.ToString()))
            .ToArray();
        var eval = new DxpFieldEval();
        eval.Context.DatabaseProvider = new FixedDatabaseProvider(
            new DxpDatabaseResult(columns, [values]));

        byte[] outputBytes = DxpDocxExport.Export(sourceBytes,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate }, Logger, eval);

        using var output = Open(outputBytes);
        var body = output.MainDocumentPart!.Document.Body!;
        Assert.Empty(body.Elements<Table>());
        Assert.Equal(61, body.Descendants<TabChar>().Count());
        Assert.Contains("after", body.InnerText, StringComparison.Ordinal);
        Assert.Empty(new OpenXmlValidator().Validate(output));
    }

    private static byte[] CreateDocument(Action<Body> populate)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var body = new Body();
            populate(body);
            main.Document = new Document(body);
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private sealed class SyntheticIncludeResolver(byte[] content) : IDxpIncludeTextResolver
    {
        public Task<DxpIncludeTextSource?> ResolveAsync(
            DxpIncludeTextRequest request,
            DxpFieldEvalContext context,
            CancellationToken cancellationToken = default)
            => Task.FromResult<DxpIncludeTextSource?>(new DxpIncludeTextSource("synthetic-child", content));
    }

    private sealed class SyntheticDatabaseProvider : IDatabaseFieldProvider
    {
        public Task<DxpDatabaseResult?> ExecuteAsync(
            DxpDatabaseRequest request,
            CancellationToken cancellationToken)
            => Task.FromResult<DxpDatabaseResult?>(new DxpDatabaseResult(
                [new DxpDatabaseColumn("Value")],
                [
                    new DxpFieldValue?[] { new("alpha") },
                    new DxpFieldValue?[] { new("beta") }
                ]));
    }

    private sealed class FixedDatabaseProvider(DxpDatabaseResult result) : IDatabaseFieldProvider
    {
        public Task<DxpDatabaseResult?> ExecuteAsync(
            DxpDatabaseRequest request,
            CancellationToken cancellationToken)
            => Task.FromResult<DxpDatabaseResult?>(result);
    }

    private static WordprocessingDocument Open(byte[] bytes)
        => WordprocessingDocument.Open(new MemoryStream(bytes), false);
}
