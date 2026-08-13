using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Fields;
using DocxportNet.Fields.Resolution;
using DocxportNet.Tests.Utils;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using Xunit.Abstractions;

namespace DocxportNet.Tests;

public sealed class SemanticFieldResultTests : TestBase<SemanticFieldResultTests>
{
    public SemanticFieldResultTests(ITestOutputHelper output) : base(output) { }

    [Fact]
    public void NestedDatabaseInSelectedIfRemainsStructuredAcrossExporters()
    {
        byte[] source = CreateNestedDatabaseDocument(condition: true);

        byte[] docx = DxpDocxExport.Export(source, SemanticOptions(), Logger, CreateEval());
        using (var output = Open(docx))
        {
            Table table = Assert.Single(output.MainDocumentPart!.Document.Body!.Elements<Table>());
            Assert.Equal(3, table.Elements<TableRow>().Count());
            Assert.Contains("Alice", table.InnerText, StringComparison.Ordinal);
            Assert.NotEmpty(table.Descendants<Break>());
            Assert.Empty(new OpenXmlValidator().Validate(output));
        }

        string html = DxpExport.ExportToString(
            source,
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig(), Logger, CreateEval()),
            SemanticOptions(),
            Logger);
        Assert.Contains("<table", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Alice", html, StringComparison.Ordinal);

        string markdown = DxpExport.ExportToString(
            source,
            new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreateRichConfig(), Logger, CreateEval()),
            SemanticOptions(),
            Logger);
        Assert.Contains("Alice", markdown, StringComparison.Ordinal);
        Assert.Contains("Montreal", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void NestedDatabaseInUnselectedIfBranchProducesNothing()
    {
        byte[] source = CreateNestedDatabaseDocument(condition: false);
        byte[] docx = DxpDocxExport.Export(source, SemanticOptions(), Logger, CreateEval());

        using var output = Open(docx);
        Assert.Empty(output.MainDocumentPart!.Document.Body!.Descendants<Table>());
        Assert.DoesNotContain("Alice", output.MainDocumentPart.Document.Body.InnerText, StringComparison.Ordinal);
        Assert.Empty(new OpenXmlValidator().Validate(output));
    }

    [Fact]
    public void SetSideEffectsAreCommittedOnlyForSelectedBranch()
    {
        byte[] selected = CreateConditionalSetDocument(condition: true);
        byte[] unselected = CreateConditionalSetDocument(condition: false);

        using var selectedOutput = Open(DxpDocxExport.Export(
            selected, SemanticOptions(), Logger, new DxpFieldEval(logger: Logger)));
        using var unselectedOutput = Open(DxpDocxExport.Export(
            unselected, SemanticOptions(), Logger, new DxpFieldEval(logger: Logger)));

        Assert.Contains("selected", selectedOutput.MainDocumentPart!.Document.Body!.InnerText, StringComparison.Ordinal);
        Assert.DoesNotContain("selected", unselectedOutput.MainDocumentPart!.Document.Body!.InnerText, StringComparison.Ordinal);
    }

    [Fact]
    public void NestedIfComposesWithAdjacentLiteralInSourceOrder()
    {
        byte[] source = CreateNestedIfDocument();
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.SetDocVariable("Count", "2");

        using var output = Open(DxpDocxExport.Export(source, SemanticOptions(), Logger, eval));
        string text = output.MainDocumentPart!.Document.Body!.InnerText;

        Assert.Contains("The items request the synthetic continuation.", text, StringComparison.Ordinal);
        Assert.DoesNotContain("continuation.The items", text, StringComparison.Ordinal);
        Assert.Empty(new OpenXmlValidator().Validate(output));
    }

    [Fact]
    public void NestedDatabaseAcrossParagraphBoundaryDoesNotLeakOrDuplicateContent()
    {
        byte[] source = CreateCrossParagraphDatabaseDocument();
        byte[] docx = DxpDocxExport.Export(source, SemanticOptions(), Logger, CreateEval());

        using var output = Open(docx);
        Body body = output.MainDocumentPart!.Document.Body!;
        Assert.Single(body.Elements<Table>());
        Assert.Equal(1, body.InnerText.Split("Alice", StringSplitOptions.None).Length - 1);
        Assert.DoesNotContain(body.Descendants<Text>(), text =>
            text.Text.IndexOfAny(new[] { '\r', '\n', '\t' }) >= 0);
        Assert.Empty(new OpenXmlValidator().Validate(output));
    }

    [Fact]
    public void SelectedIfBranchPreservesSemanticParagraphBoundary()
    {
        byte[] source = CreateCrossParagraphLiteralDocument();
        using var output = Open(DxpDocxExport.Export(
            source, SemanticOptions(), Logger, new DxpFieldEval(logger: Logger)));

        string[] paragraphs = output.MainDocumentPart!.Document.Body!
            .Elements<Paragraph>()
            .Select(static paragraph => paragraph.InnerText)
            .Where(static text => text.Length > 0)
            .ToArray();
        Assert.Equal(new[] { "FIRST", "SECOND", "AFTER" }, paragraphs);
        Assert.Empty(new OpenXmlValidator().Validate(output));
    }

    private static DxpExportOptions SemanticOptions() => new()
    {
        FieldEvalMode = DxpFieldEvalExportMode.Evaluate,
        UseSemanticFieldResults = true
    };

    private DxpFieldEval CreateEval()
    {
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.DatabaseProvider = new SyntheticDatabaseProvider();
        return eval;
    }

    private static byte[] CreateNestedDatabaseDocument(bool condition)
        => CreateDocument(body => body.Append(
            new Paragraph(
                Begin(),
                Code($" IF 1 = {(condition ? 1 : 0)} \""),
                Begin(),
                Code(" DATABASE \\h \\s \"SELECT Name, Address FROM SyntheticPeople\" "),
                End(),
                Code("\" \"\" "),
                End()),
            new Paragraph(new Run(new Text("AFTER")))));

    private static byte[] CreateConditionalSetDocument(bool condition)
        => CreateDocument(body => body.Append(
            new Paragraph(
                Begin(),
                Code($" IF 1 = {(condition ? 1 : 0)} \""),
                Begin(),
                Code(" SET Result \"selected\" "),
                End(),
                Code("\" \"\" "),
                End(),
                Begin(),
                Code(" REF Result "),
                End())));

    private static byte[] CreateNestedIfDocument()
        => CreateDocument(body => body.Append(
            new Paragraph(
                Begin(),
                Code(" IF 1 = 1 \""),
                Begin(),
                Code(" IF "),
                Begin(),
                Code(" DOCVARIABLE Count "),
                End(),
                Code(" = 1 \"The item requests\" \"The items request\" "),
                End(),
                Code(" the synthetic continuation.\" \"\" "),
                End())));

    private static byte[] CreateCrossParagraphDatabaseDocument()
        => CreateDocument(body => body.Append(
            new Paragraph(
                Begin(),
                Code(" IF 1 = 1 \""),
                Begin(),
                Code(" DATABASE \\h \\s \"SELECT Name, Address FROM SyntheticPeople\" ")),
            new Paragraph(
                End(),
                Code("\" \"\" "),
                End()),
            new Paragraph(new Run(new Text("AFTER")))));

    private static byte[] CreateCrossParagraphLiteralDocument()
        => CreateDocument(body => body.Append(
            new Paragraph(
                Begin(),
                Code(" IF 1 = 1 \"FIRST")),
            new Paragraph(
                Code("SECOND\" \"\" "),
                End()),
            new Paragraph(new Run(new Text("AFTER")))));

    private static Run Begin() => new(new FieldChar { FieldCharType = FieldCharValues.Begin });
    private static Run End() => new(new FieldChar { FieldCharType = FieldCharValues.End });
    private static Run Code(string text) => new(new FieldCode { Text = text });

    private static byte[] CreateDocument(Action<Body> build)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var body = new Body();
            build(body);
            body.AppendChild(new SectionProperties());
            main.Document = new Document(body);
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static WordprocessingDocument Open(byte[] bytes)
        => WordprocessingDocument.Open(new MemoryStream(bytes), false);

    private sealed class SyntheticDatabaseProvider : IDatabaseFieldProvider
    {
        public Task<DxpDatabaseResult?> ExecuteAsync(
            DxpDatabaseRequest request,
            CancellationToken cancellationToken)
        {
            _ = request;
            _ = cancellationToken;
            DxpDatabaseResult result = new(
                new[] { new DxpDatabaseColumn("Name"), new DxpDatabaseColumn("Address") },
                new IReadOnlyList<DxpFieldValue?>[]
                {
                    new DxpFieldValue?[] { new("Alice"), new("Montreal\nQuebec") },
                    new DxpFieldValue?[] { new("Bob"), new("Toronto") }
                });
            return Task.FromResult<DxpDatabaseResult?>(result);
        }
    }
}
