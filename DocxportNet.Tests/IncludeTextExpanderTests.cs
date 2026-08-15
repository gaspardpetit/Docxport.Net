using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Fields;
using DocxportNet.Fields.Resolution;

namespace DocxportNet.Tests;

public class IncludeTextExpanderTests
{
    [Fact]
    public void Expand_ExpandsIncludeTextAndPreservesOtherFields()
    {
        byte[] child = CreateDocument(new Paragraph(new Run(new Text("CHILD"))));
        byte[] parent = CreateDocument(
            new Paragraph(
                new Run(new Text("Before ")),
                new SimpleField(new Run(new Text("include-cache"))) { Instruction = " INCLUDETEXT \"memory:child\" " },
                new Run(new Text(" after"))),
            new Paragraph(
                new SimpleField(new Run(new Text("variable-cache"))) { Instruction = " DOCVARIABLE KeepMe " }));

        byte[] result = DxpIncludeTextExpander.Expand(parent, new StaticResolver(child));

        using var stream = new MemoryStream(result);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, false);
        Body body = document.MainDocumentPart!.Document.Body!;
        Assert.Contains("Before CHILD after", body.InnerText, StringComparison.Ordinal);
        SimpleField variable = Assert.Single(body.Descendants<SimpleField>());
        Assert.Contains("DOCVARIABLE", variable.Instruction!.Value, StringComparison.Ordinal);
        Assert.Equal("variable-cache", variable.InnerText);
    }

    [Fact]
    public void Expand_PropagatesCancellationToResolver()
    {
        byte[] parent = CreateDocument(new Paragraph(
            new SimpleField(new Run(new Text("cached"))) { Instruction = " INCLUDETEXT \"memory:child\" " }));
        using var cancellation = new CancellationTokenSource();
        var resolver = new CancellingResolver(cancellation);

        Assert.ThrowsAny<OperationCanceledException>(() =>
            DxpIncludeTextExpander.Expand(parent, resolver, cancellation.Token));
        Assert.True(resolver.ReceivedToken);
    }

    [Fact]
    public void Expand_ImportsAndRemapsNumberingDefinitions()
    {
        byte[] child = CreateNumberedDocument("CHILD", 1);
        byte[] parent = CreateNumberedParent(CreateIncludeField("memory:child"), 1);

        byte[] result = DxpIncludeTextExpander.Expand(parent, new StaticResolver(child));

        using var stream = new MemoryStream(result);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, false);
        MainDocumentPart main = document.MainDocumentPart!;
        Paragraph childParagraph = main.Document.Body!.Elements<Paragraph>()
            .Single(value => value.InnerText == "CHILD");
        int importedId = childParagraph.ParagraphProperties!.NumberingProperties!
            .NumberingId!.Val!.Value;
        Assert.NotEqual(1, importedId);
        Numbering numbering = main.NumberingDefinitionsPart!.Numbering!;
        NumberingInstance instance = Assert.Single(numbering.Elements<NumberingInstance>(),
            value => value.NumberID!.Value == importedId);
        Assert.Contains(numbering.Elements<AbstractNum>(),
            value => value.AbstractNumberId!.Value == instance.AbstractNumId!.Val!.Value);
    }

    private static byte[] CreateDocument(params OpenXmlElement[] blocks)
    {
        using var stream = new MemoryStream();
        using (WordprocessingDocument document = WordprocessingDocument.Create(
            stream, WordprocessingDocumentType.Document, true))
        {
            MainDocumentPart main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(blocks));
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static SimpleField CreateIncludeField(string path) =>
        new(new Run(new Text("cached"))) { Instruction = $" INCLUDETEXT \"{path}\" " };

    private static byte[] CreateNumberedDocument(string text, int numberId)
    {
        using var stream = new MemoryStream();
        using (WordprocessingDocument document = WordprocessingDocument.Create(
            stream, WordprocessingDocumentType.Document, true))
        {
            MainDocumentPart main = document.AddMainDocumentPart();
            AddNumbering(main, numberId);
            main.Document = new Document(new Body(new Paragraph(
                new ParagraphProperties(new NumberingProperties(
                    new NumberingLevelReference { Val = 0 }, new NumberingId { Val = numberId })),
                new Run(new Text(text)))));
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] CreateNumberedParent(SimpleField include, int numberId)
    {
        using var stream = new MemoryStream();
        using (WordprocessingDocument document = WordprocessingDocument.Create(
            stream, WordprocessingDocumentType.Document, true))
        {
            MainDocumentPart main = document.AddMainDocumentPart();
            AddNumbering(main, numberId);
            main.Document = new Document(new Body(new Paragraph(include)));
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static void AddNumbering(MainDocumentPart main, int numberId)
    {
        NumberingDefinitionsPart part = main.AddNewPart<NumberingDefinitionsPart>();
        part.Numbering = new Numbering(
            new AbstractNum(new Level(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = NumberFormatValues.Decimal },
                new LevelText { Val = "%1." }) { LevelIndex = 0 }) { AbstractNumberId = numberId },
            new NumberingInstance(new AbstractNumId { Val = numberId }) { NumberID = numberId });
    }

    private sealed class StaticResolver(byte[] content) : IDxpIncludeTextResolver
    {
        public Task<DxpIncludeTextSource?> ResolveAsync(DxpIncludeTextRequest request,
            DxpFieldEvalContext context, CancellationToken cancellationToken = default)
            => Task.FromResult<DxpIncludeTextSource?>(new(request.Path, content)
            {
                Format = DxpIncludeTextSourceFormat.Docx
            });
    }

    private sealed class CancellingResolver(CancellationTokenSource cancellation) : IDxpIncludeTextResolver
    {
        public bool ReceivedToken { get; private set; }

        public Task<DxpIncludeTextSource?> ResolveAsync(DxpIncludeTextRequest request,
            DxpFieldEvalContext context, CancellationToken cancellationToken = default)
        {
            ReceivedToken = cancellationToken == cancellation.Token;
            cancellation.Cancel();
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult<DxpIncludeTextSource?>(null);
        }
    }
}
