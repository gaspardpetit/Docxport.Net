using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Tests.Utils;
using Xunit.Abstractions;

namespace DocxportNet.Tests;

public sealed class DxpTransformTests : TestBase<DxpTransformTests>
{
    public DxpTransformTests(ITestOutputHelper output) : base(output)
    {
    }

    [Fact]
    public void Transform_NoOpPreservesSupportedPartsAndLeavesCommentsUntouched()
    {
        var before = CreateFixtureBytes();
        var after = DxpTransform.Transform(before, new NoOpTransformer(), Logger);

        using var beforeDoc = OpenRead(before);
        using var afterDoc = OpenRead(after);

        Assert.Equal(
            NormalizeXml(beforeDoc.MainDocumentPart!.Document!.OuterXml),
            NormalizeXml(afterDoc.MainDocumentPart!.Document!.OuterXml));
        Assert.Equal(
            NormalizeXml(beforeDoc.MainDocumentPart.HeaderParts.Single().Header!.OuterXml),
            NormalizeXml(afterDoc.MainDocumentPart.HeaderParts.Single().Header!.OuterXml));
        Assert.Equal(
            NormalizeXml(beforeDoc.MainDocumentPart.FooterParts.Single().Footer!.OuterXml),
            NormalizeXml(afterDoc.MainDocumentPart.FooterParts.Single().Footer!.OuterXml));
        Assert.Equal(
            NormalizeXml(beforeDoc.MainDocumentPart.FootnotesPart!.Footnotes!.OuterXml),
            NormalizeXml(afterDoc.MainDocumentPart.FootnotesPart!.Footnotes!.OuterXml));
        Assert.Equal(
            NormalizeXml(beforeDoc.MainDocumentPart.EndnotesPart!.Endnotes!.OuterXml),
            NormalizeXml(afterDoc.MainDocumentPart.EndnotesPart!.Endnotes!.OuterXml));
        Assert.Equal(
            NormalizeXml(beforeDoc.MainDocumentPart.WordprocessingCommentsPart!.Comments!.OuterXml),
            NormalizeXml(afterDoc.MainDocumentPart.WordprocessingCommentsPart!.Comments!.OuterXml));
    }

    [Fact]
    public void Transform_KeepWithoutDescendSkipsChildrenAndPreservesSubtree()
    {
        var transformer = new SkipParagraphChildrenTransformer();
        var bytes = DxpTransform.Transform(CreateFixtureBytes(), transformer, Logger);

        using var doc = OpenRead(bytes);
        var paragraphs = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToList();

        Assert.DoesNotContain("B1", transformer.VisitedRunTexts);
        Assert.DoesNotContain("B2", transformer.VisitedRunTexts);
        Assert.Contains("C1", transformer.VisitedRunTexts);
        Assert.Equal("B1B2", paragraphs[0].InnerText);
    }

    [Fact]
    public void Transform_RemoveDeletesTargetedNodeOnly()
    {
        var bytes = DxpTransform.Transform(CreateFixtureBytes(), new RemoveBodyParagraphTransformer("C1"), Logger);

        using var doc = OpenRead(bytes);
        var paragraphs = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Select(GetParagraphText).ToList();

        Assert.Equal(new[] { "B1B2" }, paragraphs);
    }

    [Fact]
    public void Transform_ReplaceInsertsNodesInSiblingPosition()
    {
        var bytes = DxpTransform.Transform(CreateFixtureBytes(), new ReplaceBodyParagraphTransformer("C1"), Logger);

        using var doc = OpenRead(bytes);
        var paragraphs = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Select(GetParagraphText).ToList();

        Assert.Equal(new[] { "B1B2", "X", "Y" }, paragraphs);
    }

    [Fact]
    public void Transform_TraversalOrderAcrossPartsIsDeterministic()
    {
        var recorder = new ParagraphRecorderTransformer();
        _ = DxpTransform.Transform(CreateFixtureBytes(), recorder, Logger);

        Assert.Equal(
            new[] { "MainDocument:B1B2", "MainDocument:C1", "Header:H1", "Footer:F1", "Footnote:FN1", "Endnote:EN1" },
            recorder.Paragraphs);
    }

    [Fact]
    public void Transform_ContextReportsPartDepthSiblingOrdinalAndPath()
    {
        var recorder = new ContextRecorderTransformer("B2");
        _ = DxpTransform.Transform(CreateFixtureBytes(), recorder, Logger);

        Assert.NotNull(recorder.Context);
        Assert.Equal(DxpTransformPartKind.MainDocument, recorder.Context!.PartKind);
        Assert.Equal(1, recorder.Context.Depth);
        Assert.Equal(1, recorder.Context.SiblingIndex);
        Assert.Equal("Body/Paragraph[0]/Run[1]", recorder.Context.Path);
        Assert.Equal(6, recorder.Context.NodeOrdinal);
        Assert.Single(recorder.Context.Ancestors);
        Assert.IsType<Paragraph>(recorder.Context.Ancestors[0]);
    }

    [Fact]
    public void Transform_FileHelperWritesOutputWithoutMutatingInput()
    {
        string tempDirectory = Path.Combine(Path.GetTempPath(), "docxport-transform-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDirectory);
        string inputPath = Path.Combine(tempDirectory, "input.docx");
        string outputPath = Path.Combine(tempDirectory, "output.docx");

        try
        {
            File.WriteAllBytes(inputPath, CreateFixtureBytes());
            var originalInput = File.ReadAllBytes(inputPath);

            DxpTransform.Transform(inputPath, outputPath, new DxpSimplifyParagraphRunsTransformer(), Logger);

            var currentInput = File.ReadAllBytes(inputPath);
            Assert.Equal(originalInput, currentInput);

            using var inputDoc = WordprocessingDocument.Open(inputPath, false);
            using var outputDoc = WordprocessingDocument.Open(outputPath, false);

            Assert.Equal(2, inputDoc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First().Elements<Run>().Count());
            Assert.Single(outputDoc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First().Elements<Run>());
        }
        finally
        {
            if (Directory.Exists(tempDirectory))
                Directory.Delete(tempDirectory, recursive: true);
        }
    }

    [Fact]
    public void CollapseEquivalentRuns_MergesAdjacentRunsWithIdenticalRunProperties()
    {
        var bytes = CreateFixtureBytes(
            new Paragraph(
                new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                StyledRun("A", new Bold()),
                StyledRun("B", new Bold()),
                new Run(new Text("C"))));

        using var transformed = OpenRead(DxpTransform.Transform(bytes, new DxpCollapseEquivalentRunsTransformer(), Logger));
        var paragraph = transformed.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
        var runs = paragraph.Elements<Run>().ToList();

        Assert.Equal(2, runs.Count);
        Assert.Equal("AB", runs[0].InnerText);
        Assert.NotNull(runs[0].RunProperties?.Bold);
        Assert.Equal("C", runs[1].InnerText);
        Assert.Equal(JustificationValues.Center, paragraph.ParagraphProperties?.Justification?.Val?.Value);
    }

    [Fact]
    public void CollapseEquivalentRuns_DoesNotMergeDifferentRunProperties()
    {
        var bytes = CreateFixtureBytes(
            new Paragraph(
                StyledRun("A", new Bold()),
                StyledRun("B", new Italic())));

        using var transformed = OpenRead(DxpTransform.Transform(bytes, new DxpCollapseEquivalentRunsTransformer(), Logger));
        var runs = transformed.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First().Elements<Run>().ToList();

        Assert.Equal(2, runs.Count);
        Assert.Equal("A", runs[0].InnerText);
        Assert.Equal("B", runs[1].InnerText);
    }

    [Fact]
    public void SimplifyParagraphRuns_CollapsesParagraphToSingleUnformattedRun()
    {
        using var transformed = OpenRead(DxpTransform.Transform(CreateFixtureBytes(), new DxpSimplifyParagraphRunsTransformer(), Logger));
        var firstParagraph = transformed.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
        var runs = firstParagraph.Elements<Run>().ToList();

        Assert.Single(runs);
        Assert.Null(runs[0].RunProperties);
        Assert.Equal("B1B2", firstParagraph.InnerText);
    }

    [Fact]
    public void SampleTransformers_OperateAcrossBodyHeaderFooterFootnoteAndEndnote()
    {
        using var collapseDoc = OpenRead(DxpTransform.Transform(CreateFixtureBytes(), new DxpCollapseEquivalentRunsTransformer(), Logger));
        using var simplifyDoc = OpenRead(DxpTransform.Transform(CreateFixtureBytes(), new DxpSimplifyParagraphRunsTransformer(), Logger));

        Assert.All(GetStoryParagraphs(collapseDoc), static p => Assert.True(p.Elements<Run>().Count() >= 1));
        Assert.All(GetStoryParagraphs(simplifyDoc), static p => Assert.Single(p.Elements<Run>()));
    }

    [Fact]
    public void SampleTransformers_AggressivelyFlattenComplexInlineContent()
    {
        byte[] bytes = CreateFixtureBytes(
            new Paragraph(
                new Hyperlink(new Run(new Text("Link"))) { Anchor = "Anchor1" },
                new SimpleField { Instruction = " DATE " },
                new Run(new Text("Tail"))));

        using var collapseDoc = OpenRead(DxpTransform.Transform(bytes, new DxpCollapseEquivalentRunsTransformer(), Logger));
        var collapseParagraph = collapseDoc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
        Assert.All(collapseParagraph.ChildElements, static child => Assert.True(child is ParagraphProperties || child is Run));
        Assert.Equal("LinkTail", collapseParagraph.InnerText);

        using var simplifyDoc = OpenRead(DxpTransform.Transform(bytes, new DxpSimplifyParagraphRunsTransformer(), Logger));
        var simplifyParagraph = simplifyDoc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
        Assert.All(simplifyParagraph.ChildElements, static child => Assert.True(child is ParagraphProperties || child is Run));
        Assert.Equal("LinkTail", simplifyParagraph.InnerText);
    }

    private static IReadOnlyList<Paragraph> GetStoryParagraphs(WordprocessingDocument doc)
    {
        var paragraphs = new List<Paragraph>();
        paragraphs.AddRange(doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>());
        paragraphs.AddRange(doc.MainDocumentPart.HeaderParts.SelectMany(static part => part.Header?.Elements<Paragraph>() ?? Enumerable.Empty<Paragraph>()));
        paragraphs.AddRange(doc.MainDocumentPart.FooterParts.SelectMany(static part => part.Footer?.Elements<Paragraph>() ?? Enumerable.Empty<Paragraph>()));
        paragraphs.AddRange(doc.MainDocumentPart.FootnotesPart!.Footnotes!.Elements<Footnote>().Where(static fn => fn.Id?.Value > 0).SelectMany(static fn => fn.Elements<Paragraph>()));
        paragraphs.AddRange(doc.MainDocumentPart.EndnotesPart!.Endnotes!.Elements<Endnote>().Where(static en => en.Id?.Value > 0).SelectMany(static en => en.Elements<Paragraph>()));
        return paragraphs;
    }

    private static string NormalizeXml(string xml) => TestCompare.Normalize(xml);

    private static string GetParagraphText(Paragraph paragraph) => paragraph.InnerText;

    private static WordprocessingDocument OpenRead(byte[] bytes)
    {
        var stream = new MemoryStream(bytes);
        return WordprocessingDocument.Open(stream, false);
    }

    private static Run StyledRun(string text, params OpenXmlElement[] properties)
    {
        var run = new Run(new Text(text));
        if (properties.Length > 0)
            run.RunProperties = new RunProperties(properties);
        return run;
    }

    private static byte[] CreateFixtureBytes(Paragraph? firstParagraph = null)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var headerPart = main.AddNewPart<HeaderPart>();
            var footerPart = main.AddNewPart<FooterPart>();
            var commentsPart = main.AddNewPart<WordprocessingCommentsPart>();
            var footnotesPart = main.AddNewPart<FootnotesPart>();
            var endnotesPart = main.AddNewPart<EndnotesPart>();

            headerPart.Header = new Header(new Paragraph(new Run(new Text("H1"))));
            footerPart.Footer = new Footer(new Paragraph(new Run(new Text("F1"))));
            commentsPart.Comments = new Comments(
                new Comment(new Paragraph(new Run(new Text("Comment1")))) { Id = "0", Author = "tester" });
            footnotesPart.Footnotes = new Footnotes(
                new Footnote(new Paragraph(new Run(new FootnoteReferenceMark()))) { Id = -1, Type = FootnoteEndnoteValues.Separator },
                new Footnote(new Paragraph(new Run(new Text("FN1")))) { Id = 1 });
            endnotesPart.Endnotes = new Endnotes(
                new Endnote(new Paragraph(new Run(new Text("EN1")))) { Id = 1 });

            string headerId = main.GetIdOfPart(headerPart);
            string footerId = main.GetIdOfPart(footerPart);
            main.Document = new Document(
                new Body(
                    (Paragraph)(firstParagraph?.CloneNode(true) ?? new Paragraph(
                        StyledRun("B1", new Bold()),
                        StyledRun("B2", new Bold()))),
                    new Paragraph(new Run(new Text("C1"))),
                    new SectionProperties(
                        new HeaderReference { Id = headerId, Type = HeaderFooterValues.Default },
                        new FooterReference { Id = footerId, Type = HeaderFooterValues.Default })));
            main.Document.Save();
            headerPart.Header.Save();
            footerPart.Footer.Save();
            commentsPart.Comments.Save();
            footnotesPart.Footnotes.Save();
            endnotesPart.Endnotes.Save();
        }

        return stream.ToArray();
    }

    private sealed class NoOpTransformer : IDxpNodeTransformer
    {
        public DxpTransformDecision Visit(OpenXmlElement node, DxpTransformContext context) => DxpTransformDecision.Keep();
    }

    private sealed class SkipParagraphChildrenTransformer : IDxpNodeTransformer
    {
        public List<string> VisitedRunTexts { get; } = new();

        public DxpTransformDecision Visit(OpenXmlElement node, DxpTransformContext context)
        {
            if (node is Paragraph paragraph && paragraph.InnerText == "B1B2")
                return DxpTransformDecision.Keep(descend: false);
            if (node is Run run)
                VisitedRunTexts.Add(run.InnerText);
            return DxpTransformDecision.Keep();
        }
    }

    private sealed class RemoveBodyParagraphTransformer : IDxpNodeTransformer
    {
        private readonly string _text;

        public RemoveBodyParagraphTransformer(string text)
        {
            _text = text;
        }

        public DxpTransformDecision Visit(OpenXmlElement node, DxpTransformContext context)
        {
            if (context.PartKind == DxpTransformPartKind.MainDocument && node is Paragraph paragraph && paragraph.InnerText == _text)
                return DxpTransformDecision.Remove();
            return DxpTransformDecision.Keep();
        }
    }

    private sealed class ReplaceBodyParagraphTransformer : IDxpNodeTransformer
    {
        private readonly string _text;

        public ReplaceBodyParagraphTransformer(string text)
        {
            _text = text;
        }

        public DxpTransformDecision Visit(OpenXmlElement node, DxpTransformContext context)
        {
            if (context.PartKind == DxpTransformPartKind.MainDocument && node is Paragraph paragraph && paragraph.InnerText == _text)
            {
                return DxpTransformDecision.Replace(
                    new Paragraph(new Run(new Text("X"))),
                    new Paragraph(new Run(new Text("Y"))));
            }

            return DxpTransformDecision.Keep();
        }
    }

    private sealed class ParagraphRecorderTransformer : IDxpNodeTransformer
    {
        public List<string> Paragraphs { get; } = new();

        public DxpTransformDecision Visit(OpenXmlElement node, DxpTransformContext context)
        {
            if (node is Paragraph paragraph)
                Paragraphs.Add($"{context.PartKind}:{paragraph.InnerText}");
            return DxpTransformDecision.Keep();
        }
    }

    private sealed class ContextRecorderTransformer : IDxpNodeTransformer
    {
        private readonly string _text;

        public ContextRecorderTransformer(string text)
        {
            _text = text;
        }

        public DxpTransformContext? Context { get; private set; }

        public DxpTransformDecision Visit(OpenXmlElement node, DxpTransformContext context)
        {
            if (node is Run run && run.InnerText == _text)
                Context = context;
            return DxpTransformDecision.Keep();
        }
    }
}
