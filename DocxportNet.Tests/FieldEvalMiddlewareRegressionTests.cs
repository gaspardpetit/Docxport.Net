using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocxportNet.Tests.Utils;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.PlainText;
using DocxportNet.Walker;
using DocxportNet.Middleware;
using Xunit.Abstractions;

namespace DocxportNet.Tests;

public sealed class FieldEvalMiddlewareRegressionTests : TestBase<FieldEvalMiddlewareRegressionTests>
{
    private static readonly string ProjectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
    private static readonly string SamplesDirectory = Path.Combine(ProjectRoot, "samples");

    public FieldEvalMiddlewareRegressionTests(ITestOutputHelper output) : base(output)
    {
    }

    [Fact]
    public void Eval_TestFields_InlineIfPreservesFormatting()
    {
        string docxPath = Path.Combine(SamplesDirectory, "TestFields.docx");
        var config = DxpHtmlVisitorConfig.CreateRichConfig();
        var visitor = new DxpHtmlVisitor(config, Logger);
        var options = new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate };

        string html = TestCompare.Normalize(DxpExport.ExportToString(docxPath, visitor, options, Logger));

        Assert.Contains("Expect No Error: Not Empty", html, StringComparison.Ordinal);
        Assert.Contains("Expect <strong class=\"dxp-bold\">one</strong> (bold):", html, StringComparison.Ordinal);
        Assert.Contains("Expect <strong class=\"dxp-bold\">1</strong><span class=\"dxp-underline\">2</span><strong class=\"dxp-bold\">3: 1</strong><span class=\"dxp-underline\">2</span><strong class=\"dxp-bold\">3</strong>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Eval_DocVariableWithoutSeparate_EmitsStyledValue()
    {
        using var doc = CreateDocVariableDoc("TokenLabel", "VALUE");
        var config = DxpHtmlVisitorConfig.CreateRichConfig();
        var visitor = new DxpHtmlVisitor(config, Logger);
        visitor.FieldEval.Context.SetDocVariable("TokenLabel", "VALUE");
        var options = new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate };

        string html = TestCompare.Normalize(DxpExport.ExportToString(doc, visitor, options, Logger));

        Assert.Contains("<strong class=\"dxp-bold\">VALUE</strong>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Factory_EvaluateMiddleware_EmitsStyledValue()
    {
        using var doc = CreateDocVariableDoc("TokenLabel", "VALUE");
        var visitor = new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig(), Logger);
        visitor.FieldEval.Context.SetDocVariable("TokenLabel", "VALUE");
        using var writer = new StringWriter();
        visitor.SetOutput(writer);

        var pipeline = DxpVisitorMiddleware.Chain(
            visitor,
            next => DxpFieldEvalMiddleware.CreateEvaluatedFieldMiddleware(next, visitor.FieldEval, logger: Logger),
            next => new DxpContextMiddleware(next, Logger));

        new DxpWalker(Logger).Accept(doc, pipeline);

        string html = TestCompare.Normalize(writer.ToString());
        Assert.Contains("<strong class=\"dxp-bold\">VALUE</strong>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Factory_CachedMiddleware_ReplaysCachedValue()
    {
        using var doc = CreateSimpleFieldDoc(" FOO ", "cached text");
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig(), Logger);
        using var writer = new StringWriter();
        visitor.SetOutput(writer);

        var pipeline = DxpVisitorMiddleware.Chain(
            visitor,
            next => DxpFieldEvalMiddleware.CreateCachedFieldMiddleware(next, visitor.FieldEval, logger: Logger),
            next => new DxpContextMiddleware(next, Logger));

        new DxpWalker(Logger).Accept(doc, pipeline);

        Assert.Contains("cached text", writer.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void Factory_CachedMiddleware_AutoNumWithoutSeparate_EmitsSequenceLabel()
    {
        using var doc = CreateComplexAutoNumDoc();
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig(), Logger);
        using var writer = new StringWriter();
        visitor.SetOutput(writer);

        var pipeline = DxpVisitorMiddleware.Chain(
            visitor,
            next => DxpFieldEvalMiddleware.CreateCachedFieldMiddleware(next, visitor.FieldEval, logger: Logger),
            next => new DxpContextMiddleware(next, Logger));

        new DxpWalker(Logger).Accept(doc, pipeline);

        Assert.Contains("1.\tClaim text", writer.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void Factory_CachedMiddleware_AutoNum_MixedCachedAndUncached_StaysSequential()
    {
        using var doc = CreateComplexMixedAutoNumDoc();
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig(), Logger);
        using var writer = new StringWriter();
        visitor.SetOutput(writer);

        var pipeline = DxpVisitorMiddleware.Chain(
            visitor,
            next => DxpFieldEvalMiddleware.CreateCachedFieldMiddleware(next, visitor.FieldEval, logger: Logger),
            next => new DxpContextMiddleware(next, Logger));

        new DxpWalker(Logger).Accept(doc, pipeline);

        var output = writer.ToString();
        Assert.Contains("1.\tFirst claim", output, StringComparison.Ordinal);
        Assert.Contains("2.\tSecond claim", output, StringComparison.Ordinal);
    }

    [Fact]
    public void Factory_CachedMiddleware_AutoNumWithoutSeparate_AppliesFormatAndSeparator()
    {
        using var doc = CreateComplexAutoNumDoc(" AUTONUM \\* Roman \\s- ");
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig(), Logger);
        using var writer = new StringWriter();
        visitor.SetOutput(writer);

        var pipeline = DxpVisitorMiddleware.Chain(
            visitor,
            next => DxpFieldEvalMiddleware.CreateCachedFieldMiddleware(next, visitor.FieldEval, logger: Logger),
            next => new DxpContextMiddleware(next, Logger));

        new DxpWalker(Logger).Accept(doc, pipeline);

        Assert.Contains("I-\tClaim text", writer.ToString(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("AUTONUMLGL", "1.|1.1.|1.1.1.")]
    [InlineData("AUTONUMOUT", "I.|I.A.|I.A.1.")]
    public void Factory_EvaluateMiddleware_AutoNumberHierarchy_FollowsHeadingLevels(string fieldType, string expected)
    {
        using var doc = CreateAutoNumberHierarchyDoc(fieldType);
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig(), Logger);
        using var writer = new StringWriter();
        visitor.SetOutput(writer);

        var pipeline = DxpVisitorMiddleware.Chain(
            visitor,
            next => DxpFieldEvalMiddleware.CreateEvaluatedFieldMiddleware(next, visitor.FieldEval, logger: Logger),
            next => new DxpContextMiddleware(next, Logger));

        new DxpWalker(Logger).Accept(doc, pipeline);

        string output = writer.ToString();
        int position = 0;
        foreach (string label in expected.Split('|'))
        {
            int found = output.IndexOf(label, position, StringComparison.Ordinal);
            Assert.True(found >= position, $"Expected label '{label}' after offset {position}. Output: {output}");
            position = found + label.Length;
        }
    }

    [Fact]
    public void Factory_CachedMiddleware_EmptySimpleAutoNum_SynthesizesLabel()
    {
        using var doc = CreateSimpleFieldDoc(" AUTONUM ", string.Empty);
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig(), Logger);
        using var writer = new StringWriter();
        visitor.SetOutput(writer);

        var pipeline = DxpVisitorMiddleware.Chain(
            visitor,
            next => DxpFieldEvalMiddleware.CreateCachedFieldMiddleware(next, visitor.FieldEval, logger: Logger),
            next => new DxpContextMiddleware(next, Logger));

        new DxpWalker(Logger).Accept(doc, pipeline);

        Assert.Contains("1.", writer.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void Eval_WordAuthoredAutoNumberCompatibilityFixture_PreservesAllFamilies()
    {
        string docxPath = Path.Combine(SamplesDirectory, "AutoNumberCompatibility.docx");
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig(), Logger);
        var options = new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate };

        string text = DxpExport.ExportToString(docxPath, visitor, options, Logger);

        Assert.DoesNotContain("Error! Invalid field code.", text, StringComparison.Ordinal);
        Assert.Contains("1.AUTONUM level 1", text, StringComparison.Ordinal);
        Assert.Contains("1.1.1.1.AUTONUMLGL level 0", text, StringComparison.Ordinal);
        Assert.Contains("I.A.1.a.AUTONUMOUT level 0", text, StringComparison.Ordinal);
    }

    private static WordprocessingDocument CreateDocVariableDoc(string name, string value)
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var runProps = new RunProperties(new Bold());
            var paragraph = new Paragraph(
                new Run(runProps.CloneNode(true), new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(runProps.CloneNode(true), new FieldCode { Text = $" DOCVARIABLE {name} " }),
                new Run(runProps.CloneNode(true), new FieldChar { FieldCharType = FieldCharValues.End })
            );

            mainPart.Document.Body!.Append(paragraph);
            mainPart.Document.Save();
            document.Save();
        }

        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }

    private static WordprocessingDocument CreateSimpleFieldDoc(string instruction, string cachedText)
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var mainPart = document.AddMainDocumentPart();
            var field = new SimpleField { Instruction = instruction };
            field.Append(new Run(new Text(cachedText)));
            mainPart.Document = new Document(new Body(new Paragraph(field)));
            mainPart.Document.Save();
            document.Save();
        }

        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }

    private static WordprocessingDocument CreateComplexAutoNumDoc(string instruction = " AUTONUM ")
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var paragraph = new Paragraph(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode { Text = instruction }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
                new Run(new TabChar(), new Text("Claim text")));

            mainPart.Document.Body!.Append(paragraph);
            mainPart.Document.Save();
            document.Save();
        }

        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }

    private static WordprocessingDocument CreateComplexMixedAutoNumDoc()
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var first = new Paragraph(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode { Text = " AUTONUM " }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("1.")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
                new Run(new TabChar(), new Text("First claim")));

            var second = new Paragraph(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode { Text = " AUTONUM " }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
                new Run(new TabChar(), new Text("Second claim")));

            mainPart.Document.Body!.Append(first, second);
            mainPart.Document.Save();
            document.Save();
        }

        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }

    private static WordprocessingDocument CreateAutoNumberHierarchyDoc(string fieldType)
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var mainPart = document.AddMainDocumentPart();
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(
                new Style(new StyleName { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true },
                new Style(new StyleName { Val = "heading 1" }, new BasedOn { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Heading1" },
                new Style(new StyleName { Val = "heading 2" }, new BasedOn { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Heading2" });

            static Paragraph FieldParagraph(string instruction, string? styleId = null)
            {
                var paragraph = new Paragraph(
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                    new Run(new FieldCode { Text = $" {instruction} " }),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
                if (styleId != null)
                    paragraph.PrependChild(new ParagraphProperties(new ParagraphStyleId { Val = styleId }));
                return paragraph;
            }

            mainPart.Document = new Document(new Body(
                FieldParagraph(fieldType, "Heading1"),
                FieldParagraph(fieldType, "Heading2"),
                FieldParagraph(fieldType)));
            stylesPart.Styles.Save();
            mainPart.Document.Save();
            document.Save();
        }

        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }
}
