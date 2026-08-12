using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocxportNet.Tests.Utils;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.PlainText;
using DocxportNet.Walker;
using DocxportNet.Middleware;
using DocxportNet.Fields.Resolution;
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

    [Fact]
    public void Eval_IncludeText_WalksChildBodyAndResumesParent()
    {
        byte[] child = CreateTextDocBytes("CHILD");
        using var parent = CreateIncludeTextDoc("child.docx", "CACHE", "BEFORE", "AFTER");
        var resolver = new MemoryIncludeTextResolver(("child.docx", "child", child));

        string output = ExportIncludeText(parent, resolver, DxpFieldEvalExportMode.Evaluate);

        Assert.Contains("BEFORE", output, StringComparison.Ordinal);
        Assert.Contains("CHILD", output, StringComparison.Ordinal);
        Assert.Contains("AFTER", output, StringComparison.Ordinal);
        Assert.True(output.IndexOf("BEFORE", StringComparison.Ordinal) < output.IndexOf("CHILD", StringComparison.Ordinal));
        Assert.True(output.IndexOf("CHILD", StringComparison.Ordinal) < output.IndexOf("AFTER", StringComparison.Ordinal));
        Assert.DoesNotContain("CACHE", output, StringComparison.Ordinal);
    }

    [Fact]
    public void Eval_IncludeText_SingleParagraphMergesInlineAndProducesWellFormedHtml()
    {
        using var parent = CreateInlineIncludeTextDoc("child.docx", "CACHE", "BEFORE-", "-AFTER");
        var resolver = new MemoryIncludeTextResolver(("child.docx", "child", CreateTextDocBytes("CHILD")));
        var visitor = new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig(), Logger);
        visitor.FieldEval.Context.IncludeTextResolver = resolver;

        string html = DxpExport.ExportToString(parent, visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate }, Logger);

        Assert.Contains("BEFORE-CHILD-AFTER", StripTags(html), StringComparison.Ordinal);
        Assert.DoesNotContain("<p class=\"dxp-paragraph\"><p", html, StringComparison.Ordinal);
        _ = System.Xml.Linq.XDocument.Parse(html);
    }

    [Fact]
    public void Eval_IncludeText_MultipleParagraphsMergeFirstAndLastSeams()
    {
        using var parent = CreateInlineIncludeTextDoc("child.docx", "CACHE", "BEFORE-", "-AFTER");
        var resolver = new MemoryIncludeTextResolver(("child.docx", "child", CreateTextDocBytes("FIRST", "LAST")));

        string output = ExportIncludeText(parent, resolver, DxpFieldEvalExportMode.Evaluate);

        Assert.Contains("BEFORE-FIRST", output, StringComparison.Ordinal);
        Assert.Contains("LAST-AFTER", output, StringComparison.Ordinal);
        Assert.True(output.IndexOf("FIRST", StringComparison.Ordinal) < output.IndexOf("LAST", StringComparison.Ordinal));
    }

    [Fact]
    public void Eval_IncludeText_TableBetweenParagraphSeamsProducesWellFormedHtml()
    {
        using var parent = CreateInlineIncludeTextDoc("child.docx", "CACHE", "BEFORE-", "-AFTER");
        var resolver = new MemoryIncludeTextResolver(("child.docx", "child", CreateParagraphTableParagraphDocBytes()));
        var visitor = new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig(), Logger);
        visitor.FieldEval.Context.IncludeTextResolver = resolver;

        string html = DxpExport.ExportToString(parent, visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate }, Logger);
        string text = StripTags(html);

        _ = System.Xml.Linq.XDocument.Parse(html);
        Assert.True(text.IndexOf("BEFORE-FIRST", StringComparison.Ordinal) < text.IndexOf("MIDDLE", StringComparison.Ordinal));
        Assert.True(text.IndexOf("MIDDLE", StringComparison.Ordinal) < text.IndexOf("LAST-AFTER", StringComparison.Ordinal));
    }

    [Fact]
    public void Eval_UnsupportedIncludeTextCacheInsideTableCellProducesWellFormedHtml()
    {
        using var stream = new MemoryStream();
        using (var created = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = created.AddMainDocumentPart();
            main.Document = new Document(new Body(new Table(new TableRow(new TableCell(new Paragraph(
                new Run(new RunProperties(new Bold()), new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new RunProperties(new Bold()), new FieldCode { Text = " INCLUDETEXT \"signature.htm\" \\c HTML " }),
                new Run(new RunProperties(new Bold()), new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new RunProperties(new Bold()), new Text("Error! Not a valid filename.")),
                new Run(new RunProperties(new Bold()), new FieldChar { FieldCharType = FieldCharValues.End })))))));
            main.Document.Save();
        }
        stream.Position = 0;
        using var document = WordprocessingDocument.Open(stream, false);
        var visitor = new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig(), Logger);
        string html = DxpExport.ExportToString(document, visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate }, Logger);

        _ = System.Xml.Linq.XDocument.Parse(html);
        Assert.Contains("Error! Not a valid filename.", html, StringComparison.Ordinal);
        Assert.DoesNotContain("</td>\n</strong>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Eval_IncludeText_UsesEvaluatedNestedPathAndEvaluatesChildFields()
    {
        byte[] child = CreateDocVariableDocBytes("ChildValue", "child-cache");
        using var parent = CreateNestedIncludeTextDoc("Root", "\\Headers\\Example.docx", "CACHE");
        var resolver = new MemoryIncludeTextResolver(("C:\\Templates\\Headers\\Example.docx", "child", child));
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig(), Logger);
        visitor.FieldEval.Context.SetDocVariable("Root", "C:\\Templates");
        visitor.FieldEval.Context.SetDocVariable("ChildValue", "CHILD-VALUE");
        visitor.FieldEval.Context.IncludeTextResolver = resolver;

        string output = DxpExport.ExportToString(parent, visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate }, Logger);

        Assert.Equal("C:\\Templates\\Headers\\Example.docx", Assert.Single(resolver.Requests));
        Assert.Contains("CHILD-VALUE", output, StringComparison.Ordinal);
    }

    [Fact]
    public void Eval_IncludeText_RecursesAndStopsCycles()
    {
        byte[] cycle = CreateIncludeTextDocBytes("cycle.docx", "CYCLE-CACHE", "CHILD-BEFORE", "CHILD-AFTER");
        using var parent = CreateIncludeTextDoc("cycle.docx", "PARENT-CACHE");
        var resolver = new MemoryIncludeTextResolver(("cycle.docx", "cycle", cycle));

        string output = ExportIncludeText(parent, resolver, DxpFieldEvalExportMode.Evaluate);

        Assert.Contains("CHILD-BEFORE", output, StringComparison.Ordinal);
        Assert.Contains("CYCLE-CACHE", output, StringComparison.Ordinal);
        Assert.Contains("CHILD-AFTER", output, StringComparison.Ordinal);
        Assert.Equal(2, resolver.Requests.Count);
    }

    [Fact]
    public void Eval_IncludeText_DepthLimitReplaysNestedCache()
    {
        byte[] second = CreateTextDocBytes("TOO-DEEP");
        byte[] first = CreateIncludeTextDocBytes("second.docx", "DEPTH-CACHE");
        using var parent = CreateIncludeTextDoc("first.docx", "PARENT-CACHE");
        var resolver = new MemoryIncludeTextResolver(
            ("first.docx", "first", first),
            ("second.docx", "second", second));
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig(), Logger);
        visitor.FieldEval.Context.IncludeTextResolver = resolver;
        visitor.FieldEval.Context.MaxIncludeTextDepth = 1;

        string output = DxpExport.ExportToString(parent, visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate }, Logger);

        Assert.Contains("DEPTH-CACHE", output, StringComparison.Ordinal);
        Assert.DoesNotContain("TOO-DEEP", output, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("missing.docx", " INCLUDETEXT \"missing.docx\" ")]
    [InlineData("page.htm", " INCLUDETEXT \"page.htm\" \\c HTML ")]
    [InlineData("child.docx", " INCLUDETEXT \"child.docx\" Bookmark1 ")]
    public void Eval_IncludeText_UnsupportedOrMissingSourceReplaysCache(string path, string instruction)
    {
        using var parent = CreateIncludeTextDoc(path, "CACHED", instruction: instruction);
        var resolver = new MemoryIncludeTextResolver();

        string output = ExportIncludeText(parent, resolver, DxpFieldEvalExportMode.Evaluate);

        Assert.Contains("CACHED", output, StringComparison.Ordinal);
    }

    [Fact]
    public void Cache_IncludeText_DoesNotInvokeResolver()
    {
        using var parent = CreateIncludeTextDoc("child.docx", "CACHED");
        var resolver = new MemoryIncludeTextResolver(("child.docx", "child", CreateTextDocBytes("CHILD")));

        string output = ExportIncludeText(parent, resolver, DxpFieldEvalExportMode.Cache);

        Assert.Contains("CACHED", output, StringComparison.Ordinal);
        Assert.Empty(resolver.Requests);
    }

    [Fact]
    public void Eval_IncludeText_MalformedDocxReplaysCache()
    {
        using var parent = CreateIncludeTextDoc("broken.docx", "CACHED");
        var resolver = new MemoryIncludeTextResolver(("broken.docx", "broken", [1, 2, 3, 4]));

        string output = ExportIncludeText(parent, resolver, DxpFieldEvalExportMode.Evaluate);

        Assert.Contains("CACHED", output, StringComparison.Ordinal);
    }

    [Fact]
    public void Eval_IncludeText_MalformedDocxReplaysCacheInsideOriginalParagraph()
    {
        using var parent = CreateInlineIncludeTextDoc("broken.docx", "CACHED", "BEFORE-", "-AFTER");
        var resolver = new MemoryIncludeTextResolver(("broken.docx", "broken", [1, 2, 3, 4]));
        var visitor = new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig(), Logger);
        visitor.FieldEval.Context.IncludeTextResolver = resolver;

        string html = DxpExport.ExportToString(parent, visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate }, Logger);

        _ = System.Xml.Linq.XDocument.Parse(html);
        Assert.Contains("BEFORE-CACHED-AFTER", StripTags(html), StringComparison.Ordinal);
    }

    [Fact]
    public void Eval_MultipleIncludeTextFieldsProduceWellFormedHtml()
    {
        using var parent = CreateTwoIncludeTextDoc();
        var resolver = new MemoryIncludeTextResolver(
            ("first.docx", "first", CreateTextDocBytes("FIRST")),
            ("second.docx", "second", CreateTextDocBytes("SECOND")));
        var visitor = new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig(), Logger);
        visitor.FieldEval.Context.IncludeTextResolver = resolver;

        string html = DxpExport.ExportToString(parent, visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate }, Logger);

        _ = System.Xml.Linq.XDocument.Parse(html);
        string text = StripTags(html);
        Assert.True(text.IndexOf("BEFORE", StringComparison.Ordinal) < text.IndexOf("FIRST", StringComparison.Ordinal));
        Assert.True(text.IndexOf("FIRST", StringComparison.Ordinal) < text.IndexOf("BETWEEN", StringComparison.Ordinal));
        Assert.True(text.IndexOf("BETWEEN", StringComparison.Ordinal) < text.IndexOf("SECOND", StringComparison.Ordinal));
        Assert.True(text.IndexOf("SECOND", StringComparison.Ordinal) < text.IndexOf("AFTER", StringComparison.Ordinal));
    }

    [Fact]
    public async Task Eval_ImplicitBookmarkFieldBehavesLikeRef()
    {
        var eval = new DocxportNet.Fields.DxpFieldEval(logger: Logger);
        eval.Context.SetBookmarkNodes("DocPathBMK", DocxportNet.Fields.DxpFieldNodeBuffer.FromText("C:\\Templates"));

        var result = await eval.EvalAsync(new DocxportNet.Fields.DxpFieldInstruction(" DocPathBMK "));

        Assert.Equal(DocxportNet.Fields.DxpFieldEvalStatus.Resolved, result.Status);
        Assert.Equal("C:\\Templates", result.Text);
    }

    [Fact]
    public async Task Eval_UnknownBareFieldWithoutBookmarkUsesCache()
    {
        var eval = new DocxportNet.Fields.DxpFieldEval(logger: Logger);

        var result = await eval.EvalAsync(new DocxportNet.Fields.DxpFieldInstruction(" MissingBMK ", "CACHED"));

        Assert.Equal(DocxportNet.Fields.DxpFieldEvalStatus.UsedCache, result.Status);
        Assert.Equal("CACHED", result.Text);
    }

    [Fact]
    public void Eval_ImplicitBookmarkNestedInIncludeTextBuildsPath()
    {
        byte[] child = CreateTextDocBytes("CHILD");
        using var parent = CreateSetAndImplicitBookmarkIncludeDoc();
        var resolver = new MemoryIncludeTextResolver(("C:\\Templates\\Headers\\Example.docx", "child", child));

        string output = ExportIncludeText(parent, resolver, DxpFieldEvalExportMode.Evaluate);

        Assert.Equal("C:\\Templates\\Headers\\Example.docx", Assert.Single(resolver.Requests));
        Assert.Contains("CHILD", output, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(" PAGE ", "7")]
    [InlineData(" SECTION ", "2")]
    [InlineData(" SECTIONPAGES ", "4")]
    [InlineData(" PAGEREF Target \\h ", "3")]
    public void Eval_LayoutDependentSimpleFieldReplaysCachedResult(string instruction, string cached)
    {
        using var document = CreateSimpleFieldDoc(instruction, cached);
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig(), Logger);

        string output = DxpExport.ExportToString(document, visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate }, Logger);

        Assert.Contains(cached, output, StringComparison.Ordinal);
        Assert.DoesNotContain("Error! Invalid field code.", output, StringComparison.Ordinal);
    }

    [Fact]
    public void Eval_LayoutDependentComplexFieldReplaysStructuredCachedResult()
    {
        using var document = CreateComplexCachedFieldDoc(" PAGE ", "12", bold: true);
        var visitor = new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig(), Logger);

        string html = DxpExport.ExportToString(document, visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate }, Logger);

        Assert.Contains("12", html, StringComparison.Ordinal);
        Assert.Contains("dxp-bold", html, StringComparison.Ordinal);
        Assert.DoesNotContain("Error! Invalid field code.", html, StringComparison.Ordinal);
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

    private static WordprocessingDocument CreateComplexCachedFieldDoc(string instruction, string cachedText, bool bold)
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var mainPart = document.AddMainDocumentPart();
            RunProperties? properties = bold ? new RunProperties(new Bold()) : null;
            Run MakeRun(OpenXmlElement child)
            {
                var run = new Run();
                if (properties != null)
                    run.RunProperties = (RunProperties)properties.CloneNode(true);
                run.Append(child);
                return run;
            }
            mainPart.Document = new Document(new Body(new Paragraph(
                MakeRun(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                MakeRun(new FieldCode { Text = instruction }),
                MakeRun(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                MakeRun(new Text(cachedText)),
                MakeRun(new FieldChar { FieldCharType = FieldCharValues.End }))));
            mainPart.Document.Save();
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

    private string ExportIncludeText(
        WordprocessingDocument document,
        MemoryIncludeTextResolver resolver,
        DxpFieldEvalExportMode mode)
    {
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig(), Logger);
        visitor.FieldEval.Context.IncludeTextResolver = resolver;
        return DxpExport.ExportToString(document, visitor, new DxpExportOptions { FieldEvalMode = mode }, Logger);
    }

    private static WordprocessingDocument CreateIncludeTextDoc(
        string path,
        string cached,
        string? before = null,
        string? after = null,
        string? instruction = null)
    {
        var stream = new MemoryStream(CreateIncludeTextDocBytes(path, cached, before, after, instruction));
        return WordprocessingDocument.Open(stream, false);
    }

    private static WordprocessingDocument CreateInlineIncludeTextDoc(
        string path, string cached, string before, string after)
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text(before)),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode { Text = $" INCLUDETEXT \"{path}\" " }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text(cached)),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
                new Run(new Text(after)))));
            main.Document.Save();
        }
        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }

    private static WordprocessingDocument CreateTwoIncludeTextDoc()
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new Text("BEFORE")),
                CreateSimpleIncludeTextField("first.docx", "FIRST-CACHE"),
                new Run(new Text("BETWEEN")),
                CreateSimpleIncludeTextField("second.docx", "SECOND-CACHE"),
                new Run(new Text("AFTER")))));
            main.Document.Save();
        }
        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }

    private static SimpleField CreateSimpleIncludeTextField(string path, string cached)
        => new(new Run(new Text(cached))) { Instruction = $" INCLUDETEXT \"{path}\" " };

    private static byte[] CreateIncludeTextDocBytes(
        string path,
        string cached,
        string? before = null,
        string? after = null,
        string? instruction = null)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body());
            if (before != null)
                main.Document.Body!.Append(new Paragraph(new Run(new Text(before))));
            main.Document.Body!.Append(new Paragraph(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode { Text = instruction ?? $" INCLUDETEXT \"{path}\" " }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text(cached)),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End })));
            if (after != null)
                main.Document.Body!.Append(new Paragraph(new Run(new Text(after))));
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static WordprocessingDocument CreateNestedIncludeTextDoc(string variable, string suffix, string cached)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode { Text = " INCLUDETEXT \"" }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode { Text = $" DOCVARIABLE {variable} " }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
                new Run(new FieldCode { Text = suffix + "\" " }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text(cached)),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }))));
            main.Document.Save();
        }
        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }

    private static byte[] CreateTextDocBytes(params string[] paragraphs)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(paragraphs.Select(text => new Paragraph(new Run(new Text(text))))));
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static string StripTags(string html)
        => System.Net.WebUtility.HtmlDecode(System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", string.Empty));

    private static byte[] CreateDocVariableDocBytes(string name, string cached)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode { Text = $" DOCVARIABLE {name} " }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text(cached)),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }))));
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] CreateParagraphTableParagraphDocBytes()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("FIRST"))),
                new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text("MIDDLE")))))),
                new Paragraph(new Run(new Text("LAST")))));
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static WordprocessingDocument CreateSetAndImplicitBookmarkIncludeDoc()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                    new Run(new FieldCode { Text = " SET DocPathBMK \"C:\\Templates\" " }),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.End })),
                new Paragraph(
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                    new Run(new FieldCode { Text = " INCLUDETEXT \"" }),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                    new Run(new FieldCode { Text = " DocPathBMK " }),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
                    new Run(new FieldCode { Text = "\\Headers\\Example.docx\" " }),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                    new Run(new Text("CACHE")),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.End }))));
            main.Document.Save();
        }
        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }

    private sealed class MemoryIncludeTextResolver : IDxpIncludeTextResolver
    {
        private readonly Dictionary<string, DxpIncludeTextSource> _sources;

        public MemoryIncludeTextResolver(params (string Path, string Identity, byte[] Content)[] sources)
        {
            _sources = sources.ToDictionary(
                source => source.Path,
                source => new DxpIncludeTextSource(source.Identity, source.Content),
                StringComparer.OrdinalIgnoreCase);
        }

        public List<string> Requests { get; } = new();

        public Task<DxpIncludeTextSource?> ResolveAsync(
            DxpIncludeTextRequest request,
            DocxportNet.Fields.DxpFieldEvalContext context,
            CancellationToken cancellationToken = default)
        {
            Requests.Add(request.Path);
            _sources.TryGetValue(request.Path, out var source);
            return Task.FromResult(source);
        }
    }
}
