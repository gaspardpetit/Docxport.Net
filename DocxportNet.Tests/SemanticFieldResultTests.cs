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
        var validationErrors = new OpenXmlValidator().Validate(output).ToArray();
        Assert.True(validationErrors.Length == 0,
            string.Join(Environment.NewLine, validationErrors.Select(error =>
                $"{error.Description} Path={error.Path?.XPath}")));
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

    [Fact]
    public void DeepExpressionTreePreservesTypedSetAndStructuredChildOrder()
    {
        byte[] source = CreateDeepCompositionDocument(outerCondition: true);
        var eval = CreateDeepEval();

        using var output = Open(DxpDocxExport.Export(source, SemanticOptions(), Logger, eval));
        Body body = output.MainDocumentPart!.Document.Body!;

        Assert.True(eval.Context.TryGetBookmarkValue("Result", out DxpFieldValue value));
        Assert.Equal(DxpFieldValueKind.Number, value.Kind);
        Assert.Equal(42d, value.NumberValue);
        Assert.Contains("Value: 42", body.InnerText, StringComparison.Ordinal);
        Assert.Single(body.Elements<Table>());
        Assert.True(body.InnerText.IndexOf("Value: 42", StringComparison.Ordinal) <
            body.InnerText.IndexOf("Alice", StringComparison.Ordinal));
        Assert.Empty(new OpenXmlValidator().Validate(output));
    }

    [Fact]
    public void DeepExpressionTreeDoesNotEvaluateUnselectedChildren()
    {
        byte[] source = CreateDeepCompositionDocument(outerCondition: false);
        var provider = new CountingDatabaseProvider();
        var eval = CreateDeepEval(provider);

        using var output = Open(DxpDocxExport.Export(source, SemanticOptions(), Logger, eval));

        Assert.False(eval.Context.TryGetBookmarkValue("Result", out _));
        Assert.Equal(0, provider.CallCount);
        Assert.DoesNotContain("Value:", output.MainDocumentPart!.Document.Body!.InnerText, StringComparison.Ordinal);
        Assert.Empty(output.MainDocumentPart.Document.Body.Elements<Table>());
    }

    [Fact]
    public void IncludeTextPathIsComposedFromExpressionChildrenWithoutFieldCodeReconstruction()
    {
        byte[] included = CreateDocument(body =>
            body.Append(new Paragraph(new Run(new Text("INCLUDED")))));
        var resolver = new RecordingIncludeResolver(included);
        var eval = new DxpFieldEval(new DxpFieldEvalDelegates
        {
            ResolveDocVariableAsync = (name, _) => Task.FromResult<DxpFieldValue?>(name switch
            {
                "VersionSource" => new DxpFieldValue("2026_"),
                "Language" => new DxpFieldValue("E"),
                _ => null
            })
        }, logger: Logger);
        eval.Context.IncludeTextResolver = resolver;

        using var output = Open(DxpDocxExport.Export(
            CreateComposedIncludeDocument(), SemanticOptions(), Logger, eval));

        Assert.Equal("Synthetic_2026_E.docx", resolver.RequestedPath);
        Assert.Contains("Before INCLUDED After", output.MainDocumentPart!.Document.Body!.InnerText,
            StringComparison.Ordinal);
        Assert.Empty(new OpenXmlValidator().Validate(output));
    }

    [Fact]
    public void DatabaseQueryIsComposedFromExpressionChildrenWithoutFieldCodeReconstruction()
    {
        var provider = new RecordingDatabaseProvider();
        var eval = new DxpFieldEval(new DxpFieldEvalDelegates
        {
            ResolveDocVariableAsync = (name, _) => Task.FromResult<DxpFieldValue?>(
                name == "SyntheticId" ? new DxpFieldValue(7d) : null)
        }, logger: Logger);
        eval.Context.DatabaseProvider = provider;

        using var output = Open(DxpDocxExport.Export(
            CreateComposedDatabaseDocument(), SemanticOptions(), Logger, eval));

        Assert.Equal("SELECT Name FROM SyntheticPeople WHERE Id = 7", provider.QueryText);
        Assert.Single(output.MainDocumentPart!.Document.Body!.Descendants<Table>());
        Assert.Empty(new OpenXmlValidator().Validate(output));
    }

    [Fact]
    public void InlineTextAroundNestedBlockIsNormalizedAcrossExporters()
    {
        byte[] source = CreateDatabaseWithInlineSiblingsDocument();

        using (var output = Open(DxpDocxExport.Export(source, SemanticOptions(), Logger, CreateEval())))
        {
            OpenXmlElement[] content = output.MainDocumentPart!.Document.Body!.ChildElements
                .Where(static element => element is Table || element is Paragraph paragraph && paragraph.InnerText.Length > 0)
                .ToArray();
            Assert.Collection(content,
                element => Assert.Equal("BEFORE", Assert.IsType<Paragraph>(element).InnerText),
                element => Assert.IsType<Table>(element),
                element => Assert.Equal("AFTER", Assert.IsType<Paragraph>(element).InnerText));
            Assert.Empty(new OpenXmlValidator().Validate(output));
        }

        string html = DxpExport.ExportToString(source,
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig(), Logger, CreateEval()),
            SemanticOptions(), Logger);
        Assert.True(html.IndexOf("BEFORE", StringComparison.Ordinal) < html.IndexOf("<table", StringComparison.OrdinalIgnoreCase));
        Assert.True(html.IndexOf("</table>", StringComparison.OrdinalIgnoreCase) < html.IndexOf("AFTER", StringComparison.Ordinal));

        string markdown = DxpExport.ExportToString(source,
            new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreateRichConfig(), Logger, CreateEval()),
            SemanticOptions(), Logger);
        Assert.True(markdown.IndexOf("BEFORE", StringComparison.Ordinal) < markdown.IndexOf("Alice", StringComparison.Ordinal));
        Assert.True(markdown.IndexOf("Alice", StringComparison.Ordinal) < markdown.IndexOf("AFTER", StringComparison.Ordinal));
    }

    [Fact]
    public void ImplicitAndExplicitBookmarkReferencesHaveEquivalentTypedResults()
    {
        byte[] source = CreateDocument(body => body.Append(new Paragraph(
            Begin(), Code(" SET SyntheticBookmark 42 "), End(),
            Begin(), Code(" SyntheticBookmark "), End(),
            new Run(new Text("|")),
            Begin(), Code(" REF SyntheticBookmark "), End())));

        using var output = Open(DxpDocxExport.Export(
            source, SemanticOptions(), Logger, new DxpFieldEval(logger: Logger)));

        Assert.Equal("42|42", output.MainDocumentPart!.Document.Body!.InnerText);
        Assert.Empty(new OpenXmlValidator().Validate(output));
    }

    [Fact]
    public void SelectedLiteralPreservesFragmentFormattingAcrossExporters()
    {
        byte[] source = CreateFormattedLiteralDocument();

        using (var output = Open(DxpDocxExport.Export(source, SemanticOptions(), Logger, new DxpFieldEval(logger: Logger))))
        {
            Run bold = Assert.Single(output.MainDocumentPart!.Document.Body!.Descendants<Run>(),
                static run => run.InnerText.Trim() == "bold");
            Run italic = Assert.Single(output.MainDocumentPart.Document.Body.Descendants<Run>(),
                static run => run.InnerText.Trim() == "italic");
            Assert.NotNull(bold.RunProperties?.Bold);
            Assert.NotNull(italic.RunProperties?.Italic);
            Assert.Empty(new OpenXmlValidator().Validate(output));
        }

        string html = DxpExport.ExportToString(source,
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig(), Logger, new DxpFieldEval(logger: Logger)),
            SemanticOptions(), Logger);
        Assert.Contains("<strong class=\"dxp-bold\">", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<em class=\"dxp-italic\">", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("bold", html, StringComparison.Ordinal);
        Assert.Contains("italic", html, StringComparison.Ordinal);

        string markdown = DxpExport.ExportToString(source,
            new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreateRichConfig(), Logger, new DxpFieldEval(logger: Logger)),
            SemanticOptions(), Logger);
        Assert.Contains("<b>", markdown, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<i>", markdown, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("bold", markdown, StringComparison.Ordinal);
        Assert.Contains("italic", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void CrossParagraphLiteralPreservesParagraphFormatting()
    {
        byte[] source = CreateDocument(body => body.Append(
            new Paragraph(Begin(), Code(" IF 1 = 1 \"FIRST")),
            new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId { Val = "SyntheticStyle" },
                    new Justification { Val = JustificationValues.Center }),
                Code("SECOND\" \"\" "), End())));

        using var output = Open(DxpDocxExport.Export(
            source, SemanticOptions(), Logger, new DxpFieldEval(logger: Logger)));

        Paragraph second = Assert.Single(output.MainDocumentPart!.Document.Body!.Elements<Paragraph>(),
            static paragraph => paragraph.InnerText == "SECOND");
        Assert.Equal("SyntheticStyle", second.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
        Assert.Equal(JustificationValues.Center, second.ParagraphProperties?.Justification?.Val?.Value);
        Assert.Empty(new OpenXmlValidator().Validate(output));
    }

    [Theory]
    [InlineData(DxpMissingDocVariableBehavior.Error, "Error! No document variable supplied.")]
    [InlineData(DxpMissingDocVariableBehavior.UseCache, "CACHED")]
    [InlineData(DxpMissingDocVariableBehavior.Skip, "")]
    public void MissingDocVariablePolicyIsHonoredBySemanticEvaluation(
        DxpMissingDocVariableBehavior behavior,
        string expected)
    {
        var eval = new DxpFieldEval(options: new DxpFieldEvalOptions
        {
            MissingDocVariableBehavior = behavior
        }, logger: Logger);

        using var output = Open(DxpDocxExport.Export(
            CreateCachedDocVariableDocument(nestedInIf: false), SemanticOptions(), Logger, eval));

        Assert.Equal(expected, output.MainDocumentPart!.Document.Body!.InnerText);
        Assert.Empty(new OpenXmlValidator().Validate(output));
    }

    [Fact]
    public void NestedMissingDocVariableCanUseItsOwnCachedResult()
    {
        var eval = new DxpFieldEval(options: new DxpFieldEvalOptions
        {
            MissingDocVariableBehavior = DxpMissingDocVariableBehavior.UseCache
        }, logger: Logger);

        using var output = Open(DxpDocxExport.Export(
            CreateCachedDocVariableDocument(nestedInIf: true), SemanticOptions(), Logger, eval));

        Assert.Equal("CACHED", output.MainDocumentPart!.Document.Body!.InnerText);
        Assert.Empty(new OpenXmlValidator().Validate(output));
    }

    [Fact]
    public void SemanticFieldsResolveAcrossBodyHeaderFooterTableAndTextBoxStories()
    {
        using var output = Open(DxpDocxExport.Export(
            CreateMultiStoryDocument(), SemanticOptions(), Logger, new DxpFieldEval(logger: Logger)));

        Assert.Contains("BODY", output.MainDocumentPart!.Document.Body!.InnerText, StringComparison.Ordinal);
        Assert.Contains("CELL", output.MainDocumentPart.Document.Body.InnerText, StringComparison.Ordinal);
        Assert.Contains("TEXTBOX", output.MainDocumentPart.Document.Body.InnerText, StringComparison.Ordinal);
        Assert.Equal("HEADER", output.MainDocumentPart.HeaderParts.Single().Header!.InnerText);
        Assert.Equal("FOOTER", output.MainDocumentPart.FooterParts.Single().Footer!.InnerText);
        Assert.DoesNotContain(output.MainDocumentPart.Document.Descendants<FieldCode>(),
            static code => code.Text.Contains("IF", StringComparison.OrdinalIgnoreCase));
        var validationErrors = new OpenXmlValidator().Validate(output).ToArray();
        Assert.True(validationErrors.Length == 0,
            string.Join(Environment.NewLine, validationErrors.Select(error =>
                $"{error.Description} Path={error.Path?.XPath}")));
    }

    [Fact]
    public void SemanticFieldsResolveInsideComments()
    {
        using var output = Open(DxpDocxExport.Export(
            CreateCommentDocument(), SemanticOptions(), Logger, new DxpFieldEval(logger: Logger)));

        var comments = output.MainDocumentPart!.WordprocessingCommentsPart!.Comments!;
        Assert.Equal("COMMENT", comments.InnerText);
        Assert.Empty(comments.Descendants<FieldCode>());
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

    private DxpFieldEval CreateDeepEval(IDatabaseFieldProvider? provider = null)
    {
        var eval = new DxpFieldEval(new DxpFieldEvalDelegates
        {
            ResolveDocVariableAsync = (name, _) => Task.FromResult<DxpFieldValue?>(name switch
            {
                "Choice" => new DxpFieldValue("YES"),
                "Payload" => new DxpFieldValue(42d),
                _ => null
            })
        }, logger: Logger);
        eval.Context.DatabaseProvider = provider ?? new SyntheticDatabaseProvider();
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

    private static byte[] CreateDeepCompositionDocument(bool outerCondition)
        => CreateDocument(body => body.Append(
            new Paragraph(
                Begin(),
                Code($" IF 1 = {(outerCondition ? 1 : 0)} \""),
                    Begin(),
                    Code(" IF "),
                        Begin(), Code(" DOCVARIABLE Choice "), End(),
                    Code(" = \"YES\" \""),
                        Begin(),
                        Code(" SET Result \""),
                            Begin(), Code(" DOCVARIABLE Payload "), End(),
                        Code("\" "),
                        End(),
                    Code("\" \""),
                        Begin(), Code(" SET Result \"wrong\" "), End(),
                    Code("\" "),
                    End(),
                Code("Value: "),
                    Begin(), Code(" REF Result "), End(),
                Code(" "),
                    Begin(),
                    Code(" DATABASE \\h \\s \"SELECT Name, Address FROM SyntheticPeople\" "),
                    End(),
                Code("\" \"\" "),
                End()),
            new Paragraph(new Run(new Text("AFTER")))));

    private static byte[] CreateComposedIncludeDocument()
        => CreateDocument(body => body.Append(
            new Paragraph(
                Begin(),
                Code(" IF 1 = 1 \""),
                    Begin(),
                    Code(" SET Version \""),
                        Begin(), Code(" DOCVARIABLE VersionSource "), End(),
                    Code("\" "),
                    End(),
                Code("Before "),
                    Begin(),
                    Code(" INCLUDETEXT \"Synthetic_"),
                        Begin(), Code(" REF Version "), End(),
                        Begin(), Code(" DOCVARIABLE Language "), End(),
                    Code(".docx\" "),
                    End(),
                Code(" After\" \"\" "),
                End())));

    private static byte[] CreateComposedDatabaseDocument()
        => CreateDocument(body => body.Append(
            new Paragraph(
                Begin(),
                Code(" IF 1 = 1 \""),
                    Begin(),
                    Code(" DATABASE \\h \\s \"SELECT Name FROM SyntheticPeople WHERE Id = "),
                        Begin(), Code(" DOCVARIABLE SyntheticId "), End(),
                    Code("\" "),
                    End(),
                Code("\" \"\" "),
                End())));

    private static byte[] CreateDatabaseWithInlineSiblingsDocument()
        => CreateDocument(body => body.Append(
            new Paragraph(
                Begin(),
                Code(" IF 1 = 1 \"BEFORE"),
                    Begin(), Code(" DATABASE \\h \\s \"SELECT Name, Address FROM SyntheticPeople\" "), End(),
                Code("AFTER\" \"\" "),
                End())));

    private static byte[] CreateFormattedLiteralDocument()
        => CreateDocument(body => body.Append(new Paragraph(
            Begin(),
            Code(" IF 1 = 1 \"plain "),
            new Run(new RunProperties(new Bold()), new FieldCode { Text = "bold" }),
            new Run(new RunProperties(new Italic()), new FieldCode { Text = " italic" }),
            Code("\" \"\" "),
            End())));

    private static byte[] CreateCachedDocVariableDocument(bool nestedInIf)
        => CreateDocument(body => body.Append(nestedInIf
            ? new Paragraph(
                Begin(), Code(" IF 1 = 1 \""),
                Begin(), Code(" DOCVARIABLE Missing "), Separate(),
                new Run(new Text("CACHED")), End(),
                Code("\" \"\" "), End())
            : new Paragraph(
                Begin(), Code(" DOCVARIABLE Missing "), Separate(),
                new Run(new Text("CACHED")), End())));

    private static byte[] CreateMultiStoryDocument()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            MainDocumentPart main = document.AddMainDocumentPart();
            HeaderPart header = main.AddNewPart<HeaderPart>();
            FooterPart footer = main.AddNewPart<FooterPart>();
            header.Header = new Header(FieldParagraph("HEADER"));
            footer.Footer = new Footer(FieldParagraph("FOOTER"));

            var textBox = new TextBoxContent(FieldParagraph("TEXTBOX"));
            var picture = new Picture(
                new DocumentFormat.OpenXml.Vml.Shape(
                    new DocumentFormat.OpenXml.Vml.TextBox(textBox))
                {
                    Id = "SyntheticTextBox",
                    Style = "position:absolute;width:100pt;height:30pt"
                });

            main.Document = new Document(new Body(
                FieldParagraph("BODY"),
                new Table(
                    new TableProperties(),
                    new TableGrid(new GridColumn()),
                    new TableRow(new TableCell(FieldParagraph("CELL")))),
                new Paragraph(new Run(picture)),
                new SectionProperties(
                    new HeaderReference
                    {
                        Id = main.GetIdOfPart(header),
                        Type = HeaderFooterValues.Default
                    },
                    new FooterReference
                    {
                        Id = main.GetIdOfPart(footer),
                        Type = HeaderFooterValues.Default
                    })));
            main.Document.Save();
            header.Header.Save();
            footer.Footer.Save();
        }
        return stream.ToArray();
    }

    private static byte[] CreateCommentDocument()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            MainDocumentPart main = document.AddMainDocumentPart();
            var commentsPart = main.AddNewPart<WordprocessingCommentsPart>();
            commentsPart.Comments = new Comments(
                new Comment(FieldParagraph("COMMENT"))
                {
                    Id = "0",
                    Author = "Synthetic"
                });
            main.Document = new Document(new Body(
                new Paragraph(
                    new CommentRangeStart { Id = "0" },
                    new Run(new Text("ANCHOR")),
                    new CommentRangeEnd { Id = "0" },
                    new Run(new CommentReference { Id = "0" })),
                new SectionProperties()));
            main.Document.Save();
            commentsPart.Comments.Save();
        }
        return stream.ToArray();
    }

    private static Paragraph FieldParagraph(string value)
        => new(Begin(), Code($" IF 1 = 1 \"{value}\" \"\" "), End());

    private static Run Begin() => new(new FieldChar { FieldCharType = FieldCharValues.Begin });
    private static Run Separate() => new(new FieldChar { FieldCharType = FieldCharValues.Separate });
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

    private sealed class CountingDatabaseProvider : IDatabaseFieldProvider
    {
        public int CallCount { get; private set; }

        public Task<DxpDatabaseResult?> ExecuteAsync(
            DxpDatabaseRequest request,
            CancellationToken cancellationToken)
        {
            CallCount++;
            return new SyntheticDatabaseProvider().ExecuteAsync(request, cancellationToken);
        }
    }

    private sealed class RecordingDatabaseProvider : IDatabaseFieldProvider
    {
        public string? QueryText { get; private set; }

        public Task<DxpDatabaseResult?> ExecuteAsync(
            DxpDatabaseRequest request,
            CancellationToken cancellationToken)
        {
            QueryText = request.QueryText;
            return new SyntheticDatabaseProvider().ExecuteAsync(request, cancellationToken);
        }
    }

    private sealed class RecordingIncludeResolver(byte[] content) : IDxpIncludeTextResolver
    {
        public string? RequestedPath { get; private set; }

        public Task<DxpIncludeTextSource?> ResolveAsync(
            DxpIncludeTextRequest request,
            DxpFieldEvalContext context,
            CancellationToken cancellationToken = default)
        {
            _ = context;
            _ = cancellationToken;
            RequestedPath = request.Path;
            return Task.FromResult<DxpIncludeTextSource?>(new DxpIncludeTextSource(request.Path, content)
            {
                Format = DxpIncludeTextSourceFormat.Docx
            });
        }
    }
}
