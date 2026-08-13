using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Fields;
using DocxportNet.Fields.Resolution;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class HtmlIncludeTextTests
{
    private static readonly byte[] s_png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+Xw4AAAAASUVORK5CYII=");

    [Fact]
    public void HtmlSource_IsSplicedThroughExistingBlockPipeline()
    {
        using var parent = CreateParent("fragment.html", "CACHE", "BEFORE-", "-AFTER");
        var resolver = new StaticResolver(new DxpIncludeTextSource("html", Encoding.UTF8.GetBytes(
            "<p><strong>FIRST</strong></p><table><tr><td>MIDDLE</td></tr></table><p>LAST</p>"))
        {
            Format = DxpIncludeTextSourceFormat.Html
        });
        var visitor = new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig());
        visitor.FieldEval.Context.IncludeTextResolver = resolver;

        string output = DxpExport.ExportToString(parent, visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate });

        _ = System.Xml.Linq.XDocument.Parse(output);
        Assert.Contains("BEFORE-", output, StringComparison.Ordinal);
        Assert.Contains("FIRST", output, StringComparison.Ordinal);
        Assert.Contains("MIDDLE", output, StringComparison.Ordinal);
        Assert.Contains("LAST-AFTER", StripTags(output), StringComparison.Ordinal);
        Assert.DoesNotContain("CACHE", output, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("alias", " INCLUDETEXT \"alias\" ", DxpIncludeTextSourceFormat.Html)]
    [InlineData("alias.bin", " INCLUDETEXT \"alias.bin\" \\c HTML ", DxpIncludeTextSourceFormat.Docx)]
    [InlineData("fragment.html", " INCLUDETEXT \"fragment.html\" ", DxpIncludeTextSourceFormat.Auto)]
    public void HtmlDetection_UsesSwitchThenResolverHintThenExtension(
        string path, string instruction, DxpIncludeTextSourceFormat format)
    {
        using var parent = CreateParent(path, "CACHE", instruction: instruction);
        var resolver = new StaticResolver(new DxpIncludeTextSource("html", Encoding.UTF8.GetBytes("<p>HTML-CONTENT</p>"))
        {
            Format = format
        });
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig());
        visitor.FieldEval.Context.IncludeTextResolver = resolver;

        string output = DxpExport.ExportToString(parent, visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate });

        Assert.Contains("HTML-CONTENT", output, StringComparison.Ordinal);
        Assert.DoesNotContain("CACHE", output, StringComparison.Ordinal);
    }

    [Fact]
    public void ExplicitDocxHint_OverridesHtmlExtension()
    {
        using var parent = CreateParent("aliased.html", "CACHE");
        var resolver = new StaticResolver(new DxpIncludeTextSource("docx", CreateDocx("DOCX-CONTENT"))
        {
            Format = DxpIncludeTextSourceFormat.Docx
        });
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig());
        visitor.FieldEval.Context.IncludeTextResolver = resolver;

        string output = DxpExport.ExportToString(parent, visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate });

        Assert.Contains("DOCX-CONTENT", output, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlInclude_RecursesFromIncludedDocx()
    {
        using var parent = CreateParent("wrapper.docx", "PARENT-CACHE");
        var resolver = new PathResolver(new Dictionary<string, DxpIncludeTextSource>
        {
            ["wrapper.docx"] = new("wrapper", CreateIncludeDocx("fragment.html", "CHILD-CACHE"))
            {
                Format = DxpIncludeTextSourceFormat.Docx
            },
            ["fragment.html"] = new("html", Encoding.UTF8.GetBytes("<p>NESTED-HTML</p>"))
            {
                Format = DxpIncludeTextSourceFormat.Html
            }
        });
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig());
        visitor.FieldEval.Context.IncludeTextResolver = resolver;

        string output = DxpExport.ExportToString(parent, visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate });

        Assert.Contains("NESTED-HTML", output, StringComparison.Ordinal);
        Assert.DoesNotContain("CHILD-CACHE", output, StringComparison.Ordinal);
    }

    [Fact]
    public void NestedFieldsInIncludeCacheDoNotBecomeBookmarkArguments()
    {
        using var parent = CreateParentWithNestedCachedField("fragment.html");
        var resolver = new PathResolver(new Dictionary<string, DxpIncludeTextSource>
        {
            ["fragment.html"] = new("html", Encoding.UTF8.GetBytes("<p>RESOLVED</p>"))
            {
                Format = DxpIncludeTextSourceFormat.Html
            }
        });
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig());
        visitor.FieldEval.Context.IncludeTextResolver = resolver;

        string output = DxpExport.ExportToString(parent, visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate });

        Assert.Contains("RESOLVED", output, StringComparison.Ordinal);
        Assert.DoesNotContain("CACHE", output, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlConversion_IsCachedByContentWithinEvaluationSession()
    {
        byte[] html = Encoding.UTF8.GetBytes("<p>SAME</p>");
        using var parent = CreateParentWithTwoIncludes("one.html", "two.html");
        var resolver = new PathResolver(new Dictionary<string, DxpIncludeTextSource>
        {
            ["one.html"] = new("one", html) { Format = DxpIncludeTextSourceFormat.Html },
            ["two.html"] = new("two", html) { Format = DxpIncludeTextSourceFormat.Html }
        });
        var converter = new CountingConverter(CreateDocx("CONVERTED"));
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig());
        visitor.FieldEval.Context.IncludeTextResolver = resolver;
        visitor.FieldEval.Context.HtmlToDocxConverter = converter;

        string output = DxpExport.ExportToString(parent, visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate });

        Assert.Equal(1, converter.Calls);
        Assert.Equal(2, Count(output, "CONVERTED"));
    }

    [Theory]
    [InlineData("https://assets.example.test/logo.png")]
    [InlineData("http://assets.example.test/logo.png")]
    [InlineData("images/logo.png")]
    [InlineData("file:///C:/Signatures/logo.png")]
    public void ExternalImage_IsNotFetchedAndIsPreserved(string source)
    {
        using var parent = CreateParent("image.html", "CACHE");
        var resolver = new StaticResolver(new DxpIncludeTextSource("html",
            Encoding.UTF8.GetBytes($"<p>IMAGE<img src=\"{source}\" alt=\"logo\"></p>"))
        {
            Format = DxpIncludeTextSourceFormat.Html
        });
        var visitor = new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig());
        visitor.FieldEval.Context.IncludeTextResolver = resolver;

        string output = DxpExport.ExportToString(parent, visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate });

        Assert.Contains(System.Net.WebUtility.HtmlEncode(source), output, StringComparison.Ordinal);
        Assert.Contains("alt=\"logo\"", output, StringComparison.Ordinal);
    }

    [Fact]
    public void ExternalImage_IsPreservedByMarkdownAndPlaceholderByPlainText()
    {
        const string source = "https://assets.example.test/logo.png";
        var resolver = new StaticResolver(new DxpIncludeTextSource("html",
            Encoding.UTF8.GetBytes($"<img src=\"{source}\" alt=\"logo\">"))
        {
            Format = DxpIncludeTextSourceFormat.Html
        });
        using var markdownParent = CreateParent("image.html", "CACHE");
        var markdown = new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreateRichConfig());
        markdown.FieldEval.Context.IncludeTextResolver = resolver;
        string markdownOutput = DxpExport.ExportToString(markdownParent, markdown,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate });

        using var textParent = CreateParent("image.html", "CACHE");
        var text = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig());
        text.FieldEval.Context.IncludeTextResolver = resolver;
        string textOutput = DxpExport.ExportToString(textParent, text,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate });

        Assert.Contains(source, markdownOutput, StringComparison.Ordinal);
        Assert.Contains("[IMAGE]: logo", textOutput, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlHyperlinks_AfterMatchingStyledTextRemainWellNested()
    {
        using var parent = CreateParent("links.html", "CACHE");
        const string html = """
            <p style="font-family:Arial">Text <a href="https://one.example">One</a> —
            <a href="https://two.example">Two</a></p>
            """;
        var resolver = new StaticResolver(new DxpIncludeTextSource("html", Encoding.UTF8.GetBytes(html))
        {
            Format = DxpIncludeTextSourceFormat.Html
        });
        var visitor = new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig());
        visitor.FieldEval.Context.IncludeTextResolver = resolver;

        string output = DxpExport.ExportToString(parent, visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate });

        _ = System.Xml.Linq.XDocument.Parse(output);
        Assert.Contains("https://one.example", output, StringComparison.Ordinal);
        Assert.Contains("https://two.example", output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ValidDataImage_IsEmbedded()
    {
        string data = Convert.ToBase64String(s_png);
        var converter = new DxpHtmlToDocxConverter();

        byte[] docx = await converter.ConvertAsync(Encoding.UTF8.GetBytes(
            $"<img src=\"data:image/png;base64,{data}\" alt=\"pixel\">"));

        using var document = WordprocessingDocument.Open(new MemoryStream(docx), false);
        Assert.NotEmpty(document.MainDocumentPart!.ImageParts);
    }

    [Fact]
    public void InvalidDataImage_ReplaysCache()
    {
        using var parent = CreateParent("invalid.html", "CACHED");
        var resolver = new StaticResolver(new DxpIncludeTextSource("html",
            Encoding.UTF8.GetBytes("<img src=\"data:image/png;base64,bm90LWEtcG5n\">"))
        {
            Format = DxpIncludeTextSourceFormat.Html
        });
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig());
        visitor.FieldEval.Context.IncludeTextResolver = resolver;

        string output = DxpExport.ExportToString(parent, visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate });

        Assert.Contains("CACHED", output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Decoder_HandlesBomMetaAndLegacyFallbacks()
    {
        Assert.Contains("é", await ConvertToText(Encoding.UTF8.GetPreamble()
            .Concat(Encoding.UTF8.GetBytes("<p>é</p>")).ToArray()), StringComparison.Ordinal);
        Assert.Contains("é", await ConvertToText(Encoding.Unicode.GetPreamble()
            .Concat(Encoding.Unicode.GetBytes("<p>é</p>")).ToArray()), StringComparison.Ordinal);
        Assert.Contains("é", await ConvertToText(
            Encoding.GetEncoding(1252).GetBytes("<meta charset=windows-1252><p>é</p>")), StringComparison.Ordinal);
        Assert.Contains("é", await ConvertToText(Encoding.GetEncoding(1252).GetBytes("<p>é</p>")), StringComparison.Ordinal);
        Assert.Contains("wide", await ConvertToText(Encoding.Unicode.GetBytes("<p>wide</p>")), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Converter_NormalizesQuotedSingleWordFontFamily()
    {
        byte[] docx = await new DxpHtmlToDocxConverter().ConvertAsync(
            Encoding.UTF8.GetBytes("<p style=\"font-family:'Arial'\">FONT</p>"));

        using var document = WordprocessingDocument.Open(new MemoryStream(docx), false);
        Assert.Contains(document.MainDocumentPart!.Document!.Descendants<RunFonts>(),
            fonts => fonts.Ascii?.Value == "Arial");
    }

    private static WordprocessingDocument CreateParent(
        string path, string cached, string? before = null, string? after = null, string? instruction = null)
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var content = new List<OpenXmlElement>();
            if (before != null)
                content.Add(new Run(new Text(before)));
            content.AddRange([
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode { Text = instruction ?? $" INCLUDETEXT \"{path}\" " }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text(cached)),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End })]);
            if (after != null)
                content.Add(new Run(new Text(after)));
            main.Document = new Document(new Body(new Paragraph(content)));
            main.Document.Save();
        }
        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }

    private static WordprocessingDocument CreateParentWithTwoIncludes(string first, string second)
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            static IEnumerable<OpenXmlElement> Field(string path) => [
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode { Text = $" INCLUDETEXT \"{path}\" " }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("CACHE")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End })];
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(Field(first)), new Paragraph(Field(second))));
            main.Document.Save();
        }
        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }

    private static WordprocessingDocument CreateParentWithNestedCachedField(string path)
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode { Text = $" INCLUDETEXT \"{path}\" " }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("CACHE-")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode { Text = " DATE \\@ \"yyyy\" " }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("2000")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }))));
            main.Document.Save();
        }
        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }

    private static byte[] CreateDocx(string text)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text(text)))));
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] CreateIncludeDocx(string path, string cached)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode { Text = $" INCLUDETEXT \"{path}\" " }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text(cached)),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }))));
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static async Task<string> ConvertToText(byte[] html)
    {
        byte[] docx = await new DxpHtmlToDocxConverter().ConvertAsync(html);
        using var document = WordprocessingDocument.Open(new MemoryStream(docx), false);
        return document.MainDocumentPart?.Document?.Body?.InnerText ?? string.Empty;
    }

    private static string StripTags(string html)
        => System.Net.WebUtility.HtmlDecode(System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", string.Empty));

    private static int Count(string text, string value)
        => (text.Length - text.Replace(value, string.Empty).Length) / value.Length;

    private sealed class StaticResolver(DxpIncludeTextSource source) : IDxpIncludeTextResolver
    {
        public Task<DxpIncludeTextSource?> ResolveAsync(DxpIncludeTextRequest request,
            DxpFieldEvalContext context, CancellationToken cancellationToken = default)
            => Task.FromResult<DxpIncludeTextSource?>(source);
    }

    private sealed class PathResolver(Dictionary<string, DxpIncludeTextSource> sources) : IDxpIncludeTextResolver
    {
        public Task<DxpIncludeTextSource?> ResolveAsync(DxpIncludeTextRequest request,
            DxpFieldEvalContext context, CancellationToken cancellationToken = default)
            => Task.FromResult(sources.TryGetValue(request.Path, out var source) ? source : null);
    }

    private sealed class CountingConverter(byte[] result) : IDxpHtmlToDocxConverter
    {
        public int Calls { get; private set; }
        public Task<byte[]> ConvertAsync(byte[] html, CancellationToken cancellationToken = default)
        {
            Calls++;
            return Task.FromResult(result);
        }
    }
}
