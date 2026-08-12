using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Fields;
using DocxportNet.Fields.Resolution;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class BookmarkIncludeTextTests
{
    [Fact]
    public void BookmarkWithinParagraph_SelectsOnlyRangeAndPreservesFormatting()
    {
        byte[] child = CreateDocx(new Paragraph(
            new Run(new Text("OUTSIDE-BEFORE")),
            new BookmarkStart { Name = "Selected", Id = "1" },
            new Run(new RunProperties(new Bold()), new Text("INSIDE")),
            new BookmarkEnd { Id = "1" },
            new Run(new Text("OUTSIDE-AFTER"))));
        using var parent = CreateParent("child.docx", "selected", "CACHE", "PARENT-", "-END");
        var visitor = CreateHtmlVisitor(child);

        string html = Export(parent, visitor);

        _ = System.Xml.Linq.XDocument.Parse(html);
        Assert.Contains("PARENT-<strong class=\"dxp-bold\">INSIDE</strong>-END", html, StringComparison.Ordinal);
        Assert.DoesNotContain("OUTSIDE", html, StringComparison.Ordinal);
        Assert.DoesNotContain("CACHE", html, StringComparison.Ordinal);
    }

    [Fact]
    public void BookmarkAcrossBlocks_PreservesParagraphsAndTableSplicing()
    {
        byte[] child = CreateDocx(
            new Paragraph(new Run(new Text("BEFORE")), new BookmarkStart { Name = "Range", Id = "2" }, new Run(new Text("FIRST"))),
            new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text("TABLE")))))),
            new Paragraph(new Run(new Text("LAST")), new BookmarkEnd { Id = "2" }, new Run(new Text("AFTER"))));
        using var parent = CreateParent("child.docx", "Range", "CACHE", "PARENT-", "-END");
        var visitor = CreateHtmlVisitor(child);

        string html = Export(parent, visitor);
        string text = StripTags(html);

        _ = System.Xml.Linq.XDocument.Parse(html);
        Assert.Contains("PARENT-FIRST", text, StringComparison.Ordinal);
        Assert.Contains("LAST-END", text, StringComparison.Ordinal);
        Assert.True(text.IndexOf("FIRST", StringComparison.Ordinal) < text.IndexOf("TABLE", StringComparison.Ordinal));
        Assert.True(text.IndexOf("TABLE", StringComparison.Ordinal) < text.IndexOf("LAST", StringComparison.Ordinal));
        Assert.DoesNotContain("BEFORE", text, StringComparison.Ordinal);
        Assert.DoesNotContain("AFTER", text, StringComparison.Ordinal);
    }

    [Fact]
    public void NestedBookmarks_SelectRequestedOuterOrInnerRange()
    {
        byte[] child = CreateDocx(new Paragraph(
            new BookmarkStart { Name = "Outer", Id = "10" },
            new Run(new Text("OUTER-START")),
            new BookmarkStart { Name = "Inner", Id = "11" },
            new Run(new Text("INNER")),
            new BookmarkEnd { Id = "11" },
            new Run(new Text("OUTER-END")),
            new BookmarkEnd { Id = "10" }));

        string inner = ExportPlain(child, "inner");
        string outer = ExportPlain(child, "OUTER");

        Assert.Contains("INNER", inner, StringComparison.Ordinal);
        Assert.DoesNotContain("OUTER-", inner, StringComparison.Ordinal);
        Assert.Contains("OUTER-STARTINNEROUTER-END", RemoveWhitespace(outer), StringComparison.Ordinal);
    }

    [Fact]
    public void SelectedRange_EvaluatesNestedFields()
    {
        byte[] child = CreateDocx(new Paragraph(
            new BookmarkStart { Name = "Range", Id = "3" },
            new SimpleField(new Run(new Text("FIELD-CACHE"))) { Instruction = " DOCVARIABLE ChildValue " },
            new BookmarkEnd { Id = "3" }));
        using var parent = CreateParent("child.docx", "Range", "CACHE");
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig());
        visitor.FieldEval.Context.IncludeTextResolver = new StaticResolver(child);
        visitor.FieldEval.Context.SetDocVariable("ChildValue", "FIELD-VALUE");

        string output = Export(parent, visitor);

        Assert.Contains("FIELD-VALUE", output, StringComparison.Ordinal);
        Assert.DoesNotContain("FIELD-CACHE", output, StringComparison.Ordinal);
    }

    [Fact]
    public void CollapsedBookmark_IsSuccessfulEmptyInclusion()
    {
        byte[] child = CreateDocx(new Paragraph(
            new Run(new Text("OUTSIDE")),
            new BookmarkStart { Name = "Empty", Id = "4" },
            new BookmarkEnd { Id = "4" }));
        using var parent = CreateParent("child.docx", "Empty", "CACHE", "BEFORE", "AFTER");
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig());
        visitor.FieldEval.Context.IncludeTextResolver = new StaticResolver(child);

        string output = Export(parent, visitor);

        Assert.Contains("BEFOREAFTER", RemoveWhitespace(output), StringComparison.Ordinal);
        Assert.DoesNotContain("OUTSIDE", output, StringComparison.Ordinal);
        Assert.DoesNotContain("CACHE", output, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("Missing", false, false)]
    [InlineData("Broken", true, false)]
    [InlineData("Reversed", false, true)]
    public void InvalidBookmark_ReplaysCacheWithoutPartialOutput(string name, bool startOnly, bool endFirst)
    {
        var content = new List<OpenXmlElement>();
        if (endFirst)
            content.Add(new BookmarkEnd { Id = "5" });
        if (startOnly || endFirst)
            content.Add(new BookmarkStart { Name = name, Id = "5" });
        content.Add(new Run(new Text("CHILD")));
        byte[] child = CreateDocx(new Paragraph(content));
        using var parent = CreateParent("child.docx", name, "CACHED", "BEFORE-", "-AFTER");
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig());
        visitor.FieldEval.Context.IncludeTextResolver = new StaticResolver(child);

        string output = Export(parent, visitor);

        Assert.Contains("BEFORE-CACHED-AFTER", RemoveWhitespace(output), StringComparison.Ordinal);
        Assert.DoesNotContain("CHILD", output, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlNamedAnchor_CanBeSelectedAfterConversion()
    {
        using var parent = CreateParent("child.html", "target", "CACHE");
        const string image = "https://assets.example.test/selected.png";
        var source = new DxpIncludeTextSource("html", Encoding.UTF8.GetBytes(
            $"<p>BEFORE</p><p id=\"target\">SELECTED<img src=\"{image}\"></p><p>AFTER</p>"))
        {
            Format = DxpIncludeTextSourceFormat.Html
        };
        var visitor = new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig());
        visitor.FieldEval.Context.IncludeTextResolver = new SourceResolver(source);

        string output = Export(parent, visitor);

        _ = System.Xml.Linq.XDocument.Parse(output);
        Assert.Contains("SELECTED", output, StringComparison.Ordinal);
        Assert.Contains(image, output, StringComparison.Ordinal);
        Assert.DoesNotContain("BEFORE", output, StringComparison.Ordinal);
        Assert.DoesNotContain("AFTER", output, StringComparison.Ordinal);
        Assert.DoesNotContain("CACHE", output, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlDroppedElementId_DoesNotBreakUnqualifiedIncludeOrLeakMarkers()
    {
        using var parent = CreateUnqualifiedParent("child.html", "CACHE");
        var source = new DxpIncludeTextSource("html", Encoding.UTF8.GetBytes(
            "<script id=\"not-rendered\">ignored()</script><p>VISIBLE</p>"))
        {
            Format = DxpIncludeTextSourceFormat.Html
        };
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig());
        visitor.FieldEval.Context.IncludeTextResolver = new SourceResolver(source);

        string output = Export(parent, visitor);

        Assert.Contains("VISIBLE", output, StringComparison.Ordinal);
        Assert.DoesNotContain("CACHE", output, StringComparison.Ordinal);
        Assert.DoesNotContain("DXPBM", output, StringComparison.Ordinal);
    }

    private static DxpHtmlVisitor CreateHtmlVisitor(byte[] child)
    {
        var visitor = new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig());
        visitor.FieldEval.Context.IncludeTextResolver = new StaticResolver(child);
        return visitor;
    }

    private static string ExportPlain(byte[] child, string bookmark)
    {
        using var parent = CreateParent("child.docx", bookmark, "CACHE");
        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig());
        visitor.FieldEval.Context.IncludeTextResolver = new StaticResolver(child);
        return Export(parent, visitor);
    }

    private static string Export(WordprocessingDocument parent, DocxportNet.API.DxpITextVisitor visitor)
        => DxpExport.ExportToString(parent, visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate });

    private static WordprocessingDocument CreateParent(
        string path, string bookmark, string cached, string? before = null, string? after = null)
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var content = new List<OpenXmlElement>();
            if (before != null)
                content.Add(new Run(new Text(before)));
            content.AddRange([
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode { Text = $" INCLUDETEXT \"{path}\" {bookmark} " }),
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

    private static WordprocessingDocument CreateUnqualifiedParent(string path, string cached)
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
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
        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }

    private static byte[] CreateDocx(params OpenXmlElement[] blocks)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(blocks));
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static string StripTags(string html)
        => System.Net.WebUtility.HtmlDecode(System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", string.Empty));

    private static string RemoveWhitespace(string value)
        => string.Concat(value.Where(character => !char.IsWhiteSpace(character)));

    private sealed class StaticResolver(byte[] content) : IDxpIncludeTextResolver
    {
        public Task<DxpIncludeTextSource?> ResolveAsync(DxpIncludeTextRequest request,
            DxpFieldEvalContext context, CancellationToken cancellationToken = default)
            => Task.FromResult<DxpIncludeTextSource?>(new DxpIncludeTextSource("child", content));
    }

    private sealed class SourceResolver(DxpIncludeTextSource source) : IDxpIncludeTextResolver
    {
        public Task<DxpIncludeTextSource?> ResolveAsync(DxpIncludeTextRequest request,
            DxpFieldEvalContext context, CancellationToken cancellationToken = default)
            => Task.FromResult<DxpIncludeTextSource?>(source);
    }
}
