using DocxportNet.Tests.Utils;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace DocxportNet.Tests;

public class HtmlExportTests : TestBase<HtmlExportTests>
{
    public sealed record Sample : IXunitSerializable
    {
        public Sample()
        {
            DocxPath = string.Empty;
        }

        public Sample(string docxPath)
        {
            DocxPath = docxPath;
        }

        public string DocxPath { get; private set; }
        public string FileName => Path.GetFileName(DocxPath);

        public void Serialize(IXunitSerializationInfo info) => info.AddValue(nameof(DocxPath), DocxPath);
        public void Deserialize(IXunitSerializationInfo info) => DocxPath = info.GetValue<string>(nameof(DocxPath));

        public override string ToString() => FileName;
    }

    private static readonly string ProjectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
    private static readonly string SamplesDirectory = Path.Combine(ProjectRoot, "samples");

    public HtmlExportTests(ITestOutputHelper output) : base(output)
    {
    }

    public static IEnumerable<object[]> SampleDocs()
    {
        return Directory.EnumerateFiles(SamplesDirectory, "*.docx", SearchOption.TopDirectoryOnly)
            .Where(path => !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
            .OrderBy(Path.GetFileName)
            .Select(path => new object[] { new Sample(path) });
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToHtml_Accept(Sample sample)
    {
        VerifyAgainstFixture(sample, DxpHtmlVisitorConfig.CreateRichConfig(), ".html", ".test.html", DxpTrackedChangeMode.AcceptChanges);
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToHtml_Reject(Sample sample)
    {
        VerifyAgainstFixture(sample, DxpHtmlVisitorConfig.CreateRichConfig(), ".reject.html", ".reject.test.html", DxpTrackedChangeMode.RejectChanges);
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToHtml_Cached(Sample sample)
    {
        VerifyCachedAgainstFixture(sample, DxpHtmlVisitorConfig.CreateRichConfig(), ".cached.html", ".cached.test.html");
    }

    private void VerifyAgainstFixture(Sample sample, DxpHtmlVisitorConfig baseConfig, string expectedExt, string actualSuffix, DxpTrackedChangeMode mode)
    {
        string expectedPath = TestPaths.GetSampleOutputPath(sample.DocxPath, expectedExt);
        string actualPath = TestPaths.GetSampleOutputPath(sample.DocxPath, actualSuffix);

        var config = CloneConfig(baseConfig, mode);
        string html = TestCompare.Normalize(ToHtml(sample.DocxPath, config));
        File.WriteAllText(actualPath, html);

        if (!File.Exists(expectedPath))
            throw new XunitException($"Expected HTML file missing for {sample.FileName} ({mode}). Add {expectedPath}. Actual output saved to {actualPath}.");

        string expectedHtml = TestCompare.Normalize(File.ReadAllText(expectedPath));

        if (!string.Equals(expectedHtml, html, StringComparison.Ordinal))
        {
            string diff = TestCompare.DescribeDifference(expectedHtml, html);
            throw new XunitException($"Mismatch for {sample.FileName} ({mode}): {diff}. Expected: {expectedPath}. Actual: {actualPath}.");
        }
    }

    private string ToHtml(string docxPath, DxpHtmlVisitorConfig config)
    {
        var visitor = new DxpHtmlVisitor(config, Logger);
        return DxpExport.ExportToString(docxPath, visitor, Logger);
    }

    private DxpHtmlVisitorConfig CloneConfig(DxpHtmlVisitorConfig source, DxpTrackedChangeMode mode)
    {
        return new DxpHtmlVisitorConfig {
            EmitImages = source.EmitImages,
            EmitStyleFont = source.EmitStyleFont,
            EmitRunColor = source.EmitRunColor,
            EmitRunBackground = source.EmitRunBackground,
            EmitTableBorders = source.EmitTableBorders,
            EmitDocumentColors = source.EmitDocumentColors,
            EmitParagraphAlignment = source.EmitParagraphAlignment,
            PreserveListSymbols = source.PreserveListSymbols,
            RichTables = source.RichTables,
            EmitSectionHeadersFooters = source.EmitSectionHeadersFooters,
            EmitUnreferencedBookmarks = source.EmitUnreferencedBookmarks,
            EmitPageNumbers = source.EmitPageNumbers,
            UsePlainComments = source.UsePlainComments,
            EmitCustomProperties = source.EmitCustomProperties,
            EmitTimeline = source.EmitTimeline,
            StylesheetHref = source.StylesheetHref,
            EmbedDefaultStylesheet = source.EmbedDefaultStylesheet,
            RootCssClass = source.RootCssClass,
            TrackedChangeMode = mode
        };
    }

    private void VerifyCachedAgainstFixture(Sample sample, DxpHtmlVisitorConfig baseConfig, string expectedExt, string actualSuffix)
    {
        string expectedPath = TestPaths.GetSampleOutputPath(sample.DocxPath, expectedExt);
        string actualPath = TestPaths.GetSampleOutputPath(sample.DocxPath, actualSuffix);

        var config = CloneConfig(baseConfig, DxpTrackedChangeMode.AcceptChanges);
        string html = TestCompare.Normalize(ToHtmlCached(sample.DocxPath, config));
        File.WriteAllText(actualPath, html);

        if (!File.Exists(expectedPath))
            throw new XunitException($"Expected HTML file missing for {sample.FileName} (CachedFields). Add {expectedPath}. Actual output saved to {actualPath}.");

        string expectedHtml = TestCompare.Normalize(File.ReadAllText(expectedPath));
        if (!string.Equals(expectedHtml, html, StringComparison.Ordinal))
        {
            string diff = TestCompare.DescribeDifference(expectedHtml, html);
            throw new XunitException($"Mismatch for {sample.FileName} (CachedFields): {diff}. Expected: {expectedPath}. Actual: {actualPath}.");
        }
    }

    private string ToHtmlCached(string docxPath, DxpHtmlVisitorConfig config)
    {
        var visitor = new DxpHtmlVisitor(config, Logger);
        using var writer = new StringWriter();
        visitor.SetOutput(writer);

        if (visitor is not DocxportNet.Fields.IDxpFieldEvalProvider provider)
            throw new XunitException("DxpHtmlVisitor should provide field evaluation context.");

        var pipeline = DocxportNet.Walker.DxpVisitorMiddleware.Chain(
            visitor,
            next => new DocxportNet.Walker.DxpFieldEvalMiddleware(next, provider.FieldEval, DocxportNet.Walker.DxpFieldEvalMode.Cache, logger: Logger),
            next => new DocxportNet.Walker.DxpContextTracker(next));

        new DocxportNet.Walker.DxpWalker(Logger).Accept(docxPath, pipeline);
        return writer.ToString();
    }
}
