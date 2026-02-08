using DocxportNet.Tests.Utils;
using DocxportNet.Visitors.Markdown;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace DocxportNet.Tests;

public class MarkdownExportTests : TestBase<MarkdownExportTests>
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
        public string ExpectedMarkdownPath => Path.ChangeExtension(DocxPath, ".md");
        public string TestOutputPath => TestPaths.GetSampleOutputPath(DocxPath, ".test.md");

        public void Serialize(IXunitSerializationInfo info) => info.AddValue(nameof(DocxPath), DocxPath);
        public void Deserialize(IXunitSerializationInfo info) => DocxPath = info.GetValue<string>(nameof(DocxPath));

        public override string ToString() => FileName; // keep theory display concise
    }

    private static readonly string ProjectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
    private static readonly string SamplesDirectory = Path.Combine(ProjectRoot, "samples");

    public MarkdownExportTests(ITestOutputHelper output) : base(output)
    {
    }

    public static IEnumerable<object[]> SampleDocs()
    {
        return Directory.EnumerateFiles(SamplesDirectory, "*.docx", SearchOption.TopDirectoryOnly)
            .Where(path => !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
            .OrderBy(Path.GetFileName) // deterministic ordering for discovery
            .Select(path => new object[] { new Sample(path) });
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToMarkdown_Rich(Sample sample)
    {
        VerifyAgainstFixture(sample, DxpMarkdownVisitorConfig.CreateRichConfig(), ".md", ".test.md");
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToMarkdown_Plain(Sample sample)
    {
        VerifyAgainstFixture(sample, DxpMarkdownVisitorConfig.CreatePlainConfig(), ".plain.md", ".plain.test.md");
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToMarkdown_Rich_Reject(Sample sample)
    {
        VerifyVariant(sample, DxpMarkdownVisitorConfig.CreateRichConfig(), ".reject.test.md", DxpTrackedChangeMode.RejectChanges);
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToMarkdown_Rich_Inline(Sample sample)
    {
        VerifyVariant(sample, DxpMarkdownVisitorConfig.CreateRichConfig(), ".inline.test.md", DxpTrackedChangeMode.InlineChanges);
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToMarkdown_Rich_Split(Sample sample)
    {
        VerifyVariant(sample, DxpMarkdownVisitorConfig.CreateRichConfig(), ".split.test.md", DxpTrackedChangeMode.SplitChanges);
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToMarkdown_Plain_Reject(Sample sample)
    {
        VerifyVariant(sample, DxpMarkdownVisitorConfig.CreatePlainConfig(), ".plain.reject.test.md", DxpTrackedChangeMode.RejectChanges);
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToMarkdown_Plain_Split(Sample sample)
    {
        VerifyVariant(sample, DxpMarkdownVisitorConfig.CreatePlainConfig(), ".plain.split.test.md", DxpTrackedChangeMode.SplitChanges);
    }


    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToMarkdown_Plain_Inline(Sample sample)
    {
        VerifyVariant(sample, DxpMarkdownVisitorConfig.CreatePlainConfig(), ".plain.inline.test.md", DxpTrackedChangeMode.InlineChanges);
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToMarkdown_Plain_Cached(Sample sample)
    {
        VerifyCachedAgainstFixture(sample, DxpMarkdownVisitorConfig.CreatePlainConfig(), ".plain.cached.md", ".plain.cached.test.md");
    }

    private void VerifyAgainstFixture(Sample sample, DxpMarkdownVisitorConfig config, string expectedExt, string actualSuffix)
    {
        string expectedPath = TestPaths.GetSampleOutputPath(sample.DocxPath, expectedExt);
        string actualPath = TestPaths.GetSampleOutputPath(sample.DocxPath, actualSuffix);

        string actualMarkdown = TestCompare.Normalize(ToMarkdown(sample.DocxPath, CloneConfig(config, DxpTrackedChangeMode.AcceptChanges)));
        File.WriteAllText(actualPath, actualMarkdown);

        if (!File.Exists(expectedPath))
        {
            throw new XunitException($"Expected markdown file missing for {sample.FileName}. Add {expectedPath}. Actual output saved to {actualPath}.");
        }

        string expectedMarkdown = TestCompare.Normalize(File.ReadAllText(expectedPath));

        if (!string.Equals(expectedMarkdown, actualMarkdown, StringComparison.Ordinal))
        {
            string diff = TestCompare.DescribeDifference(expectedMarkdown, actualMarkdown);
            throw new XunitException($"Mismatch for {sample.FileName}: {diff}. Expected: {expectedPath}. Actual: {actualPath}.");
        }

        // Emit additional tracked-change variants for inspection.
        WriteVariant(sample, config, DxpTrackedChangeMode.RejectChanges, actualSuffix.Replace(".test", ".reject.test"));
        WriteVariant(sample, config, DxpTrackedChangeMode.InlineChanges, actualSuffix.Replace(".test", ".inline.test"));
    }

    private void VerifyVariant(Sample sample, DxpMarkdownVisitorConfig config, string suffix, DxpTrackedChangeMode mode)
    {
        string expectedPath = TestPaths.GetSampleOutputPath(sample.DocxPath, suffix.Replace(".test", string.Empty));
        string actualPath = TestPaths.GetSampleOutputPath(sample.DocxPath, suffix);

        var cfg = CloneConfig(config, mode);
        string markdown = TestCompare.Normalize(ToMarkdown(sample.DocxPath, cfg));
        File.WriteAllText(actualPath, markdown);

        if (!File.Exists(expectedPath))
            throw new XunitException($"Expected markdown file missing for {sample.FileName} ({mode}). Add {expectedPath}. Actual output saved to {actualPath}.");

        string expectedMarkdown = TestCompare.Normalize(File.ReadAllText(expectedPath));
        if (!string.Equals(expectedMarkdown, markdown, StringComparison.Ordinal))
        {
            string diff = TestCompare.DescribeDifference(expectedMarkdown, markdown);
            throw new XunitException($"Mismatch for {sample.FileName} ({mode}): {diff}. Expected: {expectedPath}. Actual: {actualPath}.");
        }
    }

    private void WriteVariant(Sample sample, DxpMarkdownVisitorConfig baseConfig, DxpTrackedChangeMode mode, string suffix)
    {
        var cfg = CloneConfig(baseConfig, mode);
        string path = TestPaths.GetSampleOutputPath(sample.DocxPath, suffix);
        string markdown = TestCompare.Normalize(ToMarkdown(sample.DocxPath, cfg));
        File.WriteAllText(path, markdown);
    }

    private string ToMarkdown(string docxPath, DxpMarkdownVisitorConfig config)
    {
        var visitor = new DxpMarkdownVisitor(config, Logger);
        return DxpExport.ExportToString(docxPath, visitor, Logger);
    }

    private void VerifyCachedAgainstFixture(Sample sample, DxpMarkdownVisitorConfig config, string expectedExt, string actualSuffix)
    {
        string expectedPath = TestPaths.GetSampleOutputPath(sample.DocxPath, expectedExt);
        string actualPath = TestPaths.GetSampleOutputPath(sample.DocxPath, actualSuffix);

        string actualMarkdown = TestCompare.Normalize(ToMarkdownCached(sample.DocxPath, CloneConfig(config, DxpTrackedChangeMode.AcceptChanges)));
        File.WriteAllText(actualPath, actualMarkdown);

        if (!File.Exists(expectedPath))
            throw new XunitException($"Expected markdown file missing for {sample.FileName} (CachedFields). Add {expectedPath}. Actual output saved to {actualPath}.");

        string expectedMarkdown = TestCompare.Normalize(File.ReadAllText(expectedPath));
        if (!string.Equals(expectedMarkdown, actualMarkdown, StringComparison.Ordinal))
        {
            string diff = TestCompare.DescribeDifference(expectedMarkdown, actualMarkdown);
            throw new XunitException($"Mismatch for {sample.FileName} (CachedFields): {diff}. Expected: {expectedPath}. Actual: {actualPath}.");
        }
    }

    private string ToMarkdownCached(string docxPath, DxpMarkdownVisitorConfig config)
    {
        var visitor = new DxpMarkdownVisitor(config, Logger);
        using var writer = new StringWriter();
        visitor.SetOutput(writer);

        if (visitor is not DocxportNet.Fields.IDxpFieldEvalProvider provider)
            throw new XunitException("DxpMarkdownVisitor should provide field evaluation context.");

        var pipeline = DocxportNet.Walker.DxpVisitorMiddleware.Chain(
            visitor,
            next => new DocxportNet.Walker.DxpFieldEvalMiddleware(next, provider.FieldEval, DocxportNet.Walker.DxpFieldEvalMode.Cache, logger: Logger),
            next => new DocxportNet.Walker.DxpContextTracker(next));

        new DocxportNet.Walker.DxpWalker(Logger).Accept(docxPath, pipeline);
        return writer.ToString();
    }

    private static DxpMarkdownVisitorConfig CloneConfig(DxpMarkdownVisitorConfig source, DxpTrackedChangeMode mode)
    {
        return new DxpMarkdownVisitorConfig {
            EmitImages = source.EmitImages,
            EmitStyleFont = source.EmitStyleFont,
            EmitRunColor = source.EmitRunColor,
            EmitRunBackground = source.EmitRunBackground,
            EmitTableBorders = source.EmitTableBorders,
            EmitDocumentColors = source.EmitDocumentColors,
            EmitParagraphAlignment = source.EmitParagraphAlignment,
            PreserveListSymbols = source.PreserveListSymbols,
            RichTables = source.RichTables,
            UsePlainCodeBlocks = source.UsePlainCodeBlocks,
            UseMarkdownInlineStyles = source.UseMarkdownInlineStyles,
            EmitSectionHeadersFooters = source.EmitSectionHeadersFooters,
            EmitUnreferencedBookmarks = source.EmitUnreferencedBookmarks,
            EmitPageNumbers = source.EmitPageNumbers,
            UsePlainComments = source.UsePlainComments,
            EmitCustomProperties = source.EmitCustomProperties,
            EmitTimeline = source.EmitTimeline,
            TrackedChangeMode = mode
        };
    }
}
