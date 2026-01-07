using DocxportNet.Tests.Utils;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Walker;
using Microsoft.Extensions.Logging;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace DocxportNet.Tests;

public class MarkdownExportTests
{
	private readonly ILogger _logger;

	public MarkdownExportTests(ITestOutputHelper output)
	{
		_logger = new TestLogger(output, nameof(MarkdownExportTests));
	}

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
		public string TestOutputPath => Path.Combine(Path.GetDirectoryName(DocxPath)!, $"{Path.GetFileNameWithoutExtension(DocxPath)}.test.md");

		public void Serialize(IXunitSerializationInfo info) => info.AddValue(nameof(DocxPath), DocxPath);
		public void Deserialize(IXunitSerializationInfo info) => DocxPath = info.GetValue<string>(nameof(DocxPath));

		public override string ToString() => FileName; // keep theory display concise
	}

	private static readonly string ProjectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
	private static readonly string SamplesDirectory = Path.Combine(ProjectRoot, "samples");

	public static IEnumerable<object[]> SampleDocs()
	{
		return Directory.EnumerateFiles(SamplesDirectory, "*.docx", SearchOption.TopDirectoryOnly)
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

	private void VerifyAgainstFixture(Sample sample, DxpMarkdownVisitorConfig config, string expectedExt, string actualSuffix)
	{
		string expectedPath = Path.ChangeExtension(sample.DocxPath, expectedExt);
		string actualPath = Path.Combine(Path.GetDirectoryName(sample.DocxPath)!, $"{Path.GetFileNameWithoutExtension(sample.DocxPath)}{actualSuffix}");

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
		string expectedPath = Path.Combine(Path.GetDirectoryName(sample.DocxPath)!, $"{Path.GetFileNameWithoutExtension(sample.DocxPath)}{suffix.Replace(".test", string.Empty)}");
		string actualPath = Path.Combine(Path.GetDirectoryName(sample.DocxPath)!, $"{Path.GetFileNameWithoutExtension(sample.DocxPath)}{suffix}");

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
		string path = Path.Combine(Path.GetDirectoryName(sample.DocxPath)!, $"{Path.GetFileNameWithoutExtension(sample.DocxPath)}{suffix}");
		string markdown = TestCompare.Normalize(ToMarkdown(sample.DocxPath, cfg));
		File.WriteAllText(path, markdown);
	}

	private string ToMarkdown(string docxPath, DxpMarkdownVisitorConfig config)
	{
		using var writer = new StringWriter();
		var visitor = new DxpMarkdownVisitor(writer, config, _logger);
		var walker = new DxpWalker();

		walker.Accept(docxPath, visitor);
		return writer.ToString();
	}

	private DxpMarkdownVisitorConfig CloneConfig(DxpMarkdownVisitorConfig source, DxpTrackedChangeMode mode)
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
