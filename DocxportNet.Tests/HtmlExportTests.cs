using DocxportNet.Tests.Utils;
using DocxportNet.Visitors.Html;
using DocxportNet.Walker;
using DocxportNet.Visitors.Markdown;
using Microsoft.Extensions.Logging;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace DocxportNet.Tests;

public class HtmlExportTests
{
	private readonly ILogger _logger;

	public HtmlExportTests(ITestOutputHelper output)
	{
		_logger = new TestLogger(output, nameof(HtmlExportTests));
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

		public void Serialize(IXunitSerializationInfo info) => info.AddValue(nameof(DocxPath), DocxPath);
		public void Deserialize(IXunitSerializationInfo info) => DocxPath = info.GetValue<string>(nameof(DocxPath));

		public override string ToString() => FileName;
	}

	private static readonly string ProjectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
	private static readonly string SamplesDirectory = Path.Combine(ProjectRoot, "samples");

	public static IEnumerable<object[]> SampleDocs()
	{
		return Directory.EnumerateFiles(SamplesDirectory, "*.docx", SearchOption.TopDirectoryOnly)
			.OrderBy(Path.GetFileName)
			.Select(path => new object[] { new Sample(path) });
	}

	[Theory]
	[MemberData(nameof(SampleDocs))]
	public void TestDocxToHtml_Accept(Sample sample)
	{
		VerifyAgainstFixture(sample, DxpHtmlVisitorConfig.RICH, ".html", ".test.html", DxpTrackedChangeMode.AcceptChanges);
	}

	[Theory]
	[MemberData(nameof(SampleDocs))]
	public void TestDocxToHtml_Reject(Sample sample)
	{
		VerifyAgainstFixture(sample, DxpHtmlVisitorConfig.RICH, ".reject.html", ".reject.test.html", DxpTrackedChangeMode.RejectChanges);
	}

	private void VerifyAgainstFixture(Sample sample, DxpHtmlVisitorConfig baseConfig, string expectedExt, string actualSuffix, DxpTrackedChangeMode mode)
	{
		string expectedPath = Path.ChangeExtension(sample.DocxPath, expectedExt);
		string actualPath = Path.Combine(Path.GetDirectoryName(sample.DocxPath)!, $"{Path.GetFileNameWithoutExtension(sample.DocxPath)}{actualSuffix}");

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
		using var writer = new StringWriter();
		var visitor = new DxpHtmlVisitor(writer, config, _logger);
		var walker = new DxpWalker();

		walker.Accept(docxPath, visitor);
		return writer.ToString();
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
}
