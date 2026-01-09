using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Walker;
using System.Xml.Linq;

namespace DocxportNet.Tests;

public class ThemeFontResolutionTests
{
	private static readonly string ProjectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));

	[Fact]
	public void DefaultRunStyle_IncludesThemeLatinFont()
	{
		string path = Path.Combine(ProjectRoot, "samples", "with_breaking_style.docx");
		using var doc = WordprocessingDocument.Open(path, false);

		string? minorLatin = null;
		var themePart = doc.MainDocumentPart?.ThemePart;
		if (themePart != null)
		{
			using var s = themePart.GetStream(FileMode.Open, FileAccess.Read);
			var xdoc = XDocument.Load(s);
			XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
			minorLatin = xdoc.Descendants(a + "fontScheme")
				.Elements(a + "minorFont")
				.Elements(a + "latin")
				.Attributes("typeface")
				.Select(a => a.Value)
				.FirstOrDefault();
		}

		var resolver = new DxpStyleResolver(doc);
		var style = resolver.GetDefaultRunStyle();

		Assert.False(string.IsNullOrWhiteSpace(minorLatin));
		Assert.False(string.IsNullOrWhiteSpace(style.FontName));
		Assert.Equal(minorLatin, style.FontName);
	}
}
