using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;

namespace DocxportNet.Walker;


public class DxpParagraphContext : DxpIParagraphContext
{
	public DxpMarker Marker { get; internal set; }
	public DxpStyleEffectiveIndentTwips Indent { get; internal set; }

	public DxpParagraphContext(DxpMarker marker, DxpStyleEffectiveIndentTwips indent)
	{
		Marker = marker;
		Indent = indent;
	}

	public static DxpParagraphContext INVALID => new DxpParagraphContext(null!, null!);
}


public class DxpParagraphs
{
	public static bool HasRenderableParagraphContent(Paragraph p)
	{
		// Render if there is any non-empty text, drawings, breaks, or tabs.
		if (p.Descendants<Text>().Any(t => !string.IsNullOrEmpty(t.Text)))
			return true;
		if (p.Descendants<Drawing>().Any())
			return true;
		if (p.Descendants<Break>().Any() || p.Descendants<CarriageReturn>().Any() || p.Descendants<TabChar>().Any())
			return true;
		return false;
	}

}
