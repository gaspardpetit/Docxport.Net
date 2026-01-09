using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;

namespace DocxportNet.Walker;


public class DxpParagraphContext : DxpIParagraphContext
{
	public DxpMarker MarkerAccept { get; internal set; }
	public DxpMarker MarkerReject { get; internal set; }
	public DxpStyleEffectiveIndentTwips Indent { get; internal set; }
	public ParagraphProperties? Properties { get; internal set; }
	public DxpComputedParagraphStyle ComputedStyle { get; internal set; }
	public DxpComputedParagraphLayout? Layout { get; internal set; }

	public DxpParagraphContext(DxpMarker markerAccept, DxpMarker markerReject, DxpStyleEffectiveIndentTwips indent, ParagraphProperties? properties, DxpComputedParagraphStyle computedStyle, DxpComputedParagraphLayout? layout)
	{
		MarkerAccept = markerAccept;
		MarkerReject = markerReject;
		Indent = indent;
		Properties = properties;
		ComputedStyle = computedStyle;
		Layout = layout;
	}

	public static DxpParagraphContext INVALID => new DxpParagraphContext(
		null!,
		null!,
		null!,
		null,
		new DxpComputedParagraphStyle(
			MarginLeftPt: null,
			MarginTopPt: null,
			MarginBottomPt: null,
			TextAlign: null,
			LineHeightCss: null,
			Borders: null,
			BackgroundColorCss: null),
		null);
}


public class DxpParagraphs
{
	public static bool HasRenderableParagraphContent(Paragraph p)
	{
		// Render if there is any non-empty text, drawings, breaks, or tabs.
		if (p.Descendants<Text>().Any(t => !string.IsNullOrEmpty(t.Text)))
			return true;
		if (p.Descendants<DeletedText>().Any(t => !string.IsNullOrEmpty(t.Text)))
			return true;
		if (p.Descendants<Drawing>().Any())
			return true;
		if (p.Descendants<Break>().Any() || p.Descendants<CarriageReturn>().Any() || p.Descendants<TabChar>().Any())
			return true;
		return false;
	}

}
