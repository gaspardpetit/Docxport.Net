using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;

namespace DocxportNet.Walker;

public class DxpHyperlinks
{
	public static DxpLinkAnchor? ResolveHyperlinkTarget(Hyperlink link, OpenXmlPart currentPart)
	{
		// Anchor links are direct
		if (!string.IsNullOrEmpty(link.Anchor?.Value))
		{
			return new DxpLinkAnchor(link.Anchor!.Value!, $"#{link.Anchor!.Value!}");
		}

		string? relId = link.Id?.Value;
		OpenXmlPart part = currentPart;
		if (string.IsNullOrEmpty(relId) || part == null)
			return null;

		HyperlinkRelationship? rel = part.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
		if (rel == null)
			return null;

		return new DxpLinkAnchor(null, rel.Uri?.ToString()!);
	}
}
