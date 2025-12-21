using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;

namespace DocxportNet.Walker;

public class DxpSmartTagContext : DxpISmartTagContext
{
	public OpenXmlUnknownElement SmartTag { get; }
	public string ElementName { get; }
	public string ElementUri { get; }
	public IReadOnlyList<CustomXmlAttribute> Attributes { get; }

	public DxpSmartTagContext(OpenXmlUnknownElement smartTag, string elementName, string elementUri, IReadOnlyList<CustomXmlAttribute> attributes)
	{
		SmartTag = smartTag;
		ElementName = elementName;
		ElementUri = elementUri;
		Attributes = attributes;
	}
}

public static class DxpSmartTags
{
	private const string WNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

	public static bool IsWSmartTag(OpenXmlUnknownElement element)
	{
		return string.Equals(element.LocalName, "smartTag", StringComparison.Ordinal)
			&& string.Equals(element.NamespaceUri, WNamespace, StringComparison.Ordinal);
	}
}
