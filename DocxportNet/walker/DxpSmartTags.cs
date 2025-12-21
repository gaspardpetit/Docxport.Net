using DocumentFormat.OpenXml;

namespace DocxportNet.Walker;

public class DxpSmartTags
{
	public static bool IsWSmartTag(OpenXmlUnknownElement unk) =>
		unk != null
		&& unk.LocalName == "smartTag"
		&& (unk.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
			|| unk.NamespaceUri == "http://purl.oclc.org/ooxml/wordprocessingml/main"); // some tools map to purl
}
