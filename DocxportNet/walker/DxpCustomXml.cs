using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;

namespace DocxportNet.Walker;

public class DxpCustomXmlContext : DxpICustomXmlContext
{
	public OpenXmlElement Element { get; }
	public CustomXmlProperties? Properties { get; }

	public DxpCustomXmlContext(OpenXmlElement element, CustomXmlProperties? properties)
	{
		Element = element;
		Properties = properties;
	}
}
