using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;

namespace DocxportNet.Walker;

public class DxpRunContext : DxpIRunContext
{
	public Run Run { get; }
	public RunProperties? Properties { get; }
	public DxpStyleEffectiveRunStyle Style { get; }
	public string? Language { get; }

	public DxpRunContext(Run run, RunProperties? properties, DxpStyleEffectiveRunStyle style, string? language)
	{
		Run = run;
		Properties = properties;
		Style = style;
		Language = language;
	}
}
