using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Walker;

namespace DocxportNet.API;


public sealed record DxpStyleInfo(
	string StyleId,
	string? Name,
	StyleValues? Type,
	string? BasedOnStyleId
);

public sealed record DxpStyleEffectiveRunStyle(
	bool Bold,
	bool Italic,
	bool Underline,
	bool Strike,
	bool DoubleStrike,
	bool Superscript,
	bool Subscript,
	bool AllCaps,
	bool SmallCaps,
	string? FontName,
	int? FontSizeHalfPoints
);


public sealed record DxpStyleEffectiveIndentTwips(
	int? Left,      // w:ind/@w:left or @w:start
	int? Right,     // w:ind/@w:right or @w:end
	int? FirstLine, // w:ind/@w:firstLine
	int? Hanging    // w:ind/@w:hanging
);

public sealed record DxpStyleEffectiveNumPr(int NumId, int Ilvl);

public interface DxpIStyleResolver
{
	DxpStyleEffectiveNumPr? ResolveEffectiveNumPr(Paragraph p);
	DxpStyleEffectiveIndentTwips GetIndentation(Paragraph p, DxpNumberingResolver? num);
	DxpStyleEffectiveRunStyle ResolveRunStyle(Paragraph paragraph, Run r);
	string? ResolveRunLanguage(Paragraph paragraph, Run r);
	int? GetHeadingLevel(Paragraph p);
	IReadOnlyList<DxpStyleInfo> GetParagraphStyleChain(Paragraph p);
	DxpStyleEffectiveRunStyle GetDefaultRunStyle();
}
