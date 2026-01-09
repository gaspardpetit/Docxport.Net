namespace DocxportNet.API;

public enum DxpComputedTextAlign
{
	Left,
	Center,
	Right,
	Justify
}

public sealed record DxpComputedParagraphStyle(
	double? MarginLeftPt,
	DxpComputedTextAlign? TextAlign,
	DxpComputedBoxBorders? Borders,
	string? BackgroundColorCss
);
