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
	double? MarginTopPt,
	double? MarginBottomPt,
	DxpComputedTextAlign? TextAlign,
	string? LineHeightCss,
	DxpComputedBoxBorders? Borders,
	string? BackgroundColorCss
);
