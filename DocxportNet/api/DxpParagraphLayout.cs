namespace DocxportNet.API;

public enum DxpComputedTabStopKind
{
	Left,
	Right,
	Center,
	Decimal
}

public sealed record DxpComputedTabStop(
	DxpComputedTabStopKind Kind,
	double PositionPt
);

public sealed record DxpComputedParagraphLayout(
	IReadOnlyList<DxpComputedTabStop> TabStops
);

