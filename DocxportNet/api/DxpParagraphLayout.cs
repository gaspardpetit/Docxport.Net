namespace DocxportNet.API;

public enum DxpComputedTabStopKind
{
    Left,
    Right,
    Center,
    Decimal
}

public enum DxpComputedTabLeaderKind
{
    None,
    Dot,
    Hyphen,
    Underscore,
    Heavy,
    MiddleDot
}

public sealed record DxpComputedTabStop(
    DxpComputedTabStopKind Kind,
    double PositionPt,
    DxpComputedTabLeaderKind Leader
);

public sealed record DxpComputedParagraphLayout(
    IReadOnlyList<DxpComputedTabStop> TabStops
);
