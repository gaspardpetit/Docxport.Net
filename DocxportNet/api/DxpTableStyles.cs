namespace DocxportNet.API;

public enum DxpComputedBorderLineStyle
{
    None,
    Solid,
    Dotted,
    Dashed,
    Double
}

public sealed record DxpComputedBorder(
    double WidthPt,
    DxpComputedBorderLineStyle LineStyle,
    string ColorCss
);

public sealed record DxpComputedBoxBorders(
    DxpComputedBorder? Top,
    DxpComputedBorder? Right,
    DxpComputedBorder? Bottom,
    DxpComputedBorder? Left
);

public enum DxpComputedVerticalAlign
{
    Top,
    Middle,
    Bottom
}

public sealed record DxpComputedTableStyle(
    DxpComputedBorder? TableBorder,
    bool BorderCollapse,
    DxpComputedBorder? DefaultCellBorder
)
{
    public DxpComputedBoxBorders? TableBorders { get; init; }
    public DxpComputedBoxBorders? DefaultCellBorders { get; init; }
}

public sealed record DxpComputedTableCellStyle(
    DxpComputedBorder? Border,
    string? BackgroundColorCss,
    DxpComputedVerticalAlign? VerticalAlign
)
{
    public DxpComputedBoxBorders? Borders { get; init; }
}
