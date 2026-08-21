namespace DocxportNet.API;

public sealed record DxpImageCrop(double Left, double Top, double Right, double Bottom)
{
    public bool IsEmpty => Left <= 0 && Top <= 0 && Right <= 0 && Bottom <= 0;
}

public sealed record DxpImagePresentation
{
    public double? FrameWidthPoints { get; init; }
    public double? FrameHeightPoints { get; init; }
    public DxpImageCrop? Crop { get; init; }
    public double RotationDegrees { get; init; }
    public bool FlipHorizontal { get; init; }
    public bool FlipVertical { get; init; }
    public string? AlternativeText { get; init; }
    public string? Title { get; init; }
    public bool IsDecorative { get; init; }
}

public sealed record DxpDrawingInfo(
    string? EmbedRelId,
    string? ContentType,
    string? FileName,
    string? AltText,
    string? DataUri
)
{
    public string? ExternalSource { get; init; }
    public DxpImagePresentation? Presentation { get; init; }
}
