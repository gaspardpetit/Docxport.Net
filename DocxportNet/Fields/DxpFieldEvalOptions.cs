namespace DocxportNet.Fields;

public sealed class DxpFieldEvalOptions
{
    public bool UseCacheOnNull { get; init; } = true;
    public bool UseCacheOnError { get; init; } = true;
    public bool ErrorOnUnsupported { get; init; } = false;
}
