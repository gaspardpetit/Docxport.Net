namespace DocxportNet.Fields;

public enum DxpFieldEvalStatus
{
    Resolved,
    UsedCache,
    Skipped,
    Failed,
}

public sealed record DxpFieldEvalResult(
    DxpFieldEvalStatus Status,
    string? Text,
    Exception? Error = null);
