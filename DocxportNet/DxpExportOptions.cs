namespace DocxportNet;

public enum DxpFieldEvalExportMode
{
    None,
    Evaluate,
    Cache
}

public sealed class DxpExportOptions
{
    public DxpFieldEvalExportMode FieldEvalMode { get; set; } = DxpFieldEvalExportMode.Evaluate;
    /// <summary>
    /// Uses the composable, format-neutral result pipeline for supported fields.
    /// Disable this only as a temporary compatibility fallback to the legacy pipeline.
    /// </summary>
    public bool UseSemanticFieldResults { get; set; } = true;
}
