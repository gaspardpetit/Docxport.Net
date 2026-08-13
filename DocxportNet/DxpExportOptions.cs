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
    /// This remains opt-in while exporter parity is being established.
    /// </summary>
    public bool UseSemanticFieldResults { get; set; }
}
