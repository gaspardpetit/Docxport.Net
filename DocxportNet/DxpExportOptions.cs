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
    /// This remains opt-in until differential parity is complete.
    /// </summary>
    public bool UseSemanticFieldResults { get; set; }
}
