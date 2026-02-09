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
}
