namespace DocxportNet;

public enum DxpFieldEvalExportMode
{
    None,
    Evaluate,
    Cache
}

/// <summary>Identifies the current stage of a DOCX export.</summary>
public enum DxpExportPhase
{
    Opening,
    Preparing,
    Converting,
    Finalizing,
    Completed
}

/// <summary>
/// Reports export progress. In the initial implementation, one unit represents
/// one source-document paragraph.
/// </summary>
public readonly record struct DxpExportProgress(
    DxpExportPhase Phase,
    long CompletedUnits,
    long TotalUnits)
{
    /// <summary>
    /// Gets the overall percentage, or <see langword="null"/> while the total is
    /// not yet known. One hundred is reserved for a successfully completed export.
    /// </summary>
    public double? Percentage => Phase switch
    {
        DxpExportPhase.Opening or DxpExportPhase.Preparing => null,
        DxpExportPhase.Completed => 100d,
        _ when TotalUnits == 0 => 0d,
        _ => Math.Min(99d, 100d * CompletedUnits / TotalUnits)
    };
}

public sealed class DxpExportOptions
{
    public DxpFieldEvalExportMode FieldEvalMode { get; set; } = DxpFieldEvalExportMode.Evaluate;
    public Func<string?, bool>? FieldEvaluationFilter { get; set; }
    /// <summary>
    /// Optional progress reporter. Supplying one enables a lightweight paragraph-counting pre-pass.
    /// </summary>
    public IProgress<DxpExportProgress>? Progress { get; set; }
}
