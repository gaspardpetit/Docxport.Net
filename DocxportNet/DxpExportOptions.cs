namespace DocxportNet;

public enum DxpFieldEvalExportMode
{
    None,
    Evaluate,
    Cache
}

/// <summary>Controls field handling when rebuilding a DOCX.</summary>
public enum DxpDocxFieldPolicy
{
    Preserve,
    Evaluate
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
    private DxpFieldEvalExportMode _fieldEvalMode = DxpFieldEvalExportMode.Evaluate;
    internal bool HasExplicitFieldEvalMode { get; private set; }
    public DxpFieldEvalExportMode FieldEvalMode
    {
        get => _fieldEvalMode;
        set
        {
            _fieldEvalMode = value;
            HasExplicitFieldEvalMode = true;
        }
    }
    /// <summary>
    /// When unset, a caller-supplied evaluator selects evaluation; otherwise
    /// DOCX rebuilding preserves reference and cached layout-sensitive fields.
    /// </summary>
    public DxpDocxFieldPolicy? DocxFieldPolicy { get; set; }
    public Func<string?, bool>? FieldEvaluationFilter { get; set; }
    /// <summary>
    /// Optional progress reporter. Supplying one enables a lightweight paragraph-counting pre-pass.
    /// </summary>
    public IProgress<DxpExportProgress>? Progress { get; set; }
}
