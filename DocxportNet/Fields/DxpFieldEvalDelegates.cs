namespace DocxportNet.Fields;

public sealed class DxpFieldEvalDelegates
{
    public Func<DxpFieldInstruction, DxpFieldEvalContext, Task<DxpFieldValue?>>? ResolveExternalAsync { get; init; }
    public Func<string, DxpFieldEvalContext, Task<DxpFieldNodeBuffer?>>? ResolveDocVariableNodesAsync { get; init; }
    public Func<string, DxpFieldEvalContext, Task<DxpFieldValue?>>? ResolveDocVariableAsync { get; init; }
    public Func<string, DxpFieldEvalContext, Task<DxpFieldValue?>>? ResolveDocumentPropertyAsync { get; init; }
    public Func<string, DxpFieldEvalContext, Task<DxpFieldValue?>>? ResolveMergeFieldAsync { get; init; }
    public Func<string, DxpFieldEvalContext, Task<DxpFieldValue?>>? AskAsync { get; init; }
}
