using DocxportNet.Fields.Resolution;

namespace DocxportNet.Fields.Eval;

public sealed class DxpEvaluateFieldMiddlewareOptions
{
    public IDxpRefResolver? RefResolver { get; set; }
    public bool PreserveLayoutDependentFields { get; set; }
    public bool EmitStructuredDatabaseResults { get; set; }
}
