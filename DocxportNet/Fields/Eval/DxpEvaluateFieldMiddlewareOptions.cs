using DocxportNet.Fields.Resolution;

namespace DocxportNet.Fields.Eval;

public sealed class DxpEvaluateFieldMiddlewareOptions
{
    public IDxpRefResolver? RefResolver { get; set; }
}
