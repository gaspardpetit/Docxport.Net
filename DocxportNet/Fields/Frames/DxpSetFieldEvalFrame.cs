using DocxportNet.API;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpSetFieldEvalFrame : DxpSetFieldFrame
{
    public DxpSetFieldEvalFrame(DxpFieldEval eval, DxpFieldEvalContext evalContext, ILogger? logger)
        : base(eval, evalContext, logger)
    {
    }

    protected override bool TryGetResultText(DxpIDocumentContext d, out string? text)
    {
        text = null;
        return false;
    }
}
