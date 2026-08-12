using DocxportNet.API;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Frames;

/// <summary>
/// Replays Word's stored result for fields whose value depends on pagination or
/// section layout, which Docxport deliberately does not compute during export.
/// </summary>
internal sealed class DxpLayoutCachedFieldEvalFrame : DxpValueFieldEvalFrame
{
    public DxpLayoutCachedFieldEvalFrame(
        DxpIVisitor? next,
        DxpFieldEval eval,
        ILogger? logger,
        string? instructionText)
        : base(next, eval, logger, instructionText)
    {
    }

    protected override bool Evaluate(DxpIDocumentContext d)
    {
        if (Next != null && CachedResultBuffer != null)
            CachedResultBuffer.Replay(Next, d);
        return true;
    }
}
