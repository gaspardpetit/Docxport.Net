using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpSetFieldCachedFrame : DxpSetFieldFrame
{
    private string? _cachedResultText;

    public DxpSetFieldCachedFrame(DxpFieldEval eval, DxpFieldEvalContext evalContext, ILogger? logger)
        : base(eval, evalContext, logger)
    {
    }

    public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
    {
        if (string.IsNullOrEmpty(text) || InResult)
            return;
        SuppressContent = true;
    }

    public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
    {
        if (string.IsNullOrEmpty(_cachedResultText))
            _cachedResultText = text;
        else
            _cachedResultText += text;
    }

    protected override bool TryGetResultText(DxpIDocumentContext d, out string? text)
    {
        text = _cachedResultText;
        return text != null;
    }
}
