using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Fields;
using DocxportNet.Fields.Eval;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpMergeFieldEvalFrame : DxpValueFieldEvalFrame
{
    public DxpMergeFieldEvalFrame(DxpIVisitor? next, DxpFieldEval eval, ILogger? logger, string? instructionText, Run? codeRun = null)
        : base(next, eval, logger, instructionText, codeRun)
    {}
}
