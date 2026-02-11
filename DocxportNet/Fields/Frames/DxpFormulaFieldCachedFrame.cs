using DocxportNet.API;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpFormulaFieldCachedFrame : DxpSimpleFieldCachedFrame, DxpIFieldEvalFrame
{
    public DxpFormulaFieldCachedFrame(DxpIVisitor next)
        : base(next)
    {}
}
