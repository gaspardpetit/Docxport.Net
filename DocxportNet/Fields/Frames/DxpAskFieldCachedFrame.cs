using DocxportNet.API;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpAskFieldCachedFrame : DxpSimpleFieldCachedFrame, DxpIFieldEvalFrame
{
    public DxpAskFieldCachedFrame(DxpIVisitor next)
        : base(next)
    {}
}
