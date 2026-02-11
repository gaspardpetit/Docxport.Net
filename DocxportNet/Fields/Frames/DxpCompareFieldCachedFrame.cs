using DocxportNet.API;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpCompareFieldCachedFrame : DxpSimpleFieldCachedFrame, DxpIFieldEvalFrame
{
    public DxpCompareFieldCachedFrame(DxpIVisitor next)
        : base(next)
    {}
}
