using DocxportNet.API;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpMergeFieldCachedFrame : DxpSimpleFieldCachedFrame, DxpIFieldEvalFrame
{
    public DxpMergeFieldCachedFrame(DxpIVisitor next)
        : base(next)
    {}
}
