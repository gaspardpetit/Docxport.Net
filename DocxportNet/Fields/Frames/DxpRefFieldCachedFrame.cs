using DocxportNet.API;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpRefFieldCachedFrame : DxpSimpleFieldCachedFrame, DxpIFieldEvalFrame
{
    public DxpRefFieldCachedFrame(DxpIVisitor next)
        : base(next)
    {}
}
