using DocxportNet.API;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpDocPropertyFieldCachedFrame : DxpSimpleFieldCachedFrame, DxpIFieldEvalFrame
{
    public DxpDocPropertyFieldCachedFrame(DxpIVisitor next)
        : base(next)
    {}
}
