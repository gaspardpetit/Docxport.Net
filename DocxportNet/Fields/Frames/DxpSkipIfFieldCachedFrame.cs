using DocxportNet.API;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpSkipIfFieldCachedFrame : DxpSimpleFieldCachedFrame, DxpIFieldEvalFrame
{
    public DxpSkipIfFieldCachedFrame(DxpIVisitor next)
        : base(next)
    {}
}
