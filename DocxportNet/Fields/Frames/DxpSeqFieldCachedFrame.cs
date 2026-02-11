using DocxportNet.API;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpSeqFieldCachedFrame : DxpSimpleFieldCachedFrame, DxpIFieldEvalFrame
{
    public DxpSeqFieldCachedFrame(DxpIVisitor next)
        : base(next)
    {}
}
