using DocxportNet.API;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpIFFieldCachedFrame : DxpFieldCachedFrameBase, DxpIFieldEvalFrame
{

    public DxpIFFieldCachedFrame(DxpIVisitor next)
        : base(next)
    {}
}
