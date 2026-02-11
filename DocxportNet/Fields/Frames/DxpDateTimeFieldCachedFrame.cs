using DocxportNet.API;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpDateTimeFieldCachedFrame : DxpSimpleFieldCachedFrame, DxpIFieldEvalFrame
{
    public DxpDateTimeFieldCachedFrame(DxpIVisitor next)
        : base(next)
    {}
}
