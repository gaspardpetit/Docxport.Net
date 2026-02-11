using DocxportNet.API;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpDocVariableFieldCachedFrame : DxpFieldCachedFrameBase, DxpIFieldEvalFrame
{
    public DxpDocVariableFieldCachedFrame(DxpIVisitor next)
        : base(next)
    {}
}
