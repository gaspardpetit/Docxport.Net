using DocxportNet.API;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpDocVariableFieldCachedFrame : DxpSimpleFieldCachedFrame, DxpIFieldEvalFrame
{
    public DxpDocVariableFieldCachedFrame(DxpIVisitor next)
        : base(next)
    {}
}
