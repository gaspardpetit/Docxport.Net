using DocxportNet.API;

namespace DocxportNet.Fields.Frames;

internal class DxpSimpleFieldCachedFrame : DxpFieldCachedFrameBase, DxpIFieldEvalFrame
{
    public DxpSimpleFieldCachedFrame(DxpIVisitor next, string? instructionText = null)
        : base(next, instructionText)
    {}
}
