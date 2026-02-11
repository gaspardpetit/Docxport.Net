using DocxportNet.API;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpRefFieldCachedFrame : DxpFieldCachedFrameBase, DxpIFieldEvalFrame
{
	public DxpRefFieldCachedFrame(DxpIVisitor next)
		: base(next)
	{}

}
