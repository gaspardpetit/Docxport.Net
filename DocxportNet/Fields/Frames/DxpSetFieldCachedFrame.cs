using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpSetFieldCachedFrame : DxpFieldCachedFrameBase, DxpIFieldEvalFrame
{
	public DxpSetFieldCachedFrame(DxpFieldEvalContext evalContext, ILogger? logger)
	: base(null)
	{}
}
