using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace DocxportNet.API;

public sealed record DxpCommentInfo
{
	public string Id { get; init; } = string.Empty;
	public string Text { get; init; } = string.Empty;
	public string? Author { get; init; }
	public string? Initials { get; init; }
	public DateTime? DateUtc { get; init; }
	public bool IsDone { get; init; }
	public bool IsReply { get; init; }
	public string? ParentId { get; init; }
	public IReadOnlyList<OpenXmlElement> Blocks { get; init; } = Array.Empty<OpenXmlElement>();
	public OpenXmlPart? Part { get; init; }
}

public sealed record DxpCommentThread
{
	/// <summary>
	/// The comment id where this thread was anchored in the document when emitted.
	/// This may be the root comment id or a reply id, depending on where the walker encountered it.
	/// </summary>
	public string AnchorCommentId { get; init; } = string.Empty;

	public IReadOnlyList<DxpCommentInfo> Comments { get; init; } = Array.Empty<DxpCommentInfo>();
}
