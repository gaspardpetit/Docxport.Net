using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Office2013.Word;
using DocxportNet.api;

namespace DocxportNet.walker;

public class DxpComments
{
	private readonly Dictionary<string, DxpCommentInfo> _commentsById = new(StringComparer.Ordinal);
	private readonly Dictionary<string, string> _rootIdByCommentId = new(StringComparer.Ordinal);
	private readonly Dictionary<string, DxpCommentThread> _threadsByRootId = new(StringComparer.Ordinal);
	private readonly HashSet<string> _emittedRootIds = new(StringComparer.Ordinal); // avoid duplicate thread emission

	private static string ExtractCommentPlainText(Comment c)
	{
		// Keep it simple: join block InnerText by newlines.
		// This ignores rich formatting inside the comment.
		var lines = new List<string>();

		foreach (var p in c.Descendants<Paragraph>())
		{
			string? t = p.InnerText?.Trim();
			if (!string.IsNullOrEmpty(t))
				lines.Add(t!);
		}

		// Fallback if there are no paragraphs (rare)
		if (lines.Count == 0)
		{
			var t = c.InnerText?.Trim();
			if (!string.IsNullOrEmpty(t))
				lines.Add(t!);
		}

		return string.Join("\n", lines);
	}

	private static IReadOnlyList<OpenXmlElement> CloneCommentBlocks(Comment c)
	{
		return c.ChildElements
			.Select(e => (OpenXmlElement)e.CloneNode(true))
			.ToList();
	}

	internal void Init(MainDocumentPart? mainDocumentPart)
	{
		_commentsById.Clear();
		_rootIdByCommentId.Clear();
		_threadsByRootId.Clear();
		_emittedRootIds.Clear();

		if (mainDocumentPart == null)
			return;

		var commentsPart = mainDocumentPart.WordprocessingCommentsPart;
		var comments = commentsPart?.Comments;
		if (comments == null)
			return;

		// Map from w14:paraId -> w:comment/@w:id for the last paragraph in each comment.
		var commentIdByParaId = new Dictionary<string, string>(StringComparer.Ordinal);
		var infos = new Dictionary<string, DxpCommentInfo>(StringComparer.Ordinal);

		// First pass: core comment info + paraId mapping.
		foreach (var c in comments.Elements<Comment>())
		{
			string? id = c.Id?.Value;
			if (string.IsNullOrEmpty(id))
				continue;

			string text = ExtractCommentPlainText(c);
			string? author = c.Author?.Value;
			string? initials = c.Initials?.Value;

			DateTime? dateUtc = null;
			if (c.DateUtc != null)
			{
				dateUtc = c.DateUtc.Value;
			}
			else if (c.Date != null)
			{
				dateUtc = c.Date.Value.ToUniversalTime();
			}

			// Per spec, CommentEx links to the paraId of the last paragraph in the comment.
			string? paraId = c.Descendants<Paragraph>().LastOrDefault()?.ParagraphId?.Value;
			if (!string.IsNullOrEmpty(paraId))
				commentIdByParaId[paraId!] = id!;

			infos[id!] = new DxpCommentInfo
			{
				Id = id!,
				Text = text,
				Author = author,
				Initials = initials,
				DateUtc = dateUtc,
				IsDone = false,
				IsReply = false,
				ParentId = null,
				Blocks = CloneCommentBlocks(c),
				Part = commentsPart
			};
		}

		// Second pass: Office 2013+ extended info (threading + done flag).
		var commentsExPart = mainDocumentPart.WordprocessingCommentsExPart;
		var commentsEx = commentsExPart?.CommentsEx;
		if (commentsEx != null)
		{
			foreach (var ex in commentsEx.Elements<CommentEx>())
			{
				string? paraId = ex.ParaId?.Value;
				if (string.IsNullOrEmpty(paraId))
					continue;

				if (!commentIdByParaId.TryGetValue(paraId!, out string? commentId))
					continue; // no matching w:comment

				if (!infos.TryGetValue(commentId, out var info))
					continue;

				bool isDone = ex.Done != null && ex.Done.Value;

				string? parentId = null;
				string? parentParaId = ex.ParaIdParent?.Value;
				if (!string.IsNullOrEmpty(parentParaId) && commentIdByParaId.TryGetValue(parentParaId!, out var parentCommentId))
					parentId = parentCommentId;

				info = info with
				{
					IsDone = isDone,
					ParentId = parentId,
					IsReply = parentId != null
				};

				infos[commentId] = info;
			}
		}

		// Finalize maps and build threads.
		foreach (var kvp in infos)
			_commentsById[kvp.Key] = kvp.Value;

		// Compute root id for each comment (follow ParentId chain).
		foreach (var id in _commentsById.Keys)
		{
			_rootIdByCommentId[id] = ResolveRootId(id);
		}

		// Group into threads by root id.
		foreach (var group in _commentsById.Values.GroupBy(info => _rootIdByCommentId[info.Id]))
		{
			var ordered = group
				.OrderBy(info => info.DateUtc ?? DateTime.MinValue)
				.ThenBy(info => info.Id, StringComparer.Ordinal)
				.ToList();

			string rootId = group.Key;
			var thread = new DxpCommentThread
			{
				AnchorCommentId = rootId,
				Comments = ordered
			};

			_threadsByRootId[rootId] = thread;
		}
	}

	private string ResolveRootId(string id)
	{
		string current = id;
		var visited = new HashSet<string>(StringComparer.Ordinal);

		while (true)
		{
			if (!_commentsById.TryGetValue(current, out var info))
				return current; // should not happen, but fail safe

			if (string.IsNullOrEmpty(info.ParentId))
				return current; // root

			if (!visited.Add(current))
				return current; // cycle protection

			current = info.ParentId!;
		}
	}

	public DxpCommentThread? GetThreadForAnchor(string anchorCommentId)
	{
		if (string.IsNullOrEmpty(anchorCommentId))
			return null;

		if (!_commentsById.ContainsKey(anchorCommentId))
			return null;

		if (!_rootIdByCommentId.TryGetValue(anchorCommentId, out var rootId))
			rootId = anchorCommentId;

		if (!_threadsByRootId.TryGetValue(rootId, out var thread))
			return null;

		// Only emit once per root/thread to avoid duplicates when multiple anchors
		// (e.g., rangeStart + reference) share the same comment id.
		if (!_emittedRootIds.Add(rootId))
			return null;

		return thread with { AnchorCommentId = anchorCommentId };
	}
}
