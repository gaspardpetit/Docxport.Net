using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace l3ia.lapi.services.documents.docx.convert;

public class DxpComments
{
	private readonly Dictionary<string, Comment> _commentsById = new();
	private readonly HashSet<string> _emittedComments = new(); // avoid duplicates

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

	public string? GetComment(string id)
	{
		// Only emit once per comment id (rangeStart + reference often both appear)
		if (string.IsNullOrEmpty(id) || !_emittedComments.Add(id))
			return null;

		if (_commentsById.TryGetValue(id, out Comment? c) == false)
			return null;

		string text = ExtractCommentPlainText(c);
		return text;
	}

	internal void Init(MainDocumentPart? mainDocumentPart)
	{
		var comments = mainDocumentPart?.WordprocessingCommentsPart?.Comments;
		if (comments != null)
		{
			foreach (var c in comments.Elements<Comment>())
			{
				string? id = c.Id?.Value;
				if (id != null)
					_commentsById[id] = c;
			}
		}

	}
}
