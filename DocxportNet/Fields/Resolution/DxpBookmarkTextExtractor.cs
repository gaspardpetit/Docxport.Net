using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Visitors;
using DocxportNet.Walker;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Resolution;

internal static class DxpBookmarkTextExtractor
{
	public static IReadOnlyDictionary<string, string> Extract(WordprocessingDocument document, ILogger? logger = null)
	{
		var visitor = new DxpBookmarkTextVisitor(logger);
		var walker = new DxpWalker(logger);
		walker.Accept(document, visitor);
		return visitor.Results;
	}

	private sealed class DxpBookmarkTextVisitor : DxpVisitor
	{
		private readonly Dictionary<string, StringBuilder> _builders = new(StringComparer.OrdinalIgnoreCase);
		private readonly Dictionary<string, string> _idToName = new();
		private readonly List<string> _activeIds = new();

		public DxpBookmarkTextVisitor(ILogger? logger) : base(logger)
		{
		}

		public IReadOnlyDictionary<string, string> Results
		{
			get
			{
				var results = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
				foreach (var kvp in _builders)
					results[kvp.Key] = kvp.Value.ToString();
				return results;
			}
		}

		public override void VisitBookmarkStart(BookmarkStart bs, DxpIDocumentContext d)
		{
			string? name = bs.Name?.Value;
			string? id = bs.Id?.Value;
			if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(id))
				return;

			var bookmarkName = name!;
			var bookmarkId = id!;
			_idToName[bookmarkId] = bookmarkName;
			_activeIds.Add(bookmarkId);
			if (!_builders.ContainsKey(bookmarkName))
				_builders[bookmarkName] = new StringBuilder();
		}

		public override void VisitBookmarkEnd(BookmarkEnd be, DxpIDocumentContext d)
		{
			string? id = be.Id?.Value;
			if (string.IsNullOrWhiteSpace(id))
				return;

			var bookmarkId = id!;
			int idx = _activeIds.LastIndexOf(bookmarkId);
			if (idx >= 0)
				_activeIds.RemoveAt(idx);
		}

		public override void VisitText(Text t, DxpIDocumentContext d) => Append(t.Text);
		public override void VisitDeletedText(DeletedText dt, DxpIDocumentContext d) => Append(dt.Text);
		public override void VisitTab(TabChar tab, DxpIDocumentContext d) => Append("\t");
		public override void VisitBreak(Break br, DxpIDocumentContext d) => Append("\n");
		public override void VisitCarriageReturn(CarriageReturn cr, DxpIDocumentContext d) => Append("\n");
		public override void VisitNoBreakHyphen(NoBreakHyphen h, DxpIDocumentContext d) => Append("-");

		private void Append(string text)
		{
			if (_activeIds.Count == 0)
				return;

			foreach (var id in _activeIds)
			{
				if (_idToName.TryGetValue(id, out var name) && _builders.TryGetValue(name, out var sb))
					sb.Append(text);
			}
		}
	}
}
