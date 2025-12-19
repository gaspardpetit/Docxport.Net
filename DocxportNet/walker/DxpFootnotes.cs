using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet.walker;

internal class DxpFootnotes
{
	private readonly Dictionary<long, (int, Footnote)> _footnotesById = new();
	private readonly HashSet<long> _emittedFootnotes = new(); // avoid duplicates
	private readonly Stack<long> _footnoteIdStack = new();

	public IEnumerable<(Footnote, int, long)> GetFootnotes()
	{
		foreach (long id in _footnotesById.Keys)
		{
			try
			{
				(int, Footnote) fn = _footnotesById[id];
				int index = _footnotesById[id].Item1;
				yield return (fn.Item2, index, id);
			}
			finally
			{
				//_footnoteIdStack.Pop();
			}
		}
	}

	internal void Init(MainDocumentPart mainPart)
	{
		var footnotes = mainPart?.FootnotesPart?.Footnotes;
		if (footnotes != null)
		{
			int index = 0;
			foreach (var fn in footnotes.Elements<Footnote>())
			{
				var id = fn.Id?.Value;
				if (id == null)
					continue;

				// Skip Word's internal separator footnotes (commonly -1 and 0)
				var type = fn.Type?.Value;
				if (type == FootnoteEndnoteValues.Separator ||
					type == FootnoteEndnoteValues.ContinuationSeparator ||
					type == FootnoteEndnoteValues.ContinuationNotice
					)
					continue;

				_footnotesById[id.Value] = (++index, fn);
			}
		}

	}

	internal bool Resolve(long id, out int index)
	{
		if (_footnotesById.TryGetValue(id, out var fn))
		{
			index = fn.Item1;
			_emittedFootnotes.Add(id);
			_footnoteIdStack.Push(id);
			return true;
		}
		index = 0;
		return false;
	}
}


internal class DocxEndnotes
{
	private readonly Dictionary<long, (int, Endnote)> _endnotesById = new();
	private readonly HashSet<long> _emittedEndnotes = new(); // avoid duplicates
	private readonly Stack<long> _endnoteIdStack = new();

	public IEnumerable<(Endnote, int, long)> GetEndnotes()
	{
		foreach (long id in _endnotesById.Keys)
		{
			try
			{
				(int, Endnote) fn = _endnotesById[id];
				int index = _endnotesById[id].Item1;
				yield return (fn.Item2, index, id);
			}
			finally
			{
				//_endnoteIdStack.Pop();
			}
		}
	}

	internal void Init(MainDocumentPart mainPart)
	{
		var endnotes = mainPart?.EndnotesPart?.Endnotes;
		if (endnotes != null)
		{
			int index = 0;
			foreach (Endnote fn in endnotes.Elements<Endnote>())
			{
				var id = fn.Id?.Value;
				if (id == null)
					continue;

				// Skip Word's internal separator endnotes (commonly -1 and 0)
				FootnoteEndnoteValues? type = fn.Type?.Value;
				if (type == FootnoteEndnoteValues.Separator ||
					type == FootnoteEndnoteValues.ContinuationSeparator ||
					type == FootnoteEndnoteValues.ContinuationNotice
					)
					continue;

				_endnotesById[id.Value] = (++index, fn);
			}
		}

	}

	internal bool Resolve(long id, out int index)
	{
		if (_endnotesById.TryGetValue(id, out var fn))
		{
			index = fn.Item1;
			_emittedEndnotes.Add(id);
			_endnoteIdStack.Push(id);
			return true;
		}
		index = 0;
		return false;
	}
}
