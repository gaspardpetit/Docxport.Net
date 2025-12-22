namespace DocxportNet.Markdown;

public class DxpMarkdownTableBuilder
{
	private readonly List<Row> _rows = new();
	private Row? _current;

	public void BeginRow(bool isHeader)
	{
		_current = new Row(isHeader);
	}

	public void AddCell(string content)
	{
		_current?.Cells.Add(content);
	}

	public void EndRow()
	{
		if (_current != null)
			_rows.Add(_current);
		_current = null;
	}

	public void Render(TextWriter writer)
	{
		if (_rows.Count == 0)
			return;

		writer.WriteLine();
		int headerIndex = _rows.FindIndex(r => r.IsHeader);
		if (headerIndex < 0)
			headerIndex = 0;

		WriteRow(writer, _rows[headerIndex]);
		WriteSeparator(writer, _rows[headerIndex].Cells.Count);

		for (int i = 0; i < _rows.Count; i++)
		{
			if (i == headerIndex)
				continue;
			WriteRow(writer, _rows[i]);
		}

		writer.WriteLine();
	}

	private static void WriteRow(TextWriter writer, Row row)
	{
		writer.Write("|");
		foreach (var cell in row.Cells)
		{
			var text = NormalizeCell(cell);
			writer.Write(" ");
			writer.Write(text);
			writer.Write(" |");
		}
		writer.WriteLine();
	}

	private static void WriteSeparator(TextWriter writer, int count)
	{
		writer.Write("|");
		for (int i = 0; i < count; i++)
			writer.Write(" --- |");
		writer.WriteLine();
	}

	private static string NormalizeCell(string cell)
	{
		var text = cell.Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ").Trim();
		return text.Replace("|", "\\|");
	}

	private sealed record Row(bool IsHeader)
	{
		public List<string> Cells { get; } = new();
	}
}
