using System.Text;


namespace DocxportNet.Visitors;


public sealed class BufferedTextWriter : TextWriter
{
	private readonly StringBuilder _sb = new();
	private readonly object _lock = new();
	private readonly int? _maxChars;

	public BufferedTextWriter(int? maxChars = null, Encoding? encoding = null)
	{
		_maxChars = maxChars;
		Encoding = encoding ?? Encoding.UTF8;
	}

	public override Encoding Encoding { get; }

	public override void Write(char value)
	{
		lock (_lock)
		{
			_sb.Append(value);
			TrimIfNeeded();
		}
	}

	public override void Write(char[] buffer, int index, int count)
	{
		if (buffer is null)
			throw new ArgumentNullException(nameof(buffer));
		lock (_lock)
		{
			_sb.Append(buffer, index, count);
			TrimIfNeeded();
		}
	}

	public override void Write(string? value)
	{
		if (value is null)
			return;
		lock (_lock)
		{
			_sb.Append(value);
			TrimIfNeeded();
		}
	}

	/// <summary>Returns current buffer contents without clearing.</summary>
	public string Peek()
	{
		lock (_lock)
			return _sb.ToString();
	}

	/// <summary>Returns buffer contents and clears it.</summary>
	public string Drain()
	{
		lock (_lock)
		{
			if (_sb.Length == 0)
				return string.Empty;
			var s = _sb.ToString();
			_sb.Clear();
			return s;
		}
	}

	/// <summary>Clears the buffer.</summary>
	public void Clear()
	{
		lock (_lock)
			_sb.Clear();
	}

	public int Length {
		get { lock (_lock) return _sb.Length; }
	}

	private void TrimIfNeeded()
	{
		if (_maxChars is null)
			return;

		int max = _maxChars.Value;
		if (max < 0)
			return;

		if (_sb.Length > max)
		{
			// Drop oldest content, keep the newest 'max' characters.
			_sb.Remove(0, _sb.Length - max);
		}
	}
}
