using System.Text;

namespace DocxportNet.Core;

public sealed class DxpMultiTextWriter : TextWriter
{
	private readonly TextWriter[] _writers;
	private readonly bool _leaveOpen;
	private readonly object _lock = new();

	public DxpMultiTextWriter(bool leaveOpen, params TextWriter[] writers)
	{
		_leaveOpen = leaveOpen;
		_writers = Normalize(writers);
	}

	public DxpMultiTextWriter(params TextWriter[] writers) : this(leaveOpen: false, writers) { }

	private static TextWriter[] Normalize(TextWriter[] writers)
	{
		if (writers is null)
			throw new ArgumentNullException(nameof(writers));

		var normalized = writers.Where(w => w != null).ToArray();
		if (normalized.Length == 0)
			throw new ArgumentException("At least one non-null TextWriter is required.", nameof(writers));

		return normalized;
	}

	public override Encoding Encoding => _writers[0].Encoding;

	public override void Flush()
	{
		lock (_lock)
		{
			foreach (var w in _writers)
				w.Flush();
		}
	}

	public override void Write(char value)
	{
		lock (_lock)
		{
			foreach (var w in _writers)
				w.Write(value);
		}
	}

	public override void Write(char[] buffer, int index, int count)
	{
		lock (_lock)
		{
			foreach (var w in _writers)
				w.Write(buffer, index, count);
		}
	}

	public override void Write(string? value)
	{
		lock (_lock)
		{
			foreach (var w in _writers)
				w.Write(value);
		}
	}

	public override Task WriteAsync(char value)
	{
		lock (_lock)
		{
			return Task.WhenAll(_writers.Select(w => w.WriteAsync(value)));
		}
	}

	public override Task WriteAsync(char[] buffer, int index, int count)
	{
		lock (_lock)
		{
			return Task.WhenAll(_writers.Select(w => w.WriteAsync(buffer, index, count)));
		}
	}

	public override Task WriteAsync(string? value)
	{
		lock (_lock)
		{
			return Task.WhenAll(_writers.Select(w => w.WriteAsync(value)));
		}
	}

	public override Task FlushAsync()
	{
		lock (_lock)
		{
			return Task.WhenAll(_writers.Select(w => w.FlushAsync()));
		}
	}

	protected override void Dispose(bool disposing)
	{
		if (disposing)
		{
			lock (_lock)
			{
				foreach (var w in _writers)
					w.Flush();

				if (!_leaveOpen)
				{
					foreach (var w in _writers)
						w.Dispose();
				}
			}
		}
		base.Dispose(disposing);
	}

#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
	public override async ValueTask DisposeAsync()
	{
		// Note: not locking across awaits. If you need strict async serialization,
		// swap _lock for a SemaphoreSlim and use WaitAsync/Release.
		Flush();
		if (!_leaveOpen)
		{
			await Task.WhenAll(_writers.Select(w => w.DisposeAsync().AsTask()))
					  .ConfigureAwait(false);
		}
		await base.DisposeAsync().ConfigureAwait(false);
	}
#endif
}
