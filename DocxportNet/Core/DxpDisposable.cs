namespace DocxportNet.Core;

public static class DxpDisposable
{
	public static IDisposable Empty { get; } = new EmptyDisposable();
	private sealed class EmptyDisposable : IDisposable
	{
		public void Dispose() { /* no-op */ }
	}

	public static IDisposable Create(Action onDispose)
	{
		if (onDispose is null)
			throw new ArgumentNullException(nameof(onDispose));
		return new AnonymousDisposable(onDispose);
	}

	private sealed class AnonymousDisposable : IDisposable
	{
		private Action? _onDispose;

		public AnonymousDisposable(Action onDispose) => _onDispose = onDispose;

		public void Dispose()
		{
			// Ensure it runs only once, even if Dispose is called multiple times.
			var action = Interlocked.Exchange(ref _onDispose, null);
			action?.Invoke();
		}
	}
}

