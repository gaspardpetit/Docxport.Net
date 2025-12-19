using Microsoft.Extensions.Logging;
using Xunit.Abstractions;

namespace DocxportNet.Tests.Utils;

internal sealed class NullScope : IDisposable
{
	public static readonly NullScope Instance = new();
	private NullScope() { }
	public void Dispose() { }
}

internal sealed class TestLogger : ILogger
{
	private readonly ITestOutputHelper _output;
	private readonly string _categoryName;

	public TestLogger(ITestOutputHelper output, string categoryName)
	{
		_output = output;
		_categoryName = categoryName;
	}

	public IDisposable BeginScope<TState>(TState state) => NullScope.Instance;
	public bool IsEnabled(LogLevel logLevel) => logLevel >= LogLevel.Information;

	public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception? exception, Func<TState, Exception?, string> formatter)
	{
		if (!IsEnabled(logLevel))
			return;

		string message = formatter(state, exception);
		if (exception != null)
			message += Environment.NewLine + exception;

		_output.WriteLine($"[{logLevel}] {_categoryName}: {message}");
	}
}
