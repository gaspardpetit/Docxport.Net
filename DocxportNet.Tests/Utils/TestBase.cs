using Microsoft.Extensions.Logging;
using Xunit.Abstractions;

internal sealed class XunitLoggerProvider : ILoggerProvider
{
    private readonly ITestOutputHelper _output;

    public XunitLoggerProvider(ITestOutputHelper output)
    {
        _output = output;
    }

    public ILogger CreateLogger(string categoryName)
        => new XunitLogger(_output, categoryName);

    public void Dispose() { }
}

internal sealed class XunitLogger : ILogger
{
    private readonly ITestOutputHelper _output;
    private readonly string _category;

    public XunitLogger(ITestOutputHelper output, string category)
    {
        _output = output;
        _category = category;
    }

    public IDisposable BeginScope<TState>(TState state) where TState : notnull => NullScope.Instance;
    public bool IsEnabled(LogLevel logLevel) => true;

    public void Log<TState>(
        LogLevel logLevel,
        EventId eventId,
        TState state,
        Exception? exception,
        Func<TState, Exception?, string> formatter)
    {
        var msg = formatter(state, exception);
        _output.WriteLine($"{logLevel}: {_category}: {msg}");

        if (exception != null)
            _output.WriteLine(exception.ToString());
    }

    private sealed class NullScope : IDisposable
    {
        public static readonly NullScope Instance = new();
        public void Dispose() { }
    }
}


public abstract class TestBase<T>
{
    protected ILogger<T> Logger { get; }

    protected TestBase(ITestOutputHelper output)
    {
        Logger = LoggerFactory
            .Create(builder => {
                builder.ClearProviders();
                builder.AddProvider(new XunitLoggerProvider(output));
                builder.SetMinimumLevel(LogLevel.Debug);
            })
            .CreateLogger<T>();
    }
}
