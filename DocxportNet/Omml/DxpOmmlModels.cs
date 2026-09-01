namespace DocxportNet.Omml;

/// <summary>Output formats supported by the standalone OMML converter.</summary>
public enum DxpOmmlOutputFormat
{
    MathMl,
    Latex,
    UnicodeMath,
    Text,
}

/// <summary>Controls how valid but unsupported OMML structures are rendered.</summary>
public enum DxpOmmlFallbackPolicy
{
    Throw,
    ExtractText,
    Placeholder,
    Omit,
}

/// <summary>The severity of an OMML conversion diagnostic.</summary>
public enum DxpOmmlDiagnosticSeverity
{
    Warning,
    Error,
}

/// <summary>Describes a lossy conversion or unsupported OMML structure.</summary>
public sealed record DxpOmmlDiagnostic(
    string Code,
    DxpOmmlDiagnosticSeverity Severity,
    string Message,
    string Path,
    string ElementName);

/// <summary>Options shared by all standalone OMML conversion methods.</summary>
public sealed class DxpOmmlConversionOptions
{
    /// <summary>How valid but unsupported OMML is represented.</summary>
    public DxpOmmlFallbackPolicy FallbackPolicy { get; set; } = DxpOmmlFallbackPolicy.ExtractText;

    /// <summary>Text emitted when <see cref="FallbackPolicy"/> is <see cref="DxpOmmlFallbackPolicy.Placeholder"/>.</summary>
    public string Placeholder { get; set; } = "[unsupported math]";

    /// <summary>Overrides inline/display mode inferred from the OMML root.</summary>
    public bool? Display { get; set; }

    /// <summary>Maximum accepted XML character count.</summary>
    public long MaxInputCharacters { get; set; } = 1_048_576;
}

/// <summary>The output and diagnostics from one standalone OMML conversion.</summary>
public sealed class DxpOmmlConversionResult
{
    internal DxpOmmlConversionResult(
        string output,
        DxpOmmlOutputFormat format,
        bool isDisplay,
        IReadOnlyList<DxpOmmlDiagnostic> diagnostics)
    {
        Output = output;
        Format = format;
        IsDisplay = isDisplay;
        Diagnostics = diagnostics;
    }

    public string Output { get; }
    public DxpOmmlOutputFormat Format { get; }
    public bool IsDisplay { get; }
    public IReadOnlyList<DxpOmmlDiagnostic> Diagnostics { get; }
    public bool IsLossy => Diagnostics.Count != 0;
}

/// <summary>Base exception for standalone OMML conversion failures.</summary>
public class DxpOmmlException : Exception
{
    public DxpOmmlException(string message) : base(message) { }
    public DxpOmmlException(string message, Exception innerException) : base(message, innerException) { }
}

/// <summary>Thrown when input is malformed XML or does not have a supported OMML root.</summary>
public sealed class DxpOmmlParseException : DxpOmmlException
{
    public DxpOmmlParseException(string message) : base(message) { }
    public DxpOmmlParseException(string message, Exception innerException) : base(message, innerException) { }
}

/// <summary>Thrown for valid OMML that is unsupported under the selected fallback policy.</summary>
public sealed class DxpOmmlUnsupportedException : DxpOmmlException
{
    internal DxpOmmlUnsupportedException(DxpOmmlDiagnostic diagnostic)
        : base(diagnostic.Message)
    {
        Diagnostic = diagnostic;
    }

    public DxpOmmlDiagnostic Diagnostic { get; }
}
