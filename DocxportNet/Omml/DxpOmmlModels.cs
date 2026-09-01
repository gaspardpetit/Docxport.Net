using DocumentFormat.OpenXml;
using System.Xml.Linq;

namespace DocxportNet.Omml;

/// <summary>Output formats supported by the standalone OMML converter.</summary>
public enum DxpOmmlOutputFormat
{
    MathMl,
    Latex,
    UnicodeMath,
    Text,
}

/// <summary>Placement used for limits on n-ary operators.</summary>
public enum DxpOmmlLimitLocation
{
    UnderOver,
    SubscriptSuperscript,
}

/// <summary>Horizontal layout of a display-math paragraph.</summary>
public enum DxpOmmlJustification
{
    Left,
    Right,
    Center,
    CenterGroup,
}

/// <summary>Placement of a binary operator when an equation wraps.</summary>
public enum DxpOmmlBreakBinary
{
    Before,
    After,
    Repeat,
}

/// <summary>Replacement used when subtraction is repeated across a break.</summary>
public enum DxpOmmlBreakBinarySubtraction
{
    MinusMinus,
    MinusPlus,
    PlusMinus,
}

/// <summary>Controls how valid but unsupported OMML structures are rendered.</summary>
public enum DxpOmmlFallbackPolicy
{
    Throw,
    ExtractText,
    Placeholder,
    Omit,
}

public enum DxpOmmlRevisionMode
{
    Accept,
    Reject,
    Preserve,
}

public enum DxpOmmlFieldMode
{
    CachedResult,
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

/// <summary>Describes embedded non-OMML content encountered while converting an equation.</summary>
public sealed record DxpOmmlEmbeddedContentRequest(
    IReadOnlyList<XElement> XmlElements,
    IReadOnlyList<OpenXmlElement> OpenXmlElements,
    string Path,
    string ElementName,
    DxpOmmlOutputFormat OutputFormat,
    DxpOmmlRevisionMode RevisionMode,
    DxpOmmlFieldMode FieldMode,
    bool IncludeHyperlinkTargets);

/// <summary>Resolves embedded WordprocessingML to visible text for an OMML output writer.</summary>
public interface IDxpOmmlEmbeddedContentResolver
{
    string? Resolve(DxpOmmlEmbeddedContentRequest request);
}

/// <summary>Options shared by all standalone OMML conversion methods.</summary>
public sealed class DxpOmmlConversionOptions
{
    /// <summary>Optional resolver for embedded WordprocessingML. The lightweight fallback is used when absent.</summary>
    public IDxpOmmlEmbeddedContentResolver? EmbeddedContentResolver { get; set; }

    /// <summary>Controls visible content selected from embedded revision containers.</summary>
    public DxpOmmlRevisionMode RevisionMode { get; set; } = DxpOmmlRevisionMode.Accept;

    /// <summary>Controls embedded field handling. Field evaluation is intentionally outside this utility.</summary>
    public DxpOmmlFieldMode FieldMode { get; set; } = DxpOmmlFieldMode.CachedResult;

    /// <summary>Appends resolved hyperlink targets to visible hyperlink text.</summary>
    public bool IncludeHyperlinkTargets { get; set; }

    /// <summary>Resolves a hyperlink relationship id and/or anchor without requiring a package.</summary>
    public Func<string?, string?, string?>? HyperlinkTargetResolver { get; set; }

    /// <summary>How valid but unsupported OMML is represented.</summary>
    public DxpOmmlFallbackPolicy FallbackPolicy { get; set; } = DxpOmmlFallbackPolicy.ExtractText;

    /// <summary>Text emitted when <see cref="FallbackPolicy"/> is <see cref="DxpOmmlFallbackPolicy.Placeholder"/>.</summary>
    public string Placeholder { get; set; } = "[unsupported math]";

    /// <summary>Overrides inline/display mode inferred from the OMML root.</summary>
    public bool? Display { get; set; }

    /// <summary>Requests compact inline fractions, corresponding to Word's document-level smallFrac setting.</summary>
    public bool SmallFractions { get; set; }

    /// <summary>Uses display-math defaults for an inline <c>m:oMath</c>, corresponding to <c>m:dispDef</c>.</summary>
    public bool DisplayDefaults { get; set; }

    /// <summary>Document math font hint, corresponding to <c>m:mathFont</c>.</summary>
    public string? MathFont { get; set; }

    /// <summary>Document default for binary operators at line breaks.</summary>
    public DxpOmmlBreakBinary? BreakBinary { get; set; }

    /// <summary>Document subtraction behavior when a binary operator is repeated across a break.</summary>
    public DxpOmmlBreakBinarySubtraction? BreakBinarySubtraction { get; set; }

    /// <summary>Document default math-paragraph justification. A local <c>m:jc</c> takes precedence.</summary>
    public DxpOmmlJustification? DefaultJustification { get; set; }

    /// <summary>Left display-math margin in twentieths of a point.</summary>
    public uint? LeftMarginTwips { get; set; }

    /// <summary>Right display-math margin in twentieths of a point.</summary>
    public uint? RightMarginTwips { get; set; }

    /// <summary>Space before a display equation in twentieths of a point.</summary>
    public uint? PreSpacingTwips { get; set; }

    /// <summary>Space after a display equation in twentieths of a point.</summary>
    public uint? PostSpacingTwips { get; set; }

    /// <summary>Space between equations in a group in twentieths of a point.</summary>
    public uint? InterSpacingTwips { get; set; }

    /// <summary>Space between lines within an equation in twentieths of a point.</summary>
    public uint? IntraSpacingTwips { get; set; }

    /// <summary>Indent applied to wrapped equation lines in twentieths of a point.</summary>
    public uint? WrapIndentTwips { get; set; }

    /// <summary>Aligns wrapped equation lines to the right margin instead of using <see cref="WrapIndentTwips"/>.</summary>
    public bool WrapRight { get; set; }

    /// <summary>Default limit placement for integral operators when local OMML does not specify it.</summary>
    public DxpOmmlLimitLocation IntegralLimitLocation { get; set; } = DxpOmmlLimitLocation.SubscriptSuperscript;

    /// <summary>Default limit placement for non-integral n-ary operators when local OMML does not specify it.</summary>
    public DxpOmmlLimitLocation NaryLimitLocation { get; set; } = DxpOmmlLimitLocation.UnderOver;

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
