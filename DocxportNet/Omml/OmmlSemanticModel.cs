namespace DocxportNet.Omml;

internal sealed class OmmlDocument
{
    public OmmlDocument(bool isDisplay, IReadOnlyList<OmmlNode> children)
    {
        IsDisplay = isDisplay;
        Children = children;
    }

    public bool IsDisplay { get; }
    public IReadOnlyList<OmmlNode> Children { get; }
}

internal abstract class OmmlNode
{
    protected OmmlNode(string path) => Path = path;
    public string Path { get; }
}

internal sealed class OmmlSequence : OmmlNode
{
    public OmmlSequence(string path, IReadOnlyList<OmmlNode> children) : base(path) => Children = children;
    public IReadOnlyList<OmmlNode> Children { get; }
}

internal enum OmmlTokenKind { Identifier, Number, Operator, Text }
internal enum OmmlMathScript { Default, Roman, Script, Fraktur, DoubleStruck, SansSerif, Monospace }
internal enum OmmlMathStyle { Default, Plain, Bold, Italic, BoldItalic }

internal sealed class OmmlToken
{
    public OmmlToken(OmmlTokenKind kind, string value) { Kind = kind; Value = value; }
    public OmmlTokenKind Kind { get; }
    public string Value { get; }
}

internal sealed class OmmlRun : OmmlNode
{
    public OmmlRun(string path, IReadOnlyList<OmmlToken> tokens, OmmlMathScript script,
        OmmlMathStyle style, bool literal, bool normal, bool alignment, string? language, bool rightToLeft)
        : base(path)
    {
        Tokens = tokens; Script = script; Style = style; Literal = literal; Normal = normal;
        Alignment = alignment; Language = language; RightToLeft = rightToLeft;
    }
    public IReadOnlyList<OmmlToken> Tokens { get; }
    public OmmlMathScript Script { get; }
    public OmmlMathStyle Style { get; }
    public bool Literal { get; }
    public bool Normal { get; }
    public bool Alignment { get; }
    public string? Language { get; }
    public bool RightToLeft { get; }
}

internal enum OmmlFractionType { Bar, Skewed, Linear, NoBar }
internal enum OmmlScriptType { Subscript, Superscript, SubSup, PreSubSup }

internal sealed class OmmlFraction : OmmlNode
{
    public OmmlFraction(string path, OmmlFractionType type, OmmlSequence numerator,
        OmmlSequence denominator, bool hasControlProperties) : base(path)
    { Type = type; Numerator = numerator; Denominator = denominator; HasControlProperties = hasControlProperties; }
    public OmmlFractionType Type { get; }
    public OmmlSequence Numerator { get; }
    public OmmlSequence Denominator { get; }
    public bool HasControlProperties { get; }
}

internal sealed class OmmlRadical : OmmlNode
{
    public OmmlRadical(string path, OmmlSequence radicand, OmmlSequence degree,
        bool hasDegree, bool degreeHidden, bool hasControlProperties) : base(path)
    { Radicand = radicand; Degree = degree; HasDegree = hasDegree; DegreeHidden = degreeHidden; HasControlProperties = hasControlProperties; }
    public OmmlSequence Radicand { get; }
    public OmmlSequence Degree { get; }
    public bool HasDegree { get; }
    public bool DegreeHidden { get; }
    public bool HasControlProperties { get; }
}

internal sealed class OmmlScript : OmmlNode
{
    public OmmlScript(string path, OmmlScriptType type, OmmlSequence @base,
        OmmlSequence subscript, OmmlSequence superscript, bool alignScripts,
        bool hasControlProperties) : base(path)
    { Type = type; Base = @base; Subscript = subscript; Superscript = superscript; AlignScripts = alignScripts; HasControlProperties = hasControlProperties; }
    public OmmlScriptType Type { get; }
    public OmmlSequence Base { get; }
    public OmmlSequence Subscript { get; }
    public OmmlSequence Superscript { get; }
    public bool AlignScripts { get; }
    public bool HasControlProperties { get; }
}

internal sealed class OmmlUnsupported : OmmlNode
{
    public OmmlUnsupported(string path, string elementName, string visibleText) : base(path)
    {
        ElementName = elementName;
        VisibleText = visibleText;
    }

    public string ElementName { get; }
    public string VisibleText { get; }
}
