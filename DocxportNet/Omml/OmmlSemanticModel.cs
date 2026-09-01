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
