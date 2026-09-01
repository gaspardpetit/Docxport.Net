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
