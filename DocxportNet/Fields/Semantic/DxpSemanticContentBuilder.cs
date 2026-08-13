namespace DocxportNet.Fields.Semantic;

internal sealed class DxpSemanticContentBuilder
{
    private readonly List<DxpSemanticNode> _nodes = new();

    public void Append(DxpSemanticContent content) => _nodes.AddRange(content.Nodes);
    public void Append(DxpSemanticNode node) => _nodes.Add(node);

    public void AppendTextWithControls(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return;

        int start = 0;
        for (int index = 0; index < text.Length; index++)
        {
            char ch = text[index];
            if (ch is not ('\r' or '\n' or '\t'))
                continue;
            if (index > start)
                _nodes.Add(new DxpSemanticText(text.Substring(start, index - start)));
            if (ch == '\t')
                _nodes.Add(new DxpSemanticTab());
            else
            {
                _nodes.Add(new DxpSemanticBreak());
                if (ch == '\r' && index + 1 < text.Length && text[index + 1] == '\n')
                    index++;
            }
            start = index + 1;
        }
        if (start < text.Length)
            _nodes.Add(new DxpSemanticText(text.Substring(start)));
    }

    public DxpSemanticContent Build() => new(_nodes.ToArray());
}
