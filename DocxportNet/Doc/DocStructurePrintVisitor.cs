namespace DocxportNet.Doc;

/// <summary>Prints each recognized structure and its attributes.</summary>
public sealed class DocStructurePrintVisitor : IDocStructureVisitor
{
    private readonly TextWriter _writer;

    public DocStructurePrintVisitor(TextWriter writer) => _writer = writer ?? throw new ArgumentNullException(nameof(writer));

    public IDisposable Enter(DocStructureNode node, int depth)
    {
        _writer.Write(new string(' ', depth * 2));
        _writer.Write(node.Kind);
        _writer.Write(' ');
        _writer.Write(Escape(node.Name));
        if (node.StreamName != null) _writer.Write($" stream={Escape(node.StreamName)}");
        if (node.Offset.HasValue) _writer.Write($" offset={node.Offset.Value}");
        if (node.Length.HasValue) _writer.Write($" length={node.Length.Value}");
        foreach (var attribute in node.Attributes)
            _writer.Write($" {attribute.Key}={Escape(attribute.Value)}");
        _writer.WriteLine();
        return DocxportNet.Core.DxpDisposable.Empty;
    }

    private static string Escape(string value)
    {
        var output = new System.Text.StringBuilder(value.Length);
        foreach (var character in value)
        {
            if (character == '\\') output.Append("\\\\");
            else if (character < ' ' || character > '~')
                output.Append("\\u").Append(((int)character).ToString("X4", System.Globalization.CultureInfo.InvariantCulture));
            else output.Append(character);
        }
        return output.ToString();
    }
}
