using System.Xml;
using DocxportNet.Core;

namespace DocxportNet.Doc;

/// <summary>Writes the structural hierarchy as diagnostic XML.</summary>
public sealed class DocStructureXmlVisitor : IDocStructureVisitor, IDisposable
{
    private readonly XmlWriter _writer;
    private bool _disposed;

    public DocStructureXmlVisitor(TextWriter output)
    {
        if (output == null) throw new ArgumentNullException(nameof(output));
        _writer = XmlWriter.Create(output, new XmlWriterSettings
        {
            Indent = true,
            OmitXmlDeclaration = false,
            CloseOutput = false
        });
        _writer.WriteStartDocument();
    }

    public IDisposable Enter(DocStructureNode node, int depth)
    {
        if (_disposed) throw new ObjectDisposedException(nameof(DocStructureXmlVisitor));
        _writer.WriteStartElement(XmlConvert.EncodeLocalName(node.Kind));
        _writer.WriteAttributeString("name", Safe(node.Name));
        if (node.StreamName != null) _writer.WriteAttributeString("stream", Safe(node.StreamName));
        if (node.Offset.HasValue) _writer.WriteAttributeString("offset", XmlConvert.ToString(node.Offset.Value));
        if (node.Length.HasValue) _writer.WriteAttributeString("length", XmlConvert.ToString(node.Length.Value));
        foreach (var attribute in node.Attributes)
        {
            _writer.WriteStartElement("attribute");
            _writer.WriteAttributeString("name", Safe(attribute.Key));
            _writer.WriteAttributeString("value", Safe(attribute.Value));
            _writer.WriteEndElement();
        }
        if (node.Kind == "Pcd" && node.Payload is DocTextPieceContent text)
        {
            _writer.WriteStartElement("Text");
            _writer.WriteString(Safe(text.Text));
            _writer.WriteEndElement();
        }
        if (node.Kind == "EnvUnicodeString" && node.Payload is DocEnvelopeText envelopeText)
        {
            _writer.WriteStartElement("Text");
            _writer.WriteString(Safe(envelopeText.Text));
            _writer.WriteEndElement();
        }
        return DxpDisposable.Create(_writer.WriteEndElement);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _writer.WriteEndDocument();
        _writer.Dispose();
    }

    private static string Safe(string value)
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
