using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;

namespace DocxportNet.Omml;

internal static class OmmlParser
{
    internal const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public static OmmlDocument Parse(string? omml, DxpOmmlConversionOptions options)
    {
        if (omml is null)
            throw new DxpOmmlParseException("OMML input cannot be null.");
        if (string.IsNullOrWhiteSpace(omml))
            throw new DxpOmmlParseException("OMML input cannot be empty or whitespace.");
        if (options.MaxInputCharacters <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxInputCharacters), "The input limit must be positive.");
        if (omml.Length > options.MaxInputCharacters)
            throw new DxpOmmlParseException(
                $"OMML input exceeds the {options.MaxInputCharacters.ToString(CultureInfo.InvariantCulture)} character limit.");

        XDocument xml;
        try
        {
            using StringReader source = new(omml);
            using XmlReader reader = XmlReader.Create(source, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = options.MaxInputCharacters,
                IgnoreComments = true,
            });
            xml = XDocument.Load(reader, LoadOptions.PreserveWhitespace | LoadOptions.SetLineInfo);
        }
        catch (XmlException exception)
        {
            throw new DxpOmmlParseException("OMML input is not secure, well-formed XML.", exception);
        }

        XElement? root = xml.Root;
        if (root is null)
            throw new DxpOmmlParseException("OMML input has no root element.");
        if (root.Name.NamespaceName != MathNamespace ||
            (root.Name.LocalName != "oMath" && root.Name.LocalName != "oMathPara"))
        {
            throw new DxpOmmlParseException(
                $"Expected an OMML oMath or oMathPara root, but found '{root.Name}'.");
        }

        bool isDisplay = root.Name.LocalName == "oMathPara";
        IReadOnlyList<OmmlNode> children = isDisplay
            ? ParseParagraph(root)
            : ParseChildren(root, "/m:oMath[1]");
        return new OmmlDocument(isDisplay, children);
    }

    public static OmmlDocument Parse(OpenXmlElement root)
    {
        bool isDisplay = root is DocumentFormat.OpenXml.Math.Paragraph;
        string rootPath = isDisplay ? "/m:oMathPara[1]" : "/m:oMath[1]";
        IReadOnlyList<OmmlNode> children = isDisplay
            ? ParseParagraph(root)
            : ParseChildren(root, rootPath);
        return new OmmlDocument(isDisplay, children);
    }

    private static IReadOnlyList<OmmlNode> ParseParagraph(XElement paragraph)
    {
        List<OmmlNode> children = new();
        int mathIndex = 0;
        Dictionary<XName, int> indexes = new();
        int textIndex = 0;
        foreach (XNode node in paragraph.Nodes())
        {
            if (node is XText text)
            {
                if (!string.IsNullOrWhiteSpace(text.Value))
                    children.Add(new OmmlUnsupported($"/m:oMathPara[1]/#text[{Index(++textIndex)}]", "#text", text.Value));
                continue;
            }
            if (node is not XElement child)
                continue;

            if (child.Name == XName.Get("oMath", MathNamespace))
            {
                mathIndex++;
                string path = $"/m:oMathPara[1]/m:oMath[{Index(mathIndex)}]";
                children.Add(new OmmlSequence(path, ParseChildren(child, path)));
            }
            else
            {
                indexes.TryGetValue(child.Name, out int index);
                indexes[child.Name] = ++index;
                string path = $"/m:oMathPara[1]/{QualifiedName(child)}[{Index(index)}]";
                children.Add(ParseUnsupported(child, path));
            }
        }

        if (mathIndex == 0)
            throw new DxpOmmlParseException("An OMML oMathPara root must contain at least one oMath child.");
        return children;
    }

    private static IReadOnlyList<OmmlNode> ParseParagraph(OpenXmlElement paragraph)
    {
        List<OmmlNode> children = new();
        int mathIndex = 0;
        Dictionary<(string Namespace, string LocalName), int> indexes = new();
        foreach (OpenXmlElement child in paragraph.ChildElements)
        {
            if (child is DocumentFormat.OpenXml.Math.OfficeMath)
            {
                mathIndex++;
                string path = $"/m:oMathPara[1]/m:oMath[{Index(mathIndex)}]";
                children.Add(new OmmlSequence(path, ParseChildren(child, path)));
            }
            else
            {
                (string Namespace, string LocalName) name = (child.NamespaceUri, child.LocalName);
                indexes.TryGetValue(name, out int index);
                indexes[name] = ++index;
                string path = $"/m:oMathPara[1]/{QualifiedName(child)}[{Index(index)}]";
                children.Add(ParseUnsupported(child, path));
            }
        }

        if (mathIndex == 0)
            throw new DxpOmmlParseException("An OMML oMathPara root must contain at least one oMath child.");
        return children;
    }

    private static IReadOnlyList<OmmlNode> ParseChildren(XElement parent, string parentPath)
    {
        List<OmmlNode> children = new();
        Dictionary<XName, int> indexes = new();
        int textIndex = 0;
        foreach (XNode node in parent.Nodes())
        {
            if (node is XText text)
            {
                if (!string.IsNullOrWhiteSpace(text.Value))
                    children.Add(new OmmlUnsupported($"{parentPath}/#text[{Index(++textIndex)}]", "#text", text.Value));
                continue;
            }
            if (node is not XElement child)
                continue;

            indexes.TryGetValue(child.Name, out int index);
            indexes[child.Name] = ++index;
            children.Add(ParseUnsupported(child, $"{parentPath}/{QualifiedName(child)}[{Index(index)}]"));
        }

        return children;
    }

    private static IReadOnlyList<OmmlNode> ParseChildren(OpenXmlElement parent, string parentPath)
    {
        List<OmmlNode> children = new();
        Dictionary<(string Namespace, string LocalName), int> indexes = new();
        foreach (OpenXmlElement child in parent.ChildElements)
        {
            (string Namespace, string LocalName) name = (child.NamespaceUri, child.LocalName);
            indexes.TryGetValue(name, out int index);
            indexes[name] = ++index;
            children.Add(ParseUnsupported(child, $"{parentPath}/{QualifiedName(child)}[{Index(index)}]"));
        }

        return children;
    }

    private static OmmlUnsupported ParseUnsupported(XElement element, string path) =>
        new(path, QualifiedName(element), ExtractVisibleText(element));

    private static OmmlUnsupported ParseUnsupported(OpenXmlElement element, string path) =>
        new(path, QualifiedName(element), ExtractVisibleText(element));

    private static string ExtractVisibleText(XElement element)
    {
        StringWriter result = new();
        foreach (XElement descendant in element.DescendantsAndSelf())
        {
            if ((descendant.Name.NamespaceName == MathNamespace || descendant.Name.NamespaceName == WordNamespace) &&
                descendant.Name.LocalName == "t")
            {
                result.Write(descendant.Value);
            }
            else if (descendant.Name.NamespaceName == WordNamespace && descendant.Name.LocalName == "tab")
            {
                result.Write('\t');
            }
            else if (descendant.Name.NamespaceName == WordNamespace && descendant.Name.LocalName == "br")
            {
                result.Write('\n');
            }
        }

        return result.ToString();
    }

    private static string ExtractVisibleText(OpenXmlElement element)
    {
        StringWriter result = new();
        WriteVisibleText(element, result);
        return result.ToString();
    }

    private static void WriteVisibleText(OpenXmlElement element, TextWriter result)
    {
        if ((element.NamespaceUri == MathNamespace || element.NamespaceUri == WordNamespace) &&
            element.LocalName == "t")
        {
            result.Write(element.InnerText);
            return;
        }
        if (element.NamespaceUri == WordNamespace && element.LocalName == "tab")
        {
            result.Write('\t');
            return;
        }
        if (element.NamespaceUri == WordNamespace && element.LocalName == "br")
        {
            result.Write('\n');
            return;
        }

        foreach (OpenXmlElement child in element.ChildElements)
            WriteVisibleText(child, result);
    }

    private static string QualifiedName(XElement element)
    {
        string prefix = element.Name.NamespaceName switch
        {
            MathNamespace => "m",
            WordNamespace => "w",
            _ => "ns",
        };
        return $"{prefix}:{element.Name.LocalName}";
    }

    private static string QualifiedName(OpenXmlElement element)
    {
        string prefix = element.NamespaceUri switch
        {
            MathNamespace => "m",
            WordNamespace => "w",
            _ => "ns",
        };
        return $"{prefix}:{element.LocalName}";
    }

    private static string Index(int value) => value.ToString(CultureInfo.InvariantCulture);
}
