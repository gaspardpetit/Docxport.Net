using System.Globalization;
using System.Text;
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
            string path = $"{parentPath}/{QualifiedName(child)}[{Index(index)}]";
            children.Add(child.Name == XName.Get("r", MathNamespace) ? ParseRun(child, path) : ParseUnsupported(child, path));
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
            string path = $"{parentPath}/{QualifiedName(child)}[{Index(index)}]";
            children.Add(child.NamespaceUri == MathNamespace && child.LocalName == "r" ? ParseRun(child, path) : ParseUnsupported(child, path));
        }

        return children;
    }

    private static OmmlRun ParseRun(XElement run, string path) => ParseRunCore(
        path,
        new[] { ExtractRunText(run) },
        run.Descendants(), e => e.Name.NamespaceName, e => e.Name.LocalName,
        e => (string?)e.Attribute(XName.Get("val", e.Name.NamespaceName)) ?? (string?)e.Attribute(XName.Get("val", WordNamespace)));

    private static OmmlRun ParseRun(OpenXmlElement run, string path)
    {
        IEnumerable<OpenXmlElement> all = run.Descendants();
        return ParseRunCore(path,
            new[] { ExtractRunText(run) },
            all, e => e.NamespaceUri, e => e.LocalName, e => Attribute(e, "val"));
    }

    private static OmmlRun ParseRunCore<T>(string path, IEnumerable<string> texts,
        IEnumerable<T> elements, Func<T, string> ns, Func<T, string> local, Func<T, string?> val)
        where T : class
    {
        List<T> all = elements.ToList();
        bool Has(string space, string name) => all.Any(e => ns(e) == space && local(e) == name && Enabled(val(e)));
        string? Value(string space, string name)
        {
            T? found = all.FirstOrDefault(e => ns(e) == space && local(e) == name);
            return found == null ? null : val(found);
        }
        bool literal = Has(MathNamespace, "lit"), normal = Has(MathNamespace, "nor");
        OmmlMathScript script = Value(MathNamespace, "scr") switch { "roman" => OmmlMathScript.Roman, "script" => OmmlMathScript.Script, "fraktur" => OmmlMathScript.Fraktur, "double-struck" => OmmlMathScript.DoubleStruck, "sans-serif" => OmmlMathScript.SansSerif, "monospace" => OmmlMathScript.Monospace, _ => OmmlMathScript.Default };
        bool wordBold = Has(WordNamespace, "b"), wordItalic = Has(WordNamespace, "i");
        OmmlMathStyle style = Value(MathNamespace, "sty") switch
        {
            "p" => OmmlMathStyle.Plain,
            "b" => OmmlMathStyle.Bold,
            "i" => OmmlMathStyle.Italic,
            "bi" => OmmlMathStyle.BoldItalic,
            _ when wordBold && wordItalic => OmmlMathStyle.BoldItalic,
            _ when wordBold => OmmlMathStyle.Bold,
            _ when wordItalic => OmmlMathStyle.Italic,
            _ => OmmlMathStyle.Default,
        };
        string text = string.Concat(texts);
        return new OmmlRun(path, OmmlTokenClassifier.Classify(text, literal || normal), script, style, literal, normal,
            Has(MathNamespace, "aln"), Value(WordNamespace, "lang"), Has(WordNamespace, "rtl"));
    }

    private static bool Enabled(string? value) => value == null ||
        !(value == "0" || value.Equals("false", StringComparison.OrdinalIgnoreCase) ||
          value.Equals("off", StringComparison.OrdinalIgnoreCase));

    private static string? Attribute(OpenXmlElement element, string localName) =>
        element.GetAttributes().FirstOrDefault(a => a.LocalName == localName).Value;

    private static string ExtractRunText(XElement run)
    {
        StringBuilder result = new();
        XElement? fonts = run.Descendants().FirstOrDefault(e => e.Name.NamespaceName == WordNamespace && e.Name.LocalName == "rFonts");
        string? font = fonts == null ? null : (string?)fonts.Attribute(XName.Get("ascii", WordNamespace)) ?? (string?)fonts.Attribute(XName.Get("hAnsi", WordNamespace));
        foreach (XElement e in run.Descendants())
        {
            if ((e.Name.NamespaceName == MathNamespace || e.Name.NamespaceName == WordNamespace) && e.Name.LocalName == "t") result.Append(global::DocxportNet.DxpFontSymbols.Substitute(font, e.Value));
            else if (e.Name.NamespaceName == WordNamespace && e.Name.LocalName == "tab") result.Append('\t');
            else if (e.Name.NamespaceName == WordNamespace && e.Name.LocalName == "br") result.Append('\n');
            else if (e.Name.NamespaceName == WordNamespace && e.Name.LocalName == "sym") result.Append(global::DocxportNet.DxpFontSymbols.TranslateWordSymbol((string?)e.Attribute(XName.Get("font", WordNamespace)), (string?)e.Attribute(XName.Get("char", WordNamespace))));
        }
        return result.ToString();
    }

    private static string ExtractRunText(OpenXmlElement run)
    {
        StringBuilder result = new();
        OpenXmlElement? fonts = run.Descendants().FirstOrDefault(e => e.NamespaceUri == WordNamespace && e.LocalName == "rFonts");
        string? font = fonts == null ? null : Attribute(fonts, "ascii") ?? Attribute(fonts, "hAnsi");
        foreach (OpenXmlElement e in run.Descendants())
        {
            if ((e.NamespaceUri == MathNamespace || e.NamespaceUri == WordNamespace) && e.LocalName == "t") result.Append(global::DocxportNet.DxpFontSymbols.Substitute(font, e.InnerText));
            else if (e.NamespaceUri == WordNamespace && e.LocalName == "tab") result.Append('\t');
            else if (e.NamespaceUri == WordNamespace && e.LocalName == "br") result.Append('\n');
            else if (e.NamespaceUri == WordNamespace && e.LocalName == "sym") result.Append(global::DocxportNet.DxpFontSymbols.TranslateWordSymbol(Attribute(e, "font"), Attribute(e, "char")));
        }
        return result.ToString();
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
