using System.Xml;
using System.Xml.Linq;

namespace DocxportNet.Tests.Omml;

internal static class OmmlTestData
{
    public const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    public static string FixtureRoot => Path.Combine(
        FindRepositoryRoot(),
        "DocxportNet.Tests",
        "Fixtures",
        "Omml");

    public static string UpstreamRoot => Path.Combine(FixtureRoot, "Upstream", "Plurimath");

    public static string NormativeRoot => Path.Combine(FixtureRoot, "Normative");

    public static IReadOnlyList<string> UpstreamFixtures() => Directory
        .EnumerateFiles(UpstreamRoot, "*.omml", SearchOption.AllDirectories)
        .OrderBy(path => path, StringComparer.Ordinal)
        .ToArray();

    public static IReadOnlyList<string> NormativeFixtures() => Directory
        .EnumerateFiles(NormativeRoot, "*.omml", SearchOption.AllDirectories)
        .OrderBy(path => path, StringComparer.Ordinal)
        .ToArray();

    public static IReadOnlyList<string> InvalidNormativeFixtures() => Directory
        .EnumerateFiles(NormativeRoot, "*.invalid.xml", SearchOption.AllDirectories)
        .OrderBy(path => path, StringComparer.Ordinal)
        .ToArray();

    public static XElement Run(string text)
    {
        XNamespace m = MathNamespace;
        return new XElement(m + "r", new XElement(m + "t", text));
    }

    public static XElement Inline(params object[] content)
    {
        XNamespace m = MathNamespace;
        return new XElement(m + "oMath", content);
    }

    public static XElement Display(params object[] content)
    {
        XNamespace m = MathNamespace;
        return new XElement(m + "oMathPara", new XElement(m + "oMath", content));
    }

    public static XElement Fraction(XElement numerator, XElement denominator)
    {
        XNamespace m = MathNamespace;
        return new XElement(
            m + "f",
            new XElement(m + "num", numerator),
            new XElement(m + "den", denominator));
    }

    public static IEnumerable<object[]> MalformedFragments()
    {
        yield return ["truncated", "<m:oMath xmlns:m=\"" + MathNamespace + "\"><m:r>"];
        yield return ["undeclared-prefix", "<m:oMath><m:r /></m:oMath>"];
        yield return ["bare-ampersand", "<m:t>a & b</m:t>"];
        yield return ["external-entity", "<!DOCTYPE x [<!ENTITY e SYSTEM \"file:///tmp/e\">]><x>&e;</x>"];
    }

    private static string FindRepositoryRoot()
    {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "DocxportNet.sln")))
                return directory.FullName;

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate the Docxport.Net repository root.");
    }
}

internal static class XmlCanonicalizer
{
    public static string Canonicalize(string xml)
    {
        using StringReader text = new(xml);
        using XmlReader reader = XmlReader.Create(text, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = false,
        });

        XDocument document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        if (document.Root is null)
            throw new XmlException("The XML document has no root element.");

        StringWriter result = new();
        WriteElement(result, document.Root);
        return result.ToString();
    }

    private static void WriteElement(TextWriter writer, XElement element)
    {
        writer.Write('<');
        WriteName(writer, element.Name);
        foreach (XAttribute attribute in element.Attributes()
                     .Where(attribute => !attribute.IsNamespaceDeclaration)
                     .OrderBy(attribute => attribute.Name.NamespaceName, StringComparer.Ordinal)
                     .ThenBy(attribute => attribute.Name.LocalName, StringComparer.Ordinal))
        {
            writer.Write('|');
            WriteName(writer, attribute.Name);
            WriteValue(writer, attribute.Value);
        }

        writer.Write('>');
        foreach (XNode node in element.Nodes())
        {
            switch (node)
            {
                case XElement child:
                    WriteElement(writer, child);
                    break;
                case XText value when !string.IsNullOrWhiteSpace(value.Value) || !element.HasElements:
                    writer.Write('#');
                    WriteValue(writer, value.Value);
                    break;
            }
        }

        writer.Write("</");
        WriteName(writer, element.Name);
        writer.Write('>');
    }

    private static void WriteName(TextWriter writer, XName name) =>
        writer.Write($"{{{name.NamespaceName}}}{name.LocalName}");

    private static void WriteValue(TextWriter writer, string value) =>
        writer.Write($"{value.Length}:{value}");
}
