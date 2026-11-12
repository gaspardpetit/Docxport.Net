using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Xml.Linq;

namespace DocxportNet.Walker;

internal static class DxpNumberingCompatibility
{
    private const string Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    internal static Numbering? NormalizeIgnorableContent(Numbering? source, Func<string, bool> understood)
    {
        if (source == null || (source.MCAttributes?.ProcessContent == null &&
            !source.Descendants().Any(element => element.MCAttributes?.ProcessContent != null ||
                element.NamespaceUri.Length > 0 && element.NamespaceUri != Mc && !understood(element.NamespaceUri))))
            return source;
        var xml = XElement.Parse(source.OuterXml, LoadOptions.PreserveWhitespace);
        var inheritedIgnorable = new HashSet<string>(StringComparer.Ordinal);
        var inheritedProcess = new HashSet<(string Namespace, string Name)>();
        var changed = false;
        Visit(xml, inheritedIgnorable, inheritedProcess);
        // Reparse after lifting children: SDK nodes below unknown wrappers can
        // themselves be unknown nodes, even when their names are core OOXML.
        return changed ? new Numbering(xml.ToString(SaveOptions.DisableFormatting)) : source;

        void Visit(XElement element, HashSet<string> inheritedIgnorable,
            HashSet<(string Namespace, string Name)> inheritedProcess)
        {
            var ignorable = new HashSet<string>(inheritedIgnorable, StringComparer.Ordinal);
            var process = new HashSet<(string Namespace, string Name)>(inheritedProcess);
            foreach (var prefix in Tokens((string?)element.Attribute(XName.Get("Ignorable", Mc))))
                if (element.GetNamespaceOfPrefix(prefix) is { } ns) ignorable.Add(ns.NamespaceName);
            foreach (var token in Tokens((string?)element.Attribute(XName.Get("ProcessContent", Mc))))
            {
                var colon = token.IndexOf(':');
                if (colon > 0 && element.GetNamespaceOfPrefix(token.Substring(0, colon)) is { } ns)
                    process.Add((ns.NamespaceName, token.Substring(colon + 1)));
            }
            var elementNs = element.Name.NamespaceName;
            if (!understood(elementNs) && ignorable.Contains(elementNs))
            {
                if (process.Contains((elementNs, element.Name.LocalName)) || process.Contains((elementNs, "*")))
                {
                    foreach (var child in element.Elements().ToArray()) Visit(child, ignorable, process);
                    foreach (var child in element.Elements())
                        foreach (var binding in element.Attributes().Where(attribute => attribute.IsNamespaceDeclaration))
                            if (child.Attribute(binding.Name) == null) child.Add(new XAttribute(binding));
                    element.ReplaceWith(element.Nodes().ToArray());
                }
                else element.Remove();
                changed = true;
                return;
            }
            foreach (var child in element.Elements().ToArray()) Visit(child, ignorable, process);
        }
    }

    private static string[] Tokens(string? value) =>
        value?.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries) ?? [];

}
