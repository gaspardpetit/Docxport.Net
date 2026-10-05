namespace DocxportNet.Doc;

/// <summary>Exposes the fixed, nested DOP version ranges and selected settings blocks.</summary>
internal static class DocDopNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode location)
    {
        if (location.Children.Any(x => x.Kind == "Dop") || location.StreamName == null ||
            location.Offset == null || location.Length == null) return;
        var length = location.Length.Value;
        if (length < 500) return;
        var dop = new DocStructureNode("Dop", "DocumentProperties", location.StreamName,
            location.Offset, length);
        var extension = structure.Root.Children.SelectMany(x => x.Children)
            .Where(x => x.Kind == "FIB").SelectMany(x => x.Children)
            .FirstOrDefault(x => x.Kind == "FibRgCswNew");
        var version = extension != null && extension.Attributes.TryGetValue("version", out var value)
            ? value : null;
        var tier = version switch
        {
            "0x00D9" => 3,
            "0x0101" => 4,
            "0x010C" => 5,
            "0x0112" => length switch { >= 694 => 8, >= 690 => 7, >= 674 => 6, _ => 5 },
            _ => 2
        };
        var names = new[] { "DopBase", "Dop95", "Dop97", "Dop2000", "Dop2002",
            "Dop2003", "Dop2007", "Dop2010", "Dop2013" };
        var sizes = new[] { 84, 88, 500, 544, 594, 616, 674, 690, 694 };
        var parent = dop;
        for (var i = tier; i >= 0; i--)
        {
            if (length < sizes[i]) continue;
            var node = new DocStructureNode(names[i], names[i], location.StreamName,
                location.Offset, sizes[i]);
            parent.Children.Add(node);
            parent = node;
        }
        var baseNode = parent;
        if (baseNode.Kind == "DopBase")
            baseNode.Children.Add(new DocStructureNode("Copts60", "Compatibility60",
                location.StreamName, location.Offset + 6, 2));
        if (tier >= 1)
        {
            var dop95 = Find(dop, "Dop95");
            dop95?.Children.Add(new DocStructureNode("Copts80", "Compatibility80",
                location.StreamName, location.Offset + 84, 4));
        }
        if (tier >= 2)
        {
            var dop97 = Find(dop, "Dop97");
            dop97?.Children.Add(new DocStructureNode("DopTypography", "Typography",
                location.StreamName, location.Offset + 90, 310));
            dop97?.Children.Add(new DocStructureNode("Dogrid", "DrawingGrid",
                location.StreamName, location.Offset + 400, 10));
            dop97?.Children.Add(new DocStructureNode("Asumyi", "AutoSummary",
                location.StreamName, location.Offset + 414, 12));
        }
        if (tier >= 3)
        {
            var dop2000 = Find(dop, "Dop2000");
            dop2000?.Children.Add(new DocStructureNode("Copts", "CompatibilityOptions",
                location.StreamName, location.Offset + 508, 32));
        }
        if (tier >= 6)
        {
            var dop2007 = Find(dop, "Dop2007");
            dop2007?.Children.Add(new DocStructureNode("DopMth", "MathSettings",
                location.StreamName, location.Offset + 640, 34));
        }
        location.Children.Add(dop);
    }

    private static DocStructureNode? Find(DocStructureNode node, string kind)
    {
        if (node.Kind == kind) return node;
        foreach (var child in node.Children)
        {
            var found = Find(child, kind);
            if (found != null) return found;
        }
        return null;
    }
}
