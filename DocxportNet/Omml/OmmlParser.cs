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
        ValidateOptions(options);
        if (omml is null)
            throw new DxpOmmlParseException("OMML input cannot be null.");
        if (string.IsNullOrWhiteSpace(omml))
            throw new DxpOmmlParseException("OMML input cannot be empty or whitespace.");
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

        ValidateTree(root, options);

        bool isDisplay = root.Name.LocalName == "oMathPara";
        IReadOnlyList<OmmlNode> children = isDisplay
            ? ParseParagraph(root)
            : ParseChildren(root, "/m:oMath[1]");
        return new OmmlDocument(isDisplay, children, isDisplay ? ParagraphJustification(root) : null);
    }

    public static OmmlDocument Parse(OpenXmlElement root, DxpOmmlConversionOptions options)
    {
        ValidateOptions(options);
        ValidateTree(root, options);
        bool isDisplay = root is DocumentFormat.OpenXml.Math.Paragraph;
        string rootPath = isDisplay ? "/m:oMathPara[1]" : "/m:oMath[1]";
        IReadOnlyList<OmmlNode> children = isDisplay
            ? ParseParagraph(root)
            : ParseChildren(root, rootPath);
        return new OmmlDocument(isDisplay, children, isDisplay ? ParagraphJustification(root) : null);
    }

    private static void ValidateOptions(DxpOmmlConversionOptions options)
    {
        if (options.MaxInputCharacters <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxInputCharacters), "The input limit must be positive.");
        if (options.MaxNestingDepth <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxNestingDepth), "The nesting-depth limit must be positive.");
        if (options.MaxElementCount <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxElementCount), "The element-count limit must be positive.");
        if (options.MaxOutputCharacters <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxOutputCharacters), "The output limit must be positive.");
    }

    private static void ValidateTree(XElement root, DxpOmmlConversionOptions options)
    {
        Stack<(XElement Element, int Depth)> pending = new();
        pending.Push((root, 1));
        int count = 0;
        while (pending.Count != 0)
        {
            (XElement element, int depth) = pending.Pop();
            ValidateResourcePosition(++count, depth, options);
            foreach (XElement child in element.Elements())
                pending.Push((child, depth + 1));
        }
    }

    private static void ValidateTree(OpenXmlElement root, DxpOmmlConversionOptions options)
    {
        Stack<(OpenXmlElement Element, int Depth)> pending = new();
        pending.Push((root, 1));
        int count = 0;
        while (pending.Count != 0)
        {
            (OpenXmlElement element, int depth) = pending.Pop();
            ValidateResourcePosition(++count, depth, options);
            foreach (OpenXmlElement child in element.ChildElements)
                pending.Push((child, depth + 1));
        }
    }

    private static void ValidateResourcePosition(int count, int depth, DxpOmmlConversionOptions options)
    {
        if (depth > options.MaxNestingDepth)
            throw new DxpOmmlResourceLimitException($"OMML input exceeds the {options.MaxNestingDepth} element nesting-depth limit.");
        if (count > options.MaxElementCount)
            throw new DxpOmmlResourceLimitException($"OMML input exceeds the {options.MaxElementCount} element-count limit.");
    }

    private static IReadOnlyList<OmmlNode> ParseParagraph(XElement paragraph)
    {
        List<OmmlNode> children = new();
        int mathIndex = 0;
        int paragraphRunIndex = 0;
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
            if (child.Name == XName.Get("ctrlPr", MathNamespace) || child.Name == XName.Get("argPr", MathNamespace))
                continue;

            if (child.Name == XName.Get("oMathParaPr", MathNamespace))
                continue;
            if (child.Name == XName.Get("oMath", MathNamespace))
            {
                mathIndex++;
                string path = $"/m:oMathPara[1]/m:oMath[{Index(mathIndex)}]";
                children.Add(new OmmlSequence(path, ParseChildren(child, path)));
            }
            else if (child.Name == XName.Get("r", MathNamespace) &&
                     child.Descendants().Any(IsWordBreak))
            {
                paragraphRunIndex++;
                string path = $"/m:oMathPara[1]/m:r[{Index(paragraphRunIndex)}]";
                children.AddRange(ParseParagraphRun(child, path));
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
        return CoalesceEmbeddedWordprocessing(children);
    }

    private static IReadOnlyList<OmmlNode> ParseParagraph(OpenXmlElement paragraph)
    {
        List<OmmlNode> children = new();
        int mathIndex = 0;
        int paragraphRunIndex = 0;
        Dictionary<(string Namespace, string LocalName), int> indexes = new();
        foreach (OpenXmlElement child in paragraph.ChildElements)
        {
            if (child is DocumentFormat.OpenXml.Math.ParagraphProperties)
                continue;
            if (child is DocumentFormat.OpenXml.Math.OfficeMath)
            {
                mathIndex++;
                string path = $"/m:oMathPara[1]/m:oMath[{Index(mathIndex)}]";
                children.Add(new OmmlSequence(path, ParseChildren(child, path)));
            }
            else if (child.NamespaceUri == MathNamespace && child.LocalName == "r" &&
                     child.Descendants().Any(IsWordBreak))
            {
                paragraphRunIndex++;
                string path = $"/m:oMathPara[1]/m:r[{Index(paragraphRunIndex)}]";
                children.AddRange(ParseParagraphRun(child, path));
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
        return CoalesceEmbeddedWordprocessing(children);
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
            if (child.Name == XName.Get("ctrlPr", MathNamespace))
                continue;

            indexes.TryGetValue(child.Name, out int index);
            indexes[child.Name] = ++index;
            string path = $"{parentPath}/{QualifiedName(child)}[{Index(index)}]";
            children.Add(ParseElement(child, path));
        }

        return CoalesceEmbeddedWordprocessing(children);
    }

    private static IReadOnlyList<OmmlNode> ParseChildren(OpenXmlElement parent, string parentPath)
    {
        List<OmmlNode> children = new();
        Dictionary<(string Namespace, string LocalName), int> indexes = new();
        foreach (OpenXmlElement child in parent.ChildElements)
        {
            if (child.NamespaceUri == MathNamespace && child.LocalName is "ctrlPr" or "argPr")
                continue;
            (string Namespace, string LocalName) name = (child.NamespaceUri, child.LocalName);
            indexes.TryGetValue(name, out int index);
            indexes[name] = ++index;
            string path = $"{parentPath}/{QualifiedName(child)}[{Index(index)}]";
            children.Add(ParseElement(child, path));
        }

        return CoalesceEmbeddedWordprocessing(children);
    }

    private static IReadOnlyList<OmmlNode> CoalesceEmbeddedWordprocessing(IReadOnlyList<OmmlNode> children)
    {
        List<OmmlNode> result = new();
        for (int i = 0; i < children.Count; i++)
        {
            if (children[i] is not OmmlUnsupported first || !first.ElementName.StartsWith("w:", StringComparison.Ordinal))
            {
                result.Add(children[i]);
                continue;
            }

            List<OmmlUnsupported> group = new() { first };
            while (i + 1 < children.Count && children[i + 1] is OmmlUnsupported next &&
                   next.ElementName.StartsWith("w:", StringComparison.Ordinal))
            {
                group.Add(next);
                i++;
            }
            if (group.Count == 1)
            {
                result.Add(first);
                continue;
            }

            result.Add(new OmmlUnsupported(first.Path, "w:embeddedContent",
                string.Concat(group.Select(item => item.VisibleText)),
                group.SelectMany(item => item.XmlElements).ToArray(),
                group.SelectMany(item => item.OpenXmlElements).ToArray()));
        }
        return result;
    }

    private static OmmlNode ParseElement(XElement element, string path)
    {
        OmmlNode node = ParseElementCore(element, path);
        node.ControlPresentation = ParseControlPresentation(element);
        node.ControlRevision = ParseControlRevision(element);
        return node;
    }

    private static OmmlNode ParseElementCore(XElement element, string path) => element.Name.NamespaceName == MathNamespace
        ? element.Name.LocalName switch
        {
            "r" => ParseRun(element, path),
            "t" => ParseBareText(element, path),
            "f" => ParseFraction(element, path),
            "rad" => ParseRadical(element, path),
            "sSub" => ParseScript(element, path, OmmlScriptType.Subscript),
            "sSup" => ParseScript(element, path, OmmlScriptType.Superscript),
            "sSubSup" => ParseScript(element, path, OmmlScriptType.SubSup),
            "sPre" => ParseScript(element, path, OmmlScriptType.PreSubSup),
            "d" => ParseDelimiter(element, path),
            "acc" => ParseDecoration(element, path, OmmlDecorationType.Accent),
            "bar" => ParseDecoration(element, path, OmmlDecorationType.Bar),
            "groupChr" => ParseDecoration(element, path, OmmlDecorationType.GroupCharacter),
            "func" => ParseFunction(element, path),
            "limLow" => ParseLimit(element, path, OmmlLimitType.Lower),
            "limUpp" => ParseLimit(element, path, OmmlLimitType.Upper),
            "nary" => ParseNary(element, path),
            "m" => ParseMatrix(element, path),
            "eqArr" => ParseEquationArray(element, path),
            "box" => ParseBox(element, path),
            "borderBox" => ParseBorderBox(element, path),
            "phant" => ParsePhantom(element, path),
            _ => ParseUnsupported(element, path),
        }
        : ParseUnsupported(element, path);

    private static OmmlNode ParseElement(OpenXmlElement element, string path)
    {
        OmmlNode node = ParseElementCore(element, path);
        node.ControlPresentation = ParseControlPresentation(element);
        node.ControlRevision = ParseControlRevision(element);
        return node;
    }

    private static OmmlNode ParseElementCore(OpenXmlElement element, string path) => element.NamespaceUri == MathNamespace
        ? element.LocalName switch
        {
            "r" => ParseRun(element, path),
            "t" => ParseBareText(element, path),
            "f" => ParseFraction(element, path),
            "rad" => ParseRadical(element, path),
            "sSub" => ParseScript(element, path, OmmlScriptType.Subscript),
            "sSup" => ParseScript(element, path, OmmlScriptType.Superscript),
            "sSubSup" => ParseScript(element, path, OmmlScriptType.SubSup),
            "sPre" => ParseScript(element, path, OmmlScriptType.PreSubSup),
            "d" => ParseDelimiter(element, path),
            "acc" => ParseDecoration(element, path, OmmlDecorationType.Accent),
            "bar" => ParseDecoration(element, path, OmmlDecorationType.Bar),
            "groupChr" => ParseDecoration(element, path, OmmlDecorationType.GroupCharacter),
            "func" => ParseFunction(element, path),
            "limLow" => ParseLimit(element, path, OmmlLimitType.Lower),
            "limUpp" => ParseLimit(element, path, OmmlLimitType.Upper),
            "nary" => ParseNary(element, path),
            "m" => ParseMatrix(element, path),
            "eqArr" => ParseEquationArray(element, path),
            "box" => ParseBox(element, path),
            "borderBox" => ParseBorderBox(element, path),
            "phant" => ParsePhantom(element, path),
            _ => ParseUnsupported(element, path),
        }
        : ParseUnsupported(element, path);

    private static OmmlFraction ParseFraction(XElement element, string path)
    {
        XElement? properties = MathChild(element, "fPr");
        string? type = properties == null ? null : MathChild(properties, "type")?.Attribute(XName.Get("val", MathNamespace))?.Value;
        return new OmmlFraction(path, FractionType(type), ParseArgument(MathChild(element, "num"), path + "/m:num[1]"),
            ParseArgument(MathChild(element, "den"), path + "/m:den[1]"), properties?.Descendants(XName.Get("ctrlPr", MathNamespace)).Any() == true);
    }

    private static OmmlFraction ParseFraction(OpenXmlElement element, string path)
    {
        OpenXmlElement? properties = MathChild(element, "fPr");
        string? type = properties == null ? null : Attribute(MathChild(properties, "type"), "val");
        return new OmmlFraction(path, FractionType(type), ParseArgument(MathChild(element, "num"), path + "/m:num[1]"),
            ParseArgument(MathChild(element, "den"), path + "/m:den[1]"), properties?.Descendants().Any(e => e.NamespaceUri == MathNamespace && e.LocalName == "ctrlPr") == true);
    }

    private static OmmlFractionType FractionType(string? value) => value switch
    { "skw" => OmmlFractionType.Skewed, "lin" => OmmlFractionType.Linear, "noBar" => OmmlFractionType.NoBar, _ => OmmlFractionType.Bar };

    private static OmmlRadical ParseRadical(XElement element, string path)
    {
        XElement? properties = MathChild(element, "radPr");
        XElement? hide = properties == null ? null : MathChild(properties, "degHide");
        XElement? degree = MathChild(element, "deg");
        return new OmmlRadical(path, ParseArgument(MathChild(element, "e"), path + "/m:e[1]"),
            ParseArgument(degree, path + "/m:deg[1]"), degree != null, hide != null && Enabled((string?)hide.Attribute(XName.Get("val", MathNamespace))),
            properties?.Descendants(XName.Get("ctrlPr", MathNamespace)).Any() == true);
    }

    private static OmmlRadical ParseRadical(OpenXmlElement element, string path)
    {
        OpenXmlElement? properties = MathChild(element, "radPr");
        OpenXmlElement? hide = properties == null ? null : MathChild(properties, "degHide");
        OpenXmlElement? degree = MathChild(element, "deg");
        return new OmmlRadical(path, ParseArgument(MathChild(element, "e"), path + "/m:e[1]"),
            ParseArgument(degree, path + "/m:deg[1]"), degree != null, hide != null && Enabled(Attribute(hide, "val")),
            properties?.Descendants().Any(e => e.NamespaceUri == MathNamespace && e.LocalName == "ctrlPr") == true);
    }

    private static OmmlScript ParseScript(XElement element, string path, OmmlScriptType type)
    {
        XElement? properties = element.Elements().FirstOrDefault(e => e.Name.NamespaceName == MathNamespace && e.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal));
        XElement? align = properties == null ? null : MathChild(properties, "alnScr");
        return new OmmlScript(path, type, ParseArgument(MathChild(element, "e"), path + "/m:e[1]"),
            ParseArgument(MathChild(element, "sub"), path + "/m:sub[1]"), ParseArgument(MathChild(element, "sup"), path + "/m:sup[1]"),
            align != null && Enabled((string?)align.Attribute(XName.Get("val", MathNamespace))), properties?.Descendants(XName.Get("ctrlPr", MathNamespace)).Any() == true);
    }

    private static OmmlScript ParseScript(OpenXmlElement element, string path, OmmlScriptType type)
    {
        OpenXmlElement? properties = element.ChildElements.FirstOrDefault(e => e.NamespaceUri == MathNamespace && e.LocalName.EndsWith("Pr", StringComparison.Ordinal));
        OpenXmlElement? align = properties == null ? null : MathChild(properties, "alnScr");
        return new OmmlScript(path, type, ParseArgument(MathChild(element, "e"), path + "/m:e[1]"),
            ParseArgument(MathChild(element, "sub"), path + "/m:sub[1]"), ParseArgument(MathChild(element, "sup"), path + "/m:sup[1]"),
            align != null && Enabled(Attribute(align, "val")), properties?.Descendants().Any(e => e.NamespaceUri == MathNamespace && e.LocalName == "ctrlPr") == true);
    }

    private static OmmlDelimiter ParseDelimiter(XElement element, string path)
    {
        XElement? properties = MathChild(element, "dPr");
        return new OmmlDelimiter(path, CharProperty(properties, "begChr", "("),
            CharProperty(properties, "sepChr", "|"), CharProperty(properties, "endChr", ")"),
            OnOffProperty(properties, "grow", true), ShapeProperty(properties),
            element.Elements(XName.Get("e", MathNamespace)).Select((e, i) => ParseArgument(e, $"{path}/m:e[{Index(i + 1)}]")).ToArray(),
            HasControlProperties(properties));
    }

    private static OmmlDelimiter ParseDelimiter(OpenXmlElement element, string path)
    {
        OpenXmlElement? properties = MathChild(element, "dPr");
        return new OmmlDelimiter(path, CharProperty(properties, "begChr", "("),
            CharProperty(properties, "sepChr", "|"), CharProperty(properties, "endChr", ")"),
            OnOffProperty(properties, "grow", true), ShapeProperty(properties),
            element.ChildElements.Where(e => e.NamespaceUri == MathNamespace && e.LocalName == "e").Select((e, i) => ParseArgument(e, $"{path}/m:e[{Index(i + 1)}]")).ToArray(),
            HasControlProperties(properties));
    }

    private static OmmlDecoration ParseDecoration(XElement element, string path, OmmlDecorationType type)
    {
        string propertyName = type switch { OmmlDecorationType.Accent => "accPr", OmmlDecorationType.Bar => "barPr", _ => "groupChrPr" };
        XElement? properties = MathChild(element, propertyName);
        OmmlVerticalPosition position = PositionProperty(properties, "pos", type == OmmlDecorationType.Accent ? OmmlVerticalPosition.Top : OmmlVerticalPosition.Bottom);
        string character = type switch { OmmlDecorationType.Accent => CharProperty(properties, "chr", "̂"), OmmlDecorationType.Bar => "―", _ => CharProperty(properties, "chr", position == OmmlVerticalPosition.Top ? "⏞" : "⏟") };
        return new OmmlDecoration(path, type, character, position, PositionProperty(properties, "vertJc", OmmlVerticalPosition.Top),
            ParseArgument(MathChild(element, "e"), path + "/m:e[1]"), HasControlProperties(properties));
    }

    private static OmmlDecoration ParseDecoration(OpenXmlElement element, string path, OmmlDecorationType type)
    {
        string propertyName = type switch { OmmlDecorationType.Accent => "accPr", OmmlDecorationType.Bar => "barPr", _ => "groupChrPr" };
        OpenXmlElement? properties = MathChild(element, propertyName);
        OmmlVerticalPosition position = PositionProperty(properties, "pos", type == OmmlDecorationType.Accent ? OmmlVerticalPosition.Top : OmmlVerticalPosition.Bottom);
        string character = type switch { OmmlDecorationType.Accent => CharProperty(properties, "chr", "̂"), OmmlDecorationType.Bar => "―", _ => CharProperty(properties, "chr", position == OmmlVerticalPosition.Top ? "⏞" : "⏟") };
        return new OmmlDecoration(path, type, character, position, PositionProperty(properties, "vertJc", OmmlVerticalPosition.Top),
            ParseArgument(MathChild(element, "e"), path + "/m:e[1]"), HasControlProperties(properties));
    }

    private static OmmlFunction ParseFunction(XElement element, string path)
    {
        XElement? properties = MathChild(element, "funcPr");
        return new OmmlFunction(path, ParseArgument(MathChild(element, "fName"), path + "/m:fName[1]"),
            ParseArgument(MathChild(element, "e"), path + "/m:e[1]"), HasControlProperties(properties));
    }

    private static OmmlFunction ParseFunction(OpenXmlElement element, string path)
    {
        OpenXmlElement? properties = MathChild(element, "funcPr");
        return new OmmlFunction(path, ParseArgument(MathChild(element, "fName"), path + "/m:fName[1]"),
            ParseArgument(MathChild(element, "e"), path + "/m:e[1]"), HasControlProperties(properties));
    }

    private static OmmlLimit ParseLimit(XElement element, string path, OmmlLimitType type)
    {
        XElement? properties = MathChild(element, type == OmmlLimitType.Lower ? "limLowPr" : "limUppPr");
        return new OmmlLimit(path, type, ParseArgument(MathChild(element, "e"), path + "/m:e[1]"),
            ParseArgument(MathChild(element, "lim"), path + "/m:lim[1]"), HasControlProperties(properties));
    }

    private static OmmlLimit ParseLimit(OpenXmlElement element, string path, OmmlLimitType type)
    {
        OpenXmlElement? properties = MathChild(element, type == OmmlLimitType.Lower ? "limLowPr" : "limUppPr");
        return new OmmlLimit(path, type, ParseArgument(MathChild(element, "e"), path + "/m:e[1]"),
            ParseArgument(MathChild(element, "lim"), path + "/m:lim[1]"), HasControlProperties(properties));
    }

    private static OmmlNary ParseNary(XElement element, string path)
    {
        XElement? properties = MathChild(element, "naryPr");
        return new OmmlNary(path, CharProperty(properties, "chr", "∫"), LimitLocationProperty(properties),
            OnOffProperty(properties, "grow", true), OnOffProperty(properties, "subHide", false),
            OnOffProperty(properties, "supHide", false), ParseArgument(MathChild(element, "sub"), path + "/m:sub[1]"),
            ParseArgument(MathChild(element, "sup"), path + "/m:sup[1]"), ParseArgument(MathChild(element, "e"), path + "/m:e[1]"),
            HasControlProperties(properties));
    }

    private static OmmlNary ParseNary(OpenXmlElement element, string path)
    {
        OpenXmlElement? properties = MathChild(element, "naryPr");
        return new OmmlNary(path, CharProperty(properties, "chr", "∫"), LimitLocationProperty(properties),
            OnOffProperty(properties, "grow", true), OnOffProperty(properties, "subHide", false),
            OnOffProperty(properties, "supHide", false), ParseArgument(MathChild(element, "sub"), path + "/m:sub[1]"),
            ParseArgument(MathChild(element, "sup"), path + "/m:sup[1]"), ParseArgument(MathChild(element, "e"), path + "/m:e[1]"),
            HasControlProperties(properties));
    }

    private static OmmlMatrix ParseMatrix(XElement element, string path)
    {
        XElement? properties = MathChild(element, "mPr");
        XElement? columnList = properties == null ? null : MathChild(properties, "mcs");
        OmmlMatrixColumn[] columns = columnList?.Elements(XName.Get("mc", MathNamespace))
            .Select(ParseMatrixColumn).ToArray() ?? Array.Empty<OmmlMatrixColumn>();
        OmmlMatrixRow[] rows = element.Elements(XName.Get("mr", MathNamespace))
            .Select((row, rowIndex) => new OmmlMatrixRow(row.Elements(XName.Get("e", MathNamespace))
                .Select((cell, cellIndex) => ParseArgument(cell, $"{path}/m:mr[{Index(rowIndex + 1)}]/m:e[{Index(cellIndex + 1)}]"))
                .ToArray())).ToArray();
        return CreateMatrix(path, properties, rows, columns);
    }

    private static OmmlMatrix ParseMatrix(OpenXmlElement element, string path)
    {
        OpenXmlElement? properties = MathChild(element, "mPr");
        OpenXmlElement? columnList = properties == null ? null : MathChild(properties, "mcs");
        OmmlMatrixColumn[] columns = columnList?.ChildElements
            .Where(e => e.NamespaceUri == MathNamespace && e.LocalName == "mc")
            .Select(ParseMatrixColumn).ToArray() ?? Array.Empty<OmmlMatrixColumn>();
        OmmlMatrixRow[] rows = element.ChildElements
            .Where(e => e.NamespaceUri == MathNamespace && e.LocalName == "mr")
            .Select((row, rowIndex) => new OmmlMatrixRow(row.ChildElements
                .Where(e => e.NamespaceUri == MathNamespace && e.LocalName == "e")
                .Select((cell, cellIndex) => ParseArgument(cell, $"{path}/m:mr[{Index(rowIndex + 1)}]/m:e[{Index(cellIndex + 1)}]"))
                .ToArray())).ToArray();
        return CreateMatrix(path, properties, rows, columns);
    }

    private static OmmlMatrixColumn ParseMatrixColumn(XElement column)
    {
        XElement? properties = MathChild(column, "mcPr");
        return new OmmlMatrixColumn(IntegerProperty(properties, "count", 1, 1, 255),
            HorizontalAlignmentProperty(properties, "mcJc"));
    }

    private static OmmlMatrixColumn ParseMatrixColumn(OpenXmlElement column)
    {
        OpenXmlElement? properties = MathChild(column, "mcPr");
        return new OmmlMatrixColumn(IntegerProperty(properties, "count", 1, 1, 255),
            HorizontalAlignmentProperty(properties, "mcJc"));
    }

    private static OmmlMatrix CreateMatrix(string path, XElement? properties,
        IReadOnlyList<OmmlMatrixRow> rows, IReadOnlyList<OmmlMatrixColumn> columns) =>
        new(path, rows, columns, VerticalAlignmentProperty(properties, "baseJc"),
            OnOffProperty(properties, "plcHide", false), UnsignedProperty(properties, "rSp"),
            IntegerProperty(properties, "rSpRule", 0, 0, 4), UnsignedProperty(properties, "cSp"),
            UnsignedProperty(properties, "cGp"), IntegerProperty(properties, "cGpRule", 0, 0, 4),
            HasControlProperties(properties));

    private static OmmlMatrix CreateMatrix(string path, OpenXmlElement? properties,
        IReadOnlyList<OmmlMatrixRow> rows, IReadOnlyList<OmmlMatrixColumn> columns) =>
        new(path, rows, columns, VerticalAlignmentProperty(properties, "baseJc"),
            OnOffProperty(properties, "plcHide", false), UnsignedProperty(properties, "rSp"),
            IntegerProperty(properties, "rSpRule", 0, 0, 4), UnsignedProperty(properties, "cSp"),
            UnsignedProperty(properties, "cGp"), IntegerProperty(properties, "cGpRule", 0, 0, 4),
            HasControlProperties(properties));

    private static OmmlEquationArray ParseEquationArray(XElement element, string path)
    {
        XElement? properties = MathChild(element, "eqArrPr");
        OmmlSequence[] rows = element.Elements(XName.Get("e", MathNamespace))
            .Select((row, i) => ParseArgument(row, $"{path}/m:e[{Index(i + 1)}]")).ToArray();
        return new OmmlEquationArray(path, rows, VerticalAlignmentProperty(properties, "baseJc"),
            OnOffProperty(properties, "maxDist", false), OnOffProperty(properties, "objDist", false),
            UnsignedProperty(properties, "rSp"), IntegerProperty(properties, "rSpRule", 0, 0, 4),
            HasControlProperties(properties));
    }

    private static OmmlEquationArray ParseEquationArray(OpenXmlElement element, string path)
    {
        OpenXmlElement? properties = MathChild(element, "eqArrPr");
        OmmlSequence[] rows = element.ChildElements.Where(e => e.NamespaceUri == MathNamespace && e.LocalName == "e")
            .Select((row, i) => ParseArgument(row, $"{path}/m:e[{Index(i + 1)}]")).ToArray();
        return new OmmlEquationArray(path, rows, VerticalAlignmentProperty(properties, "baseJc"),
            OnOffProperty(properties, "maxDist", false), OnOffProperty(properties, "objDist", false),
            UnsignedProperty(properties, "rSp"), IntegerProperty(properties, "rSpRule", 0, 0, 4),
            HasControlProperties(properties));
    }

    private static OmmlBox ParseBox(XElement element, string path)
    {
        XElement? properties = MathChild(element, "boxPr");
        XElement? manualBreak = properties == null ? null : MathChild(properties, "brk");
        return new OmmlBox(path, ParseArgument(MathChild(element, "e"), path + "/m:e[1]"),
            OnOffProperty(properties, "opEmu", false), OnOffProperty(properties, "noBreak", false),
            OnOffProperty(properties, "diff", false), manualBreak == null ? null :
                ParseInteger((string?)manualBreak.Attribute("alnAt") ??
                    (string?)manualBreak.Attribute(XName.Get("alnAt", MathNamespace)), 0, 0, 255),
            OnOffProperty(properties, "aln", false), HasControlProperties(properties));
    }

    private static OmmlBox ParseBox(OpenXmlElement element, string path)
    {
        OpenXmlElement? properties = MathChild(element, "boxPr");
        OpenXmlElement? manualBreak = properties == null ? null : MathChild(properties, "brk");
        return new OmmlBox(path, ParseArgument(MathChild(element, "e"), path + "/m:e[1]"),
            OnOffProperty(properties, "opEmu", false), OnOffProperty(properties, "noBreak", false),
            OnOffProperty(properties, "diff", false), manualBreak == null ? null :
                ParseInteger(Attribute(manualBreak, "alnAt"), 0, 0, 255),
            OnOffProperty(properties, "aln", false), HasControlProperties(properties));
    }

    private static OmmlBorderBox ParseBorderBox(XElement element, string path)
    {
        XElement? properties = MathChild(element, "borderBoxPr");
        return new OmmlBorderBox(path, ParseArgument(MathChild(element, "e"), path + "/m:e[1]"),
            OnOffProperty(properties, "hideTop", false), OnOffProperty(properties, "hideBot", false),
            OnOffProperty(properties, "hideLeft", false), OnOffProperty(properties, "hideRight", false),
            OnOffProperty(properties, "strikeH", false), OnOffProperty(properties, "strikeV", false),
            OnOffProperty(properties, "strikeBLTR", false), OnOffProperty(properties, "strikeTLBR", false),
            HasControlProperties(properties));
    }

    private static OmmlBorderBox ParseBorderBox(OpenXmlElement element, string path)
    {
        OpenXmlElement? properties = MathChild(element, "borderBoxPr");
        return new OmmlBorderBox(path, ParseArgument(MathChild(element, "e"), path + "/m:e[1]"),
            OnOffProperty(properties, "hideTop", false), OnOffProperty(properties, "hideBot", false),
            OnOffProperty(properties, "hideLeft", false), OnOffProperty(properties, "hideRight", false),
            OnOffProperty(properties, "strikeH", false), OnOffProperty(properties, "strikeV", false),
            OnOffProperty(properties, "strikeBLTR", false), OnOffProperty(properties, "strikeTLBR", false),
            HasControlProperties(properties));
    }

    private static OmmlPhantom ParsePhantom(XElement element, string path)
    {
        XElement? properties = MathChild(element, "phantPr");
        return new OmmlPhantom(path, ParseArgument(MathChild(element, "e"), path + "/m:e[1]"),
            OnOffProperty(properties, "show", true), OnOffProperty(properties, "zeroWid", false),
            OnOffProperty(properties, "zeroAsc", false), OnOffProperty(properties, "zeroDesc", false),
            OnOffProperty(properties, "transp", false), HasControlProperties(properties));
    }

    private static OmmlPhantom ParsePhantom(OpenXmlElement element, string path)
    {
        OpenXmlElement? properties = MathChild(element, "phantPr");
        return new OmmlPhantom(path, ParseArgument(MathChild(element, "e"), path + "/m:e[1]"),
            OnOffProperty(properties, "show", true), OnOffProperty(properties, "zeroWid", false),
            OnOffProperty(properties, "zeroAsc", false), OnOffProperty(properties, "zeroDesc", false),
            OnOffProperty(properties, "transp", false), HasControlProperties(properties));
    }

    private static string CharProperty(XElement? properties, string name, string defaultValue)
    {
        XElement? property = properties == null ? null : MathChild(properties, name);
        return property == null ? defaultValue : (string?)property.Attribute(XName.Get("val", MathNamespace)) ?? string.Empty;
    }

    private static string CharProperty(OpenXmlElement? properties, string name, string defaultValue)
    {
        OpenXmlElement? property = properties == null ? null : MathChild(properties, name);
        return property == null ? defaultValue : Attribute(property, "val") ?? string.Empty;
    }

    private static bool OnOffProperty(XElement? properties, string name, bool defaultValue)
    { XElement? property = properties == null ? null : MathChild(properties, name); return property == null ? defaultValue : Enabled((string?)property.Attribute(XName.Get("val", MathNamespace))); }
    private static bool OnOffProperty(OpenXmlElement? properties, string name, bool defaultValue)
    { OpenXmlElement? property = properties == null ? null : MathChild(properties, name); return property == null ? defaultValue : Enabled(Attribute(property, "val")); }
    private static int IntegerProperty(XElement? properties, string name, int defaultValue, int minimum, int maximum)
    { XElement? property = properties == null ? null : MathChild(properties, name); return ParseInteger(property == null ? null : (string?)property.Attribute(XName.Get("val", MathNamespace)), defaultValue, minimum, maximum); }
    private static int IntegerProperty(OpenXmlElement? properties, string name, int defaultValue, int minimum, int maximum)
    { OpenXmlElement? property = properties == null ? null : MathChild(properties, name); return ParseInteger(Attribute(property, "val"), defaultValue, minimum, maximum); }
    private static int ParseInteger(string? value, int defaultValue, int minimum, int maximum) =>
        int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
            ? parsed < minimum ? minimum : parsed > maximum ? maximum : parsed
            : defaultValue;
    private static uint UnsignedProperty(XElement? properties, string name)
    { XElement? property = properties == null ? null : MathChild(properties, name); return ParseUnsigned(property == null ? null : (string?)property.Attribute(XName.Get("val", MathNamespace))); }
    private static uint UnsignedProperty(OpenXmlElement? properties, string name)
    { OpenXmlElement? property = properties == null ? null : MathChild(properties, name); return ParseUnsigned(Attribute(property, "val")); }
    private static uint ParseUnsigned(string? value) => uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out uint parsed) ? parsed : 0;
    private static OmmlHorizontalAlignment HorizontalAlignmentProperty(XElement? properties, string name) =>
        CharProperty(properties, name, "center") switch { "left" => OmmlHorizontalAlignment.Left, "right" => OmmlHorizontalAlignment.Right, _ => OmmlHorizontalAlignment.Center };
    private static OmmlHorizontalAlignment HorizontalAlignmentProperty(OpenXmlElement? properties, string name) =>
        CharProperty(properties, name, "center") switch { "left" => OmmlHorizontalAlignment.Left, "right" => OmmlHorizontalAlignment.Right, _ => OmmlHorizontalAlignment.Center };
    private static OmmlVerticalAlignment VerticalAlignmentProperty(XElement? properties, string name) =>
        CharProperty(properties, name, "center") switch { "top" => OmmlVerticalAlignment.Top, "bot" => OmmlVerticalAlignment.Bottom, _ => OmmlVerticalAlignment.Center };
    private static OmmlVerticalAlignment VerticalAlignmentProperty(OpenXmlElement? properties, string name) =>
        CharProperty(properties, name, "center") switch { "top" => OmmlVerticalAlignment.Top, "bot" => OmmlVerticalAlignment.Bottom, _ => OmmlVerticalAlignment.Center };
    private static OmmlDelimiterShape ShapeProperty(XElement? properties) => properties == null || CharProperty(properties, "shp", "centered") != "match" ? OmmlDelimiterShape.Centered : OmmlDelimiterShape.Match;
    private static OmmlDelimiterShape ShapeProperty(OpenXmlElement? properties) => properties == null || CharProperty(properties, "shp", "centered") != "match" ? OmmlDelimiterShape.Centered : OmmlDelimiterShape.Match;
    private static OmmlVerticalPosition PositionProperty(XElement? properties, string name, OmmlVerticalPosition defaultValue) => properties == null || CharProperty(properties, name, defaultValue == OmmlVerticalPosition.Top ? "top" : "bot") != "top" ? (properties == null ? defaultValue : OmmlVerticalPosition.Bottom) : OmmlVerticalPosition.Top;
    private static OmmlVerticalPosition PositionProperty(OpenXmlElement? properties, string name, OmmlVerticalPosition defaultValue) => properties == null || CharProperty(properties, name, defaultValue == OmmlVerticalPosition.Top ? "top" : "bot") != "top" ? (properties == null ? defaultValue : OmmlVerticalPosition.Bottom) : OmmlVerticalPosition.Top;
    private static DxpOmmlLimitLocation? LimitLocationProperty(XElement? properties)
    { XElement? property = properties == null ? null : MathChild(properties, "limLoc"); return property == null ? null : LimitLocation((string?)property.Attribute(XName.Get("val", MathNamespace))); }
    private static DxpOmmlLimitLocation? LimitLocationProperty(OpenXmlElement? properties)
    { OpenXmlElement? property = properties == null ? null : MathChild(properties, "limLoc"); return property == null ? null : LimitLocation(Attribute(property, "val")); }
    private static DxpOmmlLimitLocation LimitLocation(string? value) => value == "subSup" ? DxpOmmlLimitLocation.SubscriptSuperscript : DxpOmmlLimitLocation.UnderOver;
    private static bool HasControlProperties(XElement? properties) => properties?.Descendants(XName.Get("ctrlPr", MathNamespace)).Any() == true;
    private static bool HasControlProperties(OpenXmlElement? properties) => properties?.Descendants().Any(e => e.NamespaceUri == MathNamespace && e.LocalName == "ctrlPr") == true;

    private static OmmlSequence ParseArgument(XElement? argument, string path)
    {
        XElement? argumentProperties = argument == null ? null : MathChild(argument, "argPr");
        XElement? argumentSize = argumentProperties == null ? null : MathChild(argumentProperties, "argSz");
        int? size = argumentSize == null || !ArgumentSizeApplies(path) ? null : ParseInteger(
            (string?)argumentSize.Attribute(XName.Get("val", MathNamespace)), 0, -2, 2);
        OmmlSequence sequence = new(path, argument == null ? Array.Empty<OmmlNode>() : ParseChildren(argument, path), size);
        if (argument != null)
        {
            sequence.ControlPresentation = ParseControlPresentation(argument);
            sequence.ControlRevision = ParseControlRevision(argument);
        }
        return sequence;
    }

    private static OmmlSequence ParseArgument(OpenXmlElement? argument, string path)
    {
        OpenXmlElement? argumentProperties = argument == null ? null : MathChild(argument, "argPr");
        OpenXmlElement? argumentSize = argumentProperties == null ? null : MathChild(argumentProperties, "argSz");
        int? size = argumentSize == null || !ArgumentSizeApplies(path)
            ? null : ParseInteger(Attribute(argumentSize, "val"), 0, -2, 2);
        OmmlSequence sequence = new(path, argument == null ? Array.Empty<OmmlNode>() : ParseChildren(argument, path), size);
        if (argument != null)
        {
            sequence.ControlPresentation = ParseControlPresentation(argument);
            sequence.ControlRevision = ParseControlRevision(argument);
        }
        return sequence;
    }

    private static bool ArgumentSizeApplies(string path)
    {
        string[] segments = path.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        if (segments.Length < 2) return false;
        static string Name(string segment)
        {
            int bracket = segment.IndexOf('[');
            return bracket < 0 ? segment : segment.Substring(0, bracket);
        }
        return (Name(segments[segments.Length - 2]), Name(segments[segments.Length - 1])) switch
        {
            ("m:box", "m:e") or ("m:groupChr", "m:e") or
            ("m:limLow", "m:lim") or ("m:limUpp", "m:lim") or
            ("m:nary", "m:sub") or ("m:nary", "m:sup") or
            ("m:rad", "m:deg") or ("m:sPre", "m:sub") or
            ("m:sPre", "m:sup") or ("m:sSub", "m:sub") or
            ("m:sSubSup", "m:sub") or ("m:sSubSup", "m:sup") or
            ("m:sSup", "m:sup") => true,
            _ => false,
        };
    }
    private static XElement? MathChild(XElement element, string name) => element.Elements(XName.Get(name, MathNamespace)).FirstOrDefault();
    private static OpenXmlElement? MathChild(OpenXmlElement element, string name) => element.ChildElements.FirstOrDefault(e => e.NamespaceUri == MathNamespace && e.LocalName == name);

    private static OmmlRun ParseRun(XElement run, string path) => ParseRunCore(
        path,
        new[] { ExtractRunText(run) },
        run.Descendants(), e => e.Name.NamespaceName, e => e.Name.LocalName,
        e => (string?)e.Attribute(XName.Get("val", e.Name.NamespaceName)) ?? (string?)e.Attribute(XName.Get("val", WordNamespace)),
        (e, name) => e.Attributes().FirstOrDefault(a => a.Name.LocalName == name)?.Value,
        RunBreakAlignment(run));

    private static OmmlRun ParseRun(OpenXmlElement run, string path)
    {
        IEnumerable<OpenXmlElement> all = run.Descendants();
        return ParseRunCore(path,
            new[] { ExtractRunText(run) },
            all, e => e.NamespaceUri, e => e.LocalName, e => Attribute(e, "val"), Attribute, RunBreakAlignment(run));
    }

    private static OmmlRun ParseBareText(XElement text, string path) => ParseRunCore(
        path, new[] { text.Value }, Array.Empty<XElement>(), e => e.Name.NamespaceName,
        e => e.Name.LocalName, _ => null, (_, _) => null, null);

    private static OmmlRun ParseBareText(OpenXmlElement text, string path) => ParseRunCore(
        path, new[] { text.InnerText }, Array.Empty<OpenXmlElement>(), e => e.NamespaceUri,
        e => e.LocalName, _ => null, (_, _) => null, null);

    private static IReadOnlyList<OmmlNode> ParseParagraphRun(XElement run, string path)
    {
        IEnumerable<XElement> all = run.Descendants();
        return ParseParagraphRunCore(path, ExtractRunText(run), segment => ParseRunCore(path,
            new[] { segment }, all, e => e.Name.NamespaceName, e => e.Name.LocalName,
            e => (string?)e.Attribute(XName.Get("val", e.Name.NamespaceName)) ??
                 (string?)e.Attribute(XName.Get("val", WordNamespace)),
            (e, name) => e.Attributes().FirstOrDefault(a => a.Name.LocalName == name)?.Value, null));
    }

    private static IReadOnlyList<OmmlNode> ParseParagraphRun(OpenXmlElement run, string path)
    {
        IEnumerable<OpenXmlElement> all = run.Descendants();
        return ParseParagraphRunCore(path, ExtractRunText(run), segment => ParseRunCore(path,
            new[] { segment }, all, e => e.NamespaceUri, e => e.LocalName, e => Attribute(e, "val"), Attribute, null));
    }

    private static IReadOnlyList<OmmlNode> ParseParagraphRunCore(string path, string text,
        Func<string, OmmlRun> parseSegment)
    {
        List<OmmlNode> result = new();
        string[] segments = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        int breakIndex = 0;
        for (int i = 0; i < segments.Length; i++)
        {
            if (i != 0) result.Add(new OmmlBreak($"{path}/w:br[{Index(++breakIndex)}]"));
            if (segments[i].Length != 0) result.Add(parseSegment(segments[i]));
        }
        return result;
    }

    private static OmmlRun ParseRunCore<T>(string path, IEnumerable<string> texts,
        IEnumerable<T> elements, Func<T, string> ns, Func<T, string> local, Func<T, string?> val,
        Func<T, string, string?> attribute,
        int? breakAlignmentAt)
        where T : class
    {
        List<T> all = elements.ToList();
        bool Has(string space, string name) => all.Any(e => ns(e) == space && local(e) == name && Enabled(val(e)));
        string? Value(string space, string name)
        {
            T? found = all.FirstOrDefault(e => ns(e) == space && local(e) == name);
            return found == null ? null : val(found);
        }
        string? AttributeValue(string space, string elementName, string attributeName)
        {
            T? found = all.FirstOrDefault(e => ns(e) == space && local(e) == elementName);
            return found == null ? null : attribute(found, attributeName);
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
        string? color = NormalizeColor(Value(WordNamespace, "color"));
        double? fontSizePoints = uint.TryParse(Value(WordNamespace, "sz"), out uint halfPoints)
            ? halfPoints / 2.0 : null;
        string? fontFamily = AttributeValue(WordNamespace, "rFonts", "ascii") ??
                             AttributeValue(WordNamespace, "rFonts", "hAnsi");
        OmmlRunVerticalAlignment verticalAlignment = Value(WordNamespace, "vertAlign") switch
        {
            "superscript" => OmmlRunVerticalAlignment.Superscript,
            "subscript" => OmmlRunVerticalAlignment.Subscript,
            _ => OmmlRunVerticalAlignment.Baseline,
        };
        return new OmmlRun(path, ClassifyRunText(text, literal || normal), script, style, literal, normal,
            Has(MathNamespace, "aln"), breakAlignmentAt, Value(WordNamespace, "lang"), Has(WordNamespace, "rtl"),
            color, fontSizePoints, fontFamily, verticalAlignment);
    }

    private static string? NormalizeColor(string? value) =>
        value != null && value.Length == 6 && value.All(Uri.IsHexDigit) ? value.ToUpperInvariant() : null;

    private static OmmlRunPresentation? ParseControlPresentation(XElement element)
    {
        XElement? properties = element.Elements().FirstOrDefault(child =>
            child.Name.NamespaceName == MathNamespace && child.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal));
        XElement? control = properties?.DescendantsAndSelf(XName.Get("ctrlPr", MathNamespace)).FirstOrDefault();
        XElement? runProperties = control?.Elements(XName.Get("rPr", WordNamespace)).FirstOrDefault() ??
                                  control?.Elements().FirstOrDefault(child => child.Name.NamespaceName == WordNamespace &&
                                      child.Name.LocalName is "ins" or "del" or "moveTo" or "moveFrom")?
                                      .Descendants(XName.Get("rPr", WordNamespace)).FirstOrDefault();
        if (runProperties == null) return null;
        string? Val(string name) => runProperties.Elements(XName.Get(name, WordNamespace)).FirstOrDefault()?
            .Attributes().FirstOrDefault(attribute => attribute.Name.LocalName == "val")?.Value;
        bool On(string name) => runProperties.Elements(XName.Get(name, WordNamespace)).Any(property => Enabled(
            property.Attributes().FirstOrDefault(attribute => attribute.Name.LocalName == "val")?.Value));
        XElement? fonts = runProperties.Element(XName.Get("rFonts", WordNamespace));
        string? font = fonts?.Attributes().FirstOrDefault(attribute => attribute.Name.LocalName == "ascii")?.Value ??
                       fonts?.Attributes().FirstOrDefault(attribute => attribute.Name.LocalName == "hAnsi")?.Value;
        return CreatePresentation(On("b"), On("i"), Val("color"), Val("sz"), font, Val("vertAlign"),
            Val("lang"), On("rtl"));
    }

    private static OmmlRunPresentation? ParseControlPresentation(OpenXmlElement element)
    {
        OpenXmlElement? properties = element.ChildElements.FirstOrDefault(child =>
            child.NamespaceUri == MathNamespace && child.LocalName.EndsWith("Pr", StringComparison.Ordinal));
        OpenXmlElement? control = properties?.Descendants().FirstOrDefault(child =>
            child.NamespaceUri == MathNamespace && child.LocalName == "ctrlPr");
        OpenXmlElement? runProperties = control?.ChildElements.FirstOrDefault(child =>
            child.NamespaceUri == WordNamespace && child.LocalName == "rPr") ??
            control?.ChildElements.FirstOrDefault(child => child.NamespaceUri == WordNamespace &&
                child.LocalName is "ins" or "del" or "moveTo" or "moveFrom")?
                .Descendants().FirstOrDefault(child => child.NamespaceUri == WordNamespace && child.LocalName == "rPr");
        if (runProperties == null) return null;
        string? Val(string name) => Attribute(runProperties.ChildElements.FirstOrDefault(child =>
            child.NamespaceUri == WordNamespace && child.LocalName == name), "val");
        bool On(string name)
        {
            OpenXmlElement? property = runProperties.ChildElements.FirstOrDefault(child =>
                child.NamespaceUri == WordNamespace && child.LocalName == name);
            return property != null && Enabled(Attribute(property, "val"));
        }
        OpenXmlElement? fonts = runProperties.ChildElements.FirstOrDefault(child =>
            child.NamespaceUri == WordNamespace && child.LocalName == "rFonts");
        return CreatePresentation(On("b"), On("i"), Val("color"), Val("sz"),
            Attribute(fonts, "ascii") ?? Attribute(fonts, "hAnsi"), Val("vertAlign"), Val("lang"), On("rtl"));
    }

    private static OmmlRunPresentation CreatePresentation(bool bold, bool italic, string? color,
        string? halfPointSize, string? font, string? verticalAlignment, string? language, bool rightToLeft) =>
        new(bold, italic, NormalizeColor(color),
            uint.TryParse(halfPointSize, out uint halfPoints) ? halfPoints / 2.0 : null,
            font,
            verticalAlignment switch
            {
                "superscript" => OmmlRunVerticalAlignment.Superscript,
                "subscript" => OmmlRunVerticalAlignment.Subscript,
                _ => OmmlRunVerticalAlignment.Baseline,
            },
            language, rightToLeft);

    private static OmmlRevisionKind? ParseControlRevision(XElement element)
    {
        XElement? revision = element.Descendants(XName.Get("ctrlPr", MathNamespace)).FirstOrDefault()?
            .Elements().FirstOrDefault(child => child.Name.NamespaceName == WordNamespace &&
                child.Name.LocalName is "ins" or "del" or "moveTo" or "moveFrom");
        return revision?.Name.LocalName switch
        {
            "ins" or "moveTo" => OmmlRevisionKind.Inserted,
            "del" or "moveFrom" => OmmlRevisionKind.Deleted,
            _ => null,
        };
    }

    private static OmmlRevisionKind? ParseControlRevision(OpenXmlElement element)
    {
        OpenXmlElement? revision = element.Descendants().FirstOrDefault(child =>
            child.NamespaceUri == MathNamespace && child.LocalName == "ctrlPr")?
            .ChildElements.FirstOrDefault(child => child.NamespaceUri == WordNamespace &&
                child.LocalName is "ins" or "del" or "moveTo" or "moveFrom");
        return revision?.LocalName switch
        {
            "ins" or "moveTo" => OmmlRevisionKind.Inserted,
            "del" or "moveFrom" => OmmlRevisionKind.Deleted,
            _ => null,
        };
    }

    private static IReadOnlyList<OmmlToken> ClassifyRunText(string text, bool textMode)
    {
        List<OmmlToken> result = new();
        string[] segments = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        for (int i = 0; i < segments.Length; i++)
        {
            if (i != 0) result.Add(new OmmlToken(OmmlTokenKind.LineBreak, "\n"));
            if (segments[i].Length != 0 || segments.Length == 1)
                result.AddRange(OmmlTokenClassifier.Classify(segments[i], textMode));
        }
        return result;
    }

    private static int? RunBreakAlignment(XElement run)
    {
        XElement? properties = MathChild(run, "rPr");
        XElement? manualBreak = properties == null ? null : MathChild(properties, "brk");
        return manualBreak == null ? null : ParseInteger((string?)manualBreak.Attribute("alnAt") ??
            (string?)manualBreak.Attribute(XName.Get("alnAt", MathNamespace)), 0, 0, 255);
    }

    private static int? RunBreakAlignment(OpenXmlElement run)
    {
        OpenXmlElement? properties = MathChild(run, "rPr");
        OpenXmlElement? manualBreak = properties == null ? null : MathChild(properties, "brk");
        return manualBreak == null ? null : ParseInteger(Attribute(manualBreak, "alnAt"), 0, 0, 255);
    }

    private static DxpOmmlJustification? ParagraphJustification(XElement paragraph)
    {
        XElement? properties = MathChild(paragraph, "oMathParaPr");
        XElement? value = properties == null ? null : MathChild(properties, "jc");
        return value == null ? null : Justification((string?)value.Attribute(XName.Get("val", MathNamespace)));
    }

    private static DxpOmmlJustification? ParagraphJustification(OpenXmlElement paragraph)
    {
        OpenXmlElement? properties = MathChild(paragraph, "oMathParaPr");
        OpenXmlElement? value = properties == null ? null : MathChild(properties, "jc");
        return value == null ? null : Justification(Attribute(value, "val"));
    }

    private static DxpOmmlJustification Justification(string? value) => value switch
    {
        "left" => DxpOmmlJustification.Left,
        "right" => DxpOmmlJustification.Right,
        "center" => DxpOmmlJustification.Center,
        _ => DxpOmmlJustification.CenterGroup,
    };

    private static bool IsWordBreak(XElement element) => element.Name.NamespaceName == WordNamespace &&
        element.Name.LocalName is "br" or "cr";
    private static bool IsWordBreak(OpenXmlElement element) => element.NamespaceUri == WordNamespace &&
        element.LocalName is "br" or "cr";

    private static bool Enabled(string? value) => value == null ||
        !(value == "0" || value.Equals("false", StringComparison.OrdinalIgnoreCase) ||
          value.Equals("off", StringComparison.OrdinalIgnoreCase));

    private static string? Attribute(OpenXmlElement? element, string localName) =>
        element?.GetAttributes().FirstOrDefault(a => a.LocalName == localName).Value;

    private static string ExtractRunText(XElement run)
    {
        StringBuilder result = new();
        XElement? fonts = run.Descendants().FirstOrDefault(e => e.Name.NamespaceName == WordNamespace && e.Name.LocalName == "rFonts");
        string? font = fonts == null ? null : (string?)fonts.Attribute(XName.Get("ascii", WordNamespace)) ?? (string?)fonts.Attribute(XName.Get("hAnsi", WordNamespace));
        foreach (XElement e in run.Descendants())
        {
            if ((e.Name.NamespaceName == MathNamespace || e.Name.NamespaceName == WordNamespace) && e.Name.LocalName == "t") result.Append(global::DocxportNet.DxpFontSymbols.Substitute(font, e.Value));
            else if (e.Name.NamespaceName == WordNamespace && e.Name.LocalName == "tab") result.Append('\t');
            else if (e.Name.NamespaceName == WordNamespace && e.Name.LocalName is "br" or "cr") result.Append('\n');
            else if (e.Name.NamespaceName == WordNamespace && e.Name.LocalName == "noBreakHyphen") result.Append('-');
            else if (e.Name.NamespaceName == WordNamespace && e.Name.LocalName == "softHyphen") result.Append('\u00AD');
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
            else if (e.NamespaceUri == WordNamespace && e.LocalName is "br" or "cr") result.Append('\n');
            else if (e.NamespaceUri == WordNamespace && e.LocalName == "noBreakHyphen") result.Append('-');
            else if (e.NamespaceUri == WordNamespace && e.LocalName == "softHyphen") result.Append('\u00AD');
            else if (e.NamespaceUri == WordNamespace && e.LocalName == "sym") result.Append(global::DocxportNet.DxpFontSymbols.TranslateWordSymbol(Attribute(e, "font"), Attribute(e, "char")));
        }
        return result.ToString();
    }

    private static OmmlUnsupported ParseUnsupported(XElement element, string path) =>
        new(path, QualifiedName(element), ExtractVisibleText(element), xmlElement: new XElement(element));

    private static OmmlUnsupported ParseUnsupported(OpenXmlElement element, string path) =>
        new(path, QualifiedName(element), ExtractVisibleText(element), openXmlElement: element.CloneNode(true));

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
            else if (descendant.Name.NamespaceName == WordNamespace && descendant.Name.LocalName is "br" or "cr")
            {
                result.Write('\n');
            }
            else if (descendant.Name.NamespaceName == WordNamespace && descendant.Name.LocalName == "noBreakHyphen") result.Write('-');
            else if (descendant.Name.NamespaceName == WordNamespace && descendant.Name.LocalName == "softHyphen") result.Write('\u00AD');
            else if (descendant.Name.NamespaceName == WordNamespace && descendant.Name.LocalName == "sym")
                result.Write(global::DocxportNet.DxpFontSymbols.TranslateWordSymbol(
                    (string?)descendant.Attribute(XName.Get("font", WordNamespace)),
                    (string?)descendant.Attribute(XName.Get("char", WordNamespace))));
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
        if (element.NamespaceUri == WordNamespace && element.LocalName is "br" or "cr")
        {
            result.Write('\n');
            return;
        }
        if (element.NamespaceUri == WordNamespace && element.LocalName == "noBreakHyphen") { result.Write('-'); return; }
        if (element.NamespaceUri == WordNamespace && element.LocalName == "softHyphen") { result.Write('\u00AD'); return; }
        if (element.NamespaceUri == WordNamespace && element.LocalName == "sym")
        {
            result.Write(global::DocxportNet.DxpFontSymbols.TranslateWordSymbol(Attribute(element, "font"), Attribute(element, "char")));
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
