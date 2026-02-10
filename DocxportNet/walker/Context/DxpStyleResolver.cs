using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Word;
using System.Globalization;
using System.Xml.Linq;

namespace DocxportNet.Walker.Context;


public struct DxpEffectiveRunStyleBuilder
{
    public bool? Bold;
    public bool? Italic;
    public bool? Underline;
    public bool? Strike;
    public bool? DoubleStrike;
    public bool? Superscript;
    public bool? Subscript;
    public bool? AllCaps;
    public bool? SmallCaps;
    public string? FontName;
    public int? FontSizeHalfPoints;

    public DxpStyleEffectiveRunStyle ToImmutable() => new(
        Bold ?? false,
        Italic ?? false,
        Underline ?? false,
        Strike ?? false,
        DoubleStrike ?? false,
        Superscript ?? false,
        Subscript ?? false,
        AllCaps ?? false,
        SmallCaps ?? false,
        FontName,
        FontSizeHalfPoints
    );

    public static DxpEffectiveRunStyleBuilder FromDefaults(RunPropertiesBaseStyle? defaults, Func<string, string?>? themeFontResolver = null)
    {
        var acc = new DxpEffectiveRunStyleBuilder();
        ApplyRunPropertiesBaseStyle(defaults, themeFontResolver, ref acc);
        return acc;
    }

    public static void ApplyStyleRunProperties(StyleRunProperties? rp, Func<string, string?>? themeFontResolver, ref DxpEffectiveRunStyleBuilder acc)
    {
        if (rp == null)
            return;

        ApplyStyle(
            rp.Bold, rp.Italic, rp.Underline, rp.Strike, rp.DoubleStrike,
            rp.VerticalTextAlignment, rp.RunFonts, rp.FontSize,
            rp.Caps, rp.SmallCaps,
            themeFontResolver,
            ref acc
        );
    }

    public static void ApplyRunPropertiesBaseStyle(RunPropertiesBaseStyle? rp, Func<string, string?>? themeFontResolver, ref DxpEffectiveRunStyleBuilder acc)
    {
        if (rp == null)
            return;

        ApplyStyle(
            rp.Bold, rp.Italic, rp.Underline, rp.Strike, rp.DoubleStrike,
            rp.VerticalTextAlignment, rp.RunFonts, rp.FontSize,
            rp.Caps, rp.SmallCaps,
            themeFontResolver,
            ref acc
        );
    }

    public static void ApplyRunProperties(RunProperties? rp, Func<string, string?>? themeFontResolver, ref DxpEffectiveRunStyleBuilder acc)
    {
        if (rp == null)
            return;

        ApplyStyle(
            rp.Bold, rp.Italic, rp.Underline, rp.Strike, rp.DoubleStrike,
            rp.VerticalTextAlignment, rp.RunFonts, rp.FontSize,
            rp.Caps, rp.SmallCaps,
            themeFontResolver,
            ref acc
        );
    }


    private static void ApplyStyle(
        Bold? bold,
        Italic? italic,
        Underline? underline,
        Strike? strike,
        DoubleStrike? doubleStrike,
        VerticalTextAlignment? vAlign,
        RunFonts? fonts,
        FontSize? fontSize,
        Caps? caps,
        SmallCaps? smallCaps,
        Func<string, string?>? themeFontResolver,
    ref DxpEffectiveRunStyleBuilder acc)
    {
        if (bold != null)
            acc.Bold = IsOn(bold.Val);
        if (italic != null)
            acc.Italic = IsOn(italic.Val);

        if (underline != null)
            acc.Underline = underline.Val != null && underline.Val != UnderlineValues.None;

        if (strike != null)
            acc.Strike = IsOn(strike.Val);
        if (doubleStrike != null)
            acc.DoubleStrike = IsOn(doubleStrike.Val);

        if (vAlign != null)
        {
            var v = vAlign.Val?.Value;
            acc.Superscript = v == VerticalPositionValues.Superscript;
            acc.Subscript = v == VerticalPositionValues.Subscript;
        }

        if (caps != null)
            acc.AllCaps = IsOn(caps.Val);
        if (smallCaps != null)
            acc.SmallCaps = IsOn(smallCaps.Val);

        if (fonts != null)
        {
            string? asciiTheme = TryGetAttributeValue(fonts, "asciiTheme");
            string? highAnsiTheme = TryGetAttributeValue(fonts, "hAnsiTheme");

            string? resolvedFont =
                fonts.Ascii?.Value
                ?? fonts.HighAnsi?.Value
                ?? TryResolveThemeFont(asciiTheme, themeFontResolver)
                ?? TryResolveThemeFont(highAnsiTheme, themeFontResolver);

            if (!string.IsNullOrWhiteSpace(resolvedFont))
                acc.FontName = resolvedFont;
        }
        if (fontSize?.Val?.Value != null && int.TryParse(fontSize.Val.Value, out var hp))
            acc.FontSizeHalfPoints = hp;
    }

    private static string? TryGetAttributeValue(OpenXmlElement el, string localName)
    {
        foreach (var attr in el.GetAttributes())
        {
            if (string.Equals(attr.LocalName, localName, StringComparison.OrdinalIgnoreCase))
                return attr.Value;
        }

        // Some SDKs may not surface these as attributes; fall back to parsing the element XML.
        var xml = el.OuterXml;
        var needle = localName + "=\"";
        int start = xml.IndexOf(needle, StringComparison.OrdinalIgnoreCase);
        if (start < 0)
            return null;
        start += needle.Length;
        int end = xml.IndexOf('"', start);
        if (end < 0)
            return null;
        return xml.Substring(start, end - start);
    }

    private static string? TryResolveThemeFont(string? theme, Func<string, string?>? themeFontResolver)
    {
        if (string.IsNullOrWhiteSpace(theme) || themeFontResolver == null)
            return null;
        return themeFontResolver(theme!);
    }


    private static bool IsOn(OnOffValue? v)
        => v == null || v.Value; // in WordprocessingML, missing val often means "on"
}


public sealed class DxpStyleResolver : DxpIStyleResolver
{
    private readonly Styles? _styles;
    private readonly Dictionary<string, Style> _byId;

    private readonly RunPropertiesBaseStyle? _docDefaultRunProps;
    private readonly ParagraphPropertiesBaseStyle? _docDefaultParaProps;
    private readonly string? _themeMajorLatinFont;
    private readonly string? _themeMinorLatinFont;

    public DxpStyleEffectiveIndentTwips GetIndentation(
    Paragraph p,
    DxpNumberingResolver? nr = null)
    {
        var acc = new IndentAcc();

        // 1) Document defaults (paragraph defaults)
        ApplyIndentation(_docDefaultParaProps?.Indentation, ref acc);

        // 2) Paragraph style chain (base -> ... -> direct style)
        var pStyleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        foreach (var style in EnumerateStyleChainRaw(pStyleId).Reverse())
        {
            var ind = style.StyleParagraphProperties?.Indentation;
            ApplyIndentation(ind, ref acc);
        }

        // 3) Numbering level indentation (if paragraph is in a list)
        if (nr != null)
        {
            var np = ResolveEffectiveNumPr(p);
            if (np != null)
            {
                var resolved = nr.ResolveLevel(np.NumId, np.Ilvl);
                if (resolved != null)
                {
                    var lvlInd = resolved.Value.lvl.PreviousParagraphProperties?.Indentation;
                    ApplyIndentation(lvlInd, ref acc);
                }
            }
        }

        // 4) Direct paragraph indentation (highest precedence)
        ApplyIndentation(p.ParagraphProperties?.Indentation, ref acc);

        return acc.ToImmutable();
    }

    // ---------------- helpers ----------------

    private struct IndentAcc
    {
        public int? Left;
        public int? Right;
        public int? FirstLine;
        public int? Hanging;

        public DxpStyleEffectiveIndentTwips ToImmutable() => new(Left, Right, FirstLine, Hanging);
    }

    private static int? ReadTwips(StringValue? v)
    {
        if (v?.Value == null)
            return null;
        return int.TryParse(v.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : null;
    }

    private static void ApplyIndentation(Indentation? ind, ref IndentAcc acc)
    {
        if (ind == null)
            return;

        int? left = ReadTwips(ind.Left) ?? ReadTwips(ind.Start);
        int? right = ReadTwips(ind.Right) ?? ReadTwips(ind.End);

        if (left != null)
            acc.Left = left;
        if (right != null)
            acc.Right = right;

        var firstLine = ReadTwips(ind.FirstLine);
        var hanging = ReadTwips(ind.Hanging);

        if (firstLine != null)
            acc.FirstLine = firstLine;
        if (hanging != null)
            acc.Hanging = hanging;
    }




    public DxpStyleResolver(WordprocessingDocument doc)
    {
        _styles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        _byId = _styles?
            .Elements<Style>()
            .Where(s => s.StyleId?.Value != null)
            .ToDictionary(s => s.StyleId!.Value!, s => s)
            ?? new Dictionary<string, Style>();

        var docDefaults = _styles?.DocDefaults;
        _docDefaultRunProps = docDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
        _docDefaultParaProps = docDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle;

        (_themeMajorLatinFont, _themeMinorLatinFont) = TryGetThemeLatinFonts(doc);
    }

    private static (string? majorLatin, string? minorLatin) TryGetThemeLatinFonts(WordprocessingDocument doc)
    {
        try
        {
            var themePart = doc.MainDocumentPart?.ThemePart;
            if (themePart == null)
                return (null, null);

            using var s = themePart.GetStream(FileMode.Open, FileAccess.Read);
            var xdoc = XDocument.Load(s);

            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            var fontScheme = xdoc.Descendants(a + "fontScheme").FirstOrDefault();
            if (fontScheme == null)
                return (null, null);

            static string? LatinTypeface(XNamespace a, XElement? font)
                => font?.Element(a + "latin")?.Attribute("typeface")?.Value;

            var majorLatin = LatinTypeface(a, fontScheme.Element(a + "majorFont"));
            var minorLatin = LatinTypeface(a, fontScheme.Element(a + "minorFont"));
            return (majorLatin, minorLatin);
        }
        catch
        {
            return (null, null);
        }
    }

    private string? ResolveThemeFont(string themeRef)
    {
        if (string.IsNullOrWhiteSpace(themeRef))
            return null;

        if (themeRef.StartsWith("minor", StringComparison.OrdinalIgnoreCase))
            return string.IsNullOrWhiteSpace(_themeMinorLatinFont) ? null : _themeMinorLatinFont;
        if (themeRef.StartsWith("major", StringComparison.OrdinalIgnoreCase))
            return string.IsNullOrWhiteSpace(_themeMajorLatinFont) ? null : _themeMajorLatinFont;

        return null;
    }

    public DxpStyleInfo? GetStyleInfo(string? styleId)
    {
        if (string.IsNullOrEmpty(styleId))
            return null;
        if (!_byId.TryGetValue(styleId!, out var s))
            return null;

        return new DxpStyleInfo(
            StyleId: styleId!,
            Name: s.StyleName?.Val?.Value,
            Type: s.Type?.Value,
            BasedOnStyleId: s.BasedOn?.Val?.Value
        );
    }

    public IReadOnlyList<DxpStyleInfo> GetStyleChain(string? styleId)
    {
        var result = new List<DxpStyleInfo>();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        var current = styleId;
        while (!string.IsNullOrEmpty(current) && seen.Add(current!))
        {
            var info = GetStyleInfo(current);
            if (info == null)
                break;

            result.Add(info);
            current = info.BasedOnStyleId;
        }

        return result; // [direct, parent, grandparent, ...]
    }

    public IReadOnlyList<DxpStyleInfo> GetParagraphStyleChain(Paragraph p)
    {
        var pStyleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        return GetStyleChain(pStyleId);
    }


    public DxpStyleEffectiveRunStyle ResolveRunStyle(Paragraph p, Run r)
    {
        // 1) Start with doc defaults
        var acc = DxpEffectiveRunStyleBuilder.FromDefaults(_docDefaultRunProps, ResolveThemeFont);

        // 2) Apply paragraph style chain (paragraph style's rPr affects runs)
        var pStyleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (string.IsNullOrWhiteSpace(pStyleId))
            pStyleId = DxpWordBuiltInStyleId.wdStyleNormal;
        ApplyParagraphStyleChainRunProps(pStyleId, ref acc);

        // 3) Apply direct paragraph run properties (pPr/rPr)
        var paraRunProps = p.ParagraphProperties?.GetFirstChild<RunProperties>();
        DxpEffectiveRunStyleBuilder.ApplyRunProperties(paraRunProps, ResolveThemeFont, ref acc);

        // 4) Apply character style chain (rStyle)
        var rStyleId = r.RunProperties?.RunStyle?.Val?.Value;
        ApplyCharacterStyleChainRunProps(rStyleId, ref acc);

        // 5) Apply direct run formatting (highest precedence)
        DxpEffectiveRunStyleBuilder.ApplyRunProperties(r.RunProperties, ResolveThemeFont, ref acc);

        return acc.ToImmutable();
    }

    public string? ResolveRunLanguage(Paragraph p, Run r)
    {
        // Highest precedence: direct run lang
        var lang = r.RunProperties?.Languages?.Val?.Value;
        if (!string.IsNullOrEmpty(lang))
            return lang;

        // Paragraph style chain (paragraph and character styles)
        string? fromStyles = ResolveLangFromStyles(p.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
        if (!string.IsNullOrEmpty(fromStyles))
            return fromStyles;

        var runStyleId = r.RunProperties?.RunStyle?.Val?.Value;
        fromStyles = ResolveLangFromStyles(runStyleId);
        if (!string.IsNullOrEmpty(fromStyles))
            return fromStyles;

        // Document defaults
        return _docDefaultRunProps?.Languages?.Val?.Value;
    }

    public string? GetDefaultLanguage()
    {
        return _docDefaultRunProps?.Languages?.Val?.Value;
    }

    public string? ResolveParagraphLanguage(Paragraph p)
    {
        // Highest precedence: paragraph mark run properties language
        var lang = p.ParagraphProperties?.GetFirstChild<ParagraphMarkRunProperties>()
            ?.GetFirstChild<Languages>()?.Val?.Value;
        if (!string.IsNullOrEmpty(lang))
            return lang;

        // Paragraph style chain (paragraph and character styles)
        string? fromStyles = ResolveLangFromStyles(p.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
        if (!string.IsNullOrEmpty(fromStyles))
            return fromStyles;

        // Document defaults
        return _docDefaultRunProps?.Languages?.Val?.Value;
    }

    private string? ResolveLangFromStyles(string? styleId)
    {
        foreach (var style in EnumerateStyleChain(styleId))
        {
            var lang = style.StyleRunProperties?.Languages?.Val?.Value;
            if (lang == null)
            {
                var rp = style.StyleParagraphProperties?.GetFirstChild<ParagraphMarkRunProperties>();
                lang = rp?.GetFirstChild<Languages>()?.Val?.Value;
            }
            if (!string.IsNullOrEmpty(lang))
                return lang;
        }

        return null;
    }

    private void ApplyParagraphStyleChainRunProps(string? styleId, ref DxpEffectiveRunStyleBuilder acc)
    {
        foreach (var style in EnumerateStyleChain(styleId).Reverse())
        {
            DxpEffectiveRunStyleBuilder.ApplyStyleRunProperties(style.StyleRunProperties, ResolveThemeFont, ref acc);
            var rp = style.StyleParagraphProperties?.GetFirstChild<RunProperties>();
            DxpEffectiveRunStyleBuilder.ApplyRunProperties(rp, ResolveThemeFont, ref acc);
        }
    }

    private void ApplyCharacterStyleChainRunProps(string? styleId, ref DxpEffectiveRunStyleBuilder acc)
    {
        foreach (var style in EnumerateStyleChain(styleId).Reverse())
            DxpEffectiveRunStyleBuilder.ApplyStyleRunProperties(style.StyleRunProperties, ResolveThemeFont, ref acc);
    }

    private IEnumerable<Style> EnumerateStyleChain(string? styleId)
    {
        // Walk basedOn chain, starting from styleId, stopping on cycles or missing.
        if (string.IsNullOrEmpty(styleId))
            yield break;

        var seen = new HashSet<string>(StringComparer.Ordinal);

        var current = styleId;
        while (!string.IsNullOrEmpty(current) && seen.Add(current!))
        {
            if (!_byId.TryGetValue(current!, out var style))
                yield break;

            yield return style;

            current = style.BasedOn?.Val?.Value;
        }
    }


    public int? GetOutlineLevel(Paragraph p)
    {
        // Direct formatting on the paragraph (highest precedence)
        var direct = p.ParagraphProperties?.OutlineLevel?.Val?.Value;
        if (direct != null)
            return direct; // 0-based

        // From style chain
        var pStyleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        foreach (var s in EnumerateStyleChainRaw(pStyleId)) // Style objects
        {
            var lvl = s.StyleParagraphProperties?.OutlineLevel?.Val?.Value;
            if (lvl != null)
                return lvl;
        }
        return null;
    }

    public JustificationValues? GetJustification(Paragraph p)
    {
        // Direct formatting on the paragraph (highest precedence)
        var direct = p.ParagraphProperties?.Justification?.Val?.Value;
        if (direct != null)
            return direct;

        // From style chain
        var pStyleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        foreach (var s in EnumerateStyleChainRaw(pStyleId))
        {
            var jc = s.StyleParagraphProperties?.Justification?.Val?.Value;
            if (jc != null)
                return jc;
        }

        // Document defaults
        return _docDefaultParaProps?.Justification?.Val?.Value;
    }

    public ParagraphBorders? GetParagraphBorders(Paragraph p)
    {
        var direct = p.ParagraphProperties?.ParagraphBorders;

        // Resolve per-side, so partially-defined borders in styles/base styles still apply.
        var pStyleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

        TopBorder? top = ResolveBorderSide(direct, pStyleId, b => b.TopBorder);
        RightBorder? right = ResolveBorderSide(direct, pStyleId, b => b.RightBorder);
        BottomBorder? bottom = ResolveBorderSide(direct, pStyleId, b => b.BottomBorder);
        LeftBorder? left = ResolveBorderSide(direct, pStyleId, b => b.LeftBorder);

        if (top == null && right == null && bottom == null && left == null)
            return null;

        return new ParagraphBorders {
            TopBorder = top,
            RightBorder = right,
            BottomBorder = bottom,
            LeftBorder = left
        };
    }

    public Shading? GetParagraphShading(Paragraph p)
    {
        var direct = p.ParagraphProperties?.Shading;
        if (HasMeaningfulShadingFill(direct))
            return direct;

        var pStyleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        foreach (var s in EnumerateStyleChainRaw(pStyleId))
        {
            var shd = s.StyleParagraphProperties?.Shading;
            if (HasMeaningfulShadingFill(shd))
                return shd;
        }

        var defaults = _docDefaultParaProps?.Shading;
        return HasMeaningfulShadingFill(defaults) ? defaults : null;
    }

    public Tabs? GetParagraphTabs(Paragraph p)
    {
        var direct = p.ParagraphProperties?.Tabs;
        if (direct != null)
            return direct;

        var pStyleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        foreach (var s in EnumerateStyleChainRaw(pStyleId))
        {
            var tabs = s.StyleParagraphProperties?.Tabs;
            if (tabs != null)
                return tabs;
        }

        return _docDefaultParaProps?.Tabs;
    }

    public SpacingBetweenLines? GetParagraphSpacing(Paragraph p)
    {
        var direct = p.ParagraphProperties?.SpacingBetweenLines;
        if (direct != null)
            return direct;

        var pStyleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        foreach (var s in EnumerateStyleChainRaw(pStyleId))
        {
            var spacing = s.StyleParagraphProperties?.SpacingBetweenLines;
            if (spacing != null)
                return spacing;
        }

        // Do not apply document defaults here; most output already has sensible CSS defaults and
        // surfacing doc defaults would introduce pervasive wrappers in "plain" visitors.
        return null;
    }

    public bool GetContextualSpacing(Paragraph p)
    {
        var direct = p.ParagraphProperties?.ContextualSpacing;
        if (direct != null)
            return direct.Val == null || direct.Val.Value;

        var pStyleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        foreach (var s in EnumerateStyleChainRaw(pStyleId))
        {
            var cs = s.StyleParagraphProperties?.ContextualSpacing;
            if (cs != null)
                return cs.Val == null || cs.Val.Value;
        }

        return false;
    }

    private static bool HasMeaningfulShadingFill(Shading? shd)
    {
        var fill = shd?.Fill?.Value;
        if (string.IsNullOrWhiteSpace(fill))
            return false;
        if (string.Equals(fill, "auto", StringComparison.OrdinalIgnoreCase))
            return false;
        return true;
    }

    private TBorder? ResolveBorderSide<TBorder>(
        ParagraphBorders? direct,
        string? pStyleId,
        Func<ParagraphBorders, TBorder?> selector)
        where TBorder : OpenXmlElement
    {
        var directSide = direct != null ? selector(direct) : null;
        if (directSide != null)
            return (TBorder)directSide.CloneNode(true);

        foreach (var s in EnumerateStyleChainRaw(pStyleId))
        {
            var b = s.StyleParagraphProperties?.ParagraphBorders;
            if (b == null)
                continue;

            var side = selector(b);
            if (side != null)
                return (TBorder)side.CloneNode(true);
        }

        var def = _docDefaultParaProps?.ParagraphBorders;
        var defSide = def != null ? selector(def) : null;
        return defSide != null ? (TBorder)defSide.CloneNode(true) : null;
    }

    // helper: raw style chain (same logic as EnumerateStyleChain)
    private IEnumerable<Style> EnumerateStyleChainRaw(string? styleId)
    {
        if (string.IsNullOrEmpty(styleId))
            yield break;
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var current = styleId;

        while (!string.IsNullOrEmpty(current) && seen.Add(current!))
        {
            if (!_byId.TryGetValue(current!, out var style))
                yield break;
            yield return style;
            current = style.BasedOn?.Val?.Value;
        }
    }

    public int? GetHeadingLevel(Paragraph p)
    {
        // OutlineLvl is 0-based; convert to 1-based heading level
        var outline = GetOutlineLevel(p);
        if (outline is >= 0 and <= 8)
            return outline.Value + 2; // shift by 1 so Title can stay at level 1

        // Fallback: name/id heuristics
        var pStyleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        return GetHeadingLevelFromStyleChain(pStyleId);
    }

    public static int? TryGetHeadingLevelFromStyleNameOrId(DxpStyleInfo s)
    {
        // StyleName examples: "Heading 1", "Heading 2"
        // StyleId examples: "Heading1", "Heading2"
        static int? ParseLevel(string? x)
        {
            if (string.IsNullOrWhiteSpace(x))
                return null;

            x = x?.Trim();

            if (string.IsNullOrWhiteSpace(x))
                return null;

            if (x!.Equals(DxpWordBuiltInStyleId.wdStyleTitle, StringComparison.OrdinalIgnoreCase))
                return 1;

            if (x.Equals(DxpWordBuiltInStyleId.wdStyleHeading1, StringComparison.OrdinalIgnoreCase))
                return 2;
            if (x.Equals(DxpWordBuiltInStyleId.wdStyleHeading2, StringComparison.OrdinalIgnoreCase))
                return 3;
            if (x.Equals(DxpWordBuiltInStyleId.wdStyleHeading3, StringComparison.OrdinalIgnoreCase))
                return 4;
            if (x.Equals(DxpWordBuiltInStyleId.wdStyleHeading4, StringComparison.OrdinalIgnoreCase))
                return 5;
            if (x.Equals(DxpWordBuiltInStyleId.wdStyleHeading5, StringComparison.OrdinalIgnoreCase))
                return 6;
            if (x.Equals(DxpWordBuiltInStyleId.wdStyleHeading6, StringComparison.OrdinalIgnoreCase))
                return 7;
            if (x.Equals(DxpWordBuiltInStyleId.wdStyleHeading7, StringComparison.OrdinalIgnoreCase))
                return 8;
            if (x.Equals(DxpWordBuiltInStyleId.wdStyleHeading8, StringComparison.OrdinalIgnoreCase))
                return 9;
            if (x.Equals(DxpWordBuiltInStyleId.wdStyleHeading9, StringComparison.OrdinalIgnoreCase))
                return 10;

            if (x.StartsWith("Heading ", StringComparison.OrdinalIgnoreCase) &&
                int.TryParse(x.Substring("Heading ".Length), out var n1))
                return n1 + 1;

            if (x.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) &&
                int.TryParse(x.Substring("Heading".Length), out var n2))
                return n2 + 1;

            return null;
        }

        return ParseLevel(s.Name) ?? ParseLevel(s.StyleId);
    }

    public int? GetHeadingLevelFromStyleChain(string? pStyleId)
    {
        foreach (var s in GetStyleChain(pStyleId))
        {
            var lvl = TryGetHeadingLevelFromStyleNameOrId(s);
            if (lvl is >= 1 and <= 9)
                return lvl;
        }
        return null;
    }

    public DxpStyleEffectiveRunStyle GetDefaultRunStyle()
    {
        var acc = DxpEffectiveRunStyleBuilder.FromDefaults(_docDefaultRunProps, ResolveThemeFont);
        ApplyParagraphStyleChainRunProps(DxpWordBuiltInStyleId.wdStyleNormal, ref acc);
        return acc.ToImmutable();
    }

    public DxpStyleEffectiveNumPr? ResolveEffectiveNumPr(Paragraph p)
    {
        // 1) direct pPr wins, including explicit "no numbering" (numId=0)
        var directNp = p.ParagraphProperties?.NumberingProperties;
        if (directNp != null)
        {
            var directNumId = directNp.NumberingId?.Val?.Value;
            if (directNumId != null)
            {
                if (directNumId.Value == 0)
                    return null; // explicit suppression: do NOT consult styles

                var directIlvl = directNp.NumberingLevelReference?.Val?.Value ?? 0;
                return new DxpStyleEffectiveNumPr(directNumId.Value, directIlvl);
            }

            // If there's a direct numPr but no numId, Word still considers it "direct";
            // in that weird case we fall through to styles.
        }

        // 2) from style chain (closest style wins)
        var pStyleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

        int? numId = null;
        int? ilvl = null;

        foreach (var style in EnumerateStyleChainRaw(pStyleId)) // direct -> parent -> ...
        {
            var np = style.StyleParagraphProperties?.NumberingProperties;
            if (np == null)
                continue;

            var sid = np.NumberingId?.Val?.Value;
            if (sid != null)
            {
                // If a style explicitly sets numId=0, treat as "no numbering" too.
                if (sid.Value == 0)
                    return null;

                if (numId == null)
                    numId = sid.Value;
            }

            if (ilvl == null)
                ilvl = np.NumberingLevelReference?.Val?.Value;

            if (numId != null)
                break;
        }

        if (numId == null || numId.Value == 0)
            return null;

        return new DxpStyleEffectiveNumPr(numId.Value, ilvl ?? 0);
    }
}
