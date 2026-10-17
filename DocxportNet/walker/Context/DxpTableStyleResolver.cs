using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;

namespace DocxportNet.Walker.Context;

internal sealed class DxpTableStyleResolver
{
    private readonly Dictionary<string, Style> _tableStylesById = new(StringComparer.Ordinal);

    public DxpTableStyleResolver(WordprocessingDocument doc)
    {
        var styles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles == null)
            return;

        foreach (var s in styles.Elements<Style>())
        {
            var id = s.StyleId?.Value;
            if (string.IsNullOrEmpty(id))
                continue;

            var type = s.Type?.Value;
            if (type == StyleValues.Table)
                _tableStylesById[id!] = s;
        }
    }

    public DxpResolvedTableStyle ResolveTableStyle(TableProperties? tableProperties, int rowCount, int colCount)
    {
        var look = ParseTblLook(tableProperties?.TableLook);
        string? styleId = tableProperties?.TableStyle?.Val?.Value;
        var styleChain = ResolveTableStyleChain(styleId);
        var direction = tableProperties?.GetFirstChild<BiDiVisual>() ??
            styleChain.Select(style => style.GetFirstChild<StyleTableProperties>()?.GetFirstChild<BiDiVisual>())
                .LastOrDefault(value => value != null);
        var rightToLeft = direction != null && direction.Val?.Value != OnOffOnlyValues.Off;

        var baseTableStyle = ComputeBaseTableStyle(styleChain, tableProperties, rightToLeft);
        var baseCell = ComputeBaseCellStyleFromStyles(styleChain, baseTableStyle,
            tableProperties, rightToLeft);

        return new DxpResolvedTableStyle(styleId, look, rowCount, colCount, styleChain,
            baseTableStyle, baseCell, rightToLeft);
    }

    public static DxpComputedTableCellStyle ComputeCellStyle(DxpResolvedTableStyle resolved, TableCellProperties? cellProperties, int rowIndex, int colIndex, int rowSpan, int colSpan)
    {
        var tableStyle = resolved.ComputedTableStyle;
        var baseCell = resolved.BaseCellStyle;

        var computedBorders = ResolveInitialBorders(resolved, baseCell, tableStyle, rowIndex, colIndex, rowSpan, colSpan);
        var computedBorder = CollapseToSingleBorder(computedBorders) ?? baseCell.Border ?? tableStyle.DefaultCellBorder;
        var computedBackground = baseCell.BackgroundColorCss;
        var computedVAlign = baseCell.VerticalAlign;
        string? conditionalTextColor = null;
        bool? conditionalBold = null;
        bool? conditionalItalic = null;
        bool? conditionalUnderline = null;
        DxpComputedTextAlign? conditionalTextAlign = null;

        // Apply conditional formatting (tblStylePr) from the table style chain.
        foreach (var type in resolved.GetApplicableConditionalTypes(rowIndex, colIndex))
        {
            var ov = ComputeConditionalOverride(resolved.StyleChain, type, resolved.RightToLeft);
            if (ov.BordersSet && ov.Borders != null)
            {
                computedBorders = ApplyBorderOverride(computedBorders, ov.Borders);
                computedBorder = CollapseToSingleBorder(computedBorders) ?? computedBorder;
            }
            else if (ov.BorderSet && ov.Border != null)
            {
                computedBorders = new DxpComputedBoxBorders(ov.Border, ov.Border, ov.Border, ov.Border);
                computedBorder = ov.Border;
            }
            if (ov.BackgroundSet)
                computedBackground = ov.BackgroundColorCss;
            if (ov.VerticalAlignSet)
                computedVAlign = ov.VerticalAlign;
            if (ov.TextColorSet)
                conditionalTextColor = ov.TextColorCss;
            conditionalBold = ov.Bold ?? conditionalBold;
            conditionalItalic = ov.Italic ?? conditionalItalic;
            conditionalUnderline = ov.Underline ?? conditionalUnderline;
            conditionalTextAlign = ov.TextAlign ?? conditionalTextAlign;
        }

        // Direct cell formatting always wins.
        var direct = DxpTableStyleComputer.ComputeDirectCellStyle(cellProperties);
        var directBorders = ComputeDirectCellBorders(cellProperties, resolved.RightToLeft);
        if (directBorders != null)
        {
            computedBorders = ApplyBorderOverride(computedBorders, directBorders);
            computedBorder = CollapseToSingleBorder(computedBorders) ?? computedBorder;
        }
        else if (direct.Border != null)
        {
            computedBorders = new DxpComputedBoxBorders(direct.Border, direct.Border, direct.Border, direct.Border);
            computedBorder = direct.Border;
        }
        if (cellProperties?.Shading?.Val?.Value == ShadingPatternValues.Clear &&
            (cellProperties.Shading.Fill?.Value == null ||
             string.Equals(cellProperties.Shading.Fill.Value, "auto", StringComparison.OrdinalIgnoreCase)))
            computedBackground = null;
        else if (direct.BackgroundColorCss != null)
            computedBackground = direct.BackgroundColorCss;
        if (direct.VerticalAlign != null)
            computedVAlign = direct.VerticalAlign;

        return new DxpComputedTableCellStyle(Border: computedBorder, BackgroundColorCss: computedBackground, VerticalAlign: computedVAlign) {
            Borders = computedBorders,
            ConditionalTextColorCss = conditionalTextColor,
            ConditionalBold = conditionalBold,
            ConditionalItalic = conditionalItalic,
            ConditionalUnderline = conditionalUnderline,
            ConditionalTextAlign = conditionalTextAlign
        };
    }

    private static DxpComputedTableStyle ComputeBaseTableStyle(IReadOnlyList<Style> styleChain,
        TableProperties? directTableProperties, bool rightToLeft)
    {
        // Merge tblBorders across the style chain (base -> derived), then overlay direct formatting.
        var merged = new MergedTableBorders();
        foreach (var s in styleChain)
        {
            var stp = s.GetFirstChild<StyleTableProperties>();
            var borders = stp?.GetFirstChild<TableBorders>();
            merged.Apply(borders, rightToLeft);
        }
        merged.Apply(directTableProperties?.TableBorders, rightToLeft);

        var tableBorders = ComputeTableBoxBorders(merged, out var anyTableBorderSpecified);
        var tableBorder = PickComputedBorder(merged).Border;

        // Keep inside borders on their specified axes. A single picked border must not
        // fill unspecified outer or vertical edges when tblBorders has explicit sides.
        var defaultCellBorders = ComputeDefaultCellBoxBordersFromTable(merged);
        var defaultCellBorder = merged.AnySpecified ? null : tableBorder;

        bool collapse = anyTableBorderSpecified && (
            HasVisibleBorder(tableBorders) ||
            HasVisibleBorder(defaultCellBorders) ||
            (tableBorder != null && tableBorder.LineStyle != DxpComputedBorderLineStyle.None));
        var result = new DxpComputedTableStyle(
            TableBorder: tableBorder,
            BorderCollapse: collapse,
            DefaultCellBorder: defaultCellBorder) {
            TableBorders = tableBorders,
            DefaultCellBorders = defaultCellBorders
        };

        return result;
    }

    private static DxpComputedTableCellStyle ComputeBaseCellStyleFromStyles(IReadOnlyList<Style> styleChain,
        DxpComputedTableStyle tableStyle, TableProperties? directTableProperties, bool rightToLeft)
    {
        var mergedBorders = new MergedCellBorders();
        ResolvedOverride mergedOverrides = default;

        foreach (var s in styleChain)
        {
            if (s.GetFirstChild<StyleTableProperties>() is { } tableProperties)
                mergedOverrides.ApplyCellProperties(tableProperties);
            var stcp = s.GetFirstChild<StyleTableCellProperties>();
            if (stcp != null)
            {
                mergedBorders.Apply(stcp.GetFirstChild<TableCellBorders>(), rightToLeft);
                mergedOverrides.ApplyCellProperties(stcp);
            }
        }
        if (directTableProperties != null)
            mergedOverrides.ApplyCellProperties(directTableProperties);

        var styleBox = ComputeCellBoxBordersFromMerged(mergedBorders);
        var border = styleBox != null ? CollapseToSingleBorder(styleBox) : null;

        return new DxpComputedTableCellStyle(
            Border: border,
            BackgroundColorCss: mergedOverrides.BackgroundSet ? mergedOverrides.BackgroundColorCss : null,
            VerticalAlign: mergedOverrides.VerticalAlignSet ? mergedOverrides.VerticalAlign : null) {
            Borders = styleBox
        };
    }

    private static ResolvedOverride ComputeConditionalOverride(IReadOnlyList<Style> styleChain,
        string type, bool rightToLeft)
    {
        ResolvedOverride result = default;
        var mergedBorders = new MergedCellBorders();
        DxpComputedBoxBorders? rawMerged = null;

        foreach (var s in styleChain)
        {
            foreach (var el in s.ChildElements)
            {
                if (!string.Equals(el.LocalName, "tblStylePr", StringComparison.Ordinal))
                    continue;

                var t = GetAttributeValue(el, "type");
                if (!string.Equals(t, type, StringComparison.OrdinalIgnoreCase))
                    continue;

                var runProperties = FindFirstChildByLocalName(el, "rPr");
                if (runProperties != null)
                {
                    result.Bold = ReadOnOff(runProperties, "b") ?? result.Bold;
                    result.Italic = ReadOnOff(runProperties, "i") ?? result.Italic;
                    var underline = FindFirstChildByLocalName(runProperties, "u");
                    if (underline != null)
                        result.Underline = !string.Equals(GetAttributeValue(underline, "val"),
                            "none", StringComparison.OrdinalIgnoreCase);
                    var color = FindFirstChildByLocalName(runProperties, "color");
                    if (color != null)
                    {
                        var value = GetAttributeValue(color, "val");
                        result.TextColorSet = true;
                        result.TextColorCss = value?.Length == 6 && value.All(Uri.IsHexDigit)
                            ? "#" + value : null;
                    }
                }
                var paragraphProperties = FindFirstChildByLocalName(el, "pPr");
                var justification = paragraphProperties == null ? null :
                    FindFirstChildByLocalName(paragraphProperties, "jc");
                if (justification != null)
                    result.TextAlign = GetAttributeValue(justification, "val") switch
                    {
                        "left" => DxpComputedTextAlign.Left,
                        "center" => DxpComputedTextAlign.Center,
                        "right" => DxpComputedTextAlign.Right,
                        "both" => DxpComputedTextAlign.Justify,
                        _ => null
                    };

                // styles.xml can surface tblStylePr content as unknown elements, so we use LocalName/attributes parsing.
                OpenXmlElement? tcPr =
                    (OpenXmlElement?)el.GetFirstChild<StyleTableCellProperties>() ??
                    el.GetFirstChild<TableCellProperties>() ??
                    FindFirstChildByLocalName(el, "tcPr");

                if (tcPr != null)
                {
                    mergedBorders.Apply(tcPr.GetFirstChild<TableCellBorders>(), rightToLeft);
                    result.ApplyCellProperties(tcPr);

                    var tcBorders = FindFirstChildByLocalName(tcPr, "tcBorders");
                    if (tcBorders != null)
                    {
                        var raw = ParseBoxBordersFromRaw(tcBorders, rightToLeft);
                        if (raw != null)
                            rawMerged = rawMerged == null ? raw : ApplyBorderOverride(rawMerged, raw);
                    }
                }
            }
        }

        var styleBox = ComputeCellBoxBordersFromMerged(mergedBorders);
        if (styleBox == null)
            styleBox = rawMerged;
        else if (rawMerged != null)
            styleBox = ApplyBorderOverride(styleBox, rawMerged);

        result.BordersSet = styleBox != null;
        result.Borders = styleBox;

        var single = styleBox != null ? CollapseToSingleBorder(styleBox) : null;
        if (single != null)
        {
            result.BorderSet = true;
            result.Border = single;
        }

        return result;
    }

    private IReadOnlyList<Style> ResolveTableStyleChain(string? styleId)
    {
        if (string.IsNullOrEmpty(styleId))
            return [];

        var chainDerivedToBase = new List<Style>();
        var visited = new HashSet<string>(StringComparer.Ordinal);

        string? currentId = styleId;
        while (!string.IsNullOrEmpty(currentId) && visited.Add(currentId!))
        {
            if (!_tableStylesById.TryGetValue(currentId!, out var s))
                break;

            chainDerivedToBase.Add(s);
            currentId = GetBasedOnStyleId(s);
        }

        chainDerivedToBase.Reverse(); // base -> derived
        return chainDerivedToBase;
    }

    private static string? GetBasedOnStyleId(Style s)
    {
        var basedOn = s.GetFirstChild<BasedOn>();
        if (basedOn == null)
            return null;

        return basedOn.Val?.Value ?? GetAttributeValue(basedOn, "val");
    }

    private static DxpTableLook ParseTblLook(TableLook? look)
    {
        if (look == null)
            return default;

        return new DxpTableLook(
            FirstRow: GetBoolAttribute(look, "firstRow"),
            LastRow: GetBoolAttribute(look, "lastRow"),
            FirstColumn: GetBoolAttribute(look, "firstColumn"),
            LastColumn: GetBoolAttribute(look, "lastColumn"),
            NoHBand: GetBoolAttribute(look, "noHBand"),
            NoVBand: GetBoolAttribute(look, "noVBand"));
    }

    private static bool GetBoolAttribute(OpenXmlElement el, string localName)
    {
        var v = GetAttributeValue(el, localName);
        if (string.IsNullOrEmpty(v))
            return false;

        if (string.Equals(v, "1", StringComparison.OrdinalIgnoreCase))
            return true;
        if (string.Equals(v, "0", StringComparison.OrdinalIgnoreCase))
            return false;
        if (bool.TryParse(v, out bool b))
            return b;
        return false;
    }

    private static string? GetAttributeValue(OpenXmlElement el, string localName)
    {
        foreach (var a in el.GetAttributes())
        {
            if (string.Equals(a.LocalName, localName, StringComparison.OrdinalIgnoreCase))
                return a.Value;
        }
        return null;
    }

    private static OpenXmlElement? FindFirstChildByLocalName(OpenXmlElement el, string localName)
    {
        foreach (var c in el.ChildElements)
        {
            if (string.Equals(c.LocalName, localName, StringComparison.Ordinal))
                return c;
        }
        return null;
    }

    private static bool? ReadOnOff(OpenXmlElement parent, string localName)
    {
        var element = FindFirstChildByLocalName(parent, localName);
        if (element == null) return null;
        var value = GetAttributeValue(element, "val");
        return !string.Equals(value, "0", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(value, "false", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(value, "off", StringComparison.OrdinalIgnoreCase);
    }

    private static DxpComputedBoxBorders? ParseBoxBordersFromRaw(OpenXmlElement tcBorders,
        bool rightToLeft)
    {
        DxpComputedBorder? top = ParseBorderFromRaw(tcBorders, "top");
        DxpComputedBorder? right = ParseBorderFromRaw(tcBorders, rightToLeft ? "start" : "end") ??
            ParseBorderFromRaw(tcBorders, "right");
        DxpComputedBorder? bottom = ParseBorderFromRaw(tcBorders, "bottom");
        DxpComputedBorder? left = ParseBorderFromRaw(tcBorders, rightToLeft ? "end" : "start") ??
            ParseBorderFromRaw(tcBorders, "left");

        if (top == null && right == null && bottom == null && left == null)
            return null;

        return new DxpComputedBoxBorders(top, right, bottom, left);
    }

    private static DxpComputedBorder? ParseBorderFromRaw(OpenXmlElement tcBorders, string sideLocalName)
    {
        var b = FindFirstChildByLocalName(tcBorders, sideLocalName);
        if (b == null)
            return null;

        var val = GetAttributeValue(b, "val");
        if (string.Equals(val, "none", StringComparison.OrdinalIgnoreCase) || string.Equals(val, "nil", StringComparison.OrdinalIgnoreCase))
            return new DxpComputedBorder(WidthPt: 0, LineStyle: DxpComputedBorderLineStyle.None, ColorCss: "#000000");

        var sz = GetAttributeValue(b, "sz");
        if (string.IsNullOrWhiteSpace(sz) || !int.TryParse(sz, out int sizeEighthPoints) || sizeEighthPoints <= 0)
            return null;

        double pt = sizeEighthPoints / 8.0;
        DxpComputedBorderLineStyle line = DxpComputedBorderLineStyle.Solid;
        if (string.Equals(val, "dotted", StringComparison.OrdinalIgnoreCase))
            line = DxpComputedBorderLineStyle.Dotted;
        else if (string.Equals(val, "dashed", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(val, "dashSmallGap", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(val, "dotDash", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(val, "dotDotDash", StringComparison.OrdinalIgnoreCase))
            line = DxpComputedBorderLineStyle.Dashed;
        else if (string.Equals(val, "double", StringComparison.OrdinalIgnoreCase))
            line = DxpComputedBorderLineStyle.Double;

        string? color = GetAttributeValue(b, "color");
        if (string.IsNullOrEmpty(color) || string.Equals(color, "auto", StringComparison.OrdinalIgnoreCase))
            color = "#000000";
        else
            color = ToCssColor(color!);

        return new DxpComputedBorder(WidthPt: pt, LineStyle: line, ColorCss: color);
    }

    private static (bool BorderSet, DxpComputedBorder? Border) PickComputedBorder(MergedTableBorders merged)
    {
        bool any = merged.AnySpecified;
        var b = merged.PickFirstBorderCandidate();
        var computed = ToComputedBorder(b);
        if (computed != null)
            return (true, computed);

        return (any, null);
    }

    private static DxpComputedBorder? ToComputedBorder(BorderType? b)
    {
        if (b == null)
            return null;

        var val = b.Val?.Value;
        if (val == BorderValues.None || val == BorderValues.Nil)
            return new DxpComputedBorder(WidthPt: 0, LineStyle: DxpComputedBorderLineStyle.None, ColorCss: "#000000");

        int sizeEighthPoints = b.Size != null ? (int)b.Size.Value : 0;
        if (sizeEighthPoints <= 0)
            return null;

        double pt = sizeEighthPoints / 8.0;
        var line = MapLineStyle(val);
        string? color = b.Color?.Value;
        if (string.IsNullOrEmpty(color) || string.Equals(color, "auto", StringComparison.OrdinalIgnoreCase))
            color = "#000000";
        else
            color = ToCssColor(color!);

        return new DxpComputedBorder(
            WidthPt: pt,
            LineStyle: line,
            ColorCss: color);
    }

    private static DxpComputedBorderLineStyle MapLineStyle(BorderValues? val)
    {
        if (val == BorderValues.Dotted)
            return DxpComputedBorderLineStyle.Dotted;
        if (val == BorderValues.DashSmallGap || val == BorderValues.Dashed || val == BorderValues.DotDash || val == BorderValues.DotDotDash)
            return DxpComputedBorderLineStyle.Dashed;
        if (val == BorderValues.Double)
            return DxpComputedBorderLineStyle.Double;
        return DxpComputedBorderLineStyle.Solid;
    }

    private static string ToCssColor(string color)
    {
        if (color.StartsWith("#", StringComparison.Ordinal))
            return color;
        if (color.Length is 6 or 3)
            return "#" + color;
        return color;
    }

    internal readonly record struct DxpTableLook(bool FirstRow, bool LastRow, bool FirstColumn, bool LastColumn, bool NoHBand, bool NoVBand);

    internal sealed class DxpResolvedTableStyle
    {
        public string? TableStyleId { get; }
        public DxpTableLook Look { get; }
        public int RowCount { get; }
        public int ColCount { get; }
        public IReadOnlyList<Style> StyleChain { get; }
        public DxpComputedTableStyle ComputedTableStyle { get; }
        public DxpComputedTableCellStyle BaseCellStyle { get; }
        public bool RightToLeft { get; }

        public DxpResolvedTableStyle(string? tableStyleId, DxpTableLook look, int rowCount, int colCount,
            IReadOnlyList<Style> styleChain, DxpComputedTableStyle computedTableStyle,
            DxpComputedTableCellStyle baseCellStyle, bool rightToLeft)
        {
            TableStyleId = tableStyleId;
            Look = look;
            RowCount = rowCount;
            ColCount = colCount;
            StyleChain = styleChain;
            ComputedTableStyle = computedTableStyle;
            BaseCellStyle = baseCellStyle;
            RightToLeft = rightToLeft;
        }

        public IEnumerable<string> GetApplicableConditionalTypes(int rowIndex, int colIndex)
        {
            bool firstRow = Look.FirstRow && rowIndex == 0;
            bool lastRow = Look.LastRow && RowCount > 0 && rowIndex == RowCount - 1;
            bool firstCol = Look.FirstColumn && colIndex == 0;
            bool lastCol = Look.LastColumn && ColCount > 0 && colIndex == ColCount - 1;

            bool hBanding = !Look.NoHBand;
            bool vBanding = !Look.NoVBand;

            // Banding is lowest precedence.
            if (hBanding)
            {
                int bodyRowIndex = rowIndex;
                if (Look.FirstRow)
                    bodyRowIndex -= 1;
                bool excluded = firstRow || lastRow;
                if (!excluded && bodyRowIndex >= 0)
                    yield return (bodyRowIndex % 2 == 0) ? "band1Horz" : "band2Horz";
            }

            if (vBanding)
            {
                int bodyColIndex = colIndex;
                if (Look.FirstColumn)
                    bodyColIndex -= 1;
                bool excluded = firstCol || lastCol;
                if (!excluded && bodyColIndex >= 0)
                    yield return (bodyColIndex % 2 == 0) ? "band1Vert" : "band2Vert";
            }

            // Edge conditions.
            if (firstRow)
                yield return "firstRow";
            if (lastRow)
                yield return "lastRow";
            if (firstCol)
                yield return "firstCol";
            if (lastCol)
                yield return "lastCol";

            // Corners are most specific.
            if (firstRow && firstCol)
                yield return "nwCell";
            if (firstRow && lastCol)
                yield return "neCell";
            if (lastRow && firstCol)
                yield return "swCell";
            if (lastRow && lastCol)
                yield return "seCell";
        }
    }

    private struct ResolvedOverride
    {
        public bool BorderSet;
        public DxpComputedBorder? Border;
        public bool BordersSet;
        public DxpComputedBoxBorders? Borders;
        public bool BackgroundSet;
        public string? BackgroundColorCss;
        public bool VerticalAlignSet;
        public DxpComputedVerticalAlign? VerticalAlign;
        public bool TextColorSet;
        public string? TextColorCss;
        public bool? Bold;
        public bool? Italic;
        public bool? Underline;
        public DxpComputedTextAlign? TextAlign;

        public void ApplyCellProperties(OpenXmlElement cellPropertiesElement)
        {
            // Background (w:shd)
            var shd = cellPropertiesElement.GetFirstChild<Shading>() ?? FindFirstChildByLocalName(cellPropertiesElement, "shd");
            if (shd != null)
            {
                var fill = GetAttributeValue(shd, "fill");
                BackgroundSet = true;
                if (!string.IsNullOrWhiteSpace(fill) && !string.Equals(fill, "auto", StringComparison.OrdinalIgnoreCase))
                    BackgroundColorCss = ToCssColor(fill!);
                else
                    BackgroundColorCss = null;
            }

            // Vertical alignment (w:vAlign)
            var vAlignEl = (OpenXmlElement?)cellPropertiesElement.GetFirstChild<TableCellVerticalAlignment>() ?? FindFirstChildByLocalName(cellPropertiesElement, "vAlign");
            string? v = null;
            if (vAlignEl is TableCellVerticalAlignment typed)
            {
                if (typed.Val != null)
                    v = typed.Val.Value.ToString();
            }
            else if (vAlignEl != null)
                v = GetAttributeValue(vAlignEl, "val");

            if (vAlignEl != null && v != null)
            {
                VerticalAlignSet = true;
                if (string.Equals(v, "top", StringComparison.OrdinalIgnoreCase))
                    VerticalAlign = DxpComputedVerticalAlign.Top;
                else if (string.Equals(v, "center", StringComparison.OrdinalIgnoreCase))
                    VerticalAlign = DxpComputedVerticalAlign.Middle;
                else if (string.Equals(v, "bottom", StringComparison.OrdinalIgnoreCase))
                    VerticalAlign = DxpComputedVerticalAlign.Bottom;
                else
                    VerticalAlign = null;
            }
        }
    }

    private static DxpComputedBoxBorders ResolveInitialBorders(DxpResolvedTableStyle resolved, DxpComputedTableCellStyle baseCell, DxpComputedTableStyle tableStyle, int rowIndex, int colIndex, int rowSpan, int colSpan)
    {
        // Start from the position-based table defaults (outer vs inside borders),
        // then apply style-level defaults (tcPr), then conditional, then direct tcPr.
        var visualColumn = resolved.RightToLeft
            ? resolved.ColCount - colIndex - Math.Max(1, colSpan) : colIndex;
        var borders = ComputePositionBasedTableCellBorders(resolved, tableStyle, rowIndex,
            visualColumn, rowSpan, colSpan);

        if (baseCell.Border != null)
            borders = new DxpComputedBoxBorders(baseCell.Border, baseCell.Border, baseCell.Border, baseCell.Border);
        if (baseCell.Borders != null)
            borders = ApplyBorderOverride(borders, baseCell.Borders);

        return borders;
    }

    private static DxpComputedBoxBorders ComputePositionBasedTableCellBorders(DxpResolvedTableStyle resolved, DxpComputedTableStyle tableStyle, int rowIndex, int colIndex, int rowSpan, int colSpan)
    {
        var inside = tableStyle.DefaultCellBorders;
        var outer = tableStyle.TableBorders;
        DxpComputedBorder? insideAll = outer == null ? tableStyle.DefaultCellBorder ?? tableStyle.TableBorder : null;
        DxpComputedBorder? outerAll = outer == null ? tableStyle.TableBorder : null;

        bool isTopEdge = rowIndex == 0;
        bool isBottomEdge = (resolved.RowCount > 0) && (rowIndex + Math.Max(1, rowSpan) - 1) == (resolved.RowCount - 1);
        bool isLeftEdge = colIndex == 0;
        bool isRightEdge = (resolved.ColCount > 0) && (colIndex + Math.Max(1, colSpan) - 1) == (resolved.ColCount - 1);

        var top = isTopEdge ? (outer?.Top ?? outerAll) : (inside?.Top ?? insideAll);
        var bottom = isBottomEdge ? (outer?.Bottom ?? outerAll) : (inside?.Bottom ?? insideAll);
        var left = isLeftEdge ? (outer?.Left ?? outerAll) : (inside?.Left ?? insideAll);
        var right = isRightEdge ? (outer?.Right ?? outerAll) : (inside?.Right ?? insideAll);

        return new DxpComputedBoxBorders(top, right, bottom, left);
    }

    private static DxpComputedBoxBorders ApplyBorderOverride(DxpComputedBoxBorders current, DxpComputedBoxBorders ov)
    {
        return new DxpComputedBoxBorders(
            Top: ov.Top ?? current.Top,
            Right: ov.Right ?? current.Right,
            Bottom: ov.Bottom ?? current.Bottom,
            Left: ov.Left ?? current.Left);
    }

    private static DxpComputedBorder? CollapseToSingleBorder(DxpComputedBoxBorders borders)
    {
        if (borders.Top == null)
            return null;
        if (!Equals(borders.Top, borders.Right) || !Equals(borders.Top, borders.Bottom) || !Equals(borders.Top, borders.Left))
            return null;
        return borders.Top;
    }

    private static bool HasVisibleBorder(DxpComputedBoxBorders? borders)
    {
        if (borders == null)
            return false;

        foreach (var b in new[] { borders.Top, borders.Right, borders.Bottom, borders.Left })
        {
            if (b != null && b.LineStyle != DxpComputedBorderLineStyle.None)
                return true;
        }
        return false;
    }

    private static DxpComputedBoxBorders? ComputeDirectCellBorders(TableCellProperties? cellProperties,
        bool rightToLeft)
    {
        var borders = cellProperties?.TableCellBorders;
        if (borders == null)
            return null;

        // Only treat explicitly-specified sides as overrides; null means "fall back".
        var top = ToComputedBorder(borders.TopBorder);
        var right = ToComputedBorder(rightToLeft ? borders.StartBorder : borders.EndBorder) ??
            ToComputedBorder(borders.RightBorder);
        var bottom = ToComputedBorder(borders.BottomBorder);
        var left = ToComputedBorder(rightToLeft ? borders.EndBorder : borders.StartBorder) ??
            ToComputedBorder(borders.LeftBorder);

        if (top == null && right == null && bottom == null && left == null)
            return null;

        return new DxpComputedBoxBorders(top, right, bottom, left);
    }

    private static DxpComputedBoxBorders? ComputeTableBoxBorders(MergedTableBorders borders, out bool anySpecified)
    {
        anySpecified = borders.AnySpecified;
        if (!borders.AnySpecified)
            return null;

        var top = ToComputedBorder(borders.Top);
        var right = ToComputedBorder(borders.Right);
        var bottom = ToComputedBorder(borders.Bottom);
        var left = ToComputedBorder(borders.Left);

        return new DxpComputedBoxBorders(top, right, bottom, left);
    }

    private static DxpComputedBoxBorders? ComputeDefaultCellBoxBordersFromTable(MergedTableBorders borders)
    {
        if (!borders.AnySpecified)
            return null;

        var insideH = ToComputedBorder(borders.InsideH);
        var insideV = ToComputedBorder(borders.InsideV);

        var topBottom = insideH;
        var leftRight = insideV;
        if (topBottom == null && leftRight == null)
            return null;

        return new DxpComputedBoxBorders(Top: topBottom, Right: leftRight, Bottom: topBottom, Left: leftRight);
    }

    private static DxpComputedBoxBorders? ComputeCellBoxBordersFromMerged(MergedCellBorders borders)
    {
        if (!borders.AnySpecified)
            return null;

        var top = ToComputedBorder(borders.Top);
        var right = ToComputedBorder(borders.Right);
        var bottom = ToComputedBorder(borders.Bottom);
        var left = ToComputedBorder(borders.Left);

        if (top == null && right == null && bottom == null && left == null)
            return null;

        return new DxpComputedBoxBorders(top, right, bottom, left);
    }

    private sealed class MergedTableBorders
    {
        public BorderType? Top;
        public BorderType? Left;
        public BorderType? Bottom;
        public BorderType? Right;
        public BorderType? InsideH;
        public BorderType? InsideV;
        public bool AnySpecified { get; private set; }

        public void Apply(TableBorders? borders, bool rightToLeft)
        {
            if (borders == null)
                return;

            Apply(ref Top, borders.TopBorder);
            Apply(ref Left, borders.LeftBorder);
            Apply(ref Bottom, borders.BottomBorder);
            Apply(ref Right, borders.RightBorder);
            if (rightToLeft)
            {
                Apply(ref Right, borders.StartBorder);
                Apply(ref Left, borders.EndBorder);
            }
            else
            {
                Apply(ref Left, borders.StartBorder);
                Apply(ref Right, borders.EndBorder);
            }
            Apply(ref InsideH, borders.InsideHorizontalBorder);
            Apply(ref InsideV, borders.InsideVerticalBorder);
        }

        private void Apply(ref BorderType? target, BorderType? incoming)
        {
            if (incoming == null)
                return;

            AnySpecified = true;
            target = incoming;
        }

        public BorderType? PickFirstBorderCandidate()
        {
            foreach (var b in new BorderType?[] { Top, Left, Bottom, Right, InsideH, InsideV })
            {
                if (b != null)
                    return b;
            }
            return null;
        }
    }

    private sealed class MergedCellBorders
    {
        public BorderType? Top;
        public BorderType? Left;
        public BorderType? Bottom;
        public BorderType? Right;
        public BorderType? InsideH;
        public BorderType? InsideV;
        public bool AnySpecified { get; private set; }

        public void Apply(TableCellBorders? borders, bool rightToLeft)
        {
            if (borders == null)
                return;

            Apply(ref Top, borders.TopBorder);
            Apply(ref Left, borders.LeftBorder);
            Apply(ref Bottom, borders.BottomBorder);
            Apply(ref Right, borders.RightBorder);
            if (rightToLeft)
            {
                Apply(ref Right, borders.StartBorder);
                Apply(ref Left, borders.EndBorder);
            }
            else
            {
                Apply(ref Left, borders.StartBorder);
                Apply(ref Right, borders.EndBorder);
            }
            Apply(ref InsideH, borders.InsideHorizontalBorder);
            Apply(ref InsideV, borders.InsideVerticalBorder);
        }

        private void Apply(ref BorderType? target, BorderType? incoming)
        {
            if (incoming == null)
                return;

            AnySpecified = true;
            target = incoming;
        }

        public BorderType? PickFirstBorderCandidate()
        {
            foreach (var b in new BorderType?[] { Top, Left, Bottom, Right, InsideH, InsideV })
            {
                if (b != null)
                    return b;
            }
            return null;
        }
    }
}
