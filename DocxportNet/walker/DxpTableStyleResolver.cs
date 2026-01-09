using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;

namespace DocxportNet.Walker;

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

		var baseTableStyle = ComputeBaseTableStyle(styleChain, tableProperties);
		var baseCell = ComputeBaseCellStyleFromStyles(styleChain, baseTableStyle);

		return new DxpResolvedTableStyle(styleId, look, rowCount, colCount, styleChain, baseTableStyle, baseCell);
	}

	public static DxpComputedTableCellStyle ComputeCellStyle(DxpResolvedTableStyle resolved, TableCellProperties? cellProperties, int rowIndex, int colIndex)
	{
		var tableStyle = resolved.ComputedTableStyle;
		var baseCell = resolved.BaseCellStyle;

		var computedBorder = baseCell.Border ?? tableStyle.DefaultCellBorder;
		var computedBackground = baseCell.BackgroundColorCss;
		var computedVAlign = baseCell.VerticalAlign;

		// Apply conditional formatting (tblStylePr) from the table style chain.
		foreach (var type in resolved.GetApplicableConditionalTypes(rowIndex, colIndex))
		{
			var ov = ComputeConditionalOverride(resolved.StyleChain, type);
			if (ov.BorderSet)
				computedBorder = ov.Border ?? tableStyle.DefaultCellBorder;
			if (ov.BackgroundSet)
				computedBackground = ov.BackgroundColorCss;
			if (ov.VerticalAlignSet)
				computedVAlign = ov.VerticalAlign;
		}

		// Direct cell formatting always wins.
		var direct = DxpTableStyleComputer.ComputeDirectCellStyle(cellProperties);
		if (direct.Border != null)
			computedBorder = direct.Border;
		if (direct.BackgroundColorCss != null)
			computedBackground = direct.BackgroundColorCss;
		if (direct.VerticalAlign != null)
			computedVAlign = direct.VerticalAlign;

		return new DxpComputedTableCellStyle(Border: computedBorder, BackgroundColorCss: computedBackground, VerticalAlign: computedVAlign);
	}

	private static DxpComputedTableStyle ComputeBaseTableStyle(IReadOnlyList<Style> styleChain, TableProperties? directTableProperties)
	{
		// Merge tblBorders across the style chain (base -> derived), then overlay direct formatting.
		var merged = new MergedTableBorders();
		foreach (var s in styleChain)
		{
			var stp = s.GetFirstChild<StyleTableProperties>();
			var borders = stp?.GetFirstChild<TableBorders>();
			merged.Apply(borders);
		}
		merged.Apply(directTableProperties?.TableBorders);

		var border = PickComputedBorder(merged).Border;
		if (border == null)
			return new DxpComputedTableStyle(TableBorder: null, BorderCollapse: false, DefaultCellBorder: null);

		// DefaultCellBorder may be overridden by style tcPr later, but keeps the prior "same border everywhere" fallback.
		return new DxpComputedTableStyle(
			TableBorder: border,
			BorderCollapse: true,
			DefaultCellBorder: border);
	}

	private static DxpComputedTableCellStyle ComputeBaseCellStyleFromStyles(IReadOnlyList<Style> styleChain, DxpComputedTableStyle tableStyle)
	{
		var mergedBorders = new MergedCellBorders();
		ResolvedOverride mergedOverrides = default;

		foreach (var s in styleChain)
		{
			var stcp = s.GetFirstChild<StyleTableCellProperties>();
			if (stcp == null)
				continue;

			mergedBorders.Apply(stcp.GetFirstChild<TableCellBorders>());
			mergedOverrides.ApplyCellProperties(stcp);
		}

		var border = PickComputedBorder(mergedBorders);
		var effectiveBorder = border.BorderSet ? border.Border : (tableStyle.DefaultCellBorder ?? tableStyle.TableBorder);

		return new DxpComputedTableCellStyle(
			Border: effectiveBorder,
			BackgroundColorCss: mergedOverrides.BackgroundSet ? mergedOverrides.BackgroundColorCss : null,
			VerticalAlign: mergedOverrides.VerticalAlignSet ? mergedOverrides.VerticalAlign : null);
	}

	private static ResolvedOverride ComputeConditionalOverride(IReadOnlyList<Style> styleChain, string type)
	{
		ResolvedOverride result = default;
		var mergedBorders = new MergedCellBorders();
		DxpComputedBorder? computedBorder = null;
		bool borderSet = false;

		foreach (var s in styleChain)
		{
			foreach (var el in s.ChildElements)
			{
				if (!string.Equals(el.LocalName, "tblStylePr", StringComparison.Ordinal))
					continue;

				var t = GetAttributeValue(el, "type");
				if (!string.Equals(t, type, StringComparison.OrdinalIgnoreCase))
					continue;

				// styles.xml can surface tblStylePr content as unknown elements, so we use LocalName/attributes parsing.
				OpenXmlElement? tcPr =
					(OpenXmlElement?)el.GetFirstChild<StyleTableCellProperties>() ??
					el.GetFirstChild<TableCellProperties>() ??
					FindFirstChildByLocalName(el, "tcPr");

				if (tcPr != null)
				{
					mergedBorders.Apply(tcPr.GetFirstChild<TableCellBorders>());
					result.ApplyCellProperties(tcPr);

					// Fallback: parse borders from raw XML if they were not materialized into TableCellBorders/BorderType.
					var tcBorders = FindFirstChildByLocalName(tcPr, "tcBorders");
					if (tcBorders != null)
					{
						var b = PickFirstBorderCandidateFromRaw(tcBorders);
						borderSet = true;
						if (computedBorder == null)
							computedBorder = b;
					}
				}
			}
		}

		var border = PickComputedBorder(mergedBorders);
		if (border.BorderSet || borderSet)
		{
			result.BorderSet = true;
			result.Border = border.Border ?? computedBorder;
		}

		return result;
	}

	private IReadOnlyList<Style> ResolveTableStyleChain(string? styleId)
	{
		if (string.IsNullOrEmpty(styleId))
			return Array.Empty<Style>();

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

	private static DxpComputedBorder? PickFirstBorderCandidateFromRaw(OpenXmlElement tcBorders)
	{
		// Mirrors the existing "pick first non-null border" behavior.
		foreach (var name in new[] { "top", "left", "bottom", "right", "insideH", "insideV" })
		{
			var b = FindFirstChildByLocalName(tcBorders, name);
			if (b == null)
				continue;

			var val = GetAttributeValue(b, "val");
			if (string.Equals(val, "none", StringComparison.OrdinalIgnoreCase) || string.Equals(val, "nil", StringComparison.OrdinalIgnoreCase))
				continue;

			var sz = GetAttributeValue(b, "sz");
			if (string.IsNullOrWhiteSpace(sz) || !int.TryParse(sz, out int sizeEighthPoints) || sizeEighthPoints <= 0)
				continue;

			double pt = sizeEighthPoints / 8.0;
			string? color = GetAttributeValue(b, "color");
			if (string.IsNullOrEmpty(color) || string.Equals(color, "auto", StringComparison.OrdinalIgnoreCase))
				color = "#000000";
			else
				color = ToCssColor(color!);

			return new DxpComputedBorder(WidthPt: pt, LineStyle: DxpComputedBorderLineStyle.Solid, ColorCss: color);
		}

		return null;
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

	private static (bool BorderSet, DxpComputedBorder? Border) PickComputedBorder(MergedCellBorders merged)
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
			return null;

		int sizeEighthPoints = b.Size != null ? (int)b.Size.Value : 0;
		if (sizeEighthPoints <= 0)
			return null;

		double pt = sizeEighthPoints / 8.0;
		string? color = b.Color?.Value;
		if (string.IsNullOrEmpty(color) || string.Equals(color, "auto", StringComparison.OrdinalIgnoreCase))
			color = "#000000";
		else
			color = ToCssColor(color!);

		return new DxpComputedBorder(
			WidthPt: pt,
			LineStyle: DxpComputedBorderLineStyle.Solid,
			ColorCss: color);
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

		public DxpResolvedTableStyle(string? tableStyleId, DxpTableLook look, int rowCount, int colCount, IReadOnlyList<Style> styleChain, DxpComputedTableStyle computedTableStyle, DxpComputedTableCellStyle baseCellStyle)
		{
			TableStyleId = tableStyleId;
			Look = look;
			RowCount = rowCount;
			ColCount = colCount;
			StyleChain = styleChain;
			ComputedTableStyle = computedTableStyle;
			BaseCellStyle = baseCellStyle;
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
			if (firstRow) yield return "firstRow";
			if (lastRow) yield return "lastRow";
			if (firstCol) yield return "firstCol";
			if (lastCol) yield return "lastCol";

			// Corners are most specific.
			if (firstRow && firstCol) yield return "nwCell";
			if (firstRow && lastCol) yield return "neCell";
			if (lastRow && firstCol) yield return "swCell";
			if (lastRow && lastCol) yield return "seCell";
		}
	}

	private struct ResolvedOverride
	{
		public bool BorderSet;
		public DxpComputedBorder? Border;
		public bool BackgroundSet;
		public string? BackgroundColorCss;
		public bool VerticalAlignSet;
		public DxpComputedVerticalAlign? VerticalAlign;

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

	private sealed class MergedTableBorders
	{
		public BorderType? Top;
		public BorderType? Left;
		public BorderType? Bottom;
		public BorderType? Right;
		public BorderType? InsideH;
		public BorderType? InsideV;
		public bool AnySpecified { get; private set; }

		public void Apply(TableBorders? borders)
		{
			if (borders == null)
				return;

			Apply(ref Top, borders.TopBorder);
			Apply(ref Left, borders.LeftBorder);
			Apply(ref Bottom, borders.BottomBorder);
			Apply(ref Right, borders.RightBorder);
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

		public void Apply(TableCellBorders? borders)
		{
			if (borders == null)
				return;

			Apply(ref Top, borders.TopBorder);
			Apply(ref Left, borders.LeftBorder);
			Apply(ref Bottom, borders.BottomBorder);
			Apply(ref Right, borders.RightBorder);
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
