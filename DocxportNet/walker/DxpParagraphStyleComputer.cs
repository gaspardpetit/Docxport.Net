using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;

namespace DocxportNet.Walker;

internal static class DxpParagraphStyleComputer
{
	public static DxpComputedParagraphStyle ComputeParagraphStyle(Paragraph p, DxpStyleEffectiveIndentTwips indent, DxpDocumentContext d)
	{
		double? marginLeftPt = null;
		if (indent.Left.HasValue)
		{
			double pt = DxpTwipValue.ToPoints(indent.Left.Value);
			var marginLeftPoints = d.CurrentSection.Layout?.MarginLeft?.Inches is double inches
				? inches * 72.0
				: (double?)null;
			if (marginLeftPoints != null)
			{
				pt -= marginLeftPoints.Value;
				if (pt < 0)
					pt = 0;
			}
			if (pt > 0.0001)
				marginLeftPt = pt;
		}

		DxpComputedTextAlign? align = null;
		var justification = p.ParagraphProperties?.Justification?.Val?.Value;
		if (justification == JustificationValues.Center)
			align = DxpComputedTextAlign.Center;
		else if (justification == JustificationValues.Right)
			align = DxpComputedTextAlign.Right;
		else if (justification == JustificationValues.Both || justification == JustificationValues.Distribute)
			align = DxpComputedTextAlign.Justify;

		var borders = ComputeBorders(p.ParagraphProperties?.ParagraphBorders);
		var background = ComputeBackground(p.ParagraphProperties?.Shading);

		return new DxpComputedParagraphStyle(
			MarginLeftPt: marginLeftPt,
			TextAlign: align,
			Borders: borders,
			BackgroundColorCss: background);
	}

	private static DxpComputedBoxBorders? ComputeBorders(ParagraphBorders? bdr)
	{
		if (bdr == null)
			return null;

		var top = ToComputedBorder(bdr.TopBorder);
		var right = ToComputedBorder(bdr.RightBorder);
		var bottom = ToComputedBorder(bdr.BottomBorder);
		var left = ToComputedBorder(bdr.LeftBorder);

		if (top == null && right == null && bottom == null && left == null)
			return null;

		return new DxpComputedBoxBorders(top, right, bottom, left);
	}

	private static string? ComputeBackground(Shading? shd)
	{
		var fill = shd?.Fill?.Value;
		if (string.IsNullOrWhiteSpace(fill) || string.Equals(fill, "auto", StringComparison.OrdinalIgnoreCase))
			return null;
		return ToCssColor(fill!);
	}

	private static DxpComputedBorder? ToComputedBorder(BorderType? b)
	{
		if (b == null)
			return null;

		var val = b.Val?.Value;
		// For paragraphs, treat "none/nil" as absent to avoid emitting redundant `border-*:none`.
		if (val == BorderValues.None || val == BorderValues.Nil)
			return null;

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

		return new DxpComputedBorder(WidthPt: pt, LineStyle: line, ColorCss: color);
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
}
