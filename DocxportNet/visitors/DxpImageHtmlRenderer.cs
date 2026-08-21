using DocxportNet.API;
using System.Globalization;
using System.Net;
using System.Text;

namespace DocxportNet.Visitors;

internal static class DxpImageHtmlRenderer
{
    public static string Render(
        string source,
        DxpDrawingInfo? info,
        string imageCssClass,
        string? contextualStyle = null)
    {
        DxpImagePresentation? presentation = info?.Presentation;
        string alt = presentation?.IsDecorative == true
            ? string.Empty
            : presentation?.AlternativeText ?? info?.AltText ?? "image";
        string? title = presentation?.Title;
        string attributes = $"src=\"{Attribute(source)}\" alt=\"{Attribute(alt)}\"" +
            (string.IsNullOrEmpty(title) ? string.Empty : $" title=\"{Attribute(title!)}\"");

        double? width = presentation?.FrameWidthPoints;
        double? height = presentation?.FrameHeightPoints;
        DxpImageCrop? crop = presentation?.Crop;
        bool hasTransform = presentation != null &&
            (Math.Abs(presentation.RotationDegrees) > 0.000001 || presentation.FlipHorizontal || presentation.FlipVertical);
        bool needsFrame = width != null && height != null && (crop != null || hasTransform);

        if (!needsFrame)
        {
            var style = new StringBuilder();
            AppendDimension(style, "width", width);
            AppendDimension(style, "height", height);
            if (width != null || height != null)
                style.Append("max-width:none;");
            AppendRaw(style, contextualStyle);
            string styleAttribute = style.Length == 0 ? string.Empty : $" style=\"{style}\"";
            return $"<img class=\"{Attribute(imageCssClass)}\" {attributes}{styleAttribute} />";
        }

        var frameStyle = new StringBuilder("display:inline-block;position:relative;");
        AppendDimension(frameStyle, "width", width);
        AppendDimension(frameStyle, "height", height);
        frameStyle.Append("overflow:hidden;");
        AppendPresentationTransform(frameStyle, presentation!);

        var imageStyle = new StringBuilder("position:absolute;max-width:none;");
        if (crop != null)
        {
            double visibleWidth = 1.0 - crop.Left - crop.Right;
            double visibleHeight = 1.0 - crop.Top - crop.Bottom;
            double fullWidth = width!.Value / visibleWidth;
            double fullHeight = height!.Value / visibleHeight;
            AppendDimension(imageStyle, "width", fullWidth);
            AppendDimension(imageStyle, "height", fullHeight);
            AppendDimension(imageStyle, "left", -width.Value * crop.Left / visibleWidth);
            AppendDimension(imageStyle, "top", -height.Value * crop.Top / visibleHeight);
        }
        else
        {
            imageStyle.Append("left:0;top:0;width:100%;height:100%;");
        }

        string frame = $"<span class=\"dxp-image-frame\" style=\"{frameStyle}\"><img class=\"{Attribute(imageCssClass)}\" {attributes} style=\"{imageStyle}\" /></span>";
        if (string.IsNullOrWhiteSpace(contextualStyle))
            return frame;
        return $"<span class=\"dxp-image-position\" style=\"{Attribute(contextualStyle!)}\">{frame}</span>";
    }

    private static void AppendPresentationTransform(StringBuilder style, DxpImagePresentation presentation)
    {
        var transform = new StringBuilder();
        if (Math.Abs(presentation.RotationDegrees) > 0.000001)
            transform.Append("rotate(").Append(Number(presentation.RotationDegrees)).Append("deg)");
        if (presentation.FlipHorizontal)
        {
            if (transform.Length > 0)
                transform.Append(' ');
            transform.Append("scaleX(-1)");
        }
        if (presentation.FlipVertical)
        {
            if (transform.Length > 0)
                transform.Append(' ');
            transform.Append("scaleY(-1)");
        }
        if (transform.Length > 0)
            style.Append("transform:").Append(transform).Append(";transform-origin:center;");
    }

    private static void AppendDimension(StringBuilder style, string name, double? points)
    {
        if (points != null)
            style.Append(name).Append(':').Append(Number(points.Value)).Append("pt;");
    }

    private static void AppendRaw(StringBuilder style, string? css)
    {
        if (string.IsNullOrWhiteSpace(css))
            return;
        style.Append(css);
        if (style[style.Length - 1] != ';')
            style.Append(';');
    }

    private static string Number(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);
    private static string Attribute(string value) => WebUtility.HtmlEncode(value);
}
