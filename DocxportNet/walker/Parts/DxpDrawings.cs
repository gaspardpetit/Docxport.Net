using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using System.Text;

namespace DocxportNet.Walker;


public class DxpDrawings
{
    public (string dataUri, string contentType)? TryBuildImageDataUri(OpenXmlPart? hostPart, Drawing drw)
    {
        if (hostPart == null)
            return null;

        var blip = drw.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
        var relId = blip?.Embed?.Value;

        if (string.IsNullOrEmpty(relId))
            return null; // not an embedded raster image (could be chart/SmartArt/etc.)

        if (hostPart.GetPartById(relId!) is not ImagePart imgPart)
            return null;

        byte[] bytes;
        using (var stream = imgPart.GetStream(FileMode.Open, FileAccess.Read))
        using (var ms = new MemoryStream())
        {
            stream.CopyTo(ms);
            bytes = ms.ToArray();
        }

        var base64 = Convert.ToBase64String(bytes);
        var contentType = imgPart.ContentType; // e.g. "image/png", "image/jpeg"

        var dataUri = $"data:{contentType};base64,{base64}";
        return (dataUri, contentType);
    }

    public DxpDrawingInfo? TryResolveDrawingInfo(OpenXmlPart? hostPart, Drawing drw)
    {
        if (hostPart == null)
            return null;

        var docPr = drw.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>()
                     .FirstOrDefault();
        var pictureProperties = drw.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties>()
            .FirstOrDefault();
        string? altText = NormalizeAltText(docPr?.Description?.Value ?? pictureProperties?.Description?.Value);
        string? title = NormalizeAltText(docPr?.Title?.Value ?? pictureProperties?.Title?.Value);
        bool isDecorative = drw.Descendants()
            .Where(element => string.Equals(element.LocalName, "decorative", StringComparison.OrdinalIgnoreCase))
            .Any(element => IsOn(element.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == "val").Value));

        var blip = drw.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
        var relId = blip?.Embed?.Value;
        var linkRelId = blip?.Link?.Value;

        string? contentType = null;
        string? fileName = null;
        string? dataUri = null;
        string? externalSource = null;

        if (!string.IsNullOrEmpty(relId))
        {
            try
            {
                var part = hostPart.GetPartById(relId!);
                contentType = part.ContentType;
                fileName = part.Uri?.ToString();

                var built = TryBuildAnyDataUri(part);
                dataUri = built?.dataUri;
            }
            catch { /* swallow and return partial info */ }
        }

        if (!string.IsNullOrEmpty(linkRelId))
        {
            try
            {
                externalSource = hostPart.ExternalRelationships
                    .FirstOrDefault(relationship => relationship.Id == linkRelId)?.Uri.OriginalString;
            }
            catch { /* swallow and return partial info */ }
        }

        var presentation = BuildImagePresentation(drw, altText, title, isDecorative);
        return new DxpDrawingInfo(relId, contentType, fileName, presentation.AlternativeText, dataUri) {
            ExternalSource = externalSource,
            Presentation = presentation
        };
    }

    public static DxpImagePresentation BuildImagePresentation(Drawing drw)
    {
        var docPr = drw.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>()
            .FirstOrDefault();
        var pictureProperties = drw.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties>()
            .FirstOrDefault();
        string? altText = NormalizeAltText(docPr?.Description?.Value ?? pictureProperties?.Description?.Value);
        string? title = NormalizeAltText(docPr?.Title?.Value ?? pictureProperties?.Title?.Value);
        bool isDecorative = drw.Descendants()
            .Where(element => string.Equals(element.LocalName, "decorative", StringComparison.OrdinalIgnoreCase))
            .Any(element => IsOn(element.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == "val").Value));
        return BuildImagePresentation(drw, altText, title, isDecorative);
    }

    private static DxpImagePresentation BuildImagePresentation(
        Drawing drw,
        string? altText,
        string? title,
        bool isDecorative)
    {
        var extent = drw.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>().FirstOrDefault();
        double? widthPoints = extent?.Cx is { } cx && cx.Value > 0 ? cx.Value / 12700.0 : null;
        double? heightPoints = extent?.Cy is { } cy && cy.Value > 0 ? cy.Value / 12700.0 : null;

        var sourceRectangle = drw.Descendants<DocumentFormat.OpenXml.Drawing.SourceRectangle>().FirstOrDefault();
        var crop = sourceRectangle == null ? null : NormalizeCrop(
            sourceRectangle.Left?.Value,
            sourceRectangle.Top?.Value,
            sourceRectangle.Right?.Value,
            sourceRectangle.Bottom?.Value);

        var transform = drw.Descendants<DocumentFormat.OpenXml.Drawing.Transform2D>().FirstOrDefault();
        double rotation = 0;
        if (transform?.Rotation?.Value is int rotationUnits)
        {
            rotation = (rotationUnits / 60000.0) % 360.0;
            if (rotation < 0)
                rotation += 360.0;
        }

        return new DxpImagePresentation {
            FrameWidthPoints = widthPoints,
            FrameHeightPoints = heightPoints,
            Crop = crop,
            RotationDegrees = rotation,
            FlipHorizontal = transform?.HorizontalFlip?.Value == true,
            FlipVertical = transform?.VerticalFlip?.Value == true,
            AlternativeText = isDecorative ? null : altText,
            Title = title,
            IsDecorative = isDecorative
        };
    }

    private static DxpImageCrop? NormalizeCrop(int? left, int? top, int? right, int? bottom)
    {
        static double Normalize(int? value) => Math.Max(0.0, Math.Min(1.0, (value ?? 0) / 100000.0));

        double l = Normalize(left);
        double t = Normalize(top);
        double r = Normalize(right);
        double b = Normalize(bottom);
        if (l <= 0 && t <= 0 && r <= 0 && b <= 0)
            return null;
        if (1.0 - l - r <= 0.000001 || 1.0 - t - b <= 0.000001)
            return null;
        return new DxpImageCrop(l, t, r, b);
    }

    private static bool IsOn(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return true;
        return value == "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "on", StringComparison.OrdinalIgnoreCase);
    }

    private static string? NormalizeAltText(string? altText)
    {
        if (string.IsNullOrWhiteSpace(altText))
            return null;

        var text = altText!;
        var sb = new StringBuilder(text.Length);
        bool previousWasWhitespace = false;

        foreach (var ch in text)
        {
            if (ch == '\r' || ch == '\n' || ch == '\t' || char.IsWhiteSpace(ch))
            {
                if (!previousWasWhitespace)
                {
                    sb.Append(' ');
                    previousWasWhitespace = true;
                }
                continue;
            }

            sb.Append(ch);
            previousWasWhitespace = false;
        }

        var normalized = sb.ToString().Trim();
        return normalized.Length == 0 ? null : normalized;
    }

    private static (string dataUri, string contentType)? TryBuildAnyDataUri(OpenXmlPart part)
    {
        try
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            var bytes = ms.ToArray();

            if (bytes.Length == 0)
                return null;

            var contentType = part.ContentType ?? "application/octet-stream";
            var base64 = Convert.ToBase64String(bytes);
            return ($"data:{contentType};base64,{base64}", contentType);
        }
        catch
        {
            return null;
        }
    }
}
