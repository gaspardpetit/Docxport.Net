using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;

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
		string? altText = docPr?.Description?.Value ?? docPr?.Title?.Value;

		var blip = drw.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
		var relId = blip?.Embed?.Value;

		string? contentType = null;
		string? fileName = null;
		string? dataUri = null;

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

		return new DxpDrawingInfo(relId, contentType, fileName, altText, dataUri);
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
