using DocumentFormat.OpenXml.Packaging;

namespace DocxportNet;

/// <summary>Identifies a Word revision within a DOCX package.</summary>
public static class DxpReviewId
{
    public static string Create(OpenXmlPart? part, string id)
    {
        if (part == null) throw new ArgumentNullException(nameof(part));
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("A revision ID is required.", nameof(id));
        return part.Uri.ToString().TrimStart('/') + "#" + id;
    }
}
