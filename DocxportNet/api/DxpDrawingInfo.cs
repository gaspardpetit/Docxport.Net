namespace DocxportNet.API;

public sealed record DxpDrawingInfo(
    string? EmbedRelId,
    string? ContentType,
    string? FileName,
    string? AltText,
    string? DataUri
);
