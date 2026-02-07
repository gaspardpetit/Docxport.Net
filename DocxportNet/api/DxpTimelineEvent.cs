namespace DocxportNet.API;

public sealed record DxpTimelineEvent(string Kind, string? Author, DateTime? DateUtc, string? Detail);
