using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using DocxportNet.API;

namespace DocxportNet.Walker;

public static class DxpTimeline
{
	public static IReadOnlyList<DxpTimelineEvent> BuildTimeline(WordprocessingDocument doc)
	{
		var events = new List<DxpTimelineEvent>();

		// Core properties: created / modified
		var core = doc.PackageProperties;
		if (!string.IsNullOrWhiteSpace(core.Creator) || core.Created != null)
			events.Add(new DxpTimelineEvent("create", core.Creator, ToUtc(core.Created), null));
		if (!string.IsNullOrWhiteSpace(core.LastModifiedBy) || core.Modified != null)
			events.Add(new DxpTimelineEvent("modify", core.LastModifiedBy, ToUtc(core.Modified), null));

		// Comments: author / date
		var comments = doc.MainDocumentPart?.WordprocessingCommentsPart?.Comments;
		if (comments != null)
		{
			foreach (var c in comments.Elements<Comment>())
				events.Add(new DxpTimelineEvent("comment", c.Author?.Value ?? c.Initials?.Value, ToUtc(c.Date?.Value), null));
		}

		// Tracked changes: ins/del/move/conflicts
		foreach (var root in EnumerateRoots(doc))
		{
			AddTrackChanges<Inserted>(root, "change", events);
			AddTrackChanges<Deleted>(root, "change", events);
			AddTrackChanges<MoveFromRangeStart>(root, "change", events);
			AddTrackChanges<MoveToRangeStart>(root, "change", events);
			AddTrackChanges<DocumentFormat.OpenXml.Office2010.Word.ConflictInsertion>(root, "change", events);
			AddTrackChanges<DocumentFormat.OpenXml.Office2010.Word.ConflictDeletion>(root, "change", events);
		}

		return events
			.OrderBy(e => e.DateUtc ?? DateTime.MaxValue)
			.ToList();
	}

	private static IEnumerable<OpenXmlElement> EnumerateRoots(WordprocessingDocument doc)
	{
		if (doc.MainDocumentPart?.Document != null)
			yield return doc.MainDocumentPart.Document;

		foreach (var h in doc.MainDocumentPart?.HeaderParts ?? Enumerable.Empty<HeaderPart>())
			if (h.Header != null)
				yield return h.Header;

		foreach (var f in doc.MainDocumentPart?.FooterParts ?? Enumerable.Empty<FooterPart>())
			if (f.Footer != null)
				yield return f.Footer;

		if (doc.MainDocumentPart?.FootnotesPart?.Footnotes != null)
			yield return doc.MainDocumentPart.FootnotesPart.Footnotes;
		if (doc.MainDocumentPart?.EndnotesPart?.Endnotes != null)
			yield return doc.MainDocumentPart.EndnotesPart.Endnotes;
	}

	private static void AddTrackChanges<T>(OpenXmlElement root, string kind, List<DxpTimelineEvent> events)
		where T : OpenXmlElement
	{
		foreach (var item in root.Descendants<T>())
		{
			(string? author, string? dateStr) = ExtractAuthorDate(item);
			events.Add(new DxpTimelineEvent(kind, author, ToUtc(dateStr), null));
		}
	}

	private static (string? Author, string? Date) ExtractAuthorDate(OpenXmlElement el)
	{
		// TrackChangeType-derived elements expose Author/Date properties; use dynamic lookup with reflection to avoid many casts.
		var type = el.GetType();
		var authorProp = type.GetProperty("Author");
		var dateProp = type.GetProperty("Date");

		string? author = authorProp?.GetValue(el) as string ?? (authorProp?.GetValue(el) as OpenXmlSimpleType)?.InnerText;
		string? date = dateProp?.GetValue(el) as string ?? (dateProp?.GetValue(el) as OpenXmlSimpleType)?.InnerText;
		return (author, date);
	}

	private static DateTime? ToUtc(DateTime? dt) => dt?.ToUniversalTime();

	private static DateTime? ToUtc(string? dt)
	{
		if (string.IsNullOrWhiteSpace(dt))
			return null;
		if (DateTime.TryParse(dt, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out var parsed))
			return parsed.ToUniversalTime();
		return null;
	}
}
