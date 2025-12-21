using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocxportNet.API;
using Microsoft.Extensions.Logging;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocxportNet.Core;
using DocxportNet.Word;
using DocxportNet.Markdown;

namespace DocxportNet.Visitors;

/// <summary>
/// Mutable state specific to DxpMarkdownVisitor; separated for clarity of intent.
/// Currently empty—fields should be added here instead of on the visitor itself.
/// </summary>
internal sealed class DxpMarkdownVisitorState
{
	public DxpContextStack<MarkdownTableBuilder> MarkdownTables { get; } = new DxpContextStack<MarkdownTableBuilder>("markdown-table");
	public bool InHeading { get; set; }
	public bool FontSpanOpen { get; set; }
	public bool AllCaps { get; set; }
	public bool IsFirstSection { get; set; } = true;
	public int SuppressDepth { get; set; }
	public Stack<StringBuilder> DeletedCaptures { get; } = new();
	public Stack<DxpChangeStream> ChangeStreams { get; } = new();
	public TextWriter PrevAcceptWriter = TextWriter.Null;
	public TextWriter PrevRejectWriter = TextWriter.Null;
	public TextWriter PrevUnchangedWriter = TextWriter.Null;
	public bool SawTrackedChange { get; set; }
}

public enum DxpTrackedChangeMode
{
	AcceptChanges,
	RejectChanges,
	InlineChanges,
	SplitChanges
}

internal enum DxpChangeStream
{
	Unchanged,
	Inserted,
	Deleted
}

public sealed record DxpMarkdownVisitorConfig
{
	public bool EmitImages = true;
	public bool EmitStyleFont = true;
	public bool EmitRunColor = true;
	public bool EmitRunBackground = true;
	public bool EmitTableBorders = true;
	public bool EmitDocumentColors = true;
	public bool EmitParagraphAlignment = true;
	public bool PreserveListSymbols = true;
	public bool RichTables = true;
	public bool UsePlainCodeBlocks = false;
	public bool UseMarkdownInlineStyles = false;
	public bool EmitSectionHeadersFooters = true;
	public bool EmitUnreferencedBookmarks = true;
	public bool EmitPageNumbers = false;
	public bool UsePlainComments = false;
	public bool EmitCustomProperties = true;
	public bool EmitTimeline = false;
	public DxpTrackedChangeMode TrackedChangeMode = DxpTrackedChangeMode.InlineChanges;

	public static DxpMarkdownVisitorConfig RICH = new();
	public static DxpMarkdownVisitorConfig PLAIN = new() {
		EmitImages = false,
		EmitStyleFont = false,
		EmitRunColor = false,
		EmitRunBackground = false,
		EmitTableBorders = false,
		EmitDocumentColors = false,
		EmitParagraphAlignment = false,
		PreserveListSymbols = false,
		RichTables = false,
		UsePlainCodeBlocks = true,
		UseMarkdownInlineStyles = false,
		EmitSectionHeadersFooters = true,
		EmitUnreferencedBookmarks = false,
		EmitPageNumbers = false,
		UsePlainComments = true,
		EmitCustomProperties = true,
		EmitTimeline = false
	};

	public static DxpMarkdownVisitorConfig DEFAULT = RICH;
}


public partial class DxpMarkdownVisitor : DxpVisitor, DxpIVisitor
{
	private TextWriter _writer;
	private readonly TextWriter _sinkWriter;
	private TextWriter _acceptWriter;
	private TextWriter _rejectWriter;
	private TextWriter _unchangedWriter;
	private DxpMarkdownVisitorConfig _config;
	private DxpMarkdownVisitorState _state = new();
	private bool AcceptMode => _config.TrackedChangeMode == DxpTrackedChangeMode.AcceptChanges;
	private bool RejectMode => _config.TrackedChangeMode == DxpTrackedChangeMode.RejectChanges;
	private bool InlineMode => _config.TrackedChangeMode == DxpTrackedChangeMode.InlineChanges;
	private bool SplitMode => _config.TrackedChangeMode == DxpTrackedChangeMode.SplitChanges;

	public DxpMarkdownVisitor(TextWriter writer, DxpMarkdownVisitorConfig config, ILogger? logger)
		: base(logger)
	{
		_config = config;
		_sinkWriter = writer;
		_acceptWriter = writer;
		_rejectWriter = writer;
		_unchangedWriter = writer;
		_writer = new RoutingWriter(this);
		_state.ChangeStreams.Push(DxpChangeStream.Unchanged);
	}

	private IDisposable SuppressOutputScope()
	{
		_state.SuppressDepth++;
		return Disposable.Create(() => _state.SuppressDepth--);
	}

	private IDisposable PushChangeStream(DxpChangeStream stream)
	{
		if (stream != DxpChangeStream.Unchanged)
			_state.SawTrackedChange = true;
		_state.ChangeStreams.Push(stream);
		return Disposable.Create(() => _state.ChangeStreams.Pop());
	}

	private IDisposable BeginSplitParagraphBuffers(out StringWriter left, out StringWriter right)
	{
		left = new StringWriter(new StringBuilder());
		right = new StringWriter(new StringBuilder());

		_state.PrevAcceptWriter = _acceptWriter;
		_state.PrevRejectWriter = _rejectWriter;
		_state.PrevUnchangedWriter = _unchangedWriter;

		_acceptWriter = right;
		_rejectWriter = left;
		_unchangedWriter = new MultiWriter(left, right);

		return Disposable.Create(() => {
			_acceptWriter = _state.PrevAcceptWriter;
			_rejectWriter = _state.PrevRejectWriter;
			_unchangedWriter = _state.PrevUnchangedWriter;
		});
	}

	private static IDisposable Combine(IDisposable first, IDisposable second)
	{
		return Disposable.Create(() => {
			second.Dispose();
			first.Dispose();
		});
	}

	private IDisposable WriteInsertedWrapper()
	{
		if (_config.RichTables && _config.EmitRunColor)
		{
			_writer.Write("""<span style="color:blue;text-decoration:underline;">""");
			return Disposable.Create(() => _writer.Write("</span>"));
		}

		_writer.Write("<u>");
		return Disposable.Create(() => _writer.Write("</u>"));
	}

	private bool IsCapturingDeleted() => _state.DeletedCaptures.Count > 0;
	private void CaptureDeleted(string text)
	{
		if (_state.DeletedCaptures.Count > 0)
			_state.DeletedCaptures.Peek().Append(text);
	}

	private IDisposable BeginDeletedCapture()
	{
		var sb = new StringBuilder();
		_state.DeletedCaptures.Push(sb);
		return Disposable.Create(() => {
			var captured = _state.DeletedCaptures.Pop().ToString();
			WriteInlineDeleted(captured);
		});
	}

	private void WriteInlineDeleted(string text)
	{
		string trimmed = text.Trim();
		bool isShort = trimmed.Length > 0 && trimmed.Length <= 3;

		if (isShort)
		{
			string payload = $"[[{trimmed}]]";
			if (_config.RichTables && _config.EmitRunColor)
			{
				_writer.Write($"""<span style="color:red;">{payload}</span>""");
			}
			else
			{
				_writer.Write(payload);
			}
			return;
		}

		if (_config.RichTables && _config.EmitRunColor)
		{
			_writer.Write($"""<span style="color:red;text-decoration:line-through;">{text}</span>""");
		}
		else
		{
			_writer.Write("<del>");
			_writer.Write(text);
			_writer.Write("</del>");
		}
	}

	private void EmitSplitRow(string left, string right)
	{
		if (string.Equals(left, right, StringComparison.Ordinal))
		{
			_sinkWriter.Write(left);
			if (!left.EndsWith("\n"))
				_sinkWriter.WriteLine();
			_sinkWriter.WriteLine();
			return;
		}

		// Prefer HTML table for consistent width in rich mode; fall back to markdown table otherwise.
		if (_config.RichTables)
		{
			_sinkWriter.WriteLine("<table style=\"width:100%;table-layout:fixed;\"><tr>");
			_sinkWriter.WriteLine($"""<td style="vertical-align:top;padding:4px;border:1px solid #ccc;width:50%;">{left}</td>""");
			_sinkWriter.WriteLine($"""<td style="vertical-align:top;padding:4px;border:1px solid #ccc;width:50%;">{right}</td>""");
			_sinkWriter.WriteLine("</tr></table>");
			_sinkWriter.WriteLine();
		}
		else
		{
			_sinkWriter.WriteLine("| Accepted | Rejected |");
			_sinkWriter.WriteLine("| --- | --- |");
			_sinkWriter.WriteLine($"| {SanitizeCell(left)} | {SanitizeCell(right)} |");
			_sinkWriter.WriteLine();
		}
	}

	private string SanitizeCell(string value)
	{
		return value.Replace("\n", "<br/>");
	}

	private void WriteSplitMarkers(DxpMarker? acceptMarker, DxpMarker? rejectMarker, bool isDeletedParagraph, bool isInsertedParagraph, bool needsParagraphWrapper)
	{
		// In split mode, common content goes to both sides; emit the appropriate marker to each side.
		if (isInsertedParagraph)
		{
			// Only the accept side should get a marker.
			if (acceptMarker?.marker != null)
			{
				var mk = NormalizeMarker(acceptMarker.marker);
				if (!needsParagraphWrapper && LooksLikeOrderedListMarker(mk))
					mk = EscapeOrderedListMarker(mk);
				_acceptWriter.Write(mk + " ");
			}
			return;
		}

		if (isDeletedParagraph)
		{
			// Only the reject side should get a marker.
			if (rejectMarker?.marker != null)
			{
				var mk = NormalizeMarker(rejectMarker.marker);
				if (!needsParagraphWrapper && LooksLikeOrderedListMarker(mk))
					mk = EscapeOrderedListMarker(mk);
				_rejectWriter.Write(mk + " ");
			}
			return;
		}

		// Normal paragraph: send to both sides (using their respective markers).
		if (acceptMarker?.marker != null)
		{
			var mk = NormalizeMarker(acceptMarker.marker);
			if (!needsParagraphWrapper && LooksLikeOrderedListMarker(mk))
				mk = EscapeOrderedListMarker(mk);
			_acceptWriter.Write(mk + " ");
		}
		if (rejectMarker?.marker != null)
		{
			var mk = NormalizeMarker(rejectMarker.marker);
			if (!needsParagraphWrapper && LooksLikeOrderedListMarker(mk))
				mk = EscapeOrderedListMarker(mk);
			_rejectWriter.Write(mk + " ");
		}
	}

	private TextWriter GetWriterFor(DxpChangeStream stream)
	{
		return stream switch {
			DxpChangeStream.Inserted => _acceptWriter,
			DxpChangeStream.Deleted => _rejectWriter,
			_ => _unchangedWriter
		};
	}

	private DxpChangeStream CurrentStream => _state.ChangeStreams.Count > 0 ? _state.ChangeStreams.Peek() : DxpChangeStream.Unchanged;

	private sealed class RoutingWriter : TextWriter
	{
		private readonly DxpMarkdownVisitor _owner;
		public RoutingWriter(DxpMarkdownVisitor owner)
		{
			_owner = owner;
		}

		public override Encoding Encoding => Encoding.UTF8;

		private TextWriter Target => _owner.GetWriterFor(_owner.CurrentStream);

		public override void Write(char value) => Target.Write(value);
		public override void Write(char[] buffer, int index, int count) => Target.Write(buffer, index, count);
		public override void Write(string? value) => Target.Write(value);
		public override void WriteLine() => Target.WriteLine();
		public override void WriteLine(string? value) => Target.WriteLine(value);
		public override void WriteLine(char value) => Target.WriteLine(value);
		public override void WriteLine(char[] buffer, int index, int count) => Target.WriteLine(buffer, index, count);
	}

	private sealed class MultiWriter : TextWriter
	{
		private readonly TextWriter _left;
		private readonly TextWriter _right;
		public MultiWriter(TextWriter left, TextWriter right)
		{
			_left = left;
			_right = right;
		}
		public override Encoding Encoding => Encoding.UTF8;

		public override void Write(char value) { _left.Write(value); _right.Write(value); }
		public override void Write(char[] buffer, int index, int count) { _left.Write(buffer, index, count); _right.Write(buffer, index, count); }
		public override void Write(string? value) { _left.Write(value); _right.Write(value); }
		public override void WriteLine() { _left.WriteLine(); _right.WriteLine(); }
		public override void WriteLine(string? value) { _left.WriteLine(value); _right.WriteLine(value); }
		public override void WriteLine(char value) { _left.WriteLine(value); _right.WriteLine(value); }
		public override void WriteLine(char[] buffer, int index, int count) { _left.WriteLine(buffer, index, count); _right.WriteLine(buffer, index, count); }
	}

	public override void VisitText(Text t, DxpIDocumentContext d)
	{
		if (ShouldSuppressOutput(d))
			return;
		string text = t.Text;
		if (InlineMode && IsCapturingDeleted())
		{
			CaptureDeleted(text);
			return;
		}
		if (_state.AllCaps)
		{
			var culture = CultureInfo.InvariantCulture;
			text = text.ToUpper(culture);
		}

		_writer.Write(text);
	}

	public override void StyleAllCapsEnd(DxpIDocumentContext d)
	{
		_state.AllCaps = false;
	}

	public override void StyleAllCapsBegin(DxpIDocumentContext d)
	{
		_state.AllCaps = true;
	}

	public override void StyleSmallCapsEnd(DxpIDocumentContext d)
	{
		_state.AllCaps = false;
	}

	public override void StyleSmallCapsBegin(DxpIDocumentContext d)
	{
		_state.AllCaps = true;
	}

	public override void StyleBoldBegin(DxpIDocumentContext d)
	{
		if (_state.InHeading)
			return;
		_writer.Write("<b>");
	}
	public override void StyleBoldEnd(DxpIDocumentContext d)
	{
		if (_state.InHeading)
			return;
		_writer.Write("</b>");
	}


	public override void StyleItalicBegin(DxpIDocumentContext d)
	{
		_writer.Write("<i>");
	}
	public override void StyleItalicEnd(DxpIDocumentContext d)
	{
		_writer.Write("</i>");
	}

	public override void StyleUnderlineBegin(DxpIDocumentContext d)
	{
		_writer.Write("<u>");
	}
	public override void StyleUnderlineEnd(DxpIDocumentContext d)
	{
		_writer.Write("</u>");
	}

	public override void StyleStrikeBegin(DxpIDocumentContext d)
	{
		_writer.Write("<del>");
	}
	public override void StyleStrikeEnd(DxpIDocumentContext d)
	{
		_writer.Write("</del>");
	}

	public override void StyleDoubleStrikeBegin(DxpIDocumentContext d)
	{
		_writer.Write("<del>");
	}
	public override void StyleDoubleStrikeEnd(DxpIDocumentContext d)
	{
		_writer.Write("</del>");
	}

	public override void StyleSuperscriptBegin(DxpIDocumentContext d)
	{
		_writer.Write("<sup>");
	}
	public override void StyleSuperscriptEnd(DxpIDocumentContext d)
	{
		_writer.Write("</sup>");
	}

	public override void StyleSubscriptBegin(DxpIDocumentContext d)
	{
		_writer.Write("<sub>");
	}
	public override void StyleSubscriptEnd(DxpIDocumentContext d)
	{
		_writer.Write("</sub>");
	}

	public override void StyleFontBegin(string? fontName, int? fontSizeHalfPoints, DxpIDocumentContext d)
	{
		if (_config.EmitStyleFont == false)
			return;

		if (IsDefaultFont(fontName, fontSizeHalfPoints, d))
		{
			_state.FontSpanOpen = false;
			return;
		}

		_state.FontSpanOpen = true;
		_writer.Write($"""<span style="font-family: {fontName}; font-size: {fontSizeHalfPoints / 2.0}pt;">""");
	}

	public override void StyleFontEnd(DxpIDocumentContext d)
	{
		if (_config.EmitStyleFont == false)
			return;

		if (_state.FontSpanOpen)
		{
			_writer.Write("</span>");
			_state.FontSpanOpen = false;
		}
	}

	private bool IsDefaultFont(string? fontName, int? fontSizeHalfPoints, DxpIDocumentContext d)
	{
		if (d.DefaultRunStyle.FontName == null && d.DefaultRunStyle.FontSizeHalfPoints == null)
			return false;

		bool nameMatch = d.DefaultRunStyle.FontName == null || string.Equals(d.DefaultRunStyle.FontName, fontName, StringComparison.OrdinalIgnoreCase);
		bool sizeMatch = d.DefaultRunStyle.FontSizeHalfPoints == null || d.DefaultRunStyle.FontSizeHalfPoints == fontSizeHalfPoints;
		return nameMatch && sizeMatch;
	}

	public override void VisitBookmarkStart(BookmarkStart bs, DxpIDocumentContext d)
	{
		string? name = bs.Name?.Value;
		string? id = bs.Id?.Value;

		if (_config.EmitUnreferencedBookmarks == false)
		{
			if (!string.IsNullOrEmpty(name) && !d.ReferencedBookmarkAnchors.Contains(name!))
				// skip unreferenced links
				return;
		}

		// Skip Word’s internal _GoBack if desired; current behavior surfaces it.
		if (!string.IsNullOrEmpty(name))
			_writer.Write($"<a id=\"{Escape(name!)}\" data-bookmark-id=\"{id}\"></a>");
	}

	public override void VisitBookmarkEnd(BookmarkEnd be, DxpIDocumentContext d)
	{
		// Usually nothing to emit; it just closes the range.
	}

	public override IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target, DxpIDocumentContext d)
	{
		string? href = target?.uri;
		_writer.Write(href != null ? $"<a href=\"{HtmlAttr(href)}\">" : "<a>");
		return Disposable.Create(() => _writer.Write("</a>"));
	}

	public string Escape(string name)
	{
		return name;
	}

	public override IDisposable VisitInsertedBegin(Inserted ins, DxpIDocumentContext d)
	{
		if (AcceptMode)
			return Disposable.Empty;
		if (RejectMode)
			return SuppressOutputScope();
		if (SplitMode)
			return PushChangeStream(DxpChangeStream.Inserted);
		if (InlineMode)
			return Combine(PushChangeStream(DxpChangeStream.Inserted), WriteInsertedWrapper());

		return PushChangeStream(DxpChangeStream.Inserted);
	}

	public override IDisposable VisitDeletedBegin(Deleted del, DxpIDocumentContext d)
	{
		if (AcceptMode)
			return SuppressOutputScope();
		if (RejectMode)
			return Disposable.Empty;
		if (SplitMode)
			return PushChangeStream(DxpChangeStream.Deleted);
		if (InlineMode)
			return Combine(PushChangeStream(DxpChangeStream.Deleted), BeginDeletedCapture());

		_writer.Write("<edit-delete>");
		return Disposable.Create(() => {
			_writer.Write("</edit-delete>");
		});
	}

	public override IDisposable VisitDeletedRunBegin(DeletedRun dr, DxpIDocumentContext d)
	{
		if (AcceptMode)
			return SuppressOutputScope();
		if (RejectMode)
			return Disposable.Empty;
		if (SplitMode)
			return PushChangeStream(DxpChangeStream.Deleted);
		if (InlineMode)
			return Combine(PushChangeStream(DxpChangeStream.Deleted), BeginDeletedCapture());

		_writer.Write("<edit-delete>");
		return Disposable.Create(() => {
			_writer.Write("</edit-delete>");
		});
	}

	public override void VisitDeletedParagraphMark(Deleted del, ParagraphProperties pPr, Paragraph? p, DxpIDocumentContext d)
	{
		// Paragraph mark deletions carry no inline content; outer VisitDeletedBegin wrapper is enough.
	}

	public override IDisposable VisitInsertedRunBegin(InsertedRun ir, DxpIDocumentContext d)
	{
		if (AcceptMode)
			return Disposable.Empty;
		if (RejectMode)
			return SuppressOutputScope();
		if (SplitMode)
			return PushChangeStream(DxpChangeStream.Inserted);
		if (InlineMode)
			return Combine(PushChangeStream(DxpChangeStream.Inserted), WriteInsertedWrapper());

		_writer.Write("<edit-insert>");
		return Disposable.Create(() => {
			_writer.Write("</edit-insert>");
		});
	}

	public override void VisitDeletedText(DeletedText dt, DxpIDocumentContext d)
	{
		if (ShouldSuppressOutput(d))
			return;
		if (InlineMode && IsCapturingDeleted())
		{
			CaptureDeleted(dt.Text);
			return;
		}
		_writer.Write(dt.Text);
	}

	public override void VisitNoBreakHyphen(NoBreakHyphen h, DxpIDocumentContext d)
	{
		if (ShouldSuppressOutput(d))
			return;
		if (InlineMode && IsCapturingDeleted())
		{
			CaptureDeleted("-");
			return;
		}
		_writer.Write("-");
	}

	public override void VisitDrawingBegin(Drawing drw, DxpDrawingInfo? info, DxpIDocumentContext d)
	{
		if (_config.EmitImages == false)
		{
			_writer.Write("[IMAGE]");
			return;
		}

		var alt = HtmlAttr(info?.AltText ?? "image");
		var dataUri = info?.DataUri;
		var contentType = info?.ContentType ?? "";

		if (!string.IsNullOrEmpty(dataUri) && contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
		{
			_writer.Write($"<img src=\"{dataUri}\" alt=\"{alt}\" />");
		}
		else if (!string.IsNullOrEmpty(dataUri))
		{
			_writer.Write($"<object data=\"{dataUri}\" type=\"{HtmlAttr(contentType)}\">[DRAWING: {alt}]</object>");
		}
		else
		{
			var meta = string.IsNullOrEmpty(contentType) ? "" : $" ({contentType})";
			_writer.Write($"[DRAWING: {alt}{meta}]");
		}
	}

	public override void VisitDocumentProperties(IPackageProperties core, IReadOnlyList<CustomFileProperty> custom, IReadOnlyList<DxpTimelineEvent> timeline)
	{
		var lines = new List<string>();

		void Add(string label, string? value)
		{
			if (!string.IsNullOrWhiteSpace(value))
				lines.Add($"<!-- {label}: {value} -->");
		}

		Add("Title", core.Title);
		Add("Subject", core.Subject);
		Add("Author", core.Creator);
		Add("Description", core.Description);
		Add("Category", core.Category);
		Add("Keywords", core.Keywords);
		Add("LastModifiedBy", core.LastModifiedBy);
		Add("Revision", core.Revision);
		Add("Created", FormatDateUtc(core.Created));
		Add("Modified", FormatDateUtc(core.Modified));

		if (_config.EmitTimeline && _config.RichTables == false && timeline != null && timeline.Count > 0)
		{
			foreach (var ev in timeline)
			{
				var date = ev.DateUtc?.ToString("yyyy-MM-ddTHH:mm:ss'Z'") ?? "unknown";
				var who = string.IsNullOrWhiteSpace(ev.Author) ? "unknown" : ev.Author;
				var detail = string.IsNullOrWhiteSpace(ev.Detail) ? "" : $" ({ev.Detail})";
				lines.Add($"<!-- Timeline: {date} - {ev.Kind} by {who}{detail} -->");
			}
		}

		if (custom != null && _config.EmitCustomProperties)
		{
			foreach (var prop in custom)
			{
				if (prop.Value != null)
				lines.Add($"<!-- {prop.Name}: {prop.Value} -->");
			}
		}

		foreach (var line in lines)
			_writer.WriteLine(line);

		if (lines.Count > 0)
			_writer.WriteLine();

		if (_config.EmitTimeline && _config.RichTables && timeline != null && timeline.Count > 0)
		{
			_writer.WriteLine();
			_writer.WriteLine("| Date | Event |");
			_writer.WriteLine("| --- | --- |");
			foreach (var ev in timeline)
			{
				var date = ev.DateUtc?.ToString("yyyy-MM-dd HH:mm:ss 'UTC'") ?? "unknown";
				var who = string.IsNullOrWhiteSpace(ev.Author) ? "unknown" : ev.Author;
				var detail = string.IsNullOrWhiteSpace(ev.Detail) ? "" : $" ({ev.Detail})";
				_writer.WriteLine($"| {date} | {ev.Kind} by {who}{detail} |");
			}
			_writer.WriteLine();
		}
	}

	static string? FormatDateUtc(DateTime? value)
	{
		if (value == null)
			return null;

		// PackageProperties returns local times; normalize to UTC to avoid timezone-dependent output.
		return value.Value.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss'Z'");
	}

	static string HtmlAttr(string s) =>
		s.Replace("&", "&amp;").Replace("\"", "&quot;").Replace("<", "&lt;").Replace(">", "&gt;");


	public override IDisposable VisitParagraphBegin(Paragraph p, DxpIDocumentContext d, DxpIParagraphContext paragraph)
	{
		if (ShouldSuppressOutput(d))
			return Disposable.Empty;

		bool splitMode = _config.TrackedChangeMode == DxpTrackedChangeMode.SplitChanges
			&& CurrentStream == DxpChangeStream.Unchanged
			&& _state.SawTrackedChange
			&& !SplitMode; // paragraph-level split only when not capturing the whole body
		IDisposable splitScope = Disposable.Empty;
		StringWriter? splitLeft = null;
		StringWriter? splitRight = null;
		if (splitMode)
			splitScope = BeginSplitParagraphBuffers(out splitLeft, out splitRight);

		bool isDeletedParagraph =
			paragraph.Properties?.GetFirstChild<Deleted>() != null ||
			paragraph.Properties?.ParagraphMarkRunProperties?.GetFirstChild<Deleted>() != null;
		bool isInsertedParagraph =
			paragraph.Properties?.ParagraphMarkRunProperties?.GetFirstChild<Inserted>() != null;
		var marker =
			RejectMode ? paragraph.MarkerReject :
			(InlineMode && isDeletedParagraph ? paragraph.MarkerReject : paragraph.MarkerAccept);
		var indent = paragraph.Indent;
		if (InlineMode && isDeletedParagraph && marker?.marker != null)
		{
			if (_config.RichTables && _config.EmitRunColor)
				marker = new DxpMarker($"""<span style="color:red;text-decoration:line-through;">{marker.marker}</span>""", marker.numId, marker.iLvl);
			else
				marker = new DxpMarker($"<del>{marker.marker}</del>", marker.numId, marker.iLvl);
		}
		else if (InlineMode && isInsertedParagraph && marker?.marker != null)
		{
			if (_config.RichTables && _config.EmitRunColor)
				marker = new DxpMarker($"""<span style="color:blue;text-decoration:underline;">{marker.marker}</span>""", marker.numId, marker.iLvl);
			else
				marker = new DxpMarker($"<u>{marker.marker}</u>", marker.numId, marker.iLvl);
		}

		if (AcceptMode && isDeletedParagraph)
			return SuppressOutputScope();
		if (RejectMode && isInsertedParagraph)
			return SuppressOutputScope();

		// Avoid splitting paragraphs with no tracked content; stick to normal rendering to prevent duplication.
		if (splitMode && !isDeletedParagraph && !isInsertedParagraph && CurrentStream == DxpChangeStream.Unchanged)
		{
			splitScope.Dispose();
			splitScope = Disposable.Empty;
			splitMode = false;
		}

		string innerText = p.InnerText;

		if (string.IsNullOrWhiteSpace(innerText))
		{
			return Disposable.Create(() => {
				_writer.WriteLine();
			});
		}


		var styleChain = d.Styles.GetParagraphStyleChain(p);

		var justification = _config.EmitParagraphAlignment
			? p.ParagraphProperties?.Justification?.Val?.Value
			: null;
		string? justify = null;
		if (justification == JustificationValues.Center)
			justify = "center";
		else if (justification == JustificationValues.Right)
			justify = "right";
		else if (justification == JustificationValues.Both || justification == JustificationValues.Distribute)
			justify = "justify";

		var headingLevel = d.Styles.GetHeadingLevel(p);
		bool hasNumbering = marker?.numId != null;
		bool isHeading = headingLevel != null && !hasNumbering;
		var paragraphStyleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

		// Subtitle -> treat as a sub heading if not already a heading
		if (!isHeading && !hasNumbering && styleChain.Any(sc => string.Equals(sc.StyleId, DxpWordBuiltInStyleId.wdStyleSubtitle, StringComparison.OrdinalIgnoreCase)))
		{
			headingLevel = 2;
			isHeading = true;
		}

		if (!_config.EmitPageNumbers && styleChain.Any(sc => string.Equals(sc.StyleId, DxpWordBuiltInStyleId.wdStylePageNumber, StringComparison.OrdinalIgnoreCase)))
		{
			return Disposable.Empty;
		}

		bool isBlockQuote = styleChain.Any(sc =>
			string.Equals(sc.StyleId, DxpWordBuiltInStyleId.wdStyleQuote, StringComparison.OrdinalIgnoreCase) ||
			string.Equals(sc.StyleId, DxpWordBuiltInStyleId.wdStyleIntenseQuote, StringComparison.OrdinalIgnoreCase) ||
			string.Equals(sc.StyleId, DxpWordBuiltInStyleId.wdStyleBlockQuotation, StringComparison.OrdinalIgnoreCase));

		bool isCaption = styleChain.Any(sc =>
			string.Equals(sc.StyleId, DxpWordBuiltInStyleId.wdStyleCaption, StringComparison.OrdinalIgnoreCase));

		bool isCode = styleChain.Any(sc =>
			string.Equals(sc.StyleId, DxpWordBuiltInStyleId.wdStyleHtmlPre, StringComparison.OrdinalIgnoreCase) ||
			string.Equals(sc.StyleId, DxpWordBuiltInStyleId.wdStyleHtmlCode, StringComparison.OrdinalIgnoreCase) ||
			string.Equals(sc.StyleId, "Code", StringComparison.OrdinalIgnoreCase));

		string? headingStyle = null;
		if (isHeading && justify != null && _config.EmitParagraphAlignment && _config.EmitStyleFont)
		{
			headingStyle = $" style=\"text-align:{justify};\"";
		}

		if (isHeading)
		{
			if (_config.EmitStyleFont)
			{
				var styleAttr = headingStyle ?? string.Empty;
				_writer.Write($"<h{headingLevel}{styleAttr}>");
			}
			else
			{
				for (int i = 0; i < headingLevel; ++i)
					_writer.Write($"#");
				_writer.Write($" ");
			}
		}

		double adjustedMargin = indent.Left.HasValue ? AdjustMarginLeft(DxpTwipValue.ToPoints(indent.Left.Value), d) : 0.0;
		bool hasMargin = adjustedMargin > 0.0001;

		bool needsParagraphWrapper = !isHeading && (hasMargin || justify != null);
		if (needsParagraphWrapper)
		{
			var para = new StringBuilder();
			para.Append("<p style=\"");
			if (hasMargin)
				para.Append("margin-left:").Append(adjustedMargin.ToString("0.###", CultureInfo.InvariantCulture)).Append("pt;");
			if (justify != null)
				para.Append("text-align:").Append(justify).Append(';');
			para.Append("\">");
			_writer.Write(para.ToString());
		}

		if (isHeading)
			marker = null;

		if (isBlockQuote)
			marker = null;

		if (isCode)
			marker = null;

		if (_config.TrackedChangeMode == DxpTrackedChangeMode.SplitChanges)
		{
			WriteSplitMarkers(paragraph.MarkerAccept, paragraph.MarkerReject, isDeletedParagraph, isInsertedParagraph, needsParagraphWrapper);
		}
		else if (marker?.marker != null)
		{
			var normalizedMarker = NormalizeMarker(marker.marker);
			if (!needsParagraphWrapper && LooksLikeOrderedListMarker(normalizedMarker))
				normalizedMarker = EscapeOrderedListMarker(normalizedMarker);
			_writer.Write($"""{normalizedMarker} """);
		}
		bool previousHeading = _state.InHeading;
		if (isHeading)
			_state.InHeading = true;

		if (isBlockQuote)
		{
			if (_config.EmitStyleFont)
				_writer.Write("<blockquote>");
			else
				_writer.Write("> ");
		}

		if (isCode)
		{
			if (_config.UsePlainCodeBlocks)
			{
				_writer.Write("```\n");
			}
			else
			{
				_writer.Write("<pre><code>");
			}
		}

		if (isCaption && !_config.UsePlainCodeBlocks && !isHeading && !isBlockQuote)
		{
			_writer.Write("<figcaption>");
		}

		var baseDispose = Disposable.Create(() => {
			if (isCaption && !_config.UsePlainCodeBlocks && !isHeading && !isBlockQuote)
			{
				_writer.Write("</figcaption>");
			}
			_state.InHeading = previousHeading;
			if (isCode)
			{
				if (_config.UsePlainCodeBlocks)
					_writer.WriteLine("\n```");
				else
					_writer.Write("</code></pre>");
			}
			if (isHeading && _config.EmitStyleFont)
			{
				_writer.Write($"</h{headingLevel}>");
			}
			if (needsParagraphWrapper)
				_writer.WriteLine("</p>");
			else if (isBlockQuote && _config.EmitStyleFont)
				_writer.WriteLine("</blockquote>");
			int newlines = 2;
			for (int i = 0; i < newlines; i++)
				_writer.WriteLine();
		});
		if (!splitMode)
			return baseDispose;

		return Disposable.Create(() => {
			baseDispose.Dispose();
			if (splitLeft != null && splitRight != null)
				EmitSplitRow(splitLeft.ToString(), splitRight.ToString());
			splitScope.Dispose();
		});

	}

	public override void VisitFootnoteReference(FootnoteReference fr, DxpIFootnoteContext footnote, DxpIDocumentContext d)
	{
		_writer.Write($"<a href=\"#fn-{footnote.Id}\" id=\"fnref-{footnote.Id}\">[{footnote.Index}]</a>");
	}

	public override IDisposable VisitSectionHeaderBegin(Header hdr, object kind, DxpIDocumentContext d)
	{
		if (_config.EmitSectionHeadersFooters == false)
			return Disposable.Empty;

		// Capture header content; if anything was rendered, wrap it with a bottom border.
		var buffer = new StringWriter();
		var previous = _writer;
		_writer = buffer;
		return Disposable.Create(() => {
			_writer = previous;
			var content = buffer.ToString();
			if (HasVisibleContent(content))
			{
				_writer.WriteLine("""<div class="header" style="border-bottom:1px solid #000;">""");
				_writer.Write(content);
				if (!content.EndsWith("\n"))
					_writer.WriteLine();
				_writer.WriteLine("</div>");
			}
		});
	}

	public override IDisposable VisitSectionFooterBegin(Footer ftr, object kind, DxpIDocumentContext d)
	{
		// Capture footer content; if anything was rendered, wrap it with a top border.
		var buffer = new StringWriter();
		var previous = _writer;
		_writer = buffer;
		return Disposable.Create(() => {
			_writer = previous;
			var content = buffer.ToString();
			if (HasVisibleContent(content))
			{
				_writer.WriteLine("""<div class="footer" style="border-top:1px solid #000;">""");
				_writer.Write(content);
				if (!content.EndsWith("\n"))
					_writer.WriteLine();
				_writer.WriteLine("</div>");
			}
		});
	}

	public override void VisitPageNumber(PageNumber pn, DxpIDocumentContext d)
	{
		// Rendered via field result; we suppress field output when EmitPageNumbers is false.
	}

	public override void VisitComplexFieldBegin(FieldChar begin, DxpIDocumentContext d)
	{
	}

	public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
	{
		if (!_config.EmitPageNumbers && LooksLikePageField(text) && d.CurrentFields.Current != null)
			d.CurrentFields.Current.SuppressResult = true;
	}

	public override IDisposable VisitComplexFieldResultBegin(DxpIDocumentContext d)
	{
		return Disposable.Empty;
	}

	public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
	{
	}

	public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
	{
		var instr = fld.Instruction?.Value;
		bool suppress = !_config.EmitPageNumbers && LooksLikePageField(instr);
		if (suppress && d.CurrentFields.Current != null)
			d.CurrentFields.Current.SuppressResult = true;
		return Disposable.Empty;
	}

	public override IDisposable VisitFootnoteBegin(Footnote fn, DxpIFootnoteContext footnote, DxpIDocumentContext d)
	{
		_writer.Write($"""\n<div class="footnote" id="fn-{footnote.Id}">\n\n""");
		return Disposable.Create(() => _writer.Write("</div>\n"));
	}

	public override void VisitFootnoteReferenceMark(FootnoteReferenceMark m, DxpIFootnoteContext footnote, DxpIDocumentContext d)
	{
		if (footnote.Index != null)
			_writer.Write($"{footnote.Index}");
	}


	public override IDisposable VisitTableBegin(Table t, DxpTableModel model, DxpIDocumentContext d, DxpITableContext table)
	{
		var styles = _config.EmitTableBorders && table.Properties != null
			? BuildTableStyle(table.Properties)
			: (null, null);

		if (!_config.RichTables)
		{
			var mdTable = new MarkdownTableBuilder();
			var mdScope = _state.MarkdownTables.Push(mdTable);
			return Disposable.Create(() => {
				mdTable.Render(_writer);
				mdScope.Dispose();
			});
		}

		var currentStyle = styles.tableStyle;

		_writer.Write("<table");
		if (!string.IsNullOrEmpty(currentStyle))
			_writer.Write($" style=\"{currentStyle}\"");
		_writer.WriteLine(">");
		return Disposable.Create(() => {
			_writer.WriteLine("</table>");
		});
	}

	public override IDisposable VisitTableRowBegin(TableRow tr, DxpITableRowContext row, DxpIDocumentContext d)
	{
		var isHeader = row.IsHeader;

		var mdTable = _state.MarkdownTables.Current;
		if (mdTable != null)
		{
			mdTable.BeginRow(isHeader);
			return Disposable.Create(() => mdTable.EndRow());
		}

		if (isHeader)
			_writer.WriteLine("  <tr class=\"header-row\">");
		else
			_writer.WriteLine("  <tr>");
		return Disposable.Create(() => _writer.WriteLine("  </tr>"));
	}

	public override IDisposable VisitTableCellBegin(TableCell tc, DxpITableCellContext cell, DxpIDocumentContext d)
	{
		var mdTable = _state.MarkdownTables.Current;
		if (mdTable != null)
		{
			var cellWriter = new StringWriter();
			var previous = _writer;
			_writer = cellWriter;
			return Disposable.Create(() => {
				_writer = previous;
				mdTable.AddCell(cellWriter.ToString());
			});
		}

		var spans = (cell.RowSpan, cell.ColSpan);
		var cellBorders = cell.Properties?.TableCellBorders;
		var cellStyle = _config.EmitTableBorders ? BuildCellStyle(cellBorders) : null;

		_writer.Write("    <td");
		if (spans.Item1 > 1)
			_writer.Write($" rowspan=\"{spans.Item1}\"");
		if (spans.Item2 > 1)
			_writer.Write($" colspan=\"{spans.Item2}\"");
		string? borderCss = null;
		if (_config.EmitTableBorders && cell.Row.Table.Properties != null)
		{
			borderCss = BuildTableStyle(cell.Row.Table.Properties).cellBorderStyle;
		}
		var effectiveCellStyle = cellStyle ?? (borderCss != null ? $"border:{borderCss};" : null);
		if (!string.IsNullOrEmpty(effectiveCellStyle))
			_writer.Write($" style=\"{effectiveCellStyle}\"");
		_writer.Write(">");

		return Disposable.Create(() => {
			_writer.WriteLine("</td>");
		});
	}

	private static string? BuildCellStyle(TableCellBorders? borders)
	{
		if (borders == null)
			return null;

		var b = PickBorder(borders);
		if (b == null)
			return null;

		string? css = BuildBorderCss(b);
		return css != null ? $"border:{css};" : null;
	}

	private static (string? tableStyle, string? cellBorderStyle) BuildTableStyle(TableProperties tp)
	{
		var b = PickBorder(tp.TableBorders);
		if (b == null)
			return (null, null);

		string? borderCss = BuildBorderCss(b);
		if (borderCss == null)
			return (null, null);

		var sb = new StringBuilder();
		sb.Append("border:").Append(borderCss).Append(";");
		sb.Append("border-collapse:collapse;");
		return (sb.ToString(), borderCss);
	}

	private static string? BuildBorderCss(BorderType? b)
	{
		if (b == null)
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

		return pt.ToString("0.###", CultureInfo.InvariantCulture) + "pt solid " + color;
	}

	private static BorderType? PickBorder(TableBorders? borders)
	{
		if (borders == null)
			return null;

		foreach (var b in new BorderType?[]
			{
				borders.TopBorder,
				borders.LeftBorder,
				borders.BottomBorder,
				borders.RightBorder,
				borders.InsideHorizontalBorder,
				borders.InsideVerticalBorder
			})
		{
			if (b != null)
				return b;
		}

		return null;
	}

	private static BorderType? PickBorder(TableCellBorders borders)
	{
		foreach (var b in new BorderType?[]
			{
				borders.TopBorder,
				borders.LeftBorder,
				borders.BottomBorder,
				borders.RightBorder,
				borders.InsideHorizontalBorder,
				borders.InsideVerticalBorder
			})
		{
			if (b != null)
				return b;
		}

		return null;
	}

	public override IDisposable VisitBlockBegin(OpenXmlElement child, DxpIDocumentContext d)
	{
		// optional: do nothing
		return Disposable.Empty;
	}

	public override IDisposable VisitCommentBegin(DxpCommentInfo c, DxpCommentThread thread, DxpIDocumentContext d)
	{
		if (_config.UsePlainComments)
		{
			var label = c.IsReply ? "REPLY BY" : "COMMENT BY";
			var who = !string.IsNullOrEmpty(c.Author)
				? c.Author!
				: (!string.IsNullOrEmpty(c.Initials) ? c.Initials! : "Unknown");
			var when = c.DateUtc?.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss'Z'");

			if (!string.IsNullOrEmpty(when))
				_writer.WriteLine($"  {label} {who} ON {when}");
			else
				_writer.WriteLine($"  {label} {who}");

			_writer.WriteLine();

			return Disposable.Create(() => {
				_writer.WriteLine();
			});
		}
		else
		{
			var commentStyle = "background:#fffcf0;border:1px solid #d9b200;border-radius:4px;padding:6px;margin-bottom:6px;";

			var label = BuildCommentLabel(c);
			if (!string.IsNullOrEmpty(label))
				_writer.WriteLine("  " + label);

			_writer.Write($"""  <div class="comment" style="{commentStyle}">""");
			_writer.WriteLine();

			return Disposable.Create(() => {
				_writer.WriteLine("  </div>");
				_writer.WriteLine();
			});
		}
	}

	public override IDisposable VisitCommentThreadBegin(string anchorId, DxpCommentThread thread, DxpIDocumentContext d)
	{
		if (thread.Comments == null || thread.Comments.Count == 0)
			return Disposable.Empty;

		if (_config.UsePlainComments)
		{
			return EmitPlainCommentThread(thread);
		}

		var commentsStyle = "background:#fff8c6;border:1px solid #e6c44a;border-radius:6px;padding:8px;margin:8px 0 8px 12px;float:right;max-width:45%;";

		_writer.Write($"""<div class="comments" style="{commentsStyle}">""");

		return Disposable.Create(() => {
			_writer.WriteLine("</div>");
			_writer.WriteLine();
		});
	}

	private IDisposable EmitPlainCommentThread(DxpCommentThread thread)
	{
		_writer.WriteLine("<!--");
		_writer.WriteLine();

		return Disposable.Create(() => {
			_writer.WriteLine("-->");
			_writer.WriteLine();
		});
	}

	private string BuildCommentLabel(DxpCommentInfo c)
	{
		var who = !string.IsNullOrEmpty(c.Author)
			? c.Author!
			: (!string.IsNullOrEmpty(c.Initials) ? c.Initials! : "Unknown");
		var when = c.DateUtc?.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss'Z'") ?? string.Empty;

		if (string.IsNullOrEmpty(who) && string.IsNullOrEmpty(when))
			return string.Empty;

		return $"<span style=\"font-size:small\">{HtmlAttr(who)} | {HtmlAttr(when)}</span>";
	}

	public override void VisitBreak(Break br, DxpIDocumentContext d)

	{
		if (ShouldSuppressOutput(d))
			return;
		_writer.Write("<br/>");
	}

	public override void VisitCarriageReturn(CarriageReturn cr, DxpIDocumentContext d)
	{
		if (ShouldSuppressOutput(d))
			return;
		_writer.Write("<br/>");
	}

	public override void VisitTab(TabChar tab, DxpIDocumentContext d)
	{
		if (ShouldSuppressOutput(d))
			return;
		if (InlineMode && IsCapturingDeleted())
		{
			CaptureDeleted("\t");
			return;
		}
		_writer.Write("&#9;"); // or &nbsp; spacing
	}

	public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
	{
		var style = BuildRunStyle(r.RunProperties);
		bool hasText = r.ChildElements.OfType<Text>().Any(t => !string.IsNullOrEmpty(t.Text));

		if (string.IsNullOrEmpty(style) || !hasText)
			return Disposable.Empty;

		_writer.Write($"<span style=\"{style}\">");
		return Disposable.Create(() => _writer.Write("</span>"));
	}

	private string BuildRunStyle(RunProperties? rp)
	{
		if (rp == null)
			return string.Empty;

		var parts = new List<string>();

		var colorVal = rp.Color?.Val?.Value;
		if (_config.EmitRunColor && !string.IsNullOrEmpty(colorVal) && !string.Equals(colorVal, "auto", StringComparison.OrdinalIgnoreCase))
			parts.Add($"color:{ToCssColor(colorVal!)}");

		if (_config.EmitRunBackground)
		{
			var highlight = HighlightToColor(rp.Highlight?.Val);
			var shading = rp.Shading?.Fill?.Value;
			string? background = highlight ?? (!string.IsNullOrEmpty(shading) ? ToCssColor(shading!) : null);
			if (!string.IsNullOrEmpty(background))
				parts.Add($"background-color:{background}");
		}

		return string.Join(";", parts);
	}

	private static string? HighlightToColor(EnumValue<HighlightColorValues>? highlight)
	{
		var value = highlight?.Value;
		if (value == HighlightColorValues.Yellow)
			return "#ffff00";
		if (value == HighlightColorValues.Green)
			return "#00ff00";
		if (value == HighlightColorValues.Cyan)
			return "#00ffff";
		if (value == HighlightColorValues.Magenta)
			return "#ff00ff";
		if (value == HighlightColorValues.Blue)
			return "#0000ff";
		if (value == HighlightColorValues.Red)
			return "#ff0000";
		if (value == HighlightColorValues.DarkBlue)
			return "#000080";
		if (value == HighlightColorValues.DarkCyan)
			return "#008080";
		if (value == HighlightColorValues.DarkGreen)
			return "#008000";
		if (value == HighlightColorValues.DarkMagenta)
			return "#800080";
		if (value == HighlightColorValues.DarkRed)
			return "#800000";
		if (value == HighlightColorValues.DarkYellow)
			return "#808000";
		if (value == HighlightColorValues.LightGray)
			return "#d3d3d3";
		if (value == HighlightColorValues.DarkGray)
			return "#a9a9a9";
		if (value == HighlightColorValues.Black)
			return "#000000";
		return null;
	}

	private static string ToCssColor(string color)
	{
		if (color.StartsWith("#", StringComparison.Ordinal))
			return color;
		if (color.Length is 6 or 3)
			return "#" + color;
		return color;
	}

	private string NormalizeMarker(string marker)
	{
		// Normalize common Word bullet glyphs to a standard round bullet
		// so markdown viewers render them consistently.
		if (marker.IndexOf("<span", StringComparison.OrdinalIgnoreCase) >= 0)
		{
			if (_config.PreserveListSymbols)
				return marker;

			string inner = StripTags(marker).Trim();
			var font = ExtractFontFamily(marker);
			var translatedSpan = TryTranslateSymbolFont(inner, font);
			if (!string.IsNullOrEmpty(translatedSpan))
				return translatedSpan!;
			return inner;
		}

		string trimmed = marker.Trim();
		if (trimmed.Length == 1)
		{
			char c = trimmed[0];
			if (c == '\u2022' || c == '•' || c == '·' || c == '')
				return _config.PreserveListSymbols ? "•" : "•";
		}

		// Attempt to translate symbol-font bullets when not preserving
		if (!_config.PreserveListSymbols)
		{
			var translated = TryTranslateSymbolFont(marker);
			if (!string.IsNullOrEmpty(translated))
				return translated!;
		}

		return trimmed;
	}

	private static bool LooksLikeOrderedListMarker(string marker)
	{
		// e.g., "1." or "12." or "1)" etc.
		return Regex.IsMatch(marker, @"^\d+[.)]$");
	}

	private static string EscapeOrderedListMarker(string marker)
	{
		// Escape the delimiter so Markdown doesn't start a list.
		if (marker.EndsWith("."))
			return marker.Insert(marker.Length - 1, "\\");
		if (marker.EndsWith(")"))
			return marker.Insert(marker.Length - 1, "\\");
		return marker;
	}

	private static string StripTags(string input)
	{
		return Regex.Replace(input, "<.*?>", string.Empty);
	}

	private static string? ExtractFontFamily(string marker)
	{
		var m = Regex.Match(marker, "font-family\\s*:\\s*([^;\">]+)", RegexOptions.IgnoreCase);
		if (!m.Success)
			return null;
		var font = m.Groups[1].Value.Trim();
		return font.Trim('"', '\'');
	}

	private static bool FontEquals(string? font, string target)
		=> !string.IsNullOrEmpty(font) && string.Equals(font, target, StringComparison.OrdinalIgnoreCase);

	private static string? TryTranslateSymbolFont(string marker, string? fontFamily = null)
	{
		// marker may be a single char in Symbol/ZapfDingbats/Webdings/Wingdings encoding. If not, return null.
		if (string.IsNullOrEmpty(marker) || marker.Length != 1)
			return null;

		var ch = marker[0];

		var converter = DxpFontSymbols.GetSymbolConverter(fontFamily);
		if (converter != null)
		{
			var translated = converter.Substitute(ch, null);
			if (!string.IsNullOrEmpty(translated) && !string.Equals(translated, marker, StringComparison.Ordinal))
				return translated;
		}

		return null;
	}

	public override IDisposable VisitDocumentBodyBegin(Body body, DxpIDocumentContext d)
	{
		if (SplitMode)
		{
			// Capture entire body into accept/reject buffers; emit at end.
			var prevAccept = _acceptWriter;
			var prevReject = _rejectWriter;
			var prevUnchanged = _unchangedWriter;

			var left = new StringWriter(new StringBuilder());
			var right = new StringWriter(new StringBuilder());

			_acceptWriter = right;   // accept path
			_rejectWriter = left;    // reject path
			_unchangedWriter = new MultiWriter(left, right);

			return Disposable.Create(() => {
				var leftText = left.ToString();
				var rightText = right.ToString();

				_acceptWriter = prevAccept;
				_rejectWriter = prevReject;
				_unchangedWriter = prevUnchanged;

				if (string.Equals(leftText, rightText, StringComparison.Ordinal))
				{
					_sinkWriter.Write(leftText);
				}
				else
				{
					EmitSplitRow(leftText, rightText);
				}
			});
		}

		return Disposable.Empty;
	}

	private double AdjustMarginLeft(double marginPt, DxpIDocumentContext d)
	{
		var marginLeftPoints = d.CurrentSection.Layout?.MarginLeft?.Inches is double inches
			? inches * 72.0
			: (double?)null;

		if (marginLeftPoints == null)
			return marginPt;
		var adjusted = marginPt - marginLeftPoints.Value;
		if (adjusted < 0)
			adjusted = 0;
		return adjusted;
	}

	private bool ShouldSuppressOutput(DxpIDocumentContext d)
	{
		if (_state.SuppressDepth > 0)
			return true;

		return !_config.EmitPageNumbers && d.CurrentFields.IsSuppressed;
	}

	private static bool LooksLikePageField(string? instr)
	{
		if (string.IsNullOrEmpty(instr))
			return false;
		return instr!.IndexOf("PAGE", StringComparison.OrdinalIgnoreCase) >= 0
			|| instr.IndexOf("NUMPAGES", StringComparison.OrdinalIgnoreCase) >= 0
			|| instr.IndexOf("SECTIONPAGES", StringComparison.OrdinalIgnoreCase) >= 0;
	}

	private static bool HasVisibleContent(string html)
	{
		if (string.IsNullOrWhiteSpace(html))
			return false;

		// Strip comments and empty paragraphs
		string cleaned = Regex.Replace(html, @"<!--.*?-->", string.Empty, RegexOptions.Singleline);
		cleaned = Regex.Replace(cleaned, @"<p[^>]*>\s*</p>", string.Empty, RegexOptions.IgnoreCase);
		cleaned = Regex.Replace(cleaned, @"&nbsp;", " ", RegexOptions.IgnoreCase);

		// Remove tags to inspect visible text
		string textOnly = Regex.Replace(cleaned, "<[^>]+>", string.Empty);
		textOnly = textOnly.Replace("\u00A0", " ").Trim();

		if (string.IsNullOrWhiteSpace(textOnly))
			return false;

		// Consider pure page-number content as non-visible for header/footer emission.
		if (Regex.IsMatch(textOnly, @"^[0-9]+$", RegexOptions.CultureInvariant))
			return false;

		return true;
	}

	IDisposable DxpIVisitor.VisitDrawingBegin(Drawing drw, DxpDrawingInfo? info, DxpIDocumentContext d)
	{
		VisitDrawingBegin(drw, info, d);
		return Disposable.Empty;
	}

	public new void VisitLegacyPictureBegin(Picture pict, DxpIDocumentContext d)
	{
		if (_config.EmitImages == false)
		{
			_writer.Write("[IMAGE]");
			return;
		}

		var alt = "image";
		_writer.Write($"[PICTURE: {alt}]");
	}

	IDisposable DxpIVisitor.VisitLegacyPictureBegin(Picture pict, DxpIDocumentContext d)
	{
		VisitLegacyPictureBegin(pict, d);
		return Disposable.Empty;
	}

	public override IDisposable VisitSectionBodyBegin(SectionProperties properties, DxpIDocumentContext d)
	{
		if (!_config.EmitDocumentColors)
			return Disposable.Empty;

		var style = new StringBuilder("flex:1 0 auto;");

		double? marginTopInches = d.CurrentSection.Layout?.MarginTop?.Inches;
		if (marginTopInches != null)
			style.Append("padding-top:").Append(marginTopInches.Value.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");

		_writer.Write($"""<div class="body" style="{style}">""" + "\n");

		return Disposable.Create(() => {
			_writer.WriteLine("</div>");
		});
	}

	public override IDisposable VisitSectionBegin(SectionProperties properties, SectionLayout layout, DxpIDocumentContext d)
	{

		if (!_config.EmitDocumentColors)
			return Disposable.Empty;

		if (_state.IsFirstSection)
		{
			_state.IsFirstSection = false;
		}
		else
		{
			_writer.WriteLine();
			_writer.WriteLine("<hr />");
			_writer.WriteLine();
		}

		var style = new StringBuilder("color:#000000;display:flex;flex-direction:column;position:relative;");
		double? pageWidthInches = d.CurrentSection.Layout?.PageWidth?.Inches;
		double? pageHeightInches = d.CurrentSection.Layout?.PageHeight?.Inches;
		if (pageWidthInches != null && pageHeightInches != null)
		{
			style.Append("width:").Append(pageWidthInches.Value.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");
			style.Append("min-height:").Append(pageHeightInches.Value.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");
		}
		double? marginLeftInches = d.CurrentSection.Layout?.MarginLeft?.Inches;
		double? marginRightInches = d.CurrentSection.Layout?.MarginRight?.Inches;
		double? marginTopInches = d.CurrentSection.Layout?.MarginTop?.Inches;
		if (marginLeftInches != null || marginRightInches != null || marginTopInches != null)
		{
			style.Append("box-sizing:border-box;");
			if (marginLeftInches != null)
				style.Append("padding-left:").Append(marginLeftInches.Value.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");
			if (marginRightInches != null)
				style.Append("padding-right:").Append(marginRightInches.Value.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");
		}


		string pageBackground = "#ffffff"; // default to white when no document background is set
		string? color = d.Background?.Color?.Value;
		if (!string.IsNullOrEmpty(color) && !string.Equals(color, "auto", StringComparison.OrdinalIgnoreCase))
		{
			pageBackground = ToCssColor(color!);
		}
		style.Append("background-color:").Append(pageBackground).Append(';');



		if (_config.EmitStyleFont)
		{
			if (!string.IsNullOrEmpty(d.DefaultRunStyle.FontName))
				style.Append("font-family:").Append(d.DefaultRunStyle.FontName).Append(';');
			if (d.DefaultRunStyle.FontSizeHalfPoints != null)
				style.Append("font-size:").Append((d.DefaultRunStyle.FontSizeHalfPoints.Value / 2.0).ToString("0.###", CultureInfo.InvariantCulture)).Append("pt;");
		}

		_writer.Write($"""<div class="section" style="{style}">""" + "\n");

		return Disposable.Create(() => {
			_writer.WriteLine("</div>");
		});
	}
}
