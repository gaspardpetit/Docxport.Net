using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocxportNet.API;
using Microsoft.Extensions.Logging;
using System.Globalization;
using System.Text.RegularExpressions;
using DocxportNet.Word;
using System.Text;
using DocxportNet.Core;

namespace DocxportNet.Visitors.Markdown;


/// <summary>
/// Mutable state specific to DxpMarkdownVisitor; separated for clarity of intent.
/// Currently empty—fields should be added here instead of on the visitor itself.
/// </summary>
internal sealed class DxpMarkdownVisitorState
{
	public enum InlineChangeMode
	{
		Unchanged,
		Inserted,
		Deleted
	}

	public bool InHeading { get; set; }
	public bool FontSpanOpen { get; set; }
	public bool AllCaps { get; set; }
	public bool IsFirstSection { get; set; } = true;
	public int TabIndex { get; set; }
	public int UnderlineDepth { get; set; }
	public bool InParagraph { get; set; }
	public double CurrentLineXPt { get; set; }
	public double CurrentFontSizePt { get; set; } = 12.0;
	public Stack<double> FontSizeStack { get; } = new();
	public DxpComputedTabStopKind? PendingAlignedTabKind { get; set; }
	public double PendingAlignedTabStopPosPt { get; set; }
	public double PendingAlignedTabStartXPt { get; set; }
	public double PendingAlignedTabSegmentWidthPt { get; set; }
	public bool PendingAlignedTabUnderline { get; set; }
	public StringBuilder PendingAlignedTabBuffer { get; } = new();
	public Stack<StringBuilder> DeletedCaptures { get; } = new();
	public InlineChangeMode CurrentInlineMode { get; set; } = InlineChangeMode.Unchanged;

	// the writer used to print deleted content
	public TextWriter DeletedTextWriter = TextWriter.Null;
	// the writer used to print inserted content
	public TextWriter InsertedTextWriter = TextWriter.Null;
	// the writer used to print unmodified content
	public TextWriter UnchangedTextWriter = TextWriter.Null;

	public bool SawTrackedChange { get; set; }
}

public enum DxpTrackedChangeMode
{
	AcceptChanges,
	RejectChanges,
	InlineChanges,
	SplitChanges
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
	public bool EmitFieldInstructions = true;
	public bool UsePlainComments = false;
	public bool EmitCustomProperties = true;
	public bool EmitTimeline = false;
	public DxpTrackedChangeMode TrackedChangeMode = DxpTrackedChangeMode.InlineChanges;

	public static DxpMarkdownVisitorConfig CreateRichConfig() => new();
	public static DxpMarkdownVisitorConfig CreatePlainConfig() => new() {
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

	public static DxpMarkdownVisitorConfig CreateConfig() => CreateRichConfig();
}


public partial class DxpMarkdownVisitor : DxpVisitor, DxpITextVisitor, IDisposable
{
	private TextWriter _sinkWriter;
	private StreamWriter? _ownedStreamWriter;
	private DxpBufferedTextWriter _rejectBufferedWriter;
	private DxpBufferedTextWriter _acceptBufferedWriter;

	private DxpMarkdownVisitorConfig _config;
	private DxpMarkdownVisitorState _state = new();

	public DxpMarkdownVisitor(TextWriter writer, DxpMarkdownVisitorConfig config, ILogger? logger)
		: base(logger)
	{
		_config = config;
		_sinkWriter = writer;
		_rejectBufferedWriter = new DxpBufferedTextWriter();
		_acceptBufferedWriter = new DxpBufferedTextWriter();
		ConfigureWriters();
	}

	public DxpMarkdownVisitor(DxpMarkdownVisitorConfig config, ILogger? logger = null)
		: this(TextWriter.Null, config, logger)
	{
	}

	public void SetOutput(TextWriter writer)
	{
		ReleaseOwnedWriter();
		_sinkWriter = writer ?? throw new ArgumentNullException(nameof(writer));
		_rejectBufferedWriter = new DxpBufferedTextWriter();
		_acceptBufferedWriter = new DxpBufferedTextWriter();
		_state = new DxpMarkdownVisitorState();
		ConfigureWriters();
	}

	public override void SetOutput(Stream stream)
	{
		ReleaseOwnedWriter();
		_ownedStreamWriter = new StreamWriter(stream, Encoding.UTF8, bufferSize: 1024, leaveOpen: true)
		{
			AutoFlush = true
		};
		var writer = _ownedStreamWriter;
		SetOutput(writer);
	}

	private void ReleaseOwnedWriter()
	{
		if (_ownedStreamWriter == null)
			return;

		_ownedStreamWriter.Flush();
		_ownedStreamWriter.Dispose();
		_ownedStreamWriter = null;
	}

	public void Dispose() => ReleaseOwnedWriter();

	private void ConfigureWriters()
	{
		switch (_config.TrackedChangeMode)
		{
			case DxpTrackedChangeMode.InlineChanges:
				_state.DeletedTextWriter = _sinkWriter;
				_state.InsertedTextWriter = _sinkWriter;
				_state.UnchangedTextWriter = _sinkWriter;
				break;
			case DxpTrackedChangeMode.SplitChanges:
				_state.DeletedTextWriter = _rejectBufferedWriter;
				_state.InsertedTextWriter = _acceptBufferedWriter;
				_state.UnchangedTextWriter = new DxpMultiTextWriter(true, _rejectBufferedWriter, _acceptBufferedWriter);
				break;
			case DxpTrackedChangeMode.AcceptChanges:
				_state.DeletedTextWriter = TextWriter.Null;
				_state.InsertedTextWriter = _sinkWriter;
				_state.UnchangedTextWriter = _sinkWriter;
				break;
			case DxpTrackedChangeMode.RejectChanges:
				_state.DeletedTextWriter = _sinkWriter;
				_state.InsertedTextWriter = TextWriter.Null;
				_state.UnchangedTextWriter = _sinkWriter;
				break;
		}
	}

	public override void VisitText(Text t, DxpIDocumentContext d)
	{
		string text = t.Text;
		if (_state.AllCaps)
		{
			var culture = CultureInfo.InvariantCulture;
			text = text.ToUpper(culture);
		}

		if (_state.PendingAlignedTabKind != null)
			_state.PendingAlignedTabSegmentWidthPt += EstimateTextWidthPt(text, _state.CurrentFontSizePt);
		else if (_state.InParagraph)
			_state.CurrentLineXPt += EstimateTextWidthPt(text, _state.CurrentFontSizePt);

		Write(d, text);
	}

	private static double EstimateTextWidthPt(string text, double fontSizePt)
	{
		if (string.IsNullOrEmpty(text))
			return 0.0;

		double width = 0.0;
		foreach (var ch in text)
		{
			if (ch == '\n' || ch == '\r')
				continue;

			if (ch == ' ' || ch == '\u00A0')
			{
				width += fontSizePt * 0.33;
				continue;
			}

			if ("ilI.,:;|!".IndexOf(ch) >= 0)
			{
				width += fontSizePt * 0.25;
				continue;
			}

			width += fontSizePt * 0.5;
		}

		return width;
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
		Write(d, "<b>");
	}
	public override void StyleBoldEnd(DxpIDocumentContext d)
	{
		if (_state.InHeading)
			return;
		Write(d, "</b>");
	}

	public override void StyleItalicBegin(DxpIDocumentContext d)
	{
		Write(d, "<i>");
	}
	public override void StyleItalicEnd(DxpIDocumentContext d)
	{
		Write(d, "</i>");
	}

	public override void StyleUnderlineBegin(DxpIDocumentContext d)
	{
		_state.UnderlineDepth++;
		Write(d, "<u>");
	}
	public override void StyleUnderlineEnd(DxpIDocumentContext d)
	{
		if (_state.UnderlineDepth > 0)
			_state.UnderlineDepth--;
		Write(d, "</u>");
	}

	public override void StyleStrikeBegin(DxpIDocumentContext d)
	{
		Write(d, "<del>");
	}
	public override void StyleStrikeEnd(DxpIDocumentContext d)
	{
		Write(d, "</del>");
	}

	public override void StyleDoubleStrikeBegin(DxpIDocumentContext d)
	{
		Write(d, "<del>");
	}
	public override void StyleDoubleStrikeEnd(DxpIDocumentContext d)
	{
		Write(d, "</del>");
	}

	public override void StyleSuperscriptBegin(DxpIDocumentContext d)
	{
		Write(d, "<sup>");
	}
	public override void StyleSuperscriptEnd(DxpIDocumentContext d)
	{
		Write(d, "</sup>");
	}

	public override void StyleSubscriptBegin(DxpIDocumentContext d)
	{
		Write(d, "<sub>");
	}
	public override void StyleSubscriptEnd(DxpIDocumentContext d)
	{
		Write(d, "</sub>");
	}

	public override void StyleFontBegin(DxpFont font, DxpIDocumentContext d)
	{
		_state.FontSizeStack.Push(_state.CurrentFontSizePt);
		if (font.fontSizeHalfPoints != null)
			_state.CurrentFontSizePt = font.fontSizeHalfPoints.Value / 2.0;

		if (_config.EmitStyleFont == false)
			return;

		if (IsDefaultFont(font.fontName, font.fontSizeHalfPoints, d))
		{
			_state.FontSpanOpen = false;
			return;
		}

		var style = new StringBuilder();
		if (!string.IsNullOrWhiteSpace(font.fontName))
			style.Append("font-family: ").Append(font.fontName).Append(';');
		if (font.fontSizeHalfPoints != null)
			style.Append(" font-size: ").Append(font.fontSizeHalfPoints.Value / 2.0).Append("pt;");

		if (style.Length == 0)
		{
			_state.FontSpanOpen = false;
			return;
		}

		_state.FontSpanOpen = true;
		Write(d, $"""<span style="{style}">""");
	}

	public override void StyleFontEnd(DxpIDocumentContext d)
	{
		if (_state.FontSizeStack.Count > 0)
			_state.CurrentFontSizePt = _state.FontSizeStack.Pop();

		if (_config.EmitStyleFont == false)
			return;

		if (_state.FontSpanOpen)
		{
			Write(d, "</span>");
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
			Write(d, $"<a id=\"{Escape(name!)}\" data-bookmark-id=\"{id}\"></a>");
	}

	public override void VisitBookmarkEnd(BookmarkEnd be, DxpIDocumentContext d)
	{
		// Usually nothing to emit; it just closes the range.
	}

	public override IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target, DxpIDocumentContext d)
	{
		string? href = target?.uri;
		Write(d, href != null ? $"<a href=\"{HtmlAttr(href)}\">" : "<a>");
		return DxpDisposable.Create(() => Write(d, "</a>"));
	}

	public string Escape(string name)
	{
		return name;
	}

	public override IDisposable VisitInsertedBegin(Inserted ins, DxpIDocumentContext d)
	{
		return DxpDisposable.Empty;
	}

	public override IDisposable VisitDeletedBegin(Deleted del, DxpIDocumentContext d)
	{
		return DxpDisposable.Empty;
	}

	public override IDisposable VisitDeletedRunBegin(DeletedRun dr, DxpIDocumentContext d)
	{
		return DxpDisposable.Empty;
	}

	public override void VisitDeletedParagraphMark(Deleted del, ParagraphProperties pPr, Paragraph? p, DxpIDocumentContext d)
	{
		// Paragraph mark deletions carry no inline content; outer VisitDeletedBegin wrapper is enough.
	}

	public override IDisposable VisitInsertedRunBegin(InsertedRun ir, DxpIDocumentContext d)
	{
		return DxpDisposable.Empty;
	}

	public override void VisitDeletedText(DeletedText dt, DxpIDocumentContext d)
	{
		Write(d, dt.Text);
	}

	public override void VisitNoBreakHyphen(NoBreakHyphen h, DxpIDocumentContext d)
	{
		Write(d, "-");
	}

	public override void VisitDrawingBegin(Drawing drw, DxpDrawingInfo? info, DxpIDocumentContext d)
	{
		if (_config.EmitImages == false)
		{
			Write(d, "[IMAGE]");
			return;
		}

		var alt = HtmlAttr(info?.AltText ?? "image");
		var dataUri = info?.DataUri;
		var contentType = info?.ContentType ?? "";

		if (!string.IsNullOrEmpty(dataUri) && contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
		{
			Write(d, $"<img src=\"{dataUri}\" alt=\"{alt}\" />");
		}
		else if (!string.IsNullOrEmpty(dataUri))
		{
			Write(d, $"<object data=\"{dataUri}\" type=\"{HtmlAttr(contentType)}\">[DRAWING: {alt}]</object>");
		}
		else
		{
			var meta = string.IsNullOrEmpty(contentType) ? "" : $" ({contentType})";
			Write(d, $"[DRAWING: {alt}{meta}]");
		}
	}

	public override IDisposable VisitDocumentBegin(WordprocessingDocument doc, DxpIDocumentContext d)
	{
		var lines = new List<string>();

		void Add(string label, string? value)
		{
			if (!string.IsNullOrWhiteSpace(value))
				lines.Add($"<!-- {label}: {value} -->");
		}

		IPackageProperties? core = d.DocumentProperties.PackageProperties;
		if (core != null)
		{
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
		}

		IReadOnlyList<DxpTimelineEvent>? timeline = d.DocumentProperties.TimelineEvents;
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

		IReadOnlyList<CustomFileProperty>? custom = d.DocumentProperties.CustomFileProperties;
		if (custom != null && _config.EmitCustomProperties)
		{
			foreach (var prop in custom)
			{
				if (prop.Value != null)
					lines.Add($"<!-- {prop.Name}: {prop.Value} -->");
			}
		}

		foreach (string line in lines)
			WriteLine(d, line);

		if (lines.Count > 0)
			WriteLine(d);

		if (_config.EmitTimeline && _config.RichTables && timeline != null && timeline.Count > 0)
		{
			WriteLine(d);
			WriteLine(d, "| Date | Event |");
			WriteLine(d, "| --- | --- |");
			foreach (var ev in timeline)
			{
				var date = ev.DateUtc?.ToString("yyyy-MM-dd HH:mm:ss 'UTC'") ?? "unknown";
				var who = string.IsNullOrWhiteSpace(ev.Author) ? "unknown" : ev.Author;
				var detail = string.IsNullOrWhiteSpace(ev.Detail) ? "" : $" ({ev.Detail})";
				WriteLine(d, $"| {date} | {ev.Kind} by {who}{detail} |");
			}
			WriteLine(d);
		}

		return DxpDisposable.Empty;
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
		_state.TabIndex = 0;
		_state.CurrentLineXPt = 0.0;
		_state.InParagraph = true;
		_state.PendingAlignedTabKind = null;
		_state.PendingAlignedTabBuffer.Clear();
		_state.PendingAlignedTabSegmentWidthPt = 0.0;
		_state.PendingAlignedTabUnderline = false;

		if (_config.TrackedChangeMode == DxpTrackedChangeMode.SplitChanges)
		{
			EmitSplitBuffersIfNeeded();
			_rejectBufferedWriter.Clear();
			_acceptBufferedWriter.Clear();
		}

		var marker =
			d.KeepAccept
			? paragraph.MarkerAccept
			: paragraph.MarkerReject;
		var indent = paragraph.Indent;

		string innerText = p.InnerText;

		if (string.IsNullOrWhiteSpace(innerText))
		{
			return DxpDisposable.Create(() => {
				WriteLine(d);
			});
		}


		var styleChain = d.Styles.GetParagraphStyleChain(p);
		string? justify = null;
		if (_config.EmitParagraphAlignment && paragraph.ComputedStyle.TextAlign != null)
		{
			justify = paragraph.ComputedStyle.TextAlign.Value switch
			{
				DxpComputedTextAlign.Center => "center",
				DxpComputedTextAlign.Right => "right",
				DxpComputedTextAlign.Justify => "justify",
				_ => null
			};
		}

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
			return DxpDisposable.Empty;
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
			headingStyle = $" style=\"text-align:{justify};\"";

		if (isHeading)
		{
			if (_config.EmitStyleFont)
			{
				var styleAttr = headingStyle ?? string.Empty;
				Write(d, $"<h{headingLevel}{styleAttr}>");
			}
			else
			{
				for (int i = 0; i < headingLevel; ++i)
					Write(d, $"#");
				Write(d, $" ");
			}
		}

		var paraCss = paragraph.ComputedStyle.ToCss(includeTextAlign: _config.EmitParagraphAlignment);
		bool needsParagraphWrapper = !isHeading && !string.IsNullOrEmpty(paraCss);
		if (needsParagraphWrapper)
		{
			Write(d, $"<p style=\"{paraCss}\">");
		}

		if (isHeading)
			marker = null;

		if (isBlockQuote)
			marker = null;

		if (isCode)
			marker = null;

		if (marker?.marker != null)
		{
			var normalizedMarker = NormalizeMarker(marker.marker);
			if (!needsParagraphWrapper && LooksLikeOrderedListMarker(normalizedMarker))
				normalizedMarker = EscapeOrderedListMarker(normalizedMarker);
			Write(d, $"""{normalizedMarker} """);
		}

		bool previousHeading = _state.InHeading;
		if (isHeading)
			_state.InHeading = true;

		if (isBlockQuote)
		{
			if (_config.EmitStyleFont)
				Write(d, "<blockquote>");
			else
				Write(d, "> ");
		}

		if (isCode)
		{
			if (_config.UsePlainCodeBlocks)
			{
				Write(d, "```\n");
			}
			else
			{
				Write(d, "<pre><code>");
			}
		}

		if (isCaption && !_config.UsePlainCodeBlocks && !isHeading && !isBlockQuote)
		{
			Write(d, "<figcaption>");
		}

		var baseDispose = DxpDisposable.Create(() => {
			FlushPendingAlignedTab(d);

			if (_config.TrackedChangeMode == DxpTrackedChangeMode.InlineChanges)
				SetInlineChangeMode(DxpMarkdownVisitorState.InlineChangeMode.Unchanged);

			if (isCaption && !_config.UsePlainCodeBlocks && !isHeading && !isBlockQuote)
			{
				Write(d, "</figcaption>");
			}
			_state.InHeading = previousHeading;
			if (isCode)
			{
				if (_config.UsePlainCodeBlocks)
					WriteLine(d, "\n```");
				else
					Write(d, "</code></pre>");
			}
			if (isHeading && _config.EmitStyleFont)
			{
				Write(d, $"</h{headingLevel}>");
			}
			if (needsParagraphWrapper)
				WriteLine(d, "</p>");
			else if (isBlockQuote && _config.EmitStyleFont)
				WriteLine(d, "</blockquote>");
			int newlines = 2;
			for (int i = 0; i < newlines; i++)
				WriteLine(d);

			if (_config.TrackedChangeMode == DxpTrackedChangeMode.SplitChanges)
				EmitSplitBuffersIfNeeded();
			_state.InParagraph = false;
		});

		return DxpDisposable.Create(() => {
			baseDispose.Dispose();
		});

	}

	public override void VisitFootnoteReference(FootnoteReference fr, DxpIFootnoteContext footnote, DxpIDocumentContext d)
	{
		Write(d, $"<a href=\"#fn-{footnote.Id}\" id=\"fnref-{footnote.Id}\">[{footnote.Index}]</a>");
	}

	public override IDisposable VisitSectionHeaderBegin(Header hdr, object kind, DxpIDocumentContext d)
	{
		if (_config.EmitSectionHeadersFooters == false)
			return DxpDisposable.Empty;

		WriteLine(d, """<div class="header" style="border-bottom:1px solid #000;">""");

		return DxpDisposable.Create(() => {
			WriteLine(d, "</div>");
		});
	}

	public override IDisposable VisitSectionFooterBegin(Footer ftr, object kind, DxpIDocumentContext d)
	{
		WriteLine(d, """<div class="footer" style="border-top:1px solid #000;">""");

		return DxpDisposable.Create(() => {
				WriteLine(d, "</div>");
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
		EmitFieldInstruction(d, text);
	}

	public override IDisposable VisitComplexFieldResultBegin(DxpIDocumentContext d)
	{
		return DxpDisposable.Empty;
	}

	public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
	{
		if (!_config.EmitPageNumbers)
		{
			var instr = d.CurrentFields.Current?.InstructionText;
			if (LooksLikePageField(instr))
				return;
		}
		Write(d, text);
	}

	public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
	{
	}

	public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
	{
		var instr = fld.Instruction?.Value;
		if (instr != null)
			EmitFieldInstruction(d, instr);
		return DxpDisposable.Empty;
	}

	public override IDisposable VisitFootnoteBegin(Footnote fn, DxpIFootnoteContext footnote, DxpIDocumentContext d)
	{
		WriteLine(d, $"""<div class="footnote" id="fn-{footnote.Id}">""");
		return DxpDisposable.Create(() => WriteLine(d, "</div>"));
	}

	public override void VisitFootnoteReferenceMark(FootnoteReferenceMark m, DxpIFootnoteContext footnote, DxpIDocumentContext d)
	{
		if (footnote.Index != null)
			Write(d, $"{footnote.Index}");
	}


	public override IDisposable VisitTableBegin(Table t, DxpTableModel model, DxpIDocumentContext d, DxpITableContext table)
	{
		var currentStyle = _config.EmitTableBorders ? table.ComputedStyle.ToCss() : null;

		Write(d, "<table");
		if (!string.IsNullOrEmpty(currentStyle))
			Write(d, $" style=\"{currentStyle}\"");
		WriteLine(d, ">");
		return DxpDisposable.Create(() => {
			WriteLine(d, "</table>");
		});
	}

	public override IDisposable VisitTableRowBegin(TableRow tr, DxpITableRowContext row, DxpIDocumentContext d)
	{
		var isHeader = row.IsHeader;

		if (isHeader)
			WriteLine(d, "  <tr class=\"header-row\">");
		else
			WriteLine(d, "  <tr>");
		return DxpDisposable.Create(() => WriteLine(d, "  </tr>"));
	}

	public override IDisposable VisitTableCellBegin(TableCell tc, DxpITableCellContext cell, DxpIDocumentContext d)
	{
		var spans = (cell.RowSpan, cell.ColSpan);
		var cellStyle = _config.EmitTableBorders ? cell.ComputedStyle.ToCss() : null;

		Write(d, "    <td");
		if (spans.Item1 > 1)
			Write(d, $" rowspan=\"{spans.Item1}\"");
		if (spans.Item2 > 1)
			Write(d, $" colspan=\"{spans.Item2}\"");
		if (!string.IsNullOrEmpty(cellStyle))
			Write(d, $" style=\"{cellStyle}\"");
		Write(d, ">");

		return DxpDisposable.Create(() => {
			WriteLine(d, "</td>");
		});
	}

	public override IDisposable VisitBlockBegin(OpenXmlElement child, DxpIDocumentContext d)
	{
		// optional: do nothing
		return DxpDisposable.Empty;
	}

	public override IDisposable VisitCommentBegin(DxpCommentInfo c, DxpCommentThread thread, DxpIDocumentContext d)
	{
		if (_config.UsePlainComments)
		{
			var label = c.IsReply ? "REPLY BY" : "COMMENT BY";
			var who = !string.IsNullOrEmpty(c.Author)
				? c.Author!
				: !string.IsNullOrEmpty(c.Initials) ? c.Initials! : "Unknown";
			var when = c.DateUtc?.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss'Z'");

			if (!string.IsNullOrEmpty(when))
				WriteLine(d, $"  {label} {who} ON {when}");
			else
				WriteLine(d, $"  {label} {who}");

			WriteLine(d);

			return DxpDisposable.Create(() => {
				WriteLine(d);
			});
		}
		else
		{
			var commentStyle = "background:#fffcf0;border:1px solid #d9b200;border-radius:4px;padding:6px;margin-bottom:6px;";

			var label = BuildCommentLabel(c);
			if (!string.IsNullOrEmpty(label))
				WriteLine(d, "  " + label);

			Write(d, $"""  <div class="comment" style="{commentStyle}">""");
			WriteLine(d);

			return DxpDisposable.Create(() => {
				WriteLine(d, "  </div>");
				WriteLine(d);
			});
		}
	}

	public override IDisposable VisitCommentThreadBegin(string anchorId, DxpCommentThread thread, DxpIDocumentContext d)
	{
		if (thread.Comments == null || thread.Comments.Count == 0)
			return DxpDisposable.Empty;

		if (_config.UsePlainComments)
		{
			return EmitPlainCommentThread(thread, d);
		}

		var commentsStyle = "background:#fff8c6;border:1px solid #e6c44a;border-radius:6px;padding:8px;margin:8px 0 8px 12px;float:right;max-width:45%;";

		Write(d, $"""<div class="comments" style="{commentsStyle}">""");

		return DxpDisposable.Create(() => {
			WriteLine(d, "</div>");
			WriteLine(d);
		});
	}

	private IDisposable EmitPlainCommentThread(DxpCommentThread thread, DxpIDocumentContext d)
	{
		WriteLine(d, "<!--");
		WriteLine(d);

		return DxpDisposable.Create(() => {
			WriteLine(d, "-->");
			WriteLine(d);
		});
	}

	private string BuildCommentLabel(DxpCommentInfo c)
	{
		var who = !string.IsNullOrEmpty(c.Author)
			? c.Author!
			: !string.IsNullOrEmpty(c.Initials) ? c.Initials! : "Unknown";
		var when = c.DateUtc?.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss'Z'") ?? string.Empty;

		if (string.IsNullOrEmpty(who) && string.IsNullOrEmpty(when))
			return string.Empty;

		return $"<span style=\"font-size:small\">{HtmlAttr(who)} | {HtmlAttr(when)}</span>";
	}

	public override void VisitBreak(Break br, DxpIDocumentContext d)
	{
		FlushPendingAlignedTab(d);
		if (_state.InParagraph)
		{
			_state.CurrentLineXPt = 0.0;
			_state.TabIndex = 0;
		}
		Write(d, "<br/>");
	}

	public override void VisitCarriageReturn(CarriageReturn cr, DxpIDocumentContext d)
	{
		FlushPendingAlignedTab(d);
		if (_state.InParagraph)
		{
			_state.CurrentLineXPt = 0.0;
			_state.TabIndex = 0;
		}
		Write(d, "<br/>");
	}

	public override void VisitTab(TabChar tab, DxpIDocumentContext d)
	{
		var stops = d.CurrentParagraph.Layout?.TabStops;
		if (stops == null || stops.Count == 0)
		{
			Write(d, "&#9;"); // best-effort fallback
			return;
		}

		FlushPendingAlignedTab(d);

		var index = _state.TabIndex++;
		var stop = index < stops.Count ? stops[index] : null;
		var kind = stop?.Kind ?? DxpComputedTabStopKind.Left;
		double stopPos = stop?.PositionPt ?? (_state.CurrentLineXPt + 36.0);

		if (kind == DxpComputedTabStopKind.Right || kind == DxpComputedTabStopKind.Center)
		{
			_state.PendingAlignedTabKind = kind;
			_state.PendingAlignedTabStopPosPt = stopPos;
			_state.PendingAlignedTabStartXPt = _state.CurrentLineXPt;
			_state.PendingAlignedTabSegmentWidthPt = 0.0;
			_state.PendingAlignedTabUnderline = _state.UnderlineDepth > 0;
			_state.PendingAlignedTabBuffer.Clear();
			return;
		}

		double width = stopPos - _state.CurrentLineXPt;
		if (width < 0) width = 0;
		_state.CurrentLineXPt = stopPos;

		var cls = kind switch
		{
			DxpComputedTabStopKind.Right => "dxp-tab dxp-tab-right",
			DxpComputedTabStopKind.Center => "dxp-tab dxp-tab-center",
			DxpComputedTabStopKind.Decimal => "dxp-tab dxp-tab-decimal",
			_ => "dxp-tab"
		};

		var extraCss = _state.UnderlineDepth > 0 ? "height:1em;border-bottom:1px solid currentColor;" : "";
		Write(d, $"<span class=\"{cls}\" style=\"display:inline-block;width:{width.ToString("0.###", CultureInfo.InvariantCulture)}pt;{extraCss}\"></span>");
	}

	private void FlushPendingAlignedTab(DxpIDocumentContext d)
	{
		if (_state.PendingAlignedTabKind == null)
			return;

		var kind = _state.PendingAlignedTabKind.Value;
		double stopPos = _state.PendingAlignedTabStopPosPt;
		double segmentWidth = _state.PendingAlignedTabSegmentWidthPt;
		double startX = _state.PendingAlignedTabStartXPt;

		double alignedStart = kind switch
		{
			DxpComputedTabStopKind.Right => stopPos - segmentWidth,
			DxpComputedTabStopKind.Center => stopPos - (segmentWidth / 2.0),
			_ => stopPos
		};

		if (alignedStart < startX)
			alignedStart = startX;

		double spacerWidth = alignedStart - startX;
		if (spacerWidth < 0) spacerWidth = 0;

		var cls = kind switch
		{
			DxpComputedTabStopKind.Right => "dxp-tab dxp-tab-right",
			DxpComputedTabStopKind.Center => "dxp-tab dxp-tab-center",
			_ => "dxp-tab"
		};

		var extraCss = _state.PendingAlignedTabUnderline ? "height:1em;border-bottom:1px solid currentColor;" : "";

		_state.PendingAlignedTabKind = null;
		_state.PendingAlignedTabUnderline = false;
		var buffered = _state.PendingAlignedTabBuffer.ToString();
		_state.PendingAlignedTabBuffer.Clear();
		_state.PendingAlignedTabSegmentWidthPt = 0.0;

		Write(d, $"<span class=\"{cls}\" style=\"display:inline-block;width:{spacerWidth.ToString("0.###", CultureInfo.InvariantCulture)}pt;{extraCss}\"></span>");
		if (buffered.Length > 0)
			Write(d, buffered);

		_state.CurrentLineXPt = alignedStart + segmentWidth;
	}

	public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
	{
		var style = BuildRunStyle(r.RunProperties);
		bool hasText = r.ChildElements.OfType<Text>().Any(t => !string.IsNullOrEmpty(t.Text));

		if (string.IsNullOrEmpty(style) || !hasText)
			return DxpDisposable.Empty;

		Write(d, $"<span style=\"{style}\">");
		return DxpDisposable.Create(() => Write(d, "</span>"));
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
		return DxpDisposable.Empty;
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

	private static bool LooksLikePageField(string? instr)
	{
		if (string.IsNullOrEmpty(instr))
			return false;
		return instr!.IndexOf("PAGE", StringComparison.OrdinalIgnoreCase) >= 0
			|| instr.IndexOf("NUMPAGES", StringComparison.OrdinalIgnoreCase) >= 0
			|| instr.IndexOf("SECTIONPAGES", StringComparison.OrdinalIgnoreCase) >= 0;
	}

	IDisposable DxpIVisitor.VisitDrawingBegin(Drawing drw, DxpDrawingInfo? info, DxpIDocumentContext d)
	{
		VisitDrawingBegin(drw, info, d);
		return DxpDisposable.Empty;
	}

	public new void VisitLegacyPictureBegin(Picture pict, DxpIDocumentContext d)
	{
		if (_config.EmitImages == false)
		{
			Write(d, "[IMAGE]");
			return;
		}

		var alt = "image";
		Write(d, $"[PICTURE: {alt}]");
	}

	IDisposable DxpIVisitor.VisitLegacyPictureBegin(Picture pict, DxpIDocumentContext d)
	{
		VisitLegacyPictureBegin(pict, d);
		return DxpDisposable.Empty;
	}

	public override IDisposable VisitSectionBodyBegin(SectionProperties properties, DxpIDocumentContext d)
	{
		if (!_config.EmitDocumentColors)
			return DxpDisposable.Empty;

		var style = new StringBuilder("flex:1 0 auto;");

		double? marginTopInches = d.CurrentSection.Layout?.MarginTop?.Inches;
		if (marginTopInches != null)
			style.Append("padding-top:").Append(marginTopInches.Value.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");

		Write(d, $"""<div class="body" style="{style}">""" + "\n");

		return DxpDisposable.Create(() => {
			WriteLine(d, "</div>");
		});
	}

	public override IDisposable VisitSectionBegin(SectionProperties properties, SectionLayout layout, DxpIDocumentContext d)
	{

		if (!_config.EmitDocumentColors)
			return DxpDisposable.Empty;

		if (_state.IsFirstSection)
		{
			_state.IsFirstSection = false;
		}
		else
		{
			WriteLine(d);
			WriteLine(d, "<hr />");
			WriteLine(d);
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

		Write(d, $"""<div class="section" style="{style}">""" + "\n");

		return DxpDisposable.Create(() => {
			WriteLine(d, "</div>");
		});
	}

	private void WriteLine(DxpIDocumentContext d)
	{
		Write(d, "\n");
	}

	private void WriteLine(DxpIDocumentContext d, string str)
	{
		Write(d, $"{str}\n");
	}

	private void EmitFieldInstruction(DxpIDocumentContext d, string instruction)
	{
		if (!_config.EmitFieldInstructions)
			return;

		var trimmed = instruction.Trim();
		if (trimmed.Length == 0)
			return;

		bool useBlock = trimmed.Contains('\n') || trimmed.Length > 80;
		if (useBlock)
		{
			Write(d, "\n\n```\n");
			Write(d, trimmed);
			if (!trimmed.EndsWith("\n", StringComparison.Ordinal))
				Write(d, "\n");
			Write(d, "```\n\n");
		}
		else
		{
			var escaped = trimmed.Replace("`", "``");
			Write(d, "\n\n`");
			Write(d, escaped);
			Write(d, "`\n\n");
		}
	}

	private void Write(DxpIDocumentContext d, string str)
	{
		if (_config.TrackedChangeMode == DxpTrackedChangeMode.InlineChanges)
		{
			var targetMode = DetermineChangeMode(d);
			SetInlineChangeMode(targetMode);
			WriteRouted(targetMode, str);
			return;
		}

		WriteRouted(DetermineChangeMode(d), str);
	}

	private void WriteRouted(DxpMarkdownVisitorState.InlineChangeMode mode, string str)
	{
		if (_state.PendingAlignedTabKind != null)
		{
			_state.PendingAlignedTabBuffer.Append(str);
			return;
		}

		if (mode == DxpMarkdownVisitorState.InlineChangeMode.Inserted)
			_state.InsertedTextWriter.Write(str);
		else if (mode == DxpMarkdownVisitorState.InlineChangeMode.Deleted)
			_state.DeletedTextWriter.Write(str);
		else
			_state.UnchangedTextWriter.Write(str);
	}

	private DxpMarkdownVisitorState.InlineChangeMode DetermineChangeMode(DxpIDocumentContext d)
	{
		if (d.KeepAccept && d.KeepReject)
			return DxpMarkdownVisitorState.InlineChangeMode.Unchanged;
		if (d.KeepAccept && !d.KeepReject)
			return DxpMarkdownVisitorState.InlineChangeMode.Inserted;
		if (!d.KeepAccept && d.KeepReject)
			return DxpMarkdownVisitorState.InlineChangeMode.Deleted;
		return DxpMarkdownVisitorState.InlineChangeMode.Unchanged;
	}

	private void SetInlineChangeMode(DxpMarkdownVisitorState.InlineChangeMode mode)
	{
		if (mode == _state.CurrentInlineMode)
			return;

		if (_state.CurrentInlineMode == DxpMarkdownVisitorState.InlineChangeMode.Inserted)
			WriteRouted(DxpMarkdownVisitorState.InlineChangeMode.Unchanged, "</u>");
		else if (_state.CurrentInlineMode == DxpMarkdownVisitorState.InlineChangeMode.Deleted)
			WriteRouted(DxpMarkdownVisitorState.InlineChangeMode.Unchanged, "</del>");

		if (mode == DxpMarkdownVisitorState.InlineChangeMode.Inserted)
			WriteRouted(DxpMarkdownVisitorState.InlineChangeMode.Unchanged, InlineInsertedTag());
		else if (mode == DxpMarkdownVisitorState.InlineChangeMode.Deleted)
			WriteRouted(DxpMarkdownVisitorState.InlineChangeMode.Unchanged, InlineDeletedTag());

		_state.CurrentInlineMode = mode;
	}

	private void EmitSplitBuffersIfNeeded()
	{
		var rejected = _rejectBufferedWriter.Drain();
		var accepted = _acceptBufferedWriter.Drain();

		if (string.IsNullOrEmpty(rejected) && string.IsNullOrEmpty(accepted))
			return;

		if (string.Equals(rejected, accepted, StringComparison.Ordinal))
		{
			_sinkWriter.Write(rejected);
		}
		else
		{
			const string border = "1px solid #ccc";
			_sinkWriter.Write($"<table style=\"width:100%;border-collapse:collapse;border:{border};\">");
			_sinkWriter.Write("<tr>");
			_sinkWriter.Write($"<td style=\"width:50%;vertical-align:top;padding:8px;border:{border};\">");
			_sinkWriter.Write(rejected);
			_sinkWriter.Write("</td>");
			_sinkWriter.Write($"<td style=\"width:50%;vertical-align:top;padding:8px;border:{border};\">");
			_sinkWriter.Write(accepted);
			_sinkWriter.Write("</td>");
			_sinkWriter.Write("</tr></table>");
		}
	}

	private string InlineInsertedTag()
	{
		if (_config.EmitRunColor)
			return "<u style=\"color:blue;\">";
		return "<u>";
	}

	private string InlineDeletedTag()
	{
		if (_config.EmitRunColor)
			return "<del style=\"color:red;\">";
		return "<del>";
	}
}
