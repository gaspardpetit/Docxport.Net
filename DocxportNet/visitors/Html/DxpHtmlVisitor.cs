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
using System.Net;
using DocxportNet.Visitors.Markdown;

namespace DocxportNet.Visitors.Html;

internal sealed class DxpHtmlVisitorState
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
	public InlineChangeMode CurrentInlineMode { get; set; } = InlineChangeMode.Unchanged;

	public TextWriter DeletedTextWriter = TextWriter.Null;
	public TextWriter InsertedTextWriter = TextWriter.Null;
	public TextWriter UnchangedTextWriter = TextWriter.Null;
}

public sealed record DxpHtmlVisitorConfig
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
	public bool EmitSectionHeadersFooters = true;
	public bool EmitUnreferencedBookmarks = true;
	public bool EmitPageNumbers = false;
	public bool EmitFieldInstructions = true;
	public bool UsePlainComments = false;
	public bool EmitCustomProperties = true;
	public bool EmitTimeline = false;
	public string? StylesheetHref = null;
	public bool EmbedDefaultStylesheet = true;
	public string RootCssClass = "dxp-root";
	public DxpTrackedChangeMode TrackedChangeMode = DxpTrackedChangeMode.InlineChanges;

	public static DxpHtmlVisitorConfig RICH = new();
	public static DxpHtmlVisitorConfig PLAIN = new() {
		EmitImages = false,
		EmitStyleFont = false,
		EmitRunColor = false,
		EmitRunBackground = false,
		EmitTableBorders = false,
		EmitDocumentColors = false,
		EmitParagraphAlignment = false,
		PreserveListSymbols = false,
		RichTables = false,
		EmitSectionHeadersFooters = true,
		EmitUnreferencedBookmarks = false,
		EmitPageNumbers = false,
		UsePlainComments = true,
		EmitCustomProperties = true,
		EmitTimeline = false
	};

	public static DxpHtmlVisitorConfig DEFAULT = RICH;
}

public sealed class DxpHtmlVisitor : DxpVisitor, DxpITextVisitor
{
	private TextWriter _sinkWriter;
	private DxpBufferedTextWriter _rejectBufferedWriter;
	private DxpBufferedTextWriter _acceptBufferedWriter;

	private readonly DxpHtmlVisitorConfig _config;
	private DxpHtmlVisitorState _state = new();

	public DxpHtmlVisitor(TextWriter writer, DxpHtmlVisitorConfig config, ILogger? logger)
		: base(logger)
	{
		_config = config;
		_sinkWriter = writer;
		_rejectBufferedWriter = new DxpBufferedTextWriter();
		_acceptBufferedWriter = new DxpBufferedTextWriter();
		ConfigureWriters();
	}

	public DxpHtmlVisitor(DxpHtmlVisitorConfig config, ILogger? logger = null)
		: this(TextWriter.Null, config, logger)
	{
	}

	public void SetOutput(TextWriter writer)
	{
		_sinkWriter = writer ?? throw new ArgumentNullException(nameof(writer));
		_rejectBufferedWriter = new DxpBufferedTextWriter();
		_acceptBufferedWriter = new DxpBufferedTextWriter();
		_state = new DxpHtmlVisitorState();
		ConfigureWriters();
	}

	public override void SetOutput(Stream stream)
	{
		var writer = new StreamWriter(stream, Encoding.UTF8, bufferSize: 1024, leaveOpen: true);
		SetOutput(writer);
	}

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

	private static string DefaultStylesheet => """
:root {
  --dxp-text: #111111;
  --dxp-background: #ffffff;
  --dxp-chrome: #f2f2f2;
  --dxp-border: #d0d0d0;
  --dxp-accent: #005fb8;
  --dxp-muted: #4a4a4a;
}

body.dxp-root {
  margin: 0;
  padding: 0;
  background: var(--dxp-chrome);
  color: var(--dxp-text);
  font-family: "Segoe UI", "Calibri", "Helvetica Neue", Arial, sans-serif;
}

.dxp-document {
  background: var(--dxp-background);
  margin: 1rem auto;
  padding: 1rem;
  max-width: 8.5in;
  box-shadow: 0 1px 4px rgba(0,0,0,0.16);
}

.dxp-section {
  position: relative;
  display: flex;
  flex-direction: column;
}

.dxp-body {
  flex: 1 0 auto;
}

.dxp-header,
.dxp-footer {
  font-size: 0.95em;
  color: var(--dxp-muted);
}

.dxp-heading {
  margin: 0.2em 0 0.4em;
  font-weight: 600;
}
.dxp-heading-1 { font-size: 1.8em; }
.dxp-heading-2 { font-size: 1.5em; }
.dxp-heading-3 { font-size: 1.3em; }
.dxp-heading-4 { font-size: 1.15em; }
.dxp-heading-5 { font-size: 1.05em; }
.dxp-heading-6 { font-size: 1em; }

.dxp-paragraph {
  margin: 0 0 0.6em;
  line-height: 1.4;
}
.dxp-paragraph.align-center { text-align: center; }
.dxp-paragraph.align-right { text-align: right; }
.dxp-paragraph.align-justify { text-align: justify; }
.dxp-paragraph .dxp-marker { display: inline-block; min-width: 1.4em; }

.dxp-blockquote {
  border-left: 3px solid var(--dxp-border);
  margin: 0.4em 0;
  padding: 0.2em 1em;
  color: var(--dxp-muted);
}

.dxp-code {
  background: #f6f8fa;
  border: 1px solid #e0e0e0;
  border-radius: 4px;
  padding: 0.5em;
  font-family: Consolas, Monaco, monospace;
  white-space: pre-wrap;
}

.dxp-table {
  border-collapse: collapse;
  width: 100%;
  margin: 0.5em 0;
}
.dxp-table td, .dxp-table th {
  border: 1px solid var(--dxp-border);
  padding: 4px 6px;
  vertical-align: top;
}
.dxp-table .header-row th {
  background: #f4f4f4;
}

.dxp-image {
  max-width: 100%;
  height: auto;
}

.dxp-caption {
  font-size: 0.95em;
  color: var(--dxp-muted);
}

.dxp-footnote {
  font-size: 0.9em;
  border-top: 1px solid var(--dxp-border);
  margin-top: 1em;
  padding-top: 0.5em;
}

.dxp-comment {
  background: #fffbeb;
  border: 1px solid #e6c44a;
  border-radius: 6px;
  padding: 6px;
  margin-bottom: 6px;
}
.dxp-comments {
  background: #fff8c6;
  border: 1px solid #e6c44a;
  border-radius: 6px;
  padding: 8px;
  margin: 8px 0 8px 12px;
}

.dxp-underline { text-decoration: underline; }
.dxp-strike { text-decoration: line-through; }
.dxp-inserted { text-decoration: underline; color: var(--dxp-accent); }
.dxp-deleted { text-decoration: line-through; color: #a40000; }

.dxp-split {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 8px;
  border: 1px solid var(--dxp-border);
}
.dxp-split-column {
  padding: 8px;
}
""";

	public override IDisposable VisitDocumentBegin(WordprocessingDocument doc, DxpIDocumentContext d)
	{
		_sinkWriter.WriteLine("<!DOCTYPE html>");
		_sinkWriter.WriteLine("<html lang=\"en\">");
		_sinkWriter.WriteLine("<head>");
		_sinkWriter.WriteLine("  <meta charset=\"utf-8\" />");

		IPackageProperties core = d.DocumentProperties.core;

		if (!string.IsNullOrWhiteSpace(core.Title))
			_sinkWriter.WriteLine($"  <title>{WebUtility.HtmlEncode(core.Title)}</title>");
		else
			_sinkWriter.WriteLine("  <title>Document</title>");

		if (!string.IsNullOrEmpty(_config.StylesheetHref))
		{
			_sinkWriter.WriteLine($"  <link rel=\"stylesheet\" href=\"{HtmlAttr(_config.StylesheetHref)}\" />");
		}
		else if (_config.EmbedDefaultStylesheet)
		{
			_sinkWriter.WriteLine("  <style>");
			_sinkWriter.WriteLine(DefaultStylesheet);
			_sinkWriter.WriteLine("  </style>");
		}

		_sinkWriter.WriteLine("</head>");
		_sinkWriter.WriteLine($"<body class=\"{_config.RootCssClass}\">");

		return DxpDisposable.Create(() => {
			_sinkWriter.WriteLine("</body>");
			_sinkWriter.WriteLine("</html>");
			_sinkWriter.Flush();
		});
	}

	public override IDisposable VisitDocumentBodyBegin(Body body, DxpIDocumentContext d)
	{
		WriteLine(d, $"""<div class="dxp-document">""");
		return DxpDisposable.Create(() => {
			WriteLine(d, "</div>");
		});
	}

	public override void VisitText(Text t, DxpIDocumentContext d)
	{
		string text = t.Text;
		if (_state.AllCaps)
		{
			var culture = CultureInfo.InvariantCulture;
			text = text.ToUpper(culture);
		}

		Write(d, WebUtility.HtmlEncode(text));
	}

	public override void StyleAllCapsEnd(DxpIDocumentContext d) => _state.AllCaps = false;
	public override void StyleAllCapsBegin(DxpIDocumentContext d) => _state.AllCaps = true;
	public override void StyleSmallCapsEnd(DxpIDocumentContext d) => _state.AllCaps = false;
	public override void StyleSmallCapsBegin(DxpIDocumentContext d) => _state.AllCaps = true;

	public override void StyleBoldBegin(DxpIDocumentContext d)
	{
		if (_state.InHeading)
			return;
		Write(d, "<strong class=\"dxp-bold\">");
	}
	public override void StyleBoldEnd(DxpIDocumentContext d)
	{
		if (_state.InHeading)
			return;
		Write(d, "</strong>");
	}

	public override void StyleItalicBegin(DxpIDocumentContext d) => Write(d, "<em class=\"dxp-italic\">");
	public override void StyleItalicEnd(DxpIDocumentContext d) => Write(d, "</em>");

	public override void StyleUnderlineBegin(DxpIDocumentContext d) => Write(d, "<span class=\"dxp-underline\">");
	public override void StyleUnderlineEnd(DxpIDocumentContext d) => Write(d, "</span>");

	public override void StyleStrikeBegin(DxpIDocumentContext d) => Write(d, "<span class=\"dxp-strike\">");
	public override void StyleStrikeEnd(DxpIDocumentContext d) => Write(d, "</span>");

	public override void StyleDoubleStrikeBegin(DxpIDocumentContext d) => Write(d, "<span class=\"dxp-strike\">");
	public override void StyleDoubleStrikeEnd(DxpIDocumentContext d) => Write(d, "</span>");

	public override void StyleSuperscriptBegin(DxpIDocumentContext d) => Write(d, "<sup>");
	public override void StyleSuperscriptEnd(DxpIDocumentContext d) => Write(d, "</sup>");
	public override void StyleSubscriptBegin(DxpIDocumentContext d) => Write(d, "<sub>");
	public override void StyleSubscriptEnd(DxpIDocumentContext d) => Write(d, "</sub>");

	public override void StyleFontBegin(DxpFont font, DxpIDocumentContext d)
	{
		if (_config.EmitStyleFont == false)
			return;

		if (IsDefaultFont(font.fontName, font.fontSizeHalfPoints, d))
		{
			_state.FontSpanOpen = false;
			return;
		}

		_state.FontSpanOpen = true;
		Write(d, $"""<span class="dxp-font" style="font-family:{WebUtility.HtmlEncode(font.fontName ?? string.Empty)};font-size:{font.fontSizeHalfPoints / 2.0}pt;">""");
	}

	public override void StyleFontEnd(DxpIDocumentContext d)
	{
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
				return;
		}

		if (!string.IsNullOrEmpty(name))
			Write(d, $"<a id=\"{HtmlAttr(name!)}\" data-bookmark-id=\"{id}\"></a>");
	}

	public override IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target, DxpIDocumentContext d)
	{
		string? href = target?.uri;
		Write(d, href != null ? $"<a class=\"dxp-link\" href=\"{HtmlAttr(href)}\">" : "<a class=\"dxp-link\">");
		return DxpDisposable.Create(() => Write(d, "</a>"));
	}

	public override IDisposable VisitInsertedBegin(Inserted ins, DxpIDocumentContext d) => DxpDisposable.Empty;
	public override IDisposable VisitDeletedBegin(Deleted del, DxpIDocumentContext d) => DxpDisposable.Empty;
	public override IDisposable VisitDeletedRunBegin(DeletedRun dr, DxpIDocumentContext d) => DxpDisposable.Empty;
	public override void VisitDeletedParagraphMark(Deleted del, ParagraphProperties pPr, Paragraph? p, DxpIDocumentContext d) { }
	public override IDisposable VisitInsertedRunBegin(InsertedRun ir, DxpIDocumentContext d) => DxpDisposable.Empty;

	public override void VisitDeletedText(DeletedText dt, DxpIDocumentContext d)
	{
		Write(d, WebUtility.HtmlEncode(dt.Text));
	}

	public override void VisitNoBreakHyphen(NoBreakHyphen h, DxpIDocumentContext d) => Write(d, "-");

	public override void VisitDrawingBegin(Drawing drw, DxpDrawingInfo? info, DxpIDocumentContext d)
	{
		if (_config.EmitImages == false)
		{
			Write(d, "<span class=\"dxp-image\">[IMAGE]</span>");
			return;
		}

		var alt = HtmlAttr(info?.AltText ?? "image");
		var dataUri = info?.DataUri;
		var contentType = info?.ContentType ?? "";

		if (!string.IsNullOrEmpty(dataUri) && contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
		{
			Write(d, $"<img class=\"dxp-image\" src=\"{dataUri}\" alt=\"{alt}\" />");
		}
		else if (!string.IsNullOrEmpty(dataUri))
		{
			Write(d, $"<object data=\"{dataUri}\" type=\"{HtmlAttr(contentType)}\">[DRAWING: {alt}]</object>");
		}
		else
		{
			var meta = string.IsNullOrEmpty(contentType) ? "" : $" ({contentType})";
			Write(d, $"<span class=\"dxp-image\">[DRAWING: {alt}{meta}]</span>");
		}
	}

	public new void VisitLegacyPictureBegin(Picture pict, DxpIDocumentContext d)
	{
		if (_config.EmitImages == false)
		{
			Write(d, "<span class=\"dxp-image\">[IMAGE]</span>");
			return;
		}

		var alt = "image";
		Write(d, $"<span class=\"dxp-image\">[PICTURE: {alt}]</span>");
	}

	static string HtmlAttr(string s) =>
		s.Replace("&", "&amp;").Replace("\"", "&quot;").Replace("<", "&lt;").Replace(">", "&gt;");

	public override IDisposable VisitParagraphBegin(Paragraph p, DxpIDocumentContext d, DxpIParagraphContext paragraph)
	{
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
			return DxpDisposable.Create(() => WriteLine(d));
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

		double adjustedMargin = indent.Left.HasValue ? AdjustMarginLeft(DxpTwipValue.ToPoints(indent.Left.Value), d) : 0.0;
		bool hasMargin = adjustedMargin > 0.0001;

		if (isHeading)
			marker = null;
		if (isBlockQuote)
			marker = null;
		if (isCode)
			marker = null;

		var paraClasses = new List<string>();
		string tag = "p";

		if (isHeading)
		{
			tag = $"h{headingLevel}";
			paraClasses.Add("dxp-heading");
			paraClasses.Add($"dxp-heading-{headingLevel}");
		}
		else if (isCaption)
		{
			tag = "figcaption";
			paraClasses.Add("dxp-caption");
		}
		else if (isCode)
		{
			tag = "pre";
			paraClasses.Add("dxp-code");
		}
		else
		{
			paraClasses.Add("dxp-paragraph");
		}

		if (justify != null && !_state.InHeading)
			paraClasses.Add($"align-{justify}");

		var style = new StringBuilder();
		if (hasMargin)
			style.Append("margin-left:").Append(adjustedMargin.ToString("0.###", CultureInfo.InvariantCulture)).Append("pt;");

		bool previousHeading = _state.InHeading;
		if (isHeading)
			_state.InHeading = true;

		if (isBlockQuote)
			WriteLine(d, """<blockquote class="dxp-blockquote">""");

		var openTag = new StringBuilder();
		openTag.Append('<').Append(tag);
		if (paraClasses.Count > 0)
			openTag.Append(" class=\"").Append(string.Join(" ", paraClasses)).Append('"');
		if (style.Length > 0)
			openTag.Append(" style=\"").Append(style).Append('"');
		openTag.Append('>');
		Write(d, openTag.ToString());

		if (marker?.marker != null)
		{
			var normalizedMarker = NormalizeMarker(marker.marker);
			Write(d, BuildMarkerHtml(normalizedMarker));
		}

		if (isCode)
			Write(d, "<code>");

		var baseDispose = DxpDisposable.Create(() => {
			if (_config.TrackedChangeMode == DxpTrackedChangeMode.InlineChanges)
				SetInlineChangeMode(DxpHtmlVisitorState.InlineChangeMode.Unchanged);

			if (isCode)
				Write(d, "</code>");

			Write(d, $"</{tag}>");

			if (isBlockQuote)
				WriteLine(d, "</blockquote>");
			else
				WriteLine(d);

			_state.InHeading = previousHeading;

			if (_config.TrackedChangeMode == DxpTrackedChangeMode.SplitChanges)
				EmitSplitBuffersIfNeeded();
		});

		return DxpDisposable.Create(() => {
			baseDispose.Dispose();
		});
	}

	public override void VisitFootnoteReference(FootnoteReference fr, DxpIFootnoteContext footnote, DxpIDocumentContext d)
	{
		Write(d, $"<a class=\"dxp-footnote-ref\" href=\"#fn-{footnote.Id}\" id=\"fnref-{footnote.Id}\">[{footnote.Index}]</a>");
	}

	public override IDisposable VisitSectionHeaderBegin(Header hdr, object kind, DxpIDocumentContext d)
	{
		if (_config.EmitSectionHeadersFooters == false)
			return DxpDisposable.Empty;

		WriteLine(d, """<div class="dxp-header">""");

		return DxpDisposable.Create(() => {
			WriteLine(d, "</div>");
		});
	}

	public override IDisposable VisitSectionFooterBegin(Footer ftr, object kind, DxpIDocumentContext d)
	{
		WriteLine(d, """<div class="dxp-footer">""");

		return DxpDisposable.Create(() => {
			WriteLine(d, "</div>");
		});
	}

	public override void VisitPageNumber(PageNumber pn, DxpIDocumentContext d)
	{
	}

	public override void VisitComplexFieldBegin(FieldChar begin, DxpIDocumentContext d) { }

	public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
	{
		EmitFieldInstruction(d, text);
	}

	public override IDisposable VisitComplexFieldResultBegin(DxpIDocumentContext d) => DxpDisposable.Empty;

	public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
	{
		if (!_config.EmitPageNumbers)
		{
			var instr = d.CurrentFields.Current?.InstructionText;
			if (LooksLikePageField(instr))
				return;
		}
		Write(d, WebUtility.HtmlEncode(text));
	}

	public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d) { }

	public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
	{
		var instr = fld.Instruction?.Value;
		if (instr != null)
			EmitFieldInstruction(d, instr);
		return DxpDisposable.Empty;
	}

	public override IDisposable VisitFootnoteBegin(Footnote fn, DxpIFootnoteContext footnote, DxpIDocumentContext d)
	{
		WriteLine(d, $"""<div class="dxp-footnote" id="fn-{footnote.Id}">""");
		return DxpDisposable.Create(() => WriteLine(d, "</div>"));
	}

	public override void VisitFootnoteReferenceMark(FootnoteReferenceMark m, DxpIFootnoteContext footnote, DxpIDocumentContext d)
	{
		if (footnote.Index != null)
			Write(d, $"{footnote.Index}");
	}

	public override IDisposable VisitTableBegin(Table t, DxpTableModel model, DxpIDocumentContext d, DxpITableContext table)
	{
		var styles = _config.EmitTableBorders && table.Properties != null
			? BuildTableStyle(table.Properties)
			: (null, null);

		var currentStyle = styles.tableStyle;

		Write(d, """
			<table class="dxp-table"
			""");
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
		var cellBorders = cell.Properties?.TableCellBorders;
		var cellStyle = _config.EmitTableBorders ? BuildCellStyle(cellBorders) : null;

		Write(d, "    <td");
		if (spans.Item1 > 1)
			Write(d, $" rowspan=\"{spans.Item1}\"");
		if (spans.Item2 > 1)
			Write(d, $" colspan=\"{spans.Item2}\"");
		string? borderCss = null;
		if (_config.EmitTableBorders && cell.Row.Table.Properties != null)
		{
			borderCss = BuildTableStyle(cell.Row.Table.Properties).cellBorderStyle;
		}
		var effectiveCellStyle = cellStyle ?? (borderCss != null ? $"border:{borderCss};" : null);
		if (!string.IsNullOrEmpty(effectiveCellStyle))
			Write(d, $" style=\"{effectiveCellStyle}\"");
		Write(d, ">");

		return DxpDisposable.Create(() => {
			WriteLine(d, "</td>");
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
				WriteLine(d, $"  {label} {WebUtility.HtmlEncode(who)} ON {when}");
			else
				WriteLine(d, $"  {label} {WebUtility.HtmlEncode(who)}");

			WriteLine(d);

			return DxpDisposable.Create(() => {
				WriteLine(d);
			});
		}
		else
		{
			var label = BuildCommentLabel(c);
			if (!string.IsNullOrEmpty(label))
				WriteLine(d, "  " + label);

			WriteLine(d, """  <div class="dxp-comment">""");

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

		WriteLine(d, """<div class="dxp-comments">""");

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

	public override void VisitBreak(Break br, DxpIDocumentContext d) => Write(d, "<br/>");
	public override void VisitCarriageReturn(CarriageReturn cr, DxpIDocumentContext d) => Write(d, "<br/>");
	public override void VisitTab(TabChar tab, DxpIDocumentContext d) => Write(d, "&#9;");

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

		if (!_config.PreserveListSymbols)
		{
			var translated = TryTranslateSymbolFont(marker);
			if (!string.IsNullOrEmpty(translated))
				return translated!;
		}

		return trimmed;
	}

	private static bool LooksLikeOrderedListMarker(string marker) => Regex.IsMatch(marker, @"^\d+[.)]$");

	private static string BuildMarkerHtml(string marker)
	{
		if (string.IsNullOrEmpty(marker))
			return string.Empty;

		bool markerIsHtml = marker.IndexOf('<') >= 0 && marker.IndexOf('>') > marker.IndexOf('<');
		var inner = markerIsHtml ? marker : WebUtility.HtmlEncode(marker);
		return $"""<span class="dxp-marker">{inner}</span> """;
	}

	private static string StripTags(string input) => Regex.Replace(input, "<.*?>", string.Empty);

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

	public override IDisposable VisitSectionBodyBegin(SectionProperties properties, DxpIDocumentContext d)
	{
		if (!_config.EmitDocumentColors)
			return DxpDisposable.Empty;

		var style = new StringBuilder("flex:1 0 auto;");

		double? marginTopInches = d.CurrentSection.Layout?.MarginTop?.Inches;
		if (marginTopInches != null)
			style.Append("padding-top:").Append(marginTopInches.Value.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");

		Write(d, $"""<div class="dxp-body" style="{style}">""" + "\n");

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

		string pageBackground = "#ffffff";
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

		Write(d, $"""<div class="dxp-section" style="{style}">""" + "\n");

		return DxpDisposable.Create(() => {
			WriteLine(d, "</div>");
		});
	}

	private void WriteLine(DxpIDocumentContext d) => Write(d, "\n");
	private void WriteLine(DxpIDocumentContext d, string str) => Write(d, $"{str}\n");

	private void EmitFieldInstruction(DxpIDocumentContext d, string instruction)
	{
		if (!_config.EmitFieldInstructions)
			return;

		var trimmed = instruction.Trim();
		if (trimmed.Length == 0)
			return;

		Write(d, $"<span class=\"dxp-field\" data-field=\"{HtmlAttr(trimmed)}\"></span>");
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

	private void WriteRouted(DxpHtmlVisitorState.InlineChangeMode mode, string str)
	{
		if (mode == DxpHtmlVisitorState.InlineChangeMode.Inserted)
			_state.InsertedTextWriter.Write(str);
		else if (mode == DxpHtmlVisitorState.InlineChangeMode.Deleted)
			_state.DeletedTextWriter.Write(str);
		else
			_state.UnchangedTextWriter.Write(str);
	}

	private DxpHtmlVisitorState.InlineChangeMode DetermineChangeMode(DxpIDocumentContext d)
	{
		if (d.KeepAccept && d.KeepReject)
			return DxpHtmlVisitorState.InlineChangeMode.Unchanged;
		if (d.KeepAccept && !d.KeepReject)
			return DxpHtmlVisitorState.InlineChangeMode.Inserted;
		if (!d.KeepAccept && d.KeepReject)
			return DxpHtmlVisitorState.InlineChangeMode.Deleted;
		return DxpHtmlVisitorState.InlineChangeMode.Unchanged;
	}

	private void SetInlineChangeMode(DxpHtmlVisitorState.InlineChangeMode mode)
	{
		if (mode == _state.CurrentInlineMode)
			return;

		if (_state.CurrentInlineMode == DxpHtmlVisitorState.InlineChangeMode.Inserted)
			WriteRouted(DxpHtmlVisitorState.InlineChangeMode.Unchanged, "</span>");
		else if (_state.CurrentInlineMode == DxpHtmlVisitorState.InlineChangeMode.Deleted)
			WriteRouted(DxpHtmlVisitorState.InlineChangeMode.Unchanged, "</span>");

		if (mode == DxpHtmlVisitorState.InlineChangeMode.Inserted)
			WriteRouted(DxpHtmlVisitorState.InlineChangeMode.Unchanged, InlineInsertedTag());
		else if (mode == DxpHtmlVisitorState.InlineChangeMode.Deleted)
			WriteRouted(DxpHtmlVisitorState.InlineChangeMode.Unchanged, InlineDeletedTag());

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
			_sinkWriter.Write("""<div class="dxp-split">""");
			_sinkWriter.Write("""<div class="dxp-split-column">""");
			_sinkWriter.Write(rejected);
			_sinkWriter.Write("</div>");
			_sinkWriter.Write("""<div class="dxp-split-column">""");
			_sinkWriter.Write(accepted);
			_sinkWriter.Write("</div>");
			_sinkWriter.Write("</div>");
		}
	}

	private string InlineInsertedTag()
	{
		if (_config.EmitRunColor)
			return "<span class=\"dxp-inserted\">";
		return "<span>";
	}

	private string InlineDeletedTag()
	{
		if (_config.EmitRunColor)
			return "<span class=\"dxp-deleted\">";
		return "<span>";
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

	IDisposable DxpIVisitor.VisitLegacyPictureBegin(Picture pict, DxpIDocumentContext d)
	{
		VisitLegacyPictureBegin(pict, d);
		return DxpDisposable.Empty;
	}
}
