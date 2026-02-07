using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocxportNet.API;
using Microsoft.Extensions.Logging;
using System.Globalization;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using DocxportNet.Word;
using System.Text;
using DocxportNet.Core;
using System.Net;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Fields;

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
	public bool SectionHasHeader { get; set; }
	public bool SectionHasFooter { get; set; }
	public bool InHeader { get; set; }
	public bool InFooter { get; set; }
	public InlineChangeMode CurrentInlineMode { get; set; } = InlineChangeMode.Unchanged;
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
	public int SuppressFieldDepth { get; set; }
	public Stack<bool> ComplexFieldEvaluated { get; } = new();

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
	public DxpHeaderFooterSelection HeaderSelection = DxpHeaderFooterSelection.First;
	public DxpHeaderFooterSelection FooterSelection = DxpHeaderFooterSelection.First;

	public static DxpHtmlVisitorConfig CreateRichConfig() => new();
	public static DxpHtmlVisitorConfig CreatePlainConfig() => new() {
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

	public static DxpHtmlVisitorConfig CreateConfig() => CreateRichConfig();
}

public sealed class DxpHtmlVisitor : DxpVisitor, DxpITextVisitor, IDxpHeaderFooterSelectionProvider, IDisposable, IDxpFieldEvalProvider
{
	private TextWriter _sinkWriter;
	private StreamWriter? _ownedStreamWriter;
	private DxpBufferedTextWriter _rejectBufferedWriter;
	private DxpBufferedTextWriter _acceptBufferedWriter;

	private readonly DxpHtmlVisitorConfig _config;
	private DxpHtmlVisitorState _state = new();
	private readonly DxpFieldEval _fieldEval;

	public DxpFieldEval FieldEval => _fieldEval;

	public DxpHtmlVisitor(TextWriter writer, DxpHtmlVisitorConfig config, ILogger? logger, DxpFieldEval? fieldEval = null)
		: base(logger)
	{
		_config = config;
		_sinkWriter = writer;
		_rejectBufferedWriter = new DxpBufferedTextWriter();
		_acceptBufferedWriter = new DxpBufferedTextWriter();
		_fieldEval = fieldEval ?? new DxpFieldEval();
		ConfigureWriters();
	}

	public DxpHtmlVisitor(DxpHtmlVisitorConfig config, ILogger? logger = null, DxpFieldEval? fieldEval = null)
		: this(TextWriter.Null, config, logger, fieldEval)
	{
	}

	public DxpHeaderFooterSelection HeaderSelection => _config.HeaderSelection;
	public DxpHeaderFooterSelection FooterSelection => _config.FooterSelection;

	public void SetOutput(TextWriter writer)
	{
		ReleaseOwnedWriter();
		_sinkWriter = writer ?? throw new ArgumentNullException(nameof(writer));
		_rejectBufferedWriter = new DxpBufferedTextWriter();
		_acceptBufferedWriter = new DxpBufferedTextWriter();
		_state = new DxpHtmlVisitorState();
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
  z-index: 0;
  overflow: hidden;
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
  margin: 0;
  font-weight: 600;
  line-height: 1.2;
}
.dxp-heading-1 { font-size: inherit; }
.dxp-heading-2 { font-size: inherit; }
.dxp-heading-3 { font-size: inherit; }
.dxp-heading-4 { font-size: inherit; }
.dxp-heading-5 { font-size: inherit; }
.dxp-heading-6 { font-size: inherit; }
.dxp-heading.align-center { text-align: center; }
.dxp-heading.align-right { text-align: right; }
.dxp-heading.align-justify { text-align: justify; }

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

		IPackageProperties? core = d.DocumentProperties.PackageProperties;
		string? title = core?.Title;

		if (!string.IsNullOrWhiteSpace(title))
			_sinkWriter.WriteLine($"  <title>{WebUtility.HtmlEncode(title)}</title>");
		else
			_sinkWriter.WriteLine("  <title>Document</title>");

		string? stylesheetHref = _config.StylesheetHref;
		if (!string.IsNullOrEmpty(stylesheetHref))
		{
			_sinkWriter.WriteLine($"  <link rel=\"stylesheet\" href=\"{HtmlAttr(stylesheetHref!)}\" />");
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

		if (_state.PendingAlignedTabKind != null)
			_state.PendingAlignedTabSegmentWidthPt += EstimateTextWidthPt(text, _state.CurrentFontSizePt);
		else if (_state.InParagraph)
			_state.CurrentLineXPt += EstimateTextWidthPt(text, _state.CurrentFontSizePt);

		Write(d, WebUtility.HtmlEncode(text));
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

	public override void StyleUnderlineBegin(DxpIDocumentContext d)
	{
		_state.UnderlineDepth++;
		Write(d, "<span class=\"dxp-underline\">");
	}

	public override void StyleUnderlineEnd(DxpIDocumentContext d)
	{
		if (_state.UnderlineDepth > 0)
			_state.UnderlineDepth--;
		Write(d, "</span>");
	}

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
			style.Append("font-family:").Append(WebUtility.HtmlEncode(font.fontName)).Append(';');
		if (font.fontSizeHalfPoints != null)
			style.Append("font-size:").Append(font.fontSizeHalfPoints.Value / 2.0).Append("pt;");

		if (style.Length == 0)
		{
			_state.FontSpanOpen = false;
			return;
		}

		_state.FontSpanOpen = true;
		Write(d, $"""<span class="dxp-font" style="{style}">""");
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
			var style = BuildDrawingImageStyle(drw, d, _state.InHeader, _state.InFooter);
			var styleAttr = string.IsNullOrEmpty(style) ? "" : $" style=\"{style}\"";
			Write(d, $"<img class=\"dxp-image\" src=\"{dataUri}\" alt=\"{alt}\"{styleAttr} />");
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

	private static string? BuildDrawingImageStyle(Drawing drw, DxpIDocumentContext d, bool inHeader, bool inFooter)
	{
		var sb = new StringBuilder();

		static void Append(StringBuilder sb, string css)
		{
			if (string.IsNullOrEmpty(css))
				return;
			if (sb.Length > 0 && sb[sb.Length - 1] != ';')
				sb.Append(';');
			sb.Append(css);
			if (sb.Length > 0 && sb[sb.Length - 1] != ';')
				sb.Append(';');
		}

		// Sizing: use wp:extent when available (EMU units).
		var extent = drw.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>().FirstOrDefault();
		if (extent?.Cx != null && extent?.Cy != null)
		{
			var widthPt = EmuToPoints(extent.Cx.Value);
			var heightPt = EmuToPoints(extent.Cy.Value);
			if (widthPt > 0.01)
				Append(sb, $"width:{widthPt.ToString("0.###", CultureInfo.InvariantCulture)}pt");
			if (heightPt > 0.01)
				Append(sb, $"height:{heightPt.ToString("0.###", CultureInfo.InvariantCulture)}pt");
			// If we have explicit dimensions, don't let global max-width force scaling.
			Append(sb, "max-width:none");
		}

		// Positioning: for floating drawings (wp:anchor), surface a best-effort approximation.
		// We prefer relative positioning (keeps layout flow; avoids overlapping following content) and
		// adjust page-relative offsets using the current section margins when available.
		var anchor = drw.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>().FirstOrDefault();
		if (anchor != null)
		{
			var positionCss = TryBuildAnchorPositionCss(anchor, d, inHeader, inFooter);
			if (!string.IsNullOrEmpty(positionCss))
			{
				Append(sb, "display:block");
				Append(sb, "position:relative");
				Append(sb, positionCss!);
			}
		}

		return sb.Length == 0 ? null : sb.ToString();
	}

	private static string? TryBuildAnchorPositionCss(DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor anchor, DxpIDocumentContext d, bool inHeader, bool inFooter)
	{
		// We use local-name based parsing to stay robust across SDK differences.
		OpenXmlElement? posH = anchor.ChildElements.FirstOrDefault(e => e.LocalName == "positionH");
		OpenXmlElement? posV = anchor.ChildElements.FirstOrDefault(e => e.LocalName == "positionV");

		string? transform = null;
		var sb = new StringBuilder();

		static string? TryGetAttribute(OpenXmlElement el, string localName)
			=> el.GetAttributes().FirstOrDefault(a => a.LocalName == localName).Value;

		static bool IsPageReference(string? relativeFrom)
		{
			return string.Equals(relativeFrom, "page", StringComparison.OrdinalIgnoreCase);
		}

		void Append(string name, string value)
		{
			if (sb.Length > 0 && sb[sb.Length - 1] != ';')
				sb.Append(';');
			sb.Append(name).Append(':').Append(value).Append(';');
		}

		static string FormatInches(double inches)
			=> inches.ToString("0.###", CultureInfo.InvariantCulture) + "in";

		double marginLeftIn = d.CurrentSection.Layout?.MarginLeft?.Inches ?? 0.0;
		double marginRightIn = d.CurrentSection.Layout?.MarginRight?.Inches ?? 0.0;
		double pageCenterShiftIn = (marginRightIn - marginLeftIn) / 2.0;

		// Note: block-level alignment is handled via margins; offsets are applied via transforms.
		if (posH != null)
		{
			var relativeFrom = TryGetAttribute(posH, "relativeFrom")?.Trim();
			var align = posH.ChildElements.FirstOrDefault(e => e.LocalName == "align")?.InnerText?.Trim().ToLowerInvariant();
			var offset = posH.ChildElements.FirstOrDefault(e => e.LocalName == "posOffset")?.InnerText?.Trim();
			long.TryParse(offset, out var offsetEmu);
			var offsetIn = EmuToInches(offsetEmu);
			var offsetPresent = !string.IsNullOrEmpty(offset);

			// Compute shift relative to the content box (paragraph/content area).
			// If relativeFrom="page", Word offsets are from the page edge; our HTML content is inset by margins.
			bool fromPage = IsPageReference(relativeFrom);

			if (align == "right")
			{
				Append("margin-left", "auto");
				var shiftIn = fromPage ? (marginRightIn - (offsetPresent ? offsetIn : 0.0)) : -(offsetPresent ? offsetIn : 0.0);
				if (Math.Abs(shiftIn) > 0.0005)
					transform = AppendTransform(transform, "translateX(" + FormatInches(shiftIn) + ")");
			}
			else if (align == "center")
			{
				Append("margin-left", "auto");
				Append("margin-right", "auto");
				var shiftIn = (offsetPresent ? offsetIn : 0.0) + (fromPage ? pageCenterShiftIn : 0.0);
				if (Math.Abs(shiftIn) > 0.0005)
					transform = AppendTransform(transform, "translateX(" + FormatInches(shiftIn) + ")");
			}
			else if (align == "left")
			{
				var shiftIn = fromPage ? ((offsetPresent ? offsetIn : 0.0) - marginLeftIn) : (offsetPresent ? offsetIn : 0.0);
				if (Math.Abs(shiftIn) > 0.0005)
					transform = AppendTransform(transform, "translateX(" + FormatInches(shiftIn) + ")");
			}
			else if (offsetPresent)
			{
				var shiftIn = fromPage ? (offsetIn - marginLeftIn) : offsetIn;
				if (Math.Abs(shiftIn) > 0.0005)
					transform = AppendTransform(transform, "translateX(" + FormatInches(shiftIn) + ")");
			}
		}

		if (posV != null)
		{
			var relativeFrom = TryGetAttribute(posV, "relativeFrom")?.Trim();
			var ignoreVertical = IsPageReference(relativeFrom);
			if (ignoreVertical)
				goto AfterVertical;

			var align = posV.ChildElements.FirstOrDefault(e => e.LocalName == "align")?.InnerText?.Trim().ToLowerInvariant();
			var offset = posV.ChildElements.FirstOrDefault(e => e.LocalName == "posOffset")?.InnerText?.Trim();
			long.TryParse(offset, out var offsetEmu);
			var offsetIn = EmuToInches(offsetEmu);
			var offsetPresent = !string.IsNullOrEmpty(offset);
			if (inFooter && relativeFrom != null && relativeFrom.Equals("paragraph", StringComparison.OrdinalIgnoreCase) && offsetPresent)
			{
				// Best-effort: in a single-page HTML layout, footer paragraphs are placed relative to the bottom edge.
				// Word's footer "distance from bottom" (w:pgMar/@w:footer) tends to already be reflected in the footer box.
				// When we apply a raw negative posOffset, it can move the object far above the footer region.
				// Heuristic: for negative paragraph-relative offsets in footers, flip sign and compensate by footer distance.
				double footerDistIn = d.CurrentSection.Layout?.MarginFooter?.Inches ?? 0.0;
				if (offsetIn < 0)
					offsetIn = (-offsetIn) - footerDistIn;
			}

			if (align == "bottom")
			{
				var shiftIn = -(offsetPresent ? offsetIn : 0.0);
				if (Math.Abs(shiftIn) > 0.0005)
					transform = AppendTransform(transform, "translateY(" + FormatInches(shiftIn) + ")");
			}
			else if (align == "center")
			{
				var shiftIn = offsetPresent ? offsetIn : 0.0;
				if (Math.Abs(shiftIn) > 0.0005)
					transform = AppendTransform(transform, "translateY(" + FormatInches(shiftIn) + ")");
			}
			else if (align == "top")
			{
				var shiftIn = offsetPresent ? offsetIn : 0.0;
				if (Math.Abs(shiftIn) > 0.0005)
					transform = AppendTransform(transform, "translateY(" + FormatInches(shiftIn) + ")");
			}
			else if (offsetPresent)
			{
				if (Math.Abs(offsetIn) > 0.0005)
					transform = AppendTransform(transform, "translateY(" + FormatInches(offsetIn) + ")");
			}

AfterVertical:
			;
		}

		// Stacking: if the anchor indicates "behind text", try to keep it behind normal content but above page background.
		var behindDoc = anchor.GetAttributes().FirstOrDefault(a => a.LocalName == "behindDoc").Value;
		if (string.Equals(behindDoc, "1", StringComparison.Ordinal) || string.Equals(behindDoc, "true", StringComparison.OrdinalIgnoreCase))
			Append("z-index", "-1");

		if (!string.IsNullOrEmpty(transform))
			Append("transform", transform!);

		return sb.Length == 0 ? null : sb.ToString();
	}

	private static string AppendTransform(string? existing, string addition)
		=> string.IsNullOrEmpty(existing) ? addition : existing + " " + addition;

	private static double EmuToPoints(long emu) => emu / 12700.0;
	private static double EmuToInches(long emu) => emu / 914400.0;

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
			return DxpDisposable.Create(() => WriteLine(d));
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

		// Margin/borders are precomputed in the walker; keep alignment as classes for existing output shape.
		var computedParaCss = paragraph.ComputedStyle.ToCss(includeTextAlign: false);
		bool hasComputedCss = !string.IsNullOrEmpty(computedParaCss);

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
		if (hasComputedCss)
			style.Append(computedParaCss);

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
			var markerCss = TryBuildMarkerCss(marker, d);
			Write(d, BuildMarkerHtml(normalizedMarker, markerCss));
		}

		if (isCode)
			Write(d, "<code>");

		var baseDispose = DxpDisposable.Create(() => {
			FlushPendingAlignedTab(d);

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
			_state.InParagraph = false;
		});

		return DxpDisposable.Create(() => {
			baseDispose.Dispose();
		});
	}

	private void WriteTab(DxpIDocumentContext d)
	{
		var layout = d.CurrentParagraph.Layout;
		var stops = layout?.TabStops;
		if (stops == null || stops.Count == 0)
		{
			// Best-effort fallback.
			Write(d, "&#9;");
			return;
		}

		// Basic heuristic: treat tab n as advancing to tab stop n, but shrink based on preceding text width
		// (prevents large fixed tab widths from wrapping to the next line).
		FlushPendingAlignedTab(d);

		var index = _state.TabIndex++;
		var stop = index < stops.Count ? stops[index] : null;
		var kind = stop?.Kind ?? DxpComputedTabStopKind.Left;
		double stopPos = stop?.PositionPt ?? (_state.CurrentLineXPt + 36.0);

		// For Right/Center tabs, the following segment is aligned to the stop, so we must buffer that segment
		// to estimate its width before emitting the spacer.
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

		// Common Word pattern: an underlined run with only a tab should render as an underline up to the next tab stop.
		// Text-decoration doesn't paint for empty spans, so render a border-bottom when underline is active.
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

	public override void VisitFootnoteReference(FootnoteReference fr, DxpIFootnoteContext footnote, DxpIDocumentContext d)
	{
		Write(d, $"<a class=\"dxp-footnote-ref\" href=\"#fn-{footnote.Id}\" id=\"fnref-{footnote.Id}\">[{footnote.Index}]</a>");
	}

	public override IDisposable VisitSectionHeaderBegin(Header hdr, object kind, DxpIDocumentContext d)
	{
		if (_config.EmitSectionHeadersFooters == false)
			return DxpDisposable.Empty;

		_state.SectionHasHeader = true;
		_state.InHeader = true;

		var style = new StringBuilder();
		var layout = d.CurrentSection.Layout;
		double? marginTopIn = layout?.MarginTop?.Inches;
		double? headerDistIn = layout?.MarginHeader?.Inches;
		double marginLeftIn = layout?.MarginLeft?.Inches ?? 0.0;
		double marginRightIn = layout?.MarginRight?.Inches ?? 0.0;
		if (marginTopIn != null)
			style.Append("height:").Append(marginTopIn.Value.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");
		if (headerDistIn != null)
			style.Append("padding-top:").Append(headerDistIn.Value.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");
		if (marginLeftIn > 0.0005 || marginRightIn > 0.0005)
		{
			// Header/footer in Word can use the full page width; don't let section margins clip content.
			style.Append("margin-left:-").Append(marginLeftIn.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");
			style.Append("margin-right:-").Append(marginRightIn.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");
			style.Append("padding-left:").Append(marginLeftIn.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");
			style.Append("padding-right:").Append(marginRightIn.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");
			style.Append("width:calc(100% + ")
				.Append(marginLeftIn.ToString("0.###", CultureInfo.InvariantCulture)).Append("in + ")
				.Append(marginRightIn.ToString("0.###", CultureInfo.InvariantCulture)).Append("in);");
		}
		style.Append("box-sizing:border-box;overflow:visible;");

		WriteLine(d, $"""<div class="dxp-header" style="{style}">""");

		return DxpDisposable.Create(() => {
			WriteLine(d, "</div>");
			_state.InHeader = false;
		});
	}

	public override IDisposable VisitSectionFooterBegin(Footer ftr, object kind, DxpIDocumentContext d)
	{
		_state.SectionHasFooter = true;
		_state.InFooter = true;

		var style = new StringBuilder("display:flex;flex-direction:column;justify-content:flex-end;");
		var layout = d.CurrentSection.Layout;
		double? marginBottomIn = layout?.MarginBottom?.Inches;
		double? footerDistIn = layout?.MarginFooter?.Inches;
		double marginLeftIn = layout?.MarginLeft?.Inches ?? 0.0;
		double marginRightIn = layout?.MarginRight?.Inches ?? 0.0;
		if (marginBottomIn != null)
			style.Append("height:").Append(marginBottomIn.Value.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");
		if (footerDistIn != null)
			style.Append("padding-bottom:").Append(footerDistIn.Value.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");
		if (marginLeftIn > 0.0005 || marginRightIn > 0.0005)
		{
			style.Append("margin-left:-").Append(marginLeftIn.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");
			style.Append("margin-right:-").Append(marginRightIn.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");
			style.Append("padding-left:").Append(marginLeftIn.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");
			style.Append("padding-right:").Append(marginRightIn.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");
			style.Append("width:calc(100% + ")
				.Append(marginLeftIn.ToString("0.###", CultureInfo.InvariantCulture)).Append("in + ")
				.Append(marginRightIn.ToString("0.###", CultureInfo.InvariantCulture)).Append("in);");
		}
		style.Append("box-sizing:border-box;overflow:visible;");

		WriteLine(d, $"""<div class="dxp-footer" style="{style}">""");

		return DxpDisposable.Create(() => {
			WriteLine(d, "</div>");
			_state.InFooter = false;
		});
	}

	public override void VisitPageNumber(PageNumber pn, DxpIDocumentContext d)
	{
	}

	public override void VisitComplexFieldBegin(FieldChar begin, DxpIDocumentContext d)
	{
		_state.ComplexFieldEvaluated.Push(false);
	}

	public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
	{
		EmitFieldInstruction(d, text);
	}

	public override IDisposable VisitComplexFieldResultBegin(DxpIDocumentContext d) => DxpDisposable.Empty;

	public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
	{
		if (TryWriteEvaluatedComplexField(d))
			return;

		if (!_config.EmitPageNumbers)
		{
			var instr = d.CurrentFields.Current?.InstructionText;
			if (LooksLikePageField(instr))
				return;
		}
		Write(d, WebUtility.HtmlEncode(text));
	}

	public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
	{
		if (_state.ComplexFieldEvaluated.Count > 0)
			_state.ComplexFieldEvaluated.Pop();
	}

	public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
	{
		var instr = d.CurrentFields.Current?.InstructionText ?? fld.Instruction?.Value;
		if (instr != null)
		{
			if (TryWriteEvaluatedSimpleField(instr, d))
				return DxpDisposable.Create(() => _state.SuppressFieldDepth--);

			EmitFieldInstruction(d, instr);
		}
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
		var currentStyle = _config.EmitTableBorders ? table.ComputedStyle.ToCss() : null;

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
	public override void VisitTab(TabChar tab, DxpIDocumentContext d) => WriteTab(d);

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

	private static string BuildMarkerHtml(string marker, string? markerCss)
	{
		if (string.IsNullOrEmpty(marker))
			return string.Empty;

		bool markerIsHtml = marker.IndexOf('<') >= 0 && marker.IndexOf('>') > marker.IndexOf('<');
		var inner = markerIsHtml ? marker : WebUtility.HtmlEncode(marker);
		var cssAttr = !string.IsNullOrEmpty(markerCss) ? " style=\"" + HtmlAttr(markerCss!) + "\"" : "";
		return $"""<span class="dxp-marker"{cssAttr}>{inner}</span> """;
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

	private static string? TryBuildMarkerCss(DxpMarker marker, DxpIDocumentContext d)
	{
		if (marker.numId == null || marker.iLvl == null)
			return null;

		if (d is not DocxportNet.Walker.DxpDocumentContext docCtx)
			return null;

		var resolved = docCtx.NumberingResolver.ResolveLevel(marker.numId.Value, marker.iLvl.Value);
		if (resolved == null)
			return null;

		var rpr = resolved.Value.lvl.NumberingSymbolRunProperties;
		if (rpr == null)
			return null;

		var decorations = new List<string>();
		var parts = new List<string>();

		if (rpr.Bold != null && IsOn(rpr.Bold.Val))
			parts.Add("font-weight:bold");
		if (rpr.Italic != null && IsOn(rpr.Italic.Val))
			parts.Add("font-style:italic");
		if (rpr.Underline?.Val != null && rpr.Underline.Val.Value != UnderlineValues.None)
			decorations.Add("underline");
		if (rpr.Strike != null && IsOn(rpr.Strike.Val))
			decorations.Add("line-through");
		if (rpr.DoubleStrike != null && IsOn(rpr.DoubleStrike.Val))
			decorations.Add("line-through");

		if (decorations.Count > 0)
			parts.Add("text-decoration:" + string.Join(" ", decorations.Distinct(StringComparer.Ordinal)));

		return parts.Count == 0 ? null : string.Join(";", parts);
	}

	private static bool IsOn(OnOffValue? v) => v == null || v.Value;

	public override IDisposable VisitSectionBodyBegin(SectionProperties properties, DxpIDocumentContext d)
	{
		if (!_config.EmitDocumentColors)
			return DxpDisposable.Empty;

		var style = new StringBuilder("flex:1 0 auto;");

		double? marginTopInches = d.CurrentSection.Layout?.MarginTop?.Inches;
		if (marginTopInches != null && _state.SectionHasHeader == false)
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

		_state.SectionHasHeader = false;
		_state.SectionHasFooter = false;

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
		if (_state.SuppressFieldDepth > 0)
			return;

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
		if (_state.PendingAlignedTabKind != null)
		{
			_state.PendingAlignedTabBuffer.Append(str);
			return;
		}

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

	private bool TryWriteEvaluatedSimpleField(string instruction, DxpIDocumentContext d)
	{
		var result = _fieldEval.EvalAsync(new DxpFieldInstruction(instruction)).GetAwaiter().GetResult();
		if (result.Status != DxpFieldEvalStatus.Resolved || result.Text == null)
			return false;

		Write(d, WebUtility.HtmlEncode(result.Text));
		_state.SuppressFieldDepth++;
		return true;
	}

	private bool TryWriteEvaluatedComplexField(DxpIDocumentContext d)
	{
		if (_state.ComplexFieldEvaluated.Count == 0)
			return false;

		if (_state.ComplexFieldEvaluated.Peek())
			return true;

		var instruction = d.CurrentFields.Current?.InstructionText;
		if (string.IsNullOrWhiteSpace(instruction))
			return false;

		var result = _fieldEval.EvalAsync(new DxpFieldInstruction(instruction)).GetAwaiter().GetResult();
		if (result.Status != DxpFieldEvalStatus.Resolved || result.Text == null)
			return false;

		Write(d, WebUtility.HtmlEncode(result.Text));
		_state.ComplexFieldEvaluated.Pop();
		_state.ComplexFieldEvaluated.Push(true);
		return true;
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
