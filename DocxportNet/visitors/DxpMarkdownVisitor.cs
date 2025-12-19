using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocxportNet.api;
using Microsoft.Extensions.Logging;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocxportNet.walker;
using DocxportNet.symbols;

namespace DocxportNet.visitors;

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

	public static DxpMarkdownVisitorConfig RICH = new();
	public static DxpMarkdownVisitorConfig PLAIN = new()
	{
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
		UsePlainComments = true
	};

	public static DxpMarkdownVisitorConfig DEFAULT = RICH;
}


public class DxpMarkdownVisitor : DxpVisitor, IDxpVisitor
{
	private TextWriter _writer;
	private DxpMarkdownVisitorConfig _config;
	private bool _currentRowIsHeader;
	private string? _pageBackground;
	private bool _hasBackgroundColor;
	private (int rowSpan, int colSpan)? _pendingCellSpan;
	private string? _currentTableStyle;
	private string? _currentTableCellBorder;
	private string? _pendingCellStyle;
	private MarkdownTableBuilder? _mdTable;
	private int _headingDepth;
	private ISet<string>? _referencedAnchors;
	private string? _normalFontName;
	private int? _normalFontSizeHp;
	private bool _fontSpanOpen;
	private bool _capturedNormal;
	private double? _pageWidthInches;
	private double? _pageHeightInches;
	private bool _bodyContainerOpen;
	private double? _marginLeftInches;
	private double? _marginRightInches;
	private double? _marginTopInches;
	private double? _marginLeftPoints;
	private readonly Stack<bool> _fieldSuppressStack = new();
	private int _suppressFieldDepth;

	public DxpMarkdownVisitor(TextWriter writer, DxpMarkdownVisitorConfig config, ILogger? logger)
		: base(logger)
	{
		_config = config;
		_writer = writer;
	}

	public bool EmitSectionHeadersFooters => _config.EmitSectionHeadersFooters;
	public bool EmitUnreferencedBookmarks => _config.EmitUnreferencedBookmarks;
	public void SetReferencedAnchors(ISet<string> anchors) => _referencedAnchors = anchors;
	public void SetDefaultSectionLayout(SectionLayout layout)
	{
		UpdatePageDimensions(layout);
	}

	public override void VisitText(Text t, IDxpStyleResolver s)
	{
			if (ShouldSuppressOutput())
				return;
			string text = t.Text;
			if (_allcaps)
			{
				var culture = CultureInfo.InvariantCulture;
				text = text.ToUpper(culture);
			}

			_writer.Write(text);
		}

	private bool _allcaps = false;

	public override void StyleAllCapsEnd()
	{
		_allcaps = false;
	}

	public override void StyleAllCapsBegin()
	{
		_allcaps = true;
	}

	public override void StyleSmallCapsEnd()
	{
		_allcaps = false;
	}

	public override void StyleSmallCapsBegin()
	{
		_allcaps = true;
	}

	public override void StyleBoldBegin()
	{
		if (_headingDepth > 0)
			return;
		_writer.Write("<b>");
	}
	public override void StyleBoldEnd()
	{
		if (_headingDepth > 0)
			return;
		_writer.Write("</b>");
	}


	public override void StyleItalicBegin()
	{
		_writer.Write("<i>");
	}
	public override void StyleItalicEnd()
	{
		_writer.Write("</i>");
	}

	public override void StyleUnderlineBegin()
	{
		_writer.Write("<u>");
	}
	public override void StyleUnderlineEnd()
	{
		_writer.Write("</u>");
	}

	public override void StyleStrikeBegin()
	{
		_writer.Write("<del>");
	}
	public override void StyleStrikeEnd()
	{
		_writer.Write("</del>");
	}

	public override void StyleDoubleStrikeBegin()
	{
		_writer.Write("<del>");
	}
	public override void StyleDoubleStrikeEnd()
	{
		_writer.Write("</del>");
	}

	public override void StyleSuperscriptBegin()
	{
		_writer.Write("<sup>");
	}
	public override void StyleSuperscriptEnd()
	{
		_writer.Write("</sup>");
	}

	public override void StyleSubscriptBegin()
	{
		_writer.Write("<sub>");
	}
	public override void StyleSubscriptEnd()
	{
		_writer.Write("</sub>");
	}

	public override void StyleFontBegin(string? fontName, int? fontSizeHalfPoints)
	{
		if (_config.EmitStyleFont == false)
			return;
		
		if (IsDefaultFont(fontName, fontSizeHalfPoints))
		{
			_fontSpanOpen = false;
			return;
		}

		_fontSpanOpen = true;
		_writer.Write($"""<span style="font-family: {fontName}; font-size: {fontSizeHalfPoints/2.0}pt;">""");
	}

	public override void StyleFontEnd()
	{
		if (_config.EmitStyleFont == false)
			return;
		
		if (_fontSpanOpen)
		{
			_writer.Write("</span>");
			_fontSpanOpen = false;
		}
	}

	private bool IsDefaultFont(string? fontName, int? fontSizeHalfPoints)
	{
		if (_normalFontName == null && _normalFontSizeHp == null)
			return false;

		bool nameMatch = _normalFontName == null || string.Equals(_normalFontName, fontName, StringComparison.OrdinalIgnoreCase);
		bool sizeMatch = _normalFontSizeHp == null || _normalFontSizeHp == fontSizeHalfPoints;
		return nameMatch && sizeMatch;
	}



	public override void VisitBookmarkStart(BookmarkStart bs, IDxpStyleResolver s)
	{
		var name = bs.Name?.Value;
		var id = bs.Id?.Value;

		if (!EmitUnreferencedBookmarks)
		{
			if (string.IsNullOrEmpty(name) || _referencedAnchors != null && !_referencedAnchors.Contains(name!))
				return;
		}

		// Skip Word’s internal _GoBack if desired; current behavior surfaces it.
		if (!string.IsNullOrEmpty(name))
			_writer.Write($"<a id=\"{Escape(name!)}\" data-bookmark-id=\"{id}\"></a>");
	}

	public override void VisitBookmarkEnd(BookmarkEnd be, IDxpStyleResolver s)
	{
		// Usually nothing to emit; it just closes the range.
	}

	public override IDisposable VisitHyperlinkBegin(Hyperlink link, string? target, IDxpStyleResolver s)
	{
		string? href = target;
		_writer.Write(href != null ? $"<a href=\"{HtmlAttr(href)}\">" : "<a>");
		return Disposable.Create(() => _writer.Write("</a>"));
	}

	public string Escape(string name)
	{
		return name;
	}

	public override IDisposable VisitDeletedRunBegin(DeletedRun dr, IDxpStyleResolver s)
	{
		_writer.Write("<edit-delete>");
		return Disposable.Create(() => {
			_writer.Write("</edit-delete>");
		});
	}

	public override IDisposable VisitInsertedRunBegin(InsertedRun ir, IDxpStyleResolver s)
	{
		_writer.Write("<edit-insert>");
		return Disposable.Create(() => {
			_writer.Write("</edit-insert>");
		});
	}

	public override void VisitDeletedText(DeletedText dt, IDxpStyleResolver s)
	{
		_writer.Write(dt.Text);
	}

	public override void VisitNoBreakHyphen(NoBreakHyphen h, IDxpStyleResolver s)
	{
		if (ShouldSuppressOutput())
			return;
		_writer.Write("-");
	}

	public override void VisitSectionProperties(SectionProperties sp, IDxpStyleResolver s)
	{
		// Example: write a boundary marker
		_writer.Write("\n<!-- SECTION BREAK -->\n");
	}

	static double TwipsToPt(int twips) => twips / 20.0;


	static string BuildParagraphSpanStart(DxpStyleEffectiveIndentTwips ind, int? level)
	{
		var sb = new StringBuilder();
		sb.Append("<p style=\"");

		if (ind.Left is int l)
			sb.Append("margin-left:").Append(TwipsToPt(l).ToString("0.###", CultureInfo.InvariantCulture)).Append("pt;");

		sb.Append("\"");

		if (level != null && level != 0)
		{
			sb.Append($" list-level=\"{level}\"");
		}

		sb.Append(">");
		sb.AppendLine();
		return sb.ToString();
	}


	public static class NbspIndent
	{
		// Typical space width is often ~0.25em–0.33em in proportional fonts.
		// Pick one and be consistent; 0.28 is a decent default for Calibri-ish docs.
		public static string BuildIndentNbsp(
			DxpStyleEffectiveIndentTwips ind,
			bool isFirstLine,
			double fontSizePt,
			double spaceWidthEm = 0.28,
			int maxSpaces = 200)
		{
			if (fontSizePt <= 0)
				fontSizePt = 12; // fallback

			// Compute desired start position in twips
			int left = ind.Left ?? 0;

			int startTwips = left;

			if (ind.Hanging is int h && h != 0)
			{
				// Hanging: first line starts left - hanging; others at left.
				if (isFirstLine)
					startTwips = left - h;
			}
			else if (ind.FirstLine is int f && f != 0)
			{
				// FirstLine: first line starts left + firstLine; others at left.
				if (isFirstLine)
					startTwips = left + f;
			}

			// If negative, don't emit negative spaces.
			if (startTwips <= 0)
				return "";

			// Convert twips to points (1 pt = 20 twips)
			double startPt = startTwips / 20.0;

			// Approximate width of one space in points
			double spacePt = fontSizePt * spaceWidthEm;
			if (spacePt <= 0.01)
				return "";

			int n = (int)Math.Round(startPt / spacePt, MidpointRounding.AwayFromZero);
			n = n < 0 ? 0 : n > maxSpaces ? maxSpaces : n;


			if (n == 0)
				return "";

			var sb = new StringBuilder(n * 6);
			for (int i = 0; i < n; i++)
				sb.Append('\u00A0'); // nbsp
			return sb.ToString();
		}
	}

	public override void VisitDrawingBegin(Drawing d, DxpDrawingInfo? info, IDxpStyleResolver s)
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

	public override void VisitDocumentSettings(Settings settings, IDxpStyleResolver s)
	{
		// no-op for now
	}

	public override void VisitDocumentBackground(object background, IDxpStyleResolver s)
	{
		if (background is DocumentBackground db)
		{
			var color = db.Color?.Value;
			if (!string.IsNullOrEmpty(color) && !string.Equals(color, "auto", StringComparison.OrdinalIgnoreCase))
			{
				_pageBackground = ToCssColor(color!);
				_hasBackgroundColor = true;
			}
		}
	}

	public override void VisitSectionLayout(SectionProperties sp, SectionLayout layout, IDxpStyleResolver s)
	{
		if (!_config.EmitDocumentColors)
			return;

		UpdatePageDimensions(layout);

		_writer.Write("\n<!-- SECTION LAYOUT -->\n");
		CloseBodyContainer();
		EmitBodyContainerStart();
	}

	public override void VisitCoreFileProperties(IPackageProperties core)
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

		foreach (var line in lines)
			_writer.WriteLine(line);

		if (lines.Count > 0)
			_writer.WriteLine();
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


	public override IDisposable VisitParagraphBegin(Paragraph p, IDxpStyleResolver s, string? marker, int? numId, int? iLvl, DxpStyleEffectiveIndentTwips indent)
	{
		string innerText = p.InnerText;

		if (string.IsNullOrWhiteSpace(innerText))
		{
			return Disposable.Create(() => {
				_writer.WriteLine();
			});
		}


		var styleChain = s.GetParagraphStyleChain(p);

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

		var headingLevel = s.GetHeadingLevel(p);
		bool hasNumbering = numId != null;
		bool isHeading = headingLevel != null && !hasNumbering;
		var paragraphStyleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

		// Subtitle -> treat as a sub heading if not already a heading
		if (!isHeading && !hasNumbering && styleChain.Any(sc => string.Equals(sc.StyleId, WordBuiltInStyleId.wdStyleSubtitle, StringComparison.OrdinalIgnoreCase)))
		{
			headingLevel = 2;
			isHeading = true;
		}

		if (!_config.EmitPageNumbers && styleChain.Any(sc => string.Equals(sc.StyleId, WordBuiltInStyleId.wdStylePageNumber, StringComparison.OrdinalIgnoreCase)))
		{
			return Disposable.Empty;
		}

		bool isBlockQuote = styleChain.Any(sc =>
			string.Equals(sc.StyleId, WordBuiltInStyleId.wdStyleQuote, StringComparison.OrdinalIgnoreCase) ||
			string.Equals(sc.StyleId, WordBuiltInStyleId.wdStyleIntenseQuote, StringComparison.OrdinalIgnoreCase) ||
			string.Equals(sc.StyleId, WordBuiltInStyleId.wdStyleBlockQuotation, StringComparison.OrdinalIgnoreCase));

		bool isCaption = styleChain.Any(sc =>
			string.Equals(sc.StyleId, WordBuiltInStyleId.wdStyleCaption, StringComparison.OrdinalIgnoreCase));

		bool isCode = styleChain.Any(sc =>
			string.Equals(sc.StyleId, WordBuiltInStyleId.wdStyleHtmlPre, StringComparison.OrdinalIgnoreCase) ||
			string.Equals(sc.StyleId, WordBuiltInStyleId.wdStyleHtmlCode, StringComparison.OrdinalIgnoreCase) ||
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

		double adjustedMargin = indent.Left.HasValue ? AdjustMarginLeft(TwipsToPt(indent.Left.Value)) : 0.0;
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

		if (marker != null)
		{
			var normalizedMarker = NormalizeMarker(marker);
			if (!needsParagraphWrapper && LooksLikeOrderedListMarker(normalizedMarker))
				normalizedMarker = EscapeOrderedListMarker(normalizedMarker);
			_writer.Write($"""{normalizedMarker} """);
		}
		if (isHeading)
			_headingDepth++;

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

		return Disposable.Create(() => {
			if (isCaption && !_config.UsePlainCodeBlocks && !isHeading && !isBlockQuote)
			{
				_writer.Write("</figcaption>");
			}
			if (isHeading && _headingDepth > 0)
				_headingDepth--;
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

	}

	public override void VisitFootnoteReference(FootnoteReference fr, long id, int index, IDxpStyleResolver s)
	{
		_writer.Write($"<a href=\"#fn-{id}\" id=\"fnref-{id}\">[{index}]</a>");
	}

	public override IDisposable VisitSectionHeaderBegin(Header hdr, object kind, IDxpStyleResolver s)
	{
		// Capture header content; if anything was rendered, wrap it with a bottom border.
		var buffer = new StringWriter();
		var previous = _writer;
		_writer = buffer;
		return Disposable.Create(() =>
		{
			_writer = previous;
			var content = buffer.ToString();
			if (HasVisibleContent(content))
			{
				_writer.Write("<div style=\"border-bottom:1px solid #000;\">\n");
				_writer.Write(content);
				if (!content.EndsWith("\n"))
					_writer.WriteLine();
				_writer.WriteLine("</div>");
			}
		});
	}

	public override IDisposable VisitSectionFooterBegin(Footer ftr, object kind, IDxpStyleResolver s)
	{
		// Capture footer content; if anything was rendered, wrap it with a top border.
		var buffer = new StringWriter();
		var previous = _writer;
		_writer = buffer;
		return Disposable.Create(() =>
		{
			_writer = previous;
			var content = buffer.ToString();
			if (HasVisibleContent(content))
			{
				_writer.Write("<div style=\"border-top:1px solid #000;\">\n");
				_writer.Write(content);
				if (!content.EndsWith("\n"))
					_writer.WriteLine();
				_writer.WriteLine("</div>");
			}
		});
	}

	public override void VisitPageNumber(PageNumber pn, IDxpStyleResolver s)
	{
		// Rendered via field result; we suppress field output when EmitPageNumbers is false.
	}

	public override void VisitComplexFieldBegin(FieldChar begin, IDxpStyleResolver s)
	{
		_fieldSuppressStack.Push(false);
	}

	public override void VisitComplexFieldInstruction(FieldCode instr, string text, IDxpStyleResolver s)
	{
		if (_fieldSuppressStack.Count > 0 && !_config.EmitPageNumbers && LooksLikePageField(text))
		{
			_fieldSuppressStack.Pop();
			_fieldSuppressStack.Push(true);
		}
	}

	public override IDisposable VisitComplexFieldResultBegin(IDxpStyleResolver s)
	{
		bool suppress = !_config.EmitPageNumbers && _fieldSuppressStack.Count > 0 && _fieldSuppressStack.Peek();
		if (suppress)
		{
			_suppressFieldDepth++;
			return Disposable.Create(() => _suppressFieldDepth--);
		}
		return Disposable.Empty;
	}

	public override void VisitComplexFieldEnd(FieldChar end, IDxpStyleResolver s)
	{
		if (_fieldSuppressStack.Count > 0)
			_fieldSuppressStack.Pop();
	}

	public override IDisposable VisitSimpleFieldBegin(SimpleField fld, IDxpStyleResolver s)
	{
		var instr = fld.Instruction?.Value;
		bool suppress = !_config.EmitPageNumbers && LooksLikePageField(instr);
		if (suppress)
		{
			_suppressFieldDepth++;
			return Disposable.Create(() => _suppressFieldDepth--);
		}
		return Disposable.Empty;
	}

	public override IDisposable VisitFootnoteBegin(Footnote fn, long id, int index, IDxpStyleResolver s)
	{
		_writer.Write($"\n<div class=\"footnote\" id=\"fn-{id}\">\n\n");
		return Disposable.Create(() => _writer.Write("</div>\n"));
	}

	public override void VisitFootnoteReferenceMark(FootnoteReferenceMark m, long? footnoteId, int index, IDxpStyleResolver s)
	{
		if (footnoteId != null)
			_writer.Write($"{index}");
	}


	public override IDisposable VisitTableBegin(Table t, DxpTableModel model, IDxpStyleResolver s)
	{
		if (!_config.RichTables)
		{
			_mdTable = new MarkdownTableBuilder();
			return Disposable.Create(() =>
			{
				_mdTable?.Render(_writer);
				_mdTable = null;
			});
		}

		var style = _currentTableStyle;
		_currentTableStyle = null;
		var previousCellBorder = _currentTableCellBorder;

		_writer.Write("<table");
		if (!string.IsNullOrEmpty(style))
			_writer.Write($" style=\"{style}\"");
		_writer.WriteLine(">");
		return Disposable.Create(() => {
			_writer.WriteLine("</table>");
			_currentTableCellBorder = previousCellBorder; // restore for outer table if nested
		});
	}

	public override IDisposable VisitTableRowBegin(TableRow tr, IDxpStyleResolver s)
	{
		var header = tr.TableRowProperties?.GetFirstChild<TableHeader>();
		_currentRowIsHeader = header != null;

		if (_mdTable != null)
		{
			_mdTable.BeginRow(_currentRowIsHeader);
			return Disposable.Create(() => _mdTable.EndRow());
		}

		if (_currentRowIsHeader)
			_writer.WriteLine("  <tr class=\"header-row\">");
		else
			_writer.WriteLine("  <tr>");
		return Disposable.Create(() => _writer.WriteLine("  </tr>"));
	}

	public override void VisitTableCellLayout(TableCell tc, int row, int col, int rowSpan, int colSpan)
	{
		_pendingCellSpan = (rowSpan, colSpan);
	}

	public override IDisposable VisitTableCellBegin(TableCell tc, IDxpStyleResolver s)
	{
		if (_mdTable != null)
		{
			var cellWriter = new StringWriter();
			var previous = _writer;
			_writer = cellWriter;
			return Disposable.Create(() =>
			{
				_writer = previous;
				_mdTable!.AddCell(cellWriter.ToString());
			});
		}

		var spans = _pendingCellSpan;
		_pendingCellSpan = null;
		var cellStyle = _pendingCellStyle;
		_pendingCellStyle = null;

		_writer.Write("    <td");
		if (spans.HasValue)
		{
			if (spans.Value.rowSpan > 1)
				_writer.Write($" rowspan=\"{spans.Value.rowSpan}\"");
			if (spans.Value.colSpan > 1)
				_writer.Write($" colspan=\"{spans.Value.colSpan}\"");
		}
		var effectiveCellStyle = cellStyle ?? (_currentTableCellBorder != null ? $"border:{_currentTableCellBorder};" : null);
		if (!string.IsNullOrEmpty(effectiveCellStyle))
			_writer.Write($" style=\"{effectiveCellStyle}\"");
		_writer.Write(">");

		return Disposable.Create(() => {
			_writer.WriteLine("</td>");
		});
	}

	public override void VisitTableCellProperties(TableCellProperties tcp, IDxpStyleResolver s)
	{
		_pendingCellStyle = _config.EmitTableBorders ? BuildCellStyle(tcp.TableCellBorders) : null;
	}

	public override void VisitParagraphProperties(ParagraphProperties pp, IDxpStyleResolver s)
	{
		// handled in VisitParagraphBegin
	}

	public override void VisitTableProperties(TableProperties tp, IDxpStyleResolver s)
	{
		var styles = _config.EmitTableBorders ? BuildTableStyle(tp) : (null, null);
		_currentTableStyle = styles.tableStyle;
		_currentTableCellBorder = styles.cellBorderStyle;
	}

	public override void VisitTableGrid(TableGrid tg, IDxpStyleResolver s)
	{
		// ignore for "simple HTML"
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

	private bool IsSimpleTable(Table t, DxpTableModel model)
	{
		// No colspan/rowspan
		for (int r = 0; r < model.RowCount; r++)
		{
			for (int c = 0; c < model.ColumnCount; c++)
			{
				var cell = model.Cells[r, c];
				if (cell == null || cell.IsCovered)
					continue;

				if (cell.ColSpan > 1 || cell.RowSpan > 1)
					return false;
			}
		}

		return true;
	}

	private void RenderCapturedTable()
	{
		if (_mdTable == null)
			return;
		_mdTable.Render(_writer);
	}

	private sealed class MarkdownTableBuilder
	{
		private readonly List<Row> _rows = new();
		private Row? _current;

		public void BeginRow(bool isHeader)
		{
			_current = new Row(isHeader);
		}

		public void AddCell(string content)
		{
			_current?.Cells.Add(content);
		}

		public void EndRow()
		{
			if (_current != null)
				_rows.Add(_current);
			_current = null;
		}

		public void Render(TextWriter writer)
		{
			if (_rows.Count == 0)
				return;

			writer.WriteLine();
			int headerIndex = _rows.FindIndex(r => r.IsHeader);
			if (headerIndex < 0)
				headerIndex = 0;

			WriteRow(writer, _rows[headerIndex]);
			WriteSeparator(writer, _rows[headerIndex].Cells.Count);

			for (int i = 0; i < _rows.Count; i++)
			{
				if (i == headerIndex)
					continue;
				WriteRow(writer, _rows[i]);
			}

			writer.WriteLine();
		}

		private static void WriteRow(TextWriter writer, Row row)
		{
			writer.Write("|");
			foreach (var cell in row.Cells)
			{
				var text = NormalizeCell(cell);
				writer.Write(" ");
				writer.Write(text);
				writer.Write(" |");
			}
			writer.WriteLine();
		}

		private static void WriteSeparator(TextWriter writer, int count)
		{
			writer.Write("|");
			for (int i = 0; i < count; i++)
				writer.Write(" --- |");
			writer.WriteLine();
		}

		private static string NormalizeCell(string cell)
		{
			var text = cell.Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ").Trim();
			return text.Replace("|", "\\|");
		}

		private sealed record Row(bool IsHeader)
		{
			public List<string> Cells { get; } = new();
		}
	}

		public override void VisitTableRowProperties(TableRowProperties trp, IDxpStyleResolver s)
	{
		// properties are applied in VisitTableRowBegin
	}

	public override IDisposable VisitBlockBegin(OpenXmlElement child, IDxpStyleResolver s)
	{
		// optional: do nothing
		return Disposable.Empty;
	}

	public void VisitCommentThread(string anchorId, DxpCommentThread thread, IDxpStyleResolver s, Action<DxpCommentInfo>? renderContent)
	{
		if (thread.Comments == null || thread.Comments.Count == 0)
			return;

		if (_config.UsePlainComments)
		{
			EmitPlainCommentThread(thread, renderContent);
			return;
		}

		var commentsStyle = "background:#fff8c6;border:1px solid #e6c44a;border-radius:6px;padding:8px;margin:8px 0 8px 12px;float:right;max-width:45%;";
		var commentStyle = "background:#fffcf0;border:1px solid #d9b200;border-radius:4px;padding:6px;margin-bottom:6px;";

		_writer.Write("<div class=\"comments\"");
		_writer.Write($" style=\"{commentsStyle}\"");
		_writer.WriteLine(">");

		foreach (var c in thread.Comments)
		{
			var label = BuildCommentLabel(c);
			if (!string.IsNullOrEmpty(label))
				_writer.WriteLine("  " + label);

			_writer.Write("  <div class=\"comment\"");
			_writer.Write($" style=\"{commentStyle}\"");
			_writer.WriteLine(">");
			_writer.WriteLine();

			if (renderContent != null)
			{
				renderContent(c);
			}
			else
			{
				_writer.WriteLine(c.Text);
			}

			_writer.WriteLine("  </div>");
			_writer.WriteLine();
		}

		_writer.WriteLine("</div>");
		_writer.WriteLine();
	}

	private void EmitPlainCommentThread(DxpCommentThread thread, Action<DxpCommentInfo>? renderContent)
	{
		_writer.WriteLine("<!--");
		_writer.WriteLine();

		foreach (var c in thread.Comments)
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

			string content = c.Text ?? string.Empty;
			if (renderContent != null)
			{
				var buffer = new StringWriter();
				var prev = _writer;
				_writer = buffer;
				renderContent(c);
				_writer = prev;
				content = buffer.ToString();
			}

			var lines = content.Split('\n');
			foreach (var line in lines)
			{
				if (line.Length == 0)
				{
					_writer.WriteLine();
				}
				else
				{
					_writer.WriteLine("  " + line.TrimEnd());
				}
			}

			_writer.WriteLine();
		}

		_writer.WriteLine("-->");
		_writer.WriteLine();
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

	public override void VisitCommentThread(string anchorId, DxpCommentThread thread, IDxpStyleResolver s)
	{
		VisitCommentThread(anchorId, thread, s, null);
	}

	public override void VisitBreak(Break br, IDxpStyleResolver s)

	{
		if (ShouldSuppressOutput())
			return;
			_writer.Write("<br/>");
	}

	public override void VisitCarriageReturn(CarriageReturn cr, IDxpStyleResolver s)
	{
		if (ShouldSuppressOutput())
			return;
		_writer.Write("<br/>");
	}

	public override void VisitTab(TabChar tab, IDxpStyleResolver s)
	{
		if (ShouldSuppressOutput())
			return;
		_writer.Write("&#9;"); // or &nbsp; spacing
	}

	public override IDisposable VisitRunBegin(Run r, IDxpStyleResolver s)
	{
		var style = BuildRunStyle(r.RunProperties);
		bool hasText = r.ChildElements.OfType<Text>().Any(t => !string.IsNullOrEmpty(t.Text));

		if (string.IsNullOrEmpty(style) || !hasText)
			return Disposable.Empty;

		_writer.Write($"<span style=\"{style}\">");
		return Disposable.Create(() => _writer.Write("</span>"));
	}

	public override void VisitRunProperties(RunProperties rp, IDxpStyleResolver s)
	{
		// handled in VisitRunBegin
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
		if (value == HighlightColorValues.Yellow) return "#ffff00";
		if (value == HighlightColorValues.Green) return "#00ff00";
		if (value == HighlightColorValues.Cyan) return "#00ffff";
		if (value == HighlightColorValues.Magenta) return "#ff00ff";
		if (value == HighlightColorValues.Blue) return "#0000ff";
		if (value == HighlightColorValues.Red) return "#ff0000";
		if (value == HighlightColorValues.DarkBlue) return "#000080";
		if (value == HighlightColorValues.DarkCyan) return "#008080";
		if (value == HighlightColorValues.DarkGreen) return "#008000";
		if (value == HighlightColorValues.DarkMagenta) return "#800080";
		if (value == HighlightColorValues.DarkRed) return "#800000";
		if (value == HighlightColorValues.DarkYellow) return "#808000";
		if (value == HighlightColorValues.LightGray) return "#d3d3d3";
		if (value == HighlightColorValues.DarkGray) return "#a9a9a9";
		if (value == HighlightColorValues.Black) return "#000000";
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
		if (string.IsNullOrEmpty(marker))
			return null;

		if (marker.Length == 1)
		{
			var ch = marker[0];

			if (!string.IsNullOrEmpty(fontFamily))
			{
				var code = (byte)ch;

				if (FontEquals(fontFamily, "Symbol"))
				{
					if (SymbolEncoding.Map.TryGetValue(code, out var symCodepoints) && symCodepoints.Length > 0)
						return char.ConvertFromUtf32(symCodepoints[0]);
				}
				else if (FontEquals(fontFamily, "ZapfDingbats") || FontEquals(fontFamily, "Zapf Dingbats"))
				{
					if (ZapfDingbatsEncoding.Map.TryGetValue(code, out var zapfCodepoints) && zapfCodepoints.Length > 0)
						return char.ConvertFromUtf32(zapfCodepoints[0]);
				}
				else if (FontEquals(fontFamily, "Webdings"))
				{
					var webd = WebdingsEncoding.ToUnicode(code);
					if (!string.IsNullOrEmpty(webd))
						return webd;
				}
				else if (FontEquals(fontFamily, "Wingdings"))
				{
					var wd1 = WingdingsEncoding.ToUnicode(code);
					if (!string.IsNullOrEmpty(wd1))
						return wd1;
				}
				else if (FontEquals(fontFamily, "Wingdings 2"))
				{
					var wd2 = Wingdings2Map.ToUnicode(code);
					if (!string.IsNullOrEmpty(wd2))
						return wd2;
				}
				else if (FontEquals(fontFamily, "Wingdings 3"))
				{
					var wd3 = Wingdings3Encoding.ToUnicode(code);
					if (!string.IsNullOrEmpty(wd3))
						return wd3;
				}
			}

			// Symbol font encoding maps directly via SymbolEncoding
			if (SymbolEncoding.Map.TryGetValue((byte)ch, out var codepoints) && codepoints.Length > 0)
				return char.ConvertFromUtf32(codepoints[0]);
			if (ZapfDingbatsEncoding.Map.TryGetValue((byte)ch, out var zcodepoints) && zcodepoints.Length > 0)
				return char.ConvertFromUtf32(zcodepoints[0]);
		}

		return null;
	}

	public override IDisposable VisitBodyBegin(Body body, IDxpStyleResolver s)
	{
		if (!_config.EmitDocumentColors)
			return Disposable.Empty;

		CaptureNormalStyle(s);

		if (!_hasBackgroundColor)
			_pageBackground = "#ffffff"; // default to white when no document background is set
		EmitBodyContainerStart();

		return Disposable.Create(() =>
		{
			CloseBodyContainer();
		});
	}

	private void CaptureNormalStyle(IDxpStyleResolver s)
	{
		if (_capturedNormal)
			return;

		_capturedNormal = true;

		if (s is DxpStyleResolver resolver)
		{
			var normal = resolver.GetDefaultRunStyle();
			_normalFontName = normal.FontName;
			_normalFontSizeHp = normal.FontSizeHalfPoints;
		}
	}

	private static double TwipsToInches(int twips) => twips / 1440.0;

	private void EmitBodyContainerStart()
	{
		if (_bodyContainerOpen)
			return;

		var style = new StringBuilder("color:#000000;");
		if (_pageWidthInches != null && _pageHeightInches != null)
		{
			style.Append("width:").Append(_pageWidthInches.Value.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");
			style.Append("min-height:").Append(_pageHeightInches.Value.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");
		}
		if (_marginLeftInches != null || _marginRightInches != null || _marginTopInches != null)
		{
			style.Append("box-sizing:border-box;");
			if (_marginTopInches != null)
				style.Append("padding-top:").Append(_marginTopInches.Value.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");
			if (_marginLeftInches != null)
				style.Append("padding-left:").Append(_marginLeftInches.Value.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");
			if (_marginRightInches != null)
				style.Append("padding-right:").Append(_marginRightInches.Value.ToString("0.###", CultureInfo.InvariantCulture)).Append("in;");
		}
		if (!string.IsNullOrEmpty(_pageBackground))
			style.Append("background-color:").Append(_pageBackground).Append(';');
		if (_config.EmitStyleFont)
		{
			if (!string.IsNullOrEmpty(_normalFontName))
				style.Append("font-family:").Append(_normalFontName).Append(';');
			if (_normalFontSizeHp != null)
				style.Append("font-size:").Append((_normalFontSizeHp.Value / 2.0).ToString("0.###", CultureInfo.InvariantCulture)).Append("pt;");
		}

		_writer.Write($"""<div style="{style}">""" + "\n");
		_bodyContainerOpen = true;
	}

	private void CloseBodyContainer()
	{
		if (!_bodyContainerOpen)
			return;
		_writer.Write("</div>\n");
		_bodyContainerOpen = false;
	}

	private void UpdatePageDimensions(SectionLayout layout)
	{
		var pg = layout.PageSize;
		if (pg?.Width != null && pg.Height != null)
		{
			_pageWidthInches = TwipsToInches((int)pg.Width.Value);
			_pageHeightInches = TwipsToInches((int)pg.Height.Value);
		}

		var margin = layout.PageMargin;
		if (margin != null)
		{
			if (margin.Left != null)
			{
				_marginLeftInches = TwipsToInches((int)margin.Left.Value);
				_marginLeftPoints = _marginLeftInches * 72.0;
			}
			if (margin.Right != null)
				_marginRightInches = TwipsToInches((int)margin.Right.Value);
			if (margin.Top != null)
				_marginTopInches = TwipsToInches(margin.Top.Value);
		}
	}

	private double AdjustMarginLeft(double marginPt)
	{
		if (_marginLeftPoints == null)
			return marginPt;
		var adjusted = marginPt - _marginLeftPoints.Value;
		if (adjusted < 0)
			adjusted = 0;
		return adjusted;
	}

	private bool ShouldSuppressOutput() => !_config.EmitPageNumbers && _suppressFieldDepth > 0;

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
		return !string.IsNullOrWhiteSpace(cleaned);
	}

	IDisposable IDxpVisitor.VisitDrawingBegin(Drawing d, DxpDrawingInfo? info, IDxpStyleResolver s)
	{
		VisitDrawingBegin(d, info, s);
		return Disposable.Empty;
	}

	public new void VisitLegacyPictureBegin(Picture pict, IDxpStyleResolver s)
	{
		if (_config.EmitImages == false)
		{
			_writer.Write("[IMAGE]");
			return;
		}

		var alt = "image";
		_writer.Write($"[PICTURE: {alt}]");
	}

	IDisposable IDxpVisitor.VisitLegacyPictureBegin(Picture pict, IDxpStyleResolver s)
	{
		VisitLegacyPictureBegin(pict, s);
		return Disposable.Empty;
	}
}
