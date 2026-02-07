using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;

namespace DocxportNet.Walker;

public class DxpDocumentContext : DxpIDocumentContext
{
	private sealed record DxpEditState(bool KeepAccept, bool KeepReject, DxpChangeInfo ChangeInfo);

	public DxpFieldFrameContext CurrentFields { get; }
	public MainDocumentPart? MainDocumentPart { get; internal set; }
	internal DxpNumberingResolver NumberingResolver { get; }
	public DxpStyleTracker StyleTracker { get; } = new DxpStyleTracker();
	public DxpComments Comments { get; } = new DxpComments();
	public DxpDrawings Drawings { get; } = new DxpDrawings();
	public DxpTables Tables { get; } = new DxpTables();
	internal DxpTableStyleResolver TableStyleResolver { get; }
	public DxpFootnotes Footnotes { get; } = new DxpFootnotes();
	public DocxEndnotes Endnotes { get; } = new DocxEndnotes();
	public DxpLists AcceptLists { get; } = new DxpLists();
	public DxpLists RejectLists { get; } = new DxpLists();
	public HashSet<string> ReferencedBookmarkAnchors { get; } = new HashSet<string>(StringComparer.Ordinal);
	public DxpIStyleResolver Styles { get; }
	public DocumentBackground? Background { get; }
	public DxpStyleEffectiveRunStyle DefaultRunStyle { get; }
	public Settings? DocumentSettings { get; internal set; }
	public IPackageProperties? CoreProperties { get; internal set; }
	public IReadOnlyList<CustomFileProperty>? CustomProperties { get; internal set; }
	public OpenXmlPart? CurrentPart { get; internal set; }
	private readonly DxpEditState _defaultEditState;
	private readonly Stack<DxpEditState> _editStateStack = new();
	public bool KeepAccept => (_editStateStack.Count == 0 ? _defaultEditState : _editStateStack.Peek()).KeepAccept;
	public bool KeepReject => (_editStateStack.Count == 0 ? _defaultEditState : _editStateStack.Peek()).KeepReject;
	public DxpChangeInfo CurrentChangeInfo => (_editStateStack.Count == 0 ? _defaultEditState : _editStateStack.Peek()).ChangeInfo;
	public DxpParagraphContext CurrentParagraph { get; internal set; } = DxpParagraphContext.INVALID;
	DxpIParagraphContext DxpIDocumentContext.CurrentParagraph => CurrentParagraph;
	public DxpRubyContext? CurrentRuby { get; internal set; }
	DxpIRubyContext? DxpIDocumentContext.CurrentRuby => CurrentRuby;
	public DxpSmartTagContext? CurrentSmartTag { get; internal set; }
	DxpISmartTagContext? DxpIDocumentContext.CurrentSmartTag => CurrentSmartTag;
	public DxpCustomXmlContext? CurrentCustomXml { get; internal set; }
	DxpICustomXmlContext? DxpIDocumentContext.CurrentCustomXml => CurrentCustomXml;
	public DxpSdtContext? CurrentSdt { get; internal set; }
	DxpISdtContext? DxpIDocumentContext.CurrentSdt => CurrentSdt;
	public DxpRunContext? CurrentRun { get; internal set; }
	DxpIRunContext? DxpIDocumentContext.CurrentRun => CurrentRun;
	public DxpITableContext? CurrentTable { get; internal set; }
	DxpITableContext? DxpIDocumentContext.CurrentTable => CurrentTable;
	public DxpITableRowContext? CurrentTableRow { get; internal set; }
	DxpITableRowContext? DxpIDocumentContext.CurrentTableRow => CurrentTableRow;
	public DxpITableCellContext? CurrentTableCell { get; internal set; }
	DxpITableCellContext? DxpIDocumentContext.CurrentTableCell => CurrentTableCell;
	public DxpTableModel? CurrentTableModel { get; internal set; }
	DxpTableModel? DxpIDocumentContext.CurrentTableModel => CurrentTableModel;
	public DxpFootnoteContext CurrentFootnote { get; internal set; } = DxpFootnoteContext.INVALID;
	public DxpSectionContext CurrentSection { get; private set; } = DxpSectionContext.INVALID;
	DxpISectionContext DxpIDocumentContext.CurrentSection => CurrentSection;

	public DxpDocumentProperties DocumentProperties { get; internal set; }
	public DxpDocumentIndex DocumentIndex { get; }

	public DxpDocumentContext(WordprocessingDocument doc)
	{
		CurrentFields = new();
		NumberingResolver = new DxpNumberingResolver(doc);
		TableStyleResolver = new DxpTableStyleResolver(doc);
		ReferencedBookmarkAnchors = CollectReferencedAnchors(doc);
		var mainPart = doc.MainDocumentPart;
		if (mainPart != null)
		{
			MainDocumentPart = mainPart;
			Footnotes.Init(mainPart);
			Endnotes.Init(mainPart);
			Comments.Init(mainPart);
			Background = mainPart.Document?.DocumentBackground;
		}
		else
		{
			Comments.Init(null);
		}
		AcceptLists.Init(doc);
		RejectLists.Init(doc);
		Styles = new DxpStyleResolver(doc);
		DefaultRunStyle = Styles.GetDefaultRunStyle();
		DocumentProperties = new(null, null, null);
		DocumentIndex = DxpDocumentIndexBuilder.Build(doc, (DxpStyleResolver)Styles);

		CoreProperties = doc.PackageProperties;
		_defaultEditState = new DxpEditState(
			KeepAccept: true,
			KeepReject: true,
			ChangeInfo: new DxpChangeInfo(CoreProperties?.LastModifiedBy, CoreProperties?.Modified));
	}

	public DxpParagraphContext CreateParagraphContext(Paragraph p, bool advanceAccept = true, bool advanceReject = true)
	{
		DxpMarker acceptMarker = advanceAccept ? AcceptLists.MaterializeMarker(p, Styles) : new DxpMarker(null, null, null);
		DxpMarker rejectMarker = advanceReject ? RejectLists.MaterializeMarker(p, Styles) : new DxpMarker(null, null, null);
		DxpStyleEffectiveIndentTwips indent = AcceptLists.GetIndentation(p, Styles);
		var computed = DxpParagraphStyleComputer.ComputeParagraphStyle(p, indent, this);

		// Word treats consecutive paragraphs with identical borders as a single bordered block:
		// - top border shows only on the first paragraph in the block
		// - bottom border shows only on the last paragraph in the block
		// This avoids "double" borders between adjacent paragraphs.
		if (computed.Borders != null && Styles is DxpStyleResolver resolver)
		{
			var prev = p.PreviousSibling<Paragraph>();
			var next = p.NextSibling<Paragraph>();

			var prevBorders = prev != null ? DxpParagraphStyleComputer.ComputeBorders(resolver.GetParagraphBorders(prev)) : null;
			var nextBorders = next != null ? DxpParagraphStyleComputer.ComputeBorders(resolver.GetParagraphBorders(next)) : null;

			bool suppressTop = computed.Borders.Top != null && prevBorders?.Top != null && computed.Borders.Top == prevBorders.Top;
			bool suppressBottom = computed.Borders.Bottom != null && nextBorders?.Bottom != null && computed.Borders.Bottom == nextBorders.Bottom;

			if (suppressTop || suppressBottom)
			{
				computed = computed with
				{
					Borders = new DxpComputedBoxBorders(
						Top: suppressTop ? null : computed.Borders.Top,
						Right: computed.Borders.Right,
						Bottom: suppressBottom ? null : computed.Borders.Bottom,
						Left: computed.Borders.Left
					)
				};
			}
		}

		// Word contextualSpacing: don't add space between paragraphs of the same style.
		// Approximation: when contextualSpacing is enabled on either paragraph, suppress the shared boundary by
		// removing the bottom margin on the first and top margin on the second.
		if (Styles is DxpStyleResolver styleResolver)
		{
			var prev = p.PreviousSibling<Paragraph>();
			var next = p.NextSibling<Paragraph>();
			var styleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

			if (!string.IsNullOrEmpty(styleId))
			{
				var contextual = styleResolver.GetContextualSpacing(p);

				if (prev != null)
				{
					var prevStyleId = prev.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
					if (string.Equals(styleId, prevStyleId, StringComparison.Ordinal))
					{
						var prevCtx = styleResolver.GetContextualSpacing(prev);
						if (contextual || prevCtx)
							computed = computed with { MarginTopPt = 0.0 };
					}
				}

				if (next != null)
				{
					var nextStyleId = next.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
					if (string.Equals(styleId, nextStyleId, StringComparison.Ordinal))
					{
						var nextCtx = styleResolver.GetContextualSpacing(next);
						if (contextual || nextCtx)
							computed = computed with { MarginBottomPt = 0.0 };
					}
				}
			}
		}

		var layout = DxpParagraphLayoutComputer.ComputeLayout(p, this);
		return new DxpParagraphContext(acceptMarker, rejectMarker, indent, p.ParagraphProperties, computed, layout);
	}

	public IDisposable PushParagraph(Paragraph p, out DxpParagraphContext ctx, bool advanceAccept = true, bool advanceReject = true)
	{
		var prev = CurrentParagraph;
		DxpParagraphContext paragraphContext = CreateParagraphContext(p, advanceAccept, advanceReject);
		ctx = paragraphContext;
		CurrentParagraph = paragraphContext;
		return DxpDisposable.Create(() => CurrentParagraph = prev);
	}

	public IDisposable PushRun(Run r, DxpStyleEffectiveRunStyle style, string? language, out DxpRunContext ctx)
	{
		var prev = CurrentRun;
		var runCtx = new DxpRunContext(r, r.RunProperties, style, language);
		ctx = runCtx;
		CurrentRun = runCtx;
		return DxpDisposable.Create(() => CurrentRun = prev);
	}

	public IDisposable PushRuby(Ruby ruby, RubyProperties? properties, out DxpRubyContext ctx)
	{
		var prev = CurrentRuby;
		var rubyCtx = new DxpRubyContext(ruby, properties);
		ctx = rubyCtx;
		CurrentRuby = rubyCtx;
		return DxpDisposable.Create(() => CurrentRuby = prev);
	}

	public IDisposable PushSmartTag(OpenXmlUnknownElement smart, string elementName, string elementUri, IReadOnlyList<CustomXmlAttribute> attrs, out DxpSmartTagContext ctx)
	{
		var prev = CurrentSmartTag;
		var smartCtx = new DxpSmartTagContext(smart, elementName, elementUri, attrs);
		ctx = smartCtx;
		CurrentSmartTag = smartCtx;
		return DxpDisposable.Create(() => CurrentSmartTag = prev);
	}

	public IDisposable PushSdt(SdtElement sdt, SdtProperties? properties, SdtEndCharProperties? endCharProperties, out DxpSdtContext ctx)
	{
		var prev = CurrentSdt;
		var sdtCtx = new DxpSdtContext(sdt, properties, endCharProperties);
		ctx = sdtCtx;
		CurrentSdt = sdtCtx;
		return DxpDisposable.Create(() => CurrentSdt = prev);
	}

	public IDisposable PushCustomXml(OpenXmlElement element, CustomXmlProperties? properties, out DxpCustomXmlContext ctx)
	{
		var prev = CurrentCustomXml;
		var cCtx = new DxpCustomXmlContext(element, properties);
		ctx = cCtx;
		CurrentCustomXml = cCtx;
		return DxpDisposable.Create(() => CurrentCustomXml = prev);
	}

	public IDisposable PushFootnote(long id, int index, out DxpFootnoteContext ctx)
	{
		var prev = CurrentFootnote;
		DxpFootnoteContext footnoteContext = CreateFootnoteContext(id, index);
		ctx = footnoteContext;
		CurrentFootnote = footnoteContext;
		return DxpDisposable.Create(() => CurrentFootnote = prev);
	}

	private DxpFootnoteContext CreateFootnoteContext(long id, int index)
	{
		return new DxpFootnoteContext(id, index);
	}

	public DxpSectionContext EnterSection(SectionProperties sp, SectionLayout layout)
	{
		DxpSectionContext ctx = new DxpSectionContext(sp, layout, BuildDxpSectionLayout(layout));
		CurrentSection = ctx;
		return ctx;
	}

	internal static DxpSectionLayout BuildDxpSectionLayout(SectionLayout layout)
	{
		var result = new DxpSectionLayout();

		var pg = layout.PageSize;
		if (pg?.Width != null && pg.Height != null)
		{
			result.PageWidth = new DxpTwipValue((int)pg.Width.Value);
			result.PageHeight = new DxpTwipValue((int)pg.Height.Value);
		}

		var margin = layout.PageMargin;
		if (margin != null)
		{
			if (margin.Left != null)
				result.MarginLeft = new DxpTwipValue((int)margin.Left.Value);
			if (margin.Right != null)
				result.MarginRight = new DxpTwipValue((int)margin.Right.Value);
			if (margin.Top != null)
				result.MarginTop = new DxpTwipValue((int)margin.Top.Value);
			if (margin.Bottom != null)
				result.MarginBottom = new DxpTwipValue((int)margin.Bottom.Value);
			if (margin.Header != null)
				result.MarginHeader = new DxpTwipValue((int)margin.Header.Value);
			if (margin.Footer != null)
				result.MarginFooter = new DxpTwipValue((int)margin.Footer.Value);
			if (margin.Gutter != null)
				result.MarginGutter = new DxpTwipValue((int)margin.Gutter.Value);
		}

		var cols = layout.Columns;
		if (cols != null)
		{
			IList<OpenXmlAttribute> attrs = cols.GetAttributes();
			OpenXmlAttribute numAttr = attrs.FirstOrDefault(a => string.Equals(a.LocalName, "num", StringComparison.OrdinalIgnoreCase));
			if (!string.IsNullOrEmpty(numAttr.Value) && int.TryParse(numAttr.Value, out int numCols))
				result.ColumnCount = numCols;

			OpenXmlAttribute spaceAttr = attrs.FirstOrDefault(a => string.Equals(a.LocalName, "space", StringComparison.OrdinalIgnoreCase));
			if (!string.IsNullOrEmpty(spaceAttr.Value) && int.TryParse(spaceAttr.Value, out int spaceTwips))
				result.ColumnSpace = new DxpTwipValue(spaceTwips);
		}

		return result;
	}

	private static HashSet<string> CollectReferencedAnchors(WordprocessingDocument doc)
	{
		static IEnumerable<OpenXmlPartRootElement> Roots(MainDocumentPart main)
		{
			if (main.Document is not null)
				yield return main.Document;

			foreach (var h in main.HeaderParts)
				if (h.Header is not null)
					yield return h.Header;

			foreach (var f in main.FooterParts)
				if (f.Footer is not null)
					yield return f.Footer;

			if (main.FootnotesPart?.Footnotes is not null)
				yield return main.FootnotesPart.Footnotes;
			if (main.EndnotesPart?.Endnotes is not null)
				yield return main.EndnotesPart.Endnotes;
		}

		var main = doc.MainDocumentPart;
		var set = new HashSet<string>(StringComparer.Ordinal);
		if (main is null)
			return set;

		foreach (var root in Roots(main))
		{
			foreach (var link in root.Descendants<Hyperlink>())
			{
				var a = link.Anchor?.Value;
				if (!string.IsNullOrEmpty(a))
					set.Add(a!);
			}
		}

		return set;
	}

	public IDisposable PushCurrentPart(OpenXmlPart? part)
	{
		var previous = CurrentPart;
		if (part != null)
			CurrentPart = part;
		return DxpDisposable.Create(() => CurrentPart = previous);
	}

	public IDisposable PushChangeScope(bool keepAccept, bool keepReject, DxpChangeInfo changeInfo)
	{
		_editStateStack.Push(new DxpEditState(keepAccept, keepReject, changeInfo));
		return DxpDisposable.Create(() => _editStateStack.Pop());
	}
}
