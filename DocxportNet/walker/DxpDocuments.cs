using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;

namespace DocxportNet.Walker;

public class DxpDocumentContext : DxpIDocumentContext
{
	public DxpFieldFrameContext CurrentFields { get; }
	public MainDocumentPart? MainDocumentPart { get; internal set; }
	public DxpStyleTracker StyleTracker { get; } = new DxpStyleTracker();
	public DxpComments Comments { get; } = new DxpComments();
	public DxpDrawings Drawings { get; } = new DxpDrawings();
	public DxpTables Tables { get; } = new DxpTables();
	public DxpFootnotes Footnotes { get; } = new DxpFootnotes();
	public DocxEndnotes Endnotes { get; } = new DocxEndnotes();
	public DxpLists Lists { get; } = new DxpLists();
	public HashSet<string> ReferencedBookmarkAnchors { get; } = new HashSet<string>(StringComparer.Ordinal);
	public DxpIStyleResolver Styles { get; }
	public DocumentBackground? Background { get; }
	public DxpStyleEffectiveRunStyle DefaultRunStyle { get; }
	public OpenXmlPart? CurrentPart { get; internal set; }
	public DxpParagraphContext CurrentParagraph { get; internal set; } = DxpParagraphContext.INVALID;
	DxpIParagraphContext DxpIDocumentContext.CurrentParagraph => CurrentParagraph;
	public DxpFootnoteContext CurrentFootnote { get; internal set; } = DxpFootnoteContext.INVALID;
	public DxpSectionContext CurrentSection { get; private set; }
	DxpISectionContext DxpIDocumentContext.CurrentSection => CurrentSection;

	public DxpDocumentContext(WordprocessingDocument doc)
	{
		CurrentFields = new();
		ReferencedBookmarkAnchors = CollectReferencedAnchors(doc);
		Footnotes.Init(doc.MainDocumentPart);
		Endnotes.Init(doc.MainDocumentPart);
		Comments.Init(doc.MainDocumentPart);
		Lists.Init(doc);
		Styles = new DxpStyleResolver(doc);
		DefaultRunStyle = Styles.GetDefaultRunStyle();

		Background = doc.MainDocumentPart.Document?.DocumentBackground;
	}

	public DxpParagraphContext CreateParagraphContext(Paragraph p)
	{
		DxpMarker marker = Lists.MaterializeMarker(p, Styles);
		DxpStyleEffectiveIndentTwips indent = Lists.GetIndentation(p, Styles);
		return new DxpParagraphContext(marker, indent);
	}

	public IDisposable PushParagraph(Paragraph p, out DxpParagraphContext ctx)
	{
		var prev = CurrentParagraph;
		DxpParagraphContext paragraphContext = CreateParagraphContext(p);
		ctx = paragraphContext;
		CurrentParagraph = paragraphContext;
		return Disposable.Create(() => CurrentParagraph = prev);
	}

	public IDisposable PushFootnote(long id, int index, out DxpFootnoteContext ctx)
	{
		var prev = CurrentFootnote;
		DxpFootnoteContext footnoteContext = CreateFootnoteContext(id, index);
		ctx = footnoteContext;
		CurrentFootnote = footnoteContext;
		return Disposable.Create(() => CurrentFootnote = prev);
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
		return Disposable.Create(() => CurrentPart = previous);
	}
}
