using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml.Linq;
using DocxportNet.Walker;

namespace DocxportNet.API;

public sealed record FieldFrame
{
	public bool SeenSeparate;
	public IDisposable? ResultScope;
	public bool InResult;
	public string? InstructionText;
}

public class DxpFieldFrameContext
{
	public readonly Stack<FieldFrame> FieldStack = new();

	public FieldFrame? Current => FieldStack.Count > 0 ? FieldStack.Peek() : null;
}


public sealed class SectionLayout
{
	public PageSize? PageSize;
	public PageMargin? PageMargin;
	public Columns? Columns;
	public DocGrid? DocGrid;
	public PageBorders? PageBorders;
	public LineNumberType? LineNumbers;
	public TextDirection? TextDirection;            // w:textDirection (if used)
	public VerticalTextAlignment? VerticalJustification;
	public FootnoteProperties? FootnoteProperties;
	public EndnoteProperties? EndnoteProperties;
}

public sealed class DxpSectionLayout
{
	public DxpTwipValue? PageWidth { get; set; }
	public DxpTwipValue? PageHeight { get; set; }

	public DxpTwipValue? MarginLeft { get; set; }
	public DxpTwipValue? MarginRight { get; set; }
	public DxpTwipValue? MarginTop { get; set; }
	public DxpTwipValue? MarginBottom { get; set; }
	public DxpTwipValue? MarginHeader { get; set; }
	public DxpTwipValue? MarginFooter { get; set; }
	public DxpTwipValue? MarginGutter { get; set; }

	public int? ColumnCount { get; set; }
	public DxpTwipValue? ColumnSpace { get; set; }
}


public interface DxpIFootnoteContext
{
	public long? Id { get; }
	int? Index { get; }
}

public sealed record DxpMarker(string? marker, int? numId, int? iLvl);

public sealed record DxpLinkAnchor(string? internalRef, string uri);

public sealed record CustomFileProperty(string Name, string? Type, object? Value);

public sealed record DxpChangeInfo(string? Author, DateTime? Date);

public interface DxpITableContext
{
	Table Table { get; }
	TableProperties? Properties { get; }
	TableGrid? Grid { get; }
	DxpComputedTableStyle ComputedStyle { get; }
}

public interface DxpITableRowContext
{
	DxpITableContext Table { get; }
	bool IsHeader { get; }
	int Index { get; }
	TableRowProperties? Properties { get; }
}

public interface DxpITableCellContext
{
	DxpITableRowContext Row { get; }
	int RowIndex { get; }
	int ColumnIndex { get; }
	int RowSpan { get; }
	int ColSpan { get; }
	TableCellProperties? Properties { get; }
	DxpComputedTableCellStyle ComputedStyle { get; }
}

public interface DxpIRubyContext
{
	Ruby Ruby { get; }
	RubyProperties? Properties { get; }
}

public interface DxpISmartTagContext
{
	OpenXmlUnknownElement SmartTag { get; }
	string ElementName { get; }
	string ElementUri { get; }
	IReadOnlyList<CustomXmlAttribute> Attributes { get; }
}

public interface DxpICustomXmlContext
{
	OpenXmlElement Element { get; }
	CustomXmlProperties? Properties { get; }
}

public interface DxpISdtContext
{
	SdtElement Sdt { get; }
	SdtProperties? Properties { get; }
	SdtEndCharProperties? EndCharProperties { get; }
}

public interface DxpIRunContext
{
	Run Run { get; }
	RunProperties? Properties { get; }
	DxpStyleEffectiveRunStyle Style { get; }
	string? Language { get; }
}

public record DxpDocumentProperties(
	IPackageProperties? PackageProperties,
	IReadOnlyList<CustomFileProperty>? CustomFileProperties,
	IReadOnlyList<DxpTimelineEvent>? TimelineEvents
);

public interface DxpIDocumentContext
{
	DxpDocumentProperties DocumentProperties { get; }
	DxpIStyleResolver Styles { get; }
	HashSet<string> ReferencedBookmarkAnchors { get; }
	DxpIParagraphContext CurrentParagraph { get; }
	DxpIRubyContext? CurrentRuby { get; }
	DxpISmartTagContext? CurrentSmartTag { get; }
	DxpICustomXmlContext? CurrentCustomXml { get; }
	DxpISdtContext? CurrentSdt { get; }
	DxpIRunContext? CurrentRun { get; }
	DxpISectionContext CurrentSection { get; }
	DocumentBackground? Background { get; }
	DxpStyleEffectiveRunStyle DefaultRunStyle { get; }
	DxpFieldFrameContext CurrentFields { get; }
	Settings? DocumentSettings { get; }
	IPackageProperties? CoreProperties { get; }
	IReadOnlyList<CustomFileProperty>? CustomProperties { get; }
	bool KeepAccept { get; }
	bool KeepReject { get; }
	DxpChangeInfo CurrentChangeInfo { get; }
}

public interface DxpIParagraphContext
{
	DxpMarker MarkerAccept { get; }
	DxpMarker MarkerReject { get; }
	DxpStyleEffectiveIndentTwips Indent { get; }
	ParagraphProperties? Properties { get; }
}

public interface DxpISectionContext
{
	SectionProperties? SectionProperties { get; }
	DxpSectionLayout? Layout { get; }
	SectionLayout? LayoutRaw { get; }
	bool IsLast { get; }
}

public sealed record DxpFont(string? fontName, int? fontSizeHalfPoints);


public interface DxpIFieldVisitor
{
	// Called when a complex field begins (w:fldChar type="begin")
	void VisitComplexFieldBegin(FieldChar begin, DxpIDocumentContext d);

	// Called for each w:instrText node (FieldCode in SDK); 'text' is the instruction content.
	void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d);

	// Called when the field hits the SEPARATE marker (w:fldChar type="separate")
	void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext d);

	// Called exactly once when entering the "result" portion (after SEPARATE, before END).
	// Return a scope (may be a no-op) that will be disposed when the field ends.
	IDisposable VisitComplexFieldResultBegin(DxpIDocumentContext d);

	// Called for the cached field result text while inside the result portion.
	void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d);

	// Called when the field ends (w:fldChar type="end")
	void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d);
}


public interface DxpIStyleVisitor
{
	void StyleBoldBegin(DxpIDocumentContext d);
	void StyleBoldEnd(DxpIDocumentContext d);

	void StyleItalicBegin(DxpIDocumentContext d);
	void StyleItalicEnd(DxpIDocumentContext d);

	void StyleUnderlineBegin(DxpIDocumentContext d);
	void StyleUnderlineEnd(DxpIDocumentContext d);

	void StyleStrikeBegin(DxpIDocumentContext d);
	void StyleStrikeEnd(DxpIDocumentContext d);

	void StyleDoubleStrikeBegin(DxpIDocumentContext d);
	void StyleDoubleStrikeEnd(DxpIDocumentContext d);

	void StyleSuperscriptBegin(DxpIDocumentContext d);
	void StyleSuperscriptEnd(DxpIDocumentContext d);

	void StyleSubscriptBegin(DxpIDocumentContext d);
	void StyleSubscriptEnd(DxpIDocumentContext d);

	void StyleSmallCapsBegin(DxpIDocumentContext d);
	void StyleSmallCapsEnd(DxpIDocumentContext d);

	void StyleAllCapsBegin(DxpIDocumentContext d);
	void StyleAllCapsEnd(DxpIDocumentContext d);

	void StyleFontBegin(DxpFont font, DxpIDocumentContext d);
	void StyleFontEnd(DxpIDocumentContext d);
}

public interface DxpIVisitor : DxpIStyleVisitor, DxpIFieldVisitor
{
	// Assign a binary sink (e.g., a Stream) when the visitor produces bytes.
	void SetOutput(Stream stream);

	bool AcceptAlternateContentChoice(AlternateContentChoice choice, IReadOnlyList<string> required, DxpIDocumentContext d);
	IDisposable VisitAlternateContentBegin(AlternateContent ac, DxpIDocumentContext d);
	void VisitAltChunk(AltChunk ac, DxpIDocumentContext d);
	void VisitAnnotationReference(AnnotationReferenceMark arm, DxpIDocumentContext d);
	IDisposable VisitBidirectionalEmbeddingBegin(BidirectionalEmbedding bdi, DxpIDocumentContext d);
	IDisposable VisitBidirectionalOverrideBegin(BidirectionalOverride bdo, DxpIDocumentContext d);
	void VisitBibliographySources(CustomXmlPart bibliographyPart, XDocument bib, DxpIDocumentContext d);
	IDisposable VisitBlockBegin(OpenXmlElement child, DxpIDocumentContext d);
	void VisitBookmarkEnd(BookmarkEnd be, DxpIDocumentContext d);
	void VisitBookmarkStart(BookmarkStart bs, DxpIDocumentContext d);
	void VisitBreak(Break br, DxpIDocumentContext d);
	void VisitCarriageReturn(CarriageReturn cr, DxpIDocumentContext d);
	IDisposable VisitCommentBegin(DxpCommentInfo c, DxpCommentThread thread, DxpIDocumentContext d);
	IDisposable VisitCommentThreadBegin(string anchorId, DxpCommentThread thread, DxpIDocumentContext d);
	void VisitConflictDeletion(ConflictDeletion cDel, DxpIDocumentContext d);
	void VisitConflictInsertion(ConflictInsertion cIns, DxpIDocumentContext d);
	void VisitContentPart(DocumentFormat.OpenXml.Wordprocessing.ContentPart cp, DxpIDocumentContext d);
	void VisitContinuationSeparatorMark(ContinuationSeparatorMark csep, DxpIDocumentContext d);
	IDisposable VisitCustomXmlBlockBegin(CustomXmlBlock cx, DxpIDocumentContext d);
	IDisposable VisitCustomXmlCellBegin(CustomXmlCell cxCell, DxpIDocumentContext d);
	void VisitCustomXmlConflictDeletionRangeEnd(CustomXmlConflictDeletionRangeEnd cxCde, DxpIDocumentContext d);
	void VisitCustomXmlConflictDeletionRangeStart(CustomXmlConflictDeletionRangeStart cxCds, DxpIDocumentContext d);
	void VisitCustomXmlConflictInsertionRangeEnd(CustomXmlConflictInsertionRangeEnd cxCie, DxpIDocumentContext d);
	void VisitCustomXmlConflictInsertionRangeStart(CustomXmlConflictInsertionRangeStart cxCis, DxpIDocumentContext d);
	void VisitCustomXmlDelRangeEnd(CustomXmlDelRangeEnd cdle, DxpIDocumentContext d);
	void VisitCustomXmlDelRangeStart(CustomXmlDelRangeStart cdls, DxpIDocumentContext d);
	void VisitCustomXmlInsRangeEnd(CustomXmlInsRangeEnd cine, DxpIDocumentContext d);
	void VisitCustomXmlInsRangeStart(CustomXmlInsRangeStart cins, DxpIDocumentContext d);
	void VisitCustomXmlMoveFromRangeEnd(CustomXmlMoveFromRangeEnd cmfe, DxpIDocumentContext d);
	void VisitCustomXmlMoveFromRangeStart(CustomXmlMoveFromRangeStart cmfs, DxpIDocumentContext d);
	void VisitCustomXmlMoveToRangeEnd(CustomXmlMoveToRangeEnd cmte, DxpIDocumentContext d);
	void VisitCustomXmlMoveToRangeStart(CustomXmlMoveToRangeStart cmts, DxpIDocumentContext d);
	IDisposable VisitCustomXmlRowBegin(CustomXmlRow cxRow, DxpIDocumentContext d);
	IDisposable VisitCustomXmlRunBegin(CustomXmlRun cxr, DxpIDocumentContext d);
	void VisitDayLong(DayLong dl, DxpIDocumentContext d);
	void VisitDayShort(DayShort ds, DxpIDocumentContext d);
	IDisposable VisitDeletedBegin(Deleted del, DxpIDocumentContext d);
	void VisitDeletedFieldCode(DeletedFieldCode dfc, DxpIDocumentContext d);
	void VisitDeletedParagraphMark(Deleted del, ParagraphProperties pPr, Paragraph? p, DxpIDocumentContext d);
	IDisposable VisitDeletedRunBegin(DeletedRun dr, DxpIDocumentContext d);
	void VisitDeletedTableRowMark(Deleted del, TableRowProperties trPr, TableRow? tr, DxpIDocumentContext d);
	void VisitDeletedText(DeletedText dt, DxpIDocumentContext d);
	IDisposable VisitDocumentBodyBegin(Body body, DxpIDocumentContext d);
	IDisposable VisitDocumentBegin(WordprocessingDocument doc, DxpIDocumentContext documentContext);
	IDisposable VisitDrawingBegin(Drawing drw, DxpDrawingInfo? info, DxpIDocumentContext d);
	void VisitEmbeddedObject(EmbeddedObject obj, DxpIDocumentContext d);
	IDisposable VisitEndnoteBegin(Endnote item1, long item3, int item2, DxpIDocumentContext d);
	void VisitEndnoteReference(EndnoteReference enr, DxpIDocumentContext d);
	void VisitEndnoteReferenceMark(EndnoteReferenceMark erm, DxpIDocumentContext d);
	void VisitFieldData(FieldData data, DxpIDocumentContext d);
	IDisposable VisitFootnoteBegin(Footnote fn, DxpIFootnoteContext footnote, DxpIDocumentContext d);
	void VisitFootnoteReference(FootnoteReference fr, DxpIFootnoteContext footnote, DxpIDocumentContext d);
	void VisitFootnoteReferenceMark(FootnoteReferenceMark m, DxpIFootnoteContext footnote, DxpIDocumentContext d);
	IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target, DxpIDocumentContext d);
	IDisposable VisitInsertedBegin(Inserted ins, DxpIDocumentContext d);
	void VisitInsertedNumbering(Inserted ins, DxpMarker? marker, DxpStyleEffectiveIndentTwips indent, Paragraph? p, DxpIDocumentContext d);
	void VisitInsertedParagraphMark(Inserted ins, ParagraphProperties pPr2, Paragraph? p, DxpIDocumentContext d);
	IDisposable VisitInsertedRunBegin(InsertedRun ir, DxpIDocumentContext d);
	void VisitInsertedTableRowMark(Inserted ins, TableRowProperties trPr, TableRow? tr, DxpIDocumentContext d);
	void VisitLastRenderedPageBreak(LastRenderedPageBreak pb, DxpIDocumentContext d);
	IDisposable VisitLegacyPictureBegin(Picture pict, DxpIDocumentContext d);
	void VisitMonthLong(MonthLong ml, DxpIDocumentContext d);
	void VisitMonthShort(MonthShort ms, DxpIDocumentContext d);
	void VisitMoveFromRangeEnd(MoveFromRangeEnd mfre, DxpIDocumentContext d);
	void VisitMoveFromRangeStart(MoveFromRangeStart mfrs, DxpIDocumentContext d);
	void VisitMoveFromRun(MoveFromRun mfr, DxpIDocumentContext d);
	void VisitMoveToRangeEnd(MoveToRangeEnd mtre, DxpIDocumentContext d);
	void VisitMoveToRangeStart(MoveToRangeStart mtrs, DxpIDocumentContext d);
	void VisitMoveToRun(MoveToRun mtr, DxpIDocumentContext d);
	void VisitNoBreakHyphen(NoBreakHyphen h, DxpIDocumentContext d);
	void VisitOMath(DocumentFormat.OpenXml.Math.OfficeMath oMath, DxpIDocumentContext d);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Accent mAccent, DxpIDocumentContext d);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Bar mBar, DxpIDocumentContext d);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.BorderBox mBorderBox, DxpIDocumentContext d);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Box mBox, DxpIDocumentContext d);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Delimiter mDelim, DxpIDocumentContext d);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.EquationArray mEqArr, DxpIDocumentContext d);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Fraction mFrac, DxpIDocumentContext d);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.GroupChar mGroupChr, DxpIDocumentContext d);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.LimitLower mLimLow, DxpIDocumentContext d);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.LimitUpper mLimUpp, DxpIDocumentContext d);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.MathFunction mFunc, DxpIDocumentContext d);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Matrix mMat, DxpIDocumentContext d);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Nary mNary, DxpIDocumentContext d);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Phantom mPhant, DxpIDocumentContext d);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.PreSubSuper mPreSubSup, DxpIDocumentContext d);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Radical mRad, DxpIDocumentContext d);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Subscript mSub, DxpIDocumentContext d);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.SubSuperscript mSubSup, DxpIDocumentContext d);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Superscript mSup, DxpIDocumentContext d);
	void VisitOMathParagraph(DocumentFormat.OpenXml.Math.Paragraph oMathPara, DxpIDocumentContext d);
	void VisitOMathRun(DocumentFormat.OpenXml.Math.Run mMathRun, DxpIDocumentContext d);
	void VisitPageNumber(PageNumber pn, DxpIDocumentContext d);
	IDisposable VisitParagraphBegin(Paragraph p, DxpIDocumentContext d, DxpIParagraphContext paragraph);
	void VisitPermEnd(PermEnd pe2, DxpIDocumentContext d);
	void VisitPermStart(PermStart ps, DxpIDocumentContext d);
	void VisitPositionalTab(PositionalTab ptab, DxpIDocumentContext d);
	void VisitProofError(ProofError pe, DxpIDocumentContext d);
	IDisposable VisitRubyBegin(Ruby ruby, DxpIDocumentContext d);
	IDisposable VisitRubyContentBegin(RubyContentType rc, bool isBase, DxpIDocumentContext d);
	IDisposable VisitRunBegin(Run r, DxpIDocumentContext d);
	IDisposable VisitSectionBegin(SectionProperties properties, SectionLayout layout, DxpIDocumentContext d);
	IDisposable VisitSectionBodyBegin(SectionProperties properties, DxpIDocumentContext d);
	IDisposable VisitSectionFooterBegin(Footer ftr, object value, DxpIDocumentContext d);
	IDisposable VisitSectionHeaderBegin(Header hdr, object value, DxpIDocumentContext d);
	void VisitSeparatorMark(SeparatorMark sep, DxpIDocumentContext d);
	IDisposable VisitSdtBlockBegin(SdtBlock sdt, DxpIDocumentContext d);
	IDisposable VisitSdtCellBegin(SdtCell sdtCell, DxpIDocumentContext d);
	IDisposable VisitSdtContentBlockBegin(SdtContentBlock content, DxpIDocumentContext d);
	IDisposable VisitSdtContentRunBegin(SdtContentRun content, DxpIDocumentContext d);
	IDisposable VisitSdtRowBegin(SdtRow sdtRow, DxpIDocumentContext d);
	IDisposable VisitSdtRunBegin(SdtRun sdtRun, DxpIDocumentContext d);
	IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d);
	void VisitSoftHyphen(SoftHyphen sh, DxpIDocumentContext d);
	IDisposable VisitSmartTagRunBegin(OpenXmlUnknownElement smart, string elementName, string elementUri, DxpIDocumentContext d);
	void VisitSubDocumentReference(SubDocumentReference subDoc, DxpIDocumentContext d);
	void VisitSymbol(SymbolChar sym, DxpIDocumentContext d);
	void VisitTab(TabChar tab, DxpIDocumentContext d);
	IDisposable VisitTableBegin(Table t, DxpTableModel model, DxpIDocumentContext d, DxpITableContext table);
	IDisposable VisitTableCellBegin(TableCell tc, DxpITableCellContext cell, DxpIDocumentContext d);
	IDisposable VisitTableRowBegin(TableRow tr, DxpITableRowContext row, DxpIDocumentContext d);
	void VisitText(Text t, DxpIDocumentContext d);
	IDisposable VisitTextBoxContentBegin(TextBoxContent txbx, DxpIDocumentContext d);
	void VisitUnknown(string context, OpenXmlElement el, DxpIDocumentContext d);
	void VisitYearLong(YearLong yl, DxpIDocumentContext d);
	void VisitYearShort(YearShort ys, DxpIDocumentContext d);
}

public interface DxpITextVisitor : DxpIVisitor
{
	// Assign a text sink; common for Markdown/HTML/plain text visitors.
	void SetOutput(TextWriter writer);
}
