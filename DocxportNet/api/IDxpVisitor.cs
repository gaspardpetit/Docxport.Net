using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using l3ia.lapi.services.documents.docx.convert;
using static l3ia.lapi.services.documents.docx.convert.DxpWalker;

namespace DocxportNet.api;

public interface IDxpFieldVisitor
{
	// Called when a complex field begins (w:fldChar type="begin")
	void VisitComplexFieldBegin(FieldChar begin, IDxpStyleResolver s);

	// Called for each w:instrText node (FieldCode in SDK); 'text' is the instruction content.
	void VisitComplexFieldInstruction(FieldCode instr, string text, IDxpStyleResolver s);

	// Called when the field hits the SEPARATE marker (w:fldChar type="separate")
	void VisitComplexFieldSeparate(FieldChar separate, IDxpStyleResolver s);

	// Called exactly once when entering the "result" portion (after SEPARATE, before END).
	// Return a scope (may be a no-op) that will be disposed when the field ends.
	IDisposable VisitComplexFieldResultBegin(IDxpStyleResolver s);

	// Called when the field ends (w:fldChar type="end")
	void VisitComplexFieldEnd(FieldChar end, IDxpStyleResolver s);
}


public interface IDxpStyleVisitor
{
	void StyleBoldBegin();
	void StyleBoldEnd();

	void StyleItalicBegin();
	void StyleItalicEnd();

	void StyleUnderlineBegin();
	void StyleUnderlineEnd();

	void StyleStrikeBegin();
	void StyleStrikeEnd();

	void StyleDoubleStrikeBegin();
	void StyleDoubleStrikeEnd();

	void StyleSuperscriptBegin();
	void StyleSuperscriptEnd();

	void StyleSubscriptBegin();
	void StyleSubscriptEnd();

	void StyleSmallCapsBegin();
	void StyleSmallCapsEnd();

	void StyleAllCapsBegin();
	void StyleAllCapsEnd();

	void StyleFontBegin(string? fontName, int? fontSizeHalfPoints);
	void StyleFontEnd();
}




public interface IDxpVisitor : IDxpStyleVisitor, IDxpFieldVisitor
{
	void VisitTableCellLayout(TableCell tc, int row, int col, int rowSpan, int colSpan);
	void VisitParagraphProperties(ParagraphProperties pp, IDxpStyleResolver s);
	void VisitBookmarkStart(BookmarkStart bs, IDxpStyleResolver s);
	void VisitBookmarkEnd(BookmarkEnd be, IDxpStyleResolver s);
	IDisposable VisitRunBegin(Run r, IDxpStyleResolver s);
	IDisposable VisitHyperlinkBegin(Hyperlink link, string? target, IDxpStyleResolver s);
	IDisposable VisitParagraphBegin(Paragraph p, IDxpStyleResolver s, string? marker, int? numId, int? iLvl, DxpStyleEffectiveIndentTwips indent);
	IDisposable VisitTableBegin(Table t, DxpTableModel model, IDxpStyleResolver s);
	IDisposable VisitTableRowBegin(TableRow tr, IDxpStyleResolver s);
	IDisposable VisitTableCellBegin(TableCell tc, IDxpStyleResolver s);
	void VisitTableCellProperties(TableCellProperties tcp, IDxpStyleResolver s);
	void VisitTableProperties(TableProperties tp, IDxpStyleResolver s);
	void VisitTableGrid(TableGrid tg, IDxpStyleResolver s);
	IDisposable VisitDeletedRunBegin(DeletedRun dr, IDxpStyleResolver s);
	IDisposable VisitInsertedRunBegin(InsertedRun ir, IDxpStyleResolver s);
	void VisitLastRenderedPageBreak(LastRenderedPageBreak pb, IDxpStyleResolver s);
	void VisitRunProperties(RunProperties rp, IDxpStyleResolver s);
	void VisitDeletedText(DeletedText dt, IDxpStyleResolver s);
	void VisitText(Text t, IDxpStyleResolver s);
	void VisitTab(TabChar tab, IDxpStyleResolver s);
	void VisitBreak(Break br, IDxpStyleResolver s);
	void VisitCarriageReturn(CarriageReturn cr, IDxpStyleResolver s);
	void VisitProofError(ProofError pe, IDxpStyleResolver s);
	void VisitNoBreakHyphen(NoBreakHyphen h, IDxpStyleResolver s);
	void VisitSectionProperties(SectionProperties sp, IDxpStyleResolver s);
	IDisposable VisitBodyBegin(Body body, IDxpStyleResolver s);
	IDisposable VisitBlockBegin(OpenXmlElement child, IDxpStyleResolver s);
	void VisitTableRowProperties(TableRowProperties trp, IDxpStyleResolver s);
	IDisposable VisitDrawingBegin(Drawing d, DxpDrawingInfo? info, IDxpStyleResolver s);
	void VisitFootnoteReference(FootnoteReference fr, long id, int index, IDxpStyleResolver s);
	IDisposable VisitFootnoteBegin(Footnote fn, long id, int index, IDxpStyleResolver s);
	void VisitFootnoteReferenceMark(FootnoteReferenceMark m, long? footnoteId, int index, IDxpStyleResolver s);
	void VisitCommentInline(string id, string text, IDxpStyleResolver s);
	void VisitDayShort(DayShort ds, IDxpStyleResolver s);
	void VisitMonthShort(MonthShort ms, IDxpStyleResolver s);
	void VisitYearShort(YearShort ys, IDxpStyleResolver s);
	void VisitDayLong(DayLong dl, IDxpStyleResolver s);
	void VisitMonthLong(MonthLong ml, IDxpStyleResolver s);
	void VisitYearLong(YearLong yl, IDxpStyleResolver s);
	void VisitPageNumber(PageNumber pn, IDxpStyleResolver s);
	void VisitAnnotationReference(AnnotationReferenceMark arm, IDxpStyleResolver s);
	void VisitEndnoteReferenceMark(EndnoteReferenceMark erm, IDxpStyleResolver s);
	void VisitEndnoteReference(EndnoteReference enr, IDxpStyleResolver s);
	void VisitSeparatorMark(SeparatorMark sep, IDxpStyleResolver s);
	void VisitContinuationSeparatorMark(ContinuationSeparatorMark csep, IDxpStyleResolver s);
	void VisitSoftHyphen(SoftHyphen sh, IDxpStyleResolver s);
	void VisitSymbol(SymbolChar sym, IDxpStyleResolver s);
	void VisitPositionalTab(PositionalTab ptab, IDxpStyleResolver s);
	IDisposable VisitRubyBegin(Ruby ruby, IDxpStyleResolver s);
	void VisitDeletedFieldCode(DeletedFieldCode dfc, IDxpStyleResolver s);
	void VisitEmbeddedObject(EmbeddedObject obj, IDxpStyleResolver s);
	IDisposable VisitLegacyPictureBegin(Picture pict, IDxpStyleResolver s);
	void VisitRubyProperties(RubyProperties pr, IDxpStyleResolver s);
	IDisposable VisitRubyContentBegin(RubyContentType rc, bool isBase, IDxpStyleResolver s);
	void VisitPermStart(PermStart ps, IDxpStyleResolver s);
	void VisitPermEnd(PermEnd pe2, IDxpStyleResolver s);
	void VisitMoveFromRangeStart(MoveFromRangeStart mfrs, IDxpStyleResolver s);
	void VisitMoveFromRangeEnd(MoveFromRangeEnd mfre, IDxpStyleResolver s);
	void VisitMoveToRangeStart(MoveToRangeStart mtrs, IDxpStyleResolver s);
	void VisitMoveToRangeEnd(MoveToRangeEnd mtre, IDxpStyleResolver s);
	IDisposable VisitInsertedBegin(Inserted ins, IDxpStyleResolver s);
	IDisposable VisitDeletedBegin(Deleted del, IDxpStyleResolver s);
	void VisitOMathParagraph(DocumentFormat.OpenXml.Math.Paragraph oMathPara, IDxpStyleResolver s);
	void VisitOMath(DocumentFormat.OpenXml.Math.OfficeMath oMath, IDxpStyleResolver s);
	void VisitDeletedTableRowMark(Deleted del, TableRowProperties trPr, TableRow? tr, IDxpStyleResolver s);
	void VisitDeletedParagraphMark(Deleted del, ParagraphProperties pPr, Paragraph? p, IDxpStyleResolver s);
	void VisitInsertedParagraphMark(Inserted ins, ParagraphProperties pPr2, Paragraph? p, IDxpStyleResolver s);
	void VisitInsertedNumberingProperties(Inserted ins, NumberingProperties numPr, ParagraphProperties? pPr, Paragraph? p, IDxpStyleResolver s);
	void VisitInsertedTableRowMark(Inserted ins, TableRowProperties trPr, TableRow? tr, IDxpStyleResolver s);
	void VisitCustomXmlInsRangeStart(CustomXmlInsRangeStart cins, IDxpStyleResolver s);
	void VisitCustomXmlInsRangeEnd(CustomXmlInsRangeEnd cine, IDxpStyleResolver s);
	void VisitCustomXmlDelRangeStart(CustomXmlDelRangeStart cdls, IDxpStyleResolver s);
	void VisitCustomXmlDelRangeEnd(CustomXmlDelRangeEnd cdle, IDxpStyleResolver s);
	void VisitCustomXmlMoveFromRangeStart(CustomXmlMoveFromRangeStart cmfs, IDxpStyleResolver s);
	void VisitCustomXmlMoveFromRangeEnd(CustomXmlMoveFromRangeEnd cmfe, IDxpStyleResolver s);
	void VisitCustomXmlMoveToRangeStart(CustomXmlMoveToRangeStart cmts, IDxpStyleResolver s);
	void VisitCustomXmlMoveToRangeEnd(CustomXmlMoveToRangeEnd cmte, IDxpStyleResolver s);
	IDisposable VisitSdtBlockBegin(SdtBlock sdt, IDxpStyleResolver s);
	IDisposable VisitCustomXmlBlockBegin(CustomXmlBlock cx, IDxpStyleResolver s);
	void VisitAltChunk(AltChunk ac, IDxpStyleResolver s);
	void VisitCommentRangeStart(CommentRangeStart crs, IDxpStyleResolver s);
	void VisitCommentRangeEnd(CommentRangeEnd cre, IDxpStyleResolver s);
	void VisitSdtProperties(SdtProperties pr, IDxpStyleResolver s);
	IDisposable VisitSdtContentBlockBegin(SdtContentBlock content, IDxpStyleResolver s);
	void VisitCustomXmlProperties(CustomXmlProperties pr, IDxpStyleResolver s);
	void VisitCustomXmlConflictInsertionRangeStart(CustomXmlConflictInsertionRangeStart cxCis, IDxpStyleResolver s);
	void VisitCustomXmlConflictInsertionRangeEnd(CustomXmlConflictInsertionRangeEnd cxCie, IDxpStyleResolver s);
	void VisitCustomXmlConflictDeletionRangeStart(CustomXmlConflictDeletionRangeStart cxCds, IDxpStyleResolver s);
	void VisitCustomXmlConflictDeletionRangeEnd(CustomXmlConflictDeletionRangeEnd cxCde, IDxpStyleResolver s);
	void VisitMoveFromRun(MoveFromRun mfr, IDxpStyleResolver s);
	void VisitMoveToRun(MoveToRun mtr, IDxpStyleResolver s);
	void VisitContentPart(DocumentFormat.OpenXml.Wordprocessing.ContentPart cp, IDxpStyleResolver s);
	IDisposable VisitCustomXmlRunBegin(CustomXmlRun cxr, IDxpStyleResolver s);
	IDisposable VisitSimpleFieldBegin(SimpleField fld, IDxpStyleResolver s);
	IDisposable VisitSdtRunBegin(SdtRun sdtRun, IDxpStyleResolver s);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Accent mAccent, IDxpStyleResolver s);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Bar mBar, IDxpStyleResolver s);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Box mBox, IDxpStyleResolver s);
	void VisitOMathRun(DocumentFormat.OpenXml.Math.Run mMathRun, IDxpStyleResolver s);
	IDisposable VisitBidirectionalOverrideBegin(BidirectionalOverride bdo, IDxpStyleResolver s);
	IDisposable VisitBidirectionalEmbeddingBegin(BidirectionalEmbedding bdi, IDxpStyleResolver s);
	void VisitSubDocumentReference(SubDocumentReference subDoc, IDxpStyleResolver s);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.BorderBox mBorderBox, IDxpStyleResolver s);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Delimiter mDelim, IDxpStyleResolver s);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.EquationArray mEqArr, IDxpStyleResolver s);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Fraction mFrac, IDxpStyleResolver s);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.MathFunction mFunc, IDxpStyleResolver s);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.GroupChar mGroupChr, IDxpStyleResolver s);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.LimitLower mLimLow, IDxpStyleResolver s);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.LimitUpper mLimUpp, IDxpStyleResolver s);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Matrix mMat, IDxpStyleResolver s);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Nary mNary, IDxpStyleResolver s);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Phantom mPhant, IDxpStyleResolver s);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Radical mRad, IDxpStyleResolver s);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.PreSubSuper mPreSubSup, IDxpStyleResolver s);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Subscript mSub, IDxpStyleResolver s);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.SubSuperscript mSubSup, IDxpStyleResolver s);
	void VisitOMathElement(DocumentFormat.OpenXml.Math.Superscript mSup, IDxpStyleResolver s);
	void VisitSdtEndCharProperties(SdtEndCharProperties endPr, IDxpStyleResolver s);
	IDisposable VisitSdtContentRunBegin(SdtContentRun content, IDxpStyleResolver s);
	void VisitFieldData(FieldData data, IDxpStyleResolver s);
	void VisitConflictInsertion(ConflictInsertion cIns, IDxpStyleResolver s);
	void VisitConflictDeletion(ConflictDeletion cDel, IDxpStyleResolver s);
	IDisposable VisitSdtRowBegin(SdtRow sdtRow, IDxpStyleResolver s);
	IDisposable VisitCustomXmlRowBegin(CustomXmlRow cxRow, IDxpStyleResolver s);
	bool AcceptAlternateContentChoice(AlternateContentChoice choice, IReadOnlyList<string> required, IDxpStyleResolver s);
	bool SupportsNamespaces(IReadOnlyList<string> mu);
	IDisposable VisitEndnoteBegin(Endnote item1, long item3, int item2, IDxpStyleResolver s);
	IDisposable VisitTextBoxContentBegin(TextBoxContent txbx, IDxpStyleResolver s);
	IDisposable VisitSmartTagRunBegin(OpenXmlUnknownElement smart, string elementName, string elementUri, IDxpStyleResolver s);
	void VisitSmartTagProperties(OpenXmlUnknownElement smartTagPr, List<CustomXmlAttribute> attrs, IDxpStyleResolver s);
	IDisposable VisitAlternateContentBegin(AlternateContent ac, IDxpStyleResolver s);
	void VisitUnknown(string context, OpenXmlElement el, IDxpStyleResolver s);
	IDisposable VisitSdtCellBegin(SdtCell sdtCell, IDxpStyleResolver s);
	IDisposable VisitCustomXmlCellBegin(CustomXmlCell cxCell, IDxpStyleResolver s);
	void VisitDocumentSettings(Settings settings, IDxpStyleResolver s);
	void VisitDocumentBackground(object background, IDxpStyleResolver s);
	void VisitSectionLayout(SectionProperties sp, SectionLayout layout, IDxpStyleResolver s);
	IDisposable VisitSectionHeaderBegin(Header hdr, object value, IDxpStyleResolver s);
	IDisposable VisitSectionFooterBegin(Footer ftr, object value, IDxpStyleResolver s);
	void VisitCoreFileProperties(IPackageProperties core);
	void VisitCustomFileProperties(IEnumerable<CustomFileProperty> custom);
	void VisitGlossaryDocumentBegin(GlossaryDocument gl, IDxpStyleResolver s);
	void VisitBibliographySources(object bibliographyPart, object bib);
}
