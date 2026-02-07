using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using System.Xml.Linq;

namespace DocxportNet.Walker;

public abstract class DxpVisitorMiddlewareBase : DxpIVisitor
{
	protected DxpIVisitor _next;

	protected DxpVisitorMiddlewareBase(DxpIVisitor next)
	{
		_next = next ?? throw new ArgumentNullException(nameof(next));
	}

	public DxpIVisitor Next => _next;

	public virtual void VisitComplexFieldBegin(FieldChar begin, DxpIDocumentContext d)
	{
		_next.VisitComplexFieldBegin(begin, d);
	}

	public virtual void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
	{
		_next.VisitComplexFieldInstruction(instr, text, d);
	}

	public virtual void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext d)
	{
		_next.VisitComplexFieldSeparate(separate, d);
	}

	public virtual IDisposable VisitComplexFieldResultBegin(DxpIDocumentContext d)
	{
		return _next.VisitComplexFieldResultBegin(d);
	}

	public virtual void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
	{
		_next.VisitComplexFieldCachedResultText(text, d);
	}

	public virtual void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
	{
		_next.VisitComplexFieldEnd(end, d);
	}

	public virtual void StyleBoldBegin(DxpIDocumentContext d)
	{
		_next.StyleBoldBegin(d);
	}

	public virtual void StyleBoldEnd(DxpIDocumentContext d)
	{
		_next.StyleBoldEnd(d);
	}

	public virtual void StyleItalicBegin(DxpIDocumentContext d)
	{
		_next.StyleItalicBegin(d);
	}

	public virtual void StyleItalicEnd(DxpIDocumentContext d)
	{
		_next.StyleItalicEnd(d);
	}

	public virtual void StyleUnderlineBegin(DxpIDocumentContext d)
	{
		_next.StyleUnderlineBegin(d);
	}

	public virtual void StyleUnderlineEnd(DxpIDocumentContext d)
	{
		_next.StyleUnderlineEnd(d);
	}

	public virtual void StyleStrikeBegin(DxpIDocumentContext d)
	{
		_next.StyleStrikeBegin(d);
	}

	public virtual void StyleStrikeEnd(DxpIDocumentContext d)
	{
		_next.StyleStrikeEnd(d);
	}

	public virtual void StyleDoubleStrikeBegin(DxpIDocumentContext d)
	{
		_next.StyleDoubleStrikeBegin(d);
	}

	public virtual void StyleDoubleStrikeEnd(DxpIDocumentContext d)
	{
		_next.StyleDoubleStrikeEnd(d);
	}

	public virtual void StyleSuperscriptBegin(DxpIDocumentContext d)
	{
		_next.StyleSuperscriptBegin(d);
	}

	public virtual void StyleSuperscriptEnd(DxpIDocumentContext d)
	{
		_next.StyleSuperscriptEnd(d);
	}

	public virtual void StyleSubscriptBegin(DxpIDocumentContext d)
	{
		_next.StyleSubscriptBegin(d);
	}

	public virtual void StyleSubscriptEnd(DxpIDocumentContext d)
	{
		_next.StyleSubscriptEnd(d);
	}

	public virtual void StyleSmallCapsBegin(DxpIDocumentContext d)
	{
		_next.StyleSmallCapsBegin(d);
	}

	public virtual void StyleSmallCapsEnd(DxpIDocumentContext d)
	{
		_next.StyleSmallCapsEnd(d);
	}

	public virtual void StyleAllCapsBegin(DxpIDocumentContext d)
	{
		_next.StyleAllCapsBegin(d);
	}

	public virtual void StyleAllCapsEnd(DxpIDocumentContext d)
	{
		_next.StyleAllCapsEnd(d);
	}

	public virtual void StyleFontBegin(DxpFont font, DxpIDocumentContext d)
	{
		_next.StyleFontBegin(font, d);
	}

	public virtual void StyleFontEnd(DxpIDocumentContext d)
	{
		_next.StyleFontEnd(d);
	}

	public virtual void SetOutput(Stream stream)
	{
		_next.SetOutput(stream);
	}

	public virtual bool AcceptAlternateContentChoice(AlternateContentChoice choice, IReadOnlyList<string> required, DxpIDocumentContext d)
	{
		return _next.AcceptAlternateContentChoice(choice, required, d);
	}

	public virtual IDisposable VisitAlternateContentBegin(AlternateContent ac, DxpIDocumentContext d)
	{
		return _next.VisitAlternateContentBegin(ac, d);
	}

	public virtual void VisitAltChunk(AltChunk ac, DxpIDocumentContext d)
	{
		_next.VisitAltChunk(ac, d);
	}

	public virtual void VisitAnnotationReference(AnnotationReferenceMark arm, DxpIDocumentContext d)
	{
		_next.VisitAnnotationReference(arm, d);
	}

	public virtual IDisposable VisitBidirectionalEmbeddingBegin(BidirectionalEmbedding bdi, DxpIDocumentContext d)
	{
		return _next.VisitBidirectionalEmbeddingBegin(bdi, d);
	}

	public virtual IDisposable VisitBidirectionalOverrideBegin(BidirectionalOverride bdo, DxpIDocumentContext d)
	{
		return _next.VisitBidirectionalOverrideBegin(bdo, d);
	}

	public virtual void VisitBibliographySources(CustomXmlPart bibliographyPart, XDocument bib, DxpIDocumentContext d)
	{
		_next.VisitBibliographySources(bibliographyPart, bib, d);
	}

	public virtual IDisposable VisitBlockBegin(OpenXmlElement child, DxpIDocumentContext d)
	{
		return _next.VisitBlockBegin(child, d);
	}

	public virtual void VisitBookmarkEnd(BookmarkEnd be, DxpIDocumentContext d)
	{
		_next.VisitBookmarkEnd(be, d);
	}

	public virtual void VisitBookmarkStart(BookmarkStart bs, DxpIDocumentContext d)
	{
		_next.VisitBookmarkStart(bs, d);
	}

	public virtual void VisitBreak(Break br, DxpIDocumentContext d)
	{
		_next.VisitBreak(br, d);
	}

	public virtual void VisitCarriageReturn(CarriageReturn cr, DxpIDocumentContext d)
	{
		_next.VisitCarriageReturn(cr, d);
	}

	public virtual IDisposable VisitCommentBegin(DxpCommentInfo c, DxpCommentThread thread, DxpIDocumentContext d)
	{
		return _next.VisitCommentBegin(c, thread, d);
	}

	public virtual IDisposable VisitCommentThreadBegin(string anchorId, DxpCommentThread thread, DxpIDocumentContext d)
	{
		return _next.VisitCommentThreadBegin(anchorId, thread, d);
	}

	public virtual void VisitConflictDeletion(ConflictDeletion cDel, DxpIDocumentContext d)
	{
		_next.VisitConflictDeletion(cDel, d);
	}

	public virtual void VisitConflictInsertion(ConflictInsertion cIns, DxpIDocumentContext d)
	{
		_next.VisitConflictInsertion(cIns, d);
	}

	public virtual void VisitContentPart(DocumentFormat.OpenXml.Wordprocessing.ContentPart cp, DxpIDocumentContext d)
	{
		_next.VisitContentPart(cp, d);
	}

	public virtual void VisitContinuationSeparatorMark(ContinuationSeparatorMark csep, DxpIDocumentContext d)
	{
		_next.VisitContinuationSeparatorMark(csep, d);
	}

	public virtual IDisposable VisitCustomXmlBlockBegin(CustomXmlBlock cx, DxpIDocumentContext d)
	{
		return _next.VisitCustomXmlBlockBegin(cx, d);
	}

	public virtual IDisposable VisitCustomXmlCellBegin(CustomXmlCell cxCell, DxpIDocumentContext d)
	{
		return _next.VisitCustomXmlCellBegin(cxCell, d);
	}

	public virtual void VisitCustomXmlConflictDeletionRangeEnd(CustomXmlConflictDeletionRangeEnd cxCde, DxpIDocumentContext d)
	{
		_next.VisitCustomXmlConflictDeletionRangeEnd(cxCde, d);
	}

	public virtual void VisitCustomXmlConflictDeletionRangeStart(CustomXmlConflictDeletionRangeStart cxCds, DxpIDocumentContext d)
	{
		_next.VisitCustomXmlConflictDeletionRangeStart(cxCds, d);
	}

	public virtual void VisitCustomXmlConflictInsertionRangeEnd(CustomXmlConflictInsertionRangeEnd cxCie, DxpIDocumentContext d)
	{
		_next.VisitCustomXmlConflictInsertionRangeEnd(cxCie, d);
	}

	public virtual void VisitCustomXmlConflictInsertionRangeStart(CustomXmlConflictInsertionRangeStart cxCis, DxpIDocumentContext d)
	{
		_next.VisitCustomXmlConflictInsertionRangeStart(cxCis, d);
	}

	public virtual void VisitCustomXmlDelRangeEnd(CustomXmlDelRangeEnd cdle, DxpIDocumentContext d)
	{
		_next.VisitCustomXmlDelRangeEnd(cdle, d);
	}

	public virtual void VisitCustomXmlDelRangeStart(CustomXmlDelRangeStart cdls, DxpIDocumentContext d)
	{
		_next.VisitCustomXmlDelRangeStart(cdls, d);
	}

	public virtual void VisitCustomXmlInsRangeEnd(CustomXmlInsRangeEnd cine, DxpIDocumentContext d)
	{
		_next.VisitCustomXmlInsRangeEnd(cine, d);
	}

	public virtual void VisitCustomXmlInsRangeStart(CustomXmlInsRangeStart cins, DxpIDocumentContext d)
	{
		_next.VisitCustomXmlInsRangeStart(cins, d);
	}

	public virtual void VisitCustomXmlMoveFromRangeEnd(CustomXmlMoveFromRangeEnd cmfe, DxpIDocumentContext d)
	{
		_next.VisitCustomXmlMoveFromRangeEnd(cmfe, d);
	}

	public virtual void VisitCustomXmlMoveFromRangeStart(CustomXmlMoveFromRangeStart cmfs, DxpIDocumentContext d)
	{
		_next.VisitCustomXmlMoveFromRangeStart(cmfs, d);
	}

	public virtual void VisitCustomXmlMoveToRangeEnd(CustomXmlMoveToRangeEnd cmte, DxpIDocumentContext d)
	{
		_next.VisitCustomXmlMoveToRangeEnd(cmte, d);
	}

	public virtual void VisitCustomXmlMoveToRangeStart(CustomXmlMoveToRangeStart cmts, DxpIDocumentContext d)
	{
		_next.VisitCustomXmlMoveToRangeStart(cmts, d);
	}

	public virtual IDisposable VisitCustomXmlRowBegin(CustomXmlRow cxRow, DxpIDocumentContext d)
	{
		return _next.VisitCustomXmlRowBegin(cxRow, d);
	}

	public virtual IDisposable VisitCustomXmlRunBegin(CustomXmlRun cxr, DxpIDocumentContext d)
	{
		return _next.VisitCustomXmlRunBegin(cxr, d);
	}

	public virtual void VisitDayLong(DayLong dl, DxpIDocumentContext d)
	{
		_next.VisitDayLong(dl, d);
	}

	public virtual void VisitDayShort(DayShort ds, DxpIDocumentContext d)
	{
		_next.VisitDayShort(ds, d);
	}

	public virtual IDisposable VisitDeletedBegin(Deleted del, DxpIDocumentContext d)
	{
		return _next.VisitDeletedBegin(del, d);
	}

	public virtual void VisitDeletedFieldCode(DeletedFieldCode dfc, DxpIDocumentContext d)
	{
		_next.VisitDeletedFieldCode(dfc, d);
	}

	public virtual void VisitDeletedParagraphMark(Deleted del, ParagraphProperties pPr, Paragraph? p, DxpIDocumentContext d)
	{
		_next.VisitDeletedParagraphMark(del, pPr, p, d);
	}

	public virtual IDisposable VisitDeletedRunBegin(DeletedRun dr, DxpIDocumentContext d)
	{
		return _next.VisitDeletedRunBegin(dr, d);
	}

	public virtual void VisitDeletedTableRowMark(Deleted del, TableRowProperties trPr, TableRow? tr, DxpIDocumentContext d)
	{
		_next.VisitDeletedTableRowMark(del, trPr, tr, d);
	}

	public virtual void VisitDeletedText(DeletedText dt, DxpIDocumentContext d)
	{
		_next.VisitDeletedText(dt, d);
	}

	public virtual IDisposable VisitDocumentBodyBegin(Body body, DxpIDocumentContext d)
	{
		return _next.VisitDocumentBodyBegin(body, d);
	}

	public virtual IDisposable VisitDocumentBegin(WordprocessingDocument doc, DxpIDocumentContext documentContext)
	{
		return _next.VisitDocumentBegin(doc, documentContext);
	}

	public virtual IDisposable VisitDrawingBegin(Drawing drw, DxpDrawingInfo? info, DxpIDocumentContext d)
	{
		return _next.VisitDrawingBegin(drw, info, d);
	}

	public virtual void VisitEmbeddedObject(EmbeddedObject obj, DxpIDocumentContext d)
	{
		_next.VisitEmbeddedObject(obj, d);
	}

	public virtual IDisposable VisitEndnoteBegin(Endnote item1, long item3, int item2, DxpIDocumentContext d)
	{
		return _next.VisitEndnoteBegin(item1, item3, item2, d);
	}

	public virtual void VisitEndnoteReference(EndnoteReference enr, DxpIDocumentContext d)
	{
		_next.VisitEndnoteReference(enr, d);
	}

	public virtual void VisitEndnoteReferenceMark(EndnoteReferenceMark erm, DxpIDocumentContext d)
	{
		_next.VisitEndnoteReferenceMark(erm, d);
	}

	public virtual void VisitFieldData(FieldData data, DxpIDocumentContext d)
	{
		_next.VisitFieldData(data, d);
	}

	public virtual IDisposable VisitFootnoteBegin(Footnote fn, DxpIFootnoteContext footnote, DxpIDocumentContext d)
	{
		return _next.VisitFootnoteBegin(fn, footnote, d);
	}

	public virtual void VisitFootnoteReference(FootnoteReference fr, DxpIFootnoteContext footnote, DxpIDocumentContext d)
	{
		_next.VisitFootnoteReference(fr, footnote, d);
	}

	public virtual void VisitFootnoteReferenceMark(FootnoteReferenceMark m, DxpIFootnoteContext footnote, DxpIDocumentContext d)
	{
		_next.VisitFootnoteReferenceMark(m, footnote, d);
	}

	public virtual IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target, DxpIDocumentContext d)
	{
		return _next.VisitHyperlinkBegin(link, target, d);
	}

	public virtual IDisposable VisitInsertedBegin(Inserted ins, DxpIDocumentContext d)
	{
		return _next.VisitInsertedBegin(ins, d);
	}

	public virtual void VisitInsertedNumbering(Inserted ins, DxpMarker? marker, DxpStyleEffectiveIndentTwips indent, Paragraph? p, DxpIDocumentContext d)
	{
		_next.VisitInsertedNumbering(ins, marker, indent, p, d);
	}

	public virtual void VisitInsertedParagraphMark(Inserted ins, ParagraphProperties pPr2, Paragraph? p, DxpIDocumentContext d)
	{
		_next.VisitInsertedParagraphMark(ins, pPr2, p, d);
	}

	public virtual IDisposable VisitInsertedRunBegin(InsertedRun ir, DxpIDocumentContext d)
	{
		return _next.VisitInsertedRunBegin(ir, d);
	}

	public virtual void VisitInsertedTableRowMark(Inserted ins, TableRowProperties trPr, TableRow? tr, DxpIDocumentContext d)
	{
		_next.VisitInsertedTableRowMark(ins, trPr, tr, d);
	}

	public virtual void VisitLastRenderedPageBreak(LastRenderedPageBreak pb, DxpIDocumentContext d)
	{
		_next.VisitLastRenderedPageBreak(pb, d);
	}

	public virtual IDisposable VisitLegacyPictureBegin(Picture pict, DxpIDocumentContext d)
	{
		return _next.VisitLegacyPictureBegin(pict, d);
	}

	public virtual void VisitMonthLong(MonthLong ml, DxpIDocumentContext d)
	{
		_next.VisitMonthLong(ml, d);
	}

	public virtual void VisitMonthShort(MonthShort ms, DxpIDocumentContext d)
	{
		_next.VisitMonthShort(ms, d);
	}

	public virtual void VisitMoveFromRangeEnd(MoveFromRangeEnd mfre, DxpIDocumentContext d)
	{
		_next.VisitMoveFromRangeEnd(mfre, d);
	}

	public virtual void VisitMoveFromRangeStart(MoveFromRangeStart mfrs, DxpIDocumentContext d)
	{
		_next.VisitMoveFromRangeStart(mfrs, d);
	}

	public virtual void VisitMoveFromRun(MoveFromRun mfr, DxpIDocumentContext d)
	{
		_next.VisitMoveFromRun(mfr, d);
	}

	public virtual void VisitMoveToRangeEnd(MoveToRangeEnd mtre, DxpIDocumentContext d)
	{
		_next.VisitMoveToRangeEnd(mtre, d);
	}

	public virtual void VisitMoveToRangeStart(MoveToRangeStart mtrs, DxpIDocumentContext d)
	{
		_next.VisitMoveToRangeStart(mtrs, d);
	}

	public virtual void VisitMoveToRun(MoveToRun mtr, DxpIDocumentContext d)
	{
		_next.VisitMoveToRun(mtr, d);
	}

	public virtual void VisitNoBreakHyphen(NoBreakHyphen h, DxpIDocumentContext d)
	{
		_next.VisitNoBreakHyphen(h, d);
	}

	public virtual void VisitOMath(DocumentFormat.OpenXml.Math.OfficeMath oMath, DxpIDocumentContext d)
	{
		_next.VisitOMath(oMath, d);
	}

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Accent mAccent, DxpIDocumentContext d)
	{
		_next.VisitOMathElement(mAccent, d);
	}

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Bar mBar, DxpIDocumentContext d)
	{
		_next.VisitOMathElement(mBar, d);
	}

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.BorderBox mBorderBox, DxpIDocumentContext d)
	{
		_next.VisitOMathElement(mBorderBox, d);
	}

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Box mBox, DxpIDocumentContext d)
	{
		_next.VisitOMathElement(mBox, d);
	}

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Delimiter mDelim, DxpIDocumentContext d)
	{
		_next.VisitOMathElement(mDelim, d);
	}

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.EquationArray mEqArr, DxpIDocumentContext d)
	{
		_next.VisitOMathElement(mEqArr, d);
	}

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Fraction mFrac, DxpIDocumentContext d)
	{
		_next.VisitOMathElement(mFrac, d);
	}

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.GroupChar mGroupChr, DxpIDocumentContext d)
	{
		_next.VisitOMathElement(mGroupChr, d);
	}

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.LimitLower mLimLow, DxpIDocumentContext d)
	{
		_next.VisitOMathElement(mLimLow, d);
	}

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.LimitUpper mLimUpp, DxpIDocumentContext d)
	{
		_next.VisitOMathElement(mLimUpp, d);
	}

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.MathFunction mFunc, DxpIDocumentContext d)
	{
		_next.VisitOMathElement(mFunc, d);
	}

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Matrix mMat, DxpIDocumentContext d)
	{
		_next.VisitOMathElement(mMat, d);
	}

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Nary mNary, DxpIDocumentContext d)
	{
		_next.VisitOMathElement(mNary, d);
	}

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Phantom mPhant, DxpIDocumentContext d)
	{
		_next.VisitOMathElement(mPhant, d);
	}

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.PreSubSuper mPreSubSup, DxpIDocumentContext d)
	{
		_next.VisitOMathElement(mPreSubSup, d);
	}

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Radical mRad, DxpIDocumentContext d)
	{
		_next.VisitOMathElement(mRad, d);
	}

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Subscript mSub, DxpIDocumentContext d)
	{
		_next.VisitOMathElement(mSub, d);
	}

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.SubSuperscript mSubSup, DxpIDocumentContext d)
	{
		_next.VisitOMathElement(mSubSup, d);
	}

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Superscript mSup, DxpIDocumentContext d)
	{
		_next.VisitOMathElement(mSup, d);
	}

	public virtual void VisitOMathParagraph(DocumentFormat.OpenXml.Math.Paragraph oMathPara, DxpIDocumentContext d)
	{
		_next.VisitOMathParagraph(oMathPara, d);
	}

	public virtual void VisitOMathRun(DocumentFormat.OpenXml.Math.Run mMathRun, DxpIDocumentContext d)
	{
		_next.VisitOMathRun(mMathRun, d);
	}

	public virtual void VisitPageNumber(PageNumber pn, DxpIDocumentContext d)
	{
		_next.VisitPageNumber(pn, d);
	}

	public virtual IDisposable VisitParagraphBegin(Paragraph p, DxpIDocumentContext d, DxpIParagraphContext paragraph)
	{
		return _next.VisitParagraphBegin(p, d, paragraph);
	}

	public virtual void VisitPermEnd(PermEnd pe2, DxpIDocumentContext d)
	{
		_next.VisitPermEnd(pe2, d);
	}

	public virtual void VisitPermStart(PermStart ps, DxpIDocumentContext d)
	{
		_next.VisitPermStart(ps, d);
	}

	public virtual void VisitPositionalTab(PositionalTab ptab, DxpIDocumentContext d)
	{
		_next.VisitPositionalTab(ptab, d);
	}

	public virtual void VisitProofError(ProofError pe, DxpIDocumentContext d)
	{
		_next.VisitProofError(pe, d);
	}

	public virtual IDisposable VisitRubyBegin(Ruby ruby, DxpIDocumentContext d)
	{
		return _next.VisitRubyBegin(ruby, d);
	}

	public virtual IDisposable VisitRubyContentBegin(RubyContentType rc, bool isBase, DxpIDocumentContext d)
	{
		return _next.VisitRubyContentBegin(rc, isBase, d);
	}

	public virtual IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
	{
		return _next.VisitRunBegin(r, d);
	}

	public virtual IDisposable VisitSectionBegin(SectionProperties properties, SectionLayout layout, DxpIDocumentContext d)
	{
		return _next.VisitSectionBegin(properties, layout, d);
	}

	public virtual IDisposable VisitSectionBodyBegin(SectionProperties properties, DxpIDocumentContext d)
	{
		return _next.VisitSectionBodyBegin(properties, d);
	}

	public virtual IDisposable VisitSectionFooterBegin(Footer ftr, object value, DxpIDocumentContext d)
	{
		return _next.VisitSectionFooterBegin(ftr, value, d);
	}

	public virtual IDisposable VisitSectionHeaderBegin(Header hdr, object value, DxpIDocumentContext d)
	{
		return _next.VisitSectionHeaderBegin(hdr, value, d);
	}

	public virtual void VisitSeparatorMark(SeparatorMark sep, DxpIDocumentContext d)
	{
		_next.VisitSeparatorMark(sep, d);
	}

	public virtual IDisposable VisitSdtBlockBegin(SdtBlock sdt, DxpIDocumentContext d)
	{
		return _next.VisitSdtBlockBegin(sdt, d);
	}

	public virtual IDisposable VisitSdtCellBegin(SdtCell sdtCell, DxpIDocumentContext d)
	{
		return _next.VisitSdtCellBegin(sdtCell, d);
	}

	public virtual IDisposable VisitSdtContentBlockBegin(SdtContentBlock content, DxpIDocumentContext d)
	{
		return _next.VisitSdtContentBlockBegin(content, d);
	}

	public virtual IDisposable VisitSdtContentRunBegin(SdtContentRun content, DxpIDocumentContext d)
	{
		return _next.VisitSdtContentRunBegin(content, d);
	}

	public virtual IDisposable VisitSdtRowBegin(SdtRow sdtRow, DxpIDocumentContext d)
	{
		return _next.VisitSdtRowBegin(sdtRow, d);
	}

	public virtual IDisposable VisitSdtRunBegin(SdtRun sdtRun, DxpIDocumentContext d)
	{
		return _next.VisitSdtRunBegin(sdtRun, d);
	}

	public virtual IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
	{
		return _next.VisitSimpleFieldBegin(fld, d);
	}

	public virtual void VisitSoftHyphen(SoftHyphen sh, DxpIDocumentContext d)
	{
		_next.VisitSoftHyphen(sh, d);
	}

	public virtual IDisposable VisitSmartTagRunBegin(OpenXmlUnknownElement smart, string elementName, string elementUri, DxpIDocumentContext d)
	{
		return _next.VisitSmartTagRunBegin(smart, elementName, elementUri, d);
	}

	public virtual void VisitSubDocumentReference(SubDocumentReference subDoc, DxpIDocumentContext d)
	{
		_next.VisitSubDocumentReference(subDoc, d);
	}

	public virtual void VisitSymbol(SymbolChar sym, DxpIDocumentContext d)
	{
		_next.VisitSymbol(sym, d);
	}

	public virtual void VisitTab(TabChar tab, DxpIDocumentContext d)
	{
		_next.VisitTab(tab, d);
	}

	public virtual IDisposable VisitTableBegin(Table t, DxpTableModel model, DxpIDocumentContext d, DxpITableContext table)
	{
		return _next.VisitTableBegin(t, model, d, table);
	}

	public virtual IDisposable VisitTableCellBegin(TableCell tc, DxpITableCellContext cell, DxpIDocumentContext d)
	{
		return _next.VisitTableCellBegin(tc, cell, d);
	}

	public virtual IDisposable VisitTableRowBegin(TableRow tr, DxpITableRowContext row, DxpIDocumentContext d)
	{
		return _next.VisitTableRowBegin(tr, row, d);
	}

	public virtual void VisitText(Text t, DxpIDocumentContext d)
	{
		_next.VisitText(t, d);
	}

	public virtual IDisposable VisitTextBoxContentBegin(TextBoxContent txbx, DxpIDocumentContext d)
	{
		return _next.VisitTextBoxContentBegin(txbx, d);
	}

	public virtual void VisitUnknown(string context, OpenXmlElement el, DxpIDocumentContext d)
	{
		_next.VisitUnknown(context, el, d);
	}

	public virtual void VisitYearLong(YearLong yl, DxpIDocumentContext d)
	{
		_next.VisitYearLong(yl, d);
	}

	public virtual void VisitYearShort(YearShort ys, DxpIDocumentContext d)
	{
		_next.VisitYearShort(ys, d);
	}

}
