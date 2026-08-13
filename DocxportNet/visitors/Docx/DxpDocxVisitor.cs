using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Fields;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Visitors.Docx;

/// <summary>
/// Naive DOCX passthrough visitor. It clones the source package, then rebuilds the
/// stories surfaced by <see cref="Walker.DxpWalker"/> from visitor events.
/// Unsupported events are intentionally left to <see cref="Visitors.DxpVisitor"/>.
/// </summary>
public sealed class DxpDocxVisitor : DxpVisitor, IDisposable, DxpIFieldEvalProvider, DxpIPreserveLayoutFields
{
    private readonly Stack<OpenXmlCompositeElement> _parents = new();
    private Stream? _output;
    private WordprocessingDocument? _outputDocument;
    private readonly HashSet<OpenXmlPart> _sourceParts = new();
    private readonly Dictionary<(OpenXmlPart Part, string RelationshipId), string> _importedRelationships = new();
    private int _suppressDepth;
    private bool _completed;
    private Footnotes? _rebuiltFootnotes;
    private Endnotes? _rebuiltEndnotes;
    private Comments? _rebuiltComments;
    private readonly HashSet<string> _rebuiltCommentIds = new(StringComparer.Ordinal);
    public DxpFieldEval FieldEval { get; }

    public DxpDocxVisitor(ILogger? logger = null, DxpFieldEval? fieldEval = null) : base(logger)
        => FieldEval = fieldEval ?? new DxpFieldEval(logger: logger);

    public override void SetOutput(Stream stream)
        => _output = stream ?? throw new ArgumentNullException(nameof(stream));

    public override IDisposable VisitDocumentBegin(WordprocessingDocument doc, DxpIDocumentContext documentContext)
    {
        if (_output == null)
            throw new InvalidOperationException("An output stream must be assigned before walking the document.");

        _outputDocument = doc.Clone(_output, true);
        CollectParts(doc, _sourceParts);
        return DxpDisposable.Create(Complete);
    }

    public override IDisposable VisitDocumentBodyBegin(Body body, DxpIDocumentContext d)
    {
        var rebuilt = (Body)body.CloneNode(false);
        var main = _outputDocument?.MainDocumentPart
            ?? throw new InvalidOperationException("The cloned package has no main document part.");
        var document = main.Document
            ?? throw new InvalidOperationException("The cloned package has no main document root.");
        document.Body = rebuilt;
        return PushRoot(rebuilt);
    }

    public override IDisposable VisitSectionHeaderBegin(Header hdr, object value, DxpIDocumentContext d)
    {
        if (value is not DxpHeaderFooterContext { Part: HeaderPart sourcePart })
            return DxpDisposable.Empty;

        var destination = _outputDocument!.MainDocumentPart!.HeaderParts
            .FirstOrDefault(part => part.Uri == sourcePart.Uri);
        if (destination == null)
            return SuppressScope();

        var rebuilt = (Header)hdr.CloneNode(false);
        destination.Header = rebuilt;
        return PushRoot(rebuilt);
    }

    public override IDisposable VisitSectionFooterBegin(Footer ftr, object value, DxpIDocumentContext d)
    {
        if (value is not DxpHeaderFooterContext { Part: FooterPart sourcePart })
            return DxpDisposable.Empty;

        var destination = _outputDocument!.MainDocumentPart!.FooterParts
            .FirstOrDefault(part => part.Uri == sourcePart.Uri);
        if (destination == null)
            return SuppressScope();

        var rebuilt = (Footer)ftr.CloneNode(false);
        destination.Footer = rebuilt;
        return PushRoot(rebuilt);
    }

    public override IDisposable VisitSectionBegin(SectionProperties properties, SectionLayout layout, DxpIDocumentContext d)
    {
        if (_suppressDepth > 0 || properties.Parent is not Body ||
            d.CurrentPart == null || !_sourceParts.Contains(d.CurrentPart))
            return DxpDisposable.Empty;

        return DxpDisposable.Create(() => AppendClone(properties));
    }

    public override IDisposable VisitFootnoteBegin(Footnote fn, DxpIFootnoteContext footnote, DxpIDocumentContext d)
    {
        var part = _outputDocument?.MainDocumentPart?.FootnotesPart;
        if (part == null)
            return SuppressScope();

        if (_rebuiltFootnotes == null)
        {
            var sourceRoot = fn.Parent as Footnotes;
            _rebuiltFootnotes = sourceRoot != null
                ? (Footnotes)sourceRoot.CloneNode(false)
                : new Footnotes();
            if (sourceRoot != null)
            {
                foreach (var internalNote in sourceRoot.Elements<Footnote>().Where(IsInternalNote))
                    _rebuiltFootnotes.AppendChild(internalNote.CloneNode(true));
            }
            part.Footnotes = _rebuiltFootnotes;
        }

        var rebuilt = (Footnote)fn.CloneNode(false);
        _rebuiltFootnotes.AppendChild(rebuilt);
        return PushRoot(rebuilt);
    }

    public override IDisposable VisitEndnoteBegin(Endnote en, long id, int index, DxpIDocumentContext d)
    {
        var part = _outputDocument?.MainDocumentPart?.EndnotesPart;
        if (part == null)
            return SuppressScope();

        if (_rebuiltEndnotes == null)
        {
            var sourceRoot = en.Parent as Endnotes;
            _rebuiltEndnotes = sourceRoot != null
                ? (Endnotes)sourceRoot.CloneNode(false)
                : new Endnotes();
            if (sourceRoot != null)
            {
                foreach (var internalNote in sourceRoot.Elements<Endnote>().Where(IsInternalNote))
                    _rebuiltEndnotes.AppendChild(internalNote.CloneNode(true));
            }
            part.Endnotes = _rebuiltEndnotes;
        }

        var rebuilt = (Endnote)en.CloneNode(false);
        _rebuiltEndnotes.AppendChild(rebuilt);
        return PushRoot(rebuilt);
    }

    public override IDisposable VisitBlockBegin(OpenXmlElement child, DxpIDocumentContext d)
        => DxpDisposable.Empty;

    public override IDisposable VisitParagraphBegin(Paragraph p, DxpIDocumentContext d, DxpIParagraphContext paragraph)
    {
        var shell = CloneShell(p);
        if (d.CurrentPart != null && !_sourceParts.Contains(d.CurrentPart))
        {
            foreach (var reference in shell.Descendants()
                         .Where(element => element is HeaderReference or FooterReference)
                         .ToArray())
                reference.Remove();
        }
        return AppendContainer(shell);
    }

    public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
    {
        if (!r.HasChildren && _parents.Count > 0 && _parents.Peek() is Body or TableCell)
            return DxpDisposable.Empty;
        return AppendContainer(CloneShell(r));
    }

    public override IDisposable VisitTableBegin(Table t, DxpTableModel model, DxpIDocumentContext d, DxpITableContext table)
        => AppendContainer(CloneShell(t));

    public override IDisposable VisitTableRowBegin(TableRow tr, DxpITableRowContext row, DxpIDocumentContext d)
        => AppendContainer(CloneShell(tr));

    public override IDisposable VisitTableCellBegin(TableCell tc, DxpITableCellContext cell, DxpIDocumentContext d)
        => AppendContainer(CloneShell(tc));

    public override IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target, DxpIDocumentContext d)
        => AppendContainer(CloneShellWithExternalRelationships(link, d));

    public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
        => AppendContainer(CloneShell(fld));

    public override IDisposable VisitInsertedBegin(Inserted ins, DxpIDocumentContext d)
        => DxpDisposable.Empty;

    public override IDisposable VisitDeletedBegin(Deleted del, DxpIDocumentContext d)
        => DxpDisposable.Empty;

    public override IDisposable VisitInsertedRunBegin(InsertedRun ir, DxpIDocumentContext d)
        => AppendContainer(CloneShell(ir));

    public override IDisposable VisitDeletedRunBegin(DeletedRun dr, DxpIDocumentContext d)
        => AppendContainer(CloneShell(dr));

    public override IDisposable VisitSdtBlockBegin(SdtBlock sdt, DxpIDocumentContext d)
        => AppendContainer(CloneShell(sdt));

    public override IDisposable VisitSdtRunBegin(SdtRun sdtRun, DxpIDocumentContext d)
        => AppendContainer(CloneShell(sdtRun));

    public override IDisposable VisitSdtRowBegin(SdtRow sdtRow, DxpIDocumentContext d)
        => AppendSdtContainer(sdtRow, sdtRow.SdtContentRow);

    public override IDisposable VisitSdtCellBegin(SdtCell sdtCell, DxpIDocumentContext d)
        => AppendSdtContainer(sdtCell, sdtCell.SdtContentCell);

    public override IDisposable VisitSdtContentBlockBegin(SdtContentBlock content, DxpIDocumentContext d)
        => AppendContainer(CloneShell(content));

    public override IDisposable VisitSdtContentRunBegin(SdtContentRun content, DxpIDocumentContext d)
        => AppendContainer(CloneShell(content));

    public override IDisposable VisitCustomXmlBlockBegin(CustomXmlBlock cx, DxpIDocumentContext d)
        => AppendContainer(CloneShell(cx));

    public override IDisposable VisitCustomXmlRunBegin(CustomXmlRun cxr, DxpIDocumentContext d)
        => AppendContainer(CloneShell(cxr));

    public override IDisposable VisitCustomXmlRowBegin(CustomXmlRow cxRow, DxpIDocumentContext d)
        => AppendContainer(CloneShell(cxRow));

    public override IDisposable VisitCustomXmlCellBegin(CustomXmlCell cxCell, DxpIDocumentContext d)
        => AppendContainer(CloneShell(cxCell));

    public override IDisposable VisitAlternateContentBegin(AlternateContent ac, DxpIDocumentContext d)
    {
        // AlternateContent traversal is selective. Preserve the complete original node.
        AppendCloneWithExternalRelationships(ac, d);
        return SuppressScope();
    }

    public override IDisposable VisitDrawingBegin(Drawing drw, DxpDrawingInfo? info, DxpIDocumentContext d)
    {
        AppendCloneWithExternalRelationships(drw, d);
        return SuppressScope();
    }

    public override IDisposable VisitLegacyPictureBegin(Picture pict, DxpIDocumentContext d)
    {
        // Preserve the opaque VML carrier, but rebuild its WordprocessingML
        // textbox content so fields inside the textbox pass through the same
        // pipeline as fields in the main story.
        var clone = (Picture)pict.CloneNode(true);
        RemapExternalRelationships(clone, d.CurrentPart);
        var textBoxContent = clone.Descendants<TextBoxContent>().FirstOrDefault();
        if (textBoxContent == null)
        {
            if (_suppressDepth == 0 && _parents.Count > 0)
                _parents.Peek().AppendChild(clone);
            return SuppressScope();
        }

        textBoxContent.RemoveAllChildren();
        if (_suppressDepth > 0 || _parents.Count == 0)
            return SuppressScope();

        _parents.Peek().AppendChild(clone);
        _parents.Push(textBoxContent);
        return DxpDisposable.Create(() => _parents.Pop());
    }

    public override IDisposable VisitTextBoxContentBegin(TextBoxContent txbx, DxpIDocumentContext d)
        => DxpDisposable.Empty;

    public override void VisitOMath(DocumentFormat.OpenXml.Math.OfficeMath oMath, DxpIDocumentContext d) => AppendClone(oMath);
    public override void VisitOMathParagraph(DocumentFormat.OpenXml.Math.Paragraph oMathPara, DxpIDocumentContext d) => AppendClone(oMathPara);

    public override void VisitText(Text t, DxpIDocumentContext d) => AppendClone(t);
    public override void VisitDeletedText(DeletedText dt, DxpIDocumentContext d) => AppendClone(dt);
    public override void VisitBreak(Break br, DxpIDocumentContext d) => AppendClone(br);
    public override void VisitTab(TabChar tab, DxpIDocumentContext d) => AppendClone(tab);
    public override void VisitCarriageReturn(CarriageReturn cr, DxpIDocumentContext d) => AppendClone(cr);
    public override void VisitNoBreakHyphen(NoBreakHyphen h, DxpIDocumentContext d) => AppendClone(h);
    public override void VisitSoftHyphen(SoftHyphen sh, DxpIDocumentContext d) => AppendClone(sh);
    public override void VisitLastRenderedPageBreak(LastRenderedPageBreak pb, DxpIDocumentContext d) => AppendClone(pb);
    public override void VisitSymbol(SymbolChar sym, DxpIDocumentContext d) => AppendClone(sym);
    public override void VisitPositionalTab(PositionalTab ptab, DxpIDocumentContext d) => AppendClone(ptab);
    public override void VisitBookmarkStart(BookmarkStart bs, DxpIDocumentContext d) => AppendClone(bs);
    public override void VisitBookmarkEnd(BookmarkEnd be, DxpIDocumentContext d) => AppendClone(be);
    public override void VisitCommentRangeStart(CommentRangeStart start, DxpIDocumentContext d) => AppendClone(start);
    public override void VisitCommentRangeEnd(CommentRangeEnd end, DxpIDocumentContext d) => AppendClone(end);
    public override void VisitCommentReference(CommentReference reference, DxpIDocumentContext d) => AppendClone(reference);
    public override IDisposable VisitCommentThreadBegin(string anchorId, DxpCommentThread thread, DxpIDocumentContext d)
        => DxpDisposable.Empty;
    public override IDisposable VisitCommentBegin(DxpCommentInfo info, DxpCommentThread thread, DxpIDocumentContext d)
    {
        _ = thread;
        var part = _outputDocument?.MainDocumentPart?.WordprocessingCommentsPart;
        if (part == null || !_rebuiltCommentIds.Add(info.Id))
            return SuppressScope();

        if (_rebuiltComments == null)
        {
            var sourceRoot = info.Blocks.FirstOrDefault()?.Ancestors<Comments>().FirstOrDefault();
            _rebuiltComments = sourceRoot != null
                ? (Comments)sourceRoot.CloneNode(true)
                : new Comments();
            _rebuiltComments.RemoveAllChildren();
            part.Comments = _rebuiltComments;
        }

        var sourceComment = info.Blocks.FirstOrDefault()?.Ancestors<Comment>().FirstOrDefault();
        var comment = sourceComment != null
            ? (Comment)sourceComment.CloneNode(true)
            : new Comment
            {
                Id = info.Id,
                Author = info.Author,
                Initials = info.Initials,
                Date = info.DateUtc
            };
        comment.RemoveAllChildren();
        _rebuiltComments.AppendChild(comment);
        return PushRoot(comment);
    }
    public override void VisitPermStart(PermStart ps, DxpIDocumentContext d) => AppendClone(ps);
    public override void VisitPermEnd(PermEnd pe, DxpIDocumentContext d) => AppendClone(pe);
    public override void VisitProofError(ProofError pe, DxpIDocumentContext d) => AppendClone(pe);
    public override void VisitFootnoteReference(FootnoteReference fr, DxpIFootnoteContext footnote, DxpIDocumentContext d) => AppendClone(fr);
    public override void VisitFootnoteReferenceMark(FootnoteReferenceMark mark, DxpIFootnoteContext footnote, DxpIDocumentContext d) => AppendClone(mark);
    public override void VisitEndnoteReference(EndnoteReference enr, DxpIDocumentContext d) => AppendClone(enr);
    public override void VisitEndnoteReferenceMark(EndnoteReferenceMark mark, DxpIDocumentContext d) => AppendClone(mark);
    public override void VisitSeparatorMark(SeparatorMark mark, DxpIDocumentContext d) => AppendClone(mark);
    public override void VisitContinuationSeparatorMark(ContinuationSeparatorMark mark, DxpIDocumentContext d) => AppendClone(mark);
    public override void VisitAnnotationReference(AnnotationReferenceMark arm, DxpIDocumentContext d) => AppendClone(arm);
    public override void VisitFieldData(FieldData data, DxpIDocumentContext d) => AppendClone(data);
    public override void VisitAltChunk(AltChunk ac, DxpIDocumentContext d) => AppendClone(ac);
    public override void VisitContentPart(ContentPart cp, DxpIDocumentContext d) => AppendClone(cp);
    public override void VisitUnknown(string context, OpenXmlElement el, DxpIDocumentContext d)
    {
        // Container property children are copied when their owning shell is created.
        // Some legacy included documents contain empty runs directly in block
        // containers. Word tolerates them in the source but rejects them after
        // package reconstruction, so omit the content-free invalid node.
        bool invalidEmptyBlockRun =
            el.LocalName == "r" &&
            el.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main" &&
            !el.HasChildren;
        if (!IsPropertyChild(el) && !invalidEmptyBlockRun)
            AppendClone(el);
    }

    public override void VisitComplexFieldBegin(FieldChar begin, DxpIDocumentContext d) => AppendClone(begin);
    public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d) => AppendClone(instr);
    public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext d) => AppendClone(separate);
    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d) => AppendClone(end);
    public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
    {
        var value = new Text(text);
        if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1])))
            value.Space = SpaceProcessingModeValues.Preserve;
        AppendClone(value);
    }

    public override void VisitMoveFromRangeStart(MoveFromRangeStart value, DxpIDocumentContext d) => AppendClone(value);
    public override void VisitMoveFromRangeEnd(MoveFromRangeEnd value, DxpIDocumentContext d) => AppendClone(value);
    public override void VisitMoveToRangeStart(MoveToRangeStart value, DxpIDocumentContext d) => AppendClone(value);
    public override void VisitMoveToRangeEnd(MoveToRangeEnd value, DxpIDocumentContext d) => AppendClone(value);

    public void Dispose()
    {
        Complete();
        GC.SuppressFinalize(this);
    }

    private IDisposable AppendContainer<T>(T element) where T : OpenXmlCompositeElement
    {
        if (_suppressDepth > 0)
            return SuppressScope();
        if (_parents.Count == 0)
            return DxpDisposable.Empty;

        _parents.Peek().AppendChild(element);
        _parents.Push(element);
        return DxpDisposable.Create(() => _parents.Pop());
    }

    private IDisposable AppendSdtContainer<T>(T source, OpenXmlCompositeElement? sourceContent)
        where T : OpenXmlCompositeElement
    {
        var shell = CloneShell(source);
        if (sourceContent == null)
            return AppendContainer(shell);

        var content = (OpenXmlCompositeElement)sourceContent.CloneNode(false);
        shell.AppendChild(content);
        var outer = AppendContainer(shell);
        if (_suppressDepth > 0 || _parents.Count == 0)
            return outer;

        _parents.Push(content);
        return DxpDisposable.Create(() => {
            _parents.Pop();
            outer.Dispose();
        });
    }

    private IDisposable PushRoot(OpenXmlCompositeElement root)
    {
        int depth = _parents.Count;
        _parents.Push(root);
        return DxpDisposable.Create(() => {
            while (_parents.Count > depth)
                _parents.Pop();
        });
    }

    private IDisposable SuppressScope()
    {
        _suppressDepth++;
        return DxpDisposable.Create(() => _suppressDepth--);
    }

    private void AppendClone(OpenXmlElement element)
    {
        if (_suppressDepth == 0 && _parents.Count > 0)
            _parents.Peek().AppendChild(element.CloneNode(true));
    }

    private static T CloneShell<T>(T source) where T : OpenXmlCompositeElement
    {
        var clone = (T)source.CloneNode(false);
        foreach (var child in source.ChildElements.Where(IsPropertyChild))
            clone.AppendChild(child.CloneNode(true));
        return clone;
    }

    private T CloneShellWithExternalRelationships<T>(T source, DxpIDocumentContext context)
        where T : OpenXmlCompositeElement
    {
        var clone = CloneShell(source);
        RemapExternalRelationships(clone, context.CurrentPart);
        return clone;
    }

    private void AppendCloneWithExternalRelationships(OpenXmlElement source, DxpIDocumentContext context)
    {
        var clone = source.CloneNode(true);
        RemapExternalRelationships(clone, context.CurrentPart);
        if (_suppressDepth == 0 && _parents.Count > 0)
            _parents.Peek().AppendChild(clone);
    }

    private void RemapExternalRelationships(OpenXmlElement clone, OpenXmlPart? sourcePart)
    {
        var destinationPart = _outputDocument?.MainDocumentPart;
        if (sourcePart == null || destinationPart == null || _sourceParts.Contains(sourcePart))
            return;

        var externalById = sourcePart.ExternalRelationships.ToDictionary(rel => rel.Id, StringComparer.Ordinal);
        var hyperlinksById = sourcePart.HyperlinkRelationships.ToDictionary(rel => rel.Id, StringComparer.Ordinal);
        foreach (var element in clone.Descendants().Prepend(clone))
        {
            foreach (var attribute in element.GetAttributes().Where(attribute =>
                         attribute.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/relationships").ToArray())
            {
                if (attribute.Value == null)
                    continue;
                string? importedId = null;
                if (externalById.TryGetValue(attribute.Value, out var relationship))
                    importedId = destinationPart.AddExternalRelationship(relationship.RelationshipType, relationship.Uri).Id;
                else if (hyperlinksById.TryGetValue(attribute.Value, out var hyperlink))
                    importedId = destinationPart.AddHyperlinkRelationship(hyperlink.Uri, hyperlink.IsExternal).Id;
                else if (_importedRelationships.TryGetValue((sourcePart, attribute.Value), out string? existingId))
                    importedId = existingId;
                else
                {
                    OpenXmlPart? relatedPart = null;
                    try { relatedPart = sourcePart.GetPartById(attribute.Value); }
                    catch (ArgumentOutOfRangeException) { }

                    if (relatedPart is ImagePart sourceImage)
                    {
                        var importedImage = destinationPart.AddImagePart(sourceImage.ContentType);
                        using var imageStream = sourceImage.GetStream(FileMode.Open, FileAccess.Read);
                        importedImage.FeedData(imageStream);
                        importedId = destinationPart.GetIdOfPart(importedImage);
                        _importedRelationships[(sourcePart, attribute.Value)] = importedId;
                    }
                }
                if (importedId != null)
                    element.SetAttribute(new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, importedId));
            }
        }
    }

    private static bool IsPropertyChild(OpenXmlElement child)
        => child is ParagraphProperties
            or RunProperties
            or TableProperties
            or TableGrid
            or TableRowProperties
            or TablePropertyExceptions
            or TableCellProperties
            or SdtProperties
            or SdtEndCharProperties
            or CustomXmlProperties
            or RubyProperties;

    private static bool IsInternalNote(Footnote note) => IsInternalNote(note.Type?.Value);
    private static bool IsInternalNote(Endnote note) => IsInternalNote(note.Type?.Value);
    private static bool IsInternalNote(FootnoteEndnoteValues? type)
        => type == FootnoteEndnoteValues.Separator
            || type == FootnoteEndnoteValues.ContinuationSeparator
            || type == FootnoteEndnoteValues.ContinuationNotice;

    private void Complete()
    {
        if (_completed)
            return;
        _completed = true;

        if (_outputDocument == null)
            return;

        NormalizeInvalidBlockMarkup(_outputDocument);
        _outputDocument.MainDocumentPart?.Document?.Save();
        foreach (var part in _outputDocument.MainDocumentPart?.HeaderParts ?? Enumerable.Empty<HeaderPart>())
            part.Header?.Save();
        foreach (var part in _outputDocument.MainDocumentPart?.FooterParts ?? Enumerable.Empty<FooterPart>())
            part.Footer?.Save();
        _outputDocument.MainDocumentPart?.FootnotesPart?.Footnotes?.Save();
        _outputDocument.MainDocumentPart?.EndnotesPart?.Endnotes?.Save();
        _outputDocument.MainDocumentPart?.WordprocessingCommentsPart?.Comments?.Save();
        _outputDocument.Dispose();
        _outputDocument = null;
        _sourceParts.Clear();
        _importedRelationships.Clear();
    }

    private static void NormalizeInvalidBlockMarkup(WordprocessingDocument document)
    {
        IEnumerable<OpenXmlPartRootElement?> roots =
        [
            document.MainDocumentPart?.Document,
            .. document.MainDocumentPart?.HeaderParts.Select(part => part.Header) ?? [],
            .. document.MainDocumentPart?.FooterParts.Select(part => part.Footer) ?? [],
            document.MainDocumentPart?.FootnotesPart?.Footnotes,
            document.MainDocumentPart?.EndnotesPart?.Endnotes,
            document.MainDocumentPart?.WordprocessingCommentsPart?.Comments
        ];

        foreach (var root in roots.OfType<OpenXmlPartRootElement>())
        {
            if (root is Comments)
            {
                foreach (var element in root.Descendants())
                {
                    foreach (var attribute in element.GetAttributes()
                                 .Where(attribute => attribute.NamespaceUri.StartsWith(
                                     "http://schemas.microsoft.com/office/word/2010/",
                                     StringComparison.Ordinal))
                                 .ToArray())
                        element.RemoveAttribute(attribute.LocalName, attribute.NamespaceUri);
                }
            }

            foreach (var tableLook in root.Descendants<TableLook>())
            {
                foreach (string name in new[] { "firstRow", "lastRow", "firstColumn", "lastColumn", "noHBand", "noVBand" })
                    tableLook.RemoveAttribute(name, tableLook.NamespaceUri);
            }

            foreach (var conditionalStyle in root.Descendants<ConditionalFormatStyle>())
            {
                foreach (var attribute in conditionalStyle.GetAttributes()
                             .Where(attribute => attribute.LocalName != "val")
                             .ToArray())
                    conditionalStyle.RemoveAttribute(attribute.LocalName, attribute.NamespaceUri);
            }

            // Legacy smart tags carry recognition metadata only. They are no
            // longer valid paragraph children in current WordprocessingML;
            // retain their runs while removing the obsolete wrappers.
            foreach (var smartTag in root.Descendants()
                         .Where(element => element.LocalName == "smartTag")
                         .Reverse()
                         .ToArray())
            {
                foreach (var child in smartTag.ChildElements.ToArray())
                    smartTag.InsertBeforeSelf(child.CloneNode(true));
                smartTag.Remove();
            }

            var drawingProperties = root
                .Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>()
                .ToArray();
            uint nextDrawingId = drawingProperties
                .Select(properties => properties.Id?.Value ?? 0U)
                .DefaultIfEmpty()
                .Max() + 1;
            var drawingIds = new HashSet<uint>();
            foreach (var properties in drawingProperties)
            {
                uint id = properties.Id?.Value ?? 0U;
                if (!drawingIds.Add(id))
                {
                    properties.Id = nextDrawingId;
                    drawingIds.Add(nextDrawingId++);
                }
            }

            foreach (var paragraph in root.Descendants<Paragraph>())
            {
                // Field replay normally places these inside runs. A field whose
                // original run scope was consumed can surface them directly.
                foreach (var fieldChild in paragraph.ChildElements
                             .Where(element => element is FieldChar or FieldCode)
                             .ToArray())
                    fieldChild.InsertBeforeSelf(new Run(fieldChild.CloneNode(true)));
                foreach (var fieldChild in paragraph.ChildElements
                             .Where(element => element is FieldChar or FieldCode)
                             .ToArray())
                    fieldChild.Remove();
            }


            // A block emitted while an outer paragraph is still active must be
            // promoted to a sibling. Split the outer paragraph so ordering and
            // paragraph formatting on either side are retained.
            foreach (var run in root.Descendants<Run>()
                         .Where(r => r.ChildElements.Any(element => element.LocalName == "p"))
                         .ToArray())
            {
                var containingParagraph = run.Ancestors<Paragraph>().FirstOrDefault();
                if (containingParagraph == null)
                    continue;
                foreach (var nested in run.ChildElements.Where(element => element.LocalName == "p").ToArray())
                    containingParagraph.InsertBeforeSelf(nested.CloneNode(true));
                run.Remove();
            }

            while (true)
            {
                var nestedParagraphs = root.Descendants<Paragraph>()
                    .Where(p => p.ChildElements.Any(element => element.LocalName == "p"))
                    .Reverse()
                    .ToArray();
                if (nestedParagraphs.Length == 0)
                    break;
                foreach (var paragraph in nestedParagraphs)
                    PromoteNestedParagraphs(paragraph);
            }

            uint nextBookmarkId = 1;
            var bookmarkIds = new Dictionary<string, Stack<string>>(StringComparer.Ordinal);
            foreach (var marker in root.Descendants().Where(element => element is BookmarkStart or BookmarkEnd))
            {
                if (marker is BookmarkStart start)
                {
                    string oldId = start.Id?.Value ?? string.Empty;
                    string newId = (nextBookmarkId++).ToString();
                    start.Id = newId;
                    if (!bookmarkIds.TryGetValue(oldId, out var ids))
                        bookmarkIds[oldId] = ids = new Stack<string>();
                    ids.Push(newId);
                }
                else if (marker is BookmarkEnd end &&
                         bookmarkIds.TryGetValue(end.Id?.Value ?? string.Empty, out var ids) &&
                         ids.Count > 0)
                    end.Id = ids.Pop();
                else if (marker is BookmarkEnd unmatchedEnd)
                    unmatchedEnd.Id = (nextBookmarkId++).ToString();
            }

            foreach (var container in root.Descendants<OpenXmlCompositeElement>()
                         .Where(element => element is Body or Header or Footer or TableCell or Footnote or Endnote))
            {
                Paragraph? inlineParagraph = null;
                foreach (var child in container.ChildElements.ToArray())
                {
                    if (child.LocalName is not ("r" or "t"))
                    {
                        inlineParagraph = null;
                        continue;
                    }

                    inlineParagraph ??= child.InsertBeforeSelf(new Paragraph());
                    inlineParagraph.AppendChild(child.LocalName == "r"
                        ? child.CloneNode(true)
                        : new Run(child.CloneNode(true)));
                    child.Remove();
                }
            }
        }
    }

    private static void PromoteNestedParagraphs(Paragraph paragraph)
    {
        var parent = paragraph.Parent as OpenXmlCompositeElement;
        if (parent == null)
            return;

        Paragraph current = (Paragraph)paragraph.CloneNode(false);
        if (paragraph.ParagraphProperties != null)
            current.AppendChild(paragraph.ParagraphProperties.CloneNode(true));

        foreach (var child in paragraph.ChildElements
                     .Where(child => child is not ParagraphProperties)
                     .ToArray())
        {
            if (child.LocalName == "p")
            {
                if (current.ChildElements.Any(element => element is not ParagraphProperties))
                    paragraph.InsertBeforeSelf(current);
                paragraph.InsertBeforeSelf(child.CloneNode(true));
                current = (Paragraph)paragraph.CloneNode(false);
                if (paragraph.ParagraphProperties != null)
                    current.AppendChild(paragraph.ParagraphProperties.CloneNode(true));
            }
            else
                current.AppendChild(child.CloneNode(true));
        }

        if (current.ChildElements.Any(element => element is not ParagraphProperties))
            paragraph.InsertBeforeSelf(current);
        paragraph.Remove();
    }

    private static void CollectParts(OpenXmlPartContainer container, ISet<OpenXmlPart> parts)
    {
        foreach (var pair in container.Parts)
        {
            if (!parts.Add(pair.OpenXmlPart))
                continue;
            CollectParts(pair.OpenXmlPart, parts);
        }
    }
}
