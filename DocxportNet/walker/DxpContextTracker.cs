using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Walker;

/// <summary>
/// Placeholder state layer for the walker pipeline.
/// Currently forwards all events unchanged.
/// </summary>
public sealed class DxpContextTracker : DxpMiddleware
{
    private OpenXmlPart? _mainPart;
    private readonly ILogger? _logger;

    public DxpContextTracker(DxpIVisitor next, ILogger? logger = null) : base(next)
    {
        _logger = logger;
    }

    public override IDisposable VisitDocumentBegin(WordprocessingDocument doc, DxpIDocumentContext documentContext)
    {
        _mainPart = doc.MainDocumentPart;
        IDisposable partScope = DxpDisposable.Empty;
        if (documentContext is IDxpMutableDocumentContext mutable)
            partScope = mutable.PushCurrentPart(_mainPart);

        var inner = _next.VisitDocumentBegin(doc, documentContext);
        return new DxpAfterScope(inner, () => {
            partScope.Dispose();
            CleanupFields(documentContext);
            _mainPart = null;
        });
    }

    public override void VisitText(Text t, DxpIDocumentContext d)
    {
        if (d.CurrentFields.IsInFieldResult)
            _next.VisitComplexFieldCachedResultText(t.Text, d);
        else
            _next.VisitText(t, d);
    }

    public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
    {
        var frame = new FieldFrame { SeenSeparate = true, InResult = true };
        var instruction = fld.Instruction?.Value;
        if (!string.IsNullOrEmpty(instruction))
            frame.InstructionText = instruction;
        d.CurrentFields.FieldStack.Push(frame);

        var inner = _next.VisitSimpleFieldBegin(fld, d);
        return new DxpAfterScope(inner, () => {
            if (d.CurrentFields.FieldStack.Count > 0)
                d.CurrentFields.FieldStack.Pop();
        });
    }

    public override void VisitComplexFieldBegin(FieldChar begin, DxpIDocumentContext d)
    {
        var frame = new FieldFrame { SeenSeparate = false, ResultScope = null, InResult = false };
        d.CurrentFields.FieldStack.Push(frame);
        _next.VisitComplexFieldBegin(begin, d);
    }

    public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
    {
        if (!string.IsNullOrEmpty(text) && d.CurrentFields.Current != null)
        {
            var current = d.CurrentFields.Current;
            current.InstructionText = current.InstructionText == null
                ? text
                : current.InstructionText + text;
        }

        _next.VisitComplexFieldInstruction(instr, text, d);
    }

    public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext d)
    {
        if (d.CurrentFields.FieldStack.Count > 0)
        {
            var top = d.CurrentFields.FieldStack.Pop();
            if (!top.SeenSeparate)
            {
                _next.VisitComplexFieldSeparate(separate, d);
                top.SeenSeparate = true;
                top.InResult = true;
                if (top.ResultScope == null)
                    top.ResultScope = _next.VisitComplexFieldResultBegin(d);
            }
            d.CurrentFields.FieldStack.Push(top);
            return;
        }

        _next.VisitComplexFieldSeparate(separate, d);
    }

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
    {
        if (d.CurrentFields.FieldStack.Count > 0)
        {
            var top = d.CurrentFields.FieldStack.Pop();
            top.InResult = false;
            top.ResultScope?.Dispose();
            _next.VisitComplexFieldEnd(end, d);
            return;
        }

        _next.VisitComplexFieldEnd(end, d);
    }

    public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
    {
        if (d is not IDxpMutableDocumentContext doc)
            return _next.VisitRunBegin(r, d);

        var para = r.Ancestors<Paragraph>().FirstOrDefault();
        var styles = d.Styles;
        var defaultRunStyle = d.DefaultRunStyle;
        var style = para != null ? styles.ResolveRunStyle(para, r) : defaultRunStyle;
        var language = para != null ? styles.ResolveRunLanguage(para, r) : null;

        bool hasRenderable = r.ChildElements.Any(child =>
            child is Text or DeletedText or NoBreakHyphen or TabChar or Break or CarriageReturn or Drawing);
        if (_logger?.IsEnabled(LogLevel.Debug) == true)
        {
            var runText = string.Concat(r.Elements<Text>().Select(t => t.Text));
            bool hasBoldProp = r.RunProperties?.Bold != null;
            _logger.LogDebug(
                "RunBegin: hasRenderable={HasRenderable}, runBoldProp={RunBoldProp}, effectiveBold={EffectiveBold}, text='{RunText}'",
                hasRenderable,
                hasBoldProp,
                style.Bold,
                runText);
        }
        if (para != null && hasRenderable)
            doc.StyleTracker.ApplyStyle(style, d, _next);

        var runScope = doc.PushRun(r, style, language, out _);
        var inner = _next.VisitRunBegin(r, d);
        return new DxpAfterScope(inner, runScope.Dispose);
    }

    public override IDisposable VisitParagraphBegin(Paragraph p, DxpIDocumentContext d, DxpIParagraphContext paragraph)
    {
        if (d is not IDxpMutableDocumentContext doc || paragraph is not DxpParagraphContext ctx)
            return _next.VisitParagraphBegin(p, d, paragraph);

        Deleted? deletedParagraph =
            p.ParagraphProperties?.GetFirstChild<Deleted>() ??
            p.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<Deleted>();
        Inserted? insertedParagraph =
            p.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<Inserted>();
        bool isDeletedParagraph = deletedParagraph != null;
        bool isInsertedParagraph = insertedParagraph != null;
        var changeScope = isDeletedParagraph
            ? doc.PushChangeScope(keepAccept: false, keepReject: true, changeInfo: ResolveChangeInfo(deletedParagraph!, d))
            : isInsertedParagraph
                ? doc.PushChangeScope(keepAccept: true, keepReject: false, changeInfo: ResolveChangeInfo(insertedParagraph!, d))
                : DxpDisposable.Empty;

        var previous = doc.CurrentParagraph;
        doc.CurrentParagraph = ctx;
        var inner = _next.VisitParagraphBegin(p, d, paragraph);
        return new DxpAroundScope(
            inner,
            () => {
                ResetStyle(d);
                doc.CurrentParagraph = previous;
            },
            changeScope.Dispose
        );
    }

    public override IDisposable VisitTableBegin(Table t, DxpTableModel model, DxpIDocumentContext d, DxpITableContext table)
    {
        if (d is not IDxpMutableDocumentContext doc)
        {
            var inner = _next.VisitTableBegin(t, model, d, table);
            return new DxpBeforeScope(inner, () => ResetStyle(d));
        }

        var previousTable = doc.CurrentTable;
        var previousModel = doc.CurrentTableModel;
        doc.CurrentTable = table;
        doc.CurrentTableModel = model;
        var innerScope = _next.VisitTableBegin(t, model, d, table);
        return new DxpBeforeScope(innerScope, () => {
            ResetStyle(d);
            doc.CurrentTable = previousTable;
            doc.CurrentTableModel = previousModel;
        });
    }

    public override IDisposable VisitTableRowBegin(TableRow tr, DxpITableRowContext row, DxpIDocumentContext d)
    {
        if (d is not IDxpMutableDocumentContext doc)
            return _next.VisitTableRowBegin(tr, row, d);

        var previousRow = doc.CurrentTableRow;
        doc.CurrentTableRow = row;
        var inner = _next.VisitTableRowBegin(tr, row, d);
        return new DxpAfterScope(inner, () => doc.CurrentTableRow = previousRow);
    }

    public override IDisposable VisitTableCellBegin(TableCell tc, DxpITableCellContext cell, DxpIDocumentContext d)
    {
        if (d is not IDxpMutableDocumentContext doc)
            return _next.VisitTableCellBegin(tc, cell, d);

        var previousCell = doc.CurrentTableCell;
        doc.CurrentTableCell = cell;
        var inner = _next.VisitTableCellBegin(tc, cell, d);
        return new DxpAfterScope(inner, () => doc.CurrentTableCell = previousCell);
    }

    public override IDisposable VisitSectionHeaderBegin(Header hdr, object value, DxpIDocumentContext d)
    {
        var part = (value as DxpHeaderFooterContext)?.Part;
        IDisposable partScope = DxpDisposable.Empty;
        if (d is IDxpMutableDocumentContext doc)
            partScope = doc.PushCurrentPart(part ?? doc.CurrentPart);

        var inner = _next.VisitSectionHeaderBegin(hdr, value, d);
        return new DxpAroundScope(inner, () => ResetStyle(d), partScope.Dispose);
    }

    public override IDisposable VisitSectionFooterBegin(Footer ftr, object value, DxpIDocumentContext d)
    {
        var part = (value as DxpHeaderFooterContext)?.Part;
        IDisposable partScope = DxpDisposable.Empty;
        if (d is IDxpMutableDocumentContext doc)
            partScope = doc.PushCurrentPart(part ?? doc.CurrentPart);

        var inner = _next.VisitSectionFooterBegin(ftr, value, d);
        return new DxpAroundScope(inner, () => ResetStyle(d), partScope.Dispose);
    }

    public override IDisposable VisitSectionBegin(SectionProperties properties, SectionLayout layout, DxpIDocumentContext d)
    {
        if (d is not IDxpMutableDocumentContext doc)
            return _next.VisitSectionBegin(properties, layout, d);

        var previous = doc.CurrentSection;
        doc.EnterSection(properties, layout);
        var inner = _next.VisitSectionBegin(properties, layout, d);
        return new DxpAfterScope(inner, () => doc.CurrentSection = previous);
    }

    public override IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target, DxpIDocumentContext d)
    {
        var inner = _next.VisitHyperlinkBegin(link, target, d);
        return new DxpBeforeScope(inner, () => ResetStyle(d));
    }

    public override IDisposable VisitCommentBegin(DxpCommentInfo c, DxpCommentThread thread, DxpIDocumentContext d)
    {
        IDisposable partScope = DxpDisposable.Empty;
        if (d is IDxpMutableDocumentContext doc)
            partScope = doc.PushCurrentPart(c.Part ?? doc.CurrentPart);

        var inner = _next.VisitCommentBegin(c, thread, d);
        return new DxpAroundScope(inner, () => ResetStyle(d), partScope.Dispose);
    }

    public override IDisposable VisitRubyBegin(Ruby ruby, DxpIDocumentContext d)
    {
        if (d is not IDxpMutableDocumentContext doc)
            return _next.VisitRubyBegin(ruby, d);

        var pr = ruby.GetFirstChild<RubyProperties>();
        var scope = doc.PushRuby(ruby, pr, out _);
        var inner = _next.VisitRubyBegin(ruby, d);
        return new DxpAfterScope(inner, scope.Dispose);
    }

    public override IDisposable VisitSmartTagRunBegin(OpenXmlUnknownElement smart, string elementName, string elementUri, DxpIDocumentContext d)
    {
        if (d is not IDxpMutableDocumentContext doc)
            return _next.VisitSmartTagRunBegin(smart, elementName, elementUri, d);

        var wNs = smart.NamespaceUri;
        var smartTagPr = smart.ChildElements
            .OfType<OpenXmlUnknownElement>()
            .FirstOrDefault(e => e.LocalName == "smartTagPr" && e.NamespaceUri == wNs);
        var attrs = smartTagPr != null ? smartTagPr.Elements<CustomXmlAttribute>().ToList() : new List<CustomXmlAttribute>();

        var scope = doc.PushSmartTag(smart, elementName, elementUri, attrs, out _);
        var inner = _next.VisitSmartTagRunBegin(smart, elementName, elementUri, d);
        return new DxpAfterScope(inner, scope.Dispose);
    }

    public override IDisposable VisitCustomXmlRunBegin(CustomXmlRun cxr, DxpIDocumentContext d)
    {
        if (d is not IDxpMutableDocumentContext doc)
            return _next.VisitCustomXmlRunBegin(cxr, d);

        var scope = doc.PushCustomXml(cxr, cxr.CustomXmlProperties, out _);
        var inner = _next.VisitCustomXmlRunBegin(cxr, d);
        return new DxpAfterScope(inner, scope.Dispose);
    }

    public override IDisposable VisitCustomXmlBlockBegin(CustomXmlBlock cx, DxpIDocumentContext d)
    {
        if (d is not IDxpMutableDocumentContext doc)
            return _next.VisitCustomXmlBlockBegin(cx, d);

        var scope = doc.PushCustomXml(cx, cx.CustomXmlProperties, out _);
        var inner = _next.VisitCustomXmlBlockBegin(cx, d);
        return new DxpAfterScope(inner, scope.Dispose);
    }

    public override IDisposable VisitSdtContentRunBegin(SdtContentRun content, DxpIDocumentContext d)
    {
        if (d is not IDxpMutableDocumentContext doc)
            return _next.VisitSdtContentRunBegin(content, d);

        if (content.Parent is not SdtRun sdtRun)
            return _next.VisitSdtContentRunBegin(content, d);

        var scope = doc.PushSdt(sdtRun, sdtRun.SdtProperties, sdtRun.SdtEndCharProperties, out _);
        var inner = _next.VisitSdtContentRunBegin(content, d);
        return new DxpAfterScope(inner, scope.Dispose);
    }

    public override IDisposable VisitSdtContentBlockBegin(SdtContentBlock content, DxpIDocumentContext d)
    {
        if (d is not IDxpMutableDocumentContext doc)
            return _next.VisitSdtContentBlockBegin(content, d);

        if (content.Parent is not SdtBlock sdtBlock)
            return _next.VisitSdtContentBlockBegin(content, d);

        var scope = doc.PushSdt(sdtBlock, sdtBlock.SdtProperties, sdtBlock.SdtEndCharProperties, out _);
        var inner = _next.VisitSdtContentBlockBegin(content, d);
        return new DxpAfterScope(inner, scope.Dispose);
    }

    public override IDisposable VisitDeletedRunBegin(DeletedRun dr, DxpIDocumentContext d)
    {
        if (d is not IDxpMutableDocumentContext doc)
            return _next.VisitDeletedRunBegin(dr, d);

        var scope = doc.PushChangeScope(keepAccept: false, keepReject: true, changeInfo: ResolveChangeInfo(dr, d));
        var inner = _next.VisitDeletedRunBegin(dr, d);
        return new DxpAfterScope(inner, scope.Dispose);
    }

    public override IDisposable VisitInsertedRunBegin(InsertedRun ir, DxpIDocumentContext d)
    {
        if (d is not IDxpMutableDocumentContext doc)
            return _next.VisitInsertedRunBegin(ir, d);

        var scope = doc.PushChangeScope(keepAccept: true, keepReject: false, changeInfo: ResolveChangeInfo(ir, d));
        var inner = _next.VisitInsertedRunBegin(ir, d);
        return new DxpAfterScope(inner, scope.Dispose);
    }

    public override IDisposable VisitDeletedBegin(Deleted del, DxpIDocumentContext d)
    {
        if (d is not IDxpMutableDocumentContext doc)
            return _next.VisitDeletedBegin(del, d);

        var scope = doc.PushChangeScope(keepAccept: false, keepReject: true, changeInfo: ResolveChangeInfo(del, d));
        var inner = _next.VisitDeletedBegin(del, d);
        return new DxpAfterScope(inner, scope.Dispose);
    }

    public override IDisposable VisitInsertedBegin(Inserted ins, DxpIDocumentContext d)
    {
        if (d is not IDxpMutableDocumentContext doc)
            return _next.VisitInsertedBegin(ins, d);

        var scope = doc.PushChangeScope(keepAccept: true, keepReject: false, changeInfo: ResolveChangeInfo(ins, d));
        var inner = _next.VisitInsertedBegin(ins, d);
        return new DxpAfterScope(inner, scope.Dispose);
    }

    public override IDisposable VisitFootnoteBegin(Footnote footnote, DxpIFootnoteContext footnoteContext, DxpIDocumentContext d)
    {
        if (d is not IDxpMutableDocumentContext doc)
        {
            var innerScope = _next.VisitFootnoteBegin(footnote, footnoteContext, d);
            return new DxpBeforeScope(innerScope, () => ResetStyle(d));
        }

        var previous = doc.CurrentFootnote;
        doc.CurrentFootnote = footnoteContext;
        var partScope = doc.PushCurrentPart((_mainPart as MainDocumentPart)?.FootnotesPart);
        var inner = _next.VisitFootnoteBegin(footnote, footnoteContext, d);
        return new DxpAroundScope(inner, () => ResetStyle(d), () => {
            partScope.Dispose();
            doc.CurrentFootnote = previous;
        });
    }

    public override IDisposable VisitEndnoteBegin(Endnote endnote, long id, int index, DxpIDocumentContext d)
    {
        if (d is not IDxpMutableDocumentContext doc)
        {
            var innerScope = _next.VisitEndnoteBegin(endnote, id, index, d);
            return new DxpBeforeScope(innerScope, () => ResetStyle(d));
        }

        var scope = doc.PushFootnote(id, index, out _);
        var partScope = doc.PushCurrentPart((_mainPart as MainDocumentPart)?.EndnotesPart);
        var inner = _next.VisitEndnoteBegin(endnote, id, index, d);
        return new DxpAroundScope(inner, () => ResetStyle(d), () => {
            partScope.Dispose();
            scope.Dispose();
        });
    }

    public override void VisitBookmarkStart(BookmarkStart bs, DxpIDocumentContext d)
    {
        _next.VisitBookmarkStart(bs, d);
        ResetStyle(d);
    }

    public override void VisitBookmarkEnd(BookmarkEnd be, DxpIDocumentContext d)
    {
        _next.VisitBookmarkEnd(be, d);
        ResetStyle(d);
    }

    private void ResetStyle(DxpIDocumentContext d)
    {
        if (d is IDxpMutableDocumentContext doc)
            doc.StyleTracker.ResetStyle(d, _next);
    }

    private static void CleanupFields(DxpIDocumentContext d)
    {
        var stack = d.CurrentFields.FieldStack;
        while (stack.Count > 0)
        {
            var frame = stack.Pop();
            frame.ResultScope?.Dispose();
        }
    }

    private static DxpChangeInfo ResolveChangeInfo(OpenXmlElement change, DxpIDocumentContext d)
    {
        var current = d.CurrentChangeInfo;
        string? author = null;
        DateTime? date = null;

        switch (change)
        {
            case TrackChangeType track:
                author = track.Author?.Value;
                date = track.Date?.Value;
                break;
            case RunTrackChangeType runTrack:
                author = runTrack.Author?.Value;
                date = runTrack.Date?.Value;
                break;
        }

        return new DxpChangeInfo(author ?? current.Author, date ?? current.Date);
    }

    private sealed class DxpAfterScope : IDisposable
    {
        private readonly IDisposable _inner;
        private readonly Action _afterDispose;
        private bool _disposed;

        public DxpAfterScope(IDisposable inner, Action afterDispose)
        {
            _inner = inner;
            _afterDispose = afterDispose;
        }

        public void Dispose()
        {
            if (_disposed)
                return;
            _disposed = true;
            _inner.Dispose();
            _afterDispose();
        }
    }

    private sealed class DxpBeforeScope : IDisposable
    {
        private readonly IDisposable _inner;
        private readonly Action _beforeDispose;
        private bool _disposed;

        public DxpBeforeScope(IDisposable inner, Action beforeDispose)
        {
            _inner = inner;
            _beforeDispose = beforeDispose;
        }

        public void Dispose()
        {
            if (_disposed)
                return;
            _disposed = true;
            _beforeDispose();
            _inner.Dispose();
        }
    }

    private sealed class DxpAroundScope : IDisposable
    {
        private readonly IDisposable _inner;
        private readonly Action _beforeDispose;
        private readonly Action _afterDispose;
        private bool _disposed;

        public DxpAroundScope(IDisposable inner, Action beforeDispose, Action afterDispose)
        {
            _inner = inner;
            _beforeDispose = beforeDispose;
            _afterDispose = afterDispose;
        }

        public void Dispose()
        {
            if (_disposed)
                return;
            _disposed = true;
            _beforeDispose();
            _inner.Dispose();
            _afterDispose();
        }
    }
}
