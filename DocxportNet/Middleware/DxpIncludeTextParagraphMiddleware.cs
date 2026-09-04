using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Fields;
using DocxportNet.Fields.Eval;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Middleware;

internal sealed class DxpIncludeTextParagraphMiddleware : DxpMiddleware,
    IDxpStructuredFieldSpliceCollector,
    IDxpEmbeddedWalkCompletion
{
    private readonly DxpIVisitor _next;
    private readonly DxpFieldEvalContext _context;
    private readonly ILogger? _logger;
    private readonly DxpEvalFieldNodeBufferRecorder _recorder = new();
    private DxpFieldNodeBuffer? _buffer;
    private Paragraph? _paragraph;
    private DxpIDocumentContext? _documentContext;
    private IDxpStructuredFieldSpliceCollector? _previousCollector;
    private bool _rootParagraphClosed;

    public DxpIncludeTextParagraphMiddleware(DxpIVisitor next, DxpFieldEval eval, ILogger? logger)
    {
        _next = next;
        _context = eval.Context;
        _logger = logger;
    }

    public override DxpIVisitor Next => _buffer == null ? _next : _recorder;

    public bool HasPendingEmbeddedWork(DxpIDocumentContext documentContext)
        => (_buffer != null && ReferenceEquals(_documentContext, documentContext))
        || (_next is IDxpEmbeddedWalkCompletion completion &&
            completion.HasPendingEmbeddedWork(documentContext));

    public override IDisposable VisitParagraphBegin(Paragraph p, DxpIDocumentContext d, DxpIParagraphContext paragraph)
    {
        if (_buffer != null || !ContainsBlockProducingField(p))
            return base.VisitParagraphBegin(p, d, paragraph);

        _buffer = new DxpFieldNodeBuffer();
        _recorder.Reset(_buffer);
        _paragraph = p;
        _documentContext = d;
        _previousCollector = _context.StructuredFieldSpliceCollector;
        _context.StructuredFieldSpliceCollector = this;
        _rootParagraphClosed = false;
        _logger?.LogDebug("Capturing field paragraph for block-aware replay.");

        return DocxportNet.Core.DxpDisposable.Create(Flush);
    }

    public bool Record(DxpIncludeTextExpansion expansion)
    {
        if (_buffer == null)
            return false;
        _recorder.RecordInclude(expansion);
        Complete();
        return true;
    }

    public bool Record(DxpFieldNodeBuffer buffer)
    {
        if (_buffer == null)
            return false;
        _recorder.Record(buffer);
        Complete();
        return true;
    }

    public void Complete()
    {
        if (_rootParagraphClosed)
            Flush(force: true);
    }

    public void CompleteEmbeddedWalk(DxpIDocumentContext documentContext)
    {
        if (_buffer != null && ReferenceEquals(_documentContext, documentContext))
            Complete();
        if (_next is IDxpEmbeddedWalkCompletion completion)
            completion.CompleteEmbeddedWalk(documentContext);
    }

    private void Flush() => Flush(force: false);

    private void Flush(bool force)
    {
        if (!force && _context.FieldDepth > 0)
        {
            _rootParagraphClosed = true;
            return;
        }
        var buffer = _buffer;
        var paragraph = _paragraph;
        var documentContext = _documentContext;
        _context.StructuredFieldSpliceCollector = _previousCollector;
        _buffer = null;
        _paragraph = null;
        _documentContext = null;
        _previousCollector = null;
        _rootParagraphClosed = false;

        if (buffer == null || paragraph == null || documentContext == null)
            return;

        var parts = buffer.SplitIncludeTextExpansions();
        _logger?.LogDebug("Replaying field paragraph with {PartCount} splice parts and {ExpansionCount} INCLUDETEXT expansions.", parts.Count, parts.Count(part => part.Expansion != null));
        if (parts.All(part => part.Expansion == null))
        {
            ReplayPart(documentContext, paragraph, buffer);
            return;
        }
        if (parts.Count == 3 && parts[1].Expansion != null)
        {
            parts[1].Expansion!.Emit(_next, documentContext, paragraph, parts[0].Inline, parts[2].Inline);
            return;
        }
        if (parts.Count == 2)
        {
            var firstExpansion = parts[0].Expansion;
            var secondExpansion = parts[1].Expansion;
            if (firstExpansion != null)
                firstExpansion.Emit(_next, documentContext, paragraph, null, parts[1].Inline);
            else if (secondExpansion != null)
                secondExpansion.Emit(_next, documentContext, paragraph, parts[0].Inline, null);
            else
                buffer.Replay(_next, documentContext);
            return;
        }
        if (parts.Count == 1 && parts[0].Expansion is { } onlyExpansion)
        {
            onlyExpansion.Emit(_next, documentContext, paragraph, null, null);
            return;
        }

        // Multiple INCLUDETEXT fields in one paragraph are uncommon. Replay each part in order
        // without nesting child blocks in an open parent paragraph. Each inline segment gets a
        // parent-formatted paragraph; exact seam coalescing across multiple child bodies is deferred.
        foreach (var part in parts)
        {
            if (part.Inline != null)
                ReplayPart(documentContext, paragraph, part.Inline);
            else
                part.Expansion?.Emit(_next, documentContext, paragraph, null, null);
        }
    }

    private void ReplayPart(DxpIDocumentContext context, Paragraph paragraph, DxpFieldNodeBuffer buffer)
    {
        if (buffer.HasBlockRoots)
            buffer.Replay(_next, context);
        else
            context.Walker.ReplayBufferedParentParagraph(paragraph, context, _next, buffer);
    }

    private static bool ContainsBlockProducingField(Paragraph paragraph)
    {
        string complexInstructions = string.Concat(
            paragraph.Descendants<FieldCode>().Select(static code => code.Text));
        if (DxpFieldInstructionClassifier.IsIncludeTextInstruction(complexInstructions) ||
            DxpFieldInstructionClassifier.IsDatabaseInstruction(complexInstructions) ||
            complexInstructions.Contains("INCLUDETEXT", StringComparison.OrdinalIgnoreCase) ||
            complexInstructions.Contains("DATABASE", StringComparison.OrdinalIgnoreCase))
            return true;

        return paragraph.Descendants<SimpleField>().Any(field =>
        {
            string instruction = field.Instruction?.Value ?? string.Empty;
            return DxpFieldInstructionClassifier.IsIncludeTextInstruction(instruction) ||
                DxpFieldInstructionClassifier.IsDatabaseInstruction(instruction);
        });
    }
}
