using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Fields.Eval;
using DocxportNet.Fields.Semantic;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpIFCaptureState
{
    public int TokenIndex = 0;
    public bool FieldTypeConsumed;
    public bool InQuote;
    public int BraceDepth;
    public bool JustClosedQuote;
    public bool TrueBranchObserved;
    public bool FalseBranchObserved;
    public readonly StringBuilder CurrentToken = new();
    public readonly DxpFieldNodeBuffer TrueBuffer = new();
    public readonly DxpFieldNodeBuffer FalseBuffer = new();
    public readonly List<DxpSemanticBranchPart> TrueSemanticParts = new();
    public readonly List<DxpSemanticBranchPart> FalseSemanticParts = new();
    public readonly DxpEvalFieldNodeBufferRecorder Recorder = new();
    private DxpFieldNodeBuffer? _paragraphOwner;
    private DxpFieldNodeBuffer? _paragraphBuffer;

    public DxpFieldNodeBuffer? GetCurrentBuffer()
    {
        var root = GetCurrentRootBuffer();
        return ReferenceEquals(root, _paragraphOwner) ? _paragraphBuffer : root;
    }

    public void MarkCurrentBranchObserved()
    {
        if (TokenIndex == 3)
            TrueBranchObserved = true;
        else if (TokenIndex == 4)
            FalseBranchObserved = true;
    }

    public bool WasBranchObserved(bool condition)
        => condition ? TrueBranchObserved : FalseBranchObserved;

    public void BeginParagraph(Paragraph paragraph)
    {
        var root = GetCurrentRootBuffer();
        // A branch that starts on this instruction paragraph is still an inline
        // field result at the field's original insertion point. Only a paragraph
        // boundary encountered after branch content has begun belongs to the result.
        _paragraphOwner = root != null && !root.IsEmpty ? root : null;
        _paragraphBuffer = _paragraphOwner?.BeginParagraph(paragraph);
        if (_paragraphOwner != null)
            GetCurrentSemanticParts()?.Add(new DxpSemanticBranchParagraphStart());
    }

    public bool AppendStructuredResult(DxpFieldNodeBuffer buffer)
    {
        var root = GetCurrentRootBuffer();
        if (root == null)
            return false;
        MarkCurrentBranchObserved();
        root.Append(buffer);
        _paragraphOwner = null;
        _paragraphBuffer = null;
        return true;
    }

    public bool AppendDeferredAction(Action<DocxportNet.API.DxpIVisitor, DocxportNet.API.DxpIDocumentContext> action)
    {
        var root = GetCurrentRootBuffer();
        if (root == null)
            return false;
        MarkCurrentBranchObserved();
        root.AddDeferredAction(action);
        _paragraphOwner = null;
        _paragraphBuffer = null;
        return true;
    }

    public bool AppendDeferredField(DxpDeferredField field)
    {
        List<DxpSemanticBranchPart>? parts = GetCurrentSemanticParts();
        if (parts == null)
            return false;
        parts.Add(new DxpSemanticBranchField(field));
        return AppendDeferredAction((visitor, context) => field.Replay(visitor, context));
    }

    public void AppendSemanticText(DxpFieldNodeBuffer target, string text)
    {
        DxpFieldNodeBuffer semanticOwner = ReferenceEquals(target, _paragraphBuffer) && _paragraphOwner != null
            ? _paragraphOwner
            : target;
        List<DxpSemanticBranchPart>? parts = ReferenceEquals(semanticOwner, TrueBuffer)
            ? TrueSemanticParts
            : ReferenceEquals(semanticOwner, FalseBuffer)
                ? FalseSemanticParts
                : GetCurrentSemanticParts();
        if (parts == null || string.IsNullOrEmpty(text))
            return;
        if (parts.Count > 0 && parts[^1] is DxpSemanticBranchText previous)
            parts[^1] = new DxpSemanticBranchText(previous.Text + text);
        else
            parts.Add(new DxpSemanticBranchText(text));
    }

    public void CompleteDeferredFieldToken()
    {
        if (InQuote)
            return;
        TokenIndex++;
        CurrentToken.Clear();
        JustClosedQuote = false;
    }

    private DxpFieldNodeBuffer? GetCurrentRootBuffer()
    {
        return TokenIndex switch
        {
            3 => TrueBuffer,
            4 => FalseBuffer,
            _ => null
        };
    }

    private List<DxpSemanticBranchPart>? GetCurrentSemanticParts()
        => TokenIndex switch
        {
            3 => TrueSemanticParts,
            4 => FalseSemanticParts,
            _ => null
        };
}
