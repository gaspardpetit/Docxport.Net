using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Middleware;
using Microsoft.Extensions.Logging;
using System.Text;

namespace DocxportNet.Fields.Frames;

internal abstract class DxpSetFieldFrame : DxpMiddleware, DxpIFieldEvalFrame
{
    public bool SuppressContent { get; set; }
    public bool Evaluated { get; set; }
    public bool SeenSeparate { get; set; }
    public bool InResult { get; set; }
    public string? InstructionText { get; set; }
    public RunProperties? CodeRunProperties { get; set; }
    public Run? CodeRun { get; set; }
    public List<Run?>? CachedResultRuns { get; set; }
    public DxpIFCaptureState? IfState { get; set; }

    public override DxpIVisitor Next => null!;
    public DxpFieldEvalContext EvalContext { get; }

    protected readonly DxpFieldEval _eval;
    protected readonly DxpFieldParser _parser = new();

    protected DxpSetFieldFrame(DxpFieldEval eval, DxpFieldEvalContext evalContext, ILogger? logger)
        : base()
    {
        _eval = eval ?? throw new ArgumentNullException(nameof(eval));
        EvalContext = evalContext ?? throw new ArgumentNullException(nameof(evalContext));
    }

    public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
    {
        if (string.IsNullOrEmpty(text) || InResult)
            return;
    }

    public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext d)
    {
        if (!SeenSeparate)
        {
            SeenSeparate = true;
            InResult = true;
        }
    }

    public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
    {
    }

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
    {
        EvaluateSet(d);
    }

    public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
    {
        EvaluateSet(d);
        return DxpDisposable.Empty;
    }

    protected override bool ShouldForwardContent(DxpIDocumentContext d)
        => false;

    protected abstract bool TryGetResultText(DxpIDocumentContext d, out string? text);

    protected void EvaluateSet(DxpIDocumentContext d)
    {
        if (Evaluated)
            return;
        if (string.IsNullOrWhiteSpace(InstructionText))
            return;

        Evaluated = true;
        SuppressContent = true;

        var parse = _parser.Parse(InstructionText!);
        var argsText = parse.Ast.ArgumentsText;
        if (!TryGetFirstToken(argsText, out var setName))
            return;

        if (TryGetResultText(d, out var cachedText) && cachedText != null)
        {
            EvalContext.SetBookmarkNodes(setName, DxpFieldNodeBuffer.FromText(cachedText));
            return;
        }

        if (d.DocumentIndex.BookmarkNodes != null &&
            d.DocumentIndex.BookmarkNodes.TryGetValue(setName, out var existingNodes))
        {
            EvalContext.SetBookmarkNodes(setName, existingNodes);
            return;
        }

        var setResult = _eval.EvalAsync(new DxpFieldInstruction(InstructionText!), d).GetAwaiter().GetResult();
        var text = setResult.Text ?? string.Empty;
        EvalContext.SetBookmarkNodes(setName, DxpFieldNodeBuffer.FromText(text));
    }

    protected static bool TryGetFirstToken(string? argsText, out string token)
    {
        token = string.Empty;
        if (string.IsNullOrWhiteSpace(argsText))
            return false;

        var tokens = TokenizeArgs(argsText);
        if (tokens.Count == 0)
            return false;
        token = tokens[0];
        return true;
    }

    protected static List<string> TokenizeArgs(string? text)
    {
        var tokens = new List<string>();
        if (string.IsNullOrEmpty(text))
            return tokens;
        bool inQuote = false;
        bool justClosedQuote = false;
        var current = new StringBuilder();
        for (int i = 0; i < text?.Length; i++)
        {
            char ch = text[i];
            if (ch == '"')
            {
                if (inQuote && i > 0 && text[i - 1] == '\\')
                {
                    if (current.Length > 0)
                        current.Length -= 1;
                    current.Append('"');
                    continue;
                }
                inQuote = !inQuote;
                if (!inQuote)
                {
                    justClosedQuote = true;
                    if (current.Length > 0)
                    {
                        tokens.Add(current.ToString());
                        current.Clear();
                    }
                }
                continue;
            }

            if (!inQuote && char.IsWhiteSpace(ch))
            {
                if (current.Length > 0 || justClosedQuote)
                {
                    if (current.Length > 0)
                    {
                        tokens.Add(current.ToString());
                        current.Clear();
                    }
                    justClosedQuote = false;
                }
                continue;
            }

            current.Append(ch);
            justClosedQuote = false;
        }

        if (current.Length > 0)
            tokens.Add(current.ToString());
        return tokens;
    }
}
