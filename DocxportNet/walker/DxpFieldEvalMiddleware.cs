using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Fields.Formatting;
using DocxportNet.Core;
using DocxportNet.Fields;
using DocxportNet.Fields.Resolution;
using DocxportNet.Visitors;
using Microsoft.Extensions.Logging;
using System.Globalization;
using System.Linq;
using System.Text;

namespace DocxportNet.Walker;

public sealed class DxpFieldEvalMiddleware : DxpMiddleware
{
    private sealed class FieldEvalFrameState
    {
        public bool SuppressContent;
        public bool Evaluated;
        public bool IsComplex;
        public bool SeenSeparate;
        public bool InResult;
        public string? InstructionText;
        public bool InstructionEmitted;
        public RunProperties? CodeRunProperties;
        public List<RunProperties?>? CachedResultRunProperties;
        public IfCaptureState? IfState;
    }

    private sealed class IfCaptureState
    {
        public int TokenIndex = 0;
        public bool FieldTypeConsumed;
        public bool InQuote;
        public int BraceDepth;
        public bool JustClosedQuote;
        public readonly StringBuilder CurrentToken = new();
        public readonly DxpFieldNodeBuffer TrueBuffer = new();
        public readonly DxpFieldNodeBuffer FalseBuffer = new();
        public readonly DxpFieldNodeBufferRecorder Recorder = new();

        public DxpFieldNodeBuffer? GetCurrentBuffer()
        {
            return TokenIndex switch {
                3 => TrueBuffer,
                4 => FalseBuffer,
                _ => null
            };
        }
    }

    private readonly DxpFieldEval _eval;
    private readonly DxpFieldEvalContext _context;
    private readonly DxpFieldEvalMode _mode;
    private readonly bool _includeDocumentProperties;
    private readonly bool _includeCustomProperties;
    private readonly Func<DateTimeOffset>? _nowProvider;
    private readonly ILogger? _logger;
    private readonly DxpFieldParser _parser = new();
    private bool _initialized;
    private int _paragraphOrder;
    private readonly Stack<FieldEvalFrameState> _fieldFrames = new();
    private FieldEvalFrameState? _outerFrame;

    public DxpFieldEvalMiddleware(
        DxpIVisitor next,
        DxpFieldEval eval,
        DxpFieldEvalMode mode = DxpFieldEvalMode.Evaluate,
        bool includeDocumentProperties = true,
        bool includeCustomProperties = false,
        Func<DateTimeOffset>? nowProvider = null,
        ILogger? logger = null)
        : base(next)
    {
        _eval = eval ?? throw new ArgumentNullException(nameof(eval));
        _context = _eval.Context;
        _mode = mode;
        _includeDocumentProperties = includeDocumentProperties;
        _includeCustomProperties = includeCustomProperties;
        _nowProvider = nowProvider;
        _logger = logger;
    }

    public override IDisposable VisitDocumentBegin(WordprocessingDocument doc, DxpIDocumentContext documentContext)
    {
        if (!_initialized)
        {
            _context.InitFromDocumentContext(documentContext, _includeDocumentProperties, _includeCustomProperties);
            if (_nowProvider != null)
                _context.SetNow(_nowProvider);
            _context.TableResolver ??= new DxpWalkerTableResolver(documentContext);
            _context.RefResolver ??= new DocxportNet.Fields.Resolution.DxpRefIndexResolver(
                documentContext.DocumentIndex.RefIndex,
                () => _context.CurrentDocumentOrder);
            var bookmarkNodes = DocxportNet.Fields.Resolution.DxpBookmarkNodeExtractor.Extract(doc, _logger);
            foreach (var kvp in bookmarkNodes)
                _context.SetBookmarkNodes(kvp.Key, kvp.Value);
            _initialized = true;
        }

        _paragraphOrder = 0;
        return _next.VisitDocumentBegin(doc, documentContext);
    }

    protected override bool ShouldForwardContent(DxpIDocumentContext d)
    {
        if (_mode == DxpFieldEvalMode.Cache)
        {
            if (_outerFrame == null)
                return true;
            if (!_outerFrame.InResult)
                return false;
            if (_outerFrame.SuppressContent)
                return false;
            return !IsNestedField;
        }

        if (!IsInFieldResult)
            return true;
        if (IsNestedField)
            return false;
        if (_fieldFrames.Count > 0 && _fieldFrames.Peek().SuppressContent)
            return false;
        return true;
    }

    private FieldEvalFrameState? CurrentField => _fieldFrames.Count > 0 ? _fieldFrames.Peek() : null;
    private bool IsNestedField => _fieldFrames.Count > 1;
    private bool IsInFieldResult => CurrentField?.InResult == true;

    private bool ShouldDeferEvaluation(FieldEvalFrameState frame)
    {
        if (frame.InstructionText == null)
            return false;
        var parse = _parser.Parse(frame.InstructionText);
        var fieldType = parse.Ast.FieldType ?? string.Empty;
        if (!fieldType.Equals("REF", StringComparison.OrdinalIgnoreCase) &&
            !fieldType.Equals("DOCVARIABLE", StringComparison.OrdinalIgnoreCase))
            return false;
        if (!TryGetCharOrMergeFormat(parse.Ast.FormatSpecs, out _, out var hasMergeFormat))
            return false;
        return hasMergeFormat;
    }

    private static bool StartsWithIf(string instruction)
    {
        var trimmed = instruction.TrimStart();
        if (!trimmed.StartsWith("IF", StringComparison.OrdinalIgnoreCase))
            return false;
        return trimmed.Length == 2 || char.IsWhiteSpace(trimmed[2]);
    }

    private void ProcessIfInstructionSegment(IfCaptureState state, string text, RunProperties? runProps)
    {
        if (string.IsNullOrEmpty(text))
            return;

        DxpFieldNodeBuffer? currentTarget = state.GetCurrentBuffer();
        var bufferText = new StringBuilder();

        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];

            if (ch == '"' && state.InQuote && i > 0 && text[i - 1] == '\\')
            {
                if (state.CurrentToken.Length > 0)
                    state.CurrentToken.Length -= 1;
                state.CurrentToken.Append('"');
                if (state.GetCurrentBuffer() != null && state.BraceDepth == 0)
                    bufferText.Append('"');
                continue;
            }

            if (ch == '"')
            {
                state.InQuote = !state.InQuote;
                if (!state.InQuote)
                    state.JustClosedQuote = true;
                continue;
            }

            if (!state.InQuote && ch == '{')
            {
                state.BraceDepth++;
                state.CurrentToken.Append(ch);
                continue;
            }

            if (state.BraceDepth > 0)
            {
                state.CurrentToken.Append(ch);
                if (ch == '{')
                    state.BraceDepth++;
                else if (ch == '}')
                    state.BraceDepth--;
                continue;
            }

            if (!state.InQuote && char.IsWhiteSpace(ch))
            {
                if (state.CurrentToken.Length > 0 || state.JustClosedQuote)
                {
                    var tokenText = state.CurrentToken.ToString();
                    if (!state.FieldTypeConsumed && tokenText.Equals("IF", StringComparison.OrdinalIgnoreCase))
                    {
                        state.FieldTypeConsumed = true;
                        state.TokenIndex = 0;
                    }
                    else
                    {
                        state.TokenIndex++;
                    }
                    state.CurrentToken.Clear();
                    state.JustClosedQuote = false;
                }
                continue;
            }

            state.CurrentToken.Append(ch);
            state.JustClosedQuote = false;
            var nextTarget = state.GetCurrentBuffer();
            if (nextTarget != currentTarget)
            {
                if (currentTarget != null && bufferText.Length > 0)
                {
                    AppendBufferText(currentTarget, bufferText.ToString(), runProps);
                    bufferText.Clear();
                }
                currentTarget = nextTarget;
            }
            if (currentTarget != null)
                bufferText.Append(ch);
        }

        if (currentTarget != null && bufferText.Length > 0)
            AppendBufferText(currentTarget, bufferText.ToString(), runProps);
    }

    private static void AppendBufferText(DxpFieldNodeBuffer buffer, string text, RunProperties? runProps)
    {
        if (string.IsNullOrEmpty(text))
            return;
        var run = new Run();
        if (runProps != null)
            run.RunProperties = (RunProperties)runProps.CloneNode(true);
        var t = new Text(text);
        if (NeedsPreserveSpace(text))
            t.Space = SpaceProcessingModeValues.Preserve;
        run.AppendChild(t);
        var child = buffer.BeginRun(run);
        child.AddText(text);
    }

    private sealed class DxpFieldNodeBufferRecorder : DxpVisitor
    {
        private readonly Stack<DxpFieldNodeBuffer> _stack = new();
        private readonly Stack<Run?> _runStack = new();

        public DxpFieldNodeBufferRecorder() : base(null)
        {
        }

        public void Reset(DxpFieldNodeBuffer root)
        {
            _stack.Clear();
            _runStack.Clear();
            _stack.Push(root);
            _runStack.Push(null);
        }

        private DxpFieldNodeBuffer Current => _stack.Peek();
        private Run? CurrentRun => _runStack.Peek();

        public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
        {
            var run = new Run();
            if (r.RunProperties != null)
                run.RunProperties = (RunProperties)r.RunProperties.CloneNode(true);
            var child = Current.BeginRun(run);
            _stack.Push(child);
            _runStack.Push(run);
            return DxpDisposable.Create(() => {
                _stack.Pop();
                _runStack.Pop();
            });
        }

        public override void VisitText(Text t, DxpIDocumentContext d)
        {
            Current.AddText(t.Text);
            var run = CurrentRun;
            if (run != null)
            {
                var text = new Text(t.Text);
                if (NeedsPreserveSpace(t.Text))
                    text.Space = SpaceProcessingModeValues.Preserve;
                run.AppendChild(text);
            }
        }
        public override void VisitDeletedText(DeletedText dt, DxpIDocumentContext d) => Current.AddDeletedText(dt.Text);
        public override void VisitBreak(Break b, DxpIDocumentContext d) => Current.AddBreak();
        public override void VisitTab(TabChar tab, DxpIDocumentContext d) => Current.AddTab();
        public override void VisitCarriageReturn(CarriageReturn cr, DxpIDocumentContext d) => Current.AddCarriageReturn();
        public override void VisitNoBreakHyphen(NoBreakHyphen nbh, DxpIDocumentContext d) => Current.AddNoBreakHyphen();

        public override IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target, DxpIDocumentContext d)
        {
            var clone = (Hyperlink)link.CloneNode(false);
            var child = Current.BeginHyperlink(clone, target);
            _stack.Push(child);
            return DxpDisposable.Create(() => _stack.Pop());
        }
    }

    private bool TryEvaluateIfAndEmit(FieldEvalFrameState frame, DxpIDocumentContext d)
    {
        if (frame.IfState == null || string.IsNullOrWhiteSpace(frame.InstructionText))
            return false;

        var ifResult = _eval.EvaluateIfConditionAsync(frame.InstructionText).GetAwaiter().GetResult();
        if (ifResult == null || !ifResult.Value.Success)
        {
            frame.Evaluated = true;
            frame.SuppressContent = true;
            EmitEvaluatedText(GetEvaluationErrorText(frame.InstructionText), d);
            return true;
        }

        var selected = ifResult.Value.Condition ? frame.IfState.TrueBuffer : frame.IfState.FalseBuffer;
        frame.Evaluated = true;
        frame.SuppressContent = true;
        if (selected.IsEmpty)
        {
            var evalResult = _eval.EvalAsync(new DxpFieldInstruction(frame.InstructionText)).GetAwaiter().GetResult();
            if (evalResult.Status == DxpFieldEvalStatus.Resolved && evalResult.Text != null)
                EmitEvaluatedText(evalResult.Text, d);
            else
                EmitEvaluatedText(GetEvaluationErrorText(frame.InstructionText), d);
            return true;
        }

        if (selected.TryGetRunSegments(out var segments))
        {
            foreach (var (text, props) in segments)
                EmitEvaluatedText(text, d, props);
            return true;
        }

        selected.Replay(_next, d);
        return true;
    }

    private FieldEvalFrameState? GetActiveIfFrame()
    {
        if (_fieldFrames.Count == 0)
            return null;

        foreach (var frame in _fieldFrames)
        {
            if (frame.IfState == null)
                continue;
            if (frame.InResult)
                continue;
            return frame;
        }

        return null;
    }

    private static string FormatIfLiteralToken(string value)
    {
        if (string.IsNullOrEmpty(value))
            return "\"\"";

        bool needsQuote = value.Any(char.IsWhiteSpace) || value.Contains('"');
        if (!needsQuote)
            return value;

        var escaped = value.Replace("\"", "\\\"");
        return $"\"{escaped}\"";
    }

    private bool TryInjectIfInstructionValue(string value)
    {
        var frame = GetActiveIfFrame();
        if (frame?.IfState == null)
            return false;
        if (frame.IfState.GetCurrentBuffer() != null)
            return false;

        var token = FormatIfLiteralToken(value);
        var needsSpace = !string.IsNullOrEmpty(frame.InstructionText) &&
            !char.IsWhiteSpace(frame.InstructionText![frame.InstructionText.Length - 1]);
        // Add a trailing space so the tokenizer commits the injected token.
        var segment = (needsSpace ? " " : string.Empty) + token + " ";
        frame.InstructionText = frame.InstructionText == null ? token : frame.InstructionText + segment;
        ProcessIfInstructionSegment(frame.IfState, segment, null);
        return true;
    }

    public override void VisitComplexFieldBegin(FieldChar begin, DxpIDocumentContext d)
    {
        var frame = new FieldEvalFrameState { IsComplex = true, InResult = false, SeenSeparate = false };
        _fieldFrames.Push(frame);
        if (_fieldFrames.Count == 1)
            _outerFrame = frame;
        d.CurrentFields.FieldStack.Push(new FieldFrame { SeenSeparate = false, ResultScope = null, InResult = false });
    }

    public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
    {
        if (_mode == DxpFieldEvalMode.Cache && _outerFrame?.SuppressContent == true)
            return;
        if (_mode == DxpFieldEvalMode.Cache && IsNestedField)
            return;

        var current = CurrentField;
        if (current != null && current.Evaluated)
            return;

        if (_mode == DxpFieldEvalMode.Evaluate)
        {
            if (current != null && current.Evaluated)
                return;
            if (current != null && ShouldDeferEvaluation(current))
                return;
            if (current != null && current.IfState != null)
                return;
            if ((!IsNestedField || IsCapturingForIf() || GetActiveIfFrame() != null) &&
                TryWriteEvaluatedCurrentField(current, d))
                return;
            // In eval mode, we do not forward cached results or instruction text.
            return;
        }

        if (current != null && !current.InstructionEmitted && !string.IsNullOrWhiteSpace(current.InstructionText))
        {
            current.InstructionEmitted = true;
            _next.VisitComplexFieldInstruction(new FieldCode(), current.InstructionText!, d);
        }

        if (!ShouldForwardContent(d))
            return;

        _next.VisitComplexFieldCachedResultText(text, d);
    }

    public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
    {
        if (!string.IsNullOrEmpty(text) && CurrentField?.InResult == true)
        {
            VisitComplexFieldCachedResultText(text, d);
            return;
        }

        if (!string.IsNullOrEmpty(text) && CurrentField != null)
        {
            var current = CurrentField;
            current.InstructionText = current.InstructionText == null
                ? text
                : current.InstructionText + text;
            if (_mode == DxpFieldEvalMode.Cache && IsSetInstruction(current.InstructionText))
                current.SuppressContent = true;
            if (current.IfState == null && StartsWithIf(current.InstructionText))
                current.IfState = new IfCaptureState();

            if (current.IfState != null)
            {
                var runProps = (instr.Parent as Run)?.RunProperties;
                ProcessIfInstructionSegment(current.IfState, text, runProps);
            }
            if (current.CodeRunProperties == null)
            {
                if (instr.Parent is Run instrRun && instrRun.RunProperties != null)
                    current.CodeRunProperties = (RunProperties)instrRun.RunProperties.CloneNode(true);
                else if (d.CurrentRun?.Properties != null)
                    current.CodeRunProperties = (RunProperties)d.CurrentRun.Properties.CloneNode(true);

                if (current.CodeRunProperties != null && _logger?.IsEnabled(LogLevel.Debug) == true)
                    _logger.LogDebug("Captured field code run properties.");
            }
        }
        if (!string.IsNullOrEmpty(text) && d.CurrentFields.Current != null)
        {
            var current = d.CurrentFields.Current;
            current.InstructionText = current.InstructionText == null
                ? text
                : current.InstructionText + text;
        }
    }

    public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext d)
    {
        if (CurrentField == null)
            return;

        if (!CurrentField.SeenSeparate)
        {
            CurrentField.SeenSeparate = true;
            CurrentField.InResult = true;
        }
        if (d.CurrentFields.Current != null && !d.CurrentFields.Current.SeenSeparate)
        {
            d.CurrentFields.Current.SeenSeparate = true;
            d.CurrentFields.Current.InResult = true;
        }

        if (_mode == DxpFieldEvalMode.Evaluate && !IsNestedField)
        {
            var frame = CurrentField;
            if (frame != null && frame.IfState != null && !frame.Evaluated)
            {
                if (TryEvaluateIfAndEmit(frame, d))
                    return;
            }
        }
    }

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
    {
        if (_mode == DxpFieldEvalMode.Evaluate && !IsNestedField)
        {
            var frame = CurrentField;
            if (frame != null && frame.IfState != null && !frame.Evaluated)
                TryEvaluateIfAndEmit(frame, d);
            if (frame != null && !frame.Evaluated)
                TryWriteEvaluatedCurrentField(frame, d);
            if (frame != null && !frame.Evaluated && ShouldDeferEvaluation(frame))
                TryWriteEvaluatedCurrentField(frame, d);
        }

        if (_fieldFrames.Count > 0)
        {
            if (_fieldFrames.Count == 1)
                _outerFrame = null;
            _fieldFrames.Pop();
        }
        if (d.CurrentFields.FieldStack.Count > 0)
            d.CurrentFields.FieldStack.Pop();
    }

    public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
    {
        var frame = new FieldEvalFrameState { IsComplex = false, InResult = true, SeenSeparate = true };
        var instruction = fld.Instruction?.Value;
        if (!string.IsNullOrEmpty(instruction))
            frame.InstructionText = instruction;
        if (frame.CodeRunProperties == null && d.CurrentRun?.Properties != null)
            frame.CodeRunProperties = (RunProperties)d.CurrentRun.Properties.CloneNode(true);
        _fieldFrames.Push(frame);
        if (_fieldFrames.Count == 1)
            _outerFrame = frame;
        var docFrame = new FieldFrame { SeenSeparate = true, InResult = true, InstructionText = instruction };
        d.CurrentFields.FieldStack.Push(docFrame);

        if (_mode == DxpFieldEvalMode.Evaluate && (!IsNestedField || IsCapturingForIf() || GetActiveIfFrame() != null))
        {
            if (frame.IfState == null && frame.InstructionText != null && StartsWithIf(frame.InstructionText))
                frame.IfState = new IfCaptureState();
            if (frame.IfState == null && !ShouldDeferEvaluation(frame) && TryWriteEvaluatedCurrentField(frame, d))
            {
                return DxpDisposable.Create(() => {
                    if (_fieldFrames.Count == 1)
                        _outerFrame = null;
                    if (_fieldFrames.Count > 0)
                        _fieldFrames.Pop();
                    if (d.CurrentFields.FieldStack.Count > 0)
                        d.CurrentFields.FieldStack.Pop();
                });
            }

            if (ShouldDeferEvaluation(frame) && !IsNestedField)
            {
                return DxpDisposable.Create(() => {
                    if (!frame.Evaluated)
                        TryWriteEvaluatedCurrentField(frame, d);
                    if (!frame.Evaluated && frame.IfState != null)
                        TryEvaluateIfAndEmit(frame, d);
                    if (_fieldFrames.Count == 1)
                        _outerFrame = null;
                    if (_fieldFrames.Count > 0)
                        _fieldFrames.Pop();
                    if (d.CurrentFields.FieldStack.Count > 0)
                        d.CurrentFields.FieldStack.Pop();
                });
            }
        }

        if (_mode == DxpFieldEvalMode.Cache &&
            frame != null &&
            !string.IsNullOrWhiteSpace(frame.InstructionText) &&
            frame.InstructionText.StartsWith("SET", StringComparison.OrdinalIgnoreCase))
        {
            return DxpDisposable.Create(() => {
                if (_fieldFrames.Count == 1)
                    _outerFrame = null;
                if (_fieldFrames.Count > 0)
                    _fieldFrames.Pop();
                if (d.CurrentFields.FieldStack.Count > 0)
                    d.CurrentFields.FieldStack.Pop();
            });
        }

        var inner = DxpDisposable.Empty;
        return new DxpCompositeScope(inner, () => {
            if (_fieldFrames.Count == 1)
                _outerFrame = null;
            if (_fieldFrames.Count > 0)
                _fieldFrames.Pop();
            if (d.CurrentFields.FieldStack.Count > 0)
                d.CurrentFields.FieldStack.Pop();
        });
    }

    public override void VisitText(Text t, DxpIDocumentContext d)
    {
        if (_mode == DxpFieldEvalMode.Cache)
        {
            if (_outerFrame?.InResult == true)
            {
                VisitComplexFieldCachedResultText(t.Text, d);
                return;
            }
        }
        else if (IsInFieldResult)
        {
            VisitComplexFieldCachedResultText(t.Text, d);
            return;
        }

        if (!ShouldForwardContent(d))
            return;

        _next.VisitText(t, d);
    }

    private bool TryWriteEvaluatedCurrentField(FieldEvalFrameState? frame, DxpIDocumentContext d)
    {
        var instruction = frame?.InstructionText;
        if (string.IsNullOrWhiteSpace(instruction))
            return false;

        var fallbackRunProps = frame?.CodeRunProperties;
        var parse = _parser.Parse(instruction);
        var fieldType = parse.Ast.FieldType ?? string.Empty;
        var argsText = parse.Ast.ArgumentsText;

        if ((fieldType.Equals("REF", StringComparison.OrdinalIgnoreCase) ||
            fieldType.Equals("DOCVARIABLE", StringComparison.OrdinalIgnoreCase)) &&
            TryGetCharOrMergeFormat(parse.Ast.FormatSpecs, out var hasCharFormat, out var hasMergeFormat) &&
            (hasCharFormat || hasMergeFormat) &&
            TryGetFirstToken(argsText, out var formatName))
        {
            RunProperties? runProps = null;
            IReadOnlyList<RunProperties?>? mergeRunProps = null;
            if (hasMergeFormat && frame?.CachedResultRunProperties != null && frame.CachedResultRunProperties.Count > 0)
                mergeRunProps = frame.CachedResultRunProperties;
            else if (hasCharFormat && frame?.CodeRunProperties != null)
                runProps = frame.CodeRunProperties;
            if (hasCharFormat && runProps == null && mergeRunProps == null && _logger?.IsEnabled(LogLevel.Debug) == true)
                _logger.LogDebug("CHARFORMAT requested but no field code run properties captured.");

            if (runProps == null && mergeRunProps == null)
            {
                if (fieldType.Equals("REF", StringComparison.OrdinalIgnoreCase))
                {
                    if (_context.TryGetBookmarkNodes(formatName, out var refNodes))
                        refNodes.TryGetFirstRunProperties(out runProps);
                }
                else if (fieldType.Equals("DOCVARIABLE", StringComparison.OrdinalIgnoreCase))
                {
                    if (_context.TryGetDocVariableNodes(formatName, out var docVarNodes))
                        docVarNodes.TryGetFirstRunProperties(out runProps);
                }
            }

            var formatResult = _eval.EvalAsync(new DxpFieldInstruction(instruction!)).GetAwaiter().GetResult();
            if (fieldType.Equals("DOCVARIABLE", StringComparison.OrdinalIgnoreCase) && runProps == null &&
                _context.TryGetDocVariableNodes(formatName, out var postEvalNodes))
            {
                postEvalNodes.TryGetFirstRunProperties(out runProps);
            }

            if (frame != null)
            {
                frame.Evaluated = true;
                frame.SuppressContent = true;
            }

            var formatText = formatResult.Status == DxpFieldEvalStatus.Resolved && formatResult.Text != null
                ? formatResult.Text
                : GetEvaluationErrorText(instruction);
            if (TryInjectIfInstructionValue(formatText))
                return true;
            if (mergeRunProps != null)
                EmitEvaluatedText(formatText, d, mergeRunProps);
            else
                EmitEvaluatedText(formatText, d, runProps ?? fallbackRunProps);
            return true;
        }

        if (fieldType.Equals("REF", StringComparison.OrdinalIgnoreCase) &&
            !HasFieldSwitches(instruction) &&
            TryGetFirstToken(argsText, out var refName) &&
            _context.TryGetBookmarkNodes(refName, out var storedNodes))
        {
            if (frame != null)
            {
                frame.Evaluated = true;
                frame.SuppressContent = true;
            }

            if (TryInjectIfInstructionValue(storedNodes.ToPlainText()))
                return true;
            var replayVisitor = TryGetCaptureVisitor(out var captureVisitor, out _) ? captureVisitor! : _next;
            storedNodes.Replay(replayVisitor, d);
            return true;
        }

        if (fieldType.Equals("DOCVARIABLE", StringComparison.OrdinalIgnoreCase) &&
            !HasFieldSwitches(instruction) &&
            TryGetFirstToken(argsText, out var docVarName))
        {
            if (!_context.TryGetDocVariableNodes(docVarName, out var docVarNodes))
            {
                _eval.EvalAsync(new DxpFieldInstruction(instruction!)).GetAwaiter().GetResult();
                _context.TryGetDocVariableNodes(docVarName, out docVarNodes);
            }

            if (docVarNodes != null)
            {
                if (frame != null)
                {
                    frame.Evaluated = true;
                    frame.SuppressContent = true;
                }

                if (TryInjectIfInstructionValue(docVarNodes.ToPlainText()))
                    return true;
                if (docVarNodes.TryGetFirstRunProperties(out var docVarProps) && docVarProps != null)
                {
                    var replayVisitor = TryGetCaptureVisitor(out var captureVisitor, out _) ? captureVisitor! : _next;
                    docVarNodes.Replay(replayVisitor, d);
                }
                else if (fallbackRunProps != null)
                {
                    EmitEvaluatedText(docVarNodes.ToPlainText(), d, fallbackRunProps);
                }
                else
                {
                    var replayVisitor = TryGetCaptureVisitor(out var captureVisitor, out _) ? captureVisitor! : _next;
                    docVarNodes.Replay(replayVisitor, d);
                }
                return true;
            }
        }

        if (fieldType.Equals("SET", StringComparison.OrdinalIgnoreCase))
        {
            if (frame != null)
            {
                frame.Evaluated = true;
                frame.SuppressContent = true;
            }

            var setResult = _eval.EvalAsync(new DxpFieldInstruction(instruction!)).GetAwaiter().GetResult();
            if (TryGetFirstToken(argsText, out var setName))
            {
                var text = setResult.Text ?? string.Empty;
                _context.SetBookmarkNodes(setName, DxpFieldNodeBuffer.FromText(text));
            }
            return true;
        }
        else
        {

            var result = _eval.EvalAsync(new DxpFieldInstruction(instruction!)).GetAwaiter().GetResult();
            if (frame != null)
            {
                frame.Evaluated = true;
                frame.SuppressContent = true;
            }

            if (result.Status != DxpFieldEvalStatus.Resolved || result.Text == null)
            {
                var errorText = GetEvaluationErrorText(instruction);
                if (TryInjectIfInstructionValue(errorText))
                    return true;
                EmitEvaluatedText(errorText, d, fallbackRunProps);
                return true;
            }

            if (TryInjectIfInstructionValue(result.Text))
                return true;
            EmitEvaluatedText(result.Text, d, fallbackRunProps);
            return true;
        }
    }

    private string GetEvaluationErrorText(string instruction)
    {
        var parse = _parser.Parse(instruction);
        var fieldType = parse.Ast.FieldType;
        if (string.IsNullOrWhiteSpace(fieldType))
            return "Error! Invalid field code.";

        switch (fieldType.Trim().ToUpperInvariant())
        {
            case "REF":
                return "Error! Reference source not found.";
            case "DOCVARIABLE":
                return "Error! No document variable supplied.";
            case "DOCPROPERTY":
                return "Error! Unknown document property name.";
            case "IF":
                return "Error! Invalid field code.";
            case "=":
                return "Error! Invalid formula.";
            default:
                return "Error! Invalid field code.";
        }
    }

    private static bool HasFieldSwitches(string instruction)
    {
        bool inQuote = false;
        int braceDepth = 0;
        for (int i = 0; i < instruction.Length; i++)
        {
            char ch = instruction[i];
            if (inQuote && ch == '\\' && i + 1 < instruction.Length && instruction[i + 1] == '"')
            {
                i++;
                continue;
            }
            if (ch == '"')
            {
                inQuote = !inQuote;
                continue;
            }

            if (!inQuote)
            {
                if (ch == '{')
                {
                    braceDepth++;
                    continue;
                }
                if (ch == '}' && braceDepth > 0)
                {
                    braceDepth--;
                    continue;
                }
            }

            if (!inQuote && braceDepth == 0 && ch == '\\')
                return true;
        }

        return false;
    }

    private static bool TryGetCharOrMergeFormat(
        IReadOnlyList<IDxpFieldFormatSpec> specs,
        out bool hasCharFormat,
        out bool hasMergeFormat)
    {
        hasCharFormat = false;
        hasMergeFormat = false;
        foreach (var spec in specs)
        {
            if (spec is not DxpTextTransformFormatSpec transform)
                continue;
            if (transform.Kind == DxpTextTransformKind.Charformat)
                hasCharFormat = true;
            else if (transform.Kind == DxpTextTransformKind.MergeFormat)
                hasMergeFormat = true;
        }
        return hasCharFormat || hasMergeFormat;
    }

    private static bool TryGetFirstToken(string? argsText, out string token)
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

    private static List<string> TokenizeArgs(string text)
    {
        var tokens = new List<string>();
        bool inQuote = false;
        bool justClosedQuote = false;
        var current = new StringBuilder();
        for (int i = 0; i < text.Length; i++)
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
                    justClosedQuote = true;
                continue;
            }

            if (!inQuote && ch == '{')
            {
                int start = i;
                int depth = 0;
                for (; i < text.Length; i++)
                {
                    if (text[i] == '{')
                        depth++;
                    else if (text[i] == '}')
                    {
                        depth--;
                        if (depth == 0)
                        {
                            i++;
                            break;
                        }
                    }
                }
                string token = text.Substring(start, i - start).Trim();
                if (current.Length > 0)
                {
                    tokens.Add(current.ToString());
                    current.Clear();
                }
                if (token.Length > 0)
                    tokens.Add(token);
                justClosedQuote = false;
                i--;
                continue;
            }

            if (!inQuote && char.IsWhiteSpace(ch))
            {
                if (current.Length > 0)
                {
                    tokens.Add(current.ToString());
                    current.Clear();
                    justClosedQuote = false;
                }
                else if (justClosedQuote)
                {
                    tokens.Add(string.Empty);
                    justClosedQuote = false;
                }
                continue;
            }

            current.Append(ch);
            justClosedQuote = false;
        }

        if (current.Length > 0)
            tokens.Add(current.ToString());
        else if (justClosedQuote)
            tokens.Add(string.Empty);

        return tokens;
    }

    private void EmitEvaluatedText(string text, DxpIDocumentContext d, RunProperties? runProperties = null)
    {
        if (string.IsNullOrEmpty(text))
            return;

        if (TryGetCaptureVisitor(out var captureVisitor, out var captureBuffer))
        {
            var captureRun = new Run();
            if (runProperties != null)
                captureRun.RunProperties = (RunProperties)runProperties.CloneNode(true);
            var captureText = new Text(text);
            if (NeedsPreserveSpace(text))
                captureText.Space = SpaceProcessingModeValues.Preserve;
            captureRun.AppendChild(captureText);
            using (captureVisitor!.VisitRunBegin(captureRun, d))
                captureVisitor.VisitText(captureText, d);
            return;
        }

        var sink = _next;

        var run = new Run();
        if (runProperties != null)
            run.RunProperties = (RunProperties)runProperties.CloneNode(true);
        // Attach a Text child so downstream style tracking sees renderable content.
        var t = new Text(text);
        if (NeedsPreserveSpace(text))
            t.Space = SpaceProcessingModeValues.Preserve;
        run.AppendChild(t);
        // Attach to a temporary paragraph so style resolution can see a paragraph ancestor.
        var tempParagraph = new Paragraph();
        tempParagraph.Append(run);
        using (sink.VisitRunBegin(run, d))
        {
            sink.VisitText(t, d);
        }
    }

    private void EmitEvaluatedText(string text, DxpIDocumentContext d, IReadOnlyList<RunProperties?> runProperties)
    {
        if (string.IsNullOrEmpty(text))
            return;
        if (runProperties == null || runProperties.Count == 0)
        {
            EmitEvaluatedText(text, d, (RunProperties?)null);
            return;
        }

        var segments = SplitTextByRuns(text, runProperties.Count);
        for (int i = 0; i < segments.Count; i++)
            EmitEvaluatedText(segments[i], d, runProperties[i]);
    }

    private static bool NeedsPreserveSpace(string text)
    {
        if (text.Length == 0)
            return false;
        if (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1]))
            return true;
        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            if (ch == '\t' || ch == '\r' || ch == '\n')
                return true;
            if (ch == ' ' && i + 1 < text.Length && text[i + 1] == ' ')
                return true;
        }
        return false;
    }

    private static DxpStyleEffectiveRunStyle BuildEffectiveRunStyle(DxpIDocumentContext d, RunProperties runProperties)
    {
        var defaults = d.DefaultRunStyle;
        var acc = new DxpEffectiveRunStyleBuilder {
            Bold = defaults.Bold,
            Italic = defaults.Italic,
            Underline = defaults.Underline,
            Strike = defaults.Strike,
            DoubleStrike = defaults.DoubleStrike,
            Superscript = defaults.Superscript,
            Subscript = defaults.Subscript,
            AllCaps = defaults.AllCaps,
            SmallCaps = defaults.SmallCaps,
            FontName = defaults.FontName,
            FontSizeHalfPoints = defaults.FontSizeHalfPoints
        };

        DxpEffectiveRunStyleBuilder.ApplyRunProperties(runProperties, null, ref acc);
        return acc.ToImmutable();
    }

    private DxpIVisitor GetSyntheticSink()
    {
        return _next is DxpMiddleware middleware ? middleware.Next : _next;
    }

    private bool TryGetCaptureVisitor(out DxpIVisitor? visitor, out DxpFieldNodeBuffer? buffer)
    {
        visitor = null;
        buffer = null;
        var state = GetActiveIfState();
        if (state == null)
            return false;
        buffer = state.GetCurrentBuffer();
        if (buffer == null)
            return false;
        state.Recorder.Reset(buffer);
        visitor = state.Recorder;
        return true;
    }

    private IfCaptureState? GetActiveIfState()
    {
        if (_fieldFrames.Count == 0)
            return null;

        foreach (var frame in _fieldFrames)
        {
            if (frame.IfState == null)
                continue;
            if (frame.InResult)
                continue;
            return frame.IfState;
        }

        return null;
    }

    private bool IsCapturingForIf()
    {
        var state = GetActiveIfState();
        return state?.GetCurrentBuffer() != null;
    }

    private static bool HasRenderableContent(Run r)
    {
        return r.ChildElements.Any(child =>
            child is Text or DeletedText or NoBreakHyphen or TabChar or Break or CarriageReturn or Drawing);
    }

    private static List<string> SplitTextByRuns(string text, int count)
    {
        var segments = new List<string>(count);
        if (count <= 1)
        {
            segments.Add(text);
            return segments;
        }

        int length = text.Length;
        int baseSize = length / count;
        int remainder = length % count;
        int offset = 0;

        for (int i = 0; i < count; i++)
        {
            int size = baseSize + (i < remainder ? 1 : 0);
            if (offset >= length)
            {
                segments.Add(string.Empty);
                continue;
            }
            if (offset + size > length)
                size = length - offset;
            segments.Add(text.Substring(offset, size));
            offset += size;
        }

        return segments;
    }

    private static bool IsSetInstruction(string? instruction)
    {
        if (string.IsNullOrWhiteSpace(instruction))
            return false;
        var trimmed = instruction.TrimStart();
        if (!trimmed.StartsWith("SET", StringComparison.OrdinalIgnoreCase))
            return false;
        return trimmed.Length == 3 || char.IsWhiteSpace(trimmed[3]);
    }


    public override IDisposable VisitParagraphBegin(Paragraph p, DxpIDocumentContext d, DxpIParagraphContext paragraph)
    {
        var previous = _context.Culture;
        var previousOutlineProvider = _context.CurrentOutlineLevelProvider;
        var previousOrder = _context.CurrentDocumentOrder;
        if (TryResolveParagraphCulture(p, d, _logger, out var culture))
            _context.Culture = culture;
        _context.CurrentOutlineLevelProvider = CreateOutlineLevelProvider(p, d);
        _context.CurrentDocumentOrder = ++_paragraphOrder;

        var inner = _next.VisitParagraphBegin(p, d, paragraph);
        return new DxpCompositeScope(inner, () => {
            _context.Culture = previous;
            _context.CurrentOutlineLevelProvider = previousOutlineProvider;
            _context.CurrentDocumentOrder = previousOrder;
        });
    }

    public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
    {
        var previous = _context.Culture;
        if (TryResolveRunCulture(r, d, _logger, out var culture))
            _context.Culture = culture;

        if (CurrentField != null && !CurrentField.InResult && CurrentField.CodeRunProperties == null && r.RunProperties != null)
        {
            CurrentField.CodeRunProperties = (RunProperties)r.RunProperties.CloneNode(true);
            if (_logger?.IsEnabled(LogLevel.Debug) == true)
                _logger.LogDebug("Captured field code run properties from run.");
        }

        if (_mode == DxpFieldEvalMode.Evaluate && IsInFieldResult && CurrentField != null)
        {
            if (HasRenderableContent(r))
            {
                CurrentField.CachedResultRunProperties ??= new List<RunProperties?>();
                RunProperties? props = r.RunProperties != null ? (RunProperties)r.RunProperties.CloneNode(true) : null;
                CurrentField.CachedResultRunProperties.Add(props);
            }
        }

        if (_mode == DxpFieldEvalMode.Evaluate && IsInFieldResult)
            return new DxpCompositeScope(DxpDisposable.Empty, () => _context.Culture = previous);

        var inner = _next.VisitRunBegin(r, d);
        return new DxpCompositeScope(inner, () => _context.Culture = previous);
    }

    private static bool TryResolveParagraphCulture(Paragraph p, DxpIDocumentContext d, ILogger? logger, out CultureInfo culture)
    {
        culture = CultureInfo.CurrentCulture;
        string? lang = null;

        if (d.Styles is DxpStyleResolver resolver)
            lang = resolver.ResolveParagraphLanguage(p) ?? resolver.GetDefaultLanguage();
        else
            lang = p.ParagraphProperties?.GetFirstChild<ParagraphMarkRunProperties>()
                ?.GetFirstChild<Languages>()?.Val?.Value;

        return TryCreateCulture(lang, logger, out culture);
    }

    private static bool TryResolveRunCulture(Run r, DxpIDocumentContext d, ILogger? logger, out CultureInfo culture)
    {
        culture = CultureInfo.CurrentCulture;
        string? lang = null;

        if (d.Styles is DxpStyleResolver resolver)
        {
            var paragraph = r.Ancestors<Paragraph>().FirstOrDefault();
            if (paragraph != null)
                lang = resolver.ResolveRunLanguage(paragraph, r);
        }

        lang ??= d.CurrentRun?.Language ?? r.RunProperties?.Languages?.Val?.Value;
        return TryCreateCulture(lang, logger, out culture);
    }

    private static bool TryCreateCulture(string? lang, ILogger? logger, out CultureInfo culture)
    {
        culture = CultureInfo.CurrentCulture;
        if (string.IsNullOrWhiteSpace(lang))
            return false;

        try
        {
            culture = new CultureInfo(lang);
            return true;
        }
        catch (CultureNotFoundException)
        {
            logger?.LogWarning("Invalid language tag '{Lang}' in document; using current culture.", lang);
            return false;
        }
    }

    private static Func<int> CreateOutlineLevelProvider(Paragraph p, DxpIDocumentContext d)
    {
        int? level = null;
        if (d.Styles is DxpStyleResolver resolver)
            level = resolver.GetOutlineLevel(p);
        else
            level = p.ParagraphProperties?.OutlineLevel?.Val?.Value;

        // Word stores outline levels as 0-based; SEQ \s expects 1-based.
        int resolved = level.HasValue ? level.Value + 1 : 0;
        return () => resolved;
    }

    private sealed class DxpCompositeScope : IDisposable
    {
        private readonly IDisposable _inner;
        private readonly Action _onDispose;
        private bool _disposed;

        public DxpCompositeScope(IDisposable inner, Action onDispose)
        {
            _inner = inner;
            _onDispose = onDispose;
        }

        public void Dispose()
        {
            if (_disposed)
                return;
            _disposed = true;
            _onDispose();
            _inner.Dispose();
        }
    }

    private sealed class DxpWalkerTableResolver : IDxpTableValueResolver
    {
        private readonly DxpIDocumentContext _document;

        public DxpWalkerTableResolver(DxpIDocumentContext document)
        {
            _document = document;
        }

        public Task<IReadOnlyList<double>> ResolveRangeAsync(string range, DxpFieldEvalContext context)
        {
            var model = _document.CurrentTableModel;
            if (model == null)
                return Task.FromResult<IReadOnlyList<double>>([]);

            if (!TryParseRange(range, out var startRow, out var startCol, out var endRow, out var endCol))
                return Task.FromResult<IReadOnlyList<double>>([]);

            var values = CollectRangeValues(model, startRow, startCol, endRow, endCol, context);
            return Task.FromResult<IReadOnlyList<double>>(values);
        }

        public Task<IReadOnlyList<double>> ResolveDirectionalRangeAsync(DxpTableRangeDirection direction, DxpFieldEvalContext context)
        {
            var model = _document.CurrentTableModel;
            var cell = _document.CurrentTableCell;
            if (model == null || cell == null)
                return Task.FromResult<IReadOnlyList<double>>([]);

            int row = cell.RowIndex;
            int col = cell.ColumnIndex;
            int startRow, endRow, startCol, endCol;

            switch (direction)
            {
                case DxpTableRangeDirection.Above:
                    startRow = 0;
                    endRow = row - 1;
                    startCol = col;
                    endCol = col;
                    break;
                case DxpTableRangeDirection.Below:
                    startRow = row + 1;
                    endRow = model.RowCount - 1;
                    startCol = col;
                    endCol = col;
                    break;
                case DxpTableRangeDirection.Left:
                    startRow = row;
                    endRow = row;
                    startCol = 0;
                    endCol = col - 1;
                    break;
                case DxpTableRangeDirection.Right:
                    startRow = row;
                    endRow = row;
                    startCol = col + 1;
                    endCol = model.ColumnCount - 1;
                    break;
                default:
                    return Task.FromResult<IReadOnlyList<double>>([]);
            }

            var values = CollectRangeValues(model, startRow, startCol, endRow, endCol, context);
            return Task.FromResult<IReadOnlyList<double>>(values);
        }

        private static List<double> CollectRangeValues(DxpTableModel model, int startRow, int startCol, int endRow, int endCol, DxpFieldEvalContext context)
        {
            var values = new List<double>();
            if (startRow > endRow || startCol > endCol)
                return values;

            startRow = Math.Max(0, startRow);
            startCol = Math.Max(0, startCol);
            endRow = Math.Min(model.RowCount - 1, endRow);
            endCol = Math.Min(model.ColumnCount - 1, endCol);

            for (int r = startRow; r <= endRow; r++)
            {
                for (int c = startCol; c <= endCol; c++)
                {
                    var cell = model.Cells[r, c];
                    if (cell == null || cell.IsCovered)
                        continue;

                    string text = ExtractCellText(cell.Cell);
                    if (TryParseNumber(text, context, out var number))
                        values.Add(number);
                }
            }

            return values;
        }

        private static bool TryParseRange(string range, out int startRow, out int startCol, out int endRow, out int endCol)
        {
            startRow = startCol = endRow = endCol = 0;
            if (string.IsNullOrWhiteSpace(range))
                return false;

            var parts = range.Split(':');
            if (parts.Length == 0 || parts.Length > 2)
                return false;

            if (!TryParseCell(parts[0], out startRow, out startCol))
                return false;

            if (parts.Length == 2)
            {
                if (!TryParseCell(parts[1], out endRow, out endCol))
                    return false;
            }
            else
            {
                endRow = startRow;
                endCol = startCol;
            }

            if (startRow > endRow)
                (startRow, endRow) = (endRow, startRow);
            if (startCol > endCol)
                (startCol, endCol) = (endCol, startCol);

            return true;
        }

        private static bool TryParseCell(string text, out int rowIndex, out int colIndex)
        {
            rowIndex = colIndex = 0;
            if (string.IsNullOrWhiteSpace(text))
                return false;

            int i = 0;
            while (i < text.Length && char.IsLetter(text[i]))
                i++;
            if (i == 0 || i == text.Length)
                return false;

            var colPart = text.Substring(0, i).ToUpperInvariant();
            var rowPart = text.Substring(i);
            if (!int.TryParse(rowPart, NumberStyles.Integer, CultureInfo.InvariantCulture, out var row))
                return false;

            colIndex = ColumnLettersToIndex(colPart) - 1;
            rowIndex = row - 1;
            return rowIndex >= 0 && colIndex >= 0;
        }

        private static int ColumnLettersToIndex(string letters)
        {
            int value = 0;
            foreach (char ch in letters)
            {
                if (ch < 'A' || ch > 'Z')
                    return 0;
                value = value * 26 + (ch - 'A' + 1);
            }
            return value;
        }

        private static string ExtractCellText(TableCell cell)
        {
            var sb = new StringBuilder();
            foreach (var text in cell.Descendants<Text>())
                sb.Append(text.Text);
            return sb.ToString().Trim();
        }

        private static bool TryParseNumber(string text, DxpFieldEvalContext context, out double number)
        {
            var culture = context.Culture ?? CultureInfo.CurrentCulture;
            if (double.TryParse(text, NumberStyles.Any, culture, out number))
                return true;
            if (context.AllowInvariantNumericFallback &&
                double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out number))
                return true;
            number = 0;
            return false;
        }
    }
}
