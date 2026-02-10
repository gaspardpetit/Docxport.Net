using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Fields.Eval;
using DocxportNet.Fields.Formatting;
using DocxportNet.Fields.Resolution;
using DocxportNet.Middleware;
using DocxportNet.Walker;
using DocxportNet.Walker.Context;
using Microsoft.Extensions.Logging;
using System.Text;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpRefFieldEvalFrame : DxpMiddleware, DxpIFieldEvalFrame
{
    public bool SuppressContent { get; set; }
    public bool Evaluated { get; set; }
    public bool SeenSeparate { get; set; }
    public bool InResult { get; set; }
    public string? InstructionText { get; set; }
    public bool InstructionEmitted { get; set; }
    public RunProperties? CodeRunProperties { get; set; }
    public Run? CodeRun { get; set; }
    public List<Run?>? CachedResultRuns { get; set; }
    public List<RunProperties?>? CachedResultRunProperties { get; set; }
    public DxpIFCaptureState? IfState { get; set; }

    public override DxpIVisitor Next { get; }
    public DxpFieldEvalContext EvalContext { get; }

    private readonly DxpFieldEval _eval;
    private readonly DxpFieldParser _parser = new();
    private readonly ILogger? _logger;

    public DxpRefFieldEvalFrame(DxpIVisitor next, DxpFieldEval eval, DxpFieldEvalContext evalContext, ILogger? logger)
        : base()
    {
		Next = next ?? throw new ArgumentNullException(nameof(next));
		_eval = eval ?? throw new ArgumentNullException(nameof(eval));
        EvalContext = evalContext ?? throw new ArgumentNullException(nameof(evalContext));
        _logger = logger;
    }

    public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
    {
        if (string.IsNullOrEmpty(text) || InResult)
            return;
        if (CodeRun == null)
        {
            CodeRun = instr.Parent as Run;
            LogRunInfo("Ref.CodeRunCaptured", CodeRun);
        }
    }

    public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext d)
    {
        if (!SeenSeparate)
        {
            SeenSeparate = true;
            InResult = true;
            if (ReferenceEquals(EvalContext.ActiveIfFrame, this))
                EvalContext.ActiveIfFrame = null;
        }

        if (CanEvaluateInCurrentScope())
            TryEvaluateAndEmit(d);
    }

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
    {
        if (CanEvaluateInCurrentScope())
        {
            TryEvaluateAndEmit(d);
            if (!Evaluated)
                EvaluateRef(d);
            if (!Evaluated && ShouldDeferEvaluation())
                EvaluateRef(d);
        }
    }

    public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
    {
        if (Evaluated)
            return;
        if (_logger?.IsEnabled(LogLevel.Debug) == true)
            _logger.LogDebug("CachedText: text='{Text}' frame={Frame} suppress={Suppress} outerSuppress={OuterSuppress} depth={Depth}",
                text,
                GetType().Name,
                SuppressContent,
                EvalContext.OuterFrame?.SuppressContent == true,
                EvalContext.FieldDepth);

        if (ShouldDeferEvaluation())
            return;
        if (IfState != null)
            return;
        if ((CanEvaluateInCurrentScope()) && EvaluateRef(d))
            return;
    }

    public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
    {
        if (CanEvaluateInCurrentScope())
        {
            if (IfState == null && !ShouldDeferEvaluation() && EvaluateRef(d))
                return DxpDisposable.Empty;

            if (ShouldDeferEvaluation() && CanEvaluateInCurrentScope())
            {
                return DxpDisposable.Create(() => {
                    if (!Evaluated)
                        EvaluateRef(d);
                });
            }
        }

        return DxpDisposable.Empty;
    }

    public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
    {
        if (InResult && DxpFieldEvalRules.HasRenderableContent(r))
        {
            CachedResultRunProperties ??= new List<RunProperties?>();
            RunProperties? props = r.RunProperties != null ? (RunProperties)r.RunProperties.CloneNode(true) : null;
            CachedResultRunProperties.Add(props);
            CachedResultRuns ??= new List<Run?>();
            CachedResultRuns.Add(r);
        }

        if (InResult)
            return DxpDisposable.Empty;

        return base.VisitRunBegin(r, d);
    }

    protected override bool ShouldForwardContent(DxpIDocumentContext d)
    {
        if (InResult != true)
            return true;
        if (EvalContext.FieldDepth > 1 && EvalContext.ActiveIfFrame == null)
            return false;
        if (SuppressContent)
            return false;
        return true;
    }

    private bool ShouldDeferEvaluation()
    {
        if (InstructionText == null)
            return false;
        return HasMergeFormat(InstructionText);
    }

    private bool TryEvaluateAndEmit(DxpIDocumentContext d)
    {
        if (IfState == null || string.IsNullOrWhiteSpace(InstructionText))
            return false;

        return DxpFieldEvalIfRunner.TryEvaluateAndEmit(this, _eval, d, Next, GetEvaluationErrorText, EmitEvaluatedText);
    }

    private bool EvaluateRef(DxpIDocumentContext d)
    {
        if (Evaluated)
            return false;
        if (string.IsNullOrWhiteSpace(InstructionText))
            return false;

        var instruction = InstructionText!;
        var parse = _parser.Parse(instruction);
        var argsText = parse.Ast.ArgumentsText;

        if (TryHandleRefWithFormat(d, instruction, argsText, parse.Ast.FormatSpecs))
            return true;

        var refResult = _eval.EvalAsync(new DxpFieldInstruction(instruction), d).GetAwaiter().GetResult();
        Evaluated = true;
        SuppressContent = true;

        if (refResult.Status != DxpFieldEvalStatus.Resolved || refResult.Text == null)
        {
            var errorText = GetEvaluationErrorText(instruction);
            if (TryInjectIfInstructionValue(errorText))
                return true;
            EmitEvaluatedText(errorText, d, CodeRunProperties, CodeRun);
            return true;
        }

        if (TryInjectIfInstructionValue(refResult.Text))
            return true;
        EmitEvaluatedText(refResult.Text, d, CodeRunProperties, CodeRun);
        return true;
    }

    private bool TryHandleRefWithFormat(
        DxpIDocumentContext d,
        string instruction,
        string? argsText,
        IReadOnlyList<IDxpFieldFormatSpec> formatSpecs)
    {
        if (!DxpFieldEvalFrameFactory.IsRefInstruction(instruction))
            return false;

        string formatName;
        if (!DxpFieldEvalRules.TryGetCharOrMergeFormat(formatSpecs, out var hasCharFormat, out var hasMergeFormat) ||
            (!hasCharFormat && !hasMergeFormat) ||
            !TryGetFirstToken(argsText, out formatName))
            return false;

        RunProperties? runProps = null;
        IReadOnlyList<RunProperties?>? mergeRunProps = null;
        IReadOnlyList<Run?>? mergeRuns = null;
        if (hasMergeFormat && CachedResultRunProperties != null && CachedResultRunProperties.Count > 0)
            mergeRunProps = CachedResultRunProperties;
        if (hasMergeFormat && CachedResultRuns != null && CachedResultRuns.Count > 0)
            mergeRuns = CachedResultRuns;
        else if (hasCharFormat && CodeRunProperties != null)
            runProps = CodeRunProperties;
        if (hasCharFormat && runProps == null && mergeRunProps == null && _logger?.IsEnabled(LogLevel.Debug) == true)
            _logger.LogDebug("CHARFORMAT requested but no field code run properties captured.");

        var formatResult = _eval.EvalAsync(new DxpFieldInstruction(instruction), d).GetAwaiter().GetResult();
        Evaluated = true;
        SuppressContent = true;

        var formatText = formatResult.Status == DxpFieldEvalStatus.Resolved && formatResult.Text != null
            ? formatResult.Text
            : GetEvaluationErrorText(instruction);
        if (TryInjectIfInstructionValue(formatText))
            return true;
        if (mergeRunProps != null || mergeRuns != null)
            EmitEvaluatedText(formatText, d, mergeRuns, mergeRunProps);
        else
            EmitEvaluatedText(formatText, d, runProps ?? CodeRunProperties, CodeRun);
        return true;
    }

    private DxpRefRecord? TryResolveRefRecord(string bookmark, DxpIDocumentContext d)
    {
        if (EvalContext.RefResolver == null)
            return null;
        var request = DxpRefRequests.Simple(bookmark);
        return EvalContext.RefResolver.ResolveAsync(request, EvalContext, d).GetAwaiter().GetResult();
    }

    private bool TryInjectIfInstructionValue(string value)
    {
        var frame = EvalContext.ActiveIfFrame;
        var ifState = frame?.IfState;
        if (ifState == null)
            return false;
        if (ifState.GetCurrentBuffer() != null)
            return false;
        if (frame == null)
            return false;

        var token = FormatIfLiteralToken(value);
        var instructionText = frame.InstructionText;
        var needsSpace = !string.IsNullOrEmpty(instructionText) &&
            !char.IsWhiteSpace(instructionText![instructionText.Length - 1]);
        var segment = (needsSpace ? " " : string.Empty) + token + " ";
        frame.InstructionText = instructionText == null ? token : instructionText + segment;
        DxpFieldEvalIfRunner.ProcessInstructionSegment(frame, segment, null, null);
        return true;
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

    private bool TryGetCaptureVisitor(out DxpIVisitor? visitor, out DxpFieldNodeBuffer? buffer)
    {
        visitor = null;
        buffer = null;
        var state = EvalContext.ActiveIfFrame?.IfState;
        if (state == null)
            return false;
        buffer = state.GetCurrentBuffer();
        if (buffer == null)
            return false;
        state.Recorder.Reset(buffer);
        visitor = state.Recorder;
        return true;
    }

    private void EmitEvaluatedText(string text, DxpIDocumentContext d, RunProperties? runProperties)
        => EmitEvaluatedText(text, d, runProperties, CodeRun);

    private void EmitEvaluatedText(string text, DxpIDocumentContext d, RunProperties? runProperties = null, Run? sourceRun = null)
    {
        if (string.IsNullOrEmpty(text))
            return;
        var effectiveRunProps = runProperties;
        if (_logger?.IsEnabled(LogLevel.Debug) == true)
            _logger.LogDebug("EmitEvaluatedText: text='{Text}' runProps={RunProps}",
                text,
                effectiveRunProps != null);
        LogRunInfo("Ref.Emit.SourceRun", sourceRun);
        LogRunInfo("Ref.Emit.CodeRun", CodeRun);

        if (TryGetCaptureVisitor(out var captureVisitor, out _))
        {
            var captureRun = sourceRun != null
                ? DxpRunCloner.CloneRunWithParagraphAncestor(sourceRun)
                : new Run();
            if (captureRun.RunProperties == null && effectiveRunProps != null)
                captureRun.RunProperties = (RunProperties)effectiveRunProps.CloneNode(true);
            var captureText = new Text(text);
            if (DxpFieldEvalMiddleware.NeedsPreserveSpace(text))
                captureText.Space = SpaceProcessingModeValues.Preserve;
            captureRun.AppendChild(captureText);
            using (captureVisitor!.VisitRunBegin(captureRun, d))
                captureVisitor.VisitText(captureText, d);
            return;
        }

        bool useGlobalStyle = Next is DxpContextMiddleware tracker &&
            (tracker.Next is DocxportNet.Visitors.Html.DxpHtmlVisitor ||
                tracker.Next is DocxportNet.Visitors.Markdown.DxpMarkdownVisitor ||
                tracker.Next is DocxportNet.Visitors.PlainText.DxpPlainTextVisitor);
        var sink = useGlobalStyle ? Next : GetSyntheticSink();
        DxpStyleTracker? localStyleTracker = null;
        if (!useGlobalStyle && effectiveRunProps != null)
        {
            var style = BuildEffectiveRunStyle(d, effectiveRunProps);
            localStyleTracker = new DxpStyleTracker();
            localStyleTracker.ApplyStyle(style, d, sink);
        }

        if (_logger?.IsEnabled(LogLevel.Warning) == true && sourceRun == null)
            _logger.LogWarning("Synthetic run emitted for evaluated REF text; no captured run available.");
        var run = sourceRun != null
            ? DxpRunCloner.CloneRunWithParagraphAncestor(sourceRun)
            : new Run();
        if (run.RunProperties == null && effectiveRunProps != null)
            run.RunProperties = (RunProperties)effectiveRunProps.CloneNode(true);
        var t = new Text(text);
        if (DxpFieldEvalMiddleware.NeedsPreserveSpace(text))
            t.Space = SpaceProcessingModeValues.Preserve;
        run.AppendChild(t);
        using (sink.VisitRunBegin(run, d))
            sink.VisitText(t, d);

        localStyleTracker?.ResetStyle(d, sink);
    }

    private void EmitEvaluatedText(
        string text,
        DxpIDocumentContext d,
        IReadOnlyList<Run?>? runs,
        IReadOnlyList<RunProperties?>? runProperties)
    {
        if (string.IsNullOrEmpty(text))
            return;
        if ((runs == null || runs.Count == 0) && (runProperties == null || runProperties.Count == 0))
        {
            EmitEvaluatedText(text, d, (RunProperties?)null, null);
            return;
        }

        int segmentCount = runs?.Count ?? runProperties?.Count ?? 0;
        if (_logger?.IsEnabled(LogLevel.Debug) == true)
            _logger.LogDebug("Ref.Emit.RunSegments: count={Count} hasRuns={HasRuns} hasProps={HasProps}",
                segmentCount,
                runs != null,
                runProperties != null);
        if (runs != null && runs.Count > 0)
            LogRunInfo("Ref.Emit.MergeRun[0]", runs[0]);
        if (segmentCount == 0)
        {
            EmitEvaluatedText(text, d, (RunProperties?)null, null);
            return;
        }

        var segments = SplitTextByRuns(text, segmentCount);
        for (int i = 0; i < segments.Count; i++)
        {
            var segmentRun = runs != null && i < runs.Count ? runs[i] : null;
            var segmentProps = runProperties != null && i < runProperties.Count ? runProperties[i] : null;
            EmitEvaluatedText(segments[i], d, segmentProps, segmentRun);
        }
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

    private void LogRunInfo(string label, Run? run)
    {
        if (_logger?.IsEnabled(LogLevel.Debug) != true)
            return;
        if (run == null)
        {
            _logger.LogDebug("{Label}: <null>", label);
            return;
        }

        var sz = run.RunProperties?.FontSize?.Val?.Value;
        var szCs = run.RunProperties?.FontSizeComplexScript?.Val?.Value;
        var para = run.Ancestors<Paragraph>().FirstOrDefault();
        var paraSz = para?.ParagraphProperties?.GetFirstChild<RunProperties>()?.FontSize?.Val?.Value;
        var paraSzCs = para?.ParagraphProperties?.GetFirstChild<RunProperties>()?.FontSizeComplexScript?.Val?.Value;

        _logger.LogDebug(
            "{Label}: runSz={RunSz} runSzCs={RunSzCs} paraSz={ParaSz} paraSzCs={ParaSzCs}",
            label,
            sz ?? "null",
            szCs ?? "null",
            paraSz ?? "null",
            paraSzCs ?? "null");
    }


    private DxpIVisitor GetSyntheticSink()
    {
        return Next is DxpMiddleware middleware ? middleware.Next : Next;
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
            SmallCaps = defaults.SmallCaps,
            AllCaps = defaults.AllCaps,
            FontName = defaults.FontName,
            FontSizeHalfPoints = defaults.FontSizeHalfPoints
        };

        DxpEffectiveRunStyleBuilder.ApplyRunProperties(runProperties, null, ref acc);
        return acc.ToImmutable();
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

    private static List<string> TokenizeArgs(string? text)
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

    private bool HasMergeFormat(string instruction)
    {
        var parse = _parser.Parse(instruction);
        if (!DxpFieldEvalRules.TryGetCharOrMergeFormat(parse.Ast.FormatSpecs, out _, out var hasMergeFormat))
            return false;
        return hasMergeFormat;
    }

    private string GetEvaluationErrorText(string instruction)
        => DxpFieldEvalRules.GetEvaluationErrorText(_parser, instruction);

    private bool CanEvaluateInCurrentScope()
        => EvalContext.FieldDepth <= 1 || EvalContext.ActiveIfFrame != null;
}
