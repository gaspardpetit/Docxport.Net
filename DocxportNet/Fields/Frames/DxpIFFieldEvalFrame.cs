using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Fields.Eval;
using DocxportNet.Middleware;
using DocxportNet.Walker;
using DocxportNet.Walker.Context;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpIFFieldEvalFrame : DxpMiddleware, DxpIFieldEvalFrame
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
    private readonly DxpEvalFieldMode _mode;
    private readonly DxpFieldParser _parser = new();
    private readonly ILogger? _logger;

    public DxpIFFieldEvalFrame(DxpIVisitor next, DxpFieldEval eval, DxpFieldEvalContext evalContext, ILogger? logger, DxpEvalFieldMode mode)
        : base()
    {
		Next = next ?? throw new ArgumentNullException(nameof(next));
		_eval = eval ?? throw new ArgumentNullException(nameof(eval));
        EvalContext = evalContext ?? throw new ArgumentNullException(nameof(evalContext));
        _logger = logger;
        _mode = mode;
    }

    public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
    {
        if (string.IsNullOrEmpty(text) || InResult)
            return;

        DxpFieldEvalIfRunner.EnsureIfState(this);
        if (EvalContext.ActiveIfFrame == null)
            EvalContext.ActiveIfFrame = this;
        var instrRun = instr.Parent as Run;
        var runProps = instrRun?.RunProperties;
        if (CodeRun == null)
        {
            CodeRun = instrRun;
            LogRunInfo("IF.CodeRunCaptured", CodeRun);
        }
        DxpFieldEvalIfRunner.ProcessInstructionSegment(this, text, instrRun, runProps);

        if (CodeRunProperties == null)
        {
            if (instr.Parent is Run parInstrRun && parInstrRun.RunProperties != null)
                CodeRunProperties = (RunProperties)parInstrRun.RunProperties.CloneNode(true);
        }
    }

    public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext d)
    {
        if (!SeenSeparate)
        {
            SeenSeparate = true;
            InResult = true;
            ClearActiveIf();
        }

        if (_mode == DxpEvalFieldMode.Evaluate && EvalContext.FieldDepth <= 1)
        {
            if (IfState != null && !Evaluated)
                DxpFieldEvalIfRunner.TryEvaluateAndEmit(this, _eval, d, Next, GetEvaluationErrorText, EmitEvaluatedText);
            if (Evaluated)
                ClearActiveIf();
        }
    }

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
    {
        if (_mode == DxpEvalFieldMode.Evaluate && EvalContext.FieldDepth <= 1)
        {
            if (IfState != null && !Evaluated)
                DxpFieldEvalIfRunner.TryEvaluateAndEmit(this, _eval, d, Next, GetEvaluationErrorText, EmitEvaluatedText);
            if (Evaluated)
                ClearActiveIf();
        }
        else if (ReferenceEquals(EvalContext.ActiveIfFrame, this))
        {
            ClearActiveIf();
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

        if (_mode == DxpEvalFieldMode.Cache)
        {
            if (EvalContext.OuterFrame?.SuppressContent == true)
                return;
            if (EvalContext.FieldDepth > 1)
                return;
            if (EvalContext.ActiveIfFrame != null)
                return;
            if (Evaluated)
                return;

            if (!InstructionEmitted && !string.IsNullOrWhiteSpace(InstructionText))
            {
                InstructionEmitted = true;
                Next.VisitComplexFieldInstruction(new FieldCode(), InstructionText!, d);
            }

            if (!ShouldForwardContent(d))
                return;

            Next.VisitComplexFieldCachedResultText(text, d);
        }
    }

    public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
    {
        if (_mode == DxpEvalFieldMode.Cache && !string.IsNullOrWhiteSpace(InstructionText))
            InstructionEmitted = true;

        if (_mode == DxpEvalFieldMode.Evaluate && CanEvaluateInCurrentScope())
        {
            DxpFieldEvalIfRunner.EnsureIfState(this);
            if (EvalContext.ActiveIfFrame == null)
                EvalContext.ActiveIfFrame = this;
            if (!Evaluated && IfState != null)
                DxpFieldEvalIfRunner.TryEvaluateAndEmit(this, _eval, d, Next, GetEvaluationErrorText, EmitEvaluatedText);
            if (Evaluated)
                ClearActiveIf();
        }

        return DxpDisposable.Empty;
    }

    protected override bool ShouldForwardContent(DxpIDocumentContext d)
    {
        if (_mode == DxpEvalFieldMode.Cache)
        {
            if (EvalContext.OuterFrame == null)
                return true;
            if (!EvalContext.OuterFrame.InResult)
                return false;
            if (EvalContext.OuterFrame.SuppressContent)
                return false;
            return EvalContext.FieldDepth <= 1;
        }

        if (InResult != true)
            return true;
        if (EvalContext.FieldDepth > 1 && EvalContext.ActiveIfFrame == null)
            return false;
        if (SuppressContent)
            return false;
        return true;
    }

    private bool TryGetCaptureVisitor(out DxpIVisitor? visitor, out DxpFieldNodeBuffer? buffer)
    {
        visitor = null;
        buffer = null;
        if (Evaluated && ReferenceEquals(EvalContext.ActiveIfFrame, this))
            return false;
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
        if (_logger?.IsEnabled(LogLevel.Debug) == true)
            _logger.LogDebug("EmitEvaluatedText: text='{Text}' runProps={RunProps}",
                text,
                runProperties != null);
        LogRunInfo("IF.Emit.SourceRun", sourceRun);
        LogRunInfo("IF.Emit.CodeRun", CodeRun);

        if (TryGetCaptureVisitor(out var captureVisitor, out _))
        {
            var captureRun = sourceRun != null
                ? DxpRunCloner.CloneRunWithParagraphAncestor(sourceRun)
                : new Run();
            if (captureRun.RunProperties == null && runProperties != null)
                captureRun.RunProperties = (RunProperties)runProperties.CloneNode(true);
            var captureText = new Text(text);
            if (DxpFieldEvalMiddleware.NeedsPreserveSpace(text))
                captureText.Space = SpaceProcessingModeValues.Preserve;
            captureRun.AppendChild(captureText);
            using (captureVisitor!.VisitRunBegin(captureRun, d))
                captureVisitor.VisitText(captureText, d);
            return;
        }

        var effectiveStyle = ApplyRunProperties(d, runProperties);
        var shouldEmitStyle = Next is not DxpContextMiddleware;
        if (d is IDxpMutableDocumentContext doc)
        {
            doc.StyleTracker.ApplyStyle(effectiveStyle, d, Next);
        }
        else if (shouldEmitStyle)
        {
            if (effectiveStyle.Bold)
				Next.StyleBoldBegin(d);
            if (effectiveStyle.Italic)
				Next.StyleItalicBegin(d);
            if (effectiveStyle.Underline)
				Next.StyleUnderlineBegin(d);
            if (effectiveStyle.Strike)
				Next.StyleStrikeBegin(d);
            if (effectiveStyle.DoubleStrike)
				Next.StyleDoubleStrikeBegin(d);
            if (effectiveStyle.Superscript)
				Next.StyleSuperscriptBegin(d);
            if (effectiveStyle.Subscript)
				Next.StyleSubscriptBegin(d);
            if (effectiveStyle.SmallCaps)
				Next.StyleSmallCapsBegin(d);
            if (effectiveStyle.AllCaps)
				Next.StyleAllCapsBegin(d);
        }

        var fallbackRun = sourceRun ?? CodeRun;
        if (_logger?.IsEnabled(LogLevel.Warning) == true && fallbackRun == null)
            _logger.LogWarning("Synthetic run emitted for evaluated IF text; no captured run available.");
        var run = fallbackRun != null
            ? DxpRunCloner.CloneRunWithParagraphAncestor(fallbackRun)
            : new Run();
        if (run.RunProperties == null && runProperties != null)
            run.RunProperties = (RunProperties)runProperties.CloneNode(true);
        var t = new Text(text);
        if (DxpFieldEvalMiddleware.NeedsPreserveSpace(text))
            t.Space = SpaceProcessingModeValues.Preserve;
        run.AppendChild(t);
        using (Next.VisitRunBegin(run, d))
			Next.VisitText(t, d);

        if (d is not IDxpMutableDocumentContext && shouldEmitStyle)
        {
            if (effectiveStyle.AllCaps)
				Next.StyleAllCapsEnd(d);
            if (effectiveStyle.SmallCaps)
				Next.StyleSmallCapsEnd(d);
            if (effectiveStyle.Subscript)
				Next.StyleSubscriptEnd(d);
            if (effectiveStyle.Superscript)
				Next.StyleSuperscriptEnd(d);
            if (effectiveStyle.DoubleStrike)
				Next.StyleDoubleStrikeEnd(d);
            if (effectiveStyle.Strike)
				Next.StyleStrikeEnd(d);
            if (effectiveStyle.Underline)
				Next.StyleUnderlineEnd(d);
            if (effectiveStyle.Italic)
				Next.StyleItalicEnd(d);
            if (effectiveStyle.Bold)
				Next.StyleBoldEnd(d);
        }
    }

    private static DxpStyleEffectiveRunStyle ApplyRunProperties(DxpIDocumentContext d, RunProperties? runProperties)
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


    private bool CanEvaluateInCurrentScope()
        => EvalContext.FieldDepth <= 1 || EvalContext.ActiveIfFrame != null;

    private static string GetEvaluationErrorText(string instructionText)
        => DxpFieldEvalRules.GetEvaluationErrorText(new DxpFieldParser(), instructionText);

    private void ClearActiveIf()
    {
        if (ReferenceEquals(EvalContext.ActiveIfFrame, this))
            EvalContext.ActiveIfFrame = null;
    }
}
