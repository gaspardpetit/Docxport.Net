using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Fields.Eval;
using DocxportNet.Fields.Formatting;
using DocxportNet.Middleware;
using DocxportNet.Walker;
using DocxportNet.Walker.Context;
using Microsoft.Extensions.Logging;
using System.Text;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpDocVariableFieldEvalFrame : DxpMiddleware, DxpIFieldEvalFrame
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

	public DxpDocVariableFieldEvalFrame(DxpIVisitor next, DxpFieldEval eval, DxpFieldEvalContext evalContext, ILogger? logger)
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
			LogRunInfo("DocVariable.CodeRunCaptured", CodeRun);
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
				EvaluateDocVariable(d);
			if (!Evaluated && ShouldDeferEvaluation())
				EvaluateDocVariable(d);
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
		if ((CanEvaluateInCurrentScope()) && EvaluateDocVariable(d))
			return;
	}

	public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
	{
		if (CanEvaluateInCurrentScope())
		{
			if (IfState == null && !ShouldDeferEvaluation() && EvaluateDocVariable(d))
				return DxpDisposable.Empty;

			if (ShouldDeferEvaluation() && CanEvaluateInCurrentScope())
			{
				return DxpDisposable.Create(() => {
					if (!Evaluated)
						EvaluateDocVariable(d);
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

	private bool EvaluateDocVariable(DxpIDocumentContext d)
	{
		if (Evaluated)
			return false;
		if (string.IsNullOrWhiteSpace(InstructionText))
			return false;

		var instruction = InstructionText!;
		var parse = _parser.Parse(instruction);
		var argsText = parse.Ast.ArgumentsText;

		if (TryHandleDocVariableWithFormat(d, instruction, argsText, parse.Ast.FormatSpecs))
			return true;

		if (!TryGetFirstToken(argsText, out var docVarName))
			return false;

		var evalResult = _eval.EvalAsync(new DxpFieldInstruction(instruction), d).GetAwaiter().GetResult();
		Evaluated = true;
		SuppressContent = true;

		if (parse.Ast.FormatSpecs.Count > 0 &&
			evalResult.Status == DxpFieldEvalStatus.Resolved &&
			evalResult.Text != null)
		{
			if (TryInjectIfInstructionValue(evalResult.Text))
				return true;
			EmitEvaluatedText(evalResult.Text, d, CodeRunProperties, CodeRun);
			return true;
		}

		if (!EvalContext.TryGetDocVariableNodes(docVarName, out var docVarNodes))
		{
			var fallbackText = evalResult.Status == DxpFieldEvalStatus.Resolved && evalResult.Text != null
				? evalResult.Text
				: GetEvaluationErrorText(instruction);
			if (TryInjectIfInstructionValue(fallbackText))
				return true;
			EmitEvaluatedText(fallbackText, d, CodeRunProperties, CodeRun);
			return true;
		}

		var text = docVarNodes.ToPlainText();
		if (TryInjectIfInstructionValue(text))
			return true;

		if (docVarNodes.TryGetFirstRunProperties(out var docVarProps) && docVarProps != null)
		{
			var replayVisitor = TryGetCaptureVisitor(out var captureVisitor, out _) ? captureVisitor! : Next;
			docVarNodes.Replay(replayVisitor, d);
		}
		else if (CodeRunProperties != null)
		{
			EmitEvaluatedText(text, d, CodeRunProperties, CodeRun);
		}
		else
		{
			var replayVisitor = TryGetCaptureVisitor(out var captureVisitor, out _) ? captureVisitor! : Next;
			docVarNodes.Replay(replayVisitor, d);
		}
		return true;
	}

	private bool TryHandleDocVariableWithFormat(
		DxpIDocumentContext d,
		string instruction,
		string? argsText,
		IReadOnlyList<IDxpFieldFormatSpec> formatSpecs)
	{
		if (!DxpFieldEvalFrameFactory.IsDocVariableInstruction(instruction))
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
		if (_logger?.IsEnabled(LogLevel.Debug) == true)
			_logger.LogDebug("EmitEvaluatedText: text='{Text}' runProps={RunProps}",
				text,
				runProperties != null);
		LogRunInfo("DocVariable.Emit.SourceRun", sourceRun);
		LogRunInfo("DocVariable.Emit.CodeRun", CodeRun);

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
        if (shouldEmitStyle)
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

		if (_logger?.IsEnabled(LogLevel.Warning) == true && sourceRun == null)
			_logger.LogWarning("Synthetic run emitted for evaluated DOCVARIABLE text; no captured run available.");
		var run = sourceRun != null
			? DxpRunCloner.CloneRunWithParagraphAncestor(sourceRun)
			: new Run();
		if (run.RunProperties == null && runProperties != null)
			run.RunProperties = (RunProperties)runProperties.CloneNode(true);
		var t = new Text(text);
		if (DxpFieldEvalMiddleware.NeedsPreserveSpace(text))
			t.Space = SpaceProcessingModeValues.Preserve;
		run.AppendChild(t);
		using (Next.VisitRunBegin(run, d))
			Next.VisitText(t, d);

        if (shouldEmitStyle)
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
			_logger.LogDebug("DocVariable.Emit.RunSegments: count={Count} hasRuns={HasRuns} hasProps={HasProps}",
				segmentCount,
				runs != null,
				runProperties != null);
		if (runs != null && runs.Count > 0)
			LogRunInfo("DocVariable.Emit.MergeRun[0]", runs[0]);
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

	private bool CanEvaluateInCurrentScope()
		=> EvalContext.FieldDepth <= 1 || EvalContext.ActiveIfFrame != null;

	private static string GetEvaluationErrorText(string instructionText)
		=> DxpFieldEvalRules.GetEvaluationErrorText(new DxpFieldParser(), instructionText);

	private static bool HasMergeFormat(string instructionText)
		=> DxpFieldEvalRules.HasMergeFormat(new DxpFieldParser(), instructionText);

	private static IReadOnlyList<string> SplitTextByRuns(string text, int runCount)
	{
		if (runCount <= 1 || text.Length == 0)
			return new List<string> { text };

		int total = text.Length;
		int baseSize = total / runCount;
		int remainder = total % runCount;
		var segments = new List<string>(runCount);
		int index = 0;
		for (int i = 0; i < runCount; i++)
		{
			int size = baseSize + (i < remainder ? 1 : 0);
			segments.Add(text.Substring(index, size));
			index += size;
		}
		return segments;
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

}
