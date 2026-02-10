using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Fields.Eval;
using DocxportNet.Middleware;
using DocxportNet.Walker;
using DocxportNet.Walker.Context;
using Microsoft.Extensions.Logging;
using System.Text;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpRefFieldCachedFrame : DxpMiddleware, DxpIFieldEvalFrame
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

	private readonly DxpFieldParser _parser = new();
	private readonly ILogger? _logger;

	public DxpRefFieldCachedFrame(DxpIVisitor next, DxpFieldEvalContext evalContext, ILogger? logger)
		: base()
	{
		Next = next ?? throw new ArgumentNullException(nameof(next));
		EvalContext = evalContext ?? throw new ArgumentNullException(nameof(evalContext));
		_logger = logger;
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
			if (ReferenceEquals(EvalContext.ActiveIfFrame, this))
				EvalContext.ActiveIfFrame = null;
		}
	}

	public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
	{
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

		if (CachedResultRunProperties != null && CachedResultRunProperties.Count > 0)
			EmitEvaluatedText(text, d, CachedResultRunProperties);
		else
			Next.VisitComplexFieldCachedResultText(text, d);
	}

	public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
	{
		return DxpDisposable.Empty;
	}

	public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
	{
		if (InResult && DxpFieldEvalRules.HasRenderableContent(r))
		{
			CachedResultRunProperties ??= new List<RunProperties?>();
			RunProperties? props = r.RunProperties != null ? (RunProperties)r.RunProperties.CloneNode(true) : null;
			CachedResultRunProperties.Add(props);
		}

		return base.VisitRunBegin(r, d);
	}

	protected override bool ShouldForwardContent(DxpIDocumentContext d)
	{
		if (EvalContext.OuterFrame == null)
			return true;
		if (!EvalContext.OuterFrame.InResult)
			return false;
		if (EvalContext.OuterFrame.SuppressContent)
			return false;
		return EvalContext.FieldDepth <= 1;
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

	private void EmitEvaluatedText(string text, DxpIDocumentContext d, RunProperties? runProperties = null)
	{
		if (string.IsNullOrEmpty(text))
			return;
		var effectiveRunProps = runProperties;
		if (_logger?.IsEnabled(LogLevel.Debug) == true)
			_logger.LogDebug("EmitEvaluatedText: text='{Text}' runProps={RunProps}",
				text,
				effectiveRunProps != null);

		if (TryGetCaptureVisitor(out var captureVisitor, out _))
		{
			var captureRun = new Run();
			if (effectiveRunProps != null)
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

		if (_logger?.IsEnabled(LogLevel.Warning) == true)
			_logger.LogWarning("Synthetic run emitted for cached REF text; no captured run available.");
		var run = new Run();
		if (effectiveRunProps != null)
			run.RunProperties = (RunProperties)effectiveRunProps.CloneNode(true);
		var t = new Text(text);
		if (DxpFieldEvalMiddleware.NeedsPreserveSpace(text))
			t.Space = SpaceProcessingModeValues.Preserve;
		run.AppendChild(t);
		using (sink.VisitRunBegin(run, d))
			sink.VisitText(t, d);

		localStyleTracker?.ResetStyle(d, sink);
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
