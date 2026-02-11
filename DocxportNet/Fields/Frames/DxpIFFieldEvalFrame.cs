using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Fields.Eval;
using DocxportNet.Middleware;
using DocxportNet.Walker;
using Microsoft.Extensions.Logging;
using System.Text;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpIFFieldEvalFrame : DxpMiddleware, DxpIFieldEvalFrame
{
	private readonly ILogger? _logger;

	private bool _inCachedResult;
	private string? _instructionText;
	private Run? _codeRun;
	private DxpIFCaptureState? _ifState;
	private Run? _currentInstructionRun;
	private StringBuilder? _literalBuffer;
	private bool _hasPendingLiteral;

	public override DxpIVisitor? Next { get; }

	private readonly DxpFieldEval _eval;

	public DxpIFFieldEvalFrame(DxpIVisitor next, DxpFieldEval eval, ILogger? logger, Run? codeRun = null)
		: base()
	{
		Next = next;
		_eval = eval ?? throw new ArgumentNullException(nameof(eval));
		_logger = logger;
		_codeRun = codeRun;
	}

	public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
	{
		if (string.IsNullOrEmpty(text) || _inCachedResult)
			return;

		FlushLiteralBuffer();
		var state = DxpFieldEvalIfRunner.EnsureIfState(ref _ifState);
		_instructionText = _instructionText == null ? text : _instructionText + text;
		var instrRun = instr.Parent as Run;
		var runProps = instrRun?.RunProperties;

		if (_codeRun == null && instrRun != null)
			_codeRun = DxpRunCloner.CloneRunWithParagraphAncestor(instrRun);

		DxpFieldEvalIfRunner.ProcessInstructionSegment(state, text, instrRun, runProps);

	}

	public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext d)
	{
		_inCachedResult = true;
	}

	public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
	{
		FlushLiteralBuffer();
		if (_ifState != null)
			DxpFieldEvalIfRunner.TryEvaluateAndEmit(_ifState, _instructionText ?? string.Empty, _eval, d, Next,
				DxpFieldEvalRules.GetEvaluationErrorText, EmitEvaluatedText);
	}

	public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
	{
		return;
	}

	public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
	{
		DxpFieldEvalIfRunner.EnsureIfState(ref _ifState);
		_inCachedResult = true;
		return DxpDisposable.Create(() => {
			FlushLiteralBuffer();
			if (_ifState != null)
				DxpFieldEvalIfRunner.TryEvaluateAndEmit(_ifState, _instructionText ?? string.Empty, _eval, d, Next, 
					DxpFieldEvalRules.GetEvaluationErrorText, EmitEvaluatedText);
			_inCachedResult = false;
		});
	}

	public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
	{
		if (!_inCachedResult)
		{
			_currentInstructionRun = r;
			return DxpDisposable.Empty;
		}
		return DxpDisposable.Empty;
	}

	public override IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target, DxpIDocumentContext d)
	{
		return DxpDisposable.Empty;
	}

	public override void VisitText(Text t, DxpIDocumentContext d)
	{
		if (!_inCachedResult)
		{
			BufferLiteralToken(t.Text);
			return;
		}
		return;
	}

	public override void VisitBreak(Break br, DxpIDocumentContext d)
	{
		if (!_inCachedResult)
		{
			BufferLiteralToken("\n");
			return;
		}
		return;
	}

	public override void VisitTab(TabChar tab, DxpIDocumentContext d)
	{
		if (!_inCachedResult)
		{
			BufferLiteralToken("\t");
			return;
		}
		return;
	}

	public override void VisitCarriageReturn(CarriageReturn cr, DxpIDocumentContext d)
	{
		if (!_inCachedResult)
		{
			BufferLiteralToken("\n");
			return;
		}
		return;
	}

	public override void VisitNoBreakHyphen(NoBreakHyphen nbh, DxpIDocumentContext d)
	{
		if (!_inCachedResult)
		{
			BufferLiteralToken("-");
			return;
		}
		return;
	}

	private void EmitEvaluatedText(string text, DxpIDocumentContext d, RunProperties? runProperties)
		=> EmitEvaluatedText(text, d, runProperties, _codeRun);

	private void BufferLiteralToken(string value)
	{
		_literalBuffer ??= new StringBuilder();
		_hasPendingLiteral = true;
		_literalBuffer.Append(value);
	}

	private void FlushLiteralBuffer()
	{
		if (!_hasPendingLiteral)
			return;
		var state = DxpFieldEvalIfRunner.EnsureIfState(ref _ifState);
		var value = _literalBuffer?.ToString() ?? string.Empty;
		_literalBuffer?.Clear();
		_hasPendingLiteral = false;
		var instructionText = _instructionText ?? string.Empty;
		var needsSpace = !state.InQuote &&
			instructionText.Length > 0 &&
			!char.IsWhiteSpace(instructionText[instructionText.Length - 1]);
		var segment = state.InQuote
			? value
			: (needsSpace ? " " : string.Empty) + FormatIfLiteralToken(value);
		_instructionText = instructionText + segment;
		DxpFieldEvalIfRunner.ProcessInstructionSegment(state, segment, _currentInstructionRun, _currentInstructionRun?.RunProperties);
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

	private void EmitEvaluatedText(string text, DxpIDocumentContext d, RunProperties? runProperties = null, Run? sourceRun = null)
	{
		var fallbackRun = sourceRun ?? _codeRun;
		var run = DxpFieldFrames.NewSyntheticRun(fallbackRun, runProperties);
		DxpFieldFrames.EmitTextInRun(text, d, run, Next);
	}
}
