using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Fields.Eval;
using DocxportNet.Fields.Formatting;
using DocxportNet.Middleware;
using DocxportNet.Walker;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpDocVariableFieldEvalFrame : DxpMiddleware, DxpIFieldEvalFrame
{
	private readonly ILogger? _logger;

	private bool _inCachedResult;
	private string? _instructionText;
	private Run? _codeRun;
	private List<Run?>? _cachedResultRuns;

	public override DxpIVisitor? Next { get; }

	private readonly DxpFieldEval _eval;

	public DxpDocVariableFieldEvalFrame(
		DxpIVisitor? next,
		DxpFieldEval eval,
		ILogger? logger,
		Run? codeRun = null,
		string? instructionText = null)
		: base()
	{
		Next = next;
		_eval = eval ?? throw new ArgumentNullException(nameof(eval));
		_logger = logger;
		_codeRun = codeRun;
		_instructionText = instructionText;
	}

	public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
	{
		if (string.IsNullOrEmpty(text) || _inCachedResult)
			return;
		if (_codeRun == null && instr.Parent is Run instrRun)
			_codeRun = DxpRunCloner.CloneRunWithParagraphAncestor(instrRun);
	}

	public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext d)
	{
		_inCachedResult = true;
	}

	public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
	{
		EvaluateDocVariable(d);
	}

	public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
	{
		return;
	}

	public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
	{
		_inCachedResult = true;
		return DxpDisposable.Create(() => {
			EvaluateDocVariable(d);
			_inCachedResult = false;
		});
	}

	public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
	{
		if (_inCachedResult && DxpFieldEvalRules.HasRenderableContent(r))
		{
			_cachedResultRuns ??= new List<Run?>();
			_cachedResultRuns.Add(DxpRunCloner.CloneRunWithParagraphAncestor(r));
		}

		return DxpDisposable.Empty;
	}

	public override IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target, DxpIDocumentContext d)
	{
		return DxpDisposable.Empty;
	}

	public override void VisitText(Text t, DxpIDocumentContext d)
	{
		return;
	}

	public override void VisitBreak(Break br, DxpIDocumentContext d)
	{
		return;
	}

	public override void VisitTab(TabChar tab, DxpIDocumentContext d)
	{
		return;
	}

	public override void VisitCarriageReturn(CarriageReturn cr, DxpIDocumentContext d)
	{
		return;
	}

	public override void VisitNoBreakHyphen(NoBreakHyphen nbh, DxpIDocumentContext d)
	{
		return;
	}

	private bool EvaluateDocVariable(DxpIDocumentContext d)
	{
		if (string.IsNullOrWhiteSpace(_instructionText))
			return false;

		string instruction = _instructionText!;

		DxpFieldParser parser = new();
		DxpFieldParseResult parse = parser.Parse(instruction);
		string? argsText = parse.Ast.ArgumentsText;

		if (!DxpFieldTokenization.TryGetFirstToken(argsText, out _))
			return false;

		DxpFieldEvalResult result = _eval.EvalAsync(new DxpFieldInstruction(instruction), d).GetAwaiter().GetResult();
		string? resultText = result.Text;

		var formatText = result.Status == DxpFieldEvalStatus.Resolved && resultText != null
			? resultText
			: DxpFieldEvalRules.GetEvaluationErrorText(instruction);


		IReadOnlyList<IDxpFieldFormatSpec> formatSpecs = parse.Ast.FormatSpecs;
		return DxpFieldFrames.EmitTextWithMergeFormat(formatText, formatSpecs, _cachedResultRuns, _codeRun, d, Next, _logger);
	}
}
