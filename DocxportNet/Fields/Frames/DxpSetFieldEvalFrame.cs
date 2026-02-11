using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Middleware;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpSetFieldEvalFrame : DxpMiddleware, DxpIFieldEvalFrame
{
	private readonly DxpFieldEval _fieldEvaluator;
	private readonly DxpFieldEvalContext _evaluationContext;

	public string? InstructionText { get; }

	public DxpSetFieldEvalFrame(DxpFieldEval eval, DxpFieldEvalContext evalContext, ILogger? logger, string? instructionText)
		: base()
	{
		_fieldEvaluator = eval ?? throw new ArgumentNullException(nameof(eval));
		_evaluationContext = evalContext ?? throw new ArgumentNullException(nameof(evalContext));
		InstructionText = instructionText;
	}

	// DxpMiddleware
	// we never forward anything for a Set Field
	public override DxpIVisitor? Next => null;

	public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
	{
		EvaluateSet(d);
	}

	public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
	{
		EvaluateSet(d);
		return DxpDisposable.Empty;
	}

	private void EvaluateSet(DxpIDocumentContext d)
	{
		if (string.IsNullOrWhiteSpace(InstructionText))
			return;

		DxpFieldParser fieldParser = new();
		DxpFieldParseResult parseResult = fieldParser.Parse(InstructionText!);

		var argsText = parseResult.Ast.ArgumentsText;
		if (!DxpFieldTokenization.TryGetFirstToken(argsText, out var setName))
			return;

		var setResult = _fieldEvaluator.EvalAsync(new DxpFieldInstruction(InstructionText!), d).GetAwaiter().GetResult();
		var text = setResult.Text ?? string.Empty;
		_evaluationContext.SetBookmarkNodes(setName, DxpFieldNodeBuffer.FromText(text));
	}
}
