using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Fields.Frames;
using DocxportNet.Walker;
using System.Text;

namespace DocxportNet.Fields.Eval;

internal static class DxpFieldEvalIfRunner
{
    public static void EnsureIfState(DxpIFieldEvalFrame frame)
    {
        frame.IfState ??= new DxpIFCaptureState();
    }

    public static void ProcessInstructionSegment(DxpIFieldEvalFrame frame, string text, Run? run, RunProperties? runProps)
    {
        var state = frame.IfState;
        if (state == null || string.IsNullOrEmpty(text))
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
                    AppendBufferText(currentTarget, bufferText.ToString(), run, runProps);
                    bufferText.Clear();
                }
                currentTarget = nextTarget;
            }
            if (currentTarget != null)
                bufferText.Append(ch);
        }

        if (currentTarget != null && bufferText.Length > 0)
            AppendBufferText(currentTarget, bufferText.ToString(), run, runProps);
    }

    public static bool TryEvaluateAndEmit(
        DxpIFieldEvalFrame frame,
        DxpFieldEval eval,
        DxpIDocumentContext documentContext,
        DxpIVisitor next,
        Func<string, string> errorTextProvider,
        Action<string, DxpIDocumentContext, RunProperties?> emitText)
    {
        if (frame.IfState == null || string.IsNullOrWhiteSpace(frame.InstructionText))
            return false;

        string instructionText = frame.InstructionText!;
        var ifResult = eval.EvaluateIfConditionAsync(instructionText, documentContext).GetAwaiter().GetResult();
        if (ifResult == null || !ifResult.Value.Success)
        {
            frame.Evaluated = true;
            frame.SuppressContent = true;
            emitText(errorTextProvider(instructionText), documentContext, null);
            return true;
        }

        var selected = ifResult.Value.Condition ? frame.IfState.TrueBuffer : frame.IfState.FalseBuffer;
        frame.Evaluated = true;
        frame.SuppressContent = true;
        if (selected.IsEmpty)
        {
            var evalResult = eval.EvalAsync(new DxpFieldInstruction(instructionText), documentContext).GetAwaiter().GetResult();
            if (evalResult.Status == DxpFieldEvalStatus.Resolved && evalResult.Text != null)
                emitText(evalResult.Text, documentContext, null);
            else
                emitText(errorTextProvider(instructionText), documentContext, null);
            return true;
        }

        selected.Replay(next, documentContext);
        return true;
    }

    private static void AppendBufferText(DxpFieldNodeBuffer buffer, string text, Run? run, RunProperties? runProps)
    {
        if (string.IsNullOrEmpty(text))
            return;
        var runClone = run != null
            ? DxpRunCloner.CloneRunWithParagraphAncestor(run)
            : new Run();
        if (runClone.RunProperties == null && runProps != null)
            runClone.RunProperties = (RunProperties)runProps.CloneNode(true);
        var t = new Text(text);
        if (DxpFieldEvalMiddleware.NeedsPreserveSpace(text))
            t.Space = SpaceProcessingModeValues.Preserve;
        runClone.AppendChild(t);
        var child = buffer.BeginRun(runClone);
        child.AddText(text);
    }
}
