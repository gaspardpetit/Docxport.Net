using DocxportNet.API;
using DocxportNet.Fields.Frames;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Eval;

internal sealed class DxpEvaluateFieldFrameFactory
{
    public DxpIFieldEvalFrame? Create(
        string? instruction,
        DxpIVisitor next,
        DxpFieldEval eval,
        DxpFieldEvalContext context,
        ILogger? logger)
    {
        if (context.PreserveLayoutDependentFields &&
            DxpFieldInstructionClassifier.IsPaginationDependentInstruction(instruction))
            return new DxpPassthroughFieldEvalFrame(next);

        if (DxpFieldInstructionClassifier.IsRefInstruction(instruction))
            return new DxpRefFieldEvalFrame(next, eval, logger, instruction);

        if (DxpFieldInstructionClassifier.IsDocVariableInstruction(instruction))
            return new DxpDocVariableFieldEvalFrame(next, eval, logger, instruction);

        if (DxpFieldInstructionClassifier.IsIfInstruction(instruction))
            return new DxpIFFieldEvalFrame(next, eval, logger);

        if (DxpFieldInstructionClassifier.IsSetInstruction(instruction))
            return new DxpSetFieldEvalFrame(eval, context, logger, instruction);

        if (DxpFieldInstructionClassifier.IsAskInstruction(instruction))
            return new DxpAskFieldEvalFrame(next, eval, logger, instruction);

        if (DxpFieldInstructionClassifier.IsFillInInstruction(instruction))
            return new DxpValueFieldEvalFrame(next, eval, logger, instruction);

        if (DxpFieldInstructionClassifier.IsIncludeTextInstruction(instruction))
            return new DxpIncludeTextFieldEvalFrame(next, eval, logger, instruction);

        if (context.EmitStructuredDatabaseResults &&
            DxpFieldInstructionClassifier.IsDatabaseInstruction(instruction))
            return new DxpDatabaseFieldEvalFrame(next, eval, instruction!);

        if (DxpFieldInstructionClassifier.IsLayoutDependentInstruction(instruction))
            return new DxpLayoutCachedFieldEvalFrame(next, eval, logger, instruction);

        if (DxpFieldInstructionClassifier.IsNextInstruction(instruction))
            return new DxpValueFieldEvalFrame(next, eval, logger, instruction);

        if (DxpFieldInstructionClassifier.IsSkipIfInstruction(instruction))
            return new DxpSkipIfFieldEvalFrame(next, eval, logger, instruction);

        if (DxpFieldInstructionClassifier.IsDocPropertyInstruction(instruction) ||
            DxpFieldInstructionClassifier.IsMergeFieldInstruction(instruction) ||
            DxpFieldInstructionClassifier.IsMergeRecInstruction(instruction) ||
            DxpFieldInstructionClassifier.IsMergeSeqInstruction(instruction) ||
            DxpFieldInstructionClassifier.IsGreetingLineInstruction(instruction) ||
            DxpFieldInstructionClassifier.IsAddressBlockInstruction(instruction) ||
            DxpFieldInstructionClassifier.IsAutoNumberInstruction(instruction) ||
            DxpFieldInstructionClassifier.IsDatabaseInstruction(instruction) ||
            DxpFieldInstructionClassifier.IsSeqInstruction(instruction) ||
            DxpFieldInstructionClassifier.IsDateTimeInstruction(instruction) ||
            DxpFieldInstructionClassifier.IsCompareInstruction(instruction) ||
            DxpFieldInstructionClassifier.IsDocumentMetricInstruction(instruction))
        {
            if (DxpFieldInstructionClassifier.IsDocPropertyInstruction(instruction))
                return new DxpDocPropertyFieldEvalFrame(next, eval, logger, instruction);

            if (DxpFieldInstructionClassifier.IsMergeFieldInstruction(instruction))
                return new DxpMergeFieldEvalFrame(next, eval, logger, instruction);

            if (DxpFieldInstructionClassifier.IsMergeRecInstruction(instruction) ||
                DxpFieldInstructionClassifier.IsMergeSeqInstruction(instruction) ||
                DxpFieldInstructionClassifier.IsGreetingLineInstruction(instruction) ||
                DxpFieldInstructionClassifier.IsAddressBlockInstruction(instruction) ||
                DxpFieldInstructionClassifier.IsAutoNumberInstruction(instruction) ||
                DxpFieldInstructionClassifier.IsDatabaseInstruction(instruction) ||
                DxpFieldInstructionClassifier.IsDocumentMetricInstruction(instruction))
            {
                return new DxpValueFieldEvalFrame(next, eval, logger, instruction);
            }

            if (DxpFieldInstructionClassifier.IsSeqInstruction(instruction))
                return new DxpSeqFieldEvalFrame(next, eval, logger, instruction);

            if (DxpFieldInstructionClassifier.IsDateTimeInstruction(instruction))
                return new DxpDateTimeFieldEvalFrame(next, eval, logger, instruction);

            if (DxpFieldInstructionClassifier.IsCompareInstruction(instruction))
                return new DxpCompareFieldEvalFrame(next, eval, logger, instruction);
        }

        if (DxpFieldInstructionClassifier.IsFormulaInstruction(instruction))
            return new DxpFormulaFieldEvalFrame(next, eval, logger, instruction);

        if (DxpFieldInstructionClassifier.TryGetImplicitRefName(instruction, context, out string bookmark))
            return new DxpRefFieldEvalFrame(
                next,
                eval,
                logger,
                DxpFieldInstructionClassifier.RewriteImplicitRefInstruction(instruction!, bookmark));

        if (logger?.IsEnabled(LogLevel.Debug) == true)
            logger.LogDebug("FieldFrameFactory: no evaluate frame for instruction '{Instruction}'.", instruction ?? string.Empty);
        return null;
    }
}
