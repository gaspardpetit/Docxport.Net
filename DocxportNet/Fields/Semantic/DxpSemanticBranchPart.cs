using DocxportNet.Fields.Eval;

namespace DocxportNet.Fields.Semantic;

internal abstract record DxpSemanticBranchPart;
internal sealed record DxpSemanticBranchText(string Text) : DxpSemanticBranchPart;
internal sealed record DxpSemanticBranchField(DxpDeferredField Field) : DxpSemanticBranchPart;
internal sealed record DxpSemanticBranchParagraphStart : DxpSemanticBranchPart;
