namespace DocxportNet.Fields;

public sealed record DxpFillInRequest(
    string InstructionText,
    string? PromptText,
    string? CachedResultText,
    string? DefaultText,
    bool AskOnce,
    string? PriorResponseText);
