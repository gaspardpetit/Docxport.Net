namespace DocxportNet.Fields;

public sealed record DxpAskRequest(
    string BookmarkName,
    string InstructionText,
    string? PromptText,
    string? DefaultText,
    bool AskOnce);
