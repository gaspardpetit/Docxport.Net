namespace DocxportNet.Doc;

/// <summary>A complete replacement for a binary Word document's email envelope.</summary>
public sealed record DocEmailEnvelope
{
    public string Subject { get; init; } = "";
    public string Introduction { get; init; } = "";
    public IReadOnlyList<DocEmailAddress> To { get; init; } = Array.Empty<DocEmailAddress>();
    public IReadOnlyList<DocEmailAddress> Cc { get; init; } = Array.Empty<DocEmailAddress>();
    public IReadOnlyList<DocEmailAddress> Bcc { get; init; } = Array.Empty<DocEmailAddress>();
    public IReadOnlyList<DocEmailAddress> ReplyTo { get; init; } = Array.Empty<DocEmailAddress>();
    public IReadOnlyList<DocEmailAttachment> Attachments { get; init; } = Array.Empty<DocEmailAttachment>();
    public DocEmailImportance Importance { get; init; } = DocEmailImportance.Normal;
    public DocEmailSensitivity Sensitivity { get; init; } = DocEmailSensitivity.Normal;
    public bool RequestDeliveryReceipt { get; init; }
    public bool RequestReadReceipt { get; init; }
    public bool Visible { get; init; } = true;
    public string Categories { get; init; } = "";
    public DateTimeOffset? DeliverAfter { get; init; }
    public DateTimeOffset? ExpiresAt { get; init; }
}

/// <summary>One SMTP mailbox and optional display name.</summary>
public sealed record DocEmailAddress(string Address, string? DisplayName = null);

/// <summary>An attachment stored by value in the envelope.</summary>
public sealed record DocEmailAttachment(string FileName, byte[] Content);

public enum DocEmailImportance { Low = 0, Normal = 1, High = 2 }
public enum DocEmailSensitivity { Normal = 0, Personal = 1, Private = 2, Confidential = 3 }
