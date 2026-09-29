# Binary DOC email envelope alignment

Source: [MS-OSHARED](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/d93502fa-5b8f-4f47-a3fe-5574046f4b8d) sections 2.3.8.1–2.3.8.19 (version 11.1, November 2025).
This inventory concerns the `MsoEnvelope` data referenced by the DOC FIB.

The structural walker identifies byte ranges. The editor's
`ReadEmailEnvelope`/`SetEmailEnvelope` methods expose a smaller settings
model. “Read” and “write” below refer to that settings model, not merely
structural navigation.

| Structure or field | Navigation | Read | Write | Current limitation |
| --- | --- | --- | --- | --- |
| Envelope CLSID and version | Yes | Yes | Yes | Settings model accepts version 8 only; version 6 remains opaque. |
| LastSentTime, FlagStatus, ReplyTime | Yes | No | Defaults | Existing values are discarded by complete replacement. |
| RequestStr | Yes | No | Empty | Follow-up request text is not modeled. |
| SentRepresentingEntryId, SentRepresentingName | Yes | No | Empty | Sender identity is not modeled. |
| InetAcctStamp, InetAcctName | Yes | No | Empty | Account integration fields are not modeled. |
| ExpiryTime, DeferredDeliveryTime | Yes | Yes | Yes | Exposed as `ExpiresAt` and `DeliverAfter`. |
| DeleteAfterSubmit, SecurityFlags | Yes | No | Zero | SecurityFlags contains signed and encrypted bits; complete replacement clears them. |
| Delivery and read receipt requests | Yes | Yes | Yes | — |
| CategoriesStr, Sensitivity, Importance, Subject | Yes | Yes | Yes | — |
| VotingOptions | Yes | No | Empty | ANSI bytes are not modeled. |
| ReplyRecipients | Yes | Partial | Partial | Only SMTP address rows fit `DocEmailAddress`; other properties are discarded. |
| ContactLinkRecipients | Yes | No | Empty | Entire collection is discarded by complete replacement. |
| Message Recipients (To/Cc/Bcc) | Yes | Partial | Partial | Role is read from `PR_RECIPIENT_TYPE`; addressless rows cannot fit `DocEmailAddress`. Extra properties are discarded. |
| EnvRecipientPropertyBlob | Yes | Partial | Partial | Walker recognizes every type listed in section 2.3.8.6 and retains its range. Settings reader decodes selected tags and skips other values; writer generates seven SMTP properties per recipient. |
| Attachments | Yes | Partial | Partial | Settings model supports by-value attachments only. |
| IntroText | Yes | Yes | Yes | Version 8 only. |
| Envelope visibility | Yes | Yes | Yes | Stored in the DOC DOP, outside `MsoEnvelope`. |

## Recipient rows

An `EnvRecipientProperties` row contains a property count, an ignored 32-bit
value, and independently tagged properties. The spec does not require
`PR_SMTP_ADDRESS` or `PR_EMAIL_ADDRESS` to appear with
`PR_RECIPIENT_TYPE`. A BCC row can therefore have role 3 and a display name
without any SMTP address. The two supplied test documents confirm this
layout; they also contain `PT_BINARY`, `PT_BOOLEAN`, and `PT_ERROR`
properties.

`DocEmailAddress` is a useful convenience for SMTP recipients, but it is
not a complete recipient-row model. Treating a display name as an email
address would misrepresent the file. The settings reader currently reports
an addressless row as unsupported; the walker can still navigate and inspect
it.

## Recommended editor model

1. Add a spec-shaped recipient row that preserves collection membership,
   property order, each property tag and encoded value, and the ignored
   32-bit row field. Decode common tags on demand, including role, display
   name, address type, email address, and SMTP address.
2. Keep `To`, `Cc`, `Bcc`, and `ReplyTo` as convenience projections for
   resolvable SMTP recipients. Expose every row separately so addressless
   and Exchange-specific recipients remain visible and round-trip.
3. Let focused edits change individual fields or rows while carrying
   untouched encoded properties forward. A complete settings replacement
   should remain explicit because it may drop unsupported data.
4. Expose the simple scalar header fields as getters and setters. Preserve
   raw bytes for sender entry ID and voting options until their semantics
   are needed. Keep version 6 and non-by-value attachments scoped separately.

The immediate compatibility check is to read and save both supplied
documents without losing the addressless BCC row or changing unrelated
recipient properties.
