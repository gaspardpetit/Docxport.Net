using System.Net.Mail;
using System.Text;

namespace DocxportNet.Doc;

/// <summary>Reads and writes the supported Unicode MS-OSHARED email envelope fields.</summary>
internal static class DocEmailEnvelopeCodec
{
    internal static readonly Guid ClassId = new Guid("0006F01A-0000-0000-C000-000000000046");
    private const uint UnspecifiedTime = 0x5AE980E0;
    private const uint RecipientMarker = 0xDCCA0123;
    private static readonly Encoding Unicode = new UnicodeEncoding(false, false, true);
    private static readonly DateTime Epoch = new DateTime(1601, 1, 1, 0, 0, 0, DateTimeKind.Utc);

    public static byte[] Serialize(DocEmailEnvelope envelope)
    {
        if (envelope == null) throw new ArgumentNullException(nameof(envelope));
        if (!Enum.IsDefined(typeof(DocEmailImportance), envelope.Importance) ||
            !Enum.IsDefined(typeof(DocEmailSensitivity), envelope.Sensitivity))
            throw new ArgumentException("The envelope importance or sensitivity is invalid.", nameof(envelope));
        using var output = new MemoryStream();
        using var writer = new BinaryWriter(output, Unicode, true);
        writer.Write(ClassId.ToByteArray());
        writer.Write(8u); // Unicode MsoEnvelope
        writer.Write(UnspecifiedTime); // LastSentTime
        writer.Write(0u); // FlagStatus
        writer.Write(UnspecifiedTime); // ReplyTime
        WriteString(writer, ""); // RequestStr
        writer.Write(0u); // SentRepresentingEntryIdSize
        WriteString(writer, ""); // SentRepresentingName
        WriteString(writer, ""); // InetAcctStamp
        WriteString(writer, ""); // InetAcctName
        writer.Write(EncodeTime(envelope.ExpiresAt));
        writer.Write(EncodeTime(envelope.DeliverAfter));
        writer.Write(0u); // DeleteAfterSubmit
        writer.Write(0u); // SecurityFlags
        writer.Write(envelope.RequestDeliveryReceipt ? 1u : 0u);
        writer.Write(envelope.RequestReadReceipt ? 1u : 0u);
        WriteString(writer, envelope.Categories);
        writer.Write((uint)envelope.Sensitivity);
        writer.Write((uint)envelope.Importance);
        WriteString(writer, envelope.Subject);
        writer.Write((ushort)0); // VotingOptionsSize
        WriteRecipients(writer, (envelope.ReplyTo ?? throw new ArgumentException("ReplyTo is null.")).Select(x => (x, 1u)));
        WriteRecipients(writer, Enumerable.Empty<(DocEmailAddress, uint)>()); // ContactLinkRecipients
        WriteRecipients(writer,
            (envelope.To ?? throw new ArgumentException("To is null.")).Select(x => (x, 1u))
                .Concat((envelope.Cc ?? throw new ArgumentException("Cc is null.")).Select(x => (x, 2u)))
                .Concat((envelope.Bcc ?? throw new ArgumentException("Bcc is null.")).Select(x => (x, 3u))));
        var attachments = envelope.Attachments ?? throw new ArgumentException("Attachments is null.");
        writer.Write(checked((uint)attachments.Count));
        foreach (var attachment in attachments)
        {
            if (attachment == null || attachment.Content == null)
                throw new ArgumentException("An attachment or its content is null.", nameof(envelope));
            CheckText(attachment.FileName);
            if (string.IsNullOrWhiteSpace(attachment.FileName) || attachment.FileName.Length > 255 ||
                attachment.FileName.IndexOfAny(new[] { '/', '\\', ':' }) >= 0 ||
                attachment.FileName == "." || attachment.FileName == "..")
                throw new ArgumentException("An attachment needs a filename of at most 255 UTF-16 units without a path.", nameof(envelope));
            writer.Write(1u); // ATTACH_BY_VALUE
            writer.Write(checked((byte)attachment.FileName.Length));
            writer.Write(Unicode.GetBytes(attachment.FileName));
            writer.Write(checked((ulong)attachment.Content.LongLength));
            writer.Write(attachment.Content);
        }
        var introduction = EncodeText(envelope.Introduction);
        writer.Write(checked((uint)introduction.Length));
        writer.Write(introduction);
        return output.ToArray();
    }

    public static DocEmailEnvelope Deserialize(byte[] bytes, bool visible)
    {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        try
        {
            using var input = new MemoryStream(bytes, false);
            using var reader = new BinaryReader(input, Unicode);
            if (new Guid(ReadBytes(reader, 16)) != ClassId)
                throw new InvalidDataException("The envelope class identifier is invalid.");
            if (reader.ReadUInt32() != 8)
                throw new NotSupportedException("Only Unicode version 8 envelopes can be read as settings.");
            reader.ReadUInt32(); // LastSentTime
            reader.ReadUInt32(); // FlagStatus
            reader.ReadUInt32(); // ReplyTime
            ReadString(reader); // RequestStr
            var representingIdLength = reader.ReadUInt32();
            ReadBytes(reader, CheckedLength(representingIdLength));
            ReadString(reader); // SentRepresentingName
            ReadString(reader); // InetAcctStamp
            ReadString(reader); // InetAcctName
            var expires = DecodeTime(reader.ReadUInt32());
            var deliver = DecodeTime(reader.ReadUInt32());
            reader.ReadUInt32(); // DeleteAfterSubmit
            reader.ReadUInt32(); // SecurityFlags
            var deliveryReceipt = reader.ReadUInt32() != 0;
            var readReceipt = reader.ReadUInt32() != 0;
            var categories = ReadString(reader);
            var sensitivity = (DocEmailSensitivity)reader.ReadUInt32();
            var importance = (DocEmailImportance)reader.ReadUInt32();
            var subject = ReadString(reader);
            var votingBytes = reader.ReadUInt16();
            ReadBytes(reader, votingBytes);
            var reply = ReadRecipients(reader);
            ReadRecipients(reader); // Contact links are outside the editable model.
            var recipients = ReadRecipients(reader);
            var attachmentCount = reader.ReadUInt32();
            if (attachmentCount > input.Length - input.Position)
                throw new InvalidDataException("The attachment count exceeds the envelope length.");
            var attachments = new List<DocEmailAttachment>();
            for (uint i = 0; i < attachmentCount; i++)
            {
                if (reader.ReadUInt32() != 1)
                    throw new NotSupportedException("Only by-value attachments can be read as settings.");
                var name = Decode(ReadBytes(reader, reader.ReadByte() * 2));
                var size = reader.ReadUInt64();
                if (size > int.MaxValue || size > (ulong)(input.Length - input.Position))
                    throw new InvalidDataException("An attachment extends beyond the envelope.");
                attachments.Add(new DocEmailAttachment(name, ReadBytes(reader, (int)size)));
            }
            var introBytes = reader.ReadUInt32();
            var introduction = Decode(ReadBytes(reader, CheckedLength(introBytes)));
            if (input.Position != input.Length)
                throw new NotSupportedException("The envelope has additional data outside the supported settings model.");
            return new DocEmailEnvelope
            {
                Subject = subject, Introduction = introduction,
                To = recipients.Where(x => x.Type == 1).Select(x => x.Address).ToArray(),
                Cc = recipients.Where(x => x.Type == 2).Select(x => x.Address).ToArray(),
                Bcc = recipients.Where(x => x.Type == 3).Select(x => x.Address).ToArray(),
                ReplyTo = reply.Select(x => x.Address).ToArray(), Attachments = attachments,
                Importance = importance, Sensitivity = sensitivity,
                RequestDeliveryReceipt = deliveryReceipt, RequestReadReceipt = readReceipt,
                Categories = categories, ExpiresAt = expires, DeliverAfter = deliver, Visible = visible
            };
        }
        catch (EndOfStreamException ex)
        {
            throw new InvalidDataException("The email envelope is truncated.", ex);
        }
    }

    private static void WriteRecipients(BinaryWriter writer, IEnumerable<(DocEmailAddress Address, uint Type)> entries)
    {
        var recipients = entries.ToArray();
        writer.Write(RecipientMarker);
        writer.Write(1u);
        writer.Write(checked((uint)recipients.Length));
        foreach (var (recipient, type) in recipients)
        {
            if (recipient == null) throw new ArgumentException("A recipient is null.");
            CheckText(recipient.Address);
            MailAddress parsed;
            try { parsed = new MailAddress(recipient.Address); }
            catch (FormatException ex) { throw new ArgumentException("A recipient needs one SMTP address.", ex); }
            if (parsed.Address != recipient.Address || !string.IsNullOrEmpty(parsed.DisplayName))
                throw new ArgumentException("A recipient needs one SMTP address without an embedded display name.");
            writer.Write(7u);
            writer.Write(0u);
            WriteLongProperty(writer, 0x0C150003, type); // PR_RECIPIENT_TYPE
            WriteUnicodeProperty(writer, 0x3001001F, recipient.DisplayName ?? recipient.Address);
            WriteUnicodeProperty(writer, 0x3002001F, "SMTP");
            WriteUnicodeProperty(writer, 0x3003001F, recipient.Address);
            WriteUnicodeProperty(writer, 0x39FE001F, recipient.Address);
            WriteLongProperty(writer, 0x0FFE0003, 6); // MAPI_MAILUSER
            WriteLongProperty(writer, 0x39000003, 0); // DT_MAILUSER
        }
    }

    private static List<(DocEmailAddress Address, uint Type)> ReadRecipients(BinaryReader reader)
    {
        if (reader.ReadUInt32() != RecipientMarker || reader.ReadUInt32() != 1)
            throw new InvalidDataException("The recipient collection header is invalid.");
        var count = reader.ReadUInt32();
        if (count > reader.BaseStream.Length - reader.BaseStream.Position)
            throw new InvalidDataException("The recipient count exceeds the envelope length.");
        var result = new List<(DocEmailAddress, uint)>();
        for (uint i = 0; i < count; i++)
        {
            var properties = reader.ReadUInt32();
            reader.ReadUInt32(); // Ignored
            if (properties > reader.BaseStream.Length - reader.BaseStream.Position)
                throw new InvalidDataException("The recipient property count is invalid.");
            uint type = 1;
            string? smtpAddress = null, emailAddress = null, name = null;
            for (uint j = 0; j < properties; j++)
            {
                var tag = reader.ReadUInt32();
                var propertyType = (ushort)(tag & 0xFFFF);
                if (propertyType is 1 or 3 or 10)
                {
                    var value = reader.ReadUInt32();
                    if (tag == 0x0C150003) type = value;
                }
                else if (propertyType == 11)
                {
                    reader.ReadUInt16();
                }
                else if (propertyType == 64)
                {
                    SkipBytes(reader, 8);
                }
                else if (propertyType is 30 or 31 or 258)
                {
                    var size = reader.ReadUInt16();
                    if (propertyType == 31 && tag is 0x3001001F or 0x3003001F or 0x39FE001F)
                    {
                        if ((size & 1) != 0) throw new InvalidDataException("A recipient Unicode property has an odd byte length.");
                        var value = Decode(ReadBytes(reader, size)).TrimEnd('\0');
                        if (tag == 0x3001001F) name = value;
                        if (tag == 0x3003001F) emailAddress = value;
                        if (tag == 0x39FE001F) smtpAddress = value;
                    }
                    else SkipBytes(reader, size);
                }
                else if (propertyType is 4126 or 4354)
                {
                    var elementCount = reader.ReadUInt32();
                    if (elementCount > (reader.BaseStream.Length - reader.BaseStream.Position) / 2)
                        throw new InvalidDataException("A recipient multi-value property count is invalid.");
                    for (uint k = 0; k < elementCount; k++) SkipBytes(reader, reader.ReadUInt16());
                }
                else throw new NotSupportedException("An envelope recipient property type is unsupported.");
            }
            var address = smtpAddress ?? emailAddress;
            if (address == null) throw new NotSupportedException(
                $"A recipient of type {type} has no SMTP or email address that the settings model can represent.");
            result.Add((new DocEmailAddress(address, name == address ? null : name), type));
        }
        return result;
    }

    private static void WriteLongProperty(BinaryWriter writer, uint tag, uint value)
    {
        writer.Write(tag);
        writer.Write(value);
    }

    private static void WriteUnicodeProperty(BinaryWriter writer, uint tag, string value)
    {
        CheckText(value);
        var bytes = Unicode.GetBytes(value + '\0');
        if (bytes.Length > ushort.MaxValue)
            throw new ArgumentException("A recipient property exceeds 65,535 bytes.");
        writer.Write(tag);
        writer.Write(checked((ushort)bytes.Length));
        writer.Write(bytes);
    }

    private static void WriteString(BinaryWriter writer, string value)
    {
        var bytes = EncodeText(value);
        if (value.Length > ushort.MaxValue)
            throw new ArgumentException("An envelope string exceeds 65,535 UTF-16 units.");
        writer.Write(checked((ushort)value.Length));
        writer.Write(bytes);
    }

    private static string ReadString(BinaryReader reader) => Decode(ReadBytes(reader, reader.ReadUInt16() * 2));

    private static byte[] ReadBytes(BinaryReader reader, int length)
    {
        if (length < 0 || length > reader.BaseStream.Length - reader.BaseStream.Position)
            throw new InvalidDataException("An envelope field exceeds its byte range.");
        var bytes = reader.ReadBytes(length);
        if (bytes.Length != length) throw new EndOfStreamException();
        return bytes;
    }

    private static void SkipBytes(BinaryReader reader, int length)
    {
        if (length < 0 || length > reader.BaseStream.Length - reader.BaseStream.Position)
            throw new InvalidDataException("An envelope field exceeds its byte range.");
        reader.BaseStream.Position += length;
    }

    private static int CheckedLength(uint length)
    {
        if (length > int.MaxValue)
            throw new InvalidDataException("An envelope field exceeds its byte range.");
        return (int)length;
    }

    private static string Decode(byte[] bytes)
    {
        try { return Unicode.GetString(bytes); }
        catch (DecoderFallbackException ex) { throw new InvalidDataException("Invalid UTF-16 in the envelope.", ex); }
    }

    private static byte[] EncodeText(string value)
    {
        CheckText(value);
        try { return Unicode.GetBytes(value); }
        catch (EncoderFallbackException ex) { throw new ArgumentException("Invalid UTF-16 in the envelope.", ex); }
    }

    private static void CheckText(string value)
    {
        if (value == null) throw new ArgumentNullException(nameof(value));
        if (value.IndexOf('\0') >= 0) throw new ArgumentException("Envelope text cannot contain NUL characters.");
    }

    private static uint EncodeTime(DateTimeOffset? value)
    {
        if (value == null) return UnspecifiedTime;
        var ticks = value.Value.UtcDateTime.Ticks - Epoch.Ticks;
        var minutes = ticks / TimeSpan.TicksPerMinute;
        if (ticks < 0 || minutes >= UnspecifiedTime)
            throw new ArgumentOutOfRangeException(nameof(value), "The envelope date is outside the supported range.");
        return checked((uint)minutes);
    }

    private static DateTimeOffset? DecodeTime(uint minutes) => minutes == UnspecifiedTime
        ? (DateTimeOffset?)null : new DateTimeOffset(Epoch.AddMinutes(minutes));
}
