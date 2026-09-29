using System.Buffers.Binary;
using System.Globalization;
using System.Text;

namespace DocxportNet.Doc;

/// <summary>Indexes Unicode email-envelope fields without copying strings or attachment data.</summary>
internal static class DocEnvelopeNavigator
{
    private const uint RecipientMarker = 0xDCCA0123;
    private static readonly Encoding Unicode = new UnicodeEncoding(false, false, true);

    public static void ExpandLocation(DocStructure structure, DocStructureNode location)
    {
        if (location.StreamName == null || location.Offset == null || location.Length == null ||
            location.Length <= 20 || location.Children.Any(x => x.Kind == "MsoEnvelope")) return;
        var header = structure.ReadRange(location.StreamName, location.Offset.Value, 20);
        var body = new DocStructureNode("MsoEnvelope", "Envelope", location.StreamName,
            location.Offset.Value + 20, location.Length.Value - 20);
        body.Attributes["version"] = BinaryPrimitives.ReadUInt32LittleEndian(header.AsSpan(16))
            .ToString(CultureInfo.InvariantCulture);
        if (new Guid(header.Take(16).ToArray()) != DocEmailEnvelopeCodec.ClassId)
            body.Attributes["navigation"] = "unrecognized-class";
        location.Children.Add(body);
    }

    public static void ExpandBody(DocStructure structure, DocStructureNode body)
    {
        if (body.Children.Count != 0 || body.StreamName == null || body.Offset == null || body.Length == null)
            return;
        var end = checked(body.Offset.Value + body.Length.Value);
        if (body.Attributes.TryGetValue("navigation", out var reason))
        {
            AddOpaque(body, body.Offset.Value, end, reason);
            return;
        }
        if (body.Attributes["version"] != "8")
        {
            AddOpaque(body, body.Offset.Value, end, "unsupported-version");
            return;
        }
        using var stream = structure.OpenStream(body.StreamName);
        var reader = new Cursor(stream, body.Offset.Value, end);
        try
        {
            UInt32(body, reader, "LastSentTime");
            UInt32(body, reader, "FlagStatus");
            UInt32(body, reader, "ReplyTime");
            UnicodeString(structure, body, reader, "RequestStr");
            Binary(body, reader, "SentRepresentingEntryId", reader.U32());
            UnicodeString(structure, body, reader, "SentRepresentingName");
            UnicodeString(structure, body, reader, "InetAcctStamp");
            UnicodeString(structure, body, reader, "InetAcctName");
            UInt32(body, reader, "ExpiresAtMinutes");
            UInt32(body, reader, "DeliverAfterMinutes");
            UInt32(body, reader, "DeleteAfterSubmit");
            UInt32(body, reader, "SecurityFlags");
            UInt32(body, reader, "RequestDeliveryReceipt");
            UInt32(body, reader, "RequestReadReceipt");
            UnicodeString(structure, body, reader, "Categories");
            UInt32(body, reader, "Sensitivity");
            UInt32(body, reader, "Importance");
            UnicodeString(structure, body, reader, "Subject");
            Binary(body, reader, "VotingOptions", reader.U16(), 2);
            body.Children.Add(Recipients(structure, reader, "ReplyRecipients"));
            body.Children.Add(Recipients(structure, reader, "ContactLinkRecipients"));
            body.Children.Add(Recipients(structure, reader, "MessageRecipients"));
            body.Children.Add(Attachments(structure, reader));
            var introStart = reader.Position;
            var introLength = reader.U32();
            if ((introLength & 1) != 0)
                throw new InvalidDataException("The envelope introduction has an odd UTF-16 byte length.");
            var introOffset = reader.Position;
            reader.Skip(introLength);
            var intro = new DocStructureNode("EnvUnicodeString", "Introduction", body.StreamName,
                introStart, reader.Position - introStart);
            TextPayload(structure, intro, introOffset, checked((int)introLength));
            body.Children.Add(intro);
            if (reader.Position < end) AddOpaque(body, reader.Position, end, "unrecognized-tail");
        }
        catch (UnsupportedPropertyException ex)
        {
            AddOpaque(body, ex.Offset, end, $"recipient-property-type-0x{ex.Type:X4}");
        }
    }

    private static void UInt32(DocStructureNode parent, Cursor reader, string name)
    {
        var offset = reader.Position;
        var value = reader.U32();
        var node = new DocStructureNode("EnvUInt32", name, parent.StreamName, offset, 4);
        node.Attributes["value"] = value.ToString(CultureInfo.InvariantCulture);
        parent.Children.Add(node);
    }

    private static void UnicodeString(DocStructure structure, DocStructureNode parent, Cursor reader, string name)
    {
        var start = reader.Position;
        var characters = reader.U16();
        var dataOffset = reader.Position;
        var byteLength = characters * 2;
        reader.Skip(byteLength);
        var node = new DocStructureNode("EnvUnicodeString", name, parent.StreamName,
            start, reader.Position - start);
        node.Attributes["characters"] = characters.ToString(CultureInfo.InvariantCulture);
        TextPayload(structure, node, dataOffset, byteLength);
        parent.Children.Add(node);
    }

    private static void Binary(DocStructureNode parent, Cursor reader, string name, long length, int prefixBytes = 4)
    {
        var dataOffset = reader.Position;
        reader.Skip(length);
        var node = new DocStructureNode("EnvBinary", name, parent.StreamName,
            dataOffset - prefixBytes, length + prefixBytes);
        node.Attributes["dataOffset"] = dataOffset.ToString(CultureInfo.InvariantCulture);
        node.Attributes["dataLength"] = length.ToString(CultureInfo.InvariantCulture);
        parent.Children.Add(node);
    }

    private static DocStructureNode Recipients(DocStructure structure, Cursor reader, string name)
    {
        var start = reader.Position;
        if (reader.U32() != RecipientMarker || reader.U32() != 1)
            throw new InvalidDataException("The envelope recipient collection header is invalid.");
        var count = reader.U32();
        if (count > (reader.End - reader.Position) / 8)
            throw new InvalidDataException("The envelope recipient count is invalid.");
        var entries = new List<DocStructureNode>();
        for (uint i = 0; i < count; i++)
        {
            var recipientStart = reader.Position;
            var propertyCount = reader.U32();
            var ignored = reader.U32();
            if (propertyCount > (reader.End - reader.Position) / 8)
                throw new InvalidDataException("The envelope recipient property count is invalid.");
            var properties = new List<DocStructureNode>();
            for (uint j = 0; j < propertyCount; j++)
            {
                var propertyStart = reader.Position;
                var tag = reader.U32();
                var type = (ushort)(tag & 0xFFFF);
                DocStructureNode property;
                if (type == 3)
                {
                    var value = reader.U32();
                    property = new DocStructureNode("EnvRecipientProperty", $"Property{j}",
                        structure.FibBase.TableStreamName, propertyStart, reader.Position - propertyStart);
                    property.Attributes["value"] = value.ToString(CultureInfo.InvariantCulture);
                }
                else if (type == 31)
                {
                    var byteLength = reader.U16();
                    if ((byteLength & 1) != 0)
                        throw new InvalidDataException("A recipient Unicode property has an odd byte length.");
                    var dataOffset = reader.Position;
                    reader.Skip(byteLength);
                    property = new DocStructureNode("EnvRecipientProperty", $"Property{j}",
                        structure.FibBase.TableStreamName, propertyStart, reader.Position - propertyStart);
                    var text = new DocStructureNode("EnvUnicodeString", "Value", property.StreamName,
                        dataOffset, byteLength);
                    TextPayload(structure, text, dataOffset, byteLength);
                    property.Children.Add(text);
                }
                else throw new UnsupportedPropertyException(start, type);
                property.Attributes["tag"] = $"0x{tag:X8}";
                property.Attributes["type"] = $"0x{type:X4}";
                properties.Add(property);
            }
            var recipient = new DocStructureNode("EnvRecipientProperties", $"Recipient{i}",
                structure.FibBase.TableStreamName, recipientStart, reader.Position - recipientStart);
            recipient.Attributes["propertyCount"] = propertyCount.ToString(CultureInfo.InvariantCulture);
            recipient.Attributes["ignored"] = ignored.ToString(CultureInfo.InvariantCulture);
            foreach (var property in properties) recipient.Children.Add(property);
            entries.Add(recipient);
        }
        var collection = new DocStructureNode("EnvRecipientCollection", name,
            structure.FibBase.TableStreamName, start, reader.Position - start);
        collection.Attributes["recipientCount"] = count.ToString(CultureInfo.InvariantCulture);
        foreach (var recipient in entries) collection.Children.Add(recipient);
        return collection;
    }

    private static DocStructureNode Attachments(DocStructure structure, Cursor reader)
    {
        var start = reader.Position;
        var count = reader.U32();
        if (count > (reader.End - reader.Position) / 13)
            throw new InvalidDataException("The envelope attachment count is invalid.");
        var entries = new List<DocStructureNode>();
        for (uint i = 0; i < count; i++)
        {
            var attachmentStart = reader.Position;
            var method = reader.U32();
            var nameCharacters = reader.U8();
            var nameOffset = reader.Position;
            var nameLength = nameCharacters * 2;
            reader.Skip(nameLength);
            var size = reader.U64();
            if (size > long.MaxValue)
                throw new InvalidDataException("An envelope attachment is too large.");
            var dataOffset = reader.Position;
            reader.Skip((long)size);
            var attachment = new DocStructureNode("EnvAttachment", $"Attachment{i}",
                structure.FibBase.TableStreamName, attachmentStart, reader.Position - attachmentStart);
            attachment.Attributes["method"] = method.ToString(CultureInfo.InvariantCulture);
            attachment.Attributes["dataLength"] = size.ToString(CultureInfo.InvariantCulture);
            var name = new DocStructureNode("EnvUnicodeString", "FileName", attachment.StreamName,
                nameOffset, nameLength);
            TextPayload(structure, name, nameOffset, nameLength);
            attachment.Children.Add(name);
            var data = new DocStructureNode("EnvAttachmentData", "Data", attachment.StreamName,
                dataOffset, (long)size);
            if (size <= int.MaxValue)
                data.SetPayloadFactory(() => new DocEnvelopeBytes(structure.ReadRange(data.StreamName!, dataOffset, (int)size)));
            attachment.Children.Add(data);
            entries.Add(attachment);
        }
        var collection = new DocStructureNode("EnvAttachmentCollection", "Attachments",
            structure.FibBase.TableStreamName, start, reader.Position - start);
        collection.Attributes["attachmentCount"] = count.ToString(CultureInfo.InvariantCulture);
        foreach (var attachment in entries) collection.Children.Add(attachment);
        return collection;
    }

    private static void TextPayload(DocStructure structure, DocStructureNode node, long offset, int byteLength)
        => node.SetPayloadFactory(() =>
        {
            var bytes = structure.ReadRange(node.StreamName!, offset, byteLength);
            try { return new DocEnvelopeText(Unicode.GetString(bytes)); }
            catch (DecoderFallbackException ex)
            {
                throw new InvalidDataException("The envelope contains invalid UTF-16 text.", ex);
            }
        });

    private static void AddOpaque(DocStructureNode body, long start, long end, string reason)
    {
        if (start >= end) return;
        var node = new DocStructureNode("EnvOpaqueBytes", "UnparsedEnvelopeData", body.StreamName,
            start, end - start);
        node.Attributes["reason"] = reason;
        body.Children.Add(node);
    }

    private sealed class UnsupportedPropertyException(long offset, ushort type) : Exception
    {
        public long Offset { get; } = offset;
        public ushort Type { get; } = type;
    }

    private sealed class Cursor(Stream stream, long start, long end)
    {
        public long End { get; } = end;
        public long Position { get; private set; } = start;

        public void Skip(long length)
        {
            Require(length);
            Position += length;
        }

        public byte U8() => Read(1)[0];
        public ushort U16() => BinaryPrimitives.ReadUInt16LittleEndian(Read(2));
        public uint U32() => BinaryPrimitives.ReadUInt32LittleEndian(Read(4));
        public ulong U64() => BinaryPrimitives.ReadUInt64LittleEndian(Read(8));

        private byte[] Read(int length)
        {
            Require(length);
            stream.Position = Position;
            var bytes = new byte[length];
            var read = 0;
            while (read < length)
            {
                var count = stream.Read(bytes, read, length - read);
                if (count == 0) throw new InvalidDataException("The email envelope ended unexpectedly.");
                read += count;
            }
            Position += length;
            return bytes;
        }

        private void Require(long length)
        {
            if (length < 0 || Position < 0 || Position > End || length > End - Position)
                throw new InvalidDataException("An email envelope field exceeds its indexed byte range.");
        }
    }
}
