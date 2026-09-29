using System.Buffers.Binary;
using System.Text;
using DocxportNet.Doc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenMcdf;
using System.Xml.Linq;

namespace DocxportNet.Tests;

public class DocEditorTests
{
    [Fact]
    public void WalkerVisitsEnvelopeFieldsAndKeepsAttachmentBytesLazy()
    {
        var payload = Enumerable.Range(0, 12000).Select(i => (byte)i).ToArray();
        using var editor = DocEditor.Open(CreatePlainDoc());
        var edited = editor.SetEmailEnvelope(new DocEmailEnvelope
        {
            Subject = "Résultats Ω",
            To = new[] { new DocEmailAddress("client@example.com", "Client") },
            Attachments = new[] { new DocEmailAttachment("report.bin", payload) }
        }).Save();

        using var input = new MemoryStream(edited);
        using var output = new StringWriter();
        using var xml = new DocStructureXmlVisitor(output);
        using var structure = new DocStructureWalker().Accept(input, xml);
        xml.Dispose();
        var dump = XDocument.Parse(output.ToString());
        Assert.Contains(dump.Descendants("MsoEnvelope"), x => (string?)x.Attribute("name") == "Envelope");
        Assert.Contains(dump.Descendants("EnvUnicodeString"), x =>
            (string?)x.Attribute("name") == "Subject" && (string?)x.Element("Text") == "R\\u00E9sultats \\u03A9");
        Assert.Single(dump.Descendants("EnvRecipientProperties"));
        Assert.Single(dump.Descendants("EnvAttachment"));
        var data = Find(structure.Root, "EnvAttachmentData");
        Assert.Equal(payload.Length, data.Length);
        Assert.True(data.HasPayload);
        Assert.False(data.IsPayloadLoaded);
        Assert.Equal(payload, ((DocEnvelopeBytes)data.Payload!).Bytes);
        Assert.True(data.IsPayloadLoaded);
        Assert.Same(data.Payload, data.Payload);
    }

    [Fact]
    public void WalkerCanSkipEnvelopeBodyBeforeParsingItsContents()
    {
        using var editor = DocEditor.Open(CreatePlainDoc());
        var edited = editor.SetEmailEnvelope(new DocEmailEnvelope { Subject = "Skip me" }).Save();
        using var input = new MemoryStream(edited);
        var visitor = new SkipEnvelopeBodyVisitor();
        using var structure = new DocStructureWalker().Accept(input, visitor);
        Assert.Contains("MsoEnvelope", visitor.Kinds);
        Assert.DoesNotContain("EnvUnicodeString", visitor.Kinds);
        Assert.Empty(Find(structure.Root, "MsoEnvelope").Children);
    }

    [Fact]
    public void WalkerSkipsMalformedBodyAndKeepsUnknownRecipientPropertiesOpaque()
    {
        using var editor = DocEditor.Open(CreatePlainDoc());
        var edited = editor.SetEmailEnvelope(new DocEmailEnvelope
        {
            To = new[] { new DocEmailAddress("client@example.com") }
        }).Save();
        long collectionOffset;
        long propertyOffset;
        using (var input = new MemoryStream(edited))
        using (var structure = new DocStructureWalker().Accept(input,
                   new DocStructurePrintVisitor(new StringWriter())))
        {
            collectionOffset = FindNamed(structure.Root, "EnvRecipientCollection", "MessageRecipients").Offset!.Value;
            propertyOffset = Find(structure.Root, "EnvRecipientProperty").Offset!.Value;
        }
        var table = ReadStream(edited, "1Table");
        table[(int)collectionOffset] = 0; // Invalid collection marker.
        var malformed = RewriteStream(edited, "1Table", table);
        using (var input = new MemoryStream(malformed))
        using (new DocStructureWalker().Accept(input, new SkipEnvelopeBodyVisitor())) { }
        using (var input = new MemoryStream(malformed))
            Assert.Throws<InvalidDataException>(() => new DocStructureWalker().Accept(input,
                new DocStructurePrintVisitor(new StringWriter())));

        table = ReadStream(edited, "1Table");
        table[(int)propertyOffset] = 0x99;
        table[(int)propertyOffset + 1] = 0x99; // Unknown recipient property type.
        var unknown = RewriteStream(edited, "1Table", table);
        using var unknownInput = new MemoryStream(unknown);
        using var unknownStructure = new DocStructureWalker().Accept(unknownInput,
            new DocStructurePrintVisitor(new StringWriter()));
        var opaque = Find(unknownStructure.Root, "EnvOpaqueBytes");
        Assert.Equal(collectionOffset, opaque.Offset);
        Assert.Contains("0x9999", opaque.Attributes["reason"]);
    }

    private sealed class SkipEnvelopeBodyVisitor : IDocStructureVisitor
    {
        public List<string> Kinds { get; } = new();
        public IDisposable? Enter(DocStructureNode node, int depth)
        {
            Kinds.Add(node.Kind);
            return node.Kind == "MsoEnvelope" ? null : DocxportNet.Core.DxpDisposable.Empty;
        }
    }

    private static DocStructureNode Find(DocStructureNode node, string kind)
    {
        if (node.Kind == kind) return node;
        foreach (var child in node.Children)
        {
            var found = FindOrNull(child, kind);
            if (found != null) return found;
        }
        throw new InvalidOperationException($"No {kind} node was found.");
    }

    private static DocStructureNode FindNamed(DocStructureNode node, string kind, string name)
    {
        if (node.Kind == kind && node.Name == name) return node;
        foreach (var child in node.Children)
        {
            var found = FindNamedOrNull(child, kind, name);
            if (found != null) return found;
        }
        throw new InvalidOperationException($"No {kind} node named {name} was found.");
    }

    private static DocStructureNode? FindNamedOrNull(DocStructureNode node, string kind, string name)
    {
        if (node.Kind == kind && node.Name == name) return node;
        foreach (var child in node.Children)
        {
            var found = FindNamedOrNull(child, kind, name);
            if (found != null) return found;
        }
        return null;
    }

    private static DocStructureNode? FindOrNull(DocStructureNode node, string kind)
    {
        if (node.Kind == kind) return node;
        foreach (var child in node.Children)
        {
            var found = FindOrNull(child, kind);
            if (found != null) return found;
        }
        return null;
    }

    [Fact]
    public void EditsEnvelopeOfTemplateFreeDocAndReadsSettingsBack()
    {
        var original = CreatePlainDoc();
        var snapshot = (byte[])original.Clone();
        var attachment = Enumerable.Range(0, 12000).Select(i => (byte)i).ToArray();
        var replacement = new DocEmailEnvelope
        {
            Subject = "Résultats — 東京",
            Introduction = "Bonjour\r\nVeuillez consulter la pièce jointe.",
            To = new[] { new DocEmailAddress("to@example.com", "Côté, Francine") },
            Cc = new[] { new DocEmailAddress("cc@example.com") },
            Bcc = new[] { new DocEmailAddress("bcc@example.com") },
            ReplyTo = new[] { new DocEmailAddress("reply@example.com") },
            Importance = DocEmailImportance.High,
            Sensitivity = DocEmailSensitivity.Confidential,
            RequestDeliveryReceipt = true,
            RequestReadReceipt = true,
            Categories = "Patents",
            DeliverAfter = new DateTimeOffset(2027, 1, 2, 15, 30, 0, TimeSpan.Zero),
            ExpiresAt = new DateTimeOffset(2027, 2, 2, 15, 30, 0, TimeSpan.Zero),
            Attachments = new[] { new DocEmailAttachment("résultats.bin", attachment) }
        };

        byte[] result;
        using (var editor = DocEditor.Open(original))
        {
            Assert.Equal(original, editor.Save());
            Assert.Null(editor.ReadEmailEnvelope());
            Assert.False(editor.Index.FindLocation("EmailEnvelope")!.IsPresent);
            Assert.False(editor.Index.FindLocation("DocumentProperties")!.IsPresent);
            editor.SetEmailEnvelope(replacement);
            Assert.Equal(replacement.Subject, editor.ReadEmailEnvelope()!.Subject);
            Assert.Equal(snapshot, original);
            result = editor.Save();
            Assert.Equal(result, editor.Save());
        }

        using var reopened = DocEditor.Open(result);
        var actual = reopened.ReadEmailEnvelope()!;
        Assert.Equal(replacement.Subject, actual.Subject);
        Assert.Equal(replacement.Introduction, actual.Introduction);
        Assert.Equal(replacement.To, actual.To);
        Assert.Equal(replacement.Cc, actual.Cc);
        Assert.Equal(replacement.Bcc, actual.Bcc);
        Assert.Equal(replacement.ReplyTo, actual.ReplyTo);
        Assert.Equal(replacement.Importance, actual.Importance);
        Assert.Equal(replacement.Sensitivity, actual.Sensitivity);
        Assert.Equal(replacement.Categories, actual.Categories);
        Assert.Equal(replacement.DeliverAfter, actual.DeliverAfter);
        Assert.Equal(replacement.ExpiresAt, actual.ExpiresAt);
        Assert.True(actual.RequestDeliveryReceipt);
        Assert.True(actual.RequestReadReceipt);
        Assert.True(actual.Visible);
        Assert.Equal(attachment, Assert.Single(actual.Attachments).Content);
        Assert.Equal("Hello Ω\r", ReadText(result));
        Assert.True(reopened.Index.FindLocation("DocumentProperties")!.IsPresent);
    }

    [Fact]
    public void VisibilityReplacementAndRemovalPreserveUnrelatedStream()
    {
        var original = CreateContainer(false);
        using var editor = DocEditor.Open(original);
        var first = editor.SetEmailEnvelope(new DocEmailEnvelope { Subject = "Old secret" }).Save();
        using var showEditor = DocEditor.Open(first);
        var hidden = showEditor.SetEmailEnvelopeVisibility(false).Save();
        Assert.Equal(EnvelopeBytes(first), EnvelopeBytes(hidden));
        using var hiddenEditor = DocEditor.Open(hidden);
        Assert.False(hiddenEditor.ReadEmailEnvelope()!.Visible);
        var replaced = hiddenEditor.SetEmailEnvelope(new DocEmailEnvelope { Subject = "New", Visible = false }).Save();
        using var replacedEditor = DocEditor.Open(replaced);
        Assert.Equal("New", replacedEditor.ReadEmailEnvelope()!.Subject);
        Assert.False(replacedEditor.ReadEmailEnvelope()!.Visible);
        Assert.DoesNotContain("Old secret", Encoding.Unicode.GetString(ReadStream(replaced, "0Table")));
        Assert.Equal(ReadStream(original, "Unrelated"), ReadStream(replaced, "Unrelated"));
        var removed = replacedEditor.RemoveEmailEnvelope().Save();
        using var removedEditor = DocEditor.Open(removed);
        Assert.Null(removedEditor.ReadEmailEnvelope());
        Assert.False(removedEditor.Index.FindLocation("EmailEnvelope")!.IsPresent);
        Assert.DoesNotContain("New", Encoding.Unicode.GetString(ReadStream(removed, "0Table")));
        Assert.Equal(ReadStream(original, "Unrelated"), ReadStream(removed, "Unrelated"));
    }

    [Fact]
    public void RejectsOverlappingOrUnrecognizedOldEnvelopeBeforeRemoving()
    {
        using var editor = DocEditor.Open(CreateContainer(true));
        var original = editor.SetEmailEnvelope(new DocEmailEnvelope { Subject = "Existing" }).Save();
        var word = ReadStream(original, "WordDocument");
        var envelope = LocateEnvelope(word);
        BinaryPrimitives.WriteUInt32LittleEndian(word.AsSpan(154 + 33 * 8), envelope.Offset);
        BinaryPrimitives.WriteUInt32LittleEndian(word.AsSpan(154 + 33 * 8 + 4), envelope.Length);
        var overlap = RewriteStream(original, "WordDocument", word);
        var overlapSnapshot = (byte[])overlap.Clone();
        using var overlappingEditor = DocEditor.Open(overlap);
        overlappingEditor.RemoveEmailEnvelope();
        Assert.Throws<InvalidDataException>(() => overlappingEditor.Save());
        Assert.Equal(overlapSnapshot, overlap);

        var table = ReadStream(original, "1Table");
        table[(int)envelope.Offset] = 0;
        var unknown = RewriteStream(original, "1Table", table);
        using var unknownEditor = DocEditor.Open(unknown);
        unknownEditor.RemoveEmailEnvelope();
        Assert.Throws<InvalidDataException>(() => unknownEditor.Save());
        // Visibility edits retain opaque envelope payloads.
        using var visibilityEditor = DocEditor.Open(unknown);
        var hidden = visibilityEditor.SetEmailEnvelopeVisibility(false).Save();
        Assert.Equal(EnvelopeBytes(unknown), EnvelopeBytes(hidden));
    }

    [Theory]
    [InlineData(false, 0)]
    [InlineData(false, 1)]
    [InlineData(true, 0)]
    [InlineData(true, 1)]
    public void RejectsPartialEnvelopeOverlapInEitherTableStream(bool tableOne, int shift)
    {
        using var editor = DocEditor.Open(CreateContainer(tableOne));
        var original = editor.SetEmailEnvelope(new DocEmailEnvelope { Subject = "Keep" }).Save();
        var word = ReadStream(original, "WordDocument");
        var envelope = LocateEnvelope(word);
        BinaryPrimitives.WriteUInt32LittleEndian(word.AsSpan(154 + 33 * 8), envelope.Offset + (uint)shift);
        BinaryPrimitives.WriteUInt32LittleEndian(word.AsSpan(154 + 33 * 8 + 4), envelope.Length - 2);
        var overlapping = RewriteStream(original, "WordDocument", word);
        using var overlappingEditor = DocEditor.Open(overlapping);
        overlappingEditor.RemoveEmailEnvelope();
        Assert.Throws<InvalidDataException>(() => overlappingEditor.Save());
    }

    [Fact]
    public void RejectsUnsupportedOldEnvelopeVersionBeforeReplacing()
    {
        using var editor = DocEditor.Open(CreateContainer(true));
        var original = editor.SetEmailEnvelope(new DocEmailEnvelope { Subject = "Keep" }).Save();
        var envelope = LocateEnvelope(ReadStream(original, "WordDocument"));
        var table = ReadStream(original, "1Table");
        BinaryPrimitives.WriteUInt32LittleEndian(table.AsSpan((int)envelope.Offset + 16), 9);
        var unsupported = RewriteStream(original, "1Table", table);
        using var replacementEditor = DocEditor.Open(unsupported);
        replacementEditor.SetEmailEnvelope(new DocEmailEnvelope { Subject = "Replace" });
        Assert.Throws<NotSupportedException>(() => replacementEditor.Save());
    }

    [Fact]
    public void VisibilityEditRejectsDopByteAliasedToEnvelope()
    {
        using var editor = DocEditor.Open(CreateContainer(true));
        var original = editor.SetEmailEnvelope(new DocEmailEnvelope { Subject = "Preserve" }).Save();
        var word = ReadStream(original, "WordDocument");
        var envelope = LocateEnvelope(word);
        var dopField = 154 + 31 * 8;
        BinaryPrimitives.WriteUInt32LittleEndian(word.AsSpan(dopField), envelope.Offset - 504);
        BinaryPrimitives.WriteUInt32LittleEndian(word.AsSpan(dopField + 4), 505);
        var aliased = RewriteStream(original, "WordDocument", word);
        using var visibilityEditor = DocEditor.Open(aliased);
        visibilityEditor.SetEmailEnvelopeVisibility(false);
        Assert.Throws<InvalidDataException>(() => visibilityEditor.Save());
        Assert.Equal(EnvelopeBytes(aliased), EnvelopeBytes(original));
    }

    [Fact]
    public void ReaderPrefersSmtpAddressEvenWhenGenericEmailComesLater()
    {
        using var editor = DocEditor.Open(CreateContainer(true));
        var original = editor.SetEmailEnvelope(new DocEmailEnvelope
        {
            To = new[] { new DocEmailAddress("a@b.co") }
        }).Save();
        long firstOffset, secondOffset;
        using (var input = new MemoryStream(original))
        using (var structure = new DocStructureWalker().Accept(input,
                   new DocStructurePrintVisitor(new StringWriter())))
        {
            var recipient = Assert.Single(FindNamed(structure.Root,
                "EnvRecipientCollection", "MessageRecipients").Children);
            firstOffset = recipient.Children.Single(x => x.Attributes.TryGetValue("tag", out var tag) &&
                tag == "0x3003001F").Offset!.Value;
            secondOffset = recipient.Children.Single(x => x.Attributes.TryGetValue("tag", out var tag) &&
                tag == "0x39FE001F").Offset!.Value;
        }
        var table = ReadStream(original, "1Table");
        BinaryPrimitives.WriteUInt32LittleEndian(table.AsSpan((int)firstOffset), 0x39FE001F);
        BinaryPrimitives.WriteUInt32LittleEndian(table.AsSpan((int)secondOffset), 0x3003001F);
        Encoding.Unicode.GetBytes("EX:foo").CopyTo(table, (int)secondOffset + 6);
        using var reader = DocEditor.Open(RewriteStream(original, "1Table", table));
        Assert.Equal("a@b.co", Assert.Single(reader.ReadEmailEnvelope()!.To).Address);
    }

    [Fact]
    public void RejectsInvalidSettingsAndHonorsCancellation()
    {
        using var editor = DocEditor.Open(CreatePlainDoc());
        Assert.Throws<ArgumentException>(() => editor.SetEmailEnvelope(new DocEmailEnvelope { Subject = "bad\0subject" }));
        Assert.Throws<ArgumentException>(() => editor.SetEmailEnvelope(new DocEmailEnvelope
        {
            Attachments = new[] { new DocEmailAttachment("../bad.txt", new byte[0]) }
        }));
        editor.SetEmailEnvelope(new DocEmailEnvelope { Subject = "Valid" });
        Assert.Throws<OperationCanceledException>(() => editor.Save(new CancellationToken(true)));
    }

    private static byte[] CreatePlainDoc()
    {
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source,
                   DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text("Hello Ω")))));
            main.Document.Save();
        }
        return DxpDocExport.Export(source.ToArray());
    }

    private static byte[] CreateContainer(bool tableOne)
    {
        using var output = new MemoryStream();
        using (var root = RootStorage.Create(output, OpenMcdf.Version.V3, StorageModeFlags.LeaveOpen))
        {
            var word = new byte[1200];
            BinaryPrimitives.WriteUInt16LittleEndian(word, 0xA5EC);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(2), 0x00C1);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(10), tableOne ? (ushort)0x0200 : (ushort)0);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(32), 14);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(62), 22);
            BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(152), 108);
            BinaryPrimitives.WriteUInt32LittleEndian(word.AsSpan(154 + 31 * 8 + 4), 600);
            using (var stream = root.CreateStream("WordDocument")) stream.Write(word);
            using (var stream = root.CreateStream(tableOne ? "1Table" : "0Table")) stream.Write(new byte[600]);
            using (var stream = root.CreateStream("Unrelated")) stream.Write(Encoding.UTF8.GetBytes("Keep this stream"));
        }
        return output.ToArray();
    }

    private static string ReadText(byte[] doc)
    {
        using var input = new MemoryStream(doc);
        using var index = new DocTextIndexWalker().Index(input);
        return string.Concat(index.GetPartSpans("Main").Select(x => x.Text));
    }

    private static byte[] EnvelopeBytes(byte[] doc)
    {
        var word = ReadStream(doc, "WordDocument");
        var (offset, length) = LocateEnvelope(word);
        var table = ReadStream(doc, (word[11] & 2) != 0 ? "1Table" : "0Table");
        return table.AsSpan((int)offset, (int)length).ToArray();
    }

    private static (uint Offset, uint Length) LocateEnvelope(byte[] word) =>
        (BinaryPrimitives.ReadUInt32LittleEndian(word.AsSpan(154 + 97 * 8)),
         BinaryPrimitives.ReadUInt32LittleEndian(word.AsSpan(154 + 97 * 8 + 4)));

    private static byte[] ReadStream(byte[] doc, string name)
    {
        using var input = new MemoryStream(doc, false);
        using var root = RootStorage.Open(input);
        using var stream = root.OpenStream(name);
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    private static byte[] RewriteStream(byte[] doc, string name, byte[] bytes)
    {
        using var output = new MemoryStream();
        output.Write(doc);
        output.Position = 0;
        using (var root = RootStorage.Open(output, StorageModeFlags.LeaveOpen))
        using (var stream = root.OpenStream(name))
        {
            stream.Position = 0;
            stream.Write(bytes);
        }
        return output.ToArray();
    }
}
