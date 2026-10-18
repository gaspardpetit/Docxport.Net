using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;

namespace DocxportNet.Visitors.Pdf;

internal static class DxpPdfSourceSpaceRestorer
{
    private static readonly Regex TextCommand = new(@"\(((?:\\.|[^\\)])*)\) Tj");
    private static readonly Regex PositionCommand = new(@"\s*[-\d.]+ [-\d.]+ Td\s*");
    private static readonly Regex ColorCommand = new(@"\s*(?:[-\d.]+ ){2}[-\d.]+ rg\s*");
    private static readonly Regex RaisedMarkerFont = new(@"/F\d+ [0-8](?:\.\d+)? Tf");
    private static readonly Regex RaisedMarkerBetween = new(
        @"^\s*/F\d+ [0-8](?:\.\d+)? Tf\s*[-\d.]+ [-\d.]+ Td\s*(?<marker>\(\d+\) Tj)\s*/F\d+ [-\d.]+ Tf\s*[-\d.]+ [-\d.]+ Td\s*$");
    private static readonly Regex UnderlineBlockPosition = new(
        @"^\s*ET\s*(?:[-\d.]+ [-\d.]+ m\s*[-\d.]+ [-\d.]+ l\s*S\s*)?BT\s*[-\d.]+ [-\d.]+ Td\s*$");

    public static int Restore(PdfDocument document, WordprocessingDocument source)
    {
        var paragraphs = source.MainDocumentPart?.Document.Body?.Elements<Paragraph>()
            .Select(paragraph => string.Concat(paragraph.Descendants()
                .Where(element => element is Text or NoBreakHyphen)
                .Select(element => element is Text text ? text.Text : "-")))
            .ToArray() ?? [];
        var starts = paragraphs.GroupBy(paragraph => string.Join("\0", Words(paragraph).Take(3)))
            .ToDictionary(group => group.Key, group => group.Count());
        var streams = document.Pages.Cast<PdfPage>()
            .Select(page => page.Contents.Elements
                .Select(entry => (entry as PdfReference)?.Value as PdfContent)
                .Where(content => content?.Stream != null && !content.Elements.ContainsKey("/Filter"))
                .Cast<PdfContent>().ToArray()).ToArray();
        var pending = new Dictionary<PdfContent, string>();
        var restored = 0;
        foreach (var paragraph in paragraphs)
        {
            var words = Words(paragraph);
            if (paragraph.Length < 80 || words.Length == 0 ||
                starts[string.Join("\0", words.Take(3))] != 1) continue;
            var cursor = 0;
            var anchor = words[0];
            var parts = new List<(PdfContent Content, string Revised, int Count)>();
            var lastPage = -1;
            while (cursor < paragraph.Length)
            {
                var matches = new List<(int Page, PdfContent Content, string Revised, int Count,
                    int Cursor)>();
                for (var page = lastPage + 1; page < streams.Length; page++)
                    foreach (var content in streams[page])
                    {
                        var text = pending.TryGetValue(content, out var revised)
                            ? revised : Encoding.Latin1.GetString(content.Stream!.Value);
                        if (TryRestore(text, paragraph, anchor, cursor, out var candidate,
                                out var count, out var nextCursor) && nextCursor > cursor)
                            matches.Add((page, content, candidate, count, nextCursor));
                    }
                if (matches.Count != 1) break;
                var match = matches[0];
                parts.Add((match.Content, match.Revised, match.Count));
                lastPage = match.Page;
                cursor = match.Cursor;
                if (cursor < paragraph.Length)
                    anchor = Words(paragraph[cursor..]).FirstOrDefault() ?? "";
                if (anchor.Length == 0 && cursor < paragraph.Length) break;
            }
            if (cursor != paragraph.Length) continue;
            foreach (var part in parts)
            {
                pending[part.Content] = part.Revised;
                restored += part.Count;
            }
        }
        foreach (var (content, revised) in pending)
            content.Stream!.Value = Encoding.Latin1.GetBytes(revised);
        return restored;
    }

    internal static bool TryRestore(string content, string source, string anchor, int startPosition,
        out string revised, out int insertionCount, out int cursor)
    {
        revised = content;
        insertionCount = 0;
        cursor = startPosition;
        var position = source.IndexOf(anchor, startPosition, StringComparison.Ordinal);
        if (position < 0 || !string.IsNullOrWhiteSpace(source[startPosition..position]))
            return false;
        var commands = TextCommand.Matches(content).Cast<Match>().ToArray();
        var candidates = new List<int>();
        for (var index = 0; index <= commands.Length - 4; index++)
        {
            var probe = position;
            var aligned = true;
            for (var offset = 0; offset < 4; offset++)
            {
                if (!TryDecode(commands[index + offset].Groups[1].Value, out var text))
                {
                    aligned = false;
                    break;
                }
                var found = source.IndexOf(text, probe, StringComparison.Ordinal);
                if (found < 0 || found - probe > 35 ||
                    !string.IsNullOrWhiteSpace(source[probe..found]))
                {
                    aligned = false;
                    break;
                }
                probe = found + text.Length;
            }
            if (aligned) candidates.Add(index);
        }
        if (candidates.Count != 1) return false;
        Match? previous = null;
        var insertions = new List<int>();
        for (var index = candidates[0]; index < commands.Length; index++)
        {
            var command = commands[index];
            if (!TryDecode(command.Groups[1].Value, out var text)) return false;
            if (text.Length > 0 && text.All(char.IsDigit) && previous != null &&
                RaisedMarkerFont.IsMatch(content[(previous.Index + previous.Length)..command.Index]))
                continue;
            var found = source.IndexOf(text, position, StringComparison.Ordinal);
            if (found < 0 || found - position > 35 ||
                !string.IsNullOrWhiteSpace(source[position..found]))
            {
                if (text.Length > 0 && text.All(char.IsDigit)) continue;
                return false;
            }
            if (previous != null && found > position)
            {
                var between = content[(previous.Index + previous.Length)..command.Index];
                var marker = RaisedMarkerBetween.Match(between);
                if (marker.Success && found - position >= 2)
                {
                    insertions.Add(previous.Index + previous.Length);
                    insertions.Add(previous.Index + previous.Length + marker.Groups["marker"].Index +
                        marker.Groups["marker"].Length);
                }
                between = ColorCommand.Replace(between, "");
                var movement = PositionCommand.Match(between);
                if (movement.Success && movement.Index == 0 && movement.Length == between.Length ||
                    UnderlineBlockPosition.IsMatch(between))
                    insertions.Add(previous.Index + previous.Length);
            }
            position = found + text.Length;
            previous = command;
            if (string.IsNullOrWhiteSpace(source[position..]))
            {
                cursor = source.Length;
                break;
            }
        }
        if (cursor != source.Length) cursor = position;
        var builder = new StringBuilder(content);
        for (var index = insertions.Count - 1; index >= 0; index--)
            builder.Insert(insertions[index], "\n( ) Tj");
        revised = builder.ToString();
        insertionCount = insertions.Count;
        return true;
    }

    private static string[] Words(string text) => text.Split((char[]?)null,
        StringSplitOptions.RemoveEmptyEntries);

    private static bool TryDecode(string literal, out string text)
    {
        var unescaped = Regex.Replace(literal, @"\\([\\()])", "$1");
        var builder = new StringBuilder(unescaped.Length);
        foreach (var character in unescaped)
        {
            var decoded = character switch
            {
                '\u0080' => '\u20ac', '\u0082' => '\u201a', '\u0083' => '\u0192',
                '\u0084' => '\u201e', '\u0085' => '\u2026', '\u0086' => '\u2020',
                '\u0087' => '\u2021', '\u0088' => '\u02c6', '\u0089' => '\u2030',
                '\u008a' => '\u0160', '\u008b' => '\u2039', '\u008c' => '\u0152',
                '\u008e' => '\u017d', '\u0091' => '\u2018', '\u0092' => '\u2019',
                '\u0093' => '\u201c', '\u0094' => '\u201d', '\u0095' => '\u2022',
                '\u0096' => '\u2013', '\u0097' => '\u2014', '\u0098' => '\u02dc',
                '\u0099' => '\u2122', '\u009a' => '\u0161', '\u009b' => '\u203a',
                '\u009c' => '\u0153', '\u009e' => '\u017e', '\u009f' => '\u0178',
                _ => character
            };
            if (decoded is '\u0081' or '\u008d' or '\u008f' or '\u0090' or '\u009d')
            {
                text = "";
                return false;
            }
            builder.Append(decoded);
        }
        text = builder.ToString();
        return true;
    }
}
