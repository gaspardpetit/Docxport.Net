using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Fields;
using DocxportNet.Fields.Eval;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpRefFieldEvalFrame : DxpValueFieldEvalFrame
{
    public DxpRefFieldEvalFrame(DxpIVisitor? next, DxpFieldEval eval, ILogger? logger, string? instructionText, Run? codeRun = null)
        : base(next, eval, logger, instructionText, codeRun)
    {}

    protected override bool Evaluate(DxpIDocumentContext d)
    {
        if (TryReplayStructuredBookmark(d))
            return true;
        return base.Evaluate(d);
    }

    private bool TryReplayStructuredBookmark(DxpIDocumentContext d)
    {
        var instruction = InstructionText;
        if (string.IsNullOrWhiteSpace(instruction))
            return false;

        var parser = new DxpFieldParser();
        var parse = parser.Parse(instruction);
        if (!string.Equals(parse.Ast.FieldType, "REF", StringComparison.OrdinalIgnoreCase))
            return false;

        if (parse.Ast.FormatSpecs.Count > 0)
            return false;

        var switches = ParseNonFormatSwitches(parse.Ast.RawText);
        if (HasRefSwitches(switches))
            return false;

        if (string.IsNullOrWhiteSpace(parse.Ast.ArgumentsText))
            return false;

        var tokens = TokenizeArgs(parse.Ast.ArgumentsText);
        if (tokens.Count == 0 || string.IsNullOrWhiteSpace(tokens[0]))
            return false;

        var bookmark = tokens[0];

        if (EvalContext.TryGetBookmarkNodes(bookmark, out var evalNodes))
        {
            ReplayStructuredBookmark(evalNodes, d);
            return true;
        }

        var docNodes = d.DocumentIndex?.BookmarkNodes;
        if (docNodes != null && docNodes.TryGetValue(bookmark, out var buffer))
        {
            ReplayStructuredBookmark(buffer, d);
            return true;
        }

        return false;
    }

    private void ReplayStructuredBookmark(DxpFieldNodeBuffer buffer, DxpIDocumentContext d)
    {
        if (Next == null)
            return;

        if (buffer.TryGetRunSegments(out var segments))
        {
            foreach (var segment in segments)
            {
                if (string.IsNullOrEmpty(segment.text))
                    continue;
                var run = new Run();
                if (segment.props != null)
                    run.RunProperties = (RunProperties)segment.props.CloneNode(true);
                var paragraph = new Paragraph();
                paragraph.AppendChild(run);
                DxpFieldFrames.EmitTextInRun(segment.text, d, run, Next);
            }
            return;
        }

        buffer.Replay(Next, d);
    }

    private static Dictionary<char, string?> ParseNonFormatSwitches(string rawText)
    {
        var switches = new Dictionary<char, string?>();
        bool inQuote = false;
        int braceDepth = 0;
        var switchStarts = new List<int>();
        for (int i = 0; i < rawText.Length; i++)
        {
            char ch = rawText[i];
            if (inQuote && ch == '\\' && i + 1 < rawText.Length && rawText[i + 1] == '"')
            {
                i++;
                continue;
            }
            if (ch == '"')
            {
                inQuote = !inQuote;
                continue;
            }
            if (!inQuote)
            {
                if (ch == '{')
                {
                    braceDepth++;
                    continue;
                }
                if (ch == '}' && braceDepth > 0)
                {
                    braceDepth--;
                    continue;
                }
            }
            if (!inQuote && braceDepth == 0 && ch == '\\')
                switchStarts.Add(i);
        }

        for (int i = 0; i < switchStarts.Count; i++)
        {
            int start = switchStarts[i];
            int end = i + 1 < switchStarts.Count ? switchStarts[i + 1] : rawText.Length;
            string seg = rawText.Substring(start, end - start).Trim();
            int index = 0;
            while (index < seg.Length && seg[index] == '\\')
                index++;
            while (index < seg.Length && char.IsWhiteSpace(seg[index]))
                index++;
            if (index >= seg.Length)
                continue;
            char kind = seg[index];
            if (kind is '*' or '#' or '@')
                continue;
            string? arg = index + 1 < seg.Length ? seg.Substring(index + 1).Trim() : null;
            string? unquoted = UnquoteSwitchArg(arg);
            switches[kind] = unquoted;
        }

        return switches;
    }

    private static string? UnquoteSwitchArg(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return value;
        string trimmed = value!.Trim();
        if (trimmed.Length >= 2 && trimmed[0] == '"' && trimmed[trimmed.Length - 1] == '"')
            return trimmed.Substring(1, trimmed.Length - 2);
        return trimmed;
    }

    private static bool HasRefSwitches(Dictionary<char, string?> switches)
    {
        foreach (var key in switches.Keys)
        {
            switch (key)
            {
                case 'd':
                case 'f':
                case 'h':
                case 'n':
                case 'p':
                case 'r':
                case 't':
                case 'w':
                    return true;
            }
        }
        return false;
    }

    private static List<string> TokenizeArgs(string text)
    {
        var tokens = new List<string>();
        bool inQuote = false;
        bool justClosedQuote = false;
        var current = new System.Text.StringBuilder();
        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            if (ch == '"')
            {
                if (inQuote && i > 0 && text[i - 1] == '\\')
                {
                    if (current.Length > 0)
                        current.Length -= 1;
                    current.Append('"');
                    continue;
                }
                inQuote = !inQuote;
                if (!inQuote)
                    justClosedQuote = true;
                continue;
            }

            if (!inQuote && ch == '{')
            {
                int start = i;
                int depth = 0;
                for (; i < text.Length; i++)
                {
                    if (text[i] == '{')
                        depth++;
                    else if (text[i] == '}')
                    {
                        depth--;
                        if (depth == 0)
                        {
                            i++;
                            break;
                        }
                    }
                }
                string token = text.Substring(start, i - start).Trim();
                if (current.Length > 0)
                {
                    tokens.Add(current.ToString());
                    current.Clear();
                }
                if (token.Length > 0)
                    tokens.Add(token);
                justClosedQuote = false;
                i--;
                continue;
            }

            if (!inQuote && char.IsWhiteSpace(ch))
            {
                if (current.Length > 0)
                {
                    tokens.Add(current.ToString());
                    current.Clear();
                    justClosedQuote = false;
                }
                else if (justClosedQuote)
                {
                    tokens.Add(string.Empty);
                    justClosedQuote = false;
                }
                continue;
            }

            current.Append(ch);
            justClosedQuote = false;
        }

        if (current.Length > 0)
            tokens.Add(current.ToString());
        else if (justClosedQuote)
            tokens.Add(string.Empty);
        return tokens;
    }
}
