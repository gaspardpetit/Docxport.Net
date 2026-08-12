using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocxportNet.API;
using DocxportNet.Fields.Resolution;
using DocxportNet.Middleware;
using DocxportNet.Walker;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpIncludeTextFieldEvalFrame : DxpValueFieldEvalFrame
{
    public DxpIncludeTextFieldEvalFrame(
        DxpIVisitor? next,
        DxpFieldEval eval,
        ILogger? logger,
        string? instructionText)
        : base(next, eval, logger, instructionText)
    {
    }

    protected override bool Evaluate(DxpIDocumentContext d)
    {
        if (Next == null || string.IsNullOrWhiteSpace(InstructionText))
            return ReplayCache(d);

        var parse = new DxpFieldParser().Parse(InstructionText!);
        if (!string.Equals(parse.Ast.FieldType, "INCLUDETEXT", StringComparison.OrdinalIgnoreCase) ||
            string.IsNullOrWhiteSpace(parse.Ast.ArgumentsText))
        {
            return ReplayCache(d);
        }

        var tokens = TokenizeArgs(parse.Ast.ArgumentsText!);
        if (tokens.Count is < 1 or > 2 || string.IsNullOrWhiteSpace(tokens[0]))
        {
            Logger?.LogInformation("INCLUDETEXT arguments are invalid; using cached result.");
            return ReplayCache(d);
        }

        string path = tokens[0];
        string? bookmark = tokens.Count == 2 ? tokens[1] : null;
        if (bookmark != null && string.IsNullOrWhiteSpace(bookmark))
            return ReplayCache(d);
        bool htmlSwitch = HasHtmlSwitch(parse.Ast.RawText);

        var resolver = EvalContext.IncludeTextResolver;
        if (resolver == null)
        {
            Logger?.LogInformation("INCLUDETEXT resolver is not configured; using cached result.");
            return ReplayCache(d);
        }

        DxpIncludeTextSource? source;
        try
        {
            source = resolver.ResolveAsync(
                new DxpIncludeTextRequest(path),
                EvalContext,
                CancellationToken.None).GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            Logger?.LogWarning(ex, "INCLUDETEXT source resolution failed for '{Path}'; using cached result.", path);
            return ReplayCache(d);
        }

        if (source == null || source.Content.Length == 0)
        {
            Logger?.LogWarning("INCLUDETEXT source '{Path}' was not resolved; using cached result.", path);
            return ReplayCache(d);
        }

        byte[] content = source.Content;
        bool isHtml = htmlSwitch
            || source.Format == DxpIncludeTextSourceFormat.Html
            || source.Format == DxpIncludeTextSourceFormat.Auto && IsHtmlPath(path);
        if (isHtml)
        {
            try
            {
                content = EvalContext.ConvertHtmlIncludeAsync(source.Content, CancellationToken.None)
                    .GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                Logger?.LogWarning(ex, "INCLUDETEXT HTML conversion failed for '{Path}'; using cached result.", path);
                return ReplayCache(d);
            }
        }

        var expansion = new DxpIncludeTextExpansion(
            path,
            source.Identity,
            content,
            bookmark,
            CachedResultBuffer,
            Eval,
            Logger);
        if (EvalContext.IncludeTextSpliceCollector?.Record(expansion) == true)
            return true;

        if (!EvalContext.TryEnterIncludeText(source.Identity, out string? recursionError))
        {
            Logger?.LogWarning("{Error} Using cached INCLUDETEXT result.", recursionError);
            return ReplayCache(d);
        }

        try
        {
            MemoryStream? stream = null;
            WordprocessingDocument? document = null;
            try
            {
                stream = new MemoryStream(content, writable: false);
                document = WordprocessingDocument.Open(stream, false);
                if (document.MainDocumentPart?.Document?.Body == null)
                    throw new InvalidOperationException("DOCX has no main document body.");
            }
            catch (Exception ex) when (ex is OpenXmlPackageException or FileFormatException or InvalidOperationException)
            {
                document?.Dispose();
                stream?.Dispose();
                Logger?.LogWarning(ex, "INCLUDETEXT source '{Path}' is not a valid DOCX; using cached result.", path);
                return ReplayCache(d);
            }

            using (stream)
            using (document)
            {
                IReadOnlyList<DocumentFormat.OpenXml.OpenXmlElement>? blocks = null;
                if (!string.IsNullOrWhiteSpace(bookmark))
                {
                    var body = document.MainDocumentPart!.Document.Body!;
                    if (!DxpBookmarkRangeProjector.TryProject(body, bookmark!, out var projected, out var error))
                    {
                        Logger?.LogWarning("{Error} Using cached INCLUDETEXT result.", error);
                        return ReplayCache(d);
                    }
                    blocks = projected;
                }

                var pipeline = DxpVisitorMiddleware.Chain(
                    Next,
                    next => DxpFieldEvalMiddleware.CreateEvaluatedFieldMiddleware(next, Eval, logger: Logger),
                    next => new DxpContextMiddleware(next, Logger));
                new DxpWalker(Logger).AcceptEmbeddedBody(document, pipeline, blocks);
                return true;
            }
        }
        finally
        {
            EvalContext.ExitIncludeText(source.Identity);
        }
    }

    private bool ReplayCache(DxpIDocumentContext d)
    {
        var collector = EvalContext.IncludeTextSpliceCollector;
        if (Next != null && CachedResultBuffer != null)
        {
            if (collector != null)
                CachedResultBuffer.Replay(Next, d);
            else
                Next.VisitText(new DocumentFormat.OpenXml.Wordprocessing.Text(CachedResultBuffer.ToPlainText()), d);
        }
        collector?.Complete();
        return true;
    }

    private static bool HasHtmlSwitch(string rawText)
        => Regex.IsMatch(rawText, @"\\c\s+(?:\""?HTML\""?)", RegexOptions.IgnoreCase);

    private static bool IsHtmlPath(string path)
        => path.EndsWith(".htm", StringComparison.OrdinalIgnoreCase)
            || path.EndsWith(".html", StringComparison.OrdinalIgnoreCase);

    private static List<string> TokenizeArgs(string text)
    {
        var tokens = new List<string>();
        var current = new StringBuilder();
        bool inQuote = false;

        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            if (ch == '"')
            {
                inQuote = !inQuote;
                continue;
            }

            if (!inQuote && char.IsWhiteSpace(ch))
            {
                if (current.Length > 0)
                {
                    tokens.Add(current.ToString());
                    current.Clear();
                }
                continue;
            }

            current.Append(ch);
        }

        if (current.Length > 0)
            tokens.Add(current.ToString());
        return tokens;
    }
}
