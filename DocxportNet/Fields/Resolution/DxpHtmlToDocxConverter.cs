using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using HtmlToOpenXml.IO;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxportNet.Fields.Resolution;

public interface IDxpHtmlToDocxConverter
{
    Task<byte[]> ConvertAsync(byte[] html, CancellationToken cancellationToken = default);
}

public sealed class DxpHtmlToDocxConverter : IDxpHtmlToDocxConverter
{
    private static readonly UTF8Encoding s_strictUtf8 = new(false, true);
    private static readonly TimeSpan s_regexTimeout = TimeSpan.FromSeconds(2);

    static DxpHtmlToDocxConverter() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    public async Task<byte[]> ConvertAsync(byte[] html, CancellationToken cancellationToken = default)
    {
        if (html == null)
            throw new ArgumentNullException(nameof(html));

        string text = DecodeHtml(html);
        var parser = new HtmlParser();
        var parsed = await parser.ParseDocumentAsync(text, cancellationToken).ConfigureAwait(false);
        ValidateDataImages(parsed.Images.Select(image => image.GetAttribute("src")));
        foreach (var element in parsed.QuerySelectorAll("[style]"))
            element.SetAttribute("style", NormalizeCssDeclarations(element.GetAttribute("style") ?? string.Empty));
        foreach (var style in parsed.QuerySelectorAll("style"))
            style.TextContent = NormalizeCssDeclarations(style.TextContent);
        foreach (var meta in parsed.QuerySelectorAll("meta[charset]"))
            meta.SetAttribute("charset", "utf-8");
        foreach (var meta in parsed.QuerySelectorAll("meta[http-equiv]"))
        {
            string? content = meta.GetAttribute("content");
            if (content?.IndexOf("charset", StringComparison.OrdinalIgnoreCase) >= 0)
                meta.SetAttribute("content", "text/html; charset=utf-8");
        }
        var externalElements = parsed.Images
            .Where(image => !string.IsNullOrEmpty(image.GetAttribute("src"))
                && !IsDataImage(image.GetAttribute("src")))
            .ToArray();
        var externalImages = externalElements
            .Select((image, index) => new ExternalImage(
                image.GetAttribute("src")!, image.GetAttribute("alt"),
                $"https://docxport.invalid/external-image/{index}"))
            .ToArray();
        for (int index = 0; index < externalImages.Length; index++)
            externalElements[index].SetAttribute("src", externalImages[index].Placeholder);

        using var output = new MemoryStream();
        using (var document = WordprocessingDocument.Create(output, WordprocessingDocumentType.Document, true))
        {
            MainDocumentPart main = document.AddMainDocumentPart();
            main.Document = new Document(new Body());
            var converter = new HtmlConverter(main, NeverFetchWebRequest.Instance)
            {
                ImageProcessing = ImageProcessingMode.LinkExternal
            };
            await converter.ParseBody(parsed.DocumentElement?.OuterHtml ?? text, cancellationToken).ConfigureAwait(false);
            main.Document ??= new Document(new Body());
            RestoreExternalImages(main, externalImages);
            main.Document.Save();
        }
        return output.ToArray();
    }

    internal static string DecodeHtml(byte[] bytes)
    {
        if (bytes.Length >= 3 && bytes[0] == 0xef && bytes[1] == 0xbb && bytes[2] == 0xbf)
            return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
        if (bytes.Length >= 2 && bytes[0] == 0xff && bytes[1] == 0xfe)
            return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
        if (bytes.Length >= 2 && bytes[0] == 0xfe && bytes[1] == 0xff)
            return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);

        int sample = Math.Min(bytes.Length, 256);
        int evenZeros = 0;
        int oddZeros = 0;
        for (int index = 0; index < sample; index++)
        {
            if (bytes[index] != 0) continue;
            if ((index & 1) == 0) evenZeros++; else oddZeros++;
        }
        if (oddZeros > sample / 8) return Encoding.Unicode.GetString(bytes);
        if (evenZeros > sample / 8) return Encoding.BigEndianUnicode.GetString(bytes);

        string header = Encoding.GetEncoding(28591).GetString(bytes, 0, Math.Min(bytes.Length, 4096));
        Match charset = Regex.Match(header,
            "(?is)<meta\\b[^>]*(?:charset\\s*=\\s*['\"]?\\s*(?<charset>[a-z0-9._-]+)|content\\s*=\\s*['\"][^'\"]*charset\\s*=\\s*(?<charset>[a-z0-9._-]+))",
            RegexOptions.None, s_regexTimeout);
        if (charset.Success)
        {
            string name = charset.Groups["charset"].Value;
            if (name.Equals("unicode", StringComparison.OrdinalIgnoreCase)) name = "windows-1252";
            try
            {
                return Encoding.GetEncoding(name, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback)
                    .GetString(bytes);
            }
            catch (DecoderFallbackException) { }
            catch (ArgumentException) { }
        }

        try { return s_strictUtf8.GetString(bytes); }
        catch (DecoderFallbackException) { return Encoding.GetEncoding(1252).GetString(bytes); }
    }

    private static void ValidateDataImages(IEnumerable<string?> sources)
    {
        foreach (string? authored in sources)
        {
            string value = WebUtility.HtmlDecode(authored ?? string.Empty).Trim();
            if (!value.StartsWith("data:image/", StringComparison.OrdinalIgnoreCase)) continue;
            int comma = value.IndexOf(',');
            if (comma < 0) throw new InvalidDataException("An HTML data image is malformed.");

            string metadata = value.Substring(0, comma);
            string payload = value.Substring(comma + 1);
            byte[] decoded;
            try
            {
                decoded = metadata.IndexOf(";base64", StringComparison.OrdinalIgnoreCase) >= 0
                    ? Convert.FromBase64String(payload)
                    : Encoding.GetEncoding(28591).GetBytes(Uri.UnescapeDataString(payload));
            }
            catch (Exception ex) when (ex is FormatException || ex is ArgumentException)
            {
                throw new InvalidDataException("An HTML data image is malformed.", ex);
            }

            using var image = new MemoryStream(decoded, writable: false);
            if (!ImageHeader.TryDetectFileType(image, out _))
                throw new InvalidDataException("An HTML data image has an unsupported or invalid format.");
        }
    }

    private static bool IsDataImage(string? source)
        => WebUtility.HtmlDecode(source ?? string.Empty).Trim()
            .StartsWith("data:image/", StringComparison.OrdinalIgnoreCase);

    private static string NormalizeCssDeclarations(string css)
        => Regex.Replace(css, "(?i)(font-family\\s*:\\s*)(['\"])([^;\"',\\s]+)\\2", "$1$3",
            RegexOptions.None, s_regexTimeout);

    private static void RestoreExternalImages(MainDocumentPart main, IReadOnlyList<ExternalImage> images)
    {
        var drawings = main.Document?.Descendants<Drawing>()
            .Where(drawing => !string.IsNullOrEmpty(drawing.Descendants<A.Blip>().FirstOrDefault()?.Link?.Value))
            .ToArray() ?? [];
        int count = Math.Min(drawings.Length, images.Count);
        for (int index = 0; index < count; index++)
        {
            Drawing drawing = drawings[index];
            ExternalImage image = images[index];
            A.Blip blip = drawing.Descendants<A.Blip>().First();
            string relationshipId = blip.Link!.Value!;
            main.DeleteExternalRelationship(relationshipId);
            main.AddExternalRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                new Uri(image.Source, UriKind.RelativeOrAbsolute),
                relationshipId);
            DW.DocProperties? properties = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
            if (properties != null && !string.IsNullOrEmpty(image.AltText))
                properties.Description = image.AltText;
        }
    }

    private sealed record ExternalImage(string Source, string? AltText, string Placeholder);

    private sealed class NeverFetchWebRequest : IWebRequest
    {
        internal static readonly NeverFetchWebRequest Instance = new();
        // LinkExternal consults this capability before creating the relationship.
        // FetchAsync still guarantees that no resource is dereferenced.
        public bool SupportsProtocol(string protocol) => true;
        public Task<Resource?> FetchAsync(Uri requestUri, CancellationToken cancellationToken)
            => Task.FromResult<Resource?>(null);
    }
}
