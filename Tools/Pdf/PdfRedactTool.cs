using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core;
using Color = System.Drawing.Color;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for redacting (blacking out) text or areas on PDF pages
/// </summary>
public class PdfRedactTool : IAsposeTool
{
    public string Description =>
        @"Redact (black out) text or area on PDF page. This permanently removes the underlying content.

Usage examples:
- Redact area: pdf_redact(path='doc.pdf', pageIndex=1, x=100, y=100, width=200, height=50)
- Redact with color: pdf_redact(path='doc.pdf', pageIndex=1, x=100, y=100, width=200, height=50, fillColor='255,0,0')
- Redact with overlay: pdf_redact(path='doc.pdf', pageIndex=1, x=100, y=100, width=200, height=50, overlayText='[REDACTED]')
- Redact by text search: pdf_redact(path='doc.pdf', textToRedact='confidential')
- Redact by text on page: pdf_redact(path='doc.pdf', pageIndex=1, textToRedact='secret', caseSensitive=false)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "PDF file path (required)"
            },
            pageIndex = new
            {
                type = "number",
                description = "Page index (1-based, optional for text search - searches all pages if not specified)"
            },
            x = new
            {
                type = "number",
                description =
                    "X position of redaction area in PDF coordinates, origin at bottom-left corner (required for area redaction)"
            },
            y = new
            {
                type = "number",
                description =
                    "Y position of redaction area in PDF coordinates, origin at bottom-left corner (required for area redaction)"
            },
            width = new
            {
                type = "number",
                description = "Width of redaction area in PDF points (required for area redaction)"
            },
            height = new
            {
                type = "number",
                description = "Height of redaction area in PDF points (required for area redaction)"
            },
            textToRedact = new
            {
                type = "string",
                description =
                    "Text to search and redact (alternative to area redaction, searches and redacts all occurrences)"
            },
            caseSensitive = new
            {
                type = "boolean",
                description = "Whether text search is case sensitive (default: true, only for textToRedact)"
            },
            fillColor = new
            {
                type = "string",
                description = "Fill color (optional, default: black, format: 'R,G,B' or color name)"
            },
            overlayText = new
            {
                type = "string",
                description = "Text to display over the redacted area (optional, e.g., '[REDACTED]')"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to input path)"
            }
        },
        required = new[] { "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return Task.Run(() =>
        {
            var textToRedact = ArgumentHelper.GetStringNullable(arguments, "textToRedact");
            var fillColor = ArgumentHelper.GetStringNullable(arguments, "fillColor");
            var overlayText = ArgumentHelper.GetStringNullable(arguments, "overlayText");

            using var document = new Document(path);

            if (!string.IsNullOrEmpty(textToRedact))
                return RedactByText(document, arguments, textToRedact, fillColor, overlayText, outputPath);

            return RedactByArea(document, arguments, fillColor, overlayText, outputPath);
        });
    }

    /// <summary>
    ///     Redacts text by searching for occurrences
    /// </summary>
    /// <param name="document">PDF document</param>
    /// <param name="arguments">JSON arguments</param>
    /// <param name="textToRedact">Text to search and redact</param>
    /// <param name="fillColor">Fill color string</param>
    /// <param name="overlayText">Overlay text</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Result message</returns>
    private string RedactByText(Document document, JsonObject? arguments, string textToRedact,
        string? fillColor, string? overlayText, string outputPath)
    {
        var pageIndex = ArgumentHelper.GetIntNullable(arguments, "pageIndex");
        var caseSensitive = ArgumentHelper.GetBool(arguments, "caseSensitive", true);

        if (pageIndex.HasValue && (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count))
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        // For case-insensitive search, use regex mode with escaped pattern
        // TextSearchOptions(bool) controls whether to use regex, not case sensitivity
        TextFragmentAbsorber absorber;
        if (caseSensitive)
        {
            absorber = new TextFragmentAbsorber(textToRedact);
        }
        else
        {
            // Escape regex special characters and use (?i) for case-insensitive matching
            var escapedPattern = Regex.Escape(textToRedact);
            var textSearchOptions = new TextSearchOptions(true); // Enable regex mode
            absorber = new TextFragmentAbsorber($"(?i){escapedPattern}", textSearchOptions);
        }

        var redactionCount = 0;
        var pagesAffected = new HashSet<int>();

        if (pageIndex.HasValue)
            document.Pages[pageIndex.Value].Accept(absorber);
        else
            document.Pages.Accept(absorber);

        foreach (var fragment in absorber.TextFragments)
        {
            var page = fragment.Page;
            var rect = fragment.Rectangle;

            var redactionAnnotation = new RedactionAnnotation(page, rect);
            ApplyRedactionStyle(redactionAnnotation, fillColor, overlayText);

            page.Annotations.Add(redactionAnnotation);
            redactionAnnotation.Redact();

            redactionCount++;
            pagesAffected.Add(document.Pages.IndexOf(page));
        }

        if (redactionCount == 0)
            return $"No occurrences of '{textToRedact}' found. No redactions applied.";

        document.Save(outputPath);
        var pageInfo = pagesAffected.Count == 1
            ? $"page {pagesAffected.First()}"
            : $"{pagesAffected.Count} pages";
        return $"Redacted {redactionCount} occurrence(s) of '{textToRedact}' on {pageInfo}. Output: {outputPath}";
    }

    /// <summary>
    ///     Redacts a specific area on a page
    /// </summary>
    /// <param name="document">PDF document</param>
    /// <param name="arguments">JSON arguments</param>
    /// <param name="fillColor">Fill color string</param>
    /// <param name="overlayText">Overlay text</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Result message</returns>
    private string RedactByArea(Document document, JsonObject? arguments,
        string? fillColor, string? overlayText, string outputPath)
    {
        var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");
        var x = ArgumentHelper.GetDouble(arguments, "x");
        var y = ArgumentHelper.GetDouble(arguments, "y");
        var width = ArgumentHelper.GetDouble(arguments, "width");
        var height = ArgumentHelper.GetDouble(arguments, "height");

        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var rect = new Rectangle(x, y, x + width, y + height);

        var redactionAnnotation = new RedactionAnnotation(page, rect);
        ApplyRedactionStyle(redactionAnnotation, fillColor, overlayText);

        page.Annotations.Add(redactionAnnotation);
        redactionAnnotation.Redact();
        document.Save(outputPath);

        return $"Redaction applied to page {pageIndex}. Output: {outputPath}";
    }

    /// <summary>
    ///     Applies fill color and overlay text to a redaction annotation
    /// </summary>
    /// <param name="annotation">Redaction annotation to style</param>
    /// <param name="fillColor">Fill color string</param>
    /// <param name="overlayText">Overlay text</param>
    private static void ApplyRedactionStyle(RedactionAnnotation annotation, string? fillColor, string? overlayText)
    {
        if (!string.IsNullOrEmpty(fillColor))
        {
            var systemColor = ColorHelper.ParseColor(fillColor, Color.Black);
            annotation.FillColor = ColorHelper.ToPdfColor(systemColor);
        }
        else
        {
            annotation.FillColor = Aspose.Pdf.Color.Black;
        }

        if (!string.IsNullOrEmpty(overlayText))
        {
            annotation.OverlayText = overlayText;
            annotation.Repeat = true;
        }
    }
}