using System.ComponentModel;
using System.Text.RegularExpressions;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;
using Color = System.Drawing.Color;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for redacting (blacking out) text or areas on PDF pages
/// </summary>
[McpServerToolType]
public class PdfRedactTool
{
    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PdfRedactTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public PdfRedactTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "pdf_redact")]
    [Description(@"Redact (black out) text or area on PDF page. This permanently removes the underlying content.

Usage examples:
- Redact area: pdf_redact(path='doc.pdf', pageIndex=1, x=100, y=100, width=200, height=50)
- Redact with color: pdf_redact(path='doc.pdf', pageIndex=1, x=100, y=100, width=200, height=50, fillColor='255,0,0')
- Redact with overlay: pdf_redact(path='doc.pdf', pageIndex=1, x=100, y=100, width=200, height=50, overlayText='[REDACTED]')
- Redact by text search: pdf_redact(path='doc.pdf', textToRedact='confidential')
- Redact by text on page: pdf_redact(path='doc.pdf', pageIndex=1, textToRedact='secret', caseSensitive=false)")]
    public string Execute(
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Page index (1-based, optional for text search - searches all pages if not specified)")]
        int? pageIndex = null,
        [Description(
            "X position of redaction area in PDF coordinates, origin at bottom-left corner (required for area redaction)")]
        double? x = null,
        [Description(
            "Y position of redaction area in PDF coordinates, origin at bottom-left corner (required for area redaction)")]
        double? y = null,
        [Description("Width of redaction area in PDF points (required for area redaction)")]
        double? width = null,
        [Description("Height of redaction area in PDF points (required for area redaction)")]
        double? height = null,
        [Description("Text to search and redact (alternative to area redaction, searches and redacts all occurrences)")]
        string? textToRedact = null,
        [Description("Whether text search is case sensitive (default: true, only for textToRedact)")]
        bool caseSensitive = true,
        [Description("Fill color (optional, default: black, format: 'R,G,B' or color name)")]
        string? fillColor = null,
        [Description("Text to display over the redacted area (optional, e.g., '[REDACTED]')")]
        string? overlayText = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);

        if (!string.IsNullOrEmpty(textToRedact))
            return RedactByText(ctx, outputPath, textToRedact, pageIndex, caseSensitive, fillColor, overlayText);

        return RedactByArea(ctx, outputPath, pageIndex, x, y, width, height, fillColor, overlayText);
    }

    /// <summary>
    ///     Redacts text by searching for occurrences and applying redaction.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="textToRedact">The text to search for and redact.</param>
    /// <param name="pageIndex">Optional 1-based page index to search on a specific page.</param>
    /// <param name="caseSensitive">Whether the search is case-sensitive.</param>
    /// <param name="fillColor">Optional fill color for the redaction.</param>
    /// <param name="overlayText">Optional overlay text to display on the redaction.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the page index is invalid.</exception>
    private static string RedactByText(DocumentContext<Document> ctx, string? outputPath, string textToRedact,
        int? pageIndex, bool caseSensitive, string? fillColor, string? overlayText)
    {
        var document = ctx.Document;

        if (pageIndex.HasValue && (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count))
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        TextFragmentAbsorber absorber;
        if (caseSensitive)
        {
            absorber = new TextFragmentAbsorber(textToRedact);
        }
        else
        {
            var escapedPattern = Regex.Escape(textToRedact);
            var textSearchOptions = new TextSearchOptions(true);
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

        ctx.Save(outputPath);
        var pageInfo = pagesAffected.Count == 1
            ? $"page {pagesAffected.First()}"
            : $"{pagesAffected.Count} pages";
        return
            $"Redacted {redactionCount} occurrence(s) of '{textToRedact}' on {pageInfo}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Redacts a specific rectangular area on a page.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <param name="x">The X position of the redaction area.</param>
    /// <param name="y">The Y position of the redaction area.</param>
    /// <param name="width">The width of the redaction area.</param>
    /// <param name="height">The height of the redaction area.</param>
    /// <param name="fillColor">Optional fill color for the redaction.</param>
    /// <param name="overlayText">Optional overlay text to display on the redaction.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    private static string RedactByArea(DocumentContext<Document> ctx, string? outputPath,
        int? pageIndex, double? x, double? y, double? width, double? height,
        string? fillColor, string? overlayText)
    {
        if (!pageIndex.HasValue)
            throw new ArgumentException("pageIndex is required for area redaction");
        if (!x.HasValue)
            throw new ArgumentException("x is required for area redaction");
        if (!y.HasValue)
            throw new ArgumentException("y is required for area redaction");
        if (!width.HasValue)
            throw new ArgumentException("width is required for area redaction");
        if (!height.HasValue)
            throw new ArgumentException("height is required for area redaction");

        var document = ctx.Document;

        if (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex.Value];
        var rect = new Rectangle(x.Value, y.Value, x.Value + width.Value, y.Value + height.Value);

        var redactionAnnotation = new RedactionAnnotation(page, rect);
        ApplyRedactionStyle(redactionAnnotation, fillColor, overlayText);

        page.Annotations.Add(redactionAnnotation);
        redactionAnnotation.Redact();

        ctx.Save(outputPath);

        return $"Redaction applied to page {pageIndex.Value}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Applies styling options to a redaction annotation.
    /// </summary>
    /// <param name="annotation">The redaction annotation to style.</param>
    /// <param name="fillColor">Optional fill color for the redaction.</param>
    /// <param name="overlayText">Optional overlay text to display on the redaction.</param>
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