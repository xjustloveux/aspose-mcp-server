using System.Text.RegularExpressions;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using Color = System.Drawing.Color;

namespace AsposeMcpServer.Handlers.Pdf.Redact;

/// <summary>
///     Handler for redacting text by searching for occurrences in a PDF document.
/// </summary>
public class RedactTextHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "text";

    /// <summary>
    ///     Redacts text by searching for occurrences and applying redaction.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: textToRedact
    ///     Optional: pageIndex (1-based), caseSensitive (default: true), fillColor, overlayText
    /// </param>
    /// <returns>Success message with redaction details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var textToRedact = parameters.GetRequired<string>("textToRedact");
        var pageIndex = parameters.GetOptional<int?>("pageIndex");
        var caseSensitive = parameters.GetOptional("caseSensitive", true);
        var fillColor = parameters.GetOptional<string?>("fillColor");
        var overlayText = parameters.GetOptional<string?>("overlayText");

        var document = context.Document;

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
            return Success($"No occurrences of '{textToRedact}' found. No redactions applied.");

        MarkModified(context);

        var pageInfo = pagesAffected.Count == 1
            ? $"page {pagesAffected.First()}"
            : $"{pagesAffected.Count} pages";
        return Success($"Redacted {redactionCount} occurrence(s) of '{textToRedact}' on {pageInfo}.");
    }

    /// <summary>
    ///     Applies styling options to a redaction annotation.
    /// </summary>
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
