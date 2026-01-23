using System.Text.RegularExpressions;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;
using Color = System.Drawing.Color;

namespace AsposeMcpServer.Handlers.Pdf.Redact;

/// <summary>
///     Handler for redacting text by searching for occurrences in a PDF document.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractTextParameters(parameters);

        var document = context.Document;

        if (p.PageIndex.HasValue && (p.PageIndex.Value < 1 || p.PageIndex.Value > document.Pages.Count))
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        TextFragmentAbsorber absorber;
        if (p.CaseSensitive)
        {
            absorber = new TextFragmentAbsorber(p.TextToRedact);
        }
        else
        {
            var escapedPattern = Regex.Escape(p.TextToRedact);
            var textSearchOptions = new TextSearchOptions(true);
            absorber = new TextFragmentAbsorber($"(?i){escapedPattern}", textSearchOptions);
        }

        var redactionCount = 0;
        var pagesAffected = new HashSet<int>();

        if (p.PageIndex.HasValue)
            document.Pages[p.PageIndex.Value].Accept(absorber);
        else
            document.Pages.Accept(absorber);

        foreach (var fragment in absorber.TextFragments)
        {
            var page = fragment.Page;
            var rect = fragment.Rectangle;

            var redactionAnnotation = new RedactionAnnotation(page, rect);
            ApplyRedactionStyle(redactionAnnotation, p.FillColor, p.OverlayText);

            page.Annotations.Add(redactionAnnotation);
            redactionAnnotation.Redact();

            redactionCount++;
            pagesAffected.Add(document.Pages.IndexOf(page));
        }

        if (redactionCount == 0)
            return new SuccessResult
                { Message = $"No occurrences of '{p.TextToRedact}' found. No redactions applied." };

        MarkModified(context);

        var pageInfo = pagesAffected.Count == 1
            ? $"page {pagesAffected.First()}"
            : $"{pagesAffected.Count} pages";
        return new SuccessResult
            { Message = $"Redacted {redactionCount} occurrence(s) of '{p.TextToRedact}' on {pageInfo}." };
    }

    /// <summary>
    ///     Extracts parameters for text operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static TextParameters ExtractTextParameters(OperationParameters parameters)
    {
        return new TextParameters(
            parameters.GetRequired<string>("textToRedact"),
            parameters.GetOptional<int?>("pageIndex"),
            parameters.GetOptional("caseSensitive", true),
            parameters.GetOptional<string?>("fillColor"),
            parameters.GetOptional<string?>("overlayText")
        );
    }

    /// <summary>
    ///     Applies styling options to a redaction annotation.
    /// </summary>
    /// <param name="annotation">The redaction annotation to style.</param>
    /// <param name="fillColor">The fill color for the redaction area.</param>
    /// <param name="overlayText">The optional text to display over the redacted area.</param>
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

    /// <summary>
    ///     Parameters for text operation.
    /// </summary>
    /// <param name="TextToRedact">The text to search and redact.</param>
    /// <param name="PageIndex">The optional 1-based page index.</param>
    /// <param name="CaseSensitive">Whether to search case-sensitively.</param>
    /// <param name="FillColor">The optional fill color for redaction.</param>
    /// <param name="OverlayText">The optional overlay text.</param>
    private sealed record TextParameters(
        string TextToRedact,
        int? PageIndex,
        bool CaseSensitive,
        string? FillColor,
        string? OverlayText);
}
