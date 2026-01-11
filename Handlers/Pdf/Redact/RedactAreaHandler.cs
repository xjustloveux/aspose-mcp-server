using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using Color = System.Drawing.Color;

namespace AsposeMcpServer.Handlers.Pdf.Redact;

/// <summary>
///     Handler for redacting a specific rectangular area on a PDF page.
/// </summary>
public class RedactAreaHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "area";

    /// <summary>
    ///     Redacts a specific rectangular area on a page.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: pageIndex (1-based), x, y, width, height
    ///     Optional: fillColor, overlayText
    /// </param>
    /// <returns>Success message with redaction details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var pageIndex = parameters.GetRequired<int>("pageIndex");
        var x = parameters.GetRequired<double>("x");
        var y = parameters.GetRequired<double>("y");
        var width = parameters.GetRequired<double>("width");
        var height = parameters.GetRequired<double>("height");
        var fillColor = parameters.GetOptional<string?>("fillColor");
        var overlayText = parameters.GetOptional<string?>("overlayText");

        var document = context.Document;

        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var rect = new Rectangle(x, y, x + width, y + height);

        var redactionAnnotation = new RedactionAnnotation(page, rect);
        ApplyRedactionStyle(redactionAnnotation, fillColor, overlayText);

        page.Annotations.Add(redactionAnnotation);
        redactionAnnotation.Redact();

        MarkModified(context);

        return Success($"Redaction applied to page {pageIndex}.");
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
