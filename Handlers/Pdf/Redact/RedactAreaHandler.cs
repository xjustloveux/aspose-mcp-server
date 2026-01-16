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
        var p = ExtractAreaParameters(parameters);

        var document = context.Document;

        if (p.PageIndex < 1 || p.PageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[p.PageIndex];
        var rect = new Rectangle(p.X, p.Y, p.X + p.Width, p.Y + p.Height);

        var redactionAnnotation = new RedactionAnnotation(page, rect);
        ApplyRedactionStyle(redactionAnnotation, p.FillColor, p.OverlayText);

        page.Annotations.Add(redactionAnnotation);
        redactionAnnotation.Redact();

        MarkModified(context);

        return Success($"Redaction applied to page {p.PageIndex}.");
    }

    /// <summary>
    ///     Extracts parameters for area operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static AreaParameters ExtractAreaParameters(OperationParameters parameters)
    {
        return new AreaParameters(
            parameters.GetRequired<int>("pageIndex"),
            parameters.GetRequired<double>("x"),
            parameters.GetRequired<double>("y"),
            parameters.GetRequired<double>("width"),
            parameters.GetRequired<double>("height"),
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
    ///     Parameters for area operation.
    /// </summary>
    /// <param name="PageIndex">The 1-based page index.</param>
    /// <param name="X">The X coordinate of the redaction area.</param>
    /// <param name="Y">The Y coordinate of the redaction area.</param>
    /// <param name="Width">The width of the redaction area.</param>
    /// <param name="Height">The height of the redaction area.</param>
    /// <param name="FillColor">The optional fill color for the redaction.</param>
    /// <param name="OverlayText">The optional overlay text.</param>
    private sealed record AreaParameters(
        int PageIndex,
        double X,
        double Y,
        double Width,
        double Height,
        string? FillColor,
        string? OverlayText);
}
