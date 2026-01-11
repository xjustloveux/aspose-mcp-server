using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Pdf.Watermark;

/// <summary>
///     Handler for adding watermarks to PDF documents.
/// </summary>
public class AddPdfWatermarkHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a watermark to the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: text.
    ///     Optional: opacity (default: 0.3), fontSize (default: 72), fontName (default: Arial),
    ///     rotation (default: 45), color (default: Gray), pageRange, isBackground (default: false),
    ///     horizontalAlignment (default: Center), verticalAlignment (default: Center).
    /// </param>
    /// <returns>Success message with the number of pages watermarked.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var text = parameters.GetOptional<string?>("text");
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text is required for add operation");

        var opacity = parameters.GetOptional("opacity", 0.3);
        var fontSize = parameters.GetOptional("fontSize", 72.0);
        var fontName = parameters.GetOptional("fontName", "Arial");
        var rotation = parameters.GetOptional("rotation", 45.0);
        var colorName = parameters.GetOptional("color", "Gray");
        var pageRange = parameters.GetOptional<string?>("pageRange");
        var isBackground = parameters.GetOptional("isBackground", false);
        var horizontalAlignment = parameters.GetOptional("horizontalAlignment", "Center");
        var verticalAlignment = parameters.GetOptional("verticalAlignment", "Center");

        var document = context.Document;

        var hAlign = horizontalAlignment.ToLower() switch
        {
            "left" => HorizontalAlignment.Left,
            "right" => HorizontalAlignment.Right,
            _ => HorizontalAlignment.Center
        };

        var vAlign = verticalAlignment.ToLower() switch
        {
            "top" => VerticalAlignment.Top,
            "bottom" => VerticalAlignment.Bottom,
            _ => VerticalAlignment.Center
        };

        var watermarkColor = PdfWatermarkHelper.ParseColor(colorName);
        var pageIndices = PdfWatermarkHelper.ParsePageRange(pageRange, document.Pages.Count);
        var appliedCount = 0;

        foreach (var pageIndex in pageIndices)
        {
            var page = document.Pages[pageIndex];
            var watermark = new WatermarkArtifact();
            var textState = new TextState
            {
                ForegroundColor = watermarkColor
            };

            FontHelper.Pdf.ApplyFontSettings(textState, fontName, fontSize);

            watermark.SetTextAndState(text, textState);
            watermark.ArtifactHorizontalAlignment = hAlign;
            watermark.ArtifactVerticalAlignment = vAlign;
            watermark.Rotation = rotation;
            watermark.Opacity = opacity;
            watermark.IsBackground = isBackground;

            page.Artifacts.Add(watermark);
            appliedCount++;
        }

        MarkModified(context);

        return Success($"Watermark added to {appliedCount} page(s).");
    }
}
