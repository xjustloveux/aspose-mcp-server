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
        var p = ExtractAddParameters(parameters);

        if (string.IsNullOrEmpty(p.Text))
            throw new ArgumentException("text is required for add operation");

        var document = context.Document;

        var hAlign = p.HorizontalAlignment.ToLower() switch
        {
            "left" => HorizontalAlignment.Left,
            "right" => HorizontalAlignment.Right,
            _ => HorizontalAlignment.Center
        };

        var vAlign = p.VerticalAlignment.ToLower() switch
        {
            "top" => VerticalAlignment.Top,
            "bottom" => VerticalAlignment.Bottom,
            _ => VerticalAlignment.Center
        };

        var watermarkColor = PdfWatermarkHelper.ParseColor(p.ColorName);
        var pageIndices = PdfWatermarkHelper.ParsePageRange(p.PageRange, document.Pages.Count);
        var appliedCount = 0;

        foreach (var pageIndex in pageIndices)
        {
            var page = document.Pages[pageIndex];
            var watermark = new WatermarkArtifact();
            var textState = new TextState
            {
                ForegroundColor = watermarkColor
            };

            FontHelper.Pdf.ApplyFontSettings(textState, p.FontName, p.FontSize);

            watermark.SetTextAndState(p.Text, textState);
            watermark.ArtifactHorizontalAlignment = hAlign;
            watermark.ArtifactVerticalAlignment = vAlign;
            watermark.Rotation = p.Rotation;
            watermark.Opacity = p.Opacity;
            watermark.IsBackground = p.IsBackground;

            page.Artifacts.Add(watermark);
            appliedCount++;
        }

        MarkModified(context);

        return Success($"Watermark added to {appliedCount} page(s).");
    }

    /// <summary>
    ///     Extracts parameters for add operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static AddParameters ExtractAddParameters(OperationParameters parameters)
    {
        return new AddParameters(
            parameters.GetOptional<string?>("text"),
            parameters.GetOptional("opacity", 0.3),
            parameters.GetOptional("fontSize", 72.0),
            parameters.GetOptional("fontName", "Arial"),
            parameters.GetOptional("rotation", 45.0),
            parameters.GetOptional("color", "Gray"),
            parameters.GetOptional<string?>("pageRange"),
            parameters.GetOptional("isBackground", false),
            parameters.GetOptional("horizontalAlignment", "Center"),
            parameters.GetOptional("verticalAlignment", "Center")
        );
    }

    /// <summary>
    ///     Parameters for add operation.
    /// </summary>
    /// <param name="Text">The watermark text.</param>
    /// <param name="Opacity">The opacity value.</param>
    /// <param name="FontSize">The font size.</param>
    /// <param name="FontName">The font name.</param>
    /// <param name="Rotation">The rotation angle.</param>
    /// <param name="ColorName">The color name.</param>
    /// <param name="PageRange">The optional page range.</param>
    /// <param name="IsBackground">Whether to place as background.</param>
    /// <param name="HorizontalAlignment">The horizontal alignment.</param>
    /// <param name="VerticalAlignment">The vertical alignment.</param>
    private sealed record AddParameters(
        string? Text,
        double Opacity,
        double FontSize,
        string FontName,
        double Rotation,
        string ColorName,
        string? PageRange,
        bool IsBackground,
        string HorizontalAlignment,
        string VerticalAlignment);
}
