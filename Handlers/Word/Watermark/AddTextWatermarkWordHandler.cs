using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Watermark;

/// <summary>
///     Handler for adding text watermarks to Word documents.
/// </summary>
public class AddTextWatermarkWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a text watermark to the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: text
    ///     Optional: fontFamily (default: Arial), fontSize (default: 72), isSemitransparent (default: true), layout (default:
    ///     Diagonal)
    /// </param>
    /// <returns>Success message with watermark details.</returns>
    /// <exception cref="ArgumentException">Thrown when text is missing.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddTextWatermarkParameters(parameters);

        if (string.IsNullOrEmpty(p.Text))
            throw new ArgumentException("Text is required for add operation");

        var doc = context.Document;

        var watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = p.FontFamily,
            FontSize = (float)p.FontSize,
            IsSemitrasparent = p.IsSemitransparent,
            Layout = string.Equals(p.Layout, "horizontal", StringComparison.OrdinalIgnoreCase)
                ? WatermarkLayout.Horizontal
                : WatermarkLayout.Diagonal
        };

        doc.Watermark.SetText(p.Text, watermarkOptions);

        MarkModified(context);

        return Success("Text watermark added to document");
    }

    private static AddTextWatermarkParameters ExtractAddTextWatermarkParameters(OperationParameters parameters)
    {
        return new AddTextWatermarkParameters(
            parameters.GetOptional<string?>("text"),
            parameters.GetOptional("fontFamily", "Arial"),
            parameters.GetOptional("fontSize", 72.0),
            parameters.GetOptional("isSemitransparent", true),
            parameters.GetOptional("layout", "Diagonal"));
    }

    private record AddTextWatermarkParameters(
        string? Text,
        string FontFamily,
        double FontSize,
        bool IsSemitransparent,
        string Layout);
}
