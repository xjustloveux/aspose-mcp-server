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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var text = parameters.GetOptional<string?>("text");
        var fontFamily = parameters.GetOptional("fontFamily", "Arial");
        var fontSize = parameters.GetOptional("fontSize", 72.0);
        var isSemitransparent = parameters.GetOptional("isSemitransparent", true);
        var layout = parameters.GetOptional("layout", "Diagonal");

        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("Text is required for add operation");

        var doc = context.Document;

        var watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = fontFamily,
            FontSize = (float)fontSize,
            IsSemitrasparent = isSemitransparent,
            Layout = string.Equals(layout, "horizontal", StringComparison.OrdinalIgnoreCase)
                ? WatermarkLayout.Horizontal
                : WatermarkLayout.Diagonal
        };

        doc.Watermark.SetText(text, watermarkOptions);

        MarkModified(context);

        return Success("Text watermark added to document");
    }
}
