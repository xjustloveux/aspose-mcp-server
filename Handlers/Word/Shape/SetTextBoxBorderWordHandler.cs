using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Word.Shape;

/// <summary>
///     Handler for setting textbox border properties in Word documents.
/// </summary>
public class SetTextBoxBorderWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set_textbox_border";

    /// <summary>
    ///     Sets border properties for a textbox.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: textboxIndex
    ///     Optional: borderVisible, borderColor, borderWidth, borderStyle
    /// </param>
    /// <returns>Success message with border details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var textboxIndex = parameters.GetOptional<int?>("textboxIndex");
        var borderVisible = parameters.GetOptional("borderVisible", true);
        var borderColor = parameters.GetOptional<string?>("borderColor");
        var borderWidth = parameters.GetOptional("borderWidth", 1.0);
        var borderStyle = parameters.GetOptional("borderStyle", "solid");

        if (!textboxIndex.HasValue)
            throw new ArgumentException("textboxIndex is required for set_textbox_border operation");

        var doc = context.Document;
        var allTextboxes = WordShapeHelper.FindAllTextboxes(doc);

        if (textboxIndex.Value < 0 || textboxIndex.Value >= allTextboxes.Count)
            throw new ArgumentException(
                $"Textbox index {textboxIndex.Value} out of range (total textboxes: {allTextboxes.Count})");

        var textBox = allTextboxes[textboxIndex.Value];
        var stroke = textBox.Stroke;

        stroke.Visible = borderVisible;

        if (borderVisible)
        {
            stroke.Color = ColorHelper.ParseColor(borderColor ?? "000000");
            stroke.Weight = borderWidth;
            stroke.DashStyle = WordShapeHelper.ParseDashStyle(borderStyle);
        }

        MarkModified(context);

        var borderDesc = borderVisible
            ? $"Border: {borderWidth}pt, Color: {borderColor ?? "000000"}, Style: {borderStyle}"
            : "No border";

        return $"Successfully set textbox {textboxIndex.Value} {borderDesc}.";
    }
}
