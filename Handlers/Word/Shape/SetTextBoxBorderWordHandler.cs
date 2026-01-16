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
    /// <exception cref="ArgumentException">Thrown when textboxIndex is missing or out of range.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractSetTextBoxBorderParameters(parameters);

        var doc = context.Document;
        var allTextboxes = WordShapeHelper.FindAllTextboxes(doc);

        if (p.TextboxIndex < 0 || p.TextboxIndex >= allTextboxes.Count)
            throw new ArgumentException(
                $"Textbox index {p.TextboxIndex} out of range (total textboxes: {allTextboxes.Count})");

        var textBox = allTextboxes[p.TextboxIndex];
        var stroke = textBox.Stroke;

        stroke.Visible = p.BorderVisible;

        if (p.BorderVisible)
        {
            stroke.Color = ColorHelper.ParseColor(p.BorderColor ?? "000000");
            stroke.Weight = p.BorderWidth;
            stroke.DashStyle = WordShapeHelper.ParseDashStyle(p.BorderStyle);
        }

        MarkModified(context);

        var borderDesc = p.BorderVisible
            ? $"Border: {p.BorderWidth}pt, Color: {p.BorderColor ?? "000000"}, Style: {p.BorderStyle}"
            : "No border";

        return $"Successfully set textbox {p.TextboxIndex} {borderDesc}.";
    }

    private static SetTextBoxBorderParameters ExtractSetTextBoxBorderParameters(OperationParameters parameters)
    {
        var textboxIndex = parameters.GetOptional<int?>("textboxIndex");
        var borderVisible = parameters.GetOptional("borderVisible", true);
        var borderColor = parameters.GetOptional<string?>("borderColor");
        var borderWidth = parameters.GetOptional("borderWidth", 1.0);
        var borderStyle = parameters.GetOptional("borderStyle", "solid");

        if (!textboxIndex.HasValue)
            throw new ArgumentException("textboxIndex is required for set_textbox_border operation");

        return new SetTextBoxBorderParameters(
            textboxIndex.Value,
            borderVisible,
            borderColor,
            borderWidth,
            borderStyle);
    }

    private sealed record SetTextBoxBorderParameters(
        int TextboxIndex,
        bool BorderVisible,
        string? BorderColor,
        double BorderWidth,
        string BorderStyle);
}
