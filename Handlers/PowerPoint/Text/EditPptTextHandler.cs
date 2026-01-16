using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Text;

/// <summary>
///     Handler for editing text in PowerPoint presentations.
/// </summary>
public class EditPptTextHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits text in a specific shape.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex.
    ///     Optional: slideIndex (default: 0), text, fontName, fontSize, bold, italic, color.
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var textParams = ExtractTextParameters(parameters);
        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, textParams.SlideIndex);
        var shape = GetTextShape(slide, textParams.ShapeIndex);

        if (!string.IsNullOrEmpty(textParams.Text))
            shape.TextFrame.Text = textParams.Text;

        ApplyFormattingToAllPortions(shape, textParams);

        MarkModified(context);

        return Success($"Text edited in shape {textParams.ShapeIndex} on slide {textParams.SlideIndex}.");
    }

    /// <summary>
    ///     Extracts text parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted text parameters.</returns>
    private static TextParameters ExtractTextParameters(OperationParameters parameters)
    {
        return new TextParameters(
            parameters.GetOptional("slideIndex", 0),
            parameters.GetRequired<int>("shapeIndex"),
            parameters.GetOptional<string?>("text"),
            parameters.GetOptional<string?>("fontName"),
            parameters.GetOptional<float?>("fontSize"),
            parameters.GetOptional<bool?>("bold"),
            parameters.GetOptional<bool?>("italic"),
            parameters.GetOptional<string?>("color")
        );
    }

    /// <summary>
    ///     Gets a text shape from the slide.
    /// </summary>
    /// <param name="slide">The slide.</param>
    /// <param name="shapeIndex">The shape index.</param>
    /// <returns>The AutoShape.</returns>
    private static IAutoShape GetTextShape(ISlide slide, int shapeIndex)
    {
        if (shapeIndex < 0 || shapeIndex >= slide.Shapes.Count)
            throw new ArgumentException(
                $"shapeIndex must be between 0 and {slide.Shapes.Count - 1}, got: {shapeIndex}");

        if (slide.Shapes[shapeIndex] is not IAutoShape shape)
            throw new ArgumentException($"Shape at index {shapeIndex} is not a text shape");

        if (shape.TextFrame == null)
            throw new ArgumentException("Shape does not support text");

        return shape;
    }

    /// <summary>
    ///     Applies formatting to all portions in the shape.
    /// </summary>
    /// <param name="shape">The AutoShape.</param>
    /// <param name="p">The text parameters.</param>
    private static void ApplyFormattingToAllPortions(IAutoShape shape, TextParameters p)
    {
        if (shape.TextFrame.Paragraphs.Count == 0) return;

        foreach (var paragraph in shape.TextFrame.Paragraphs)
        foreach (var portion in paragraph.Portions)
            ApplyPortionFormatting(portion, p);
    }

    /// <summary>
    ///     Applies formatting to a text portion.
    /// </summary>
    /// <param name="portion">The text portion.</param>
    /// <param name="p">The text parameters.</param>
    private static void ApplyPortionFormatting(IPortion portion, TextParameters p)
    {
        if (!string.IsNullOrEmpty(p.FontName))
            portion.PortionFormat.LatinFont = new FontData(p.FontName);

        if (p.FontSize.HasValue)
            portion.PortionFormat.FontHeight = p.FontSize.Value;

        if (p.Bold.HasValue)
            portion.PortionFormat.FontBold = p.Bold.Value ? NullableBool.True : NullableBool.False;

        if (p.Italic.HasValue)
            portion.PortionFormat.FontItalic = p.Italic.Value ? NullableBool.True : NullableBool.False;

        if (!string.IsNullOrEmpty(p.Color))
        {
            portion.PortionFormat.FillFormat.FillType = FillType.Solid;
            portion.PortionFormat.FillFormat.SolidFillColor.Color = ColorHelper.ParseColor(p.Color);
        }
    }

    /// <summary>
    ///     Record for holding text editing parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ShapeIndex">The shape index.</param>
    /// <param name="Text">The optional text content.</param>
    /// <param name="FontName">The optional font name.</param>
    /// <param name="FontSize">The optional font size.</param>
    /// <param name="Bold">The optional bold setting.</param>
    /// <param name="Italic">The optional italic setting.</param>
    /// <param name="Color">The optional text color.</param>
    private sealed record TextParameters(
        int SlideIndex,
        int ShapeIndex,
        string? Text,
        string? FontName,
        float? FontSize,
        bool? Bold,
        bool? Italic,
        string? Color);
}
