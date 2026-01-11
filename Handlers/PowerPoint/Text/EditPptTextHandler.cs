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
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var shapeIndex = parameters.GetRequired<int>("shapeIndex");
        var text = parameters.GetOptional<string?>("text");
        var fontName = parameters.GetOptional<string?>("fontName");
        var fontSize = parameters.GetOptional<float?>("fontSize");
        var bold = parameters.GetOptional<bool?>("bold");
        var italic = parameters.GetOptional<bool?>("italic");
        var color = parameters.GetOptional<string?>("color");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        if (shapeIndex < 0 || shapeIndex >= slide.Shapes.Count)
            throw new ArgumentException(
                $"shapeIndex must be between 0 and {slide.Shapes.Count - 1}, got: {shapeIndex}");

        if (slide.Shapes[shapeIndex] is not IAutoShape shape)
            throw new ArgumentException($"Shape at index {shapeIndex} is not a text shape");

        if (shape.TextFrame == null)
            throw new ArgumentException("Shape does not support text");

        if (!string.IsNullOrEmpty(text))
            shape.TextFrame.Text = text;

        if (shape.TextFrame.Paragraphs.Count > 0)
            foreach (var paragraph in shape.TextFrame.Paragraphs)
            foreach (var portion in paragraph.Portions)
            {
                if (!string.IsNullOrEmpty(fontName))
                    portion.PortionFormat.LatinFont = new FontData(fontName);

                if (fontSize.HasValue)
                    portion.PortionFormat.FontHeight = fontSize.Value;

                if (bold.HasValue)
                    portion.PortionFormat.FontBold = bold.Value ? NullableBool.True : NullableBool.False;

                if (italic.HasValue)
                    portion.PortionFormat.FontItalic = italic.Value ? NullableBool.True : NullableBool.False;

                if (!string.IsNullOrEmpty(color))
                {
                    portion.PortionFormat.FillFormat.FillType = FillType.Solid;
                    portion.PortionFormat.FillFormat.SolidFillColor.Color = ColorHelper.ParseColor(color);
                }
            }

        MarkModified(context);

        return Success($"Text edited in shape {shapeIndex} on slide {slideIndex}.");
    }
}
