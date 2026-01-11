using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Text;

/// <summary>
///     Handler for adding text to PowerPoint presentations.
/// </summary>
public class AddPptTextHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds text to a shape in the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: text.
    ///     Optional: slideIndex (default: 0), shapeIndex, x, y, width, height, fontName, fontSize, bold, italic, color.
    /// </param>
    /// <returns>Success message with text addition details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var text = parameters.GetRequired<string>("text");
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var shapeIndex = parameters.GetOptional<int?>("shapeIndex");
        var x = parameters.GetOptional("x", 100f);
        var y = parameters.GetOptional("y", 100f);
        var width = parameters.GetOptional("width", 400f);
        var height = parameters.GetOptional("height", 100f);
        var fontName = parameters.GetOptional<string?>("fontName");
        var fontSize = parameters.GetOptional<float?>("fontSize");
        var bold = parameters.GetOptional<bool?>("bold");
        var italic = parameters.GetOptional<bool?>("italic");
        var color = parameters.GetOptional<string?>("color");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        IAutoShape shape;
        if (shapeIndex.HasValue)
        {
            if (slide.Shapes[shapeIndex.Value] is not IAutoShape existingShape)
                throw new ArgumentException($"Shape at index {shapeIndex.Value} is not a text shape");
            shape = existingShape;
        }
        else
        {
            shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, x, y, width, height);
            shape.FillFormat.FillType = FillType.NoFill;
            shape.LineFormat.FillFormat.FillType = FillType.NoFill;
        }

        if (shape.TextFrame == null)
            throw new ArgumentException("Shape does not support text");

        shape.TextFrame.Text = text;

        if (shape.TextFrame.Paragraphs.Count > 0)
        {
            var portion = shape.TextFrame.Paragraphs[0].Portions.Count > 0
                ? shape.TextFrame.Paragraphs[0].Portions[0]
                : null;

            if (portion != null)
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
        }

        MarkModified(context);

        return Success($"Text added to slide {slideIndex}.");
    }
}
