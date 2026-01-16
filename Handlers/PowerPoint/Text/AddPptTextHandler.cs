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
        var fontParams = ExtractFontParameters(parameters);
        var shapeParams = ExtractShapeParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        var shape = GetOrCreateShape(slide, shapeIndex, shapeParams);

        if (shape.TextFrame == null)
            throw new ArgumentException("Shape does not support text");

        shape.TextFrame.Text = text;
        ApplyFontFormatting(shape.TextFrame, fontParams);

        MarkModified(context);

        return Success($"Text added to slide {slideIndex}.");
    }

    private static FontParameters ExtractFontParameters(OperationParameters parameters)
    {
        return new FontParameters(
            parameters.GetOptional<string?>("fontName"),
            parameters.GetOptional<float?>("fontSize"),
            parameters.GetOptional<bool?>("bold"),
            parameters.GetOptional<bool?>("italic"),
            parameters.GetOptional<string?>("color")
        );
    }

    private static ShapeParameters ExtractShapeParameters(OperationParameters parameters)
    {
        return new ShapeParameters(
            parameters.GetOptional("x", 100f),
            parameters.GetOptional("y", 100f),
            parameters.GetOptional("width", 400f),
            parameters.GetOptional("height", 100f)
        );
    }

    private static IAutoShape GetOrCreateShape(ISlide slide, int? shapeIndex, ShapeParameters shapeParams)
    {
        if (shapeIndex.HasValue)
        {
            if (slide.Shapes[shapeIndex.Value] is not IAutoShape existingShape)
                throw new ArgumentException($"Shape at index {shapeIndex.Value} is not a text shape");
            return existingShape;
        }

        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, shapeParams.X, shapeParams.Y,
            shapeParams.Width, shapeParams.Height);
        shape.FillFormat.FillType = FillType.NoFill;
        shape.LineFormat.FillFormat.FillType = FillType.NoFill;
        return shape;
    }

    private static void ApplyFontFormatting(ITextFrame textFrame, FontParameters fontParams)
    {
        if (textFrame.Paragraphs.Count == 0) return;

        var portion = textFrame.Paragraphs[0].Portions.Count > 0 ? textFrame.Paragraphs[0].Portions[0] : null;
        if (portion == null) return;

        ApplyFontProperties(portion, fontParams);
    }

    private static void ApplyFontProperties(IPortion portion, FontParameters fontParams)
    {
        if (!string.IsNullOrEmpty(fontParams.FontName))
            portion.PortionFormat.LatinFont = new FontData(fontParams.FontName);

        if (fontParams.FontSize.HasValue)
            portion.PortionFormat.FontHeight = fontParams.FontSize.Value;

        if (fontParams.Bold.HasValue)
            portion.PortionFormat.FontBold = fontParams.Bold.Value ? NullableBool.True : NullableBool.False;

        if (fontParams.Italic.HasValue)
            portion.PortionFormat.FontItalic = fontParams.Italic.Value ? NullableBool.True : NullableBool.False;

        if (!string.IsNullOrEmpty(fontParams.Color))
        {
            portion.PortionFormat.FillFormat.FillType = FillType.Solid;
            portion.PortionFormat.FillFormat.SolidFillColor.Color = ColorHelper.ParseColor(fontParams.Color);
        }
    }

    private record FontParameters(string? FontName, float? FontSize, bool? Bold, bool? Italic, string? Color);

    private record ShapeParameters(float X, float Y, float Width, float Height);
}
