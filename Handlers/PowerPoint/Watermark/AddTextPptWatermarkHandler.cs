using System.Drawing;
using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Watermark;

/// <summary>
///     Handler for adding a text watermark to a PowerPoint presentation.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddTextPptWatermarkHandler : OperationHandlerBase<Presentation>
{
    /// <summary>
    ///     The prefix for watermark shape names.
    /// </summary>
    internal const string WatermarkPrefix = "WM_ASPOSE_";

    /// <inheritdoc />
    public override string Operation => "add_text";

    /// <summary>
    ///     Adds a text watermark to all slides.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: text.
    ///     Optional: fontSize (default: 48), fontColor (default: gray), opacity (default: 128), rotation (default: -45).
    /// </param>
    /// <returns>Success message with watermark details.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var text = parameters.GetRequired<string>("text");
        var fontSize = parameters.GetOptional("fontSize", 48f);
        var fontColor = parameters.GetOptional("fontColor", "128,128,128");
        var opacity = parameters.GetOptional("opacity", 128);
        var rotation = parameters.GetOptional("rotation", -45f);

        var presentation = context.Document;
        var color = ColorHelper.ParseColor(fontColor);
        var alphaColor = Color.FromArgb(Math.Clamp(opacity, 0, 255), color.R, color.G, color.B);

        var slideCount = 0;
        foreach (var slide in presentation.Slides)
        {
            AddTextWatermarkToSlide(slide, text, fontSize, alphaColor, rotation);
            slideCount++;
        }

        MarkModified(context);
        return new SuccessResult { Message = $"Text watermark '{text}' added to {slideCount} slide(s)." };
    }

    /// <summary>
    ///     Adds a text watermark shape to a single slide.
    /// </summary>
    /// <param name="slide">The slide to add the watermark to.</param>
    /// <param name="text">The watermark text.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="color">The font color with alpha transparency.</param>
    /// <param name="rotation">The rotation angle in degrees.</param>
    internal static void AddTextWatermarkToSlide(ISlide slide, string text, float fontSize, Color color, float rotation)
    {
        var slideWidth = slide.Presentation.SlideSize.Size.Width;
        var slideHeight = slide.Presentation.SlideSize.Size.Height;

        var wmWidth = slideWidth * 0.8f;
        var wmHeight = fontSize * 2f;
        var x = (slideWidth - wmWidth) / 2f;
        var y = (slideHeight - wmHeight) / 2f;

        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, x, y, wmWidth, wmHeight);
        shape.Name = $"{WatermarkPrefix}TEXT_{Guid.NewGuid():N}";
        shape.FillFormat.FillType = FillType.NoFill;
        shape.LineFormat.FillFormat.FillType = FillType.NoFill;
        shape.Rotation = rotation;

        shape.TextFrame.Text = text;
        shape.TextFrame.Paragraphs[0].ParagraphFormat.Alignment = TextAlignment.Center;

        var portion = shape.TextFrame.Paragraphs[0].Portions[0];
        portion.PortionFormat.FontHeight = fontSize;
        portion.PortionFormat.FillFormat.FillType = FillType.Solid;
        portion.PortionFormat.FillFormat.SolidFillColor.Color = color;

        shape.ShapeLock.SelectLocked = true;
        shape.ShapeLock.SizeLocked = true;
        shape.ShapeLock.PositionLocked = true;
    }
}
