using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.PageSetup;

/// <summary>
///     Handler for setting slide size in PowerPoint presentations.
/// </summary>
public class SetSlideSizeHandler : OperationHandlerBase<Presentation>
{
    private const float MinSizePoints = 1f;
    private const float MaxSizePoints = 5000f;

    /// <inheritdoc />
    public override string Operation => "set_size";

    /// <summary>
    ///     Sets the slide size for the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Optional: preset (OnScreen16x9, OnScreen16x10, Letter, A4, Banner, Custom),
    ///     width, height (required when preset=Custom)
    /// </param>
    /// <returns>Success message with slide size information.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var preset = parameters.GetOptional("preset", "OnScreen16x9");
        var width = parameters.GetOptional<double?>("width");
        var height = parameters.GetOptional<double?>("height");

        var presentation = context.Document;
        var slideSize = presentation.SlideSize;

        var type = preset.ToLower() switch
        {
            "onscreen16x10" => SlideSizeType.OnScreen16x10,
            "a4" => SlideSizeType.A4Paper,
            "banner" => SlideSizeType.Banner,
            "custom" => SlideSizeType.Custom,
            _ => SlideSizeType.OnScreen
        };

        if (type == SlideSizeType.Custom)
        {
            if (!width.HasValue || !height.HasValue)
                throw new ArgumentException("Custom size requires width and height.");

            ValidateSizeRange(width.Value, height.Value);
            slideSize.SetSize((float)width.Value, (float)height.Value, SlideSizeScaleType.DoNotScale);
        }
        else
        {
            slideSize.SetSize(type, SlideSizeScaleType.DoNotScale);
        }

        MarkModified(context);

        var sizeInfo = slideSize.Type == SlideSizeType.Custom
            ? $" ({slideSize.Size.Width}x{slideSize.Size.Height})"
            : "";

        return Success($"Slide size set to {slideSize.Type}{sizeInfo}.");
    }

    /// <summary>
    ///     Validates that the width and height values are within acceptable range.
    /// </summary>
    /// <param name="width">The width value in points.</param>
    /// <param name="height">The height value in points.</param>
    /// <exception cref="ArgumentException">Thrown when width or height is outside the valid range.</exception>
    private static void ValidateSizeRange(double width, double height)
    {
        if (width < MinSizePoints || width > MaxSizePoints)
            throw new ArgumentException($"Width must be between {MinSizePoints} and {MaxSizePoints} points.");

        if (height < MinSizePoints || height > MaxSizePoints)
            throw new ArgumentException($"Height must be between {MinSizePoints} and {MaxSizePoints} points.");
    }
}
