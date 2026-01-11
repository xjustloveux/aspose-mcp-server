using Aspose.Slides;
using Aspose.Slides.SmartArt;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.SmartArt;

/// <summary>
///     Handler for adding SmartArt shapes to PowerPoint slides.
/// </summary>
public class AddSmartArtHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a SmartArt shape to a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex, layout.
    ///     Optional: x, y, width, height.
    /// </param>
    /// <returns>Success message with SmartArt creation details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var layoutStr = parameters.GetRequired<string>("layout");
        var x = parameters.GetOptional("x", 100f);
        var y = parameters.GetOptional("y", 100f);
        var width = parameters.GetOptional("width", 400f);
        var height = parameters.GetOptional("height", 300f);

        if (!Enum.TryParse<SmartArtLayoutType>(layoutStr, true, out var layoutType))
            throw new ArgumentException(
                $"Invalid SmartArt layout: '{layoutStr}'. Valid layouts include: BasicProcess, Cycle, Hierarchy, etc.");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        slide.Shapes.AddSmartArt(x, y, width, height, layoutType);

        MarkModified(context);

        return Success($"SmartArt '{layoutStr}' added to slide {slideIndex}.");
    }
}
