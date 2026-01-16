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
        var smartArtParams = ExtractSmartArtParameters(parameters);

        if (!Enum.TryParse<SmartArtLayoutType>(smartArtParams.Layout, true, out var layoutType))
            throw new ArgumentException(
                $"Invalid SmartArt layout: '{smartArtParams.Layout}'. Valid layouts include: BasicProcess, Cycle, Hierarchy, etc.");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, smartArtParams.SlideIndex);

        slide.Shapes.AddSmartArt(smartArtParams.X, smartArtParams.Y, smartArtParams.Width, smartArtParams.Height,
            layoutType);

        MarkModified(context);

        return Success($"SmartArt '{smartArtParams.Layout}' added to slide {smartArtParams.SlideIndex}.");
    }

    /// <summary>
    ///     Extracts SmartArt parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted SmartArt parameters.</returns>
    private static SmartArtParameters ExtractSmartArtParameters(OperationParameters parameters)
    {
        return new SmartArtParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetRequired<string>("layout"),
            parameters.GetOptional("x", 100f),
            parameters.GetOptional("y", 100f),
            parameters.GetOptional("width", 400f),
            parameters.GetOptional("height", 300f)
        );
    }

    /// <summary>
    ///     Record for holding SmartArt creation parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="Layout">The SmartArt layout type.</param>
    /// <param name="X">The X position.</param>
    /// <param name="Y">The Y position.</param>
    /// <param name="Width">The width.</param>
    /// <param name="Height">The height.</param>
    private sealed record SmartArtParameters(
        int SlideIndex,
        string Layout,
        float X,
        float Y,
        float Width,
        float Height);
}
