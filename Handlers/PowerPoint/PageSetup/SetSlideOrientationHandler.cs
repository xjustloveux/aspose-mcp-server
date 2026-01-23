using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.PageSetup;

/// <summary>
///     Handler for setting slide orientation in PowerPoint presentations.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetSlideOrientationHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "set_orientation";

    /// <summary>
    ///     Sets the slide orientation by swapping width and height while preserving the aspect ratio.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: orientation (Portrait or Landscape)
    /// </param>
    /// <returns>Success message with orientation information.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractSetSlideOrientationParameters(parameters);
        var isPortrait = p.Orientation.Equals("Portrait", StringComparison.OrdinalIgnoreCase);
        var presentation = context.Document;
        var currentSize = presentation.SlideSize.Size;
        var currentWidth = currentSize.Width;
        var currentHeight = currentSize.Height;

        var needsSwap = isPortrait ? currentWidth > currentHeight : currentHeight > currentWidth;

        if (needsSwap)
            presentation.SlideSize.SetSize(currentHeight, currentWidth, SlideSizeScaleType.EnsureFit);

        MarkModified(context);

        var finalSize = presentation.SlideSize.Size;
        return new SuccessResult
            { Message = $"Slide orientation set to {p.Orientation} ({finalSize.Width}x{finalSize.Height})." };
    }

    private static SetSlideOrientationParameters ExtractSetSlideOrientationParameters(OperationParameters parameters)
    {
        return new SetSlideOrientationParameters(
            parameters.GetRequired<string>("orientation"));
    }

    private sealed record SetSlideOrientationParameters(string Orientation);
}
