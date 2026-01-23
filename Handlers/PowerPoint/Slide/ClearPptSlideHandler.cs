using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Slide;

/// <summary>
///     Handler for clearing slide content in PowerPoint presentations.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class ClearPptSlideHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "clear";

    /// <summary>
    ///     Clears all content from a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex (0-based)
    /// </param>
    /// <returns>Success message with clear details.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractClearPptSlideParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

        while (slide.Shapes.Count > 0)
            slide.Shapes.RemoveAt(0);

        MarkModified(context);

        return new SuccessResult { Message = $"Cleared all shapes from slide {p.SlideIndex}." };
    }

    /// <summary>
    ///     Extracts parameters for clear slide operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static ClearPptSlideParameters ExtractClearPptSlideParameters(OperationParameters parameters)
    {
        return new ClearPptSlideParameters(parameters.GetRequired<int>("slideIndex"));
    }

    /// <summary>
    ///     Parameters for clear slide operation.
    /// </summary>
    /// <param name="SlideIndex">The slide index (0-based).</param>
    private sealed record ClearPptSlideParameters(int SlideIndex);
}
