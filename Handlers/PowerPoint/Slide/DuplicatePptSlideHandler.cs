using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Slide;

/// <summary>
///     Handler for duplicating slides in PowerPoint presentations.
/// </summary>
public class DuplicatePptSlideHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "duplicate";

    /// <summary>
    ///     Duplicates a slide in the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex (0-based)
    ///     Optional: insertAt (0-based target index, default: append)
    /// </param>
    /// <returns>Success message with duplicate details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractDuplicatePptSlideParameters(parameters);

        var presentation = context.Document;
        var count = presentation.Slides.Count;

        if (p.SlideIndex < 0 || p.SlideIndex >= count)
            throw new ArgumentException($"slideIndex must be between 0 and {count - 1}");

        if (p.InsertAt.HasValue)
        {
            if (p.InsertAt.Value < 0 || p.InsertAt.Value > count)
                throw new ArgumentException($"insertAt must be between 0 and {count}");

            presentation.Slides.InsertClone(p.InsertAt.Value, presentation.Slides[p.SlideIndex]);
        }
        else
        {
            presentation.Slides.AddClone(presentation.Slides[p.SlideIndex]);
        }

        MarkModified(context);

        return Success($"Slide {p.SlideIndex} duplicated (total: {presentation.Slides.Count}).");
    }

    /// <summary>
    ///     Extracts parameters for duplicate slide operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static DuplicatePptSlideParameters ExtractDuplicatePptSlideParameters(OperationParameters parameters)
    {
        return new DuplicatePptSlideParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetOptional<int?>("insertAt"));
    }

    /// <summary>
    ///     Parameters for duplicate slide operation.
    /// </summary>
    /// <param name="SlideIndex">The slide index (0-based).</param>
    /// <param name="InsertAt">The target index to insert the duplicate.</param>
    private sealed record DuplicatePptSlideParameters(int SlideIndex, int? InsertAt);
}
