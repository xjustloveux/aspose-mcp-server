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
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var insertAt = parameters.GetOptional<int?>("insertAt");
        var presentation = context.Document;
        var count = presentation.Slides.Count;

        if (slideIndex < 0 || slideIndex >= count)
            throw new ArgumentException($"slideIndex must be between 0 and {count - 1}");

        if (insertAt.HasValue)
        {
            if (insertAt.Value < 0 || insertAt.Value > count)
                throw new ArgumentException($"insertAt must be between 0 and {count}");

            presentation.Slides.InsertClone(insertAt.Value, presentation.Slides[slideIndex]);
        }
        else
        {
            presentation.Slides.AddClone(presentation.Slides[slideIndex]);
        }

        MarkModified(context);

        return Success($"Slide {slideIndex} duplicated (total: {presentation.Slides.Count}).");
    }
}
