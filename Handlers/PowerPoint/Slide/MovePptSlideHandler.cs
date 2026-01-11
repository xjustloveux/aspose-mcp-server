using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Slide;

/// <summary>
///     Handler for moving slides in PowerPoint presentations.
/// </summary>
public class MovePptSlideHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "move";

    /// <summary>
    ///     Moves a slide to a different position using Reorder method.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: fromIndex (0-based source index), toIndex (0-based target index)
    /// </param>
    /// <returns>Success message with move details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var fromIndex = parameters.GetRequired<int>("fromIndex");
        var toIndex = parameters.GetRequired<int>("toIndex");
        var presentation = context.Document;
        var count = presentation.Slides.Count;

        if (fromIndex < 0 || fromIndex >= count)
            throw new ArgumentException($"fromIndex must be between 0 and {count - 1}");
        if (toIndex < 0 || toIndex >= count)
            throw new ArgumentException($"toIndex must be between 0 and {count - 1}");

        var slide = presentation.Slides[fromIndex];
        presentation.Slides.Reorder(toIndex, slide);

        MarkModified(context);

        return Success($"Slide moved from {fromIndex} to {toIndex} (total: {count}).");
    }
}
