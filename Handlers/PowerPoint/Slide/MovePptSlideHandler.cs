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
        var p = ExtractMovePptSlideParameters(parameters);

        var presentation = context.Document;
        var count = presentation.Slides.Count;

        if (p.FromIndex < 0 || p.FromIndex >= count)
            throw new ArgumentException($"fromIndex must be between 0 and {count - 1}");
        if (p.ToIndex < 0 || p.ToIndex >= count)
            throw new ArgumentException($"toIndex must be between 0 and {count - 1}");

        var slide = presentation.Slides[p.FromIndex];
        presentation.Slides.Reorder(p.ToIndex, slide);

        MarkModified(context);

        return Success($"Slide moved from {p.FromIndex} to {p.ToIndex} (total: {count}).");
    }

    /// <summary>
    ///     Extracts parameters for move slide operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static MovePptSlideParameters ExtractMovePptSlideParameters(OperationParameters parameters)
    {
        return new MovePptSlideParameters(
            parameters.GetRequired<int>("fromIndex"),
            parameters.GetRequired<int>("toIndex"));
    }

    /// <summary>
    ///     Parameters for move slide operation.
    /// </summary>
    /// <param name="FromIndex">The source slide index (0-based).</param>
    /// <param name="ToIndex">The target slide index (0-based).</param>
    private record MovePptSlideParameters(int FromIndex, int ToIndex);
}
