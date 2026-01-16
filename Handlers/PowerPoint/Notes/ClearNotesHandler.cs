using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Notes;

/// <summary>
///     Handler for clearing notes from PowerPoint slides.
/// </summary>
public class ClearNotesHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "clear";

    /// <summary>
    ///     Clears notes from slides. Only clears if notes slide exists.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Optional: slideIndices (if not provided, clears all slides)
    /// </param>
    /// <returns>Success message indicating how many slides were cleared.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractClearNotesParameters(parameters);

        var presentation = context.Document;
        var targets = p.SlideIndices?.Length > 0
            ? p.SlideIndices
            : Enumerable.Range(0, presentation.Slides.Count).ToArray();

        ValidateSlideIndices(targets, presentation.Slides.Count);

        var clearedCount = 0;
        foreach (var idx in targets)
        {
            var slide = presentation.Slides[idx];
            var notesSlide = slide.NotesSlideManager.NotesSlide;
            if (notesSlide?.NotesTextFrame != null)
            {
                notesSlide.NotesTextFrame.Text = string.Empty;
                clearedCount++;
            }
        }

        MarkModified(context);

        return Success($"Cleared speaker notes for {clearedCount} slides (of {targets.Length} targeted).");
    }

    /// <summary>
    ///     Validates that all slide indices are within the valid range.
    /// </summary>
    /// <param name="indices">The array of slide indices to validate.</param>
    /// <param name="slideCount">The total number of slides in the presentation.</param>
    /// <exception cref="ArgumentException">Thrown when any index is outside the valid range.</exception>
    private static void ValidateSlideIndices(int[] indices, int slideCount)
    {
        var invalidIndices = indices.Where(idx => idx < 0 || idx >= slideCount).ToList();
        if (invalidIndices.Count > 0)
            throw new ArgumentException(
                $"Invalid slide indices: [{string.Join(", ", invalidIndices)}]. Valid range: 0 to {slideCount - 1}");
    }

    /// <summary>
    ///     Extracts clear notes parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted clear notes parameters.</returns>
    private static ClearNotesParameters ExtractClearNotesParameters(OperationParameters parameters)
    {
        return new ClearNotesParameters(parameters.GetOptional<int[]?>("slideIndices"));
    }

    /// <summary>
    ///     Record for holding clear notes parameters.
    /// </summary>
    /// <param name="SlideIndices">The array of slide indices to clear.</param>
    private record ClearNotesParameters(int[]? SlideIndices);
}
