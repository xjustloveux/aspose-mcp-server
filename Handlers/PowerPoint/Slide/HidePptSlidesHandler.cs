using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Slide;

/// <summary>
///     Handler for hiding or showing slides in PowerPoint presentations.
/// </summary>
public class HidePptSlidesHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "hide";

    /// <summary>
    ///     Hides or shows slides in the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: hidden (true to hide, false to show)
    ///     Optional: slideIndices (JSON array of indices, default: all slides)
    /// </param>
    /// <returns>Success message with hide/show details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var hidden = parameters.GetOptional("hidden", false);
        var slideIndicesJson = parameters.GetOptional<string?>("slideIndices");
        var presentation = context.Document;

        int[] targets;
        if (!string.IsNullOrWhiteSpace(slideIndicesJson))
        {
            var indices = JsonSerializer.Deserialize<int[]>(slideIndicesJson);
            targets = indices ?? Enumerable.Range(0, presentation.Slides.Count).ToArray();
        }
        else
        {
            targets = Enumerable.Range(0, presentation.Slides.Count).ToArray();
        }

        foreach (var idx in targets)
            if (idx < 0 || idx >= presentation.Slides.Count)
                throw new ArgumentException($"slide index {idx} out of range");

        foreach (var idx in targets)
            presentation.Slides[idx].Hidden = hidden;

        MarkModified(context);

        return Success($"Set {targets.Length} slide(s) hidden={hidden}.");
    }
}
